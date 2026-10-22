from __future__ import annotations

import functools
import hmac
import logging
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urlsplit

import httplib2

os.environ.setdefault(
    "ENABLED_TOOLS",
    "create_spreadsheet,get_sheet_data,get_multiple_sheet_data,update_cells,"
    "batch_update_cells,list_spreadsheets,list_sheets,create_sheet,search_spreadsheets",
)

from mcp.server.fastmcp import Context, FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from mcp.types import ToolAnnotations
from mcp_google_sheets import server as google_tools
from mcp_google_sheets.server import _configure_logging
from starlette.responses import JSONResponse
from starlette.types import Receive, Scope, Send

from work_report import GoogleSheetReport, WorkReport, assignments_from_items, normalize_text
from technical_spec import Assignment, TechnicalSpecification, PriceItem, prepare_drafts, build_spec_rows, build_price_rows
from google_connection import LOCAL_CONTEXT, UsersRegistry, current_user, google_lifespan, users_file_path
from portal import Portal, PortalConfig
import price_search


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
WORKBOOK_PATH = Path(os.environ.get("WORKBOOK_PATH", DATA_DIR / "requests.xlsx"))
OKPD_PATH = Path(os.environ.get("OKPD_PATH", DATA_DIR / "okpd.xlsx"))
WORK_SHEET_NAME = os.environ.get("WORK_SHEET_NAME", "незакр 202526")
LOCAL_MODE = os.environ.get("GOOGLE_AUTH_MODE") == "local"
# На ПК сотрудника заявки читаются из живой Google-таблицы его же аккаунтом; на сервере — из выгрузки.
REQUESTS_SPREADSHEET_ID = os.environ.get("REQUESTS_SPREADSHEET_ID", "")
if REQUESTS_SPREADSHEET_ID:
    REPORT = GoogleSheetReport(REQUESTS_SPREADSHEET_ID, WORK_SHEET_NAME, OKPD_PATH,
                               lambda: LOCAL_CONTEXT.sheets_service)
else:
    REPORT = WorkReport(WORKBOOK_PATH, OKPD_PATH, WORK_SHEET_NAME)
LOGGER = logging.getLogger("google-sheets-mcp")

INSTRUCTIONS = """Работайте по заявкам только со статусом «В работе».
1. get_work_positions: точное ФИО сотрудника в сегментах S; руководитель — первый сегмент.
Получите все страницы, сохраните source_row и source_version.
2. search_okpd2: ищите по назначению, при необходимости используйте синонимы.
Выбирайте код семантически, различайте товар и услугу; score не подтверждает выбор.
Не придумывайте коды, не выбирайте автоматически первый результат.
3. Для ТЗ вызовите prepare_technical_specification и заполните specifications по исходным
данным и требованиям пользователя. Не выдумывайте размеры, ГОСТ, сроки или гарантию;
неизвестные требования перечисляйте в missing_information. Количество в заявке —
потребность, оно не обязательно равно остатку к закупке. ТЗ всегда является черновиком.
4. create_work_report: передайте ровно один assignment на каждую позицию всей выборки,
source_version из шага 1; если нужно ТЗ — также specifications на каждую позицию.
Файл содержит лист «В работе» и при наличии specifications лист «ТЗ».
Повторный вызов создаёт новый файл, поэтому после успеха верните URL и не дублируйте вызов.
"""
if os.environ.get("GOOGLE_AUTH_MODE") == "disabled":
    INSTRUCTIONS = ("Сервер находится в режиме подготовки: Google пока не подключен. "
                    "Используйте get_work_positions, search_okpd2 и prepare_technical_specification. "
                    "Создание и чтение Google-таблиц недоступны до настройки авторизации.\n" + INSTRUCTIONS)
elif os.environ.get("GOOGLE_AUTH_MODE") == "shared":
    INSTRUCTIONS = ("Сотрудник вошёл корпоративной учётной записью. В начале работы вызовите whoami: "
                    "он вернёт ФИО, логин и ссылку на личную папку сотрудника с таблицами; "
                    "обращайтесь к сотруднику по имени. Если он просит «мои позиции», используйте "
                    "fio из whoami как employee в get_work_positions. Таблицы создаются общим "
                    "Google-аккаунтом компании в папке сотрудника; открыть можно только таблицы из "
                    "этой папки — чужую или старую таблицу по ссылке открыть нельзя, так и объясните. "
                    "После создания таблицы дайте сотруднику её URL и ссылку на папку.\n" + INSTRUCTIONS)
elif os.environ.get("GOOGLE_AUTH_MODE") == "multiuser":
    INSTRUCTIONS = ("Google подключен под аккаунтом вызывающего сотрудника: таблицы читаются "
                    "и создаются на его Диске и от его имени. В начале работы вызовите whoami, "
                    "чтобы знать, с кем говорите, и обращайтесь к сотруднику по имени.\n" + INSTRUCTIONS)
elif LOCAL_MODE:
    INSTRUCTIONS = ("MCP работает на ПК сотрудника. В начале работы вызовите whoami: он вернёт ФИО "
                    "(из учётной записи Windows), логин и Google-аккаунт; обращайтесь к сотруднику по имени. "
                    "Если он просит «мои позиции», передайте fio из whoami как employee в get_work_positions. "
                    "Заявки читаются из общей Google-таблицы; отчёты создаются на Google Диске сотрудника. "
                    "Цены: search_prices по каждой позиции даёт ориентировочные цены из интернета со ссылками; "
                    "передайте выбранные цены в create_work_report (prices) — появится лист «Цены». "
                    "Всегда указывайте источник цены и что это ориентир, не коммерческое предложение.\n"
                    + INSTRUCTIONS)
mcp = FastMCP("Google Sheets Procurement", instructions=INSTRUCTIONS, lifespan=google_lifespan)


def _guarded(name: str, tool):
    """В режиме shared пускает только к таблицам из папки вызывающего сотрудника."""
    @functools.wraps(tool)
    def wrapper(*args, **kwargs):
        ctx = kwargs.get("ctx")
        context = ctx.request_context.lifespan_context if ctx is not None else None
        check = getattr(context, "check_access", None)
        if check is not None:
            if name == "search_spreadsheets":
                raise PermissionError("Поиск по всему Диску недоступен: используйте list_spreadsheets "
                                      "(таблицы вашей папки).")
            if kwargs.get("folder_id") and kwargs["folder_id"] != context.folder_id:
                raise PermissionError("Можно работать только в своей папке: не передавайте folder_id.")
            if kwargs.get("spreadsheet_id"):
                check(kwargs["spreadsheet_id"])
            for query in kwargs.get("queries") or []:
                check(query.get("spreadsheet_id") if isinstance(query, dict) else query.spreadsheet_id)
        return tool(*args, **kwargs)
    return wrapper


for _name in os.environ["ENABLED_TOOLS"].split(","):
    _name = _name.strip()
    if not _name:
        continue
    if _name not in {"create_spreadsheet", "get_sheet_data", "get_multiple_sheet_data", "update_cells",
                     "batch_update_cells", "list_spreadsheets", "list_sheets", "create_sheet", "search_spreadsheets"}:
        raise ValueError(f"Неизвестный ENABLED_TOOLS: {_name}")
    mcp.add_tool(_guarded(_name, getattr(google_tools, _name)), annotations=ToolAnnotations(
        readOnlyHint=_name.startswith(("get_", "list_", "search_"))))


@mcp.tool(annotations=ToolAnnotations(title="Кто я", readOnlyHint=True))
def whoami(ctx: Context = None) -> dict[str, Any]:
    """Возвращает сотрудника, от имени которого подключён MCP: ФИО, логин, почту и папку.

    fio — в формате заявок («Фамилия И.О.»), подходит как employee для get_work_positions.
    folder_url — личная папка сотрудника с созданными для него Google-таблицами.
    Если login=null, это служебный токен без сотрудника: создавать таблицы нельзя.
    """
    caller = current_user.get()
    if caller is None:
        return {"login": None, "fio": None, "name": None, "email": None, "folder_url": None}
    result = {"login": caller.login, "fio": caller.fio or None, "name": caller.name or None,
              "email": caller.email or None, "folder_url": None}
    context = ctx.request_context.lifespan_context if ctx is not None else None
    if getattr(context, "check_access", None) is not None or isinstance(context, type(LOCAL_CONTEXT)):
        # Google нужен только для ссылки на папку: при его недоступности whoami всё равно отвечает.
        try:
            result["folder_url"] = f"https://drive.google.com/drive/folders/{context.folder_id}"
            if LOCAL_MODE:
                result["google_account"] = context.account_email()
        except Exception as error:
            result["google_error"] = google_error_text(error)
    if LOCAL_MODE and not caller.fio:
        result["note"] = ("ФИО не определено (Active Directory недоступен вне корпоративной сети). "
                          "Спросите сотрудника, как он записан в столбце S заявок («Фамилия И.О.»).")
    return result


def google_error_text(error: Exception) -> str:
    """Понятная причина сбоя Google вместо трассировки."""
    if isinstance(error, (TimeoutError, OSError, httplib2.HttpLib2Error)):
        return ("Нет связи с Google (googleapis.com): проверьте интернет на этом компьютере. "
                f"Подробности: {error.__class__.__name__}")
    return str(error)


def workbook_version() -> str:
    return REPORT.source_version()


@mcp.tool(
    annotations=ToolAnnotations(
        title="Позиции в работе сотрудника",
        readOnlyHint=True,
    )
)
def get_work_positions(
    employee: str,
    manager: Optional[str] = None,
    offset: int = 0,
    limit: int = 200,
) -> dict[str, Any]:
    """Возвращает позиции со статусом «В работе» из листа «незакр 202526».

    Сотрудник ищется по точному ФИО в любом сегменте столбца S, разделённом '/'.
    Можно передать полную пару «Кузнецова А.Б./Сидоров С.С.» либо employee=
    «Сидоров С.С.» и manager=«Кузнецова А.Б.». manager сравнивается только с
    первым сегментом S. Завершённые и отработанные снабжением строки не возвращаются.
    Сохраните source_row каждой позиции: он нужен для выбора ОКПД2 и создания отчёта.
    """
    if offset < 0:
        raise ValueError("offset не может быть отрицательным")
    if not 1 <= limit <= 500:
        raise ValueError("limit должен быть от 1 до 500")
    version = workbook_version()
    positions = REPORT.positions(employee, manager)
    if version != workbook_version():
        raise ValueError("Файл заявок изменился; повторите запрос")
    modified = REPORT.source_modified()
    return {
        "employee": normalize_text(employee),
        "manager": normalize_text(manager),
        "status": "В работе",
        "total": len(positions),
        "offset": offset,
        "limit": limit,
        "source_file": REPORT.source_name,
        "source_sheet": WORK_SHEET_NAME,
        "source_version": version,
        "source_modified_utc": modified.isoformat() if modified else None,
        "positions": positions[offset : offset + limit],
    }


@mcp.tool(
    annotations=ToolAnnotations(
        title="Поиск по классификатору ОКПД2",
        readOnlyHint=True,
    )
)
def search_okpd2(query: str, limit: int = 5) -> dict[str, Any]:
    """Ищет коды только в приложенном okpd.xlsx и возвращает до 20 кандидатов.

    Для каждой позиции выберите семантически подходящий наиболее конкретный код.
    Служба и товар различаются: для ремонта ищите услугу ремонта, для покупки — товар.
    Если выдача нерелевантна, повторите поиск с функциональным названием или синонимом
    (например, фаркоп — тягово-сцепное устройство). Не считайте score подтверждением
    правильности. При создании отчёта код будет проверен по классификатору ещё раз.
    """
    return {
        "query": normalize_text(query),
        "candidates": REPORT.search_okpd2(query, limit),
        "source_file": OKPD_PATH.name,
    }


def search_prices(query: str, limit: int = 8) -> dict[str, Any]:
    """Ищет в интернете ориентировочные цены позиции (₽) и возвращает ссылки на источники.

    query — наименование с ключевыми характеристиками (марка, размер, модель), без номера
    заявки. Возвращает results (заголовок, url, найденные цены) и summary_rub (min/median/max).
    Цены из выдачи и страниц магазинов: это ориентир, не коммерческое предложение — всегда
    показывайте источник. Маркетплейсы (Ozon, WB, Avito) часто не открываются программам.
    Вызывайте по одной позиции; при «ограничил частые запросы» повторите позже.
    """
    return price_search.search_prices(query, limit)


# Интернет есть только у ПК сотрудника: на сервере инструмент не регистрируется.
if LOCAL_MODE:
    mcp.tool(annotations=ToolAnnotations(title="Ориентировочные цены из интернета", readOnlyHint=True,
                                         openWorldHint=True))(search_prices)


@mcp.tool(annotations=ToolAnnotations(title="Подготовить черновики ТЗ", readOnlyHint=True))
def prepare_technical_specification(
    employee: str, manager: Optional[str] = None,
) -> dict[str, Any]:
    """Возвращает исходные данные и редактируемые черновики specifications для create_work_report.

    Заполните требования на основании заявки и уточнений пользователя. Неизвестные
    характеристики оставьте в missing_information; не выдавайте черновик за утвержденное ТЗ.
    ОКПД2 подбирается отдельно через search_okpd2. Файлы Google этот вызов не создаёт.
    """
    version = workbook_version()
    positions = REPORT.positions(employee, manager)
    if not positions:
        raise ValueError("Позиции 'В работе' для указанного сотрудника не найдены")
    if version != workbook_version():
        raise ValueError("Файл заявок изменился; повторите запрос")
    return {"source_version": version, "positions": positions,
            "specifications": prepare_drafts(positions), "review_required": True}


@mcp.tool(
    annotations=ToolAnnotations(
        title="Создать Google-таблицу позиций в работе",
        destructiveHint=True,
    )
)
def create_work_report(
    employee: str,
    assignments: list[Assignment],
    manager: Optional[str] = None,
    title: Optional[str] = None,
    folder_id: Optional[str] = None,
    specifications: Optional[list[TechnicalSpecification]] = None,
    source_version: Optional[str] = None,
    prices: Optional[list[PriceItem]] = None,
    ctx: Context = None,
) -> dict[str, Any]:
    """Создаёт новую отдельную Google-таблицу с позициями «В работе» и ОКПД2.

    Перед вызовом получите позиции через get_work_positions и подберите коды через
    search_okpd2. assignments — список объектов {"source_row": 123, "okpd2":
    "28.22.20.000"}; укажите ровно один существующий код для каждой возвращённой
    source_row, без пропусков и посторонних строк. В отчёт записываются официальное
    название кода и отметка о необходимости проверки специалистом.
    Для отдельного листа «ТЗ» передайте specifications из prepare_technical_specification,
    заполненные ИИ по подтвержденным данным. source_version из get_work_positions
    защищает от создания отчёта по изменившейся выгрузке.
    prices — ориентировочные цены из search_prices: {"source_row", "unit_price", "source_url",
    "note"}; можно не на все позиции. Появится лист «Цены» с суммой по потребности.
    """
    version = workbook_version()
    if source_version is not None and source_version != version:
        raise ValueError("Выгрузка изменилась; повторите get_work_positions и подбор кодов")
    selected = assignments_from_items([item.model_dump() for item in assignments])
    positions, values = REPORT.build_rows(employee, selected, manager)
    spec_values = build_spec_rows(positions, values, specifications) if specifications is not None else None
    price_values = build_price_rows(positions, prices) if prices else None
    if version != workbook_version():
        raise ValueError("Выгрузка изменилась во время подготовки; повторите запрос")
    report_title = normalize_text(title) or (
        f"Позиции в работе — {normalize_text(employee).replace('/', ' — ')} — "
        f"{datetime.now():%d.%m.%Y}"
    )
    if len(report_title) > 200:
        raise ValueError("Название таблицы не должно превышать 200 символов")
    if ctx is None:
        raise ValueError("Инструмент должен вызываться через MCP с Google-контекстом")
    context = ctx.request_context.lifespan_context
    sheets_service = context.sheets_service
    drive_service = context.drive_service
    if folder_id and getattr(context, "check_access", None) is not None and folder_id != context.folder_id:
        raise PermissionError("Отчёт создаётся только в вашей папке: не передавайте folder_id.")
    target_folder_id = folder_id or context.folder_id
    spreadsheet_id = None
    spec_sheet_id = 1
    try:
        body = {
                "properties": {"title": report_title},
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 0,
                            "title": "В работе",
                            "gridProperties": {
                                "rowCount": len(values) + 20,
                                "columnCount": len(values[0]),
                                "frozenRowCount": 1,
                            },
                        }
                    }
                ],
            }
        if spec_values:
            body["sheets"].append({"properties": {"sheetId": 1, "title": "ТЗ",
                "gridProperties": {"rowCount": len(spec_values) + 20,
                                   "columnCount": len(spec_values[0]), "frozenRowCount": 1}}})
        if target_folder_id:
            # Create directly in the folder, including a shared drive. A service
            # account has no personal Drive storage to create then move a file.
            created = drive_service.files().create(
                body={"name": report_title, "mimeType": "application/vnd.google-apps.spreadsheet",
                      "parents": [target_folder_id]},
                supportsAllDrives=True, fields="id",
            ).execute()
            spreadsheet_id = created["id"]
            existing = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id, fields="sheets.properties.sheetId").execute()
            initial_sheets = existing.get("sheets", [])
            requests = []
            if initial_sheets:
                sheet_id = initial_sheets[0]["properties"]["sheetId"]
                properties = dict(body["sheets"][0]["properties"], sheetId=sheet_id)
                requests.append({"updateSheetProperties": {"properties": properties,
                                "fields": "title,gridProperties"}})
            else:
                sheet_id = 0
                requests.append({"addSheet": body["sheets"][0]})
            if spec_values:
                # Do not assume the Drive-created sheet has ID zero.
                spec_properties = dict(body["sheets"][1]["properties"])
                spec_sheet_id = max([sheet_id] + [s["properties"]["sheetId"] for s in initial_sheets]) + 1
                spec_properties["sheetId"] = spec_sheet_id
                requests.append({"addSheet": {"properties": spec_properties}})
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
        else:
            spreadsheet = sheets_service.spreadsheets().create(
                body=body, fields="spreadsheetId,sheets.properties").execute()
            spreadsheet_id = spreadsheet["spreadsheetId"]
            sheet_id = spreadsheet["sheets"][0]["properties"]["sheetId"]
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range="'В работе'!A1",
            valueInputOption="RAW",
            body={"majorDimension": "ROWS", "values": values},
        ).execute()
        if spec_values:
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range="'ТЗ'!A1", valueInputOption="RAW",
                body={"majorDimension": "ROWS", "values": spec_values}).execute()
            spec_range = {"sheetId": spec_sheet_id, "startRowIndex": 0,
                          "endRowIndex": len(spec_values), "startColumnIndex": 0,
                          "endColumnIndex": len(spec_values[0])}
            sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [
                {"setBasicFilter": {"filter": {"range": spec_range}}},
                {"repeatCell": {"range": spec_range,
                    "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"}},
                    "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)"}},
                {"repeatCell": {"range": dict(spec_range, endRowIndex=1),
                    "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                    "fields": "userEnteredFormat.textFormat.bold"}},
                {"updateDimensionProperties": {
                    "range": {"sheetId": spec_sheet_id, "dimension": "COLUMNS", "startIndex": 8, "endIndex": 14},
                    "properties": {"pixelSize": 360}, "fields": "pixelSize"}},
            ]}).execute()
        sheet_range = {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": len(values),
            "startColumnIndex": 0,
            "endColumnIndex": len(values[0]),
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {"setBasicFilter": {"filter": {"range": sheet_range}}},
                    {
                        "repeatCell": {
                            "range": dict(sheet_range, endRowIndex=1),
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {"red": 0.16, "green": 0.36, "blue": 0.56},
                                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                                    "wrapStrategy": "WRAP",
                                    "verticalAlignment": "MIDDLE",
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)",
                        }
                    },
                    {
                        "repeatCell": {
                            "range": dict(sheet_range, startRowIndex=1),
                            "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"}},
                            "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 8},
                            "properties": {"pixelSize": 420},
                            "fields": "pixelSize",
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 16, "endIndex": 17},
                            "properties": {"pixelSize": 420},
                            "fields": "pixelSize",
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 20, "endIndex": 22},
                            "properties": {"pixelSize": 420},
                            "fields": "pixelSize",
                        }
                    },
                ]
            },
        ).execute()
        if price_values:
            sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [
                {"addSheet": {"properties": {"title": "Цены", "gridProperties": {
                    "rowCount": len(price_values) + 20, "columnCount": len(price_values[0]), "frozenRowCount": 1}}}},
            ]}).execute()
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range="'Цены'!A1", valueInputOption="RAW",
                body={"majorDimension": "ROWS", "values": price_values}).execute()
    except Exception as error:
        if spreadsheet_id:
            try:
                drive_service.files().update(fileId=spreadsheet_id, body={"trashed": True}, supportsAllDrives=True).execute()
            except Exception:
                LOGGER.exception("Не удалось отправить в корзину незавершённый отчёт %s", spreadsheet_id)
        if isinstance(error, (TimeoutError, OSError, httplib2.HttpLib2Error)):
            raise RuntimeError(google_error_text(error)) from None
        raise
    return {
        "spreadsheetId": spreadsheet_id,
        "title": report_title,
        "sheet": "В работе",
        "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
        "row_count": len(positions),
        "employee": normalize_text(employee),
        "manager": normalize_text(manager),
        "okpd2_review_required": True,
        "technical_specification_sheet": "ТЗ" if spec_values else None,
        "technical_specification_review_required": bool(spec_values),
        "prices_sheet": "Цены" if price_values else None,
        "source_version": version,
    }


@mcp.custom_route("/health", methods=["GET"])
async def health(_request) -> JSONResponse:
    return JSONResponse({"status": "ok", "service": "google-sheets-mcp",
                         "google_auth_mode": os.environ.get("GOOGLE_AUTH_MODE", "personal"),
                         "multiuser_ready": os.environ.get("GOOGLE_AUTH_MODE") == "multiuser",
                         "portal_ready": bool(os.environ.get("PORTAL_PUBLIC_URL")
                                              and os.environ.get("OIDC_ISSUER"))})


class BearerAuthMiddleware:
    """Bearer обязателен только для MCP; портал и /health проверяют доступ сами."""

    def __init__(self, app, token: bytes, registry: UsersRegistry | None = None,
                 portal: Portal | None = None) -> None:
        self.app = app
        self.token = token
        self.registry = registry
        self.portal = portal

    def _caller(self, credential: bytes):
        """None — токен неизвестен; False — админ-токен без Google; иначе Caller."""
        if hmac.compare_digest(credential, self.token):
            return False
        text = credential.decode("ascii", "ignore")
        if self.portal is not None:
            caller = self.portal.store.caller_for_token(text)
            if caller is not None:
                return caller
        if self.registry is not None:
            return self.registry.caller_for_token(text)
        return None

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        path = scope.get("path", "")
        if scope["type"] != "http" or not (path == "/mcp" or path.startswith("/mcp/")):
            await self.app(scope, receive, send)
            return
        headers = dict(scope.get("headers", []))
        scheme, _, credential = headers.get(b"authorization", b"").partition(b" ")
        caller = None
        if scheme.lower() == b"bearer" and credential.strip():
            caller = self._caller(credential.strip())
        if caller is None:
            response = JSONResponse(
                {"jsonrpc": "2.0", "id": None,
                 "error": {"code": -32001, "message": "Unauthorized"}},
                status_code=401,
                headers={"WWW-Authenticate": "Bearer"},
            )
            await response(scope, receive, send)
            return
        caller = caller or None  # админ-токен: доступ к MCP есть, Google-аккаунта нет
        scope["mcp.user"] = caller.login if caller else ""
        context_token = current_user.set(caller)
        try:
            await self.app(scope, receive, send)
        finally:
            current_user.reset(context_token)


def create_app():
    token = os.environ.get("MCP_AUTH_TOKEN", "").encode("ascii")
    if len(token) < 32:
        raise RuntimeError("MCP_AUTH_TOKEN должен содержать не менее 32 символов")
    port = int(os.environ.get("PORT", "8082"))
    public_host = os.environ.get("MCP_PUBLIC_HOST", "10.0.0.10")
    allowed_hosts = [f"{public_host}:{port}", f"127.0.0.1:{port}", f"localhost:{port}"]
    registry = None
    portal = None
    mode = os.environ.get("GOOGLE_AUTH_MODE", "personal")
    if mode == "multiuser":
        registry = UsersRegistry(users_file_path())
    # В personal-режиме у всех один Google без разделения папок: портал не включаем.
    config = PortalConfig.from_env() if mode in {"shared", "disabled", "multiuser"} else None
    if config is not None:
        portal = Portal(config)
        netloc = urlsplit(config.public_url).netloc
        allowed_hosts += [netloc, netloc.split(":")[0]]
    mcp.settings.stateless_http = True
    mcp.settings.json_response = True
    mcp.settings.transport_security = TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=allowed_hosts,
        allowed_origins=[],
    )
    # Менеджер сессий запускается один раз на экземпляр: новый на каждое приложение.
    mcp._session_manager = None
    app = mcp.streamable_http_app()
    if portal is not None:
        app.router.routes.extend(portal.routes())
    return BearerAuthMiddleware(app, token, registry, portal)


def local_state_dir() -> Path:
    return Path(os.environ.get("LOCAL_STATE_DIR") or Path.home() / ".config" / "zakupki-mcp")


def run_local() -> None:
    """MCP на ПК сотрудника: OpenCode запускает его сам (stdio), сотрудник — текущий пользователь Windows."""
    # stdout занят протоколом MCP: журнал только в stderr.
    logging.basicConfig(level=logging.WARNING, stream=sys.stderr, format="%(levelname)s %(message)s")
    from local_identity import windows_caller
    # Контекст копируется в задачи asyncio: все вызовы инструментов видят этого сотрудника.
    current_user.set(windows_caller(local_state_dir() / "identity.json"))
    mcp.run(transport="stdio")


def self_check() -> int:
    """Проверка установки: кто я, доступ к Google и к таблице заявок. Для установщика."""
    from local_identity import windows_caller
    caller = windows_caller(local_state_dir() / "identity.json")
    ok = True
    print(f"Сотрудник: {caller.name or '?'} (логин {caller.login}), ФИО для заявок: {caller.fio or 'не определено'}")
    if not caller.fio:
        print("  ! ФИО не найдено в Active Directory: ИИ спросит его у вас при первом запросе.")
    try:
        print(f"Google-аккаунт: {LOCAL_CONTEXT.account_email()}")
    except Exception as error:
        print(f"  ! Google недоступен: {google_error_text(error)}")
        return 1
    try:
        total = len(REPORT.positions(caller.fio)) if caller.fio else None
        print(f"Таблица заявок открывается: {REPORT.source_name}"
              + (f"; ваших позиций «В работе»: {total}" if total is not None else ""))
    except Exception as error:
        ok = False
        print(f"  ! Таблица заявок не открылась: {google_error_text(error)}")
        print("    Проверьте, что у вашего Google-аккаунта есть доступ к таблице заявок.")
    print(f"Классификатор ОКПД2: {len(REPORT.search_okpd2('перчатки', 1)) and 'загружен'}")
    return 0 if ok else 1


def main() -> None:
    if "--stdio" in sys.argv:
        run_local()
        return
    if "--check" in sys.argv:
        raise SystemExit(self_check())
    _configure_logging()
    import uvicorn

    uvicorn.run(
        create_app(),
        host=os.environ.get("HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "8082")),
        proxy_headers=False,
        server_header=False,
        ssl_certfile=os.environ.get("SSL_CERTFILE") or None,
        ssl_keyfile=os.environ.get("SSL_KEYFILE") or None,
    )


if __name__ == "__main__":
    main()
