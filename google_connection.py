"""Non-interactive authentication for the systemd service."""
from __future__ import annotations

import contextvars
import hmac
import json
import logging
import os
import tempfile
import threading
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path

import httplib2
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_httplib2 import AuthorizedHttp
from googleapiclient.discovery import build
from mcp_google_sheets.server import SpreadsheetContext, SCOPES

LOGGER = logging.getLogger("google-sheets-mcp")
FOLDER_MIME = "application/vnd.google-apps.folder"
# Локальный MCP: читать таблицу заявок и создавать свои отчёты. drive.file вместо drive —
# только файлы, созданные приложением (не restricted-доступ ко всему Диску).
LOCAL_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"]


def build_services(creds):
    """Sheets и Drive с таймаутом: без него недоступный Google вешает вызов на минуты."""
    timeout = int(os.environ.get("GOOGLE_HTTP_TIMEOUT", "30"))

    def http():
        return AuthorizedHttp(creds, http=httplib2.Http(timeout=timeout))
    return (build("sheets", "v4", http=http(), cache_discovery=False),
            build("drive", "v3", http=http(), cache_discovery=False))


def reports_folder_name() -> str:
    return os.environ.get("MCP_DRIVE_FOLDER", "Отчёты MCP").strip()


def _quote(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


class GoogleNotConnected:
    """Keep local MCP tools usable until account authorization is configured."""
    def __getattr__(self, _name):
        raise RuntimeError(
            "Google пока не подключен: сервер опубликован в режиме подготовки. "
            "Доступны заявки, ОКПД2 и черновики ТЗ; запись в Google Диск будет "
            "доступна после настройки авторизации."
        )


@dataclass(frozen=True)
class Caller:
    """Сотрудник, от имени которого идёт MCP-запрос."""
    login: str
    email: str = ""
    name: str = ""
    fio: str = ""  # «Фамилия И.О.», как в столбце S заявок
    google_token_path: Path | None = None


# Вызывающий сотрудник; выставляет BearerAuthMiddleware до входа в MCP-приложение.
# None — админ-токен без привязки к сотруднику.
current_user: contextvars.ContextVar[Caller | None] = contextvars.ContextVar(
    "mcp_google_user", default=None)


def users_file_path() -> Path:
    explicit = os.environ.get("USERS_FILE", "")
    if explicit:
        return Path(explicit)
    return Path(os.environ.get("TOKEN_PATH", "token.json")).parent / "users.json"


class UsersRegistry:
    """Карта «MCP-токен → логин → Google token.json» с перечитыванием по mtime."""

    def __init__(self, path: Path) -> None:
        self._path = path
        self._mtime = None
        self._by_token: dict[str, str] = {}
        self._entries: dict[str, dict] = {}
        self._lock = threading.Lock()

    def _reload(self) -> None:
        try:
            mtime = self._path.stat().st_mtime_ns
        except OSError:
            mtime = None
        if mtime == self._mtime:
            return
        by_token: dict[str, str] = {}
        entries: dict[str, dict] = {}
        if mtime is not None:
            try:
                with self._path.open(encoding="utf-8") as stream:
                    data = json.load(stream)
            except (OSError, ValueError):
                LOGGER.exception("Не удалось прочитать %s", self._path)
            else:
                for login, info in (data.get("users") or {}).items():
                    if not isinstance(info, dict):
                        continue
                    entries[str(login)] = info
                    token = str(info.get("token", ""))
                    if len(token) >= 32:
                        by_token[token] = str(login)
        with self._lock:
            self._by_token = by_token
            self._entries = entries
            self._mtime = mtime

    def user_for_token(self, token: str) -> str | None:
        self._reload()
        with self._lock:
            pairs = list(self._by_token.items())
        for known, login in pairs:
            if hmac.compare_digest(token, known):
                return login
        return None

    def caller_for_token(self, token: str) -> Caller | None:
        login = self.user_for_token(token)
        if login is None:
            return None
        with self._lock:
            info = self._entries.get(login) or {}
        return Caller(login=login, email=str(info.get("email") or ""),
                      google_token_path=self.google_token_path(login))

    def google_token_path(self, login: str) -> Path:
        self._reload()
        with self._lock:
            info = self._entries.get(login) or {}
        relative = str(info.get("google_token") or f"users/{login}.json")
        if os.path.isabs(relative) or ".." in Path(relative).parts:
            raise RuntimeError(f"Некорректный путь Google token для {login}")
        return self._path.parent / relative


def save_token(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent,
                                         prefix=".google-token-", delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(text)
        os.chmod(temporary, 0o600)
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def load_credentials(path: Path | None = None):
    if path is None:
        account = os.environ.get("SERVICE_ACCOUNT_PATH", "service_account.json")
        if account and Path(account).is_file():
            # Do not silently fall back to a different Google identity on failure.
            return service_account.Credentials.from_service_account_file(account, scopes=SCOPES)
    # Токен конкретного сотрудника обновляем с теми scope, что он выдал.
    scopes = None if path is not None else SCOPES
    path = path or Path(os.environ.get("TOKEN_PATH", "token.json"))
    if not path.is_file():
        raise RuntimeError(f"Нет Google token {path.name}. Администратору нужно авторизовать Google заново.")
    with path.open(encoding="utf-8-sig") as stream:
        credentials = Credentials.from_authorized_user_info(json.load(stream), scopes)
    if not credentials.valid:
        if not credentials.refresh_token:
            raise RuntimeError(f"Нет refresh_token Google ({path.name}). Администратору нужно "
                               "авторизовать Google заново.")
        try:
            credentials.refresh(Request())
        except Exception:
            raise RuntimeError("Не удалось обновить доступ Google: проверьте сеть сервера или "
                               "авторизуйте Google заново (доступ мог быть отозван).") from None
        save_token(path, credentials.to_json())
    return credentials


def _caller_or_error() -> Caller:
    caller = current_user.get()
    if caller is None:
        raise RuntimeError("Этот MCP-токен не привязан к сотруднику. Войдите на портал под своей "
                           "учётной записью и переустановите подключение.")
    return caller


class SharedAccountContext:
    """Один Google-аккаунт компании; у каждого сотрудника своя папка.

    Таблицы создаются в «<MCP_DRIVE_FOLDER>/<ФИО> (<логин>)». Папка помечена
    appProperties.mcp_login, поэтому находится и после переименования. Доступ к
    таблице разрешён, только если она лежит в папке вызывающего сотрудника.
    Сотруднику папка открывается по ссылке (MCP_SHARE_LINK) и по почте
    (MCP_SHARE_EMAIL); права на папку наследуются вложенными файлами.
    """

    def __init__(self) -> None:
        self._services = None
        self._folders: dict[str, tuple[str, str]] = {}  # логин → (id папки, имя)
        self._lock = threading.Lock()

    def _build(self):
        with self._lock:
            if self._services is None:
                self._services = build_services(load_credentials())
            return self._services

    @property
    def sheets_service(self):
        _caller_or_error()
        return self._build()[0]

    @property
    def drive_service(self):
        _caller_or_error()
        return self._build()[1]

    def _find_folder(self, drive, query: str) -> str | None:
        found = drive.files().list(q=query + f" and mimeType = '{FOLDER_MIME}' and trashed = false",
                                   spaces="drive", fields="files(id)", pageSize=1).execute().get("files", [])
        return found[0]["id"] if found else None

    def _root_folder(self, drive) -> str:
        name = reports_folder_name() or "Отчёты MCP"
        query = "appProperties has { key='mcp_role' and value='root' }"
        root = self._find_folder(drive, query)
        if root is None:
            root = drive.files().create(body={"name": name, "mimeType": FOLDER_MIME,
                                              "appProperties": {"mcp_role": "root"}},
                                        fields="id").execute()["id"]
            LOGGER.info("Создана общая папка «%s»", name)
        return root

    def _share(self, drive, folder_id: str, caller: Caller) -> None:
        link_role = os.environ.get("MCP_SHARE_LINK", "reader")
        email_role = os.environ.get("MCP_SHARE_EMAIL", "writer")
        if link_role in {"reader", "writer"}:
            drive.permissions().create(fileId=folder_id, body={"type": "anyone", "role": link_role},
                                       fields="id").execute()
        if email_role in {"reader", "writer"} and caller.email:
            try:
                drive.permissions().create(
                    fileId=folder_id, sendNotificationEmail=False, fields="id",
                    body={"type": "user", "role": email_role, "emailAddress": caller.email}).execute()
            except Exception as error:  # почта может не быть Google-аккаунтом
                LOGGER.warning("Не удалось выдать доступ %s к папке: %s", caller.email, error)

    def folder_for(self, caller: Caller) -> str:
        title = f"{caller.fio or caller.name or caller.login} ({caller.login})"
        cached = self._folders.get(caller.login)
        if cached and cached[1] == title:  # ФИО поменялось — проверим имя папки заново
            return cached[0]
        drive = self._build()[1]
        login = _quote(caller.login)
        folder = self._find_folder(drive, f"appProperties has {{ key='mcp_login' and value='{login}' }}")
        if folder is not None:
            # Имя могло быть собрано по неполным данным (например, без отчества): поправляем.
            current = drive.files().get(fileId=folder, fields="name").execute().get("name")
            if current != title:
                drive.files().update(fileId=folder, body={"name": title}, fields="id").execute()
                LOGGER.info("Папка «%s» переименована в «%s»", current, title)
        if folder is None:
            folder = drive.files().create(body={
                "name": title, "mimeType": FOLDER_MIME, "parents": [self._root_folder(drive)],
                "appProperties": {"mcp_login": caller.login},
            }, fields="id").execute()["id"]
            self._share(drive, folder, caller)
            LOGGER.info("Создана папка «%s»", title)
        self._folders[caller.login] = (folder, title)
        return folder

    @property
    def folder_id(self) -> str:
        return self.folder_for(_caller_or_error())

    def check_access(self, spreadsheet_id: str) -> None:
        """Разрешает работу только с файлами из папки вызывающего сотрудника."""
        folder = self.folder_id
        parents = self._build()[1].files().get(
            fileId=spreadsheet_id, fields="parents", supportsAllDrives=True).execute().get("parents", [])
        if folder not in parents:
            raise PermissionError("Эта таблица не из вашей папки: открывать можно только таблицы, "
                                  "созданные через MCP для вас.")


class MultiUserSpreadsheetContext:
    """Устаревший режим enroll_user.py: у каждого сотрудника свой Google token.

    Инструменты читают lifespan_context синхронно внутри запроса, поэтому прокси
    строит clients лениво по ContextVar и кеширует их, пока не изменится файл
    Google-токена.
    """

    def __init__(self) -> None:
        self._services: dict[Path, tuple] = {}
        self._lock = threading.Lock()

    def _services_for_caller(self):
        caller = current_user.get()
        if caller is None or caller.google_token_path is None:
            raise RuntimeError(
                "Этот MCP-токен не привязан к Google-аккаунту сотрудника. "
                "Попросите администратора выполнить enroll_user.py для вашей учётной записи."
            )
        path = caller.google_token_path
        try:
            version = path.stat().st_mtime_ns
        except OSError:
            version = None
        with self._lock:
            cached = self._services.get(path)
            if cached is None or cached[0] != version:
                creds = load_credentials(path)
                try:
                    # load_credentials мог обновить и переписать файл.
                    version = path.stat().st_mtime_ns
                except OSError:
                    pass
                cached = (version, *build_services(creds))
                self._services[path] = cached
                LOGGER.info("Google подключен для %s", caller.email or caller.login)
            return cached[1], cached[2]

    @property
    def sheets_service(self):
        return self._services_for_caller()[0]

    @property
    def drive_service(self):
        return self._services_for_caller()[1]

    @property
    def folder_id(self) -> str | None:
        """Папка отчётов на Диске сотрудника: находит или создаёт при первом обращении."""
        name = reports_folder_name()
        if not name:
            return None
        drive = self.drive_service
        found = drive.files().list(
            q=f"name = '{_quote(name)}' and mimeType = '{FOLDER_MIME}' and trashed = false",
            spaces="drive", fields="files(id)", pageSize=1,
        ).execute().get("files", [])
        if found:
            return found[0]["id"]
        created = drive.files().create(body={"name": name, "mimeType": FOLDER_MIME}, fields="id").execute()
        return created["id"]


class LocalAccountContext:
    """MCP на ПК сотрудника (stdio): его собственный Google-аккаунт.

    Токен лежит в TOKEN_PATH и обновляется с теми scope, что выдал сотрудник
    (LOCAL_SCOPES). Отчёты создаются в папке MCP_DRIVE_FOLDER на его Диске.
    """

    def __init__(self) -> None:
        self._services = None
        self._folder: str | None = None
        self._email: str | None = None
        self._lock = threading.Lock()

    def _build(self):
        with self._lock:
            if self._services is None:
                path = Path(os.environ.get("TOKEN_PATH", "token.json"))
                self._services = build_services(load_credentials(path))
            return self._services

    @property
    def sheets_service(self):
        return self._build()[0]

    @property
    def drive_service(self):
        return self._build()[1]

    @property
    def folder_id(self) -> str | None:
        name = reports_folder_name()
        if not name:
            return None
        if self._folder is None:
            drive = self.drive_service
            # С drive.file поиск видит только файлы приложения: найдём свою папку, где бы она ни лежала.
            found = drive.files().list(
                q=f"name = '{_quote(name)}' and mimeType = '{FOLDER_MIME}' and trashed = false",
                spaces="drive", fields="files(id)", pageSize=1,
            ).execute().get("files", [])
            self._folder = found[0]["id"] if found else drive.files().create(
                body={"name": name, "mimeType": FOLDER_MIME}, fields="id").execute()["id"]
        return self._folder

    def account_email(self) -> str | None:
        if self._email is None:
            self._email = self.drive_service.about().get(fields="user(emailAddress)").execute()["user"]["emailAddress"]
        return self._email


# Stateless HTTP входит в lifespan на каждый запрос: контексты общие на процесс.
_MULTIUSER_CONTEXT = MultiUserSpreadsheetContext()
_SHARED_CONTEXT = SharedAccountContext()
LOCAL_CONTEXT = LocalAccountContext()


@asynccontextmanager
async def google_lifespan(_server):
    mode = os.environ.get("GOOGLE_AUTH_MODE", "personal")
    if mode == "disabled":
        yield SpreadsheetContext(sheets_service=GoogleNotConnected(),
                                 drive_service=GoogleNotConnected(), folder_id=None)
        return
    if mode == "shared":
        yield _SHARED_CONTEXT
        return
    if mode == "multiuser":
        yield _MULTIUSER_CONTEXT
        return
    if mode == "local":
        yield LOCAL_CONTEXT
        return
    if mode != "personal":
        raise RuntimeError("Неизвестный GOOGLE_AUTH_MODE")
    sheets, drive = build_services(load_credentials())
    try:
        yield SpreadsheetContext(sheets_service=sheets, drive_service=drive,
                                 folder_id=os.environ.get("DRIVE_FOLDER_ID") or None)
    finally:
        sheets.close()
        drive.close()
