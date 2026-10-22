"""Structured drafts for the AI client; no invented procurement characteristics."""
from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, StrictInt, field_validator

from work_report import normalize_text


class Assignment(BaseModel):
    model_config = ConfigDict(extra="forbid")
    source_row: StrictInt = Field(ge=2, description="Номер строки исходного Excel")
    okpd2: str = Field(min_length=2, max_length=32, description="Код из search_okpd2")


class TechnicalSpecification(BaseModel):
    model_config = ConfigDict(extra="forbid")
    source_row: StrictInt = Field(ge=2)
    kind: Literal["товар", "работа", "услуга", "уточнить"] = "уточнить"
    requirements: list[str] = Field(default_factory=list, max_length=100)
    acceptance: str = Field(default="", max_length=10000)
    delivery: str = Field(default="", max_length=10000)
    warranty: str = Field(default="", max_length=10000)
    missing_information: list[str] = Field(default_factory=list, max_length=100)

    @field_validator("requirements", "missing_information")
    @classmethod
    def validate_lines(cls, values: list[str]) -> list[str]:
        for value in values:
            if not value.strip() or len(value) > 10000:
                raise ValueError("Пункт должен содержать от 1 до 10000 символов")
        if sum(len(value) + 1 for value in values) > 45000:
            raise ValueError("Суммарный текст пунктов слишком длинный для ячейки Google Sheets")
        return values


SPEC_HEADERS = [
    "Строка источника", "№ с/з", "Наименование", "ОКПД2", "Наименование ОКПД2",
    "Вид закупки", "Количество по заявке", "Ед. изм.", "Технические требования",
    "Приемка", "Условия поставки / выполнения", "Гарантия", "Требуется уточнить",
    "Статус ТЗ",
]


def prepare_drafts(positions: list[dict]) -> list[dict]:
    return [TechnicalSpecification(
        source_row=position["source_row"],
        requirements=[f"Наименование по заявке: {position['item_name']}"],
        missing_information=[
            "Вид закупки: товар, работа или услуга",
            "Подтвержденные технические характеристики и допустимые эквиваленты",
            "Адрес и срок поставки / выполнения (срок ОМТС не является сроком поставки)",
            "Порядок приемки и подтверждающие документы",
            "Требования к гарантии",
        ],
    ).model_dump() for position in positions]


def build_spec_rows(positions: list[dict], report_rows: list[list],
                    specifications: list[TechnicalSpecification]) -> list[list]:
    by_row = {}
    for spec in specifications:
        if spec.source_row in by_row:
            raise ValueError(f"ТЗ для строки {spec.source_row} указано повторно")
        by_row[spec.source_row] = spec
    required = {position["source_row"] for position in positions}
    if set(by_row) != required:
        raise ValueError(f"Нужен один черновик ТЗ на каждую позицию; пропуски: "
                         f"{sorted(required - set(by_row))}; лишние: {sorted(set(by_row) - required)}")
    rows = [SPEC_HEADERS.copy()]
    for position, report_row in zip(positions, report_rows[1:], strict=True):
        spec = by_row[position["source_row"]]
        missing = list(spec.missing_information)
        for absent, label in [
            (spec.kind == "уточнить", "Вид закупки"),
            (not spec.requirements, "Технические требования"),
            (not normalize_text(spec.acceptance), "Порядок приемки"),
            (not normalize_text(spec.delivery), "Условия поставки / выполнения"),
            (not normalize_text(spec.warranty), "Требования к гарантии"),
            (not position["unit"], "Единица измерения"),
            (position["quantity"] == "", "Количество"),
        ]:
            if absent and label not in missing:
                missing.append(label)
        rows.append([
            position["source_row"], position["request_number"], position["item_name"],
            report_row[19], report_row[20], spec.kind, position["quantity"], position["unit"],
            "\n".join(spec.requirements) or "Требуется уточнить", spec.acceptance or "Требуется уточнить",
            spec.delivery or "Требуется уточнить", spec.warranty or "Требуется уточнить",
            "\n".join(missing), "Черновик — требуется согласование специалистом",
        ])
    return rows


class PriceItem(BaseModel):
    model_config = ConfigDict(extra="forbid")
    source_row: StrictInt = Field(ge=2)
    unit_price: float = Field(gt=0, le=1e10, description="Цена за единицу, ₽, из search_prices")
    source_url: str = Field(min_length=8, max_length=2000, description="Ссылка на источник цены")
    note: str = Field(default="", max_length=2000, description="Например: аналог, цена от, без НДС")

    @field_validator("source_url")
    @classmethod
    def validate_url(cls, value: str) -> str:
        if not value.startswith(("http://", "https://")):
            raise ValueError("Источник цены — ссылка http(s)")
        return value


PRICE_HEADERS = ["Строка источника", "№ с/з", "Наименование", "Ед. изм.", "Потребность",
                 "Цена за ед., ₽ (ориентир)", "Сумма, ₽ (ориентир)", "Источник", "Примечание", "Статус"]


def _number(value) -> float | None:
    try:
        return float(str(value).replace(" ", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def build_price_rows(positions: list[dict], prices: list[PriceItem]) -> list[list]:
    """Лист «Цены»: ориентировочные цены по позициям отчёта; не все позиции обязательны."""
    by_row = {}
    for price in prices:
        if price.source_row in by_row:
            raise ValueError(f"Цена для строки {price.source_row} указана повторно")
        by_row[price.source_row] = price
    unknown = set(by_row) - {position["source_row"] for position in positions}
    if unknown:
        raise ValueError(f"Цены для строк не из отчёта: {sorted(unknown)}")
    rows = [PRICE_HEADERS.copy()]
    for position in positions:
        price = by_row.get(position["source_row"])
        if price is None:
            continue
        quantity = _number(position["quantity"])
        total = round(price.unit_price * quantity, 2) if quantity is not None else ""
        rows.append([position["source_row"], position["request_number"], position["item_name"],
                     position["unit"], position["quantity"], price.unit_price, total, price.source_url,
                     price.note, "Ориентир из интернета — требуется проверка специалистом"])
    return rows
