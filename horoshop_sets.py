from __future__ import annotations

import io
import json
import os
from dataclasses import dataclass
from datetime import UTC, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


DEFAULT_EXPORT_LIMIT = 500


class HoroshopSetsError(RuntimeError):
    pass


@dataclass(frozen=True)
class Settings:
    domain: str
    host: str
    port: int
    currency: str
    title: str
    batch_size: int
    request_timeout_seconds: int
    state_file: Path


@dataclass(frozen=True)
class Credentials:
    login: str
    password: str
    token: str = ""


@dataclass(frozen=True)
class SetRow:
    article: str
    display_articles: tuple[str, ...]
    discounted_price: Decimal
    row_number: int


@dataclass(frozen=True)
class CatalogProduct:
    article: str
    article_for_display: str


@dataclass(frozen=True)
class PlanItem:
    article: str
    display_articles: tuple[str, ...]
    discounted_price: Decimal | None
    products: tuple[str, ...]
    row_number: int
    error: str = ""

    @property
    def ready(self) -> bool:
        return not self.error


def utc_now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


def normalize(value: Any) -> str:
    return str(value or "").strip()


def endpoint_url(domain: str, endpoint: str) -> str:
    return urljoin(f"{domain.rstrip('/')}/", endpoint.lstrip("/"))


def load_settings(config_file: Path) -> Settings:
    with config_file.open("r", encoding="utf-8-sig") as file:
        raw = json.load(file)
    if not isinstance(raw, dict):
        raise ValueError("config.json must contain an object.")

    horoshop = raw.get("horoshop") or {}
    server = raw.get("server") or {}
    if not isinstance(horoshop, dict) or not isinstance(server, dict):
        raise ValueError("Sections horoshop and server must contain objects.")

    domain = normalize(horoshop.get("domain"))
    if not domain:
        raise ValueError("Set horoshop.domain in config.json.")
    currency = normalize(horoshop.get("currency", "UAH")).upper()
    if len(currency) != 3:
        raise ValueError("horoshop.currency must be a three-letter ISO code.")

    state_value = normalize(server.get("state_file", "data/sets_state.json"))
    state_file = Path(state_value)
    if not state_file.is_absolute():
        state_file = config_file.parent / state_file

    return Settings(
        domain=domain.rstrip("/"),
        host=normalize(server.get("host", "0.0.0.0")) or "0.0.0.0",
        port=max(1, min(65535, int(server.get("port", 8093)))),
        currency=currency,
        title=normalize(horoshop.get("title", "Разом дешевше")) or "Разом дешевше",
        batch_size=max(1, int(horoshop.get("batch_size", 50))),
        request_timeout_seconds=max(
            1, int(horoshop.get("request_timeout_seconds", 60))
        ),
        state_file=state_file,
    )


def parse_price(value: Any) -> Decimal:
    text = normalize(value).replace(" ", "").replace(",", ".")
    try:
        price = Decimal(text)
    except (InvalidOperation, ValueError) as error:
        raise ValueError("ціна має бути числом, більшим за нуль.") from error
    if price <= 0:
        raise ValueError("ціна має бути більшою за нуль.")
    return price.quantize(Decimal("0.01"))


def split_display_articles(value: Any) -> tuple[str, ...]:
    values = tuple(item.strip() for item in normalize(value).split(";") if item.strip())
    if len(values) < 2:
        raise ValueError("у наборі має бути щонайменше два артикули товарів.")
    if len({item.casefold() for item in values}) != len(values):
        raise ValueError("артикули товарів у наборі не можуть повторюватися.")
    return values


def parse_excel_sets(data: bytes) -> list[SetRow]:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        rows: list[SetRow] = []
        seen_articles: set[str] = set()
        for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            if not row or all(value is None for value in row[:3]):
                continue
            article = normalize(row[0] if len(row) > 0 else "")
            if article.casefold() in {"article", "артикул", "артикул набору", "артикул набора"}:
                continue
            if not article:
                raise ValueError(f"Рядок {row_number}: вкажіть артикул набору.")
            key = article.casefold()
            if key in seen_articles:
                raise ValueError(f"Рядок {row_number}: артикул набору '{article}' повторюється.")
            seen_articles.add(key)
            try:
                display_articles = split_display_articles(row[1] if len(row) > 1 else "")
                price = parse_price(row[2] if len(row) > 2 else "")
            except ValueError as error:
                raise ValueError(f"Рядок {row_number}: {error}") from error
            rows.append(SetRow(article, display_articles, price, row_number))
        if not rows:
            raise ValueError("Excel не містить жодного набору.")
        return rows
    finally:
        workbook.close()


def build_excel_template() -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Набори"
    worksheet.append(["Артикул набору", "Артикули товарів", "Ціна набору"])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = "A1:C1"
    worksheet.column_dimensions["A"].width = 28
    worksheet.column_dimensions["B"].width = 58
    worksheet.column_dimensions["C"].width = 18

    header_fill = PatternFill("solid", fgColor="166534")
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for row in range(2, 102):
        worksheet.cell(row=row, column=1).number_format = "@"
        worksheet.cell(row=row, column=2).number_format = "@"
        worksheet.cell(row=row, column=3).number_format = "0.00"

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class CatalogIndex:
    def __init__(self, products: list[CatalogProduct]):
        self.products = products
        self.product_articles = {product.article for product in products}
        self.by_display: dict[str, list[CatalogProduct]] = {}
        self.by_display_folded: dict[str, list[CatalogProduct]] = {}
        for product in products:
            if not product.article_for_display:
                continue
            self.by_display.setdefault(product.article_for_display, []).append(product)
            self.by_display_folded.setdefault(
                product.article_for_display.casefold(), []
            ).append(product)

    @classmethod
    def from_raw(cls, raw_products: list[dict[str, Any]]) -> "CatalogIndex":
        products = []
        for raw in raw_products:
            if not isinstance(raw, dict):
                continue
            article = normalize(raw.get("article"))
            if article:
                products.append(
                    CatalogProduct(
                        article=article,
                        article_for_display=normalize(raw.get("article_for_display")),
                    )
                )
        return cls(products)

    def resolve_display_article(self, display_article: str) -> tuple[str | None, str]:
        exact = self.by_display.get(display_article, [])
        if len(exact) == 1:
            return exact[0].article, ""
        if len(exact) > 1:
            return None, f"Артикул відображення '{display_article}' не є унікальним."
        insensitive = self.by_display_folded.get(display_article.casefold(), [])
        if len(insensitive) == 1:
            return insensitive[0].article, ""
        if len(insensitive) > 1:
            return None, f"Артикул відображення '{display_article}' не є унікальним."
        return None, f"Артикул відображення '{display_article}' не знайдений у каталозі."


def prepare_plan(rows: list[SetRow], catalog: CatalogIndex) -> list[PlanItem]:
    plan: list[PlanItem] = []
    for row in rows:
        if row.article in catalog.product_articles:
            plan.append(
                PlanItem(
                    article=row.article,
                    display_articles=row.display_articles,
                    discounted_price=row.discounted_price,
                    products=(),
                    row_number=row.row_number,
                    error="Артикул набору збігається з артикулом наявного товару.",
                )
            )
            continue

        products: list[str] = []
        error = ""
        for display_article in row.display_articles:
            product_article, error = catalog.resolve_display_article(display_article)
            if error:
                break
            if product_article:
                products.append(product_article)
        if not error and len(set(products)) != len(products):
            error = "Кілька артикулів відображення вказують на один товар."
        plan.append(
            PlanItem(
                article=row.article,
                display_articles=row.display_articles,
                discounted_price=row.discounted_price,
                products=tuple(products) if not error else (),
                row_number=row.row_number,
                error=error,
            )
        )
    return plan


def import_payload(items: list[PlanItem], settings: Settings) -> list[dict[str, Any]]:
    return [
        {
            "article": item.article,
            "title": settings.title,
            "discountedPrice": float(item.discounted_price or Decimal("0")),
            "currency": settings.currency,
            "enabled": True,
            "products": list(item.products),
        }
        for item in items
        if item.ready
    ]


class HoroshopClient:
    def __init__(self, settings: Settings, credentials: Credentials):
        self.settings = settings
        self.credentials = credentials
        self.session = requests.Session()
        self._token = credentials.token

    def _post(self, endpoint: str, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            response = self.session.post(
                endpoint_url(self.settings.domain, endpoint),
                json=payload,
                timeout=self.settings.request_timeout_seconds,
            )
            response.raise_for_status()
            data = response.json()
        except requests.RequestException as error:
            raise HoroshopSetsError(f"Horoshop API request failed: {error}") from error
        except ValueError as error:
            raise HoroshopSetsError("Horoshop API returned a non-JSON response.") from error
        if not isinstance(data, dict):
            raise HoroshopSetsError("Horoshop API returned an invalid JSON response.")
        if str(data.get("status", "")).upper() in {"ERROR", "EXCEPTION"}:
            raise HoroshopSetsError(str(data))
        return data

    def token(self) -> str:
        if self._token:
            return self._token
        if not self.credentials.login or not self.credentials.password:
            raise ValueError("Enter Horoshop login and password.")
        response = self._post(
            "/api/auth/",
            {"login": self.credentials.login, "password": self.credentials.password},
        )
        token = response.get("response", {}).get("token")
        if not token:
            raise HoroshopSetsError("Horoshop did not return an authorization token.")
        self._token = str(token)
        return self._token

    def export_catalog(self) -> list[dict[str, Any]]:
        offset = 0
        products: list[dict[str, Any]] = []
        while True:
            response = self._post(
                "/api/catalog/export/",
                {
                    "token": self.token(),
                    "offset": offset,
                    "limit": DEFAULT_EXPORT_LIMIT,
                    "includedParams": ["article_for_display"],
                },
            )
            nested = response.get("response")
            page = nested.get("products") if isinstance(nested, dict) else response.get("products")
            if not isinstance(page, list):
                raise HoroshopSetsError("Catalog export did not contain products.")
            products.extend(item for item in page if isinstance(item, dict))
            if len(page) < DEFAULT_EXPORT_LIMIT:
                return products
            offset += DEFAULT_EXPORT_LIMIT

    def import_sets(self, items: list[dict[str, Any]]) -> dict[str, Any]:
        if not items:
            return {"status": "OK", "response": {"log": []}}
        return self._post(
            "/api/productSet/import/",
            {"token": self.token(), "items": items},
        )


def import_results(response: dict[str, Any]) -> dict[str, tuple[bool, str]]:
    status = str(response.get("status", "")).upper()
    nested = response.get("response")
    log_items = nested.get("log", []) if isinstance(nested, dict) else []
    results: dict[str, tuple[bool, str]] = {}
    if isinstance(log_items, list):
        for entry in log_items:
            if not isinstance(entry, dict):
                continue
            article = normalize(entry.get("article"))
            info = entry.get("info", [])
            messages = []
            codes = []
            if isinstance(info, list):
                for item in info:
                    if isinstance(item, dict):
                        codes.append(item.get("code"))
                        messages.append(normalize(item.get("message")))
            success = 0 in codes or (status == "OK" and not codes)
            results[article] = (success, "; ".join(message for message in messages if message))
    return results


class StateStore:
    def __init__(self, path: Path):
        self.path = path
        self.sets: dict[str, dict[str, Any]] = {}
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            return
        with self.path.open("r", encoding="utf-8") as file:
            data = json.load(file)
        if isinstance(data, dict) and isinstance(data.get("sets"), dict):
            self.sets = {
                normalize(article): value
                for article, value in data["sets"].items()
                if normalize(article) and isinstance(value, dict)
            }

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(f"{self.path.suffix}.tmp")
        with temp.open("w", encoding="utf-8") as file:
            json.dump(
                {"updated_at": utc_now(), "sets": self.sets},
                file,
                ensure_ascii=False,
                indent=2,
            )
        os.replace(temp, self.path)

    def record(self, item: PlanItem, status: str, message: str) -> None:
        self.sets[item.article] = {
            "article": item.article,
            "display_articles": list(item.display_articles),
            "products": list(item.products),
            "discounted_price": str(item.discounted_price or ""),
            "status": status,
            "message": message,
            "updated_at": utc_now(),
        }

    def snapshot(self) -> list[dict[str, Any]]:
        return sorted(
            self.sets.values(),
            key=lambda item: str(item.get("updated_at", "")),
            reverse=True,
        )
