from __future__ import annotations

import asyncio
import logging
import shutil
import sys
import threading
from dataclasses import replace
from decimal import Decimal
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from horoshop_sets import (
    ArticleMappings,
    CatalogIndex,
    Credentials,
    HoroshopClient,
    HoroshopSetsError,
    PRICE_BASE_CURRENT,
    PlanItem,
    SetRow,
    SET_TITLE,
    Settings,
    StateStore,
    build_excel_template,
    build_mappings_excel,
    build_set_article,
    build_state_excel,
    format_discount,
    import_payload,
    import_results,
    load_article_mappings,
    load_settings,
    normalize,
    parse_article_mappings,
    parse_discount_percent,
    parse_excel_sets,
    parse_price,
    parse_price_base,
    prepare_plan,
    remove_results,
    save_article_mappings,
    split_display_articles,
)


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAPPINGS_FILE = PROJECT_DIR / "data" / "article_mappings.json"

settings: Settings | None = None
state_store: StateStore | None = None
state_lock = threading.Lock()
mapping_lock = threading.Lock()
logger = logging.getLogger(__name__)
service_output_stream: "PublicLogStream | None" = None


class PublicLogStream:
    def __init__(self, path: Path, fallback_path: Path) -> None:
        self.path = path
        self.fallback_path = fallback_path
        self.encoding = "utf-8"

    def write(self, message: str) -> int:
        if not message:
            return 0
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with self.path.open("a", encoding=self.encoding) as file:
                file.write(message)
        except OSError:
            try:
                self.fallback_path.parent.mkdir(parents=True, exist_ok=True)
                with self.fallback_path.open("a", encoding=self.encoding) as file:
                    file.write(message)
            except OSError:
                pass
        return len(message)

    def flush(self) -> None:
        return None

    def isatty(self) -> bool:
        return False


def configure_service_output(runtime_settings: Settings) -> None:
    global service_output_stream

    requested_path = runtime_settings.public_log_file
    fallback_path = PROJECT_DIR / "logs" / "horoshop_sets.log"
    selected_path = requested_path
    fallback_message = ""
    try:
        selected_path.parent.mkdir(parents=True, exist_ok=True)
        selected_path.touch(exist_ok=True)
    except OSError as error:
        selected_path = fallback_path
        selected_path.parent.mkdir(parents=True, exist_ok=True)
        selected_path.touch(exist_ok=True)
        fallback_message = f"Public log path was unavailable ({requested_path}): {error}. Using {selected_path}."

    service_output_stream = PublicLogStream(selected_path, fallback_path)
    sys.stdout = service_output_stream
    sys.stderr = service_output_stream
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        stream=service_output_stream,
        force=True,
    )
    if fallback_message:
        logger.warning(fallback_message)
    logger.info("Service output is writing to %s", selected_path)


def get_runtime() -> tuple[Settings, StateStore]:
    global settings, state_store
    if settings is None or state_store is None:
        ensure_config_file()
        settings = load_settings(CONFIG_FILE)
        state_store = StateStore(settings.state_file)
    return settings, state_store


def ensure_config_file() -> None:
    if CONFIG_FILE.exists():
        return
    example_file = PROJECT_DIR / "config.example.json"
    if not example_file.exists():
        raise RuntimeError(f"Configuration template was not found: {example_file}")
    shutil.copyfile(example_file, CONFIG_FILE)
    print(f"Created configuration file from template: {CONFIG_FILE}")


def parse_flag(value: Any, default: bool = False) -> bool:
    text = normalize(value).casefold()
    if not text:
        return default
    return text not in {"0", "false", "ні", "нет", "no", "off"}


def stored_mappings() -> ArticleMappings:
    with mapping_lock:
        return load_article_mappings(MAPPINGS_FILE)


def add_to_stored_mappings(uploaded: ArticleMappings) -> ArticleMappings:
    with mapping_lock:
        combined = load_article_mappings(MAPPINGS_FILE).merged_with(uploaded)
        save_article_mappings(MAPPINGS_FILE, combined)
        return combined


async def uploaded_mappings(form: Any) -> ArticleMappings | None:
    mapping_file = form.get("mapping_file")
    if (
        mapping_file is None
        or not hasattr(mapping_file, "read")
        or not normalize(getattr(mapping_file, "filename", ""))
    ):
        return None
    if not str(getattr(mapping_file, "filename", "")).lower().endswith((".xlsx", ".xlsm")):
        raise HoroshopSetsError("Файл сопоставлень має бути .xlsx або .xlsm.")
    mapping_contents = await mapping_file.read()
    if not mapping_contents or len(mapping_contents) > MAX_UPLOAD_BYTES:
        raise HoroshopSetsError("Файл сопоставлень порожній або перевищує 20 МБ.")
    return await asyncio.to_thread(parse_article_mappings, mapping_contents)


def credentials_from_form(form: Any) -> Credentials:
    login = str(form.get("login", "")).strip()
    password = str(form.get("password", "")).strip()
    token = str(form.get("token", "")).strip()
    if not token and (not login or not password):
        raise HoroshopSetsError("Вкажіть логін і пароль API або чинний токен.")
    return Credentials(login=login, password=password, token=token)


async def upload_bytes(
    request: Request,
) -> tuple[bytes, Credentials, Decimal | None, str, bool, ArticleMappings]:
    form = await request.form()
    uploaded = form.get("file")
    if uploaded is None or not hasattr(uploaded, "read"):
        raise HoroshopSetsError("Оберіть Excel-файл .xlsx або .xlsm.")
    filename = str(getattr(uploaded, "filename", "")).lower()
    if not filename.endswith((".xlsx", ".xlsm")):
        raise HoroshopSetsError("Підтримуються лише Excel-файли .xlsx та .xlsm.")
    contents = await uploaded.read()
    if not contents:
        raise HoroshopSetsError("Завантажений Excel-файл порожній.")
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HoroshopSetsError("Excel-файл перевищує обмеження 20 МБ.")
    default_discount = parse_discount_percent(form.get("discount_percent", ""))
    default_price_base = parse_price_base(form.get("price_base", ""), default=PRICE_BASE_CURRENT)
    normalize_articles = parse_flag(form.get("normalize_articles"), True)
    mappings = stored_mappings()
    additional_mappings = await uploaded_mappings(form)
    if additional_mappings is not None:
        mappings = await asyncio.to_thread(add_to_stored_mappings, additional_mappings)
    return (
        contents,
        credentials_from_form(form),
        default_discount,
        default_price_base,
        normalize_articles,
        mappings,
    )


def serialise_item(item: PlanItem) -> dict[str, Any]:
    discount_str = format_discount(item.discount_percent)
    return {
        "article": item.article,
        "display_articles": list(item.display_articles),
        "products": list(item.products),
        "discounted_price": str(item.discounted_price) if item.discounted_price is not None else None,
        "row_number": item.row_number,
        "action": item.action,
        "title": item.title,
        "enabled": item.enabled,
        "sort_order": item.sort_order,
        "discount_percent": discount_str or None,
        "price_base": item.price_base if discount_str else "",
        "base_price_sum": str(item.base_price_sum) if item.base_price_sum is not None else None,
        "currency": item.currency,
        "status": "ready" if item.ready else "error",
        "message": item.error,
    }


def create_plan(
    contents: bytes,
    credentials: Credentials,
    default_discount_percent: Decimal | None = None,
    default_price_base: str = PRICE_BASE_CURRENT,
    normalize_articles: bool = True,
    mappings: ArticleMappings | None = None,
) -> tuple[list[PlanItem], HoroshopClient]:
    runtime_settings, _ = get_runtime()
    rows = parse_excel_sets(
        contents,
        default_discount_percent=default_discount_percent,
        default_price_base=default_price_base,
    )
    client = HoroshopClient(runtime_settings, credentials)
    catalog = (
        CatalogIndex.from_raw(client.export_catalog())
        if any(row.action != "delete" for row in rows)
        else CatalogIndex([])
    )
    active_mappings = mappings if mappings is not None else stored_mappings()
    return (
        prepare_plan(
            rows,
            catalog,
            normalize_articles=normalize_articles,
            mappings=active_mappings,
        ),
        client,
    )


def preview_excel(
    contents: bytes,
    credentials: Credentials,
    default_discount_percent: Decimal | None = None,
    default_price_base: str = PRICE_BASE_CURRENT,
    normalize_articles: bool = True,
    mappings: ArticleMappings | None = None,
) -> dict[str, Any]:
    plan, _ = create_plan(
        contents,
        credentials,
        default_discount_percent=default_discount_percent,
        default_price_base=default_price_base,
        normalize_articles=normalize_articles,
        mappings=mappings,
    )
    return {
        "items": [serialise_item(item) for item in plan],
        "ready": sum(item.ready for item in plan),
        "errors": sum(not item.ready for item in plan),
    }


def import_excel(
    contents: bytes,
    credentials: Credentials,
    default_discount_percent: Decimal | None = None,
    default_price_base: str = PRICE_BASE_CURRENT,
    normalize_articles: bool = True,
    mappings: ArticleMappings | None = None,
) -> dict[str, Any]:
    runtime_settings, store = get_runtime()
    plan, client = create_plan(
        contents,
        credentials,
        default_discount_percent=default_discount_percent,
        default_price_base=default_price_base,
        normalize_articles=normalize_articles,
        mappings=mappings,
    )
    upsert_items = [item for item in plan if item.ready and item.action == "upsert"]
    register_items = [item for item in plan if item.ready and item.action == "register"]
    delete_items = [item for item in plan if item.ready and item.action == "delete"]
    results: dict[str, tuple[bool, str]] = {}

    for start in range(0, len(upsert_items), runtime_settings.batch_size):
        batch = upsert_items[start : start + runtime_settings.batch_size]
        response = client.import_sets(import_payload(batch, runtime_settings))
        results.update(import_results(response))

    for item in register_items:
        results[item.article] = (True, "Набор принят на учет без изменения в Хорошоп.")

    known_delete_items = [item for item in delete_items if store.contains(item.article)]
    for item in delete_items:
        if not store.contains(item.article):
            results[item.article] = (False, "Видалення заборонено: набору немає в локальному реєстрі.")
    for start in range(0, len(known_delete_items), runtime_settings.batch_size):
        batch = known_delete_items[start : start + runtime_settings.batch_size]
        results.update(remove_results(client.remove_sets([item.article for item in batch])))

    with state_lock:
        for item in plan:
            if item.error:
                store.record_failed_attempt(item, item.error)
                continue
            success, message = results.get(
                item.article,
                (False, "API не повернуло результат для цього набору."),
            )
            if item.action == "delete":
                store.record_deletion(item.article, "deleted" if success else "delete_error", message)
            elif item.action == "register":
                store.record(item, status="registered" if success else "error", message=message, source="registry")
            else:
                if success or store.contains(item.article):
                    store.record(item, status="synced" if success else "error", message=message)
                else:
                    store.record_failed_attempt(item, message)
        store.save()

    response_items = []
    for item in plan:
        response_item = serialise_item(item)
        if item.error:
            response_item["status"] = "invalid"
        else:
            success, message = results.get(
                item.article,
                (False, "API не повернуло результат для цього набору."),
            )
            response_item["status"] = (
                "deleted" if success and item.action == "delete"
                else "registered" if success and item.action == "register"
                else "synced" if success else "error"
            )
            response_item["message"] = message
        response_items.append(response_item)

    return {
        "items": response_items,
        "imported": sum(results.get(item.article, (False, ""))[0] for item in plan if item.ready),
        "errors": sum(not item.ready for item in plan)
        + sum(not results.get(item.article, (False, ""))[0] for item in plan if item.ready),
    }


def credentials_from_json(data: dict[str, Any]) -> Credentials:
    login = normalize(data.get("login"))
    password = normalize(data.get("password"))
    token = normalize(data.get("token"))
    if not token and (not login or not password):
        raise HoroshopSetsError("Вкажіть логін і пароль API або чинний токен.")
    return Credentials(login=login, password=password, token=token)


def editor_plan(data: dict[str, Any], credentials: Credentials) -> tuple[list[PlanItem], HoroshopClient]:
    runtime_settings, store = get_runtime()
    article = normalize(data.get("article"))
    with state_lock:
        existing = dict(store.get(article) or {})
    if not existing:
        raise HoroshopSetsError("Редагувати можна лише набір з локального реєстру.")

    existing_display = tuple(
        normalize(value) for value in existing.get("display_articles", []) if normalize(value)
    )
    supplied_display_articles = data.get("display_articles")
    if isinstance(supplied_display_articles, list):
        display_articles = tuple(normalize(value) for value in supplied_display_articles if normalize(value))
    elif supplied_display_articles is None:
        display_articles = existing_display
    else:
        display_articles = split_display_articles(supplied_display_articles)
    if len(display_articles) < 2:
        raise HoroshopSetsError("У наборі має бути щонайменше два товари.")

    raw_price = normalize(data.get("discounted_price")) if "discounted_price" in data else normalize(existing.get("discounted_price"))
    has_discount_arg = "discount_percent" in data and normalize(data.get("discount_percent")) != ""
    discount_percent = (
        parse_discount_percent(data.get("discount_percent"))
        if "discount_percent" in data
        else parse_discount_percent(existing.get("discount_percent"))
    )
    price_base = parse_price_base(
        data.get("price_base") if "price_base" in data else existing.get("price_base"),
        default=PRICE_BASE_CURRENT,
    )

    if raw_price:
        price: Decimal | None = parse_price(raw_price)
        if not has_discount_arg:
            discount_percent = None
    elif discount_percent is not None:
        price = None
    elif "discount_percent" in data or "price_base" in data:
        price = None
        discount_percent = Decimal("0.00")
    else:
        raise ValueError("Вкажіть кінцеву ціну або знижку у відсотках.")

    title = SET_TITLE
    enabled = bool(existing.get("enabled", True))
    currency = normalize(existing.get("currency")) or runtime_settings.currency
    normalize_articles = parse_flag(data.get("normalize_articles"), True)
    composition_changed = tuple(a.casefold() for a in display_articles) != tuple(
        a.casefold() for a in existing_display
    )

    client = HoroshopClient(runtime_settings, credentials)
    catalog = CatalogIndex.from_raw(client.export_catalog())
    row = SetRow(
        article,
        display_articles,
        price,
        row_number=1,
        title=title,
        enabled=enabled,
        sort_order=None,
        discount_percent=discount_percent,
        currency=currency,
        price_base=price_base,
    )
    active_mappings = stored_mappings() if composition_changed else ArticleMappings.empty()
    items = prepare_plan(
        [row],
        catalog,
        normalize_articles=normalize_articles,
        mappings=active_mappings,
    )
    if not composition_changed and len(items) == 1 and items[0].error:
        # Fallback if a previously stored display_article needs mapping resolution
        fallback_items = prepare_plan(
            [row],
            catalog,
            normalize_articles=normalize_articles,
            mappings=stored_mappings(),
        )
        matching = [it for it in fallback_items if it.article.casefold() == article.casefold()]
        items = matching[:1] if matching else fallback_items[:1]
    if len(items) == 1 and not items[0].error and items[0].article != article and not composition_changed:
        items = [replace(items[0], article=article)]
    for item in items:
        if item.error:
            raise HoroshopSetsError(item.error)
    return items, client


def update_tracked_set(data: dict[str, Any]) -> dict[str, Any]:
    credentials = credentials_from_json(data)
    runtime_settings, store = get_runtime()
    items, client = editor_plan(data, credentials)
    results: dict[str, tuple[bool, str]] = {}
    for start in range(0, len(items), runtime_settings.batch_size):
        batch = items[start : start + runtime_settings.batch_size]
        results.update(import_results(client.import_sets(import_payload(batch, runtime_settings))))
    with state_lock:
        for item in items:
            res = results.get(item.article, (False, "API не повернуло результат для цього набору."))
            store.record(item, "synced" if res[0] else "error", res[1])
        store.save()
    first = items[0]
    first_res = results.get(first.article, (False, "API не повернуло результат для цього набору."))
    all_ok = all(results.get(it.article, (False, ""))[0] for it in items)
    return {
        "item": serialise_item(first),
        "items": [
            {
                **serialise_item(it),
                "status": "synced" if results.get(it.article, (False, ""))[0] else "error",
                "message": results.get(it.article, (False, "API не повернуло результат для цього набору."))[1],
            }
            for it in items
        ],
        "imported": sum(results.get(it.article, (False, ""))[0] for it in items),
        "errors": sum(not results.get(it.article, (False, ""))[0] for it in items),
        "success": all_ok,
        "message": (
            first_res[1]
            if len(items) == 1
            else f"Оновлено / створено варіацій наборів: {sum(results.get(it.article, (False, ''))[0] for it in items)} із {len(items)}."
        ),
    }


def create_manual_set(data: dict[str, Any]) -> dict[str, Any]:
    credentials = credentials_from_json(data)
    runtime_settings, store = get_runtime()
    raw_article = normalize(data.get("article"))
    display_articles = split_display_articles(data.get("display_articles", ""))
    article = raw_article or build_set_article(display_articles)

    raw_price = normalize(data.get("discounted_price"))
    discount_percent = parse_discount_percent(data.get("discount_percent"))
    price_base = parse_price_base(data.get("price_base"), default=PRICE_BASE_CURRENT)
    normalize_articles = parse_flag(data.get("normalize_articles"), True)
    if raw_price:
        price: Decimal | None = parse_price(raw_price)
        discount_percent = None
    else:
        price = None
        if discount_percent is None:
            discount_percent = Decimal("0.00")

    row = SetRow(
        article,
        display_articles,
        price,
        row_number=1,
        discount_percent=discount_percent,
        price_base=price_base,
    )
    client = HoroshopClient(runtime_settings, credentials)
    items = prepare_plan(
        [row],
        CatalogIndex.from_raw(client.export_catalog()),
        normalize_articles=normalize_articles,
        mappings=stored_mappings(),
    )
    for item in items:
        if item.error:
            raise HoroshopSetsError(item.error)

    with state_lock:
        if len(items) == 1 and store.contains(items[0].article):
            raise HoroshopSetsError("Цей набір уже є в реєстрі. Скористайтеся редагуванням у таблиці нижче.")

    results: dict[str, tuple[bool, str]] = {}
    for start in range(0, len(items), runtime_settings.batch_size):
        batch = items[start : start + runtime_settings.batch_size]
        results.update(import_results(client.import_sets(import_payload(batch, runtime_settings))))

    with state_lock:
        for item in items:
            res = results.get(item.article, (False, "API не повернуло результат для цього набору."))
            if res[0] or store.contains(item.article):
                store.record(item, "synced" if res[0] else "error", res[1])
            else:
                store.record_failed_attempt(item, res[1])
        store.save()

    first = items[0]
    first_res = results.get(first.article, (False, "API не повернуло результат для цього набору."))
    imported_count = sum(results.get(it.article, (False, ""))[0] for it in items)
    error_count = len(items) - imported_count
    return {
        "item": serialise_item(first),
        "items": [
            {
                **serialise_item(it),
                "status": "synced" if results.get(it.article, (False, ""))[0] else "error",
                "message": results.get(it.article, (False, "API не повернуло результат для цього набору."))[1],
            }
            for it in items
        ],
        "imported": imported_count,
        "errors": error_count,
        "success": error_count == 0,
        "message": (
            first_res[1]
            if len(items) == 1
            else f"Створено варіацій наборів: {imported_count} із {len(items)}."
        ),
    }


def delete_tracked_sets(data: dict[str, Any]) -> dict[str, Any]:
    credentials = credentials_from_json(data)
    raw_articles = data.get("articles")
    if not isinstance(raw_articles, list):
        raise HoroshopSetsError("Передайте список артикулів наборів для видалення.")
    articles = list(dict.fromkeys(normalize(article) for article in raw_articles if normalize(article)))
    if not articles:
        raise HoroshopSetsError("Виберіть хоча б один набір.")
    runtime_settings, store = get_runtime()
    with state_lock:
        known = [article for article in articles if store.contains(article)]
    results: dict[str, tuple[bool, str]] = {
        article: (False, "Видалення заборонено: набору немає в локальному реєстрі.")
        for article in articles if article not in known
    }
    client = HoroshopClient(runtime_settings, credentials)
    for start in range(0, len(known), runtime_settings.batch_size):
        response = client.remove_sets(known[start : start + runtime_settings.batch_size])
        results.update(remove_results(response))
    with state_lock:
        for article in articles:
            success, message = results.get(article, (False, "API не повернуло результат для цього набору."))
            if store.contains(article):
                store.record_deletion(article, "deleted" if success else "delete_error", message)
        store.save()
    return {
        "items": [
            {"article": article, "success": results.get(article, (False, ""))[0], "message": results.get(article, (False, ""))[1]}
            for article in articles
        ]
    }


def page_html() -> str:
    return (PROJECT_DIR / "web_ui.html").read_text(encoding="utf-8")


app = FastAPI(title="Набори Хорошоп")


@app.exception_handler(Exception)
async def unexpected_error(_: Request, error: Exception) -> JSONResponse:
    logger.exception("Unhandled web request error", exc_info=error)
    return JSONResponse(
        status_code=500,
        content={"detail": f"Внутрішня помилка сервера: {error}"},
    )


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse(
        content=page_html(),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


@app.get("/api/template")
def download_template() -> Response:
    return Response(
        content=build_excel_template(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="horoshop_sets_template_v2.xlsx"',
            "Cache-Control": "no-store",
        },
    )


@app.get("/api/state")
def get_state() -> dict[str, Any]:
    _, store = get_runtime()
    with state_lock:
        snapshot = store.snapshot()
    return {"sets": snapshot, "count": len(snapshot)}


@app.get("/api/sets/export")
def export_sets() -> Response:
    _, store = get_runtime()
    with state_lock:
        contents = build_state_excel(store.snapshot())
    return Response(
        content=contents,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="horoshop_sets_registry.xlsx"',
            "Cache-Control": "no-store",
        },
    )


@app.post("/api/sets/update")
async def update_set(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        if not isinstance(data, dict):
            raise HoroshopSetsError("Дані редагування мають бути об'єктом JSON.")
        return await asyncio.to_thread(update_tracked_set, data)
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/sets/create")
async def create_set(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        if not isinstance(data, dict):
            raise HoroshopSetsError("Дані набору мають бути об'єктом JSON.")
        return await asyncio.to_thread(create_manual_set, data)
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/sets/delete")
async def delete_sets(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        if not isinstance(data, dict):
            raise HoroshopSetsError("Дані видалення мають бути об'єктом JSON.")
        return await asyncio.to_thread(delete_tracked_sets, data)
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.get("/api/mappings/info")
def mappings_info() -> dict[str, int]:
    mappings = stored_mappings()
    return {
        "sources": len(mappings.entries),
        "rules": sum(len(targets) for targets in mappings.entries.values()),
    }


@app.get("/api/mappings/export")
def export_mappings() -> Response:
    content = build_mappings_excel(stored_mappings())
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="horoshop_article_mappings.xlsx"',
            "Cache-Control": "no-store",
        },
    )


@app.post("/api/mappings/clear")
async def clear_mappings() -> dict[str, Any]:
    with mapping_lock:
        backup_name = ""
        if MAPPINGS_FILE.exists():
            import datetime

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = MAPPINGS_FILE.with_name(f"article_mappings_{timestamp}.json.bak")
            latest_bak = MAPPINGS_FILE.with_suffix(".json.bak")
            try:
                content = MAPPINGS_FILE.read_text(encoding="utf-8")
                backup_file.write_text(content, encoding="utf-8")
                latest_bak.write_text(content, encoding="utf-8")
                backup_name = backup_file.name
            except OSError as error:
                logger.warning("Could not backup mappings file: %s", error)
        save_article_mappings(MAPPINGS_FILE, ArticleMappings.empty())
        return {"status": "cleared", "backup": backup_name}


@app.post("/api/mappings")
async def upload_mappings_endpoint(request: Request) -> dict[str, int]:
    try:
        mappings = await uploaded_mappings(await request.form())
        if mappings is None:
            raise HoroshopSetsError("Оберіть Excel-файл сопоставлень.")
        stored = await asyncio.to_thread(add_to_stored_mappings, mappings)
        return {
            "sources": len(stored.entries),
            "rules": sum(len(targets) for targets in stored.entries.values()),
        }
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/preview")
async def preview(request: Request) -> dict[str, Any]:
    try:
        (
            contents,
            credentials,
            default_discount,
            default_price_base,
            normalize_articles,
            mappings,
        ) = await upload_bytes(request)
        return await asyncio.to_thread(
            preview_excel,
            contents,
            credentials,
            default_discount,
            default_price_base,
            normalize_articles,
            mappings,
        )
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/import")
async def import_sets(request: Request) -> dict[str, Any]:
    try:
        (
            contents,
            credentials,
            default_discount,
            default_price_base,
            normalize_articles,
            mappings,
        ) = await upload_bytes(request)
        return await asyncio.to_thread(
            import_excel,
            contents,
            credentials,
            default_discount,
            default_price_base,
            normalize_articles,
            mappings,
        )
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


def run_server() -> None:
    import uvicorn

    runtime_settings, _ = get_runtime()
    configure_service_output(runtime_settings)
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)


if __name__ == "__main__":
    run_server()
