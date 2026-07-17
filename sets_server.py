from __future__ import annotations

import asyncio
import logging
import shutil
import threading
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from horoshop_sets import (
    CatalogIndex,
    Credentials,
    HoroshopClient,
    HoroshopSetsError,
    PlanItem,
    Settings,
    StateStore,
    build_excel_template,
    import_payload,
    import_results,
    load_settings,
    parse_excel_sets,
    prepare_plan,
)


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024

settings: Settings | None = None
state_store: StateStore | None = None
state_lock = threading.Lock()
logger = logging.getLogger(__name__)


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


def credentials_from_form(form: Any) -> Credentials:
    login = str(form.get("login", "")).strip()
    password = str(form.get("password", "")).strip()
    token = str(form.get("token", "")).strip()
    if not token and (not login or not password):
        raise HoroshopSetsError("Вкажіть логін і пароль API або чинний токен.")
    return Credentials(login=login, password=password, token=token)


async def upload_bytes(request: Request) -> tuple[bytes, Credentials]:
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
    return contents, credentials_from_form(form)


def serialise_item(item: PlanItem) -> dict[str, Any]:
    return {
        "article": item.article,
        "display_articles": list(item.display_articles),
        "products": list(item.products),
        "discounted_price": str(item.discounted_price) if item.discounted_price is not None else None,
        "row_number": item.row_number,
        "status": "ready" if item.ready else "error",
        "message": item.error,
    }


def create_plan(contents: bytes, credentials: Credentials) -> tuple[list[PlanItem], HoroshopClient]:
    runtime_settings, _ = get_runtime()
    rows = parse_excel_sets(contents)
    client = HoroshopClient(runtime_settings, credentials)
    catalog = CatalogIndex(client.export_catalog())
    return prepare_plan(rows, catalog), client


def preview_excel(contents: bytes, credentials: Credentials) -> dict[str, Any]:
    plan, _ = create_plan(contents, credentials)
    return {
        "items": [serialise_item(item) for item in plan],
        "ready": sum(item.ready for item in plan),
        "errors": sum(not item.ready for item in plan),
    }


def import_excel(contents: bytes, credentials: Credentials) -> dict[str, Any]:
    runtime_settings, store = get_runtime()
    plan, client = create_plan(contents, credentials)
    ready_items = [item for item in plan if item.ready]
    results: dict[str, tuple[bool, str]] = {}

    for start in range(0, len(ready_items), runtime_settings.batch_size):
        batch = ready_items[start : start + runtime_settings.batch_size]
        response = client.import_sets(import_payload(batch, runtime_settings))
        results.update(import_results(response))

    with state_lock:
        for item in plan:
            if item.error:
                store.record(item, status="invalid", message=item.error)
                continue
            success, message = results.get(
                item.article,
                (False, "API не повернуло результат для цього набору."),
            )
            store.record(
                item,
                status="synced" if success else "error",
                message=message,
            )
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
            response_item["status"] = "synced" if success else "error"
            response_item["message"] = message
        response_items.append(response_item)

    return {
        "items": response_items,
        "imported": sum(results.get(item.article, (False, ""))[0] for item in ready_items),
        "errors": sum(not item.ready for item in plan)
        + sum(not results.get(item.article, (False, ""))[0] for item in ready_items),
    }


def page_html() -> str:
    return """<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Набори Хорошоп</title>
  <style>
    :root { color-scheme: light; font-family: Arial, sans-serif; color: #15251d; background: #f3f6f3; }
    * { box-sizing: border-box; }
    body { margin: 0; }
    header { background: #135c3c; color: #fff; padding: 22px max(20px, calc((100% - 1180px) / 2)); }
    h1 { font-size: 25px; margin: 0; font-weight: 700; }
    header p { margin: 6px 0 0; color: #dcefe4; }
    main { max-width: 1180px; margin: 0 auto; padding: 22px 20px 36px; }
    section { background: #fff; border: 1px solid #d6e1d9; border-radius: 6px; margin-bottom: 16px; padding: 18px; }
    h2 { margin: 0 0 14px; font-size: 18px; }
    .grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
    label { display: grid; gap: 6px; font-size: 14px; font-weight: 600; }
    label small { font-weight: 400; color: #52665a; }
    input { width: 100%; min-height: 38px; border: 1px solid #aebeb4; border-radius: 4px; padding: 8px; font: inherit; }
    .actions { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-top: 16px; }
    button, .download { border: 0; border-radius: 4px; padding: 10px 14px; font: inherit; font-weight: 700; cursor: pointer; text-decoration: none; }
    button { background: #135c3c; color: #fff; }
    button.secondary, .download { background: #e8f0eb; color: #153d29; border: 1px solid #b7cabb; }
    button:disabled { opacity: .6; cursor: wait; }
    .hint { margin: 14px 0 0; color: #405548; line-height: 1.45; }
    .message { min-height: 22px; margin: 12px 0 0; font-weight: 600; }
    .message.error { color: #a51d2a; }
    .message.ok { color: #14633d; }
    .table-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 720px; font-size: 14px; }
    th, td { padding: 10px; text-align: left; vertical-align: top; border-bottom: 1px solid #e0e8e2; }
    th { color: #496354; font-size: 12px; text-transform: uppercase; }
    .tag { display: inline-block; border-radius: 3px; padding: 3px 6px; font-size: 12px; font-weight: 700; }
    .tag.ready, .tag.synced { color: #0c5532; background: #dff3e7; }
    .tag.error, .tag.invalid { color: #8f1825; background: #fae4e6; }
    .empty { color: #6c7b71; padding: 8px 0; }
    @media (max-width: 720px) { .grid { grid-template-columns: 1fr; } main { padding: 14px; } section { padding: 14px; } }
  </style>
</head>
<body>
  <header><h1>Набори Хорошоп</h1><p>Створення та оновлення комплектів товарів з Excel.</p></header>
  <main>
    <section>
      <h2>Доступ до API</h2>
      <div class="grid">
        <label>Логін API<input id="login" autocomplete="username"></label>
        <label>Пароль API<input id="password" type="password" autocomplete="current-password"></label>
        <label>Токен <small>Необов'язково, діє 10 хвилин.</small><input id="token"></label>
      </div>
    </section>
    <section>
      <h2>Excel з наборами</h2>
      <div class="grid">
        <label>Файл Excel<input id="file" type="file" accept=".xlsx,.xlsm"></label>
        <div class="actions"><a class="download" href="/api/template">Завантажити шаблон Excel</a></div>
      </div>
      <p class="hint">Перший стовпець - артикул набору. Другий - артикули товарів через <b>;</b>, саме у вигляді «Артикул відображення на сайті». Третій - ціна набору. Артикули вводьте як текст, щоб зберегти початкові нулі. Перед відправленням кожен товар звіряється з каталогом Хорошопа.</p>
      <div class="actions"><button class="secondary" id="preview">Перевірити Excel</button><button id="submit">Створити / оновити набори</button></div>
      <div class="message" id="message"></div>
    </section>
    <section>
      <h2>Результат перевірки</h2><div class="table-wrap" id="result"><p class="empty">Excel ще не перевірявся.</p></div>
    </section>
    <section>
      <h2>Збережений стан наборів</h2><div class="table-wrap" id="state"><p class="empty">Завантаження стану...</p></div>
    </section>
  </main>
  <script>
    const message = document.getElementById('message');
    const result = document.getElementById('result');
    const state = document.getElementById('state');
    const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    const formData = () => { const data = new FormData(); const file = document.getElementById('file').files[0]; if (file) data.append('file', file); data.append('login', document.getElementById('login').value); data.append('password', document.getElementById('password').value); data.append('token', document.getElementById('token').value); return data; };
    const setMessage = (text, kind='') => { message.textContent = text; message.className = 'message ' + kind; };
    const renderItems = items => {
      if (!items.length) { result.innerHTML = '<p class="empty">У файлі немає наборів.</p>'; return; }
      result.innerHTML = `<table><thead><tr><th>Рядок</th><th>Набір</th><th>Артикули з Excel</th><th>Знайдені товари</th><th>Ціна</th><th>Стан</th></tr></thead><tbody>${items.map(item => `<tr><td>${item.row_number}</td><td>${esc(item.article)}</td><td>${item.display_articles.map(esc).join('; ')}</td><td>${item.products.map(esc).join('; ') || '-'}</td><td>${esc(item.discounted_price || '-')}</td><td><span class="tag ${item.status}">${item.status === 'ready' ? 'готово' : item.status === 'synced' ? 'імпортовано' : 'помилка'}</span>${item.message ? '<br>' + esc(item.message) : ''}</td></tr>`).join('')}</tbody></table>`;
    };
    const renderState = data => {
      const items = data.sets || [];
      if (!items.length) { state.innerHTML = '<p class="empty">Набори ще не імпортувалися.</p>'; return; }
      state.innerHTML = `<table><thead><tr><th>Набір</th><th>Товари</th><th>Ціна</th><th>Стан</th><th>Оновлено</th><th>Повідомлення</th></tr></thead><tbody>${items.map(item => `<tr><td>${esc(item.article)}</td><td>${(item.products || []).map(esc).join('; ') || '-'}</td><td>${esc(item.discounted_price || '-')}</td><td><span class="tag ${esc(item.status)}">${esc(item.status)}</span></td><td>${esc(item.updated_at)}</td><td>${esc(item.message)}</td></tr>`).join('')}</tbody></table>`;
    };
    async function call(endpoint, importing) {
      const file = document.getElementById('file').files[0];
      if (!file) { setMessage('Оберіть Excel-файл.', 'error'); return; }
      const buttons = [document.getElementById('preview'), document.getElementById('submit')]; buttons.forEach(button => button.disabled = true);
      setMessage(importing ? 'Імпорт наборів...' : 'Перевірка Excel...');
      try {
        const response = await fetch(endpoint, {method: 'POST', body: formData()});
        const raw = await response.text();
        let data;
        try { data = JSON.parse(raw); }
        catch { data = {detail: raw || 'Сервер повернув порожню відповідь.'}; }
        if (!response.ok) throw new Error(data.detail || 'Помилка запиту.');
        renderItems(data.items || []);
        setMessage(importing ? `Готово: імпортовано ${data.imported}, помилок ${data.errors}.` : `Готово до імпорту: ${data.ready}, помилок ${data.errors}.`, importing && data.errors ? 'error' : 'ok');
        if (importing) await loadState();
      } catch (error) { setMessage(error.message, 'error'); }
      finally { buttons.forEach(button => button.disabled = false); }
    }
    async function loadState() { try { const response = await fetch('/api/state'); renderState(await response.json()); } catch { state.innerHTML = '<p class="empty">Не вдалося завантажити збережений стан.</p>'; } }
    document.getElementById('preview').addEventListener('click', () => call('/api/preview', false));
    document.getElementById('submit').addEventListener('click', () => call('/api/import', true));
    loadState();
  </script>
</body></html>"""


app = FastAPI(title="Набори Хорошоп")


@app.exception_handler(Exception)
async def unexpected_error(_: Request, error: Exception) -> JSONResponse:
    logger.exception("Unhandled web request error", exc_info=error)
    return JSONResponse(
        status_code=500,
        content={"detail": "Внутрішня помилка сервера. Перевірте logs/server-error.log."},
    )


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return page_html()


@app.get("/api/template")
def download_template() -> Response:
    return Response(
        content=build_excel_template(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="horoshop_sets_template.xlsx"'},
    )


@app.get("/api/state")
def get_state() -> dict[str, Any]:
    _, store = get_runtime()
    with state_lock:
        snapshot = store.snapshot()
    return {"sets": snapshot, "count": len(snapshot)}


@app.post("/api/preview")
async def preview(request: Request) -> dict[str, Any]:
    try:
        contents, credentials = await upload_bytes(request)
        return await asyncio.to_thread(preview_excel, contents, credentials)
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/import")
async def import_sets(request: Request) -> dict[str, Any]:
    try:
        contents, credentials = await upload_bytes(request)
        return await asyncio.to_thread(import_excel, contents, credentials)
    except (HoroshopSetsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


if __name__ == "__main__":
    import uvicorn

    runtime_settings, _ = get_runtime()
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)
