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
    CatalogIndex,
    Credentials,
    HoroshopClient,
    HoroshopSetsError,
    PlanItem,
    SetRow,
    SET_TITLE,
    Settings,
    StateStore,
    build_excel_template,
    build_set_article,
    build_state_excel,
    import_payload,
    import_results,
    load_settings,
    parse_excel_sets,
    parse_price,
    prepare_plan,
    remove_results,
    split_display_articles,
    normalize,
)


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024

settings: Settings | None = None
state_store: StateStore | None = None
state_lock = threading.Lock()
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
        "action": item.action,
        "title": item.title,
        "enabled": item.enabled,
        "sort_order": item.sort_order,
        "discount_percent": item.discount_percent,
        "currency": item.currency,
        "status": "ready" if item.ready else "error",
        "message": item.error,
    }


def create_plan(contents: bytes, credentials: Credentials) -> tuple[list[PlanItem], HoroshopClient]:
    runtime_settings, _ = get_runtime()
    rows = parse_excel_sets(contents)
    client = HoroshopClient(runtime_settings, credentials)
    catalog = (
        CatalogIndex.from_raw(client.export_catalog())
        if any(row.action != "delete" for row in rows)
        else CatalogIndex([])
    )
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


def editor_plan(data: dict[str, Any], credentials: Credentials) -> tuple[PlanItem, HoroshopClient]:
    runtime_settings, store = get_runtime()
    article = normalize(data.get("article"))
    with state_lock:
        existing = dict(store.get(article) or {})
    if not existing:
        raise HoroshopSetsError("Редагувати можна лише набір з локального реєстру.")

    supplied_display_articles = data.get("display_articles")
    if isinstance(supplied_display_articles, list):
        display_articles = tuple(normalize(value) for value in supplied_display_articles if normalize(value))
    elif supplied_display_articles is None:
        display_articles = tuple(existing.get("display_articles", []))
    else:
        display_articles = split_display_articles(supplied_display_articles)
    if len(display_articles) < 2:
        raise HoroshopSetsError("У наборі має бути щонайменше два товари.")

    price = parse_price(data.get("discounted_price", existing.get("discounted_price", "")))
    title = SET_TITLE
    enabled = bool(existing.get("enabled", True))
    currency = normalize(existing.get("currency")) or runtime_settings.currency

    client = HoroshopClient(runtime_settings, credentials)
    row = SetRow(
        article,
        display_articles,
        price,
        row_number=1,
        title=title,
        enabled=enabled,
        sort_order=None,
        discount_percent=None,
        currency=currency,
    )
    item = prepare_plan([row], CatalogIndex.from_raw(client.export_catalog()))[0]
    if item.error:
        raise HoroshopSetsError(item.error)
    return item, client


def update_tracked_set(data: dict[str, Any]) -> dict[str, Any]:
    credentials = credentials_from_json(data)
    runtime_settings, store = get_runtime()
    item, client = editor_plan(data, credentials)
    result = import_results(client.import_sets(import_payload([item], runtime_settings))).get(
        item.article,
        (False, "API не повернуло результат для цього набору."),
    )
    with state_lock:
        store.record(item, "synced" if result[0] else "error", result[1])
        store.save()
    return {"item": serialise_item(item), "success": result[0], "message": result[1]}


def create_manual_set(data: dict[str, Any]) -> dict[str, Any]:
    credentials = credentials_from_json(data)
    runtime_settings, store = get_runtime()
    article = normalize(data.get("article"))
    display_articles = split_display_articles(data.get("display_articles", ""))
    article = article or build_set_article(display_articles)
    with state_lock:
        if store.contains(article):
            raise HoroshopSetsError("Цей набір уже є в реєстрі. Скористайтеся редагуванням у таблиці нижче.")

    price = parse_price(data.get("discounted_price", ""))
    row = SetRow(article, display_articles, price, row_number=1)
    client = HoroshopClient(runtime_settings, credentials)
    item = prepare_plan([row], CatalogIndex.from_raw(client.export_catalog()))[0]
    if item.error:
        raise HoroshopSetsError(item.error)

    result = import_results(client.import_sets(import_payload([item], runtime_settings))).get(
        item.article,
        (False, "API не повернуло результат для цього набору."),
    )
    with state_lock:
        if result[0]:
            store.record(item, "synced", result[1])
        else:
            store.record_failed_attempt(item, result[1])
        store.save()
    return {"item": serialise_item(item), "success": result[0], "message": result[1]}


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

    return """<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Наборы Хорошоп</title>
  <style>
    :root { font-family: Arial, sans-serif; color:#183126; background:#f3f6f4; }
    * { box-sizing:border-box; } body { margin:0; } header { background:#135c3c; color:#fff; padding:22px max(20px, calc((100% - 1180px)/2)); }
    h1 { margin:0; font-size:26px; } header p { margin:6px 0 0; color:#dcefe4; } main { max-width:1180px; margin:auto; padding:20px; }
    section { background:#fff; border:1px solid #d5e1d9; border-radius:6px; padding:18px; margin-bottom:16px; } h2 { margin:0 0 14px; font-size:18px; }
    .grid { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; } .grid.two { grid-template-columns:repeat(2,minmax(0,1fr)); }
    label { display:grid; gap:6px; font-size:13px; font-weight:700; } input, select { min-height:38px; width:100%; border:1px solid #adc1b3; border-radius:4px; padding:8px; font:inherit; } input.input-error { border-color:#b42332; background:#fff1f2; outline:2px solid #f8c6cc; }
    .actions { display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-top:14px; } button,.link-button { border:1px solid #135c3c; border-radius:4px; padding:9px 13px; font:inherit; font-weight:700; background:#135c3c; color:#fff; cursor:pointer; text-decoration:none; }
    button.secondary,.link-button.secondary { background:#f1f6f2; color:#15482e; border-color:#b8cabb; } button.danger { background:#a52834; border-color:#a52834; } button:disabled { opacity:.55; cursor:wait; }
    .note { margin:12px 0 0; padding:10px 12px; border-left:4px solid #d09120; background:#fff8e7; color:#5b481a; line-height:1.45; } .hint { color:#4f6456; line-height:1.45; }
    .message { min-height:20px; margin-top:12px; font-weight:700; } .message.error { color:#a11d2a; } .message.ok { color:#14623d; }
    .table-wrap { overflow-x:auto; } table { min-width:980px; width:100%; border-collapse:collapse; font-size:13px; } th,td { padding:9px; text-align:left; vertical-align:top; border-bottom:1px solid #e0e9e3; } th { color:#4b6355; font-size:11px; text-transform:uppercase; } .price-cell { display:flex; min-width:135px; gap:5px; } .price-cell input { min-height:32px; min-width:75px; padding:5px; } .price-cell button { padding:5px 8px; }
    .tag { display:inline-block; border-radius:3px; padding:3px 6px; font-size:11px; font-weight:700; background:#e8eef0; color:#354b55; } .tag.synced,.tag.registered { background:#dff3e7; color:#0c5532; } .tag.error,.tag.delete_error { background:#fae4e6; color:#8d1826; } .empty { color:#6b7b70; }
    #editor { display:none; } #editor.visible { display:block; } @media(max-width:720px) { main { padding:14px; } section { padding:14px; } .grid,.grid.two { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header><h1>Наборы Хорошоп</h1><p>Создание, изменение и удаление наборов товаров.</p></header>
  <main>
    <section>
      <h2>Доступ к API</h2>
      <div class="grid"><label>Логин API<input id="login" autocomplete="username"></label><label>Пароль API<input id="password" type="password" autocomplete="current-password"></label><label>Токен API<input id="token"><span class="hint">Можно указать вместо логина и пароля.</span></label></div>
    </section>
    <section>
      <h2>Массовое управление Excel</h2>
      <div class="grid"><label>Excel-файл наборов<input id="file" type="file" accept=".xlsx,.xlsm"></label><div class="actions"><a class="link-button secondary" href="/api/template">Скачать шаблон</a><a class="link-button secondary" href="/api/sets/export">Выгрузить реестр</a></div></div>
      <p class="hint">Первые три колонки: артикул набора, артикула отображения товаров через <b>;</b>, цена. Дополнительная колонка «Действие»: <b>обновить</b>, <b>удалить</b> или <b>принять на учет</b>. Для удаления достаточно артикула и действия.</p>
      <div class="actions"><button class="secondary" id="preview" type="button">Проверить Excel</button><button id="import" type="button">Выполнить операции</button></div><div class="message" id="message"></div>
    </section>
    <section>
      <h2>Результат файла</h2><div class="table-wrap" id="result"><p class="empty">Файл еще не проверялся.</p></div>
    </section>
    <section>
      <h2>Реестр наборов</h2>
      <p class="note">Здесь показаны только наборы, созданные через сервис или принятые на учет. Хорошоп не дает массово выгрузить существующие наборы, поэтому неизвестные сервису наборы не удаляются автоматически.</p>
      <div class="actions"><button class="danger" id="delete-selected" type="button">Удалить выбранные</button><a class="link-button secondary" href="/api/sets/export">Скачать таблицу реестра</a></div>
      <div class="table-wrap" id="state"><p class="empty">Загрузка реестра...</p></div>
    </section>
    <section id="editor">
      <h2 id="editor-title">Редактирование набора</h2>
      <div class="grid two"><label>Артикул набора<input id="edit-article" readonly></label><label>Название<input id="edit-title"></label><label>Артикулы отображения товаров<input id="edit-products"><span class="hint">Через ;</span></label><label>Цена набора<input id="edit-price" inputmode="decimal"></label><label>Валюта<input id="edit-currency" maxlength="3"></label><label>Порядок сортировки<input id="edit-sort" inputmode="numeric"></label><label>Скидка %<input id="edit-discount" inputmode="numeric"></label><label>Активен<select id="edit-enabled"><option value="true">Да</option><option value="false">Нет</option></select></label></div>
      <div class="actions"><button id="save-editor" type="button">Сохранить изменения</button><button class="secondary" id="cancel-editor" type="button">Закрыть</button></div>
    </section>
  </main>
  <script>
    const message=document.getElementById('message'), result=document.getElementById('result'), state=document.getElementById('state'); let registry=[];
    const esc=value=>String(value??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    const credentials=()=>({login:document.getElementById('login').value,password:document.getElementById('password').value,token:document.getElementById('token').value});
    function clearCredentialErrors(){['login','password','token'].forEach(id=>document.getElementById(id).classList.remove('input-error'))}
    function validateCredentials(){const data=credentials();clearCredentialErrors();if(data.token.trim())return true;let valid=true;if(!data.login.trim()){document.getElementById('login').classList.add('input-error');valid=false}if(!data.password.trim()){document.getElementById('password').classList.add('input-error');valid=false}if(!valid)setMessage('Укажите логин и пароль API или токен.','error');return valid}
    function handleApiError(error){if(/логин|парол|токен/i.test(error.message))validateCredentials();setMessage(error.message,'error')}
    const setMessage=(text,kind='')=>{message.textContent=text;message.className='message '+kind};
    async function readResponse(response){const raw=await response.text();try{return JSON.parse(raw)}catch{return {detail:raw||'Сервер вернул пустой ответ.'}}}
    async function jsonApi(url,body){const response=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});const data=await readResponse(response);if(!response.ok)throw new Error(data.detail||'Ошибка сервера.');return data}
    function renderResult(items){if(!items.length){result.innerHTML='<p class="empty">В файле нет операций.</p>';return}result.innerHTML=`<table><thead><tr><th>Строка</th><th>Действие</th><th>Набор</th><th>Товары</th><th>Цена</th><th>Статус</th></tr></thead><tbody>${items.map(item=>`<tr><td>${item.row_number}</td><td>${esc(item.action)}</td><td>${esc(item.article)}</td><td>${item.display_articles.map(esc).join('; ')||'-'}</td><td>${esc(item.discounted_price||'-')}</td><td><span class="tag ${esc(item.status)}">${esc(item.status)}</span>${item.message?'<br>'+esc(item.message):''}</td></tr>`).join('')}</tbody></table>`}
    function renderState(data){registry=data.sets||[];if(!registry.length){state.innerHTML='<p class="empty">В реестре пока нет наборов.</p>';return}state.innerHTML=`<table><thead><tr><th></th><th>Артикул</th><th>Товары</th><th>Цена</th><th>Название</th><th>Активен</th><th>Статус</th><th>Изменено</th><th></th></tr></thead><tbody>${registry.map(item=>`<tr><td><input type="checkbox" data-select="${esc(item.article)}"></td><td>${esc(item.article)}</td><td>${(item.display_articles||[]).map(esc).join('; ')}</td><td><div class="price-cell"><input data-price="${esc(item.article)}" value="${esc(item.discounted_price)}"><button class="secondary" data-save-price="${esc(item.article)}" type="button">OK</button></div></td><td>${esc(item.title||'Вместе дешевле')}</td><td>${item.enabled?'Да':'Нет'}</td><td><span class="tag ${esc(item.status)}">${esc(item.status)}</span><br>${esc(item.message||'')}</td><td>${esc(item.updated_at)}</td><td><div class="actions"><button class="secondary" data-edit="${esc(item.article)}" type="button">Изменить</button><button class="danger" data-delete-one="${esc(item.article)}" type="button">Удалить</button></div></td></tr>`).join('')}</tbody></table>`;document.querySelectorAll('[data-edit]').forEach(button=>button.addEventListener('click',()=>openEditor(button.dataset.edit)));document.querySelectorAll('[data-save-price]').forEach(button=>button.addEventListener('click',()=>quickPrice(button.dataset.savePrice)));document.querySelectorAll('[data-delete-one]').forEach(button=>button.addEventListener('click',()=>deleteArticles([button.dataset.deleteOne])))}
    async function loadState(){try{const response=await fetch('/api/state');const data=await readResponse(response);if(!response.ok)throw new Error(data.detail);renderState(data)}catch(error){state.innerHTML='<p class="empty">Не удалось загрузить реестр: '+esc(error.message)+'</p>'}}
    function openEditor(article){const item=registry.find(value=>value.article===article);if(!item)return;document.getElementById('editor').classList.add('visible');document.getElementById('edit-article').value=item.article;document.getElementById('edit-title').value=item.title||'Вместе дешевле';document.getElementById('edit-products').value=(item.display_articles||[]).join('; ');document.getElementById('edit-price').value=item.discounted_price||'';document.getElementById('edit-currency').value=item.currency||'UAH';document.getElementById('edit-sort').value=item.sort_order??'';document.getElementById('edit-discount').value=item.discount_percent??'';document.getElementById('edit-enabled').value=String(item.enabled!==false);document.getElementById('editor').scrollIntoView({behavior:'smooth',block:'start'})}
    async function quickPrice(article){if(!validateCredentials())return;const item=registry.find(value=>value.article===article);const price=document.querySelector(`[data-price="${CSS.escape(article)}"]`).value;try{setMessage('Сохранение цены...');const data=await jsonApi('/api/sets/update',{...credentials(),...item,article,discounted_price:price});if(!data.success)throw new Error(data.message);setMessage('Цена обновлена.','ok');await loadState()}catch(error){handleApiError(error)}}
    async function saveEditor(){if(!validateCredentials())return;try{setMessage('Сохранение набора...');const data=await jsonApi('/api/sets/update',{...credentials(),article:document.getElementById('edit-article').value,title:document.getElementById('edit-title').value,display_articles:document.getElementById('edit-products').value,discounted_price:document.getElementById('edit-price').value,currency:document.getElementById('edit-currency').value,sort_order:document.getElementById('edit-sort').value,discount_percent:document.getElementById('edit-discount').value,enabled:document.getElementById('edit-enabled').value});if(!data.success)throw new Error(data.message);setMessage('Набор обновлен.','ok');document.getElementById('editor').classList.remove('visible');await loadState()}catch(error){handleApiError(error)}}
    async function deleteArticles(articles){if(!validateCredentials())return;if(!confirm(`Удалить наборов: ${articles.length}? Это действие будет отправлено в Хорошоп.`))return;try{setMessage('Удаление наборов...');const data=await jsonApi('/api/sets/delete',{...credentials(),articles});const failed=data.items.filter(item=>!item.success);setMessage(failed.length?`Удаление завершено с ошибками: ${failed.map(item=>item.article).join(', ')}`:'Выбранные наборы удалены. ',failed.length?'error':'ok');await loadState()}catch(error){handleApiError(error)}}
    async function deleteSelected(){const articles=[...document.querySelectorAll('[data-select]:checked')].map(node=>node.dataset.select);if(!articles.length){setMessage('Выберите наборы для удаления.','error');return}deleteArticles(articles)}
    async function fileAction(endpoint, importing){const file=document.getElementById('file').files[0];if(!file){setMessage('Выберите Excel-файл.','error');return}if(!validateCredentials())return;const body=new FormData();body.append('file',file);for(const [key,value] of Object.entries(credentials()))body.append(key,value);try{setMessage(importing?'Выполнение операций...':'Проверка файла...');const response=await fetch(endpoint,{method:'POST',body});const data=await readResponse(response);if(!response.ok)throw new Error(data.detail||'Ошибка сервера.');renderResult(data.items||[]);setMessage(importing?`Готово: успешно ${data.imported}, ошибок ${data.errors}.`:`Готово: без ошибок ${data.ready}, ошибок ${data.errors}.`,data.errors?'error':'ok');if(importing)await loadState()}catch(error){handleApiError(error)}}
    document.getElementById('preview').addEventListener('click',()=>fileAction('/api/preview',false));document.getElementById('import').addEventListener('click',()=>fileAction('/api/import',true));document.getElementById('delete-selected').addEventListener('click',deleteSelected);document.getElementById('save-editor').addEventListener('click',saveEditor);document.getElementById('cancel-editor').addEventListener('click',()=>document.getElementById('editor').classList.remove('visible'));['login','password','token'].forEach(id=>document.getElementById(id).addEventListener('input',clearCredentialErrors));loadState();
  </script>
</body></html>"""
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


def page_html_uk() -> str:
    return """<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Набори Хорошоп</title>
  <style>
    :root { font-family:Arial,sans-serif; color:#183126; background:#f3f6f4; } * { box-sizing:border-box; } body { margin:0; }
    header { background:#135c3c; color:#fff; padding:22px max(20px,calc((100% - 1180px)/2)); } h1 { margin:0; font-size:26px; } header p { margin:6px 0 0; color:#dcefe4; }
    main { max-width:1180px; margin:auto; padding:20px; } section { background:#fff; border:1px solid #d5e1d9; border-radius:6px; padding:18px; margin-bottom:16px; } h2 { margin:0 0 14px; font-size:18px; }
    .grid { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; } .grid.two { grid-template-columns:repeat(2,minmax(0,1fr)); }
    label { display:grid; gap:6px; font-size:13px; font-weight:700; } input,select { min-height:38px; width:100%; border:1px solid #adc1b3; border-radius:4px; padding:8px; font:inherit; } input.input-error { border-color:#b42332; background:#fff1f2; outline:2px solid #f8c6cc; }
    .actions { display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-top:14px; } button,.link-button { border:1px solid #135c3c; border-radius:4px; padding:9px 13px; font:inherit; font-weight:700; background:#135c3c; color:#fff; cursor:pointer; text-decoration:none; } button.secondary,.link-button.secondary { background:#f1f6f2; color:#15482e; border-color:#b8cabb; } button.danger { background:#a52834; border-color:#a52834; }
    .note { margin:12px 0 0; padding:10px 12px; border-left:4px solid #d09120; background:#fff8e7; color:#5b481a; line-height:1.45; } .hint { color:#4f6456; line-height:1.45; } .table-wrap { overflow-x:auto; }
    table { min-width:980px; width:100%; border-collapse:collapse; font-size:13px; } th,td { padding:9px; text-align:left; vertical-align:top; border-bottom:1px solid #e0e9e3; } th { color:#4b6355; font-size:11px; text-transform:uppercase; }
    .price-cell { display:flex; min-width:135px; gap:5px; } .price-cell input { min-height:32px; min-width:75px; padding:5px; } .price-cell button { padding:5px 8px; } .tag { display:inline-block; border-radius:3px; padding:3px 6px; font-size:11px; font-weight:700; background:#e8eef0; color:#354b55; } .tag.synced,.tag.registered { background:#dff3e7; color:#0c5532; } .tag.error,.tag.delete_error { background:#fae4e6; color:#8d1826; }
    .empty { color:#6b7b70; } #editor { display:none; } #editor.visible { display:block; } #activity-log { display:grid; gap:7px; max-height:260px; overflow:auto; } .log-entry { padding:9px 10px; border-left:4px solid #7b8c81; background:#f5f8f6; line-height:1.4; } .log-entry.ok { border-color:#177245; background:#eaf7ee; } .log-entry.error { border-color:#b42332; background:#fff0f1; }
    @media(max-width:720px) { main { padding:14px; } section { padding:14px; } .grid,.grid.two { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header><h1>Набори Хорошоп</h1><p>Створення та керування наборами товарів.</p></header>
  <main>
    <section><h2>Доступ до API</h2><div class="grid"><label>Логін API<input id="login" autocomplete="username"></label><label>Пароль API<input id="password" type="password" autocomplete="current-password"></label><label>Токен API<input id="token"><span class="hint">Можна вказати замість логіна та пароля.</span></label></div></section>
    <section><h2>Масове керування через Excel</h2><div class="grid"><label>Excel-файл наборів<input id="file" type="file" accept=".xlsx,.xlsm"></label><div class="actions"><a class="link-button secondary" href="/api/template">Завантажити шаблон</a><a class="link-button secondary" href="/api/sets/export">Вивантажити реєстр</a></div></div><p class="hint">Перші три колонки: артикул набору, артикули відображення товарів через <b>;</b>, ціна. У колонці «Дія» доступні: <b>оновити</b>, <b>видалити</b>, <b>прийняти на облік</b>.</p><div class="actions"><button class="secondary" id="preview" type="button">Перевірити Excel</button><button id="import" type="button">Виконати операції</button></div></section>
    <section><h2>Результат обробки файлу</h2><div class="table-wrap" id="result"><p class="empty">Файл ще не перевірявся.</p></div></section>
    <section><h2>Реєстр наборів</h2><p class="note">Тут показані лише набори, створені через сервіс або прийняті на облік. Хорошоп не надає масового експорту наборів, тому невідомі реєстру набори автоматично не видаляються.</p><div class="actions"><button class="danger" id="delete-selected" type="button">Видалити вибрані</button><a class="link-button secondary" href="/api/sets/export">Завантажити таблицю реєстру</a></div><div class="table-wrap" id="state"><p class="empty">Завантаження реєстру...</p></div></section>
    <section id="editor"><h2>Редагування набору</h2><div class="grid two"><label>Артикул набору<input id="edit-article" readonly></label><label>Назва<input id="edit-title"></label><label>Артикули відображення товарів<input id="edit-products"><span class="hint">Розділяйте значення крапкою з комою.</span></label><label>Ціна набору<input id="edit-price" inputmode="decimal"></label><label>Валюта<input id="edit-currency" maxlength="3"></label><label>Порядок сортування<input id="edit-sort" inputmode="numeric"></label><label>Знижка %<input id="edit-discount" inputmode="numeric"></label><label>Активний<select id="edit-enabled"><option value="true">Так</option><option value="false">Ні</option></select></label></div><div class="actions"><button id="save-editor" type="button">Зберегти зміни</button><button class="secondary" id="cancel-editor" type="button">Закрити</button></div></section>
    <section><h2>Журнал операцій</h2><div id="activity-log"><p class="empty">Операцій ще не було.</p></div></section>
  </main>
  <script>
    const activityLog=document.getElementById('activity-log'), result=document.getElementById('result'), state=document.getElementById('state'); let registry=[];
    const esc=value=>String(value??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    const credentials=()=>({login:document.getElementById('login').value,password:document.getElementById('password').value,token:document.getElementById('token').value});
    const actionText=value=>({upsert:'оновити',delete:'видалити',register:'прийняти на облік'}[value]||value);
    const statusText=value=>({ready:'готово',synced:'синхронізовано',registered:'на обліку',deleted:'видалено',invalid:'некоректно',error:'помилка',delete_error:'помилка видалення'}[value]||value);
    function setMessage(text,kind=''){const empty=activityLog.querySelector('.empty');if(empty)empty.remove();const entry=document.createElement('div');entry.className='log-entry '+kind;entry.textContent=`[${new Date().toLocaleTimeString()}] ${text}`;activityLog.prepend(entry)}
    function clearCredentialErrors(){['login','password','token'].forEach(id=>document.getElementById(id).classList.remove('input-error'))}
    function validateCredentials(){const data=credentials();clearCredentialErrors();if(data.token.trim())return true;let valid=true;if(!data.login.trim()){document.getElementById('login').classList.add('input-error');valid=false}if(!data.password.trim()){document.getElementById('password').classList.add('input-error');valid=false}if(!valid)setMessage('Вкажіть логін і пароль API або токен.','error');return valid}
    function handleApiError(error){if(/логин|логін|парол|токен/i.test(error.message))validateCredentials();setMessage(error.message,'error')}
    async function readResponse(response){const raw=await response.text();try{return JSON.parse(raw)}catch{return {detail:raw||'Сервер повернув порожню відповідь.'}}}
    async function jsonApi(url,body){const response=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});const data=await readResponse(response);if(!response.ok)throw new Error(data.detail||'Помилка сервера.');return data}
    function renderResult(items){if(!items.length){result.innerHTML='<p class="empty">У файлі немає операцій.</p>';return}result.innerHTML=`<table><thead><tr><th>Рядок</th><th>Дія</th><th>Набір</th><th>Товари</th><th>Ціна</th><th>Стан</th></tr></thead><tbody>${items.map(item=>`<tr><td>${item.row_number}</td><td>${esc(actionText(item.action))}</td><td>${esc(item.article)}</td><td>${item.display_articles.map(esc).join('; ')||'-'}</td><td>${esc(item.discounted_price||'-')}</td><td><span class="tag ${esc(item.status)}">${esc(statusText(item.status))}</span>${item.message?'<br>'+esc(item.message):''}</td></tr>`).join('')}</tbody></table>`}
    function renderState(data){registry=data.sets||[];if(!registry.length){state.innerHTML='<p class="empty">У реєстрі поки немає наборів.</p>';return}state.innerHTML=`<table><thead><tr><th></th><th>Артикул</th><th>Товари</th><th>Ціна</th><th>Назва</th><th>Активний</th><th>Стан</th><th>Змінено</th><th></th></tr></thead><tbody>${registry.map(item=>`<tr><td><input type="checkbox" data-select="${esc(item.article)}"></td><td>${esc(item.article)}</td><td>${(item.display_articles||[]).map(esc).join('; ')}</td><td><div class="price-cell"><input data-price="${esc(item.article)}" value="${esc(item.discounted_price)}"><button class="secondary" data-save-price="${esc(item.article)}" type="button">OK</button></div></td><td>${esc(item.title||'Вместе дешевле')}</td><td>${item.enabled?'Так':'Ні'}</td><td><span class="tag ${esc(item.status)}">${esc(statusText(item.status))}</span><br>${esc(item.message||'')}</td><td>${esc(item.updated_at)}</td><td><div class="actions"><button class="secondary" data-edit="${esc(item.article)}" type="button">Змінити</button><button class="danger" data-delete-one="${esc(item.article)}" type="button">Видалити</button></div></td></tr>`).join('')}</tbody></table>`;document.querySelectorAll('[data-edit]').forEach(button=>button.addEventListener('click',()=>openEditor(button.dataset.edit)));document.querySelectorAll('[data-save-price]').forEach(button=>button.addEventListener('click',()=>quickPrice(button.dataset.savePrice)));document.querySelectorAll('[data-delete-one]').forEach(button=>button.addEventListener('click',()=>deleteArticles([button.dataset.deleteOne])))}
    async function loadState(){try{const response=await fetch('/api/state');const data=await readResponse(response);if(!response.ok)throw new Error(data.detail);renderState(data)}catch(error){state.innerHTML='<p class="empty">Не вдалося завантажити реєстр: '+esc(error.message)+'</p>';setMessage(error.message,'error')}}
    function openEditor(article){const item=registry.find(value=>value.article===article);if(!item)return;document.getElementById('editor').classList.add('visible');document.getElementById('edit-article').value=item.article;document.getElementById('edit-title').value=item.title||'Вместе дешевле';document.getElementById('edit-products').value=(item.display_articles||[]).join('; ');document.getElementById('edit-price').value=item.discounted_price||'';document.getElementById('edit-currency').value=item.currency||'UAH';document.getElementById('edit-sort').value=item.sort_order??'';document.getElementById('edit-discount').value=item.discount_percent??'';document.getElementById('edit-enabled').value=String(item.enabled!==false);document.getElementById('editor').scrollIntoView({behavior:'smooth',block:'start'})}
    async function quickPrice(article){if(!validateCredentials())return;const item=registry.find(value=>value.article===article);const price=document.querySelector(`[data-price="${CSS.escape(article)}"]`).value;try{setMessage('Збереження ціни...');const data=await jsonApi('/api/sets/update',{...credentials(),...item,article,discounted_price:price});if(!data.success)throw new Error(data.message);setMessage('Ціну оновлено.','ok');await loadState()}catch(error){handleApiError(error)}}
    async function saveEditor(){if(!validateCredentials())return;try{setMessage('Збереження набору...');const data=await jsonApi('/api/sets/update',{...credentials(),article:document.getElementById('edit-article').value,title:document.getElementById('edit-title').value,display_articles:document.getElementById('edit-products').value,discounted_price:document.getElementById('edit-price').value,currency:document.getElementById('edit-currency').value,sort_order:document.getElementById('edit-sort').value,discount_percent:document.getElementById('edit-discount').value,enabled:document.getElementById('edit-enabled').value});if(!data.success)throw new Error(data.message);setMessage('Набір оновлено.','ok');document.getElementById('editor').classList.remove('visible');await loadState()}catch(error){handleApiError(error)}}
    async function deleteArticles(articles){if(!validateCredentials())return;if(!confirm(`Видалити наборів: ${articles.length}? Дію буде надіслано до Хорошоп.`))return;try{setMessage('Видалення наборів...');const data=await jsonApi('/api/sets/delete',{...credentials(),articles});const failed=data.items.filter(item=>!item.success);setMessage(failed.length?`Видалення завершено з помилками: ${failed.map(item=>item.article).join(', ')}`:'Вибрані набори видалено.',failed.length?'error':'ok');await loadState()}catch(error){handleApiError(error)}}
    async function deleteSelected(){const articles=[...document.querySelectorAll('[data-select]:checked')].map(node=>node.dataset.select);if(!articles.length){setMessage('Виберіть набори для видалення.','error');return}deleteArticles(articles)}
    async function fileAction(endpoint,importing){const file=document.getElementById('file').files[0];if(!file){setMessage('Виберіть Excel-файл.','error');return}if(!validateCredentials())return;const body=new FormData();body.append('file',file);for(const [key,value] of Object.entries(credentials()))body.append(key,value);try{setMessage(importing?'Виконання операцій...':'Перевірка файлу...');const response=await fetch(endpoint,{method:'POST',body});const data=await readResponse(response);if(!response.ok)throw new Error(data.detail||'Помилка сервера.');renderResult(data.items||[]);setMessage(importing?`Готово: успішно ${data.imported}, помилок ${data.errors}.`:`Готово до імпорту: ${data.ready}, помилок ${data.errors}.`,data.errors?'error':'ok');if(importing)await loadState()}catch(error){handleApiError(error)}}
    document.getElementById('preview').addEventListener('click',()=>fileAction('/api/preview',false));document.getElementById('import').addEventListener('click',()=>fileAction('/api/import',true));document.getElementById('delete-selected').addEventListener('click',deleteSelected);document.getElementById('save-editor').addEventListener('click',saveEditor);document.getElementById('cancel-editor').addEventListener('click',()=>document.getElementById('editor').classList.remove('visible'));['login','password','token'].forEach(id=>document.getElementById(id).addEventListener('input',clearCredentialErrors));loadState();
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


def run_server() -> None:
    import uvicorn

    runtime_settings, _ = get_runtime()
    configure_service_output(runtime_settings)
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)


if __name__ == "__main__":
    run_server()
