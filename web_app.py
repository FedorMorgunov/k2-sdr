#!/usr/bin/env python3
"""Локальный веб-интерфейс для рассылки писем из OWA (Exchange/EWS).

Запускает маленький сервер на 127.0.0.1 и открывает страницу в браузере.
Работает одинаково на macOS и Windows. Вся логика отправки переиспользуется
из mailer.py.

Запуск из исходников:
    pip install -r requirements.txt
    python web_app.py

Сборка в приложение (.app для macOS) — см. build_mac.sh и README.

Особенности:
  * Пароль хранится только в оперативной памяти процесса, никуда не пишется.
  * Настройки подключения (без пароля) и шаблоны сохраняются в папке
    ~/K2-Mailer, туда же кладётся состояние рассылки (sent_state.json),
    чтобы можно было дослать письма в ту же ветку даже после перезапуска.
"""

from __future__ import annotations

import os
import threading
import time
import webbrowser
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, request, jsonify

import mailer
from mailer import (
    ColumnMap,
    ExchangeClient,
    ExchangeConfig,
    MailerError,
    SentRecord,
    load_state,
    read_contacts,
    render_template,
    save_state,
    upsert_state,
)

APP_NAME = "K2 Mailer"
HOST = "127.0.0.1"
PORT = int(os.environ.get("K2_MAILER_PORT", "8765"))

DEFAULT_LETTER = (
    "Здравствуйте, {имя}!\n\n"
    "Меня зовут Фёдор, я представляю компанию K2 SDR. Обращаюсь к вам как к\n"
    "представителю компании «{компания}».\n\n"
    "Будет здорово обсудить, насколько наше предложение актуально для вас —\n"
    "для этого достаточно короткого звонка на 15 минут.\n\n"
    "С уважением,\nФёдор Моргунов\nK2 SDR\n"
)
DEFAULT_FOLLOWUP = (
    "Здравствуйте, {имя}!\n\n"
    "Возвращаюсь к своему предыдущему письму — возможно, оно затерялось в потоке.\n"
    "Всё ещё считаю, что наше предложение может быть полезно «{компания}».\n\n"
    "Подскажите, есть ли интерес обсудить детали?\n\n"
    "С уважением,\nФёдор Моргунов\nK2 SDR\n"
)


# --------------------------------------------------------------------------- #
# Рабочая папка и файлы
# --------------------------------------------------------------------------- #
def workdir() -> Path:
    base = Path(os.environ.get("K2_MAILER_HOME") or (Path.home() / "K2-Mailer"))
    base.mkdir(parents=True, exist_ok=True)
    return base


def state_file() -> Path:
    return workdir() / "sent_state.json"


def excel_file() -> Path:
    return workdir() / "contacts.xlsx"


def conn_file() -> Path:
    return workdir() / "connection.json"


def template_file(kind: str) -> Path:
    return workdir() / (f"{kind}.txt")


def default_template(kind: str) -> str:
    saved = template_file(kind)
    if saved.exists():
        return saved.read_text(encoding="utf-8")
    return DEFAULT_LETTER if kind == "letter" else DEFAULT_FOLLOWUP


# --------------------------------------------------------------------------- #
# Хранилище настроек подключения (без пароля)
# --------------------------------------------------------------------------- #
import json


def read_conn() -> dict:
    if conn_file().exists():
        return json.loads(conn_file().read_text(encoding="utf-8"))
    return {}


def write_conn(data: dict) -> None:
    conn_file().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def current_config() -> ExchangeConfig:
    d = read_conn()
    email = (d.get("email") or "").strip()
    if not email:
        raise MailerError("Не заполнен адрес почты в разделе «Подключение».")
    return ExchangeConfig(
        email=email,
        username=(d.get("username") or "").strip() or email,
        ews_url=(d.get("ews_url") or "").strip(),
        auth_type=(d.get("auth_type") or "auto").strip() or "auto",
        verify_ssl=bool(d.get("verify_ssl", True)),
    )


def current_cols() -> ColumnMap:
    c = (read_conn().get("columns") or {})
    base = ColumnMap()
    return ColumnMap(
        name=(c.get("name") or base.name),
        company=(c.get("company") or base.company),
        email=(c.get("email") or base.email),
        subject=(c.get("subject") or base.subject),
    )


# --------------------------------------------------------------------------- #
# Состояние в памяти: пароль, загруженные контакты, текущая задача
# --------------------------------------------------------------------------- #
SESSION = {"password": "", "contacts": []}  # contacts: list[mailer.Contact]

JOB_LOCK = threading.Lock()
JOB = {
    "running": False, "kind": "", "total": 0, "done": 0,
    "sent": 0, "failed": 0, "log": [], "error": "", "finished": False,
}


def job_reset(kind: str) -> None:
    with JOB_LOCK:
        JOB.update(running=True, kind=kind, total=0, done=0,
                   sent=0, failed=0, log=[], error="", finished=False)


def job_log(line: str) -> None:
    stamp = datetime.now().strftime("%H:%M:%S")
    with JOB_LOCK:
        JOB["log"].append(f"[{stamp}] {line}")


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# --------------------------------------------------------------------------- #
# Фоновые задачи: рассылка и повторная отправка
# --------------------------------------------------------------------------- #
def worker_send(template: str, delay: float, limit: int, resume: bool) -> None:
    try:
        cfg = current_config()
        contacts = list(SESSION["contacts"])
        if not contacts:
            raise MailerError("Сначала загрузите Excel со списком контактов.")
        if not SESSION["password"]:
            raise MailerError("Сначала укажите пароль и нажмите «Проверить подключение».")

        records = load_state(state_file())
        already = {r.email.lower() for r in records if r.status == "sent"}
        todo = [c for c in contacts if not (resume and c.email.lower() in already)]
        if limit > 0:
            todo = todo[:limit]
        if not todo:
            raise MailerError("Нет писем к отправке (возможно, все уже отправлены).")

        with JOB_LOCK:
            JOB["total"] = len(todo)
        job_log(f"Подключение к Exchange как {cfg.email} …")
        client = ExchangeClient(cfg, SESSION["password"])
        job_log("Подключение установлено. Начинаю рассылку.")

        for i, c in enumerate(todo, start=1):
            body = render_template(template, c)
            subject = c.subject or "(без темы)"
            try:
                mid, conv = client.send_new(c.email, subject, body)
                upsert_state(records, SentRecord(
                    email=c.email, name=c.name, company=c.company, subject=subject,
                    status="sent", internet_message_id=mid, conversation_id=conv,
                    sent_at=_now_iso(),
                ))
                with JOB_LOCK:
                    JOB["sent"] += 1
                job_log(f"✓ [{i}/{len(todo)}] {c.email} — отправлено")
            except Exception as exc:
                upsert_state(records, SentRecord(
                    email=c.email, name=c.name, company=c.company, subject=subject,
                    status="failed", error=str(exc), sent_at=_now_iso(),
                ))
                with JOB_LOCK:
                    JOB["failed"] += 1
                job_log(f"✗ [{i}/{len(todo)}] {c.email} — ошибка: {exc}")

            save_state(state_file(), records)
            with JOB_LOCK:
                JOB["done"] = i
            if delay and i < len(todo):
                time.sleep(delay)

        job_log(f"Готово. Отправлено: {JOB['sent']}, ошибок: {JOB['failed']}.")
    except MailerError as exc:
        with JOB_LOCK:
            JOB["error"] = str(exc)
        job_log(f"Ошибка: {exc}")
    except Exception as exc:  # неожиданное
        with JOB_LOCK:
            JOB["error"] = str(exc)
        job_log(f"Непредвиденная ошибка: {exc}")
    finally:
        with JOB_LOCK:
            JOB["running"] = False
            JOB["finished"] = True


def worker_followup(template: str, delay: float, limit: int, only: str) -> None:
    try:
        cfg = current_config()
        if not SESSION["password"]:
            raise MailerError("Сначала укажите пароль и нажмите «Проверить подключение».")

        records = load_state(state_file())
        if not records:
            raise MailerError("Нет данных о первой рассылке. Сначала выполните рассылку.")

        targets = [r for r in records if r.status == "sent"]
        if only.strip():
            wanted = {e.strip().lower() for e in only.split(",") if e.strip()}
            targets = [r for r in targets if r.email.lower() in wanted]
        if not targets:
            raise MailerError("Нет подходящих адресатов для повторного письма.")
        if limit > 0:
            targets = targets[:limit]

        with JOB_LOCK:
            JOB["total"] = len(targets)
        job_log(f"Подключение к Exchange как {cfg.email} …")
        client = ExchangeClient(cfg, SESSION["password"])
        job_log("Подключение установлено. Досылаю письма в ту же ветку.")

        for i, r in enumerate(targets, start=1):
            body = render_template(template, r.context_contact())
            try:
                client.send_followup(r.email, r.subject, body,
                                     in_reply_to=r.internet_message_id)
                with JOB_LOCK:
                    JOB["sent"] += 1
                job_log(f"✓ [{i}/{len(targets)}] {r.email} — дослано в ту же ветку")
            except Exception as exc:
                with JOB_LOCK:
                    JOB["failed"] += 1
                job_log(f"✗ [{i}/{len(targets)}] {r.email} — ошибка: {exc}")
            with JOB_LOCK:
                JOB["done"] = i
            if delay and i < len(targets):
                time.sleep(delay)

        job_log(f"Готово. Дослано: {JOB['sent']}, ошибок: {JOB['failed']}.")
    except MailerError as exc:
        with JOB_LOCK:
            JOB["error"] = str(exc)
        job_log(f"Ошибка: {exc}")
    except Exception as exc:
        with JOB_LOCK:
            JOB["error"] = str(exc)
        job_log(f"Непредвиденная ошибка: {exc}")
    finally:
        with JOB_LOCK:
            JOB["running"] = False
            JOB["finished"] = True


def start_job(target, *args) -> bool:
    with JOB_LOCK:
        if JOB["running"]:
            return False
    kind = "followup" if target is worker_followup else "send"
    job_reset(kind)
    threading.Thread(target=target, args=args, daemon=True).start()
    return True


# --------------------------------------------------------------------------- #
# Flask
# --------------------------------------------------------------------------- #
app = Flask(__name__)


@app.get("/")
def index():
    return INDEX_HTML


@app.get("/api/config")
def api_get_config():
    d = read_conn()
    cols = current_cols()
    records = load_state(state_file())
    sent = sum(1 for r in records if r.status == "sent")
    return jsonify({
        "email": d.get("email", ""),
        "username": d.get("username", ""),
        "ews_url": d.get("ews_url", ""),
        "auth_type": d.get("auth_type", "NTLM"),
        "verify_ssl": bool(d.get("verify_ssl", True)),
        "columns": {"name": cols.name, "company": cols.company,
                    "email": cols.email, "subject": cols.subject},
        "letter": default_template("letter"),
        "followup": default_template("followup"),
        "has_password": bool(SESSION["password"]),
        "contacts_count": len(SESSION["contacts"]),
        "sent_count": sent,
        "workdir": str(workdir()),
    })


@app.post("/api/connect")
def api_connect():
    data = request.get_json(force=True)
    conn = {
        "email": (data.get("email") or "").strip(),
        "username": (data.get("username") or "").strip(),
        "ews_url": (data.get("ews_url") or "").strip(),
        "auth_type": (data.get("auth_type") or "NTLM").strip(),
        "verify_ssl": bool(data.get("verify_ssl", True)),
        "columns": data.get("columns") or read_conn().get("columns") or {},
    }
    write_conn(conn)
    password = data.get("password") or ""
    if password:
        SESSION["password"] = password
    if not SESSION["password"]:
        return jsonify({"ok": False, "error": "Введите пароль."}), 400
    try:
        cfg = current_config()
        client = ExchangeClient(cfg, SESSION["password"])
        # лёгкая проверка, что доступ есть
        _ = client.account.inbox.total_count
        return jsonify({"ok": True, "message": f"Подключено: {cfg.email}"})
    except MailerError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Не удалось подключиться: {exc}"}), 400


@app.post("/api/save_settings")
def api_save_settings():
    """Сохранить настройки/колонки без проверки подключения."""
    data = request.get_json(force=True)
    conn = read_conn()
    for k in ("email", "username", "ews_url", "auth_type"):
        if k in data:
            conn[k] = (data.get(k) or "").strip()
    if "verify_ssl" in data:
        conn["verify_ssl"] = bool(data["verify_ssl"])
    if data.get("columns"):
        conn["columns"] = data["columns"]
    write_conn(conn)
    if data.get("password"):
        SESSION["password"] = data["password"]
    return jsonify({"ok": True})


@app.post("/api/upload")
def api_upload():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    path = excel_file()
    f.save(path)
    try:
        contacts = read_contacts(path, current_cols())
    except MailerError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    SESSION["contacts"] = contacts
    preview = [{"name": c.name, "company": c.company, "email": c.email,
                "subject": c.subject} for c in contacts[:5]]
    return jsonify({"ok": True, "count": len(contacts), "preview": preview})


@app.post("/api/save_templates")
def api_save_templates():
    data = request.get_json(force=True)
    if "letter" in data:
        template_file("letter").write_text(data["letter"], encoding="utf-8")
    if "followup" in data:
        template_file("followup").write_text(data["followup"], encoding="utf-8")
    return jsonify({"ok": True})


@app.post("/api/preview")
def api_preview():
    data = request.get_json(force=True)
    template = data.get("template", "")
    which = data.get("which", "letter")
    count = int(data.get("count", 3))

    if which == "followup":
        records = [r for r in load_state(state_file()) if r.status == "sent"]
        items = [r.context_contact() for r in records[:count]]
        if not items:
            return jsonify({"ok": False, "error": "Нет данных о первой рассылке для предпросмотра."}), 400
    else:
        items = SESSION["contacts"][:count]
        if not items:
            return jsonify({"ok": False, "error": "Сначала загрузите Excel."}), 400

    result = []
    for c in items:
        subject = c.subject or "(без темы)"
        if which == "followup" and not subject.lower().startswith("re:"):
            subject = f"RE: {subject}"
        result.append({"email": c.email, "subject": subject,
                       "body": render_template(template, c)})
    return jsonify({"ok": True, "items": result})


@app.post("/api/send")
def api_send():
    data = request.get_json(force=True)
    template = data.get("template", "")
    template_file("letter").write_text(template, encoding="utf-8")
    delay = float(data.get("delay", 1.0))
    limit = int(data.get("limit", 0))
    resume = bool(data.get("resume", False))
    if start_job(worker_send, template, delay, limit, resume):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Уже выполняется другая задача."}), 409


@app.post("/api/followup")
def api_followup():
    data = request.get_json(force=True)
    template = data.get("template", "")
    template_file("followup").write_text(template, encoding="utf-8")
    delay = float(data.get("delay", 1.0))
    limit = int(data.get("limit", 0))
    only = data.get("only", "")
    if start_job(worker_followup, template, delay, limit, only):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Уже выполняется другая задача."}), 409


@app.get("/api/status")
def api_status():
    with JOB_LOCK:
        return jsonify(dict(JOB))


# --------------------------------------------------------------------------- #
# Страница (одностраничное приложение, без внешних ресурсов/CDN)
# --------------------------------------------------------------------------- #
INDEX_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>K2 Mailer</title>
<style>
  :root { --bg:#f5f6f8; --card:#fff; --line:#e3e6ea; --pri:#2b6cb0; --pri2:#234e7d;
          --ok:#2f855a; --err:#c53030; --mut:#667085; }
  * { box-sizing: border-box; }
  body { margin:0; font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
         background:var(--bg); color:#1a202c; }
  header { background:var(--pri2); color:#fff; padding:16px 24px; }
  header h1 { margin:0; font-size:18px; }
  header .sub { opacity:.8; font-size:13px; }
  main { max-width:920px; margin:0 auto; padding:20px; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px;
          padding:18px 20px; margin-bottom:18px; }
  .card h2 { margin:0 0 14px; font-size:16px; display:flex; align-items:center; gap:8px; }
  .step { background:var(--pri); color:#fff; width:24px; height:24px; border-radius:50%;
          display:inline-flex; align-items:center; justify-content:center; font-size:13px; }
  label { display:block; font-size:13px; color:var(--mut); margin:8px 0 4px; }
  input[type=text], input[type=password], input[type=number], select, textarea {
      width:100%; padding:9px 10px; border:1px solid var(--line); border-radius:7px;
      font:inherit; background:#fff; }
  textarea { min-height:150px; resize:vertical; font-family:ui-monospace,Menlo,Consolas,monospace; font-size:13px; }
  .row { display:flex; gap:14px; flex-wrap:wrap; }
  .row > div { flex:1; min-width:200px; }
  .inline { display:flex; align-items:center; gap:8px; margin-top:10px; }
  .inline input[type=checkbox] { width:auto; }
  button { font:inherit; cursor:pointer; border:none; border-radius:7px; padding:9px 16px;
           background:var(--pri); color:#fff; }
  button.secondary { background:#edf2f7; color:#2d3748; border:1px solid var(--line); }
  button.danger { background:var(--err); }
  button:disabled { opacity:.5; cursor:not-allowed; }
  .btns { display:flex; gap:10px; flex-wrap:wrap; margin-top:14px; }
  .msg { margin-top:10px; font-size:13px; }
  .msg.ok { color:var(--ok); } .msg.err { color:var(--err); }
  table { width:100%; border-collapse:collapse; font-size:13px; margin-top:10px; }
  th,td { text-align:left; padding:6px 8px; border-bottom:1px solid var(--line); }
  th { color:var(--mut); font-weight:600; }
  .hint { font-size:12px; color:var(--mut); margin-top:6px; }
  .log { background:#0b1220; color:#cde3ff; border-radius:8px; padding:12px;
         font-family:ui-monospace,Menlo,Consolas,monospace; font-size:12.5px;
         height:240px; overflow:auto; white-space:pre-wrap; }
  .bar { height:8px; background:#e2e8f0; border-radius:6px; overflow:hidden; margin:10px 0; }
  .bar > i { display:block; height:100%; width:0; background:var(--ok); transition:width .3s; }
  .pill { font-size:12px; padding:2px 8px; border-radius:20px; background:#edf2f7; color:#2d3748; }
  .preview { background:#f8fafc; border:1px solid var(--line); border-radius:8px; padding:10px 12px;
             margin-top:8px; }
  .preview .subj { font-weight:600; }
  .preview pre { white-space:pre-wrap; margin:6px 0 0; font-size:13px; }
  code { background:#eef2f7; padding:1px 5px; border-radius:4px; }
</style>
</head>
<body>
<header>
  <h1>K2 Mailer — рассылка из рабочей почты</h1>
  <div class="sub">Локально, на вашем компьютере. Письма уходят из вашего ящика OWA.</div>
</header>
<main>

  <!-- 1. Подключение -->
  <div class="card">
    <h2><span class="step">1</span> Подключение к почте</h2>
    <div class="row">
      <div><label>Ваш email (отправитель)</label><input id="email" type="text" placeholder="ivanov@company.ru"></div>
      <div><label>Логин (часто DOMAIN\\user, можно оставить пустым)</label><input id="username" type="text" placeholder="COMPANY\\ivanov"></div>
    </div>
    <div class="row">
      <div><label>Адрес EWS (пусто = автоопределение)</label><input id="ews_url" type="text" placeholder="https://mail.company.ru/EWS/Exchange.asmx"></div>
      <div style="max-width:160px"><label>Аутентификация</label>
        <select id="auth_type"><option>NTLM</option><option>basic</option><option value="auto">auto</option></select>
      </div>
    </div>
    <div class="row">
      <div><label>Пароль (нигде не сохраняется)</label><input id="password" type="password" placeholder="••••••••"></div>
    </div>
    <div class="inline"><input id="verify_ssl" type="checkbox" checked><label style="margin:0">Проверять SSL-сертификат (снимите только при самоподписанном)</label></div>
    <div class="btns">
      <button id="btnConnect">Проверить подключение</button>
    </div>
    <div id="connMsg" class="msg"></div>
  </div>

  <!-- 2. Контакты -->
  <div class="card">
    <h2><span class="step">2</span> Список контактов (Excel)</h2>
    <div class="hint">Колонки: <code>Имя</code>, <code>Компания</code>, <code>Почта</code>, <code>Тема письма</code>. Первая строка — заголовки.</div>
    <div class="btns">
      <input id="file" type="file" accept=".xlsx">
      <button id="btnUpload" class="secondary">Загрузить</button>
      <span id="contactsPill" class="pill">не загружено</span>
    </div>
    <div id="uploadMsg" class="msg"></div>
    <div id="previewTable"></div>
  </div>

  <!-- 3. Шаблоны -->
  <div class="card">
    <h2><span class="step">3</span> Шаблоны писем</h2>
    <div class="hint">Переменные: <code>{имя}</code> и <code>{компания}</code> (можно <code>{name}</code>/<code>{company}</code>).</div>
    <label>Текст первого письма</label>
    <textarea id="letter"></textarea>
    <label style="margin-top:12px">Текст повторного письма (в ту же ветку)</label>
    <textarea id="followup"></textarea>
    <div class="btns">
      <button class="secondary" id="btnPreviewLetter">Предпросмотр первого</button>
      <button class="secondary" id="btnPreviewFollow">Предпросмотр повторного</button>
      <button class="secondary" id="btnSaveTpl">Сохранить шаблоны</button>
    </div>
    <div id="previewBox"></div>
  </div>

  <!-- 4. Отправка -->
  <div class="card">
    <h2><span class="step">4</span> Отправка</h2>
    <div class="row">
      <div style="max-width:170px"><label>Пауза между письмами, сек</label><input id="delay" type="number" value="1" min="0" step="0.5"></div>
      <div style="max-width:170px"><label>Лимит (0 = все)</label><input id="limit" type="number" value="0" min="0"></div>
    </div>
    <div class="inline"><input id="resume" type="checkbox" checked><label style="margin:0">Пропускать уже отправленные (для рассылки)</label></div>
    <div class="btns">
      <button id="btnSend">Отправить рассылку</button>
      <button id="btnFollow" class="secondary">Дослать в ту же ветку</button>
    </div>
    <div class="hint">Повторное письмо уходит тем, кому первая рассылка прошла успешно (данные берутся из истории отправки).</div>
    <div class="bar"><i id="progBar"></i></div>
    <div id="jobStatus" class="msg"></div>
    <div class="log" id="log"></div>
  </div>

  <div class="hint" id="workdir"></div>
</main>

<script>
const $ = id => document.getElementById(id);
async function api(url, opts){ const r = await fetch(url, opts); return await r.json(); }
async function postJSON(url, body){
  return api(url, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)});
}
function setMsg(el, text, ok){ el.textContent = text; el.className = 'msg ' + (ok ? 'ok' : 'err'); }

function collectConn(){
  return {
    email: $('email').value, username: $('username').value, ews_url: $('ews_url').value,
    auth_type: $('auth_type').value, verify_ssl: $('verify_ssl').checked,
    password: $('password').value,
  };
}

async function loadConfig(){
  const c = await api('/api/config');
  $('email').value = c.email || ''; $('username').value = c.username || '';
  $('ews_url').value = c.ews_url || ''; $('auth_type').value = c.auth_type || 'NTLM';
  $('verify_ssl').checked = c.verify_ssl !== false;
  $('letter').value = c.letter || ''; $('followup').value = c.followup || '';
  $('contactsPill').textContent = c.contacts_count ? (c.contacts_count + ' контактов') : 'не загружено';
  $('workdir').textContent = 'Данные и история хранятся в папке: ' + c.workdir;
  if (c.sent_count) setMsg($('jobStatus'), 'В истории отправки: ' + c.sent_count + ' получателей (доступно «дослать»).', true);
}

$('btnConnect').onclick = async () => {
  setMsg($('connMsg'), 'Проверяю подключение…', true);
  const r = await postJSON('/api/connect', collectConn());
  setMsg($('connMsg'), r.ok ? r.message : ('Ошибка: ' + r.error), r.ok);
};

$('btnUpload').onclick = async () => {
  const f = $('file').files[0];
  if (!f){ setMsg($('uploadMsg'), 'Выберите файл .xlsx', false); return; }
  // сохраним настройки колонок/подключения перед чтением
  await postJSON('/api/save_settings', collectConn());
  const fd = new FormData(); fd.append('file', f);
  const r = await api('/api/upload', {method:'POST', body: fd});
  if (!r.ok){ setMsg($('uploadMsg'), 'Ошибка: ' + r.error, false); $('previewTable').innerHTML=''; return; }
  setMsg($('uploadMsg'), 'Загружено контактов: ' + r.count, true);
  $('contactsPill').textContent = r.count + ' контактов';
  let h = '<table><tr><th>Имя</th><th>Компания</th><th>Почта</th><th>Тема</th></tr>';
  for (const p of r.preview) h += `<tr><td>${esc(p.name)}</td><td>${esc(p.company)}</td><td>${esc(p.email)}</td><td>${esc(p.subject)}</td></tr>`;
  h += '</table><div class="hint">Показаны первые строки.</div>';
  $('previewTable').innerHTML = h;
};

$('btnSaveTpl').onclick = async () => {
  await postJSON('/api/save_templates', {letter: $('letter').value, followup: $('followup').value});
  setMsg($('jobStatus'), 'Шаблоны сохранены.', true);
};

async function preview(which){
  const template = which === 'followup' ? $('followup').value : $('letter').value;
  const r = await postJSON('/api/preview', {template, which, count: 3});
  if (!r.ok){ $('previewBox').innerHTML = `<div class="msg err">${esc(r.error)}</div>`; return; }
  let h = '';
  for (const it of r.items)
    h += `<div class="preview"><div class="subj">${esc(it.email)} — ${esc(it.subject)}</div><pre>${esc(it.body)}</pre></div>`;
  $('previewBox').innerHTML = h;
}
$('btnPreviewLetter').onclick = () => preview('letter');
$('btnPreviewFollow').onclick = () => preview('followup');

let polling = null;
function startPolling(){
  if (polling) return;
  polling = setInterval(async () => {
    const j = await api('/api/status');
    const pct = j.total ? Math.round(j.done / j.total * 100) : 0;
    $('progBar').style.width = pct + '%';
    $('log').textContent = j.log.join('\n');
    $('log').scrollTop = $('log').scrollHeight;
    let s = j.running ? `Выполняется (${j.done}/${j.total})…` : 'Готово.';
    if (j.sent || j.failed) s += `  Успешно: ${j.sent}, ошибок: ${j.failed}.`;
    setMsg($('jobStatus'), j.error ? ('Ошибка: ' + j.error) : s, !j.error);
    if (j.finished && !j.running){ clearInterval(polling); polling = null; setButtons(false); }
  }, 1000);
}
function setButtons(running){
  for (const id of ['btnSend','btnFollow']) $(id).disabled = running;
}

$('btnSend').onclick = async () => {
  if (!confirm('Отправить рассылку всем загруженным контактам?')) return;
  await postJSON('/api/save_settings', collectConn());
  const r = await postJSON('/api/send', {
    template: $('letter').value, delay: +$('delay').value,
    limit: +$('limit').value, resume: $('resume').checked,
  });
  if (!r.ok){ setMsg($('jobStatus'), 'Ошибка: ' + r.error, false); return; }
  setButtons(true); $('log').textContent=''; startPolling();
};

$('btnFollow').onclick = async () => {
  if (!confirm('Дослать повторное письмо в ту же ветку (тем, кому уже отправляли)?')) return;
  await postJSON('/api/save_settings', collectConn());
  const r = await postJSON('/api/followup', {
    template: $('followup').value, delay: +$('delay').value, limit: +$('limit').value, only: '',
  });
  if (!r.ok){ setMsg($('jobStatus'), 'Ошибка: ' + r.error, false); return; }
  setButtons(true); $('log').textContent=''; startPolling();
};

function esc(s){ return (s||'').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }
loadConfig();
</script>
</body>
</html>
"""


def open_browser_later() -> None:
    time.sleep(1.0)
    try:
        webbrowser.open(f"http://{HOST}:{PORT}")
    except Exception:
        pass


def main() -> None:
    print(f"{APP_NAME} запущен. Откройте в браузере: http://{HOST}:{PORT}")
    print(f"Рабочая папка: {workdir()}")
    threading.Thread(target=open_browser_later, daemon=True).start()
    app.run(host=HOST, port=PORT, threaded=True)


if __name__ == "__main__":
    main()
