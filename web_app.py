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

import html as htmllib
import logging
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, request, jsonify, send_file, abort

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

# Шаблоны-«открывашки». Могут содержать HTML-разметку (жирный/курсив/подчёркивание,
# абзацы) — она сохраняется в письме. Переменные {имя}/{компания} подставляются как
# обычно. Если в шаблоне нет HTML-тегов, текст уходит как обычный (с переносами строк).
DEFAULT_LETTER = (
    "<p>{имя}, добрый день!</p>\n"
    "<p>В последние 5 лет мы помогаем многим фармкомпаниям создать и разместить "
    "сайт в России. Мы обратили внимание, что сайт компании {компания} "
    "<u>размещен не в РФ</u>. Данные о размещении сайта являются публичными, что "
    "создает риск несоответствия требованиям ФЗ-152 «О персональных данных». Наши "
    "заказчики отмечают, что РКН проводит автоматизированные проверки с помощью ИИ "
    "без предварительного уведомления, штрафы за нарушение могут достигать 4 млн "
    "рублей.</p>\n"
    "<p>Мы предлагаем решение задачи по размещению сайта на территории РФ «под ключ», "
    "которое <i><u>полностью снимает риск штрафов</u></i>. Наша команда создаст сайт, "
    "визуально не отличающийся от текущего, при этом размещенном в российском ЦОДе, "
    "<b>соответствующем требованиям 152-ФЗ</b>.</p>\n"
    "<p>Мы решили задачу с размещением сайтов для более десятка иностранных "
    "фармкомпаний: порталы для личных кабинетов врачей, сайты-визитки, продуктовые "
    "лендинги. В некоторых случаях мы решали задачу локализации сайта, когда "
    "заказчик уже получал уведомление от РКН.</p>\n"
    "<p>Предлагаем показать референсы и рассказать про риски по 152-ФЗ с приглашенным "
    "юристом. Готовы ли Вы обсудить задачу локализации?</p>\n"
)
DEFAULT_FOLLOWUP = (
    "<p>{имя}, здравствуйте!</p>\n"
    "<p>Возможно, предыдущее письмо затерялось. Коротко дополню: мы реализуем "
    "подобные проекты по локализации и размещению сайтов в российском ЦОДе в "
    "течение ~<b>1</b> месяца.</p>\n"
    "<p>Подробное описание нашего предложения, этапов работ, референсы и условия "
    "размещения описаны в презентации.</p>\n"
    "<p>Подскажите, актуальна ли задача по локализации вашего сайта?</p>\n"
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
# Подпись (HTML) и фото для неё. Хранятся локально => у каждого сотрудника своя.
# --------------------------------------------------------------------------- #
PHOTO_CID = "sigphoto"
PHOTO_STEM = "signature_photo"


def signature_file() -> Path:
    return workdir() / "signature.html"


def read_signature() -> str:
    f = signature_file()
    return f.read_text(encoding="utf-8") if f.exists() else ""


def photo_path() -> Path | None:
    for p in workdir().glob(PHOTO_STEM + ".*"):
        return p
    return None


import re as _re

# Признак того, что шаблон уже содержит HTML-разметку (тогда не экранируем).
# Тег должен идти сразу после "<" или "</" и завершаться пробелом, "/" или ">",
# чтобы обычный текст вроде "a < b" не принимался за HTML.
_HTML_TAG_RE = _re.compile(
    r"</?(?:p|br|b|i|u|em|strong|div|span|ul|ol|li|a|table|tr|td|h[1-6]|blockquote)(?:\s|/|>)",
    _re.IGNORECASE,
)


def looks_like_html(text: str) -> bool:
    return bool(_HTML_TAG_RE.search(text or ""))


def text_to_html(text: str) -> str:
    """Готовит тело письма в HTML.

    Если в шаблоне есть HTML-разметку (абзацы, <b>/<i>/<u> и т.п.) — используем
    как есть, чтобы сохранить форматирование «открывашек». Если это обычный
    текст — экранируем и переносы строк превращаем в <br>.
    """
    inner = text if looks_like_html(text) else htmllib.escape(text).replace("\n", "<br>")
    return ('<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;'
            'color:#222;line-height:1.5;">' + inner + "</div>")


def build_email_html(rendered_body: str, photo_ref: str, signature=None) -> str:
    """Собирает HTML письма: текст письма + HTML-подпись.

    signature — текст подписи; если None, берётся сохранённая (для предпросмотра
    передаём текущее содержимое поля, чтобы видеть несохранённые правки).
    photo_ref подставляется вместо токена {photo} в подписи:
      * при отправке  — 'cid:sigphoto' (картинка вложена в письмо);
      * в предпросмотре — URL '/api/sigphoto' (видно в браузере).
    """
    parts = [text_to_html(rendered_body)]
    sig = (read_signature() if signature is None else signature).strip()
    if sig:
        parts.append("<br>" + sig.replace("{photo}", photo_ref or ""))
    return ("<html><body style=\"margin:0;padding:0;\">"
            + "".join(parts) + "</body></html>")


def signature_active() -> bool:
    return bool(read_signature().strip())


def inline_photo_for_send():
    """Возвращает (имя, байты, cid) если подпись использует {photo} и фото есть."""
    sig = read_signature()
    p = photo_path()
    if p and "{photo}" in sig:
        return (p.name, p.read_bytes(), PHOTO_CID)
    return None


# --------------------------------------------------------------------------- #
# Вложение для повторного письма (например, презентация в PDF).
# Хранится локально; оригинальное имя файла — в connection.json.
# --------------------------------------------------------------------------- #
ATTACH_STEM = "followup_attachment"


def attachment_path() -> Path | None:
    for p in workdir().glob(ATTACH_STEM + ".*"):
        return p
    return None


def attachment_display_name() -> str:
    p = attachment_path()
    if not p:
        return ""
    return read_conn().get("attachment_name") or p.name


def attachment_for_send():
    """Возвращает [(имя, байты)] для отправки или None, если вложения нет."""
    p = attachment_path()
    if not p:
        return None
    return [(attachment_display_name(), p.read_bytes())]


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
                html = build_email_html(body, "cid:" + PHOTO_CID)
                mid, conv = client.send_new(c.email, subject, body,
                                            html_body=html, inline_image=inline_photo_for_send())
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


def worker_followup(template: str, delay: float, limit: int, only: str,
                    skip_done: bool = True) -> None:
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
        # Защита от случайной повторной досылки: пропускаем тех, кому повторное
        # письмо уже уходило (отметка хранится в sent_state.json — то есть «глобально»,
        # а не только в рамках текущей сессии приложения).
        if skip_done:
            before = len(targets)
            targets = [r for r in targets if r.followup_status != "sent"]
            skipped = before - len(targets)
            if skipped:
                job_log(f"Пропущено уже досланных (им повторное письмо уже уходило): {skipped}")
        if not targets:
            raise MailerError("Нет адресатов для повторного письма: либо нет успешной "
                              "первой рассылки, либо всем уже дослали. Чтобы отправить "
                              "повторно намеренно — снимите галочку «Не досылать повторно».")
        if limit > 0:
            targets = targets[:limit]

        attachments = attachment_for_send()

        with JOB_LOCK:
            JOB["total"] = len(targets)
        job_log(f"Подключение к Exchange как {cfg.email} …")
        client = ExchangeClient(cfg, SESSION["password"])
        if attachments:
            job_log(f"К повторному письму прикреплено вложение: {attachments[0][0]}")
        job_log("Подключение установлено. Досылаю письма в ту же ветку.")

        for i, r in enumerate(targets, start=1):
            body = render_template(template, r.context_contact())
            try:
                html = build_email_html(body, "cid:" + PHOTO_CID)
                new_id, _conv = client.send_followup(
                    r.email, r.subject, body,
                    in_reply_to=r.internet_message_id,
                    html_body=html, inline_image=inline_photo_for_send(),
                    attachments=attachments,
                )
                # Помечаем в состоянии, что повторное письмо ушло (чтобы не выслать снова).
                r.followup_status = "sent"
                r.followup_message_id = new_id
                r.followup_at = _now_iso()
                save_state(state_file(), records)
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
    followed = sum(1 for r in records if r.status == "sent" and r.followup_status == "sent")
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
        "signature": read_signature(),
        "has_photo": photo_path() is not None,
        "photo_name": photo_path().name if photo_path() else "",
        "has_attachment": attachment_path() is not None,
        "attachment_name": attachment_display_name(),
        "has_password": bool(SESSION["password"]),
        "contacts_count": len(SESSION["contacts"]),
        "sent_count": sent,
        "followed_count": followed,
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


@app.post("/api/save_signature")
def api_save_signature():
    data = request.get_json(force=True)
    signature_file().write_text(data.get("signature", ""), encoding="utf-8")
    return jsonify({"ok": True})


ALLOWED_PHOTO_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}


@app.post("/api/upload_photo")
def api_upload_photo():
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_PHOTO_EXT:
        return jsonify({"ok": False, "error": "Допустимы только изображения (png, jpg, gif, bmp)."}), 400
    # удаляем прежнее фото (любого расширения), сохраняем новое
    for old in workdir().glob(PHOTO_STEM + ".*"):
        old.unlink()
    f.save(workdir() / (PHOTO_STEM + ext))
    return jsonify({"ok": True, "name": PHOTO_STEM + ext})


@app.post("/api/delete_photo")
def api_delete_photo():
    for old in workdir().glob(PHOTO_STEM + ".*"):
        old.unlink()
    return jsonify({"ok": True})


@app.get("/api/sigphoto")
def api_sigphoto():
    p = photo_path()
    if not p:
        abort(404)
    return send_file(p)


ALLOWED_ATTACH_EXT = {".pdf"}


@app.post("/api/upload_attachment")
def api_upload_attachment():
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_ATTACH_EXT:
        return jsonify({"ok": False, "error": "Допустим только файл PDF."}), 400
    # удаляем прежнее вложение (любого расширения), сохраняем новое
    for old in workdir().glob(ATTACH_STEM + ".*"):
        old.unlink()
    f.save(workdir() / (ATTACH_STEM + ext))
    conn = read_conn()
    conn["attachment_name"] = os.path.basename(f.filename)
    write_conn(conn)
    return jsonify({"ok": True, "name": os.path.basename(f.filename)})


@app.post("/api/delete_attachment")
def api_delete_attachment():
    for old in workdir().glob(ATTACH_STEM + ".*"):
        old.unlink()
    conn = read_conn()
    conn.pop("attachment_name", None)
    write_conn(conn)
    return jsonify({"ok": True})


@app.post("/api/preview")
def api_preview():
    data = request.get_json(force=True)
    template = data.get("template", "")
    which = data.get("which", "letter")
    signature = data.get("signature")  # текущее содержимое поля подписи
    count = int(data.get("count", 1))

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
        body = render_template(template, c)
        html = build_email_html(body, "/api/sigphoto", signature=signature)
        result.append({"email": c.email, "subject": subject, "html": html})
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
    skip_done = bool(data.get("skip_done", True))
    if start_job(worker_followup, template, delay, limit, only, skip_done):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Уже выполняется другая задача."}), 409


@app.get("/api/status")
def api_status():
    with JOB_LOCK:
        return jsonify(dict(JOB))


# --------------------------------------------------------------------------- #
# Корректное завершение: кнопка «Выход» + авто-остановка при закрытии браузера
# --------------------------------------------------------------------------- #
# Страница раз в несколько секунд шлёт «пинг». Если пингов нет дольше таймаута
# (браузер закрыли) и сейчас не идёт рассылка — приложение само завершается.
HEARTBEAT_TIMEOUT = float(os.environ.get("K2_MAILER_IDLE_TIMEOUT", "15"))
LAST_BEAT = {"t": None}  # None = браузер ещё ни разу не подключался


@app.post("/api/heartbeat")
def api_heartbeat():
    LAST_BEAT["t"] = time.time()
    return jsonify({"ok": True})


def shutdown_now() -> None:
    """Завершает процесс приложения (вместе с локальным сервером)."""
    def _exit():
        time.sleep(0.4)  # дать отдать HTTP-ответ
        os._exit(0)
    threading.Thread(target=_exit, daemon=True).start()


@app.post("/api/quit")
def api_quit():
    shutdown_now()
    return jsonify({"ok": True})


def should_shutdown() -> bool:
    """True, если браузер уже подключался, но давно молчит, и рассылка не идёт."""
    last = LAST_BEAT["t"]
    if last is None:
        return False  # браузер ещё не подключался — не выключаемся
    with JOB_LOCK:
        if JOB["running"]:
            return False  # идёт рассылка — не прерываем
    return (time.time() - last) > HEARTBEAT_TIMEOUT


def watchdog() -> None:
    while True:
        time.sleep(5)
        if should_shutdown():
            shutdown_now()
            return


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
  .mailhtml { background:#fff; border:1px dashed #cbd5e0; border-radius:6px; padding:12px; margin-top:8px; }
  .mailhtml img { max-width:100%; height:auto; }
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
    <div class="hint">Переменные: <code>{имя}</code> и <code>{компания}</code> (можно <code>{name}</code>/<code>{company}</code>).
      Поддерживается HTML-форматирование: <code>&lt;b&gt;жирный&lt;/b&gt;</code>,
      <code>&lt;i&gt;курсив&lt;/i&gt;</code>, <code>&lt;u&gt;подчёркнутый&lt;/u&gt;</code>,
      абзацы <code>&lt;p&gt;…&lt;/p&gt;</code>. Нажмите «Предпросмотр», чтобы увидеть письмо целиком.</div>
    <label>Текст первого письма</label>
    <textarea id="letter"></textarea>
    <label style="margin-top:12px">Текст повторного письма (в ту же ветку)</label>
    <textarea id="followup"></textarea>
    <label style="margin-top:12px">PDF-вложение для повторного письма (например, презентация)</label>
    <div class="btns">
      <input id="attachfile" type="file" accept="application/pdf,.pdf">
      <button id="btnUploadAttach" class="secondary">Прикрепить PDF</button>
      <button id="btnDeleteAttach" class="secondary">Убрать PDF</button>
      <span id="attachPill" class="pill">PDF не прикреплён</span>
    </div>
    <div id="attachMsg" class="msg"></div>
    <div class="btns">
      <button class="secondary" id="btnPreviewLetter">Предпросмотр первого</button>
      <button class="secondary" id="btnPreviewFollow">Предпросмотр повторного</button>
      <button class="secondary" id="btnSaveTpl">Сохранить шаблоны</button>
    </div>
    <div id="previewBox"></div>
  </div>

  <!-- Подпись -->
  <div class="card">
    <h2><span class="step">✍</span> Корпоративная подпись с фото</h2>
    <div class="hint">Вставьте готовый HTML-код подписи. Где должно быть фото — впишите
      <code>{photo}</code> в адрес картинки, например:
      <code>&lt;img src="{photo}" width="120"&gt;</code>. Фото встроится прямо в письмо.</div>
    <label>HTML-код подписи</label>
    <textarea id="signature" placeholder='&lt;table&gt;&lt;tr&gt;
  &lt;td&gt;&lt;img src="{photo}" width="110" style="border-radius:8px"&gt;&lt;/td&gt;
  &lt;td style="padding-left:14px;font-family:Arial"&gt;
    &lt;b&gt;Фёдор Моргунов&lt;/b&gt;&lt;br&gt;Менеджер, K2 SDR&lt;br&gt;
    +7 999 000-00-00 &amp;middot; fmorgunov@k2.cloud
  &lt;/td&gt;
&lt;/tr&gt;&lt;/table&gt;'></textarea>
    <div class="btns">
      <input id="photofile" type="file" accept="image/*">
      <button id="btnUploadPhoto" class="secondary">Загрузить фото</button>
      <button id="btnDeletePhoto" class="secondary">Убрать фото</button>
      <span id="photoPill" class="pill">фото не загружено</span>
      <button id="btnSaveSig" class="secondary">Сохранить подпись</button>
    </div>
    <div id="sigMsg" class="msg"></div>
  </div>

  <!-- 4. Отправка -->
  <div class="card">
    <h2><span class="step">4</span> Отправка</h2>
    <div class="row">
      <div style="max-width:170px"><label>Пауза между письмами, сек</label><input id="delay" type="number" value="1" min="0" step="0.5"></div>
      <div style="max-width:170px"><label>Лимит (0 = все)</label><input id="limit" type="number" value="0" min="0"></div>
    </div>
    <div class="inline"><input id="resume" type="checkbox" checked><label style="margin:0">Пропускать уже отправленные (для рассылки)</label></div>
    <div class="inline"><input id="skipdone" type="checkbox" checked><label style="margin:0">Не досылать повторно тем, кому уже дослали (защита от двойной отправки)</label></div>
    <div class="btns">
      <button id="btnSend">Отправить рассылку</button>
      <button id="btnFollow" class="secondary">Дослать в ту же ветку</button>
    </div>
    <div class="hint">Повторное письмо уходит тем, кому первая рассылка прошла успешно (данные берутся из истории отправки).
      Кому уже досылали — отмечается в истории, поэтому повторный клик не отправит письмо второй раз.</div>
    <div class="bar"><i id="progBar"></i></div>
    <div id="jobStatus" class="msg"></div>
    <div class="log" id="log"></div>
  </div>

  <div class="hint" id="workdir"></div>
  <div class="btns" style="margin-top:8px">
    <button id="btnQuit" class="danger">Завершить работу</button>
    <span class="hint" style="align-self:center">Сервер также сам остановится, если закрыть это окно браузера.</span>
  </div>
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
  $('signature').value = c.signature || '';
  setPhotoPill(c.has_photo, c.photo_name);
  setAttachPill(c.has_attachment, c.attachment_name);
  $('contactsPill').textContent = c.contacts_count ? (c.contacts_count + ' контактов') : 'не загружено';
  $('workdir').textContent = 'Данные и история хранятся в папке: ' + c.workdir;
  if (c.sent_count){
    let s = 'В истории отправки: ' + c.sent_count + ' получателей (доступно «дослать»).';
    if (c.followed_count) s += ' Из них повторное письмо уже получили: ' + c.followed_count + '.';
    setMsg($('jobStatus'), s, true);
  }
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

function setPhotoPill(has, name){
  $('photoPill').textContent = has ? ('фото: ' + (name || 'загружено')) : 'фото не загружено';
}

$('btnSaveSig').onclick = async () => {
  await postJSON('/api/save_signature', {signature: $('signature').value});
  setMsg($('sigMsg'), 'Подпись сохранена.', true);
};

$('btnUploadPhoto').onclick = async () => {
  const f = $('photofile').files[0];
  if (!f){ setMsg($('sigMsg'), 'Выберите файл изображения.', false); return; }
  const fd = new FormData(); fd.append('file', f);
  const r = await api('/api/upload_photo', {method:'POST', body: fd});
  if (!r.ok){ setMsg($('sigMsg'), 'Ошибка: ' + r.error, false); return; }
  setPhotoPill(true, r.name); setMsg($('sigMsg'), 'Фото загружено.', true);
};

$('btnDeletePhoto').onclick = async () => {
  await postJSON('/api/delete_photo', {});
  setPhotoPill(false, ''); setMsg($('sigMsg'), 'Фото убрано.', true);
};

function setAttachPill(has, name){
  $('attachPill').textContent = has ? ('PDF: ' + (name || 'прикреплён')) : 'PDF не прикреплён';
}

$('btnUploadAttach').onclick = async () => {
  const f = $('attachfile').files[0];
  if (!f){ setMsg($('attachMsg'), 'Выберите файл PDF.', false); return; }
  const fd = new FormData(); fd.append('file', f);
  const r = await api('/api/upload_attachment', {method:'POST', body: fd});
  if (!r.ok){ setMsg($('attachMsg'), 'Ошибка: ' + r.error, false); return; }
  setAttachPill(true, r.name); setMsg($('attachMsg'), 'PDF прикреплён к повторному письму.', true);
};

$('btnDeleteAttach').onclick = async () => {
  await postJSON('/api/delete_attachment', {});
  setAttachPill(false, ''); setMsg($('attachMsg'), 'PDF убран.', true);
};

async function preview(which){
  const template = which === 'followup' ? $('followup').value : $('letter').value;
  const r = await postJSON('/api/preview', {template, which, signature: $('signature').value, count: 1});
  if (!r.ok){ $('previewBox').innerHTML = `<div class="msg err">${esc(r.error)}</div>`; return; }
  let h = '';
  for (const it of r.items)
    h += `<div class="preview"><div class="subj">${esc(it.email)} — ${esc(it.subject)}</div>`
       + `<div class="mailhtml">${it.html}</div></div>`;
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
  await postJSON('/api/save_signature', {signature: $('signature').value});
  const r = await postJSON('/api/send', {
    template: $('letter').value, delay: +$('delay').value,
    limit: +$('limit').value, resume: $('resume').checked,
  });
  if (!r.ok){ setMsg($('jobStatus'), 'Ошибка: ' + r.error, false); return; }
  setButtons(true); $('log').textContent=''; startPolling();
};

$('btnFollow').onclick = async () => {
  const skipDone = $('skipdone').checked;
  const warn = skipDone
    ? 'Дослать повторное письмо в ту же ветку? Тем, кому уже досылали, письмо повторно НЕ уйдёт.'
    : 'ВНИМАНИЕ: галочка защиты снята — повторное письмо уйдёт ВСЕМ, включая тех, кому уже досылали. Продолжить?';
  if (!confirm(warn)) return;
  await postJSON('/api/save_settings', collectConn());
  await postJSON('/api/save_signature', {signature: $('signature').value});
  const r = await postJSON('/api/followup', {
    template: $('followup').value, delay: +$('delay').value, limit: +$('limit').value,
    only: '', skip_done: skipDone,
  });
  if (!r.ok){ setMsg($('jobStatus'), 'Ошибка: ' + r.error, false); return; }
  setButtons(true); $('log').textContent=''; startPolling();
};

function esc(s){ return (s||'').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }

// «Пульс»: пока окно открыто, сервер живёт. Закрыли — сам остановится.
function beat(){ fetch('/api/heartbeat', {method:'POST'}).catch(()=>{}); }
beat(); setInterval(beat, 5000);

$('btnQuit').onclick = async () => {
  if (!confirm('Завершить работу приложения? Локальный сервер остановится.')) return;
  try { await fetch('/api/quit', {method:'POST'}); } catch(e){}
  document.body.innerHTML = '<div style="padding:48px;font:16px -apple-system,sans-serif;color:#2d3748">'
    + 'Приложение остановлено. Эту вкладку можно закрыть.</div>';
};

loadConfig();
</script>
</body>
</html>
"""


def pick_free_port(preferred: int) -> int:
    """Возвращает свободный порт, начиная с preferred (вдруг старый экземпляр висит)."""
    for port in [preferred, *range(preferred + 1, preferred + 50)]:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((HOST, port))
                return port
            except OSError:
                continue
    return preferred


def wait_and_open(url: str) -> None:
    """Ждёт, пока сервер начнёт отвечать, затем открывает браузер."""
    for _ in range(50):  # до ~10 секунд
        try:
            urllib.request.urlopen(url, timeout=1).read()
            break
        except Exception:
            time.sleep(0.2)
    try:
        webbrowser.open(url)
    except Exception:
        logging.exception("не удалось открыть браузер")


def show_error_dialog(message: str) -> None:
    """Показывает нативное окно с ошибкой (чтобы сбой не выглядел как «ничего не происходит»)."""
    if sys.platform != "darwin":
        return
    safe = message.replace('"', "'").replace("\\", "/")
    try:
        subprocess.run(
            ["osascript", "-e",
             f'display dialog "{safe}" with title "K2 Mailer" buttons {{"OK"}} with icon stop'],
            check=False,
        )
    except Exception:
        pass


def main() -> None:
    logfile = workdir() / "app.log"
    logging.basicConfig(
        filename=str(logfile), level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    try:
        global PORT
        PORT = pick_free_port(PORT)
        url = f"http://{HOST}:{PORT}"
        logging.info("Запуск %s (рабочая папка %s)", url, workdir())
        print(f"{APP_NAME} запущен. Откройте в браузере: {url}")
        print(f"Рабочая папка: {workdir()}")
        threading.Thread(target=wait_and_open, args=(url,), daemon=True).start()
        threading.Thread(target=watchdog, daemon=True).start()
        app.run(host=HOST, port=PORT, threaded=True)
    except Exception as exc:
        logging.exception("Сбой при запуске")
        show_error_dialog(f"Не удалось запустить K2 Mailer:\n{exc}\n\nПодробности в файле:\n{logfile}")
        raise


if __name__ == "__main__":
    main()
