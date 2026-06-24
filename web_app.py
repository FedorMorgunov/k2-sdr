#!/usr/bin/env python3
"""Веб-интерфейс для рассылки писем из OWA (Exchange/EWS).

Два режима работы одним кодом:

  * Локально (двойной клик по .app/.exe) — открывает страницу в браузере,
    сам выключается при закрытии окна. Удобно одному сотруднику на своей машине.

  * Сервер (хостинг, напр. на VM в K2 Cloud) — несколько сотрудников заходят по
    сети, КАЖДЫЙ под своим почтовым ящиком и паролем. Включается переменной
    окружения K2_MAILER_SERVER=1. В этом режиме сервер не выключается сам и
    слушает все интерфейсы (за обратным прокси/файрволом).

Безопасность многопользовательского режима:
  * «Вход» = подключение к Exchange по логину/паролю сотрудника. Пароль живёт
    ТОЛЬКО в оперативной памяти сервера, привязан к сессии (cookie) и нигде не
    записывается на диск.
  * Данные каждого пользователя (настройки, шаблоны, подпись, история отправки)
    лежат в отдельной папке ~/K2-Mailer/users/<email> — пользователи друг друга
    не видят.
  * Запускать только во внутренней сети и по HTTPS (см. README, раздел про
    хостинг) — пароли ходят по сети.
"""

from __future__ import annotations

import html as htmllib
import json
import logging
import os
import re as _re
import secrets
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from datetime import datetime, timezone
from functools import wraps
from pathlib import Path

from flask import Flask, request, jsonify, send_file, abort, session

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

# --------------------------------------------------------------------------- #
# Режим и сетевые настройки (управляются переменными окружения)
# --------------------------------------------------------------------------- #
def _envbool(name: str, default: bool) -> bool:
    v = os.environ.get(name)
    if v is None:
        return default
    return v.strip().lower() not in ("0", "false", "no", "off", "")


SERVER_MODE = _envbool("K2_MAILER_SERVER", False)
# В серверном режиме по умолчанию слушаем все интерфейсы (за прокси/файрволом),
# локально — только loopback.
HOST = os.environ.get("K2_MAILER_HOST") or ("0.0.0.0" if SERVER_MODE else "127.0.0.1")
PORT = int(os.environ.get("K2_MAILER_PORT", "8765"))

# Серверные значения подключения по умолчанию (Exchange у всех обычно один и тот
# же) — чтобы сотрудники вводили только email и пароль. Можно переопределить в форме.
ENV_EWS_URL = os.environ.get("K2_EWS_URL", "").strip()
ENV_AUTH_TYPE = os.environ.get("K2_AUTH_TYPE", "NTLM").strip() or "NTLM"
ENV_VERIFY_SSL = _envbool("K2_VERIFY_SSL", True)

HEARTBEAT_TIMEOUT = float(os.environ.get("K2_MAILER_IDLE_TIMEOUT", "15"))
# Сколько держать неактивную сессию в памяти, прежде чем выкинуть пароль (сек).
SESSION_TTL = float(os.environ.get("K2_MAILER_SESSION_TTL", str(12 * 3600)))


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

PHOTO_CID = "sigphoto"
PHOTO_STEM = "signature_photo"
ATTACH_STEM = "followup_attachment"
ALLOWED_PHOTO_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}
ALLOWED_ATTACH_EXT = {".pdf"}


# --------------------------------------------------------------------------- #
# Папки данных: общая база + отдельная папка на каждого пользователя
# --------------------------------------------------------------------------- #
def base_dir() -> Path:
    base = Path(os.environ.get("K2_MAILER_HOME") or (Path.home() / "K2-Mailer"))
    base.mkdir(parents=True, exist_ok=True)
    return base


def _safe_email(email: str) -> str:
    s = _re.sub(r"[^a-z0-9._@+-]+", "_", (email or "").strip().lower())
    return s or "default"


def udir(email: str) -> Path:
    d = base_dir() / "users" / _safe_email(email)
    d.mkdir(parents=True, exist_ok=True)
    return d


def state_file(email: str) -> Path:
    return udir(email) / "sent_state.json"


def excel_file(email: str) -> Path:
    return udir(email) / "contacts.xlsx"


def conn_file(email: str) -> Path:
    return udir(email) / "connection.json"


def template_file(email: str, kind: str) -> Path:
    return udir(email) / f"{kind}.txt"


def signature_file(email: str) -> Path:
    return udir(email) / "signature.html"


def photo_path(email: str) -> Path | None:
    for p in udir(email).glob(PHOTO_STEM + ".*"):
        return p
    return None


def attachment_path(email: str) -> Path | None:
    for p in udir(email).glob(ATTACH_STEM + ".*"):
        return p
    return None


def read_conn(email: str) -> dict:
    f = conn_file(email)
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8"))
    return {}


def write_conn(email: str, data: dict) -> None:
    conn_file(email).write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                encoding="utf-8")


def default_template(email: str, kind: str) -> str:
    saved = template_file(email, kind)
    if saved.exists():
        return saved.read_text(encoding="utf-8")
    return DEFAULT_LETTER if kind == "letter" else DEFAULT_FOLLOWUP


def read_signature(email: str) -> str:
    f = signature_file(email)
    return f.read_text(encoding="utf-8") if f.exists() else ""


def signature_active(email: str) -> bool:
    return bool(read_signature(email).strip())


def attachment_display_name(email: str) -> str:
    p = attachment_path(email)
    if not p:
        return ""
    return read_conn(email).get("attachment_name") or p.name


def inline_photo_for_send(email: str):
    sig = read_signature(email)
    p = photo_path(email)
    if p and "{photo}" in sig:
        return (p.name, p.read_bytes(), PHOTO_CID)
    return None


def attachment_for_send(email: str):
    p = attachment_path(email)
    if not p:
        return None
    return [(attachment_display_name(email), p.read_bytes())]


# --------------------------------------------------------------------------- #
# HTML тела письма
# --------------------------------------------------------------------------- #
# Тег должен идти сразу после "<"/"</" и завершаться пробелом, "/" или ">",
# чтобы обычный текст вроде "a < b" не принимался за HTML.
_HTML_TAG_RE = _re.compile(
    r"</?(?:p|br|b|i|u|em|strong|div|span|ul|ol|li|a|table|tr|td|h[1-6]|blockquote)(?:\s|/|>)",
    _re.IGNORECASE,
)


def looks_like_html(text: str) -> bool:
    return bool(_HTML_TAG_RE.search(text or ""))


def text_to_html(text: str) -> str:
    inner = text if looks_like_html(text) else htmllib.escape(text).replace("\n", "<br>")
    return ('<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;'
            'color:#222;line-height:1.5;">' + inner + "</div>")


def build_email_html(email: str, rendered_body: str, photo_ref: str, signature=None) -> str:
    parts = [text_to_html(rendered_body)]
    sig = (read_signature(email) if signature is None else signature).strip()
    if sig:
        parts.append("<br>" + sig.replace("{photo}", photo_ref or ""))
    return ("<html><body style=\"margin:0;padding:0;\">"
            + "".join(parts) + "</body></html>")


def current_config(st: "UserState") -> ExchangeConfig:
    d = read_conn(st.email)
    return ExchangeConfig(
        email=st.email,
        username=(d.get("username") or "").strip() or st.email,
        ews_url=(d.get("ews_url") or "").strip(),
        auth_type=(d.get("auth_type") or ENV_AUTH_TYPE).strip() or "NTLM",
        verify_ssl=bool(d.get("verify_ssl", True)),
    )


def current_cols(email: str) -> ColumnMap:
    c = (read_conn(email).get("columns") or {})
    base = ColumnMap()
    return ColumnMap(
        name=(c.get("name") or base.name),
        company=(c.get("company") or base.company),
        email=(c.get("email") or base.email),
        subject=(c.get("subject") or base.subject),
    )


# --------------------------------------------------------------------------- #
# Сессии пользователей (в памяти). Пароль НЕ пишется на диск.
# --------------------------------------------------------------------------- #
def _empty_job() -> dict:
    return {"running": False, "kind": "", "total": 0, "done": 0,
            "sent": 0, "failed": 0, "log": [], "error": "", "finished": False}


class UserState:
    def __init__(self, email: str):
        self.email = email
        self.password = ""
        self.contacts: list = []
        self.job = _empty_job()
        self.job_lock = threading.Lock()
        self.last_beat = time.time()


USERS: dict[str, UserState] = {}
USERS_LOCK = threading.Lock()


def current_state() -> UserState | None:
    uid = session.get("uid")
    if not uid:
        return None
    with USERS_LOCK:
        return USERS.get(uid)


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if current_state() is None:
            return jsonify({"ok": False, "authed": False,
                            "error": "Сессия не найдена — войдите заново."}), 401
        return fn(*args, **kwargs)
    return wrapper


def job_reset(st: UserState, kind: str) -> None:
    with st.job_lock:
        st.job.update(running=True, kind=kind, total=0, done=0,
                      sent=0, failed=0, log=[], error="", finished=False)


def job_log(st: UserState, line: str) -> None:
    stamp = datetime.now().strftime("%H:%M:%S")
    with st.job_lock:
        st.job["log"].append(f"[{stamp}] {line}")


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# --------------------------------------------------------------------------- #
# Фоновые задачи (работают в контексте конкретного пользователя)
# --------------------------------------------------------------------------- #
def worker_send(st: UserState, template: str, delay: float, limit: int, resume: bool) -> None:
    email = st.email
    try:
        cfg = current_config(st)
        contacts = list(st.contacts)
        if not contacts:
            raise MailerError("Сначала загрузите Excel со списком контактов.")
        if not st.password:
            raise MailerError("Сначала войдите (укажите пароль).")

        records = load_state(state_file(email))
        already = {r.email.lower() for r in records if r.status == "sent"}
        todo = [c for c in contacts if not (resume and c.email.lower() in already)]
        if limit > 0:
            todo = todo[:limit]
        if not todo:
            raise MailerError("Нет писем к отправке (возможно, все уже отправлены).")

        with st.job_lock:
            st.job["total"] = len(todo)
        job_log(st, f"Подключение к Exchange как {cfg.email} …")
        client = ExchangeClient(cfg, st.password)
        job_log(st, "Подключение установлено. Начинаю рассылку.")

        for i, c in enumerate(todo, start=1):
            body = render_template(template, c)
            subject = c.subject or "(без темы)"
            try:
                html = build_email_html(email, body, "cid:" + PHOTO_CID)
                mid, conv = client.send_new(c.email, subject, body,
                                            html_body=html,
                                            inline_image=inline_photo_for_send(email))
                upsert_state(records, SentRecord(
                    email=c.email, name=c.name, company=c.company, subject=subject,
                    status="sent", internet_message_id=mid, conversation_id=conv,
                    sent_at=_now_iso(),
                ))
                with st.job_lock:
                    st.job["sent"] += 1
                job_log(st, f"✓ [{i}/{len(todo)}] {c.email} — отправлено")
            except Exception as exc:
                upsert_state(records, SentRecord(
                    email=c.email, name=c.name, company=c.company, subject=subject,
                    status="failed", error=str(exc), sent_at=_now_iso(),
                ))
                with st.job_lock:
                    st.job["failed"] += 1
                job_log(st, f"✗ [{i}/{len(todo)}] {c.email} — ошибка: {exc}")

            save_state(state_file(email), records)
            with st.job_lock:
                st.job["done"] = i
            if delay and i < len(todo):
                time.sleep(delay)

        job_log(st, f"Готово. Отправлено: {st.job['sent']}, ошибок: {st.job['failed']}.")
    except MailerError as exc:
        with st.job_lock:
            st.job["error"] = str(exc)
        job_log(st, f"Ошибка: {exc}")
    except Exception as exc:
        with st.job_lock:
            st.job["error"] = str(exc)
        job_log(st, f"Непредвиденная ошибка: {exc}")
    finally:
        with st.job_lock:
            st.job["running"] = False
            st.job["finished"] = True


def worker_followup(st: UserState, template: str, delay: float, limit: int,
                    only: str, skip_done: bool = True) -> None:
    email = st.email
    try:
        cfg = current_config(st)
        if not st.password:
            raise MailerError("Сначала войдите (укажите пароль).")

        records = load_state(state_file(email))
        if not records:
            raise MailerError("Нет данных о первой рассылке. Сначала выполните рассылку.")

        targets = [r for r in records if r.status == "sent"]
        if only.strip():
            wanted = {e.strip().lower() for e in only.split(",") if e.strip()}
            targets = [r for r in targets if r.email.lower() in wanted]
        if skip_done:
            before = len(targets)
            targets = [r for r in targets if r.followup_status != "sent"]
            skipped = before - len(targets)
            if skipped:
                job_log(st, f"Пропущено уже досланных (им повторное письмо уже уходило): {skipped}")
        if not targets:
            raise MailerError("Нет адресатов для повторного письма: либо нет успешной "
                              "первой рассылки, либо всем уже дослали. Чтобы отправить "
                              "повторно намеренно — снимите галочку «Не досылать повторно».")
        if limit > 0:
            targets = targets[:limit]

        attachments = attachment_for_send(email)

        with st.job_lock:
            st.job["total"] = len(targets)
        job_log(st, f"Подключение к Exchange как {cfg.email} …")
        client = ExchangeClient(cfg, st.password)
        if attachments:
            job_log(st, f"К повторному письму прикреплено вложение: {attachments[0][0]}")
        job_log(st, "Подключение установлено. Досылаю письма в ту же ветку.")

        for i, r in enumerate(targets, start=1):
            body = render_template(template, r.context_contact())
            try:
                html = build_email_html(email, body, "cid:" + PHOTO_CID)
                new_id, _conv = client.send_followup(
                    r.email, r.subject, body,
                    in_reply_to=r.internet_message_id,
                    html_body=html, inline_image=inline_photo_for_send(email),
                    attachments=attachments,
                )
                r.followup_status = "sent"
                r.followup_message_id = new_id
                r.followup_at = _now_iso()
                save_state(state_file(email), records)
                with st.job_lock:
                    st.job["sent"] += 1
                job_log(st, f"✓ [{i}/{len(targets)}] {r.email} — дослано в ту же ветку")
            except Exception as exc:
                with st.job_lock:
                    st.job["failed"] += 1
                job_log(st, f"✗ [{i}/{len(targets)}] {r.email} — ошибка: {exc}")
            with st.job_lock:
                st.job["done"] = i
            if delay and i < len(targets):
                time.sleep(delay)

        job_log(st, f"Готово. Дослано: {st.job['sent']}, ошибок: {st.job['failed']}.")
    except MailerError as exc:
        with st.job_lock:
            st.job["error"] = str(exc)
        job_log(st, f"Ошибка: {exc}")
    except Exception as exc:
        with st.job_lock:
            st.job["error"] = str(exc)
        job_log(st, f"Непредвиденная ошибка: {exc}")
    finally:
        with st.job_lock:
            st.job["running"] = False
            st.job["finished"] = True


def start_job(st: UserState, target, *args) -> bool:
    with st.job_lock:
        if st.job["running"]:
            return False
    job_reset(st, "followup" if target is worker_followup else "send")
    threading.Thread(target=target, args=(st, *args), daemon=True).start()
    return True


# --------------------------------------------------------------------------- #
# Flask
# --------------------------------------------------------------------------- #
app = Flask(__name__)
# Ключ подписи cookie сессии. Лучше задать постоянным через K2_MAILER_SECRET,
# иначе при перезапуске сервера все сессии (входы) сбросятся.
app.secret_key = os.environ.get("K2_MAILER_SECRET") or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    # Cookie помечается Secure в серверном режиме (предполагается HTTPS за прокси).
    SESSION_COOKIE_SECURE=SERVER_MODE,
)


@app.get("/")
def index():
    return INDEX_HTML


@app.get("/api/me")
def api_me():
    st = current_state()
    return jsonify({
        "authed": st is not None,
        "email": st.email if st else "",
        "server_mode": SERVER_MODE,
    })


@app.get("/api/defaults")
def api_defaults():
    """Значения для предзаполнения формы входа (без пароля)."""
    return jsonify({
        "ews_url": ENV_EWS_URL,
        "auth_type": ENV_AUTH_TYPE,
        "verify_ssl": ENV_VERIFY_SSL,
    })


@app.post("/api/login")
def api_login():
    data = request.get_json(force=True)
    email = (data.get("email") or "").strip()
    password = data.get("password") or ""
    if not email:
        return jsonify({"ok": False, "error": "Укажите email."}), 400
    if not password:
        return jsonify({"ok": False, "error": "Введите пароль."}), 400

    # Сохраняем настройки подключения этого пользователя (без пароля).
    conn = read_conn(email)
    conn["email"] = email
    conn["username"] = (data.get("username") or "").strip()
    conn["ews_url"] = (data.get("ews_url") or ENV_EWS_URL).strip()
    conn["auth_type"] = (data.get("auth_type") or ENV_AUTH_TYPE).strip()
    conn["verify_ssl"] = bool(data.get("verify_ssl", ENV_VERIFY_SSL))
    if data.get("columns"):
        conn["columns"] = data["columns"]
    write_conn(email, conn)

    st = UserState(email)
    st.password = password
    try:
        cfg = current_config(st)
        client = ExchangeClient(cfg, password)
        _ = client.account.inbox.total_count  # лёгкая проверка доступа
    except MailerError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Не удалось войти: {exc}"}), 400

    uid = secrets.token_hex(16)
    with USERS_LOCK:
        USERS[uid] = st
    session["uid"] = uid
    return jsonify({"ok": True, "email": email, "message": f"Вход выполнен: {email}"})


@app.post("/api/logout")
def api_logout():
    uid = session.pop("uid", None)
    if uid:
        with USERS_LOCK:
            USERS.pop(uid, None)  # удаляем пароль из памяти
    return jsonify({"ok": True})


@app.get("/api/config")
@login_required
def api_get_config():
    st = current_state()
    email = st.email
    d = read_conn(email)
    cols = current_cols(email)
    records = load_state(state_file(email))
    sent = sum(1 for r in records if r.status == "sent")
    followed = sum(1 for r in records if r.status == "sent" and r.followup_status == "sent")
    return jsonify({
        "email": email,
        "username": d.get("username", ""),
        "ews_url": d.get("ews_url", ""),
        "auth_type": d.get("auth_type", ENV_AUTH_TYPE),
        "verify_ssl": bool(d.get("verify_ssl", True)),
        "columns": {"name": cols.name, "company": cols.company,
                    "email": cols.email, "subject": cols.subject},
        "letter": default_template(email, "letter"),
        "followup": default_template(email, "followup"),
        "signature": read_signature(email),
        "has_photo": photo_path(email) is not None,
        "photo_name": photo_path(email).name if photo_path(email) else "",
        "has_attachment": attachment_path(email) is not None,
        "attachment_name": attachment_display_name(email),
        "contacts_count": len(st.contacts),
        "sent_count": sent,
        "followed_count": followed,
        "workdir": str(udir(email)),
        "server_mode": SERVER_MODE,
    })


@app.post("/api/save_settings")
@login_required
def api_save_settings():
    st = current_state()
    data = request.get_json(force=True)
    conn = read_conn(st.email)
    for k in ("username", "ews_url", "auth_type"):
        if k in data:
            conn[k] = (data.get(k) or "").strip()
    if "verify_ssl" in data:
        conn["verify_ssl"] = bool(data["verify_ssl"])
    if data.get("columns"):
        conn["columns"] = data["columns"]
    write_conn(st.email, conn)
    return jsonify({"ok": True})


@app.post("/api/upload")
@login_required
def api_upload():
    st = current_state()
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    f = request.files["file"]
    path = excel_file(st.email)
    f.save(path)
    try:
        contacts = read_contacts(path, current_cols(st.email))
    except MailerError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    st.contacts = contacts
    preview = [{"name": c.name, "company": c.company, "email": c.email,
                "subject": c.subject} for c in contacts[:5]]
    return jsonify({"ok": True, "count": len(contacts), "preview": preview})


@app.post("/api/save_templates")
@login_required
def api_save_templates():
    st = current_state()
    data = request.get_json(force=True)
    if "letter" in data:
        template_file(st.email, "letter").write_text(data["letter"], encoding="utf-8")
    if "followup" in data:
        template_file(st.email, "followup").write_text(data["followup"], encoding="utf-8")
    return jsonify({"ok": True})


@app.post("/api/save_signature")
@login_required
def api_save_signature():
    st = current_state()
    data = request.get_json(force=True)
    signature_file(st.email).write_text(data.get("signature", ""), encoding="utf-8")
    return jsonify({"ok": True})


@app.post("/api/upload_photo")
@login_required
def api_upload_photo():
    st = current_state()
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_PHOTO_EXT:
        return jsonify({"ok": False, "error": "Допустимы только изображения (png, jpg, gif, bmp)."}), 400
    for old in udir(st.email).glob(PHOTO_STEM + ".*"):
        old.unlink()
    f.save(udir(st.email) / (PHOTO_STEM + ext))
    return jsonify({"ok": True, "name": PHOTO_STEM + ext})


@app.post("/api/delete_photo")
@login_required
def api_delete_photo():
    st = current_state()
    for old in udir(st.email).glob(PHOTO_STEM + ".*"):
        old.unlink()
    return jsonify({"ok": True})


@app.get("/api/sigphoto")
@login_required
def api_sigphoto():
    st = current_state()
    p = photo_path(st.email)
    if not p:
        abort(404)
    return send_file(p)


@app.post("/api/upload_attachment")
@login_required
def api_upload_attachment():
    st = current_state()
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"ok": False, "error": "Файл не выбран."}), 400
    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_ATTACH_EXT:
        return jsonify({"ok": False, "error": "Допустим только файл PDF."}), 400
    for old in udir(st.email).glob(ATTACH_STEM + ".*"):
        old.unlink()
    f.save(udir(st.email) / (ATTACH_STEM + ext))
    conn = read_conn(st.email)
    conn["attachment_name"] = os.path.basename(f.filename)
    write_conn(st.email, conn)
    return jsonify({"ok": True, "name": os.path.basename(f.filename)})


@app.post("/api/delete_attachment")
@login_required
def api_delete_attachment():
    st = current_state()
    for old in udir(st.email).glob(ATTACH_STEM + ".*"):
        old.unlink()
    conn = read_conn(st.email)
    conn.pop("attachment_name", None)
    write_conn(st.email, conn)
    return jsonify({"ok": True})


@app.post("/api/preview")
@login_required
def api_preview():
    st = current_state()
    email = st.email
    data = request.get_json(force=True)
    template = data.get("template", "")
    which = data.get("which", "letter")
    signature = data.get("signature")
    count = int(data.get("count", 1))

    if which == "followup":
        records = [r for r in load_state(state_file(email)) if r.status == "sent"]
        items = [r.context_contact() for r in records[:count]]
        if not items:
            return jsonify({"ok": False, "error": "Нет данных о первой рассылке для предпросмотра."}), 400
    else:
        items = st.contacts[:count]
        if not items:
            return jsonify({"ok": False, "error": "Сначала загрузите Excel."}), 400

    result = []
    for c in items:
        subject = c.subject or "(без темы)"
        if which == "followup" and not subject.lower().startswith("re:"):
            subject = f"RE: {subject}"
        body = render_template(template, c)
        html = build_email_html(email, body, "/api/sigphoto", signature=signature)
        result.append({"email": c.email, "subject": subject, "html": html})
    return jsonify({"ok": True, "items": result})


@app.post("/api/send")
@login_required
def api_send():
    st = current_state()
    data = request.get_json(force=True)
    template = data.get("template", "")
    template_file(st.email, "letter").write_text(template, encoding="utf-8")
    delay = float(data.get("delay", 1.0))
    limit = int(data.get("limit", 0))
    resume = bool(data.get("resume", False))
    if start_job(st, worker_send, template, delay, limit, resume):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Уже выполняется другая задача."}), 409


@app.post("/api/followup")
@login_required
def api_followup():
    st = current_state()
    data = request.get_json(force=True)
    template = data.get("template", "")
    template_file(st.email, "followup").write_text(template, encoding="utf-8")
    delay = float(data.get("delay", 1.0))
    limit = int(data.get("limit", 0))
    only = data.get("only", "")
    skip_done = bool(data.get("skip_done", True))
    if start_job(st, worker_followup, template, delay, limit, only, skip_done):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Уже выполняется другая задача."}), 409


@app.get("/api/status")
@login_required
def api_status():
    st = current_state()
    with st.job_lock:
        return jsonify(dict(st.job))


@app.post("/api/heartbeat")
def api_heartbeat():
    st = current_state()
    if st:
        st.last_beat = time.time()
    return jsonify({"ok": True})


# --------------------------------------------------------------------------- #
# Завершение/обслуживание
# --------------------------------------------------------------------------- #
def shutdown_now() -> None:
    def _exit():
        time.sleep(0.4)
        os._exit(0)
    threading.Thread(target=_exit, daemon=True).start()


@app.post("/api/quit")
def api_quit():
    # В серверном режиме «выход» = выход из учётной записи, а не остановка сервера.
    if SERVER_MODE:
        return api_logout()
    shutdown_now()
    return jsonify({"ok": True})


def _any_job_running() -> bool:
    with USERS_LOCK:
        return any(s.job["running"] for s in USERS.values())


def _last_beat() -> float | None:
    with USERS_LOCK:
        beats = [s.last_beat for s in USERS.values() if s.last_beat]
    return max(beats) if beats else None


def should_shutdown() -> bool:
    """Авто-остановка (только локальный режим): браузер закрыли и задач нет."""
    if SERVER_MODE:
        return False
    lb = _last_beat()
    if lb is None:
        return False
    if _any_job_running():
        return False
    return (time.time() - lb) > HEARTBEAT_TIMEOUT


def watchdog() -> None:
    while True:
        time.sleep(5)
        if should_shutdown():
            shutdown_now()
            return


def session_reaper() -> None:
    """Серверный режим: выкидываем неактивные сессии (и пароли) из памяти."""
    while True:
        time.sleep(300)
        now = time.time()
        with USERS_LOCK:
            dead = [uid for uid, s in USERS.items()
                    if not s.job["running"] and s.last_beat
                    and (now - s.last_beat) > SESSION_TTL]
            for uid in dead:
                USERS.pop(uid, None)


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
  header { background:var(--pri2); color:#fff; padding:16px 24px; display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:18px; }
  header .sub { opacity:.8; font-size:13px; }
  header .who { font-size:13px; text-align:right; }
  header .who button { margin-top:6px; }
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
  <div>
    <h1>K2 Mailer — рассылка из рабочей почты</h1>
    <div class="sub">Письма уходят из вашего ящика OWA. Пароль хранится только в памяти сервера.</div>
  </div>
  <div class="who" id="whoBox" style="display:none">
    <div>Вы вошли как <b id="whoEmail"></b></div>
    <button class="secondary" id="btnLogout">Выйти</button>
  </div>
</header>
<main>

  <!-- Вход -->
  <div class="card" id="loginCard">
    <h2><span class="step">1</span> Вход в почтовый ящик</h2>
    <div class="hint">Введите данные своей рабочей почты. Это и есть вход: приложение
      подключается к Exchange под вашим логином. Пароль никуда не записывается.</div>
    <div class="row">
      <div><label>Ваш email (отправитель)</label><input id="email" type="text" placeholder="ivanov@k2.cloud"></div>
      <div><label>Логин (часто DOMAIN\\user, можно оставить пустым)</label><input id="username" type="text" placeholder="K2\\ivanov"></div>
    </div>
    <div class="row">
      <div><label>Адрес EWS (можно оставить пустым)</label><input id="ews_url" type="text" placeholder="https://mail.k2.cloud/EWS/Exchange.asmx"></div>
      <div style="max-width:160px"><label>Аутентификация</label>
        <select id="auth_type"><option>NTLM</option><option>basic</option><option value="auto">auto</option></select>
      </div>
    </div>
    <div class="row">
      <div><label>Пароль (нигде не сохраняется)</label><input id="password" type="password" placeholder="••••••••"></div>
    </div>
    <div class="inline"><input id="verify_ssl" type="checkbox" checked><label style="margin:0">Проверять SSL-сертификат (снимите только при самоподписанном)</label></div>
    <div class="btns">
      <button id="btnLogin">Войти</button>
    </div>
    <div id="connMsg" class="msg"></div>
  </div>

  <!-- Рабочая область (доступна после входа) -->
  <div id="appArea" style="display:none">

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
    <span class="hint" id="quitHint" style="align-self:center"></span>
  </div>

  </div><!-- /appArea -->
</main>

<script>
const $ = id => document.getElementById(id);
let SERVER_MODE = false;
async function api(url, opts){ const r = await fetch(url, opts); return await r.json(); }
async function postJSON(url, body){
  return api(url, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)});
}
function setMsg(el, text, ok){ el.textContent = text; el.className = 'msg ' + (ok ? 'ok' : 'err'); }
function esc(s){ return (s||'').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }

function collectLogin(){
  return {
    email: $('email').value, username: $('username').value, ews_url: $('ews_url').value,
    auth_type: $('auth_type').value, verify_ssl: $('verify_ssl').checked,
    password: $('password').value,
  };
}

async function boot(){
  const me = await api('/api/me');
  SERVER_MODE = !!me.server_mode;
  if (me.authed){
    showApp(me.email);
    await loadConfig();
  } else {
    // не вошли — подставим значения по умолчанию для формы входа
    const d = await api('/api/defaults');
    $('ews_url').value = d.ews_url || '';
    if (d.auth_type) $('auth_type').value = d.auth_type;
    $('verify_ssl').checked = d.verify_ssl !== false;
    showLogin();
  }
}

function showLogin(){
  $('appArea').style.display = 'none';
  $('whoBox').style.display = 'none';
  $('loginCard').style.display = '';
}
function showApp(email){
  $('whoEmail').textContent = email;
  $('whoBox').style.display = '';
  $('appArea').style.display = '';
  $('loginCard').style.display = 'none';
  $('btnQuit').textContent = SERVER_MODE ? 'Выйти из учётной записи' : 'Завершить работу';
  $('quitHint').textContent = SERVER_MODE
    ? 'На сервере кнопка завершает только ваш сеанс, остальные продолжают работать.'
    : 'Сервер также сам остановится, если закрыть это окно браузера.';
}

async function loadConfig(){
  const c = await api('/api/config');
  if (c.email === undefined){ showLogin(); return; }
  $('letter').value = c.letter || ''; $('followup').value = c.followup || '';
  $('signature').value = c.signature || '';
  setPhotoPill(c.has_photo, c.photo_name);
  setAttachPill(c.has_attachment, c.attachment_name);
  $('contactsPill').textContent = c.contacts_count ? (c.contacts_count + ' контактов') : 'не загружено';
  $('workdir').textContent = 'Ваши данные и история хранятся в папке: ' + c.workdir;
  if (c.sent_count){
    let s = 'В истории отправки: ' + c.sent_count + ' получателей (доступно «дослать»).';
    if (c.followed_count) s += ' Из них повторное письмо уже получили: ' + c.followed_count + '.';
    setMsg($('jobStatus'), s, true);
  }
}

$('btnLogin').onclick = async () => {
  setMsg($('connMsg'), 'Вхожу…', true);
  const r = await postJSON('/api/login', collectLogin());
  if (r.ok){ $('password').value=''; showApp(r.email); await loadConfig(); }
  else setMsg($('connMsg'), 'Ошибка: ' + r.error, false);
};
$('password').addEventListener('keydown', e => { if (e.key === 'Enter') $('btnLogin').click(); });

$('btnLogout').onclick = async () => {
  await postJSON('/api/logout', {});
  location.reload();
};

$('btnUpload').onclick = async () => {
  const f = $('file').files[0];
  if (!f){ setMsg($('uploadMsg'), 'Выберите файл .xlsx', false); return; }
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

function setPhotoPill(has, name){ $('photoPill').textContent = has ? ('фото: ' + (name || 'загружено')) : 'фото не загружено'; }
function setAttachPill(has, name){ $('attachPill').textContent = has ? ('PDF: ' + (name || 'прикреплён')) : 'PDF не прикреплён'; }

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
    $('log').textContent = (j.log||[]).join('\n');
    $('log').scrollTop = $('log').scrollHeight;
    let s = j.running ? `Выполняется (${j.done}/${j.total})…` : 'Готово.';
    if (j.sent || j.failed) s += `  Успешно: ${j.sent}, ошибок: ${j.failed}.`;
    setMsg($('jobStatus'), j.error ? ('Ошибка: ' + j.error) : s, !j.error);
    if (j.finished && !j.running){ clearInterval(polling); polling = null; setButtons(false); }
  }, 1000);
}
function setButtons(running){ for (const id of ['btnSend','btnFollow']) $(id).disabled = running; }

$('btnSend').onclick = async () => {
  if (!confirm('Отправить рассылку всем загруженным контактам?')) return;
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
  await postJSON('/api/save_signature', {signature: $('signature').value});
  const r = await postJSON('/api/followup', {
    template: $('followup').value, delay: +$('delay').value, limit: +$('limit').value,
    only: '', skip_done: skipDone,
  });
  if (!r.ok){ setMsg($('jobStatus'), 'Ошибка: ' + r.error, false); return; }
  setButtons(true); $('log').textContent=''; startPolling();
};

// «Пульс» (нужен локальному режиму для авто-остановки при закрытии окна).
function beat(){ fetch('/api/heartbeat', {method:'POST'}).catch(()=>{}); }
beat(); setInterval(beat, 5000);

$('btnQuit').onclick = async () => {
  if (SERVER_MODE){
    if (!confirm('Выйти из учётной записи?')) return;
    try { await fetch('/api/quit', {method:'POST'}); } catch(e){}
    location.reload();
    return;
  }
  if (!confirm('Завершить работу приложения? Локальный сервер остановится.')) return;
  try { await fetch('/api/quit', {method:'POST'}); } catch(e){}
  document.body.innerHTML = '<div style="padding:48px;font:16px -apple-system,sans-serif;color:#2d3748">'
    + 'Приложение остановлено. Эту вкладку можно закрыть.</div>';
};

boot();
</script>
</body>
</html>
"""


def pick_free_port(preferred: int) -> int:
    for port in [preferred, *range(preferred + 1, preferred + 50)]:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((HOST, port))
                return port
            except OSError:
                continue
    return preferred


def wait_and_open(url: str) -> None:
    for _ in range(50):
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
    global PORT
    logfile = base_dir() / "app.log"
    logging.basicConfig(
        filename=str(logfile), level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    try:
        if SERVER_MODE:
            # Серверный режим: production-сервер, без браузера и авто-остановки.
            url = f"http://{HOST}:{PORT}"
            logging.info("Запуск в СЕРВЕРНОМ режиме %s (данные в %s)", url, base_dir())
            print(f"{APP_NAME} (серверный режим) слушает {url}")
            if not os.environ.get("K2_MAILER_SECRET"):
                print("ВНИМАНИЕ: K2_MAILER_SECRET не задан — при перезапуске все входы "
                      "сбросятся. Задайте постоянный секрет в окружении.")
            threading.Thread(target=session_reaper, daemon=True).start()
            try:
                from waitress import serve
            except ImportError:
                raise MailerError("Не установлен waitress. Выполните: pip install waitress")
            serve(app, host=HOST, port=PORT, threads=8, ident=APP_NAME)
        else:
            PORT = pick_free_port(PORT)
            url = f"http://{HOST}:{PORT}"
            logging.info("Запуск %s (данные в %s)", url, base_dir())
            print(f"{APP_NAME} запущен. Откройте в браузере: {url}")
            print(f"Папка данных: {base_dir()}")
            threading.Thread(target=wait_and_open, args=(url,), daemon=True).start()
            threading.Thread(target=watchdog, daemon=True).start()
            app.run(host=HOST, port=PORT, threaded=True)
    except Exception as exc:
        logging.exception("Сбой при запуске")
        show_error_dialog(f"Не удалось запустить K2 Mailer:\n{exc}\n\nПодробности в файле:\n{logfile}")
        raise


if __name__ == "__main__":
    main()
