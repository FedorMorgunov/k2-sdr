#!/usr/bin/env python3
"""Рассылка писем из рабочего аккаунта OWA (локальный Exchange, EWS) с поддержкой
ответов в ту же ветку переписки.

Возможности:
  * send     — разослать типовое письмо по списку из Excel, подставив переменные.
  * followup — дослать повторное письмо тем же адресатам, в ту же ветку.

Письма первой рассылки сохраняются в state-файл (sent_state.json), откуда команда
followup берёт адресатов и идентификаторы исходных писем, чтобы ответ попал в тот
же разговор (по теме + заголовкам In-Reply-To/References).

Подключение к Exchange (exchangelib) импортируется лениво — режим --dry-run и чтение
Excel работают даже без доступа к серверу.
"""

from __future__ import annotations

import argparse
import getpass
import json
import os
import sys
import time
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path


class MailerError(Exception):
    """Ожидаемая ошибка (неверные данные/настройки/подключение).

    В CLI перехватывается в main() и печатается пользователю; в веб-интерфейсе
    превращается в понятное сообщение об ошибке вместо падения процесса."""


try:
    import configparser
    from openpyxl import load_workbook
except ImportError as exc:  # pragma: no cover
    raise MailerError(f"Не хватает зависимости: {exc}. Установите: pip install -r requirements.txt")


# --------------------------------------------------------------------------- #
# Конфигурация подключения
# --------------------------------------------------------------------------- #
@dataclass
class ExchangeConfig:
    email: str
    username: str
    ews_url: str
    auth_type: str
    verify_ssl: bool


@dataclass
class ColumnMap:
    name: str = "Имя"
    company: str = "Компания"
    email: str = "Почта"
    subject: str = "Тема письма"


def load_config(path: Path) -> tuple[ExchangeConfig, ColumnMap]:
    if not path.exists():
        raise MailerError(
            f"Не найден файл конфигурации: {path}\n"
            f"Скопируйте пример и заполните: cp config.example.ini config.ini"
        )
    parser = configparser.ConfigParser()
    parser.read(path, encoding="utf-8")

    ex = parser["exchange"]
    email = ex.get("email", "").strip()
    if not email:
        raise MailerError("В config.ini не указан email в секции [exchange].")

    cfg = ExchangeConfig(
        email=email,
        username=ex.get("username", "").strip() or email,
        ews_url=ex.get("ews_url", "").strip(),
        auth_type=ex.get("auth_type", "auto").strip() or "auto",
        verify_ssl=ex.getboolean("verify_ssl", fallback=True),
    )

    cols = ColumnMap()
    if parser.has_section("columns"):
        c = parser["columns"]
        cols = ColumnMap(
            name=c.get("name", cols.name).strip(),
            company=c.get("company", cols.company).strip(),
            email=c.get("email", cols.email).strip(),
            subject=c.get("subject", cols.subject).strip(),
        )
    return cfg, cols


# --------------------------------------------------------------------------- #
# Чтение Excel
# --------------------------------------------------------------------------- #
@dataclass
class Contact:
    name: str
    company: str
    email: str
    subject: str
    row_number: int = 0

    def render_context(self) -> dict:
        """Словарь для подстановки в шаблон. Поддерживаются и русские, и английские
        названия переменных, чтобы шаблон можно было писать как угодно."""
        return {
            "имя": self.name,
            "компания": self.company,
            "почта": self.email,
            "тема": self.subject,
            "name": self.name,
            "company": self.company,
            "email": self.email,
            "subject": self.subject,
        }


def _norm(value) -> str:
    return str(value).strip().lower() if value is not None else ""


def read_contacts(path: Path, cols: ColumnMap) -> list[Contact]:
    if not path.exists():
        raise MailerError(f"Не найден Excel-файл: {path}")

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)

    try:
        header = next(rows)
    except StopIteration:
        raise MailerError(f"Файл {path} пустой.")

    # Заголовок -> индекс колонки (нормализованно).
    header_index = {_norm(h): i for i, h in enumerate(header) if h is not None}

    def find_col(human_name: str) -> int:
        idx = header_index.get(_norm(human_name))
        if idx is None:
            available = ", ".join(str(h) for h in header if h is not None)
            raise MailerError(
                f"В Excel не найдена колонка «{human_name}».\n"
                f"Доступные колонки: {available}\n"
                f"Поправьте заголовки в файле или раздел [columns] в config.ini."
            )
        return idx

    i_name = find_col(cols.name)
    i_company = find_col(cols.company)
    i_email = find_col(cols.email)
    i_subject = find_col(cols.subject)

    contacts: list[Contact] = []
    for row_no, row in enumerate(rows, start=2):  # строка 1 — заголовок
        def cell(i: int) -> str:
            v = row[i] if i < len(row) else None
            return str(v).strip() if v is not None else ""

        email = cell(i_email)
        if not email:
            continue  # пустые строки пропускаем молча
        contacts.append(
            Contact(
                name=cell(i_name),
                company=cell(i_company),
                email=email,
                subject=cell(i_subject),
                row_number=row_no,
            )
        )

    wb.close()
    if not contacts:
        raise MailerError(f"В файле {path} не найдено ни одной строки с адресом почты.")
    return contacts


# --------------------------------------------------------------------------- #
# Шаблон письма
# --------------------------------------------------------------------------- #
class _SafeDict(dict):
    """Не падает на неизвестных плейсхолдерах — оставляет их как есть."""

    def __missing__(self, key):
        return "{" + key + "}"


def render_template(template_text: str, contact: Contact) -> str:
    return template_text.format_map(_SafeDict(contact.render_context()))


def load_template(path: Path) -> str:
    if not path.exists():
        raise MailerError(f"Не найден файл шаблона: {path}")
    return path.read_text(encoding="utf-8")


# --------------------------------------------------------------------------- #
# Состояние рассылки (для повторных писем в ту же ветку)
# --------------------------------------------------------------------------- #
@dataclass
class SentRecord:
    email: str
    name: str
    company: str
    subject: str
    status: str = "sent"                 # sent | failed | dry-run
    error: str = ""
    internet_message_id: str = ""        # для In-Reply-To/References
    conversation_id: str = ""            # для группировки разговора
    sent_at: str = ""
    # Повторное письмо (followup). Заполняется после успешной досылки, чтобы
    # случайно не отправить второе письмо одному адресату несколько раз.
    followup_status: str = ""            # "" (не досылали) | sent | failed
    followup_message_id: str = ""        # msg-id повторного письма (для 3-го касания)
    followup_at: str = ""

    def context_contact(self) -> Contact:
        return Contact(name=self.name, company=self.company,
                       email=self.email, subject=self.subject)


def load_state(path: Path) -> list[SentRecord]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return [SentRecord(**rec) for rec in data.get("entries", [])]


def save_state(path: Path, records: list[SentRecord]) -> None:
    payload = {
        "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "entries": [asdict(r) for r in records],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def upsert_state(records: list[SentRecord], new: SentRecord) -> None:
    """Обновляет запись по email (последняя успешная отправка перетирает старую)."""
    for i, r in enumerate(records):
        if r.email.lower() == new.email.lower():
            records[i] = new
            return
    records.append(new)


# --------------------------------------------------------------------------- #
# Клиент Exchange (EWS)
# --------------------------------------------------------------------------- #
class ExchangeClient:
    def __init__(self, cfg: ExchangeConfig, password: str):
        # Ленивый импорт: нужен только при реальной отправке.
        from exchangelib import Credentials, Account, Configuration, DELEGATE
        from exchangelib import NTLM, BASIC
        from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

        self._ex = sys.modules["exchangelib"]

        if not cfg.verify_ssl:
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

        auth_map = {"ntlm": NTLM, "basic": BASIC}
        auth = auth_map.get(cfg.auth_type.lower())  # None => exchangelib подберёт сам

        creds = Credentials(username=cfg.username, password=password)

        if cfg.ews_url:
            conf_kwargs = dict(service_endpoint=cfg.ews_url, credentials=creds)
            if auth is not None:
                conf_kwargs["auth_type"] = auth
            configuration = Configuration(**conf_kwargs)
            self.account = Account(
                primary_smtp_address=cfg.email,
                config=configuration,
                autodiscover=False,
                access_type=DELEGATE,
            )
        else:
            self.account = Account(
                primary_smtp_address=cfg.email,
                credentials=creds,
                autodiscover=True,
                access_type=DELEGATE,
            )

    def send_new(self, to_email: str, subject: str, body: str,
                 html_body: str = "", inline_image=None,
                 attachments=None) -> tuple[str, str]:
        """Отправляет новое письмо и сохраняет копию в «Отправленные».

        Если задан html_body — письмо уходит как HTML (нужно для подписи с фото).
        inline_image — кортеж (имя_файла, байты, content_id) для встроенной картинки.
        attachments — список (имя_файла, байты) для обычных вложений (например, PDF).
        Возвращает (internet_message_id, conversation_id) для будущих ответов."""
        msg = self._build_message(subject, body, html_body, [to_email],
                                  inline_image, attachments)
        msg.send_and_save()
        return self._read_thread_ids(msg, subject, to_email)

    def send_followup(self, to_email: str, subject: str, body: str,
                      in_reply_to: str, html_body: str = "",
                      inline_image=None, attachments=None) -> tuple[str, str]:
        """Отправляет повторное письмо в ту же ветку.

        Связка ветки достигается двумя способами одновременно:
          1) одинаковая тема разговора (Outlook группирует по теме);
          2) заголовки In-Reply-To / References на исходное письмо (RFC-стандарт).

        Заголовки треда поддерживаются не всеми версиями exchangelib. Поэтому
        и СБОРКА письма с этими заголовками, и его отправка выполняются под
        общим try: если что-то пошло не так (старая версия библиотеки, сервер
        отклонил заголовки) — пересобираем письмо без них и всё равно отправляем.
        Без этого у части коллег повторное письмо «совсем не отправлялось».
        """
        reply_subject = subject if subject.lower().startswith("re:") else f"RE: {subject}"
        headers = {"in_reply_to": in_reply_to, "references": in_reply_to} if in_reply_to else {}

        try:
            msg = self._build_message(reply_subject, body, html_body, [to_email],
                                      inline_image, attachments, **headers)
            msg.send_and_save()
        except Exception:
            if not headers:
                raise  # заголовков и не было — проблема в чём-то ещё, пробрасываем
            # Пересобираем без заголовков треда — письмо всё равно уйдёт, а ветка
            # сложится по одинаковой теме разговора.
            msg = self._build_message(reply_subject, body, html_body, [to_email],
                                      inline_image, attachments)
            msg.send_and_save()
        return self._read_thread_ids(msg, reply_subject, to_email)

    def _build_message(self, subject: str, body: str, html_body: str,
                       recipients: list[str], inline_image=None,
                       attachments=None, **headers):
        """Собирает Message: HTML или обычный текст, при наличии — с inline-картинкой
        и обычными вложениями (attachments: список кортежей (имя_файла, байты))."""
        from exchangelib import Message, Mailbox, HTMLBody, FileAttachment

        msg = Message(
            account=self.account,
            subject=subject,
            body=HTMLBody(html_body) if html_body else body,
            to_recipients=[Mailbox(email_address=e) for e in recipients],
            **headers,
        )
        if inline_image:
            name, content, cid = inline_image
            msg.attach(FileAttachment(name=name, content=content,
                                      is_inline=True, content_id=cid))
        for att in (attachments or []):
            name, content = att
            msg.attach(FileAttachment(name=name, content=content))
        return msg

    @staticmethod
    def _extract(msg) -> tuple[str, str]:
        # В exchangelib поле InternetMessageId называется message_id.
        internet_id = getattr(msg, "message_id", None) or getattr(msg, "internet_message_id", None)
        conv = getattr(msg, "conversation_id", None)
        conv_id = getattr(conv, "id", "") if conv else ""
        return internet_id or "", conv_id or ""

    def _read_thread_ids(self, msg, subject: str, to_email: str) -> tuple[str, str]:
        """Достаёт InternetMessageId и ConversationId из сохранённой копии письма
        (нужны, чтобы повторное письмо ушло в ту же ветку)."""
        internet_id, conv_id = self._extract(msg)
        if not internet_id:
            try:
                msg.refresh()
                internet_id, conv_id = self._extract(msg)
            except Exception:
                pass
        return internet_id, conv_id


# --------------------------------------------------------------------------- #
# Команды
# --------------------------------------------------------------------------- #
def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _connect(cfg: ExchangeConfig) -> ExchangeClient:
    password = os.environ.get("EWS_PASSWORD")
    if not password:
        password = getpass.getpass(f"Пароль для {cfg.username}: ")
    if not password:
        raise MailerError("Пароль не введён — отмена.")
    print(f"Подключение к Exchange как {cfg.email} ...")
    try:
        client = ExchangeClient(cfg, password)
    except Exception as exc:
        raise MailerError(f"Не удалось подключиться к Exchange: {exc}")
    print("Подключение установлено.\n")
    return client


def cmd_send(args) -> None:
    cfg, cols = load_config(Path(args.config))
    contacts = read_contacts(Path(args.excel), cols)
    template = load_template(Path(args.template))
    state_path = Path(args.state)
    records = load_state(state_path)

    already = {r.email.lower() for r in records if r.status == "sent"}
    if args.resume:
        before = len(contacts)
        contacts = [c for c in contacts if c.email.lower() not in already]
        skipped = before - len(contacts)
        if skipped:
            print(f"Пропущено уже отправленных (--resume): {skipped}")

    if args.limit:
        contacts = contacts[: args.limit]

    print(f"К отправке: {len(contacts)} писем. "
          f"{'РЕЖИМ ПРОВЕРКИ (--dry-run), ничего не отправляется.' if args.dry_run else ''}\n")

    client = None if args.dry_run else _connect(cfg)

    sent = failed = 0
    for idx, c in enumerate(contacts, start=1):
        body = render_template(template, c)
        subject = c.subject or "(без темы)"
        prefix = f"[{idx}/{len(contacts)}] {c.email}"

        if args.dry_run:
            print(f"{prefix}  ТЕМА: {subject}")
            print("  --- тело письма ---")
            for line in body.splitlines():
                print(f"  | {line}")
            print("  -------------------")
            upsert_state(records, SentRecord(
                email=c.email, name=c.name, company=c.company,
                subject=subject, status="dry-run", sent_at=_now_iso(),
            ))
            continue

        try:
            internet_id, conv_id = client.send_new(c.email, subject, body)
            upsert_state(records, SentRecord(
                email=c.email, name=c.name, company=c.company, subject=subject,
                status="sent", internet_message_id=internet_id,
                conversation_id=conv_id, sent_at=_now_iso(),
            ))
            sent += 1
            print(f"{prefix}  ✓ отправлено")
        except Exception as exc:
            upsert_state(records, SentRecord(
                email=c.email, name=c.name, company=c.company, subject=subject,
                status="failed", error=str(exc), sent_at=_now_iso(),
            ))
            failed += 1
            print(f"{prefix}  ✗ ошибка: {exc}")

        save_state(state_path, records)  # сохраняем прогресс после каждого письма
        if args.delay and idx < len(contacts):
            time.sleep(args.delay)

    save_state(state_path, records)
    if args.dry_run:
        print(f"\nГотово (проверка). Состояние записано в {state_path}")
    else:
        print(f"\nГотово. Отправлено: {sent}, ошибок: {failed}. Состояние: {state_path}")


def cmd_followup(args) -> None:
    cfg, _cols = load_config(Path(args.config))
    template = load_template(Path(args.template))
    state_path = Path(args.state)
    records = load_state(state_path)

    if not records:
        raise MailerError(f"В {state_path} нет данных о первой рассылке. "
                 f"Сначала выполните команду send.")

    # Кому досылаем: успешно отправленным в первой рассылке.
    targets = [r for r in records if r.status == "sent"]

    if args.only:
        wanted = {e.strip().lower() for e in args.only.split(",") if e.strip()}
        targets = [r for r in targets if r.email.lower() in wanted]

    # Защита от повторной досылки: по умолчанию пропускаем тех, кому повторное
    # письмо уже уходило (см. поле followup_status в sent_state.json). Это и есть
    # «глобальная» защита — она живёт в файле состояния, а не только в рамках
    # текущего запуска. Снять можно флагом --resend.
    if not args.resend:
        before = len(targets)
        targets = [r for r in targets if r.followup_status != "sent"]
        skipped = before - len(targets)
        if skipped:
            print(f"Пропущено уже досланных (повторное письмо им уже уходило): {skipped}")

    if not targets:
        raise MailerError("Нет подходящих адресатов для повторного письма "
                          "(возможно, всем уже дослали; для принудительной "
                          "повторной отправки используйте --resend).")

    if args.limit:
        targets = targets[: args.limit]

    # Опциональное вложение (например, презентация в PDF).
    attachments = None
    if args.attach:
        att_path = Path(args.attach)
        if not att_path.exists():
            raise MailerError(f"Не найден файл вложения: {att_path}")
        attachments = [(att_path.name, att_path.read_bytes())]

    print(f"Повторных писем к отправке: {len(targets)}. "
          f"{'РЕЖИМ ПРОВЕРКИ (--dry-run).' if args.dry_run else ''}\n")

    client = None if args.dry_run else _connect(cfg)

    followup_path = Path(args.followup_state) if args.followup_state else None
    followup_records = load_state(followup_path) if followup_path else []

    sent = failed = 0
    for idx, r in enumerate(targets, start=1):
        body = render_template(template, r.context_contact())
        prefix = f"[{idx}/{len(targets)}] {r.email}"

        if args.dry_run:
            subj = r.subject if r.subject.lower().startswith("re:") else f"RE: {r.subject}"
            att_note = f"  +вложение: {attachments[0][0]}" if attachments else ""
            print(f"{prefix}  ТЕМА: {subj}  (в ветку msg-id={r.internet_message_id or '—'}){att_note}")
            print("  --- тело письма ---")
            for line in body.splitlines():
                print(f"  | {line}")
            print("  -------------------")
            continue

        try:
            new_id, conv_id = client.send_followup(
                r.email, r.subject, body, in_reply_to=r.internet_message_id,
                attachments=attachments,
            )
            sent += 1
            print(f"{prefix}  ✓ дослано в ту же ветку")
            # Отмечаем в основном состоянии, что повторное письмо ушло — чтобы
            # при следующем запуске не выслать его повторно.
            r.followup_status = "sent"
            r.followup_message_id = new_id
            r.followup_at = _now_iso()
            save_state(state_path, records)
            if followup_path is not None:
                upsert_state(followup_records, SentRecord(
                    email=r.email, name=r.name, company=r.company,
                    subject=r.subject, status="sent",
                    internet_message_id=new_id, conversation_id=conv_id,
                    sent_at=_now_iso(),
                ))
                save_state(followup_path, followup_records)
        except Exception as exc:
            failed += 1
            print(f"{prefix}  ✗ ошибка: {exc}")

        if args.delay and idx < len(targets):
            time.sleep(args.delay)

    if args.dry_run:
        print("\nГотово (проверка).")
    else:
        print(f"\nГотово. Дослано: {sent}, ошибок: {failed}.")


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Рассылка писем из OWA (Exchange/EWS) с ответами в ту же ветку.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Примеры:\n"
            "  Проверка без отправки:\n"
            "    python mailer.py send --excel contacts.xlsx --template templates/letter.txt --dry-run\n"
            "  Рассылка:\n"
            "    python mailer.py send --excel contacts.xlsx --template templates/letter.txt\n"
            "  Повторное письмо в ту же ветку:\n"
            "    python mailer.py followup --template templates/followup.txt\n"
        ),
    )
    sub = p.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--config", default="config.ini", help="INI с настройками (по умолчанию config.ini)")
    common.add_argument("--state", default="sent_state.json",
                        help="JSON с состоянием рассылки (по умолчанию sent_state.json)")
    common.add_argument("--template", required=True, help="Текстовый файл-шаблон письма")
    common.add_argument("--delay", type=float, default=1.0,
                        help="Пауза между письмами, сек (по умолчанию 1.0)")
    common.add_argument("--limit", type=int, default=0, help="Ограничить число писем (0 = без лимита)")
    common.add_argument("--dry-run", action="store_true",
                        help="Показать, что будет отправлено, но не отправлять")

    s = sub.add_parser("send", parents=[common], help="Первая рассылка по Excel")
    s.add_argument("--excel", required=True, help="Excel со столбцами Имя, Компания, Почта, Тема письма")
    s.add_argument("--resume", action="store_true",
                   help="Пропустить адреса, которым уже успешно отправлено")
    s.set_defaults(func=cmd_send)

    f = sub.add_parser("followup", parents=[common], help="Повторное письмо в ту же ветку")
    f.add_argument("--only", default="",
                   help="Слать только указанным адресам (через запятую)")
    f.add_argument("--attach", default="",
                   help="Путь к файлу-вложению (например, презентация в PDF)")
    f.add_argument("--resend", action="store_true",
                   help="Отправить повторно даже тем, кому повторное письмо уже уходило")
    f.add_argument("--followup-state", default="",
                   help="Опционально: сохранить msg-id повторных писем в отдельный JSON "
                        "(чтобы потом досылать третье письмо в ту же ветку)")
    f.set_defaults(func=cmd_followup)

    return p


def main(argv=None) -> None:
    args = build_parser().parse_args(argv)
    try:
        args.func(args)
    except MailerError as exc:
        sys.exit(str(exc))


if __name__ == "__main__":
    main()
