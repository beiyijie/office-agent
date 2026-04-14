from __future__ import annotations

import email
import imaplib
import mimetypes
import smtplib
from email.header import decode_header, make_header
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List

from office_agent.config import settings


def _require_email_config() -> None:
    if not settings.qq_email.configured:
        raise RuntimeError("QQ 邮箱未配置。请在环境变量或 config.json 中设置 QQ_EMAIL/QQ_AUTH_CODE。")


def _decode(value: bytes | str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="ignore")
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


def connect_email() -> imaplib.IMAP4_SSL:
    _require_email_config()
    client = imaplib.IMAP4_SSL(settings.qq_email.imap_host, settings.qq_email.imap_port)
    client.login(settings.qq_email.email, settings.qq_email.auth_code)
    return client


def connect_smtp() -> smtplib.SMTP_SSL:
    _require_email_config()
    client = smtplib.SMTP_SSL(settings.qq_email.smtp_host, settings.qq_email.smtp_port)
    client.login(settings.qq_email.email, settings.qq_email.auth_code)
    return client


def get_email_list(count: int = 10, folder: str = "INBOX") -> List[Dict[str, Any]]:
    with connect_email() as client:
        client.select(folder)
        _, data = client.search(None, "ALL")
        ids = data[0].split()[-count:]
        results: List[Dict[str, Any]] = []
        for msg_id in reversed(ids):
            _, message_data = client.fetch(msg_id, "(RFC822 FLAGS)")
            message = email.message_from_bytes(message_data[0][1])
            results.append(
                {
                    "id": msg_id.decode(),
                    "from": _decode(message.get("From")),
                    "subject": _decode(message.get("Subject")),
                    "date": _decode(message.get("Date")),
                    "read": b"\\Seen" in message_data[1],
                }
            )
        return results


def read_email(message_id: str) -> Dict[str, Any]:
    with connect_email() as client:
        client.select("INBOX")
        _, message_data = client.fetch(message_id.encode(), "(RFC822)")
        message = email.message_from_bytes(message_data[0][1])
        body_parts: List[str] = []
        attachments: List[str] = []
        if message.is_multipart():
            for part in message.walk():
                content_disposition = str(part.get("Content-Disposition", ""))
                if "attachment" in content_disposition.lower():
                    attachments.append(_decode(part.get_filename()))
                    continue
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True) or b""
                    body_parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="ignore"))
        else:
            payload = message.get_payload(decode=True) or b""
            body_parts.append(payload.decode(message.get_content_charset() or "utf-8", errors="ignore"))
        return {
            "from": _decode(message.get("From")),
            "to": _decode(message.get("To")),
            "subject": _decode(message.get("Subject")),
            "date": _decode(message.get("Date")),
            "body": "\n".join(part.strip() for part in body_parts if part.strip()),
            "attachments": attachments,
        }


def search_emails(query: str) -> List[Dict[str, Any]]:
    with connect_email() as client:
        client.select("INBOX")
        _, data = client.search(None, f'TEXT "{query}"')
        ids = data[0].split()
        return [read_email(msg_id.decode()) for msg_id in ids[-10:]]


def send_email(to: str, subject: str, body: str, attachments: List[str] | None = None) -> bool:
    attachments = attachments or []
    message = EmailMessage()
    message["From"] = settings.qq_email.email
    message["To"] = to
    message["Subject"] = subject
    message.set_content(body)
    for file_path in attachments:
        path = Path(file_path).expanduser().resolve()
        mime_type, _ = mimetypes.guess_type(path.name)
        maintype, subtype = (mime_type or "application/octet-stream").split("/", 1)
        with path.open("rb") as file:
            message.add_attachment(file.read(), maintype=maintype, subtype=subtype, filename=path.name)
    with connect_smtp() as client:
        client.send_message(message)
    return True


def mark_as_read(message_id: str) -> bool:
    with connect_email() as client:
        client.select("INBOX")
        client.store(message_id.encode(), "+FLAGS", "\\Seen")
    return True


def delete_email(message_id: str) -> bool:
    with connect_email() as client:
        client.select("INBOX")
        client.store(message_id.encode(), "+FLAGS", "\\Deleted")
        client.expunge()
    return True
