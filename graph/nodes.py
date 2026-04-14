from __future__ import annotations

import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict

from office_agent.config import settings
from office_agent.state import AgentState
from office_agent.tools import (
    cancel_meeting,
    create_excel_document,
    create_meeting,
    create_word_document,
    get_meeting_detail,
    get_email_list,
    list_directory,
    read_email,
    read_excel_document,
    read_file,
    read_word_document,
    search_emails,
    send_email,
    write_file,
)
from office_agent.utils import normalize_whitespace, safe_json_dumps, truncate_text

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None


EMAIL_KEYWORDS = (
    "\u90ae\u4ef6",
    "\u90ae\u7bb1",
    "\u53d1\u4fe1",
    "\u53d1\u9001\u90ae\u4ef6",
    "\u6536\u4ef6",
    "smtp",
    "imap",
    "email",
    "mail",
)
DOC_KEYWORDS = ("word", "excel", "\u6587\u6863", "\u5468\u62a5", "\u62a5\u544a", "\u8868\u683c", ".docx", ".xlsx")
FILE_KEYWORDS = ("\u6587\u4ef6", "\u76ee\u5f55", "\u6587\u4ef6\u5939", "\u8bfb\u53d6", "\u5199\u5165", "\u4fdd\u5b58", "ls", "list")
ANALYZE_KEYWORDS = ("\u5206\u6790", "\u89e3\u8bfb", "\u603b\u7ed3", "\u6458\u8981")
SEND_KEYWORDS = ("\u53d1\u7ed9", "\u53d1\u9001", "\u53d1\u9001\u7ed9", "\u8f6c\u53d1", "\u901a\u77e5", "\u90ae\u4ef6\u901a\u77e5")
READ_KEYWORDS = ("\u8bfb\u53d6", "\u67e5\u770b", "\u6253\u5f00", "\u6700\u8fd1", "\u8be6\u60c5")
MEETING_KEYWORDS = ("\u817e\u8baf\u4f1a\u8bae", "\u4f1a\u8bae", "\u5f00\u4f1a", "\u7ea6\u4f1a", "\u4f1a\u8bae\u5ba4", "meeting")
CANCEL_KEYWORDS = ("\u53d6\u6d88", "\u53d6\u6d88\u4f1a\u8bae", "\u53d6\u6d88\u817e\u8baf\u4f1a\u8bae", "\u5220\u9664\u4f1a\u8bae")


def _update_message(state: AgentState, content: str) -> None:
    state["messages"].append({"role": "assistant", "content": content})
    state["response"] = content


def _extract_number(text: str, default: int = 5) -> int:
    match = re.search(r"(\d+)", text)
    return int(match.group(1)) if match else default


def _extract_email_fields(text: str) -> Dict[str, Any]:
    to_match = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", text)
    subject_match = re.search(
        r"\u4e3b\u9898[\u662f\u4e3a:]?\s*(.+?)(?=(?:\s*\u5e76(?:\u90ae\u4ef6)?\u901a\u77e5|\s*\u53d1\u7ed9|\s*@|\n|$))",
        text,
    )
    body_match = re.search(r"(?:\u5185\u5bb9|\u6b63\u6587)[\u662f\u4e3a:]?\s*(.+)", text)
    return {
        "to": to_match.group(1) if to_match else "",
        "subject": subject_match.group(1).strip() if subject_match else "Office Agent \u81ea\u52a8\u90ae\u4ef6",
        "body": body_match.group(1).strip() if body_match else "",
    }


def _extract_path(text: str, suffixes: tuple[str, ...] = ()) -> str | None:
    quoted = re.findall(r"['\"]([^'\"]+)['\"]", text)
    for item in quoted:
        if not suffixes or item.lower().endswith(suffixes):
            return item
    token_matches = re.findall(r"([A-Za-z]:\\[^\s]+|(?:\.{1,2}[\\/][^\s]+)|(?:[^\s]+\.[A-Za-z0-9]{1,8})|(?:[^\s]*[\\/][^\s]+))", text)
    for candidate in token_matches:
        if not suffixes or candidate.lower().endswith(suffixes):
            return candidate
    return None


def _workspace_default_path(filename: str) -> str:
    workspace = settings.workspace_dir or "."
    return str(Path(workspace).expanduser().resolve() / filename)


def _ensure_intermediate(state: AgentState) -> Dict[str, Any]:
    if "intermediate_results" not in state or state["intermediate_results"] is None:
        state["intermediate_results"] = {}
    return state["intermediate_results"]


def _has_any(text: str, keywords: tuple[str, ...]) -> bool:
    return any(keyword in text for keyword in keywords)


def _extract_meeting_topic(text: str) -> str:
    quoted = re.findall(r"['\"]([^'\"]+)['\"]", text)
    if quoted:
        return normalize_whitespace(quoted[0])
    explicit = re.search(
        r"(?:\u4e3b\u9898|\u6807\u9898|topic)[\u662f\u4e3a:]?\s*(.+?)(?=(?:\s*\u5e76(?:\u90ae\u4ef6)?\u901a\u77e5|\s*\u53d1\u7ed9|\s*@|\n|$))",
        text,
        flags=re.IGNORECASE,
    )
    if explicit:
        return normalize_whitespace(explicit.group(1))
    about = re.search(r"\u5173\u4e8e(.+?)(?:\u7684)?(?:\u817e\u8baf\u4f1a\u8bae|\u4f1a\u8bae)", text)
    if about:
        return normalize_whitespace(about.group(1))
    match = re.search(r"(?:\u521b\u5efa|\u5b89\u6392|\u5f00|\u9884\u5b9a)(.+?)(?:\u817e\u8baf\u4f1a\u8bae|\u4f1a\u8bae)", text)
    if match:
        candidate = normalize_whitespace(match.group(1))
        candidate = re.sub(r"[\u7684\s]+$", "", candidate)
        candidate = re.sub(r"(?:\u4eca\u5929|\u660e\u5929|\u540e\u5929|\u4e0a\u5348|\u4e0b\u5348|\u665a\u4e0a|\d+(?::\d+)?\u70b9(?:\u5230|\u81f3|-)?\d*(?::\d+)?\u70b9?)", "", candidate)
        candidate = normalize_whitespace(candidate).strip()
        if candidate:
            return candidate
    return "Office Agent \u4f1a\u8bae"


def _extract_attendees(text: str) -> list[str]:
    return re.findall(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", text)


def _parse_time_expression(text: str) -> tuple[datetime, datetime]:
    now = datetime.now().replace(second=0, microsecond=0)
    base_date = now.date()
    if "\u660e\u5929" in text:
        base_date = (now + timedelta(days=1)).date()
    elif "\u540e\u5929" in text:
        base_date = (now + timedelta(days=2)).date()

    meridiem_offset = 0
    if "\u4e0b\u5348" in text or "\u665a\u4e0a" in text:
        meridiem_offset = 12

    range_match = re.search(r"(\d{1,2})(?::(\d{1,2}))?\u70b9(?:\u5230|\u81f3|-)(\d{1,2})(?::(\d{1,2}))?\u70b9?", text)
    if range_match:
        start_hour = int(range_match.group(1))
        start_minute = int(range_match.group(2) or 0)
        end_hour = int(range_match.group(3))
        end_minute = int(range_match.group(4) or 0)
    else:
        single_match = re.search(r"(\d{1,2})(?::(\d{1,2}))?\u70b9", text)
        if single_match:
            start_hour = int(single_match.group(1))
            start_minute = int(single_match.group(2) or 0)
        else:
            start_hour = 15
            start_minute = 0
        end_hour = start_hour + 1
        end_minute = start_minute

    if meridiem_offset and start_hour < 12:
        start_hour += meridiem_offset
    if meridiem_offset and end_hour < 12:
        end_hour += meridiem_offset

    start_time = datetime.combine(base_date, datetime.min.time()).replace(hour=start_hour, minute=start_minute)
    end_time = datetime.combine(base_date, datetime.min.time()).replace(hour=end_hour, minute=end_minute)
    if end_time <= start_time:
        end_time = start_time + timedelta(hours=1)
    return start_time, end_time


def _build_meeting_email_body(meeting_info: Dict[str, Any]) -> str:
    return (
        "\u60a8\u597d\uff0c\n\n"
        "\u5df2\u4e3a\u60a8\u521b\u5efa\u817e\u8baf\u4f1a\u8bae\uff0c\u4fe1\u606f\u5982\u4e0b\uff1a\n\n"
        f"\u4f1a\u8bae\u4e3b\u9898\uff1a{meeting_info.get('subject', '')}\n"
        f"\u5f00\u59cb\u65f6\u95f4\uff1a{meeting_info.get('start_time', '')}\n"
        f"\u7ed3\u675f\u65f6\u95f4\uff1a{meeting_info.get('end_time', '')}\n"
        f"\u4f1a\u8bae\u53f7\uff1a{meeting_info.get('meeting_code', '')}\n"
        f"\u5165\u4f1a\u94fe\u63a5\uff1a{meeting_info.get('join_url', '')}\n\n"
        "\u8bf7\u6309\u65f6\u53c2\u4f1a\u3002\n"
    )


def _build_cancel_meeting_email_body(meeting_info: Dict[str, Any]) -> str:
    return (
        "\u60a8\u597d\uff0c\n\n"
        "\u817e\u8baf\u4f1a\u8bae\u5df2\u53d6\u6d88\uff0c\u4fe1\u606f\u5982\u4e0b\uff1a\n\n"
        f"\u4f1a\u8bae\u4e3b\u9898\uff1a{meeting_info.get('subject', '')}\n"
        f"\u539f\u5b9a\u5f00\u59cb\u65f6\u95f4\uff1a{meeting_info.get('start_time', '')}\n"
        f"\u539f\u5b9a\u7ed3\u675f\u65f6\u95f4\uff1a{meeting_info.get('end_time', '')}\n"
        f"\u4f1a\u8bae\u53f7\uff1a{meeting_info.get('meeting_code', '')}\n\n"
        "\u8bf7\u4ee5\u6700\u65b0\u5b89\u6392\u4e3a\u51c6\u3002\n"
    )


def _extract_meeting_identifier(text: str) -> str:
    meeting_id_match = re.search(r"(?:meeting[_\s-]?id|\u4f1a\u8baeID|\u4f1a\u8bae id)[\u662f\u4e3a:： ]*([A-Za-z0-9-]+)", text, flags=re.IGNORECASE)
    if meeting_id_match:
        return meeting_id_match.group(1)
    code_match = re.search(r"(?:\u4f1a\u8bae\u53f7|\u4f1a\u8bae\u7801)[\u662f\u4e3a:： ]*(\d+)", text)
    if code_match:
        return code_match.group(1)
    mock_match = re.search(r"(mock-\d+)", text)
    if mock_match:
        return mock_match.group(1)
    return ""


def _meeting_info_for_cancel(task: str) -> Dict[str, Any]:
    meeting_id = _extract_meeting_identifier(task)
    if meeting_id:
        try:
            detail = get_meeting_detail(meeting_id)
            detail["meeting_id"] = meeting_id
            return detail
        except Exception:
            pass

    subject = _extract_meeting_topic(task)
    start_time, end_time = _parse_time_expression(task)
    return {
        "meeting_id": meeting_id or f"mock-{start_time.strftime('%Y%m%d%H%M%S')}",
        "meeting_code": f"8{start_time.strftime('%m%d%H%M')}",
        "subject": subject,
        "start_time": start_time.isoformat(),
        "end_time": end_time.isoformat(),
        "join_url": "",
        "provider": "tencent_meeting",
        "status": "fetched",
    }


def _plan_actions(task: str) -> list[str]:
    message = normalize_whitespace(task).lower()
    has_email = _has_any(message, EMAIL_KEYWORDS)
    has_doc = _has_any(message, DOC_KEYWORDS)
    has_file = _has_any(message, FILE_KEYWORDS)
    has_analyze = _has_any(message, ANALYZE_KEYWORDS)
    has_send = _has_any(message, SEND_KEYWORDS) or "@" in message
    has_read = _has_any(message, READ_KEYWORDS)
    has_forward = "\u8f6c\u53d1" in message
    has_meeting = _has_any(message, MEETING_KEYWORDS)
    has_cancel = _has_any(message, CANCEL_KEYWORDS)

    if has_meeting and has_cancel and has_send and has_email:
        return ["meeting_node", "email_node"]
    if has_meeting and has_cancel:
        return ["meeting_node"]
    if has_meeting and has_send and has_email:
        return ["meeting_node", "email_node"]
    if has_meeting:
        return ["meeting_node"]
    if has_analyze and has_doc:
        return ["document_node", "general_node"]
    if has_analyze and has_email:
        return ["email_node", "general_node"]
    if ("\u603b\u7ed3" in message or "\u6458\u8981" in message) and has_email and has_doc:
        return ["email_node", "document_node"]
    if has_read and has_doc and has_send:
        return ["document_node", "email_node"]
    if has_forward and has_email:
        return ["email_node", "email_node"]
    if has_email:
        return ["email_node"]
    if has_doc:
        return ["document_node"]
    if has_file:
        return ["file_ops_node"]
    return ["general_node"]


def _prepare_general_prompt(task: str, intermediate: Dict[str, Any]) -> str:
    content = truncate_text(str(intermediate.get("content", "")), 8000)
    source = intermediate.get("source", "\u672a\u77e5\u6765\u6e90")
    path = intermediate.get("path", "")
    summary = intermediate.get("summary", "")
    extra = f"\n\u6587\u4ef6\u8def\u5f84: {path}" if path else ""
    summary_text = f"\n\u5df2\u6709\u6458\u8981: {summary}" if summary else ""
    return (
        f"\u7528\u6237\u4efb\u52a1: {task}\n"
        f"\u8bf7\u57fa\u4e8e\u4e0b\u9762\u7684{source}\u5185\u5bb9\u7ed9\u51fa\u4e2d\u6587\u5206\u6790\u3002"
        f"{extra}{summary_text}\n\n"
        f"\u5185\u5bb9:\n{content}"
    )


def _format_email_digest(emails: list[Dict[str, Any]], details: list[Dict[str, Any]]) -> tuple[str, str]:
    blocks = []
    brief_lines = []
    for index, (email_meta, detail) in enumerate(zip(emails, details), 1):
        sender = email_meta.get("from") or detail.get("from") or "\u672a\u77e5\u53d1\u4ef6\u4eba"
        subject = email_meta.get("subject") or detail.get("subject") or "\u65e0\u4e3b\u9898"
        body = normalize_whitespace(detail.get("body", ""))
        body_preview = truncate_text(body, 300)
        brief_lines.append(f"{index}. {sender} | {subject}")
        blocks.append(f"\u90ae\u4ef6{index}\n\u53d1\u4ef6\u4eba: {sender}\n\u4e3b\u9898: {subject}\n\u6b63\u6587: {body_preview}")
    return "\n".join(brief_lines), "\n\n".join(blocks)


def router_node(state: AgentState) -> AgentState:
    state["current_node"] = "router"
    state["error"] = None

    pending = state.get("pending_actions", [])
    if pending:
        next_action = pending[0]
        state["pending_actions"] = pending[1:]
        state["next_node"] = next_action
        return state

    if state.get("completed_actions"):
        state["next_node"] = "END"
        return state

    actions = _plan_actions(state["task"])
    state["pending_actions"] = actions[1:]
    state["completed_actions"] = []
    state["next_node"] = actions[0] if actions else "END"
    return state


def meeting_node(state: AgentState) -> AgentState:
    state["current_node"] = "meeting_node"
    task = state["task"]
    intermediate = _ensure_intermediate(state)

    try:
        is_cancel = _has_any(task, CANCEL_KEYWORDS)
        email_fields = _extract_email_fields(task)
        if is_cancel:
            meeting_info = _meeting_info_for_cancel(task)
            cancel_meeting(str(meeting_info.get("meeting_id", "")))
            meeting_info["status"] = "canceled"
            intermediate.update(
                {
                    "content": _build_cancel_meeting_email_body(meeting_info),
                    "source": "meeting",
                    "summary": meeting_info.get("subject", ""),
                    "meeting": meeting_info,
                    "time_range": {
                        "start_time": meeting_info.get("start_time"),
                        "end_time": meeting_info.get("end_time"),
                    },
                }
            )
            if email_fields["to"]:
                intermediate["recipients"] = [email_fields["to"]]
            state["tools_results"]["meeting"] = meeting_info
            _update_message(
                state,
                (
                    f"\u817e\u8baf\u4f1a\u8bae\u5df2\u53d6\u6d88\uff1a{meeting_info.get('subject', '')}\n"
                    f"\u4f1a\u8bae\u53f7\uff1a{meeting_info.get('meeting_code', '')}\n"
                    f"\u65f6\u95f4\uff1a{meeting_info.get('start_time', '')} - {meeting_info.get('end_time', '')}"
                ),
            )
            return state

        subject = _extract_meeting_topic(task)
        start_time, end_time = _parse_time_expression(task)
        attendees = _extract_attendees(task)
        meeting_info = create_meeting(subject=subject, start_time=start_time, end_time=end_time, attendees=attendees)

        meeting_body = _build_meeting_email_body(meeting_info)
        intermediate.update(
            {
                "content": meeting_body,
                "source": "meeting",
                "summary": meeting_info.get("subject", ""),
                "meeting": meeting_info,
                "time_range": {
                    "start_time": meeting_info.get("start_time"),
                    "end_time": meeting_info.get("end_time"),
                },
                "recipients": attendees,
            }
        )
        if email_fields["to"]:
            intermediate["recipients"] = [email_fields["to"]]

        state["tools_results"]["meeting"] = meeting_info
        _update_message(
            state,
            (
                f"\u817e\u8baf\u4f1a\u8bae\u5df2\u521b\u5efa\uff1a{meeting_info.get('subject', '')}\n"
                f"\u65f6\u95f4\uff1a{meeting_info.get('start_time', '')} - {meeting_info.get('end_time', '')}\n"
                f"\u4f1a\u8bae\u53f7\uff1a{meeting_info.get('meeting_code', '')}\n"
                f"\u5165\u4f1a\u94fe\u63a5\uff1a{meeting_info.get('join_url', '')}"
            ),
        )
    except Exception as exc:
        state["error"] = str(exc)
        _update_message(state, f"\u4f1a\u8bae\u64cd\u4f5c\u5931\u8d25\uff1a{exc}")
    return state


def email_node(state: AgentState) -> AgentState:
    state["current_node"] = "email_node"
    task = state["task"]
    normalized = normalize_whitespace(task).lower()
    intermediate = _ensure_intermediate(state)

    try:
        email_fields = _extract_email_fields(task)
        wants_forward = "\u8f6c\u53d1" in normalized
        wants_analyze = any(keyword in normalized for keyword in ("\u5206\u6790", "\u89e3\u8bfb"))
        wants_summary_to_doc = any(keyword in normalized for keyword in ("\u603b\u7ed3", "\u6458\u8981")) and any(
            keyword in normalized for keyword in ("\u5468\u62a5", "\u6587\u6863", "word", ".docx")
        )
        wants_recent = "\u6700\u8fd1" in normalized or "\u67e5\u770b" in normalized or "\u8bfb\u53d6\u90ae\u4ef6" in normalized

        if email_fields["to"] and intermediate.get("content"):
            body = email_fields["body"] or str(intermediate.get("content") or "")
            subject = email_fields["subject"]
            meeting_info = intermediate.get("meeting") or {}
            if meeting_info and subject == "Office Agent \u81ea\u52a8\u90ae\u4ef6":
                default_subject = "\u817e\u8baf\u4f1a\u8bae"
                prefix = "\u4f1a\u8bae\u53d6\u6d88\u901a\u77e5" if meeting_info.get("status") == "canceled" else "\u4f1a\u8bae\u901a\u77e5"
                subject = f"{prefix}\uff1a{meeting_info.get('subject', default_subject)}"

            attachments = []
            if intermediate.get("path") and Path(str(intermediate["path"])).suffix.lower() in {".docx", ".xlsx"}:
                attachments.append(str(intermediate["path"]))

            send_email(to=email_fields["to"], subject=subject, body=truncate_text(body, 8000), attachments=attachments)
            _update_message(state, f"\u90ae\u4ef6\u5df2\u53d1\u9001\u81f3 {email_fields['to']}\uff0c\u4e3b\u9898\uff1a{subject}")
            return state

        if wants_summary_to_doc:
            count = _extract_number(task, default=3)
            emails = get_email_list(count=count)
            details = [read_email(item["id"]) for item in emails if item.get("id")]
            brief, digest = _format_email_digest(emails[: len(details)], details)
            if not digest:
                _update_message(state, "\u6ca1\u6709\u53ef\u4f9b\u603b\u7ed3\u7684\u90ae\u4ef6\u5185\u5bb9\u3002")
                return state
            intermediate.update({"content": digest, "source": "email", "summary": brief})
            state["tools_results"]["email_summaries"] = digest
            return state

        if wants_forward:
            emails = get_email_list(count=1)
            if not emails:
                _update_message(state, "\u6ca1\u6709\u627e\u5230\u53ef\u8f6c\u53d1\u7684\u90ae\u4ef6\u3002")
                return state
            detail = read_email(emails[0]["id"])
            forward_body = (
                "\u8f6c\u53d1\u90ae\u4ef6\n"
                f"\u53d1\u4ef6\u4eba: {detail.get('from', '')}\n"
                f"\u4e3b\u9898: {detail.get('subject', '')}\n"
                f"\u65f6\u95f4: {detail.get('date', '')}\n\n"
                f"{detail.get('body', '')}"
            )
            intermediate.update({"content": forward_body, "source": "email", "summary": detail.get("subject", "")})
            state["tools_results"]["email_detail"] = detail
            return state

        if wants_analyze:
            count = _extract_number(task, default=3)
            emails = get_email_list(count=count)
            details = [read_email(item["id"]) for item in emails if item.get("id")]
            _, digest = _format_email_digest(emails[: len(details)], details)
            if not digest:
                _update_message(state, "\u6ca1\u6709\u53ef\u5206\u6790\u7684\u90ae\u4ef6\u5185\u5bb9\u3002")
                return state
            intermediate.update({"content": digest, "source": "email", "summary": f"\u6700\u8fd1 {len(details)} \u5c01\u90ae\u4ef6"})
            state["tools_results"]["email_analysis_source"] = digest
            if "general_node" not in state.get("pending_actions", []) and "general_node" not in state.get("completed_actions", []):
                state["pending_actions"] = ["general_node", *state.get("pending_actions", [])]
            return state

        if "\u641c\u7d22" in task:
            query = task.replace("\u641c\u7d22", "").replace("\u90ae\u4ef6", "").strip()
            emails = search_emails(query)
            state["tools_results"]["emails"] = emails
            _update_message(state, safe_json_dumps(emails))
            return state

        if wants_recent:
            count = _extract_number(task, default=5)
            emails = get_email_list(count=count)
            state["tools_results"]["emails"] = emails
            if "\u8be6\u60c5" in task and emails:
                detail = read_email(emails[0]["id"])
                _update_message(state, safe_json_dumps(detail))
                return state
            lines = [f"{idx}. {item['from']} | {item['subject']} | {item['date']}" for idx, item in enumerate(emails, 1)]
            _update_message(state, "\u6700\u8fd1\u90ae\u4ef6\uff1a\n" + "\n".join(lines) if lines else "\u6ca1\u6709\u8bfb\u53d6\u5230\u90ae\u4ef6\u3002")
            return state

        if email_fields["to"]:
            send_email(to=email_fields["to"], subject=email_fields["subject"], body=email_fields["body"] or task)
            _update_message(state, f"\u90ae\u4ef6\u5df2\u53d1\u9001\u81f3 {email_fields['to']}\uff0c\u4e3b\u9898\uff1a{email_fields['subject']}")
            return state

        _update_message(state, "\u5df2\u8fdb\u5165\u90ae\u4ef6\u8282\u70b9\uff0c\u4f46\u6ca1\u6709\u8bc6\u522b\u51fa\u5177\u4f53\u52a8\u4f5c\u3002")
    except Exception as exc:
        state["error"] = str(exc)
        _update_message(state, f"\u90ae\u4ef6\u64cd\u4f5c\u5931\u8d25\uff1a{exc}")
    return state


def document_node(state: AgentState) -> AgentState:
    state["current_node"] = "document_node"
    task = state["task"]
    lower_task = normalize_whitespace(task).lower()
    intermediate = _ensure_intermediate(state)

    try:
        has_intermediate_content = bool(intermediate.get("content"))
        wants_analyze = any(keyword in lower_task for keyword in ("\u5206\u6790", "\u89e3\u8bfb"))
        wants_send = any(keyword in lower_task for keyword in ("\u53d1\u7ed9", "\u53d1\u9001", "\u53d1\u9001\u7ed9")) or "@" in lower_task
        need_create_doc = any(keyword in lower_task for keyword in ("\u751f\u6210", "\u521b\u5efa", "\u5236\u4f5c", "\u5199", "\u8f93\u51fa"))
        is_word_task = any(keyword in lower_task for keyword in (".docx", "word", "\u6587\u6863", "\u5468\u62a5", "\u62a5\u544a", "\u901a\u77e5"))
        is_excel_task = any(keyword in lower_task for keyword in (".xlsx", "excel", "\u8868\u683c"))

        if has_intermediate_content and intermediate.get("source") == "email" and is_word_task:
            path = _extract_path(task, (".docx",)) or _workspace_default_path("weekly_report.docx")
            title = "\u90ae\u4ef6\u6458\u8981\u5468\u62a5" if "\u5468\u62a5" in task else "\u90ae\u4ef6\u6574\u7406\u6587\u6863"
            content = str(intermediate.get("summary") or "") + "\n\n" + str(intermediate.get("content") or "")
            create_word_document(content=content.strip(), path=path, title=title)
            intermediate["path"] = path
            intermediate["summary"] = title
            state["tools_results"]["created_doc_path"] = path
            _update_message(state, f"\u6587\u6863\u5df2\u751f\u6210\uff1a{path}")
            return state

        if (("\u8bfb\u53d6" in task) or wants_analyze or wants_send) and (".docx" in lower_task or "word" in lower_task or "\u6587\u6863" in task):
            path = _extract_path(task, (".docx",))
            if not path:
                raise RuntimeError("\u672a\u627e\u5230\u8981\u8bfb\u53d6\u7684 Word \u6587\u4ef6\u8def\u5f84\u3002")
            content = read_word_document(path)
            state["tools_results"]["document"] = content
            intermediate.update({"content": content, "source": "document", "path": path})
            if wants_analyze:
                if "general_node" not in state.get("pending_actions", []) and "general_node" not in state.get("completed_actions", []):
                    state["pending_actions"] = ["general_node", *state.get("pending_actions", [])]
                return state
            if wants_send and "email_node" not in state.get("pending_actions", []) and "email_node" not in state.get("completed_actions", []):
                state["pending_actions"] = ["email_node", *state.get("pending_actions", [])]
                return state
            _update_message(state, truncate_text(content, 1200))
            return state

        if need_create_doc and is_word_task:
            path = _extract_path(task, (".docx",)) or _workspace_default_path("generated_document.docx")
            title = "\u901a\u77e5" if "\u901a\u77e5" in task else "Office Agent \u6587\u6863"
            if "\u901a\u77e5" in task:
                content = _generate_notification_content(task)
            elif has_intermediate_content:
                content = str(intermediate.get("content"))
            else:
                content = _generate_report_content(task) if "\u5468\u62a5" in task else task
            create_word_document(content=content, path=path, title=title)
            intermediate.update({"path": path, "source": intermediate.get("source", "document"), "summary": title})
            state["tools_results"]["created_doc_path"] = path
            _update_message(state, f"Word \u6587\u6863\u5df2\u751f\u6210\uff1a{path}")
            return state

        if need_create_doc and is_excel_task:
            path = _extract_path(task, (".xlsx",)) or _workspace_default_path("output.xlsx")
            create_excel_document([["\u793a\u4f8b", "\u5df2\u751f\u6210"]], path=path, headers=["\u9879\u76ee", "\u72b6\u6001"])
            intermediate.update({"path": path, "source": "document", "summary": "Excel \u6587\u6863"})
            _update_message(state, f"Excel \u6587\u6863\u5df2\u751f\u6210\uff1a{path}")
            return state

        if (("\u8bfb\u53d6" in task) or wants_analyze) and (".xlsx" in lower_task or "excel" in lower_task or "\u8868\u683c" in task):
            path = _extract_path(task, (".xlsx",))
            if not path:
                raise RuntimeError("\u672a\u627e\u5230\u8981\u8bfb\u53d6\u7684 Excel \u6587\u4ef6\u8def\u5f84\u3002")
            content = read_excel_document(path)
            state["tools_results"]["excel"] = content
            intermediate.update({"content": safe_json_dumps(content), "source": "document", "path": path})
            if wants_analyze and "general_node" not in state.get("pending_actions", []) and "general_node" not in state.get("completed_actions", []):
                state["pending_actions"] = ["general_node", *state.get("pending_actions", [])]
                return state
            _update_message(state, safe_json_dumps(content))
            return state

        _update_message(state, "\u5df2\u8fdb\u5165\u6587\u6863\u8282\u70b9\uff0c\u4f46\u6ca1\u6709\u8bc6\u522b\u51fa\u5177\u4f53\u52a8\u4f5c\u3002")
    except Exception as exc:
        state["error"] = str(exc)
        _update_message(state, f"\u6587\u6863\u64cd\u4f5c\u5931\u8d25\uff1a{exc}")
    return state


def file_ops_node(state: AgentState) -> AgentState:
    state["current_node"] = "file_ops_node"
    task = state["task"]
    try:
        if any(keyword in task for keyword in ("\u76ee\u5f55", "\u5217\u51fa", "list", "ls")):
            path = _extract_path(task) or "."
            files = list_directory(path)
            state["tools_results"]["files"] = files
            _update_message(state, "\n".join(files) if files else "\u76ee\u5f55\u4e3a\u7a7a\u3002")
            return state
        if "\u8bfb\u53d6" in task or "\u6253\u5f00\u6587\u4ef6" in task:
            path = _extract_path(task)
            if not path:
                raise RuntimeError("\u672a\u627e\u5230\u8981\u8bfb\u53d6\u7684\u6587\u4ef6\u8def\u5f84\u3002")
            content = read_file(path)
            state["tools_results"]["file_content"] = content
            _update_message(state, truncate_text(content, 1200))
            return state
        if any(keyword in task for keyword in ("\u5199\u5165", "\u4fdd\u5b58", "\u521b\u5efa\u6587\u4ef6")):
            path = _extract_path(task) or "note.txt"
            content_match = re.search(r"(?:\u5185\u5bb9|\u5199\u5165)[\u662f\u4e3a:]?\s*(.+)", task)
            content = content_match.group(1).strip() if content_match else task
            write_file(path, content)
            _update_message(state, f"\u6587\u4ef6\u5df2\u5199\u5165\uff1a{path}")
            return state
        _update_message(state, "\u5df2\u8fdb\u5165\u6587\u4ef6\u8282\u70b9\uff0c\u4f46\u6ca1\u6709\u8bc6\u522b\u51fa\u5177\u4f53\u52a8\u4f5c\u3002")
    except Exception as exc:
        state["error"] = str(exc)
        _update_message(state, f"\u6587\u4ef6\u64cd\u4f5c\u5931\u8d25\uff1a{exc}")
    return state


def general_node(state: AgentState) -> AgentState:
    state["current_node"] = "general_node"
    try:
        intermediate = state.get("intermediate_results", {})
        if intermediate.get("content"):
            prompt = _prepare_general_prompt(state["task"], intermediate)
            response = _call_minimax_or_fallback(prompt, [{"role": "user", "content": prompt}], intermediate=intermediate)
        else:
            response = _call_minimax_or_fallback(state["task"], state["messages"], intermediate=intermediate)
        _update_message(state, response)
    except Exception as exc:
        state["error"] = str(exc)
        _update_message(state, f"\u901a\u7528\u5904\u7406\u5931\u8d25\uff1a{exc}")
    return state


def _generate_report_content(task: str) -> str:
    return (
        "\u672c\u5468\u5de5\u4f5c\u6982\u89c8\n"
        "1. \u5df2\u5b8c\u6210\u4e8b\u9879\uff1a\n"
        "- \u6839\u636e\u5f53\u524d\u6307\u4ee4\u751f\u6210\u529e\u516c\u5468\u62a5\u8349\u7a3f\u3002\n"
        "- \u6574\u7406\u5f85\u529e\u4e8b\u9879\u4e0e\u8f93\u51fa\u7269\u3002\n\n"
        "2. \u98ce\u9669\u4e0e\u95ee\u9898\uff1a\n"
        "- \u5982\u9700\u81ea\u52a8\u586b\u5145\u771f\u5b9e\u4e1a\u52a1\u6570\u636e\uff0c\u8bf7\u8865\u5145\u6570\u636e\u6e90\u6216\u6a21\u677f\u3002\n\n"
        "3. \u4e0b\u5468\u8ba1\u5212\uff1a\n"
        "- \u7ee7\u7eed\u5b8c\u5584\u81ea\u52a8\u5316\u529e\u516c\u6d41\u7a0b\u3002\n\n"
        f"\u8865\u5145\u8bf4\u660e\uff1a{task}"
    )


def _generate_notification_content(task: str) -> str:
    title_match = re.search(r"\u5173\u4e8e(.+?)(?:\u7684)?\u901a\u77e5", task)
    title = title_match.group(1) if title_match else "\u76f8\u5173\u4e8b\u9879"
    return f"\u5173\u4e8e{title}\u7684\u901a\u77e5\n\n{task}\n\n\u8bf7\u76f8\u5173\u4eba\u5458\u6309\u901a\u77e5\u5185\u5bb9\u6267\u884c\uff0c\u5982\u6709\u7591\u95ee\u8bf7\u53ca\u65f6\u6c9f\u901a\u3002\n"


def _call_minimax_or_fallback(
    prompt: str,
    messages: list[dict[str, Any]],
    intermediate: Dict[str, Any] | None = None,
) -> str:
    intermediate = intermediate or {}
    if not settings.minimax.configured or requests is None:
        content = truncate_text(str(intermediate.get("content", "")), 1200)
        if content:
            source = intermediate.get("source", "\u5185\u5bb9")
            return f"\u5f53\u524d\u672a\u914d\u7f6e Minimax API\uff0c\u5df2\u8fd4\u56de\u672c\u5730\u5206\u6790\u7ed3\u679c\u3002\n\u6765\u6e90\uff1a{source}\n\n{content}"
        return (
            "\u5f53\u524d\u672a\u914d\u7f6e Minimax API\uff0c\u5df2\u4f7f\u7528\u672c\u5730\u964d\u7ea7\u6a21\u5f0f\u3002\n"
            "\u6211\u53ef\u4ee5\u76f4\u63a5\u5904\u7406\u6587\u4ef6\u3001\u6587\u6863\u548c\u90ae\u4ef6\u7c7b\u6307\u4ee4\uff1b\u5982\u9700\u66f4\u5f3a\u7684\u5206\u6790\u80fd\u529b\uff0c\u8bf7\u914d\u7f6e MINIMAX_API_KEY\u3002"
        )

    payload = {
        "model": settings.minimax.model,
        "messages": messages,
        "temperature": 0.7,
        "stream": True,
    }
    headers = {
        "Authorization": f"Bearer {settings.minimax.api_key}",
        "Content-Type": "application/json",
    }
    if settings.minimax.group_id:
        headers["GroupId"] = settings.minimax.group_id

    full_content = []
    reasoning_content = []

    try:
        response = requests.post(
            f"{settings.minimax.base_url.rstrip('/')}/text/chatcompletion_v2",
            headers=headers,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            timeout=settings.minimax.timeout_seconds,
            stream=True,
        )
        response.raise_for_status()

        for line in response.iter_lines(decode_unicode=True):
            if not line.startswith("data: "):
                continue
            data_str = line[6:].strip()
            if not data_str or data_str == "[DONE]":
                continue
            try:
                chunk = json.loads(data_str)
            except json.JSONDecodeError:
                continue
            choices = chunk.get("choices") or []
            if not choices:
                continue
            delta = choices[0].get("delta", {})
            content = delta.get("content", "")
            reason = delta.get("reasoning_content", "")
            if content:
                full_content.append(content)
            if reason:
                reasoning_content.append(reason)
    except Exception as exc:
        return f"Minimax \u8c03\u7528\u5931\u8d25\uff1a{exc}"

    content_text = "".join(full_content)
    if not content_text and intermediate.get("content"):
        source = intermediate.get("source", "\u5185\u5bb9")
        content_text = f"Minimax \u672a\u8fd4\u56de\u5185\u5bb9\uff0c\u5df2\u964d\u7ea7\u4e3a\u672c\u5730\u7ed3\u679c\u3002\n\u6765\u6e90\uff1a{source}\n\n{truncate_text(str(intermediate.get('content', '')), 1200)}"

    result = {
        "content": content_text or "Minimax \u672a\u8fd4\u56de\u5185\u5bb9\u3002",
        "reasoning": "".join(reasoning_content),
    }
    return json.dumps(result, ensure_ascii=False)
