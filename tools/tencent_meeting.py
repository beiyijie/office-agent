from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any, Dict, List

from office_agent.config import settings

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None


def _mock_meeting(subject: str, start_time: datetime, end_time: datetime, attendees: List[str] | None = None) -> Dict[str, Any]:
    meeting_code = f"8{start_time.strftime('%m%d%H%M')}"
    return {
        "meeting_id": f"mock-{start_time.strftime('%Y%m%d%H%M%S')}",
        "meeting_code": meeting_code,
        "subject": subject,
        "start_time": start_time.isoformat(),
        "end_time": end_time.isoformat(),
        "join_url": f"https://meeting.tencent.com/dm/{meeting_code}",
        "attendees": attendees or [],
        "provider": "tencent_meeting",
        "status": "mock",
    }


def create_meeting(
    subject: str,
    start_time: datetime,
    end_time: datetime,
    attendees: List[str] | None = None,
) -> Dict[str, Any]:
    config = settings.tencent_meeting
    attendees = attendees or []

    if config.dry_run or not config.configured or requests is None:
        return _mock_meeting(subject, start_time, end_time, attendees)

    payload = {
        "userid": config.user_id,
        "instanceid": 1,
        "subject": subject,
        "type": 0,
        "start_time": int(start_time.timestamp()),
        "end_time": int(end_time.timestamp()),
        "invitees": [{"userid": attendee} for attendee in attendees],
    }
    headers = {
        "Authorization": f"Bearer {config.access_token}",
        "Content-Type": "application/json",
        "X-TC-Registered": "1" if config.registered else "0",
    }
    response = requests.post(
        f"{config.base_url.rstrip('/')}/v1/meetings",
        json=payload,
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()

    meeting_info = data.get("meeting_info") or data.get("meeting") or data
    meeting_code = str(meeting_info.get("meeting_code") or meeting_info.get("meeting_id") or "")
    join_url = (
        meeting_info.get("join_url")
        or meeting_info.get("join_meeting_url")
        or (f"https://meeting.tencent.com/dm/{meeting_code}" if meeting_code else "")
    )

    return {
        "meeting_id": str(meeting_info.get("meeting_id") or ""),
        "meeting_code": meeting_code,
        "subject": meeting_info.get("subject") or subject,
        "start_time": meeting_info.get("start_time") or start_time.isoformat(),
        "end_time": meeting_info.get("end_time") or end_time.isoformat(),
        "join_url": join_url,
        "attendees": attendees,
        "provider": "tencent_meeting",
        "status": "created",
        "raw": data,
    }


def cancel_meeting(meeting_id: str) -> bool:
    config = settings.tencent_meeting
    if config.dry_run or not config.configured or requests is None:
        return True

    headers = {
        "Authorization": f"Bearer {config.access_token}",
        "X-TC-Registered": "1" if config.registered else "0",
    }
    response = requests.delete(
        f"{config.base_url.rstrip('/')}/v1/meetings/{meeting_id}",
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    return True


def get_meeting_detail(meeting_id: str) -> Dict[str, Any]:
    config = settings.tencent_meeting
    if config.dry_run or not config.configured or requests is None:
        now = datetime.now()
        return _mock_meeting("模拟腾讯会议", now + timedelta(hours=1), now + timedelta(hours=2))

    headers = {
        "Authorization": f"Bearer {config.access_token}",
        "X-TC-Registered": "1" if config.registered else "0",
    }
    response = requests.get(
        f"{config.base_url.rstrip('/')}/v1/meetings/{meeting_id}",
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    meeting_info = data.get("meeting_info") or data.get("meeting") or data
    return {
        "meeting_id": str(meeting_info.get("meeting_id") or meeting_id),
        "meeting_code": str(meeting_info.get("meeting_code") or ""),
        "subject": meeting_info.get("subject") or "",
        "start_time": meeting_info.get("start_time") or "",
        "end_time": meeting_info.get("end_time") or "",
        "join_url": meeting_info.get("join_url") or meeting_info.get("join_meeting_url") or "",
        "provider": "tencent_meeting",
        "status": "fetched",
        "raw": data,
    }
