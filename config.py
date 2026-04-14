from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = DATA_DIR / "cache"
CONFIG_FILE = BASE_DIR / "config.json"


def _load_config_file() -> dict[str, Any]:
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


FILE_CONFIG = _load_config_file()


def _get_config_value(*keys: str, env_name: str, default: Any = "") -> Any:
    current: Any = FILE_CONFIG
    for key in keys:
        if not isinstance(current, dict) or key not in current:
            current = None
            break
        current = current[key]
    env_value = os.getenv(env_name)
    if env_value not in (None, ""):
        return env_value
    if current not in (None, ""):
        return current
    return default


def _as_bool(value: Any, default: bool = False) -> bool:
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


@dataclass(slots=True)
class QQEmailConfig:
    email: str = str(_get_config_value("qq_email", "email", env_name="QQ_EMAIL", default=""))
    imap_host: str = str(_get_config_value("qq_email", "imap_host", env_name="QQ_IMAP_HOST", default="imap.qq.com"))
    imap_port: int = int(_get_config_value("qq_email", "imap_port", env_name="QQ_IMAP_PORT", default=993))
    smtp_host: str = str(_get_config_value("qq_email", "smtp_host", env_name="QQ_SMTP_HOST", default="smtp.qq.com"))
    smtp_port: int = int(_get_config_value("qq_email", "smtp_port", env_name="QQ_SMTP_PORT", default=465))
    auth_code: str = str(_get_config_value("qq_email", "auth_code", env_name="QQ_AUTH_CODE", default=""))

    @property
    def configured(self) -> bool:
        return bool(self.email and self.auth_code)


@dataclass(slots=True)
class MinimaxConfig:
    api_key: str = str(_get_config_value("minimax", "api_key", env_name="MINIMAX_API_KEY", default=""))
    base_url: str = str(
        _get_config_value("minimax", "base_url", env_name="MINIMAX_BASE_URL", default="https://api.minimax.chat/v1")
    )
    model: str = str(_get_config_value("minimax", "model", env_name="MINIMAX_MODEL", default="abab6-chat"))
    group_id: str = str(_get_config_value("minimax", "group_id", env_name="MINIMAX_GROUP_ID", default=""))
    timeout_seconds: int = int(
        _get_config_value("minimax", "timeout_seconds", env_name="MINIMAX_TIMEOUT_SECONDS", default=60)
    )

    @property
    def configured(self) -> bool:
        return bool(self.api_key)


@dataclass(slots=True)
class TencentMeetingConfig:
    base_url: str = str(
        _get_config_value(
            "tencent_meeting",
            "base_url",
            env_name="TENCENT_MEETING_BASE_URL",
            default="https://api.meeting.qq.com",
        )
    )
    access_token: str = str(
        _get_config_value("tencent_meeting", "access_token", env_name="TENCENT_MEETING_ACCESS_TOKEN", default="")
    )
    user_id: str = str(_get_config_value("tencent_meeting", "user_id", env_name="TENCENT_MEETING_USER_ID", default=""))
    user_id_type: int = int(
        _get_config_value("tencent_meeting", "user_id_type", env_name="TENCENT_MEETING_USER_ID_TYPE", default=1)
    )
    registered: bool = _as_bool(
        _get_config_value("tencent_meeting", "registered", env_name="TENCENT_MEETING_REGISTERED", default=True),
        default=True,
    )
    dry_run: bool = _as_bool(
        _get_config_value("tencent_meeting", "dry_run", env_name="TENCENT_MEETING_DRY_RUN", default=True),
        default=True,
    )

    @property
    def configured(self) -> bool:
        return bool(self.access_token and self.user_id)


@dataclass(slots=True)
class Settings:
    qq_email: QQEmailConfig = field(default_factory=QQEmailConfig)
    minimax: MinimaxConfig = field(default_factory=MinimaxConfig)
    tencent_meeting: TencentMeetingConfig = field(default_factory=TencentMeetingConfig)
    workspace_dir: Path = Path(
        str(_get_config_value("office_agent", "workspace_dir", env_name="OFFICE_AGENT_WORKSPACE", default=str(BASE_DIR)))
    ).resolve()
    max_retry_count: int = int(_get_config_value("office_agent", "max_retry_count", env_name="OFFICE_AGENT_MAX_RETRY", default=2))


settings = Settings()

for path in (DATA_DIR, CACHE_DIR):
    path.mkdir(parents=True, exist_ok=True)
