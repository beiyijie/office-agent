from __future__ import annotations

from pathlib import Path
from typing import List

from office_agent.config import settings
from office_agent.utils import ensure_parent_dir


def _resolve(path: str) -> Path:
    target = Path(path).expanduser()
    if not target.is_absolute():
        target = settings.workspace_dir / target
    return target.resolve()


def list_directory(path: str = ".") -> List[str]:
    target = _resolve(path)
    return sorted(item.name for item in target.iterdir())


def read_file(path: str) -> str:
    target = _resolve(path)
    return target.read_text(encoding="utf-8")


def write_file(path: str, content: str) -> bool:
    target = ensure_parent_dir(_resolve(path))
    target.write_text(content, encoding="utf-8")
    return True


def file_exists(path: str) -> bool:
    return _resolve(path).exists()
