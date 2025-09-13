# utils/io_utils.py
import re
from datetime import datetime
from pathlib import Path

SAFE_CHARS = re.compile(r"[^A-Za-z0-9._-]+")

def sanitize_filename(name: str) -> str:
    """
    파일명/폴더명에 안전하지 않은 문자를 -로 치환
    """
    name = name.strip().replace(" ", "-")
    return SAFE_CHARS.sub("-", name)

def timestamp_utc() -> str:
    return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

def ensure_parent(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
