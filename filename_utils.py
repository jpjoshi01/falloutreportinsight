from __future__ import annotations

import hashlib
import re
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional


TIMESTAMP_RE = re.compile(r"(?<!\d)(20\d{12})(?!\d)")


def extract_timestamp_from_filename(filename: str) -> Optional[datetime]:
    match = TIMESTAMP_RE.search(Path(filename).name)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y%m%d%H%M%S")
    except ValueError:
        return None


def decide_latest_report(filenames: Iterable[str]) -> Optional[str]:
    names = list(filenames)
    if not names:
        return None
    with_dates = [(name, extract_timestamp_from_filename(name)) for name in names]
    dated = [(name, value) for name, value in with_dates if value is not None]
    if dated:
        return max(dated, key=lambda item: item[1])[0]
    return max(names)


def safe_filename(filename: str) -> str:
    name = Path(filename).name.strip()
    name = re.sub(r"[^\w.\- ()]+", "_", name)
    return name or "report.xlsx"


def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()
