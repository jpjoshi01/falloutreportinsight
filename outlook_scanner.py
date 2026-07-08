from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

from config import ATTACHMENT_DIR, EXCEL_EXTENSIONS, ensure_data_dirs
from email_matcher import flexible_match
from filename_utils import safe_filename


@dataclass
class OutlookMatch:
    subject: str
    sender: str
    received_time: str
    matched_keywords: list[str]
    attachments: list[str]


def date_range_from_preset(preset: str, start: datetime | None = None, end: datetime | None = None) -> tuple[datetime, datetime]:
    now = datetime.now()
    normalized = preset.lower()
    if normalized == "last day":
        return now - timedelta(days=1), now
    if normalized == "last week":
        return now - timedelta(days=7), now
    if normalized == "last month":
        return now - timedelta(days=30), now
    return start or now - timedelta(days=1), end or now


def _outlook_namespace():
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for Outlook scanning. Install requirements.txt on Windows.") from exc
    try:
        return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as exc:
        message = str(exc)
        if "Invalid class string" in message or "-2147221005" in message:
            raise RuntimeError(
                "Microsoft Outlook desktop is not registered for COM automation on this Windows profile. "
                "Open Outlook desktop once, finish account setup, and confirm you are not using only Outlook Web/New Outlook. "
                "This scanner needs classic Outlook desktop because it reads the local MAPI inbox."
            ) from exc
        raise RuntimeError(f"Could not open Outlook desktop through Windows COM: {message}") from exc


def scan_outlook(
    keywords: str | Iterable[str],
    unread_only: bool = True,
    start: datetime | None = None,
    end: datetime | None = None,
    download_attachments: bool = True,
) -> list[OutlookMatch]:
    ensure_data_dirs()
    namespace = _outlook_namespace()
    inbox = namespace.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    matches: list[OutlookMatch] = []
    for item in items:
        try:
            received = item.ReceivedTime.replace(tzinfo=None)
            if start and received < start:
                break
            if end and received > end:
                continue
            if unread_only and not bool(item.UnRead):
                continue
            attachment_names = [item.Attachments.Item(i).FileName for i in range(1, item.Attachments.Count + 1)]
            haystack = f"{item.Subject}\n{getattr(item, 'Body', '')}\n{' '.join(attachment_names)}"
            is_match, matched_keywords = flexible_match(haystack, keywords)
            if not is_match:
                continue
            saved_paths: list[str] = []
            if download_attachments:
                for i in range(1, item.Attachments.Count + 1):
                    attachment = item.Attachments.Item(i)
                    name = safe_filename(attachment.FileName)
                    if Path(name).suffix.lower() not in EXCEL_EXTENSIONS:
                        continue
                    target = ATTACHMENT_DIR / f"{received:%Y%m%d%H%M%S}_{name}"
                    attachment.SaveAsFile(str(target))
                    saved_paths.append(str(target))
            matches.append(
                OutlookMatch(
                    subject=str(item.Subject or ""),
                    sender=str(getattr(item, "SenderName", "")),
                    received_time=received.isoformat(timespec="seconds"),
                    matched_keywords=matched_keywords,
                    attachments=saved_paths,
                )
            )
        except Exception:
            continue
    return matches
