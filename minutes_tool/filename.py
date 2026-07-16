from __future__ import annotations

from datetime import datetime
import re


INVALID_FILENAME_CHARS = r'[\\/:*?"<>|]'


def build_output_filename(meeting_name: str, meeting_datetime: str, now: datetime | None = None) -> str:
    date_part = _date_part(meeting_datetime, now or datetime.now())
    name_part = _safe_meeting_name(meeting_name)
    return f"{date_part}_{name_part}.xlsx"


def _date_part(meeting_datetime: str, now: datetime) -> str:
    text = meeting_datetime or ""
    match = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", text)
    if match:
        return f"{int(match.group(2))}.{int(match.group(3))}"

    match = re.search(r"(\d{1,2})月\s*(\d{1,2})日", text)
    if match:
        return f"{int(match.group(1))}.{int(match.group(2))}"

    match = re.search(r"\d{4}[/-](\d{1,2})[/-](\d{1,2})", text)
    if match:
        return f"{int(match.group(1))}.{int(match.group(2))}"

    match = re.search(r"(\d{1,2})[/-](\d{1,2})", text)
    if match:
        return f"{int(match.group(1))}.{int(match.group(2))}"

    return f"{now.month}.{now.day}"


def _safe_meeting_name(meeting_name: str) -> str:
    name = (meeting_name or "").strip() or "会議議事録"
    name = re.sub(INVALID_FILENAME_CHARS, "", name)
    name = re.sub(r"\s+", "", name)
    name = name.strip(". ")
    if not name:
        name = "会議議事録"
    return name[:20]
