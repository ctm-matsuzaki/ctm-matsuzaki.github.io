from __future__ import annotations

from datetime import datetime, timedelta, timezone
import re


INVALID_FILENAME_CHARS = r'[\\/:*?"<>|]'
JST = timezone(timedelta(hours=9))
MAX_MEETING_NAME_LENGTH = 15


def build_output_filename(
    meeting_name: str,
    meeting_datetime: str,
    now: datetime | None = None,
    short_meeting_name: str = "",
) -> str:
    date_part = build_date_part(meeting_datetime, now)
    name_part = build_display_meeting_name(meeting_name, short_meeting_name)
    return f"{date_part}_{name_part}.xlsx"


def build_sheet_name(meeting_datetime: str, now: datetime | None = None) -> str:
    return build_date_part(meeting_datetime, now)


def build_date_part(meeting_datetime: str, now: datetime | None = None) -> str:
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

    current = _jst_now(now)
    return f"{current.month}.{current.day}"


def build_display_meeting_name(meeting_name: str, short_meeting_name: str = "") -> str:
    source = short_meeting_name.strip() or meeting_name.strip() or "会議議事録"
    name = _sanitize_meeting_name(source)
    if not name:
        name = "会議議事録"
    if len(name) <= MAX_MEETING_NAME_LENGTH:
        return name

    shortened = _remove_redundant_words(name)
    if len(shortened) <= MAX_MEETING_NAME_LENGTH:
        return shortened
    return shortened[:MAX_MEETING_NAME_LENGTH].rstrip(". ")


def _jst_now(now: datetime | None) -> datetime:
    if now is None:
        return datetime.now(JST)
    if now.tzinfo is None:
        return now.replace(tzinfo=JST)
    return now.astimezone(JST)


def _sanitize_meeting_name(name: str) -> str:
    name = re.sub(INVALID_FILENAME_CHARS, "", name)
    name = re.sub(r"\s+", "", name)
    return name.strip(". ")


def _remove_redundant_words(name: str) -> str:
    shortened = name
    for word in ("について", "に関する", "に関して", "の件", "打ち合わせ", "打合せ", "ミーティング"):
        shortened = shortened.replace(word, "")
    shortened = re.sub(r"[「」『』【】（）()［］\\[\\]]", "", shortened)
    shortened = shortened.strip(". ")
    return shortened or name
