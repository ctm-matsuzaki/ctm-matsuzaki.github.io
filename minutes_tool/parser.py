from __future__ import annotations

import re
from dataclasses import replace

from .models import AgendaItem, MeetingMinutes, NextMeeting, TodoItem


AGENDA_MARKS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
SECTION_STOP_RE = re.compile(r"^(ToDo|ＴｏＤｏ|次回会議|次回日程|全体要約|基本情報)\b")
AGENDA_RE = re.compile(r"^議題\s*([①-⑳]|\d+|[０-９]+)[\s　:：]*(.*)$")
SUBSECTION_NAMES = ("内容", "決定事項", "未決事項・継続検討事項", "未決事項", "継続検討事項")
TODO_DUE_RE = re.compile(
    r"^(?P<assignee>.+?)[\s　]+(?P<due>不明|未定|なし|"
    r"\d{4}[/-]\d{1,2}[/-]\d{1,2}|"
    r"\d{1,2}月(?:\d{1,2}日|中|末|上旬|中旬|下旬)?|"
    r"来週中|今週中|本日中|明日中|月内|訪問時|大阪訪問前または訪問時)"
    r"[\s　]+(?P<content>.+)$"
)
DUE_WORD = (
    r"不明|未定|なし|\d{4}[/-]\d{1,2}[/-]\d{1,2}|"
    r"\d{1,2}月(?:\d{1,2}日|中|末|上旬|中旬|下旬)?|"
    r"来週中|今週中|本日中|明日中|月内|訪問時|大阪訪問前または(?:訪問時)?"
)
DUE_CONTENT_RE = re.compile(rf"^(?P<due>{DUE_WORD})[\s　]+(?P<content>.+)$")
ASSIGNEE_DUE_RE = re.compile(rf"^(?P<assignee>.+?)[\s　]+(?P<due>{DUE_WORD})$")


class MinutesParseError(ValueError):
    """Raised only when there is no text to process."""


class MinutesParser:
    """Parse AI-generated Japanese meeting minutes into structured data.

    The parser intentionally keeps rules small and replaceable. Future input
    formats can add a new parser class without touching Excel output code.
    """

    def parse(self, text: str) -> MeetingMinutes:
        lines = self._normalize_lines(text)
        if not lines:
            return MeetingMinutes()

        minutes = MeetingMinutes(
            meeting_name=self._find_basic_value(lines, "会議名"),
            date_time=self._find_basic_value(lines, "日時"),
            place=self._find_basic_value(lines, "場所"),
            attendees=self._find_basic_value(lines, "出席者"),
            agendas=self._parse_agendas(lines),
            todos=self._parse_todos(lines),
            next_meeting=self._parse_next_meeting(lines),
            summary=self._parse_summary(lines),
        )

        if not minutes.agendas:
            minutes.agendas = [AgendaItem(number=1, title="", content=lines)]
        return minutes

    def _normalize_lines(self, text: str) -> list[str]:
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        normalized = normalized.replace("ＴｏＤｏ", "ToDo")
        lines: list[str] = []
        for raw in normalized.split("\n"):
            line = raw.strip()
            if not line:
                continue
            if re.fullmatch(r"\d+", line):
                continue
            line = self._clean_markdown_line(line)
            line = re.sub(r"^[・•\-●]\s*", "", line)
            line = line.strip()
            if not line:
                continue
            lines.append(line)
        return lines

    def _clean_markdown_line(self, line: str) -> str:
        if line.startswith("|"):
            return re.sub(r"\*\*(.*?)\*\*", r"\1", line).strip()
        line = re.sub(r"^#{1,6}\s*", "", line)
        line = re.sub(r"\*\*(.*?)\*\*", r"\1", line)
        line = re.sub(r"__(.*?)__", r"\1", line)
        line = line.strip()
        match = re.match(r"^(.+?)[\s　]*[:：][\s　]*(.*)$", line)
        if match and match.group(1).strip() in {"会議名", "日時", "場所", "出席者"}:
            return f"{match.group(1).strip()} {match.group(2).strip()}".strip()
        return line

    def _find_basic_value(self, lines: list[str], label: str) -> str:
        for index, line in enumerate(lines):
            if line == label:
                return self._next_value(lines, index)
            match = re.match(rf"^{re.escape(label)}[\s　:：]+(.+)$", line)
            if match:
                return match.group(1).strip()
        return ""

    def _next_value(self, lines: list[str], index: int) -> str:
        skip_labels = {"基本情報", "会議名", "日時", "場所", "出席者", *SUBSECTION_NAMES}
        for next_line in lines[index + 1 :]:
            if self._is_section_heading(next_line):
                return ""
            if next_line not in skip_labels:
                return next_line
        return ""

    def _is_section_heading(self, line: str) -> bool:
        return bool(
            AGENDA_RE.match(line)
            or line in {"基本情報", "ToDo", "次回会議", "次回日程", "全体要約", *SUBSECTION_NAMES}
        )

    def _parse_agendas(self, lines: list[str]) -> list[AgendaItem]:
        agendas: list[AgendaItem] = []
        agenda_indices = [
            (index, AGENDA_RE.match(line)) for index, line in enumerate(lines) if AGENDA_RE.match(line)
        ]
        for position, (start, match) in enumerate(agenda_indices):
            assert match is not None
            end = agenda_indices[position + 1][0] if position + 1 < len(agenda_indices) else len(lines)
            block = []
            for line in lines[start + 1 : end]:
                if SECTION_STOP_RE.match(line):
                    break
                block.append(line)
            agenda = AgendaItem(
                number=self._agenda_number(match.group(1)),
                title=match.group(2).strip(),
            )
            agendas.append(self._fill_agenda_sections(agenda, block))
        return agendas

    def _fill_agenda_sections(self, agenda: AgendaItem, block: list[str]) -> AgendaItem:
        current = "content"
        content: list[str] = []
        decisions: list[str] = []
        pending: list[str] = []

        for line in block:
            if line == "内容":
                current = "content"
                continue
            if line == "決定事項":
                current = "decisions"
                continue
            if line in {"未決事項・継続検討事項", "未決事項", "継続検討事項"}:
                current = "pending"
                continue

            if current == "decisions":
                decisions.append(line)
            elif current == "pending":
                pending.append(line)
            else:
                content.append(line)

        return replace(agenda, content=content, decisions=decisions, pending=pending)

    def _agenda_number(self, value: str) -> int:
        if value in AGENDA_MARKS:
            return AGENDA_MARKS.index(value) + 1
        return int(value.translate(str.maketrans("０１２３４５６７８９", "0123456789")))

    def _parse_todos(self, lines: list[str]) -> list[TodoItem]:
        start = self._find_line_index(lines, {"ToDo"})
        if start is None:
            return []

        end = len(lines)
        for index in range(start + 1, len(lines)):
            if lines[index].startswith(("次回会議", "次回日程", "全体要約")):
                end = index
                break

        rows = [
            line
            for line in lines[start + 1 : end]
            if line not in {"担当 期限 内容", "担当　期限　内容"}
            and not self._is_markdown_todo_header(line)
            and not self._is_markdown_separator_row(line)
        ]
        return self._parse_todo_rows(rows)

    def _parse_todo_rows(self, rows: list[str]) -> list[TodoItem]:
        todos: list[TodoItem] = []
        pending_fragments: list[str] = []
        index = 0

        while index < len(rows):
            line = rows[index]
            parsed = self._parse_single_todo_line(line)
            if parsed:
                if pending_fragments:
                    parsed.content = f"{''.join(pending_fragments)}{parsed.content}"
                    pending_fragments.clear()
                todos.append(parsed)
                index += 1
                continue

            due_content = DUE_CONTENT_RE.match(line)
            if due_content and pending_fragments:
                assignee = pending_fragments.pop(0)
                if index + 1 < len(rows) and self._looks_like_name_fragment(rows[index + 1]):
                    assignee = f"{assignee}{rows[index + 1]}"
                    index += 1
                todos.append(
                    TodoItem(
                        assignee=assignee,
                        due_date=due_content.group("due"),
                        content=f"{''.join(pending_fragments)}{due_content.group('content')}",
                    )
                )
                pending_fragments.clear()
                index += 1
                continue

            assignee_due = ASSIGNEE_DUE_RE.match(line)
            if assignee_due:
                continuation = rows[index + 1] if index + 1 < len(rows) else ""
                todos.append(
                    TodoItem(
                        assignee=assignee_due.group("assignee"),
                        due_date=assignee_due.group("due"),
                        content=f"{''.join(pending_fragments)}{continuation}",
                    )
                )
                pending_fragments.clear()
                index += 2 if continuation else 1
                continue

            if due_content and not pending_fragments and index + 2 < len(rows):
                assignee = rows[index + 1]
                next_due_content = DUE_CONTENT_RE.match(rows[index + 2])
                if self._looks_like_name_fragment(assignee) and next_due_content:
                    todos.append(
                        TodoItem(
                            assignee=assignee,
                            due_date=f"{due_content.group('due')}{next_due_content.group('due')}",
                            content=f"{due_content.group('content')}{next_due_content.group('content')}",
                        )
                    )
                    index += 3
                    continue

            pending_fragments.append(line)
            index += 1

        return todos

    def _parse_single_todo_line(self, line: str) -> TodoItem | None:
        if self._is_markdown_separator_row(line):
            return None
        pipe_parts = [part.strip() for part in re.split(r"\s*\|\s*", line.strip("| ")) if part.strip()]
        if len(pipe_parts) >= 3 and {"担当", "期限", "内容"}.issubset(set(pipe_parts[:3])):
            return None
        if len(pipe_parts) >= 3:
            return TodoItem(assignee=pipe_parts[0], due_date=pipe_parts[1], content=" ".join(pipe_parts[2:]))

        tab_parts = [part.strip() for part in line.split("\t") if part.strip()]
        if len(tab_parts) >= 3:
            return TodoItem(assignee=tab_parts[0], due_date=tab_parts[1], content=" ".join(tab_parts[2:]))

        match = TODO_DUE_RE.match(line)
        if match:
            return TodoItem(
                assignee=match.group("assignee").strip(),
                due_date=match.group("due").strip(),
                content=match.group("content").strip(),
            )
        return None

    def _is_markdown_separator_row(self, line: str) -> bool:
        parts = [part.strip() for part in line.strip("| ").split("|") if part.strip()]
        return bool(parts) and all(re.fullmatch(r":?-{3,}:?", part) for part in parts)

    def _is_markdown_todo_header(self, line: str) -> bool:
        parts = [part.strip() for part in line.strip("| ").split("|") if part.strip()]
        return len(parts) >= 3 and parts[:3] == ["担当", "期限", "内容"]

    def _looks_like_name_fragment(self, line: str) -> bool:
        return bool(line) and not any(token in line for token in ("、", "。", "・", "を", "する", "確認", "整理"))

    def _parse_next_meeting(self, lines: list[str]) -> NextMeeting:
        start = self._find_line_index(lines, {"次回会議", "次回日程", "【次回日程】"})
        if start is None:
            return NextMeeting()

        end = len(lines)
        for index in range(start + 1, len(lines)):
            if lines[index].startswith("全体要約"):
                end = index
                break
        block = lines[start + 1 : end]
        date_time, date_note = self._find_value_with_notes(block, "日時")
        place, place_note = self._find_value_with_notes(block, "場所")
        return NextMeeting(
            date_time=date_time,
            place=place,
            note=date_note,
            place_note=place_note,
        )

    def _parse_summary(self, lines: list[str]) -> list[str]:
        start = self._find_line_index(lines, {"全体要約"})
        if start is None:
            return []
        return [line for line in lines[start + 1 :] if line and not line.startswith("※")]

    def _find_line_index(self, lines: list[str], candidates: set[str]) -> int | None:
        for index, line in enumerate(lines):
            if line in candidates:
                return index
        return None

    def _find_value_with_notes(self, lines: list[str], label: str) -> tuple[str, str]:
        for index, line in enumerate(lines):
            value = ""
            if line == label:
                value = self._next_value(lines, index)
            else:
                match = re.match(rf"^{re.escape(label)}[\s　:：]+(.+)$", line)
                if match:
                    value = match.group(1).strip()
            if not value:
                continue

            notes: list[str] = []
            for next_line in lines[index + 1 :]:
                if next_line in {"日時", "場所"} or next_line.startswith(("日時", "場所")):
                    break
                if next_line.startswith("※"):
                    notes.append(next_line)
            return value, " / ".join(notes)
        return "", ""
