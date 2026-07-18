from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class AgendaItem:
    number: int
    title: str
    content: list[str] = field(default_factory=list)
    decisions: list[str] = field(default_factory=list)
    pending: list[str] = field(default_factory=list)
    speaker: str = ""


@dataclass
class TodoItem:
    content: str
    assignee: str = ""
    due_date: str = ""
    agenda_label: str = ""


@dataclass
class NextMeeting:
    date_time: str = ""
    place: str = ""
    note: str = ""
    place_note: str = ""


@dataclass
class MeetingMinutes:
    meeting_name: str = ""
    short_meeting_name: str = ""
    date_time: str = ""
    place: str = ""
    attendees: str = ""
    agendas: list[AgendaItem] = field(default_factory=list)
    todos: list[TodoItem] = field(default_factory=list)
    next_meeting: NextMeeting = field(default_factory=NextMeeting)
    summary: list[str] = field(default_factory=list)
