from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

from .models import AgendaItem, MeetingMinutes, TodoItem


class MinutesExcelError(RuntimeError):
    """User-facing Excel generation error."""


class MinutesExcelWriter:
    """Write parsed minutes into the fixed CTM Excel template."""

    DYNAMIC_START_ROW = 10
    DYNAMIC_END_ROW = 39
    AGENDA_BLOCK_HEIGHT = 5
    TODO_MIN_ROWS = 1

    def __init__(self, template_path: str | Path) -> None:
        self.template_path = Path(template_path)

    def write(self, minutes: MeetingMinutes, output_path: str | Path) -> Path:
        if not self.template_path.exists():
            raise FileNotFoundError(f"テンプレートが見つかりません: {self.template_path}")

        output = Path(output_path)
        if output.suffix.lower() != ".xlsx":
            output = output.with_suffix(".xlsx")
        output.parent.mkdir(parents=True, exist_ok=True)

        try:
            workbook = load_workbook(self.template_path)
        except InvalidFileException as exc:
            raise MinutesExcelError(
                "選択されたテンプレートをExcelファイルとして読み込めませんでした。\n"
                "拡張子が .xlsx の会議議事録テンプレートを選択してください。"
            ) from exc
        except PermissionError:
            raise
        except Exception as exc:
            raise MinutesExcelError(
                "Excelテンプレートを読み込めませんでした。\n"
                "テンプレートファイルが破損していないか確認してください。"
            ) from exc

        sheet = workbook.active
        self._validate_template(sheet)

        self._write_basic_info(sheet, minutes)
        self._rebuild_dynamic_area(sheet, minutes)
        self._write_next_meeting(sheet, minutes)

        try:
            workbook.save(output)
        except PermissionError:
            raise
        except Exception as exc:
            raise MinutesExcelError(
                "Excelファイルの保存中にエラーが発生しました。\n"
                "保存先、ファイル名、Excelで同名ファイルを開いていないかを確認してください。"
            ) from exc
        return output

    def _validate_template(self, sheet: Worksheet) -> None:
        required_values = {
            "B2": "会議議事録",
            "B4": "会議名",
            "B5": "日時",
            "B6": "場所",
            "B7": "出席者",
            "B9": "議題",
            "C9": "打合せ内容",
            "E9": "発言者",
            "B26": "議題",
            "C26": "ToDoタスク",
            "D26": "担当者",
            "E26": "期限",
        }
        missing = [cell for cell, value in required_values.items() if sheet[cell].value != value]
        if missing:
            raise MinutesExcelError(
                "選択されたExcelは、CTM議事録テンプレートとして認識できませんでした。\n"
                "添付の「会議議事録テンプレ.xlsx」を選択してください。"
            )

    def _write_basic_info(self, sheet: Worksheet, minutes: MeetingMinutes) -> None:
        sheet["C4"] = minutes.meeting_name
        sheet["C5"] = minutes.date_time
        sheet["C6"] = minutes.place
        sheet["C7"] = minutes.attendees

    def _rebuild_dynamic_area(self, sheet: Worksheet, minutes: MeetingMinutes) -> None:
        agenda_templates = self._capture_rows(sheet, 10, 14)
        spacer_after_agenda = self._capture_rows(sheet, 25, 25)[0]
        todo_header_template = self._capture_rows(sheet, 26, 26)[0]
        todo_first_template = self._capture_rows(sheet, 27, 27)[0]
        todo_body_template = self._capture_rows(sheet, 28, 28)[0]
        spacer_after_todo = self._capture_rows(sheet, 39, 39)[0]

        self._unmerge_rows(sheet, self.DYNAMIC_START_ROW, self.DYNAMIC_END_ROW)
        sheet.delete_rows(self.DYNAMIC_START_ROW, self.DYNAMIC_END_ROW - self.DYNAMIC_START_ROW + 1)

        agenda_count = max(len(minutes.agendas), 1)
        todo_count = max(len(minutes.todos), self.TODO_MIN_ROWS)
        insert_count = agenda_count * self.AGENDA_BLOCK_HEIGHT + 1 + 1 + todo_count + 1
        sheet.insert_rows(self.DYNAMIC_START_ROW, insert_count)

        row = self.DYNAMIC_START_ROW
        agendas = minutes.agendas or [AgendaItem(number=1, title="")]
        for agenda in agendas:
            for offset, template in enumerate(agenda_templates):
                self._apply_row_template(sheet, row + offset, template)
                sheet.merge_cells(start_row=row + offset, start_column=3, end_row=row + offset, end_column=4)
            self._write_agenda_block(sheet, row, agenda)
            row += self.AGENDA_BLOCK_HEIGHT

        self._apply_row_template(sheet, row, spacer_after_agenda)
        row += 1

        self._apply_row_template(sheet, row, todo_header_template)
        sheet.cell(row=row, column=2).value = "議題"
        sheet.cell(row=row, column=3).value = "ToDoタスク"
        sheet.cell(row=row, column=4).value = "担当者"
        sheet.cell(row=row, column=5).value = "期限"
        row += 1

        todos = minutes.todos or [TodoItem(content="")]
        previous_agenda = None
        for index, todo in enumerate(todos):
            template = todo_first_template if index == 0 else todo_body_template
            self._apply_row_template(sheet, row, template)
            self._write_todo_row(sheet, row, todo, previous_agenda)
            previous_agenda = todo.agenda_label or previous_agenda
            row += 1

        self._apply_row_template(sheet, row, spacer_after_todo)

    def _write_agenda_block(self, sheet: Worksheet, row: int, agenda: AgendaItem) -> None:
        sheet.cell(row=row, column=2).value = f"議題{self._circled_number(agenda.number)}"
        sheet.cell(row=row, column=3).value = agenda.title
        sheet.cell(row=row, column=5).value = agenda.speaker

        content_lines = self._fit_lines(agenda.content, 3)
        for offset, value in enumerate(content_lines, start=1):
            sheet.cell(row=row + offset, column=3).value = value

        sheet.cell(row=row + 4, column=2).value = "決定事項"
        sheet.cell(row=row + 4, column=3).value = self._join_items(agenda.decisions)

    def _write_todo_row(self, sheet: Worksheet, row: int, todo: TodoItem, previous_agenda: str | None) -> None:
        label = todo.agenda_label
        sheet.cell(row=row, column=1).value = False
        sheet.cell(row=row, column=2).value = "" if label == previous_agenda else label
        sheet.cell(row=row, column=3).value = todo.content
        sheet.cell(row=row, column=4).value = todo.assignee
        sheet.cell(row=row, column=5).value = todo.due_date

    def _write_next_meeting(self, sheet: Worksheet, minutes: MeetingMinutes) -> None:
        next_row = self._find_next_meeting_row(sheet)
        if next_row is None:
            return
        date_time = minutes.next_meeting.date_time
        if minutes.next_meeting.note:
            date_time = f"{date_time}（{minutes.next_meeting.note}）" if date_time else minutes.next_meeting.note
        sheet.cell(row=next_row + 1, column=2).value = f"日時：{date_time}" if date_time else "日時："
        place = minutes.next_meeting.place
        if minutes.next_meeting.place_note:
            place = f"{place}（{minutes.next_meeting.place_note}）" if place else minutes.next_meeting.place_note
        sheet.cell(row=next_row + 2, column=2).value = (
            f"場所：{place}" if place else "場所："
        )

    def _capture_rows(self, sheet: Worksheet, start: int, end: int) -> list[dict]:
        rows = []
        for row_number in range(start, end + 1):
            rows.append(
                {
                    "height": sheet.row_dimensions[row_number].height,
                    "cells": [self._capture_cell(sheet.cell(row=row_number, column=col)) for col in range(1, 6)],
                }
            )
        return rows

    def _capture_cell(self, cell) -> dict:
        return {
            "style": copy(cell._style),
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format,
            "protection": copy(cell.protection),
            "value": None if isinstance(cell, MergedCell) else cell.value,
        }

    def _apply_row_template(self, sheet: Worksheet, row_number: int, template: dict) -> None:
        sheet.row_dimensions[row_number].height = template["height"]
        for column, cell_template in enumerate(template["cells"], start=1):
            cell = sheet.cell(row=row_number, column=column)
            cell._style = copy(cell_template["style"])
            cell.font = copy(cell_template["font"])
            cell.fill = copy(cell_template["fill"])
            cell.border = copy(cell_template["border"])
            cell.alignment = copy(cell_template["alignment"])
            cell.number_format = cell_template["number_format"]
            cell.protection = copy(cell_template["protection"])
            cell.value = cell_template["value"]

    def _unmerge_rows(self, sheet: Worksheet, start_row: int, end_row: int) -> None:
        for merged_range in list(sheet.merged_cells.ranges):
            if start_row <= merged_range.min_row and merged_range.max_row <= end_row:
                sheet.unmerge_cells(str(merged_range))

    def _find_next_meeting_row(self, sheet: Worksheet) -> int | None:
        for row in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row, column=2).value
            if isinstance(value, str) and "次回日程" in value:
                return row
        return None

    def _fit_lines(self, items: list[str], max_rows: int) -> list[str]:
        if not items:
            return [""] * max_rows
        rows = items[: max_rows - 1]
        rows.append(self._join_items(items[max_rows - 1 :]))
        return rows[:max_rows]

    def _join_items(self, items: list[str]) -> str:
        return "\n".join(f"・{item}" for item in items if item)

    def _circled_number(self, number: int) -> str:
        if 1 <= number <= len("①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"):
            return "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"[number - 1]
        return str(number)
