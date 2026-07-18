from __future__ import annotations

from datetime import datetime
from pathlib import Path
import sys
from tkinter import BOTH, END, LEFT, X, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import tkinter as tk

from .excel_writer import MinutesExcelError, MinutesExcelWriter
from .filename import build_output_filename
from .models import AgendaItem, MeetingMinutes
from .parser import MinutesParseError, MinutesParser
from .version import __version__


class MinutesToolApp(tk.Tk):
    TEMPLATE_FILENAME = "会議議事録テンプレのコピー.xlsx"
    COLOR_BG = "#1A1A1A"
    COLOR_TITLE = "#FFFFFF"
    COLOR_DESCRIPTION = "#DDDDDD"
    COLOR_STATUS = "#AAAAAA"
    COLOR_INPUT_BG = "#1E1E1E"
    COLOR_INPUT_TEXT = "#FFFFFF"
    COLOR_SELECTION = "#2F6B55"
    COLOR_BUTTON_BG = "#3A3A3A"
    COLOR_BUTTON_HOVER = "#4A4A4A"
    COLOR_SCROLLBAR_BG = "#2C2C2C"
    COLOR_SCROLLBAR_THUMB = "#555555"
    APP_DISPLAY_NAME = "議事録作成ツール"

    def __init__(self) -> None:
        super().__init__()
        self.title(f"{self.APP_DISPLAY_NAME} v{__version__}")
        self.geometry("960x720+120+80")
        self.minsize(780, 560)
        self.configure(bg=self.COLOR_BG)

        self.parser = MinutesParser()
        self.template_path = self._default_template_path()

        self._build_ui()
        self.after(100, self._show_window)

    def _build_ui(self) -> None:
        title_label = tk.Label(
            self,
            text=f"{self.APP_DISPLAY_NAME} v{__version__}",
            bg=self.COLOR_BG,
            fg=self.COLOR_TITLE,
            font=("Helvetica", 16, "bold"),
            anchor="w",
        )
        title_label.pack(fill=X, padx=16, pady=(16, 8))

        description = (
            "AIで作成した議事録テキストを中央の入力欄に貼り付け、"
            "画面下部の「Excel作成」を押してください。"
        )
        description_label = tk.Label(
            self,
            text=description,
            bg=self.COLOR_BG,
            fg=self.COLOR_DESCRIPTION,
            font=("Helvetica", 12),
            justify="left",
            anchor="w",
        )
        description_label.pack(fill=X, padx=16, pady=(0, 8))

        self.text = ScrolledText(
            self,
            wrap="word",
            undo=True,
            bg=self.COLOR_INPUT_BG,
            fg=self.COLOR_INPUT_TEXT,
            insertbackground=self.COLOR_INPUT_TEXT,
            selectbackground=self.COLOR_SELECTION,
            selectforeground=self.COLOR_INPUT_TEXT,
            relief="solid",
            bd=1,
            padx=8,
            pady=8,
            font=("Helvetica", 12),
            width=80,
            height=24,
        )
        if hasattr(self.text, "frame"):
            self.text.frame.configure(bg=self.COLOR_BG)
        if hasattr(self.text, "vbar"):
            self.text.vbar.configure(
                bg=self.COLOR_SCROLLBAR_THUMB,
                activebackground=self.COLOR_SCROLLBAR_THUMB,
                troughcolor=self.COLOR_SCROLLBAR_BG,
                highlightbackground=self.COLOR_BG,
                relief="flat",
                bd=0,
            )
        self.text.pack(fill=BOTH, expand=True, padx=16, pady=(0, 8))

        self.status = tk.StringVar(value="準備完了")
        status_label = tk.Label(
            self,
            textvariable=self.status,
            bg=self.COLOR_BG,
            fg=self.COLOR_STATUS,
            font=("Helvetica", 11),
            anchor="w",
        )
        status_label.pack(fill=X, padx=16, pady=(0, 8))

        button_bar = tk.Frame(self, bg=self.COLOR_BG)
        button_bar.pack(fill=X, padx=16, pady=(0, 16))

        create_button = self._make_button(
            button_bar,
            text="Excel作成",
            command=self._create_excel,
        )
        create_button.pack(side=LEFT, padx=(0, 8))

        clear_button = self._make_button(button_bar, text="クリア", command=self._clear_text)
        clear_button.pack(side=LEFT, padx=8)

        close_button = self._make_button(button_bar, text="終了", command=self.destroy)
        close_button.pack(side=LEFT, padx=8)

    def _make_button(
        self,
        parent: tk.Widget,
        text: str,
        command,
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            width=14,
            height=1,
            bg=self.COLOR_BUTTON_BG,
            fg="#000000",
            activebackground=self.COLOR_BUTTON_HOVER,
            activeforeground=self.COLOR_TITLE,
            relief="raised",
            bd=1,
            highlightthickness=0,
            font=("Helvetica", 12),
        )

    def _show_window(self) -> None:
        self.update_idletasks()
        self.deiconify()
        self.lift()
        self.focus_force()
        try:
            self.attributes("-topmost", True)
            self.after(800, lambda: self.attributes("-topmost", False))
        except tk.TclError:
            pass

    def _create_excel(self) -> None:
        source_text = self.text.get("1.0", END).strip()

        try:
            minutes = self.parser.parse(source_text)
        except MinutesParseError as exc:
            messagebox.showwarning("確認", "一部の項目を自動判定できませんでした。黄色セルを確認してください。")
            minutes = MeetingMinutes(agendas=[AgendaItem(number=1, title="", content=[source_text] if source_text else [])])
        except Exception as exc:
            messagebox.showwarning(
                "確認",
                "一部の項目を自動判定できませんでした。黄色セルを確認してください。\n"
                f"詳細: {exc}",
            )
            minutes = MeetingMinutes(agendas=[AgendaItem(number=1, title="", content=[source_text] if source_text else [])])
        self.status.set("一部の項目を自動判定できない場合は、Excelの黄色セルを確認してください。")

        template_path = self._resolve_template_path()
        if template_path is None:
            self.status.set("テンプレート選択をキャンセルしました。")
            return

        default_name = build_output_filename(
            minutes.meeting_name,
            minutes.date_time,
            short_meeting_name=minutes.short_meeting_name,
        )
        output_path = filedialog.asksaveasfilename(
            title="Excelファイルの保存先を選択",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            self.status.set("保存をキャンセルしました。")
            return

        try:
            saved_path = MinutesExcelWriter(template_path).write(minutes, output_path)
        except FileNotFoundError:
            messagebox.showerror(
                "テンプレートエラー",
                "Excelテンプレートが見つかりませんでした。\n"
                "「会議議事録テンプレのコピー.xlsx」を選択して、もう一度作成してください。",
            )
            self.status.set("テンプレートが見つかりません。")
            return
        except PermissionError:
            messagebox.showerror(
                "保存エラー",
                "Excelファイルを保存できませんでした。\n\n"
                "同じ名前のExcelファイルが開かれている場合は閉じてください。\n"
                "また、保存先フォルダに書き込み権限があるか確認してください。",
            )
            self.status.set("保存エラーがあります。")
            return
        except MinutesExcelError as exc:
            messagebox.showerror("Excel作成エラー", str(exc))
            self.status.set("Excel作成エラーがあります。")
            return
        except Exception as exc:
            messagebox.showerror(
                "Excel作成エラー",
                "Excelファイルを作成できませんでした。\n\n"
                "入力内容、保存先、テンプレートファイルを確認してください。\n"
                f"詳細: {exc}",
            )
            self.status.set("Excel作成エラーがあります。")
            return

        messagebox.showinfo("作成完了", f"Excelファイルを作成しました。\n\n保存先:\n{saved_path}")
        self.status.set(f"作成完了: {saved_path}")

    def _clear_text(self) -> None:
        self.text.delete("1.0", END)
        self.status.set("入力欄をクリアしました。")

    def _default_template_path(self) -> Path:
        candidates = []
        bundled_root = getattr(sys, "_MEIPASS", None)
        if bundled_root:
            candidates.append(Path(bundled_root) / "minutes_tool" / "resources" / self.TEMPLATE_FILENAME)
            candidates.append(Path(bundled_root) / "resources" / self.TEMPLATE_FILENAME)
        candidates.append(Path(__file__).resolve().parent / "resources" / self.TEMPLATE_FILENAME)

        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[-1]

    def _resolve_template_path(self) -> Path | None:
        if self.template_path.exists():
            return self.template_path

        messagebox.showwarning(
            "テンプレート確認",
            "Excelテンプレートが見つかりません。\n"
            "次の画面で「会議議事録テンプレのコピー.xlsx」を選択してください。",
        )
        selected = filedialog.askopenfilename(
            title="Excelテンプレートを選択",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not selected:
            return None

        self.template_path = Path(selected)
        self.status.set(f"テンプレート選択済み: {self.template_path}")
        return self.template_path


def run() -> None:
    app = MinutesToolApp()
    app.mainloop()
