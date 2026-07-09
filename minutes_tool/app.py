from __future__ import annotations

from datetime import datetime
from pathlib import Path
import sys
from tkinter import END, filedialog, messagebox, ttk
import tkinter as tk

from .excel_writer import MinutesExcelError, MinutesExcelWriter
from .parser import MinutesParseError, MinutesParser
from .version import __version__


class MinutesToolApp(tk.Tk):
    TEMPLATE_FILENAME = "会議議事録テンプレ.xlsx"

    def __init__(self) -> None:
        super().__init__()
        self.title(f"CTM 議事録 Excel 作成ツール v{__version__}")
        self.geometry("960x720")
        self.minsize(780, 560)

        self.parser = MinutesParser()
        self.template_path = self._default_template_path()

        self._build_ui()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.grid(row=0, column=0, sticky="nsew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        root.columnconfigure(0, weight=1)

        header = ttk.Frame(root)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)

        title_label = ttk.Label(header, text=f"CTM 議事録 Excel 作成ツール v{__version__}", font=("", 16, "bold"))
        title_label.grid(row=0, column=0, sticky="w")

        description = (
            "AIで作成した議事録テキストを中央の入力欄に貼り付け、"
            "画面下部の「Excel作成」を押してください。"
        )
        description_label = ttk.Label(header, text=description, wraplength=860, justify="left")
        description_label.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        self.status = tk.StringVar(value="準備完了")
        status_label = ttk.Label(header, textvariable=self.status, foreground="#555555")
        status_label.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        text_frame = ttk.Frame(root)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.text = tk.Text(text_frame, wrap="word", undo=True)
        y_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=y_scroll.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")

        button_bar = ttk.Frame(root)
        button_bar.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        button_bar.columnconfigure(0, weight=1)
        button_bar.columnconfigure(4, weight=1)

        create_button = ttk.Button(button_bar, text="Excel作成", command=self._create_excel, width=16)
        create_button.grid(row=0, column=1, padx=6)

        clear_button = ttk.Button(button_bar, text="クリア", command=self._clear_text, width=16)
        clear_button.grid(row=0, column=2, padx=6)

        close_button = ttk.Button(button_bar, text="終了", command=self.destroy, width=16)
        close_button.grid(row=0, column=3, padx=6)

    def _create_excel(self) -> None:
        source_text = self.text.get("1.0", END).strip()
        if not source_text:
            messagebox.showerror("入力エラー", "議事録テキストを貼り付けてください。")
            return

        try:
            minutes = self.parser.parse(source_text)
        except MinutesParseError as exc:
            messagebox.showerror("解析エラー", str(exc))
            self.status.set("解析エラーがあります。入力形式を確認してください。")
            return
        except Exception as exc:
            messagebox.showerror("解析エラー", f"議事録の解析中にエラーが発生しました。\n{exc}")
            self.status.set("解析エラーがあります。")
            return

        template_path = self._resolve_template_path()
        if template_path is None:
            self.status.set("テンプレート選択をキャンセルしました。")
            return

        default_name = f"会議議事録_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
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
                "「会議議事録テンプレ.xlsx」を選択して、もう一度作成してください。",
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
            "次の画面で「会議議事録テンプレ.xlsx」を選択してください。",
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
