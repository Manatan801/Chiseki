#!/usr/bin/env python3
"""
BM CSV差分検出ツールの簡易GUI。

同じフォルダに diff_bm_csv.py を置いて実行してください。
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path
try:
    from tkinter import (
        BOTH,
        DISABLED,
        END,
        LEFT,
        NORMAL,
        RIGHT,
        X,
        Button,
        Entry,
        Frame,
        Label,
        LabelFrame,
        StringVar,
        Tk,
        messagebox,
    )
    from tkinter import filedialog
    from tkinter.scrolledtext import ScrolledText
except ModuleNotFoundError:
    print("tkinter が見つからないためGUIを起動できません。")
    print("Windowsでは、Python公式インストーラーで tcl/tk and IDLE を含めてインストールしてください。")
    print("Ubuntu/Debian系では sudo apt install python3-tk が必要です。")
    sys.exit(1)

from diff_bm_csv import compare_properties, count_by_type, parse_bm_csv, write_diff_csv


class DiffCsvApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("BM CSV 差分検出")
        self.root.geometry("820x520")
        self.root.minsize(720, 460)

        self.old_csv = StringVar()
        self.new_csv = StringVar()
        self.output_csv = StringVar(value=str(Path.cwd() / "bm_csv_diff.csv"))
        self.last_output: Path | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        main = Frame(self.root, padx=14, pady=12)
        main.pack(fill=BOTH, expand=True)

        files = LabelFrame(main, text="CSV選択", padx=10, pady=10)
        files.pack(fill=X)

        self._file_row(files, "旧CSV", self.old_csv, self.select_old_csv, 0)
        self._file_row(files, "新CSV", self.new_csv, self.select_new_csv, 1)
        self._file_row(files, "出力CSV", self.output_csv, self.select_output_csv, 2, save=True)

        controls = Frame(main, pady=10)
        controls.pack(fill=X)

        self.run_button = Button(controls, text="差分を出力", width=16, command=self.run_diff)
        self.run_button.pack(side=LEFT)

        self.open_button = Button(
            controls,
            text="出力フォルダを開く",
            width=18,
            command=self.open_output_folder,
            state=DISABLED,
        )
        self.open_button.pack(side=LEFT, padx=(8, 0))

        Button(controls, text="終了", width=10, command=self.root.destroy).pack(side=RIGHT)

        result = LabelFrame(main, text="結果", padx=10, pady=10)
        result.pack(fill=BOTH, expand=True)

        self.log = ScrolledText(result, height=12, wrap="word")
        self.log.pack(fill=BOTH, expand=True)
        self.write_log("旧CSVと新CSVを選択し、出力先を確認してから「差分を出力」を押してください。")

    def _file_row(
        self,
        parent: Frame,
        label: str,
        variable: StringVar,
        command,
        row_index: int,
        save: bool = False,
    ) -> None:
        row = Frame(parent)
        row.pack(fill=X, pady=3)
        Label(row, text=label, width=9, anchor="w").pack(side=LEFT)
        Entry(row, textvariable=variable).pack(side=LEFT, fill=X, expand=True)
        button_text = "保存先" if save else "選択"
        Button(row, text=button_text, width=9, command=command).pack(side=LEFT, padx=(8, 0))

    def select_old_csv(self) -> None:
        path = self.ask_csv_file("旧CSVを選択")
        if path:
            self.old_csv.set(str(path))
            self.suggest_output_path()

    def select_new_csv(self) -> None:
        path = self.ask_csv_file("新CSVを選択")
        if path:
            self.new_csv.set(str(path))
            self.suggest_output_path()

    def select_output_csv(self) -> None:
        initial = Path(self.output_csv.get()).name or "bm_csv_diff.csv"
        path = filedialog.asksaveasfilename(
            title="差分CSVの保存先",
            defaultextension=".csv",
            initialfile=initial,
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if path:
            self.output_csv.set(path)

    def ask_csv_file(self, title: str) -> Path | None:
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        return Path(path) if path else None

    def suggest_output_path(self) -> None:
        old_path = Path(self.old_csv.get()) if self.old_csv.get() else None
        new_path = Path(self.new_csv.get()) if self.new_csv.get() else None
        if not old_path or not new_path:
            return
        output_dir = new_path.parent
        output_name = f"{old_path.stem}__{new_path.stem}__diff.csv"
        self.output_csv.set(str(output_dir / output_name))

    def run_diff(self) -> None:
        old_path = Path(self.old_csv.get())
        new_path = Path(self.new_csv.get())
        output_path = Path(self.output_csv.get())

        if not old_path.is_file():
            messagebox.showerror("入力エラー", "旧CSVを選択してください。")
            return
        if not new_path.is_file():
            messagebox.showerror("入力エラー", "新CSVを選択してください。")
            return
        if not output_path.name:
            messagebox.showerror("入力エラー", "出力CSVの保存先を指定してください。")
            return

        self.run_button.config(state=DISABLED)
        self.open_button.config(state=DISABLED)
        self.clear_log()
        self.write_log("差分を検出しています...")

        thread = threading.Thread(
            target=self._run_diff_worker,
            args=(old_path, new_path, output_path),
            daemon=True,
        )
        thread.start()

    def _run_diff_worker(self, old_path: Path, new_path: Path, output_path: Path) -> None:
        try:
            old = parse_bm_csv(old_path)
            new = parse_bm_csv(new_path)
            rows = compare_properties(old, new)
            write_diff_csv(output_path, rows)

            lines = [
                "完了しました。",
                "",
                f"旧ファイル: {old_path}",
                f"新ファイル: {new_path}",
                f"旧物件数: {len(old)}",
                f"新物件数: {len(new)}",
                f"差分件数: {len(rows)}",
                f"  追加: {count_by_type(rows, '追加')}",
                f"  削除: {count_by_type(rows, '削除')}",
                f"  変更: {count_by_type(rows, '変更')}",
                "",
                f"出力: {output_path}",
            ]
            self.last_output = output_path
            self.root.after(0, lambda: self._finish_success("\n".join(lines)))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_error(exc))

    def _finish_success(self, message: str) -> None:
        self.clear_log()
        self.write_log(message)
        self.run_button.config(state=NORMAL)
        self.open_button.config(state=NORMAL)
        messagebox.showinfo("完了", "差分CSVを出力しました。")

    def _finish_error(self, exc: Exception) -> None:
        self.clear_log()
        self.write_log(f"エラーが発生しました。\n\n{type(exc).__name__}: {exc}")
        self.run_button.config(state=NORMAL)
        self.open_button.config(state=DISABLED)
        messagebox.showerror("エラー", str(exc))

    def open_output_folder(self) -> None:
        if not self.last_output:
            return
        folder = self.last_output.parent
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            messagebox.showerror("エラー", f"出力フォルダを開けませんでした。\n{exc}")

    def write_log(self, text: str) -> None:
        self.log.insert(END, text + "\n")
        self.log.see(END)

    def clear_log(self) -> None:
        self.log.delete("1.0", END)


def main() -> None:
    root = Tk()
    app = DiffCsvApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
