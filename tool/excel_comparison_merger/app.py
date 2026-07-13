"""Tkinter desktop interface for comparing two experiment workbooks."""

from __future__ import annotations

import queue
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from merge_excel import merge_experiment_workbooks


TOOL_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = TOOL_DIR / "output"


class ComparisonMergerApp:
    """Small desktop interface for choosing and merging two XLSX files."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("情绪预测实验对比工具")
        self.root.geometry("760x330")
        self.root.minsize(680, 300)
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()

        self.file1 = tk.StringVar()
        self.file2 = tk.StringVar()
        default_name = f"emotion_comparison_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        self.output_name = tk.StringVar(value=default_name)
        self.status = tk.StringVar(value="请选择两份实验结果文件。")

        self._build_ui()
        self.root.after(100, self._poll_events)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=24)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="情绪预测实验对比", font=("Arial", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 22)
        )
        self._file_row(frame, 1, "实验文件 1", self.file1)
        self._file_row(frame, 2, "实验文件 2", self.file2)

        ttk.Label(frame, text="合并文件名").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.output_name).grid(
            row=3, column=1, columnspan=2, sticky="ew", padx=(12, 0), pady=8
        )

        ttk.Label(
            frame,
            text=f"输出目录：{OUTPUT_DIR}",
            foreground="#555555",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 16))

        self.merge_button = ttk.Button(frame, text="开始合并", command=self._start_merge)
        self.merge_button.grid(row=5, column=0, sticky="w")
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=180)
        self.progress.grid(row=5, column=1, sticky="w", padx=12)
        ttk.Label(frame, textvariable=self.status, wraplength=680).grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(18, 0)
        )

    def _file_row(
        self, frame: ttk.Frame, row: int, label: str, variable: tk.StringVar
    ) -> None:
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=variable).grid(
            row=row, column=1, sticky="ew", padx=12, pady=8
        )
        ttk.Button(
            frame,
            text="选择文件",
            command=lambda: self._choose_file(variable),
        ).grid(row=row, column=2, pady=8)

    def _choose_file(self, variable: tk.StringVar) -> None:
        selected = filedialog.askopenfilename(
            title="选择实验结果 Excel",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if selected:
            variable.set(selected)

    def _start_merge(self) -> None:
        file1 = Path(self.file1.get().strip())
        file2 = Path(self.file2.get().strip())
        output_name = self.output_name.get().strip()

        if not file1.is_file() or not file2.is_file():
            messagebox.showerror("无法合并", "请选择两份有效的 .xlsx 文件。")
            return
        if not output_name:
            messagebox.showerror("无法合并", "请输入合并文件名。")
            return
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"
            self.output_name.set(output_name)
        if Path(output_name).name != output_name:
            messagebox.showerror("无法合并", "文件名中不能包含目录路径。")
            return

        output_path = OUTPUT_DIR / output_name
        if output_path.exists() and not messagebox.askyesno(
            "文件已存在", f"{output_name} 已存在，是否覆盖？"
        ):
            return

        self.merge_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("正在读取并合并工作簿，请勿关闭窗口……")
        threading.Thread(
            target=self._merge_worker,
            args=(file1, file2, output_path),
            daemon=True,
        ).start()

    def _merge_worker(self, file1: Path, file2: Path, output_path: Path) -> None:
        try:
            report = merge_experiment_workbooks(file1, file2, output_path)
            self.events.put(("success", (output_path, report)))
        except Exception as exc:
            self.events.put(("error", exc))

    def _poll_events(self) -> None:
        try:
            event, payload = self.events.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_events)
            return

        self.progress.stop()
        self.merge_button.configure(state="normal")
        if event == "success":
            output_path, report = payload
            self.status.set(f"合并完成：{output_path.name}")
            messagebox.showinfo(
                "合并完成",
                f"输出文件：\n{output_path}\n\n{report.summary()}",
            )
        else:
            self.status.set("合并失败，请检查所选文件。")
            messagebox.showerror("合并失败", str(payload))
        self.root.after(100, self._poll_events)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    root = tk.Tk()
    ComparisonMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
