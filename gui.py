import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from main import load_replacements, REPLACEMENTS_MD
from replacer import replace_in_docx, replace_multiple_in_docx


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("docx 文字替換工具")
        self.resizable(False, False)
        self._build_ui()

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── 輸入檔案 ──────────────────────────────────────────────────────────
        file_frame = tk.LabelFrame(self, text="輸入檔案", **pad)
        file_frame.grid(row=0, column=0, sticky="ew", **pad)

        self.input_path = tk.StringVar()
        tk.Entry(file_frame, textvariable=self.input_path, width=46, state="readonly").grid(
            row=0, column=0, padx=(6, 4), pady=6
        )
        tk.Button(file_frame, text="選擇 .docx", command=self._pick_file).grid(
            row=0, column=1, padx=(0, 6), pady=6
        )

        # ── 模式選擇 ──────────────────────────────────────────────────────────
        mode_frame = tk.LabelFrame(self, text="替換模式", **pad)
        mode_frame.grid(row=1, column=0, sticky="ew", **pad)

        self.mode = tk.StringVar(value="batch")
        tk.Radiobutton(
            mode_frame, text="批次模式（套用 replacements.md 所有規則）",
            variable=self.mode, value="batch", command=self._on_mode_change,
        ).grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
        tk.Radiobutton(
            mode_frame, text="單一模式（手動指定替換字串）",
            variable=self.mode, value="single", command=self._on_mode_change,
        ).grid(row=1, column=0, sticky="w", padx=6, pady=(2, 6))

        # ── 單一模式輸入 ──────────────────────────────────────────────────────
        self.single_frame = tk.LabelFrame(self, text="單一模式設定", **pad)
        self.single_frame.grid(row=2, column=0, sticky="ew", **pad)

        tk.Label(self.single_frame, text="目標字串").grid(
            row=0, column=0, sticky="e", padx=(6, 4), pady=6
        )
        self.target_var = tk.StringVar()
        tk.Entry(self.single_frame, textvariable=self.target_var, width=20).grid(
            row=0, column=1, sticky="w", pady=6
        )

        tk.Label(self.single_frame, text="替換字串").grid(
            row=0, column=2, sticky="e", padx=(12, 4)
        )
        self.replacement_var = tk.StringVar()
        tk.Entry(self.single_frame, textvariable=self.replacement_var, width=20).grid(
            row=0, column=3, sticky="w", padx=(0, 6)
        )

        # ── 輸出檔案名稱 ──────────────────────────────────────────────────────
        out_frame = tk.LabelFrame(self, text="輸出檔案名稱", **pad)
        out_frame.grid(row=3, column=0, sticky="ew", **pad)

        self.output_name = tk.StringVar()
        tk.Entry(out_frame, textvariable=self.output_name, width=46).grid(
            row=0, column=0, padx=6, pady=6, sticky="w"
        )
        tk.Label(out_frame, text="（留空則自動命名為 原檔名_modified.docx）", fg="gray").grid(
            row=0, column=1, padx=(0, 6)
        )

        # ── 執行按鈕 ──────────────────────────────────────────────────────────
        tk.Button(
            self, text="執行替換", bg="#2563EB", fg="white",
            font=("", 12, "bold"), padx=16, pady=6,
            command=self._run,
        ).grid(row=4, column=0, pady=(4, 8))

        # ── 結果顯示 ──────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self, text="執行結果", **pad)
        log_frame.grid(row=5, column=0, sticky="ew", **pad)

        self.log = tk.Text(log_frame, height=10, width=62, state="disabled",
                           bg="#F8F8F8", relief="flat")
        self.log.grid(row=0, column=0, padx=6, pady=6)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=6)
        self.log.configure(yscrollcommand=scrollbar.set)

        self._on_mode_change()

    # ── Event handlers ────────────────────────────────────────────────────────

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="選擇 .docx 檔案",
            filetypes=[("Word 文件", "*.docx"), ("所有檔案", "*.*")],
        )
        if path:
            self.input_path.set(path)
            stem = Path(path).stem
            if not self.output_name.get():
                self.output_name.set(f"{stem}_modified")

    def _on_mode_change(self):
        state = "normal" if self.mode.get() == "single" else "disabled"
        for child in self.single_frame.winfo_children():
            if isinstance(child, tk.Entry):
                child.configure(state=state)

    def _run(self):
        input_file = self.input_path.get()
        if not input_file:
            messagebox.showwarning("缺少輸入", "請先選擇 .docx 檔案。")
            return

        # Resolve output path
        input_path = Path(input_file)
        out_name = self.output_name.get().strip() or f"{input_path.stem}_modified"
        if not out_name.endswith(".docx"):
            out_name += ".docx"
        output_path_override = input_path.parent / out_name

        try:
            if self.mode.get() == "batch":
                self._run_batch(input_file, output_path_override)
            else:
                self._run_single(input_file, output_path_override)
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("錯誤", str(e))

    def _run_batch(self, input_file: str, output_path: Path):
        replacements = load_replacements(REPLACEMENTS_MD)
        if not replacements:
            messagebox.showerror("錯誤", "replacements.md 中找不到任何規則。")
            return

        raw_output, counts = replace_multiple_in_docx(input_file, replacements)

        # Rename to user-specified output path
        Path(raw_output).rename(output_path)

        lines = [f"輸出至：{output_path}\n"]
        total = 0
        for target, count in counts.items():
            if count > 0:
                replacement = next(r for t, r in replacements if t == target)
                lines.append(f"  「{target}」→「{replacement}」：{count} 處")
                total += count
        lines.append(f"\n共替換 {total} 處")
        self._log("\n".join(lines))

    def _run_single(self, input_file: str, output_path: Path):
        target = self.target_var.get()
        replacement = self.replacement_var.get()
        if not target:
            messagebox.showwarning("缺少輸入", "請輸入目標字串。")
            return

        raw_output, count = replace_in_docx(input_file, target, replacement)
        Path(raw_output).rename(output_path)

        self._log(
            f"輸出至：{output_path}\n"
            f"「{target}」→「{replacement}」：{count} 處"
        )

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.insert(tk.END, text)
        self.log.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
