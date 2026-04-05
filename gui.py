import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path

from main import load_replacements, REPLACEMENTS_MD
from replacer import replace_in_docx, replace_multiple_in_docx

# ── Theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BLUE      = "#2563EB"
BLUE_DARK = "#1D4ED8"
BG        = "#F1F5F9"
CARD      = "#FFFFFF"
TEXT      = "#1E293B"
MUTED     = "#94A3B8"
SUCCESS   = "#16A34A"
BORDER    = "#E2E8F0"


class Card(ctk.CTkFrame):
    """White card with subtle shadow-like border."""
    def __init__(self, master, **kw):
        super().__init__(master, fg_color=CARD, corner_radius=12,
                         border_width=1, border_color=BORDER, **kw)


class SectionLabel(ctk.CTkLabel):
    def __init__(self, master, text, **kw):
        super().__init__(master, text=text, text_color=TEXT,
                         font=ctk.CTkFont(size=13, weight="bold"), **kw)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("文字替換工具")
        self.geometry("620x720")
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self._build_ui()

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color=BLUE, corner_radius=0, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header, text="📄  docx 文字替換工具",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white",
        ).place(relx=0.5, rely=0.5, anchor="center")

        scroll = ctk.CTkScrollableFrame(self, fg_color=BG, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        def section(title):
            card = Card(scroll)
            card.pack(fill="x", padx=20, pady=(14, 0))
            SectionLabel(card, title).pack(anchor="w", padx=16, pady=(12, 6))
            return card

        # ── 輸入檔案 ──────────────────────────────────────────────────────────
        c1 = section("① 選擇輸入檔案")
        row = ctk.CTkFrame(c1, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 14))

        self.input_path = ctk.StringVar()
        self._file_display = ctk.CTkEntry(
            row, textvariable=self.input_path, state="disabled",
            placeholder_text="尚未選擇檔案…",
            fg_color="#F8FAFC", border_color=BORDER,
            text_color=TEXT, font=ctk.CTkFont(size=13),
            height=38,
        )
        self._file_display.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            row, text="選擇 .docx", width=110, height=38,
            fg_color=BLUE, hover_color=BLUE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=8, command=self._pick_file,
        ).pack(side="left")

        # ── 替換模式 ──────────────────────────────────────────────────────────
        c2 = section("② 替換模式")
        self.mode = ctk.StringVar(value="batch")

        self._rb_batch = ctk.CTkRadioButton(
            c2, text="批次模式  ─  自動套用 replacements.md 所有規則",
            variable=self.mode, value="batch",
            command=self._on_mode_change,
            font=ctk.CTkFont(size=13), text_color=TEXT,
            fg_color=BLUE, hover_color=BLUE_DARK,
        )
        self._rb_batch.pack(anchor="w", padx=16, pady=(0, 8))

        self._rb_single = ctk.CTkRadioButton(
            c2, text="單一模式  ─  手動指定替換字串",
            variable=self.mode, value="single",
            command=self._on_mode_change,
            font=ctk.CTkFont(size=13), text_color=TEXT,
            fg_color=BLUE, hover_color=BLUE_DARK,
        )
        self._rb_single.pack(anchor="w", padx=16, pady=(0, 14))

        # ── 單一模式欄位 ──────────────────────────────────────────────────────
        self._single_card = section("③ 單一模式設定")
        pair = ctk.CTkFrame(self._single_card, fg_color="transparent")
        pair.pack(fill="x", padx=16, pady=(0, 14))
        pair.columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(pair, text="目標字串", text_color=MUTED,
                     font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(pair, text="替換字串", text_color=MUTED,
                     font=ctk.CTkFont(size=12)).grid(row=0, column=2, sticky="w", padx=(16, 0))

        self.target_var = ctk.StringVar()
        self._entry_target = ctk.CTkEntry(
            pair, textvariable=self.target_var, height=38,
            fg_color="#F8FAFC", border_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(size=13), placeholder_text="例：如果",
        )
        self._entry_target.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        self.replacement_var = ctk.StringVar()
        self._entry_repl = ctk.CTkEntry(
            pair, textvariable=self.replacement_var, height=38,
            fg_color="#F8FAFC", border_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(size=13), placeholder_text="例：若",
        )
        self._entry_repl.grid(row=1, column=2, columnspan=2, sticky="ew",
                              padx=(16, 0), pady=(4, 0))

        # ── 輸出名稱 ──────────────────────────────────────────────────────────
        c4 = section("④ 輸出檔案名稱")
        ctk.CTkLabel(c4, text="留空則自動命名為「原檔名_modified.docx」",
                     text_color=MUTED, font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=16, pady=(0, 6))

        self.output_name = ctk.StringVar()
        ctk.CTkEntry(
            c4, textvariable=self.output_name, height=38,
            fg_color="#F8FAFC", border_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(size=13), placeholder_text="例：thesis_final.docx",
        ).pack(fill="x", padx=16, pady=(0, 14))

        # ── 執行按鈕 ──────────────────────────────────────────────────────────
        ctk.CTkButton(
            scroll, text="執 行 替 換", height=48,
            fg_color=BLUE, hover_color=BLUE_DARK,
            font=ctk.CTkFont(size=15, weight="bold"),
            corner_radius=10, command=self._run,
        ).pack(fill="x", padx=20, pady=16)

        # ── 執行結果 ──────────────────────────────────────────────────────────
        c5 = section("執行結果")
        self.log = ctk.CTkTextbox(
            c5, height=180, fg_color="#F8FAFC",
            border_color=BORDER, border_width=1,
            font=ctk.CTkFont(family="Menlo", size=12),
            text_color=TEXT, corner_radius=8,
            state="disabled",
        )
        self.log.pack(fill="x", padx=16, pady=(0, 14))

        self._on_mode_change()

    # ── Events ────────────────────────────────────────────────────────────────

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="選擇 .docx 檔案",
            filetypes=[("Word 文件", "*.docx"), ("所有檔案", "*.*")],
        )
        if path:
            self.input_path.set(path)
            if not self.output_name.get():
                self.output_name.set(f"{Path(path).stem}_modified")

    def _on_mode_change(self):
        is_single = self.mode.get() == "single"
        state = "normal" if is_single else "disabled"
        fg = "#F8FAFC" if is_single else "#F1F5F9"
        for entry in (self._entry_target, self._entry_repl):
            entry.configure(state=state, fg_color=fg)
        alpha = 1.0 if is_single else 0.4
        self._single_card.configure(
            border_color=BLUE if is_single else BORDER
        )

    def _run(self):
        input_file = self.input_path.get()
        if not input_file:
            messagebox.showwarning("缺少輸入", "請先選擇 .docx 檔案。")
            return

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
        Path(raw_output).rename(output_path)

        total = sum(counts.values())
        lines = [f"✅  完成！輸出至：{output_path}", f"共替換 {total} 處\n"]
        for target, count in counts.items():
            if count > 0:
                replacement = next(r for t, r in replacements if t == target)
                lines.append(f"  「{target}」→「{replacement}」：{count} 處")
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
            f"✅  完成！輸出至：{output_path}\n\n"
            f"  「{target}」→「{replacement}」：{count} 處"
        )

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", text)
        self.log.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
