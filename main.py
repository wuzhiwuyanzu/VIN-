"""VIN batch generator for Windows."""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


TRANSLITERATION = {
    **dict.fromkeys("AJ", 1),
    **dict.fromkeys("BKS", 2),
    **dict.fromkeys("CLT", 3),
    **dict.fromkeys("DMU", 4),
    **dict.fromkeys("ENV", 5),
    **dict.fromkeys("FW", 6),
    **dict.fromkeys("GPX", 7),
    **dict.fromkeys("HY", 8),
    **dict.fromkeys("RZ", 9),
    **{str(n): n for n in range(10)},
}
WEIGHTS = (8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2)
ALLOWED = set(TRANSLITERATION)


def check_digit(vin_without_check: str) -> str:
    total = sum(TRANSLITERATION[ch] * weight for ch, weight in zip(vin_without_check, WEIGHTS))
    remainder = total % 11
    return "X" if remainder == 10 else str(remainder)


def generate(prefix: str, post_check: str, serial_start: int, quantity: int) -> list[str]:
    rows = []
    for serial in range(serial_start, serial_start + quantity):
        partial = prefix + "0" + post_check + f"{serial:06d}"
        rows.append(prefix + check_digit(partial) + post_check + f"{serial:06d}")
    return rows


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VIN 批量生成器")
        self.geometry("900x760")
        self.minsize(760, 650)
        self.configure(bg="#f1f5f9")
        self.records: list[str] = []
        self._style()
        self._build()

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f1f5f9")
        style.configure("Card.TFrame", background="white")
        style.configure("TLabel", background="#f1f5f9", foreground="#172033", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background="white", foreground="#1e293b", font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", background="white", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("Section.TLabel", background="white", foreground="#0f172a", font=("Segoe UI", 13, "bold"))
        style.configure("Status.TLabel", background="white", foreground="#0f766e", font=("Segoe UI", 9, "bold"))
        style.configure("TEntry", padding=(10, 9), fieldbackground="#f8fafc", foreground="#0f172a")
        style.map("TEntry", fieldbackground=[("focus", "white")])
        style.configure("Accent.TButton", background="#0f766e", foreground="white", padding=(18, 11), font=("Segoe UI", 10, "bold"), borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#115e59"), ("disabled", "#94a3b8")], foreground=[("disabled", "white")])
        style.configure("Secondary.TButton", background="white", foreground="#334155", padding=(14, 10), font=("Segoe UI", 9, "bold"), borderwidth=1)
        style.map("Secondary.TButton", background=[("active", "#f1f5f9"), ("disabled", "#f8fafc")], foreground=[("disabled", "#94a3b8")])

    def _build(self):
        root = ttk.Frame(self, padding=(26, 24, 26, 16))
        root.pack(fill="both", expand=True)

        banner = tk.Frame(root, bg="#10243a", padx=22, pady=18)
        banner.pack(fill="x", pady=(0, 18))
        tk.Label(banner, text="VIN 批量生成器", bg="#10243a", fg="white", font=("Segoe UI", 21, "bold")).pack(anchor="w")
        tk.Label(banner, text="填写车辆固定信息，自动计算校验位并生成连续流水号", bg="#10243a", fg="#cbd5e1", font=("Segoe UI", 10)).pack(anchor="w", pady=(5, 0))

        card = ttk.Frame(root, style="Card.TFrame", padding=(20, 18))
        card.pack(fill="x", pady=(0, 14))
        ttk.Label(card, text="生成参数", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        self.serial = tk.StringVar(value="1")
        self.quantity = tk.StringVar(value="20")
        self.prefix_vars = []
        self.post_vars = []
        self._character_fields(card, "前 8 位主体", self.prefix_vars, 1, 0)
        self._character_fields(card, "第 10–11 位", self.post_vars, 1, 1)
        self._field(card, "流水号起始值", self.serial, "范围 0–999999", 2, 0)
        self._field(card, "生成数量", self.quantity, "每批最多 5,000 条", 2, 1)
        check_hint = ttk.Frame(card, style="Card.TFrame")
        check_hint.grid(row=3, column=0, sticky="w", pady=(2, 0))
        ttk.Label(check_hint, text="第 9 位校验码", style="Hint.TLabel").pack(side="left")
        tk.Label(check_hint, text=" ? ", bg="#e0f2fe", fg="#0369a1", font=("Segoe UI", 11, "bold"), padx=5, pady=1).pack(side="left", padx=(6, 0))
        ttk.Label(card, text="每格输入 1 个字符；I、O、Q 不可使用", style="Hint.TLabel").grid(row=3, column=1, sticky="w", padx=(18, 0), pady=(2, 0))

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(0, 14))
        self.generate_button = ttk.Button(actions, text="生成 VIN", style="Accent.TButton", command=self._generate)
        self.generate_button.pack(side="left")
        self.copy_button = ttk.Button(actions, text="复制全部", style="Secondary.TButton", command=self._copy, state="disabled")
        self.copy_button.pack(side="left", padx=(10, 8))
        self.export_button = ttk.Button(actions, text="导出 Excel（.xlsx）", style="Secondary.TButton", command=self._export, state="disabled")
        self.export_button.pack(side="left")

        result_card = ttk.Frame(root, style="Card.TFrame", padding=(18, 16))
        result_card.pack(fill="both", expand=True)
        result_header = ttk.Frame(result_card, style="Card.TFrame")
        result_header.pack(fill="x", pady=(0, 10))
        ttk.Label(result_header, text="生成结果", style="Section.TLabel").pack(side="left")
        self.status = ttk.Label(result_header, text="等待生成", style="Status.TLabel")
        self.status.pack(side="right")
        text_frame = tk.Frame(result_card, bg="#f8fafc", highlightbackground="#e2e8f0", highlightthickness=1)
        text_frame.pack(fill="both", expand=True)
        self.output = tk.Text(text_frame, font=("Consolas", 11), bg="#f8fafc", fg="#0f172a", insertbackground="#0f766e", relief="flat", padx=14, pady=12, wrap="none", borderwidth=0, state="disabled")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.output.yview)
        self.output.configure(yscrollcommand=scrollbar.set)
        self.output.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        footer = ttk.Frame(root)
        footer.pack(fill="x", pady=(12, 0))
        ttk.Label(footer, text="VIN 共 17 位  ·  前 8 位 + 校验位 + 第 10–11 位 + 6 位流水号", style="Hint.TLabel").pack(side="left")
        ttk.Label(footer, text="此软件由工艺技术部开发", style="Hint.TLabel").pack(side="right")

    @staticmethod
    def _field(parent, label, variable, placeholder, row, col):
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 18, 0), pady=(0, 12))
        ttk.Label(frame, text=label, style="Card.TLabel").pack(anchor="w", pady=(0, 6))
        ttk.Entry(frame, textvariable=variable).pack(fill="x")
        parent.columnconfigure(col, weight=1)

    def _character_fields(self, parent, label, variables, row, col):
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 18, 0), pady=(0, 12))
        ttk.Label(frame, text=label, style="Card.TLabel").pack(anchor="w", pady=(0, 7))
        boxes = ttk.Frame(frame, style="Card.TFrame")
        boxes.pack(anchor="w")
        count = 8 if label == "前 8 位主体" else 2
        for idx in range(count):
            value = tk.StringVar()
            variables.append(value)
            entry = ttk.Entry(boxes, textvariable=value, width=3, justify="center", font=("Segoe UI", 11, "bold"))
            entry.grid(row=0, column=idx, padx=(0 if idx == 0 else 5, 0))
            entry.bind("<KeyRelease>", lambda event, group=variables, pos=idx: self._advance_character(event, group, pos))
            entry.bind("<BackSpace>", lambda event, group=variables, pos=idx: self._previous_character(event, group, pos))
            entry.bind("<<Paste>>", lambda event, group=variables, pos=idx, widget=entry: self.after_idle(lambda: self._spread_pasted_text(group, pos, widget)))
        parent.columnconfigure(col, weight=1)

    @staticmethod
    def _advance_character(event, variables, index):
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Left", "Right", "Tab", "BackSpace", "Delete", "Home", "End"):
            return
        value = variables[index].get().upper()
        if value:
            variables[index].set(value[-1])
            if index + 1 < len(variables):
                event.widget.tk_focusNext().focus_set()

    @staticmethod
    def _previous_character(event, variables, index):
        if not variables[index].get() and index > 0:
            event.widget.tk_focusPrev().focus_set()

    def _spread_pasted_text(self, variables, start_index, widget):
        text = widget.get().upper()
        if len(text) <= 1:
            return
        chars = list(text)
        for offset, char in enumerate(chars):
            target = start_index + offset
            if target >= len(variables):
                break
            variables[target].set(char)
        widget.tk_focusNext().focus_set()

    def _read_character_fields(self, variables):
        return "".join(value.get().strip().upper() for value in variables)

    def _generate(self):
        prefix = self._read_character_fields(self.prefix_vars)
        post = self._read_character_fields(self.post_vars)
        try:
            start, quantity = int(self.serial.get()), int(self.quantity.get())
        except ValueError:
            messagebox.showerror("输入有误", "起始流水号和生成数量必须是整数。")
            return
        if len(prefix) != 8 or any(ch not in ALLOWED for ch in prefix):
            messagebox.showerror("输入有误", "前 8 位必须恰好为 8 个有效 VIN 字符（不可含 I、O、Q）。")
            return
        if len(post) != 2 or any(ch not in ALLOWED for ch in post):
            messagebox.showerror("输入有误", "第 10–11 位必须恰好为 2 个有效 VIN 字符（不可含 I、O、Q）。")
            return
        if start < 0 or start > 999999:
            messagebox.showerror("输入有误", "起始流水号范围为 0–999999。")
            return
        if quantity < 1 or quantity > 5000:
            messagebox.showerror("输入有误", "生成数量范围为 1–5000。")
            return
        if start + quantity - 1 > 999999:
            messagebox.showerror("输入有误", "流水号递增后不能超过 999999。")
            return
        self.records = generate(prefix, post, start, quantity)
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.insert("1.0", "\n".join(self.records))
        self.output.configure(state="disabled")
        self.copy_button.configure(state="normal")
        self.export_button.configure(state="normal")
        self.status.configure(text=f"已生成 {len(self.records):,} 个")

    def _copy(self):
        if not self.records:
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(self.records))
        self.status.configure(text=f"已复制 {len(self.records):,} 个 VIN")

    def _export(self):
        if not self.records:
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 工作簿", "*.xlsx")], initialfile="vin_batch.xlsx")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "VIN列表"
            sheet.append(["VIN"])
            for vin in self.records:
                sheet.append([vin])
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = f"A1:A{len(self.records) + 1}"
            sheet.column_dimensions["A"].width = 24
            sheet["A1"].font = Font(bold=True, color="FFFFFF")
            sheet["A1"].fill = PatternFill("solid", fgColor="0F766E")
            workbook.save(path)
        except (OSError, PermissionError, ValueError) as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self.status.configure(text=f"已导出 Excel：{len(self.records):,} 个 VIN")


if __name__ == "__main__":
    App().mainloop()
