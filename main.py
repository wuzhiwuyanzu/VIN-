"""VIN batch generator for Windows; standard-library-only Tkinter application."""

import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


TRANSLITERATION = {
    **dict.fromkeys("ABCDEFG", 1),
    **dict.fromkeys("H", 8),
    **dict.fromkeys("JKLM", 2),
    **dict.fromkeys("N", 5),
    **dict.fromkeys("P", 7),
    **dict.fromkeys("R", 9),
    **dict.fromkeys("ST", 3),
    **dict.fromkeys("UVW", 4),
    **dict.fromkeys("XYZ", 6),
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
        self.geometry("820x680")
        self.minsize(700, 560)
        self.configure(bg="#f4f6fa")
        self.records: list[str] = []
        self._style()
        self._build()

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6fa")
        style.configure("Card.TFrame", background="white")
        style.configure("TLabel", background="#f4f6fa", foreground="#172033", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground="#172033")
        style.configure("Sub.TLabel", foreground="#64748b")
        style.configure("Card.TLabel", background="white", font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", background="white", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("TEntry", padding=9, fieldbackground="white")
        style.configure("Accent.TButton", background="#2563eb", foreground="white", padding=(16, 10), font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#1d4ed8")])
        style.configure("TButton", padding=(12, 9), font=("Segoe UI", 9))

    def _build(self):
        root = ttk.Frame(self, padding=24)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="VIN 批量生成器", style="Title.TLabel").pack(anchor="w")
        ttk.Label(root, text="按固定信息生成 VIN，并自动计算校验位", style="Sub.TLabel").pack(anchor="w", pady=(3, 18))

        card = ttk.Frame(root, style="Card.TFrame", padding=18)
        card.pack(fill="x")
        self.prefix = tk.StringVar()
        self.post = tk.StringVar()
        self.serial = tk.StringVar(value="1")
        self.quantity = tk.StringVar(value="20")
        self._field(card, "前 8 位主体", self.prefix, "例如 1HGCM826", 0, 0)
        self._field(card, "第 10–11 位", self.post, "例如 5A", 0, 1)
        ttk.Label(card, text="校验位", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(16, 6))
        ttk.Label(card, text="自动计算", style="Hint.TLabel").grid(row=3, column=0, sticky="w")
        self._field(card, "流水号起始值", self.serial, "例如 1", 2, 1)
        self._field(card, "生成数量", self.quantity, "最多 5,000", 4, 0)
        ttk.Label(card, text="I、O、Q 不可使用；末尾 6 位由流水号生成", style="Hint.TLabel").grid(row=5, column=0, columnspan=2, sticky="w", pady=(12, 0))

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=14)
        ttk.Button(actions, text="生成并复核", style="Accent.TButton", command=self._generate).pack(side="left")
        ttk.Button(actions, text="复制全部", command=self._copy).pack(side="left", padx=8)
        ttk.Button(actions, text="导出 CSV…", command=self._export).pack(side="left")
        self.status = ttk.Label(actions, text="填写前 8 位主体和流水号范围", style="Sub.TLabel")
        self.status.pack(side="right", pady=8)

        result_card = ttk.Frame(root, style="Card.TFrame", padding=14)
        result_card.pack(fill="both", expand=True)
        ttk.Label(result_card, text="生成结果", style="Card.TLabel").pack(anchor="w", pady=(0, 8))
        self.output = tk.Text(result_card, font=("Consolas", 10), bg="#fbfcfe", fg="#172033", relief="flat", padx=10, pady=8, wrap="none")
        self.output.pack(fill="both", expand=True)
        ttk.Label(root, text="VIN 共 17 位：前 8 位 + 自动校验位 + 第 10–11 位 + 6 位流水号", style="Sub.TLabel").pack(anchor="w", pady=(12, 0))

    @staticmethod
    def _field(parent, label, variable, placeholder, row, col):
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 18, 0), pady=(0, 5))
        ttk.Label(frame, text=label, style="Card.TLabel").pack(anchor="w", pady=(0, 6))
        ttk.Entry(frame, textvariable=variable).pack(fill="x")
        parent.columnconfigure(col, weight=1)

    def _generate(self):
        prefix = self.prefix.get().strip().upper()
        post = self.post.get().strip().upper()
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
        self.output.delete("1.0", "end")
        self.output.insert("1.0", "\n".join(self.records))
        self.status.configure(text=f"已生成并复核 {len(self.records):,} 个 VIN")

    def _copy(self):
        if not self.records:
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(self.records))
        self.status.configure(text=f"已复制 {len(self.records):,} 个 VIN")

    def _export(self):
        if not self.records:
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV 文件", "*.csv")], initialfile="vin_batch.csv")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.writer(handle)
                writer.writerow(["VIN"])
                writer.writerows([vin] for vin in self.records)
        except OSError as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self.status.configure(text=f"已导出 {len(self.records):,} 个 VIN")


if __name__ == "__main__":
    App().mainloop()
