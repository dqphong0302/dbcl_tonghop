#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI Tổng hợp Khảo sát Excel → Word  v2
Refactored: Tab layout, file list có xóa từng dòng, log timestamp, settings persistence.
"""

import sys
import os
import threading
import glob
import json
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from excel_to_word import (
    extract_data_from_excel,
    create_word_report,
    merge_sheets_data,
    parse_open_responses,
    get_sheet_count,
)
from merge_feedbacks import merge_feedbacks

# ═══════════════════════════════════════════════════════════════
# SETTINGS PERSISTENCE
# ═══════════════════════════════════════════════════════════════

SETTINGS_FILE = Path.home() / ".dbcl_tonghop_settings.json"

def load_settings():
    try:
        if SETTINGS_FILE.exists():
            return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def save_settings(data: dict):
    try:
        SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

# ═══════════════════════════════════════════════════════════════
# THEME
# ═══════════════════════════════════════════════════════════════

C = {
    "bg":           "#16172b",
    "bg_card":      "#1e2040",
    "bg_input":     "#252849",
    "bg_row_even":  "#1e2040",
    "bg_row_odd":   "#232548",
    "accent":       "#6c63ff",
    "accent2":      "#4ade80",
    "accent3":      "#f59e0b",
    "text":         "#e2e8f0",
    "text_dim":     "#64748b",
    "text_bright":  "#ffffff",
    "border":       "#2e3260",
    "success":      "#10b981",
    "warning":      "#f59e0b",
    "error":        "#ef4444",
    "log_bg":       "#0f1020",
    "tag_success":  "#4ade80",
    "tag_error":    "#f87171",
    "tag_warn":     "#fbbf24",
    "tag_info":     "#818cf8",
    "tag_dim":      "#475569",
}

FF  = "Segoe UI"
FFM = "Consolas"


# ═══════════════════════════════════════════════════════════════
# FILE LIST WIDGET
# ═══════════════════════════════════════════════════════════════

class FileListWidget(tk.Frame):
    """Danh sách file với nút xóa từng dòng và đếm file."""

    def __init__(self, parent, on_change=None, **kwargs):
        super().__init__(parent, bg=C["bg_card"], **kwargs)
        self._files = []
        self._on_change = on_change
        self._build()

    def _build(self):
        # Toolbar
        bar = tk.Frame(self, bg=C["bg_card"])
        bar.pack(fill="x", pady=(0, 6))

        self._btn_add = self._btn(bar, "➕ Thêm file", self._add_files, C["accent"])
        self._btn_add.pack(side="left")
        self._btn_folder = self._btn(bar, "📁 Thư mục", self._add_folder, C["bg_input"])
        self._btn_folder.pack(side="left", padx=(6, 0))
        self._btn_clear = self._btn(bar, "🗑 Xóa tất cả", self._clear_all, C["bg_input"])
        self._btn_clear.pack(side="left", padx=(6, 0))

        self._count_lbl = tk.Label(bar, text="Chưa có file", font=(FF, 9),
                                   bg=C["bg_card"], fg=C["text_dim"])
        self._count_lbl.pack(side="right")

        # Treeview
        frame = tk.Frame(self, bg=C["border"], bd=1, relief="flat")
        frame.pack(fill="both", expand=True)

        cols = ("name", "size")
        self._tree = ttk.Treeview(frame, columns=cols, show="headings",
                                  height=5, selectmode="extended")
        self._tree.heading("name", text="Tên file")
        self._tree.heading("size", text="Kích thước")
        self._tree.column("name", width=380, stretch=True)
        self._tree.column("size", width=90, anchor="e", stretch=False)

        sb = ttk.Scrollbar(frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True)

        self._tree.bind("<Delete>", lambda e: self._remove_selected())
        self._tree.bind("<BackSpace>", lambda e: self._remove_selected())

        # Remove selected button
        bot = tk.Frame(self, bg=C["bg_card"])
        bot.pack(fill="x", pady=(4, 0))
        self._btn_remove = self._btn(bot, "✕ Xóa dòng đã chọn", self._remove_selected, C["bg_input"])
        self._btn_remove.pack(side="left")
        tk.Label(bot, text="(hoặc nhấn Delete)", font=(FF, 8),
                 bg=C["bg_card"], fg=C["text_dim"]).pack(side="left", padx=6)

    def _btn(self, parent, text, cmd, bg):
        b = ttk.Button(parent, text=text, command=cmd)
        return b

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        added = 0
        for f in files:
            if f not in self._files:
                self._files.append(f)
                added += 1
        if added:
            self._refresh()

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file Excel")
        if folder:
            added = 0
            for ext in ("*.xls", "*.xlsx"):
                for f in sorted(glob.glob(os.path.join(folder, ext))):
                    if f not in self._files:
                        self._files.append(f)
                        added += 1
            if added:
                self._refresh()

    def _remove_selected(self):
        selected = self._tree.selection()
        if not selected:
            return
        indices = sorted([self._tree.index(s) for s in selected], reverse=True)
        for i in indices:
            if 0 <= i < len(self._files):
                del self._files[i]
        self._refresh()

    def _clear_all(self):
        self._files.clear()
        self._refresh()

    def _refresh(self):
        self._tree.delete(*self._tree.get_children())
        for f in self._files:
            size = os.path.getsize(f) if os.path.exists(f) else 0
            size_str = f"{size // 1024} KB" if size >= 1024 else f"{size} B"
            self._tree.insert("", "end", values=(os.path.basename(f), size_str))
        n = len(self._files)
        if n == 0:
            self._count_lbl.config(text="Chưa có file", fg=C["text_dim"])
        else:
            self._count_lbl.config(text=f"📄 {n} file đã chọn", fg=C["accent"])
        if self._on_change:
            self._on_change(n)

    @property
    def files(self):
        return list(self._files)


# ═══════════════════════════════════════════════════════════════
# LOG WIDGET
# ═══════════════════════════════════════════════════════════════

class LogWidget(tk.Frame):
    """Log area với timestamp, màu tag, nút xóa và copy."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C["bg_card"], **kwargs)
        self._build()

    def _build(self):
        bar = tk.Frame(self, bg=C["bg_card"])
        bar.pack(fill="x", pady=(0, 4))
        tk.Label(bar, text="📋 Nhật ký xử lý", font=(FF, 10, "bold"),
                 bg=C["bg_card"], fg=C["text"]).pack(side="left")
        ttk.Button(bar, text="🗑 Xóa", command=self.clear, width=8).pack(side="right")
        ttk.Button(bar, text="📋 Copy", command=self._copy, width=8).pack(side="right", padx=(0, 4))

        frame = tk.Frame(self, bg=C["border"], bd=1, relief="flat")
        frame.pack(fill="both", expand=True)

        self._txt = tk.Text(frame, wrap="word", font=(FFM, 9),
                            bg=C["log_bg"], fg=C["text"],
                            insertbackground=C["text"],
                            selectbackground=C["accent"],
                            relief="flat", bd=8, state="disabled")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self._txt.yview)
        self._txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._txt.pack(fill="both", expand=True)

        for tag, fg in [("success", C["tag_success"]), ("error", C["tag_error"]),
                        ("warning", C["tag_warn"]), ("info", C["tag_info"]),
                        ("dim", C["tag_dim"])]:
            self._txt.tag_configure(tag, foreground=fg)

    def log(self, msg, tag=None):
        ts = datetime.now().strftime("%H:%M:%S")
        def _do():
            self._txt.configure(state="normal")
            self._txt.insert("end", f"[{ts}] ", "dim")
            if tag:
                self._txt.insert("end", msg + "\n", tag)
            else:
                self._txt.insert("end", msg + "\n")
            self._txt.see("end")
            self._txt.configure(state="disabled")
        self._txt.after(0, _do)

    def clear(self):
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.configure(state="disabled")

    def _copy(self):
        content = self._txt.get("1.0", "end").strip()
        self._txt.clipboard_clear()
        self._txt.clipboard_append(content)


# ═══════════════════════════════════════════════════════════════
# OUTPUT DIR ROW
# ═══════════════════════════════════════════════════════════════

class OutputDirRow(tk.Frame):
    def __init__(self, parent, var: tk.StringVar, **kwargs):
        super().__init__(parent, bg=C["bg_card"], **kwargs)
        self._var = var
        self._build()

    def _build(self):
        tk.Label(self, text="📂 Thư mục kết quả:", font=(FF, 10, "bold"),
                 bg=C["bg_card"], fg=C["text"]).pack(anchor="w", pady=(0, 4))
        row = tk.Frame(self, bg=C["bg_card"])
        row.pack(fill="x")
        entry = tk.Entry(row, textvariable=self._var, font=(FF, 10),
                         bg=C["bg_input"], fg=C["text"],
                         insertbackground=C["text"], relief="flat", bd=0)
        entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 8))
        ttk.Button(row, text="Chọn...", command=self._choose, width=10).pack(side="right")

    def _choose(self):
        folder = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if folder:
            self._var.set(folder)


# ═══════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════

class App(tk.Tk):
    APP_VERSION = "v2"

    def __init__(self):
        super().__init__()
        self._settings = load_settings()
        self.title(f"📊 Tổng hợp Khảo sát Excel → Word  ({self.APP_VERSION})")
        self.geometry("860x740")
        self.minsize(720, 620)
        self.configure(bg=C["bg"])
        self._setup_styles()
        self._build_ui()
        self._center_window()
        self._load_settings_to_ui()

    def _resource_path(self, rel):
        try:
            base = sys._MEIPASS
        except AttributeError:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, rel)

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # General
        style.configure(".", background=C["bg"], foreground=C["text"],
                        font=(FF, 10), borderwidth=0, focuscolor=C["accent"])
        style.configure("TFrame", background=C["bg"])
        style.configure("TLabel", background=C["bg"], foreground=C["text"])

        # Notebook
        style.configure("TNotebook", background=C["bg"], borderwidth=0, tabmargins=0)
        style.configure("TNotebook.Tab", background=C["bg_input"], foreground=C["text_dim"],
                        padding=[18, 8], font=(FF, 10))
        style.map("TNotebook.Tab",
                  background=[("selected", C["accent"]), ("active", C["bg_card"])],
                  foreground=[("selected", C["text_bright"]), ("active", C["text"])])

        # Buttons
        style.configure("TButton", background=C["bg_input"], foreground=C["text"],
                        padding=[10, 6], relief="flat", borderwidth=0)
        style.map("TButton",
                  background=[("active", C["bg_card"]), ("disabled", C["bg_input"])],
                  foreground=[("disabled", C["text_dim"])])
        style.configure("Accent.TButton", background=C["accent"], foreground=C["text_bright"],
                        padding=[16, 8], font=(FF, 11, "bold"))
        style.map("Accent.TButton",
                  background=[("active", "#8b83ff"), ("disabled", C["bg_input"])])
        style.configure("Green.TButton", background=C["success"], foreground=C["text_bright"],
                        padding=[16, 8], font=(FF, 11, "bold"))
        style.map("Green.TButton",
                  background=[("active", "#34d399"), ("disabled", C["bg_input"])])

        # Treeview
        style.configure("Treeview", background=C["bg_input"], foreground=C["text"],
                        fieldbackground=C["bg_input"], borderwidth=0,
                        rowheight=26, font=(FF, 9))
        style.configure("Treeview.Heading", background=C["bg_card"], foreground=C["text_dim"],
                        font=(FF, 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", C["accent"])],
                  foreground=[("selected", C["text_bright"])])

        # Progressbar
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor=C["bg_input"], background=C["accent"],
                        borderwidth=0, thickness=10)
        style.configure("Green.Horizontal.TProgressbar",
                        troughcolor=C["bg_input"], background=C["success"],
                        borderwidth=0, thickness=10)

        # Scrollbar
        style.configure("TScrollbar", background=C["bg_card"], troughcolor=C["bg_input"],
                        borderwidth=0, arrowsize=12)

    def _build_ui(self):
        # ── HEADER ──
        hdr = tk.Frame(self, bg=C["bg"], padx=24, pady=14)
        hdr.pack(fill="x")

        tk.Label(hdr, text="📊", font=(FF, 26), bg=C["bg"], fg=C["accent"]).pack(side="left")
        tf = tk.Frame(hdr, bg=C["bg"])
        tf.pack(side="left", padx=(10, 0))
        tk.Label(tf, text="Tổng hợp Khảo sát", font=(FF, 17, "bold"),
                 bg=C["bg"], fg=C["text_bright"]).pack(anchor="w")
        tk.Label(tf, text="Excel (.xls) → Word (.docx)  •  Hỗ trợ nhiều sheet  •  " + self.APP_VERSION,
                 font=(FF, 9), bg=C["bg"], fg=C["text_dim"]).pack(anchor="w")

        sep = tk.Frame(self, bg=C["border"], height=1)
        sep.pack(fill="x")

        # ── NOTEBOOK ──
        nb_frame = tk.Frame(self, bg=C["bg"], padx=20, pady=12)
        nb_frame.pack(fill="both", expand=True)

        self._nb = ttk.Notebook(nb_frame)
        self._nb.pack(fill="both", expand=True)

        self._tab1 = self._build_tab1()
        self._tab2 = self._build_tab2()
        self._nb.add(self._tab1, text="  📄 Tổng hợp báo cáo  ")
        self._nb.add(self._tab2, text="  💬 Gộp Feedbacks  ")

        # ── STATUS BAR ──
        sep2 = tk.Frame(self, bg=C["border"], height=1)
        sep2.pack(fill="x")
        self._status_var = tk.StringVar(value="Sẵn sàng  •  Chọn tab và thêm file Excel để bắt đầu")
        tk.Label(self, textvariable=self._status_var, font=(FF, 9),
                 bg=C["bg"], fg=C["text_dim"], anchor="w", padx=24, pady=6).pack(fill="x")

    # ── TAB 1: TỔNG HỢP BÁO CÁO ──────────────────────────────────

    def _build_tab1(self):
        tab = tk.Frame(self._nb, bg=C["bg"])

        # File list
        card1 = self._card(tab, "📁 File Excel đầu vào")
        self._fl1 = FileListWidget(card1, on_change=self._on_fl1_change)
        self._fl1.pack(fill="both", expand=False, pady=(0, 4))

        # Output dir
        card2 = self._card(tab, "")
        self._outdir1 = tk.StringVar()
        OutputDirRow(card2, self._outdir1).pack(fill="x")
        self._outdir1.trace_add("write", lambda *_: self._save_settings())

        # Action button
        act = tk.Frame(tab, bg=C["bg"])
        act.pack(fill="x", padx=20, pady=(4, 8))
        self._btn1 = ttk.Button(act, text="▶  Bắt đầu xử lý",
                                style="Accent.TButton", command=self._start_tool1)
        self._btn1.pack(side="left")
        self._open1 = ttk.Button(act, text="📂 Mở kết quả", command=lambda: self._open_dir(self._outdir1.get()))
        self._open1.pack(side="left", padx=(8, 0))

        # Progress
        prog_frame = tk.Frame(tab, bg=C["bg"], padx=20)
        prog_frame.pack(fill="x")
        self._prog1_lbl = tk.Label(prog_frame, text="", font=(FF, 9),
                                   bg=C["bg"], fg=C["text_dim"], anchor="w")
        self._prog1_lbl.pack(fill="x")
        self._prog1 = ttk.Progressbar(prog_frame, style="Accent.Horizontal.TProgressbar",
                                      mode="determinate", length=300)
        self._prog1.pack(fill="x", pady=(2, 8))

        # Log
        log_card = tk.Frame(tab, bg=C["bg"], padx=20, pady=0)
        log_card.pack(fill="both", expand=True)
        self._log1 = LogWidget(log_card)
        self._log1.pack(fill="both", expand=True)

        return tab

    # ── TAB 2: GỘP FEEDBACKS ─────────────────────────────────────

    def _build_tab2(self):
        tab = tk.Frame(self._nb, bg=C["bg"])

        info = tk.Frame(tab, bg=C["bg_card"], padx=16, pady=10)
        info.pack(fill="x", padx=20, pady=(12, 0))
        tk.Label(info, text="ℹ️  Tool này đọc cột AK (index 36) từ các file Excel, làm sạch và xuất Feedbacks.docx",
                 font=(FF, 9), bg=C["bg_card"], fg=C["text_dim"], wraplength=700, justify="left").pack(anchor="w")

        card1 = self._card(tab, "📁 File Excel đầu vào")
        self._fl2 = FileListWidget(card1, on_change=self._on_fl2_change)
        self._fl2.pack(fill="both", expand=False, pady=(0, 4))

        card2 = self._card(tab, "")
        self._outdir2 = tk.StringVar()
        OutputDirRow(card2, self._outdir2).pack(fill="x")
        self._outdir2.trace_add("write", lambda *_: self._save_settings())

        act = tk.Frame(tab, bg=C["bg"], padx=20, pady=4)
        act.pack(fill="x")
        self._btn2 = ttk.Button(act, text="💬  Gộp Feedbacks",
                                style="Green.TButton", command=self._start_tool2)
        self._btn2.pack(side="left")
        self._open2 = ttk.Button(act, text="📂 Mở kết quả", command=lambda: self._open_dir(self._outdir2.get()))
        self._open2.pack(side="left", padx=(8, 0))

        prog_frame = tk.Frame(tab, bg=C["bg"], padx=20)
        prog_frame.pack(fill="x")
        self._prog2_lbl = tk.Label(prog_frame, text="", font=(FF, 9),
                                   bg=C["bg"], fg=C["text_dim"], anchor="w")
        self._prog2_lbl.pack(fill="x")
        self._prog2 = ttk.Progressbar(prog_frame, style="Green.Horizontal.TProgressbar",
                                      mode="determinate", length=300)
        self._prog2.pack(fill="x", pady=(2, 8))

        log_card = tk.Frame(tab, bg=C["bg"], padx=20, pady=0)
        log_card.pack(fill="both", expand=True)
        self._log2 = LogWidget(log_card)
        self._log2.pack(fill="both", expand=True)

        return tab

    # ── HELPERS ──────────────────────────────────────────────────

    def _card(self, parent, title):
        outer = tk.Frame(parent, bg=C["bg"], padx=20, pady=4)
        outer.pack(fill="x")
        if title:
            tk.Label(outer, text=title, font=(FF, 10, "bold"),
                     bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(0, 6))
        inner = tk.Frame(outer, bg=C["bg_card"], padx=12, pady=10,
                         highlightbackground=C["border"], highlightthickness=1)
        inner.pack(fill="x")
        return inner

    def _on_fl1_change(self, n):
        self._status_var.set(f"Tab 1: {n} file đã chọn" if n else "Sẵn sàng  •  Chọn tab và thêm file Excel để bắt đầu")

    def _on_fl2_change(self, n):
        self._status_var.set(f"Tab 2: {n} file đã chọn" if n else "Sẵn sàng  •  Chọn tab và thêm file Excel để bắt đầu")

    def _open_dir(self, path):
        if path and os.path.isdir(path):
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        else:
            messagebox.showinfo("Thông báo", "Thư mục kết quả chưa tồn tại.")

    def _ensure_outdir(self, var, files):
        d = var.get().strip()
        if not d and files:
            d = os.path.join(os.path.dirname(files[0]), "output")
            var.set(d)
        os.makedirs(d, exist_ok=True)
        return d

    def _set_progress(self, bar, lbl, value, text=""):
        def _do():
            bar["value"] = value
            if text:
                lbl.config(text=text)
        self.after(0, _do)

    def _set_status(self, text):
        self.after(0, lambda: self._status_var.set(text))

    def _save_settings(self):
        save_settings({
            "outdir1": self._outdir1.get(),
            "outdir2": self._outdir2.get(),
        })

    def _load_settings_to_ui(self):
        self._outdir1.set(self._settings.get("outdir1", ""))
        self._outdir2.set(self._settings.get("outdir2", ""))


    # ── TOOL 1 WORKER ────────────────────────────────────────────

    def _start_tool1(self):
        files = self._fl1.files
        if not files:
            messagebox.showwarning("Chưa có file", "Vui lòng thêm file Excel trước khi xử lý.")
            return
        out_dir = self._ensure_outdir(self._outdir1, files)
        self._btn1.state(["disabled"])
        self._log1.clear()
        self._prog1["value"] = 0
        threading.Thread(target=self._worker_tool1, args=(files, out_dir), daemon=True).start()

    def _worker_tool1(self, files, out_dir):
        total = len(files)
        success = fail = total_out = 0
        self._log1.log(f"{'═'*48}", "dim")
        self._log1.log(f"Bắt đầu xử lý {total} file Excel", "info")
        self._log1.log(f"Thư mục kết quả: {out_dir}", "dim")
        self._log1.log(f"{'═'*48}", "dim")

        for fi, filepath in enumerate(files):
            name = os.path.basename(filepath)
            pct = int(fi / total * 100)
            self._set_progress(self._prog1, self._prog1_lbl, pct,
                               f"Đang xử lý {fi+1}/{total}: {name[:55]}...")
            self._set_status(f"Đang xử lý file {fi+1}/{total}...")
            self._log1.log(f"", "dim")
            self._log1.log(f"📄 {name}", "info")

            try:
                n_sheets, sheet_names = get_sheet_count(filepath)
                basename = os.path.splitext(name)[0]
                all_data = []
                missing_sheets = []

                for si in range(n_sheets):
                    try:
                        data = extract_data_from_excel(filepath, sheet_index=si)
                    except Exception as e:
                        self._log1.log(f'   ⚠ Sheet {si+1} "{sheet_names[si]}": lỗi đọc – {e}', "warning")
                        continue

                    if not data["survey_questions"]:
                        self._log1.log(f'   ⏭ Sheet {si+1} "{sheet_names[si]}": bỏ qua (không có dữ liệu)', "dim")
                        continue

                    lop = data["meta"].get("lop_hoc", sheet_names[si])
                    if data["meta"].get("tong_so_sv") is None:
                        missing_sheets.append(f"Sheet {si+1} \"{lop}\"")
                        self._log1.log(f"   ❌ Sheet {si+1} \"{lop}\": Không xác định được Tổng SV → bỏ qua", "error")
                        continue

                    all_data.append(data)
                    n_q = len(data["survey_questions"])

                    if n_sheets > 1:
                        out_path = os.path.join(out_dir, f"{basename}_sheet{si+1}.docx")
                        create_word_report(data, out_path)
                        total_out += 1
                        self._log1.log(f"   ✅ Sheet {si+1} \"{lop}\": {n_q} câu, {data['meta']['tong_so_sv']} SV → {os.path.basename(out_path)}", "success")
                    else:
                        out_path = os.path.join(out_dir, f"{basename}.docx")
                        create_word_report(data, out_path)
                        total_out += 1
                        n_resp = len(parse_open_responses(data["open_responses"]))
                        self._log1.log(f"   ✅ {n_q} câu, {n_resp} ý kiến, {data['meta']['tong_so_sv']} SV → {os.path.basename(out_path)}", "success")

                if missing_sheets:
                    sheets_str = ", ".join(missing_sheets)
                    msg = (f"File: {name}\n\nCác sheet sau thiếu Tổng số SV:\n{sheets_str}\n\n"
                           f"Vui lòng ghi số SV không tham gia ở cuối tên file\n"
                           f"hoặc điền vào ô F5 trong Excel rồi thử lại.")
                    self.after(0, lambda m=msg: messagebox.showwarning("Thiếu thông tin Tổng SV", m))

                if len(all_data) > 1:
                    merged = merge_sheets_data(all_data)
                    if merged:
                        sum_path = os.path.join(out_dir, f"{basename}_TONG_HOP.docx")
                        create_word_report(merged, sum_path)
                        total_out += 1
                        total_sv = merged["meta"].get("tong_so_sv", 0)
                        self._log1.log(f"   📊 TỔNG HỢP: {len(all_data)} lớp, {total_sv} SV → {os.path.basename(sum_path)}", "info")

                success += 1
            except Exception as e:
                fail += 1
                self._log1.log(f"   ❌ Lỗi: {e}", "error")

        self._set_progress(self._prog1, self._prog1_lbl, 100, "Hoàn tất!")
        self._log1.log(f"", "dim")
        self._log1.log(f"{'═'*48}", "dim")
        self._log1.log(f"✅ Hoàn tất: {success} thành công, {fail} thất bại  •  {total_out} file tạo ra", "success")
        self._log1.log(f"📂 Kết quả: {out_dir}", "dim")
        self._set_status(f"✅ Hoàn tất  •  {total_out} file tạo ra  •  {out_dir}")
        self.after(0, lambda: self._btn1.state(["!disabled"]))

    # ── TOOL 2 WORKER ────────────────────────────────────────────

    def _start_tool2(self):
        files = self._fl2.files
        if not files:
            messagebox.showwarning("Chưa có file", "Vui lòng thêm file Excel trước khi xử lý.")
            return
        out_dir = self._ensure_outdir(self._outdir2, files)
        self._btn2.state(["disabled"])
        self._log2.clear()
        self._prog2["value"] = 0
        threading.Thread(target=self._worker_tool2, args=(files, out_dir), daemon=True).start()

    def _worker_tool2(self, files, out_dir):
        self._log2.log(f"{'═'*48}", "dim")
        self._log2.log(f"Bắt đầu gộp feedbacks từ {len(files)} file", "info")
        self._log2.log(f"{'═'*48}", "dim")
        self._set_progress(self._prog2, self._prog2_lbl, 30, "Đang đọc dữ liệu...")
        self._set_status("Đang tổng hợp feedbacks...")

        try:
            for f in files:
                self._log2.log(f"📄 Đọc: {os.path.basename(f)}", "dim")
            output_path = merge_feedbacks(files, out_dir)
            self._set_progress(self._prog2, self._prog2_lbl, 100, "Hoàn tất!")
            if output_path:
                self._log2.log(f"", "dim")
                self._log2.log(f"✅ Hoàn tất gộp feedbacks", "success")
                self._log2.log(f"📄 File kết quả: {os.path.basename(output_path)}", "info")
                self._log2.log(f"📂 Thư mục: {out_dir}", "dim")
                self._set_status(f"✅ Hoàn tất gộp feedbacks  •  {os.path.basename(output_path)}")
            else:
                self._log2.log("⚠️ Không tìm thấy feedback nào trong cột AK", "warning")
                self._set_status("⚠️ Không có dữ liệu feedback")
        except Exception as e:
            self._log2.log(f"❌ Lỗi: {e}", "error")
            self._set_status("❌ Lỗi xử lý")
            self._set_progress(self._prog2, self._prog2_lbl, 0, "")

        self.after(0, lambda: self._btn2.state(["!disabled"]))


# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
