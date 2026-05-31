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
from app_logging import get_logger, attach_gui_callback, get_log_file_path
import analytics
import report_analytics
import benchmark
import history_store
import feedback_classifier
import charts as charts_mod

_LOG = get_logger()

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
        
        # Cấu hình variables cho Word Output
        self._university_var = tk.StringVar()
        self._faculty_var = tk.StringVar()
        self._country_header_var = tk.StringVar()
        self._country_slogan_var = tk.StringVar()
        self._font_name_var = tk.StringVar()
        self._font_size_var = tk.StringVar()
        self._margin_top_var = tk.StringVar()
        self._margin_bottom_var = tk.StringVar()
        self._margin_left_var = tk.StringVar()
        self._margin_right_var = tk.StringVar()
        self._table_header_bg_var = tk.StringVar()
        self._open_response_lines_var = tk.StringVar()
        
        # Cấu hình variables cho Feedback Keywords
        self._feedback_search_keywords_var = tk.StringVar()
        self._feedback_skip_keywords_var = tk.StringVar()
        
        self.title(f"📊 Tổng hợp Khảo sát Excel → Word  ({self.APP_VERSION})")
        self.geometry("960x840")
        self.minsize(860, 720)
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
        self._tab4 = self._build_tab4()
        self._tab3 = self._build_tab3()
        self._nb.add(self._tab1, text="  📄 Tổng hợp báo cáo  ")
        self._nb.add(self._tab2, text="  💬 Gộp Feedbacks  ")
        # self._nb.add(self._tab4, text="  📈 Phân tích & Xu hướng  ") # Đã ẩn tính năng phân tích & xu hướng theo yêu cầu
        self._nb.add(self._tab3, text="  ⚙️ Cấu hình  ")

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

    # ── TAB 4: PHÂN TÍCH & XU HƯỚNG ──────────────────────────────

    def _build_tab4(self):
        tab = tk.Frame(self._nb, bg=C["bg"])

        info = tk.Frame(tab, bg=C["bg_card"], padx=16, pady=10)
        info.pack(fill="x", padx=20, pady=(12, 0))
        tk.Label(info,
                 text=("ℹ️  Mỗi lần tạo báo cáo, phần mềm tự lưu chỉ số (điểm TB, % hài lòng) theo học phần & kỳ. "
                       "Chọn học phần để xem xu hướng qua các kỳ và xuất biểu đồ minh chứng cải tiến."),
                 font=(FF, 9), bg=C["bg_card"], fg=C["text_dim"],
                 wraplength=760, justify="left").pack(anchor="w")

        card1 = self._card(tab, "📚 Chọn học phần đã có lịch sử")
        row = tk.Frame(card1, bg=C["bg_card"])
        row.pack(fill="x", pady=(0, 4))
        self._course_var = tk.StringVar()
        self._course_combo = ttk.Combobox(row, textvariable=self._course_var,
                                           state="readonly", width=48)
        self._course_combo.pack(side="left")
        ttk.Button(row, text="🔄 Tải lại", command=self._refresh_courses).pack(side="left", padx=(8, 0))
        self._course_combo.bind("<<ComboboxSelected>>", lambda *_: self._show_course_history())

        card2 = self._card(tab, "")
        self._outdir4 = tk.StringVar()
        OutputDirRow(card2, self._outdir4).pack(fill="x")
        self._outdir4.trace_add("write", lambda *_: self._save_settings())

        act = tk.Frame(tab, bg=C["bg"], padx=20, pady=4)
        act.pack(fill="x")
        self._btn4 = ttk.Button(act, text="📈  Xuất biểu đồ xu hướng",
                                style="Green.TButton", command=self._start_tool4)
        self._btn4.pack(side="left")
        ttk.Button(act, text="📂 Mở kết quả",
                   command=lambda: self._open_dir(self._outdir4.get())).pack(side="left", padx=(8, 0))

        log_card = tk.Frame(tab, bg=C["bg"], padx=20, pady=0)
        log_card.pack(fill="both", expand=True)
        self._log4 = LogWidget(log_card)
        self._log4.pack(fill="both", expand=True)

        self._refresh_courses()
        return tab

    def _refresh_courses(self):
        try:
            courses = history_store.list_courses()
        except Exception:
            courses = []
        self._course_combo["values"] = courses
        if courses and not self._course_var.get():
            self._course_var.set(courses[0])
        if hasattr(self, "_log4"):
            self._show_course_history()

    def _show_course_history(self):
        course = self._course_var.get()
        if not course or not hasattr(self, "_log4"):
            return
        self._log4.clear()
        records = history_store.get_history(course)
        self._log4.log(f"Học phần: {course}", "info")
        self._log4.log(f"Số kỳ đã ghi nhận: {len(records)}", "dim")
        self._log4.log("─" * 48, "dim")
        for r in records:
            self._log4.log(
                f"• {r.get('period')}: TB {r.get('overall_average')}/5, "
                f"hài lòng {r.get('overall_satisfied_pct')}%, "
                f"{r.get('n_responses')} phiếu", "info")
        if len(records) < 2:
            self._log4.log("", "dim")
            self._log4.log("⚠ Cần ít nhất 2 kỳ để vẽ biểu đồ xu hướng.", "warning")

    def _start_tool4(self):
        course = self._course_var.get()
        if not course:
            messagebox.showwarning("Chưa chọn học phần", "Vui lòng chọn một học phần.")
            return
        out_dir = self._outdir4.get().strip()
        if not out_dir:
            messagebox.showwarning("Thiếu thư mục", "Vui lòng chọn thư mục kết quả.")
            return
        records = history_store.get_history(course)
        if len(records) < 2:
            messagebox.showinfo("Chưa đủ dữ liệu",
                                "Cần ít nhất 2 kỳ khảo sát của học phần này để vẽ xu hướng.")
            return
        self._btn4.state(["disabled"])
        threading.Thread(target=self._worker_tool4, args=(course, out_dir), daemon=True).start()

    def _worker_tool4(self, course, out_dir):
        try:
            os.makedirs(out_dir, exist_ok=True)
            trend = history_store.history_for_trend(course)
            safe = "".join(ch if ch.isalnum() else "_" for ch in course)[:40]
            out_path = os.path.join(out_dir, f"XuHuong_{safe}.png")
            res = charts_mod.chart_trend(
                trend, out_path, title=f"Xu hướng khảo sát - {course}")
            if res:
                first, last = trend[0], trend[-1]
                delta = round(last["overall_average"] - first["overall_average"], 2)
                self._log4.log(f"✅ Đã xuất biểu đồ: {os.path.basename(out_path)}", "success")
                self._log4.log(
                    f"   Thay đổi điểm TB: {first['overall_average']} → "
                    f"{last['overall_average']} ({'+' if delta >= 0 else ''}{delta})",
                    "info")
                self._set_status(f"✅ Đã xuất biểu đồ xu hướng • {out_dir}")
            else:
                self._log4.log("⚠ Không tạo được biểu đồ (cần ≥2 kỳ).", "warning")
        except Exception as e:
            _LOG.error("Lỗi xuất xu hướng: %s", e)
            self._log4.log(f"❌ Lỗi: {e}", "error")
        finally:
            self.after(0, lambda: self._btn4.state(["!disabled"]))

    # ── TAB 3: CẤU HÌNH ──────────────────────────────────────────


    def _build_tab3(self):
        tab = tk.Frame(self._nb, bg=C["bg"])
        
        # Tạo canvas cuộn cho giao diện cấu hình
        canvas = tk.Canvas(tab, bg=C["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=C["bg"])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def _configure_canvas(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind('<Configure>', _configure_canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bắt sự kiện MouseWheel
        def _on_mousewheel(event):
            if sys.platform == "darwin":
                canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            
        tab.bind("<Enter>", _bind_mousewheel)
        tab.bind("<Leave>", _unbind_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Helper widgets thiết kế GUI Cấu hình
        def create_card(title):
            outer = tk.Frame(scrollable_frame, bg=C["bg"], padx=20, pady=8)
            outer.pack(fill="x")
            tk.Label(outer, text=title, font=(FF, 11, "bold"), bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(0, 6))
            inner = tk.Frame(outer, bg=C["bg_card"], padx=18, pady=14,
                             highlightbackground=C["border"], highlightthickness=1)
            inner.pack(fill="both", expand=True)
            return inner
            
        # ── CARD 1: CẤU HÌNH ĐỊNH DẠNG FILE WORD ──
        c1 = create_card("📝 Định dạng file Word xuất ra")
        c1.columnconfigure(1, weight=1)
        c1.columnconfigure(3, weight=1)
        
        # Hàng 0: Tên trường & Tên khoa
        tk.Label(c1, text="Tên trường:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
        entry_uni = tk.Entry(c1, textvariable=self._university_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_uni.grid(row=0, column=1, sticky="ew", pady=6, ipady=4, padx=(0, 20))
        
        tk.Label(c1, text="Tên khoa/Đơn vị:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=0, column=2, sticky="w", pady=6, padx=(0, 10))
        entry_fac = tk.Entry(c1, textvariable=self._faculty_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_fac.grid(row=0, column=3, sticky="ew", pady=6, ipady=4)
        
        # Hàng 1: Quốc hiệu & Tiêu ngữ
        tk.Label(c1, text="Quốc hiệu:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 10))
        entry_qg = tk.Entry(c1, textvariable=self._country_header_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_qg.grid(row=1, column=1, sticky="ew", pady=6, ipady=4, padx=(0, 20))
        
        tk.Label(c1, text="Tiêu ngữ:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=1, column=2, sticky="w", pady=6, padx=(0, 10))
        entry_sl = tk.Entry(c1, textvariable=self._country_slogan_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_sl.grid(row=1, column=3, sticky="ew", pady=6, ipady=4)
        
        # Hàng 2: Tiêu đề báo cáo (nhiều dòng)
        tk.Label(c1, text="Tiêu đề báo cáo:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=2, column=0, sticky="nw", pady=10, padx=(0, 10))
        self._report_title_txt = tk.Text(c1, font=(FF, 10), height=3, bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        self._report_title_txt.grid(row=2, column=1, columnspan=3, sticky="ew", pady=6)
        
        # Hàng 3: Font chữ & Cỡ chữ
        tk.Label(c1, text="Font chữ:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=3, column=0, sticky="w", pady=6, padx=(0, 10))
        font_box = ttk.Combobox(c1, textvariable=self._font_name_var, values=["Times New Roman", "Arial", "Calibri", "Segoe UI"], state="readonly", font=(FF, 10))
        font_box.grid(row=3, column=1, sticky="ew", pady=6, padx=(0, 20))
        
        tk.Label(c1, text="Cỡ chữ (pt):", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=3, column=2, sticky="w", pady=6, padx=(0, 10))
        size_spin = tk.Spinbox(c1, textvariable=self._font_size_var, from_=8, to=24, increment=1, font=(FF, 10), bg=C["bg_input"], fg=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        size_spin.grid(row=3, column=3, sticky="ew", pady=6, ipady=2)
        
        # Hàng 4: Căn lề lề (Trái, Phải, Trên, Dưới)
        tk.Label(c1, text="Căn lề (cm):", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=4, column=0, sticky="w", pady=6, padx=(0, 10))
        margin_frame = tk.Frame(c1, bg=C["bg_card"])
        margin_frame.grid(row=4, column=1, columnspan=3, sticky="ew", pady=6)
        margin_frame.columnconfigure((1, 3, 5, 7), weight=1)
        
        tk.Label(margin_frame, text="Trái:", font=(FF, 9), bg=C["bg_card"], fg=C["text_dim"]).grid(row=0, column=0, padx=(0, 4))
        entry_ml = tk.Entry(margin_frame, textvariable=self._margin_left_var, width=6, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_ml.grid(row=0, column=1, sticky="ew", padx=(0, 15), ipady=3)
        
        tk.Label(margin_frame, text="Phải:", font=(FF, 9), bg=C["bg_card"], fg=C["text_dim"]).grid(row=0, column=2, padx=(0, 4))
        entry_mr = tk.Entry(margin_frame, textvariable=self._margin_right_var, width=6, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_mr.grid(row=0, column=3, sticky="ew", padx=(0, 15), ipady=3)
        
        tk.Label(margin_frame, text="Trên:", font=(FF, 9), bg=C["bg_card"], fg=C["text_dim"]).grid(row=0, column=4, padx=(0, 4))
        entry_mt = tk.Entry(margin_frame, textvariable=self._margin_top_var, width=6, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_mt.grid(row=0, column=5, sticky="ew", padx=(0, 15), ipady=3)
        
        tk.Label(margin_frame, text="Dưới:", font=(FF, 9), bg=C["bg_card"], fg=C["text_dim"]).grid(row=0, column=6, padx=(0, 4))
        entry_mb = tk.Entry(margin_frame, textvariable=self._margin_bottom_var, width=6, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_mb.grid(row=0, column=7, sticky="ew", ipady=3)
        
        # Hàng 5: Màu nền bảng & Số dòng trống ý kiến
        tk.Label(c1, text="Màu nền bảng (HEX):", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=5, column=0, sticky="w", pady=6, padx=(0, 10))
        entry_bg = tk.Entry(c1, textvariable=self._table_header_bg_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_bg.grid(row=5, column=1, sticky="ew", pady=6, ipady=4, padx=(0, 20))
        
        tk.Label(c1, text="Dòng trống ý kiến:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=5, column=2, sticky="w", pady=6, padx=(0, 10))
        lines_spin = tk.Spinbox(c1, textvariable=self._open_response_lines_var, from_=0, to=20, increment=1, font=(FF, 10), bg=C["bg_input"], fg=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        lines_spin.grid(row=5, column=3, sticky="ew", pady=6, ipady=2)
        
        # ── CARD 2: CẤU HÌNH TỪ KHÓA FEEDBACK ──
        c2 = create_card("💬 Cấu hình Từ khóa Feedback")
        c2.columnconfigure(1, weight=1)
        
        # Hàng 0: Cột feedback tìm kiếm
        tk.Label(c2, text="Tìm cột feedback:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
        entry_sf = tk.Entry(c2, textvariable=self._feedback_search_keywords_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_sf.grid(row=0, column=1, sticky="ew", pady=6, ipady=4)
        tk.Label(c2, text="(Cách nhau bằng dấu phẩy. Ví dụ: ý kiến đóng góp, feedback, nhận xét)", font=(FF, 8), bg=C["bg_card"], fg=C["text_dim"]).grid(row=1, column=1, sticky="w", pady=(0, 6))
        
        # Hàng 2: Từ khóa bỏ qua
        tk.Label(c2, text="Từ khóa bỏ qua:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=2, column=0, sticky="w", pady=6, padx=(0, 10))
        entry_kf = tk.Entry(c2, textvariable=self._feedback_skip_keywords_var, font=(FF, 10), bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", bd=0, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        entry_kf.grid(row=2, column=1, sticky="ew", pady=6, ipady=4)
        tk.Label(c2, text="(Cách nhau bằng dấu phẩy. Loại bỏ các phản hồi khớp hoàn toàn với cụm từ này)", font=(FF, 8), bg=C["bg_card"], fg=C["text_dim"]).grid(row=3, column=1, sticky="w", pady=(0, 6))
        
        # Hàng 4: Regex bỏ qua (nhiều dòng)
        tk.Label(c2, text="Regex bỏ qua:", font=(FF, 9, "bold"), bg=C["bg_card"], fg=C["text"]).grid(row=4, column=0, sticky="nw", pady=10, padx=(0, 10))
        self._feedback_skip_patterns_txt = tk.Text(c2, font=(FF, 10), height=4, bg=C["bg_input"], fg=C["text"], insertbackground=C["text"], relief="flat", highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        self._feedback_skip_patterns_txt.grid(row=4, column=1, sticky="ew", pady=6)
        tk.Label(c2, text="(Mỗi dòng một biểu thức chính quy. Dùng để lọc bỏ các phản hồi trùng khớp định dạng)", font=(FF, 8), bg=C["bg_card"], fg=C["text_dim"]).grid(row=5, column=1, sticky="w", pady=(0, 6))
        
        # ── HÀNH ĐỘNG ──
        act_frame = tk.Frame(scrollable_frame, bg=C["bg"], padx=20, pady=12)
        act_frame.pack(fill="x")
        
        btn_save = ttk.Button(act_frame, text="💾 Lưu Cấu Hinh", style="Accent.TButton", command=self._save_config_from_ui)
        btn_save.pack(side="left")
        
        btn_reset = ttk.Button(act_frame, text="🔄 Khôi phục mặc định", command=self._reset_config_to_defaults)
        btn_reset.pack(side="left", padx=(10, 0))
        
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
        self._settings["outdir1"] = self._outdir1.get()
        self._settings["outdir2"] = self._outdir2.get()
        save_settings(self._settings)

    def _load_settings_to_ui(self):
        self._outdir1.set(self._settings.get("outdir1", ""))
        self._outdir2.set(self._settings.get("outdir2", ""))
        
        # Tải cấu hình word_format
        wf = self._settings.get("word_format", {})
        from excel_to_word import CONFIG
        
        self._university_var.set(wf.get("university_name", CONFIG["university_name"]))
        self._faculty_var.set(wf.get("faculty_name", CONFIG["faculty_name"]))
        self._country_header_var.set(wf.get("country_header", CONFIG["country_header"]))
        self._country_slogan_var.set(wf.get("country_slogan", CONFIG["country_slogan"]))
        
        title_lines = CONFIG["report_title_lines"]
        default_title = "\n".join(title_lines)
        report_title_val = wf.get("report_title", default_title)
        self._report_title_txt.delete("1.0", "end")
        self._report_title_txt.insert("1.0", report_title_val)
        
        self._font_name_var.set(wf.get("font_name", CONFIG["font_name"]))
        self._font_size_var.set(str(wf.get("font_size", CONFIG["font_size"])))
        self._margin_top_var.set(str(wf.get("margin_top", CONFIG["margin_top"])))
        self._margin_bottom_var.set(str(wf.get("margin_bottom", CONFIG["margin_bottom"])))
        self._margin_left_var.set(str(wf.get("margin_left", CONFIG["margin_left"])))
        self._margin_right_var.set(str(wf.get("margin_right", CONFIG["margin_right"])))
        self._table_header_bg_var.set(wf.get("table_header_bg", CONFIG["table_header_bg"]))
        self._open_response_lines_var.set(str(wf.get("open_response_lines", CONFIG["open_response_lines"])))
        
        # Tải cấu hình feedback_config
        fc = self._settings.get("feedback_config", {})
        
        search_kw = fc.get("search_keywords", CONFIG["feedback_search_keywords"])
        if isinstance(search_kw, list):
            search_kw = ", ".join(search_kw)
        self._feedback_search_keywords_var.set(search_kw)
        
        skip_kw = fc.get("skip_keywords", CONFIG["feedback_skip_keywords"])
        if isinstance(skip_kw, list):
            skip_kw = ", ".join(skip_kw)
        self._feedback_skip_keywords_var.set(skip_kw)
        
        skip_pat = fc.get("skip_patterns", CONFIG["feedback_skip_patterns"])
        if isinstance(skip_pat, list):
            skip_pat = "\n".join(skip_pat)
        self._feedback_skip_patterns_txt.delete("1.0", "end")
        self._feedback_skip_patterns_txt.insert("1.0", skip_pat)

    def _save_config_from_ui(self):
        try:
            margin_left = float(self._margin_left_var.get())
            margin_right = float(self._margin_right_var.get())
            margin_top = float(self._margin_top_var.get())
            margin_bottom = float(self._margin_bottom_var.get())
            open_lines = int(self._open_response_lines_var.get())
            font_size = int(self._font_size_var.get())
        except ValueError:
            messagebox.showerror("Lỗi nhập liệu", "Vui lòng kiểm tra lại các trường số (Lề trang, Cỡ chữ, Số dòng trống) phải là số hợp lệ.")
            return
            
        search_kw = [k.strip() for k in self._feedback_search_keywords_var.get().split(",") if k.strip()]
        skip_kw = [k.strip() for k in self._feedback_skip_keywords_var.get().split(",") if k.strip()]
        skip_pat = [p.strip() for p in self._feedback_skip_patterns_txt.get("1.0", "end").split("\n") if p.strip()]
        report_title = self._report_title_txt.get("1.0", "end").strip()
        
        self._settings["word_format"] = {
            "university_name": self._university_var.get().strip(),
            "faculty_name": self._faculty_var.get().strip(),
            "country_header": self._country_header_var.get().strip(),
            "country_slogan": self._country_slogan_var.get().strip(),
            "report_title": report_title,
            "font_name": self._font_name_var.get().strip(),
            "font_size": font_size,
            "margin_top": margin_top,
            "margin_bottom": margin_bottom,
            "margin_left": margin_left,
            "margin_right": margin_right,
            "table_header_bg": self._table_header_bg_var.get().strip(),
            "open_response_lines": open_lines,
        }
        
        self._settings["feedback_config"] = {
            "search_keywords": search_kw,
            "skip_keywords": skip_kw,
            "skip_patterns": skip_pat,
        }
        
        save_settings(self._settings)
        
        # Reload vào excel_to_word config
        from excel_to_word import load_settings_into_config
        load_settings_into_config()
        
        messagebox.showinfo("Thành công", "Đã lưu cấu hình thành công!")

    def _reset_config_to_defaults(self):
        if not messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn khôi phục cấu hình về mặc định của nhà trường?"):
            return
            
        if "word_format" in self._settings:
            del self._settings["word_format"]
        if "feedback_config" in self._settings:
            del self._settings["feedback_config"]
            
        save_settings(self._settings)
        
        # Reload defaults vào CONFIG
        from excel_to_word import load_settings_into_config
        load_settings_into_config()
        
        # Cập nhật UI
        self._load_settings_to_ui()
        messagebox.showinfo("Thành công", "Đã khôi phục cấu hình mặc định!")


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
        _start_time = datetime.now()
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

                analysis_source = None
                if len(all_data) > 1:
                    merged = merge_sheets_data(all_data)
                    if merged:
                        sum_path = os.path.join(out_dir, f"{basename}_TONG_HOP.docx")
                        create_word_report(merged, sum_path)
                        total_out += 1
                        total_sv = merged["meta"].get("tong_so_sv", 0)
                        self._log1.log(f"   📊 TỔNG HỢP: {len(all_data)} lớp, {total_sv} SV → {os.path.basename(sum_path)}", "info")
                        analysis_source = merged
                elif len(all_data) == 1:
                    analysis_source = all_data[0]

                # Báo cáo phân tích kèm biểu đồ (Round 2 - số liệu hiệu quả SKKN)
                if False:  # Đã lược bỏ tính năng phân tích định lượng nâng cao theo yêu cầu
                    try:
                        a = analytics.analyze_survey(analysis_source)
                        ana_path = os.path.join(out_dir, f"{basename}_PHAN_TICH.docx")
                        report_analytics.create_analysis_report(
                            a, ana_path,
                            chart_dir=os.path.join(out_dir, "_charts"))
                        total_out += 1
                        self._log1.log(
                            f"   📈 PHÂN TÍCH: TB {a.overall_average:.2f}/5, "
                            f"hài lòng {a.overall_satisfied_pct:.1f}% → {os.path.basename(ana_path)}",
                            "success")
                        # Lưu lịch sử để so sánh xu hướng giữa các kỳ (Round 3)
                        try:
                            rec = history_store.record_analysis(
                                a, meta=analysis_source.get("meta", {}))
                            self._log1.log(
                                f"   💾 Đã lưu lịch sử kỳ '{rec.period}' cho học phần '{rec.course}'",
                                "dim")
                        except Exception as he:
                            _LOG.warning("Không lưu được lịch sử: %s", he)
                    except Exception as e:
                        _LOG.warning("Không tạo được báo cáo phân tích cho %s: %s", name, e)
                        self._log1.log(f"   ⚠ Không tạo được báo cáo phân tích: {e}", "warning")

                success += 1
            except Exception as e:
                fail += 1
                self._log1.log(f"   ❌ Lỗi: {e}", "error")

        self._set_progress(self._prog1, self._prog1_lbl, 100, "Hoàn tất!")
        _elapsed = (datetime.now() - _start_time).total_seconds()
        bench = benchmark.build_result(n_reports=total_out, auto_seconds=_elapsed)
        self._log1.log(f"", "dim")
        self._log1.log(f"{'═'*48}", "dim")
        self._log1.log(f"✅ Hoàn tất: {success} thành công, {fail} thất bại  •  {total_out} file tạo ra", "success")
        self._log1.log(f"⏱ Hiệu quả (ước lượng so với làm thủ công):", "info")
        for line in bench.summary_lines():
            self._log1.log(f"   • {line}", "dim")
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
