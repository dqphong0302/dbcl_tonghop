#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI Tổng hợp Khảo sát Excel → Word.

Giao diện Tkinter gồm 3 tab: Tổng hợp báo cáo, Gộp Feedbacks, Cấu hình.
Xử lý chạy trên luồng nền để không khóa giao diện.
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
import traceback

try:
    from excel_to_word import (
        extract_data_from_excel,
        create_word_report,
        merge_sheets_data,
        parse_open_responses,
        get_sheet_count,
    )
    from merge_feedbacks import merge_feedbacks
    from app_logging import get_logger
    import benchmark
except Exception as e:
    # Initialize a temporary tkinter root to show messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Lỗi Khởi động",
        f"Không thể khởi động ứng dụng do lỗi import/khởi tạo:\n\n{traceback.format_exc()}"
    )
    with open("crash_error.txt", "w", encoding="utf-8") as f:
        traceback.print_exc(file=f)
    sys.exit(1)

try:
    _LOG = get_logger()
except Exception as e:
    _LOG = None

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

def _pick_font(candidates, fallback):
    """Chọn font đầu tiên có sẵn trên hệ thống để giao diện đồng nhất đa nền tảng."""
    try:
        from tkinter import font as _tkfont
        available = {f.lower() for f in _tkfont.families()}
        for name in candidates:
            if name.lower() in available:
                return name
    except Exception:
        pass
    return fallback


# Font giao diện và font monospace cho nhật ký, chọn theo nền tảng.
if sys.platform == "darwin":
    FF = _pick_font(["SF Pro Text", "Helvetica Neue", "Segoe UI"], "Helvetica")
    FFM = _pick_font(["SF Mono", "Menlo", "Consolas"], "Courier")
elif sys.platform == "win32":
    FF = _pick_font(["Segoe UI", "Calibri"], "Arial")
    FFM = _pick_font(["Consolas", "Cascadia Mono"], "Courier New")
else:
    FF = _pick_font(["Ubuntu", "Noto Sans", "DejaVu Sans", "Segoe UI"], "TkDefaultFont")
    FFM = _pick_font(["Ubuntu Mono", "DejaVu Sans Mono", "Consolas"], "TkFixedFont")


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
        # Cho phép mở sẵn một tab khi khởi động (phục vụ chụp ảnh tài liệu).
        _start_tab = os.environ.get("DBCL_START_TAB")
        if _start_tab is None and "--tab" in sys.argv:
            try:
                _start_tab = sys.argv[sys.argv.index("--tab") + 1]
            except (ValueError, IndexError):
                _start_tab = None
        if _start_tab is not None:
            try:
                self._nb.select(int(_start_tab))
            except Exception:
                pass

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
        style.configure("TNotebook", background=C["bg"], borderwidth=0, tabmargins=[0, 0, 0, 0])
        style.configure("TNotebook.Tab", background=C["bg_input"], foreground=C["text_dim"],
                        padding=[20, 10], font=(FF, 10), borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", C["accent"]), ("active", C["bg_card"])],
                  foreground=[("selected", C["text_bright"]), ("active", C["text"])])

        # Buttons
        style.configure("TButton", background=C["bg_input"], foreground=C["text"],
                        padding=[12, 7], relief="flat", borderwidth=0, font=(FF, 10))
        style.map("TButton",
                  background=[("active", C["border"]), ("disabled", C["bg_input"])],
                  foreground=[("disabled", C["text_dim"])])
        style.configure("Accent.TButton", background=C["accent"], foreground=C["text_bright"],
                        padding=[18, 9], font=(FF, 11, "bold"))
        style.map("Accent.TButton",
                  background=[("active", "#8b83ff"), ("pressed", "#564fd1"),
                              ("disabled", C["bg_input"])],
                  foreground=[("disabled", C["text_dim"])])
        style.configure("Green.TButton", background=C["success"], foreground=C["text_bright"],
                        padding=[18, 9], font=(FF, 11, "bold"))
        style.map("Green.TButton",
                  background=[("active", "#34d399"), ("pressed", "#0d9268"),
                              ("disabled", C["bg_input"])],
                  foreground=[("disabled", C["text_dim"])])

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
        hdr = tk.Frame(self, bg=C["bg"], padx=28, pady=18)
        hdr.pack(fill="x")

        tk.Label(hdr, text="📊", font=(FF, 28), bg=C["bg"], fg=C["accent"]).pack(side="left")
        tf = tk.Frame(hdr, bg=C["bg"])
        tf.pack(side="left", padx=(14, 0))
        tk.Label(tf, text="Tổng hợp Khảo sát", font=(FF, 18, "bold"),
                 bg=C["bg"], fg=C["text_bright"]).pack(anchor="w")
        tk.Label(tf, text="Excel (.xls / .xlsx)  →  Word (.docx)   •   Hỗ trợ nhiều sheet   •   Xử lý cục bộ, offline",
                 font=(FF, 9), bg=C["bg"], fg=C["text_dim"]).pack(anchor="w", pady=(2, 0))

        tk.Label(hdr, text=self.APP_VERSION, font=(FF, 9, "bold"),
                 bg=C["bg_input"], fg=C["accent"], padx=10, pady=4).pack(side="right")

        sep = tk.Frame(self, bg=C["border"], height=1)
        sep.pack(fill="x")

        # ── NOTEBOOK ──
        nb_frame = tk.Frame(self, bg=C["bg"], padx=20, pady=12)
        nb_frame.pack(fill="both", expand=True)

        self._nb = ttk.Notebook(nb_frame)
        self._nb.pack(fill="both", expand=True)

        self._tab1 = self._build_tab1()
        self._tab2 = self._build_tab2()
        self._tab3 = self._build_tab3()
        self._nb.add(self._tab1, text="  📄 Tổng hợp báo cáo  ")
        self._nb.add(self._tab2, text="  💬 Gộp Feedbacks  ")
        self._nb.add(self._tab3, text="  ⚙️ Cấu hình  ")

        # ── STATUS BAR ──
        sep2 = tk.Frame(self, bg=C["border"], height=1)
        sep2.pack(fill="x")
        status_bar = tk.Frame(self, bg=C["bg_card"])
        status_bar.pack(fill="x")
        self._status_var = tk.StringVar(value="Sẵn sàng  •  Chọn tab và thêm file Excel để bắt đầu")
        tk.Label(status_bar, textvariable=self._status_var, font=(FF, 9),
                 bg=C["bg_card"], fg=C["text_dim"], anchor="w", padx=24, pady=7).pack(side="left")

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
        self._field_label(c1, "Tên trường:", row=0, column=0, sticky="w", pady=6)
        self._entry(c1, self._university_var).grid(row=0, column=1, sticky="ew", pady=6, ipady=4, padx=(0, 20))
        self._field_label(c1, "Tên khoa/Đơn vị:", row=0, column=2, sticky="w", pady=6)
        self._entry(c1, self._faculty_var).grid(row=0, column=3, sticky="ew", pady=6, ipady=4)

        # Hàng 1: Quốc hiệu & Tiêu ngữ
        self._field_label(c1, "Quốc hiệu:", row=1, column=0, sticky="w", pady=6)
        self._entry(c1, self._country_header_var).grid(row=1, column=1, sticky="ew", pady=6, ipady=4, padx=(0, 20))
        self._field_label(c1, "Tiêu ngữ:", row=1, column=2, sticky="w", pady=6)
        self._entry(c1, self._country_slogan_var).grid(row=1, column=3, sticky="ew", pady=6, ipady=4)

        # Hàng 2: Tiêu đề báo cáo (nhiều dòng)
        self._field_label(c1, "Tiêu đề báo cáo:", row=2, column=0, sticky="nw", pady=10)
        self._report_title_txt = self._text(c1, height=3)
        self._report_title_txt.grid(row=2, column=1, columnspan=3, sticky="ew", pady=6)

        # Hàng 3: Font chữ & Cỡ chữ
        self._field_label(c1, "Font chữ:", row=3, column=0, sticky="w", pady=6)
        font_box = ttk.Combobox(c1, textvariable=self._font_name_var,
                                values=["Times New Roman", "Arial", "Calibri", "Segoe UI"],
                                state="readonly", font=(FF, 10))
        font_box.grid(row=3, column=1, sticky="ew", pady=6, padx=(0, 20))
        self._field_label(c1, "Cỡ chữ (pt):", row=3, column=2, sticky="w", pady=6)
        self._spin(c1, self._font_size_var, 8, 24).grid(row=3, column=3, sticky="ew", pady=6, ipady=2)

        # Hàng 4: Căn lề (Trái, Phải, Trên, Dưới)
        self._field_label(c1, "Căn lề (cm):", row=4, column=0, sticky="w", pady=6)
        margin_frame = tk.Frame(c1, bg=C["bg_card"])
        margin_frame.grid(row=4, column=1, columnspan=3, sticky="ew", pady=6)
        margin_frame.columnconfigure((1, 3, 5, 7), weight=1)
        for col, (label, var) in enumerate([
            ("Trái:", self._margin_left_var), ("Phải:", self._margin_right_var),
            ("Trên:", self._margin_top_var), ("Dưới:", self._margin_bottom_var),
        ]):
            tk.Label(margin_frame, text=label, font=(FF, 9), bg=C["bg_card"],
                     fg=C["text_dim"]).grid(row=0, column=col * 2, padx=(0, 4))
            pad = (0, 15) if col < 3 else (0, 0)
            self._entry(margin_frame, var, width=6).grid(
                row=0, column=col * 2 + 1, sticky="ew", padx=pad, ipady=3)

        # Hàng 5: Màu nền bảng & Số dòng trống ý kiến
        self._field_label(c1, "Màu nền bảng (HEX):", row=5, column=0, sticky="w", pady=6)
        self._entry(c1, self._table_header_bg_var).grid(row=5, column=1, sticky="ew", pady=6, ipady=4, padx=(0, 20))
        self._field_label(c1, "Dòng trống ý kiến:", row=5, column=2, sticky="w", pady=6)
        self._spin(c1, self._open_response_lines_var, 0, 20).grid(row=5, column=3, sticky="ew", pady=6, ipady=2)
        
        # ── CARD 2: CẤU HÌNH TỪ KHÓA FEEDBACK ──
        c2 = create_card("💬 Cấu hình Từ khóa Feedback")
        c2.columnconfigure(1, weight=1)
        
        # Hàng 0: Cột feedback tìm kiếm
        self._field_label(c2, "Tìm cột feedback:", row=0, column=0, sticky="w", pady=6)
        self._entry(c2, self._feedback_search_keywords_var).grid(row=0, column=1, sticky="ew", pady=6, ipady=4)
        self._hint(c2, "(Cách nhau bằng dấu phẩy. Ví dụ: ý kiến đóng góp, feedback, nhận xét)", row=1, column=1)

        # Hàng 2: Từ khóa bỏ qua
        self._field_label(c2, "Từ khóa bỏ qua:", row=2, column=0, sticky="w", pady=6)
        self._entry(c2, self._feedback_skip_keywords_var).grid(row=2, column=1, sticky="ew", pady=6, ipady=4)
        self._hint(c2, "(Cách nhau bằng dấu phẩy. Loại bỏ các phản hồi khớp hoàn toàn với cụm từ này)", row=3, column=1)

        # Hàng 4: Regex bỏ qua (nhiều dòng)
        self._field_label(c2, "Regex bỏ qua:", row=4, column=0, sticky="nw", pady=10)
        self._feedback_skip_patterns_txt = self._text(c2, height=4)
        self._feedback_skip_patterns_txt.grid(row=4, column=1, sticky="ew", pady=6)
        self._hint(c2, "(Mỗi dòng một biểu thức chính quy. Dùng để lọc bỏ các phản hồi trùng khớp định dạng)", row=5, column=1)

        # ── HÀNH ĐỘNG ──
        act_frame = tk.Frame(scrollable_frame, bg=C["bg"], padx=20, pady=12)
        act_frame.pack(fill="x")

        btn_save = ttk.Button(act_frame, text="💾 Lưu cấu hình", style="Accent.TButton", command=self._save_config_from_ui)
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

    # ── Form field factories (dùng chung cho tab Cấu hình) ───────

    @staticmethod
    def _field_label(parent, text, **grid):
        lbl = tk.Label(parent, text=text, font=(FF, 9, "bold"),
                       bg=C["bg_card"], fg=C["text"])
        lbl.grid(padx=(0, 10), **grid)
        return lbl

    @staticmethod
    def _hint(parent, text, **grid):
        lbl = tk.Label(parent, text=text, font=(FF, 8),
                       bg=C["bg_card"], fg=C["text_dim"])
        lbl.grid(sticky="w", pady=(0, 6), **grid)
        return lbl

    @staticmethod
    def _entry(parent, textvariable, **kwargs):
        return tk.Entry(parent, textvariable=textvariable, font=(FF, 10),
                        bg=C["bg_input"], fg=C["text"], insertbackground=C["text"],
                        relief="flat", bd=0, highlightthickness=1,
                        highlightbackground=C["border"], highlightcolor=C["accent"],
                        **kwargs)

    @staticmethod
    def _text(parent, height, **kwargs):
        return tk.Text(parent, font=(FF, 10), height=height,
                       bg=C["bg_input"], fg=C["text"], insertbackground=C["text"],
                       relief="flat", highlightthickness=1,
                       highlightbackground=C["border"], highlightcolor=C["accent"],
                       **kwargs)

    @staticmethod
    def _spin(parent, textvariable, frm, to, **kwargs):
        return tk.Spinbox(parent, textvariable=textvariable, from_=frm, to=to,
                          increment=1, font=(FF, 10), bg=C["bg_input"], fg=C["text"],
                          relief="flat", bd=0, highlightthickness=1,
                          highlightbackground=C["border"], highlightcolor=C["accent"],
                          **kwargs)

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

                for si in range(n_sheets):
                    try:
                        data = extract_data_from_excel(filepath, sheet_index=si, apply_filename_suffix=False)
                        if data["survey_questions"]:
                            all_data.append(data)
                        else:
                            self._log1.log(f'   ⏭ Sheet {si+1} "{sheet_names[si]}": bỏ qua (không có dữ liệu)', "dim")
                    except Exception as e:
                        self._log1.log(f'   ⚠ Sheet {si+1} "{sheet_names[si]}": lỗi đọc – {e}', "warning")
                        continue

                if not all_data:
                    self._log1.log("   ⚠️ Không có sheet nào chứa dữ liệu khảo sát", "warning")
                    fail += 1
                    continue

                # Lấy số SV KHÔNG tham gia khảo sát từ tên file
                match = re.search(r'\s+(\d+)$', basename)
                sv_khong_khao_sat = int(match.group(1)) if match else None

                is_valid = True
                if len(all_data) == 1:
                    data = all_data[0]
                    total_phieu = data["survey_questions"][0].get("total", 0) if data["survey_questions"] else 0
                    if sv_khong_khao_sat is not None:
                        data["meta"]["tong_so_sv"] = total_phieu + sv_khong_khao_sat
                    
                    if data["meta"].get("tong_so_sv") is None:
                        is_valid = False
                    else:
                        out_path = os.path.join(out_dir, f"{basename}.docx")
                        create_word_report(data, out_path)
                        total_out += 1
                        n_q = len(data["survey_questions"])
                        n_resp = len(parse_open_responses(data["open_responses"]))
                        self._log1.log(f"   ✅ {n_q} câu, {n_resp} ý kiến, {data['meta']['tong_so_sv']} SV → {os.path.basename(out_path)}", "success")
                else:
                    merged = merge_sheets_data(all_data)
                    if merged:
                        total_tong_sv = 0
                        has_any_tong_sv = False
                        for d in all_data:
                            sv = d["meta"].get("tong_so_sv", None) # Đây là các F5 do apply_filename_suffix=False
                            if sv is not None:
                                total_tong_sv += sv
                                has_any_tong_sv = True

                        total_phieu = merged["survey_questions"][0].get("total", 0) if merged["survey_questions"] else 0
                        if sv_khong_khao_sat is not None:
                            merged["meta"]["tong_so_sv"] = total_phieu + sv_khong_khao_sat
                        elif has_any_tong_sv:
                            merged["meta"]["tong_so_sv"] = total_tong_sv
                        else:
                            merged["meta"]["tong_so_sv"] = None
                            merged["warnings"].append("⚠️ Không lấy được số SV không tham gia từ tên file và ô F5 cũng trống ở tất cả các sheet!")

                        if merged["meta"].get("tong_so_sv") is None:
                            is_valid = False
                        else:
                            out_path = os.path.join(out_dir, f"{basename}.docx")
                            create_word_report(merged, out_path)
                            total_out += 1
                            total_sv = merged["meta"].get("tong_so_sv", 0)
                            self._log1.log(f"   📊 ĐÃ GỘP: {len(all_data)} sheets, {total_sv} SV → {os.path.basename(out_path)}", "info")

                if not is_valid:
                    self._log1.log(f"   ❌ Không xác định được Tổng số SV cho file {name}", "error")
                    msg = (f"File: {name}\n\nKhông xác định được Tổng số SV cho file này.\n\n"
                           f"Vui lòng ghi số SV không tham gia ở cuối tên file\n"
                           f"hoặc điền vào ô F5 trong Excel rồi thử lại.")
                    self.after(0, lambda m=msg: messagebox.showwarning("Thiếu thông tin Tổng SV", m))
                    fail += 1
                    continue

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
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        import traceback
        # Attempt to show a graphical error dialog
        try:
            messagebox.showerror(
                "Lỗi Hệ Thống",
                f"Ứng dụng gặp lỗi nghiêm trọng khi đang hoạt động:\n\n{traceback.format_exc()}"
            )
        except Exception:
            pass
        # Write error log to file
        with open("crash_error.txt", "w", encoding="utf-8") as f:
            traceback.print_exc(file=f)

