#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI Tổng hợp Khảo sát Excel → Word.

Giao diện CustomTkinter (chế độ sáng) gồm thanh điều hướng bên trái
và 3 trang: Tổng hợp báo cáo, Gộp Feedbacks, Cấu hình.
Xử lý chạy trên luồng nền; luồng nền chỉ gửi cập nhật qua hàng đợi để
luồng giao diện tự vẽ lại (Tkinter không an toàn đa luồng).
"""

import glob
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk


def _fatal(title, detail):
    """Hiện lỗi khởi động bằng messagebox và ghi crash_error.txt cạnh thư mục home."""
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, detail)
        root.destroy()
    except Exception:
        pass
    try:
        (Path.home() / "dbcl_tonghop_crash.txt").write_text(detail, encoding="utf-8")
    except Exception:
        pass


try:
    import customtkinter as ctk
    import excel_to_word as e2w
    from merge_feedbacks import merge_feedbacks
    from app_logging import get_logger, get_log_dir
    import benchmark
except Exception:
    _fatal("Lỗi khởi động",
           f"Không thể khởi động ứng dụng do lỗi import:\n\n{traceback.format_exc()}")
    sys.exit(1)

try:
    _LOG = get_logger()
except Exception:
    _LOG = None

APP_NAME = "Tổng hợp Khảo sát"
APP_VERSION = "3.0.0"

# ═══════════════════════════════════════════════════════════════
# SETTINGS PERSISTENCE (dùng chung file với excel_to_word.CONFIG)
# ═══════════════════════════════════════════════════════════════

SETTINGS_FILE = e2w.SETTINGS_FILE


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


def resource_path(rel):
    """Đường dẫn tài nguyên, chạy được cả khi đóng gói bằng PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def open_path(path):
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ═══════════════════════════════════════════════════════════════
# THEME — giao diện sáng
# ═══════════════════════════════════════════════════════════════

C = {
    "bg":          "#f4f5fb",
    "sidebar":     "#ffffff",
    "card":        "#ffffff",
    "card_alt":    "#f7f8fc",
    "border":      "#e4e6f0",
    "text":        "#1e2033",
    "muted":       "#6b7085",
    "accent":      "#4f46e5",
    "accent_hover": "#4338ca",
    "accent_soft": "#eef0ff",
    "green":       "#059669",
    "green_hover": "#047857",
    "danger":      "#dc2626",
    "danger_soft": "#fdecec",
    "neutral_btn": "#eef0f6",
    "neutral_hover": "#e2e5ef",
    "log_bg":      "#fafbff",
}

# Màu chữ nhật ký theo mức độ.
LOG_TAGS = {
    "success": "#16a34a",
    "error":   "#e5484d",
    "warning": "#d97706",
    "info":    "#6366f1",
    "dim":     "#8a8fa3",
}


def pick_font(root, candidates, fallback):
    try:
        from tkinter import font as tkfont
        available = {f.lower() for f in tkfont.families(root)}
        for name in candidates:
            if name.lower() in available:
                return name
    except Exception:
        pass
    return fallback


class Fonts:
    """Font chọn theo nền tảng (khởi tạo sau khi có cửa sổ gốc)."""
    ui = "TkDefaultFont"
    mono = "TkFixedFont"

    @classmethod
    def init(cls, root):
        if sys.platform == "darwin":
            cls.ui = pick_font(root, ["SF Pro Text", "Helvetica Neue"], "Helvetica")
            cls.mono = pick_font(root, ["SF Mono", "Menlo"], "Courier")
        elif sys.platform == "win32":
            cls.ui = pick_font(root, ["Segoe UI Variable Text", "Segoe UI"], "Arial")
            cls.mono = pick_font(root, ["Cascadia Mono", "Consolas"], "Courier New")
        else:
            cls.ui = pick_font(root, ["Inter", "Ubuntu", "Noto Sans", "DejaVu Sans"], "TkDefaultFont")
            cls.mono = pick_font(root, ["JetBrains Mono", "Ubuntu Mono", "DejaVu Sans Mono"], "TkFixedFont")


def F(size=13, weight="normal"):
    return ctk.CTkFont(family=Fonts.ui, size=size, weight=weight)


def card(parent, **kwargs):
    return ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=14,
                        border_width=1, border_color=C["border"], **kwargs)


def section_title(parent, text, subtitle=None):
    box = ctk.CTkFrame(parent, fg_color="transparent")
    ctk.CTkLabel(box, text=text, font=F(14, "bold"), text_color=C["text"],
                 anchor="w").pack(anchor="w")
    if subtitle:
        ctk.CTkLabel(box, text=subtitle, font=F(12), text_color=C["muted"],
                     anchor="w", justify="left").pack(anchor="w")
    return box


def button(parent, text, command, kind="neutral", **kwargs):
    styles = {
        "primary": dict(fg_color=C["accent"], hover_color=C["accent_hover"], text_color="#ffffff"),
        "green":   dict(fg_color=C["green"], hover_color=C["green_hover"], text_color="#ffffff"),
        "neutral": dict(fg_color=C["neutral_btn"], hover_color=C["neutral_hover"], text_color=C["text"]),
        "ghost":   dict(fg_color="transparent", hover_color=C["neutral_btn"], text_color=C["muted"]),
    }
    big = kind in ("primary", "green")
    opts = dict(height=40 if big else 34, corner_radius=10,
                font=F(14 if big else 13, "bold" if big else "normal"))
    opts.update(styles[kind])
    opts.update(kwargs)
    return ctk.CTkButton(parent, text=text, command=command, **opts)


def human_size(n):
    if n >= 1024 * 1024:
        return f"{n / 1024 / 1024:.1f} MB"
    if n >= 1024:
        return f"{n // 1024} KB"
    return f"{n} B"


def shorten(text, limit=44):
    if len(text) <= limit:
        return text
    half = (limit - 1) // 2
    return text[:half] + "…" + text[-half:]


# ═══════════════════════════════════════════════════════════════
# WIDGETS
# ═══════════════════════════════════════════════════════════════

class FileList(ctk.CTkFrame):
    """Danh sách file Excel: thêm file/thư mục, xóa từng dòng, đếm file."""

    EXTENSIONS = ("*.xls", "*.xlsx")

    def __init__(self, parent, on_change=None):
        super().__init__(parent, fg_color="transparent")
        self._files = []
        self._on_change = on_change

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 10))
        button(bar, "＋  Thêm file", self._add_files, width=120).pack(side="left")
        button(bar, "📁  Thêm thư mục", self._add_folder, width=140).pack(side="left", padx=(8, 0))
        self._clear_btn = button(bar, "Xóa tất cả", self.clear, kind="ghost", width=90)
        self._clear_btn.pack(side="right")

        self._list = ctk.CTkScrollableFrame(self, fg_color=C["card_alt"], corner_radius=10,
                                            border_width=1, border_color=C["border"])
        self._list.pack(fill="both", expand=True)
        self._list.grid_columnconfigure(0, weight=1)

        self._count = ctk.CTkLabel(self, text="", font=F(12), text_color=C["muted"], anchor="w")
        self._count.pack(fill="x", pady=(8, 0))
        self._render(notify=False)

    @property
    def files(self):
        return list(self._files)

    def _add(self, paths):
        new = [p for p in paths if p not in self._files]
        if new:
            self._files.extend(new)
            self._render()

    def _add_files(self):
        self._add(filedialog.askopenfilenames(
            title="Chọn file Excel",
            filetypes=[("Excel", "*.xls *.xlsx"), ("Tất cả", "*.*")]))

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file Excel")
        if folder:
            found = []
            for ext in self.EXTENSIONS:
                found += sorted(glob.glob(os.path.join(folder, ext)))
            self._add([f for f in found if not os.path.basename(f).startswith("~$")])

    def _remove(self, path):
        self._files.remove(path)
        self._render()

    def clear(self):
        self._files.clear()
        self._render()

    def _render(self, notify=True):
        for w in self._list.winfo_children():
            w.destroy()

        if not self._files:
            empty = ctk.CTkFrame(self._list, fg_color="transparent")
            empty.grid(row=0, column=0, sticky="nsew", pady=36)
            ctk.CTkLabel(empty, text="📄", font=F(30)).pack()
            ctk.CTkLabel(empty, text="Chưa có file nào", font=F(14, "bold"),
                         text_color=C["text"]).pack(pady=(6, 0))
            ctk.CTkLabel(empty, text="Bấm “Thêm file” hoặc “Thêm thư mục” để chọn file .xls / .xlsx",
                         font=F(12), text_color=C["muted"]).pack()
        for i, path in enumerate(self._files):
            self._row(i, path)

        n = len(self._files)
        self._count.configure(text=f"{n} file đã chọn" if n else "")
        self._clear_btn.configure(state="normal" if n else "disabled")
        if notify and self._on_change:
            self._on_change(n)

    def _row(self, i, path):
        row = ctk.CTkFrame(self._list, fg_color=C["card"], corner_radius=8, height=40)
        row.grid(row=i, column=0, sticky="ew", padx=2, pady=2)
        row.grid_columnconfigure(1, weight=1)

        ext = os.path.splitext(path)[1].lstrip(".").upper() or "?"
        ctk.CTkLabel(row, text=ext, width=46, height=24, corner_radius=6,
                     fg_color=C["accent_soft"], text_color=C["accent"],
                     font=F(11, "bold")).grid(row=0, column=0, padx=(10, 10), pady=8)
        ctk.CTkLabel(row, text=shorten(os.path.basename(path)), font=F(13),
                     text_color=C["text"], anchor="w").grid(row=0, column=1, sticky="ew")
        size = os.path.getsize(path) if os.path.exists(path) else 0
        ctk.CTkLabel(row, text=human_size(size), font=F(12), text_color=C["muted"],
                     width=70, anchor="e").grid(row=0, column=2, padx=(8, 4))
        ctk.CTkButton(row, text="✕", width=28, height=28, corner_radius=6,
                      fg_color="transparent", hover_color=C["danger_soft"],
                      text_color=C["muted"], font=F(13),
                      command=lambda p=path: self._remove(p)).grid(row=0, column=3, padx=(0, 6))


class LogView(ctk.CTkFrame):
    """Nhật ký xử lý có màu theo mức độ, nút sao chép và xóa."""

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 8))
        section_title(bar, "Nhật ký xử lý").pack(side="left")
        button(bar, "Xóa", self.clear, kind="ghost", width=56, height=28).pack(side="right")
        button(bar, "Sao chép", self._copy, kind="ghost", width=80, height=28).pack(side="right")

        self._txt = ctk.CTkTextbox(self, fg_color=C["log_bg"], text_color=C["text"],
                                   border_width=1, border_color=C["border"], corner_radius=10,
                                   font=ctk.CTkFont(family=Fonts.mono, size=12),
                                   wrap="word", state="disabled")
        self._txt.pack(fill="both", expand=True)
        for tag, color in LOG_TAGS.items():
            self._txt.tag_config(tag, foreground=color)

    def log(self, msg, tag=None):
        """Chỉ gọi từ luồng giao diện."""
        if _LOG and msg.strip():
            level = {"error": _LOG.error, "warning": _LOG.warning,
                     "success": _LOG.info, "info": _LOG.info}.get(tag, _LOG.debug)
            level(msg.strip())
        self._txt.configure(state="normal")
        self._txt.insert("end", datetime.now().strftime("%H:%M:%S  "), "dim")
        self._txt.insert("end", msg + "\n", tag or ())
        self._txt.see("end")
        self._txt.configure(state="disabled")

    def clear(self):
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.configure(state="disabled")

    def _copy(self):
        self.clipboard_clear()
        self.clipboard_append(self._txt.get("1.0", "end").strip())


class StatTile(ctk.CTkFrame):
    def __init__(self, parent, label, color):
        super().__init__(parent, fg_color=C["card_alt"], corner_radius=10)
        self._value = ctk.CTkLabel(self, text="–", font=F(22, "bold"), text_color=color)
        self._value.pack(anchor="w", padx=14, pady=(10, 0))
        ctk.CTkLabel(self, text=label, font=F(12), text_color=C["muted"]).pack(
            anchor="w", padx=14, pady=(0, 10))

    def set(self, value):
        self._value.configure(text=str(value))


# ═══════════════════════════════════════════════════════════════
# PAGES
# ═══════════════════════════════════════════════════════════════

class ToolPage(ctk.CTkFrame):
    """Khung chung cho trang xử lý: chọn file, thư mục kết quả, chạy, nhật ký."""

    TITLE = ""
    SUBTITLE = ""
    RUN_TEXT = "Bắt đầu"
    RUN_KIND = "primary"
    OUTDIR_KEY = ""
    STAT_LABELS = ("Tổng file", "Thành công", "Lỗi")

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self._busy = False
        self.outdir = tk.StringVar(value=app.settings.get(self.OUTDIR_KEY, ""))
        self.outdir.trace_add("write", lambda *_: app.update_setting(self.OUTDIR_KEY, self.outdir.get()))

        self.grid_columnconfigure(0, weight=5, uniform="col")
        self.grid_columnconfigure(1, weight=4, uniform="col")
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 18))
        ctk.CTkLabel(header, text=self.TITLE, font=F(24, "bold"), text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(header, text=self.SUBTITLE, font=F(13), text_color=C["muted"],
                     justify="left", anchor="w").pack(anchor="w", pady=(2, 0))

        # ── Cột trái: file đầu vào + thư mục kết quả + nút chạy
        left = ctk.CTkFrame(self, fg_color="transparent")
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 16))
        left.grid_rowconfigure(0, weight=1)
        left.grid_columnconfigure(0, weight=1)

        files_card = card(left)
        files_card.grid(row=0, column=0, sticky="nsew")
        section_title(files_card, "File Excel đầu vào").pack(anchor="w", padx=18, pady=(16, 10))
        self.file_list = FileList(files_card, on_change=self._on_files_changed)
        self.file_list.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        run_card = card(left)
        run_card.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        run_card.grid_columnconfigure(0, weight=1)
        section_title(run_card, "Thư mục lưu kết quả",
                      "Để trống để lưu vào thư mục “output” cạnh file Excel đầu tiên").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(16, 8))
        ctk.CTkEntry(run_card, textvariable=self.outdir, height=36, corner_radius=8,
                     border_color=C["border"], fg_color=C["card_alt"], text_color=C["text"],
                     font=F(13)).grid(row=1, column=0, sticky="ew", padx=(18, 8))
        button(run_card, "Chọn…", self._choose_outdir, width=80).grid(row=1, column=1, padx=(0, 18))

        actions = ctk.CTkFrame(run_card, fg_color="transparent")
        actions.grid(row=2, column=0, columnspan=2, sticky="ew", padx=18, pady=(14, 4))
        self.run_btn = button(actions, self.RUN_TEXT, self._start, kind=self.RUN_KIND, width=200)
        self.run_btn.pack(side="left")
        button(actions, "📂  Mở thư mục kết quả", self._open_outdir).pack(side="left", padx=(10, 0))

        self.progress = ctk.CTkProgressBar(run_card, height=8, corner_radius=4,
                                           progress_color=C["accent"] if self.RUN_KIND == "primary" else C["green"],
                                           fg_color=C["neutral_btn"])
        self.progress.set(0)
        self.progress.grid(row=3, column=0, columnspan=2, sticky="ew", padx=18, pady=(12, 4))
        self.progress_lbl = ctk.CTkLabel(run_card, text="Sẵn sàng", font=F(12),
                                         text_color=C["muted"], anchor="w")
        self.progress_lbl.grid(row=4, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 14))

        # ── Cột phải: thống kê + nhật ký
        right = card(self)
        right.grid(row=1, column=1, sticky="nsew")
        stats = ctk.CTkFrame(right, fg_color="transparent")
        stats.pack(fill="x", padx=18, pady=(16, 0))
        colors = (C["text"], C["green"], C["danger"])
        self.stats = []
        for i, (label, color) in enumerate(zip(self.STAT_LABELS, colors)):
            stats.grid_columnconfigure(i, weight=1, uniform="stat")
            tile = StatTile(stats, label, color)
            tile.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
            self.stats.append(tile)
        self.log = LogView(right)
        self.log.pack(fill="both", expand=True, padx=18, pady=16)

    # ── Hành động ───────────────────────────────────────────────

    def _on_files_changed(self, n):
        if not self._busy:
            self.stats[0].set(n or "–")
            self.app.set_status(f"{n} file đã chọn" if n else "Sẵn sàng")

    def _choose_outdir(self):
        folder = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if folder:
            self.outdir.set(folder)

    def _open_outdir(self):
        path = self.outdir.get().strip()
        if path and os.path.isdir(path):
            open_path(path)
        else:
            messagebox.showinfo("Thông báo", "Thư mục kết quả chưa tồn tại — hãy chạy xử lý trước.")

    def _start(self):
        files = self.file_list.files
        if not files:
            messagebox.showwarning("Chưa có file", "Vui lòng thêm file Excel trước khi xử lý.")
            return
        out_dir = self.outdir.get().strip()
        if not out_dir:
            out_dir = os.path.join(os.path.dirname(files[0]), "output")
            self.outdir.set(out_dir)
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Lỗi thư mục", f"Không tạo được thư mục kết quả:\n{e}")
            return

        self._busy = True
        self.run_btn.configure(state="disabled", text="Đang xử lý…")
        self.log.clear()
        self.progress.set(0)
        self.stats[0].set(len(files))
        self.stats[1].set(0)
        self.stats[2].set(0)
        threading.Thread(target=self._run_safe, args=(files, out_dir), daemon=True).start()

    def _run_safe(self, files, out_dir):
        try:
            self.run(files, out_dir)
        except Exception as e:
            self.post_log(f"❌ Lỗi không mong muốn: {e}", "error")
            if _LOG:
                _LOG.exception("Lỗi luồng xử lý")
        finally:
            self.app.post(self._finish)

    def _finish(self):
        self._busy = False
        self.run_btn.configure(state="normal", text=self.RUN_TEXT)

    # ── Cập nhật giao diện từ luồng nền ─────────────────────────

    def post_log(self, msg, tag=None):
        self.app.post(lambda: self.log.log(msg, tag))

    def post_progress(self, fraction, text):
        def _do():
            self.progress.set(fraction)
            self.progress_lbl.configure(text=text)
        self.app.post(_do)

    def post_stat(self, index, value):
        self.app.post(lambda: self.stats[index].set(value))

    def run(self, files, out_dir):
        raise NotImplementedError


class ReportPage(ToolPage):
    TITLE = "Tổng hợp báo cáo"
    SUBTITLE = ("Chuyển file kết quả khảo sát Excel thành báo cáo Word theo mẫu. "
                "File nhiều sheet được tự động gộp thành một báo cáo.")
    RUN_TEXT = "▶   Tạo báo cáo"
    OUTDIR_KEY = "outdir1"

    def run(self, files, out_dir):
        total, ok, failed = len(files), 0, 0
        missing_total = []
        started = time.perf_counter()
        self.post_log(f"Bắt đầu xử lý {total} file → {out_dir}", "info")

        for i, path in enumerate(files):
            name = os.path.basename(path)
            self.post_progress(i / total, f"Đang xử lý {i + 1}/{total}: {shorten(name, 50)}")
            self.app.post_status(f"Đang xử lý file {i + 1}/{total}…")
            self.post_log("")
            self.post_log(f"📄 {name}", "info")

            result = e2w.generate_report(path, out_dir,
                                         on_progress=lambda m: self.post_log(f"   ⏭ {m}", "dim"))
            if result.ok:
                ok += 1
                merged = f"gộp {result.n_sheets} sheet, " if result.merged else ""
                self.post_log(f"   ✅ {merged}{result.n_questions} câu, {result.n_responses} ý kiến, "
                              f"{result.total_students} SV → {os.path.basename(result.output_path)}",
                              "success")
                for w in result.warnings:
                    self.post_log(f"   ⚠️ {w}", "warning")
            else:
                failed += 1
                self.post_log(f"   ❌ {result.error}", "error")
                if result.missing_total:
                    missing_total.append(name)
            self.post_stat(1, ok)
            self.post_stat(2, failed)

        elapsed = time.perf_counter() - started
        self.post_progress(1, f"Hoàn tất • {ok}/{total} file trong {elapsed:.1f} giây")
        self.post_log("")
        self.post_log(f"Hoàn tất: {ok} thành công, {failed} lỗi", "success" if not failed else "warning")
        if ok:
            self.post_log("Hiệu quả ước lượng so với làm thủ công:", "info")
            for line in benchmark.build_result(n_reports=ok, auto_seconds=elapsed).summary_lines():
                self.post_log(f"   • {line}", "dim")
        self.app.post_status(f"Hoàn tất • {ok} báo cáo đã tạo • {out_dir}")

        if missing_total:
            names = "\n".join(f"• {n}" for n in missing_total)
            self.app.post(lambda: messagebox.showwarning(
                "Thiếu thông tin Tổng số SV",
                f"Không xác định được Tổng số SV cho:\n\n{names}\n\n"
                "Hãy ghi số SV không tham gia ở cuối tên file (vd: “… HKI 3.xls”) "
                "hoặc điền vào ô F5 trong Excel rồi thử lại."))


class FeedbackPage(ToolPage):
    TITLE = "Gộp Feedbacks"
    SUBTITLE = ("Gom ý kiến đóng góp mở từ nhiều file khảo sát, làm sạch, loại trùng, "
                "phân loại theo chủ đề và xuất ra Feedbacks.docx.")
    RUN_TEXT = "💬   Gộp Feedbacks"
    RUN_KIND = "green"
    OUTDIR_KEY = "outdir2"

    def run(self, files, out_dir):
        self.post_log(f"Bắt đầu gộp feedbacks từ {len(files)} file", "info")
        for f in files:
            self.post_log(f"📄 {os.path.basename(f)}", "dim")
        self.post_progress(0.3, "Đang đọc dữ liệu…")
        self.app.post_status("Đang tổng hợp feedbacks…")

        output_path = merge_feedbacks(files, out_dir)
        if output_path:
            self.post_stat(1, len(files))
            self.post_progress(1, "Hoàn tất")
            self.post_log("")
            self.post_log(f"✅ Đã tạo {os.path.basename(output_path)}", "success")
            self.app.post_status(f"Hoàn tất • {output_path}")
        else:
            self.post_stat(2, len(files))
            self.post_progress(0, "Không có dữ liệu")
            self.post_log("⚠️ Không tìm thấy ý kiến hợp lệ nào trong các file đã chọn", "warning")
            self.app.post_status("Không có dữ liệu feedback")


class SettingsPage(ctk.CTkFrame):
    """Cấu hình định dạng Word và bộ lọc feedback (lưu vào SETTINGS_FILE)."""

    WORD_FIELDS = [
        ("university_name", "Tên trường"),
        ("faculty_name", "Tên khoa / đơn vị"),
        ("country_header", "Quốc hiệu"),
        ("country_slogan", "Tiêu ngữ"),
        ("report_title", "Tiêu đề báo cáo (mỗi dòng một hàng; {ten_mon} = tên học phần)"),
        ("signer_title", "Chức danh ký tên cuối báo cáo"),
        ("section_layouts", "Nhóm tiêu chí trong bảng kết quả"),
    ]
    NUMBER_FIELDS = [
        ("font_size", "Cỡ chữ (pt)", int),
        ("margin_left", "Lề trái (cm)", float),
        ("margin_right", "Lề phải (cm)", float),
        ("margin_top", "Lề trên (cm)", float),
        ("margin_bottom", "Lề dưới (cm)", float),
        ("open_response_lines", "Dòng trống ý kiến", int),
    ]
    FONTS = ["Times New Roman", "Arial", "Calibri", "Segoe UI"]

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self._widgets = {}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        ctk.CTkLabel(header, text="Cấu hình", font=F(24, "bold"), text_color=C["text"]).pack(side="left")
        button(header, "💾  Lưu cấu hình", self._save, kind="primary", width=150).pack(side="right")
        button(header, "Khôi phục mặc định", self._reset).pack(side="right", padx=(0, 10))

        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)

        # ── Định dạng Word
        c1 = card(body)
        c1.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        section_title(c1, "Định dạng file Word",
                      "Thông tin đầu trang và định dạng của báo cáo xuất ra").pack(
            anchor="w", padx=20, pady=(18, 12))
        form = ctk.CTkFrame(c1, fg_color="transparent")
        form.pack(fill="x", padx=20, pady=(0, 20))
        form.grid_columnconfigure((0, 1), weight=1, uniform="f")

        # (khóa, hàng, cột, số cột chiếm, số dòng textbox — 0 là ô nhập 1 dòng)
        layout = [
            ("university_name", 0, 0, 1, 2), ("faculty_name", 0, 1, 1, 0),
            ("country_header", 1, 0, 1, 0), ("country_slogan", 1, 1, 1, 0),
            ("report_title", 2, 0, 2, 3),
        ]
        labels = {key: label for key, label in self.WORD_FIELDS}
        for key, row, col, span, lines in layout:
            cell = self._field(form, labels[key])
            cell.grid(row=row, column=col, columnspan=span, sticky="new",
                      padx=(0, 12 if col == 0 and span == 1 else 0), pady=6)
            self._widgets[key] = self._textbox(cell, lines) if lines else self._entry(cell)
        r = 3
        font_cell = self._field(form, "Font chữ")
        font_cell.grid(row=r, column=0, sticky="ew", padx=(0, 12), pady=6)
        self._font_var = tk.StringVar()
        ctk.CTkOptionMenu(font_cell, variable=self._font_var, values=self.FONTS, height=36,
                          corner_radius=8, fg_color=C["card_alt"], button_color=C["accent"],
                          button_hover_color=C["accent_hover"], text_color=C["text"],
                          font=F(13), dropdown_font=F(13)).pack(fill="x")
        color_cell = self._field(form, "Màu nền tiêu đề bảng (HEX)")
        color_cell.grid(row=r, column=1, sticky="ew", pady=6)
        self._widgets["table_header_bg"] = self._entry(color_cell)

        nums = ctk.CTkFrame(form, fg_color="transparent")
        nums.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        for i, (key, label, _) in enumerate(self.NUMBER_FIELDS):
            nums.grid_columnconfigure(i, weight=1, uniform="n")
            cell = self._field(nums, label)
            cell.grid(row=0, column=i, sticky="ew", padx=(0, 0 if i == len(self.NUMBER_FIELDS) - 1 else 10))
            self._widgets[key] = self._entry(cell)

        # ── Bảng kết quả & ký tên
        c3 = card(body)
        c3.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        section_title(c3, "Bảng kết quả & chữ ký",
                      "Nhóm tiêu chí: mỗi dòng “Tên nhóm | từ-đến” theo số câu (8.1, 8.2 tính là câu 8). "
                      "Mỗi mẫu phiếu là một khối, cách nhau một dòng trống;\nphần mềm tự chọn khối có "
                      "số câu cuối trùng với file Excel.").pack(anchor="w", padx=20, pady=(18, 12))
        form3 = ctk.CTkFrame(c3, fg_color="transparent")
        form3.pack(fill="x", padx=20, pady=(0, 20))
        form3.grid_columnconfigure(0, weight=1)
        labels3 = dict(self.WORD_FIELDS)
        cell = self._field(form3, labels3["section_layouts"])
        cell.grid(row=0, column=0, sticky="ew", pady=6)
        self._widgets["section_layouts"] = self._textbox(cell, 9)
        cell = self._field(form3, labels3["signer_title"])
        cell.grid(row=1, column=0, sticky="ew", pady=6)
        self._widgets["signer_title"] = self._entry(cell)

        # ── Bộ lọc feedback
        c2 = card(body)
        c2.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        section_title(c2, "Bộ lọc Feedback",
                      "Cách nhận diện cột ý kiến và loại bỏ các câu trả lời vô nghĩa").pack(
            anchor="w", padx=20, pady=(18, 12))
        form2 = ctk.CTkFrame(c2, fg_color="transparent")
        form2.pack(fill="x", padx=20, pady=(0, 20))
        form2.grid_columnconfigure(0, weight=1)
        for row, (key, label, hint, height) in enumerate([
            ("search_keywords", "Từ khóa tìm cột feedback",
             "Cách nhau bằng dấu phẩy — vd: ý kiến đóng góp, feedback, nhận xét", 0),
            ("skip_keywords", "Từ khóa bỏ qua",
             "Cách nhau bằng dấu phẩy — loại phản hồi trùng khớp hoàn toàn", 0),
            ("skip_patterns", "Biểu thức chính quy bỏ qua",
             "Mỗi dòng một regex — loại phản hồi khớp mẫu", 4),
        ]):
            cell = self._field(form2, label, hint)
            cell.grid(row=row, column=0, sticky="ew", pady=6)
            self._widgets[key] = self._textbox(cell, height) if height else self._entry(cell)

        self.load()

    # ── Factory ─────────────────────────────────────────────────

    def _field(self, parent, label, hint=None):
        cell = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(cell, text=label, font=F(12, "bold"), text_color=C["text"],
                     anchor="w").pack(anchor="w", pady=(0, 4))
        if hint:
            ctk.CTkLabel(cell, text=hint, font=F(11), text_color=C["muted"],
                         anchor="w").pack(anchor="w", pady=(0, 4))
        return cell

    def _entry(self, parent):
        e = ctk.CTkEntry(parent, height=36, corner_radius=8, border_color=C["border"],
                         fg_color=C["card_alt"], text_color=C["text"], font=F(13))
        e.pack(fill="x")
        return e

    def _textbox(self, parent, lines):
        t = ctk.CTkTextbox(parent, height=24 * lines + 14, corner_radius=8, border_width=2,
                           border_color=C["border"], fg_color=C["card_alt"],
                           text_color=C["text"], font=F(13))
        t.pack(fill="x")
        return t

    # ── Đọc / ghi giá trị ───────────────────────────────────────

    def _set(self, key, value):
        w = self._widgets[key]
        if isinstance(w, ctk.CTkTextbox):
            w.delete("1.0", "end")
            w.insert("1.0", value)
        else:
            w.delete(0, "end")
            w.insert(0, value)

    def _get(self, key):
        w = self._widgets[key]
        if isinstance(w, ctk.CTkTextbox):
            return w.get("1.0", "end").strip()
        return w.get().strip()

    def load(self):
        cfg = e2w.CONFIG
        self._set("university_name", cfg["university_name"])
        self._set("faculty_name", cfg["faculty_name"])
        self._set("country_header", cfg["country_header"])
        self._set("country_slogan", cfg["country_slogan"])
        self._set("report_title", "\n".join(cfg["report_title_lines"]))
        self._set("table_header_bg", cfg["table_header_bg"])
        self._set("signer_title", cfg["signer_title"])
        self._set("section_layouts", cfg["section_layouts"])
        for key, _, _ in self.NUMBER_FIELDS:
            self._set(key, str(cfg[key]))
        self._font_var.set(cfg["font_name"])
        self._set("search_keywords", ", ".join(cfg["feedback_search_keywords"]))
        self._set("skip_keywords", ", ".join(cfg["feedback_skip_keywords"]))
        self._set("skip_patterns", "\n".join(cfg["feedback_skip_patterns"]))

    def _save(self):
        word = {key: self._get(key) for key, _ in self.WORD_FIELDS}
        word["font_name"] = self._font_var.get()
        word["table_header_bg"] = self._get("table_header_bg").lstrip("#")
        if not re.fullmatch(r"[0-9A-Fa-f]{6}", word["table_header_bg"]):
            messagebox.showerror("Giá trị không hợp lệ",
                                 "Màu nền tiêu đề bảng phải là mã HEX 6 ký tự, ví dụ D9E2F3.")
            return
        try:
            e2w.parse_section_layouts(word["section_layouts"])
        except ValueError as e:
            messagebox.showerror("Nhóm tiêu chí không hợp lệ", str(e))
            return
        for key, label, cast in self.NUMBER_FIELDS:
            try:
                word[key] = cast(self._get(key).replace(",", "."))
            except ValueError:
                messagebox.showerror("Giá trị không hợp lệ", f"“{label}” phải là một số.")
                return
        self.app.settings["word_format"] = word
        self.app.settings["feedback_config"] = {
            "search_keywords": self._get("search_keywords"),
            "skip_keywords": self._get("skip_keywords"),
            "skip_patterns": self._get("skip_patterns"),
        }
        save_settings(self.app.settings)
        e2w.load_settings_into_config()
        self.load()
        self.app.set_status("Đã lưu cấu hình")
        messagebox.showinfo("Đã lưu", "Cấu hình đã được lưu và áp dụng cho các lần xử lý tiếp theo.")

    def _reset(self):
        if not messagebox.askyesno("Khôi phục mặc định",
                                   "Khôi phục toàn bộ cấu hình về mặc định của nhà trường?"):
            return
        self.app.settings.pop("word_format", None)
        self.app.settings.pop("feedback_config", None)
        save_settings(self.app.settings)
        e2w.load_settings_into_config()
        self.load()
        self.app.set_status("Đã khôi phục cấu hình mặc định")


# ═══════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════

class App(ctk.CTk):
    PAGES = [
        ("report", "📄", "Tổng hợp báo cáo", ReportPage),
        ("feedback", "💬", "Gộp Feedbacks", FeedbackPage),
        ("settings", "⚙️", "Cấu hình", SettingsPage),
    ]

    def __init__(self):
        self.settings = load_settings()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        super().__init__(fg_color=C["bg"])
        Fonts.init(self)

        self._queue = queue.Queue()
        self.title(f"{APP_NAME} — Excel → Word")
        self.geometry("1180x760")
        self.minsize(1000, 660)
        self._set_icon()

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=0, column=1, sticky="nsew", padx=28, pady=(24, 0))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self._pages = {}
        for key, _, _, cls in self.PAGES:
            page = cls(content, self)
            page.grid(row=0, column=0, sticky="nsew")
            self._pages[key] = page

        self._status = ctk.CTkLabel(self, text="Sẵn sàng", font=F(12), text_color=C["muted"],
                                    anchor="w")
        self._status.grid(row=1, column=1, sticky="ew", padx=30, pady=(8, 10))

        self._show(self._initial_page())
        self._center()
        self.after(50, self._drain_queue)

    # ── Layout ──────────────────────────────────────────────────

    def _build_sidebar(self):
        side = ctk.CTkFrame(self, width=236, corner_radius=0, fg_color=C["sidebar"],
                            border_width=0)
        side.grid(row=0, column=0, rowspan=2, sticky="nsw")
        side.grid_propagate(False)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(2, weight=1)

        brand = ctk.CTkFrame(side, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="ew", padx=20, pady=(26, 28))
        logo = self._logo_image()
        if logo:
            ctk.CTkLabel(brand, text="", image=logo).pack(side="left")
        else:
            ctk.CTkLabel(brand, text="📊", font=F(28)).pack(side="left")
        names = ctk.CTkFrame(brand, fg_color="transparent")
        names.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(names, text=APP_NAME, font=F(15, "bold"), text_color=C["text"],
                     anchor="w").pack(anchor="w")
        ctk.CTkLabel(names, text="Excel → Word • ĐBCL", font=F(11), text_color=C["muted"],
                     anchor="w").pack(anchor="w")

        nav = ctk.CTkFrame(side, fg_color="transparent")
        nav.grid(row=1, column=0, sticky="ew", padx=14)
        self._nav = {}
        for key, icon, label, _ in self.PAGES:
            b = ctk.CTkButton(nav, text=f"  {icon}    {label}", anchor="w", height=42,
                              corner_radius=10, font=F(14), fg_color="transparent",
                              hover_color=C["neutral_btn"], text_color=C["text"],
                              command=lambda k=key: self._show(k))
            b.pack(fill="x", pady=2)
            self._nav[key] = b

        bottom = ctk.CTkFrame(side, fg_color="transparent")
        bottom.grid(row=3, column=0, sticky="ew", padx=20, pady=20)
        links = ctk.CTkFrame(bottom, fg_color="transparent")
        links.pack(fill="x")
        ctk.CTkLabel(links, text=f"Phiên bản {APP_VERSION}", font=F(11),
                     text_color=C["muted"]).pack(side="left")
        ctk.CTkButton(links, text="Mở nhật ký", width=10, height=22, font=F(11),
                      fg_color="transparent", hover_color=C["neutral_btn"],
                      text_color=C["accent"], command=self._open_logs).pack(side="right")

    def _logo_image(self):
        try:
            from PIL import Image
            img = Image.open(resource_path(os.path.join("assets", "icon_256.png")))
            return ctk.CTkImage(light_image=img, size=(40, 40))
        except Exception:
            return None

    def _set_icon(self):
        try:
            if sys.platform == "win32":
                self.iconbitmap(resource_path(os.path.join("assets", "icon.ico")))
            else:
                self._icon = tk.PhotoImage(file=resource_path(os.path.join("assets", "icon_256.png")))
                self.iconphoto(True, self._icon)
        except Exception:
            pass

    def _center(self):
        self.update_idletasks()
        w, h = 1180, 760
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _initial_page(self):
        """Cho phép mở sẵn một trang (DBCL_START_TAB hoặc --tab N) — phục vụ chụp ảnh tài liệu."""
        start = os.environ.get("DBCL_START_TAB")
        if start is None and "--tab" in sys.argv:
            idx = sys.argv.index("--tab") + 1
            start = sys.argv[idx] if idx < len(sys.argv) else None
        try:
            return self.PAGES[int(start)][0]
        except (TypeError, ValueError, IndexError):
            return self.PAGES[0][0]

    def _show(self, key):
        self._pages[key].tkraise()
        for k, b in self._nav.items():
            active = k == key
            b.configure(fg_color=C["accent_soft"] if active else "transparent",
                        text_color=C["accent"] if active else C["text"],
                        font=F(14, "bold" if active else "normal"))

    def _open_logs(self):
        try:
            open_path(get_log_dir())
        except Exception as e:
            messagebox.showinfo("Nhật ký", f"Không mở được thư mục nhật ký:\n{e}")

    # ── Trạng thái & cài đặt ────────────────────────────────────

    def update_setting(self, key, value):
        self.settings[key] = value
        save_settings(self.settings)

    def set_status(self, text):
        self._status.configure(text=text)

    # ── Hàng đợi cập nhật từ luồng nền ──────────────────────────

    def post(self, fn):
        """Lên lịch chạy fn trên luồng giao diện (an toàn khi gọi từ luồng nền)."""
        self._queue.put(fn)

    def post_status(self, text):
        self.post(lambda: self.set_status(text))

    def _drain_queue(self):
        try:
            while True:
                self._queue.get_nowait()()
        except queue.Empty:
            pass
        self.after(50, self._drain_queue)


# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════

def main():
    """Điểm khởi chạy ứng dụng GUI."""
    try:
        App().mainloop()
    except Exception:
        _fatal("Lỗi hệ thống",
               f"Ứng dụng gặp lỗi nghiêm trọng:\n\n{traceback.format_exc()}")


if __name__ == "__main__":
    main()
