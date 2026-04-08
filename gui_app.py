#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI Tổng hợp Khảo sát Excel → Word
Giao diện đồ họa cho script excel_to_word.py
Đóng gói chạy trên Windows không cần cài Python.
"""

import sys
import os
import threading
import glob
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# Import logic từ excel_to_word
from excel_to_word import (
    extract_data_from_excel,
    create_word_report,
    merge_sheets_data,
    parse_open_responses,
    get_sheet_count,
)


# ═══════════════════════════════════════════════════════════════
# THEME & COLORS
# ═══════════════════════════════════════════════════════════════

COLORS = {
    "bg": "#1a1b2e",
    "bg_card": "#232545",
    "bg_input": "#2a2d52",
    "bg_hover": "#33366a",
    "accent": "#6c63ff",
    "accent_hover": "#8b83ff",
    "accent_green": "#4ade80",
    "accent_red": "#f87171",
    "accent_yellow": "#fbbf24",
    "text": "#e2e8f0",
    "text_dim": "#94a3b8",
    "text_bright": "#ffffff",
    "border": "#3d4070",
    "progress_bg": "#2a2d52",
    "progress_fill": "#6c63ff",
    "log_bg": "#12132a",
    "success": "#10b981",
    "warning": "#f59e0b",
    "error": "#ef4444",
}

FONT_FAMILY = "Segoe UI"
FONT_FAMILY_MONO = "Consolas"


# ═══════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("📊 Tổng hợp Khảo sát Excel → Word")
        self.geometry("820x720")
        self.minsize(700, 600)
        self.configure(bg=COLORS["bg"])
        
        # Icon (nếu có)
        try:
            if sys.platform == "win32":
                self.iconbitmap(self._resource_path("icon.ico"))
        except Exception:
            pass
        
        # State
        self.input_files = []
        self.output_dir = tk.StringVar(value="")
        self.is_processing = False
        
        self._setup_styles()
        self._build_ui()
        self._center_window()
    
    def _resource_path(self, relative_path):
        """Lấy path tương thích PyInstaller."""
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)
    
    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"+{x}+{y}")
    
    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        
        # Progress bar
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=COLORS["progress_bg"],
                        background=COLORS["accent"],
                        borderwidth=0,
                        thickness=8)
    
    def _build_ui(self):
        # Main container with padding
        main = tk.Frame(self, bg=COLORS["bg"], padx=24, pady=16)
        main.pack(fill="both", expand=True)
        
        # ── HEADER ──
        header = tk.Frame(main, bg=COLORS["bg"])
        header.pack(fill="x", pady=(0, 16))
        
        tk.Label(header, text="📊", font=(FONT_FAMILY, 28),
                 bg=COLORS["bg"], fg=COLORS["accent"]).pack(side="left")
        
        title_frame = tk.Frame(header, bg=COLORS["bg"])
        title_frame.pack(side="left", padx=(12, 0))
        
        tk.Label(title_frame, text="Tổng hợp Khảo sát",
                 font=(FONT_FAMILY, 18, "bold"),
                 bg=COLORS["bg"], fg=COLORS["text_bright"]).pack(anchor="w")
        tk.Label(title_frame, text="Excel (.xls) → Word (.docx) • Hỗ trợ nhiều sheet",
                 font=(FONT_FAMILY, 10),
                 bg=COLORS["bg"], fg=COLORS["text_dim"]).pack(anchor="w")
        
        # ── INPUT SECTION ──
        self._build_section(main, "📁 File Excel đầu vào", self._build_input_area)
        
        # ── OUTPUT SECTION ──
        self._build_section(main, "📂 Thư mục lưu kết quả", self._build_output_area)
        
        # ── ACTION BUTTONS ──
        btn_frame = tk.Frame(main, bg=COLORS["bg"])
        btn_frame.pack(fill="x", pady=(8, 8))
        
        self.btn_process = self._make_button(
            btn_frame, "▶  Bắt đầu xử lý", self._start_processing,
            bg=COLORS["accent"], fg="#ffffff", font_size=12, padx=32, pady=10
        )
        self.btn_process.pack(side="left")
        
        self.btn_clear = self._make_button(
            btn_frame, "🗑 Xóa danh sách", self._clear_files,
            bg=COLORS["bg_card"], fg=COLORS["text_dim"], font_size=10, padx=16, pady=8
        )
        self.btn_clear.pack(side="right")
        
        self.btn_open_output = self._make_button(
            btn_frame, "📂 Mở thư mục kết quả", self._open_output_dir,
            bg=COLORS["bg_card"], fg=COLORS["text_dim"], font_size=10, padx=16, pady=8
        )
        self.btn_open_output.pack(side="right", padx=(0, 8))
        
        # ── PROGRESS ──
        self.progress_frame = tk.Frame(main, bg=COLORS["bg"])
        self.progress_frame.pack(fill="x", pady=(0, 4))
        
        self.progress_label = tk.Label(self.progress_frame, text="",
                                       font=(FONT_FAMILY, 9),
                                       bg=COLORS["bg"], fg=COLORS["text_dim"])
        self.progress_label.pack(anchor="w")
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, 
                                            style="Custom.Horizontal.TProgressbar",
                                            mode="determinate", length=300)
        self.progress_bar.pack(fill="x", pady=(2, 0))
        
        # ── LOG OUTPUT ──
        log_label = tk.Label(main, text="📋 Nhật ký xử lý",
                             font=(FONT_FAMILY, 10, "bold"),
                             bg=COLORS["bg"], fg=COLORS["text"], anchor="w")
        log_label.pack(fill="x", pady=(8, 4))
        
        log_container = tk.Frame(main, bg=COLORS["border"], bd=1, relief="solid")
        log_container.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(log_container, wrap="word",
                                font=(FONT_FAMILY_MONO, 9),
                                bg=COLORS["log_bg"], fg=COLORS["text"],
                                insertbackground=COLORS["text"],
                                selectbackground=COLORS["accent"],
                                relief="flat", bd=8,
                                state="disabled", height=8)
        
        scrollbar = tk.Scrollbar(log_container, command=self.log_text.yview,
                                 bg=COLORS["bg_card"], troughcolor=COLORS["bg_card"])
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        
        # Configure log tags
        self.log_text.tag_configure("success", foreground=COLORS["success"])
        self.log_text.tag_configure("error", foreground=COLORS["error"])
        self.log_text.tag_configure("warning", foreground=COLORS["warning"])
        self.log_text.tag_configure("info", foreground=COLORS["accent"])
        self.log_text.tag_configure("dim", foreground=COLORS["text_dim"])
        
        # ── STATUS BAR ──
        self.status_var = tk.StringVar(value="Sẵn sàng • Kéo thả hoặc chọn file Excel để bắt đầu")
        status = tk.Label(main, textvariable=self.status_var,
                          font=(FONT_FAMILY, 9),
                          bg=COLORS["bg"], fg=COLORS["text_dim"], anchor="w")
        status.pack(fill="x", pady=(6, 0))
        
        # ── DRAG & DROP (Windows) ──
        self._setup_dnd()
    
    def _build_section(self, parent, title, builder_fn):
        """Build a collapsible card section."""
        card = tk.Frame(parent, bg=COLORS["bg_card"], bd=1,
                        highlightbackground=COLORS["border"],
                        highlightthickness=1)
        card.pack(fill="x", pady=(0, 10))
        
        # Header
        hdr = tk.Frame(card, bg=COLORS["bg_card"], padx=14, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text=title, font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["bg_card"], fg=COLORS["text"]).pack(anchor="w")
        
        # Content
        content = tk.Frame(card, bg=COLORS["bg_card"], padx=14, pady=0)
        content.pack(fill="x", pady=(0, 10))
        builder_fn(content)
    
    def _build_input_area(self, parent):
        row = tk.Frame(parent, bg=COLORS["bg_card"])
        row.pack(fill="x")
        
        self._make_button(row, "➕ Thêm file", self._add_files,
                          bg=COLORS["accent"], fg="#fff", font_size=9, padx=14, pady=5
                          ).pack(side="left")
        
        self._make_button(row, "📁 Thêm thư mục", self._add_folder,
                          bg=COLORS["bg_input"], fg=COLORS["text"], font_size=9, padx=14, pady=5
                          ).pack(side="left", padx=(6, 0))
        
        self.file_count_label = tk.Label(row, text="Chưa có file nào",
                                         font=(FONT_FAMILY, 9),
                                         bg=COLORS["bg_card"], fg=COLORS["text_dim"])
        self.file_count_label.pack(side="right")
        
        # File listbox
        list_frame = tk.Frame(parent, bg=COLORS["bg_input"], bd=1, relief="flat")
        list_frame.pack(fill="x", pady=(8, 0))
        
        self.file_listbox = tk.Listbox(list_frame, height=4,
                                       font=(FONT_FAMILY, 9),
                                       bg=COLORS["bg_input"], fg=COLORS["text"],
                                       selectbackground=COLORS["accent"],
                                       selectforeground="#fff",
                                       highlightthickness=0,
                                       relief="flat", bd=0)
        self.file_listbox.pack(fill="x", padx=8, pady=4)
    
    def _build_output_area(self, parent):
        row = tk.Frame(parent, bg=COLORS["bg_card"])
        row.pack(fill="x")
        
        entry = tk.Entry(row, textvariable=self.output_dir,
                         font=(FONT_FAMILY, 10),
                         bg=COLORS["bg_input"], fg=COLORS["text"],
                         insertbackground=COLORS["text"],
                         relief="flat", bd=0)
        entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        
        self._make_button(row, "Chọn...", self._choose_output_dir,
                          bg=COLORS["bg_input"], fg=COLORS["text"], font_size=9, padx=14, pady=5
                          ).pack(side="right")
    
    def _make_button(self, parent, text, command, bg, fg, font_size=10, padx=16, pady=6):
        btn = tk.Label(parent, text=text, font=(FONT_FAMILY, font_size),
                       bg=bg, fg=fg, padx=padx, pady=pady, cursor="hand2")
        btn.bind("<Button-1>", lambda e: command())
        btn.bind("<Enter>", lambda e: btn.configure(
            bg=COLORS["accent_hover"] if bg == COLORS["accent"] else COLORS["bg_hover"]))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn
    
    # ── FILE OPERATIONS ──
    
    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        for f in files:
            if f not in self.input_files:
                self.input_files.append(f)
        self._update_file_list()
    
    def _add_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file Excel")
        if folder:
            for ext in ("*.xls", "*.xlsx"):
                for f in glob.glob(os.path.join(folder, ext)):
                    if f not in self.input_files:
                        self.input_files.append(f)
            self._update_file_list()
    
    def _choose_output_dir(self):
        folder = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if folder:
            self.output_dir.set(folder)
    
    def _clear_files(self):
        self.input_files.clear()
        self._update_file_list()
    
    def _update_file_list(self):
        self.file_listbox.delete(0, "end")
        for f in self.input_files:
            self.file_listbox.insert("end", f"  {os.path.basename(f)}")
        
        n = len(self.input_files)
        if n == 0:
            self.file_count_label.config(text="Chưa có file nào", fg=COLORS["text_dim"])
        else:
            self.file_count_label.config(text=f"📄 {n} file", fg=COLORS["accent"])
    
    def _open_output_dir(self):
        out = self.output_dir.get()
        if out and os.path.isdir(out):
            if sys.platform == "win32":
                os.startfile(out)
            elif sys.platform == "darwin":
                os.system(f'open "{out}"')
            else:
                os.system(f'xdg-open "{out}"')
        else:
            messagebox.showinfo("Thông báo", "Chưa có thư mục kết quả hoặc thư mục chưa tồn tại.")
    
    def _setup_dnd(self):
        """Setup drag & drop - works natively on some platforms."""
        # Basic drop support via tkdnd if available
        try:
            self.drop_target_register = None  # Placeholder
        except Exception:
            pass
    
    # ── LOG ──
    
    def _log(self, msg, tag=None):
        """Thread-safe logging."""
        def _do():
            self.log_text.configure(state="normal")
            if tag:
                self.log_text.insert("end", msg + "\n", tag)
            else:
                self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, _do)
    
    def _set_progress(self, value, text=""):
        def _do():
            self.progress_bar["value"] = value
            if text:
                self.progress_label.config(text=text)
        self.after(0, _do)
    
    def _set_status(self, text):
        self.after(0, lambda: self.status_var.set(text))
    
    # ── PROCESSING ──
    
    def _start_processing(self):
        if self.is_processing:
            return
        
        if not self.input_files:
            messagebox.showwarning("Chưa có file", "Vui lòng thêm file Excel trước khi xử lý.")
            return
        
        out_dir = self.output_dir.get().strip()
        if not out_dir:
            # Default: cùng thư mục với file đầu tiên
            out_dir = os.path.join(os.path.dirname(self.input_files[0]), "output")
            self.output_dir.set(out_dir)
        
        os.makedirs(out_dir, exist_ok=True)
        
        self.is_processing = True
        self.btn_process.configure(bg=COLORS["text_dim"])
        
        # Clear log
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        
        thread = threading.Thread(target=self._process_worker, args=(list(self.input_files), out_dir), daemon=True)
        thread.start()
    
    def _process_worker(self, files, out_dir):
        """Worker thread xử lý file."""
        total_files = len(files)
        success_count = 0
        fail_count = 0
        total_outputs = 0
        
        self._log(f"{'═' * 50}", "dim")
        self._log(f"  Bắt đầu xử lý {total_files} file Excel", "info")
        self._log(f"  Thư mục kết quả: {out_dir}", "dim")
        self._log(f"{'═' * 50}", "dim")
        
        for fi, filepath in enumerate(files):
            basename = os.path.splitext(os.path.basename(filepath))[0]
            progress_pct = int((fi / total_files) * 100)
            self._set_progress(progress_pct, f"Đang xử lý {fi + 1}/{total_files}: {os.path.basename(filepath)}")
            self._set_status(f"Đang xử lý file {fi + 1}/{total_files}...")
            
            try:
                n_sheets, sheet_names = get_sheet_count(filepath)
                self._log(f"\n📄 {os.path.basename(filepath)} ({n_sheets} sheet)", "info")
                
                all_data = []
                
                for si in range(n_sheets):
                    try:
                        data = extract_data_from_excel(filepath, sheet_index=si)
                    except Exception as e:
                        self._log(f"   ⚠ Sheet {si} \"{sheet_names[si]}\": lỗi đọc - {e}", "warning")
                        continue
                    
                    if not data["survey_questions"]:
                        self._log(f"   ⏭ Sheet {si} \"{sheet_names[si]}\": bỏ qua (không có dữ liệu)", "dim")
                        continue
                    
                    all_data.append(data)
                    n_q = len(data["survey_questions"])
                    n_resp = len(parse_open_responses(data["open_responses"]))
                    lop = data["meta"].get("lop_hoc", sheet_names[si])
                    
                    # Hiển thị cảnh báo (ví dụ: ô F5 trống)
                    for warn_msg in data.get("warnings", []):
                        self._log(f"   {warn_msg}", "warning")
                    
                    if n_sheets > 1:
                        out_path = os.path.join(out_dir, f"{basename}_sheet{si + 1}.docx")
                        create_word_report(data, out_path)
                        total_outputs += 1
                        self._log(f"   ✅ Sheet {si} \"{lop}\": {n_q} câu → {os.path.basename(out_path)}", "success")
                    else:
                        out_path = os.path.join(out_dir, f"{basename}.docx")
                        create_word_report(data, out_path)
                        total_outputs += 1
                        self._log(f"   ✅ {n_q} câu, {n_resp} ý kiến → {os.path.basename(out_path)}", "success")
                
                # Tạo file tổng hợp
                if len(all_data) > 1:
                    merged = merge_sheets_data(all_data)
                    if merged:
                        sum_path = os.path.join(out_dir, f"{basename}_TONG_HOP.docx")
                        create_word_report(merged, sum_path)
                        total_outputs += 1
                        total_sv = merged["survey_questions"][0]["total"] if merged["survey_questions"] else 0
                        self._log(f"   📊 TỔNG HỢP: {len(all_data)} lớp, {total_sv} SV → {os.path.basename(sum_path)}", "info")
                
                success_count += 1
                
            except Exception as e:
                fail_count += 1
                self._log(f"   ❌ Lỗi: {e}", "error")
        
        # Done
        self._set_progress(100, "Hoàn tất!")
        self._log(f"\n{'═' * 50}", "dim")
        self._log(f"  ✅ Hoàn tất: {success_count} thành công, {fail_count} thất bại", "success")
        self._log(f"  📁 Tổng file tạo ra: {total_outputs}", "info")
        self._log(f"  📂 Kết quả: {out_dir}", "dim")
        self._log(f"{'═' * 50}", "dim")
        
        self._set_status(f"✅ Hoàn tất • {total_outputs} file đã tạo tại {out_dir}")
        
        self.after(0, lambda: self.btn_process.configure(bg=COLORS["accent"]))
        self.is_processing = False


# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
