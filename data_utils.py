#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tiện ích đọc dữ liệu & kiểm tra (validate) cho phần mềm Tổng hợp Khảo sát.

Mục đích (phục vụ tiêu chí "Tính khoa học" của SKKN):
- Dò cột TỰ ĐỘNG theo TIÊU ĐỀ thay vì hard-code index (cột AK = 36, ô F5...).
  -> Bền vững khi cấu trúc file thay đổi, có fallback an toàn.
- VALIDATE dữ liệu đầu vào: phát hiện file lỗi, thiếu cột, tổng % bất thường.
  -> Tạo "báo cáo kiểm tra dữ liệu" minh bạch, đáng tin cậy.

Hỗ trợ cả .xls (xlrd) và .xlsx (openpyxl) qua một lớp đọc thống nhất.
"""

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

# ───────────────────────────────────────────────────────────────
# LỚP ĐỌC THỐNG NHẤT .xls / .xlsx
# ───────────────────────────────────────────────────────────────


class _Cell:
    """Đối tượng ô tương thích xlrd: cho phép truy cập .value."""
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class SheetReader:
    """Bọc 1 sheet để truy cập đồng nhất bất kể .xls hay .xlsx.

    Tương thích ngược với API của xlrd: hỗ trợ cả cell(r,c).value,
    cell_value(r,c), row_values(r), nrows, ncols, name.
    """

    def __init__(self, name: str, rows: List[List]):
        self.name = name
        self._rows = rows
        self.nrows = len(rows)
        self.ncols = max((len(r) for r in rows), default=0)

    def cell_value(self, r: int, c: int):
        if 0 <= r < self.nrows:
            row = self._rows[r]
            if 0 <= c < len(row):
                return row[c]
        return ""

    def cell(self, r: int, c: int) -> "_Cell":
        """Tương thích xlrd: sheet.cell(r, c).value."""
        return _Cell(self.cell_value(r, c))

    def row_values(self, r: int) -> List:
        if 0 <= r < self.nrows:
            row = list(self._rows[r])
            if len(row) < self.ncols:
                row += [""] * (self.ncols - len(row))
            return row
        return [""] * self.ncols


class WorkbookReader:
    """Đọc workbook .xls/.xlsx thành danh sách SheetReader."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.sheets: List[SheetReader] = []
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".xlsx":
            self._load_xlsx(filepath)
        else:
            self._load_xls(filepath)

    def _load_xls(self, filepath):
        import xlrd
        wb = xlrd.open_workbook(filepath)
        for sh in wb.sheets():
            rows = [sh.row_values(r) for r in range(sh.nrows)]
            self.sheets.append(SheetReader(sh.name, rows))

    def _load_xlsx(self, filepath):
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append(["" if v is None else v for v in row])
            self.sheets.append(SheetReader(ws.title, rows))
        wb.close()

    @property
    def nsheets(self) -> int:
        return len(self.sheets)

    def sheet_names(self) -> List[str]:
        return [s.name for s in self.sheets]

    def sheet_by_index(self, idx: int) -> SheetReader:
        return self.sheets[idx]


def open_workbook(filepath: str) -> WorkbookReader:
    """Mở file Excel (.xls hoặc .xlsx) trả về WorkbookReader thống nhất."""
    return WorkbookReader(filepath)


# ───────────────────────────────────────────────────────────────
# DÒ CỘT TỰ ĐỘNG THEO TIÊU ĐỀ
# ───────────────────────────────────────────────────────────────


def _norm(text) -> str:
    """Chuẩn hóa chuỗi để so khớp: bỏ khoảng trắng thừa, về chữ thường."""
    return re.sub(r"\s+", " ", str(text or "").strip().lower())


def find_column_by_keywords(
    sheet: SheetReader,
    keywords: List[str],
    max_scan_rows: int = 12,
) -> Optional[int]:
    """Tìm chỉ số CỘT đầu tiên mà tiêu đề chứa 1 trong các keyword.

    Quét tối đa `max_scan_rows` dòng đầu (vùng header). Trả về index cột hoặc None.
    """
    kws = [_norm(k) for k in keywords if _norm(k)]
    if not kws:
        return None
    for r in range(min(max_scan_rows, sheet.nrows)):
        for c in range(sheet.ncols):
            val = _norm(sheet.cell_value(r, c))
            if not val:
                continue
            if any(kw in val for kw in kws):
                return c
    return None


def find_header_row(
    sheet: SheetReader,
    keywords: List[str],
    max_scan_rows: int = 12,
) -> Optional[int]:
    """Tìm DÒNG header chứa 1 trong các keyword (vd dòng có 'STT', 'Họ đệm')."""
    kws = [_norm(k) for k in keywords if _norm(k)]
    for r in range(min(max_scan_rows, sheet.nrows)):
        row_text = " ".join(_norm(sheet.cell_value(r, c)) for c in range(sheet.ncols))
        if any(kw in row_text for kw in kws):
            return r
    return None


def detect_feedback_column(
    sheet: SheetReader,
    search_keywords: List[str],
) -> tuple:
    """Dò cột chứa ý kiến mở trong FILE FEEDBACK (raw, mỗi dòng 1 SV).

    Trả về (col_index, start_row).
    Chiến lược:
      1) Tìm theo keyword tiêu đề (vd 'ý kiến đóng góp').
      2) Fallback: cột cuối cùng, bắt đầu sau dòng header dò được (hoặc dòng 5).
    """
    col = find_column_by_keywords(sheet, search_keywords)
    if col is not None:
        # start_row = ngay sau dòng chứa tiêu đề
        header_row = find_header_row(sheet, ["stt", "họ đệm", "mã số sinh viên"])
        start = (header_row + 1) if header_row is not None else 5
        return col, start

    # Fallback an toàn
    header_row = find_header_row(sheet, ["stt", "họ đệm", "mã số sinh viên"])
    start = (header_row + 1) if header_row is not None else 5
    return sheet.ncols - 1, start


# ───────────────────────────────────────────────────────────────
# VALIDATE DỮ LIỆU
# ───────────────────────────────────────────────────────────────


@dataclass
class ValidationIssue:
    level: str            # "error" | "warning" | "info"
    location: str         # vd "Sheet 1", "C5"
    message: str


@dataclass
class ValidationReport:
    filepath: str
    issues: List[ValidationIssue] = field(default_factory=list)

    def add(self, level: str, location: str, message: str):
        self.issues.append(ValidationIssue(level, location, message))

    @property
    def errors(self):
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self):
        return [i for i in self.issues if i.level == "warning"]

    @property
    def ok(self) -> bool:
        return len(self.errors) == 0

    def summary(self) -> str:
        name = os.path.basename(self.filepath)
        return (f"{name}: {len(self.errors)} lỗi, {len(self.warnings)} cảnh báo, "
                f"{len(self.issues)} mục kiểm tra")


def validate_survey_sheet(sheet: SheetReader, sheet_label: str) -> ValidationReport:
    """Kiểm tra 1 sheet khảo sát (dạng đã tổng hợp, 8 cột).

    Kiểm tra: có vùng 'Thông tin khảo sát', tổng tỷ lệ ~ 100%, số cột tối thiểu.
    """
    report = ValidationReport(filepath=sheet_label)

    if sheet.nrows < 6:
        report.add("error", sheet_label, "Sheet quá ít dòng, không giống file khảo sát.")
        return report

    if sheet.ncols < 8:
        report.add("warning", sheet_label,
                   f"Chỉ có {sheet.ncols} cột (kỳ vọng >= 8). Có thể sai định dạng.")

    # Tìm vùng câu hỏi khảo sát
    has_survey = False
    for r in range(sheet.nrows):
        if _norm(sheet.cell_value(r, 0)) == "thông tin khảo sát":
            has_survey = True
            break
    if not has_survey:
        report.add("warning", sheet_label,
                   "Không tìm thấy mục 'Thông tin khảo sát'. Báo cáo có thể thiếu dữ liệu.")

    # Kiểm tra dòng 'Tỷ lệ' tổng ~ 1.0
    checked = 0
    for r in range(sheet.nrows):
        if _norm(sheet.cell_value(r, 1)) == "tỷ lệ":
            try:
                total = float(sheet.cell_value(r, 7))
            except (ValueError, TypeError):
                continue
            checked += 1
            if total > 0 and abs(total - 1.0) > 0.05:
                report.add("warning", f"{sheet_label} dòng {r+1}",
                           f"Tổng tỷ lệ = {total:.4f} (lệch khỏi 100%).")
    if checked:
        report.add("info", sheet_label, f"Đã kiểm tra {checked} dòng tỷ lệ.")

    return report
