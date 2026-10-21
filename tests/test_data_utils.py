#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kiểm thử data_utils.py (SheetReader, Validation, Keyword Search).
"""
import os
import sys
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import data_utils


def test_sheet_reader_basic():
    rows = [
        ["STT", "Họ đệm", "Tên"],
        [1, "Nguyễn Văn", "An"],
        [2, "Trần Thị", "Bình", "Ghi chú thêm"]
    ]
    sheet = data_utils.SheetReader("TestSheet", rows)
    assert sheet.name == "TestSheet"
    assert sheet.nrows == 3
    assert sheet.ncols == 4
    assert sheet.cell_value(0, 1) == "Họ đệm"
    assert sheet.cell(1, 2).value == "An"
    # Out of range test
    assert sheet.cell_value(10, 10) == ""
    # row_values test
    r0 = sheet.row_values(0)
    assert len(r0) == 4
    assert r0[3] == ""


def test_find_column_and_header():
    rows = [
        ["Tiêu đề trường", "Khoa KHCB"],
        ["STT", "Mã số sinh viên", "Ý kiến đóng góp"],
        [1, "202601", "Tốt"]
    ]
    sheet = data_utils.SheetReader("TestSheet", rows)
    header_r = data_utils.find_header_row(sheet, ["stt", "mã số sinh viên"])
    assert header_r == 1

    col_idx = data_utils.find_column_by_keywords(sheet, ["ý kiến đóng góp"])
    assert col_idx == 2

    # Không tìm thấy
    assert data_utils.find_column_by_keywords(sheet, ["không tồn tại"]) is None


def test_validation_report():
    report = data_utils.ValidationReport(filepath="test.xls")
    assert report.ok
    report.add("warning", "A1", "Cảnh báo nhẹ")
    assert report.ok
    assert len(report.warnings) == 1
    report.add("error", "B2", "Lỗi thiếu dữ liệu")
    assert not report.ok
    assert len(report.errors) == 1
    summary = report.summary()
    assert "1 lỗi" in summary


# ── Đọc file lỗi / sai định dạng ─────────────────────────────────
import shutil
from data_utils import ExcelReadError, open_workbook

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REAL_XLS = os.path.join(ROOT_DIR, "dulieu", "D25 Giáo dục thể chất 2 25-26 HKI 2.xls")


@pytest.mark.parametrize("content, message", [
    (b"", "rỗng"),
    (b"hello world", "Không phải file Excel"),
])
def test_bad_files_give_vietnamese_error(tmp_path, content, message):
    f = tmp_path / "x.xls"
    f.write_bytes(content)
    with pytest.raises(ExcelReadError, match=message):
        open_workbook(str(f))


def test_truncated_xls_is_reported_as_corrupt(tmp_path):
    f = tmp_path / "cut.xls"
    f.write_bytes(open(REAL_XLS, "rb").read()[:3000])
    with pytest.raises(ExcelReadError, match="hỏng"):
        open_workbook(str(f))


def test_format_detected_by_content_not_extension(tmp_path):
    f = tmp_path / "doi_duoi.xlsx"            # thực chất là .xls
    shutil.copy(REAL_XLS, f)
    assert open_workbook(str(f)).nsheets == open_workbook(REAL_XLS).nsheets


def test_html_disguised_as_xls(tmp_path):
    f = tmp_path / "web.xls"
    f.write_text('<html><meta charset="utf-8"><table><tr><td colspan="2">Lớp</td><td>3</td></tr>'
                 '<tr><td>A</td><td>B</td><td>1,5</td></tr></table></html>', encoding="utf-8")
    sh = open_workbook(str(f)).sheet_by_index(0)
    assert sh.cell_value(0, 0) == "Lớp" and sh.cell_value(0, 2) == 3.0
    assert sh.cell_value(1, 2) == 1.5                      # dấu phẩy thập phân kiểu Việt
