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
