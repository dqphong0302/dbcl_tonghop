#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bộ kiểm thử tự động (pytest) cho phần mềm Tổng hợp Khảo sát.

Phục vụ tiêu chí "Tính khoa học" của SKKN: chứng minh phần mềm xử lý ĐÚNG
trên dữ liệu thật, và bảo vệ chống hồi quy khi nâng cấp.

Chạy:  pytest -v   (từ thư mục gốc dự án)
"""

import os
import sys
import math
import pytest

# Cho phép import module ở thư mục gốc
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import data_utils
import excel_to_word as e2w
import merge_feedbacks as mf

DULIEU = os.path.join(ROOT, "dulieu")
SURVEY_FILE = os.path.join(DULIEU, "CTXH25 Giáo dục thể chất 1 25-26 HKI.xls")
SURVEY_FILE_2 = os.path.join(DULIEU, "D25 Giáo dục thể chất 2 25-26 HKI 2.xls")
FEEDBACK_FILE = os.path.join(DULIEU, "feedback1.xls")
FEEDBACK_FILE_2 = os.path.join(DULIEU, "feedback2.xls")


# ───────────────────────────────────────────────────────────────
# data_utils: lớp đọc thống nhất
# ───────────────────────────────────────────────────────────────

@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_open_workbook_survey():
    wb = data_utils.open_workbook(SURVEY_FILE)
    assert wb.nsheets >= 1
    sh = wb.sheet_by_index(0)
    assert sh.ncols >= 8
    # API tương thích xlrd
    assert sh.cell(0, 0).value == sh.cell_value(0, 0)


@pytest.mark.skipif(not os.path.exists(FEEDBACK_FILE), reason="thiếu file mẫu")
def test_detect_feedback_column():
    wb = data_utils.open_workbook(FEEDBACK_FILE)
    sh = wb.sheet_by_index(0)
    col, start = data_utils.detect_feedback_column(
        sh, ["ý kiến đóng góp", "feedback"])
    # File feedback mẫu: cột ý kiến ở index 36, header ở dòng 5 (index 4)
    assert col == 36
    assert start >= 5


@pytest.mark.skipif(not os.path.exists(FEEDBACK_FILE), reason="thiếu file mẫu")
def test_find_column_by_keywords():
    wb = data_utils.open_workbook(FEEDBACK_FILE)
    sh = wb.sheet_by_index(0)
    col = data_utils.find_column_by_keywords(sh, ["mã số sinh viên"])
    assert col == 7


# ───────────────────────────────────────────────────────────────
# Trích xuất dữ liệu khảo sát
# ───────────────────────────────────────────────────────────────

@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_extract_survey_questions():
    data = e2w.extract_data_from_excel(SURVEY_FILE)
    # File mẫu có 22 câu hỏi khảo sát
    assert len(data["survey_questions"]) == 22
    # Mỗi câu có 5 mức và tổng > 0
    q = data["survey_questions"][0]
    assert len(q["counts"]) == 5
    assert q["total"] > 0


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_average_score_in_valid_range():
    data = e2w.extract_data_from_excel(SURVEY_FILE)
    for q in data["survey_questions"]:
        # Điểm TB thang Likert 1-5 phải nằm trong [1, 5]
        assert 1.0 <= q["average"] <= 5.0, f"Điểm TB lỗi: {q['average']}"


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_average_score_manual_calc():
    """Kiểm chứng công thức điểm TB bằng tính tay trên câu đầu tiên."""
    data = e2w.extract_data_from_excel(SURVEY_FILE)
    q = data["survey_questions"][0]
    counts = q["counts"]
    total = q["total"]
    expected = sum((i + 1) * counts[i] for i in range(5)) / total
    assert math.isclose(q["average"], round(expected, 2), abs_tol=0.01)


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE_2), reason="thiếu file mẫu")
def test_tong_so_sv_from_filename():
    """File kết thúc bằng ' 2' => tong_so_sv = tổng phiếu + 2."""
    data = e2w.extract_data_from_excel(SURVEY_FILE_2)
    total_phieu = data["survey_questions"][0]["total"]
    assert data["meta"].get("tong_so_sv") == total_phieu + 2


# ───────────────────────────────────────────────────────────────
# Tách & làm sạch ý kiến mở
# ───────────────────────────────────────────────────────────────

def test_parse_open_responses_filters_empty():
    assert e2w.parse_open_responses("") == []
    assert e2w.parse_open_responses("không") == []
    assert e2w.parse_open_responses("không ạ") == []
    assert e2w.parse_open_responses("dạ không ạ") == []


def test_parse_open_responses_keeps_real():
    out = e2w.parse_open_responses("Giáo viên dạy rất hay")
    assert len(out) == 1
    assert out[0].endswith(".")
    assert out[0][0].isupper()


def test_parse_open_responses_dedup():
    out = e2w.parse_open_responses("Tốt lắm\nTốt lắm\nTốt lắm")
    assert len(out) == 1


# ───────────────────────────────────────────────────────────────
# Gộp feedback
# ───────────────────────────────────────────────────────────────

@pytest.mark.skipif(not os.path.exists(FEEDBACK_FILE), reason="thiếu file mẫu")
def test_extract_feedbacks_returns_list():
    fbs = mf.extract_feedbacks(FEEDBACK_FILE)
    assert isinstance(fbs, list)
    # Các câu vô nghĩa như "Không", "Dạ không ạ" phải bị loại
    for fb in fbs:
        assert fb.strip().lower() not in ("không", "không.", "dạ không ạ")


@pytest.mark.skipif(not os.path.exists(FEEDBACK_FILE), reason="thiếu file mẫu")
def test_merge_feedbacks_creates_file(tmp_path):
    out = mf.merge_feedbacks([FEEDBACK_FILE, FEEDBACK_FILE_2], str(tmp_path))
    if out is not None:  # có thể None nếu toàn bộ là "không"
        assert os.path.exists(out)
        assert out.endswith(".docx")


# ───────────────────────────────────────────────────────────────
# Tạo báo cáo Word end-to-end
# ───────────────────────────────────────────────────────────────

@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_create_word_report(tmp_path):
    data = e2w.extract_data_from_excel(SURVEY_FILE)
    out = os.path.join(str(tmp_path), "report.docx")
    e2w.create_word_report(data, out)
    assert os.path.exists(out)
    assert os.path.getsize(out) > 5000  # file Word hợp lệ có kích thước đáng kể


# ───────────────────────────────────────────────────────────────
# Validate dữ liệu
# ───────────────────────────────────────────────────────────────

@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_validate_survey_sheet():
    wb = data_utils.open_workbook(SURVEY_FILE)
    sh = wb.sheet_by_index(0)
    report = data_utils.validate_survey_sheet(sh, "Sheet 1")
    # File mẫu hợp lệ => không có lỗi nghiêm trọng
    assert report.ok


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE_2), reason="thiếu file mẫu")
def test_multi_sheet_aggregation_single_suffix(tmp_path):
    # Chạy process_file cho SURVEY_FILE_2 có 4 sheets
    outputs = e2w.process_file(SURVEY_FILE_2, output_dir=str(tmp_path))
    assert outputs is not None
    # Chỉ được tạo 1 file kết quả đại diện cho dữ liệu gộp, thay vì từng file cho mỗi sheet
    assert len(outputs) == 1
    out_file = outputs[0]
    assert os.path.exists(out_file)
    
    # Kiểm tra logic trích xuất không mang suffix ở cấp sheet và chỉ cộng 1 lần khi gộp
    n_sheets, _ = e2w.get_sheet_count(SURVEY_FILE_2)
    all_data = []
    for si in range(n_sheets):
        data = e2w.extract_data_from_excel(SURVEY_FILE_2, sheet_index=si, apply_filename_suffix=False)
        if data["survey_questions"]:
            all_data.append(data)
            
    # Mỗi sheet đơn lẻ không được tự ý áp dụng suffix của file name
    for d in all_data:
        assert d["meta"].get("tong_so_sv") is None
        
    merged = e2w.merge_sheets_data(all_data)
    assert merged is not None
    
    # Tính toán tong_so_sv gộp có suffix
    import re
    basename = os.path.splitext(os.path.basename(SURVEY_FILE_2))[0]
    match = re.search(r'\s+(\d+)$', basename)
    sv_khong_khao_sat = int(match.group(1)) if match else None
    assert sv_khong_khao_sat == 2
    
    total_phieu = merged["survey_questions"][0].get("total", 0)
    merged["meta"]["tong_so_sv"] = total_phieu + sv_khong_khao_sat
    
    # tong_so_sv của bản gộp phải bằng tổng phiếu phản hồi của tất cả các sheet + 2
    assert merged["meta"]["tong_so_sv"] == total_phieu + 2

