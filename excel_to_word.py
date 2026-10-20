#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script tổng hợp kết quả khảo sát từ file Excel (.xls) sang báo cáo Word (.docx)
Dùng cho: Báo cáo tổng hợp kết quả khảo sát sinh viên sau khi kết thúc học phần

Cách sử dụng:
  python3 excel_to_word.py                          # Xử lý tất cả file .xls trong thư mục hiện tại
  python3 excel_to_word.py "file1.xls"              # Xử lý 1 file cụ thể
  python3 excel_to_word.py "file1.xls" "file2.xls"  # Xử lý nhiều file
  python3 excel_to_word.py --input-dir ./data        # Xử lý tất cả .xls trong thư mục chỉ định
  python3 excel_to_word.py --output-dir ./output     # Lưu kết quả vào thư mục chỉ định
"""

import sys
import os
import glob
import argparse
import re
import xlrd
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import copy
from xml.sax.saxutils import escape as xml_escape
import json
from dataclasses import dataclass, field
from typing import List, Optional
from pathlib import Path
from data_utils import open_workbook as _open_workbook
try:
    from app_logging import get_logger
    log = get_logger()
except Exception:
    import logging
    log = logging.getLogger("dbcl")

XML_CLASS_BOX_TEMPLATE = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
       xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
       xmlns:v="urn:schemas-microsoft-com:vml"
       xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:r>
    <w:rPr>
      <w:rFonts w:cs="Times New Roman"/>
      <w:szCs w:val="24"/>
    </w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column">
              <wp:posOffset>-795020</wp:posOffset>
            </wp:positionH>
            <wp:positionV relativeFrom="paragraph">
              <wp:posOffset>-1365885</wp:posOffset>
            </wp:positionV>
            <wp:extent cx="771525" cy="285750"/>
            <wp:effectExtent l="0" t="0" r="28575" b="19050"/>
            <wp:wrapNone/>
            <wp:docPr id="7" name="Text Box 7"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1">
                    <a:spLocks/>
                  </wps:cNvSpPr>
                  <wps:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="771525" cy="285750"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                    <a:solidFill>
                      <a:sysClr val="window" lastClr="FFFFFF"/>
                    </a:solidFill>
                    <a:ln w="6350">
                      <a:solidFill>
                        <a:prstClr val="black"/>
                      </a:solidFill>
                    </a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:jc w:val="center"/>
                        </w:pPr>
                        <w:r>
                          <w:rPr>
                            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                            <w:b/>
                            <w:sz w:val="24"/>
                            <w:szCs w:val="24"/>
                          </w:rPr>
                          <w:t>{CLASS_CODE}</w:t>
                        </w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1">
                    <a:prstTxWarp prst="textNoShape">
                      <a:avLst/>
                    </a:prstTxWarp>
                    <a:noAutofit/>
                  </wps:bodyPr>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
      <mc:Fallback>
        <w:pict>
          <v:shapetype id="_x0000_t202" coordsize="21600,21600" spt="202" path="m,l,21600r21600,l21600,xe">
            <v:stroke joinstyle="miter"/>
            <v:path gradientshapeok="t" connecttype="rect"/>
          </v:shapetype>
          <v:shape id="Text Box 7" type="#_x0000_t202" style="position:absolute;margin-left:-62.6pt;margin-top:-107.55pt;width:60.75pt;height:22.5pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-position-horizontal:absolute;mso-position-vertical:absolute;v-text-anchor:top" fillcolor="window" strokeweight=".5pt">
            <v:textbox>
              <w:txbxContent>
                <w:p>
                  <w:pPr>
                    <w:jc w:val="center"/>
                  </w:pPr>
                  <w:r>
                    <w:rPr>
                      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                      <w:b/>
                      <w:sz w:val="24"/>
                      <w:szCs w:val="24"/>
                    </w:rPr>
                    <w:t>{CLASS_CODE}</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent>
            </v:textbox>
          </v:shape>
        </w:pict>
      </mc:Fallback>
    </mc:AlternateContent>
  </w:r>
</w:p>"""


# ═══════════════════════════════════════════════════════════════════
# CẤU HÌNH - Điều chỉnh theo nhu cầu
# ═══════════════════════════════════════════════════════════════════

DEFAULT_CONFIG = {
    "university_name": "ĐẠI HỌC Y DƯỢC\nTHÀNH PHỐ HỒ CHÍ MINH",
    "faculty_name": "KHOA KHCB",                # Tên khoa/đơn vị
    "country_header": "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
    "country_slogan": "Độc lập – Tự do – Hạnh phúc",
    "report_title_lines": [
        "BÁO CÁO TỔNG HỢP KẾT QUẢ",
        "KHẢO SÁT SINH VIÊN SAU KHI KẾT THÚC HỌC PHẦN CƠ SỞ",
        "& ĐỀ XUẤT KẾ HOẠCH CẢI TIẾN",
    ],
    "font_name": "Times New Roman",
    "font_size": 12,                             # pt
    "margin_top": 2.0,                           # cm
    "margin_bottom": 2.0,                        # cm
    "margin_left": 2.5,                          # cm
    "margin_right": 2.0,                         # cm
    "table_header_bg": "D9E2F3",                 # hex
    "open_response_lines": 4,                    # số dòng trống

    # Cấu hình keywords của feedback
    "feedback_search_keywords": ["ý kiến đóng góp", "feedback"],
    "feedback_skip_keywords": ["không", "không có", "khong", "none", "n/a", "chưa có", "không có ý kiến", "không có ý kiến gì"],
    "feedback_skip_patterns": [
        r'^không\s*(ạ|có|cóa)?\.?$',
        r'^ko\.?$',
        r'^khong\.?$',
        r'^dạ\s*không\s*(ạ|có)?\.?$',
        r'^không\s*có\s*(góp\s*ý|ý\s*kiến)?\.?$',
        r'^dạ\s*không\s*có\.?$',
        r'^không\s*có\s*ý\s*kiến\.?$',
        r'^em\s*không\s*có\s*ý\s*kiến\s*(ạ)?\.?$',
        # Chuỗi chỉ gồm từ từ chối: "KHÔNG Dạ không ạ!", "Hong ạ", "Em không có ạ.không ạ"
        r'^((dạ|thưa|em|mình|tôi|không|khong|khoong|hong|khum|ko|k|có|ạ|aj|nha|nhé)[\s.,!]*)+$',
        # "Em/Mình/Tôi (k/ko/không) có (ý kiến/góp ý) (gì) (ạ/nữa/thêm)"
        r'^(dạ\s+)?(em|mình|tôi)\s+(k|ko|không|chưa)\s+có(\s+(ý\s*kiến|góp\s*ý|đóng\s*góp|gì|ạ|nha|nhé|nữa|thêm|hết))*\s*[.!]*$',
        # "Không có ý kiến…" kèm phần lặp lại dính liền ("Không có ý kiếnem không có ý kiến gì nữa")
        r'^không\s*có\s*(ý\s*kiến|đóng\s*góp|góp\s*ý)\S*(\s+(em|gì|ạ|nha|nhé|nữa|thêm|không|có|ý|kiến|đóng|góp))*\s*[.!]*$',
    ],

    "scale_labels": [                            # Nhãn thang đo Likert 5 mức
        "1. Hoàn toàn không đồng ý",
        "2. Không đồng ý",
        "3. Đồng ý một phần",
        "4. Đồng ý",
        "5. Hoàn toàn đồng ý",
    ],
    # Nhóm tiêu chí trong bảng kết quả. Mỗi khối (cách nhau dòng trống) là một
    # mẫu phiếu; dòng có dạng "Tên nhóm | từ-đến" theo số câu chính (8.1 → 8).
    # Khối được chọn là khối có số câu lớn nhất trùng với số câu lớn nhất của file.
    "section_layouts": (
        "I. CUNG CẤP THÔNG TIN VỀ MÔN HỌC/HỌC PHẦN | 1-10\n"
        "II. PHƯƠNG PHÁP GIẢNG DẠY | 11-14\n"
        "III. LƯỢNG GIÁ HỌC PHẦN | 15-18\n"
        "IV. CẢM NHẬN CHUNG | 19-22\n"
        "\n"
        "I. CUNG CẤP THÔNG TIN VỀ MÔN HỌC/HỌC PHẦN | 1-7\n"
        "II. PHƯƠNG PHÁP GIẢNG DẠY | 8\n"
        "III. LƯỢNG GIÁ HỌC PHẦN | 9-15\n"
        "IV. CẢM NHẬN CHUNG | 16\n"
        "\n"
        "I. CUNG CẤP THÔNG TIN VỀ MÔN HỌC/HỌC PHẦN | 1-10\n"
        "II. PHƯƠNG PHÁP GIẢNG DẠY | 11\n"
        "III. LƯỢNG GIÁ HỌC PHẦN | 12-18\n"
        "IV. CẢM NHẬN CHUNG | 19"
    ),
    "signer_title": "TRƯỞNG MODULE/TRƯỞNG BỘ MÔN",
}

CONFIG = copy.deepcopy(DEFAULT_CONFIG)

SETTINGS_FILE = Path.home() / ".dbcl_tonghop_settings.json"

# Ánh xạ khóa trong file settings (mục "word_format") → (khóa CONFIG, hàm chuyển kiểu)
_WORD_FORMAT_KEYS = {
    "university_name": ("university_name", str),
    "faculty_name": ("faculty_name", str),
    "country_header": ("country_header", str),
    "country_slogan": ("country_slogan", str),
    "report_title": ("report_title_lines",
                     lambda v: [ln.strip() for ln in str(v).split("\n") if ln.strip()]),
    "font_name": ("font_name", str),
    "font_size": ("font_size", int),
    "margin_top": ("margin_top", float),
    "margin_bottom": ("margin_bottom", float),
    "margin_left": ("margin_left", float),
    "margin_right": ("margin_right", float),
    "table_header_bg": ("table_header_bg", str),
    "open_response_lines": ("open_response_lines", int),
    "section_layouts": ("section_layouts", str),
    "signer_title": ("signer_title", str),
}

# Mục "feedback_config": (khóa CONFIG, ký tự phân tách khi lưu dạng chuỗi)
_FEEDBACK_KEYS = {
    "search_keywords": ("feedback_search_keywords", ","),
    "skip_keywords": ("feedback_skip_keywords", ","),
    "skip_patterns": ("feedback_skip_patterns", "\n"),
}


def _as_list(value, sep):
    if isinstance(value, list):
        return value
    return [v.strip() for v in str(value).split(sep) if v.strip()]


def load_settings_into_config():
    """Nạp lại CONFIG = mặc định + các giá trị người dùng đã lưu trong SETTINGS_FILE."""
    CONFIG.clear()
    CONFIG.update(copy.deepcopy(DEFAULT_CONFIG))
    try:
        if not SETTINGS_FILE.exists():
            return
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        for key, (cfg_key, cast) in _WORD_FORMAT_KEYS.items():
            if key in data.get("word_format", {}):
                CONFIG[cfg_key] = cast(data["word_format"][key])
        for key, (cfg_key, sep) in _FEEDBACK_KEYS.items():
            if key in data.get("feedback_config", {}):
                CONFIG[cfg_key] = _as_list(data["feedback_config"][key], sep)
    except Exception as e:
        log.warning("Không đọc được cấu hình người dùng: %s", e)

# Tải cấu hình tự động khi import
load_settings_into_config()


# ═══════════════════════════════════════════════════════════════════
# TRÍCH XUẤT DỮ LIỆU TỪ EXCEL
# ═══════════════════════════════════════════════════════════════════

def extract_data_from_excel(filepath, sheet_index=0, apply_filename_suffix=True):
    """Đọc và trích xuất toàn bộ dữ liệu từ file Excel khảo sát."""
    wb = _open_workbook(filepath)
    sheet = wb.sheet_by_index(sheet_index)
    
    data = {
        "meta": {},
        "background_info": [],    # Thông tin nền (giới tính, hình thức đào tạo...)
        "survey_questions": [],   # Các câu hỏi khảo sát với dữ liệu
        "open_responses": "",     # Câu hỏi mở
        "signer_name": "",        # Người lập
        "signer_date": "",        # Ngày lập
        "warnings": [],           # Danh sách cảnh báo
    }
    
    # --- Trích xuất tổng số sinh viên từ ô F5 (fallback) ---
    tong_so_sv = None
    if sheet.nrows > 4 and sheet.ncols > 5:
        f5_val = sheet.cell(4, 5).value  # F5 = row index 4, col index 5
        if f5_val is not None and str(f5_val).strip() != "":
            try:
                tong_so_sv = int(float(f5_val))
            except (ValueError, TypeError):
                data["warnings"].append(
                    f"⚠️ Ô F5 có giá trị '{f5_val}' nhưng không phải số. "
                )
    
    # --- Lấy số SV KHÔNG tham gia khảo sát từ tên file ---
    basename = os.path.splitext(os.path.basename(filepath))[0]
    match = re.search(r'\s+(\d+)$', basename)
    sv_khong_khao_sat = None
    if apply_filename_suffix and match:
        sv_khong_khao_sat = int(match.group(1))
    
    # --- Trích xuất metadata từ header ---
    for r in range(min(10, sheet.nrows)):
        row_text = " ".join(str(sheet.cell(r, c).value).strip() for c in range(sheet.ncols))
        
        # Đợt khảo sát
        if "Đợt khảo sát" in row_text or "đợt khảo sát" in row_text.lower():
            data["meta"]["dot_khao_sat"] = str(sheet.cell(r, 0).value).replace("Đợt khảo sát:", "").strip().rstrip(" ")
        
        # Phiếu khảo sát
        for c in range(sheet.ncols):
            val = str(sheet.cell(r, c).value).strip()
            if "Phiếu khảo sát" in val:
                # Phiếu khảo sát name thường ở ô kế tiếp
                for cc in range(c + 1, sheet.ncols):
                    pval = str(sheet.cell(r, cc).value).strip().rstrip(" ")
                    if pval:
                        data["meta"]["phieu_khao_sat"] = pval
                        break
        
        # Thời gian khảo sát
        if "Thời gian KS" in row_text or "Thời gian" in row_text:
            val = str(sheet.cell(r, 0).value).strip()
            match = re.search(r'Từ\s+(.+)', val)
            if match:
                data["meta"]["thoi_gian_ks"] = match.group(1).strip().rstrip(" ")
            else:
                time_part = val.replace("Thời gian KS:", "").replace("Thời gian:", "").strip()
                data["meta"]["thoi_gian_ks"] = time_part.rstrip(" ")
        
        # Lớp học
        if "Lớp học" in row_text:
            for c in range(sheet.ncols):
                if "Lớp học" in str(sheet.cell(r, c).value):
                    for cc in range(c + 1, sheet.ncols):
                        lop = str(sheet.cell(r, cc).value).strip().rstrip(" ")
                        if lop:
                            data["meta"]["lop_hoc"] = lop
                            break
                    break
        
        # Môn học
        if "Môn học" in row_text:
            val = str(sheet.cell(r, 0).value).strip()
            mon = val.replace("Môn học:", "").strip().rstrip(" ")
            if mon:
                data["meta"]["mon_hoc"] = mon
        
        # Giảng viên
        if "Giảng viên" in row_text:
            for c in range(sheet.ncols):
                if "Giảng viên" in str(sheet.cell(r, c).value):
                    for cc in range(c + 1, sheet.ncols):
                        gv = str(sheet.cell(r, cc).value).strip().rstrip(" ")
                        if gv:
                            data["meta"]["giang_vien"] = gv
                            break
                    break

    # --- Parse dữ liệu chính: thông tin nền + câu hỏi khảo sát ---
    section = None  # "thong_tin_chung" hoặc "khao_sat"
    current_question = None
    
    for r in range(sheet.nrows):
        col0 = str(sheet.cell(r, 0).value).strip()
        col1 = str(sheet.cell(r, 1).value).strip()
        
        # Phát hiện section headers
        if col0 == "Thông tin chung" and not col1:
            section = "thong_tin_chung"
            continue
        elif col0 == "Thông tin khảo sát" and not col1:
            section = "khao_sat"
            continue
        
        # Skip header rows, empty rows
        if not col0 and not col1:
            continue
        if col0 == "" and col1 == "" :
            continue
            
        # --- THÔNG TIN NỀN ---
        if section == "thong_tin_chung":
            if col0.startswith("C") and ":" in col0 and col1 == "Số SV trả lời":
                # Extract question label
                label = col0.split(":", 1)[1].strip() if ":" in col0 else col0
                
                # Get values from columns 2 onwards
                values = []
                for c in range(2, sheet.ncols):
                    val = sheet.cell(r, c).value
                    try:
                        values.append(float(val))
                    except (ValueError, TypeError):
                        values.append(0)
                
                # Get percentages from next row
                if r + 1 < sheet.nrows and str(sheet.cell(r + 1, 1).value).strip() == "Tỷ lệ":
                    percentages = []
                    for c in range(2, sheet.ncols):
                        val = sheet.cell(r + 1, c).value
                        try:
                            percentages.append(float(val))
                        except (ValueError, TypeError):
                            percentages.append(0)
                else:
                    percentages = []
                
                data["background_info"].append({
                    "label": label,
                    "counts": values,
                    "percentages": percentages,
                })
            elif col0.startswith("Ngành học"):
                nghanh = str(sheet.cell(r, 2).value).strip().rstrip(" ") if sheet.ncols > 2 else ""
                data["meta"]["nganh_hoc_raw"] = nghanh
        
        # --- CÂU HỎI KHẢO SÁT ---
        elif section == "khao_sat":
            if col0.startswith("C") and ":" in col0 and col1 == "Số SV trả lời":
                # Parse question
                q_match = re.match(r'C\d+:\s*(\d+\.?\d*\.?\s*)(.*)', col0)
                if q_match:
                    q_number = q_match.group(1).strip().rstrip(".")
                    q_text = q_match.group(2).strip()
                else:
                    # Fallback: just use the text after ":"
                    parts = col0.split(":", 1)
                    q_number = parts[0].replace("C", "").strip()
                    q_text = parts[1].strip() if len(parts) > 1 else col0
                
                # Counts: cols 2-6 = mức 1-5, col 7 = tổng
                counts = []
                for c in range(2, min(7, sheet.ncols)):
                    val = sheet.cell(r, c).value
                    try:
                        counts.append(int(float(val)))
                    except (ValueError, TypeError):
                        counts.append(0)
                
                total = 0
                if sheet.ncols > 7:
                    try:
                        total = int(float(sheet.cell(r, 7).value))
                    except (ValueError, TypeError):
                        total = sum(counts)
                else:
                    total = sum(counts)
                
                # Percentages from next row
                percentages = []
                if r + 1 < sheet.nrows and str(sheet.cell(r + 1, 1).value).strip() == "Tỷ lệ":
                    for c in range(2, min(7, sheet.ncols)):
                        val = sheet.cell(r + 1, c).value
                        try:
                            percentages.append(float(val))
                        except (ValueError, TypeError):
                            percentages.append(0)
                
                # Thiếu dòng "Tỷ lệ" → tự tính từ số lượng
                answered = sum(counts)
                if not percentages and total > 0:
                    percentages = [c / total for c in counts]

                # Điểm trung bình trên số SV có chọn mức (bỏ phiếu trống câu này)
                if answered > 0 and len(counts) == 5:
                    avg = sum((i + 1) * counts[i] for i in range(5)) / answered
                else:
                    avg = 0
                
                current_question = {
                    "number": q_number,
                    "text": q_text,
                    "counts": counts,
                    "total": total,
                    "percentages": percentages,
                    "average": round(avg, 2),
                }
                data["survey_questions"].append(current_question)
            
            # Câu hỏi mở (câu 20)
            elif "ý kiến đóng góp" in col0.lower() or "ý kiến" in col0.lower() and "đóng góp" in col0.lower():
                raw = str(sheet.cell(r, 2).value).strip().rstrip(" ") if sheet.ncols > 2 else ""
                data["open_responses"] = raw
    
    # --- Tính toán Tổng số sinh viên (Dựa vào tổng phiếu + SV không tham gia) ---
    total_phieu = 0
    if data["survey_questions"]:
        total_phieu = data["survey_questions"][0].get("total", 0)
        
    if sv_khong_khao_sat is not None:
        data["meta"]["tong_so_sv"] = total_phieu + sv_khong_khao_sat
    elif tong_so_sv is not None:
        data["meta"]["tong_so_sv"] = tong_so_sv
    else:
        data["meta"]["tong_so_sv"] = None
        if apply_filename_suffix:
            data["warnings"].append("⚠️ Không lấy được số SV không tham gia từ tên file và ô F5 cũng trống!")
    
    # --- Trích xuất thông tin cuối file (ngày, người ký) ---
    for r in range(max(0, sheet.nrows - 10), sheet.nrows):
        for c in range(sheet.ncols):
            val = str(sheet.cell(r, c).value).strip()
            if "ngày" in val.lower() and "tháng" in val.lower():
                fixed_val = val.rstrip(" ").replace(" nam ", " năm ")
                data["signer_date"] = fixed_val
            if "Người lập" in val:
                # Tên người lập thường ở vài dòng sau
                for rr in range(r + 1, min(r + 5, sheet.nrows)):
                    name = str(sheet.cell(rr, c).value).strip()
                    if name:
                        data["signer_name"] = name
                        break

    return data


# ═══════════════════════════════════════════════════════════════════
# XỬ LÝ Ý KIẾN ĐÓNG GÓP (TÁCH CÂU TRẢ LỜI)
# ═══════════════════════════════════════════════════════════════════

# Từ xưng hô/danh từ thường đứng ngay trước tên riêng viết hoa ("thầy Nam", "quận Bình Thạnh")
_NAME_PREFIXES = re.compile(
    r"(?:^|\s)(?:thầy|cô|thầy cô|anh|chị|bạn|giảng viên|gv|trường|khoa|bộ môn|môn|"
    r"quận|phường|tp|thành phố|cơ sở|bệnh viện|bv|đại học|ĐH)$", re.IGNORECASE)
# "không"/"dạ" dính liền chữ trước (vd "Khôngkhông có", "ạdạ không") — không có âm tiết
# tiếng Việt hợp lệ nào kết thúc bằng các chuỗi này nên tách an toàn.
_GLUED_REFUSAL = re.compile(r"(?<=[^\W\d_])(?=(?:không|khong|dạ)(?:\s|$|[.,!?]))", re.IGNORECASE)


def _split_glued_refusals(parts):
    out = []
    for p in parts:
        out.extend(_GLUED_REFUSAL.split(p))
    return out


def _rejoin_names(parts):
    """Nối lại các đoạn bị tách nhầm trước tên riêng (vd "Em mong thầy" + "Nam giảng chậm")."""
    out = []
    for p in parts:
        if out and _NAME_PREFIXES.search(out[-1].strip()):
            out[-1] = out[-1].rstrip() + " " + p.lstrip()
        else:
            out.append(p)
    return out


def parse_open_responses(raw_text):
    """Tách các câu trả lời mở từ chuỗi nối liền, loại bỏ câu trả lời trống/vô nghĩa."""
    if not raw_text:
        return []
    
    # Các pattern chỉ "không có ý kiến"
    default_patterns = [
        r'^không\s*(ạ|có|cóa)?\.?$',
        r'^ko\.?$',
        r'^khong\.?$',
        r'^dạ\s*không\s*(ạ|có)?\.?$',
        r'^không\s*có\s*(góp\s*ý|ý\s*kiến)?\.?$',
        r'^dạ\s*không\s*có\.?$',
        r'^không\s*có\s*ý\s*kiến\.?$',
        r'^em\s*không\s*có\s*ý\s*kiến\s*(ạ)?\.?$',
    ]
    skip_patterns = CONFIG.get("feedback_skip_patterns", default_patterns)
    
    # Thử tách bằng nhiều cách
    # Cách 1: Tách bằng các cụm keyword phân cách 
    # (Do Excel gộp tất cả vào 1 ô, cần heuristic để tách)
    
    # Tách bằng các chữ cái viết hoa đầu câu sau chữ thường
    # hoặc tách bằng dấu chấm câu + viết hoa
    responses = []
    
    # Xử lý \r\n
    raw_text = raw_text.replace("\\r\\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    
    # Tách đơn giản: mỗi câu trả lời thường bắt đầu bằng chữ cái viết hoa
    # sau khi kết thúc bằng dấu chấm hoặc không có dấu ngắt
    # Heuristic: tách tại vị trí có chữ cái viết hoa sau chữ cái thường (không phải giữa từ thông thường)
    
    # Trước tiên, tách bằng newline
    parts = raw_text.split("\n")
    
    for part in parts:
        part = part.strip()
        if not part:
            continue
        
        # Tách tiếp bằng pattern: kết thúc câu trước + bắt đầu câu mới
        # Pattern: sau dấu chấm/không dấu + chữ cái viết hoa
        sub_parts = re.split(r'(?<=[a-zàáảãạăắằẳẵặâấầẩẫậèéẻẽẹêếềểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵđ\.\!\?])\s*(?=[A-ZÀÁẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÈÉẺẼẸÊẾỀỂỄỆÌÍỈĨỊÒÓỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÙÚỦŨỤƯỨỪỬỮỰỲÝỶỸỴĐ])', part)
        
        for sp in _rejoin_names(_split_glued_refusals(sub_parts)):
            sp = sp.strip().rstrip(".")
            if not sp or len(sp) < 3:
                continue
            
            # Kiểm tra có phải câu trả lời rỗng/vô nghĩa không
            is_skip = False
            for pattern in skip_patterns:
                if re.match(pattern, sp, re.IGNORECASE):
                    is_skip = True
                    break
            
            if not is_skip:
                # Thêm dấu chấm cuối câu nếu chưa có
                if not sp.endswith(('.', '!', '?')):
                    sp += '.'
                # Viết hoa chữ cái đầu
                sp = sp[0].upper() + sp[1:]
                responses.append(sp)
    
    # Loại bỏ trùng lặp, giữ thứ tự
    seen = set()
    unique = []
    for r in responses:
        r_lower = r.lower().strip()
        if r_lower not in seen:
            seen.add(r_lower)
            unique.append(r)
    
    return unique


# ═══════════════════════════════════════════════════════════════════
# TẠO FILE WORD
# ═══════════════════════════════════════════════════════════════════

def set_cell_shading(cell, color_hex):
    """Đặt màu nền cho ô trong bảng Word."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_border(cell, **kwargs):
    """Đặt border cho ô bảng."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}></w:tcBorders>')
    for edge, val in kwargs.items():
        element = parse_xml(
            f'<w:{edge} {nsdecls("w")} w:val="{val.get("val", "single")}" '
            f'w:sz="{val.get("sz", "4")}" w:space="0" '
            f'w:color="{val.get("color", "000000")}"/>'
        )
        tcBorders.append(element)
    tcPr.append(tcBorders)


def add_formatted_paragraph(doc, text, bold=False, italic=False, alignment=None, 
                            font_size=None, font_name=None, space_after=None, space_before=None):
    """Thêm paragraph với định dạng."""
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = font_name or CONFIG["font_name"]
    run.font.size = Pt(font_size or CONFIG["font_size"])
    
    # Đảm bảo font cho cả ký tự CJK/Vietnamese
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name or CONFIG["font_name"])
    
    if alignment is not None:
        para.alignment = alignment
    if space_after is not None:
        para.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        para.paragraph_format.space_before = Pt(space_before)
    
    return para


def format_cell_text(cell, text, bold=False, italic=False, font_size=None, alignment=None):
    """Format text trong ô bảng."""
    cell.text = ""
    para = cell.paragraphs[0]
    run = para.add_run(str(text))
    run.bold = bold
    run.font.name = CONFIG["font_name"]
    run.font.size = Pt(font_size or CONFIG["font_size"])
    run._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    if italic:
        run.italic = True
    if alignment is not None:
        para.alignment = alignment


def parse_section_layouts(text):
    """"Tên | 1-7" (mỗi dòng), các mẫu phiếu cách nhau bằng dòng trống → list[list[(tên, từ, đến)]]."""
    layouts, current = [], []
    for line in str(text).splitlines() + [""]:
        line = line.strip()
        if not line:
            if current:
                layouts.append(current)
                current = []
            continue
        name, _, rng = line.rpartition("|")
        m = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+))?\s*", rng)
        if not name.strip() or not m:
            raise ValueError(f"Dòng nhóm tiêu chí không hợp lệ: “{line}” (đúng dạng: Tên nhóm | 1-7)")
        lo = int(m.group(1))
        current.append((name.strip(), lo, int(m.group(2) or lo)))
    return layouts


def _major_number(q):
    m = re.match(r"(\d+)", str(q.get("number", "")))
    return int(m.group(1)) if m else None


def build_sections(questions):
    """Chia câu hỏi (đã sắp theo thứ tự trong phiếu) thành nhóm theo CONFIG["section_layouts"].
    Trả về list[{"name", "start", "end"}] (chỉ số trong questions) hoặc None nếu không mẫu nào khớp."""
    majors = [_major_number(q) for q in questions]
    if not questions or None in majors:
        return None
    try:
        layouts = parse_section_layouts(CONFIG.get("section_layouts", ""))
    except ValueError as e:
        log.warning("%s", e)
        return None
    for layout in layouts:
        if layout[-1][2] != max(majors):
            continue
        sections = []
        for name, lo, hi in layout:
            idx = [i for i, n in enumerate(majors) if lo <= n <= hi]
            if idx:
                sections.append({"name": name, "start": idx[0], "end": idx[-1] + 1})
        covered = sum(s["end"] - s["start"] for s in sections)
        if covered == len(questions):
            return sections
    return None


def create_word_report(data, output_path):
    """Tạo file Word báo cáo từ dữ liệu đã trích xuất."""
    doc = Document()
    
    # --- Cấu hình trang ---
    section = doc.sections[0]
    section.page_width = Cm(21)     # A4
    section.page_height = Cm(29.7)
    section.left_margin = Cm(CONFIG.get("margin_left", 2.5))
    section.right_margin = Cm(CONFIG.get("margin_right", 2.0))
    section.top_margin = Cm(CONFIG.get("margin_top", 2.0))
    section.bottom_margin = Cm(CONFIG.get("margin_bottom", 2.0))
    
    # --- Set default font + giãn dòng ---
    style = doc.styles['Normal']
    font = style.font
    font.name = CONFIG["font_name"]
    font.size = Pt(CONFIG["font_size"])
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    
    # ===== HEADER: Bảng tiêu đề =====
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Ép chế độ fixed layout + chiều rộng cột cố định
    tbl = header_table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(tblLayout)
    
    # Cột trái 6cm, cột phải 10.5cm
    for ci, w_cm in enumerate([6, 10.5]):
        cell = header_table.cell(0, ci)
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(w_cm * 567)}" w:type="dxa"/>')
        tcPr.append(tcW)
    
    # Cột trái: Tên trường + khoa
    left_cell = header_table.cell(0, 0)
    format_cell_text(left_cell, CONFIG['university_name'], bold=False, font_size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    p_faculty = left_cell.add_paragraph()
    p_faculty.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_fac = p_faculty.add_run(CONFIG['faculty_name'])
    run_fac.bold = True
    run_fac.font.name = CONFIG["font_name"]
    run_fac.font.size = Pt(12)
    run_fac._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    
    # Cột phải: Quốc hiệu
    right_cell = header_table.cell(0, 1)
    right_cell.text = ""
    p = right_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run1 = p.add_run(CONFIG["country_header"])
    run1.bold = True
    run1.font.name = CONFIG["font_name"]
    run1.font.size = Pt(12)
    run1._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    
    p2 = right_cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(CONFIG["country_slogan"])
    run2.italic = True
    run2.font.name = CONFIG["font_name"]
    run2.font.size = Pt(12)
    run2._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    
    # Ngày tháng năm
    date_text = data["signer_date"] if data["signer_date"] else "Thành phố Hồ Chí Minh, ngày ... tháng ... năm ..."
    p3 = right_cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run3 = p3.add_run(date_text)
    run3.font.name = CONFIG["font_name"]
    run3.font.size = Pt(12)
    run3._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    
    # Xóa border header table
    for row in header_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                pass  # Keep as is
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)
    
    # Trích xuất metadata trước để dùng cho tiêu đề và text box mã lớp
    meta = data.get("meta", {})
    mon_hoc = meta.get("mon_hoc", "")
    mon_hoc_match = re.match(r'(\d+)\s*-\s*(.*)', mon_hoc)
    if mon_hoc_match:
        ma_mon = mon_hoc_match.group(1).strip()
        ten_mon = mon_hoc_match.group(2).strip()
        mon_display = f"{ten_mon} ({ma_mon})"
    else:
        mon_display = mon_hoc
        ten_mon = mon_hoc

    if not ten_mon and meta.get("dot_khao_sat"):
        parts = meta.get("dot_khao_sat").split()
        if len(parts) >= 2:
            sub_parts = parts[1:]
            if sub_parts and re.match(r'\d+HK\w+', sub_parts[-1]):
                sub_parts = sub_parts[:-1]
            ten_mon = " ".join(sub_parts)

    lop_hoc = meta.get("lop_hoc", "").strip()
    if not lop_hoc and meta.get("dot_khao_sat"):
        lop_hoc = meta.get("dot_khao_sat").split()[0]

    # Spacer + Text Box mã viết tắt của lớp ở góc trên bên trái đầu trang giấy (theo mẫu mới - Hình 1)
    if lop_hoc:
        p_box = parse_xml(XML_CLASS_BOX_TEMPLATE.format(CLASS_CODE=xml_escape(lop_hoc)))
        header_table._tbl.addnext(p_box)
    else:
        doc.add_paragraph()  # Spacer

    # ===== TIÊU ĐỀ BÁO CÁO =====
    for line in CONFIG["report_title_lines"]:
        title_line = line
        if ten_mon:
            if "HỌC PHẦN CƠ SỞ" in title_line.upper():
                title_line = re.sub(r'HỌC PHẦN CƠ SỞ', f'HỌC PHẦN {ten_mon.upper()}', title_line, flags=re.IGNORECASE)
            elif "{mon_hoc}" in title_line or "{ten_mon}" in title_line:
                title_line = title_line.replace("{mon_hoc}", ten_mon.upper()).replace("{ten_mon}", ten_mon.upper())
            elif "KHẢO SÁT SINH VIÊN SAU KHI KẾT THÚC HỌC PHẦN" in title_line:
                title_line = f"KHẢO SÁT SINH VIÊN SAU KHI KẾT THÚC HỌC PHẦN {ten_mon.upper()}"
        add_formatted_paragraph(doc, title_line, bold=True, font_size=12,
                                alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    
    doc.add_paragraph()  # Spacer
    
    # ===== PHẦN 1: THÔNG TIN CHUNG =====
    add_formatted_paragraph(doc, "Phần 1. Thông tin chung", bold=True, font_size=12, space_after=6)
    
    thoi_gian = meta.get("thoi_gian_ks", "")
    phieu_ks = meta.get("phieu_khao_sat", "")
    
    # Tổng SV phản hồi (từ dữ liệu khảo sát)
    total_phieu = 0
    if data["survey_questions"]:
        total_phieu = data["survey_questions"][0].get("total", 0)
    
    # Tổng SV đăng ký (từ ô F5)
    tong_so_sv = meta.get("tong_so_sv", None)
    
    # Tính tỷ lệ khảo sát
    if tong_so_sv and tong_so_sv > 0:
        ty_le = total_phieu / tong_so_sv * 100
        ty_le_str = f"{ty_le:.1f}%"
        tong_sv_display = str(tong_so_sv)
    else:
        ty_le_str = f"{total_phieu}/{total_phieu} (100%)" if total_phieu > 0 else "N/A"
        tong_sv_display = str(total_phieu)  # Fallback: dùng số phiếu khảo sát
    
    info_lines = [
        ("Module/học phần khảo sát: ", mon_display, True, False),
        ("Thời gian khảo sát: ", thoi_gian, False, False),
        ("Đối tượng khảo sát: ", lop_hoc, True, False),
        ("Mẫu khảo sát: ", phieu_ks, False, False),
        (f"Tổng số người học môn học/học phần được khảo sát: ", tong_sv_display, False, True),
        (f"Tổng số người học được lấy ý kiến: ", str(total_phieu), False, True),
        (f"Tỷ lệ người học được lấy ý kiến: ", ty_le_str, False, True),
    ]
    
    for label, value, value_bold, has_bullet in info_lines:
        para = doc.add_paragraph()
        if has_bullet:
            para.paragraph_format.left_indent = Cm(1.27)
            run_dash = para.add_run("- ")
            run_dash.font.name = CONFIG["font_name"]
            run_dash.font.size = Pt(CONFIG["font_size"])
            run_dash._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
        
        run_label = para.add_run(label)
        run_label.font.name = CONFIG["font_name"]
        run_label.font.size = Pt(CONFIG["font_size"])
        run_label._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
        
        run_val = para.add_run(value)
        run_val.bold = value_bold
        run_val.font.name = CONFIG["font_name"]
        run_val.font.size = Pt(CONFIG["font_size"])
        run_val._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
        
        para.paragraph_format.space_after = Pt(2)
    

    # ===== PHẦN 2: KẾT QUẢ KHẢO SÁT =====
    add_formatted_paragraph(doc, "Phần 2. Kết quả khảo sát", bold=True, font_size=12, 
                            space_before=12, space_after=6)
    
    # --- 2.1 Thông tin nền ---
    add_formatted_paragraph(doc, "2.1. Thông tin nền", bold=True, font_size=12, space_after=6)
    
    bg = data["background_info"]
    if bg:
        # Tạo bảng thông tin nền
        bg_table = doc.add_table(rows=1, cols=4)
        bg_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        bg_table.style = 'Table Grid'
        
        # Header
        headers = ["", "", "Số lượng", "Tỷ lệ (%)"]
        for i, h in enumerate(headers):
            format_cell_text(bg_table.cell(0, i), h, bold=True, font_size=11,
                            alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(bg_table.cell(0, i), CONFIG.get("table_header_bg", "D9E2F3"))
        
        # Mapping thông tin nền
        bg_mapping = {
            "Hình thức đào tạo": {"labels": ["Chính quy", "Khác"], "col_indices": [0, 1]},
            "Giới tính": {"labels": ["Nam", "Nữ"], "col_indices": [0, 1]},
            "Dân tộc thiểu số": {"labels": ["Có", "Không"], "col_indices": [0, 1]},
            "Sinh viên nước ngoài": {"labels": ["Có", "Không"], "col_indices": [0, 1]},
        }
        
        for bg_item in bg:
            label = bg_item["label"]
            counts = bg_item["counts"]
            pcts = bg_item["percentages"]
            
            # Tìm mapping phù hợp
            matched_key = None
            for key in bg_mapping:
                if key.lower() in label.lower():
                    matched_key = key
                    break
            
            if matched_key:
                mapping = bg_mapping[matched_key]
                for j, sub_label in enumerate(mapping["labels"]):
                    row = bg_table.add_row()
                    
                    if j == 0:
                        format_cell_text(row.cells[0], matched_key, font_size=11)
                    else:
                        format_cell_text(row.cells[0], "", font_size=11)
                    
                    format_cell_text(row.cells[1], sub_label, font_size=11)
                    
                    idx = mapping["col_indices"][j]
                    count_val = int(counts[idx]) if idx < len(counts) else 0
                    pct_val = pcts[idx] * 100 if idx < len(pcts) else 0
                    
                    format_cell_text(row.cells[2], str(count_val), font_size=11,
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER)
                    format_cell_text(row.cells[3], f"{pct_val:.2f}%", font_size=11,
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER)
        
        # Merge ô trùng lặp ở cột đầu tiên
        # (Thực hiện merge vertical cho các ô cùng nhóm)
        row_idx = 1  # Bỏ qua header
        for bg_item in bg:
            label = bg_item["label"]
            matched_key = None
            for key in bg_mapping:
                if key.lower() in label.lower():
                    matched_key = key
                    break
            if matched_key:
                n_rows = len(bg_mapping[matched_key]["labels"])
                if n_rows > 1 and row_idx + n_rows - 1 < len(bg_table.rows):
                    # Merge cells
                    a = bg_table.cell(row_idx, 0)
                    b = bg_table.cell(row_idx + n_rows - 1, 0)
                    a.merge(b)
                    # Center text vertically
                    for para in a.paragraphs:
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                row_idx += n_rows
        
        # Giảm giãn dòng bảng thông tin nền về 1.0
        for row in bg_table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.line_spacing = 1.0
    
    doc.add_paragraph()  # Spacer
    
    # --- 2.2 Thông tin cụ thể ---
    add_formatted_paragraph(doc, "2.2. Thông tin cụ thể", bold=True, font_size=12, space_after=2)
    
    para_note = doc.add_paragraph()
    para_note.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_note = para_note.add_run("Báo cáo tỉ lệ phần trăm phản hồi của sinh viên ở từng tiêu chí")
    run_note.italic = True
    run_note.font.name = CONFIG["font_name"]
    run_note.font.size = Pt(CONFIG["font_size"])
    run_note._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    
    # Tạo bảng kết quả khảo sát
    questions = data["survey_questions"]
    if questions:
        # Tính số dòng = header + section headers + questions
        # Phân chia thành sections dựa trên số lượng câu hỏi
        n_q = len(questions)
        
        sections = build_sections(questions)
        if sections is None:
            data.setdefault("warnings", []).append(
                "Không có nhóm tiêu chí nào khớp với số câu của phiếu — gộp tất cả vào một nhóm "
                "(chỉnh tại Cấu hình → Nhóm tiêu chí)")
            sections = [{"name": "NỘI DUNG KHẢO SÁT", "start": 0, "end": n_q}]

        total_rows = 1 + len(sections) + n_q  # header + section headers + questions
        survey_table = doc.add_table(rows=total_rows, cols=8)
        survey_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        survey_table.style = 'Table Grid'
        
        # Set column widths
        for row in survey_table.rows:
            row.cells[0].width = Cm(1)      # TT
            row.cells[1].width = Cm(9)      # Nội dung (rộng hơn 30%)
            row.cells[2].width = Cm(1.3)    # Mức 1
            row.cells[3].width = Cm(1.3)    # Mức 2
            row.cells[4].width = Cm(1.3)    # Mức 3
            row.cells[5].width = Cm(1.3)    # Mức 4
            row.cells[6].width = Cm(1.3)    # Mức 5
            row.cells[7].width = Cm(1.5)    # Điểm TB
        
        # Header row
        header_row = survey_table.rows[0]
        header_texts = ["TT", "Nội dung đánh giá"] + CONFIG["scale_labels"] + ["Điểm trung bình"]
        for i, h in enumerate(header_texts):
            format_cell_text(header_row.cells[i], h, bold=True, font_size=10,
                            alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(header_row.cells[i], CONFIG.get("table_header_bg", "D9E2F3"))
        
        # Fill data
        current_row = 1
        question_counter = 1
        
        for sec in sections:
            # Section header row
            sec_row = survey_table.rows[current_row]
            
            # Merge all columns for section header
            a = sec_row.cells[0]
            b = sec_row.cells[7]
            a.merge(b)
            format_cell_text(sec_row.cells[0], sec["name"], bold=True, italic=True, 
                            font_size=11)
            set_cell_shading(sec_row.cells[0], "F2F2F2")
            current_row += 1
            
            # Question rows
            for qi in range(sec["start"], sec["end"]):
                if qi >= len(questions):
                    break
                q = questions[qi]
                q_row = survey_table.rows[current_row]
                
                # TT
                format_cell_text(q_row.cells[0], str(question_counter), font_size=10,
                                alignment=WD_ALIGN_PARAGRAPH.CENTER)
                
                # Nội dung - Thêm mã câu hỏi
                q_label = f"C{q['number']}. " if q['number'] else ""
                format_cell_text(q_row.cells[1], f"{q_label}{q['text']}", font_size=10)
                
                # Tỷ lệ % cho 5 mức (làm tròn không thập phân)
                for j in range(5):
                    if j < len(q["percentages"]):
                        pct = round(q["percentages"][j] * 100)
                        format_cell_text(q_row.cells[2 + j], f"{pct}%", font_size=10,
                                        alignment=WD_ALIGN_PARAGRAPH.CENTER)
                    else:
                        format_cell_text(q_row.cells[2 + j], "0%", font_size=10,
                                        alignment=WD_ALIGN_PARAGRAPH.CENTER)
                
                # Điểm trung bình
                format_cell_text(q_row.cells[7], f"{q['average']:.2f}", font_size=10,
                                alignment=WD_ALIGN_PARAGRAPH.CENTER)
                
                question_counter += 1
                current_row += 1
    
        # Giảm giãn dòng trong bảng khảo sát về 1.0
        for row in survey_table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.line_spacing = 1.0
    
    # Spacer trước Phần 3: Thêm dấu * theo mẫu mới (Hình 3)
    p_star = doc.add_paragraph()
    run_star = p_star.add_run("*")
    run_star.font.name = CONFIG["font_name"]
    run_star.font.size = Pt(CONFIG["font_size"])
    run_star._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    p_star.paragraph_format.space_before = Pt(0)
    p_star.paragraph_format.space_after = Pt(0)
    p_star.paragraph_format.line_spacing = 1.5
    
    # ===== PHẦN 3: Ý KIẾN ĐÓNG GÓP =====
    add_formatted_paragraph(doc, "Phần 3. Các ý kiến đóng góp từ câu hỏi mở:", 
                            bold=True, font_size=12, space_after=6)
    
    # Liệt kê ý kiến thực tế; nếu không có thì ghi "*Không có ý kiến." theo mẫu mới (Hình 3)
    responses = parse_open_responses(data.get("open_responses", ""))
    for text in responses or ["*Không có ý kiến."]:
        p_op = add_formatted_paragraph(doc, f"- {text}" if responses else text,
                                       space_after=0, space_before=0)
        if responses:
            p_op.paragraph_format.left_indent = Cm(0.63)
    
    # Các dòng trống tiếp theo để dự phòng cho người dùng viết thêm ghi chú
    for _ in range(max(0, CONFIG.get("open_response_lines", 4) - max(1, len(responses)))):
        doc.add_paragraph()
    
    
    # ===== PHẦN 4: ĐỀ XUẤT KẾ HOẠCH CẢI TIẾN =====
    # Số tiêu chí = tổng số câu trừ nhóm cuối (CẢM NHẬN CHUNG)
    n_criteria = len(questions)
    if questions and len(sections) > 1:
        n_criteria -= sections[-1]["end"] - sections[-1]["start"]
    add_formatted_paragraph(doc, "Phần 4. Đề xuất kế hoạch cải tiến:", 
                            bold=True, font_size=12, space_after=2)
    
    para_note2 = doc.add_paragraph()
    run_n2 = para_note2.add_run(
        "(Bộ môn, đơn vị đào tạo đề xuất kế hoạch cải tiến cho tối thiểu 2 tiêu chí "
        f"trong {n_criteria} tiêu chí khảo sát ở trên, không tính phần CẢM NHẬN CHUNG)"
    )
    run_n2.font.name = CONFIG["font_name"]
    run_n2.font.size = Pt(CONFIG["font_size"])
    run_n2._element.rPr.rFonts.set(qn('w:eastAsia'), CONFIG["font_name"])
    
    # Bảng cải tiến (để trống cho người dùng điền, 2 hàng theo mẫu mới)
    improve_table = doc.add_table(rows=3, cols=4)
    improve_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    improve_table.style = 'Table Grid'
    
    improve_headers = [
        "Tiêu chí\nmuốn cải tiến",
        "Lý do sinh viên chưa hài lòng",
        "Mục tiêu cải tiến cho năm sau",
        "Giải pháp đề xuất",
    ]
    for i, h in enumerate(improve_headers):
        format_cell_text(improve_table.cell(0, i), h, bold=True, font_size=11,
                        alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(improve_table.cell(0, i), CONFIG.get("table_header_bg", "D9E2F3"))
    
    # Để trống bảng cải tiến cho người dùng tự điền
    
    # Set minimum row height for improvement table
    for row in improve_table.rows[1:]:
        row.height = Cm(2.5)
    
    doc.add_paragraph()  # Spacer
    doc.add_paragraph()  # Spacer
    
    # ===== FOOTER: Ký tên =====
    footer_table = doc.add_table(rows=3, cols=2)
    footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Xóa border
    for row in footer_table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)
    
    # Row 0: TRƯỞNG MODULE/TRƯỞNG BỘ MÔN (bên phải)
    format_cell_text(footer_table.cell(0, 0), "", font_size=12)
    format_cell_text(footer_table.cell(0, 1), CONFIG.get("signer_title", ""), bold=True, font_size=12,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Row 1-2: Dòng trống cho chữ ký
    for ri in range(1, 3):
        format_cell_text(footer_table.cell(ri, 0), "", font_size=12)
        format_cell_text(footer_table.cell(ri, 1), "", font_size=12)
    
    # --- Lưu file ---
    doc.save(output_path)
    return output_path


# ═══════════════════════════════════════════════════════════════════
# TỔNG HỢP NHIỀU SHEET
# ═══════════════════════════════════════════════════════════════════

def merge_sheets_data(all_data_list):
    """Gộp dữ liệu từ nhiều sheet thành 1 bộ dữ liệu tổng hợp."""
    if not all_data_list:
        return None
    if len(all_data_list) == 1:
        return all_data_list[0]
    
    base = all_data_list[0]
    merged = {
        "meta": dict(base["meta"]),
        "background_info": [],
        "survey_questions": [],
        "open_responses": "",
        "signer_name": base.get("signer_name", ""),
        "signer_date": base.get("signer_date", ""),
        "sheet_names": [],  # Danh sách tên lớp/sheet
        "warnings": [],     # Gộp cảnh báo
    }
    
    # Thu thập tên các lớp
    for d in all_data_list:
        lop = d["meta"].get("lop_hoc", d["meta"].get("dot_khao_sat", ""))
        merged["sheet_names"].append(lop)
    
    merged["meta"]["lop_hoc"] = ", ".join(dict.fromkeys(filter(None, merged["sheet_names"])))
    
    # Gộp tổng số SV (F5) từ tất cả sheet
    total_tong_sv = 0
    has_any_tong_sv = False
    for d in all_data_list:
        sv = d["meta"].get("tong_so_sv", None)
        if sv is not None:
            total_tong_sv += sv
            has_any_tong_sv = True
    
    if has_any_tong_sv:
        merged["meta"]["tong_so_sv"] = total_tong_sv
    else:
        merged["warnings"].append(
            "⚠️ Không có sheet nào có ô F5 (Tổng số sinh viên). "
            "Không thể tính tỷ lệ khảo sát tổng hợp."
        )
    
    # === Gộp thông tin nền: cộng dồn số lượng ===
    n_bg = len(base["background_info"])
    for i in range(n_bg):
        item = {
            "label": base["background_info"][i]["label"],
            "counts": list(base["background_info"][i]["counts"]),
            "percentages": [],
        }
        for d in all_data_list[1:]:
            if i < len(d["background_info"]):
                for j in range(len(item["counts"])):
                    if j < len(d["background_info"][i]["counts"]):
                        item["counts"][j] += d["background_info"][i]["counts"][j]
        # Tính lại tổng (cột cuối) và tỷ lệ
        # counts thường có format: [val1, val2, ..., empty, empty, total]
        # Với thông tin nền 2 cột: [chính_quy, khác, 0, 0, 0, tổng]
        # Lấy 2 giá trị đầu làm dữ liệu, cột cuối là tổng
        if len(item["counts"]) >= 2:
            # Recalculate total = sum of first N-1 non-zero position values
            # For bg info: chỉ có 2 giá trị đầu là có nghĩa
            sub_total = item["counts"][0] + item["counts"][1]
            item["counts"][-1] = sub_total  # Cập nhật tổng
        total = item["counts"][-1] if item["counts"] else 1
        if total > 0:
            item["percentages"] = [c / total for c in item["counts"]]
        else:
            item["percentages"] = [0] * len(item["counts"])
        merged["background_info"].append(item)
    
    # === Gộp câu hỏi: cộng dồn counts, tính lại % và TB ===
    n_q = len(base["survey_questions"])
    if any(len(d["survey_questions"]) != n_q for d in all_data_list):
        merged["warnings"].append("Các sheet có số câu hỏi khác nhau — chỉ gộp theo số câu của sheet đầu tiên")
    for i in range(n_q):
        bq = base["survey_questions"][i]
        q = {
            "number": bq["number"],
            "text": bq["text"],
            "counts": list(bq["counts"]),
            "total": bq["total"],
            "percentages": [],
            "average": 0,
        }
        for d in all_data_list[1:]:
            if i < len(d["survey_questions"]):
                dq = d["survey_questions"][i]
                if dq["number"] != bq["number"]:
                    merged["warnings"].append(
                        f"Câu thứ {i + 1} không khớp giữa các sheet (C{bq['number']} ≠ C{dq['number']}) "
                        "— các sheet có thể khác mẫu phiếu")
                for j in range(len(q["counts"])):
                    if j < len(dq["counts"]):
                        q["counts"][j] += dq["counts"][j]
                q["total"] += dq["total"]
        
        # Tính lại tỷ lệ và điểm TB
        if q["total"] > 0:
            q["percentages"] = [c / q["total"] for c in q["counts"]]
            answered = sum(q["counts"])
            if len(q["counts"]) == 5 and answered > 0:
                q["average"] = round(sum((j + 1) * q["counts"][j] for j in range(5)) / answered, 2)
        else:
            q["percentages"] = [0] * len(q["counts"])
        
        merged["survey_questions"].append(q)
    
    # === Gộp ý kiến mở ===
    all_responses = []
    for d in all_data_list:
        if d["open_responses"]:
            all_responses.append(d["open_responses"])
    merged["open_responses"] = "\n".join(all_responses)
    
    return merged


def get_sheet_count(filepath):
    """Đếm số sheet trong file Excel."""
    wb = _open_workbook(filepath)
    return wb.nsheets, wb.sheet_names()


@dataclass
class ReportResult:
    """Kết quả tạo báo cáo cho 1 file Excel (dùng chung cho CLI và GUI)."""
    input_path: str
    output_path: Optional[str] = None
    n_sheets: int = 0
    n_questions: int = 0
    n_responses: int = 0
    total_students: Optional[int] = None
    error: Optional[str] = None
    missing_total: bool = False
    skipped_sheets: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.output_path is not None

    @property
    def merged(self) -> bool:
        return self.n_sheets > 1


def _absent_count_from_name(basename):
    """Số SV KHÔNG tham gia khảo sát ghi ở cuối tên file (vd: "... HKI 3" → 3)."""
    match = re.search(r'\s+(\d+)$', basename)
    return int(match.group(1)) if match else None


def _resolve_total_students(data, all_data, absent):
    """Tổng SV = số phiếu + SV không tham gia (từ tên file); nếu không có thì
    dùng ô F5 (cộng dồn qua các sheet khi gộp). Trả về None nếu không xác định được."""
    if absent is not None:
        total_phieu = data["survey_questions"][0].get("total", 0) if data["survey_questions"] else 0
        return total_phieu + absent
    f5_values = [d["meta"].get("tong_so_sv") for d in all_data]
    f5_values = [v for v in f5_values if v is not None]
    return sum(f5_values) if f5_values else None


def generate_report(input_path, output_dir=None, on_progress=None):
    """Đọc 1 file Excel (gộp tất cả sheet có dữ liệu) và xuất báo cáo Word.

    on_progress(msg): callback tùy chọn nhận thông báo từng sheet bị bỏ qua.
    Không ném lỗi — mọi lỗi được ghi vào ReportResult.error.
    """
    result = ReportResult(input_path=input_path)
    notify = on_progress or (lambda msg: None)
    if not os.path.exists(input_path):
        result.error = "File không tồn tại"
        return result

    basename = os.path.splitext(os.path.basename(input_path))[0]
    clean_basename = re.sub(r'\s+\d+$', '', basename)
    out_dir = output_dir or os.path.dirname(os.path.abspath(input_path))

    try:
        n_sheets, sheet_names = get_sheet_count(input_path)
        all_data = []
        for si in range(n_sheets):
            try:
                data = extract_data_from_excel(input_path, sheet_index=si, apply_filename_suffix=False)
            except Exception as e:
                notify(f'Sheet {si + 1} "{sheet_names[si]}": lỗi đọc – {e}')
                result.skipped_sheets.append(sheet_names[si])
                continue
            if data["survey_questions"]:
                all_data.append(data)
            else:
                notify(f'Sheet {si + 1} "{sheet_names[si]}": bỏ qua (không có dữ liệu)')
                result.skipped_sheets.append(sheet_names[si])

        if not all_data:
            result.error = "Không có sheet nào chứa dữ liệu khảo sát"
            return result

        result.n_sheets = len(all_data)
        if len(all_data) == 1:
            data, out_name = all_data[0], clean_basename
        else:
            data, out_name = merge_sheets_data(all_data), f"{clean_basename} (gộp)"
            if not data:
                result.error = "Không gộp được dữ liệu các sheet"
                return result

        total = _resolve_total_students(data, all_data, _absent_count_from_name(basename))
        data["meta"]["tong_so_sv"] = total
        if total is None:
            result.missing_total = True
            result.error = ("Không xác định được Tổng số SV — hãy ghi số SV không tham gia "
                            "ở cuối tên file hoặc điền vào ô F5")
            return result

        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f"{out_name}.docx")
        try:
            create_word_report(data, output_path)
        except PermissionError:
            result.error = (f"Không ghi được “{os.path.basename(output_path)}” — "
                            "file đang mở trong Word? Hãy đóng file rồi chạy lại")
            return result
        result.warnings = [w for w in data.get("warnings", []) if "F5" not in w]

        result.output_path = output_path
        result.total_students = total
        result.n_questions = len(data["survey_questions"])
        result.n_responses = len(parse_open_responses(data.get("open_responses", "")))
        log.info("Đã tạo %s từ %s", os.path.basename(output_path), os.path.basename(input_path))
    except Exception as e:
        log.exception("Lỗi khi xử lý file %s", input_path)
        result.error = str(e)
    return result


def process_file(input_path, output_dir=None):
    """CLI: xử lý 1 file Excel → Word, in tiến trình ra console.
    Trả về danh sách file Word đã tạo, hoặc None nếu thất bại."""
    print(f"  📄 Đang xử lý: {os.path.basename(input_path)}")
    result = generate_report(input_path, output_dir,
                             on_progress=lambda msg: print(f"     ├─ ⏭ {msg}"))
    if not result.ok:
        print(f"     └─ ❌ {result.error}")
        return None
    prefix = f"📊 ĐÃ GỘP {result.n_sheets} sheets, " if result.merged else ""
    print(f"     ├─ {prefix}{result.n_questions} câu, {result.n_responses} ý kiến, "
          f"{result.total_students} SV")
    for w in result.warnings:
        print(f"     ├─ ⚠️ {w}")
    print(f"     └─ ✅ {os.path.basename(result.output_path)}")
    return [result.output_path]


def main():
    parser = argparse.ArgumentParser(
        description="Tổng hợp kết quả khảo sát từ Excel (.xls) sang Word (.docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ví dụ sử dụng:
  python3 excel_to_word.py                            # Xử lý tất cả .xls trong thư mục hiện tại
  python3 excel_to_word.py "file1.xls" "file2.xls"    # Xử lý các file cụ thể  
  python3 excel_to_word.py --input-dir ./data          # Xử lý tất cả .xls trong thư mục ./data
  python3 excel_to_word.py --output-dir ./output       # Lưu kết quả vào ./output
        """
    )
    parser.add_argument("files", nargs="*", help="Danh sách file Excel (.xls) cần xử lý")
    parser.add_argument("--input-dir", "-i", help="Thư mục chứa file Excel đầu vào")
    parser.add_argument("--output-dir", "-o", help="Thư mục lưu file Word đầu ra")
    
    args = parser.parse_args()
    
    # Thu thập danh sách file cần xử lý
    files_to_process = []
    
    if args.files:
        files_to_process = [os.path.abspath(f) for f in args.files]
    elif args.input_dir:
        pattern = os.path.join(args.input_dir, "*.xls")
        files_to_process = sorted(glob.glob(pattern))
        # Cũng tìm .xlsx
        pattern2 = os.path.join(args.input_dir, "*.xlsx")
        files_to_process += sorted(glob.glob(pattern2))
    else:
        # Mặc định: tìm tất cả .xls trong thư mục hiện tại
        files_to_process = sorted(glob.glob("*.xls"))
        files_to_process += sorted(glob.glob("*.xlsx"))
    
    # Loại bỏ file output (*.docx) khỏi danh sách nếu có
    files_to_process = [f for f in files_to_process if f.endswith(('.xls', '.xlsx'))]
    
    if not files_to_process:
        print("⚠️  Không tìm thấy file Excel nào để xử lý.")
        print("   Sử dụng: python3 excel_to_word.py --help để xem hướng dẫn")
        return
    
    print(f"{'═' * 60}")
    print(f"  📊 TỔNG HỢP KHẢO SÁT EXCEL → WORD")
    print(f"  Số file cần xử lý: {len(files_to_process)}")
    print(f"{'═' * 60}")
    
    success = 0
    failed = 0
    
    for filepath in files_to_process:
        print()
        result = process_file(filepath, args.output_dir)
        if result:
            success += 1
        else:
            failed += 1
    
    print(f"\n{'═' * 60}")
    print(f"  📋 KẾT QUẢ: {success} thành công, {failed} thất bại / {len(files_to_process)} file")
    print(f"{'═' * 60}")


if __name__ == "__main__":
    main()
