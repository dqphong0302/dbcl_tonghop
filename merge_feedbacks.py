#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os
import argparse
import xlrd
from docx import Document
from docx.shared import Pt, Cm
from excel_to_word import parse_open_responses, CONFIG

def extract_feedbacks(filepath):
    feedbacks = []
    try:
        wb = xlrd.open_workbook(filepath)
        for sheet in wb.sheets():
            if sheet.nrows == 0:
                continue
            
            # Tìm cột chứa feedback bằng cách duyệt vài dòng đầu
            feedback_col_idx = -1
            start_row = 1
            
            for r in range(min(10, sheet.nrows)):
                row_vals = sheet.row_values(r)
                for c, val in enumerate(row_vals):
                    val_str = str(val).lower()
                    if "ý kiến đóng góp" in val_str or "feedback" in val_str:
                        feedback_col_idx = c
                        start_row = r + 1
                        break
                if feedback_col_idx != -1:
                    break
                    
            # Nếu không tìm thấy bằng keyword, lấy cột cuối cùng
            if feedback_col_idx == -1:
                feedback_col_idx = sheet.ncols - 1
                start_row = 5 # Mặc định dòng bắt đầu dữ liệu của file form
                
            for r in range(start_row, sheet.nrows):
                val = str(sheet.cell(r, feedback_col_idx).value).strip()
                # Loại bỏ các câu trả lời rỗng hoặc không có ý nghĩa
                if val and val.lower() not in ["không", "không có", "khong", "none", "n/a", "chưa có", "không có ý kiến", "không có ý kiến gì"]:
                    # Tách các câu trả lời và làm sạch bằng hàm có sẵn
                    cleaned_responses = parse_open_responses(val)
                    feedbacks.extend(cleaned_responses)
    except Exception as e:
        print(f"  ❌ Lỗi đọc file {filepath}: {e}")
    return feedbacks

def merge_feedbacks(input_files, output_dir=None):
    if not input_files:
        return None
    
    out_dir = output_dir or os.path.dirname(input_files[0])
    os.makedirs(out_dir, exist_ok=True)
    output_path = os.path.join(out_dir, "Feedbacks.docx")
    
    all_feedbacks = []
    for filepath in input_files:
        print(f"  📄 Đang đọc feedbacks từ: {os.path.basename(filepath)}")
        fb = extract_feedbacks(filepath)
        all_feedbacks.extend(fb)
        
    # Loại bỏ trùng lặp nếu cần
    seen = set()
    unique_feedbacks = []
    for fb in all_feedbacks:
        fb_lower = fb.lower().strip()
        if fb_lower not in seen:
            seen.add(fb_lower)
            unique_feedbacks.append(fb)
            
    if not unique_feedbacks:
        print("  ⚠️ Không tìm thấy feedback nào trong các cột AK.")
        return None
        
    # Tạo file Word
    doc = Document()
    
    # Cấu hình trang
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    
    # Style
    style = doc.styles['Normal']
    font = style.font
    font.name = CONFIG["font_name"]
    font.size = Pt(CONFIG["font_size"])
    style.paragraph_format.line_spacing = 1.5
    
    heading = doc.add_paragraph("TỔNG HỢP FEEDBACKS", style='Heading 1')
    heading.alignment = 1 # Center
    
    doc.add_paragraph(f"Tổng số ý kiến đóng góp: {len(unique_feedbacks)}")
    doc.add_paragraph()
    
    for i, fb in enumerate(unique_feedbacks, 1):
        p = doc.add_paragraph(f"{i}. {fb}")
        p.paragraph_format.space_after = Pt(6)
        
    doc.save(output_path)
    print(f"  ✅ Đã lưu {len(unique_feedbacks)} feedbacks vào {os.path.basename(output_path)}")
    return output_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Gộp feedbacks từ cột AK của nhiều file Excel")
    parser.add_argument("files", nargs="*", help="Danh sách file Excel (.xls)")
    args = parser.parse_args()
    
    files = args.files if args.files else [f for f in os.listdir('.') if f.endswith('.xls')]
    merge_feedbacks(files)
