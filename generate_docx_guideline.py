#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ═══════════════════════════════════════════════════════════════
# CONSTANTS & STYLES HELPERS
# ═══════════════════════════════════════════════════════════════
FONT_NAME = "Times New Roman"
COLOR_PRIMARY = RGBColor(30, 41, 59)     # Dark Slate
COLOR_SECONDARY = RGBColor(108, 99, 255) # Accent Purple
COLOR_BG_HEX = "F1F5F9"                  # Light grey/blue for notes
COLOR_HEADER_HEX = "D9E2F3"              # Blue for table headers
COLOR_BORDER_HEX = "CBD5E1"              # Soft grey border

def set_cell_shading(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_borders(cell, **kwargs):
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

def add_paragraph(doc, text="", bold=False, italic=False, font_size=12, 
                  color=COLOR_PRIMARY, alignment=WD_ALIGN_PARAGRAPH.LEFT, 
                  space_before=0, space_after=6, line_spacing=1.5):
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = FONT_NAME
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    return p

def add_heading(doc, text, level=1, space_before=12, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.keep_with_next = True
    
    run = p.add_run(text)
    run.bold = True
    run.font.name = FONT_NAME
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    if level == 1:
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(108, 99, 255) # Accent Purple
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif level == 2:
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(30, 41, 59)   # Dark slate
    else:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(71, 85, 105)  # Muted slate
        
    return p

def add_bullet_item(doc, bold_prefix="", text="", space_after=3):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.2
    
    if bold_prefix:
        r_prefix = p.add_run(bold_prefix)
        r_prefix.bold = True
        r_prefix.font.name = FONT_NAME
        r_prefix.font.size = Pt(12)
        r_prefix.font.color.rgb = COLOR_PRIMARY
        r_prefix._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
        
    if text:
        r_text = p.add_run(text)
        r_text.font.name = FONT_NAME
        r_text.font.size = Pt(12)
        r_text.font.color.rgb = COLOR_PRIMARY
        r_text._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    return p

def add_callout(doc, title, content):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_shading(cell, COLOR_BG_HEX)
    
    # Left border only (accent border style)
    border_spec = {"val": "single", "sz": "18", "color": "6C63FF"} # 2.25pt border
    set_cell_borders(cell, left=border_spec, top={"val": "none"}, bottom={"val": "none"}, right={"val": "none"})
    
    # Add title and content inside cell
    p_title = cell.paragraphs[0]
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = 1.15
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.name = FONT_NAME
    r_title.font.size = Pt(11)
    r_title.font.color.rgb = RGBColor(108, 99, 255)
    r_title._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    p_content = cell.add_paragraph()
    p_content.paragraph_format.space_after = Pt(4)
    p_content.paragraph_format.line_spacing = 1.15
    r_content = p_content.add_run(content)
    r_content.italic = True
    r_content.font.name = FONT_NAME
    r_content.font.size = Pt(11)
    r_content.font.color.rgb = COLOR_PRIMARY
    r_content._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    # Add an empty paragraph after the table to return to normal flow
    add_paragraph(doc, space_after=6)

# ═══════════════════════════════════════════════════════════════
# MAIN GENERATION
# ═══════════════════════════════════════════════════════════════
def generate_document():
    doc = Document()
    
    # Margins Setup
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2)
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        
    # Normal Style configuration
    style = doc.styles['Normal']
    font = style.font
    font.name = FONT_NAME
    font.size = Pt(12)
    
    # Header block
    add_paragraph(doc, "KHOA KHOA HỌC CƠ BẢN - ĐẠI HỌC Y DƯỢC TP.HCM", bold=True, font_size=10, color=RGBColor(100, 116, 139), alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph(doc, "BỘ PHẬN ĐẢM BẢO CHẤT LƯỢNG", bold=False, font_size=10, color=RGBColor(100, 116, 139), alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    
    # Title
    add_paragraph(doc, "HƯỚNG DẪN SỬ DỤNG PHẦN MỀM", bold=True, font_size=18, color=RGBColor(30, 41, 59), alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_paragraph(doc, "TỔNG HỢP KHẢO SÁT VÀ GỘP FEEDBACKS", bold=True, font_size=16, color=RGBColor(108, 99, 255), alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_paragraph(doc, "Phiên bản GUI v2  •  Định dạng Excel (.xls) → Word (.docx)", italic=True, font_size=11, color=RGBColor(100, 116, 139), alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    
    # ── PHẦN I: GIỚI THIỆU CHUNG ──
    add_heading(doc, "Phần I. Giới thiệu chung", level=1, space_before=18)
    add_paragraph(doc, "Phần mềm Tổng hợp Khảo sát là công cụ nội bộ được xây dựng nhằm tự động hóa quy trình xử lý dữ liệu khảo sát sinh viên sau khi kết thúc học phần của Khoa Khoa học Cơ bản (KHCB) – Đại học Y Dược TP.HCM. Phần mềm giúp chuyển đổi dữ liệu thô định lượng và định tính từ tệp bảng tính Excel (.xls) sang tệp báo cáo Word (.docx) chuyên nghiệp theo đúng biểu mẫu chuẩn của nhà trường.")
    
    add_paragraph(doc, "Phần mềm tích hợp hai công cụ chính:")
    add_bullet_item(doc, "1. Tổng hợp dữ liệu khảo sát (Tab 1): ", "Đọc dữ liệu điểm số Likert (1-5) và thông tin chung để tính toán tỷ lệ, điểm trung bình cho từng tiêu chí đánh giá, tự động nhóm các nội dung đánh giá theo phân bộ câu hỏi và xuất báo cáo Word hoàn chỉnh cho từng lớp và báo cáo tổng hợp.")
    add_bullet_item(doc, "2. Gộp ý kiến đóng góp (Tab 2): ", "Tổng hợp, lọc sạch các câu trả lời trùng lặp hoặc vô nghĩa (ví dụ: 'không', 'ko có', 'không ạ'...) từ cột ý kiến mở (mặc định cột AK hoặc tìm kiếm theo từ khóa) từ nhiều tệp Excel và gộp chung thành một tệp báo cáo duy nhất.")

    # ── PHẦN II: CHUẨN BỊ DỮ LIỆU ĐẦU VÀO ──
    add_heading(doc, "Phần II. Chuẩn bị dữ liệu đầu vào", level=1)
    add_paragraph(doc, "Để phần mềm hoạt động chính xác, tệp Excel khảo sát đầu vào cần tuân thủ một số quy định sau:")
    
    add_paragraph(doc, "1. Định dạng tệp:", bold=True)
    add_paragraph(doc, "Tệp dữ liệu phải ở định dạng Excel (.xls hoặc .xlsx). Phần mềm được tối ưu nhất cho cấu trúc kết xuất từ hệ thống khảo sát của trường.")
    
    add_paragraph(doc, "2. Quy tắc đặt tên tệp và Số sinh viên không tham gia:", bold=True)
    add_paragraph(doc, "Tổng số sinh viên của học phần được tự động tính toán theo công thức:")
    add_paragraph(doc, "Tổng số SV = Số phiếu khảo sát (thu về) + Số SV không tham gia khảo sát", italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_paragraph(doc, "Để xác định số sinh viên không tham gia khảo sát, hãy đặt tên tệp kết thúc bằng một khoảng trắng và một số nguyên tương ứng. Ví dụ:")
    add_bullet_item(doc, "D25 Giao duc the chat 1 25-26 HKI 2.xls ", "-> Phần mềm nhận diện có 2 sinh viên không tham gia khảo sát.")
    add_bullet_item(doc, "Lop Y2025 Hoa sinh hoc 0.xls ", "-> Phần mềm nhận diện 100% sinh viên tham gia khảo sát (0 sinh viên không tham gia).")

    add_callout(doc, "💡 Quy tắc dự phòng (Fallback logic) tính Tổng số SV", 
                "Nếu tên tệp không kết thúc bằng chữ số, phần mềm sẽ tự động đọc giá trị tại ô F5 của tệp Excel để xác định tổng số sinh viên đăng ký học phần. Nếu cả tên tệp và ô F5 đều không có thông tin này, phần mềm sẽ hiển thị cảnh báo và bỏ qua việc xử lý để tránh sai lệch dữ liệu.")

    # ── PHẦN III: HƯỚNG DẪN SỬ DỤNG GIAO DIỆN (GUI) ──
    add_heading(doc, "Phần III. Hướng dẫn sử dụng Giao diện Đồ họa (GUI)", level=1)
    add_paragraph(doc, "Khởi chạy ứng dụng bằng cách nhấp đúp vào biểu tượng TongHopKhaoSat.app (trên macOS) hoặc chạy tệp script build_mac.sh.")
    
    add_heading(doc, "1. Chức năng 1: Tổng hợp dữ liệu khảo sát (Tab 'Tổng hợp báo cáo')", level=2)
    add_paragraph(doc, "Bước 1: ", bold=True, space_after=2)
    add_paragraph(doc, "Mở tab \"Tổng hợp báo cáo\" (Tab thứ nhất). Click nút \"➕ Thêm file\" để chọn các tệp Excel cần tổng hợp, hoặc nút \"📁 Thư mục\" để chọn cả thư mục chứa các tệp khảo sát.")
    add_paragraph(doc, "Bước 2: ", bold=True, space_after=2)
    add_paragraph(doc, "Kiểm tra danh sách tệp được tải lên trong bảng hiển thị. Bạn có thể chọn hàng và nhấn nút \"✕ Xóa dòng đã chọn\" (hoặc phím Delete) để loại bỏ các tệp không cần thiết.")
    add_paragraph(doc, "Bước 3: ", bold=True, space_after=2)
    add_paragraph(doc, "Tại mục \"Thư mục kết quả\", click nút \"Chọn...\" để chỉ định nơi lưu tệp Word đầu ra. Nếu để trống, phần mềm tự tạo thư mục \"output\" nằm cùng cấp với tệp Excel đầu vào.")
    add_paragraph(doc, "Bước 4: ", bold=True, space_after=2)
    add_paragraph(doc, "Nhấp nút \"▶ Bắt đầu xử lý\". Thanh tiến trình sẽ hiển thị phần trăm hoàn tất và bảng nhật ký (Log) sẽ ghi lại kết quả xử lý chi tiết (thành công, cảnh báo, lỗi) theo thời gian thực.")
    add_paragraph(doc, "Bước 5: ", bold=True, space_after=2)
    add_paragraph(doc, "Khi xử lý xong, click nút \"📂 Mở kết quả\" để mở nhanh thư mục chứa tệp Word báo cáo.")

    add_heading(doc, "2. Chức năng 2: Tổng hợp Ý kiến mở (Tab 'Gộp Feedbacks')", level=2)
    add_paragraph(doc, "Tab này được thiết kế để đọc cột ý kiến phản hồi định tính (mặc định cột AK, hoặc cột chứa tiêu đề 'ý kiến đóng góp') của toàn bộ sinh viên từ các lớp học phần khác nhau và gom chung thành một báo cáo duy nhất giúp Giáo vụ/Trưởng Module dễ dàng theo dõi.")
    
    add_paragraph(doc, "Quy trình thực hiện tương tự Tab 1:")
    add_bullet_item(doc, "• ", "Thêm các tệp Excel khảo sát.")
    add_bullet_item(doc, "• ", "Chọn thư mục đầu ra và nhấp nút \"💬 Gộp Feedbacks\".")
    add_bullet_item(doc, "• ", "Hệ thống sẽ lọc bỏ các câu trả lời ngắn, trống hoặc vô nghĩa như: 'không', 'dạ không ạ', 'ko có', 'không có ý kiến gì',... tự động loại trừ trùng lặp nội dung giống nhau, viết hoa chữ cái đầu và thêm dấu chấm câu.")
    add_bullet_item(doc, "• ", "Kết quả xuất ra tệp Feedbacks.docx gồm danh sách các ý kiến được đánh số thứ tự chỉn chu.")

    # ── PHẦN IV: CẤU TRÚC BÁO CÁO ĐẦU RA ──
    add_heading(doc, "Phần IV. Cấu trúc báo cáo đầu ra (Word)", level=1)
    add_paragraph(doc, "Tệp báo cáo Word đầu ra được xây dựng theo định dạng chuẩn:")
    add_bullet_item(doc, "• Font chữ: ", "Times New Roman, kích thước 12pt.")
    add_bullet_item(doc, "• Giãn dòng: ", "1.5 line spacing cho văn bản thường, 1.0 cho dữ liệu trong bảng.")
    add_bullet_item(doc, "• Cấu trúc nội dung gồm 4 phần chính: ", "")
    add_bullet_item(doc, "  - Phần 1. Thông tin chung: ", "Trình bày thông tin tên học phần, đối tượng lớp học, giảng viên, thời gian khảo sát, số lượng sinh viên tham gia và tỷ lệ phản hồi (%).")
    add_bullet_item(doc, "  - Phần 2. Kết quả khảo sát: ", "Bao gồm bảng 2.1 Thông tin nền (Giới tính, hình thức học...) và bảng 2.2 Thông tin cụ thể hiển thị tỷ lệ phần trăm phân bố từ Mức 1 đến Mức 5 cho từng tiêu chí, kèm theo điểm trung bình (thang 5.0) làm tròn 2 chữ số thập phân.")
    add_bullet_item(doc, "  - Phần 3. Ý kiến đóng góp của người học: ", "Tạo sẵn khoảng trắng định dạng chuẩn (4 dòng trống) để giảng viên tự điền ghi chú hoặc chèn nội dung nhận xét.")
    add_bullet_item(doc, "  - Phần 4. Đề xuất kế hoạch cải tiến: ", "Bảng trống gồm 4 cột (Tiêu chí muốn cải tiến, Lý do sinh viên chưa hài lòng, Mục tiêu cải tiến cho năm sau, Giải pháp đề xuất) để trưởng học phần trực tiếp điền vào.")

    # ── PHẦN V: XỬ LÝ LỖI VÀ LƯU Ý ──
    add_heading(doc, "Phần V. Xử lý lỗi thường gặp", level=1)
    
    # Table of errors
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    
    headers = ["Hiện tượng / Thông báo lỗi", "Nguyên nhân", "Hướng khắc phục"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        set_cell_shading(cell, COLOR_HEADER_HEX)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(11)
        run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
        
    error_cases = [
        (
            "Cảnh báo: Thiếu thông tin Tổng SV\n(Popup hoặc ghi trong Log)",
            "Phần mềm không tìm thấy số lượng sinh viên không tham gia ở cuối tên tệp và ô F5 trong Excel cũng bị để trống.",
            "Điền số lượng sinh viên không tham gia khảo sát ở cuối tên tệp (ví dụ: 'HoaSinhHoc 2.xls') hoặc điền tổng số SV lớp vào ô F5 của file Excel, sau đó thực hiện lại."
        ),
        (
            "Thông báo: 'TongHopKhaoSat' cannot be opened because the developer cannot be verified (Lỗi bảo mật macOS)",
            "macOS kích hoạt cơ chế Gatekeeper cách ly các tệp tải về từ Internet (quarantine) không có chữ ký số của Apple Developer.",
            "Cách 1: Nhấp chuột phải vào ứng dụng TongHopKhaoSat.app và chọn Open (Mở).\nCách 2: Mở Terminal, gõ lệnh 'xattr -cr ' rồi kéo ứng dụng vào và nhấn Enter."
        ),
        (
            "Lỗi: Không hiển thị được chữ tiếng Việt hoặc lỗi định dạng bảng",
            "Sử dụng phiên bản Office hoặc trình xem văn bản quá cũ không tương thích tốt với mã XML của file docx.",
            "Khuyên dùng Microsoft Word phiên bản 2016 trở lên hoặc Office 365 để hiển thị văn bản với định dạng và font Times New Roman chuẩn xác nhất."
        )
    ]
    
    for case in error_cases:
        row = table.add_row()
        for i, text in enumerate(case):
            cell = row.cells[i]
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(text)
            run.font.name = FONT_NAME
            run.font.size = Pt(10)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
            
    # Margin settings for columns in the error table
    for row in table.rows:
        row.cells[0].width = Cm(5)
        row.cells[1].width = Cm(5)
        row.cells[2].width = Cm(6.5)
        for cell in row.cells:
            for para in cell.paragraphs:
                para.paragraph_format.space_after = Pt(2)
                para.paragraph_format.space_before = Pt(2)

    doc.add_paragraph() # Spacer
    add_paragraph(doc, "Chúc bạn có trải nghiệm làm việc hiệu quả cùng phần mềm!", italic=True, font_size=11, color=RGBColor(108, 99, 255), alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=12)
    
    # Save document
    output_filename = "Huong_Dan_Su_Dung_Phan_Mem_Tong_Hop_Khao_Sat.docx"
    doc.save(output_filename)
    print(f"SUCCESS: Generated {output_filename}")

if __name__ == "__main__":
    generate_document()
