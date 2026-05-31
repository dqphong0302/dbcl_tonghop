# Tool 1: Tổng hợp dữ liệu khảo sát (Excel → Word)

## Mục đích

Đọc file Excel khảo sát sinh viên (`.xls`), trích xuất dữ liệu từng sheet, tổng hợp và xuất báo cáo Word (`.docx`) theo mẫu chuẩn của Khoa KHCB – Đại học Y Dược TP.HCM.

---

## Nguồn dữ liệu đầu vào

- **Định dạng file:** `.xls` (xlrd)
- **Cấu trúc file:** Mỗi file có thể có 1 hoặc nhiều sheet. Mỗi sheet tương ứng với 1 lớp học.
- **Ví dụ tên file:** `D25 Giáo dục thể chất 2 25-26 HKI 2.xls`
  - Số `2` ở cuối tên file = số sinh viên **KHÔNG** tham gia khảo sát.
  - Nếu 100% sinh viên tham gia → số đuôi = `0`.

---

## Quy tắc tính Tổng số sinh viên

Đây là logic **ưu tiên** để xác định tổng số sinh viên của học phần:

| Ưu tiên | Nguồn | Mô tả |
|---------|-------|-------|
| 1 | **Tên file** | Lấy số nguyên ở cuối tên file (sau khoảng trắng cuối cùng). `tong_so_sv = tong_phieu_khao_sat + sv_khong_tham_gia` |
| 2 | **Ô F5** trong Excel | Fallback nếu tên file không có số đuôi. Ô F5 chứa trực tiếp tổng số SV. |
| 3 | **Lỗi / bỏ qua** | Nếu cả 2 nguồn đều không có → hiển thị cảnh báo, **không tạo file** cho sheet đó. |

> **Lưu ý:** Nếu số đuôi tên file = `0` → `tong_so_sv = tong_phieu_khao_sat` (100% tham gia).

---

## Cấu trúc dữ liệu trích xuất từ Excel

### Metadata (header file)
- Đợt khảo sát
- Phiếu khảo sát (tên học phần)
- Thời gian khảo sát
- Lớp học
- Môn học (mã + tên)
- Giảng viên

### Thông tin nền (section "Thông tin chung")
Các câu hỏi dạng `C1:`, `C2:`... với cột `Số SV trả lời` và `Tỷ lệ`:
- Hình thức đào tạo (Chính quy / Khác)
- Giới tính (Nam / Nữ)
- Dân tộc thiểu số (Có / Không)
- Sinh viên nước ngoài (Có / Không)

### Câu hỏi khảo sát (section "Thông tin khảo sát")
- Thang đo Likert 5 mức (cột 2–6 = mức 1–5, cột 7 = tổng)
- Dòng tiếp theo chứa `Tỷ lệ` (tỷ lệ phần trăm)
- Tự động phân nhóm theo số lượng câu hỏi:
  - ≤10 câu → 1 section
  - 11–14 câu → 2 sections
  - 15–18 câu → 3 sections
  - ≥19 câu → 4 sections (I, II, III, IV)

### Câu hỏi mở
- Nhận diện bằng từ khóa "ý kiến đóng góp" trong cột A
- Nội dung ở cột C (index 2)

### Thông tin ký tên
- Ngày tháng năm (nhận diện bằng "ngày ... tháng ... năm")
- Người lập (dòng sau "Người lập")

---

## Xử lý nhiều sheet

Khi file có nhiều sheet:
1. Tạo file Word riêng cho từng sheet: `{tên_file}_sheet{N}.docx`
2. Tạo thêm file tổng hợp gộp tất cả sheet: `{tên_file}_TONG_HOP.docx`
3. Chỉ xử lý sheet có đủ dữ liệu khảo sát VÀ xác định được `tong_so_sv`

Khi gộp nhiều sheet:
- **Thông tin nền:** Cộng dồn số lượng, tính lại tỷ lệ
- **Câu hỏi khảo sát:** Cộng dồn counts, tính lại % và điểm TB
- **Ý kiến mở:** Nối chuỗi bằng `\n`
- **Tổng SV:** Cộng tổng từ tất cả sheet

---

## Cấu trúc file Word đầu ra

### Trang bìa / Header
- Cột trái: Tên trường + Khoa (Times New Roman 12pt, căn giữa)
- Cột phải: Quốc hiệu + Ngày tháng (căn giữa / phải)

### Tiêu đề báo cáo
```
BÁO CÁO TỔNG HỢP KẾT QUẢ
KHẢO SÁT SINH VIÊN SAU KHI KẾT THÚC HỌC PHẦN CƠ SỞ
& ĐỀ XUẤT KẾ HOẠCH CẢI TIẾN
```
(In đậm, căn giữa)

### Phần 1: Thông tin chung
Các dòng thông tin:
- Module/học phần khảo sát *(in đậm)*
- Thời gian khảo sát
- Đối tượng khảo sát *(in đậm)*
- Mẫu khảo sát
- Tổng số người học được khảo sát
- Tổng số người học được lấy ý kiến
- Tỷ lệ người học được lấy ý kiến

### Phần 2: Kết quả khảo sát
**2.1 Thông tin nền** – Bảng 4 cột (nhóm, phân loại, số lượng, tỷ lệ %)  
Header màu `#D9E2F3`, merge ô dọc cho nhóm.

**2.2 Thông tin cụ thể** – Bảng 8 cột:

| TT | Nội dung đánh giá | Mức 1 | Mức 2 | Mức 3 | Mức 4 | Mức 5 | Điểm TB |
|----|-------------------|-------|-------|-------|-------|-------|---------|

- Section header: in nghiêng, nền `#F2F2F2`, merge toàn hàng
- Tỷ lệ % làm tròn không thập phân
- Điểm TB: 2 chữ số thập phân

### Phần 3: Ý kiến đóng góp
Để trống (4 dòng trắng) cho người dùng tự điền.

### Phần 4: Đề xuất kế hoạch cải tiến
Bảng 4 cột để trống:
- Tiêu chí muốn cải tiến
- Lý do sinh viên chưa hài lòng
- Mục tiêu cải tiến cho năm sau
- Giải pháp đề xuất

### Footer
Chữ ký: `TRƯỞNG MODULE/TRƯỞNG BỘ MÔN` (căn phải)

---

## Định dạng chung

- Font: **Times New Roman**, 12pt
- Giãn dòng: 1.5 (văn bản), 1.0 (trong bảng)
- Trang A4: lề trái 2.5cm, phải 2cm, trên/dưới 2cm
- Tên file output: `{tên_file_gốc_bỏ_số_đuôi} (gộp).docx`

---

## Xử lý ý kiến mở (`parse_open_responses`)

Hàm tách các câu trả lời từ chuỗi nối liền trong 1 ô Excel:
1. Chuẩn hóa xuống dòng (`\r\n` → `\n`)
2. Tách theo `\n`
3. Tách tiếp theo pattern: chữ thường/dấu câu → chữ hoa đầu câu mới
4. Lọc bỏ câu trả lời vô nghĩa: "không", "ko", "không có ý kiến"...
5. Loại bỏ trùng lặp (case-insensitive)
6. Viết hoa chữ đầu, thêm dấu chấm cuối nếu thiếu

---

## Tích hợp GUI (`gui_app.py`)

- Nút **"▶ Bắt đầu xử lý"** → gọi `process_file()` cho từng file trong danh sách
- Hỗ trợ thêm file đơn lẻ hoặc cả thư mục
- Hiển thị progress bar + log realtime (thread-safe)
- Popup cảnh báo nếu sheet thiếu thông tin tổng SV
- Nút **"📂 Mở thư mục kết quả"** để mở thư mục output

---

## Dependencies

```
xlrd==1.2.0
python-docx
```
