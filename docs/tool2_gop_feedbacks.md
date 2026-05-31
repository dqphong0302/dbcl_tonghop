# Tool 2: Tổng hợp Feedbacks (Cột AK)

## Mục đích

Đọc nhiều file Excel khảo sát (`.xls`), trích xuất toàn bộ câu trả lời định tính từ **cột AK** (index 36), làm sạch, loại trùng lặp và xuất ra 1 file Word (`Feedbacks.docx`) tổng hợp.

---

## Nguồn dữ liệu đầu vào

- **Định dạng file:** `.xls` (xlrd)
- **Cột dữ liệu:** Cột AK = index 36 (0-based)
- **Ví dụ nguồn:** `1.xls`, `2.xls`, `3.xls`, `4.xls`
- Bỏ qua dòng đầu tiên (dòng tiêu đề, `r == 0`)
- Xử lý tất cả sheet trong mỗi file (nếu file có nhiều sheet)

---

## Quy trình xử lý

### Bước 1: Đọc dữ liệu từ cột AK

```
Với mỗi file .xls:
  Với mỗi sheet:
    Nếu sheet có > 36 cột:
      Với mỗi dòng (bỏ dòng 0):
        Lấy giá trị ô [r][36]
        Nếu không rỗng → đưa vào danh sách
```

### Bước 2: Làm sạch câu trả lời

Tái sử dụng hàm `parse_open_responses()` từ `excel_to_word.py`:

1. Chuẩn hóa xuống dòng (`\r\n`, `\r` → `\n`)
2. Tách theo `\n`
3. Tách tiếp theo pattern regex: chữ thường/dấu câu → chữ hoa đầu câu mới
4. Lọc bỏ câu trả lời vô nghĩa (xem bảng bên dưới)
5. Viết hoa chữ đầu câu
6. Thêm dấu chấm cuối nếu thiếu

**Danh sách pattern bị lọc bỏ:**

| Pattern | Ví dụ |
|---------|-------|
| `^không\s*(ạ\|có)?\.$` | "không", "không ạ" |
| `^ko\.?$` | "ko" |
| `^khong\.?$` | "khong" |
| `^dạ\s*không\s*(ạ\|có)?\.?$` | "dạ không ạ" |
| `^không\s*có\s*(góp\s*ý\|ý\s*kiến)?\.?$` | "không có ý kiến" |
| `^em\s*không\s*có\s*ý\s*kiến\s*(ạ)?\.?$` | "em không có ý kiến ạ" |

### Bước 3: Loại bỏ trùng lặp toàn cục

- So sánh case-insensitive (`.lower().strip()`)
- Giữ thứ tự xuất hiện đầu tiên
- Áp dụng sau khi gộp tất cả file

---

## File đầu ra

### Tên file
`Feedbacks.docx` — lưu vào thư mục output được chỉ định (hoặc cùng thư mục file đầu tiên nếu không chỉ định).

### Cấu trúc nội dung

```
TỔNG HỢP FEEDBACKS          ← Heading 1, căn giữa
Tổng số ý kiến đóng góp: N  ← Dòng thông tin
                             ← Dòng trống
1. [Câu trả lời 1]
2. [Câu trả lời 2]
...
```

### Định dạng

- Font: **Times New Roman**, 12pt (lấy từ `CONFIG` của `excel_to_word.py`)
- Giãn dòng: 1.5
- Mỗi mục cách nhau 6pt (`space_after = Pt(6)`)
- Trang A4: lề trái 2.5cm, phải 2cm, trên/dưới 2cm

---

## Tích hợp GUI (`gui_app.py`)

- Nút **"💬 Gộp Feedbacks (Cột AK)"** → gọi `merge_feedbacks(files, out_dir)`
- Dùng chung danh sách file đầu vào với Tool 1
- Chạy trong thread riêng (không block UI)
- Log realtime: tên từng file đang đọc, số feedback tìm được
- Hiển thị cảnh báo nếu không tìm thấy feedback nào

---

## Interface hàm

### `extract_feedbacks(filepath) → list[str]`
- Input: đường dẫn file `.xls`
- Output: danh sách câu trả lời đã làm sạch (có thể trùng lặp)
- Xử lý lỗi: in cảnh báo, trả về `[]`

### `merge_feedbacks(input_files, output_dir=None) → str | None`
- Input: danh sách đường dẫn file, thư mục output (tùy chọn)
- Output: đường dẫn file `Feedbacks.docx` đã tạo, hoặc `None` nếu không có dữ liệu
- Tự tạo thư mục output nếu chưa tồn tại

---

## CLI (chạy độc lập)

```bash
# Xử lý tất cả .xls trong thư mục hiện tại
python3 merge_feedbacks.py

# Xử lý file cụ thể
python3 merge_feedbacks.py 1.xls 2.xls 3.xls 4.xls
```

---

## Dependencies

```
xlrd==1.2.0
python-docx
```

Tái sử dụng từ `excel_to_word.py`:
- `parse_open_responses(raw_text)` — hàm làm sạch câu trả lời
- `CONFIG` — cấu hình font, cỡ chữ
