# Tổng hợp Khảo sát Excel → Word

## Hướng dẫn đóng gói trên Windows

### Yêu cầu
- Máy Windows có cài **Python 3.10+** (tải từ https://www.python.org/downloads/)
- Khi cài Python, **bắt buộc** tích chọn **"Add Python to PATH"**

### Cách build file .exe

1. Mở thư mục này trong **File Explorer**
2. **Double-click** vào file `build_windows.bat`
3. Chờ khoảng 1-3 phút
4. File `TongHopKhaoSat.exe` sẽ xuất hiện trong thư mục `dist/`

### Cách sử dụng (sau khi build)

1. Copy file `dist/TongHopKhaoSat.exe` ra bất kỳ đâu
2. Double-click để mở
3. Thêm file Excel (.xls) → Chọn thư mục lưu → Bấm "Bắt đầu xử lý"
4. File Word (.docx) sẽ được tạo tự động

### Lưu ý
- File .exe chạy **không cần cài Python** trên máy đích
- Chỉ cần gửi file `TongHopKhaoSat.exe` (khoảng 30-50MB) cho người dùng
- Nếu Windows báo "Windows protected your PC", bấm **"More info" → "Run anyway"**
