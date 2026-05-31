#!/bin/bash

# Thiết lập màu sắc hiển thị
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}══════════════════════════════════════════════════${NC}"
echo -e "${BLUE}   ĐÓNG GÓI: Tổng hợp Khảo sát Excel → Word (Mac) ${NC}"
echo -e "${BLUE}   Tạo file .app chạy trực tiếp trên macOS        ${NC}"
echo -e "${BLUE}══════════════════════════════════════════════════${NC}"
echo ""

# 1. Kiểm tra Python
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Python3 chưa được cài đặt!${NC}"
    echo -e "Hãy cài đặt Python tại: https://www.python.org/downloads/"
    exit 1
fi
PYTHON_VERSION=$(python3 -V)
echo -e "ℹ️  Sử dụng: ${YELLOW}${PYTHON_VERSION}${NC}"

# 2. Cài đặt các thư viện phụ thuộc
echo -e "\n[1/4] Đang cài đặt/cập nhật thư viện..."
python3 -m pip install -r requirements.txt --quiet
if [ $? -ne 0 ]; then
    echo -e "${RED}❌ Lỗi cài đặt thư viện phụ thuộc!${NC}"
    exit 1
fi
echo -e "  ${GREEN}✅ Cài đặt thư viện thành công!${NC}"

# 3. Đóng gói ứng dụng bằng PyInstaller
echo -e "\n[2/4] Đang đóng gói ứng dụng bằng PyInstaller..."
echo -e "  (Quá trình này có thể mất 1-3 phút)"
python3 -m PyInstaller build_mac.spec --clean --noconfirm
if [ $? -ne 0 ]; then
    echo -e "${RED}❌ Lỗi đóng gói ứng dụng!${NC}"
    exit 1
fi
echo -e "  ${GREEN}✅ Đóng gói thành công!${NC}"

# 4. Gỡ bỏ thuộc tính cách ly (quarantine) Gatekeeper trên file build nội bộ
echo -e "\n[3/4] Tối ưu hóa quyền truy cập Gatekeeper..."
if [ -d "dist/TongHopKhaoSat.app" ]; then
    xattr -cr dist/TongHopKhaoSat.app
    if [ $? -eq 0 ]; then
        echo -e "  ${GREEN}✅ Đã loại bỏ cờ quarantine của macOS.${NC}"
    else
        echo -e "  ${YELLOW}⚠️ Không thể gỡ quarantine (có thể do quyền hạn).${NC}"
    fi
else
    echo -e "${RED}❌ Không tìm thấy thư mục dist/TongHopKhaoSat.app!${NC}"
    exit 1
fi

# 5. Hoàn tất
echo -e "\n${BLUE}══════════════════════════════════════════════════${NC}"
echo -e "  ${GREEN}🎉 HOÀN TẤT THÀNH CÔNG!${NC}"
echo -e "  Tệp ứng dụng: ${YELLOW}dist/TongHopKhaoSat.app${NC}"
echo -e "${BLUE}══════════════════════════════════════════════════${NC}"
echo ""

# Mở thư mục kết quả trong Finder
open dist

# Hướng dẫn chạy
echo -e "${YELLOW}💡 Mẹo khi chạy lần đầu:${NC}"
echo -e "Nếu xuất hiện cảnh báo Gatekeeper ngăn chặn mở ứng dụng trên các máy Mac khác,"
echo -e "người dùng có thể chạy lệnh sau trong Terminal:"
echo -e "  ${BLUE}xattr -cr /path/to/TongHopKhaoSat.app${NC}"
echo -e "Hoặc chuột phải vào ứng dụng và chọn ${YELLOW}Open (Mở)${NC}."
