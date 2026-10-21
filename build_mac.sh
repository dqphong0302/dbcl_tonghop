#!/bin/bash
# Đóng gói ứng dụng macOS (.app) và nén thành dist/TongHopKhaoSat-macOS.zip
set -euo pipefail
cd "$(dirname "$0")"

echo "[1/3] Cài đặt thư viện..."
python3 -m pip install -q -r requirements-dev.txt

echo "[2/3] Đóng gói bằng PyInstaller (1-3 phút)..."
python3 -m PyInstaller TongHopKhaoSat.spec --clean --noconfirm

echo "[3/3] Nén file .app..."
xattr -cr dist/TongHopKhaoSat.app || true
rm -f dist/TongHopKhaoSat-macOS.zip
ditto -c -k --sequesterRsrc --keepParent dist/TongHopKhaoSat.app dist/TongHopKhaoSat-macOS.zip

echo "✅ Hoàn tất: dist/TongHopKhaoSat.app và dist/TongHopKhaoSat-macOS.zip"
echo "💡 Lần đầu mở trên máy khác: chuột phải vào ứng dụng → Open (Mở)."
