@echo off
chcp 65001 >nul
title Đóng gói Tổng hợp Khảo sát → EXE

echo ══════════════════════════════════════════════════
echo   ĐÓNG GÓI: Tổng hợp Khảo sát Excel → Word
echo   Tạo file .exe chạy trên Windows không cần Python
echo ══════════════════════════════════════════════════
echo.

:: Kiểm tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python chưa được cài đặt!
    echo    Tải Python từ: https://www.python.org/downloads/
    echo    Chọn "Add Python to PATH" khi cài đặt.
    pause
    exit /b 1
)

:: Cài đặt dependencies
echo [1/3] Cài đặt thư viện cần thiết...
pip install -r requirements-dev.txt --quiet
if errorlevel 1 (
    echo ❌ Lỗi cài đặt thư viện!
    pause
    exit /b 1
)
echo   ✅ Thư viện OK

:: Build
echo.
echo [2/3] Đang đóng gói thành file EXE...
echo   (Quá trình này mất khoảng 1-3 phút)
echo.
pyinstaller build.spec --clean --noconfirm
if errorlevel 1 (
    echo ❌ Lỗi đóng gói!
    pause
    exit /b 1
)

:: Done
echo.
echo ══════════════════════════════════════════════════
echo   ✅ THÀNH CÔNG!
echo   File EXE: dist\TongHopKhaoSat.exe
echo ══════════════════════════════════════════════════
echo.

:: Mở thư mục dist
explorer dist

pause
