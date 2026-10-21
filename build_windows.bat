@echo off
chcp 65001 >nul
title Đóng gói Tổng hợp Khảo sát → EXE
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python chưa được cài đặt! Tải tại https://www.python.org/downloads/ ^(chọn "Add Python to PATH"^)
    pause
    exit /b 1
)

echo [1/2] Cài đặt thư viện...
python -m pip install -q -r requirements-dev.txt || goto :error

echo [2/2] Đóng gói thành file EXE ^(1-3 phút^)...
python -m PyInstaller TongHopKhaoSat.spec --clean --noconfirm || goto :error

echo.
echo Hoàn tất: dist\TongHopKhaoSat.exe
explorer dist
pause
exit /b 0

:error
echo Đóng gói thất bại!
pause
exit /b 1
