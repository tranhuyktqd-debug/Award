@echo off
chcp 65001 >nul
echo ============================================================
echo   KHỞI ĐỘNG ỨNG DỤNG XỬ LÝ MÃ CERT ASMO
echo ============================================================
echo.
echo Đang khởi động ứng dụng...
echo.

python awards_processing_app.py

if errorlevel 1 (
    echo.
    echo ❌ Lỗi khi khởi động ứng dụng!
    echo.
    pause
)
