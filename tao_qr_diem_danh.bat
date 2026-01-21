@echo off
chcp 65001 >nul
echo ============================================================
echo 🎯 TẠO QR CODE CHO DANH SÁCH ĐIỂM DANH
echo ============================================================
echo.
python create_qr_diem_danh.py
echo.
echo ============================================================
echo Nhấn phím bất kỳ để đóng cửa sổ...
pause >nul

