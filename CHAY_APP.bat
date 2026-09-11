@echo off
title FormsMobile - App Khảo Sát chuẩn Google Forms
chcp 65001 > nul
cls
echo ===================================================================
echo    🚀 ĐANG KHỞI CHẠY ỨNG DỤNG KHẢO SÁT FORMSMOBILE (MOBILE-FIRST)
echo ===================================================================
echo.
echo  📍 Địa chỉ App: http://localhost:8080
echo  ☁️ Supabase Cloud: Đã kết nối tự động
echo.
echo  [LƯU Ý]: Giữ cửa sổ này mở để App hoạt động. Đóng cửa sổ để tắt App.
echo ===================================================================
echo.

start http://localhost:8080
py server.py
