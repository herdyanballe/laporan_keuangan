@echo off
title Laporan Keuangan Web - Kelompok Narogong
cd /d "%~dp0"

echo Starting Streamlit web application...
echo.
echo Aplikasi akan terbuka di browser secara otomatis
echo Tekan Ctrl+C untuk menghentikan server
echo.

streamlit run app_streamlit.py