@echo off
title Laporan Keuangan Kas - Kelompok Narogong
cd /d "%~dp0"

:: Cek apakah Python tersedia
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERROR] Python tidak ditemukan!
    echo  Silakan install Python dari https://www.python.org/downloads/
    echo  Pastikan centang "Add Python to PATH" saat install.
    echo.
    pause
    exit /b 1
)

:: Install dependensi jika belum ada
echo  Memeriksa dependensi...
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo  Menginstall library yang dibutuhkan, harap tunggu...
    pip install -r requirements.txt --quiet
    echo  Instalasi selesai!
)

:: Jalankan aplikasi
echo  Membuka Aplikasi Laporan Keuangan...
python app.py

:: Jika ada error saat jalan
if errorlevel 1 (
    echo.
    echo  [ERROR] Aplikasi berhenti dengan error.
    echo  Coba jalankan: pip install -r requirements.txt
    pause
)
