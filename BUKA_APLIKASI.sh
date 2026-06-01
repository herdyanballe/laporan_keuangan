#!/bin/bash
# Launcher Aplikasi Laporan Keuangan Kas - Kelompok Narogong

cd "$(dirname "$0")"

# Cek Python
if ! command -v python3 &> /dev/null; then
    osascript -e 'display dialog "Python tidak ditemukan!\nSilakan install Python dari https://www.python.org/downloads/" buttons {"OK"} default button "OK" with icon stop' 2>/dev/null || \
    echo "[ERROR] Python3 tidak ditemukan. Install dari https://www.python.org/"
    exit 1
fi

# Install dependensi jika belum
python3 -c "import openpyxl" 2>/dev/null || {
    echo "Menginstall dependensi..."
    pip3 install -r requirements.txt --quiet
}

# Jalankan aplikasi
python3 app.py
