# Aplikasi Laporan Keuangan Kas
## Kelompok Narogong — GKJ

---

## 📋 Cara Install & Jalankan

### 1. Pastikan Python sudah terinstall
Download Python 3.10+ dari https://www.python.org/downloads/
✅ Centang "Add Python to PATH" saat install

### 2. Install dependensi (hanya sekali)
Buka Command Prompt, masuk ke folder aplikasi, lalu jalankan:
```
pip install -r requirements.txt
```

### 3. Buka aplikasi — cukup double klik!

| File | Sistem Operasi | Keterangan |
|------|---------------|------------|
| BUKA_APLIKASI.vbs | Windows ✅ | Tanpa jendela CMD (DIREKOMENDASIKAN) |
| BUKA_APLIKASI.bat | Windows | Dengan jendela CMD |
| BUKA_APLIKASI.sh  | Mac / Linux | Klik kanan → Open |

💡 TIP WINDOWS: Klik kanan BUKA_APLIKASI.vbs → "Send to" → "Desktop (create shortcut)"
   Lalu double klik shortcut di Desktop kapan saja!

---

## 🖥️ Fitur Aplikasi

| Fitur | Keterangan |
|-------|-----------|
| Input Transaksi | Tambah uang masuk/keluar + keterangan secara manual |
| Import Excel | Import file Excel laporan yang sudah ada (format Narogong) |
| Filter | Filter tampilan per bulan dan tahun |
| Export Excel | Laporan rapi dengan logo, warna, dan saldo otomatis |
| Export PDF | Laporan siap cetak landscape A4 dengan logo |
| Pengaturan | Ganti nama kelompok, gereja, dan logo |

---

## 📁 Struktur File

```
laporan_keuangan/
├── app.py              ← File utama aplikasi
├── requirements.txt    ← Daftar library yang dibutuhkan
├── README.txt          ← Panduan ini
├── data_kas.json       ← Data transaksi (otomatis dibuat)
├── config.json         ← Pengaturan aplikasi (otomatis dibuat)
└── logo.png/jpg        ← Logo (dipilih via menu Pengaturan)
```

---

## 💡 Tips Import Excel

- File Excel harus memiliki kolom: **Tanggal, Keterangan, Kas Masuk, Kas Keluar**
- Format sudah disesuaikan dengan file laporan Kelompok Narogong
- Data duplikat otomatis dilewati saat import

---

## ⚠️ Catatan File yang Sudah Dianalisis

Dari file-file yang diupload, berikut file yang **DIREKOMENDASIKAN** untuk diimport
(urut dari yang paling awal, hindari duplikat):

1. `Laporan_Keuangan_...Mei_2024_sd_Juni_2024.xlsx` ← Data awal
2. `Laporan_Keuangan_..._September_2024.xlsx`
3. `Laporan_Keuangan_..._November_2024.xlsx`
4. `Laporan_Keuangan_..._Februari_2025.xlsx`
5. `Laporan_Keuangan_..._Mei_2025.xlsx`
6. `Laporan_Keuangan_..._Agustus_2025.xlsx`
7. `Rev_1_-_Laporan_..._Sept_-_Nov_2025.xlsx` ← Pakai Rev_1 (sudah dikoreksi)
8. `Laporan_Keuangan_..._Des_25-_Maret_2026.xlsx`

❌ JANGAN import file berikut (sudah tercakup / versi lama):
- `Laporan_..._Juni_2024.xlsx` (duplikat dengan Mei-Juni 2024)
- `Laporan_..._Oktober_2024.xlsx` (duplikat data)
- `Laporan_..._Sept_-_Nov_2025.xlsx` (pakai Rev_1 saja)
