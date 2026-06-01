"""
Generate Icons untuk PWA Kas Narogong
Jalankan: python generate_icons.py
"""

from PIL import Image, ImageDraw, ImageFont
import os

# Buat folder icons
os.makedirs("icons", exist_ok=True)

# Ukuran icon yang dibutuhkan
sizes = [72, 96, 128, 144, 152, 192, 384, 512]

# Warna
bg_color = "#1F4E79"  # Biru navy
text_color = "#FFFFFF"  # Putih
accent_color = "#FFD700"  # Emas

print("Mulai membuat icon...")
print("-" * 40)

for size in sizes:
    # Buat gambar dengan background biru
    img = Image.new('RGB', (size, size), color=bg_color)
    draw = ImageDraw.Draw(img)
    
    # Gambar lingkaran putih tipis sebagai border
    border_width = max(2, int(size * 0.02))
    for i in range(border_width):
        draw.ellipse([i, i, size-1-i, size-1-i], outline=text_color, width=1)
    
    # Untuk semua ukuran, tulis "Rp"
    try:
        font_size = int(size * 0.35)
        # Coba pakai font default
        font = ImageFont.load_default()
    except:
        font = ImageFont.load_default()
    
    # Tulis "Rp" di tengah
    text = "Rp"
    # Ukuran teks (perkiraan)
    text_width = size // 3
    text_height = size // 4
    x = (size - text_width) // 2
    y = (size - text_height) // 2
    
    # Gambar teks dengan cara sederhana
    for i, char in enumerate(text):
        draw.text((x + i * (text_width//2), y), char, fill=accent_color, font=font)
    
    # Simpan gambar
    filename = f"icons/icon-{size}.png"
    img.save(filename)
    print(f"  Membuat {filename}")

print("-" * 40)
print("Selesai! Semua icon telah dibuat di folder 'icons/'")