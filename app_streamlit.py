"""
Aplikasi Laporan Keuangan Kas Kelompok Narogong - Versi Web
Fitur: Input transaksi, import Excel, export PDF (Potrait), mobile friendly
"""

import streamlit as st
import pandas as pd
import json, os, io
from datetime import datetime, date
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image as RLImage,
                                 HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import tempfile

# Konfigurasi halaman (wajib di awal)
st.set_page_config(
    page_title="Laporan Kas Narogong",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="auto"
)

# Custom CSS untuk mobile view yang lebih baik
st.markdown("""
<style>
    /* Untuk mobile view */
    @media (max-width: 768px) {
        .stMetric {
            font-size: 14px !important;
        }
        .stDataFrame {
            font-size: 12px !important;
        }
        .stButton button {
            width: 100% !important;
            padding: 10px !important;
        }
        .stNumberInput input {
            font-size: 16px !important;
        }
        .stTextInput input {
            font-size: 16px !important;
        }
        h1 {
            font-size: 24px !important;
        }
        h2 {
            font-size: 20px !important;
        }
        h3 {
            font-size: 18px !important;
        }
    }
    
    /* Card style untuk mobile */
    .card-mobile {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    /* Warna saldo */
    .saldo-positive {
        color: #2E7D32;
        font-weight: bold;
    }
    .saldo-negative {
        color: #C62828;
        font-weight: bold;
    }
    
    /* Tombol di mobile */
    .stButton > button {
        border-radius: 8px;
        transition: all 0.3s;
    }
    .stButton > button:hover {
        transform: scale(1.02);
    }
</style>
""", unsafe_allow_html=True)

# ── Data storage ──────────────────────────────────────────────────────────────
DATA_FILE = "data_kas.json"
CONFIG_FILE = "config.json"

BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni",
            "Juli","Agustus","September","Oktober","November","Desember"]

def fmt_rp(val):
    try:
        return f"Rp {int(float(val)):,}".replace(",",".")
    except:
        return "Rp 0"

def parse_tgl(tgl_str):
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(tgl_str)[:10], fmt)
        except:
            continue
    return None

def fmt_tgl(tgl_str):
    d = parse_tgl(tgl_str)
    return d.strftime("%d-%m-%Y") if d else tgl_str

def tgl_to_sort_key(tgl_str):
    d = parse_tgl(tgl_str)
    return d if d else datetime(1900,1,1)

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"logo_path": "", "nama_kelompok": "Kelompok Narogong", "nama_gereja": "GKJ"}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def get_saldo_hari_ini(data):
    today = date.today()
    saldo = 0
    for r in sorted(data, key=lambda x: tgl_to_sort_key(x.get("tanggal", ""))):
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl and tgl.date() <= today:
            saldo += r.get("masuk", 0) - r.get("keluar", 0)
    return saldo

def filter_data(data, bulan, tahun):
    result = []
    for r in data:
        tgl = r.get("tanggal","")
        d = parse_tgl(tgl)
        if d:
            if tahun and str(d.year) != str(tahun):
                continue
            if bulan and str(d.month) != str(bulan):
                continue
        result.append(r)
    return sorted(result, key=lambda r: tgl_to_sort_key(r.get("tanggal","")))

def get_data_before_period(data, bulan_filter, tahun_filter):
    result = []
    for r in data:
        tgl = r.get("tanggal", "")
        d = parse_tgl(tgl)
        if d:
            if tahun_filter and d.year < int(tahun_filter):
                result.append(r)
            elif tahun_filter and d.year == int(tahun_filter) and bulan_filter:
                if d.month < int(bulan_filter):
                    result.append(r)
            elif tahun_filter and not bulan_filter:
                if d.year < int(tahun_filter):
                    result.append(r)
    return result

def build_periode_label(rows, bulan_filter, tahun_filter):
    if bulan_filter and tahun_filter:
        bulan_int = int(bulan_filter)
        return f"Bulan {BULAN_ID[bulan_int-1]} {tahun_filter}"
    if tahun_filter:
        return f"Tahun {tahun_filter}"
    if rows:
        try:
            tanggals = [parse_tgl(r["tanggal"]) for r in rows if r.get("tanggal")]
            tanggals = [d for d in tanggals if d]
            if tanggals:
                mn, mx = min(tanggals), max(tanggals)
                if mn.year == mx.year and mn.month == mx.month:
                    return f"Bulan {BULAN_ID[mn.month-1]} {mn.year}"
                return f"{BULAN_ID[mn.month-1]} {mn.year} s.d. {BULAN_ID[mx.month-1]} {mx.year}"
        except:
            pass
    return "Semua Periode"

# ── Fungsi Export PDF (Potrait) ──────────────────────────────────────────────
def export_pdf(data, rows, config, bulan_filter=None, tahun_filter=None):
    """Export ke PDF dengan orientasi Potrait"""
    
    # Buat file sementara
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_file.close()
    
    nama_kel = config.get("nama_kelompok", "Kelompok Narogong")
    judul_periode = build_periode_label(rows, bulan_filter, tahun_filter)
    
    # Orientasi Potrait (A4)
    doc = SimpleDocTemplate(temp_file.name, pagesize=A4,
                             leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    navy   = colors.HexColor("#1F4E79")
    silver = colors.HexColor("#EEF4FB")
    white  = colors.white
    
    style_title = ParagraphStyle("title", fontName="Helvetica-Bold",
                                  fontSize=14, textColor=navy, alignment=TA_CENTER,
                                  spaceAfter=4)
    style_sub   = ParagraphStyle("sub", fontName="Helvetica",
                                  fontSize=10, textColor=colors.grey,
                                  alignment=TA_CENTER, spaceAfter=2)
    style_cell  = ParagraphStyle("cell", fontName="Helvetica", fontSize=8,
                                  leading=10)
    
    story = []
    
    # Header dengan logo
    logo_path = config.get("logo_path", "")
    if logo_path and os.path.exists(logo_path):
        logo = RLImage(logo_path, width=2*cm, height=2*cm)
        header_data = [[logo,
                        [Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title),
                         Paragraph(judul_periode, style_sub)]]]
        header_tbl = Table(header_data, colWidths=[2.5*cm, 12*cm])
        header_tbl.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
        ]))
        story.append(header_tbl)
    else:
        story.append(Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title))
        story.append(Paragraph(judul_periode, style_sub))
    
    story.append(HRFlowable(width="100%", thickness=2, color=navy, spaceAfter=8))
    
    # Table data
    table_data = [["No","Tanggal","Keterangan","Kas Masuk","Kas Keluar","Saldo Akhir"]]
    
    # Hitung saldo
    if bulan_filter or tahun_filter:
        all_before = get_data_before_period(data, bulan_filter, tahun_filter)
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
    else:
        saldo = 0
    
    for idx, row in enumerate(rows, 1):
        saldo += row.get("masuk", 0) - row.get("keluar", 0)
        table_data.append([
            str(idx),
            row["tanggal"],
            Paragraph(row["keterangan"][:60], style_cell),
            fmt_rp(row.get("masuk", 0)) if row.get("masuk", 0) else "-",
            fmt_rp(row.get("keluar", 0)) if row.get("keluar", 0) else "-",
            fmt_rp(saldo),
        ])
    
    total_masuk = sum(r.get("masuk", 0) for r in rows)
    total_keluar = sum(r.get("keluar", 0) for r in rows)
    
    if bulan_filter or tahun_filter:
        all_before = get_data_before_period(data, bulan_filter, tahun_filter)
        saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
        table_data.append(["", "", "", "", "", ""])
        table_data.append(["", "", f"SALDO AWAL: {fmt_rp(saldo_awal)}", "", "", ""])
        table_data.append(["", "", "GRAND TOTAL", fmt_rp(total_masuk), fmt_rp(total_keluar), fmt_rp(saldo)])
    else:
        table_data.append(["", "", "GRAND TOTAL", fmt_rp(total_masuk), fmt_rp(total_keluar), fmt_rp(saldo)])
    
    # Lebar kolom untuk Potrait A4
    col_w = [0.8*cm, 2*cm, 7*cm, 2.5*cm, 2.5*cm, 2.8*cm]
    
    tbl = Table(table_data, colWidths=col_w, repeatRows=1)
    n = len(table_data)
    tbl_style = TableStyle([
        ("BACKGROUND", (0,0), (-1,0), navy),
        ("TEXTCOLOR", (0,0), (-1,0), white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 8),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("FONTNAME", (0,1), (-1,n-2), "Helvetica"),
        ("FONTSIZE", (0,1), (-1,n-2), 7),
        ("VALIGN", (0,1), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,1), (1,n-2), "CENTER"),
        ("ALIGN", (3,1), (5,n-2), "RIGHT"),
        *[("BACKGROUND", (0,i), (-1,i), silver) for i in range(1, n-1) if i%2==0],
        ("BACKGROUND", (0,n-1), (-1,n-1), navy),
        ("TEXTCOLOR", (0,n-1), (-1,n-1), white),
        ("FONTNAME", (0,n-1), (-1,n-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,n-1), (-1,n-1), 8),
        ("ALIGN", (2,n-1), (2,n-1), "CENTER"),
        ("ALIGN", (3,n-1), (5,n-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.4, colors.HexColor("#AAAAAA")),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ])
    tbl.setStyle(tbl_style)
    story.append(tbl)
    
    # Footer
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    tgl_cetak = datetime.now().strftime("%d %B %Y pukul %H:%M")
    story.append(Paragraph(f"Dicetak pada: {tgl_cetak}", style_sub))
    
    doc.build(story)
    
    # Baca file untuk di-download
    with open(temp_file.name, "rb") as f:
        pdf_data = f.read()
    
    os.unlink(temp_file.name)
    return pdf_data

# Inisialisasi session state
if 'data' not in st.session_state:
    st.session_state.data = load_data()
if 'config' not in st.session_state:
    st.session_state.config = load_config()

# ==================== UI COMPONENTS ====================

def show_dashboard():
    """Tampilan Dashboard dengan Kartu Saldo - Mobile Friendly"""
    
    today = date.today().strftime("%d %B %Y")
    saldo_hari_ini = get_saldo_hari_ini(st.session_state.data)
    total_masuk = sum(r.get("masuk", 0) for r in st.session_state.data)
    total_keluar = sum(r.get("keluar", 0) for r in st.session_state.data)
    
    # Menggunakan columns yang responsif
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="card-mobile">
            <small>📅 TANGGAL</small>
            <h3>{today}</h3>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        warna_saldo = "saldo-positive" if saldo_hari_ini >= 0 else "saldo-negative"
        st.markdown(f"""
        <div class="card-mobile">
            <small>💰 SALDO HARI INI</small>
            <h3 class="{warna_saldo}">{fmt_rp(saldo_hari_ini)}</h3>
        </div>
        """, unsafe_allow_html=True)
    
    col3, col4 = st.columns(2)
    with col3:
        st.markdown(f"""
        <div class="card-mobile">
            <small>📈 TOTAL PEMASUKAN</small>
            <h3 style="color:#2E7D32;">{fmt_rp(total_masuk)}</h3>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="card-mobile">
            <small>📉 TOTAL PENGELUARAN</small>
            <h3 style="color:#C62828;">{fmt_rp(total_keluar)}</h3>
        </div>
        """, unsafe_allow_html=True)

def show_transaksi_form():
    """Form Input Transaksi - Mobile Friendly"""
    with st.expander("➕ TAMBAH TRANSAKSI BARU", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            tgl = st.date_input("📅 Tanggal", value=date.today(), format="DD/MM/YYYY")
            keterangan = st.text_area("📝 Keterangan", height=80, 
                                       placeholder="Contoh: Iuran bulanan, Pembelian perlengkapan, dll")
        
        with col2:
            masuk = st.number_input("💰 Kas Masuk (Rp)", min_value=0, value=0, step=10000)
            keluar = st.number_input("💸 Kas Keluar (Rp)", min_value=0, value=0, step=10000)
        
        if st.button("💾 SIMPAN TRANSAKSI", type="primary", use_container_width=True):
            if not keterangan:
                st.error("❌ Keterangan wajib diisi!")
            elif masuk > 0 and keluar > 0:
                st.error("❌ Hanya boleh mengisi salah satu: Kas Masuk ATAU Kas Keluar!")
            elif masuk == 0 and keluar == 0:
                st.error("❌ Harap isi Kas Masuk atau Kas Keluar!")
            else:
                tgl_str = tgl.strftime("%d-%m-%Y")
                new_data = {
                    "tanggal": tgl_str,
                    "keterangan": keterangan.strip(),
                    "masuk": float(masuk),
                    "keluar": float(keluar)
                }
                st.session_state.data.append(new_data)
                save_data(st.session_state.data)
                st.success("✅ Transaksi berhasil ditambahkan!")
                st.rerun()

def show_data_table():
    """Tampilan Tabel Data - Mobile Friendly"""
    st.markdown("### 📋 DAFTAR TRANSAKSI")
    
    if not st.session_state.data:
        st.info("📭 Belum ada data. Silakan tambah transaksi baru.")
        return
    
    # Filter - responsive
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        bulan_filter = st.selectbox("📆 Filter Bulan", ["Semua"] + BULAN_ID)
    with col_f2:
        years = sorted({parse_tgl(r["tanggal"]).year for r in st.session_state.data 
                       if r.get("tanggal") and parse_tgl(r["tanggal"])}, reverse=True)
        tahun_filter = st.selectbox("📅 Filter Tahun", ["Semua"] + [str(y) for y in years])
    
    bulan = str(BULAN_ID.index(bulan_filter)+1) if bulan_filter != "Semua" else None
    tahun = tahun_filter if tahun_filter != "Semua" else None
    
    rows = filter_data(st.session_state.data, bulan, tahun)
    
    if not rows:
        st.info(f"Tidak ada data untuk periode {bulan_filter} {tahun_filter}.")
        return
    
    # Tombol export PDF
    col_pdf1, col_pdf2 = st.columns([3, 1])
    with col_pdf2:
        if st.button("📄 EXPORT PDF", type="primary", use_container_width=True):
            with st.spinner("Sedang membuat PDF..."):
                pdf_data = export_pdf(st.session_state.data, rows, st.session_state.config, bulan, tahun)
                st.download_button(
                    label="📥 DOWNLOAD PDF",
                    data=pdf_data,
                    file_name=f"Laporan_Kas_{bulan_filter}_{tahun_filter}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
    
    # Hitung saldo
    if bulan or tahun:
        all_before = get_data_before_period(st.session_state.data, bulan, tahun)
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
    else:
        saldo = 0
    
    # Buat dataframe
    table_data = []
    for idx, row in enumerate(rows, 1):
        saldo += row.get("masuk", 0) - row.get("keluar", 0)
        table_data.append({
            "No": idx,
            "Tgl": row["tanggal"],
            "Keterangan": row["keterangan"][:50] + "..." if len(row["keterangan"]) > 50 else row["keterangan"],
            "Masuk": fmt_rp(row.get("masuk", 0)) if row.get("masuk", 0) else "-",
            "Keluar": fmt_rp(row.get("keluar", 0)) if row.get("keluar", 0) else "-",
            "Saldo": fmt_rp(saldo)
        })
    
    df = pd.DataFrame(table_data)
    st.dataframe(df, use_container_width=True, height=400)
    
    # Ringkasan
    total_masuk = sum(r.get("masuk", 0) for r in rows)
    total_keluar = sum(r.get("keluar", 0) for r in rows)
    
    st.markdown("---")
    st.markdown("### 📊 RINGKASAN")
    
    if bulan or tahun:
        all_before = get_data_before_period(st.session_state.data, bulan, tahun)
        saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
        
        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
        with col_r1:
            st.metric("💰 Saldo Awal", fmt_rp(saldo_awal))
        with col_r2:
            st.metric("📈 Total Masuk", fmt_rp(total_masuk))
        with col_r3:
            st.metric("📉 Total Keluar", fmt_rp(total_keluar))
        with col_r4:
            warna = "green" if saldo >= 0 else "red"
            st.markdown(f'<div style="text-align:center"><small>🏁 Saldo Akhir</small><br><span style="color:{warna};font-size:24px;font-weight:bold">{fmt_rp(saldo)}</span></div>', unsafe_allow_html=True)
    else:
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.metric("📈 Total Masuk", fmt_rp(total_masuk))
        with col_r2:
            st.metric("📉 Total Keluar", fmt_rp(total_keluar))
        with col_r3:
            warna = "green" if saldo >= 0 else "red"
            st.markdown(f'<div style="text-align:center"><small>🏁 Saldo Akhir</small><br><span style="color:{warna};font-size:24px;font-weight:bold">{fmt_rp(saldo)}</span></div>', unsafe_allow_html=True)

def show_charts():
    """Tampilan Grafik"""
    st.markdown("### 📊 GRAFIK KEUANGAN")
    
    if not st.session_state.data:
        st.info("Belum ada data untuk ditampilkan.")
        return
    
    df = pd.DataFrame(st.session_state.data)
    df['tanggal_dt'] = df['tanggal'].apply(parse_tgl)
    df['bulan'] = df['tanggal_dt'].apply(lambda x: f"{BULAN_ID[x.month-1]} {x.year}" if x else None)
    df = df.dropna(subset=['bulan'])
    
    monthly = df.groupby('bulan').agg({
        'masuk': 'sum',
        'keluar': 'sum'
    }).reset_index()
    
    # Import plotly
    import plotly.graph_objects as go
    
    fig = go.Figure()
    fig.add_trace(go.Bar(name='📈 Pemasukan', x=monthly['bulan'], y=monthly['masuk'], 
                         marker_color='#2E7D32'))
    fig.add_trace(go.Bar(name='📉 Pengeluaran', x=monthly['bulan'], y=monthly['keluar'],
                         marker_color='#C62828'))
    fig.update_layout(
        barmode='group', 
        title="Pemasukan vs Pengeluaran per Bulan",
        xaxis_title="Bulan", 
        yaxis_title="Jumlah (Rp)",
        height=450,
        template="plotly_white"
    )
    st.plotly_chart(fig, use_container_width=True)

def show_settings():
    """Halaman Pengaturan"""
    st.markdown("### ⚙️ PENGATURAN")
    
    with st.form("settings_form"):
        nama_kelompok = st.text_input("Nama Kelompok", 
                                      value=st.session_state.config.get("nama_kelompok", "Kelompok Narogong"))
        nama_gereja = st.text_input("Nama Gereja", 
                                    value=st.session_state.config.get("nama_gereja", "GKJ"))
        
        uploaded_logo = st.file_uploader("Upload Logo (PNG/JPG)", type=['png', 'jpg', 'jpeg'])
        
        if uploaded_logo:
            st.image(uploaded_logo, width=100)
        
        if st.form_submit_button("💾 SIMPAN PENGATURAN", type="primary", use_container_width=True):
            st.session_state.config["nama_kelompok"] = nama_kelompok
            st.session_state.config["nama_gereja"] = nama_gereja
            
            if uploaded_logo:
                logo_path = Path("logo" + Path(uploaded_logo.name).suffix)
                with open(logo_path, "wb") as f:
                    f.write(uploaded_logo.getbuffer())
                st.session_state.config["logo_path"] = str(logo_path)
            
            save_config(st.session_state.config)
            st.success("✅ Pengaturan berhasil disimpan!")
            st.rerun()

# ==================== MAIN APP ====================

def main():
    # Header
    st.markdown(f"""
    <div style="text-align:center; padding:10px;">
        <h1>💰 {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')}</h1>
        <p style="color:gray;">{st.session_state.config.get('nama_gereja', 'GKJ')} • Sistem Laporan Keuangan Kas</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    # Sidebar untuk mobile (menggunakan radio button)
    with st.sidebar:
        st.markdown("### 📱 MENU")
        menu = st.radio(
            "Pilih Menu",
            ["🏠 Dashboard", "📝 Input Transaksi", "📋 Data Transaksi", 
             "📊 Grafik", "⚙️ Pengaturan"],
            label_visibility="collapsed"
        )
        st.divider()
        st.caption(f"📊 Total Transaksi: **{len(st.session_state.data)}**")
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.session_state.data = load_data()
            st.session_state.config = load_config()
            st.rerun()
    
    # Routing menu
    if menu == "🏠 Dashboard":
        show_dashboard()
        st.markdown("---")
        show_charts()
    elif menu == "📝 Input Transaksi":
        show_transaksi_form()
        st.markdown("---")
        show_data_table()
    elif menu == "📋 Data Transaksi":
        show_data_table()
    elif menu == "📊 Grafik":
        show_charts()
    elif menu == "⚙️ Pengaturan":
        show_settings()

if __name__ == "__main__":
    main()