"""
Aplikasi Laporan Keuangan Kas Kelompok Narogong - Versi Web
Fitur: Input transaksi, Export PDF dengan logo, Mobile Friendly
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

# Konfigurasi halaman
st.set_page_config(
    page_title="Kas Narogong",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# CSS minimal untuk mobile
st.markdown("""
<style>
    .main > div { padding: 0.5rem !important; }
    .stButton > button { width: 100%; border-radius: 8px; }
    @media (max-width: 768px) {
        .stMarkdown h1 { font-size: 22px !important; }
        .stMarkdown h2 { font-size: 18px !important; }
        .stMarkdown h3 { font-size: 16px !important; }
    }
    .block-container { padding-top: 1rem !important; padding-bottom: 0rem !important; }
</style>
""", unsafe_allow_html=True)

# ==================== DATA STORAGE ====================

DATA_FILE = "data_kas.json"
CONFIG_FILE = "config.json"

BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni",
            "Juli","Agustus","September","Oktober","November","Desember"]

@st.cache_data(ttl=60)
def load_data_cached():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def load_data():
    return load_data_cached()

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    st.cache_data.clear()

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"logo_path": "", "nama_kelompok": "Kelompok Narogong", "nama_gereja": "GKJ"}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

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
    return "Semua Periode"

# ==================== EXPORT PDF ====================

def export_pdf(data, rows, config, bulan_filter=None, tahun_filter=None):
    """Export ke PDF dengan orientasi Potrait dan logo"""
    
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_file.close()
    
    nama_kel = config.get("nama_kelompok", "Kelompok Narogong")
    judul_periode = build_periode_label(rows, bulan_filter, tahun_filter)
    
    doc = SimpleDocTemplate(temp_file.name, pagesize=A4,
                             leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    navy   = colors.HexColor("#1F4E79")
    silver = colors.HexColor("#EEF4FB")
    white  = colors.white
    
    style_logo_text = ParagraphStyle("logo_text", fontName="Helvetica",
                                      fontSize=9, textColor=colors.grey,
                                      alignment=TA_CENTER, spaceAfter=2)
    style_title = ParagraphStyle("title", fontName="Helvetica-Bold",
                                  fontSize=14, textColor=navy, alignment=TA_CENTER,
                                  spaceAfter=2)
    style_periode = ParagraphStyle("periode", fontName="Helvetica",
                                    fontSize=12, textColor=navy, alignment=TA_CENTER,
                                    spaceAfter=8)
    style_sub   = ParagraphStyle("sub", fontName="Helvetica",
                                  fontSize=10, textColor=colors.grey,
                                  alignment=TA_CENTER, spaceAfter=2)
    style_cell  = ParagraphStyle("cell", fontName="Helvetica", fontSize=8,
                                  leading=10)
    
    story = []
    
    # HEADER DENGAN LOGO
    logo_path = config.get("logo_path", "")
    
    possible_logo_paths = [
        logo_path,
        "logo.png",
        "logo.jpg",
        "logo.jpeg",
        os.path.join(os.path.dirname(__file__), "logo.png"),
        os.path.join(os.path.dirname(__file__), "logo.jpg"),
    ]
    
    logo_found = None
    for path in possible_logo_paths:
        if path and os.path.exists(path):
            logo_found = path
            break
    
    if logo_found and os.path.exists(logo_found):
        try:
            logo = RLImage(logo_found, width=2.2*cm, height=2.2*cm, kind='proportional')
            
            header_data = [[
                logo,
                [Paragraph("GKJ BEKASI", style_logo_text)]
            ]]
            
            header_tbl = Table(header_data, colWidths=[2.5*cm, 12*cm])
            header_tbl.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ALIGN", (0,0), (0,0), "CENTER"),
                ("ALIGN", (1,0), (1,0), "CENTER"),
                ("LEFTPADDING", (0,0), (-1,-1), 0),
                ("RIGHTPADDING", (0,0), (-1,-1), 0),
                ("TOPPADDING", (0,0), (-1,-1), 0),
                ("BOTTOMPADDING", (0,0), (-1,-1), 0),
            ]))
            story.append(header_tbl)
            
        except Exception as e:
            story.append(Paragraph("GKJ BEKASI", style_logo_text))
    else:
        story.append(Paragraph("GKJ BEKASI", style_logo_text))
    
    # JUDUL 2 BARIS
    story.append(Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title))
    story.append(Paragraph(judul_periode, style_periode))
    
    story.append(HRFlowable(width="100%", thickness=2, color=navy, spaceAfter=8))
    
    # TABLE DATA
    table_data = [["No", "Tgl", "Keterangan", "Masuk", "Keluar", "Saldo"]]
    
    if bulan_filter or tahun_filter:
        all_before = get_data_before_period(data, bulan_filter, tahun_filter)
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
    else:
        saldo = 0
    
    for idx, row in enumerate(rows, 1):
        saldo += row.get("masuk", 0) - row.get("keluar", 0)
        ket = row["keterangan"]
        if len(ket) > 55:
            ket = ket[:52] + "..."
        
        table_data.append([
            str(idx),
            row["tanggal"],
            Paragraph(ket, style_cell),
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
        table_data.append(["", "", "TOTAL", fmt_rp(total_masuk), fmt_rp(total_keluar), fmt_rp(saldo)])
    else:
        table_data.append(["", "", "TOTAL", fmt_rp(total_masuk), fmt_rp(total_keluar), fmt_rp(saldo)])
    
    col_w = [0.7*cm, 1.8*cm, 6.8*cm, 2.3*cm, 2.3*cm, 2.5*cm]
    
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
        ("ALIGN", (0,1), (1,-1), "CENTER"),
        ("ALIGN", (3,1), (5,-1), "RIGHT"),
        *[("BACKGROUND", (0,i), (-1,i), silver) for i in range(1, n-1) if i%2 == 0],
        ("BACKGROUND", (0,n-1), (-1,n-1), navy),
        ("TEXTCOLOR", (0,n-1), (-1,n-1), white),
        ("FONTNAME", (0,n-1), (-1,n-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,n-1), (-1,n-1), 8),
        ("ALIGN", (2,n-1), (2,n-1), "CENTER"),
        ("ALIGN", (3,n-1), (5,n-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ])
    
    if bulan_filter or tahun_filter and len(table_data) > 2:
        tbl_style.add("BACKGROUND", (0, n-2), (-1, n-2), colors.HexColor("#E8F0FE"))
        tbl_style.add("FONTNAME", (0, n-2), (-1, n-2), "Helvetica-Bold")
        tbl_style.add("FONTSIZE", (0, n-2), (-1, n-2), 7)
        tbl_style.add("ALIGN", (2, n-2), (2, n-2), "LEFT")
    
    tbl.setStyle(tbl_style)
    story.append(tbl)
    
    # FOOTER
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    tgl_cetak = datetime.now().strftime("%d %B %Y %H:%M")
    story.append(Paragraph(f"Dicetak: {tgl_cetak}", style_sub))
    
    doc.build(story)
    
    with open(temp_file.name, "rb") as f:
        pdf_data = f.read()
    
    os.unlink(temp_file.name)
    return pdf_data

# ==================== UI COMPONENTS ====================

def show_dashboard():
    today = date.today().strftime("%d/%m/%Y")
    saldo = get_saldo_hari_ini(st.session_state.data)
    total_masuk = sum(r.get("masuk", 0) for r in st.session_state.data)
    total_keluar = sum(r.get("keluar", 0) for r in st.session_state.data)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📅 Hari Ini", today)
    with col2:
        st.metric("💰 Saldo", fmt_rp(saldo))
    
    col3, col4 = st.columns(2)
    with col3:
        st.metric("📈 Masuk", fmt_rp(total_masuk))
    with col4:
        st.metric("📉 Keluar", fmt_rp(total_keluar))

def show_transaksi_form():
    with st.expander("➕ Tambah Transaksi", expanded=False):
        tgl = st.date_input("Tanggal", value=date.today(), format="DD/MM/YYYY")
        keterangan = st.text_input("Keterangan", placeholder="Contoh: Iuran bulanan")
        col1, col2 = st.columns(2)
        with col1:
            masuk = st.number_input("Kas Masuk", min_value=0, value=0, step=10000)
        with col2:
            keluar = st.number_input("Kas Keluar", min_value=0, value=0, step=10000)
        
        if st.button("💾 Simpan", type="primary"):
            if not keterangan:
                st.error("Keterangan wajib diisi!")
            elif masuk > 0 and keluar > 0:
                st.error("Isi salah satu: Masuk atau Keluar!")
            elif masuk == 0 and keluar == 0:
                st.error("Isi nominal!")
            else:
                tgl_str = tgl.strftime("%d-%m-%Y")
                st.session_state.data.append({
                    "tanggal": tgl_str,
                    "keterangan": keterangan,
                    "masuk": float(masuk),
                    "keluar": float(keluar)
                })
                save_data(st.session_state.data)
                st.success("✅ Tersimpan!")
                st.rerun()

def show_data_table():
    st.subheader("📋 Transaksi")
    
    if not st.session_state.data:
        st.info("Belum ada data")
        return
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        bulan_filter = st.selectbox("Bulan", ["Semua"] + BULAN_ID, key="bulan_filter")
    with col_f2:
        years = sorted({parse_tgl(r["tanggal"]).year for r in st.session_state.data 
                       if r.get("tanggal") and parse_tgl(r["tanggal"])}, reverse=True)
        tahun_filter = st.selectbox("Tahun", ["Semua"] + [str(y) for y in years], key="tahun_filter")
    
    bulan = str(BULAN_ID.index(bulan_filter)+1) if bulan_filter != "Semua" else None
    tahun = tahun_filter if tahun_filter != "Semua" else None
    
    rows = filter_data(st.session_state.data, bulan, tahun)
    
    if rows:
        col_pdf, _ = st.columns([1, 3])
        with col_pdf:
            if st.button("📄 Export PDF", use_container_width=True):
                with st.spinner("Membuat PDF..."):
                    pdf_data = export_pdf(st.session_state.data, rows, st.session_state.config, bulan, tahun)
                    st.download_button(
                        label="📥 Download PDF",
                        data=pdf_data,
                        file_name=f"laporan_{bulan_filter}_{tahun_filter}.pdf",
                        mime="application/pdf"
                    )
        
        df_display = []
        saldo = 0
        if bulan or tahun:
            saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in get_data_before_period(st.session_state.data, bulan, tahun))
            saldo = saldo_awal
        
        for i, row in enumerate(rows, 1):
            saldo += row.get("masuk", 0) - row.get("keluar", 0)
            df_display.append({
                "No": i,
                "Tgl": row["tanggal"],
                "Ket": row["keterangan"][:35] + ".." if len(row["keterangan"]) > 35 else row["keterangan"],
                "Masuk": fmt_rp(row.get("masuk", 0)) if row.get("masuk") else "-",
                "Keluar": fmt_rp(row.get("keluar", 0)) if row.get("keluar") else "-",
            })
        
        st.dataframe(pd.DataFrame(df_display), use_container_width=True, height=400)
        
        total_masuk = sum(r.get("masuk", 0) for r in rows)
        total_keluar = sum(r.get("keluar", 0) for r in rows)
        
        st.markdown("---")
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.metric("Total Masuk", fmt_rp(total_masuk))
        with col_r2:
            st.metric("Total Keluar", fmt_rp(total_keluar))
        with col_r3:
            st.metric("Saldo Akhir", fmt_rp(saldo))

def show_charts():
    st.subheader("📊 Grafik")
    
    if not st.session_state.data:
        st.info("Belum ada data")
        return
    
    df = pd.DataFrame(st.session_state.data)
    df['tanggal_dt'] = df['tanggal'].apply(parse_tgl)
    df['bulan'] = df['tanggal_dt'].apply(lambda x: f"{BULAN_ID[x.month-1]}" if x else None)
    df = df.dropna(subset=['bulan'])
    
    monthly = df.groupby('bulan').agg({'masuk': 'sum', 'keluar': 'sum'}).reset_index()
    
    if not monthly.empty:
        import plotly.graph_objects as go
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Masuk', x=monthly['bulan'], y=monthly['masuk'], marker_color='green'))
        fig.add_trace(go.Bar(name='Keluar', x=monthly['bulan'], y=monthly['keluar'], marker_color='red'))
        fig.update_layout(height=400, margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)

def show_settings():
    st.subheader("⚙️ Pengaturan")
    
    with st.form("settings"):
        nama = st.text_input("Nama Kelompok", value=st.session_state.config.get("nama_kelompok", "Kelompok Narogong"))
        gereja = st.text_input("Nama Gereja", value=st.session_state.config.get("nama_gereja", "GKJ"))
        
        uploaded_logo = st.file_uploader("Upload Logo", type=['png', 'jpg', 'jpeg'])
        if uploaded_logo:
            st.image(uploaded_logo, width=80)
        
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            st.session_state.config["nama_kelompok"] = nama
            st.session_state.config["nama_gereja"] = gereja
            if uploaded_logo:
                with open("logo.png", "wb") as f:
                    f.write(uploaded_logo.getbuffer())
                st.session_state.config["logo_path"] = "logo.png"
            save_config(st.session_state.config)
            st.success("Tersimpan!")
            st.rerun()

# ==================== INISIALISASI ====================

if 'data' not in st.session_state:
    st.session_state.data = load_data()
if 'config' not in st.session_state:
    st.session_state.config = load_config()

# ==================== MAIN ====================

def main():
    st.markdown(f"# 💰 {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')}")
    st.caption(f"{st.session_state.config.get('nama_gereja', 'GKJ')}")
    
    menu = st.selectbox(
        "Menu",
        ["🏠 Dashboard", "➕ Input", "📋 Data", "📊 Grafik", "⚙️ Setting"]
    )
    
    if menu == "🏠 Dashboard":
        show_dashboard()
        st.markdown("---")
        show_charts()
    elif menu == "➕ Input":
        show_transaksi_form()
    elif menu == "📋 Data":
        show_data_table()
    elif menu == "📊 Grafik":
        show_charts()
    elif menu == "⚙️ Setting":
        show_settings()
    
    st.markdown("---")
    st.caption(f"Total: {len(st.session_state.data)} transaksi")

if __name__ == "__main__":
    main()