"""
Aplikasi Laporan Keuangan Kas Kelompok Narogong - Versi Web
Fitur: Input transaksi, Export PDF dengan logo, Mobile Friendly, Filter Rentang Tanggal
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

# ==================== KONFIGURASI HALAMAN ====================
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

# ==================== FUNGSI DASAR ====================

def fmt_rp(val):
    try:
        return f"Rp {int(float(val)):,}".replace(",",".")
    except:
        return "Rp 0"

def parse_tgl(tgl_str):
    if not tgl_str:
        return None
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

def filter_data_by_date(data, tgl_awal, tgl_akhir):
    """Filter data berdasarkan rentang tanggal"""
    if not tgl_awal and not tgl_akhir:
        return data
    result = []
    for r in data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl:
            if tgl_awal and tgl < tgl_awal:
                continue
            if tgl_akhir and tgl > tgl_akhir:
                continue
            result.append(r)
    return sorted(result, key=lambda r: tgl_to_sort_key(r.get("tanggal", "")))

def get_data_before_date(data, tgl_akhir):
    """Ambil data sebelum tanggal tertentu"""
    if not tgl_akhir:
        return []
    result = []
    for r in data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl and tgl < tgl_akhir:
            result.append(r)
    return result

def build_periode_label(rows, tgl_awal=None, tgl_akhir=None):
    """Buat label periode untuk judul PDF"""
    if tgl_awal and tgl_akhir:
        return f"{tgl_awal.strftime('%d %B %Y')} - {tgl_akhir.strftime('%d %B %Y')}"
    if rows:
        try:
            tanggals = [parse_tgl(r.get("tanggal", "")) for r in rows if r.get("tanggal")]
            tanggals = [d for d in tanggals if d]
            if tanggals:
                tgl_awal = min(tanggals)
                tgl_akhir = max(tanggals)
                return f"{tgl_awal.strftime('%d %B %Y')} - {tgl_akhir.strftime('%d %B %Y')}"
        except:
            pass
    return "Semua Periode"

# ==================== FUNGSI EXPORT PDF ====================

def export_pdf(data, rows, config, tgl_awal=None, tgl_akhir=None):
    """Export ke PDF dengan orientasi Potrait dan logo"""
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_file.close()
    
    nama_kel = config.get("nama_kelompok", "Kelompok Narogong")
    judul_periode = build_periode_label(rows, tgl_awal, tgl_akhir)
    
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
    possible_logo_paths = [logo_path, "logo.png", "logo.jpg", "logo.jpeg"]
    logo_found = None
    for path in possible_logo_paths:
        if path and os.path.exists(path):
            logo_found = path
            break
    
    if logo_found:
        try:
            logo = RLImage(logo_found, width=2.2*cm, height=2.2*cm, kind='proportional')
            header_data = [[logo, [Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title),
                                    Paragraph(judul_periode, style_sub),
                                    Paragraph("GKJ BEKASI", style_sub)]]]
            header_tbl = Table(header_data, colWidths=[2.5*cm, 12*cm])
            header_tbl.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                            ("ALIGN", (0,0), (0,0), "CENTER"),
                                            ("ALIGN", (1,0), (1,0), "CENTER"),
                                            ("LEFTPADDING", (0,0), (-1,-1), 0),
                                            ("RIGHTPADDING", (0,0), (-1,-1), 0)]))
            story.append(header_tbl)
        except:
            story.append(Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title))
            story.append(Paragraph(judul_periode, style_sub))
            story.append(Paragraph("GKJ BEKASI", style_sub))
    else:
        story.append(Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title))
        story.append(Paragraph(judul_periode, style_sub))
        story.append(Paragraph("GKJ BEKASI", style_sub))
    
    story.append(HRFlowable(width="100%", thickness=2, color=navy, spaceAfter=8))
    
    # Table data
    table_data = [["No", "Tgl", "Keterangan", "Masuk", "Keluar", "Saldo"]]
    
    # Hitung saldo awal (sebelum periode)
    if tgl_awal:
        all_before = get_data_before_date(data, tgl_awal)
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
    else:
        saldo = 0
    
    for idx, row in enumerate(rows, 1):
        saldo += row.get("masuk", 0) - row.get("keluar", 0)
        ket = row.get("keterangan", "")
        if len(ket) > 55:
            ket = ket[:52] + "..."
        table_data.append([str(idx), row.get("tanggal", ""), Paragraph(ket, style_cell),
                          fmt_rp(row.get("masuk", 0)) if row.get("masuk", 0) else "-",
                          fmt_rp(row.get("keluar", 0)) if row.get("keluar", 0) else "-",
                          fmt_rp(saldo)])
    
    total_masuk = sum(r.get("masuk", 0) for r in rows)
    total_keluar = sum(r.get("keluar", 0) for r in rows)
    
    # Tampilkan saldo awal jika ada
    if tgl_awal:
        all_before = get_data_before_date(data, tgl_awal)
        saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
        if saldo_awal != 0:
            table_data.append(["", "", "", "", "", ""])
            table_data.append(["", "", f"SALDO AWAL: {fmt_rp(saldo_awal)}", "", "", ""])
    
    table_data.append(["", "", "TOTAL", fmt_rp(total_masuk), fmt_rp(total_keluar), fmt_rp(saldo)])
    
    col_w = [0.7*cm, 1.8*cm, 6.8*cm, 2.3*cm, 2.3*cm, 2.5*cm]
    tbl = Table(table_data, colWidths=col_w, repeatRows=1)
    n = len(table_data)
    
    tbl_style = TableStyle([
        ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,0), 8),
        ("ALIGN", (0,0), (-1,0), "CENTER"), ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("FONTNAME", (0,1), (-1,n-2), "Helvetica"), ("FONTSIZE", (0,1), (-1,n-2), 7),
        ("VALIGN", (0,1), (-1,-1), "MIDDLE"), ("ALIGN", (0,1), (1,-1), "CENTER"),
        ("ALIGN", (3,1), (5,-1), "RIGHT"),
        *[("BACKGROUND", (0,i), (-1,i), silver) for i in range(1, n-1) if i%2 == 0],
        ("BACKGROUND", (0,n-1), (-1,n-1), navy), ("TEXTCOLOR", (0,n-1), (-1,n-1), white),
        ("FONTNAME", (0,n-1), (-1,n-1), "Helvetica-Bold"), ("FONTSIZE", (0,n-1), (-1,n-1), 8),
        ("ALIGN", (2,n-1), (2,n-1), "CENTER"), ("ALIGN", (3,n-1), (5,n-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ("TOPPADDING", (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 3), ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ])
    
    tbl.setStyle(tbl_style)
    story.append(tbl)
    
    # Footer
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
    today = date.today()
    today_str = today.strftime("%d/%m/%Y")
    saldo = get_saldo_hari_ini(st.session_state.data)
    total_masuk = sum(r.get("masuk", 0) for r in st.session_state.data)
    total_keluar = sum(r.get("keluar", 0) for r in st.session_state.data)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📅 Hari Ini", today_str)
    with col2:
        st.metric("💰 Saldo", fmt_rp(saldo))
    
    col3, col4 = st.columns(2)
    with col3:
        st.metric("📈 Total Masuk", fmt_rp(total_masuk))
    with col4:
        st.metric("📉 Total Keluar", fmt_rp(total_keluar))
    
    st.markdown("---")
    st.markdown(f"### 📊 Grafik {today.year}")
    
    current_year = today.year
    year_data = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and parse_tgl(r.get("tanggal", "")).year == current_year]
    
    if year_data:
        df = pd.DataFrame(year_data)
        df['tanggal_dt'] = df['tanggal'].apply(parse_tgl)
        df['bulan_num'] = df['tanggal_dt'].apply(lambda x: x.month if x else None)
        df['bulan'] = df['bulan_num'].apply(lambda x: BULAN_ID[x-1] if x else None)
        df = df.dropna(subset=['bulan'])
        
        if not df.empty:
            bulan_order = BULAN_ID.copy()
            monthly = df.groupby('bulan').agg({'masuk': 'sum', 'keluar': 'sum'}).reset_index()
            monthly['bulan_num'] = monthly['bulan'].apply(lambda x: bulan_order.index(x) if x in bulan_order else 0)
            monthly = monthly.sort_values('bulan_num').drop('bulan_num', axis=1)
            
            if not monthly.empty:
                import plotly.graph_objects as go
                fig = go.Figure()
                fig.add_trace(go.Bar(name='Masuk', x=monthly['bulan'], y=monthly['masuk'], marker_color='#2E7D32'))
                fig.add_trace(go.Bar(name='Keluar', x=monthly['bulan'], y=monthly['keluar'], marker_color='#C62828'))
                fig.update_layout(height=350, margin=dict(l=20, r=20, t=40, b=20), template="plotly_white")
                fig.update_yaxes(tickformat=',.0f')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"Belum ada data transaksi untuk tahun {current_year}")
        else:
            st.info(f"Belum ada data untuk tahun {current_year}")
    else:
        st.info(f"Belum ada data untuk tahun {current_year}")

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
                new_data = {
                    "tanggal": tgl_str,
                    "keterangan": keterangan,
                    "masuk": float(masuk),
                    "keluar": float(keluar)
                }
                st.session_state.data.append(new_data)
                save_data(st.session_state.data)
                st.success("✅ Tersimpan!")
                st.rerun()

def show_data_table():
    st.subheader("📋 Transaksi")
    
    if not st.session_state.data:
        st.info("Belum ada data")
        return
    
    # Dapatkan tanggal min dan max dari semua data
    all_dates = [parse_tgl(r.get("tanggal", "")) for r in st.session_state.data if r.get("tanggal")]
    all_dates = [d for d in all_dates if d]
    
    if not all_dates:
        st.info("Belum ada data transaksi")
        return
    
    min_date = min(all_dates).date()
    max_date = max(all_dates).date()
    
    # Filter Rentang Tanggal
    st.markdown("### 📆 Filter Periode")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        tgl_awal = st.date_input(
            "📅 Dari Tanggal", 
            value=min_date, 
            min_value=min_date, 
            max_value=max_date, 
            key="tgl_awal"
        )
    with col_f2:
        tgl_akhir = st.date_input(
            "📅 Sampai Tanggal", 
            value=max_date, 
            min_value=min_date, 
            max_value=max_date, 
            key="tgl_akhir"
        )
    
    # Validasi tanggal
    if tgl_awal > tgl_akhir:
        st.error("⚠️ Tanggal 'Dari' tidak boleh lebih besar dari 'Sampai'")
        return
    
    # Konversi ke datetime
    tgl_awal_dt = datetime.combine(tgl_awal, datetime.min.time())
    tgl_akhir_dt = datetime.combine(tgl_akhir, datetime.min.time())
    
    # Filter data berdasarkan rentang tanggal
    rows = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
            tgl_awal_dt <= parse_tgl(r.get("tanggal", "")) <= tgl_akhir_dt]
    
    # Urutkan berdasarkan tanggal
    rows = sorted(rows, key=lambda r: parse_tgl(r.get("tanggal", "")))
    
    # Tampilkan info periode yang dipilih
    periode_text = f"{tgl_awal.strftime('%d %B %Y')} - {tgl_akhir.strftime('%d %B %Y')}"
    st.caption(f"Menampilkan data periode: **{periode_text}**")
    
    if rows:
        col_pdf, _ = st.columns([1, 3])
        with col_pdf:
            if st.button("📄 Export PDF", use_container_width=True):
                with st.spinner("Membuat PDF..."):
                    pdf_data = export_pdf(
                        st.session_state.data, rows, st.session_state.config,
                        tgl_awal_dt, tgl_akhir_dt
                    )
                    st.download_button(
                        label="📥 Download PDF",
                        data=pdf_data,
                        file_name=f"laporan_{tgl_awal.strftime('%Y%m%d')}_{tgl_akhir.strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        key="download_pdf"
                    )
        
        # Tampilkan data dalam bentuk dataframe
        df_display = []
        
        # Hitung saldo awal (sebelum periode yang dipilih)
        before_data = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                       parse_tgl(r.get("tanggal", "")) < tgl_awal_dt]
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in before_data)
        
        for i, row in enumerate(rows, 1):
            saldo += row.get("masuk", 0) - row.get("keluar", 0)
            ket = row.get("keterangan", "")
            df_display.append({
                "No": i,
                "Tanggal": row.get("tanggal", ""),
                "Keterangan": ket[:40] + ".." if len(ket) > 40 else ket,
                "Kas Masuk": fmt_rp(row.get("masuk", 0)) if row.get("masuk", 0) else "-",
                "Kas Keluar": fmt_rp(row.get("keluar", 0)) if row.get("keluar", 0) else "-",
            })
        
        st.dataframe(pd.DataFrame(df_display), use_container_width=True, height=400)
        
        # Ringkasan
        total_masuk = sum(r.get("masuk", 0) for r in rows)
        total_keluar = sum(r.get("keluar", 0) for r in rows)
        saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in before_data)
        
        st.markdown("---")
        st.markdown("### 📊 Ringkasan")
        
        if saldo_awal != 0:
            col_r1, col_r2, col_r3, col_r4 = st.columns(4)
            with col_r1:
                st.metric("💰 Saldo Awal", fmt_rp(saldo_awal))
            with col_r2:
                st.metric("📈 Total Masuk", fmt_rp(total_masuk))
            with col_r3:
                st.metric("📉 Total Keluar", fmt_rp(total_keluar))
            with col_r4:
                st.metric("🏁 Saldo Akhir", fmt_rp(saldo))
        else:
            col_r1, col_r2, col_r3 = st.columns(3)
            with col_r1:
                st.metric("📈 Total Masuk", fmt_rp(total_masuk))
            with col_r2:
                st.metric("📉 Total Keluar", fmt_rp(total_keluar))
            with col_r3:
                st.metric("🏁 Saldo Akhir", fmt_rp(saldo))
    else:
        st.warning(f"Tidak ada transaksi untuk periode {periode_text}")

def show_charts():
    st.subheader("📊 Grafik Keuangan")
    
    if not st.session_state.data:
        st.info("Belum ada data untuk ditampilkan.")
        return
    
    # Ambil daftar tahun yang tersedia
    available_years = []
    for r in st.session_state.data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl:
            year = tgl.year
            if year not in available_years:
                available_years.append(year)
    available_years.sort(reverse=True)
    
    if not available_years:
        available_years = [date.today().year]
    
    current_year = date.today().year
    if current_year not in available_years:
        current_year = available_years[0]
    
    col_year, _ = st.columns([1, 3])
    with col_year:
        selected_year = st.selectbox(
            "📅 Pilih Tahun", 
            available_years, 
            index=available_years.index(current_year) if current_year in available_years else 0,
            key="chart_year"
        )
    
    # Filter data berdasarkan tahun
    filtered_data = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and parse_tgl(r.get("tanggal", "")).year == selected_year]
    
    if not filtered_data:
        st.info(f"Tidak ada data untuk tahun {selected_year}")
        return
    
    df = pd.DataFrame(filtered_data)
    df['tanggal_dt'] = df['tanggal'].apply(parse_tgl)
    df['bulan_num'] = df['tanggal_dt'].apply(lambda x: x.month if x else None)
    df['bulan'] = df['bulan_num'].apply(lambda x: BULAN_ID[x-1] if x else None)
    df = df.dropna(subset=['bulan'])
    
    if df.empty:
        st.info(f"Tidak ada data transaksi untuk tahun {selected_year}")
        return
    
    bulan_order = BULAN_ID.copy()
    monthly = df.groupby('bulan').agg({'masuk': 'sum', 'keluar': 'sum'}).reset_index()
    monthly['bulan_num'] = monthly['bulan'].apply(lambda x: bulan_order.index(x) if x in bulan_order else 0)
    monthly = monthly.sort_values('bulan_num').drop('bulan_num', axis=1)
    
    if monthly.empty:
        st.info(f"Tidak ada data transaksi untuk tahun {selected_year}")
        return
    
    import plotly.graph_objects as go
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Pemasukan', x=monthly['bulan'], y=monthly['masuk'], marker_color='#2E7D32',
                         text=monthly['masuk'].apply(lambda x: fmt_rp(x)), textposition='outside', textfont=dict(size=9)))
    fig.add_trace(go.Bar(name='Pengeluaran', x=monthly['bulan'], y=monthly['keluar'], marker_color='#C62828',
                         text=monthly['keluar'].apply(lambda x: fmt_rp(x)), textposition='outside', textfont=dict(size=9)))
    fig.update_layout(title=f"Pemasukan vs Pengeluaran Tahun {selected_year}", barmode='group',
                      xaxis_title="Bulan", yaxis_title="Jumlah (Rp)", height=450, template="plotly_white")
    fig.update_yaxes(tickformat=',.0f')
    st.plotly_chart(fig, use_container_width=True)
    
    # Ringkasan
    total_masuk = sum(r.get("masuk", 0) for r in filtered_data)
    total_keluar = sum(r.get("keluar", 0) for r in filtered_data)
    saldo_akhir = total_masuk - total_keluar
    
    st.markdown("---")
    st.markdown(f"### 📊 Ringkasan Tahun {selected_year}")
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        st.metric("💰 Total Pemasukan", fmt_rp(total_masuk))
    with col_t2:
        st.metric("💸 Total Pengeluaran", fmt_rp(total_keluar))
    with col_t3:
        warna = "green" if saldo_akhir >= 0 else "red"
        st.markdown(f'<div style="text-align:center"><p style="font-size:14px;margin:0">🏁 Surplus/Defisit</p><p style="color:{warna};font-size:24px;font-weight:bold">{fmt_rp(saldo_akhir)}</p></div>', unsafe_allow_html=True)

def show_settings():
    st.subheader("⚙️ Pengaturan")
    
    with st.form("settings"):
        nama = st.text_input("Nama Kelompok", value=st.session_state.config.get("nama_kelompok", "Kelompok Narogong"))
        
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            st.session_state.config["nama_kelompok"] = nama
            save_config(st.session_state.config)
            st.success("Tersimpan!")
            st.rerun()

# ==================== INISIALISASI SESSION STATE ====================

if 'data' not in st.session_state:
    st.session_state.data = load_data()
if 'config' not in st.session_state:
    st.session_state.config = load_config()

# ==================== MAIN APP ====================

def main():
    st.markdown(f"# 💰 {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')}")
    st.caption("GKJ BEKASI")
    
    menu = st.selectbox(
        "Menu",
        ["🏠 Dashboard", "➕ Input", "📋 Data", "📊 Grafik", "⚙️ Setting"]
    )
    
    if menu == "🏠 Dashboard":
        show_dashboard()
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