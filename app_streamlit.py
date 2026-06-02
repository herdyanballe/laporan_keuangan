"""
Aplikasi Laporan Keuangan Kas Kelompok Narogong - Versi Final
Fitur: Dashboard interaktif, Edit/Hapus Transaksi, Backup Excel, Export PDF
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
import plotly.express as px
import plotly.graph_objects as go
import base64

# ==================== KONFIGURASI HALAMAN ====================
st.set_page_config(
    page_title="Kas Narogong",
    page_icon="⛪",
    layout="wide",
    initial_sidebar_state="auto"
)

# Custom CSS untuk tampilan premium
st.markdown("""
<style>
    .main { background: linear-gradient(135deg, #f5f7fa 0%, #e9ecef 100%); }
    .metric-card {
        background: white;
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        transition: transform 0.3s ease;
    }
    .metric-card:hover { transform: translateY(-5px); }
    .metric-title { font-size: 14px; color: #6c757d; letter-spacing: 1px; margin-bottom: 10px; }
    .metric-value { font-size: 32px; font-weight: bold; color: #1F4E79; }
    .metric-change { font-size: 12px; margin-top: 8px; }
    .positive { color: #28a745; }
    .negative { color: #dc3545; }
    .main-header {
        background: linear-gradient(135deg, #1F4E79 0%, #2E6DA4 100%);
        border-radius: 20px;
        padding: 25px;
        margin-bottom: 30px;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 28px; }
    .main-header p { margin: 10px 0 0 0; opacity: 0.9; }
    .stButton > button { border-radius: 12px; font-weight: 500; transition: all 0.3s ease; }
    .stButton > button:hover { transform: scale(1.02); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    hr { margin: 30px 0; border: none; height: 1px; background: linear-gradient(90deg, transparent, #ccc, transparent); }
    @media (max-width: 768px) {
        .metric-value { font-size: 24px; }
        .main-header h1 { font-size: 22px; }
    }
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
    """Menyimpan data ke file JSON - LANGSUNG UPDATE"""
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
        print(f"✅ Data saved: {len(data)} transactions")
        return True
    except Exception as e:
        print(f"❌ Error saving data: {e}")
        return False

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
    if not tgl_akhir:
        return []
    result = []
    for r in data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl and tgl < tgl_akhir:
            result.append(r)
    return result

def build_periode_label(rows, tgl_awal=None, tgl_akhir=None):
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
    
    # Header
    story.append(Paragraph(f"LAPORAN KAS {nama_kel.upper()}", style_title))
    story.append(Paragraph(judul_periode, style_sub))
    story.append(Paragraph("GKJ BEKASI", style_sub))
    story.append(HRFlowable(width="100%", thickness=2, color=navy, spaceAfter=8))
    
    # Table data
    table_data = [["No", "Tgl", "Keterangan", "Masuk", "Keluar", "Saldo"]]
    
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
    
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    tgl_cetak = datetime.now().strftime("%d %B %Y %H:%M")
    story.append(Paragraph(f"Dicetak: {tgl_cetak}", style_sub))
    
    doc.build(story)
    
    with open(temp_file.name, "rb") as f:
        pdf_data = f.read()
    os.unlink(temp_file.name)
    return pdf_data

# ==================== KOMPONEN DASHBOARD ====================

def show_metric_cards():
    today = date.today()
    saldo_hari_ini = get_saldo_hari_ini(st.session_state.data)
    total_masuk = sum(r.get("masuk", 0) for r in st.session_state.data)
    total_keluar = sum(r.get("keluar", 0) for r in st.session_state.data)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📅 PERIODE</div>
            <div class="metric-value">{today.strftime('%B %Y')}</div>
            <div class="metric-change">{today.strftime('%d %B %Y')}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">💰 SALDO HARI INI</div>
            <div class="metric-value">{fmt_rp(saldo_hari_ini)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📈 TOTAL PEMASUKAN</div>
            <div class="metric-value">{fmt_rp(total_masuk)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📉 TOTAL PENGELUARAN</div>
            <div class="metric-value">{fmt_rp(total_keluar)}</div>
        </div>
        """, unsafe_allow_html=True)

def show_trend_chart():
    st.markdown("### 📈 Trend Pemasukan vs Pengeluaran")
    
    monthly_data = {}
    for r in st.session_state.data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl:
            key = f"{BULAN_ID[tgl.month-1]} {tgl.year}"
            if key not in monthly_data:
                monthly_data[key] = {"masuk": 0, "keluar": 0, "bulan_num": tgl.month, "tahun": tgl.year}
            monthly_data[key]["masuk"] += r.get("masuk", 0)
            monthly_data[key]["keluar"] += r.get("keluar", 0)
    
    if monthly_data:
        sorted_items = sorted(monthly_data.items(), key=lambda x: (x[1]["tahun"], x[1]["bulan_num"]))
        months = [item[0] for item in sorted_items]
        masuk_values = [item[1]["masuk"] for item in sorted_items]
        keluar_values = [item[1]["keluar"] for item in sorted_items]
        
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Pemasukan', x=months, y=masuk_values, marker_color='#28a745'))
        fig.add_trace(go.Bar(name='Pengeluaran', x=months, y=keluar_values, marker_color='#dc3545'))
        fig.update_layout(barmode='group', height=450, template='plotly_white')
        fig.update_yaxes(tickformat=',.0f')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Belum ada data")

def show_category_breakdown():
    st.markdown("### 📊 Kategori Transaksi")
    
    categories = {
        "Persembahan": ["persembahan", "iuran", "sumbangan", "dana"],
        "Kesehatan": ["subsidi kesehatan", "pk orang sakit", "pk kelahiran", "pk sripahan"],
        "Operasional": ["fc", "foto copy", "liturgi", "ibadah", "sarasehan"],
        "Sosial": ["duka", "bantuan"],
        "Lainnya": []
    }
    
    category_data = {cat: {"masuk": 0, "keluar": 0} for cat in categories.keys()}
    
    for r in st.session_state.data:
        ket = r.get("keterangan", "").lower()
        masuk = r.get("masuk", 0)
        keluar = r.get("keluar", 0)
        
        categorized = False
        for cat, keywords in categories.items():
            if cat != "Lainnya":
                for kw in keywords:
                    if kw in ket:
                        category_data[cat]["masuk"] += masuk
                        category_data[cat]["keluar"] += keluar
                        categorized = True
                        break
                if categorized:
                    break
        
        if not categorized and (masuk > 0 or keluar > 0):
            category_data["Lainnya"]["masuk"] += masuk
            category_data["Lainnya"]["keluar"] += keluar
    
    col1, col2 = st.columns(2)
    
    with col1:
        masuk_data = {k: v["masuk"] for k, v in category_data.items() if v["masuk"] > 0}
        if masuk_data:
            fig_pie = px.pie(values=list(masuk_data.values()), names=list(masuk_data.keys()), 
                            title="Pemasukan per Kategori", hole=0.4)
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        keluar_data = {k: v["keluar"] for k, v in category_data.items() if v["keluar"] > 0}
        if keluar_data:
            fig_pie = px.pie(values=list(keluar_data.values()), names=list(keluar_data.keys()), 
                            title="Pengeluaran per Kategori", hole=0.4)
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)

# ==================== FORM INPUT ====================

def show_transaksi_form():
    with st.expander("➕ Tambah Transaksi Baru", expanded=False):
        tgl = st.date_input("📅 Tanggal", value=date.today(), format="DD/MM/YYYY")
        keterangan = st.text_input("📝 Keterangan", placeholder="Contoh: Iuran bulanan")
        col1, col2 = st.columns(2)
        with col1:
            masuk = st.number_input("💰 Kas Masuk", min_value=0, value=0, step=10000)
        with col2:
            keluar = st.number_input("💸 Kas Keluar", min_value=0, value=0, step=10000)
        
        if st.button("💾 SIMPAN TRANSAKSI", type="primary", use_container_width=True):
            if not keterangan.strip():
                st.error("❌ Keterangan wajib diisi!")
            elif masuk > 0 and keluar > 0:
                st.error("❌ Hanya boleh mengisi salah satu: Kas Masuk ATAU Kas Keluar!")
            elif masuk == 0 and keluar == 0:
                st.error("❌ Harap isi nominal!")
            else:
                tgl_str = tgl.strftime("%d-%m-%Y")
                
                new_data = {
                    "tanggal": tgl_str,
                    "keterangan": keterangan.strip(),
                    "masuk": float(masuk),
                    "keluar": float(keluar)
                }
                
                st.session_state.data.append(new_data)
                if save_data(st.session_state.data):
                    st.success(f"✅ Transaksi berhasil disimpan! Total data: {len(st.session_state.data)}")
                    st.balloons()
                    st.rerun()
                else:
                    st.error("❌ Gagal menyimpan data!")

# ==================== DATA TABLE DENGAN BACKUP EXCEL ====================

def show_data_table():
    """Tampilan Data Transaksi dengan Edit, Hapus, dan Backup Excel"""
    st.subheader("📋 Daftar Transaksi")
    
    if not st.session_state.data:
        st.info("📭 Belum ada data. Silakan tambah transaksi baru.")
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
    
    if tgl_awal > tgl_akhir:
        st.error("⚠️ Tanggal 'Dari' tidak boleh lebih besar dari 'Sampai'")
        return
    
    tgl_awal_dt = datetime.combine(tgl_awal, datetime.min.time())
    tgl_akhir_dt = datetime.combine(tgl_akhir, datetime.min.time())
    
    rows = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
            tgl_awal_dt <= parse_tgl(r.get("tanggal", "")) <= tgl_akhir_dt]
    rows = sorted(rows, key=lambda r: parse_tgl(r.get("tanggal", "")))
    
    periode_text = f"{tgl_awal.strftime('%d %B %Y')} - {tgl_akhir.strftime('%d %B %Y')}"
    st.caption(f"📊 Menampilkan {len(rows)} transaksi periode: **{periode_text}**")
    
    if rows:
        # Tiga tombol berjajar: EXPORT PDF, BACKUP EXCEL, (kosong)
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
        
        with col_btn1:
            if st.button("📄 EXPORT PDF", use_container_width=True, key="btn_export"):
                with st.spinner("📑 Sedang membuat PDF..."):
                    pdf_data = export_pdf(
                        st.session_state.data, rows, st.session_state.config,
                        tgl_awal_dt, tgl_akhir_dt
                    )
                    st.download_button(
                        label="📥 DOWNLOAD PDF",
                        data=pdf_data,
                        file_name=f"laporan_kas_{tgl_awal.strftime('%Y%m%d')}_{tgl_akhir.strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        key="download_pdf"
                    )
        
        with col_btn2:
            if st.button("💾 BACKUP EXCEL", use_container_width=True, key="btn_backup"):
                with st.spinner("📑 Membuat file backup..."):
                    try:
                        # Gunakan data dari session state (yang terbaru)
                        backup_data = st.session_state.data
                        df = pd.DataFrame(backup_data)
                        df_display = pd.DataFrame()
                        df_display["No"] = range(1, len(df) + 1)
                        df_display["Tanggal"] = df["tanggal"]
                        df_display["Keterangan"] = df["keterangan"]
                        df_display["Kas Masuk"] = df["masuk"].apply(lambda x: fmt_rp(x) if x > 0 else "-")
                        df_display["Kas Keluar"] = df["keluar"].apply(lambda x: fmt_rp(x) if x > 0 else "-")
                        
                        # Simpan ke memory
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_display.to_excel(writer, sheet_name="Data Transaksi", index=False)
                        
                        excel_data = output.getvalue()
                        
                        # Download button
                        st.download_button(
                            label="📥 DOWNLOAD EXCEL",
                            data=excel_data,
                            file_name=f"backup_transaksi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_excel"
                        )
                        st.success(f"✅ Backup {len(backup_data)} transaksi siap!")
                    except Exception as e:
                        st.error(f"❌ Gagal backup: {e}")
        
        st.markdown("### 📝 Data Transaksi")
        
        # Hitung saldo awal
        before_data = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                       parse_tgl(r.get("tanggal", "")) < tgl_awal_dt]
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in before_data)
        
        # Tampilkan setiap baris dengan tombol Edit dan Hapus
        for idx, row in enumerate(rows):
            col1, col2, col3, col4, col5, col6, col7 = st.columns([0.5, 1.2, 3, 1.2, 1.2, 1, 1])
            
            saldo += row.get("masuk", 0) - row.get("keluar", 0)
            
            with col1:
                st.write(f"{idx+1}")
            with col2:
                st.write(row.get("tanggal", ""))
            with col3:
                st.write(row.get("keterangan", "")[:40])
            with col4:
                st.write(fmt_rp(row.get("masuk", 0)) if row.get("masuk", 0) else "-")
            with col5:
                st.write(fmt_rp(row.get("keluar", 0)) if row.get("keluar", 0) else "-")
            with col6:
                if st.button("✏️ Edit", key=f"edit_{idx}", use_container_width=True):
                    st.session_state.edit_index = idx
                    st.session_state.edit_data = row.copy()
                    st.session_state.edit_mode = True
                    st.rerun()
            with col7:
                if st.button("🗑️ Hapus", key=f"del_{idx}", use_container_width=True):
                    st.session_state.delete_index = idx
                    st.session_state.delete_data = row.copy()
                    st.session_state.delete_mode = True
                    st.rerun()
            
            st.divider()
        
        # Modal Edit
        if st.session_state.get('edit_mode', False):
            with st.expander("✏️ EDIT TRANSAKSI", expanded=True):
                edit_data = st.session_state.edit_data
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    tgl_edit = st.date_input("Tanggal", value=parse_tgl(edit_data.get("tanggal", "")).date(), format="DD/MM/YYYY")
                    ket_edit = st.text_input("Keterangan", value=edit_data.get("keterangan", ""))
                with col_e2:
                    masuk_edit = st.number_input("Kas Masuk", min_value=0, value=int(edit_data.get("masuk", 0)), step=10000)
                    keluar_edit = st.number_input("Kas Keluar", min_value=0, value=int(edit_data.get("keluar", 0)), step=10000)
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("💾 SIMPAN PERUBAHAN", type="primary", use_container_width=True):
                        tgl_str = tgl_edit.strftime("%d-%m-%Y")
                        # Update data
                        index = st.session_state.edit_index
                        rows_filtered = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                                        tgl_awal_dt <= parse_tgl(r.get("tanggal", "")) <= tgl_akhir_dt]
                        original_row = rows_filtered[index]
                        # Cari dan update di data utama
                        for i, item in enumerate(st.session_state.data):
                            if item == original_row:
                                st.session_state.data[i] = {
                                    "tanggal": tgl_str,
                                    "keterangan": ket_edit,
                                    "masuk": float(masuk_edit),
                                    "keluar": float(keluar_edit)
                                }
                                break
                        save_data(st.session_state.data)
                        st.success("✅ Transaksi berhasil diupdate!")
                        st.session_state.edit_mode = False
                        st.rerun()
                with col_btn2:
                    if st.button("❌ BATAL", use_container_width=True):
                        st.session_state.edit_mode = False
                        st.rerun()
        
        # Modal Hapus
        if st.session_state.get('delete_mode', False):
            st.warning("⚠️ Apakah Anda yakin ingin menghapus transaksi ini?")
            delete_data = st.session_state.delete_data
            st.info(f"📅 {delete_data.get('tanggal')} | 📝 {delete_data.get('keterangan')} | 💰 {fmt_rp(delete_data.get('masuk', 0)) if delete_data.get('masuk', 0) else fmt_rp(delete_data.get('keluar', 0))}")
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("✅ YA, HAPUS", type="primary", use_container_width=True):
                    index = st.session_state.delete_index
                    rows_filtered = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                                    tgl_awal_dt <= parse_tgl(r.get("tanggal", "")) <= tgl_akhir_dt]
                    original_row = rows_filtered[index]
                    st.session_state.data.remove(original_row)
                    save_data(st.session_state.data)
                    st.success("🗑️ Transaksi berhasil dihapus!")
                    st.session_state.delete_mode = False
                    st.rerun()
            with col_btn2:
                if st.button("❌ TIDAK, BATAL", use_container_width=True):
                    st.session_state.delete_mode = False
                    st.rerun()
        
        # Ringkasan
        total_masuk = sum(r.get("masuk", 0) for r in rows)
        total_keluar = sum(r.get("keluar", 0) for r in rows)
        saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in before_data)
        
        st.markdown("---")
        st.markdown("### 📊 RINGKASAN")
        
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
        st.warning(f"⚠️ Tidak ada transaksi untuk periode {periode_text}")

def show_charts():
    show_trend_chart()
    st.markdown("---")
    show_category_breakdown()

def show_settings():
    st.subheader("⚙️ Pengaturan")
    with st.form("settings"):
        nama = st.text_input("Nama Kelompok", value=st.session_state.config.get("nama_kelompok", "Kelompok Narogong"))
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            st.session_state.config["nama_kelompok"] = nama
            save_config(st.session_state.config)
            st.success("✅ Pengaturan tersimpan!")

# ==================== INISIALISASI ====================

if 'data' not in st.session_state:
    st.session_state.data = load_data()
if 'config' not in st.session_state:
    st.session_state.config = load_config()
if 'edit_mode' not in st.session_state:
    st.session_state.edit_mode = False
if 'delete_mode' not in st.session_state:
    st.session_state.delete_mode = False

# ==================== MAIN APP ====================

def main():
    st.markdown(f"""
    <div class="main-header">
        <h1>⛪ {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')}</h1>
        <p>GKJ BEKASI • Sistem Laporan Keuangan Kas • Data Real-time</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🧭 NAVIGASI")
        menu = st.radio("", ["🏠 DASHBOARD", "➕ INPUT", "📋 DATA", "📊 CHARTS", "⚙️ SETTING"], label_visibility="collapsed")
        st.markdown("---")
        st.caption(f"📊 Total Transaksi: **{len(st.session_state.data)}**")
    
    if menu == "🏠 DASHBOARD":
        show_metric_cards()
        st.markdown("---")
        show_trend_chart()
        st.markdown("---")
        show_category_breakdown()
    elif menu == "➕ INPUT":
        show_transaksi_form()
    elif menu == "📋 DATA":
        show_data_table()
    elif menu == "📊 CHARTS":
        show_charts()
    elif menu == "⚙️ SETTING":
        show_settings()
    
    st.markdown("---")
    st.caption(f"© {datetime.now().year} {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')} • GKJ BEKASI")

if __name__ == "__main__":
    main()
