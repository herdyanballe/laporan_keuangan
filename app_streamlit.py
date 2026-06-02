"""
Aplikasi Laporan Keuangan Kas Kelompok Narogong - Versi Premium
Fitur: Dashboard interaktif, Chart modern, Filter dinamis, Mobile Friendly
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
from plotly.subplots import make_subplots

# ==================== KONFIGURASI HALAMAN ====================
st.set_page_config(
    page_title="Kas Narogong",
    page_icon="💰",
    layout="wide",  # Ubah ke wide untuk tampilan lebih luas
    initial_sidebar_state="auto"
)

# Custom CSS untuk tampilan premium
st.markdown("""
<style>
    /* Global styling */
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #e9ecef 100%);
    }
    
    /* Card styling */
    .metric-card {
        background: white;
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        transition: transform 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
    }
    
    .metric-title {
        font-size: 14px;
        color: #6c757d;
        letter-spacing: 1px;
        margin-bottom: 10px;
    }
    
    .metric-value {
        font-size: 32px;
        font-weight: bold;
        color: #1F4E79;
    }
    
    .metric-change {
        font-size: 12px;
        margin-top: 8px;
    }
    
    .positive {
        color: #28a745;
    }
    
    .negative {
        color: #dc3545;
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, #1F4E79 0%, #2E6DA4 100%);
        border-radius: 20px;
        padding: 25px;
        margin-bottom: 30px;
        color: white;
    }
    
    .main-header h1 {
        margin: 0;
        font-size: 28px;
    }
    
    .main-header p {
        margin: 10px 0 0 0;
        opacity: 0.9;
    }
    
    /* Table styling */
    .dataframe {
        border-radius: 15px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    
    /* Button styling */
    .stButton > button {
        border-radius: 12px;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: scale(1.02);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    
    /* Divider styling */
    hr {
        margin: 30px 0;
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #ccc, transparent);
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .metric-value {
            font-size: 24px;
        }
        .main-header h1 {
            font-size: 22px;
        }
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

# ==================== KOMPONEN DASHBOARD PREMIUM ====================

def show_metric_cards():
    """Menampilkan metric cards seperti dashboard profesional"""
    
    today = date.today()
    saldo_hari_ini = get_saldo_hari_ini(st.session_state.data)
    total_masuk = sum(r.get("masuk", 0) for r in st.session_state.data)
    total_keluar = sum(r.get("keluar", 0) for r in st.session_state.data)
    
    # Hitung persentase perubahan (dari bulan lalu)
    current_month = today.month
    current_year = today.year
    last_month_data = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                       parse_tgl(r.get("tanggal", "")).month == current_month - 1 and
                       parse_tgl(r.get("tanggal", "")).year == current_year]
    last_month_masuk = sum(r.get("masuk", 0) for r in last_month_data)
    
    perubahan = ((total_masuk - last_month_masuk) / last_month_masuk * 100) if last_month_masuk > 0 else 0
    
    # Row 1: 4 cards utama
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
            <div class="metric-change {'positive' if saldo_hari_ini >= 0 else 'negative'}">
                {'▲' if saldo_hari_ini >= 0 else '▼'} Saldo {fmt_rp(abs(saldo_hari_ini))}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📈 TOTAL PEMASUKAN</div>
            <div class="metric-value">{fmt_rp(total_masuk)}</div>
            <div class="metric-change {'positive' if perubahan >= 0 else 'negative'}">
                {'▲' if perubahan >= 0 else '▼'} {abs(perubahan):.1f}% dari bulan lalu
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📉 TOTAL PENGELUARAN</div>
            <div class="metric-value">{fmt_rp(total_keluar)}</div>
            <div class="metric-change">Rasio: {(total_keluar/total_masuk*100):.1f}% dari pemasukan</div>
        </div>
        """, unsafe_allow_html=True)

def show_trend_chart():
    """Menampilkan grafik trend pemasukan vs pengeluaran"""
    
    st.markdown("### 📈 How are we trending?")
    st.caption("Track revenue and profit over time with smart aggregation")
    
    # Siapkan data per bulan
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
        # Urutkan berdasarkan tahun dan bulan
        sorted_items = sorted(monthly_data.items(), key=lambda x: (x[1]["tahun"], x[1]["bulan_num"]))
        months = [item[0] for item in sorted_items]
        masuk_values = [item[1]["masuk"] for item in sorted_items]
        keluar_values = [item[1]["keluar"] for item in sorted_items]
        
        # Buat chart dengan plotly
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            name='Pemasukan',
            x=months,
            y=masuk_values,
            marker_color='#28a745',
            marker_line_color='#1f8b4c',
            marker_line_width=1,
            opacity=0.9,
            text=[fmt_rp(v) for v in masuk_values],
            textposition='outside',
            textfont=dict(size=10)
        ))
        
        fig.add_trace(go.Bar(
            name='Pengeluaran',
            x=months,
            y=keluar_values,
            marker_color='#dc3545',
            marker_line_color='#b02a37',
            marker_line_width=1,
            opacity=0.9,
            text=[fmt_rp(v) for v in keluar_values],
            textposition='outside',
            textfont=dict(size=10)
        ))
        
        fig.update_layout(
            barmode='group',
            height=450,
            template='plotly_white',
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            margin=dict(l=40, r=40, t=60, b=40),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
        
        fig.update_yaxes(
            title="Jumlah (Rp)",
            tickformat=',.0f',
            gridcolor='#e9ecef',
            gridwidth=1
        )
        
        fig.update_xaxes(
            title="Bulan",
            gridcolor='#e9ecef',
            gridwidth=1
        )
        
        st.plotly_chart(fig, use_container_width=True)

def show_category_breakdown():
    """Menampilkan breakdown per kategori"""
    
    st.markdown("### 📊 Where & What?")
    st.caption("Break down performance by category to see where revenue and margins concentrate")
    
    # Kumpulkan data per kategori (menggunakan kata kunci dari keterangan)
    categories = {
        "Persembahan": ["persembahan", "iuran", "sumbangan", "dana"],
        "Kesehatan": ["subsidi kesehatan", "pk orang sakit", "pk kelahiran", "pk sripahan"],
        "Operasional": ["fc", "foto copy", "liturgi", "ibadah", "sarasehan"],
        "Sosial": ["duka", "sumbangan", "bantuan"],
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
    
    # Buat 2 kolom untuk chart
    col1, col2 = st.columns(2)
    
    with col1:
        # Pie chart pemasukan per kategori
        masuk_data = {k: v["masuk"] for k, v in category_data.items() if v["masuk"] > 0}
        if masuk_data:
            fig_pie = px.pie(
                values=list(masuk_data.values()),
                names=list(masuk_data.keys()),
                title="Pemasukan per Kategori",
                color_discrete_sequence=px.colors.qualitative.Set2,
                hole=0.4
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=400, template='plotly_white')
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Belum ada data pemasukan")
    
    with col2:
        # Pie chart pengeluaran per kategori
        keluar_data = {k: v["keluar"] for k, v in category_data.items() if v["keluar"] > 0}
        if keluar_data:
            fig_pie = px.pie(
                values=list(keluar_data.values()),
                names=list(keluar_data.keys()),
                title="Pengeluaran per Kategori",
                color_discrete_sequence=px.colors.qualitative.Set3,
                hole=0.4
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=400, template='plotly_white')
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Belum ada data pengeluaran")

@st.fragment
def show_interactive_chart():
    """Grafik interaktif dengan fragment untuk performa lebih baik"""
    
    st.markdown("### 🎯 Deep Dive - Click to Explore")
    st.caption("Click on data points to see detailed trends")
    
    # Siapkan data per bulan
    monthly_summary = {}
    for r in st.session_state.data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl:
            key = f"{tgl.year}-{tgl.month:02d}"
            if key not in monthly_summary:
                monthly_summary[key] = {
                    "bulan": BULAN_ID[tgl.month-1],
                    "tahun": tgl.year,
                    "masuk": 0,
                    "keluar": 0,
                    "profit": 0
                }
            monthly_summary[key]["masuk"] += r.get("masuk", 0)
            monthly_summary[key]["keluar"] += r.get("keluar", 0)
            monthly_summary[key]["profit"] = monthly_summary[key]["masuk"] - monthly_summary[key]["keluar"]
    
    if monthly_summary:
        df = pd.DataFrame([
            {
                "Bulan": v["bulan"],
                "Tahun": v["tahun"],
                "Pemasukan": v["masuk"],
                "Pengeluaran": v["keluar"],
                "Profit": v["profit"]
            }
            for v in monthly_summary.values()
        ])
        
        # Bubble chart
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=df["Pemasukan"],
            y=df["Pengeluaran"],
            mode='markers+text',
            marker=dict(
                size=df["Profit"].abs() / 100000,
                sizeref=2.*max(df["Profit"].abs()) / (50**2),
                sizemin=10,
                color=df["Profit"],
                colorscale='RdYlGn',
                showscale=True,
                colorbar=dict(title="Profit (Rp)")
            ),
            text=df["Bulan"],
            textposition="top center",
            hoverinfo='text',
            hovertext=[
                f"{row['Bulan']} {row['Tahun']}<br>"
                f"Pemasukan: {fmt_rp(row['Pemasukan'])}<br>"
                f"Pengeluaran: {fmt_rp(row['Pengeluaran'])}<br>"
                f"Profit: {fmt_rp(row['Profit'])}"
                for _, row in df.iterrows()
            ]
        ))
        
        fig.update_layout(
            title="Profit Analysis: Size = Profit Margin",
            xaxis_title="Pemasukan (Rp)",
            yaxis_title="Pengeluaran (Rp)",
            height=500,
            template='plotly_white',
            hovermode='closest'
        )
        
        fig.update_xaxis(tickformat=',.0f')
        fig.update_yaxis(tickformat=',.0f')
        
        # Event click
        selected = st.plotly_chart(fig, use_container_width=True, key="bubble_chart", on_select="rerun")
        
        # Tampilkan detail bulan yang dipilih
        if selected and selected.get('selection'):
            st.info("👆 Klik salah satu bubble untuk melihat detail trend bulan tersebut")
    else:
        st.info("Belum ada data untuk ditampilkan")

def show_operational_insights():
    """Menampilkan operational insights"""
    
    st.markdown("### 📦 Operational Insights")
    st.caption("Examine transaction patterns and distribution")
    
    # Analisis transaksi per hari dalam seminggu
    day_names = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    day_data = {day: {"count": 0, "total": 0} for day in day_names}
    
    for r in st.session_state.data:
        tgl = parse_tgl(r.get("tanggal", ""))
        if tgl:
            day_idx = tgl.weekday()
            day_name = day_names[day_idx]
            day_data[day_name]["count"] += 1
            day_data[day_name]["total"] += r.get("masuk", 0) + r.get("keluar", 0)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Bar chart frekuensi transaksi per hari
        fig_bar = go.Figure(data=[
            go.Bar(
                x=day_names,
                y=[day_data[day]["count"] for day in day_names],
                marker_color='#1F4E79',
                text=[day_data[day]["count"] for day in day_names],
                textposition='auto'
            )
        ])
        fig_bar.update_layout(
            title="Frekuensi Transaksi per Hari",
            xaxis_title="Hari",
            yaxis_title="Jumlah Transaksi",
            height=350,
            template='plotly_white'
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    
    with col2:
        # Ringkasan distribusi
        total_transaksi = sum(day_data[day]["count"] for day in day_names)
        if total_transaksi > 0:
            busiest_day = max(day_data.items(), key=lambda x: x[1]["count"])
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">📊 INSIGHT</div>
                <div class="metric-value">{busiest_day[0]}</div>
                <div class="metric-change">adalah hari tersibuk dengan {busiest_day[1]['count']} transaksi</div>
                <hr>
                <div>💰 Rata-rata nominal transaksi: {fmt_rp(sum(day_data[day]['total'] for day in day_names) / total_transaksi)}</div>
                <div>📝 Total transaksi: {total_transaksi}</div>
            </div>
            """, unsafe_allow_html=True)

def show_transaksi_form():
    """Form Input Transaksi dengan pencegahan double input"""
    with st.expander("➕ Tambah Transaksi Baru", expanded=False):
        tgl = st.date_input("📅 Tanggal", value=date.today(), format="DD/MM/YYYY")
        keterangan = st.text_input("📝 Keterangan", placeholder="Contoh: Iuran bulanan, Pembelian perlengkapan, dll")
        col1, col2 = st.columns(2)
        with col1:
            masuk = st.number_input("💰 Kas Masuk", min_value=0, value=0, step=10000)
        with col2:
            keluar = st.number_input("💸 Kas Keluar", min_value=0, value=0, step=10000)
        
        if st.button("💾 SIMPAN TRANSAKSI", type="primary", use_container_width=True, key="btn_simpan"):
            if not keterangan.strip():
                st.error("❌ Keterangan wajib diisi!")
            elif masuk > 0 and keluar > 0:
                st.error("❌ Hanya boleh mengisi salah satu: Kas Masuk ATAU Kas Keluar!")
            elif masuk == 0 and keluar == 0:
                st.error("❌ Harap isi nominal Kas Masuk atau Kas Keluar!")
            else:
                tgl_str = tgl.strftime("%d-%m-%Y")
                
                is_duplicate = False
                for existing in st.session_state.data:
                    if (existing.get("tanggal") == tgl_str and 
                        existing.get("keterangan") == keterangan.strip() and
                        existing.get("masuk", 0) == float(masuk) and
                        existing.get("keluar", 0) == float(keluar)):
                        is_duplicate = True
                        break
                
                if is_duplicate:
                    st.error("⚠️ Transaksi ini sudah ada! Tidak boleh double input.")
                else:
                    new_data = {
                        "tanggal": tgl_str,
                        "keterangan": keterangan.strip(),
                        "masuk": float(masuk),
                        "keluar": float(keluar)
                    }
                    st.session_state.data.append(new_data)
                    save_data(st.session_state.data)
                    st.success(f"✅ Transaksi berhasil disimpan!\n\n📅 {tgl_str}\n📝 {keterangan}\n💰 {fmt_rp(masuk if masuk > 0 else keluar)}")
                    st.balloons()
                    st.rerun()

def show_data_table():
    """Tampilan Data Transaksi dengan Edit dan Hapus"""
    st.subheader("📋 Daftar Transaksi")
    
    if not st.session_state.data:
        st.info("📭 Belum ada data. Silakan tambah transaksi baru.")
        return
    
    all_dates = [parse_tgl(r.get("tanggal", "")) for r in st.session_state.data if r.get("tanggal")]
    all_dates = [d for d in all_dates if d]
    
    if not all_dates:
        st.info("Belum ada data transaksi")
        return
    
    min_date = min(all_dates).date()
    max_date = max(all_dates).date()
    
    st.markdown("### 📆 Filter Periode")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        tgl_awal = st.date_input("📅 Dari Tanggal", value=min_date, min_value=min_date, max_value=max_date, key="tgl_awal")
    with col_f2:
        tgl_akhir = st.date_input("📅 Sampai Tanggal", value=max_date, min_value=min_date, max_value=max_date, key="tgl_akhir")
    
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
        col_pdf, col_spacer = st.columns([1, 3])
        with col_pdf:
            if st.button("📄 EXPORT PDF", use_container_width=True, key="btn_export"):
                with st.spinner("📑 Sedang membuat PDF..."):
                    pdf_data = export_pdf(st.session_state.data, rows, st.session_state.config, tgl_awal_dt, tgl_akhir_dt)
                    st.download_button(
                        label="📥 DOWNLOAD PDF",
                        data=pdf_data,
                        file_name=f"laporan_kas_{tgl_awal.strftime('%Y%m%d')}_{tgl_akhir.strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        key="download_pdf"
                    )
        
        st.markdown("### 📝 Data Transaksi")
        
        before_data = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                       parse_tgl(r.get("tanggal", "")) < tgl_awal_dt]
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in before_data)
        
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
                        index = st.session_state.edit_index
                        rows_filtered = [r for r in st.session_state.data if parse_tgl(r.get("tanggal", "")) and 
                                        tgl_awal_dt <= parse_tgl(r.get("tanggal", "")) <= tgl_akhir_dt]
                        original_row = rows_filtered[index]
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
    """Halaman Charts lengkap"""
    show_trend_chart()
    st.markdown("---")
    show_category_breakdown()
    st.markdown("---")
    show_interactive_chart()
    st.markdown("---")
    show_operational_insights()

def show_settings():
    st.subheader("⚙️ Pengaturan")
    
    with st.form("settings"):
        nama = st.text_input("Nama Kelompok", value=st.session_state.config.get("nama_kelompok", "Kelompok Narogong"))
        
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            st.session_state.config["nama_kelompok"] = nama
            save_config(st.session_state.config)
            st.success("✅ Pengaturan tersimpan!")
            st.rerun()

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
    # Header Premium
    st.markdown(f"""
    <div class="main-header">
        <h1>💰 {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')}</h1>
        <p>GKJ BEKASI • Sistem Laporan Keuangan Kas • Data Real-time</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar menu
    with st.sidebar:
        st.markdown("### 🧭 NAVIGASI")
        menu = st.radio(
            "",
            ["🏠 DASHBOARD", "➕ INPUT", "📋 DATA", "📊 CHARTS", "⚙️ SETTING"],
            label_visibility="collapsed"
        )
        st.markdown("---")
        st.caption(f"📊 Total Transaksi: **{len(st.session_state.data)}**")
        
        # Status indicator
        if st.session_state.data:
            st.caption("✅ Data tersimpan")
        else:
            st.caption("⚠️ Belum ada data")
    
    # Routing
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
    
    # Footer
    st.markdown("---")
    st.caption(f"© {datetime.now().year} {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')} • GKJ BEKASI • Laporan Keuangan Kas")

if __name__ == "__main__":
    main()