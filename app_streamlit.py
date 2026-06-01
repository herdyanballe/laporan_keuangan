"""
Aplikasi Laporan Keuangan Kas Kelompok Narogong - Versi Web
Fitur: Input transaksi, import Excel, export Excel & PDF dengan logo
"""

import streamlit as st
import pandas as pd
import json, os, io
from datetime import datetime, date
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go

# Konfigurasi halaman
st.set_page_config(
    page_title="Laporan Keuangan Kas Narogong",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Data storage ──────────────────────────────────────────────────────────────
DATA_FILE = "data_kas.json"
CONFIG_FILE = "config.json"

BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni",
            "Juli","Agustus","September","Oktober","November","Desember"]

# ==================== FUNGSI HARUS DIDEKLARASIKAN DULU ====================

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

# ==================== INISIALISASI SESSION STATE ====================

# Inisialisasi session state setelah fungsi didefinisikan
if 'data' not in st.session_state:
    st.session_state.data = load_data()
if 'config' not in st.session_state:
    st.session_state.config = load_config()

# ==================== UI COMPONENTS ====================

def show_dashboard():
    """Tampilan Dashboard dengan Kartu Saldo"""
    col1, col2, col3, col4 = st.columns(4)
    
    today = date.today().strftime("%d %B %Y")
    saldo_hari_ini = get_saldo_hari_ini(st.session_state.data)
    
    total_masuk = sum(r.get("masuk", 0) for r in st.session_state.data)
    total_keluar = sum(r.get("keluar", 0) for r in st.session_state.data)
    
    with col1:
        st.metric("📅 Tanggal", today)
    with col2:
        st.metric("💰 Saldo Hari Ini", fmt_rp(saldo_hari_ini))
    with col3:
        st.metric("📈 Total Pemasukan", fmt_rp(total_masuk))
    with col4:
        st.metric("📉 Total Pengeluaran", fmt_rp(total_keluar))

def show_transaksi_form():
    """Form Input Transaksi"""
    with st.expander("➕ Tambah Transaksi Baru", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            tgl = st.date_input("Tanggal", value=date.today(), format="DD/MM/YYYY")
            keterangan = st.text_area("Keterangan", height=80, placeholder="Contoh: Iuran bulanan, Pembelian perlengkapan, dll")
        
        with col2:
            masuk = st.number_input("Kas Masuk (Rp)", min_value=0, value=0, step=10000)
            keluar = st.number_input("Kas Keluar (Rp)", min_value=0, value=0, step=10000)
        
        col_btn1, col_btn2, col_btn3 = st.columns([1,1,2])
        with col_btn1:
            if st.button("💾 Simpan", type="primary", use_container_width=True):
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
    """Tampilan Tabel Data"""
    st.subheader("📋 Daftar Transaksi")
    
    if not st.session_state.data:
        st.info("📭 Belum ada data. Silakan tambah transaksi baru atau import dari Excel.")
        return
    
    # Filter
    col_f1, col_f2, col_f3 = st.columns([2,2,1])
    with col_f1:
        bulan_filter = st.selectbox("Filter Bulan", ["Semua"] + BULAN_ID)
    with col_f2:
        years = sorted({parse_tgl(r["tanggal"]).year for r in st.session_state.data 
                       if r.get("tanggal") and parse_tgl(r["tanggal"])}, reverse=True)
        tahun_filter = st.selectbox("Filter Tahun", ["Semua"] + [str(y) for y in years])
    with col_f3:
        st.write("")
        st.write("")
        if st.button("🔄 Refresh", use_container_width=True):
            st.rerun()
    
    bulan = str(BULAN_ID.index(bulan_filter)+1) if bulan_filter != "Semua" else None
    tahun = tahun_filter if tahun_filter != "Semua" else None
    
    rows = filter_data(st.session_state.data, bulan, tahun)
    
    if not rows:
        st.info(f"Tidak ada data untuk periode {bulan_filter} {tahun_filter}.")
        return
    
    # Hitung saldo
    if bulan or tahun:
        all_before = get_data_before_period(st.session_state.data, bulan, tahun)
        saldo = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
    else:
        saldo = 0
    
    # Buat dataframe untuk ditampilkan
    table_data = []
    for idx, row in enumerate(rows, 1):
        saldo += row.get("masuk", 0) - row.get("keluar", 0)
        table_data.append({
            "No": idx,
            "Tanggal": row["tanggal"],
            "Keterangan": row["keterangan"],
            "Kas Masuk": fmt_rp(row.get("masuk", 0)) if row.get("masuk", 0) else "-",
            "Kas Keluar": fmt_rp(row.get("keluar", 0)) if row.get("keluar", 0) else "-",
            "Saldo": fmt_rp(saldo)
        })
    
    df = pd.DataFrame(table_data)
    st.dataframe(df, use_container_width=True, height=400)
    
    # Ringkasan
    total_masuk = sum(r.get("masuk", 0) for r in rows)
    total_keluar = sum(r.get("keluar", 0) for r in rows)
    
    if bulan or tahun:
        all_before = get_data_before_period(st.session_state.data, bulan, tahun)
        saldo_awal = sum(r.get("masuk", 0) - r.get("keluar", 0) for r in all_before)
        
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.metric("💰 Saldo Awal", fmt_rp(saldo_awal))
        with col_s2:
            st.metric("📈 Total Masuk", fmt_rp(total_masuk))
        with col_s3:
            st.metric("📉 Total Keluar", fmt_rp(total_keluar))
        with col_s4:
            st.metric("🏁 Saldo Akhir", fmt_rp(saldo))
    else:
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            st.metric("📈 Total Masuk", fmt_rp(total_masuk))
        with col_s2:
            st.metric("📉 Total Keluar", fmt_rp(total_keluar))
        with col_s3:
            st.metric("🏁 Saldo Akhir", fmt_rp(saldo))

def show_charts():
    """Tampilan Grafik"""
    st.subheader("📊 Visualisasi Data")
    
    if not st.session_state.data:
        st.info("Belum ada data untuk ditampilkan.")
        return
    
    df = pd.DataFrame(st.session_state.data)
    df['tanggal_dt'] = df['tanggal'].apply(parse_tgl)
    df['bulan'] = df['tanggal_dt'].apply(lambda x: f"{BULAN_ID[x.month-1]} {x.year}" if x else None)
    df = df.dropna(subset=['bulan'])
    
    # Grafik per bulan
    monthly = df.groupby('bulan').agg({
        'masuk': 'sum',
        'keluar': 'sum'
    }).reset_index()
    
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
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Pie chart untuk kategori (jika ada)
    st.subheader("📊 Komposisi Pengeluaran")
    
    # Ambil 10 pengeluaran terbesar untuk pie chart
    pengeluaran = df[df['keluar'] > 0].nlargest(10, 'keluar')[['keterangan', 'keluar']]
    if not pengeluaran.empty:
        fig_pie = px.pie(pengeluaran, values='keluar', names='keterangan', 
                         title="10 Pengeluaran Terbesar")
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("Belum ada data pengeluaran.")

def show_export():
    """Halaman Export"""
    st.subheader("📤 Export Laporan")
    
    if not st.session_state.data:
        st.info("Belum ada data untuk diexport.")
        return
    
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        bulan_export = st.selectbox("Bulan", ["Semua"] + BULAN_ID, key="export_bulan")
    with col_e2:
        years = sorted({parse_tgl(r["tanggal"]).year for r in st.session_state.data 
                       if r.get("tanggal") and parse_tgl(r["tanggal"])}, reverse=True)
        tahun_export = st.selectbox("Tahun", ["Semua"] + [str(y) for y in years], key="export_tahun")
    
    bulan = str(BULAN_ID.index(bulan_export)+1) if bulan_export != "Semua" else None
    tahun = tahun_export if tahun_export != "Semua" else None
    
    rows = filter_data(st.session_state.data, bulan, tahun)
    
    if rows:
        st.success(f"📊 Ditemukan **{len(rows)}** transaksi untuk periode yang dipilih.")
        
        # Preview data yang akan diexport
        with st.expander("🔍 Preview Data"):
            preview_df = pd.DataFrame(rows)
            st.dataframe(preview_df, use_container_width=True)
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            # Export ke Excel
            output = io.BytesIO()
            df_export = pd.DataFrame(rows)
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Laporan Kas')
            output.seek(0)
            
            st.download_button(
                label="📊 Download Excel",
                data=output,
                file_name=f"laporan_kas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        
        with col_btn2:
            # Export ke CSV
            csv = pd.DataFrame(rows).to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📄 Download CSV",
                data=csv,
                file_name=f"laporan_kas_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.warning("⚠️ Tidak ada data untuk periode yang dipilih.")

def show_settings():
    """Halaman Pengaturan"""
    st.subheader("⚙️ Pengaturan Aplikasi")
    
    with st.form("settings_form"):
        nama_kelompok = st.text_input("Nama Kelompok", 
                                      value=st.session_state.config.get("nama_kelompok", "Kelompok Narogong"))
        nama_gereja = st.text_input("Nama Gereja", 
                                    value=st.session_state.config.get("nama_gereja", "GKJ"))
        
        uploaded_logo = st.file_uploader("Upload Logo (PNG/JPG)", type=['png', 'jpg', 'jpeg'])
        
        if uploaded_logo:
            st.image(uploaded_logo, width=100)
        
        if st.form_submit_button("💾 Simpan Pengaturan", type="primary"):
            st.session_state.config["nama_kelompok"] = nama_kelompok
            st.session_state.config["nama_gereja"] = nama_gereja
            
            if uploaded_logo:
                # Simpan logo
                logo_path = Path("logo" + Path(uploaded_logo.name).suffix)
                with open(logo_path, "wb") as f:
                    f.write(uploaded_logo.getbuffer())
                st.session_state.config["logo_path"] = str(logo_path)
            
            save_config(st.session_state.config)
            st.success("✅ Pengaturan berhasil disimpan!")
            st.rerun()

def show_data_management():
    """Manajemen Data"""
    st.subheader("🗂️ Manajemen Data")
    
    # Import Excel
    with st.expander("📂 Import dari Excel", expanded=False):
        st.info("Format Excel harus memiliki kolom: Tanggal, Keterangan, Kas Masuk, Kas Keluar")
        uploaded_file = st.file_uploader("Pilih file Excel", type=['xlsx', 'xls'])
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.write("Preview data yang akan diimport:")
                st.dataframe(df.head(), use_container_width=True)
                
                # Cek kolom yang diperlukan
                required_cols = ['Tanggal', 'Keterangan', 'Kas Masuk', 'Kas Keluar']
                df_columns = [col for col in df.columns if any(key in str(col).lower() for key in ['tanggal', 'keterangan', 'masuk', 'keluar'])]
                
                if st.button("✅ Konfirmasi Import", type="primary"):
                    added = 0
                    skipped = 0
                    for _, row in df.iterrows():
                        # Cari kolom yang sesuai
                        tgl_val = row[[c for c in df.columns if 'tanggal' in str(c).lower()][0]] if any('tanggal' in str(c).lower() for c in df.columns) else None
                        ket_val = row[[c for c in df.columns if 'keterangan' in str(c).lower()][0]] if any('keterangan' in str(c).lower() for c in df.columns) else None
                        masuk_val = row[[c for c in df.columns if 'masuk' in str(c).lower()][0]] if any('masuk' in str(c).lower() for c in df.columns) else 0
                        keluar_val = row[[c for c in df.columns if 'keluar' in str(c).lower()][0]] if any('keluar' in str(c).lower() for c in df.columns) else 0
                        
                        if pd.notna(tgl_val) and pd.notna(ket_val):
                            tgl_str = fmt_tgl(str(tgl_val))
                            new_row = {
                                "tanggal": tgl_str,
                                "keterangan": str(ket_val),
                                "masuk": float(masuk_val) if pd.notna(masuk_val) else 0,
                                "keluar": float(keluar_val) if pd.notna(keluar_val) else 0
                            }
                            # Cek duplikat
                            exists = any(
                                r.get("tanggal") == tgl_str and 
                                r.get("keterangan") == str(ket_val)
                                for r in st.session_state.data
                            )
                            if not exists:
                                st.session_state.data.append(new_row)
                                added += 1
                            else:
                                skipped += 1
                    
                    save_data(st.session_state.data)
                    st.success(f"✅ Import selesai! {added} data baru ditambahkan, {skipped} data duplikat dilewati.")
                    st.rerun()
            except Exception as e:
                st.error(f"Error membaca file: {e}")
    
    # Hapus semua data
    with st.expander("⚠️ Hapus Semua Data", expanded=False):
        st.warning("⚠️ Tindakan ini tidak dapat dibatalkan! Semua data transaksi akan dihapus.")
        confirm = st.text_input("Ketik 'HAPUS' untuk konfirmasi:")
        if confirm == "HAPUS":
            if st.button("🗑️ Hapus Semua Data", type="secondary"):
                st.session_state.data = []
                save_data(st.session_state.data)
                st.success("✅ Semua data telah dihapus!")
                st.rerun()

# ==================== MAIN APP ====================

def main():
    # Header dengan logo jika ada
    col_title, col_logo = st.columns([4, 1])
    with col_title:
        st.title(f"💰 {st.session_state.config.get('nama_kelompok', 'Kelompok Narogong')}")
        st.caption(f"{st.session_state.config.get('nama_gereja', 'GKJ')} • Sistem Laporan Keuangan Kas")
    
    if st.session_state.config.get("logo_path") and os.path.exists(st.session_state.config["logo_path"]):
        with col_logo:
            st.image(st.session_state.config["logo_path"], width=100)
    
    st.divider()
    
    # Sidebar Navigation
    with st.sidebar:
        st.markdown(f"### 📊 Menu")
        
        menu = st.radio(
            "Pilih Menu",
            ["🏠 Dashboard", "📝 Input Transaksi", "📋 Data Transaksi", 
             "📊 Grafik", "📤 Export", "⚙️ Pengaturan", "🗂️ Manajemen Data"],
            label_visibility="collapsed"
        )
        
        st.divider()
        st.caption(f"📊 Total Transaksi: **{len(st.session_state.data)}**")
        st.caption(f"💻 Versi 2.0 • Web App")
        
        # Tombol refresh data
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.session_state.data = load_data()
            st.session_state.config = load_config()
            st.rerun()
    
    # Routing menu
    if menu == "🏠 Dashboard":
        show_dashboard()
        show_charts()
    elif menu == "📝 Input Transaksi":
        show_transaksi_form()
        show_data_table()
    elif menu == "📋 Data Transaksi":
        show_data_table()
    elif menu == "📊 Grafik":
        show_charts()
    elif menu == "📤 Export":
        show_export()
    elif menu == "⚙️ Pengaturan":
        show_settings()
    elif menu == "🗂️ Manajemen Data":
        show_data_management()

if __name__ == "__main__":
    main()