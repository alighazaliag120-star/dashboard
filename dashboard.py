import streamlit as st
import pandas as pd
from datetime import date, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json

# Konfigurasi Halaman
st.set_page_config(layout="wide", page_title="Dashboard Monitoring", initial_sidebar_state="expanded")

def load_gsheet_all():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    spreadsheet_id = "10lpdIeAkhQj1Rv2tnP2V806edBnpCICbU2bsf5OslKc"
    
    try:
        # LOGIKA DETEKSI KUNCI (LAPTOP VS CLOUD)
        if os.path.exists("kunci_data.json"):
            # JALAN DI LAPTOP (LOCALHOST) - Memakai file baru 'kunci_data.json'
            creds = ServiceAccountCredentials.from_json_keyfile_name("kunci_data.json", scope)
        else:
            # JALAN DI STREAMLIT CLOUD
            # Mengambil string JSON utuh dari Secrets
            raw_json_str = st.secrets["gcp_service_account"]["content"]
            creds_info = json.loads(raw_json_str)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
            
        client = gspread.authorize(creds)
        sheet = client.open_by_key(spreadsheet_id).worksheet("ALL")
        data = sheet.get_all_values()
        
        # Header baris 4 (index 3), Data mulai baris 5 (index 4)
        if len(data) > 3:
            df = pd.DataFrame(data[4:], columns=data[3])
            df.columns = df.columns.str.strip()
            df = df.loc[:, df.columns != '']
            return df
        return pd.DataFrame()
    except Exception as e:
        # Menampilkan error agar kita tahu di mana masalahnya (JWT Signature, Permission, dll)
        st.error(f"Gagal koneksi database: {e}")
        return pd.DataFrame()

def load_gsheet_bpv():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    spreadsheet_id = "1fjr-r_FlaAE-WOrHmoC9Ai2-Kxbafzxt1Mr5MciIGOU"
    
    try:
        if os.path.exists("kunci_data.json"):
            creds = ServiceAccountCredentials.from_json_keyfile_name("kunci_data.json", scope)
        else:
            raw_json_str = st.secrets["gcp_service_account"]["content"]
            creds_info = json.loads(raw_json_str)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
            
        client = gspread.authorize(creds)
        # Mengambil sheet pertama ("BPV")
        sheet = client.open_by_key(spreadsheet_id).get_worksheet(0)
        data = sheet.get_all_values()
        
        # Header di baris 9 (Indeks 8), Data mulai baris 10 (Indeks 9)
        if len(data) > 8:
            df = pd.DataFrame(data[9:], columns=data[8])
            
            # Bersihkan spasi dan paksa huruf besar agar cocok dengan logika filter
            df.columns = df.columns.str.strip().str.upper()
            
            # Hapus kolom yang tidak punya nama (kolom kosong di kanan)
            df = df.loc[:, df.columns != '']
            
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Gagal koneksi database BPV: {e}")
        return pd.DataFrame()

# =================================================================
# SIDEBAR NAVIGATION
# =================================================================
with st.sidebar:
    st.title("Main Menu")
    menu_pilihan = st.radio(
        "Pilih Dashboard:",
        # --- TAMBAHAN MENU BARU DI SINI ---
        ["HOME", "NPR", "PUR", "SQ to SO", "KPI Marketing", "Laporan Weekly", "Status BPV", "History Penjualan Terakhir", "Tracking Vendor"], 
        index=0 
    )
    st.divider()
    st.info("Klik panah '>' di pojok kiri atas untuk menutup/membuka menu samping.")

today = date.today()

# --- FUNGSI FORMAT TANGGAL BAHASA INDONESIA ---
hari_indo = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
bulan_indo = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", 
              "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

nama_hari = hari_indo[today.weekday()]
nama_bulan = bulan_indo[today.month]
tanggal_sekarang_str = f"{nama_hari}, {today.day} {nama_bulan} {today.year}"
# ----------------------------------------------

# =================================================================
# LOGIKA TAMPILAN BERDASARKAN MENU SIDEBAR
# =================================================================

# --- MENU UTAMA: HOME ---
if menu_pilihan == "HOME":
    st.markdown("<h1 style='text-align: center;'>DASHBOARD MONITORING</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='text-align: center; color: gray; font-size: 18px;'>📅 {tanggal_sekarang_str}</p>", unsafe_allow_html=True)
    st.divider()

    # --- QUICK ACCESS LINKS ---
    st.subheader("🔗 Quick Access Links")
    c_link1, c_link2, c_link3 = st.columns(3)
    with c_link1: st.link_button("🌐 BRP SIBIMA", "https://eas.sibima.id/", width="stretch")
    with c_link2: st.link_button("📊 Monitoring SIBIMA", "https://s.id/DashboardSibima", width="stretch")
    with c_link3: st.link_button("📈 OTS ALL Marketing", "https://s.id/MonitoringOTSAll", width="stretch")

    c_link4, c_link5, c_link6 = st.columns(3)
    with c_link4: st.link_button("📝 ALL BPV SIBIMA", "https://docs.google.com/spreadsheets/d/1fjr-r_FlaAE-WOrHmoC9Ai2-Kxbafzxt1Mr5MciIGOU/edit?gid=0#gid=0", width="stretch")
    with c_link5: st.link_button("📨 OTS ALL RFQ", "https://s.id/RFQSIBIMA", width="stretch")
    with c_link6: st.link_button("📧 Hostinger", "https://mail.hostinger.com", width="stretch")

    st.divider()
    
    # --- SECTION CHART ---
    st.subheader("📊 Business Overview")

    # 1. DATA BPV (Pie Chart - Persentase Bayar vs Belum)
    try:
        df_bpv_home = load_gsheet_bpv()
        if not df_bpv_home.empty:
            st.write("**💰 Persentase Status Pembayaran BPV**")
            
            # Pastikan kolom dibersihkan lagi di sini untuk jaga-jaga
            df_bpv_home.columns = df_bpv_home.columns.str.strip().str.upper()
            
            # Gunakan nama kolom yang tepat sesuai hasil load_gsheet_bpv
            target_col = 'TANGGAL BAYAR' 
            
            if target_col in df_bpv_home.columns:
                # Logika: Cek apakah kolom kosong, berisi '-', atau 'nan'
                def cek_status(val):
                    v = str(val).strip().lower()
                    if v in ['', '-', 'nan', 'none', 'null', '0']:
                        return 'Belum Dibayar'
                    return 'Dibayar'

                df_bpv_home['Status Real'] = df_bpv_home[target_col].apply(cek_status)
                status_pie = df_bpv_home['Status Real'].value_counts()
                
                # Gunakan st.plotly_chart agar lebih cantik
                import plotly.express as px
                fig = px.pie(
                    values=status_pie.values, 
                    names=status_pie.index, 
                    hole=0.4,
                    color=status_pie.index,
                    color_discrete_map={'Dibayar':'#2ca02c', 'Belum Dibayar':'#d62728'}
                )
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=300)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning(f"Kolom {target_col} tidak ditemukan untuk grafik.")
        else:
            st.info("Data BPV kosong di Spreadsheet.")
    except Exception as e:
        # Tampilkan error aslinya agar kita tahu masalahnya apa
        st.error(f"Gagal memuat chart BPV: {e}")

    # --- FILTER UNTUK NOMOR 2 & 3 ---
    st.write("⚙️ **Filter Grafik KPI & Trend**")
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        # Filter Tahun
        sel_year_home = st.selectbox("Pilih Tahun Dashboard", [2026, 2025], key="h_year")
    with f_col2:
        # Filter Customer untuk Trend PO
        df_all_raw = load_gsheet_all()
        list_cust_home = ["Semua"] + sorted(df_all_raw["Customer"].unique().tolist()) if not df_all_raw.empty else ["Semua"]
        sel_cust_home = st.selectbox("Filter Customer (Trend PO)", list_cust_home, key="h_cust")

    # 2. KPI MARKETING (SI vs SQ Terbanyak)
    # Sumber: Excel KPI
    st.write(f"**🏆 Performa Salesman - {sel_year_home}**")
    try:
        df_si_home = pd.read_excel("data_kpi.xlsx", sheet_name="SI")
        df_sq_home = pd.read_excel("data_kpi.xlsx", sheet_name="SQ")
        
        # Filter Tahun
        df_si_home['Tanggal'] = pd.to_datetime(df_si_home['Tanggal'], errors='coerce')
        df_sq_home['Tanggal'] = pd.to_datetime(df_sq_home['Tanggal'], errors='coerce')
        df_si_f = df_si_home[df_si_home['Tanggal'].dt.year == sel_year_home]
        df_sq_f = df_sq_home[df_sq_home['Tanggal'].dt.year == sel_year_home]

        kpi1, kpi2 = st.columns(2)
        with kpi1:
            st.write("Penyumbang SI Terbanyak")
            si_chart = df_si_f.groupby('Salesman')['Total Nilai'].sum().sort_values(ascending=False).head(5)
            st.bar_chart(si_chart)
        with kpi2:
            st.write("Penyumbang SQ Terbanyak")
            sq_chart = df_sq_f.groupby('Sales')['Sub Total'].sum().sort_values(ascending=False).head(5)
            st.bar_chart(sq_chart)
    except:
        st.error("Gagal memuat data KPI Marketing")

    # 3. TREND JUMLAH PO JHONLIN (Minggu ke Minggu)
    # Sumber: Spreadsheet All
    st.write(f"**📈 Trend Weekly PO Jhonlin - {sel_cust_home}**")
    try:
        if not df_all_raw.empty:
            df_trend = df_all_raw.copy()
            # Filter Customer jika bukan 'Semua'
            if sel_cust_home != "Semua":
                df_trend = df_trend[df_trend["Customer"] == sel_cust_home]
            
            # Hitung jumlah PO per Week
            trend_data = df_trend.groupby('Week').size()
            st.line_chart(trend_data)
    except:
        st.info("Data Trend PO sedang diproses...")


# --- MENU 1: NPR (EXCEL) ---
elif menu_pilihan == "NPR":
    st.header("Dashboard NPR")
    st.caption(f"📅 {tanggal_sekarang_str}") # Memunculkan tanggal
    
    df_npr = pd.read_excel("data_npr.xlsx")
    df_npr.columns = df_npr.columns.str.strip()
    df_npr["Tanggal Complete"] = pd.to_datetime(df_npr["Tanggal Complete"], errors="coerce")
    
    col_f_npr1, col_f_npr2 = st.columns(2)
    with col_f_npr1:
        mode_npr = st.radio("Mode Tampilan NPR:", ["Selesai Hari Ini", "Pilih Tanggal Selesai (Tempo Lalu)"], key="mode_npr")
    
    if mode_npr == "Selesai Hari Ini":
        df_npr_filtered = df_npr[(df_npr["Status"] == "Complete") & (df_npr["Tanggal Complete"].dt.date == today)]
        label_tabel_npr = "Data NPR Selesai Hari Ini"
    else:
        with col_f_npr2:
            date_range_npr = st.date_input("Pilih Rentang Tanggal Selesai:", value=(today - timedelta(days=7), today), key="date_npr")
        if len(date_range_npr) == 2:
            start_date, end_date = date_range_npr
            df_npr_filtered = df_npr[(df_npr["Status"] == "Complete") & (df_npr["Tanggal Complete"].dt.date >= start_date) & (df_npr["Tanggal Complete"].dt.date <= end_date)]
            label_tabel_npr = f"Data NPR Selesai Periode {start_date} s/d {end_date}"
        else:
            df_npr_filtered = df_npr[df_npr["Status"] == "Complete"]
            label_tabel_npr = "Silakan pilih rentang tanggal"

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Semua NPR", len(df_npr))
    c2.metric("Total Semua Complete", len(df_npr[df_npr["Status"] == "Complete"]))
    c3.metric("NPR Terfilter (Complete)", len(df_npr_filtered))
    st.divider()
    st.subheader(label_tabel_npr)
    st.dataframe(df_npr_filtered, use_container_width=True)

# --- MENU 2: PUR (EXCEL) ---
elif menu_pilihan == "PUR":
    st.header("Dashboard PUR")
    st.caption(f"📅 {tanggal_sekarang_str}") # Memunculkan tanggal
    
    df_pur = pd.read_excel("data_pur.xlsx")
    df_pur.columns = df_pur.columns.str.strip()
    df_pur["Tanggal Complete"] = pd.to_datetime(df_pur["Tanggal Complete"], errors="coerce")

    col_f_pur1, col_f_pur2 = st.columns(2)
    with col_f_pur1:
        mode_pur = st.radio("Mode Tampilan PUR:", ["Selesai Hari Ini", "Pilih Tanggal Selesai (Tempo Lalu)"], key="mode_pur")
    
    if mode_pur == "Selesai Hari Ini":
        df_pur_filtered = df_pur[(df_pur["Status"] == "Complete") & (df_pur["Tanggal Complete"].dt.date == today)]
        label_tabel_pur = "Data PUR Selesai Hari Ini"
    else:
        with col_f_pur2:
            date_range_pur = st.date_input("Pilih Rentang Tanggal Selesai:", value=(today - timedelta(days=7), today), key="date_pur")
        if len(date_range_pur) == 2:
            start_p, end_p = date_range_pur
            df_pur_filtered = df_pur[(df_pur["Status"] == "Complete") & (df_pur["Tanggal Complete"].dt.date >= start_p) & (df_pur["Tanggal Complete"].dt.date <= end_p)]
            label_tabel_pur = f"Data PUR Selesai Periode {start_p} s/d {end_p}"
        else:
            df_pur_filtered = df_pur[df_pur["Status"] == "Complete"]
            label_tabel_pur = "Silakan pilih rentang tanggal"

    cp1, cp2, cp3 = st.columns(3)
    cp1.metric("Total Semua PUR", len(df_pur))
    cp2.metric("Total Semua Complete", len(df_pur[df_pur["Status"] == "Complete"]))
    cp3.metric("PUR Terfilter (Complete)", len(df_pur_filtered))
    st.divider()
    st.subheader(label_tabel_pur)
    st.dataframe(df_pur_filtered, use_container_width=True)

# --- MENU 3: SQ TO SO (EXCEL) ---
elif menu_pilihan == "SQ to SO":
    st.header("Dashboard SQ to SO")
    st.caption(f"📅 {tanggal_sekarang_str}")
    
    # --- MEMBACA KEDUA DATABASE (PENTING!) ---
    try:
        df_sq_to_so = pd.read_excel("data_sq_to_so.xlsx")
        df_sq_baru = pd.read_excel("data_sq.xlsx") 
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        st.stop()
    
    # --- BAGIAN 1: MONITORING SQ TO SO ---
    st.subheader("1. Monitoring SQ to SO")
    cust_list1 = df_sq_to_so["Customer"].dropna().unique().tolist()
    filter_cust1 = st.selectbox("Filter Customer (SQ to SO)", ["Semua"] + sorted(cust_list1), key="f_cust1")
    df1_f = df_sq_to_so if filter_cust1 == "Semua" else df_sq_to_so[df_sq_to_so["Customer"] == filter_cust1]
    
    status_valid = ["Complete", "In Progress"]
    
    # A. KELOMPOK: SQ TO SO
    df1_to_so = df1_f[df1_f["Status"].str.strip().isin(status_valid)]
    qty_to_so = df1_to_so["No Transaksi"].nunique()
    val_to_so = df1_to_so["Total Barang"].sum()

    # B. KELOMPOK: SQ PENDING
    df1_pending = df1_f[~df1_f["Status"].str.strip().isin(status_valid)]
    qty_pending = df1_pending["No Transaksi"].nunique()
    val_pending = df1_pending["Total Barang"].sum()

    # TAMPILAN METRIK BAGIAN 1
    st.markdown("#### ✅ SQ To SO (Complete & In Progress)")
    m1, m2 = st.columns(2)
    m1.metric("Qty SQ (Unik)", f"{qty_to_so} SQ")
    m2.metric("Total Value", f"Rp {val_to_so:,.0f}".replace(",", "."))

    st.markdown("#### ⏳ SQ Pending (Draft & Others)")
    m3, m4 = st.columns(2)
    m3.metric("Qty SQ (Unik)", f"{qty_pending} SQ")
    m4.metric("Total Value", f"Rp {val_pending:,.0f}".replace(",", "."))
    
    st.divider()
    st.dataframe(df1_f, width="stretch")

   # --- BAGIAN 2: MONITORING DATA SQ BARU ---
    st.subheader("2. Monitoring Data SQ Baru")
    
    # Filter untuk Bagian 2
    cf1, cf2 = st.columns(2)
    with cf1:
        # Pastikan df_sq_baru sudah terbaca di atas
        list_cust2 = ["Semua"] + sorted(df_sq_baru["Customer"].dropna().unique().tolist())
        filter_cust2 = st.selectbox("Filter Customer", list_cust2, key="f_cust2")
    with cf2:
        list_week2 = ["Semua"] + sorted(df_sq_baru["Week"].dropna().unique().tolist())
        filter_week = st.selectbox("Filter Week", list_week2, key="f_week")
    
    df2_f = df_sq_baru.copy()
    if filter_cust2 != "Semua": 
        df2_f = df2_f[df2_f["Customer"] == filter_cust2]
    if filter_week != "Semua": 
        df2_f = df2_f[df2_f["Week"] == filter_week]
    
    # --- PERBAIKAN LOGIKA SESUAI INSTRUKSI ---
    # 1. Filter Selain Draft (Gunakan .str sebelum .lower())
    df2_non_draft = df2_f[df2_f["Status"].str.strip().str.lower() != "draft"]
    
    # 2. QTY SQ Selain Draft (Unik)
    # Kita cek dulu apakah kolom 'No Transaksi' ada, jika tidak pakai kolom indeks ke-1
    col_id_sq2 = "No Transaksi" if "No Transaksi" in df2_f.columns else df2_f.columns[1]
    qty_sq_non_draft = df2_non_draft[col_id_sq2].nunique()
    
    # 3. Total Value (Menjumlahkan semua baris/tidak unik)
    # Menggunakan kolom 'Sub Total' sesuai file data_sq.xlsx
    val_sq_non_draft = df2_non_draft["Sub Total"].sum()

    # --- TAMPILAN METRIK BAGIAN 2 ---
    st.markdown("#### 📊 Ringkasan SQ Baru (Selain Draft)")
    b1, b2 = st.columns(2)
    b1.metric("Qty SQ (Unik)", f"{qty_sq_non_draft} SQ")
    b2.metric("Total Value", f"Rp {val_sq_non_draft:,.0f}".replace(",", "."))
    
    st.divider()
    st.dataframe(df2_f, use_container_width=True)

# --- MENU 4: KPI MARKETING (EXCEL) ---
elif menu_pilihan == "KPI Marketing":
    st.header("KPI Marketing Performance")
    st.caption(f"📅 {tanggal_sekarang_str}") # Memunculkan tanggal
    
    df_si = pd.read_excel("data_kpi.xlsx", sheet_name="SI")
    df_sq_kpi = pd.read_excel("data_kpi.xlsx", sheet_name="SQ")
    df_si.columns = df_si.columns.str.strip()
    df_sq_kpi.columns = df_sq_kpi.columns.str.strip()

    df_si["Tanggal"] = pd.to_datetime(df_si["Tanggal"], errors="coerce")
    df_sq_kpi["Tanggal"] = pd.to_datetime(df_sq_kpi["Tanggal"], errors="coerce")

    # ==========================================
    # PEMBERSIHAN DATA (MENCEGAH ERROR NOMINAL)
    # ==========================================
    if "Status" in df_si.columns:
        df_si["Status"] = df_si["Status"].astype(str).str.strip().str.title()
    if "Status" in df_sq_kpi.columns:
        df_sq_kpi["Status"] = df_sq_kpi["Status"].astype(str).str.strip().str.title()

    if "Total Nilai" in df_si.columns:
        df_si["Total Nilai"] = pd.to_numeric(df_si["Total Nilai"], errors='coerce').fillna(0)
    
    if "Sub Total" in df_sq_kpi.columns:
        df_sq_kpi["Sub Total"] = pd.to_numeric(df_sq_kpi["Sub Total"], errors='coerce').fillna(0)
    # ==========================================

    f_col1, f_col2, f_col3 = st.columns(3)
    with f_col1:
        years = sorted(df_si["Tanggal"].dt.year.dropna().unique().astype(int).tolist(), reverse=True)
        sel_year = st.selectbox("Tahun", years, key="kpi_y")
    
    with f_col2:
        months = {0: "Semua", 1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni", 
                  7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"}
        sel_month = st.selectbox("Bulan", list(months.keys()), format_func=lambda x: months[x], key="kpi_m")
    
    with f_col3:
        list_s = sorted(list(set(df_si["Salesman"].dropna().unique().tolist() + df_sq_kpi["Sales"].dropna().unique().tolist())))
        sel_sales = st.selectbox("Salesman", ["Semua"] + list_s, key="kpi_s")

    mask_si = (df_si["Tanggal"].dt.year == sel_year)
    mask_sq = (df_sq_kpi["Tanggal"].dt.year == sel_year)

    if sel_month != 0:
        mask_si &= (df_si["Tanggal"].dt.month == sel_month)
        mask_sq &= (df_sq_kpi["Tanggal"].dt.month == sel_month)
    
    if sel_sales != "Semua":
        mask_si &= (df_si["Salesman"] == sel_sales)
        mask_sq &= (df_sq_kpi["Sales"] == sel_sales)

    # Menjumlahkan murni kolom "Total Nilai"
    val_si = df_si[mask_si & (df_si["Status"] != "Draft")]["Total Nilai"].sum()
    val_sq = df_sq_kpi[mask_sq & (df_sq_kpi["Status"] != "Draft")]["Sub Total"].sum()

    st.divider()
    label_bulan = "Seluruh Bulan" if sel_month == 0 else months[sel_month]
    st.subheader(f"Ringkasan KPI: {label_bulan} {sel_year}")
    
    m1, m2 = st.columns(2)
    m1.metric(f"Total SI - {sel_sales}", f"Rp {val_si:,.0f}".replace(",", "."))
    m2.metric(f"Total SQ - {sel_sales}", f"Rp {val_sq:,.0f}".replace(",", "."))

# --- MENU 5: LAPORAN WEEKLY ---
elif menu_pilihan == "Laporan Weekly":
    st.header("Laporan Weekly - Monitoring PO Jhonlin")
    st.caption(f"📅 {tanggal_sekarang_str}") # Memunculkan tanggal
    
    try:
        df_all = load_gsheet_all()
        if df_all.empty:
            st.warning("Data tidak tersedia atau gagal memuat.")
        else:
            df_all.columns = df_all.columns.str.strip()

            # --- FIX: Bersihkan baris kosong pada No PO agar tidak ikut terhitung ---
            if "No PO" in df_all.columns:
                df_all = df_all.dropna(subset=["No PO"])
                df_all = df_all[df_all["No PO"].astype(str).str.strip() != ""]
                df_all = df_all[~df_all["No PO"].astype(str).str.lower().isin(["nan", "none", "null"])]

            st.subheader("Filter Breakdown")
            f_bln, f_cust, f_week, f_tgl = st.columns(4)
            
            with f_bln:
                list_bulan = ["Semua"] + sorted(df_all["Bulan"].unique().astype(str).tolist())
                sel_bulan_w = st.selectbox("Pilih Bulan", list_bulan, key="w_bln")
                
            with f_cust:
                list_cust = ["Semua"] + sorted(df_all["Customer"].unique().astype(str).tolist())
                sel_cust_w = st.selectbox("Pilih Customer", list_cust, key="w_cust")

            with f_week:
                list_week = ["Semua"] + sorted(df_all["Week"].unique().astype(str).tolist())
                sel_week_w = st.selectbox("Pilih Week", list_week, key="w_week")

            with f_tgl:
                list_tgl = ["Semua"] + sorted(df_all["Tgl Terima Email"].unique().astype(str).tolist())
                sel_tgl_w = st.selectbox("Pilih Tgl Terima Email", list_tgl, key="w_tgl")

            # --- PROSES FILTER ---
            df_filtered = df_all.copy()
            if sel_bulan_w != "Semua":
                df_filtered = df_filtered[df_filtered["Bulan"].astype(str) == sel_bulan_w]
            if sel_cust_w != "Semua":
                df_filtered = df_filtered[df_filtered["Customer"].astype(str) == sel_cust_w]
            if sel_week_w != "Semua":
                df_filtered = df_filtered[df_filtered["Week"].astype(str) == sel_week_w]
            if sel_tgl_w != "Semua":
                df_filtered = df_filtered[df_filtered["Tgl Terima Email"].astype(str) == sel_tgl_w]

            # --- METRIC ---
            st.divider()
            m1, m2 = st.columns(2)
            
            col_po = "No PO"
            total_po_unik = df_filtered[col_po].nunique() if col_po in df_filtered.columns else 0

            col_nom = "Nominal PO"
            if col_nom in df_filtered.columns:
                nominal_val = pd.to_numeric(df_filtered[col_nom].astype(str).str.replace('[^0-9]', '', regex=True), errors='coerce').sum()
            else:
                nominal_val = 0

            m1.metric("Total Unit PO (Unik)", f"{total_po_unik} PO")
            m2.metric("Total Nominal PO", f"Rp {nominal_val:,.0f}".replace(",", "."))

            st.subheader(f"Detail Data: {sel_cust_w}")
            st.dataframe(df_filtered, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan teknis tampilan: {e}")

# --- MENU BARU: STATUS BPV ---
elif menu_pilihan == "Status BPV":
    st.header("🔔 Monitoring & Notifikasi BPV")
    st.caption(f"📅 {tanggal_sekarang_str}")
    
    try:
        df_bpv = load_gsheet_bpv()
        
        if df_bpv.empty:
            st.warning("Data BPV tidak tersedia atau gagal memuat.")
        else:
            # --- PROSES CLEANING TANGGAL ---
            # Pastikan nama kolom sesuai dengan hasil strip() dan upper() di fungsi load
            col_tgl_bpv = 'TANGGAL BPV'
            col_tgl_bayar = 'TANGGAL BAYAR'

            if col_tgl_bpv in df_bpv.columns and col_tgl_bayar in df_bpv.columns:
                df_bpv[col_tgl_bpv] = pd.to_datetime(df_bpv[col_tgl_bpv], format='mixed', errors='coerce')
                df_bpv[col_tgl_bayar] = pd.to_datetime(df_bpv[col_tgl_bayar], format='mixed', errors='coerce')
                
                # --- FILTER RENTANG WAKTU ---
                st.subheader("🔍 Filter Periode")
                tgl_filter = st.date_input(
                    "Pilih Rentang Tanggal:",
                    value=(today - timedelta(days=3), today), # Default melihat 3 hari terakhir
                    key="filter_tgl_bpv"
                )

                if len(tgl_filter) == 2:
                    start_date, end_date = tgl_filter
                    
                    # Logika Filter Berdasarkan Rentang yang Dipilih
                    bpv_baru = df_bpv[
                        (df_bpv[col_tgl_bpv].dt.date >= start_date) & 
                        (df_bpv[col_tgl_bpv].dt.date <= end_date)
                    ]
                    
                    bpv_dibayar = df_bpv[
                        (df_bpv[col_tgl_bayar].dt.date >= start_date) & 
                        (df_bpv[col_tgl_bayar].dt.date <= end_date)
                    ]
                    
                    # --- UI: ANGKA HIGHLIGHT ---
                    st.divider()
                    c1, c2 = st.columns(2)
                    c1.metric(f"BPV Masuk ({start_date} s/d {end_date})", f"{len(bpv_baru)} Dokumen")
                    c2.metric(f"BPV Dibayar ({start_date} s/d {end_date})", f"{len(bpv_dibayar)} Dokumen")
                    
                    # --- UI: TABEL BPV BARU ---
                    st.subheader("📥 Rincian BPV Masuk di Periode Ini")
                    if not bpv_baru.empty:
                        kolom_tampil_baru = [col for col in ['PO TRANSAKSI', 'PIC', 'TANGGAL BPV', 'CUSTOMER', 'SUPPLIER', 'TOTAL'] if col in df_bpv.columns]
                        st.dataframe(bpv_baru[kolom_tampil_baru] if kolom_tampil_baru else bpv_baru, use_container_width=True)
                    else:
                        st.info("Tidak ada BPV baru yang masuk pada rentang tanggal ini.")
                        
                    st.write("") 
                    
                    # --- UI: TABEL BPV DIBAYAR ---
                    st.subheader("💸 Rincian BPV Dibayar di Periode Ini")
                    if not bpv_dibayar.empty:
                        kolom_tampil_bayar = [col for col in ['PO TRANSAKSI', 'TANGGAL BAYAR', 'CUSTOMER', 'SUPPLIER', 'TOTAL', 'STATUS PO'] if col in df_bpv.columns]
                        st.dataframe(bpv_dibayar[kolom_tampil_bayar] if kolom_tampil_bayar else bpv_dibayar, use_container_width=True)
                    else:
                        st.info("Tidak ada rekaman BPV yang dibayar pada rentang tanggal ini.")
                else:
                    st.warning("Silakan pilih rentang tanggal (klik tanggal mulai dan tanggal akhir).")
            else:
                st.error(f"Kolom tanggal tidak ditemukan. Kolom yang tersedia: {df_bpv.columns.tolist()}")

    except Exception as e:
        st.error(f"Terjadi kesalahan teknis saat memproses data BPV: {e}")

        # =====================================================================
    # --- SCRIPT BARU DIMULAI DARI SINI (LETAKKAN DI PALING BAWAH) ---
    # =====================================================================
    
    st.markdown("---") # Garis pembatas antara fitur filter tanggal dan tracker PO
    st.subheader("🧾 Tracker Status PO & BPV")
    st.caption("Cari nomor PO untuk mengecek kelengkapan BPV dan status pembayarannya.")

    # 1. Load Data PO (Ganti nama file Excelnya sesuai dengan milik Anda)
    @st.cache_data
    def load_data_po():
        try:
            # Ganti 'database_po.xlsx' dengan file yang menyimpan data PO Anda
            df_po = pd.read_excel("database_po.xlsx") 
            df_po.columns = df_po.columns.str.strip()
            return df_po
        except Exception as e:
            return pd.DataFrame()

    df_po = load_data_po()

    if df_po.empty:
        st.error("File database PO tidak ditemukan. Pastikan file tersedia di folder.")
    else:
        # 2. Kotak Input Pencarian PO
        col_po1, col_po2 = st.columns([4, 1])
        with col_po1:
            input_po = st.text_input("Ketik Nomor PO:", placeholder="Contoh: PO-12345...", key="search_po_input")
        with col_po2:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_cek_po = st.button("Cek Status", use_container_width=True)

        # 3. Logika Pencarian
        if btn_cek_po or input_po:
            if input_po:
                # PENTING: Sesuaikan 3 nama ini dengan nama kolom di file Excel Anda!
                kolom_po = "No PO" 
                kolom_bpv = "No BPV" 
                kolom_bayar = "Status Pembayaran"
                
                if kolom_po in df_po.columns:
                    # Cari PO
                    hasil_po = df_po[df_po[kolom_po].astype(str).str.contains(input_po, case=False, na=False)]
                    
                    if not hasil_po.empty:
                        st.success(f"✅ Dokumen PO '{input_po}' ditemukan!")
                        
                        data_terpilih = hasil_po.iloc[0]
                        
                        # --- LOGIKA CEK BPV ---
                        punya_bpv = False
                        if kolom_bpv in data_terpilih:
                            val_bpv = str(data_terpilih[kolom_bpv]).strip().lower()
                            if val_bpv not in ['nan', 'none', '', 'belum ada', 'belum']:
                                punya_bpv = True
                                
                        # --- LOGIKA CEK PEMBAYARAN ---
                        sudah_bayar = False
                        if kolom_bayar in data_terpilih:
                            val_bayar = str(data_terpilih[kolom_bayar]).strip().lower()
                            if val_bayar in ['lunas', 'sudah terbayar', 'paid', 'yes', 'sudah']:
                                sudah_bayar = True

                        # 4. Tampilkan Hasil
                        st.markdown("### Hasil Pengecekan:")
                        if punya_bpv:
                            if sudah_bayar:
                                st.info("🟢 **SUDAH ADA BPV & SUDAH TERBAYAR**")
                            else:
                                st.warning("🟡 **SUDAH ADA BPV, TAPI BELUM TERBAYAR**")
                        else:
                            st.error("🔴 **BELUM ADA BPV**")
                            
                        # Tampilkan tabel detailnya
                        st.write("Detail Dokumen:")
                        st.dataframe(hasil_po, use_container_width=True)
                        
                    else:
                        st.error(f"❌ Nomor PO '{input_po}' tidak terdaftar di database.")
                else:
                    st.error(f"Sistem error: Kolom '{kolom_po}' tidak ada di file Excel Anda.")
            else:
                st.warning("👆 Silakan ketik Nomor PO terlebih dahulu.")

# --- MENU 6: HISTORY PENJUALAN TERAKHIR ---
elif menu_pilihan == "History Penjualan Terakhir":
    st.header("🔍 History Penjualan Terakhir")
    st.caption(f"📅 {tanggal_sekarang_str} | Menampilkan data transaksi dengan tanggal terupdate")

    # 1. Load Database dengan Cache agar cepat
    @st.cache_data
    def load_penjualan_terakhir():
        try:
            # Menggunakan read_excel untuk file .xlsx
            df = pd.read_excel("data_penjualan_terakhir.xlsx") 
            df.columns = df.columns.str.strip() # Bersihkan spasi di judul kolom
            
            # Pastikan kolom Tanggal berformat datetime
            if "Tanggal" in df.columns:
                df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
                
            return df
        except Exception as e:
            st.error(f"Gagal memuat file: {e}")
            return pd.DataFrame() 

    df_penjualan = load_penjualan_terakhir()

    if df_penjualan.empty:
        st.error("File 'data_penjualan_terakhir.xlsx' tidak ditemukan atau kosong.")
    else:
        # 2. Kotak Pencarian dan Tombol Search
        col_search, col_btn = st.columns([4, 1])
        with col_search:
            search_query = st.text_input("Ketik Nama Barang yang dicari:", placeholder="Contoh: Gerobak sorong...", key="search_barang_jual")
        
        with col_btn:
            st.markdown("<br>", unsafe_allow_html=True)
            tombol_search = st.button("Search", use_container_width=True)

        # 3. Logika ketika tombol diklik atau user menekan Enter
        if tombol_search or search_query:
            if search_query:
                nama_kolom_barang = "Nama Barang" 
                
                if nama_kolom_barang in df_penjualan.columns:
                    # Filter data berdasarkan input
                    df_result = df_penjualan[df_penjualan[nama_kolom_barang].astype(str).str.contains(search_query, case=False, na=False)]

                    if not df_result.empty:
                        # --- LOGIKA TANGGAL PALING UPDATE ---
                        kolom_acuan = 'Item Id' if 'Item Id' in df_result.columns else nama_kolom_barang
                        
                        # Urutkan: Terbaru di atas
                        df_terupdate = df_result.sort_values(by=[kolom_acuan, 'Tanggal'], ascending=[True, False])
                        
                        # Ambil hanya baris terbaru untuk setiap barang
                        df_terupdate = df_terupdate.drop_duplicates(subset=[kolom_acuan], keep='first')

                        st.success(f"✅ Ditemukan {len(df_terupdate)} jenis barang terupdate untuk: '{search_query}'")
                        
                        # 4. Menentukan Kolom yang Tampil (Termasuk kolom DATABASE)
                        kolom_target = [
                            "No Transaksi", 
                            "Tanggal", 
                            "Customer", 
                            "Sales", 
                            "Kode Barang", 
                            "Nama Barang", 
                            "Harga", 
                            "Database"  # Kolom yang baru Anda tambahkan
                        ]
                        
                        # Filter kolom yang benar-benar ada di file Excel
                        kolom_tampil = [col for col in kolom_target if col in df_terupdate.columns]
                        
                        # Format tanggal agar enak dibaca (YYYY-MM-DD)
                        if "Tanggal" in kolom_tampil:
                            df_temp_display = df_terupdate.copy()
                            df_temp_display["Tanggal"] = df_temp_display["Tanggal"].dt.strftime('%Y-%m-%d')
                            st.dataframe(df_temp_display[kolom_tampil], use_container_width=True)
                        else:
                            st.dataframe(df_terupdate[kolom_tampil], use_container_width=True)
                    else:
                        st.warning(f"Belum ada riwayat penjualan untuk barang '{search_query}'.")
                else:
                    st.error(f"Kolom '{nama_kolom_barang}' tidak ditemukan di file.")
            else:
                st.info("👆 Silakan ketik nama barang terlebih dahulu.")
                
# --- MENU 7: TRACKING VENDOR ---
elif menu_pilihan == "Tracking Vendor":
    st.header("🏢 Tracking Vendor")
    st.caption(f"📅 {tanggal_sekarang_str} | Sumber: data_po_sbm.xlsx")
    
    # 1. Membaca Database
    try:
        df_vendor = pd.read_excel("data_po_sbm.xlsx")
        # Membersihkan spasi tak terlihat di nama kolom
        df_vendor.columns = df_vendor.columns.str.strip() 
        
        # Merapikan format Tanggal agar tidak muncul jam 00:00:00
        if "Tanggal" in df_vendor.columns:
            df_vendor["Tanggal"] = pd.to_datetime(df_vendor["Tanggal"], errors="coerce").dt.strftime('%Y-%m-%d')
            
    except Exception as e:
        st.error(f"Gagal membaca file 'data_po_sbm.xlsx': {e}")
        st.stop()

    # 2. Persiapan Kolom Sesuai Request
    kolom_wajib = ["No Transaksi", "Tanggal", "Supplier", "Nama Barang"]
    
    # Memastikan kolom-kolom tersebut benar-benar ada di Excel agar tidak error
    kolom_ada = [col for col in kolom_wajib if col in df_vendor.columns]
    
    if len(kolom_ada) < len(kolom_wajib):
        st.warning(f"⚠️ Beberapa kolom tidak ditemukan di database. Kolom yang ada: {', '.join(df_vendor.columns)}")

    st.subheader("🔍 Pencarian Riwayat Vendor")
    
    # 3. Fitur Filter & Pencarian
    if "Supplier" in df_vendor.columns:
        col_v1, col_v2 = st.columns(2)
        
        with col_v1:
            # Dropdown untuk memilih Supplier
            list_supplier = ["Semua"] + sorted(df_vendor["Supplier"].dropna().astype(str).unique().tolist())
            pilih_supplier = st.selectbox("Pilih Supplier:", list_supplier, key="filter_supplier")
        
        with col_v2:
            # Kotak pencarian tambahan untuk Nama Barang
            cari_barang_vendor = st.text_input("Cari Nama Barang (Opsional):", placeholder="Ketik nama barang...")
        
        # 4. Eksekusi Logika Filter
        df_filter = df_vendor.copy()
        
        # Filter jika Supplier dipilih
        if pilih_supplier != "Semua":
            df_filter = df_filter[df_filter["Supplier"].astype(str) == pilih_supplier]
            
        # Filter jika kotak pencarian barang diisi
        if cari_barang_vendor and "Nama Barang" in df_filter.columns:
            df_filter = df_filter[df_filter["Nama Barang"].astype(str).str.contains(cari_barang_vendor, case=False, na=False)]
            
        st.divider()
        
        # 5. Menampilkan Hasil Tabel
        if not df_filter.empty:
            st.success(f"✅ Ditemukan {len(df_filter)} riwayat transaksi.")
            # Hanya menampilkan 4 kolom yang Anda minta
            st.dataframe(df_filter[kolom_ada], use_container_width=True)
        else:
            st.info("Tidak ada data transaksi yang sesuai dengan pencarian Anda.")
            
    else:
        st.error("Kolom 'Supplier' tidak ditemukan di dalam file Excel.")