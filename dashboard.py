import streamlit as st
import pandas as pd
from datetime import date, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json  # Tambahan untuk proteksi format data

# Konfigurasi Halaman
st.set_page_config(layout="wide", page_title="Dashboard Monitoring", initial_sidebar_state="expanded")

def load_gsheet_all():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    spreadsheet_id = "10lpdIeAkhQj1Rv2tnP2V806edBnpCICbU2bsf5OslKc"
    
    try:
        # LOGIKA DETEKSI KUNCI (LAPTOP VS CLOUD)
        if os.path.exists("kunci_database.json"):
            # Jika jalan di laptop (Local)
            creds = ServiceAccountCredentials.from_json_keyfile_name("kunci_database.json", scope)
        else:
            # Jika jalan di Streamlit Cloud
            # 1. Ambil data dari Secrets dan paksa jadi Dictionary murni
            creds_info = dict(st.secrets["gcp_service_account"])
            
            # 2. PERBAIKAN KRUSIAL: Memastikan karakter \n dibaca benar oleh Google
            if "private_key" in creds_info:
                creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
            
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
            
        client = gspread.authorize(creds)
        sheet = client.open_by_key(spreadsheet_id).worksheet("ALL")
        
        data = sheet.get_all_values()
        
        # Header baris 4 (index 3), Data mulai baris 5 (index 4)
        if len(data) > 3:
            header = data[3] 
            rows = data[4:] 
            df = pd.DataFrame(rows, columns=header)
            df.columns = df.columns.str.strip()
            df = df.loc[:, df.columns != '']
            return df
        else:
            return pd.DataFrame()

    except Exception as e:
        # Menampilkan error spesifik agar mudah didebug
        st.error(f"Gagal koneksi database: {e}")
        return pd.DataFrame()

# =================================================================
# SIDEBAR NAVIGATION
# =================================================================
with st.sidebar:
    st.title("Main Menu")
    menu_pilihan = st.radio(
        "Pilih Dashboard:",
        ["NPR", "PUR", "SQ to SO", "KPI Marketing", "Laporan Weekly"],
        index=0 
    )
    st.divider()
    st.info("Klik panah '>' di pojok kiri atas untuk menutup/membuka menu samping.")

today = date.today()

# =================================================================
# LOGIKA TAMPILAN BERDASARKAN MENU SIDEBAR
# =================================================================

# --- MENU 1: NPR (EXCEL) ---
if menu_pilihan == "NPR":
    st.header("Dashboard NPR")
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
    df_sq_to_so = pd.read_excel("data_sq_to_so.xlsx")
    df_sq_baru = pd.read_excel("data_sq.xlsx")
    
    st.subheader("1. Monitoring SQ to SO")
    cust_list1 = df_sq_to_so["Customer"].dropna().unique().tolist()
    filter_cust1 = st.selectbox("Filter Customer (SQ to SO)", ["Semua"] + sorted(cust_list1), key="f_cust1")
    df1_f = df_sq_to_so if filter_cust1 == "Semua" else df_sq_to_so[df_sq_to_so["Customer"] == filter_cust1]
    
    subtotal1 = df1_f[df1_f["Status"] != "Draft"]["Total Barang"].sum()
    c1a, c1b = st.columns(2)
    c1a.metric("Total SQ to SO", len(df1_f))
    c1b.metric("Subtotal (Non-Draft)", f"Rp {subtotal1:,.0f}".replace(",", "."))
    st.dataframe(df1_f, use_container_width=True)

    st.divider()
    st.subheader("2. Monitoring Data SQ Baru")
    cf1, cf2 = st.columns(2)
    with cf1:
        cust_list2 = df_sq_baru["Customer"].dropna().unique().tolist()
        filter_cust2 = st.selectbox("Filter Customer", ["Semua"] + sorted(cust_list2), key="f_cust2")
    with cf2:
        week_list = df_sq_baru["Week"].dropna().unique().tolist()
        filter_week = st.selectbox("Filter Week", ["Semua"] + sorted(week_list), key="f_week")
    
    df2_f = df_sq_baru.copy()
    if filter_cust2 != "Semua": df2_f = df2_f[df2_f["Customer"] == filter_cust2]
    if filter_week != "Semua": df2_f = df2_f[df2_f["Week"] == filter_week]
    
    subtotal2 = df2_f[df2_f["Status"] != "Draft"]["Sub Total"].sum()
    c2a, c2b = st.columns(2)
    c2a.metric("Total SQ Baru", len(df2_f))
    c2b.metric("Subtotal (Non-Draft)", f"Rp {subtotal2:,.0f}".replace(",", "."))
    st.dataframe(df2_f, use_container_width=True)

# --- MENU 4: KPI MARKETING (EXCEL) ---
elif menu_pilihan == "KPI Marketing":
    st.header("KPI Marketing Performance")
    df_si = pd.read_excel("data_kpi.xlsx", sheet_name="SI")
    df_sq_kpi = pd.read_excel("data_kpi.xlsx", sheet_name="SQ")
    df_si.columns = df_si.columns.str.strip()
    df_sq_kpi.columns = df_sq_kpi.columns.str.strip()

    df_si["Tanggal"] = pd.to_datetime(df_si["Tanggal"], errors="coerce")
    df_sq_kpi["Tanggal"] = pd.to_datetime(df_sq_kpi["Tanggal"], errors="coerce")

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

    val_si = df_si[mask_si & (df_si["Status"] != "Draft")]["Total Harga Jual"].sum()
    val_sq = df_sq_kpi[mask_sq & (df_sq_kpi["Status"] != "Draft")]["Sub Total"].sum()

    st.divider()
    label_bulan = "Seluruh Bulan" if sel_month == 0 else months[sel_month]
    st.subheader(f"Ringkasan KPI: {label_bulan} {sel_year}")
    
    m1, m2 = st.columns(2)
    m1.metric(f"Total SI - {sel_sales}", f"Rp {val_si:,.0f}".replace(",", "."))
    m2.metric(f"Total SQ - {sel_sales}", f"Rp {val_sq:,.0f}".replace(",", "."))

# --- MENU 5: LAPORAN WEEKLY (HYBRID SUPPORT) ---
elif menu_pilihan == "Laporan Weekly":
    st.header("Laporan Weekly - Monitoring PO Jhonlin")
    
    try:
        df_all = load_gsheet_all()
        if df_all.empty:
            st.warning("Data tidak tersedia atau gagal memuat.")
        else:
            df_all.columns = df_all.columns.str.strip()

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

            col_nom = "SUM of ON SITE"
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