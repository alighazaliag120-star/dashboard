import streamlit as st
import pandas as pd
from datetime import date, timedelta

# Konfigurasi Halaman
st.set_page_config(layout="wide", page_title="Dashboard Monitoring Monitoring")

st.title("Dashboard Monitoring NPR, PUR, SQ & KPI")

# Definisi Tab
tab1, tab2, tab3, tab4 = st.tabs(["NPR", "PUR", "SQ to SO", "KPI Marketing"])

today = date.today()

# =================================================================
# TAB 1: NPR (Update Filter Tanggal Selesai)
# =================================================================
with tab1:
    st.header("Dashboard NPR")
    df_npr = pd.read_excel("data_npr.xlsx")
    df_npr.columns = df_npr.columns.str.strip()
    
    # Pastikan kolom Tanggal Complete dalam format datetime
    df_npr["Tanggal Complete"] = pd.to_datetime(df_npr["Tanggal Complete"], errors="coerce")
    
    # --- LOGIKA FILTER TANGGAL BARU ---
    col_f_npr1, col_f_npr2 = st.columns(2)
    with col_f_npr1:
        mode_npr = st.radio("Mode Tampilan NPR:", ["Selesai Hari Ini", "Pilih Tanggal Selesai (Tempo Lalu)"], key="mode_npr")
    
    if mode_npr == "Selesai Hari Ini":
        df_npr_filtered = df_npr[(df_npr["Status"] == "Complete") & (df_npr["Tanggal Complete"].dt.date == today)]
        label_tabel_npr = "Data NPR Selesai Hari Ini"
    else:
        with col_f_npr2:
            # Filter rentang tanggal (default 7 hari terakhir)
            date_range_npr = st.date_input("Pilih Rentang Tanggal Selesai:", 
                                           value=(today - timedelta(days=7), today), 
                                           key="date_npr")
        if len(date_range_npr) == 2:
            start_date, end_date = date_range_npr
            df_npr_filtered = df_npr[
                (df_npr["Status"] == "Complete") & 
                (df_npr["Tanggal Complete"].dt.date >= start_date) & 
                (df_npr["Tanggal Complete"].dt.date <= end_date)
            ]
            label_tabel_npr = f"Data NPR Selesai Periode {start_date} s/d {end_date}"
        else:
            df_npr_filtered = df_npr[df_npr["Status"] == "Complete"]
            label_tabel_npr = "Silakan pilih rentang tanggal"

    # Metric
    total_all_npr = len(df_npr)
    total_complete_npr = len(df_npr[df_npr["Status"] == "Complete"])
    total_filtered_npr = len(df_npr_filtered)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Semua NPR", total_all_npr)
    c2.metric("Total Semua Complete", total_complete_npr)
    c3.metric("NPR Terfilter (Complete)", total_filtered_npr)

    st.divider()
    st.subheader(label_tabel_npr)
    if total_filtered_npr > 0:
        if mode_npr == "Selesai Hari Ini":
            st.error(f"Ada {total_filtered_npr} NPR selesai hari ini!")
        st.dataframe(df_npr_filtered, use_container_width=True)
    else:
        st.info("Tidak ada data NPR yang cocok dengan filter tanggal tersebut.")

# =================================================================
# TAB 2: PUR (Update Filter Tanggal Selesai)
# =================================================================
with tab2:
    st.header("Dashboard PUR")
    df_pur = pd.read_excel("data_pur.xlsx")
    df_pur.columns = df_pur.columns.str.strip()
    df_pur["Tanggal Complete"] = pd.to_datetime(df_pur["Tanggal Complete"], errors="coerce")

    # --- LOGIKA FILTER TANGGAL BARU ---
    col_f_pur1, col_f_pur2 = st.columns(2)
    with col_f_pur1:
        mode_pur = st.radio("Mode Tampilan PUR:", ["Selesai Hari Ini", "Pilih Tanggal Selesai (Tempo Lalu)"], key="mode_pur")
    
    if mode_pur == "Selesai Hari Ini":
        df_pur_filtered = df_pur[(df_pur["Status"] == "Complete") & (df_pur["Tanggal Complete"].dt.date == today)]
        label_tabel_pur = "Data PUR Selesai Hari Ini"
    else:
        with col_f_pur2:
            date_range_pur = st.date_input("Pilih Rentang Tanggal Selesai:", 
                                           value=(today - timedelta(days=7), today), 
                                           key="date_pur")
        if len(date_range_pur) == 2:
            start_p, end_p = date_range_pur
            df_pur_filtered = df_pur[
                (df_pur["Status"] == "Complete") & 
                (df_pur["Tanggal Complete"].dt.date >= start_p) & 
                (df_pur["Tanggal Complete"].dt.date <= end_p)
            ]
            label_tabel_pur = f"Data PUR Selesai Periode {start_p} s/d {end_p}"
        else:
            df_pur_filtered = df_pur[df_pur["Status"] == "Complete"]
            label_tabel_pur = "Silakan pilih rentang tanggal"

    # Metric
    total_all_pur = len(df_pur)
    total_complete_pur = len(df_pur[df_pur["Status"] == "Complete"])
    total_filtered_pur = len(df_pur_filtered)

    cp1, cp2, cp3 = st.columns(3)
    cp1.metric("Total Semua PUR", total_all_pur)
    cp2.metric("Total Semua Complete", total_complete_pur)
    cp3.metric("PUR Terfilter (Complete)", total_filtered_pur)

    st.divider()
    st.subheader(label_tabel_pur)
    if total_filtered_pur > 0:
        if mode_pur == "Selesai Hari Ini":
            st.error(f"Ada {total_filtered_pur} PUR selesai hari ini!")
        st.dataframe(df_pur_filtered, use_container_width=True)
    else:
        st.info("Tidak ada data PUR yang cocok dengan filter tanggal tersebut.")


# =================================================================
# TAB 3: SQ TO SO (Dua Tabel & Filter Week)
# =================================================================
with tab3:
    st.header("Dashboard SQ to SO")
    df_sq_to_so = pd.read_excel("data_sq_to_so.xlsx")
    df_sq_baru = pd.read_excel("data_sq.xlsx")
    df_sq_to_so.columns = df_sq_to_so.columns.str.strip()
    df_sq_baru.columns = df_sq_baru.columns.str.strip()

    # --- BAGIAN 1: SQ TO SO ---
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

    # --- BAGIAN 2: DATA SQ BARU (Filter Customer & Week) ---
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
    c2a.metric(f"Total SQ Baru", len(df2_f))
    c2b.metric("Subtotal (Non-Draft)", f"Rp {subtotal2:,.0f}".replace(",", "."))
    st.dataframe(df2_f, use_container_width=True)


# =================================================================
# TAB 4: KPI MARKETING (Scorecard Tanpa List)
# =================================================================
with tab4:
    st.header("KPI Marketing Performance")
    # Membaca 2 sheet dari data_kpi.xlsx
    df_si = pd.read_excel("data_kpi.xlsx", sheet_name="SI")
    df_sq_kpi = pd.read_excel("data_kpi.xlsx", sheet_name="SQ")
    df_si.columns = df_si.columns.str.strip()
    df_sq_kpi.columns = df_sq_kpi.columns.str.strip()

    # Konversi Tanggal
    df_si["Tanggal"] = pd.to_datetime(df_si["Tanggal"], errors="coerce")
    df_sq_kpi["Tanggal"] = pd.to_datetime(df_sq_kpi["Tanggal"], errors="coerce")

    # Filter Area
    f_col1, f_col2, f_col3 = st.columns(3)
    with f_col1:
        years = sorted(df_si["Tanggal"].dt.year.dropna().unique().astype(int).tolist(), reverse=True)
        sel_year = st.selectbox("Tahun", years, key="kpi_y")
    with f_col2:
        months = {1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni", 
                  7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"}
        sel_month = st.selectbox("Bulan", list(months.keys()), format_func=lambda x: months[x], key="kpi_m")
    with f_col3:
        list_s = sorted(list(set(df_si["Salesman"].dropna().unique().tolist() + df_sq_kpi["Sales"].dropna().unique().tolist())))
        sel_sales = st.selectbox("Salesman", ["Semua"] + list_s, key="kpi_s")

    # Logika Filter
    mask_si = (df_si["Tanggal"].dt.year == sel_year) & (df_si["Tanggal"].dt.month == sel_month)
    if sel_sales != "Semua": mask_si &= (df_si["Salesman"] == sel_sales)
    
    mask_sq = (df_sq_kpi["Tanggal"].dt.year == sel_year) & (df_sq_kpi["Tanggal"].dt.month == sel_month)
    if sel_sales != "Semua": mask_sq &= (df_sq_kpi["Sales"] == sel_sales)

    # Hitung Values (Kecuali Draft)
    val_si = df_si[mask_si & (df_si["Status"] != "Draft")]["Total Harga Jual"].sum()
    val_sq = df_sq_kpi[mask_sq & (df_sq_kpi["Status"] != "Draft")]["Sub Total"].sum()

    # Tampilan Scorecard
    st.divider()
    st.subheader(f"Ringkasan KPI: {months[sel_month]} {sel_year}")
    m1, m2 = st.columns(2)
    m1.metric(f"Total SI - {sel_sales}", f"Rp {val_si:,.0f}".replace(",", "."))
    m2.metric(f"Total SQ - {sel_sales}", f"Rp {val_sq:,.0f}".replace(",", "."))

    st.divider()
    # Aktivitas Transaksi (Opsional Info)
    i1, i2 = st.columns(2)
    i1.info(f"Transaksi SI Berjalan: {len(df_si[mask_si & (df_si['Status'] != 'Draft')])}")
    i2.info(f"Transaksi SQ Berjalan: {len(df_sq_kpi[mask_sq & (df_sq_kpi['Status'] != 'Draft')])}")