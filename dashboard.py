import streamlit as st
import pandas as pd
from datetime import date

# Konfigurasi Halaman
st.set_page_config(layout="wide", page_title="Dashboard Monitoring Monitoring")

st.title("Dashboard Monitoring")

# Definisi Tab (Menambahkan Tab 4)
tab1, tab2, tab3, tab4 = st.tabs(["NPR", "PUR", "SQ to SO", "KPI Marketing"])

# =================================================================
# TAB 1: NPR
# =================================================================
with tab1:
    st.header("Dashboard NPR")
    df_npr = pd.read_excel("data_npr.xlsx")
    df_npr.columns = df_npr.columns.str.strip()
    df_npr["Tanggal Complete"] = pd.to_datetime(df_npr["Tanggal Complete"], errors="coerce")
    today = date.today()

    nama_list = df_npr["Penanggung Jawab"].dropna().unique().tolist()
    nama = st.selectbox("Filter Penanggung Jawab NPR", ["Semua"] + sorted(nama_list))

    if nama != "Semua":
        df_npr = df_npr[df_npr["Penanggung Jawab"] == nama]

    total_npr = len(df_npr)
    total_complete = len(df_npr[df_npr["Status"] == "Complete"])
    selesai_hari_ini = df_npr[(df_npr["Status"] == "Complete") & (df_npr["Tanggal Complete"].dt.date == today)]
    total_today = len(selesai_hari_ini)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total NPR", total_npr)
    col2.metric("Total NPR Complete", total_complete)
    col3.metric("Selesai Hari Ini", total_today)

    st.divider()
    st.subheader("Data NPR selesai hari ini")
    if total_today > 0:
        st.error(f"Ada {total_today} NPR selesai hari ini")
        st.dataframe(selesai_hari_ini, use_container_width=True)
    else:
        st.success("Tidak ada NPR selesai hari ini")

    st.divider()
    st.subheader("Semua Data NPR")
    st.dataframe(df_npr, use_container_width=True)


# =================================================================
# TAB 2: PUR
# =================================================================
with tab2:
    st.header("Dashboard PUR")
    df_pur = pd.read_excel("data_pur.xlsx")
    df_pur.columns = df_pur.columns.str.strip()
    df_pur["Tanggal Complete"] = pd.to_datetime(df_pur["Tanggal Complete"], errors="coerce")
    today = date.today()

    nama_list_pur = df_pur["Penanggung Jawab"].dropna().unique().tolist()
    nama_pur = st.selectbox("Filter Penanggung Jawab PUR", ["Semua"] + sorted(nama_list_pur))

    if nama_pur != "Semua":
        df_pur = df_pur[df_pur["Penanggung Jawab"] == nama_pur]

    total_pur = len(df_pur)
    total_complete_pur = len(df_pur[df_pur["Status"] == "Complete"])
    pur_today = df_pur[(df_pur["Status"] == "Complete") & (df_pur["Tanggal Complete"].dt.date == today)]
    total_today_pur = len(pur_today)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total PUR", total_pur)
    col2.metric("Total PUR Complete", total_complete_pur)
    col3.metric("Selesai Hari Ini", total_today_pur)

    st.divider()
    st.subheader("Data PUR selesai hari ini")
    if total_today_pur > 0:
        st.error(f"Ada {total_today_pur} PUR selesai hari ini")
        st.dataframe(pur_today, use_container_width=True)
    else:
        st.success("Tidak ada PUR selesai hari ini")

    st.divider()
    st.subheader("Semua Data PUR")
    st.dataframe(df_pur, use_container_width=True)


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