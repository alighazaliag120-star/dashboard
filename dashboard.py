import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(layout="wide")

st.title("Dashboard Monitoring NPR, PUR & SQ")

# Tambahkan tab ketiga: SQ to SO
tab1, tab2, tab3 = st.tabs(["NPR", "PUR", "SQ to SO"])

# =============================
# TAB NPR
# =============================
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

    selesai_hari_ini = df_npr[
        (df_npr["Status"] == "Complete") &
        (df_npr["Tanggal Complete"].dt.date == today)
    ]

    total_today = len(selesai_hari_ini)

    col1, col2, col3 = st.columns(3)

    col1.metric("Total NPR", total_npr)
    col2.metric("Total NPR Complete", total_complete)
    col3.metric("Selesai Hari Ini", total_today)

    st.divider()

    st.subheader("Data NPR selesai hari ini")

    if total_today > 0:
        st.error(f"Ada {total_today} NPR selesai hari ini")
        st.dataframe(selesai_hari_ini)
    else:
        st.success("Tidak ada NPR selesai hari ini")

    st.divider()
    st.subheader("Semua Data NPR")
    st.dataframe(df_npr)


# =============================
# TAB PUR
# =============================
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

    pur_today = df_pur[
        (df_pur["Status"] == "Complete") &
        (df_pur["Tanggal Complete"].dt.date == today)
    ]

    total_today_pur = len(pur_today)

    col1, col2, col3 = st.columns(3)

    col1.metric("Total PUR", total_pur)
    col2.metric("Total PUR Complete", total_complete_pur)
    col3.metric("Selesai Hari Ini", total_today_pur)

    st.divider()

    st.subheader("Data PUR selesai hari ini")

    if total_today_pur > 0:
        st.error(f"Ada {total_today_pur} PUR selesai hari ini")
        st.dataframe(pur_today)
    else:
        st.success("Tidak ada PUR selesai hari ini")

    st.divider()
    st.subheader("Semua Data PUR")
    st.dataframe(df_pur)

# =============================
# TAB SQ TO SO
# =============================
# =============================
# TAB SQ TO SO
# =============================
with tab3:
    st.header("Dashboard SQ to SO")

    # Baca data SQ
    df_sq = pd.read_excel("data_sq_to_so.xlsx")
    df_sq.columns = df_sq.columns.str.strip()

    # Filter Customer
    customer_list = df_sq["Customer"].dropna().unique().tolist()
    customer = st.selectbox("Filter Customer SQ", ["Semua"] + sorted(customer_list))

    # Terapkan Filter Customer
    if customer != "Semua":
        df_sq_filtered = df_sq[df_sq["Customer"] == customer]
    else:
        df_sq_filtered = df_sq

    # --- LOGIKA REVISI: Filter Status ---
    # Kita hanya ambil data yang statusnya Complete atau In Progress
    df_valid_status = df_sq_filtered[df_sq_filtered["Status"].isin(["Complete", "In Progress"])]

    # 1. Hitung Jumlah Transaksi Unik (hanya yang Complete & In Progress)
    sq_complete = df_valid_status[df_valid_status["Status"] == "Complete"]["No Transaksi"].nunique()
    sq_inprogress = df_valid_status[df_valid_status["Status"] == "In Progress"]["No Transaksi"].nunique()
    
    # 2. Hitung Total Nilai Barang (Hanya dari data yang Complete & In Progress)
    total_nilai = df_valid_status["Total Barang"].sum()

    # --- TAMPILAN METRIC ---
    col1, col2, col3 = st.columns(3)
    col1.metric("SQ Complete", sq_complete)
    col2.metric("SQ In Progress", sq_inprogress)
    
    # Menampilkan total harga dengan format Rupiah (Pemisah Titik)
    col3.metric("Total Nilai (Comp & In-Prog)", f"Rp {total_nilai:,.0f}".replace(",", "."))

    st.divider()
    
    st.subheader(f"Detail Data SQ: {customer}")
    
    # Menampilkan tabel (tetap menampilkan data asli hasil filter customer agar Anda bisa lihat semua status)
    st.dataframe(df_sq_filtered)