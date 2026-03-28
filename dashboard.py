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
with tab3:
    st.header("Dashboard SQ to SO")

    # 1. BACA DATA
    df_sq_to_so = pd.read_excel("data_sq_to_so.xlsx")
    df_sq_baru = pd.read_excel("data_sq.xlsx")

    # Bersihkan nama kolom
    df_sq_to_so.columns = df_sq_to_so.columns.str.strip()
    df_sq_baru.columns = df_sq_baru.columns.str.strip()

    # ==========================================
    # BAGIAN ATAS: DATA SQ TO SO (Tabel 1)
    # ==========================================
    st.subheader("1. Monitoring SQ to SO")
    
    cust_list1 = df_sq_to_so["Customer"].dropna().unique().tolist()
    filter_cust1 = st.selectbox("Filter Customer (SQ to SO)", ["Semua"] + sorted(cust_list1), key="filter_cust1")

    df1_filtered = df_sq_to_so if filter_cust1 == "Semua" else df_sq_to_so[df_sq_to_so["Customer"] == filter_cust1]

    # Metric Tabel 1
    subtotal1 = df1_filtered[df1_filtered["Status"] != "Draft"]["Total Barang"].sum()
    col_a, col_b = st.columns(2)
    col_a.metric("Total SQ to SO", len(df1_filtered))
    col_b.metric("Subtotal (Non-Draft)", f"Rp {subtotal1:,.0f}".replace(",", "."))

    st.dataframe(df1_filtered, use_container_width=True)

    st.divider()

    # ==========================================
    # BAGIAN BAWAH: DATA SQ BARU (Tabel 2 - Dengan Filter Week)
    # ==========================================
    st.subheader("2. Monitoring Data SQ Baru")

    # Membuat dua kolom untuk filter agar sejajar
    col_filter1, col_filter2 = st.columns(2)

    with col_filter1:
        cust_list2 = df_sq_baru["Customer"].dropna().unique().tolist()
        filter_cust2 = st.selectbox("Filter Customer", ["Semua"] + sorted(cust_list2), key="filter_cust2")

    with col_filter2:
        # Menambahkan filter Week
        week_list = df_sq_baru["Week"].dropna().unique().tolist()
        # Mengurutkan week (asumsi data week adalah angka atau string yang bisa diurutkan)
        filter_week = st.selectbox("Filter Week", ["Semua"] + sorted(week_list), key="filter_week2")

    # Logika Filter Berlapis untuk Tabel 2
    df2_filtered = df_sq_baru.copy()

    if filter_cust2 != "Semua":
        df2_filtered = df2_filtered[df2_filtered["Customer"] == filter_cust2]
    
    if filter_week != "Semua":
        df2_filtered = df2_filtered[df2_filtered["Week"] == filter_week]

    # Hitung Subtotal Tabel 2 (Kecuali Draft)
    subtotal2 = df2_filtered[df2_filtered["Status"] != "Draft"]["Sub Total"].sum()

    # Metric Tabel 2
    col_c, col_d = st.columns(2)
    col_c.metric(f"Total SQ Baru (Week {filter_week})", len(df2_filtered))
    col_d.metric("Subtotal (Non-Draft)", f"Rp {subtotal2:,.0f}".replace(",", "."))

    st.dataframe(df2_filtered, use_container_width=True)