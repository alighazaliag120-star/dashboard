import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(layout="wide")

st.title("Dashboard Monitoring NPR & PUR")

tab1, tab2 = st.tabs(["NPR", "PUR"])

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