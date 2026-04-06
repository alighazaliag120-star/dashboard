import streamlit as st
import pandas as pd
from datetime import date, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json

# ==========================================
# 1. KONFIGURASI HALAMAN (Wajib paling atas)
# ==========================================
st.set_page_config(layout="wide", page_title="Dashboard Monitoring", initial_sidebar_state="expanded")

# ==========================================
# 2. FUNGSI DATABASE
# ==========================================
def load_gsheet_all():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    spreadsheet_id = "10lpdIeAkhQj1Rv2tnP2V806edBnpCICbU2bsf5OslKc"
    
    try:
        if os.path.exists("kunci_data.json"):
            creds = ServiceAccountCredentials.from_json_keyfile_name("kunci_data.json", scope)
        else:
            raw_json_str = st.secrets["gcp_service_account"]["content"]
            creds_info = json.loads(raw_json_str)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
            
        client = gspread.authorize(creds)
        sheet = client.open_by_key(spreadsheet_id).worksheet("ALL")
        data = sheet.get_all_values()
        
        if len(data) > 3:
            df = pd.DataFrame(data[4:], columns=data[3])
            df.columns = df.columns.str.strip()
            df = df.loc[:, df.columns != '']
            return df
        return pd.DataFrame()
    except Exception as e:
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
        sheet = client.open_by_key(spreadsheet_id).get_worksheet(0)
        data = sheet.get_all_values()
        
        if len(data) > 8:
            df = pd.DataFrame(data[9:], columns=data[8])
            df.columns = df.columns.str.strip().str.upper()
            df = df.loc[:, df.columns != '']
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Gagal koneksi database BPV: {e}")
        return pd.DataFrame()

# ==========================================
# FUNGSI BANTUAN UNTUK FORMAT DOWNLOAD
# ==========================================
# Fungsi ini akan mengubah format tanggal menjadi YYYY-MM-DD
# sehingga Excel bisa mengurutkannya dengan sempurna.
def format_tgl_internasional(df, list_kolom):
    df_copy = df.copy()
    for col in list_kolom:
        if col in df_copy.columns:
            df_copy[col] = pd.to_datetime(df_copy[col], errors="coerce").dt.strftime('%Y-%m-%d').fillna('-')
    return df_copy


# ==========================================
# 3. LOGIKA PASSWORD & TAMPILAN AWAL
# ==========================================

if 'login_sukses' not in st.session_state:
    st.session_state['login_sukses'] = False

def cek_password():
    if st.session_state['password_input'] == "sibimabatulicin":
        st.session_state['login_sukses'] = True
    else:
        st.error("😕 Password salah bro! Coba lagi.")

# --- HALAMAN LOGIN ---
if not st.session_state['login_sukses']:
    st.sidebar.warning("🔒 Silakan login di layar utama untuk mengakses menu.")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔒 Login Dashboard")
        st.write("Silakan masukkan password untuk mengakses data.")
        st.text_input("Password", type="password", key="password_input")
        st.button("Masuk", on_click=cek_password)

# --- HALAMAN UTAMA ---
else:
    header_col1, header_col2 = st.columns([4, 1])
    with header_col1:
        st.title("📊 Dashboard Monitoring")
    with header_col2:
        st.write("") 
        if st.button("Keluar / Logout", use_container_width=True):
            st.session_state['login_sukses'] = False
            st.rerun() 

    st.success("Berhasil masuk! Selamat bekerja.")
    st.markdown("---")

    # =================================================================
    # SIDEBAR NAVIGATION
    # =================================================================
    with st.sidebar:
        st.title("Main Menu")
        menu_pilihan = st.radio(
            "Pilih Dashboard:",
            ["HOME", "NPR", "PUR", "SQ to SO", "KPI Marketing", "Laporan Weekly", "Status BPV", "History Penjualan Terakhir", "Tracking Vendor"], 
            index=0 
        )
        st.divider()
        st.info("Klik panah '>' di pojok kiri atas untuk menutup/membuka menu samping.")

    today = date.today()

    hari_indo = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    bulan_indo = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

    nama_hari = hari_indo[today.weekday()]
    nama_bulan = bulan_indo[today.month]
    tanggal_sekarang_str = f"{nama_hari}, {today.day} {nama_bulan} {today.year}"

    # =================================================================
    # LOGIKA TAMPILAN BERDASARKAN MENU SIDEBAR
    # =================================================================

    # --- MENU UTAMA: HOME ---
    if menu_pilihan == "HOME":
        st.markdown("<h1 style='text-align: center;'>DASHBOARD MONITORING</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: gray; font-size: 18px;'>📅 {tanggal_sekarang_str}</p>", unsafe_allow_html=True)
        st.divider()

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
        st.subheader("📊 Business Overview")

        try:
            df_bpv_home = load_gsheet_bpv()
            if not df_bpv_home.empty:
                st.write("**💰 Persentase Status Pembayaran BPV**")
                df_bpv_home.columns = df_bpv_home.columns.str.strip().str.upper()
                target_col = 'TANGGAL BAYAR' 
                if target_col in df_bpv_home.columns:
                    def cek_status(val):
                        v = str(val).strip().lower()
                        if v in ['', '-', 'nan', 'none', 'null', '0', 'nat']:
                            return 'Belum Dibayar'
                        return 'Dibayar'

                    df_bpv_home['Status Real'] = df_bpv_home[target_col].apply(cek_status)
                    status_pie = df_bpv_home['Status Real'].value_counts()
                    
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
            st.error(f"Gagal memuat chart BPV: {e}")

        st.write("⚙️ **Filter Grafik KPI & Trend**")
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            sel_year_home = st.selectbox("Pilih Tahun Dashboard", [2026, 2025], key="h_year")
        with f_col2:
            df_all_raw = load_gsheet_all()
            list_cust_home = ["Semua"] + sorted(df_all_raw["Customer"].unique().tolist()) if not df_all_raw.empty else ["Semua"]
            sel_cust_home = st.selectbox("Filter Customer (Trend PO)", list_cust_home, key="h_cust")

        st.write(f"**🏆 Performa Salesman - {sel_year_home}**")
        try:
            df_si_home = pd.read_excel("data_kpi.xlsx", sheet_name="SI")
            df_sq_home = pd.read_excel("data_kpi.xlsx", sheet_name="SQ")
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

        st.write(f"**📈 Trend Weekly PO Jhonlin - {sel_cust_home}**")
        try:
            if not df_all_raw.empty:
                df_trend = df_all_raw.copy()
                if sel_cust_home != "Semua":
                    df_trend = df_trend[df_trend["Customer"] == sel_cust_home]
                trend_data = df_trend.groupby('Week').size()
                st.line_chart(trend_data)
        except:
            st.info("Data Trend PO sedang diproses...")


    # --- MENU 1: NPR (EXCEL) ---
    elif menu_pilihan == "NPR":
        st.header("Dashboard NPR")
        st.caption(f"📅 {tanggal_sekarang_str}") 
        
        df_npr = pd.read_excel("data_npr.xlsx")
        df_npr.columns = df_npr.columns.str.strip()
        df_npr["Tanggal Complete"] = pd.to_datetime(df_npr["Tanggal Complete"], errors="coerce")
        
        df_npr_belum_complete = df_npr[df_npr["Tanggal Complete"].isna()]
        
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

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Semua NPR", len(df_npr))
        c2.metric("Total Semua Complete", len(df_npr[df_npr["Status"] == "Complete"]))
        c3.metric("Total NPR Belum Complete", len(df_npr_belum_complete))
        c4.metric("NPR Terfilter (Complete)", len(df_npr_filtered))
        
        st.divider()
        st.subheader(label_tabel_npr)
        # Format tanggal sebelum ditampilkan dan disiapkan untuk tombol download
        df_npr_filtered_disp = format_tgl_internasional(df_npr_filtered, ["Tanggal", "Tanggal Complete", "TANGGAL"])
        st.dataframe(df_npr_filtered_disp, use_container_width=True)
        st.download_button("📥 Download NPR Selesai (CSV)", df_npr_filtered_disp.to_csv(index=False).encode('utf-8'), "NPR_Selesai.csv", "text/csv", key="dl_npr1")

        st.divider()
        st.subheader("Data NPR Belum Complete")
        df_npr_belum_disp = format_tgl_internasional(df_npr_belum_complete, ["Tanggal", "Tanggal Complete", "TANGGAL"])
        st.dataframe(df_npr_belum_disp, use_container_width=True)
        st.download_button("📥 Download NPR Belum Complete (CSV)", df_npr_belum_disp.to_csv(index=False).encode('utf-8'), "NPR_Belum_Complete.csv", "text/csv", key="dl_npr2")


    # --- MENU 2: PUR (EXCEL) ---
    elif menu_pilihan == "PUR":
        st.header("Dashboard PUR")
        st.caption(f"📅 {tanggal_sekarang_str}") 
        
        df_pur = pd.read_excel("data_pur.xlsx")
        df_pur.columns = df_pur.columns.str.strip()
        df_pur["Tanggal Complete"] = pd.to_datetime(df_pur["Tanggal Complete"], errors="coerce")

        df_pur_belum_complete = df_pur[df_pur["Tanggal Complete"].isna()]

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

        cp1, cp2, cp3, cp4 = st.columns(4)
        cp1.metric("Total Semua PUR", len(df_pur))
        cp2.metric("Total Semua Complete", len(df_pur[df_pur["Status"] == "Complete"]))
        cp3.metric("Total PUR Belum Complete", len(df_pur_belum_complete))
        cp4.metric("PUR Terfilter (Complete)", len(df_pur_filtered))
        
        st.divider()
        st.subheader(label_tabel_pur)
        df_pur_filtered_disp = format_tgl_internasional(df_pur_filtered, ["Tanggal", "Tanggal Complete", "TANGGAL"])
        st.dataframe(df_pur_filtered_disp, use_container_width=True)
        st.download_button("📥 Download PUR Selesai (CSV)", df_pur_filtered_disp.to_csv(index=False).encode('utf-8'), "PUR_Selesai.csv", "text/csv", key="dl_pur1")

        st.divider()
        st.subheader("Data PUR Belum Complete")
        df_pur_belum_disp = format_tgl_internasional(df_pur_belum_complete, ["Tanggal", "Tanggal Complete", "TANGGAL"])
        st.dataframe(df_pur_belum_disp, use_container_width=True)
        st.download_button("📥 Download PUR Belum Complete (CSV)", df_pur_belum_disp.to_csv(index=False).encode('utf-8'), "PUR_Belum_Complete.csv", "text/csv", key="dl_pur2")


    # --- MENU 3: SQ TO SO (EXCEL) ---
    elif menu_pilihan == "SQ to SO":
        st.header("Dashboard SQ to SO")
        st.caption(f"📅 {tanggal_sekarang_str}")
        
        try:
            df_sq_to_so = pd.read_excel("data_sq_to_so.xlsx")
            df_sq_baru = pd.read_excel("data_sq.xlsx") 
            df_sq_to_so["Tanggal"] = pd.to_datetime(df_sq_to_so["Tanggal"], errors="coerce")
            df_sq_baru["Tanggal"] = pd.to_datetime(df_sq_baru["Tanggal"], errors="coerce")
        except Exception as e:
            st.error(f"Gagal membaca file Excel: {e}")
            st.stop()
        
        st.subheader("1. Monitoring SQ to SO")
        col_f1_1, col_f1_2 = st.columns(2)
        with col_f1_1:
            cust_list1 = df_sq_to_so["Customer"].dropna().unique().tolist()
            filter_cust1 = st.selectbox("Filter Customer (SQ to SO)", ["Semua"] + sorted(cust_list1), key="f_cust1")
        with col_f1_2:
            min_date1 = df_sq_to_so["Tanggal"].min()
            max_date1 = df_sq_to_so["Tanggal"].max()
            val_date1 = (min_date1.date(), max_date1.date()) if pd.notna(min_date1) else ()
            filter_date1 = st.date_input("Filter Rentang Tanggal (SQ to SO)", value=val_date1, key="f_date1")

        df1_f = df_sq_to_so.copy()
        if filter_cust1 != "Semua":
            df1_f = df1_f[df1_f["Customer"] == filter_cust1]
        if len(filter_date1) == 2:
            start_date1, end_date1 = filter_date1
            df1_f = df1_f[(df1_f["Tanggal"].dt.date >= start_date1) & (df1_f["Tanggal"].dt.date <= end_date1)]
        
        status_valid = ["Complete", "In Progress"]
        df1_to_so = df1_f[df1_f["Status"].str.strip().isin(status_valid)]
        qty_to_so = df1_to_so["No Transaksi"].nunique()
        val_to_so = df1_to_so["Total Barang"].sum()

        df1_pending = df1_f[~df1_f["Status"].str.strip().isin(status_valid)]
        qty_pending = df1_pending["No Transaksi"].nunique()
        val_pending = df1_pending["Total Barang"].sum()

        qty_all = df1_f["No Transaksi"].nunique()
        val_all = df1_f["Total Barang"].sum()

        st.markdown("#### ✅ SQ To SO (Complete & In Progress)")
        m1, m2 = st.columns(2)
        m1.metric("Qty SQ (Unik)", f"{qty_to_so} SQ")
        m2.metric("Total Value", f"Rp {val_to_so:,.0f}".replace(",", "."))

        st.markdown("#### ⏳ SQ Pending (Draft & Others)")
        m3, m4 = st.columns(2)
        m3.metric("Qty SQ (Unik)", f"{qty_pending} SQ")
        m4.metric("Total Value", f"Rp {val_pending:,.0f}".replace(",", "."))
        
        st.markdown("#### 📦 ALL SQ (Seluruh Status)")
        m5, m6 = st.columns(2)
        m5.metric("Qty SQ (Unik)", f"{qty_all} SQ")
        m6.metric("Total Value", f"Rp {val_all:,.0f}".replace(",", "."))
        
        st.divider()
        df1_disp = format_tgl_internasional(df1_f, ["Tanggal", "TANGGAL"])
        st.dataframe(df1_disp, width="stretch")
        st.download_button("📥 Download Data SQ to SO (CSV)", df1_disp.to_csv(index=False).encode('utf-8'), "SQ_to_SO.csv", "text/csv", key="dl_sq1")

        st.subheader("2. Monitoring Data SQ Baru")
        cf1, cf2, cf3 = st.columns(3)
        with cf1:
            list_cust2 = ["Semua"] + sorted(df_sq_baru["Customer"].dropna().unique().tolist())
            filter_cust2 = st.selectbox("Filter Customer", list_cust2, key="f_cust2")
        with cf2:
            list_week2 = ["Semua"] + sorted(df_sq_baru["Week"].dropna().unique().tolist())
            filter_week = st.selectbox("Filter Week", list_week2, key="f_week")
        with cf3:
            min_date2 = df_sq_baru["Tanggal"].min()
            max_date2 = df_sq_baru["Tanggal"].max()
            val_date2 = (min_date2.date(), max_date2.date()) if pd.notna(min_date2) else ()
            filter_date2 = st.date_input("Filter Rentang Tanggal (SQ Baru)", value=val_date2, key="f_date2")
        
        df2_f = df_sq_baru.copy()
        if filter_cust2 != "Semua": 
            df2_f = df2_f[df2_f["Customer"] == filter_cust2]
        if filter_week != "Semua": 
            df2_f = df2_f[df2_f["Week"] == filter_week]
        if len(filter_date2) == 2:
            start_date2, end_date2 = filter_date2
            df2_f = df2_f[(df2_f["Tanggal"].dt.date >= start_date2) & (df2_f["Tanggal"].dt.date <= end_date2)]
        
        df2_non_draft = df2_f[df2_f["Status"].str.strip().str.lower() != "draft"]
        col_id_sq2 = "No Transaksi" if "No Transaksi" in df2_f.columns else df2_f.columns[1]
        qty_sq_non_draft = df2_non_draft[col_id_sq2].nunique()
        val_sq_non_draft = df2_non_draft["Sub Total"].sum()

        st.markdown("#### 📊 Ringkasan SQ Baru (Selain Draft)")
        b1, b2 = st.columns(2)
        b1.metric("Qty SQ (Unik)", f"{qty_sq_non_draft} SQ")
        b2.metric("Total Value", f"Rp {val_sq_non_draft:,.0f}".replace(",", "."))
        
        st.divider()
        df2_disp = format_tgl_internasional(df2_f, ["Tanggal", "TANGGAL"])
        st.dataframe(df2_disp, use_container_width=True)
        st.download_button("📥 Download Data SQ Baru (CSV)", df2_disp.to_csv(index=False).encode('utf-8'), "SQ_Baru.csv", "text/csv", key="dl_sq2")


    # --- MENU 4: KPI MARKETING (EXCEL) ---
    elif menu_pilihan == "KPI Marketing":
        st.header("KPI Marketing Performance")
        st.caption(f"📅 {tanggal_sekarang_str}")
        
        df_si = pd.read_excel("data_kpi.xlsx", sheet_name="SI")
        df_sq_kpi = pd.read_excel("data_kpi.xlsx", sheet_name="SQ")
        df_si.columns = df_si.columns.str.strip()
        df_sq_kpi.columns = df_sq_kpi.columns.str.strip()

        df_si["Tanggal"] = pd.to_datetime(df_si["Tanggal"], errors="coerce")
        df_sq_kpi["Tanggal"] = pd.to_datetime(df_sq_kpi["Tanggal"], errors="coerce")

        if "Status" in df_si.columns:
            df_si["Status"] = df_si["Status"].astype(str).str.strip().str.title()
        if "Status" in df_sq_kpi.columns:
            df_sq_kpi["Status"] = df_sq_kpi["Status"].astype(str).str.strip().str.title()

        if "Total Nilai" in df_si.columns:
            df_si["Total Nilai"] = pd.to_numeric(df_si["Total Nilai"], errors='coerce').fillna(0)
        
        if "Sub Total" in df_sq_kpi.columns:
            df_sq_kpi["Sub Total"] = pd.to_numeric(df_sq_kpi["Sub Total"], errors='coerce').fillna(0)

        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            years = ["Semua"] + sorted(df_si["Tanggal"].dt.year.dropna().unique().astype(int).tolist(), reverse=True)
            sel_year = st.selectbox("Tahun", years, key="kpi_y")
        with f_col2:
            months = {0: "Semua", 1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni", 
                      7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"}
            sel_month = st.selectbox("Bulan", list(months.keys()), format_func=lambda x: months[x], key="kpi_m")
        with f_col3:
            list_s = sorted(list(set(df_si["Salesman"].dropna().unique().tolist() + df_sq_kpi["Sales"].dropna().unique().tolist())))
            sel_sales = st.selectbox("Salesman", ["Semua"] + list_s, key="kpi_s")

        mask_si = pd.Series(True, index=df_si.index)
        mask_sq = pd.Series(True, index=df_sq_kpi.index)

        if sel_year != "Semua":
            mask_si &= (df_si["Tanggal"].dt.year == sel_year)
            mask_sq &= (df_sq_kpi["Tanggal"].dt.year == sel_year)

        if sel_month != 0:
            mask_si &= (df_si["Tanggal"].dt.month == sel_month)
            mask_sq &= (df_sq_kpi["Tanggal"].dt.month == sel_month)
        
        if sel_sales != "Semua":
            mask_si &= (df_si["Salesman"] == sel_sales)
            mask_sq &= (df_sq_kpi["Sales"] == sel_sales)

        val_si = df_si[mask_si & (df_si["Status"] != "Draft")]["Total Nilai"].sum()
        val_sq = df_sq_kpi[mask_sq & (df_sq_kpi["Status"] != "Draft")]["Sub Total"].sum()
        val_po = df_sq_kpi[mask_sq & df_sq_kpi["Status"].isin(["Complete", "In Progress"])]["Sub Total"].sum()

        st.divider()
        label_bulan = "Seluruh Bulan" if sel_month == 0 else months[sel_month]
        label_tahun = "Semua Tahun" if sel_year == "Semua" else str(sel_year)
        
        st.subheader(f"Ringkasan KPI: {label_bulan} - {label_tahun}")
        
        custom_css = """
        <style>
        .metric-card {
            background-color: #262730; 
            border: 1px solid #4B4C5A; 
            padding: 15px;
            border-radius: 10px; 
            text-align: center;
            box-shadow: 2px 2px 10px rgba(0,0,0,0.2); 
        }
        .metric-title { color: #FAFAFA; font-size: 16px; margin-bottom: 5px; }
        .metric-value { color: #4CAF50; font-size: 24px; font-weight: bold; margin-top: 0px; }
        </style>
        """
        st.markdown(custom_css, unsafe_allow_html=True)

        m1, m2, m3 = st.columns(3)
        format_sq = f"Rp {val_sq:,.0f}".replace(",", ".")
        format_po = f"Rp {val_po:,.0f}".replace(",", ".")
        format_si = f"Rp {val_si:,.0f}".replace(",", ".")

        with m1:
            st.markdown(f'<div class="metric-card"><p class="metric-title">Total SQ - {sel_sales}</p><p class="metric-value">{format_sq}</p></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-card"><p class="metric-title">Total PO Diterima - {sel_sales}</p><p class="metric-value">{format_po}</p></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="metric-card"><p class="metric-title">Total SI - {sel_sales}</p><p class="metric-value">{format_si}</p></div>', unsafe_allow_html=True)


    # --- MENU 5: LAPORAN WEEKLY ---
    elif menu_pilihan == "Laporan Weekly":
        st.header("Laporan Weekly - Monitoring PO Jhonlin")
        st.caption(f"📅 {tanggal_sekarang_str}")
        
        try:
            df_all = load_gsheet_all()
            if df_all.empty:
                st.warning("Data tidak tersedia atau gagal memuat.")
            else:
                df_all.columns = df_all.columns.str.strip()

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

                df_filtered = df_all.copy()
                if sel_bulan_w != "Semua":
                    df_filtered = df_filtered[df_filtered["Bulan"].astype(str) == sel_bulan_w]
                if sel_cust_w != "Semua":
                    df_filtered = df_filtered[df_filtered["Customer"].astype(str) == sel_cust_w]
                if sel_week_w != "Semua":
                    df_filtered = df_filtered[df_filtered["Week"].astype(str) == sel_week_w]
                if sel_tgl_w != "Semua":
                    df_filtered = df_filtered[df_filtered["Tgl Terima Email"].astype(str) == sel_tgl_w]

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
                df_weekly_disp = format_tgl_internasional(df_filtered, ["Tgl Terima Email", "Tanggal", "TANGGAL"])
                st.dataframe(df_weekly_disp, use_container_width=True)
                st.download_button("📥 Download Laporan Weekly (CSV)", df_weekly_disp.to_csv(index=False).encode('utf-8'), "Laporan_Weekly.csv", "text/csv", key="dl_weekly")

        except Exception as e:
            st.error(f"Terjadi kesalahan teknis tampilan: {e}")

    # --- MENU 6: STATUS BPV & TRACKER PO ---
    elif menu_pilihan == "Status BPV":
        st.header("🔔 Monitoring & Notifikasi BPV")
        st.caption(f"📅 {tanggal_sekarang_str}")
        
        try:
            df_bpv = load_gsheet_bpv()
            
            if df_bpv.empty:
                st.warning("Data BPV tidak tersedia atau gagal memuat.")
            else:
                col_tgl_bpv = 'TANGGAL BPV'
                col_tgl_bayar = 'TANGGAL BAYAR'

                if col_tgl_bpv in df_bpv.columns and col_tgl_bayar in df_bpv.columns:
                    df_bpv[col_tgl_bpv] = pd.to_datetime(df_bpv[col_tgl_bpv], format='mixed', errors='coerce')
                    df_bpv[col_tgl_bayar] = pd.to_datetime(df_bpv[col_tgl_bayar], format='mixed', errors='coerce')
                    
                    st.subheader("🔍 Filter Periode")
                    tgl_filter = st.date_input("Pilih Rentang Tanggal:", value=(today - timedelta(days=3), today), key="filter_tgl_bpv")

                    if len(tgl_filter) == 2:
                        start_date, end_date = tgl_filter
                        
                        bpv_baru = df_bpv[(df_bpv[col_tgl_bpv].dt.date >= start_date) & (df_bpv[col_tgl_bpv].dt.date <= end_date)]
                        bpv_dibayar = df_bpv[(df_bpv[col_tgl_bayar].dt.date >= start_date) & (df_bpv[col_tgl_bayar].dt.date <= end_date)]
                        
                        st.divider()
                        c1, c2 = st.columns(2)
                        c1.metric(f"BPV Masuk ({start_date} s/d {end_date})", f"{len(bpv_baru)} Dokumen")
                        c2.metric(f"BPV Dibayar ({start_date} s/d {end_date})", f"{len(bpv_dibayar)} Dokumen")
                        
                        st.subheader("📥 Rincian BPV Masuk di Periode Ini")
                        if not bpv_baru.empty:
                            kolom_tampil_baru = [col for col in ['PO TRANSAKSI', 'PIC', 'TANGGAL BPV', 'CUSTOMER', 'SUPPLIER', 'TOTAL','NILAI BAYAR'] if col in df_bpv.columns]
                            bpv_baru_disp = format_tgl_internasional(bpv_baru[kolom_tampil_baru] if kolom_tampil_baru else bpv_baru, ["TANGGAL BPV", "TANGGAL BAYAR"])
                            st.dataframe(bpv_baru_disp, use_container_width=True)
                            st.download_button("📥 Download BPV Masuk (CSV)", bpv_baru_disp.to_csv(index=False).encode('utf-8'), "BPV_Masuk.csv", "text/csv", key="dl_bpv1")
                        else:
                            st.info("Tidak ada BPV baru yang masuk pada rentang tanggal ini.")
                            
                        st.write("") 
                        
                        st.subheader("💸 Rincian BPV Dibayar di Periode Ini")
                        if not bpv_dibayar.empty:
                            kolom_tampil_bayar = [col for col in ['PO TRANSAKSI', 'TANGGAL BAYAR', 'CUSTOMER', 'SUPPLIER', 'NILAI BAYAR', 'TOTAL', 'STATUS PO'] if col in df_bpv.columns]
                            bpv_dibayar_disp = format_tgl_internasional(bpv_dibayar[kolom_tampil_bayar] if kolom_tampil_bayar else bpv_dibayar, ["TANGGAL BAYAR", "TANGGAL BPV"])
                            st.dataframe(bpv_dibayar_disp, use_container_width=True)
                            st.download_button("📥 Download BPV Dibayar (CSV)", bpv_dibayar_disp.to_csv(index=False).encode('utf-8'), "BPV_Dibayar.csv", "text/csv", key="dl_bpv2")
                        else:
                            st.info("Tidak ada rekaman BPV yang dibayar pada rentang tanggal ini.")
                    else:
                        st.warning("Silakan pilih rentang tanggal (klik tanggal mulai dan tanggal akhir).")
                else:
                    st.error(f"Kolom tanggal tidak ditemukan. Kolom yang tersedia: {df_bpv.columns.tolist()}")

        except Exception as e:
            st.error(f"Terjadi kesalahan teknis saat memproses data BPV: {e}")

        # --- FITUR TAMBAHAN: TRACKER STATUS PO ---
        st.markdown("---") 
        st.subheader("🧾 Tracker Status PO & BPV")
        st.caption("Cari nomor PO untuk mengecek kelengkapan BPV dan status pembayarannya (Terhubung ke GSheet).")

        @st.cache_data(ttl=60) 
        def load_data_po():
            try:
                url_gsheet = "https://docs.google.com/spreadsheets/d/1fjr-r_FlaAE-WOrHmoC9Ai2-Kxbafzxt1Mr5MciIGOU/export?format=csv&gid=0"
                df_po = pd.read_csv(url_gsheet, header=8) 
                df_po.columns = df_po.columns.str.strip().str.upper() 
                return df_po
            except Exception as e:
                st.error(f"Gagal menarik data PO dari Google Sheets: {e}")
                return pd.DataFrame()

        df_po = load_data_po()

        if df_po.empty:
            st.error("File database PO (Google Sheets) tidak dapat diakses. Pastikan file tidak diprivate.")
        else:
            col_po1, col_po2 = st.columns([4, 1])
            with col_po1:
                input_po = st.text_input("Ketik Nomor PO:", placeholder="Contoh: PO-12345...", key="search_po_input")
            with col_po2:
                st.markdown("<br>", unsafe_allow_html=True)
                btn_cek_po = st.button("Cek Status", use_container_width=True)

            if btn_cek_po or input_po:
                if input_po:
                    kolom_po = "PO TRANSAKSI" 
                    kolom_bayar = "TANGGAL BAYAR"
                    
                    df_po.columns = df_po.columns.str.strip().str.upper()
                    
                    if kolom_po in df_po.columns:
                        hasil_po = df_po[df_po[kolom_po].astype(str).str.contains(input_po, case=False, na=False)]
                        
                        st.markdown("### Hasil Pengecekan:")
                        if not hasil_po.empty:
                            data_terpilih = hasil_po.iloc[0]
                            no_po_asli = data_terpilih[kolom_po]
                            
                            st.success(f"✅ Dokumen ditemukan untuk PO: **{no_po_asli}**")
                            st.info("🟢 **Status BPV: BPV Ada**")
                            
                            sudah_bayar = False
                            if kolom_bayar in data_terpilih:
                                val_bayar = str(data_terpilih[kolom_bayar]).strip().lower()
                                if val_bayar not in ['nan', 'nat', 'none', '', '-', 'null']:
                                    sudah_bayar = True
                                    
                            if sudah_bayar:
                                tgl_info = data_terpilih[kolom_bayar]
                                st.success(f"🟢 **Status Pembayaran: Sudah Terbayar** (Tanggal: {tgl_info})")
                            else:
                                st.warning("🟡 **Status Pembayaran: Belum Terbayar**")
                                
                            st.write("Detail Data:")
                            hasil_po_disp = format_tgl_internasional(hasil_po, ["TANGGAL BAYAR", "TANGGAL BPV"])
                            st.dataframe(hasil_po_disp, use_container_width=True)
                            st.download_button("📥 Download Hasil Pencarian PO (CSV)", hasil_po_disp.to_csv(index=False).encode('utf-8'), f"Tracker_PO_{input_po}.csv", "text/csv", key="dl_po_track")
                        else:
                            st.error(f"Pencarian untuk PO: **{input_po}**")
                            st.error("🔴 **Status BPV: BPV Tidak Ada**")
                            st.error("🔴 **Status Pembayaran: Belum Terbayar** (Karena BPV belum ada)")
                    else:
                        st.error(f"Sistem error: Kolom '{kolom_po}' tidak ditemukan di Google Sheets Anda.")
                else:
                    st.warning("👆 Silakan ketik Nomor PO terlebih dahulu pada kotak pencarian.")

    # --- MENU 7: HISTORY PENJUALAN TERAKHIR ---
    elif menu_pilihan == "History Penjualan Terakhir":
        st.header("🔍 History Penjualan Terakhir")
        st.caption(f"📅 {tanggal_sekarang_str} | Menampilkan data transaksi dengan tanggal terupdate")

        @st.cache_data
        def load_penjualan_terakhir():
            try:
                df = pd.read_excel("data_penjualan_terakhir.xlsx") 
                df.columns = df.columns.str.strip() 
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
            col_search, col_btn = st.columns([4, 1])
            with col_search:
                search_query = st.text_input("Ketik Nama Barang yang dicari:", placeholder="Contoh: Gerobak sorong...", key="search_barang_jual")
            with col_btn:
                st.markdown("<br>", unsafe_allow_html=True)
                tombol_search = st.button("Search", use_container_width=True)

            if tombol_search or search_query:
                if search_query:
                    nama_kolom_barang = "Nama Barang" 
                    if nama_kolom_barang in df_penjualan.columns:
                        df_result = df_penjualan[df_penjualan[nama_kolom_barang].astype(str).str.contains(search_query, case=False, na=False)]

                        if not df_result.empty:
                            kolom_acuan = 'Item Id' if 'Item Id' in df_result.columns else nama_kolom_barang
                            df_terupdate = df_result.sort_values(by=[kolom_acuan, 'Tanggal'], ascending=[True, False])
                            df_terupdate = df_terupdate.drop_duplicates(subset=[kolom_acuan], keep='first')

                            st.success(f"✅ Ditemukan {len(df_terupdate)} jenis barang terupdate untuk: '{search_query}'")
                            
                            kolom_target = [
                                "No Transaksi", "Tanggal", "Customer", "Sales", 
                                "Kode Barang", "Nama Barang", "Harga", "Database"
                            ]
                            kolom_tampil = [col for col in kolom_target if col in df_terupdate.columns]
                            
                            df_hist_disp = format_tgl_internasional(df_terupdate[kolom_tampil], ["Tanggal", "TANGGAL"])
                            st.dataframe(df_hist_disp, use_container_width=True)
                            st.download_button("📥 Download History Penjualan (CSV)", df_hist_disp.to_csv(index=False).encode('utf-8'), "History_Penjualan.csv", "text/csv", key="dl_hist")
                        else:
                            st.warning(f"Belum ada riwayat penjualan untuk barang '{search_query}'.")
                    else:
                        st.error(f"Kolom '{nama_kolom_barang}' tidak ditemukan di file.")
                else:
                    st.info("👆 Silakan ketik nama barang terlebih dahulu.")
                    
    # --- MENU 8: TRACKING VENDOR ---
    elif menu_pilihan == "Tracking Vendor":
        st.header("🏢 Tracking Vendor")
        st.caption(f"📅 {tanggal_sekarang_str} | Sumber: data_po_sbm.xlsx")
        
        try:
            df_vendor = pd.read_excel("data_po_sbm.xlsx")
            df_vendor.columns = df_vendor.columns.str.strip() 
            if "Tanggal" in df_vendor.columns:
                df_vendor["Tanggal"] = pd.to_datetime(df_vendor["Tanggal"], errors="coerce")
        except Exception as e:
            st.error(f"Gagal membaca file 'data_po_sbm.xlsx': {e}")
            st.stop()

        kolom_wajib = ["No Transaksi", "Tanggal", "Supplier", "Nama Barang"]
        kolom_ada = [col for col in kolom_wajib if col in df_vendor.columns]
        
        if len(kolom_ada) < len(kolom_wajib):
            st.warning(f"⚠️ Beberapa kolom tidak ditemukan di database. Kolom yang ada: {', '.join(df_vendor.columns)}")

        st.subheader("🔍 Pencarian Riwayat Vendor")
        
        if "Supplier" in df_vendor.columns:
            col_v1, col_v2 = st.columns(2)
            with col_v1:
                list_supplier = ["Semua"] + sorted(df_vendor["Supplier"].dropna().astype(str).unique().tolist())
                pilih_supplier = st.selectbox("Pilih Supplier:", list_supplier, key="filter_supplier")
            with col_v2:
                cari_barang_vendor = st.text_input("Cari Nama Barang (Opsional):", placeholder="Ketik nama barang...")
            
            df_filter = df_vendor.copy()
            if pilih_supplier != "Semua":
                df_filter = df_filter[df_filter["Supplier"].astype(str) == pilih_supplier]
            if cari_barang_vendor and "Nama Barang" in df_filter.columns:
                df_filter = df_filter[df_filter["Nama Barang"].astype(str).str.contains(cari_barang_vendor, case=False, na=False)]
                
            st.divider()
            if not df_filter.empty:
                st.success(f"✅ Ditemukan {len(df_filter)} riwayat transaksi.")
                
                df_vendor_disp = format_tgl_internasional(df_filter[kolom_ada], ["Tanggal", "TANGGAL"])
                st.dataframe(df_vendor_disp, use_container_width=True)
                st.download_button("📥 Download Riwayat Vendor (CSV)", df_vendor_disp.to_csv(index=False).encode('utf-8'), "Riwayat_Vendor.csv", "text/csv", key="dl_vendor")
            else:
                st.info("Tidak ada data transaksi yang sesuai dengan pencarian Anda.")
        else:
            st.error("Kolom 'Supplier' tidak ditemukan di dalam file Excel.")