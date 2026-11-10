import streamlit as st
import math
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from scipy.interpolate import make_interp_spline
import io
import os
import tempfile
import time
from datetime import datetime
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.chart import ScatterChart, Reference, Series
import folium
from streamlit_folium import st_folium

# ==========================================
# KONFIGURASI HALAMAN & TEMA
# ==========================================
st.set_page_config(page_title="HYDRO PRO - Enterprise Analytics", layout="wide", page_icon="💧")

st.markdown("""
    <style>
    /* Global & Typography */
    .stApp { background-color: #F4F7FB; font-family: 'Plus Jakarta Sans', 'Inter', sans-serif; }
    
    /* Sidebar Deep Navy */
    [data-testid="stSidebar"] {
        background-color: #142D4E !important;
        background-image: linear-gradient(180deg, #142D4E 0%, #0B192C 100%);
        box-shadow: 4px 0 15px rgba(0,0,0,0.05);
    }
    [data-testid="stSidebar"] * { color: #F4F7FB !important; }
    
    /* Cards & Containers */
    .dashboard-card {
        background-color: #FFFFFF; padding: 25px; border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03); border: 1px solid #e2e8f0;
        transition: transform 0.2s ease, box-shadow 0.2s ease; cursor: pointer; height: 100%;
    }
    .dashboard-card:hover {
        transform: translateY(-4px); box-shadow: 0 10px 20px rgba(0,0,0,0.08); border-color: #2478B7;
    }
    
    .input-panel {
        background-color: #FFFFFF; padding: 25px; border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02); border: 1px solid #e2e8f0; margin-bottom: 20px;
    }
    
    /* Metrics */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF; border-left: 5px solid #20A69A; padding: 20px;
        border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.03);
    }
    
    /* Buttons */
    button[kind="primary"] {
        background-color: #2478B7 !important; border: none !important; border-radius: 8px !important;
        font-weight: 600 !important; letter-spacing: 0.5px !important; transition: 0.3s !important;
    }
    button[kind="primary"]:hover { background-color: #142D4E !important; box-shadow: 0 4px 12px rgba(36,120,183,0.3) !important; }
    
    /* Scientific Table */
    .sci-table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 0.9em; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    .sci-table thead tr { background-color: #142D4E; color: #ffffff; text-align: left; }
    .sci-table th, .sci-table td { padding: 12px 15px; border-bottom: 1px solid #e2e8f0; }
    .sci-table tbody tr { background-color: #ffffff; }
    .sci-table tbody tr:nth-of-type(even) { background-color: #f8fafc; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# MEMORI APLIKASI (STATE MANAGEMENT)
# ==========================================
# Navigasi State
if 'current_page' not in st.session_state: st.session_state.current_page = "📊 Dashboard Beranda"

def change_page(page_name):
    st.session_state.current_page = page_name

# Parameter State (Diatur ke None agar form tampil KOSONG di awal)
if 'metode_tc' not in st.session_state: st.session_state.metode_tc = "Kirpich"
if 'L' not in st.session_state: st.session_state.L = None
if 'S' not in st.session_state: st.session_state.S = None
if 'R24' not in st.session_state: st.session_state.R24 = None
if 'tipe_lahan' not in st.session_state: st.session_state.tipe_lahan = "Homogen (1 Jenis Lahan)"
if 'luas_total' not in st.session_state: st.session_state.luas_total = None
if 'c_final' not in st.session_state: st.session_state.c_final = None
if 'a1' not in st.session_state: st.session_state.a1 = None
if 'a2' not in st.session_state: st.session_state.a2 = None
if 'a3' not in st.session_state: st.session_state.a3 = None
if 'c_val1' not in st.session_state: st.session_state.c_val1 = None
if 'c_val2' not in st.session_state: st.session_state.c_val2 = None
if 'c_val3' not in st.session_state: st.session_state.c_val3 = None

# Koordinat Inlet & Outlet (Tetap ada default agar Peta bisa merender lokasi awal)
if 'lat_out' not in st.session_state: st.session_state.lat_out = -7.47050
if 'lon_out' not in st.session_state: st.session_state.lon_out = 110.21780
if 'lat_in' not in st.session_state: st.session_state.lat_in = -7.45000
if 'lon_in' not in st.session_state: st.session_state.lon_in = 110.19000

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #69BFE8 !important; margin-bottom: 0;'>💧 HYDRO PRO</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 0.8rem; letter-spacing: 1.5px; color: #8982D5 !important;'>ENGINEERING PLATFORM</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    menu_options = ["📊 Dashboard Beranda", "📑 Referensi Nilai C", "⚙️ Kalkulator Hidrologi"]
    selected_menu = st.radio(
        "MAIN NAVIGATION", 
        menu_options, 
        index=menu_options.index(st.session_state.current_page)
    )
    if selected_menu != st.session_state.current_page:
        st.session_state.current_page = selected_menu
        st.rerun()
        
    st.markdown("---")
    st.caption("v6.1 Enterprise | Core Validated ✅")

menu = st.session_state.current_page

# ==========================================
# HALAMAN 1: DASHBOARD BERANDA
# ==========================================
if menu == "📊 Dashboard Beranda":
    st.markdown("""
        <div style="background: linear-gradient(135deg, #142D4E 0%, #2478B7 100%); padding: 60px 40px; border-radius: 12px; color: white; margin-bottom: 30px;">
            <h1 style="margin: 0; font-size: 2.8rem; font-weight: 800; letter-spacing: -1px;">HYDRO PRO</h1>
            <h3 style="margin: 10px 0 20px 0; color: #69BFE8; font-weight: 400; font-size: 1.3rem;">Professional Hydrological Analysis & Watershed Calculator</h3>
            <p style="font-size: 1.05rem; opacity: 0.9; max-width: 800px; line-height: 1.6;">
                Platform analisis hidrologi terintegrasi untuk asesmen Daerah Aliran Sungai (DAS). 
                Dilengkapi dengan perhitungan Waktu Konsentrasi (Tc) presisi tinggi, pemetaan Inlet-Outlet, interpolasi hidrograf natural, dan ekspor <i>spreadsheet engineering</i>.
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<h3 style='color: #142D4E;'>Core Engineering Modules</h3>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class="dashboard-card" style="border-top: 4px solid #20A69A;">
            <h3 style="color: #20A69A; margin-top:0;">🗺️ Geospatial Watershed</h3>
            <p style="color: #475569; font-size: 0.95rem;">Tentukan titik <b>Inlet (Hulu)</b> dan <b>Outlet (Hilir)</b> DAS secara interaktif. Sistem akan memvisualisasikan skema arah aliran <i>(flow direction)</i> di atas peta geospasial.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Buka Modul Pemetaan →", key="btn_map"): change_page("⚙️ Kalkulator Hidrologi"); st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("""
        <div class="dashboard-card" style="border-top: 4px solid #8982D5;">
            <h3 style="color: #8982D5; margin-top:0;">📈 Natural Hydrograph</h3>
            <p style="color: #475569; font-size: 0.95rem;">Visualisasi hidrograf menggunakan interpolasi spline matematis untuk menghasilkan kurva <i>rising limb</i> dan <i>recession limb</i> yang mulus dan natural, mempertahankan titik puncak absolut.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Buka Modul Visualisasi →", key="btn_viz"): change_page("⚙️ Kalkulator Hidrologi"); st.rerun()
        
    with col2:
        st.markdown("""
        <div class="dashboard-card" style="border-top: 4px solid #F0A45D;">
            <h3 style="color: #F0A45D; margin-top:0;">⏱️ Concentration Time (Tc)</h3>
            <p style="color: #475569; font-size: 0.95rem;">Komputasi kecepatan aliran menggunakan persamaan empiris <b>Kirpich</b> dan <b>California</b> dengan akurasi <i>floating-point</i> penuh tanpa pembulatan internal.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Buka Modul Kalkulator →", key="btn_calc"): change_page("⚙️ Kalkulator Hidrologi"); st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("""
        <div class="dashboard-card" style="border-top: 4px solid #36A779;">
            <h3 style="color: #36A779; margin-top:0;">📑 Interconnected Reports</h3>
            <p style="color: #475569; font-size: 0.95rem;">Ekspor hasil perhitungan ke format <b>Excel</b> dengan lembar kerja yang saling terhubung (berisi formula Excel asli) untuk keperluan audit teknis dan verifikasi.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Buka Modul Laporan →", key="btn_rep"): change_page("⚙️ Kalkulator Hidrologi"); st.rerun()

# ==========================================
# HALAMAN 2: REFERENSI NILAI C
# ==========================================
elif menu == "📑 Referensi Nilai C":
    st.markdown("<h2 style='color: #142D4E;'>Tabel Referensi Koefisien Pengaliran (C)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #475569; margin-bottom: 20px;'>Standar rasio limpasan permukaan berdasarkan pedoman hidrologi terverifikasi.</p>", unsafe_allow_html=True)
    
    st.markdown("""
    <table class="sci-table">
        <thead>
            <tr>
                <th>Kondisi Permukaan / Tata Guna Lahan</th>
                <th>Rentang C</th>
                <th>Sumber Referensi Utama</th>
            </tr>
        </thead>
        <tbody>
            <tr style="background-color: #e2e8f0; font-weight: bold;"><td colspan="3" style="color: #142D4E;">Kawasan Perkotaan & Bangunan</td></tr>
            <tr><td>Kawasan Bisnis / Perkotaan Padat</td><td>0.70 – 0.95</td><td>ASCE (1992) / NRCS</td></tr>
            <tr><td>Perumahan Padat</td><td>0.60 – 0.75</td><td>FHWA HEC-22 (2001)</td></tr>
            <tr><td>Perumahan Renggang (banyak pekarangan)</td><td>0.40 – 0.60</td><td>FHWA HEC-22 (2001)</td></tr>
            <tr style="background-color: #e2e8f0; font-weight: bold;"><td colspan="3" style="color: #142D4E;">Permukaan Kedap Air (Impermeable)</td></tr>
            <tr><td>Jalan Aspal / Beton / Atap Gedung</td><td>0.85 – 0.95</td><td>Chow, V.T. (1988)</td></tr>
            <tr><td>Jalan Paving / Conblock</td><td>0.70 – 0.85</td><td>SNI 03-3424-1994</td></tr>
            <tr style="background-color: #e2e8f0; font-weight: bold;"><td colspan="3" style="color: #142D4E;">Kawasan Terbuka Hijau & Alami</td></tr>
            <tr><td>Taman / Padang Rumput</td><td>0.10 – 0.30</td><td>ASCE (1992)</td></tr>
            <tr><td>Lahan Pertanian / Sawah</td><td>0.15 – 0.30</td><td>SNI 03-3424-1994</td></tr>
            <tr><td>Hutan Lebat</td><td>0.05 – 0.20</td><td>Chow, V.T. (1988)</td></tr>
        </tbody>
    </table>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div style="background-color: #F4F7FB; border-left: 5px solid #2478B7; padding: 20px; border-radius: 8px; margin-top: 20px;">
        <h4 style="color: #142D4E; margin-top: 0;">📚 Daftar Pustaka (Bibliography)</h4>
        <ul style="color: #475569; font-size: 0.9rem; line-height: 1.6;">
            <li><b>Chow, V.T., Maidment, D.R., & Mays, L.W. (1988).</b> <i>Applied Hydrology</i>. McGraw-Hill.</li>
            <li><b>FHWA (2001).</b> <i>Hydraulic Engineering Circular No. 22 (HEC-22): Urban Drainage Design Manual</i>. US Department of Transportation.</li>
            <li><b>ASCE (1992).</b> <i>Design and Construction of Urban Stormwater Management Systems</i>. American Society of Civil Engineers.</li>
            <li><b>BSN (1994).</b> <i>SNI 03-3424-1994: Tata Cara Perencanaan Drainase Permukaan Jalan</i>. Badan Standardisasi Nasional.</li>
        </ul>
        <p style="font-size: 0.85rem; color: #64748b; font-style: italic; margin-bottom: 0;">Catatan: Pemilihan nilai koefisien aktual harus mempertimbangkan intensitas hujan, kemiringan lereng (slope), dan kelompok hidrologi tanah (Hydrologic Soil Group).</p>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# HALAMAN 3: KALKULATOR UTAMA
# ==========================================
elif menu == "⚙️ Kalkulator Hidrologi":
    st.markdown("<h2 style='color: #142D4E; margin-bottom: 5px;'>Hydrological Analysis Engine</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b; margin-bottom: 25px;'>Sistem komputasi berbasis formula empiris absolut tanpa degradasi presisi.</p>", unsafe_allow_html=True)

    # --- A. LOKASI INLET & OUTLET DAS ---
    st.markdown("<div class='input-panel'>", unsafe_allow_html=True)
    st.markdown("<h4 style='color: #2478B7; margin-top: 0;'>📍 Geospatial Watershed Boundaries (Inlet - Outlet)</h4>", unsafe_allow_html=True)
    
    col_map1, col_map2 = st.columns([1, 2.5])
    with col_map1:
        st.markdown("**Titik Hulu (Inlet)**")
        st.session_state.lat_in = st.number_input("Lat Inlet [°]", value=st.session_state.lat_in, format="%.5f")
        st.session_state.lon_in = st.number_input("Lon Inlet [°]", value=st.session_state.lon_in, format="%.5f")
        st.markdown("**Titik Hilir (Outlet)**")
        st.session_state.lat_out = st.number_input("Lat Outlet [°]", value=st.session_state.lat_out, format="%.5f")
        st.session_state.lon_out = st.number_input("Lon Outlet [°]", value=st.session_state.lon_out, format="%.5f")
        st.caption("Peta memvisualisasikan indikasi *flow direction* secara skematis.")
    with col_map2:
        center_lat = (st.session_state.lat_in + st.session_state.lat_out) / 2
        center_lon = (st.session_state.lon_in + st.session_state.lon_out) / 2
        m = folium.Map(location=[center_lat, center_lon], zoom_start=13, tiles="OpenStreetMap")
        
        folium.Marker([st.session_state.lat_in, st.session_state.lon_in], tooltip="INLET (Hulu)", icon=folium.Icon(color="green", icon="arrow-down")).add_to(m)
        folium.Marker([st.session_state.lat_out, st.session_state.lon_out], tooltip="OUTLET (Hilir)", icon=folium.Icon(color="red", icon="flag")).add_to(m)
        
        folium.PolyLine(
            locations=[[st.session_state.lat_in, st.session_state.lon_in], [st.session_state.lat_out, st.session_state.lon_out]],
            color="#2478B7", weight=4, opacity=0.8, dash_array="10",
            tooltip="Indikasi Arah Aliran (Skematis)"
        ).add_to(m)
        st_folium(m, height=300, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # --- B. PARAMETER HIDROLOGI (Input Kosong / None) ---
    col_kiri, col_kanan = st.columns(2)
    
    with col_kiri:
        st.markdown("<div class='input-panel'>", unsafe_allow_html=True)
        st.markdown("<h4 style='color: #2478B7; margin-top: 0;'>📐 Channel & Rainfall Parameters</h4>", unsafe_allow_html=True)
        
        idx_tc = ["Kirpich", "California"].index(st.session_state.metode_tc)
        st.session_state.metode_tc = st.selectbox("Algoritma Waktu Konsentrasi (Tc)", ["Kirpich", "California"], index=idx_tc)
        
        st.session_state.L = st.number_input("Panjang Saluran Utama (L) [m]", min_value=0.001, value=st.session_state.L, placeholder="Masukkan angka...")
        st.session_state.S = st.number_input("Kemiringan Saluran Rata-rata (S) [m/m]", min_value=0.00001, value=st.session_state.S, format="%.5f" if st.session_state.S is not None else None, placeholder="Masukkan angka...")
        st.session_state.R24 = st.number_input("Curah Hujan Rancangan 24 Jam (R24) [mm]", min_value=0.1, value=st.session_state.R24, placeholder="Masukkan angka...")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_kanan:
        st.markdown("<div class='input-panel' style='height: 100%;'>", unsafe_allow_html=True)
        st.markdown("<h4 style='color: #2478B7; margin-top: 0;'>🌿 Land Use Characteristics (C)</h4>", unsafe_allow_html=True)
        
        idx_lahan = ["Homogen (1 Jenis Lahan)", "Campuran (Hitung C-Komposit)"].index(st.session_state.tipe_lahan)
        st.session_state.tipe_lahan = st.radio("Kondisi Tata Guna Lahan:", ["Homogen (1 Jenis Lahan)", "Campuran (Hitung C-Komposit)"], index=idx_lahan)
        
        luas_total_calc = 0.0
        c_final_calc = 0.0
        
        if st.session_state.tipe_lahan == "Homogen (1 Jenis Lahan)":
            st.session_state.luas_total = st.number_input("Luas Total DAS (A) [km²]", min_value=0.001, value=st.session_state.luas_total, placeholder="Masukkan angka...")
            st.session_state.c_final = st.number_input("Koefisien Pengaliran (C)", min_value=0.01, max_value=1.0, value=st.session_state.c_final, placeholder="Masukkan angka...")
            luas_total_calc = st.session_state.luas_total if st.session_state.luas_total else 0.0
            c_final_calc = st.session_state.c_final if st.session_state.c_final else 0.0
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.session_state.a1 = st.number_input("Luas Area 1 [km²]", min_value=0.0, value=st.session_state.a1, placeholder="Opsional...")
                st.session_state.a2 = st.number_input("Luas Area 2 [km²]", min_value=0.0, value=st.session_state.a2, placeholder="Opsional...")
                st.session_state.a3 = st.number_input("Luas Area 3 [km²]", min_value=0.0, value=st.session_state.a3, placeholder="Opsional...")
            with c2:
                st.session_state.c_val1 = st.number_input("Nilai C Area 1", min_value=0.0, max_value=1.0, value=st.session_state.c_val1, placeholder="Opsional...")
                st.session_state.c_val2 = st.number_input("Nilai C Area 2", min_value=0.0, max_value=1.0, value=st.session_state.c_val2, placeholder="Opsional...")
                st.session_state.c_val3 = st.number_input("Nilai C Area 3", min_value=0.0, max_value=1.0, value=st.session_state.c_val3, placeholder="Opsional...")
            
            # Penanganan nilai None agar tidak error saat dihitung
            a1_v = st.session_state.a1 or 0.0
            a2_v = st.session_state.a2 or 0.0
            a3_v = st.session_state.a3 or 0.0
            c1_v = st.session_state.c_val1 or 0.0
            c2_v = st.session_state.c_val2 or 0.0
            c3_v = st.session_state.c_val3 or 0.0
            
            luas_total_calc = a1_v + a2_v + a3_v
            if luas_total_calc > 0:
                c_final_calc = ((a1_v * c1_v) + (a2_v * c2_v) + (a3_v * c3_v)) / luas_total_calc
                st.info(f"Geometri Komposit: Luas Total = **{luas_total_calc:.4f} km²** | C-Eq = **{c_final_calc:.4f}**")
        st.markdown("</div>", unsafe_allow_html=True)

    # --- C. DOKUMENTASI MATEMATIS ---
    with st.expander("🔬 Math Formula Documentation (Strict Compliance)"):
        st.markdown("""
        **1. Persamaan Koefisien Komposit (C-Composite)**
        $$C_{komposit} = \\frac{\\sum_{i=1}^{n} (A_i \\times C_i)}{A_{total}}$$
        
        **2. Persamaan Waktu Konsentrasi (Kirpich Method)**
        $$t_c = 0.0195 \\times L^{0.77} \\times S^{-0.385}$$
        *(L = Panjang saluran utama dalam meter, S = Kemiringan rata-rata dalam m/m)*
        
        **3. Persamaan Waktu Konsentrasi (California Method)**
        $$t_c = \\left( \\frac{0.87 \\times L_{km}^3}{H} \\right)^{0.385}$$
        *(L_{km} = Panjang saluran dalam km, H = Beda elevasi dalam meter)*
        
        **4. Intensitas Curah Hujan (Mononobe Method)**
        $$I = \\frac{R_{24}}{24} \\times \\left( \\frac{24}{t_c} \\right)^{2/3}$$
        
        **5. Persamaan Debit Puncak (Metode Rasional)**
        $$Q_p = 0.278 \\times C \\times I \\times A$$
        """)

    # --- D. EKSEKUSI PERHITUNGAN (Dengan Validasi Data Kosong) ---
    st.markdown("---")
    if st.button("🚀 EXECUTE SCIENTIFIC ANALYSIS", type="primary", use_container_width=True):
        
        # Validasi keamanan form agar tidak crash jika data masih kosong
        if st.session_state.L is None or st.session_state.S is None or st.session_state.R24 is None:
            st.error("⚠️ SYSTEM HALTED: Harap isi semua **Parameter Saluran & Hujan** (L, S, R24) terlebih dahulu.")
        elif st.session_state.tipe_lahan == "Homogen (1 Jenis Lahan)" and (st.session_state.luas_total is None or st.session_state.c_final is None):
            st.error("⚠️ SYSTEM HALTED: Harap isi **Luas Total DAS** dan **Koefisien Pengaliran**.")
        elif luas_total_calc <= 0:
            st.error("⚠️ SYSTEM HALTED: Total Luas Area DAS tidak boleh nol. Harap masukkan area minimal pada satu jenis lahan.")
        else:
            with st.spinner('Memproses algoritma hidrologi presisi penuh...'):
                time.sleep(0.5) 
                
                # VARIABEL ASLI (Precision Preserved)
                metode_tc = st.session_state.metode_tc
                L = float(st.session_state.L)
                S = float(st.session_state.S)
                R24 = float(st.session_state.R24)
                
                # LOGIKA PERHITUNGAN ASLI (DILARANG KERAS DIUBAH)
                if metode_tc == "Kirpich":
                    tc_menit = 0.0195 * math.pow(L, 0.77) * math.pow(S, -0.385)
                else: 
                    tc_jam_cal = math.pow((0.87 * ((L/1000)**3) / (L*S)), 0.385)
                    tc_menit = tc_jam_cal * 60

                tc_jam = tc_menit / 60
                I = (R24 / 24) * math.pow((24 / tc_jam), 2/3)
                Q = 0.278 * c_final_calc * I * luas_total_calc

                # Visual Logic
                if tc_menit < 60: respon, hex_color = "Fast Response", "#e74c3c"
                elif tc_menit <= 360: respon, hex_color = "Medium Response", "#f39c12"
                else: respon, hex_color = "Slow Response", "#2478B7"

                # --- DASHBOARD HASIL METRIK ---
                st.markdown("<h3 style='color: #142D4E;'>📊 Analytical Results Dashboard</h3>", unsafe_allow_html=True)
                
                met1, met2, met3, met4 = st.columns(4)
                met1.metric("Waktu Konsentrasi (Tc)", f"{tc_menit:.3f} Menit")
                met2.metric("Intensitas Hujan (I)", f"{I:.3f} mm/jam")
                met3.metric("Debit Puncak (Qp)", f"{Q:.3f} m³/s")
                met4.metric("Karakteristik Respon", respon)
                
                # --- GRAFIK HIDROGRAF (INTERPOLASI NATURAL) ---
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("<div class='input-panel'>", unsafe_allow_html=True)
                st.markdown("<h4 style='color: #142D4E; margin-top: 0; text-align: center;'>Smoothed Synthetic Unit Hydrograph</h4>", unsafe_allow_html=True)
                st.caption("<div style='text-align:center;'>*Visual smoothing diterapkan tanpa mengubah titik puncak komputasi (Qp dan Tc absolut).*</div>", unsafe_allow_html=True)
                
                fig, ax = plt.subplots(figsize=(10, 4.5), dpi=200)
                fig.patch.set_facecolor('#FFFFFF')
                ax.set_facecolor('#F4F7FB')
                
                x_control = np.array([0, tc_menit*0.5, tc_menit, tc_menit*1.5, tc_menit*3])
                y_control = np.array([0, Q*0.35, Q, Q*0.45, 0]) 
                
                spline = make_interp_spline(x_control, y_control, k=2)
                x_smooth = np.linspace(0, tc_menit*3, 200)
                y_smooth = spline(x_smooth)
                y_smooth[y_smooth < 0] = 0 
                
                ax.plot(x_smooth, y_smooth, color=hex_color, linewidth=3, zorder=3)
                ax.fill_between(x_smooth, y_smooth, color=hex_color, alpha=0.15, zorder=2)
                
                ax.plot(tc_menit, Q, marker='o', markersize=8, color='#142D4E', zorder=4)
                ax.axvline(x=tc_menit, color='#64748b', linestyle='--', alpha=0.6, linewidth=1.5, zorder=1)
                ax.text(tc_menit * 1.03, Q * 0.4, f'Time to Peak: {tc_menit:.2f} mnt', color='#475569', fontsize=9, rotation=90)
                
                ax.annotate(f' Peak: {Q:.2f} m³/s', xy=(tc_menit, Q), xytext=(tc_menit + (tc_menit*0.1), Q),
                            arrowprops=dict(facecolor='#142D4E', shrink=0.05, width=1, headwidth=5),
                            fontsize=10, fontweight='bold', color='#142D4E', zorder=5)
                
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['left'].set_color('#cbd5e1')
                ax.spines['bottom'].set_color('#cbd5e1')
                ax.set_xlabel("Time (Minutes)", fontsize=10, fontweight='bold', color='#475569')
                ax.set_ylabel("Discharge (m³/s)", fontsize=10, fontweight='bold', color='#475569')
                ax.grid(color='#e2e8f0', linestyle='-', linewidth=0.8, zorder=0)
                
                st.pyplot(fig)
                st.markdown("</div>", unsafe_allow_html=True)

                # ==========================================
                # INTERCONNECTED EXCEL EXPORT & PDF
                # ==========================================
                st.markdown("<h4 style='color: #142D4E; margin-top: 20px;'>🖨️ Engineering Reports Export</h4>", unsafe_allow_html=True)
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                    fig.savefig(tmpfile.name, bbox_inches='tight', dpi=200)
                    tmpfile_path = tmpfile.name

                # --- EXCEL GENERATION (OPENPYXL FORMULAS) ---
                wb = openpyxl.Workbook()
                
                ws_input = wb.active
                ws_input.title = "1_INPUT_DATA"
                input_headers = ["Parameter", "Simbol", "Nilai", "Satuan"]
                ws_input.append(input_headers)
                ws_input.append(["Metode Tc", "Metode", metode_tc, "-"])
                ws_input.append(["Panjang Sungai Utama", "L", L, "m"])
                ws_input.append(["Kemiringan Saluran", "S", S, "m/m"])
                ws_input.append(["Curah Hujan Maksimum", "R24", R24, "mm"])
                ws_input.append(["Luas Total DAS", "A", luas_total_calc, "km2"])
                ws_input.append(["Koefisien Pengaliran", "C", c_final_calc, "-"])
                
                for col in ["A", "B", "C", "D"]: ws_input.column_dimensions[col].width = 25
                for cell in ws_input["1:1"]: cell.font = Font(bold=True)
                
                ws_tc = wb.create_sheet("2_PERHITUNGAN_TC")
                ws_tc.append(["Deskripsi", "Formula Excel yang Digunakan", "Hasil Evaluasi", "Satuan"])
                
                if metode_tc == "Kirpich":
                    ws_tc.append(["Tc Kirpich (Menit)", "=(0.0195*(1_INPUT_DATA!C3^0.77))*(1_INPUT_DATA!C4^-0.385)"])
                    ws_tc["C2"] = f"=(0.0195*(1_INPUT_DATA!C3^0.77))*(1_INPUT_DATA!C4^-0.385)"
                else:
                    ws_tc.append(["Tc California (Jam)", "=(0.87*((1_INPUT_DATA!C3/1000)^3)/(1_INPUT_DATA!C3*1_INPUT_DATA!C4))^0.385"])
                    ws_tc["C2"] = f"=(0.87*((1_INPUT_DATA!C3/1000)^3)/(1_INPUT_DATA!C3*1_INPUT_DATA!C4))^0.385"
                    ws_tc.append(["Tc California (Menit)", "=C2 * 60"])
                    ws_tc["C3"] = f"=C2*60"
                
                ws_tc.column_dimensions['A'].width = 30
                ws_tc.column_dimensions['B'].width = 75
                ws_tc.column_dimensions['C'].width = 20
                for cell in ws_tc["1:1"]: cell.font = Font(bold=True)
                
                ws_q = wb.create_sheet("3_DEBIT_PUNCAK")
                ws_q.append(["Parameter", "Formula", "Nilai Evaluasi", "Satuan"])
                
                cell_ref_tc = "C2" if metode_tc == "Kirpich" else "C3"
                
                ws_q.append(["Tc Jam", f"=2_PERHITUNGAN_TC!{cell_ref_tc} / 60"])
                ws_q["C2"] = f"=2_PERHITUNGAN_TC!{cell_ref_tc}/60"
                ws_q["D2"] = "jam"
                
                ws_q.append(["Intensitas (I)", "=(1_INPUT_DATA!C5/24)*((24/C2)^(2/3))"])
                ws_q["C3"] = f"=(1_INPUT_DATA!C5/24)*((24/C2)^(2/3))"
                ws_q["D3"] = "mm/jam"
                
                ws_q.append(["Debit Puncak (Q)", "=0.278 * 1_INPUT_DATA!C7 * C3 * 1_INPUT_DATA!C6"])
                ws_q["C4"] = f"=0.278 * 1_INPUT_DATA!C7 * C3 * 1_INPUT_DATA!C6"
                ws_q["D4"] = "m3/s"
                
                ws_q.column_dimensions['A'].width = 25
                ws_q.column_dimensions['B'].width = 50
                ws_q.column_dimensions['C'].width = 20
                for cell in ws_q["1:1"]: cell.font = Font(bold=True)

                ws_hydro = wb.create_sheet("4_HIDROGRAF")
                ws_hydro.append(["Time (Menit)", "Discharge (m3/s)"])
                for idx, (t_val, q_val) in enumerate(zip(x_smooth, y_smooth), start=2):
                    ws_hydro.cell(row=idx, column=1, value=float(t_val))
                    ws_hydro.cell(row=idx, column=2, value=float(q_val))
                
                chart = ScatterChart()
                chart.title = "Synthetic Smoothed Hydrograph"
                chart.style = 10
                chart.x_axis.title = "Time (Minutes)"
                chart.y_axis.title = "Discharge (m3/s)"
                xvalues = Reference(ws_hydro, min_col=1, min_row=2, max_row=len(x_smooth)+1)
                yvalues = Reference(ws_hydro, min_col=2, min_row=2, max_row=len(y_smooth)+1)
                series = Series(yvalues, xvalues, title="Discharge")
                chart.series.append(series)
                ws_hydro.add_chart(chart, "D2")
                
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                
                # --- PDF EXPORT ---
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", style="B", size=16)
                pdf.cell(0, 10, txt="HYDROLOGICAL ANALYSIS REPORT - HYDRO PRO", ln=True, align="C")
                pdf.set_font("Arial", size=10)
                pdf.cell(0, 6, txt=f"Generated Date: {datetime.now().strftime('%d %B %Y - %H:%M')}", ln=True, align="C")
                pdf.ln(5)
                
                pdf.set_font("Arial", style="B", size=11)
                pdf.cell(0, 8, txt="1. Technical Parameters & Location:", ln=True)
                pdf.set_font("Arial", size=10)
                pdf.cell(0, 5, txt=f"- Inlet  (Lat, Lon) : {st.session_state.lat_in}, {st.session_state.lon_in}", ln=True)
                pdf.cell(0, 5, txt=f"- Outlet (Lat, Lon) : {st.session_state.lat_out}, {st.session_state.lon_out}", ln=True)
                pdf.cell(0, 5, txt=f"- Tc Method         : {metode_tc}", ln=True)
                pdf.cell(0, 5, txt=f"- Length (L)        : {L} meter", ln=True)
                pdf.cell(0, 5, txt=f"- Slope (S)         : {S} m/m", ln=True)
                pdf.cell(0, 5, txt=f"- Area (A)          : {luas_total_calc:.4f} km2", ln=True)
                pdf.cell(0, 5, txt=f"- Coefficient (C)   : {c_final_calc:.4f}", ln=True)
                pdf.cell(0, 5, txt=f"- Rainfall (R24)    : {R24} mm", ln=True)
                pdf.ln(5)
                
                pdf.set_font("Arial", style="B", size=11)
                pdf.cell(0, 8, txt="2. Mathematical Results (Verified):", ln=True)
                pdf.set_font("Arial", size=10)
                pdf.cell(0, 5, txt=f"- Time of Concentration (Tc) : {tc_menit:.5f} minutes", ln=True)
                pdf.cell(0, 5, txt=f"- Rainfall Intensity (I)     : {I:.5f} mm/hour", ln=True)
                pdf.cell(0, 5, txt=f"- Peak Discharge (Qp)        : {Q:.5f} m3/second", ln=True)
                pdf.ln(5)
                
                pdf.set_font("Arial", style="B", size=11)
                pdf.cell(0, 8, txt="3. Visual Hydrograph Interpolation:", ln=True)
                pdf.image(tmpfile_path, x=15, w=180)
                
                pdf_bytes = pdf.output()
                os.remove(tmpfile_path)

                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    st.download_button("📊 Unduh Laporan Audit Excel (.xlsx)", data=excel_buffer.getvalue(), file_name="HydroPro_Audit.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with btn_col2:
                    st.download_button("📑 Unduh Executive Summary PDF (.pdf)", data=bytes(pdf_bytes), file_name="HydroPro_Summary.pdf", mime="application/pdf", use_container_width=True)