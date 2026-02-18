import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Manning Deployment", layout="wide")

# Custom CSS untuk tampilan kartu (Card)
st.markdown("""
    <style>
    .stApp { background-color: #f4f7f6; }
    .card {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #007bff;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
        margin-bottom: 10px;
    }
    .dermaga-title {
        background-color: #0e1117;
        color: white;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 15px;
        font-size: 18px;
        text-align: center;
    }
    .unit-label { color: #666; font-size: 12px; margin-bottom: 2px; }
    .name-label { font-weight: bold; font-size: 16px; color: #1f1f1f; }
    .id-label { color: #007bff; font-size: 13px; }
    </style>
    """, unsafe_allow_html=True)

def load_data():
    # Membaca CSV (Sesuaikan nama file jika perlu)
    df = pd.read_csv("rekap_manning_deployment.xlsx - Rekap_Manning.csv", header=None)
    return df

def display_personnel(id_val, name):
    if pd.notna(name) and str(name).strip() != "" and str(name).upper() != 'N':
        st.markdown(f"""
            <div class="card">
                <div class="unit-label">ITV Unit</div>
                <div class="name-label">{name}</div>
                <div class="id-label">ID: {int(id_val) if isinstance(id_val, (int, float)) else id_val}</div>
            </div>
            """, unsafe_allow_html=True)

# Judul Utama
st.title("🚢 Real-time Manning Deployment")
st.write("Daftar Penempatan Personil per Dermaga")

try:
    df = load_data()
    
    # Mencari Shift Leader (Contoh: M. EFENDI dari baris 11 di CSV Anda)
    shift_leader = "M. EFENDI" # Ini bisa didinamiskan dari pencarian string di DF
    
    col_info1, col_info2 = st.columns(2)
    col_info1.info(f"👤 **Shift Leader Berth:** {shift_leader}")
    col_info2.success("📅 **Status:** Operasional Aktif")

    st.markdown("---")

    # Membuat 3 Kolom untuk Dermaga
    col1, col2, col3 = st.columns(3)

    # Logika Pengelompokan Berdasarkan Struktur CSV Anda
    # Dermaga 1: Baris 2 ke bawah
    with col1:
        st.markdown('<div class="dermaga-title">📍 Dermaga 1 KOTA HIDAYAH</div>', unsafe_allow_html=True)
        # Ambil data dari kolom 1 & 2 (index 1, 2) dan 3 & 4 (index 3, 4)
        d1_data = df.iloc[2:12] # Range baris untuk Dermaga 1
        for _, row in d1_data.iterrows():
            display_personnel(row[1], row[2])
            display_personnel(row[3], row[4])

    # Dermaga 2: Baris 14 ke bawah
    with col2:
        st.markdown('<div class="dermaga-title">📍 Dermaga 2 INTERASIA ENGAGE</div>', unsafe_allow_html=True)
        d2_data = df.iloc[14:24] 
        for _, row in d2_data.iterrows():
            display_personnel(row[1], row[2])
            display_personnel(row[3], row[4])

    # Dermaga 4: Baris 25 ke bawah
    with col3:
        st.markdown('<div class="dermaga-title">📍 Dermaga 4 XIN YAN TAI</div>', unsafe_allow_html=True)
        d4_data = df.iloc[25:35]
        for _, row in d4_data.iterrows():
            display_personnel(row[1], row[2])
            display_personnel(row[3], row[4])
            display_personnel(row[5], row[6]) # Dermaga 4 punya 3 kolom personil di CSV

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
    st.warning("Pastikan file CSV 'rekap_manning_deployment.xlsx - Rekap_Manning.csv' sudah ada di folder yang sama.")
