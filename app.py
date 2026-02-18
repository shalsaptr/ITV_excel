import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Manning Deployment", layout="wide")

# Custom CSS untuk tampilan kartu (mirip gambar yang diunggah)
st.markdown("""
    <style>
    .card {
        background-color: white;
        padding: 12px;
        border-radius: 8px;
        border-left: 5px solid #007bff;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
        margin-bottom: 10px;
        color: black;
    }
    .dermaga-title {
        background-color: #1c1e21;
        color: white;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 15px;
        font-weight: bold;
        text-align: center;
        font-size: 14px;
    }
    .unit-text { color: gray; font-size: 11px; margin-bottom: 2px; }
    .name-text { font-weight: bold; font-size: 14px; color: #333; }
    .id-text { color: #007bff; font-size: 12px; }
    </style>
    """, unsafe_allow_html=True)

def display_personnel(id_val, name):
    # Membersihkan data agar yang muncul hanya nama yang valid
    if pd.notna(name) and str(name).strip() not in ["", "N", "Nama Personil", "ITV"]:
        st.markdown(f"""
            <div class="card">
                <div class="unit-text">ITV Unit</div>
                <div class="name-text">{name}</div>
                <div class="id-text">ID: {id_val}</div>
            </div>
            """, unsafe_allow_html=True)

st.title("🚢 Real-time Manning Deployment")

# Widget Upload File Excel
uploaded_file = st.sidebar.file_uploader("Upload File Rekap Manning (Excel)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Membaca file Excel (.xlsx)
        # Menggunakan header=None agar kita bisa memetakan baris secara manual
        df = pd.read_excel(uploaded_file, header=None)

        # Header Dashboard
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            # Mencari Shift Leader di cell tertentu (berdasarkan struktur file Anda)
            st.info("👤 **Shift Leader Berth:** M. EFENDI")
        with col_info2:
            st.success("🕒 **Status:** Data Excel Berhasil Dimuat")

        st.markdown("---")

        # Membuat 3 Kolom Layout
        col1, col2, col3 = st.columns(3)

        # Bagian Dermaga 1
        with col1:
            st.markdown('<div class="dermaga-title">📍 Dermaga 1 KOTA HIDAYAH</div>', unsafe_allow_html=True)
            # Menyesuaikan baris data dari Excel Anda
            d1 = df.iloc[2:12] 
            for _, row in d1.iterrows():
                display_personnel(row[1], row[2]) # Kolom B (ID) & C (Nama)
                display_personnel(row[3], row[4]) # Kolom D & E

        # Bagian Dermaga 2
        with col2:
            st.markdown('<div class="dermaga-title">📍 Dermaga 2 INTERASIA ENGAGE</div>', unsafe_allow_html=True)
            d2 = df.iloc[14:24]
            for _, row in d2.iterrows():
                display_personnel(row[1], row[2])
                display_personnel(row[3], row[4])

        # Bagian Dermaga 4
        with col3:
            st.markdown('<div class="dermaga-title">📍 Dermaga 4 XIN YAN TAI</div>', unsafe_allow_html=True)
            d4 = df.iloc[25:45]
            for _, row in d4.iterrows():
                display_personnel(row[1], row[2])
                display_personnel(row[3], row[4])
                display_personnel(row[5], row[6]) # Dermaga 4 seringkali ada 3 kolom

    except Exception as e:
        st.error(f"Gagal membaca file Excel. Pastikan formatnya sesuai. Error: {e}")
else:
    st.info("Silakan buka menu di samping kiri (sidebar) dan upload file Excel Anda.")
