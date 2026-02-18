import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Manning Deployment Dashboard", layout="wide")

st.title("🚢 Real-time Manning Deployment")
st.markdown("---")

# Load Data
@st.cache_data
def load_data():
    # Load CSV dan bersihkan baris yang benar-benar kosong
    df = pd.read_csv("rekap_manning_deployment.xlsx - Rekap_Manning.csv")
    return df

try:
    df = load_data()

    # Layouting: Menampilkan data dalam bentuk kartu atau tabel per Dermaga
    # Kita akan melakukan filter sederhana untuk mendemonstrasikan output
    
    # Header Info
    col_header1, col_header2 = st.columns(2)
    with col_header1:
        st.info("**Shift Leader Berth:** M. EFENDI")
    with col_header2:
        st.success("**Status Operasional:** Active")

    # Membuat Grid untuk Dermaga (Mirip tampilan gambar)
    dermagas = ["Dermaga 1 KOTA HIDAYAH", "Dermaga 2 INTERASIA ENGAGE", "Dermaga 4 XIN YAN TAI"]
    
    cols = st.columns(len(dermagas))

    for idx, dermaga in enumerate(dermagas):
        with cols[idx]:
            st.subheader(f"📍 {dermaga}")
            
            # Logika filter sederhana berdasarkan data CSV Anda
            # Mencari baris yang mengandung nama dermaga
            mask = df.iloc[:, 0].str.contains(dermaga, na=False)
            start_idx = df[mask].index[0]
            
            # Mengambil 10 baris setelah nama dermaga untuk personil
            sub_df = df.iloc[start_idx:start_idx+10, [1, 2, 3, 4]] 
            sub_df.columns = ['ID', 'Nama Personil', 'ID_2', 'Nama_2']
            
            # Menampilkan data dengan styling
            for i, row in sub_df.iterrows():
                if pd.notna(row['Nama Personil']):
                    st.markdown(f"""
                    <div style="border: 1px solid #e6e9ef; padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #f9f9f9;">
                        <small style="color: gray;">ITV Unit</small><br>
                        <strong>{row['Nama Personil']}</strong><br>
                        <small>ID: {row['ID']}</small>
                    </div>
                    """, unsafe_allow_html=True)

except Exception as e:
    st.error(f"Gagal memuat data: {e}")
    st.info("Pastikan file CSV berada di folder yang sama dengan app.py")

# Footer
st.markdown("---")
st.caption("Manning Deployment System v1.0")
