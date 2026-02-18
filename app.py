import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter", layout="wide")

st.title("🚢 Manning Deployment Converter")
st.write("Upload file Excel untuk menghasilkan 3 kolom: **No ITV, No ID, dan Nama Operator**.")

def process_data(file):
    # Baca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Berdasarkan data Anda:
    # Nama/ID ada di baris r
    # No ITV ada di baris r + 1 (tepat di bawahnya)
    
    for r in range(len(df) - 1):
        for c in [1, 3, 5]: # Kolom B, D, F
            try:
                # 1. Ambil ID & Nama dari baris saat ini
                id_val = df.iloc[r, c]
                nama_val = df.iloc[r, c+1]
                
                # 2. Ambil ITV dari baris DI BAWAHNYA
                itv_val = df.iloc[r+1, c]
                
                # Kriteria Validasi:
                # - ITV harus berupa angka
                # - Nama bukan 'N' atau kosong
                if pd.notna(itv_val) and str(itv_val).replace('.0','').isdigit():
                    if pd.notna(nama_val) and str(nama_val).strip().upper() != 'N':
                        # Validasi tambahan: ID biasanya angka 4 digit
                        if pd.notna(id_val) and len(str(id_val)) >= 4:
                            clean_rows.append({
                                "No ITV": int(float(itv_val)),
                                "No ID": id_val,
                                "Nama Operator": str(nama_val).strip()
                            })
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Menghapus duplikat jika ada
        result_df = result_df.drop_duplicates()
        
        st.success(f"Ditemukan {len(result_df)} data operator.")
        st.subheader("Preview Data (Sudah Sesuai Gambar)")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Download Button
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="📥 Download Excel (3 Kolom)",
            data=output.getvalue(),
            file_name="Manning_Clean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tetap tidak terbaca. Mohon pastikan nomor ITV ada di sel tepat di bawah No ID.")
