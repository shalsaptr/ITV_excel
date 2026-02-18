import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Fixer", layout="wide")

st.title("🚢 Manning Deployment Converter")

def process_data(file):
    # Membaca excel tanpa header agar index kolom murni (0, 1, 2...)
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Berdasarkan data mentah Anda:
    # ITV (245) ada di Baris 1, Kolom index 2 (Kolom C)
    # ID (7094) ada di Baris 2, Kolom index 1 (Kolom B)
    # Nama (TRI BAGUS) ada di Baris 2, Kolom index 2 (Kolom C)
    
    # Kolom tempat ITV dan Nama berada (C, E, G)
    itv_name_cols = [2, 4, 6] 

    for r in range(len(df) - 1):
        for c in itv_name_cols:
            try:
                # 1. Ambil ITV dari baris r (Kolom C, E, atau G)
                itv_val = df.iloc[r, c]
                
                # Cek apakah sel tersebut berisi nomor ITV (misal 245)
                if pd.notna(itv_val) and str(itv_val).replace('.0','').isdigit():
                    itv_num = str(itv_val).replace('.0','')
                    
                    if 100 <= int(itv_num) <= 999:
                        # 2. Ambil ID dari baris r+1 (Satu kolom ke KIRI dari Nama)
                        id_val = df.iloc[r+1, c-1] 
                        
                        # 3. Ambil NAMA dari baris r+1 (Kolom yang SAMA dengan ITV tadi)
                        nama_val = df.iloc[r+1, c]
                        
                        if pd.notna(id_val) and pd.notna(nama_val):
                            nama_clean = str(nama_val).strip().upper()
                            
                            # Filter sampah data
                            if nama_clean not in ["N", "", "NAMA PERSONIL", "UAT"]:
                                clean_rows.append({
                                    "ITV": itv_num,
                                    "ID": str(id_val).replace('.0',''),
                                    "Nama Operator": nama_clean
                                })
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Hapus duplikat dan tampilkan
        result_df = result_df.drop_duplicates()
        
        st.success(f"Ditemukan {len(result_df)} data operator.")
        st.subheader("Preview Output (3 Kolom)")
        st.table(result_df.head(10)) # Gunakan table agar terlihat jelas per kolomnya
        
        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Data_Manning')
        
        st.download_button(
            label="📥 Download Excel Rapi",
            data=output.getvalue(),
            file_name="Manning_Sesuai_Gambar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak terbaca. Pastikan format sel sudah benar.")
