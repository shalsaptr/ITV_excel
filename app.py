import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter Full Fix", layout="wide")

st.title("🚢 Manning Deployment Full Extractor")
st.write("Mengekstrak seluruh data: Dermaga 1, 2, 3, dan 4 tanpa terpotong.")

def process_data(file):
    # Baca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Target Kolom ITV: index 2, 4, 6 (Kolom C, E, G)
    itv_cols = [2, 4, 6] 

    for r in range(len(df) - 1):
        for c in itv_cols:
            val_itv = df.iloc[r, c]
            
            # 1. CEK APAKAH INI ANGKA ITV (3 Digit)
            if pd.notna(val_itv) and str(val_itv).replace('.0','').isdigit():
                itv_num = int(float(val_itv))
                
                if 100 <= itv_num <= 999:
                    # 2. CARI ID & NAMA DI BAWAHNYA
                    # Kita cek sampai 4 baris ke bawah untuk jaga-jaga ada jeda baris kosong
                    for offset in range(1, 5):
                        if r + offset >= len(df): break
                        
                        id_val = df.iloc[r + offset, c - 1] # ID di kiri (Kolom B, D, F)
                        nama_val = df.iloc[r + offset, c]   # Nama di kolom yang sama (C, E, G)
                        
                        # Validasi: ID harus angka dan Nama bukan 'N' atau judul
                        if pd.notna(id_val) and pd.notna(nama_val):
                            nama_str = str(nama_val).strip().upper()
                            id_str = str(id_val).replace('.0','').strip()
                            
                            if nama_str not in ["N", "", "NAMA PERSONIL", "UAT"] and id_str.isdigit():
                                clean_rows.append({
                                    "ITV": itv_num,
                                    "ID": id_str,
                                    "Nama Operator": nama_str
                                })
                                break # Sudah ketemu pasangan untuk ITV ini, lanjut ke ITV berikutnya

    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Hapus duplikat dan urutkan
        result_df = result_df.drop_duplicates(subset=['ITV', 'ID']).sort_values(by="ITV")
        
        st.success(f"Berhasil menarik {len(result_df)} data operator. Data sudah melewati baris 'Darmadi' dan 'Suwarno'.")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Data_Clean')
        
        st.download_button(
            label="📥 Download Excel Full Data",
            data=output.getvalue(),
            file_name="Manning_Full_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak ditemukan. Pastikan format file benar.")
