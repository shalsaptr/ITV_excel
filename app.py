import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter Full", layout="wide")

st.title("🚢 Manning Deployment Converter (Full Data)")
st.write("Mengekstrak seluruh data dari Dermaga 1 sampai Dermaga 4.")

def process_data(file):
    # Membaca excel tanpa header agar index kolom murni
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Kolom tempat ITV berada (C, E, G) -> Index 2, 4, 6
    # Kolom tempat ID berada (B, D, F) -> Index 1, 3, 5
    # Kolom tempat Nama berada (C, E, G) -> Index 2, 4, 6
    target_cols = [2, 4, 6] 

    for r in range(len(df)):
        for c in target_cols:
            val = df.iloc[r, c]
            
            # 1. CARI NOMOR ITV (Angka 3 digit antara 100-500)
            if pd.notna(val) and str(val).replace('.0','').isdigit():
                itv_num = int(float(val))
                
                if 100 <= itv_num <= 600:
                    # 2. CARI ID & NAMA di baris-baris bawahnya (Maksimal cek 5 baris ke bawah)
                    # Ini kunci agar data di bawah tulisan "Dermaga" atau "CC" tetap ketemu
                    found = False
                    for search_r in range(r + 1, min(r + 6, len(df))):
                        id_val = df.iloc[search_r, c-1] # ID di sebelah kiri Nama
                        nama_val = df.iloc[search_r, c]   # Nama sejajar dengan kolom ITV
                        
                        if pd.notna(id_val) and pd.notna(nama_val):
                            nama_str = str(nama_val).strip().upper()
                            # Filter sampah data
                            if nama_str not in ["N", "", "NAMA PERSONIL", "UAT", "ITV"]:
                                # Pastikan ID bukan angka ITV yang sama
                                if str(id_val) != str(itv_num):
                                    clean_rows.append({
                                        "ITV": itv_num,
                                        "ID": str(id_val).replace('.0',''),
                                        "Nama Operator": nama_str
                                    })
                                    found = True
                                    break
                    if found: continue

    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    with st.spinner('Menarik seluruh data...'):
        result_df = process_data(uploaded_file)
        
        if not result_df.empty:
            # Hapus duplikat dan urutkan
            result_df = result_df.drop_duplicates(subset=['ITV', 'ID']).sort_values(by="ITV")
            
            st.success(f"Berhasil menarik {len(result_df)} data operator (Dermaga 1 - Dermaga 4).")
            
            # Preview Data
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            
            # Download Button
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            st.download_button(
                label="📥 Download Excel Full Data",
                data=output.getvalue(),
                file_name="Manning_Deployment_Full.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Data tidak ditemukan. Periksa kembali format file Anda.")
