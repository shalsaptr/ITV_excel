import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter Fix", layout="wide")

st.title("🚢 Manning Deployment Converter")
st.write("Format: **ITV** (Baris Atas), **ID** (Baris Bawah, Kolom Sama), **Nama** (Baris Bawah, Samping ID)")

def process_data(file):
    # Membaca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Berdasarkan file Anda, data ID ada di kolom B(1), D(3), F(5)
    # Data Nama ada di sampingnya: C(2), E(4), G(6)
    target_id_cols = [1, 3, 5] 
    
    for r in range(len(df) - 1):
        for c in target_id_cols:
            try:
                # 1. Ambil ITV dari baris saat ini (r)
                itv_val = df.iloc[r, c+1] # Di file Anda, ITV sering di kolom Nama tapi baris atas
                # Jika tidak ketemu, cek di kolom ID-nya sendiri (r, c)
                if not str(itv_val).replace('.0','').isdigit():
                    itv_val = df.iloc[r, c]

                # 2. Ambil ID dan Nama dari baris BAWAHNYA (r+1)
                id_val = df.iloc[r+1, c]
                nama_val = df.iloc[r+1, c+1]
                
                # Validasi: ITV harus angka, ID harus ada, Nama bukan 'N'
                if pd.notna(itv_val) and str(itv_val).replace('.0','').isdigit():
                    if pd.notna(id_val) and pd.notna(nama_val):
                        nama_str = str(nama_val).strip().upper()
                        
                        if nama_str not in ["N", "", "NAMA PERSONIL", "UAT"]:
                            clean_rows.append({
                                "ITV": str(itv_val).replace('.0',''),
                                "ID": str(id_val).replace('.0',''),
                                "Nama Operator": nama_str
                            })
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Menghapus baris yang ID-nya sama dengan ITV (mencegah salah tangkap)
        result_df = result_df[result_df['ITV'] != result_df['ID']]
        result_df = result_df.drop_duplicates().sort_values(by="ITV")
        
        st.success(f"Ditemukan {len(result_df)} data operator.")
        st.subheader("Preview Output (3 Kolom)")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(label="📥 Download Excel Rapi", 
                           data=output.getvalue(), 
                           file_name="Manning_Rapi.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Data tidak ditemukan. Pastikan format file sesuai gambar.")
