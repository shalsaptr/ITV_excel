import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter Fix", layout="wide")

st.title("🚢 Manning Deployment Converter")

def process_data(file):
    # Membaca excel tanpa header agar koordinat kolom murni
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Berdasarkan file Anda, ITV ada di baris ganjil dan data di baris genap (atau sebaliknya)
    # Kita scan kolom B(2), D(4), F(6) berdasarkan koordinat file Anda
    # Catatan: Kolom B di Excel sering terbaca sebagai index 1 atau 2 di pandas tergantung format
    # Kita scan semua kolom agar tidak meleset
    
    for r in range(len(df) - 1):
        for c in range(len(df.columns) - 1):
            itv_candidate = df.iloc[r, c]
            
            # Ciri ITV: Angka bulat 3 digit (245, 257, dll)
            if pd.notna(itv_candidate) and str(itv_candidate).replace('.0','').isdigit():
                itv_str = str(itv_candidate).replace('.0','')
                
                if 100 <= int(itv_str) <= 999:
                    # AMBIL ID (Tepat di bawah ITV)
                    id_val = df.iloc[r+1, c]
                    # AMBIL NAMA (Di sebelah kanan ID)
                    nama_val = df.iloc[r+1, c+1]
                    
                    if pd.notna(id_val) and pd.notna(nama_val):
                        nama_clean = str(nama_val).strip().upper()
                        
                        # Validasi agar bukan teks sampah
                        if nama_clean not in ["N", "", "NAMA PERSONIL", "UAT"]:
                            clean_rows.append({
                                "ITV": itv_str,
                                "ID": str(id_val).replace('.0',''),
                                "Nama Operator": nama_clean
                            })
            
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Hapus duplikat dan pastikan ID bukan angka ITV yang tertangkap dua kali
        result_df = result_df[result_df['ITV'] != result_df['ID']]
        result_df = result_df.drop_duplicates()
        
        st.success(f"Ditemukan {len(result_df)} data operator.")
        st.subheader("Preview Output (3 Kolom)")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Data_Manning')
        
        st.download_button(
            label="📥 Download Excel Rapi",
            data=output.getvalue(),
            file_name="Manning_Rapi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak terbaca. Pastikan format: ITV (Atas), ID (Bawah ITV), Nama (Samping ID).")
