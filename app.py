import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter", layout="wide")

st.title("🚢 Manning Deployment to Clean Excel")
st.write("Aplikasi ini mendeteksi pola: **Baris ITV** di atas, lalu **ID & Nama** tepat di baris bawahnya.")

def process_data(file):
    # Baca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    results = []
    
    # Scan seluruh baris dan kolom B(2), D(4), F(6) - disesuaikan dengan index pandas
    # Di file Anda, ITV ada di kolom index 2, 4, 6 (Kolom C, E, G di Excel)
    target_cols = [2, 4, 6] 

    for r in range(len(df) - 1):
        for c in target_cols:
            val_itv = df.iloc[r, c]
            
            # 1. Cek apakah sel ini berisi nomor ITV (Angka 200-400an)
            if pd.notna(val_itv) and str(val_itv).replace('.0','').isdigit():
                itv_num = int(float(val_itv))
                
                # Saring agar hanya angka unit ITV yang diambil (misal 100 - 999)
                if 100 <= itv_num <= 999:
                    
                    # 2. Ambil ID & Nama tepat di baris bawahnya (r+1)
                    val_id = df.iloc[r+1, c]
                    val_nama = df.iloc[r+1, c+1]
                    
                    if pd.notna(val_id) and pd.notna(val_nama):
                        nama_str = str(val_nama).strip().upper()
                        # Lewati jika isinya cuma 'N' (Kosong)
                        if nama_str not in ["N", "", "NAMA PERSONIL"]:
                            results.append({
                                "ITV": itv_num,
                                "ID": str(val_id).replace('.0',''),
                                "Nama Operator": nama_str
                            })

    return pd.DataFrame(results)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    df_clean = process_data(uploaded_file)
    
    if not df_clean.empty:
        # Hapus duplikat dan urutkan berdasarkan ITV
        df_clean = df_clean.drop_duplicates().sort_values(by="ITV")
        
        st.success(f"Ditemukan {len(df_clean)} data operator.")
        st.dataframe(df_clean, use_container_width=True, hide_index=True)
        
        # Download Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_clean.to_excel(writer, index=False, sheet_name='Data_Manning')
        
        st.download_button(
            label="📥 Download Excel (3 Kolom)",
            data=output.getvalue(),
            file_name="Manning_Cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak terbaca. Pastikan format ITV di atas dan ID/Nama di bawahnya.")
