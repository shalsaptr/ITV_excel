import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter Fix", layout="wide")

st.title("🚢 Manning Deployment Converter")

def process_data(file):
    # Membaca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Scan seluruh baris
    for r in range(len(df) - 1):
        # Cek kolom B(1), D(3), F(5) untuk mencari Nomor ITV
        for c in [1, 3, 5]: 
            try:
                itv_val = df.iloc[r, c]
                id_val = df.iloc[r+1, c]
                nama_val = df.iloc[r+1, c+1]
                
                # LOGIKA: 
                # Jika sel saat ini (r) adalah angka ITV (200-400an)
                # Dan sel di bawahnya (r+1) adalah No ID
                if pd.notna(itv_val) and str(itv_val).replace('.0','').isdigit():
                    itv_num = int(float(itv_val))
                    
                    # Pastikan ini memang area data (ITV biasanya 3 digit)
                    if 100 <= itv_num <= 999:
                        if pd.notna(id_val) and pd.notna(nama_val):
                            nama_str = str(nama_val).strip().upper()
                            if nama_str not in ["N", "", "NAMA PERSONIL"]:
                                clean_rows.append({
                                    "ITV": itv_num,
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
        # Urutkan berdasarkan No ITV agar rapi
        result_df = result_df.sort_values(by="ITV").drop_duplicates()
        
        st.success(f"Berhasil menarik {len(result_df)} data operator.")
        st.subheader("Hasil Sesuai Permintaan (3 Kolom)")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Simpan ke Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="📥 Download Excel Rapi",
            data=output.getvalue(),
            file_name="Manning_Rapi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak terbaca. Pastikan format kolom ITV berada tepat di atas ID.")
