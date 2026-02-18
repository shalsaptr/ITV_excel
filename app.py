import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Converter", layout="wide")

st.title("🚢 Manning Deployment Converter")
st.write("Format: **ITV** (Atas), **ID** (Bawah ITV), **Nama** (Samping Kanan ID)")

def process_data(file):
    # Baca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    results = []
    
    # Kolom target tempat ITV dan ID berada: B (index 2), D (index 4), F (index 6)
    # Berdasarkan file CSV Anda, data mulai dari kolom index 2, 4, 6
    target_cols = [2, 4, 6] 

    for r in range(len(df) - 1):
        for c in target_cols:
            val_itv = df.iloc[r, c]
            
            # 1. Cek apakah ini sel ITV (berupa angka)
            if pd.notna(val_itv) and str(val_itv).replace('.0','').isdigit():
                itv_num = str(val_itv).replace('.0','')
                
                # Saring angka yang masuk akal sebagai ITV (misal > 100)
                if len(itv_num) >= 2:
                    
                    # 2. Ambil ID di bawahnya (baris r+1, kolom yang sama c)
                    # 3. Ambil Nama di bawahnya (baris r+1, kolom sebelah kanan c+1)
                    val_id = df.iloc[r+1, c]
                    val_nama = df.iloc[r+1, c+1]
                    
                    if pd.notna(val_id) and pd.notna(val_nama):
                        nama_str = str(val_nama).strip().upper()
                        
                        # Validasi agar bukan data sampah
                        if nama_str not in ["N", "", "NAMA PERSONIL", "UAT"]:
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
        # Urutkan berdasarkan ITV
        df_clean = df_clean.drop_duplicates().sort_values(by="ITV")
        
        st.success(f"Ditemukan {len(df_clean)} data operator.")
        st.subheader("Preview Output (3 Kolom)")
        st.dataframe(df_clean, use_container_width=True, hide_index=True)
        
        # Download Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_clean.to_excel(writer, index=False, sheet_name='Manning_Clean')
        
        st.download_button(
            label="📥 Download Hasil Excel",
            data=output.getvalue(),
            file_name="Manning_Rapi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak ditemukan. Cek kembali format file Anda.")
