import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Data Converter", layout="wide")

st.title("🚢 Manning Deployment Converter")
st.write("Aplikasi ini akan mengubah format laporan manning yang berantakan menjadi tabel Excel 3 kolom yang rapi.")

def process_data(file):
    # Membaca file excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Kita scan baris demi baris
    # Berdasarkan file Anda, ITV ada di baris r, sedangkan ID & Nama ada di baris r+1
    for r in range(len(df) - 1):
        # Kolom tempat data berada: B(1), D(3), F(5)
        for c in [1, 3, 5]: 
            try:
                itv_val = df.iloc[r, c]
                id_val = df.iloc[r+1, c]
                nama_val = df.iloc[r+1, c+1]
                
                # Kriteria pengambilan: 
                # 1. itv_val harus angka (Nomor ITV)
                # 2. id_val harus ada (Nomor ID)
                # 3. nama_val bukan 'N' atau kosong
                if pd.notna(itv_val) and str(itv_val).isdigit():
                    if pd.notna(nama_val) and str(nama_val).strip().upper() != 'N':
                        clean_rows.append({
                            "No ITV": int(itv_val),
                            "No ID": id_val,
                            "Nama Operator": str(nama_val).strip()
                        })
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning Anda", type=["xlsx"])

if uploaded_file:
    with st.spinner('Sedang memproses...'):
        result_df = process_data(uploaded_file)
        
        if not result_df.empty:
            st.success(f"Berhasil mengekstrak {len(result_df)} data!")
            
            # Tampilkan Tabel Hasil
            st.subheader("Preview Output (3 Kolom)")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            
            # Persiapan Download Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Data_Manning_Rapi')
            
            st.download_button(
                label="📥 Download Hasil Excel (.xlsx)",
                data=output.getvalue(),
                file_name="Manning_Deployment_Rapi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Gagal mendeteksi data. Pastikan Anda mengupload file yang benar sesuai format gambar.")
            st.info("Tips: Pastikan nomor ITV berada tepat di atas sel ID Operator.")
