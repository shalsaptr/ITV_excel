import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Data Fixer", layout="wide")

st.title("🚢 Manning Deployment Converter")
st.write("Aplikasi ini merapikan format Excel Manning menjadi 3 kolom: **ITV, ID, dan Nama Operator**.")

def process_data(file):
    # Membaca excel tanpa header agar koordinat sel murni
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Looping baris
    # Berdasarkan data: Nama/ID di baris 'r', ITV di baris 'r+1'
    for r in range(len(df) - 1):
        # Kolom B (1), D (3), F (5) adalah letak No ID
        for c in [1, 3, 5]: 
            try:
                id_val = df.iloc[r, c]
                nama_val = df.iloc[r, c+1] # Nama di sebelah kanan ID
                itv_val = df.iloc[r+1, c]  # ITV di bawah ID
                
                # Validasi:
                # 1. ID harus ada (biasanya angka 4 digit)
                # 2. Nama bukan 'N' atau kosong
                # 3. ITV harus ada
                if pd.notna(id_val) and pd.notna(nama_val) and pd.notna(itv_val):
                    name_str = str(nama_val).strip().upper()
                    if name_str not in ["", "N", "NAMA PERSONIL", "NAMA OPERATOR"]:
                        # Pastikan itv_val adalah angka
                        if str(itv_val).replace('.0','').isdigit():
                            clean_rows.append({
                                "ITV": str(itv_val).replace('.0',''),
                                "ID": str(id_val).replace('.0',''),
                                "Nama Operator": name_str
                            })
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    with st.spinner('Memproses data...'):
        result_df = process_data(uploaded_file)
        
        if not result_df.empty:
            # Hapus duplikat untuk memastikan data bersih
            result_df = result_df.drop_duplicates()
            
            st.success(f"Berhasil mengekstrak {len(result_df)} data operator.")
            
            # Preview Tabel
            st.subheader("Preview Output")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            
            # Tombol Download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Hasil_Manning')
            
            st.download_button(
                label="📥 Download Excel (3 Kolom)",
                data=output.getvalue(),
                file_name="Manning_Clean_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Data tidak ditemukan. Pastikan format file sesuai dengan contoh (Nama/ID di atas, ITV di bawah).")
