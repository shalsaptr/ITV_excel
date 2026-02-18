import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Full Extractor", layout="wide")

st.title("🚢 Manning Deployment Converter (Full Data)")
st.info("Logika: Mengambil ITV, ID, dan Nama. Mengabaikan teks 'Shift Leader', 'Dermaga', dll.")

def process_data(file):
    # Baca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Kolom ITV/Nama: index 2, 4, 6 (C, E, G)
    # Kolom ID: index 1, 3, 5 (B, D, F)
    target_cols = [2, 4, 6] 

    for r in range(len(df)):
        for c in target_cols:
            val_itv = df.iloc[r, c]
            
            # 1. CARI NOMOR ITV (Angka bulat 3 digit)
            if pd.notna(val_itv) and str(val_itv).replace('.0','').isdigit():
                itv_num = str(val_itv).replace('.0','')
                
                # Saring hanya angka unit (100-999)
                if 100 <= int(itv_num) <= 999:
                    
                    # 2. CARI ID & NAMA di baris-baris bawahnya
                    # Kita scan sampai 10 baris ke bawah untuk melewati teks gangguan (Shift Leader, dll)
                    for search_r in range(r + 1, min(r + 10, len(df))):
                        id_raw = df.iloc[search_r, c-1] # Kolom ID (Kiri)
                        nama_raw = df.iloc[search_r, c]   # Kolom Nama (Sama dengan ITV)
                        
                        if pd.notna(id_raw) and pd.notna(nama_raw):
                            id_str = str(id_raw).replace('.0','').strip()
                            nama_str = str(nama_raw).strip().upper()
                            
                            # KRITERIA PENTING:
                            # ID harus mengandung angka (bukan teks seperti 'SUWARNO')
                            # Nama bukan 'N' atau 'NAMA PERSONIL'
                            has_digit = any(char.isdigit() for char in id_str)
                            
                            if has_digit and nama_str not in ["N", "", "NAMA PERSONIL", "UAT"]:
                                clean_rows.append({
                                    "ITV": itv_num,
                                    "ID": id_str,
                                    "Nama Operator": nama_str
                                })
                                break # Ketemu! Lanjut cari ITV berikutnya.

    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    with st.spinner('Menyisir seluruh data dari Dermaga 1 sampai 4...'):
        result_df = process_data(uploaded_file)
        
        if not result_df.empty:
            # Hapus duplikat agar data benar-benar bersih
            result_df = result_df.drop_duplicates(subset=['ITV', 'ID']).sort_values(by="ITV")
            
            st.success(f"✅ Berhasil! Ditemukan {len(result_df)} data operator.")
            
            # Preview tabel hasil
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            
            # Download Button
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Data_Manning_Rapi')
            
            st.download_button(
                label="📥 Download Excel Hasil (Lengkap)",
                data=output.getvalue(),
                file_name="Manning_Full_Clean.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Data tidak ditemukan. Pastikan format file sesuai.")
