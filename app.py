import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Full Converter", layout="wide")

st.title("🚢 Manning Deployment Converter (Full Data)")
st.write("Mengekstrak data dari Kolom B ke kanan. Mengabaikan teks gangguan seperti 'Shift Leader'.")

def process_data(file):
    # Membaca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Kita mulai scan dari Kolom B (index 1) untuk mengabaikan Kolom A (index 0)
    # Berdasarkan file Anda, ITV ada di kolom index 1, 3, 5, 7
    itv_cols = [1, 3, 5, 7] 

    for r in range(len(df)):
        for c in itv_cols:
            try:
                val_itv = df.iloc[r, c]
                
                # 1. CARI NOMOR ITV (Harus angka 3 digit)
                if pd.notna(val_itv) and str(val_itv).replace('.0','').isdigit():
                    itv_num = str(val_itv).replace('.0','')
                    
                    if 100 <= int(itv_num) <= 999:
                        # 2. CARI ID & NAMA di baris-baris bawahnya (Cek hingga 10 baris ke bawah)
                        # Ini untuk melompati baris "SHIFT LEADER", "SUWARNO", atau baris kosong
                        for search_r in range(r + 1, min(r + 11, len(df))):
                            id_val = df.iloc[search_r, c]     # ID ada di bawah ITV (Kolom yang sama)
                            nama_val = df.iloc[search_r, c+1] # Nama ada di samping kanan ID
                            
                            if pd.notna(id_val) and pd.notna(nama_val):
                                id_str = str(id_val).replace('.0','').strip()
                                nama_str = str(nama_val).strip().upper()
                                
                                # Kriteria: ID harus angka & Nama bukan teks sampah
                                if id_str.isdigit() and nama_str not in ["N", "", "NAMA PERSONIL", "UAT"]:
                                    # Pastikan ID bukan nomor ITV yang sama
                                    if id_str != itv_num:
                                        clean_rows.append({
                                            "ITV": itv_num,
                                            "ID": id_str,
                                            "Nama Operator": nama_str
                                        })
                                        break # Ketemu pasangan ID, lanjut cari ITV berikutnya
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Manning (Contoh: 16 JANUARI SHIFT 1D.xlsx)", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Hapus duplikat dan urutkan berdasarkan ITV
        result_df = result_df.drop_duplicates(subset=['ITV', 'ID']).sort_values(by="ITV")
        
        st.success(f"✅ Berhasil! Mengekstrak {len(result_df)} data operator dari seluruh dermaga.")
        
        # Tabel Preview
        st.subheader("Preview Output (3 Kolom)")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Clean_Data')
        
        st.download_button(
            label="📥 Download Excel Hasil",
            data=output.getvalue(),
            file_name="Manning_Cleaned_Full.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak ditemukan. Pastikan format file sudah sesuai.")
