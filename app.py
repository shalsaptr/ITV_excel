import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Full Extractor", layout="wide")

st.title("🚢 Manning Deployment Converter (Full Data)")
st.write("Membaca data mulai dari Kolom 2 (B) ke kanan. Mengabaikan baris teks gangguan.")

def process_data(file):
    # Membaca excel tanpa header
    df = pd.read_excel(file, header=None)
    
    clean_rows = []
    
    # Fokus pada kolom tempat ITV/Nama berada: Index 2, 4, 6 (Kolom C, E, G)
    # Kolom ID ada di sebelah kirinya: Index 1, 3, 5 (Kolom B, D, F)
    target_cols = [2, 4, 6] 

    for r in range(len(df)):
        for c in target_cols:
            try:
                val_itv = df.iloc[r, c]
                
                # 1. CARI NOMOR ITV (Harus angka 3 digit)
                if pd.notna(val_itv) and str(val_itv).replace('.0','').isdigit():
                    itv_num = str(val_itv).replace('.0','')
                    
                    if 100 <= int(itv_num) <= 999:
                        # 2. CARI ID & NAMA di baris-baris bawahnya (Mencari hingga 8 baris ke bawah)
                        # Ini gunanya untuk melompati baris "SHIFT LEADER", "SUWARNO", dll.
                        for search_r in range(r + 1, min(r + 9, len(df))):
                            id_val = df.iloc[search_r, c-1] # Kolom kiri (ID)
                            nama_val = df.iloc[search_r, c]   # Kolom yang sama (Nama)
                            
                            if pd.notna(id_val) and pd.notna(nama_val):
                                id_str = str(id_val).replace('.0','').strip()
                                nama_str = str(nama_val).strip().upper()
                                
                                # Kriteria: ID harus mengandung angka & Nama bukan 'N' atau 'SHIFT LEADER'
                                if any(char.isdigit() for char in id_str) and nama_str not in ["N", "", "NAMA PERSONIL", "UAT", "SHIFT LEADER BERTH"]:
                                    # Pastikan bukan mengambil ID yang isinya malah nomor ITV lagi
                                    if id_str != itv_num:
                                        clean_rows.append({
                                            "ITV": itv_num,
                                            "ID": id_str,
                                            "Nama Operator": nama_str
                                        })
                                        break # Sudah ketemu pasangan ID/Nama, lanjut cari ITV berikutnya
            except:
                continue
                
    return pd.DataFrame(clean_rows)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        # Hapus duplikat dan urutkan berdasarkan No ITV
        result_df = result_df.drop_duplicates(subset=['ITV', 'ID']).sort_values(by="ITV")
        
        st.success(f"✅ Berhasil! Ditemukan {len(result_df)} data operator (Dermaga 1 s/d 4).")
        
        # Tampilkan Tabel Preview
        st.subheader("Preview Output (3 Kolom)")
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Simpan ke Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Data_Manning')
        
        st.download_button(
            label="📥 Download Excel Full Data",
            data=output.getvalue(),
            file_name="Manning_Lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak terbaca. Pastikan format kolom sudah sesuai.")
