import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Manning Data Fixer", layout="wide")

st.title("📑 Manning to Clean Excel")
st.write("Upload file untuk merapikan data menjadi kolom: ITV, No ID, dan Nama Operator.")

def process_data(file):
    # Membaca data tanpa header agar koordinat sel murni
    df = pd.read_excel(file, header=None)
    
    rows_list = []
    
    # Looping baris untuk mencari ITV (Angka ITV biasanya ada di baris genap/ganjil tertentu)
    # Kita akan memindai baris demi baris
    for r in range(len(df) - 1):
        for c in [1, 3, 5]:  # Kolom B, D, dan F (Tempat ITV dan ID berada)
            val = df.iloc[r, c]
            
            # Cek jika sel saat ini adalah angka ITV (biasanya 3 digit seperti 238, 245, dll)
            if pd.notna(val) and isinstance(val, (int, float)) and 100 <= val <= 999:
                itv_no = int(val)
                
                # Mengambil ID dan NAMA dari baris tepat DI BAWAHNYA (r + 1)
                id_op = df.iloc[r+1, c]
                nama_op = df.iloc[r+1, c+1] # Nama ada di sebelah kanan ID
                
                # Validasi: Masukkan jika ada Nama dan bukan "N"
                if pd.notna(nama_op) and str(nama_op).strip().upper() != "N":
                    rows_list.append({
                        "No ITV": itv_no,
                        "No ID": id_op,
                        "Nama Operator": nama_op
                    })
                    
    return pd.DataFrame(rows_list)

uploaded_file = st.file_uploader("Upload File Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    result_df = process_data(uploaded_file)
    
    if not result_df.empty:
        st.subheader("Hasil Ekstraksi (Sesuai Gambar)")
        # Tampilkan tabel di Streamlit
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Clean_Data')
        
        st.download_button(
            label="📥 Download Hasil Excel",
            data=output.getvalue(),
            file_name="Manning_Cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data tidak ditemukan. Pastikan format baris ITV dan Nama sesuai dengan template.")
