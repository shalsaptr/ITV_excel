import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Cleaner - Manning Deployment", layout="wide")

st.title("📊 Manning Data Extractor")
st.markdown("Upload file Excel manning untuk dikonversi menjadi daftar ITV, ID, dan Nama yang rapi.")

# --- Fungsi Processing Data ---
def process_manning_data(file):
    # Baca excel tanpa header karena formatnya tidak standar
    df = pd.read_excel(file, header=None)
    
    extracted_data = []

    # Iterasi per baris untuk mengambil data
    # Sesuai gambar/file: ID ada di kolom 1,3,5 dan Nama ada di kolom 2,4,6
    # No ITV biasanya ada di baris tepat di atas baris Nama (atau di kolom tertentu)
    
    for r in range(len(df)):
        for c in [1, 3, 5]: # Kolom B, D, F (ID)
            try:
                id_val = df.iloc[r, c]
                name_val = df.iloc[r, c+1] # Kolom C, E, G (Nama)
                
                # Ambil No ITV (biasanya ada di sel atasnya atau sel ID itu sendiri di baris berbeda)
                # Berdasarkan pola file Anda: ITV ada di baris sebelum Nama/ID
                itv_val = df.iloc[r-1, c] 

                # Validasi: Hanya ambil jika Nama tidak kosong dan bukan "N" atau "Nama Operator"
                if pd.notna(name_val) and str(name_val).strip() not in ["", "N", "Nama Personil", "Nama Operator"]:
                    extracted_data.append({
                        "No ITV": itv_val if pd.notna(itv_val) else "N/A",
                        "No ID": id_val,
                        "Nama Operator": name_val
                    })
            except:
                continue

    return pd.DataFrame(extracted_data)

# --- UI Streamlit ---
uploaded_file = st.file_uploader("Pilih file Excel Rekap Manning", type=["xlsx"])

if uploaded_file:
    with st.spinner('Sedang memproses data...'):
        try:
            # Jalankan pemrosesan
            result_df = process_manning_data(uploaded_file)
            
            if not result_df.empty:
                st.success(f"Berhasil mengekstrak {len(result_df)} data operator!")
                
                # Tampilkan Preview
                st.subheader("Preview Data Bersih")
                st.dataframe(result_df, use_container_width=True)

                # --- Tombol Download Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Data_Manning')
                
                processed_data = output.getvalue()

                st.download_button(
                    label="📥 Download Hasil (.xlsx)",
                    data=processed_data,
                    file_name="Data_Manning_Clean.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Data tidak ditemukan. Pastikan format file sesuai.")
                
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

else:
    st.info("Menunggu file diupload...")
