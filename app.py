import streamlit as st
import pandas as pd
from io import BytesIO

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

def AUTO_RKM(data):
    required_cols = {'Instansi', 'Topik', 'Channel', 'Status'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"File Excel harus mengandung kolom: {required_cols}")

    data_rkm = data[list(required_cols)]
    data_selesai = data_rkm[data_rkm['Status'] == 'Selesai']

    rkm = pd.DataFrame()
    rkm['Instansi'] = data_selesai['Instansi'].unique()
    rkm['Topik'] = rkm['Instansi'].apply(
        lambda x: list(set(data_selesai[data_selesai['Instansi'] == x]['Topik']))
    )
    rkm['Topik'] = rkm['Topik'].apply(lambda x: ', '.join(x))
    
    rkm['Jumlah'] = rkm['Instansi'].apply(
        lambda x: len(data_selesai[data_selesai['Instansi'] == x])
    )
    rkm['Ditindaklanjuti'] = rkm['Jumlah']
    rkm['Belum Ditindaklanjuti'] = rkm['Jumlah'] - rkm['Ditindaklanjuti']

    rkm = rkm[['Instansi', 'Jumlah', 'Ditindaklanjuti', 'Belum Ditindaklanjuti', 'Topik']]
    rkm = rkm.sort_values(by='Jumlah', ascending=False).reset_index(drop=True)

    kategori_keluhan = pd.DataFrame()
    kategori_keluhan = data_selesai['Channel'].value_counts().reset_index()
    kategori_keluhan.columns = ['Channel', 'Jumlah']

    return rkm, kategori_keluhan


st.title("Analisis RKM dari File Excel")
st.markdown("""
✅ **Pastikan file memenuhi kriteria berikut:**
- 📄 Format file: `.xlsx`
- 📊 Kolom wajib: `Instansi`, `Topik`, `Channel`, `Status`
""")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        data = pd.read_excel(uploaded_file)
        rkm, kategori = AUTO_RKM(data)

        st.subheader("Hasil RKM")
        st.dataframe(rkm)

        st.subheader("Visualisasi Jumlah Keluhan per Instansi (Top 5)")
        top5_rkm = rkm.nlargest(5, 'Jumlah') 
        top5_rkm_sorted = top5_rkm.set_index('Instansi').sort_values(by='Jumlah', ascending=False)
        st.bar_chart(top5_rkm_sorted['Jumlah'], use_container_width=True)

        st.subheader("Jumlah Keluhan berdasarkan Kategori")
        st.dataframe(kategori)

        # Tombol download untuk hasil rkm
        excel_rkm = to_excel(rkm)
        st.download_button(
            label="Download Hasil RKM (.xlsx)",
            data=excel_rkm,
            file_name='hasil_rkm.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        # Tombol download untuk kategori keluhan
        excel_kategori = to_excel(kategori)
        st.download_button(
            label="Download Kategori Keluhan (.xlsx)",
            data=excel_kategori,
            file_name='kategori_keluhan.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except ValueError as ve:
        st.error(f"Error: {ve}")
    except Exception as e:
        st.error("Terjadi kesalahan saat memproses file. Pastikan file Excel valid dan formatnya benar.")
        st.error(f"Detail error: {e}")
else:
    st.info("Silakan upload file Excel terlebih dahulu.")
