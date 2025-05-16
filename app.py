import streamlit as st
import altair as alt
import pandas as pd
import time
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

def vis_kecamatan(data):
    required_cols = {'Kecamatan', 'Instansi','Status'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")

    data_kecamatan = data[list(required_cols)]
    selesai = data_kecamatan[data_kecamatan['Status'] == 'Selesai']

    rkm_kecamatan = selesai[selesai['Instansi'].str.startswith('Kecamatan')]
    rkm_kecamatan = rkm_kecamatan['Instansi'].value_counts().reset_index()
    rkm_kecamatan.columns = ['Instansi', 'Jumlah']

    rkm_kecamatan = rkm_kecamatan.sort_values(by='Jumlah', ascending=False).head(5)

    # Grafik batang
    bars_kecamatan = alt.Chart(rkm_kecamatan).mark_bar(
        cornerRadiusTopLeft=5,
        cornerRadiusTopRight=5
    ).encode(
        x=alt.X('Instansi:N', sort='-y', title='Kecamatan'),
        y=alt.Y('Jumlah:Q', title='Jumlah Keluhan'),
        color=alt.Color('Instansi:N', legend=None, scale=alt.Scale(scheme='category20')),
        tooltip=['Instansi', 'Jumlah']
    )

    # Label jumlah
    text_kecamatan = alt.Chart(rkm_kecamatan).mark_text(
        align='center',
        baseline='bottom',
        dy=-5,
        fontSize=12,
        color='white'
    ).encode(
        x=alt.X('Instansi:N', sort='-y'),
        y='Jumlah:Q',
        text='Jumlah:Q'
    )

    # Gabungkan grafik batang dan label
    chart_kecamatan = (bars_kecamatan + text_kecamatan).properties(
        width=700,
        height=400
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=14
    ).configure_title(
        fontSize=18,
        anchor='start',
        color='gray'
    ).configure_axisX(
        labelLimit=0,
        labelAngle=45 
    )
    st.altair_chart(chart_kecamatan, use_container_width=True)

st.title("AUTO-RKM")
st.markdown("""
**Pastikan file memenuhi kriteria berikut:**
- 📄 Format file: `.xlsx`
- 📊 Kolom wajib: `Instansi`, `Topik`, `Channel`, `Status`
""")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        data = pd.read_excel(uploaded_file)
        rkm, kategori = AUTO_RKM(data)

        # Progres bar
        with st.spinner('Sedang memproses...'):
        # Simulasi proses berat
            time.sleep(2)

        st.success('Proses selesai!')

        st.header("Hasil RKM")
        st.dataframe(rkm)

        st.header("Jumlah Keluhan berdasarkan Kategori")
        st.dataframe(kategori)

        col1, col2 = st.columns(2)
        # Tombol download untuk hasil rkm
        with col1:
            excel_rkm = to_excel(rkm)
            st.download_button(
                label="Download Hasil RKM (.xlsx)",
                data=excel_rkm,
                file_name='hasil_rkm.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

        # Tombol download untuk kategori keluhan
        with col2:
            excel_kategori = to_excel(kategori)
            st.download_button(
                label="Download Kategori Keluhan (.xlsx)",
                data=excel_kategori,
                file_name='kategori_keluhan.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        st.header("Beberapa Visualisasi Data")
        st.subheader("Jumlah Keluhan Masyarakat berdasarkan Media")
        
        # Grafik batang
        kategori = kategori.sort_values(by='Jumlah', ascending=False)
        bars = alt.Chart(kategori).mark_bar(
            cornerRadiusTopLeft=5,
            cornerRadiusTopRight=5
        ).encode(
            x=alt.X('Channel:N', sort='-y', title='Jenis Media'),
            y=alt.Y('Jumlah:Q', title='Jumlah'),
            color=alt.Color('Channel:N', legend=None, scale=alt.Scale(scheme='category10')),
            tooltip=['Channel', 'Jumlah']
        )

        # jumlah kategori
        text = alt.Chart(kategori).mark_text(
            align='center',
            baseline='bottom',
            dy=-5,
            fontSize=12,
            color='white'
        ).encode(
            x=alt.X('Channel:N', sort='-y'),
            y='Jumlah:Q',
            text='Jumlah:Q'
        )

        chart = (bars + text).properties(
            width=600,
            height=400,
        ).configure_axis(
            labelFontSize=12,
            titleFontSize=14
        ).configure_title(
            fontSize=18,
            anchor='start',
            color='gray'
        ).configure_axisX(
            labelLimit=0 
        )
        st.altair_chart(chart, use_container_width=True)

        #Visualisasi Kecamatab
        st.subheader("5 Kecamatan dengan Keluhan Masyarakat Terbanyak")
        vis_kecamatan(uploaded_file)
        

    except ValueError as ve:
        st.error(f"Error: {ve}")
    except Exception as e:
        st.error("Terjadi kesalahan saat memproses file. Pastikan file Excel valid dan formatnya benar.")
        st.error(f"Detail error: {e}")
else:
    st.info("Silakan upload file Excel terlebih dahulu.")
