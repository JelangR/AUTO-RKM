import streamlit as st
import pandas as pd
import altair as alt
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
    required_cols = {'Kategori','Status'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")
    
    data_kategori = data[list(required_cols)]
    rkm_kategori = pd.DataFrame()
    selesai_kel = data_kategori[data_kategori['Status'] == 'Selesai']

    rkm_kategori['Jumlah'] = selesai_kel['Kategori'].apply(
        lambda x: len(selesai_kel[selesai_kel['Kategori'] == x])
    )
    rkm_kategori = selesai_kel['Kategori'].value_counts().reset_index()
    rkm_kategori.columns = ['Kategori', 'Jumlah']

    rkm_kategori = rkm_kategori.sort_values(by='Jumlah', ascending=False).head(5)

    # Grafik batang
    bars_kecamatan = alt.Chart(rkm_kategori).mark_bar(
        cornerRadiusTopLeft=5,
        cornerRadiusTopRight=5
    ).encode(
        x=alt.X('Kategori:N', sort='-y', title=None),
        y=alt.Y('Jumlah:Q', title=None),
        color=alt.Color('Kategori:N', legend=None, scale=alt.Scale(scheme='category20')),
        tooltip=['Kategori', 'Jumlah']
    )

    # Label jumlah
    text_kecamatan = alt.Chart(rkm_kategori).mark_text(
        align='center',
        baseline='bottom',
        dy=-5,
        fontSize=12,
        color='white'
    ).encode(
        x=alt.X('Kategori:N', sort='-y'),
        y='Jumlah:Q',
        text='Jumlah:Q'
    )

    # Gabungkan grafik batang dan label
    chart_kecamatan = (bars_kecamatan + text_kecamatan).properties(
        width=700,
        height=400
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=18
    ).configure_title(
        fontSize=18,
        anchor='start',
        color='gray'
    ).configure_axisX(
        labelLimit=0,
        labelAngle=45 
    )
    st.altair_chart(chart_kecamatan, use_container_width=True)

def vis_kelurahan(data):
    required_cols = {'Kelurahan','Status'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")
    
    data_kelurahan = data[list(required_cols)]
    rkm_kelurahan = pd.DataFrame()
    selesai_kel = data_kelurahan[data_kelurahan['Status'] == 'Selesai']

    rkm_kelurahan['Jumlah'] = selesai_kel['Kelurahan'].apply(
        lambda x: len(selesai_kel[selesai_kel['Kelurahan'] == x])
    )
    rkm_kelurahan = selesai_kel['Kelurahan'].value_counts().reset_index()
    rkm_kelurahan.columns = ['Kelurahan', 'Jumlah']

    rkm_kelurahan = rkm_kelurahan.sort_values(by='Jumlah', ascending=False).head(5)

    # Grafik batang
    bars_kelurahan = alt.Chart(rkm_kelurahan).mark_bar(
        cornerRadiusTopLeft=5,
        cornerRadiusTopRight=5
    ).encode(
        x=alt.X('Kelurahan:N', sort='-y', title=None),
        y=alt.Y('Jumlah:Q', title=None),
        color=alt.Color('Kelurahan:N', legend=None, scale=alt.Scale(scheme='category20')),
        tooltip=['Kelurahan', 'Jumlah']
    )

    # Label jumlah
    text_kelurahan = alt.Chart(rkm_kelurahan).mark_text(
        align='center',
        baseline='bottom',
        dy=-5,
        fontSize=12,
        color='white'
    ).encode(
        x=alt.X('Kelurahan:N', sort='-y'),
        y='Jumlah:Q',
        text='Jumlah:Q'
    )

    # Gabungkan grafik batang dan label
    chart_kelurahan = (bars_kelurahan + text_kelurahan).properties(
        width=700,
        height=400
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=18
    ).configure_title(
        fontSize=18,
        anchor='start',
        color='gray'
    ).configure_axisX(
        labelLimit=0,
        labelAngle=45 
    )
    st.altair_chart(chart_kelurahan, use_container_width=True)


def persen_kategori(data):
    required_cols = {'Kategori', 'Status'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")

    selesai = data[data['Status'] == 'Selesai'].copy()
    df = selesai['Kategori'].value_counts().reset_index()
    df.columns = ['Kategori', 'Jumlah']
    df['Persentase'] = (df['Jumlah'] / df['Jumlah'].sum()) * 100
    df['PersenLabel'] = df['Persentase'].map(lambda x: f"{x:.1f}%")

    pie = alt.Chart(df).mark_arc(innerRadius=50).encode(
        theta=alt.Theta(field="Jumlah", type="quantitative"),
        color=alt.Color(field="Kategori", type="nominal", scale=alt.Scale(scheme='category20b')),
        tooltip=['Kategori', 'Jumlah', alt.Tooltip('Persentase:Q', format='.1f')]
    ).properties(
        width=300,
        height=300
    )

    table = df[['Kategori', 'Jumlah', 'PersenLabel']].rename(columns={'PersenLabel': 'Persentase'})

    # Tampilkan di Streamlit sebagai dua kolom
    col1, col2 = st.columns([1, 1])
    with col1:
        st.altair_chart(pie, use_container_width=True)
    with col2:
        st.dataframe(table, use_container_width=True)

def tren_keluhan(data):
    required_cols = {'Tanggal Keluhan','Status','Kategori'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")
    
    data_tren = data[list(required_cols)]
    rkm_tren = pd.DataFrame()
    selesai_tren = data_tren[data_tren['Status'] == 'Selesai']
    selesai_tren = data_tren[data_tren['Kategori'] == 'Keluhan']

    rkm_tren['Jumlah'] = selesai_tren['Tanggal Keluhan'].apply(
        lambda x: len(selesai_tren[selesai_tren['Tanggal Keluhan'] == x])
    )
    selesai_tren['Tanggal Keluhan'] = selesai_tren['Tanggal Keluhan'].dt.date
    rkm_tren = selesai_tren['Tanggal Keluhan'].value_counts().reset_index()
    rkm_tren.columns = ['Tanggal Keluhan', 'Jumlah']

    tren_chart = alt.Chart(rkm_tren).mark_line(point=True).encode(
        x=alt.X('Tanggal Keluhan:T', title=None),
        y=alt.Y('Jumlah:Q', title=None),
        tooltip=['Tanggal Keluhan:T', 'Jumlah']
    ).properties(
        width=700,
        height=400
    ).configure_axisX(
        labelLimit=0,
        labelAngle=90
    )
    st.altair_chart(tren_chart, use_container_width=True)

def tren_permohonan_info(data):
    required_cols = {'Tanggal Keluhan','Status','Kategori'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")
    
    data_tren_info = data[list(required_cols)]
    rkm_tren_info = pd.DataFrame()
    selesai_tren_info = data_tren_info[data_tren_info['Status'] == 'Selesai']
    selesai_tren_info = data_tren_info[data_tren_info['Kategori'] == 'Permohonan Informasi']

    rkm_tren_info['Jumlah'] = selesai_tren_info['Tanggal Keluhan'].apply(
        lambda x: len(selesai_tren_info[selesai_tren_info['Tanggal Keluhan'] == x])
    )
    selesai_tren_info['Tanggal Keluhan'] = selesai_tren_info['Tanggal Keluhan'].dt.date
    rkm_tren_info = selesai_tren_info['Tanggal Keluhan'].value_counts().reset_index()
    rkm_tren_info.columns = ['Tanggal Keluhan', 'Jumlah']

    trenInfo_chart = alt.Chart(rkm_tren_info).mark_line(point=True).encode(
        x=alt.X('Tanggal Keluhan:T', title=None),
        y=alt.Y('Jumlah:Q', title=None),
        tooltip=['Tanggal Keluhan:T', 'Jumlah']
    ).properties(
        width=700,
        height=400
    ).configure_axisX(
        labelLimit=0,
        labelAngle=90
    )

    st.altair_chart(trenInfo_chart, use_container_width=True)
    
def opd_vis(data):
    required_cols = {'Kategori','Status','Instansi'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")
    
    opd_keluhan = data[list(required_cols)]
    selesai_opd = opd_keluhan[
        (opd_keluhan['Status'] == 'Selesai') &
        (opd_keluhan['Kategori'] == 'Keluhan') &
        (opd_keluhan['Instansi'].str.startswith('Dinas'))
    ]
    top5 = selesai_opd['Instansi'].value_counts().reset_index().head(5)
    top5.columns = ['Instansi', 'Jumlah']

    bars = alt.Chart(top5).mark_bar(
        cornerRadiusBottomRight=4,
        cornerRadiusTopRight=4
    ).encode(
        x=alt.X('Jumlah:Q',title=None),
        y=alt.Y('Instansi:N',title=None, sort='-x'),
        color=alt.Color('Instansi:N', legend=None, scale=alt.Scale(scheme='category10')),
        tooltip=['Instansi', 'Jumlah']
    )

    # Tambahkan label jumlah
    text = alt.Chart(top5).mark_text(
        align='left',
        baseline='middle',
        dx=3,
        fontSize=11,
        color='white'
    ).encode(
        x='Jumlah:Q',
        y=alt.Y('Instansi:N', sort='-x'),
        text='Jumlah:Q'
    )

    # Gabungkan dan tampilkan
    chart = (bars + text).properties(
        width=700,
        height=400,
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=14
    ).configure_title(
        fontSize=16,
        anchor='start',
        color='gray'
    ).configure_axisY(
        labelLimit=0
    )

    st.altair_chart(chart, use_container_width=True)

def opdInfo_vis(data):
    required_cols = {'Kategori','Status','Instansi'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")
    
    opd_info = data[list(required_cols)]
    selesai_opdInfo = opd_info[
        (opd_info['Status'] == 'Selesai') &
        (opd_info['Kategori'] == 'Permohonan Informasi') &
        (opd_info['Instansi'].str.startswith('Dinas'))
    ]
    top5 = selesai_opdInfo['Instansi'].value_counts().reset_index().head(5)
    top5.columns = ['Instansi', 'Jumlah']

    bars = alt.Chart(top5).mark_bar(
        cornerRadiusBottomRight=4,
        cornerRadiusTopRight=4
    ).encode(
        x=alt.X('Jumlah:Q',title=None),
        y=alt.Y('Instansi:N',title=None,sort='-x'),
        color=alt.Color('Instansi:N', legend=None, scale=alt.Scale(scheme='category10')),
        tooltip=['Instansi', 'Jumlah']
    )

    # Tambahkan label jumlah
    text = alt.Chart(top5).mark_text(
        align='left',
        baseline='middle',
        dx=3,
        fontSize=11,
        color='white'
    ).encode(
        x='Jumlah:Q',
        y=alt.Y('Instansi:N', sort='-x'),
        text='Jumlah:Q'
    )

    # Gabungkan dan tampilkan
    chart = (bars + text).properties(
        width=700,
        height=400,
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=14
    ).configure_title(
        fontSize=16,
        anchor='start',
        color='gray'
    ).configure_axisY(
        labelLimit=0
    )

    st.altair_chart(chart, use_container_width=True)
    
def top5Opd_keluhan_vis(data):
    required_cols = {'Kategori', 'Status', 'Instansi', 'Topik'}
    if not required_cols.issubset(data.columns):
        raise ValueError(f"Tidak Dapat Melakukan Visualisasi Karena Tidak Terdapat Kolom: {required_cols}")

    filtered = data[
        (data['Status'] == 'Selesai') &
        (data['Kategori'] == 'Keluhan') &
        (data['Instansi'].str.startswith('Dinas'))
    ]

    top_5_instansi = filtered['Instansi'].value_counts().reset_index().head(5)
    top_5_instansi.columns = ['Instansi', 'Jumlah'] 

    if top_5_instansi.empty:
        st.warning("Tidak ada data yang memenuhi syarat untuk divisualisasikan.")
        return

    
    instansi_terpilih = st.selectbox("Pilih Instansi:", top_5_instansi['Instansi'])

    data_instansi = filtered[filtered['Instansi'] == instansi_terpilih]
    topik_count = data_instansi['Topik'].value_counts().reset_index()
    topik_count.columns = ['Topik', 'Jumlah']

    bar = alt.Chart(topik_count).mark_bar().encode(
        x=alt.X('Jumlah:Q', title=None),
        y=alt.Y('Topik:N', sort='-x', title=None),
        color=alt.Color('Jumlah:Q', scale=alt.Scale(scheme='viridis'), legend=None)
    ).properties(
        width=600,
        height=300,
        title=f'Topik Keluhan - {instansi_terpilih}'
    )
    text = bar.mark_text(
        align='left',
        baseline='middle',
        dx=5
    ).encode(
        text='Jumlah:Q',
        color='white'
    )
    chart= (bar+text).configure_axisY(
        labelLimit=200
    )
    st.altair_chart(chart, use_container_width=True)

#--------APP
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
            x=alt.X('Channel:N', sort='-y', title=None),
            y=alt.Y('Jumlah:Q', title=None),
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
            titleFontSize=18
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
        vis_kecamatan(data)

        #Visualisasi Kelurahan
        st.subheader("5 Kelurahan dengan Keluhan Masyarakat Terbanyak")
        vis_kelurahan(data)

        #Visualisasi Kategori
        st.subheader("Persentase Jenis Kategori")
        persen_kategori(data)
        
        #Visualisasi tren harian keluhan
        st.subheader("Jumlah Keluhan per Hari")
        tren_keluhan(data)

        # Visualisasi OPD Keluhan terbanyak
        st.subheader('5 OPD Teratas yang Mendapatkan Keluhan')
        opd_vis(data)

        # Visualisasi tiap opd yang mendapatkan keluhan terbanyak
        st.subheader('Keluhan terhadap OPD')
        top5Opd_keluhan_vis(data)

        #visualisasi tren permohonan informasi
        st.subheader('Jumlah Permohonan Informasi per Hari')
        tren_permohonan_info(data)

        # Visualisasi OPD Permohonan Info terbanyak
        st.subheader('5 OPD Teratas yang Mendapatkan Permohonan Informasi')
        opdInfo_vis(data)
        
    except ValueError as ve:
        st.error(f"Error: {ve}")
    except Exception as e:
        st.error("Terjadi kesalahan saat memproses file. Pastikan file Excel valid dan formatnya benar.")
        st.error(f"Detail error: {e}")
else:
    st.info("Silakan upload file Excel terlebih dahulu.")
