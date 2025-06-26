#Bismillah : 28 Dec 2024 @home saturday in the old date
#Create APPs for scraping data in RECAPITULATION.xlsm
#Dedicated to stamping dept. PT. KPD  (after Excel Version earlier in 2020)

import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import streamlit as st
import base64
import plotly.express as px

st.set_page_config(page_title="Quality Stamping Report", page_icon=":bar_chart:",layout="wide")

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

# Fungsi untuk mengubah gambar menjadi base64
def get_image_as_base64(image_path):
	with open(image_path, "rb") as img_file:
		return base64.b64encode(img_file.read()).decode()

# Login Page added 12May2025 20.08 WIb @home

# Inisialisasi session state login
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False


# def login_page():
if not st.session_state["logged_in"]:	

    kol1 ,kol3,kol5 = st.columns((1,1,1))

    with kol1:
        # Adjust the file path based on the current directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        logo_KPD = os.path.join(current_dir, 'logoKPD.png')
        # Memuat gambar dan mengubahnya menjadi base64
        # logo_KPD ='logoKPD.png'
        image_base64 = get_image_as_base64(logo_KPD)
        
        # Menampilkan gambar dan teks di kolom kanan dengan posisi berdampingan
        st.markdown(
            f"""
            <style>
            .container {{
                display: flex;
                align-items:center;
                justify-content: flex-end;
                flex-wrap: wrap;
            }}
            .container img {{
                width: 50px;
                margin-top: -20px;
            }}
            .container h2 {{
                color: grey;
                font-size: 20px;
                margin-top: -20px;
                margin-right: 10px;
                margin-bottom: 0px;
            }}
            @media (min-width: 600px) {{
                .container {{
                    justify-content: center;
                }}
                .container img {{
                    margin-top: 0;
                }}
                .container h2 {{
                    margin-top: 0;
                    text-align: center;
                }}
            }}
            </style>
            <div class="container">
                <img src='data:image/png;base64,{image_base64}'/>
                <h2 style="color:blue;">PT. KARYAPRATAMA DUNIA</h2>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown("<div style='text-align: center; font-weight: bold;color:orange;'>QUALITY DEPARTMENT</div>", unsafe_allow_html=True)
    with kol3:
        # Form login
        st.info("Please log in to access the application.")
        st.markdown('---')
        st.markdown('<div class="login-container"><div class="login-form">', unsafe_allow_html=True)
        username = st.text_input("Username", key="username")
        password = st.text_input("Password", type="password", key="password")
        if st.button("Login"):
            if username == "kpd" and password == "kpd080808":
                st.session_state["logged_in"] = True
                # Reload halaman dengan mengatur ulang parameter URL
                st.query_params.clear()
            else:
                st.error("Invalid username or password!")

        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown('---')
    with kol5:
        st.markdown("""<h2 style="color:green;margin-top:-10px;margin-bottom:0px;"> 📊 SUMMARY REPORT </h2>""", unsafe_allow_html=True)
        
        st.markdown("<div style='text-align: center; font-weight: bold;color:blue;'>Production Performance Stamping Part</div>", unsafe_allow_html=True)
    st.stop()  # Hentikan eksekusi jika belum login



		
# heading
kolkir,kolnan=st.columns((2,1))	#artinya kolom sebelahkiri lebih lebar 2x dari kanan

with kolkir:#Judul Summary Report
	st.markdown("""<h2 style="color:green;margin-top:-10px;margin-bottom:0px;"> 📊 SUMMARY REPORT </h2>""", unsafe_allow_html=True)
	st.write("Production Performance Stamping Part")
	st.markdown("""<p style="margin-top:-10px;margin-bottom:0px;font-size:10px"> ©️2024 e-WeYe</p>""", unsafe_allow_html=True)

	
with kolnan:#logo KPD
	# Adjust the file path based on the current directory
	current_dir = os.path.dirname(os.path.abspath(__file__))
	logo_KPD = os.path.join(current_dir, 'logoKPD.png')
	# Memuat gambar dan mengubahnya menjadi base64
	# logo_KPD ='logoKPD.png'
	image_base64 = get_image_as_base64(logo_KPD)
	
    # Menampilkan gambar dan teks di kolom kanan dengan posisi berdampingan
	st.markdown(
        f"""
        <style>
        .container {{
            display: flex;
            align-items:center;
            justify-content: flex-end;
            flex-wrap: wrap;
        }}
        .container img {{
            width: 50px;
            margin-top: -20px;
        }}
        .container h2 {{
            color: grey;
            font-size: 20px;
            margin-top: -20px;
            margin-right: 10px;
            margin-bottom: 0px;
        }}
        @media (min-width: 600px) {{
            .container {{
                justify-content: center;
            }}
            .container img {{
                margin-top: 0;
            }}
            .container h2 {{
                margin-top: 0;
                text-align: center;
            }}
        }}
        </style>

        <div class="container">
            <h2 style="color:blue;">PT. KARYAPRATAMA DUNIA</h2>
            <img src='data:image/png;base64,{image_base64}'/>
            <br>
            <br>
            <br>
            <h5 style="color:brown;font-size:12px;">Quality Dept. - Stamping Part</h5>
        </div>
        """,
        unsafe_allow_html=True
	)
        
	# st.markdown("---")
 


#START SEDOT
# Fungsi untuk membuka dialog multi-file selection
st.markdown("---")

sisi_kiri,sisi_kanan=st.columns((1,1))    
# Menampilkan petunjuk di sisi kiri
with sisi_kiri:
        st.markdown(
        """
        <p style="margin-top:-10px;margin-bottom:0px;font-size:14px">Sebelum upload file:
                <br>         
                1. Siapkan file RECAPITULATION dari folder PAMOR<br>
                2. Pastikan file extensi excelnya adalah .xlsm<br>
                3. Copy file RECAPITULATION ke dalam folder yang sama<br>
                4. Beri identitas (rename) pada setiap nama filenya, misal "1Jan2024.xlsm </p>
                <br>
              
        """,
        unsafe_allow_html=True
    )
# Menampilkan petunjuk di sisi kanan
with sisi_kanan:
    
    st.markdown(
        """
        <p style="margin-top:-10px;margin-bottom:0px;font-size:14px">Petunjuk upload file:
                <br>         
                1. Klik tombol 'Browse Files' atau drag file ke area yang tersedia<br>
                2. Pilih file Recapitulation yang sudah diberi identitas pada 1 folder<br>
                3. Memilih banyak file (multi select) dengan cara menekan tombol SHIFT+ Mouse<br>
                   (pilih file berurutan) atau CTRL+ Mouse (pilih multi file secara tidak berurutan)<br>
                4. Klik 'Open' dan Tunggu proses upload file selesai<br>
                </p>
                
        """,
        unsafe_allow_html=True
    )

#Added 18Mar2025 to make this apps more user friendly and globally accessible
st.warning(f"Jika sumber file yang ingin dibersihkan berada di folder Google Drive, unduh/download lewat link berikut ini: [Link Folder](https://drive.google.com/drive/folders/1wdumCde2NJaeH-cgdxdukGvTKpjbS3qT?usp=sharing)  |  Download file Excel (hanya .xlsm) dari folder tersebut ke perangkat Anda, lalu unggah/upload file lewat menu Browse Files di bawah ini:")
    

# Function to select files using Streamlit 28Dec2024
def select_files():
    # st.title("Upload Excel Files")
    uploaded_files = st.file_uploader(
        "Pilih file excel berekstensi .xlsm:",
        type=["xlsm"],
        accept_multiple_files=True
    )
    return uploaded_files

uploaded_files = select_files()
    
# Start Cleaning
#region untuk mengambil data dari file yang diupload
if uploaded_files:  # Jika user telah memilih file
    # Alamat sel yang akan diambil
    cols_mor = ['E7', 'E8', 'E9', 'E11', 'E12', 'E13', 'E14', 'E15', 'E18', 'E19', 'E20']
    cols_ng = ['P7', 'P8', 'P9', 'P11', 'P12', 'P13', 'P14', 'P15', 'P18', 'P19', 'P20']
    cols_qty = ['H7', 'H8', 'H9', 'H11', 'H12', 'H13', 'H14', 'H15','H16','H17', 'H18', 'H19', 'H20']

    header_names = [
        'GR#01', 'GR#02', 'GR#04', 'GR#03', 'GR#09',
        'PW#5', 'RING#7', 'PW#10', 'CR#12', 'CR#13', 'CR#14'
    ]

    # Untuk menyimpan data
    data_mor = []
    data_ng = []
    data_qty = []

    for uploaded_file in uploaded_files:
        # file_name = uploaded_file.name  # Hanya mengambil nama file
        # Ambil nama file tanpa ekstensi
        file_name = os.path.splitext(uploaded_file.name)[0]  # Mengambil nama file tanpa ekstensi

        workbook_data_mor = {'Nama File': file_name}
        workbook_data_ng = {'Nama File': file_name}
        workbook_data_qty = {'Nama File': file_name}

        try:
            # Membuka workbook menggunakan openpyxl
            wb = load_workbook(uploaded_file, data_only=True)
            sheet = wb['REKAP']  # Pastikan nama sheet sesuai

            # Mengambil data dari sel yang sesuai untuk MOR, NG dan QTY
            for i, (mor_cell, ng_cell, qty_cell) in enumerate(zip(cols_mor, cols_ng, cols_qty)):
                workbook_data_mor[header_names[i]] = sheet[mor_cell].value
                workbook_data_ng[header_names[i]] = sheet[ng_cell].value
                if header_names[i] == 'PW#10':
                    workbook_data_qty[header_names[i]] = (
                        sheet['H15'].value + sheet['H16'].value + sheet['H17'].value
                    )
                else:
                    workbook_data_qty[header_names[i]] = sheet[qty_cell].value

            data_mor.append(workbook_data_mor)
            data_ng.append(workbook_data_ng)
            data_qty.append(workbook_data_qty)

        except Exception as e:
            st.error(f"Error reading {file_name}: {e}")

    # Membuat DataFrame dari data
    mor_table = pd.DataFrame(data_mor)
    ng_table = pd.DataFrame(data_ng)
    qty_table = pd.DataFrame(data_qty)

    # Menambahkan rata-rata baris ('Avg.')
    mor_table['Avg.'] = mor_table.iloc[:, 1:].mean(axis=1)

    # Tambahkan baris terakhir berisi nilai teks 'MOR' untuk semua kolom, dinamakan 'Category'
    category_row = {col: 'MOR' for col in mor_table.columns}
    category_row['Nama File'] = 'Category'
    # Simpan baris Average, lalu Category, lalu gabungkan ulang
    if 'Average' in mor_table.index:
        average_row = mor_table.loc[['Average']]
        mor_table = mor_table.drop('Average')
        mor_table = pd.concat([mor_table, average_row, pd.DataFrame([category_row], index=['Category'])])
    else:
        mor_table = pd.concat([mor_table, pd.DataFrame([category_row], index=['Category'])])

    ng_table['Avg.'] = ng_table.iloc[:, 1:].mean(axis=1)
    qty_table['Total'] = qty_table.iloc[:, 1:].sum(axis=1)

    # Menambahkan rata-rata kolom
    mor_table.loc['Average'] = mor_table.mean(numeric_only=True)
    mor_table.loc['Average', 'Nama File'] = 'Average'

    ng_table.loc['Average'] = ng_table.mean(numeric_only=True)
    ng_table.loc['Average', 'Nama File'] = 'Average'

    qty_table.loc['Sum'] = qty_table.sum(numeric_only=True)
    qty_table.loc['Sum', 'Nama File'] = 'Total'

    st.markdown("---")

#end region

    #start SUMMARY REPORT
    st.subheader("SUMMARY REPORT")

    tab_mor, tab_ng, tab_qty = st.tabs(["MOR", "NG", "Qty"])

    with tab_mor:# Tab untuk MOR
        st.subheader("Recapitulation MOR (%) - Target 85%")
        # Menampilkan tabel di Streamlit
        st.write("")
        st.dataframe(mor_table, use_container_width=True,hide_index=True)

        # Grafik batang rata-rata MOR per mesin
        avg_mor_per_machine = mor_table.loc['Average', header_names]
        fig_avg_mor = px.bar(
            x=header_names,
            y=avg_mor_per_machine.values,
            labels={'x': 'Nomor Mesin', 'y': 'Average MOR (%)'},
            title='Average MOR per Mesin',
            text_auto=True,
            color_discrete_sequence=['#A0C878'] * len(header_names)
        )
        fig_avg_mor.add_shape(
            type='line',
            x0=-0.5,
            x1=len(header_names)-0.5,
            y0=85,
            y1=85,
            line=dict(color='red', width=2, dash='dash'),
            xref='x',
            yref='y'
        )
        st.plotly_chart(fig_avg_mor, use_container_width=True)

        st.markdown("---")

        # Membuat grafik garis interaktif MOR
        mor_melted = mor_table.drop(index='Average').melt(
            id_vars=['Nama File'], 
            value_vars=header_names,
            var_name='MC', 
            value_name='MOR (%)'
        )
        # st.subheader("Grafik Tren MOR by Machine & Month")
        fig = px.line(
            mor_melted, 
            x='Nama File', 
            y='MOR (%)', 
            color='MC',
            title="Trendline MOR by Machine & Month",
            markers=True
        )
        fig.update_layout(
            xaxis_title="Nama File",
            yaxis_title="MOR (%)",
            legend_title="MC",
            template="plotly_white"
        )
        st.plotly_chart(fig, use_container_width=True)

        # Membuat grafik batang interaktif untuk MOR
        st.subheader("GRAFIK MOR")

        # Membuat grafik batang interaktif untuk MOR GR#01
        with st.expander("Grafik MOR GR#01"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='GR#01',
                color='Nama File',
                title='Grafik MOR GR#01',
                labels={'GR#01': 'GR#01 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR GR#02
        with st.expander("Grafik MOR GR#02"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='GR#02',
                color='Nama File',
                title='Grafik MOR GR#02',
                labels={'GR#02': 'GR#02 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR GR#03
        with st.expander("Grafik MOR GR#03"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='GR#03',
                color='Nama File',
                title='Grafik MOR GR#03',
                labels={'GR#03': 'GR#03 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR GR#04
        with st.expander("Grafik MOR GR#04"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='GR#04',
                color='Nama File',
                title='Grafik MOR GR#04',
                labels={'GR#04': 'GR#04 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR GR#09
        with st.expander("Grafik MOR GR#09"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='GR#09',
                color='Nama File',
                title='Grafik MOR GR#09',
                labels={'GR#09': 'GR#09 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR PW#5
        with st.expander("Grafik MOR PW#5"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='PW#5',
                color='Nama File',
                title='Grafik MOR PW#5',
                labels={'PW#5': 'PW#5 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR RING#7
        with st.expander("Grafik MOR RING#7"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='RING#7',
                color='Nama File',
                title='Grafik MOR RING#7',
                labels={'RING#7': 'RING#7 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR PW#10
        with st.expander("Grafik MOR PW#10"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='PW#10',
                color='Nama File',
                title='Grafik MOR PW#10',
                labels={'PW#10': 'PW#10 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR CR#12
        with st.expander("Grafik MOR CR#12"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='CR#12',
                color='Nama File',
                title='Grafik MOR CR#12',
                labels={'CR#12': 'CR#12 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # st.markdown("---")

        # Membuat grafik batang interaktif untuk MOR CR#13
        with st.expander("Grafik MOR CR#13"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='CR#13',
                color='Nama File',
                title='Grafik MOR CR#13',
                labels={'CR#13': 'CR#13 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)



        # Membuat grafik batang interaktif untuk MOR CR#14

        with st.expander("Grafik MOR CR#14"):
            fig = px.bar(
                mor_table,
                x='Nama File',
                y='CR#14',
                color='Nama File',
                title='Grafik MOR CR#14',
                labels={'CR#14': 'CR#14 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menambahkan garis horizontal pada nilai 85%
            fig.add_shape(
                type='line',
                x0=0,
                x1=1,
                y0=85,
                y1=85,
                line=dict(color='red', width=2),
                xref='paper',
                yref='y'
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

    with tab_ng: #Tab untuk NG
        st.subheader("Recapitulation NG (%) - Target 0.5%")
        st.write("")
        st.dataframe(ng_table, use_container_width=True,hide_index=True)
        # Grafik batang rata-rata NG per mesin (warna grey)
        avg_ng_per_machine = ng_table.loc['Average', header_names]
        fig_avg_ng = px.bar(
            x=header_names,
            y=avg_ng_per_machine.values,
            labels={'x': 'Nomor Mesin', 'y': 'Average NG (%)'},
            title='Average NG per Mesin',
            text_auto=True,
            color_discrete_sequence=['#7F8CAA'] * len(header_names)
        )
        fig_avg_ng.add_shape(
            type='line',
            x0=-0.5,
            x1=len(header_names)-0.5,
            y0=0.5,
            y1=0.5,
            line=dict(color='red', width=2, dash='dash'),
            xref='x',
            yref='y'
        )
        st.plotly_chart(fig_avg_ng, use_container_width=True)

        st.markdown("---")

        # Membuat grafik garis interaktif untuk NG
        # with st.expander("Grafik Tren NG by Machine & Month"):
        ng_melted = ng_table.drop(index='Average').melt(
            id_vars=['Nama File'], 
            value_vars=header_names,
            var_name='MC', 
            value_name='NG (%)')
        # st.subheader("Grafik Tren NG by Machine & Month")
        fig_ng = px.line(
            ng_melted, 
            x='Nama File', 
            y='NG (%)', 
            color='MC',
            title="Trendline NG by Machine & Month",
            markers=True)
        fig_ng.update_layout(
            xaxis_title="Nama File",
            yaxis_title="NG (%)",
            legend_title="MC",
            template="plotly_white")
        st.plotly_chart(fig_ng, use_container_width=True)
            # Membuat grafik batang interaktif untuk NG
        st.subheader("GRAFIK NG")
    
        # Membuat grafik batang interaktif untuk NG GR#01
        with st.expander("Grafik NG GR#01"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='GR#01',
                color='Nama File',
                title='Grafik NG GR#01',
                labels={'GR#01': 'GR#01 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG GR#02
        with st.expander("Grafik NG GR#02"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='GR#02',
                color='Nama File',
                title='Grafik NG GR#02',
                labels={'GR#02': 'GR#02 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG GR#03
        with st.expander("Grafik NG GR#03"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='GR#03',
                color='Nama File',
                title='Grafik NG GR#03',
                labels={'GR#03': 'GR#03 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG GR#04
        with st.expander("Grafik NG GR#04"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='GR#04',
                color='Nama File',
                title='Grafik NG GR#04',
                labels={'GR#04': 'GR#04 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG GR#09
        with st.expander("Grafik NG GR#09"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='GR#09',
                color='Nama File',
                title='Grafik NG GR#09',
                labels={'GR#09': 'GR#09 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG PW#5
        with st.expander("Grafik NG PW#5"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='PW#5',
                color='Nama File',
                title='Grafik NG PW#5',
                labels={'PW#5': 'PW#5 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG RING#7
        with st.expander("Grafik NG RING#7"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='RING#7',
                color='Nama File',
                title='Grafik NG RING#7',
                labels={'RING#7': 'RING#7 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG PW#10
        with st.expander("Grafik NG PW#10"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='PW#10',
                color='Nama File',
                title='Grafik NG PW#10',
                labels={'PW#10': 'PW#10 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG CR#12
        with st.expander("Grafik NG CR#12"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='CR#12',
                color='Nama File',
                title='Grafik NG CR#12',
                labels={'CR#12': 'CR#12 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG CR#13
        with st.expander("Grafik NG CR#13"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='CR#13',
                color='Nama File',
                title='Grafik NG CR#13',
                labels={'CR#13': 'CR#13 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk NG CR#14
        with st.expander("Grafik NG CR#14"):
            fig = px.bar(
                ng_table,
                x='Nama File',
                y='CR#14',
                color='Nama File',
                title='Grafik NG CR#14',
                labels={'CR#14': 'CR#14 (%)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        #end NG

    with tab_qty:#Tab untuk Qty
        st.subheader("Recapitulation Qty (pcs)")
        
        st.dataframe(qty_table, use_container_width=True,hide_index=True)

        # Grafik batang total Qty per mesin (warna light brown)
        sum_qty_per_machine = qty_table.loc['Sum', header_names]
        fig_sum_qty = px.bar(
            x=header_names,
            y=sum_qty_per_machine.values,
            labels={'x': 'Nomor Mesin', 'y': 'Total Qty (pcs)'},
            title='Total Qty per Mesin',
            text=sum_qty_per_machine.apply(lambda x: f"{int(x):,}".replace(",", ".")),  # format with dot as thousand separator
            text_auto=False,
            color_discrete_sequence=['#CAE8BD'] * len(header_names)  # light brown
        )
        fig_sum_qty.update_traces(textposition='outside')
        fig_sum_qty.update_layout(
            yaxis_tickformat=',d',  # show full integer with thousand separator
            uniformtext_minsize=8,
            uniformtext_mode='show'
        )
        st.plotly_chart(fig_sum_qty, use_container_width=True)
    
        # Grafik garis interaktif untuk Qty
        st.markdown("---")
        qty_melted = qty_table.drop(index='Sum').melt(
            id_vars=['Nama File'], 
            value_vars=header_names,
            var_name='MC', 
            value_name='Qty (pcs)'
        )
        # st.subheader("Grafik Tren Qty by Machine & Month")
        fig_qty = px.line(
            qty_melted, 
            x='Nama File', 
            y='Qty (pcs)', 
            color='MC',
            title="Trendline Qty by Machine & Month",
            markers=True
        )
        fig_qty.update_layout(
            xaxis_title="Nama File",
            yaxis_title="Qty (pcs)",
            legend_title="MC",
            template="plotly_white"
        )
        st.plotly_chart(fig_qty, use_container_width=True)

        st.markdown("---")

        # Membuat grafik batang interaktif untuk Qty 200125
        st.subheader("GRAFIK QTY")
        # Membuat grafik batang interaktif untuk Qty GR#01
        with st.expander("Grafik Qty GR#01"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='GR#01',
                color='Nama File',
                title='Grafik Qty GR#01',
                labels={'GR#01': 'GR#01 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty GR#02
        with st.expander("Grafik Qty GR#02"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='GR#02',
                color='Nama File',
                title='Grafik Qty GR#02',
                labels={'GR#02': 'GR#02 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty GR#03
        with st.expander("Grafik Qty GR#03"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='GR#03',
                color='Nama File',
                title='Grafik Qty GR#03',
                labels={'GR#03': 'GR#03 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty GR#04
        with st.expander("Grafik Qty GR#04"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='GR#04',
                color='Nama File',
                title='Grafik Qty GR#04',
                labels={'GR#04': 'GR#04 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty GR#09
        with st.expander("Grafik Qty GR#09"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='GR#09',
                color='Nama File',
                title='Grafik Qty GR#09',
                labels={'GR#09': 'GR#09 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty PW#5
        with st.expander("Grafik Qty PW#5"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='PW#5',
                color='Nama File',
                title='Grafik Qty PW#5',
                labels={'PW#5': 'PW#5 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty RING#7
        with st.expander("Grafik Qty RING#7"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='RING#7',
                color='Nama File',
                title='Grafik Qty RING#7',
                labels={'RING#7': 'RING#7 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty PW#10
        with st.expander("Grafik Qty PW#10"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='PW#10',
                color='Nama File',
                title='Grafik Qty PW#10',
                labels={'PW#10': 'PW#10 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty CR#12
        with st.expander("Grafik Qty CR#12"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='CR#12',
                color='Nama File',
                title='Grafik Qty CR#12',
                labels={'CR#12': 'CR#12 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty CR#13
        with st.expander("Grafik Qty CR#13"):
            fig = px.bar(
                qty_table.drop(index='Sum'),
                x='Nama File',
                y='CR#13',
                color='Nama File',
                title='Grafik Qty CR#13',
                labels={'CR#13': 'CR#13 (pcs)', 'Nama File': 'Bulan-Tahun'},
                text_auto=True,
            )

            # Menghilangkan legend
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig)

        # Membuat grafik batang interaktif untuk Qty CR#14
        with st.expander("Grafik Qty CR#14"):
                fig = px.bar(
                    qty_table.drop(index='Sum'),
                    x='Nama File',
                    y='CR#14',
                    color='Nama File',
                    title='Grafik Qty CR#14',
                    labels={'CR#14': 'CR#14 (pcs)', 'Nama File': 'Bulan-Tahun'},
                    text_auto=True,
                )

                # Menghilangkan legend
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig)

        
        #end Qty

    #end SUMMARY REPORT
    # ---- FOOTER ----
    
    #Footer diisi foto ditaruh ditengah
    st.markdown("---")
    kaki_kiri,kaki_kiri2, kaki_tengah,kaki_kanan2, kaki_kanan=st.columns((2,2,1,2,2))

    with kaki_kiri:
        st.write("")

    with kaki_kiri2:
        st.write("")

    with kaki_tengah:
        # kontener_photo=st.container(border=True)
        # Adjust the file path based on the current directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        e_WeYe = os.path.join(current_dir, 'eweye.png')
        # Memuat gambar dan mengubahnya menjadi base64
        # logo_KPD ='logoKPD.png'
        image_base64 = get_image_as_base64(e_WeYe)
        st.image(e_WeYe,"©️2024 | e-WeYe",use_column_width="always")

    with kaki_kanan2:
        st.write("")

    with kaki_kanan:
        st.write("")
else:
        # Jika user belum memilih file, tampilkan pesan info
        st.info("Menunggu file di-upload...")






# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
