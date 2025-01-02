#Bismillah : 28 Dec 2024 @home saturday in the old date
#Create APPs for scraping data in RECAPITULATION.xlsm
#Dedicated to stamping dept. after Excel Version earlier in 2020

import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import streamlit as st
import base64
import plotly.express as px

st.set_page_config(page_title="Quality Stamping Report", page_icon=":bar_chart:",layout="wide")

# Fungsi untuk mengubah gambar menjadi base64
def get_image_as_base64(image_path):
	with open(image_path, "rb") as img_file:
		return base64.b64encode(img_file.read()).decode()
		
# heading
kolkir,kolnan=st.columns((2,1))	#artinya kolom sebelahkiri lebih lebar 2x dari kanan

with kolkir:
	st.markdown("""<h2 style="color:green;margin-top:-10px;margin-bottom:0px;"> 📊 SUMMARY RECAPITULATION </h2>""", unsafe_allow_html=True)
	st.write("Stamping Part Summary Report")
	st.markdown("""<p style="margin-top:-10px;margin-bottom:0px;font-size:14px">Dedicated to STAMPING PART ©️2024 e-WeYe - All Rights Reserved</p>""", unsafe_allow_html=True)

	
with kolnan:
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
            <p style="margin-top:-10px;margin-bottom:0px;font-size:14px">Sebelum memulai:
                <br>         
                1. Siapkan file RECAPITULATION dari folder PAMOR<br>
                2. Pastikan file extensi excelnya adalah .xlsm<br>
                3. Beri identitas pada setiap nama filenya, misal "1JanRecap.xlsm"
        </div>
        """,
        unsafe_allow_html=True
	)
        
	# st.markdown("---")


#START SEDOT
# Fungsi untuk membuka dialog multi-file selection
st.markdown("---")
    
# def select_files(): ------> diganti dengan streamlit uploader, krn tkinter tdk bisa dideploy di cloud streamlit - 28Dec2024 11.46WIB 
#     root = Tk()
#     root.withdraw()  # Menyembunyikan jendela utama Tkinter
#     root.attributes('-topmost', True)  # Membawa dialog ke depan
#     file_paths = filedialog.askopenfilenames(
#         title="Pilih file Excel",
#         filetypes=[("Excel files", "*.xlsm")]
#     )
#     root.destroy()
#     return file_paths

# Function to select files using Streamlit 28Dec2024
def select_files():
    # st.title("Upload Excel Files")
    uploaded_files = st.file_uploader(
        "Pilih file excel berekstensi .xlsm:",
        type=["xlsm"],
        accept_multiple_files=True
    )
    return uploaded_files
    
if __name__ == "__main__":
    uploaded_files = select_files()
    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.write(f"Uploaded file: {uploaded_file.name}")

if uploaded_files:  # Jika user telah memilih file
    # Alamat sel yang akan diambil
    cols_mor = ['E7', 'E8', 'E9', 'E11', 'E12', 'E13', 'E14', 'E15', 'E18', 'E19', 'E20']
    cols_ng = ['P7', 'P8', 'P9', 'P11', 'P12', 'P13', 'P14', 'P15', 'P18', 'P19', 'P20']

    header_names = [
        'GR#01', 'GR#02', 'GR#04', 'GR#03', 'GR#09',
        'PW#5', 'RING#7', 'PW#10', 'CR#12', 'CR#13', 'CR#14'
    ]

    # Untuk menyimpan data
    data_mor = []
    data_ng = []

    for uploaded_file in uploaded_files:
        # file_name = uploaded_file.name  # Hanya mengambil nama file
        # Ambil nama file tanpa ekstensi
        file_name = os.path.splitext(uploaded_file.name)[0]  # Mengambil nama file tanpa ekstensi

        workbook_data_mor = {'Nama File': file_name}
        workbook_data_ng = {'Nama File': file_name}

        try:
            # Membuka workbook menggunakan openpyxl
            wb = load_workbook(uploaded_file, data_only=True)
            sheet = wb['REKAP']  # Pastikan nama sheet sesuai

            # Mengambil data dari sel yang sesuai untuk MOR dan NG
            for i, (mor_cell, ng_cell) in enumerate(zip(cols_mor, cols_ng)):
                workbook_data_mor[header_names[i]] = sheet[mor_cell].value
                workbook_data_ng[header_names[i]] = sheet[ng_cell].value

            data_mor.append(workbook_data_mor)
            data_ng.append(workbook_data_ng)

            # Pengecekan jika cols_mor atau cols_ng kosong
            # if cols_mor:
            #     for col in cols_mor:
            #         cell_value = sheet[col].value
            #         workbook_data_mor[col] = str(cell_value) if cell_value is not None else ''
            #     data_mor.append(workbook_data_mor)

            # if cols_ng:
            #     for col in cols_ng:
            #         cell_value = sheet[col].value
            #         workbook_data_ng[col] = str(cell_value) if cell_value is not None else ''
            #     data_ng.append(workbook_data_ng)

        except Exception as e:
            st.error(f"Error reading {file_name}: {e}")

    # Membuat DataFrame dari data
    mor_table = pd.DataFrame(data_mor)
    ng_table = pd.DataFrame(data_ng)

    # # Pastikan kolom header_names ada di DataFrame
    # mor_columns = [col for col in header_names if col in mor_table.columns]
    # ng_columns = [col for col in header_names if col in ng_table.columns]

    # # Mengubah nilai menjadi numerik, nilai yang tidak dapat dikonversi akan menjadi NaN
    # mor_table[header_names] = mor_table[header_names].apply(pd.to_numeric, errors='coerce')
    # ng_table[header_names] = ng_table[header_names].apply(pd.to_numeric, errors='coerce')

    # # Menghitung rata-rata, mengabaikan NaN
    # mor_table.loc['Average'] = mor_table.mean(numeric_only=True)
    # ng_table.loc['Average'] = ng_table.mean(numeric_only=True)

    # mor_table.loc['Average', 'Nama File'] = 'Average'
    # ng_table.loc['Average', 'Nama File'] = 'Average'

    # Menambahkan rata-rata baris ('Avg.')
    mor_table['Avg.'] = mor_table.iloc[:, 1:].mean(axis=1)
    ng_table['Avg.'] = ng_table.iloc[:, 1:].mean(axis=1)

    # Menambahkan rata-rata kolom
    mor_table.loc['Average'] = mor_table.mean(numeric_only=True)
    mor_table.loc['Average', 'Nama File'] = 'Average'

    ng_table.loc['Average'] = ng_table.mean(numeric_only=True)
    ng_table.loc['Average', 'Nama File'] = 'Average'

    st.markdown("---")
    st.subheader("SUMMARY REPORT")
    # Menampilkan tabel di Streamlit
    st.write("Recapitulation MOR (%)")
    st.dataframe(mor_table)

    st.write("Recapitulation NG (%)")
    st.dataframe(ng_table)

    st.markdown("---")
     # Membuat grafik garis interaktif MOR
    mor_melted = mor_table.melt(
        id_vars=['Nama File'], 
        value_vars=header_names,
        var_name='MC', 
        value_name='MOR (%)'
    )

    st.subheader("Grafik Tren MOR by Machine & Month")
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

    st.markdown("---")
    # Membuat grafik garis interaktif untuk NG
    ng_melted = ng_table.melt(
        id_vars=['Nama File'], 
        value_vars=header_names,
        var_name='MC', 
        value_name='NG (%)'
    )
    st.subheader("Grafik Tren NG  by Machine & Month")
    fig_ng = px.line(
        ng_melted, 
        x='Nama File', 
        y='NG (%)', 
        color='MC',
        title="Trendline NG  by Machine & Month",
        markers=True
    )
    fig_ng.update_layout(
        xaxis_title="Nama File",
        yaxis_title="ng (%)",
        legend_title="MC",
        template="plotly_white"
    )
    st.plotly_chart(fig_ng, use_container_width=True)

    # Akhir Membuat grafik garis interaktif

        #Footer
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
        st.image(e_WeYe,"Web Developer - eWeYe ©️2024",use_column_width="always")

    with kaki_kanan2:
        st.write("")

    with kaki_kanan:
        st.write("")
else:
    st.info("Klik tombol 'Pilih file Excel berekstensi .xlsm' untuk memulai...")






# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
