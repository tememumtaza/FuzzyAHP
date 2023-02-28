#Import library yang dibutuhkan
import pandas as pd 
import numpy as np
import xlsxwriter
from io import BytesIO
import plotly.graph_objects as go
import streamlit as st
from fahp.py import read_excel_file, pairwiseComp, isConsistent, FAHP

st.set_page_config(page_title="FAHP : Seleksi Keringanan UKT", layout="wide",menu_items=None)


def filedownload(df, filename='output.xlsx'):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter') # ubah engine ke xlsxwriter
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}">Download file</a>'


# Membuat fungsi untuk pengelompokan berdasarkan score dan batas skor yang ditentukan oleh pengguna
def kelompokkan_score(Score):
    if Score >= batas_keringanan_50:
        return 'Keringanan 50%'
    elif Score >= batas_keringanan_30:
        return 'Keringanan 30%'
    elif Score >= batas_keringanan_20:
        return 'Keringanan 20%'
    else:
        return 'Tanpa Keringanan'

st.title("Fuzzy AHP untuk Seleksi Keringanan UKT")

with st.sidebar:
    st.write("## Upload File \n")
    st.write('Sampel file dapat diakses [disini!](https://github.com/tememumtaza/FuzzyAHP/tree/main/Data%20Sample)\n')
    file_criteria = st.file_uploader("Upload File Nilai Kriteria", type=['xlsx'], key="criteria")
    file_alternatives = st.file_uploader("Upload File Nilai Alternatif", type=['xlsx'], key="alternatives")

st.sidebar.markdown(" © 2023 Github [@temamumtaza](https://github.com/temamumtaza)")

if file_criteria is not None and file_alternatives is not None:
    criteriaDict = read_excel_file(file_criteria, 0)
    alternativesName = read_excel_file(file_alternatives, 0)

    criteria = read_excel_file(file_criteria, 1)
    for i in range(1, len(criteriaDict)+1):
        exec(f"altc{i} = read_excel_file(file_alternatives, {i})")

    def compare(*items):
        n = len(items)
        matrix = np.zeros((n, n, 3))
        for i, (c_i, v_i) in enumerate(items):
            for j, (c_j, v_j) in enumerate(items):
                if c_i == c_j or v_i == v_j:
                    matrix[i][j] = [1, 1, 3]
                else:
                    diff = abs(v_i - v_j)
                    if diff == 1:
                        matrix[i][j] = [1, 3, 5]
                    elif diff == 2:
                        matrix[i][j] = [3, 5, 7]
                    elif diff == 3:
                        matrix[i][j] = [5, 7, 9]
                    elif diff >= 4:
                        matrix[i][j] = [7, 9, 9]
                    if v_i < v_j:
                        matrix[i][j] = 1 / matrix[i][j][::-1]
        return matrix
    
    crxcr = np.array(compare(*criteria))
    for i in range(1, len(criteriaDict)+1):
        alt = eval(f"altc{i}")
        cr = compare(*alt)
        exec(f"altxalt_cr{i} = np.array(cr)")
    
    #Membuat array numpy untuk altxalt dengan mengambil nilai dari variabel global
    altxalt = np.stack([globals()[f"altxalt_cr{i+1}"] for i in range(len(criteriaDict))])

    # Membuat checkbox untuk menampilkan komputasi lengkap (konsistensi matrix)
    show_comp = st.checkbox("Tampilkan komputasi lengkap")

    #Memanggil fungsi FAHP dengan parameter yang telah didefinisikan sebelumnya
    #printComp di-set False agar tidak menampilkan komputasi konsistensi matrix
    output = FAHP(crxcr, altxalt, alternativesName, show_comp)

    #Menampilkan rangking alternatif dengan output dari fungsi FAHP
    st.write("\n RANGKING ALTERNATIF:\n", output)

    # tampilkan widget
    st.header("Pengelompokan Data Berdasarkan Skor Tertinggi")

    # Menambahkan widget untuk memungkinkan pengguna menentukan batas skor untuk masing-masing kelompok
    batas_keringanan_50 = st.slider('Batas skor Keringanan 50%:', min_value=0.00, max_value=0.01, value=0.0056,step=0.0001, format="%.4f")
    batas_keringanan_30 = st.slider('Batas skor Keringanan 30%:', min_value=0.00, max_value=0.01, value=0.0048,step=0.0001, format="%.4f")
    batas_keringanan_20 = st.slider('Batas skor Keringanan 20%:', min_value=0.00, max_value=0.01, value=0.0035,step=0.0001, format="%.4f")

    # Melakukan pengelompokan dan pengurutan dataframe
    output['kelompok'] = output['Score'].apply(kelompokkan_score)
    output = output.sort_values(by='Score', ascending=False)

    # Menghitung presentase untuk masing-masing kelompok
    count = output.groupby('kelompok')[output.columns[0]].count()
    labels = count.index.tolist()
    values = count.values.tolist()
    total = sum(values)
    percentages = [round(value/total*100,2) for value in values]

    # Menampilkan dataframe yang sudah diurutkan dan dikelompokkan
    st.write(output)

    # Membuat diagram pie
    fig = go.Figure(data=[go.Pie(labels=labels, values=percentages)])
    fig.update_layout(title='Presentase Kelompok Keringanan')
    st.plotly_chart(fig, use_container_width=True)

    st.markdown(filedownload(output), unsafe_allow_html=True)

else:
    st.write("Mohon upload kedua file terlebih dahulu.")