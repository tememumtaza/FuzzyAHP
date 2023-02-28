#Import library yang dibutuhkan
import pandas as pd 
import numpy as np 
import base64
import xlsxwriter
from io import BytesIO
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="FAHP : Seleksi Keringanan UKT", layout="wide",menu_items=None)

#Fungsi untuk membaca file dan menyimpan dalam bentuk array / tuple 
def read_excel_file(filename, n):
    df = pd.read_excel(filename)
    items = np.array(df.iloc[:, 0].tolist()) if n == 0 else tuple(zip(df.iloc[:, 0].tolist(), df.iloc[:, n].tolist()))
    return items

def filedownload(df, filename='output.xlsx'):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter') # ubah engine ke xlsxwriter
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}">Download file</a>'

def isConsistent(matrix, printComp=True):
    mat_len = len(matrix)

    midMatrix = np.array([m[1] for row in matrix for m in row]).reshape(mat_len, mat_len)
    if(printComp): st.write("mid-value matrix: \n", midMatrix, "\n")
    
    eigenvalue = np.linalg.eigvals(midMatrix)
    lambdaMax = max(eigenvalue)
    if(printComp): st.write("eigenvalue: ", eigenvalue)
    if(printComp): st.write("lambdaMax: ", lambdaMax)
    if(printComp): st.write("\n")

    RIValue = 0.1*(mat_len-1)/mat_len + 0.9
    if(printComp): st.write("R.I. Value: ", RIValue)

    CIValue = (lambdaMax-mat_len)/(mat_len - 1)
    if(printComp): st.write("C.I. Value: ", CIValue)

    CRValue = CIValue/RIValue
    if(printComp): st.write("C.R. Value: ", CRValue)

    if(printComp): st.write("\n")
    if(CRValue<=0.1):
        if(printComp): st.write("Matrix reasonably consistent, we could continue")
        return True
    else:
        if(printComp): st.write("Consistency Ratio is greater than 10%, we need to revise the subjective judgment")
        return False

#Parameter: matrix = Matrix yang akan dihitung konsistensinya, printComp = opsi untuk menampilkan komputasi konsistensi matrix
def pairwiseComp(matrix, printComp=True):
    matrix_len = len(matrix)

    #menghitung fuzzy geometric mean value
    geoMean = np.zeros((matrix_len,3))

    for i in range(matrix_len):
        for j in range(3):
            temp = 1
            for tfn in matrix[i]:
                temp *= tfn[j]
            temp = pow(temp, 1/matrix_len)
            geoMean[i,j] = temp

    if(printComp): 
        st.write("Fuzzy Geometric Mean Value: \n", geoMean, "\n")

    #menghitung total fuzzy geometric mean value
    geoMean_sum = np.sum(geoMean, axis=0)

    if(printComp): 
        st.write("Fuzzy Geometric Mean Sum:", geoMean_sum, "\n")

    #menghitung weights
    weights = np.zeros(matrix_len)

    for i in range(matrix_len):
        weights[i] = np.sum(geoMean[i] / geoMean_sum)

    if(printComp): 
        st.write("Weights: \n", weights, "\n")

    #menghitung normalized weights
    normWeights = weights / np.sum(weights)

    if(printComp): 
        st.write("Normalized Weights: ", normWeights,"\n")

    return normWeights

#Parameter: crxcr = Pairwise comparison matrix criteria X criteria, altxalt = Pairwise comparison matrices alternatif X alternatif , 
#       alternativesName = Nama dari setiap alternatif, printComp = opsi untuk menampilkan komputasi konsistensi matrix
def FAHP(crxcr, altxalt, alternativesName, printComp=True):
    

    # Cek konsistensi pairwise comparison matrix criteria x criteria
    crxcr_cons = isConsistent(crxcr, False)
    if(printComp): st.write(f'<p style="font-size:28px">MENGHITUNG KONSISTENSI MATRIKS : \n</p>', unsafe_allow_html=True)
    if(crxcr_cons):
        if(printComp): st.write("criteria X criteria comparison matrix reasonably consistent, we could continue")
    else: 
        if(printComp): st.write("criteria X criteria comparison matrix consistency ratio is greater than 10%, we need to revise the subjective judgment")

    # Cek konsistensi pairwise comparison matrix alternative x alternative untuk setiap criteria
    for i, altxalt_cr in enumerate(altxalt):
        isConsistent(altxalt_cr, False)
        if(crxcr_cons):
            if(printComp): st.write("alternatives X alternatives comparison matrix for criteria",i+1," is reasonably consistent, we could continue")
        else: 
            if(printComp): st.write("alternatives X alternatives comparison matrix for criteria",i+1,"'s consistency ratio is greater than 10%, we need to revise the subjective judgment")

    if(printComp): st.write("\n")

    if(printComp): st.write(f'<p style="font-size:28px">KRITERIA X KRITERIA : \n</p>', unsafe_allow_html=True)
    # Hitung nilai pairwise comparison weight untuk criteria x criteria
    crxcr_weights = pairwiseComp(crxcr, printComp)
    if(printComp): st.write("criteria X criteria weights: ", crxcr_weights)

    if(printComp): st.write("\n")
    if(printComp): st.write(f'<p style="font-size:28px">ALTERNATIF X ALTERNATIF : \n</p>', unsafe_allow_html=True)

    # Hitung nilai pairwise comparison weight untuk setiap alternative x alternative dalam setiap criteria
    altxalt_weights = np.zeros((len(altxalt),len(altxalt[0])))
    for i, altxalt_cr in enumerate(altxalt):
        if(printComp): st.write("alternative x alternative untuk criteria", criteriaDict[i],"\n")
        altxalt_weights[i] =  pairwiseComp(altxalt_cr, printComp)

    # Transpose matrix altxalt_weights
    altxalt_weights = altxalt_weights.transpose(1, 0)
    if(printComp): st.write("alternative x alternative weights:")
    if(printComp): st.write(altxalt_weights)

    # Hitung nilai jumlah dari perkalian crxcr_weights dengan altxalt_weights pada setiap kolom
    sumProduct = np.zeros(len(altxalt[0]))
    for i  in range(len(altxalt[0])):
        sumProduct[i] = np.dot(crxcr_weights, altxalt_weights[i])

    # Buat output dataframe
    output_df = pd.DataFrame(data=[alternativesName, sumProduct]).T
    output_df = output_df.rename(columns={0: "Alternatif", 1: "Score"})
    output_df = output_df.sort_values(by=['Score'],ascending = False)
    output_df.index = np.arange(1,len(output_df)+1)

    # Simpan DataFrame ke dalam file CSV
    output_df.to_csv("\n output_fahp.csv", index=False)

    return output_df

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