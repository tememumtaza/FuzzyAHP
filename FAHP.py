import pandas as pd 
import numpy as np 
import base64
import streamlit as st

# Fungsi untuk membaca file dan menyimpan dalam bentuk array / tuple 
def read_excel_file(filename, n):
    df = pd.read_excel(filename)
    items = np.array(df.iloc[:, 0].tolist()) if n == 0 else tuple(zip(df.iloc[:, 0].tolist(), df.iloc[:, n].tolist()))
    return items

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