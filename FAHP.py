#Import library yang dibutuhkan
import pandas as pd 
import numpy as np 
import openpyxl
import csv
import math

#Fungsi untuk membaca file dan menyimpan dalam bentuk array / tuple 
def read_excel_file(filename, n):
    df = pd.read_excel(filename)
    items = np.array(df.iloc[:, 0].tolist()) if n == 0 else tuple(zip(df.iloc[:, 0].tolist(), df.iloc[:, n].tolist()))
    return items

criteriaDict = read_excel_file('/content/drive/MyDrive/Colab Notebooks/Data for FAHP/NilaiKriteria.xlsx', 0)
alternativesName = read_excel_file('/content/drive/MyDrive/Colab Notebooks/Data for FAHP/NilaiAlternatif.xlsx', 0)

criteria = read_excel_file('/content/drive/MyDrive/Colab Notebooks/Data for FAHP/NilaiKriteria.xlsx', 1)
for i in range(1, 14):
    exec(f"altc{i} = read_excel_file('/content/drive/MyDrive/Colab Notebooks/Data for FAHP/NilaiAlternatif.xlsx', {i})")

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
for i in range(1, 14):
    alt = eval(f"altc{i}")
    cr = compare(*alt)
    exec(f"altxalt_cr{i} = np.array(cr)")

def isConsistent(matrix, printComp=True):
    mat_len = len(matrix)

    midMatrix = np.array([m[1] for row in matrix for m in row]).reshape(mat_len, mat_len)
    if(printComp): print("mid-value matrix: \n", midMatrix, "\n")
    
    eigenvalue = np.linalg.eigvals(midMatrix)
    lambdaMax = max(eigenvalue)
    if(printComp): print("eigenvalue: ", eigenvalue)
    if(printComp): print("lambdaMax: ", lambdaMax)
    if(printComp): print("\n")

    RIValue = 0.1*(mat_len-1)/mat_len + 0.9
    if(printComp): print("R.I. Value: ", RIValue)

    CIValue = (lambdaMax-mat_len)/(mat_len - 1)
    if(printComp): print("C.I. Value: ", CIValue)

    CRValue = CIValue/RIValue
    if(printComp): print("C.R. Value: ", CRValue)

    if(printComp): print("\n")
    if(CRValue<=0.1):
        if(printComp): print("Matrix reasonably consistent, we could continue")
        return True
    else:
        if(printComp): print("Consistency Ratio is greater than 10%, we need to revise the subjective judgment")
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
        print("Fuzzy Geometric Mean Value: \n", geoMean, "\n")

    #menghitung total fuzzy geometric mean value
    geoMean_sum = np.sum(geoMean, axis=0)

    if(printComp): 
        print("Fuzzy Geometric Mean Sum:", geoMean_sum, "\n")

    #menghitung weights
    weights = np.zeros(matrix_len)

    for i in range(matrix_len):
        weights[i] = np.sum(geoMean[i] / geoMean_sum)

    if(printComp): 
        print("Weights: \n", weights, "\n")

    #menghitung normalized weights
    normWeights = weights / np.sum(weights)

    if(printComp): 
        print("Normalized Weights: ", normWeights,"\n")

    return normWeights

#Parameter: crxcr = Pairwise comparison matrix criteria X criteria, altxalt = Pairwise comparison matrices alternatif X alternatif , 
#       alternativesName = Nama dari setiap alternatif, printComp = opsi untuk menampilkan komputasi konsistensi matrix
def FAHP(crxcr, altxalt, alternativesName, printComp=True):
    

    # Cek konsistensi pairwise comparison matrix criteria x criteria
    crxcr_cons = isConsistent(crxcr, False)
    if(crxcr_cons):
        if(printComp): print("criteria X criteria comparison matrix reasonably consistent, we could continue")
    else: 
        if(printComp): print("criteria X criteria comparison matrix consistency ratio is greater than 10%, we need to revise the subjective judgment")

    # Cek konsistensi pairwise comparison matrix alternative x alternative untuk setiap criteria
    for i, altxalt_cr in enumerate(altxalt):
        isConsistent(altxalt_cr, False)
        if(crxcr_cons):
            if(printComp): print("alternatives X alternatives comparison matrix for criteria",i+1," is reasonably consistent, we could continue")
        else: 
            if(printComp): print("alternatives X alternatives comparison matrix for criteria",i+1,"'s consistency ratio is greater than 10%, we need to revise the subjective judgment")

    if(printComp): print("\n")

    if(printComp): print("criteria X criteria ======================================================\n")
    # Hitung nilai pairwise comparison weight untuk criteria x criteria
    crxcr_weights = pairwiseComp(crxcr, printComp)
    if(printComp): print("criteria X criteria weights: ", crxcr_weights)

    if(printComp): print("\n")
    if(printComp): print("alternative x alternative ======================================================\n")

    # Hitung nilai pairwise comparison weight untuk setiap alternative x alternative dalam setiap criteria
    altxalt_weights = np.zeros((len(altxalt),len(altxalt[0])))
    for i, altxalt_cr in enumerate(altxalt):
        if(printComp): print("alternative x alternative for criteria", criteriaDict[i],"---------------\n")
        altxalt_weights[i] =  pairwiseComp(altxalt_cr, printComp)

    # Transpose matrix altxalt_weights
    altxalt_weights = altxalt_weights.transpose(1, 0)
    if(printComp): print("alternative x alternative weights:")
    if(printComp): print(altxalt_weights)

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

    print("\n Output telah disimpan dalam file 'output_fahp.csv'")

    return output_df
    

#Membuat array numpy untuk altxalt dengan mengambil nilai dari variabel global
altxalt = np.stack([globals()[f"altxalt_cr{i+1}"] for i in range(len(criteriaDict))])

#Memanggil fungsi FAHP dengan parameter yang telah didefinisikan sebelumnya
#printComp di-set False agar tidak menampilkan komputasi konsistensi matrix
output = FAHP(crxcr, altxalt, alternativesName, False)

#Menampilkan rangking alternatif dengan output dari fungsi FAHP
print("\n RANGKING ALTERNATIF:\n", output)