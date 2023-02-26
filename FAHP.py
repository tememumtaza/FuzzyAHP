import streamlit as st
import pandas as pd 
import numpy as np 
import openpyxl

#Fungsi untuk membaca file dan menyimpan dalam bentuk array / tuple 
def read_excel_file(df, n):
    items = np.array(df.iloc[:, 0].tolist()) if n == 0 else tuple(zip(df.iloc[:, 0].tolist(), df.iloc[:, n].tolist()))
    return items

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

# Fungsi utama Streamlit
def main():
    st.title("Aplikasi untuk Menghitung Nilai FAHP")

    # Minta pengguna mengunggah file NilaiKriteria.xlsx
    st.write("### Mengunggah file NilaiKriteria.xlsx")
    file_criteria = st.file_uploader("Silakan unggah file Excel", type=["xlsx"])
    if file_criteria is not None:
        df_criteria = pd.read_excel(file_criteria)
        criteriaDict = read_excel_file(df_criteria, 0)
        criteria = read_excel_file(df_criteria, 1)

        # Minta pengguna mengunggah file NilaiAlternatif.xlsx
        st.write("### Mengunggah file NilaiAlternatif.xlsx")
        file_alternatives = st.file_uploader("Silakan unggah file Excel", type=["xlsx"])
        if file_alternatives is not None:
            df_alternatives = pd.read_excel(file_alternatives)
            alternativesName = read_excel_file(df_alternatives, 0)

            # Hitung nilai pairwise comparisons
            crxcr = np.array(compare(*criteria))
            alt_matrices = {}
            for i in range(1, df_alternatives.shape[1]):
                alt = read_excel_file(df_alternatives.iloc[:, i:i+1], 0)
                alt_matrices[i] = np.array(compare(*alt))

            st.write("### Hasil Perhitungan")
            st.write("**Pairwise Comparison Matrix of Criteria**")
            st.write(crxcr)
            st.write("")

            for i in range(1, df_alternatives.shape[1]):
                st.write(f"**Pairwise Comparison Matrix of Alternatives ({df_alternatives.columns[i]})**")
                st.write(alt_matrices[i])
                st.write("")

if __name__ == "__main__":
    main()