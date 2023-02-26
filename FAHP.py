import streamlit as st
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

def main():
    st.title("Aplikasi FAHP")

    st.sidebar.header("Pilih file data")
    criteria_file = st.sidebar.file_uploader("Upload file kriteria", type=["xlsx"])
    alternatives_file = st.sidebar.file_uploader("Upload file alternatif", type=["xlsx"])
    
    if criteria_file and alternatives_file:
        criteriaDict = read_excel_file(criteria_file, 0)
        alternativesName = read_excel_file(alternatives_file, 0)

        criteria = read_excel_file(criteria_file, 1)
        for i in range(1, 14):
            exec(f"altc{i} = read_excel_file(alternatives_file, {i})")

        crxcr = np.array(compare(*criteria))
        for i in range(1, 14):
            alt = eval(f"altc{i}")
            cr = compare(*alt)
            exec(f"altxalt_cr{i} = np.array(cr)")

        st.header("Data Kriteria")
        st.write(criteriaDict)

        st.header("Data Alternatif")
        st.write(alternativesName)

        st.header("Matriks Perbandingan Kriteria")
        st.write(crxcr)

        for i in range(1, 14):
            st.header(f"Matriks Perbandingan Alternatif-{i}")
            alt_cr = eval(f"altxalt_cr{i}")
            st.write(alt_cr)

if __name__ == "__main__":
    main()
