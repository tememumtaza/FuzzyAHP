import streamlit as st
import pandas as pd
import numpy as np
from FuzzyAHP import FuzzyAHP

# Fungsi untuk membandingkan nilai
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

# Fungsi untuk menghitung Fuzzy AHP
def calculate_fuzzy_ahp(criteria_df, alternatives_df):
    # Hitung nilai bobot kriteria dengan Fuzzy AHP
    ahp = FuzzyAHP(criteria_df, compare)
    ahp.calculate_weights()
    criteria_weights = ahp.get_weights()

    # Hitung nilai bobot alternatif dengan Fuzzy AHP
    ahp.calculate_alternative_weights(alternatives_df, compare)
    alternative_weights = ahp.get_alternative_weights()

    # Hitung nilai relatif dari setiap alternatif
    ahp.calculate_relative_weights()
    relative_weights = ahp.get_relative_weights()

    return criteria_weights, alternative_weights, relative_weights

# Fungsi untuk menampilkan hasil perhitungan
def display_results(criteria_weights, alternative_weights, relative_weights):
    st.subheader('Hasil Perhitungan')
    
    # Tampilkan nilai bobot kriteria
    st.write('Nilai Bobot Kriteria:')
    st.write(criteria_weights.to_string(index=False))
    
    # Tampilkan nilai bobot alternatif
    st.write('Nilai Bobot Alternatif:')
    st.write(alternative_weights.to_string(index=False))
    
    # Tampilkan nilai relatif dari setiap alternatif
    st.write('Nilai Relatif Alternatif:')
    st.write(relative_weights.to_string(index=False))

# Buat tampilan aplikasi dengan Streamlit
def main():
    st.title('Seleksi Keringanan UKT dengan Fuzzy AHP')
    st.write('Upload file Excel yang berisi data kriteria dan alternatif untuk menghitung Seleksi Keringanan UKT dengan Fuzzy AHP.')
    
    # Tambahkan widget untuk mengupload file excel kriteria
    criteria_file = st.file_uploader('Upload file Excel Kriteria', type=['xlsx'])
    
    # Tambahkan widget untuk mengupload file excel alternatif
    alternatives_file = st.file_uploader('Upload file Excel Alternatif', type=['xlsx'])
    
    if criteria_file is not None and alternatives_file is not None:
        # Muat data dari file excel kriteria ke dalam dataframe
        df = pd.read_excel(uploaded_file, sheet_name='Data', header=0, index_col=0)
        
        # Tampilkan data kriteria dan alternatif
        st.subheader('Data Kriteria dan Alternatif')
        st.write(df.to_string())

        # Tambahkan widget untuk mengatur bobot kriteria
        criteria_weights = st.sidebar.slider('Bobot Kriteria', 0.0, 1.0, (0.5, 0.5), 0.1)

        # Hitung Fuzzy AHP dan tampilkan hasil perhitungan
        if st.button('Hitung'):
            criteria_weights_df = pd.DataFrame({'Kriteria': df.columns, 'Bobot': criteria_weights})
            criteria_weights_df = criteria_weights_df[['Kriteria', 'Bobot']]
            
            criteria_weights, alternative_weights, relative_weights = calculate_fuzzy_ahp(df, criteria_weights_df)
            display_results(criteria_weights, alternative_weights, relative_weights)

# Jalankan aplikasi
if __name__ == '__main__':
    main()