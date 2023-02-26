import streamlit as st
import FAHP

# Set Streamlit page title
st.set_page_config(page_title="FAHP Calculator")

# Define Streamlit app layout
st.title("FAHP Calculator")
st.write("This app calculates the fuzzy analytic hierarchy process (FAHP) for a given set of criteria and alternatives.")

# Add input fields for criteria and alternatives
criteria = st.text_input("Enter criteria, separated by commas")
alternatives = st.text_input("Enter alternatives, separated by commas")

# Add input fields for criteria weights
st.write("Enter criteria weights:")
weights = {}
for criterion in criteria.split(","):
    weight = st.slider(criterion, 0.0, 1.0, 0.5, 0.01)
    weights[criterion.strip()] = weight

# Add input fields for pairwise comparison matrices
st.write("Enter pairwise comparison matrices:")
matrices = {}
for criterion in criteria.split(","):
    matrix = []
    for i in range(len(alternatives.split(","))):
        row = []
        for j in range(len(alternatives.split(","))):
            if i == j:
                value = 1.0
            elif j > i:
                value = st.slider(criterion + " - " + alternatives.split(",")[i] + " vs. " + alternatives.split(",")[j], 0.0, 1.0, 0.5, 0.01)
            else:
                value = 1.0 / matrices[criterion][j][i]
            row.append(value)
        matrix.append(row)
    matrices[criterion.strip()] = matrix

# Calculate FAHP result
if st.button("Calculate FAHP"):
    result = FAHP.calculate_fahp(criteria.split(","), alternatives.split(","), matrices, weights)
    st.write("FAHP Result:")
    st.write(result)