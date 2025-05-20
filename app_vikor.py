import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt

st.set_page_config(page_title="VIKOR Method for MCDM", layout="wide")
st.title("VIKOR (VIseKriterijumska Optimizacija I Kompromisno Resenje)")

st.markdown("""
### Instructions
- Upload an Excel file with **3 sheets**:
  1. `Data`: Alternatives with criteria values
  2. `CriteriaInfo`: Each criterion's type (`benefit`/`cost`) and weight
  3. `TargetValues` *(optional)*: Not needed for VIKOR but kept for format compatibility
""")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df_data = pd.read_excel(uploaded_file, sheet_name="Data")
        df_criteria = pd.read_excel(uploaded_file, sheet_name="CriteriaInfo")

        alternatives = df_data.iloc[:, 0].tolist()
        matrix = df_data.iloc[:, 1:].to_numpy(dtype=float)
        criteria = df_data.columns[1:]
        weights = df_criteria['Weight'].to_numpy(dtype=float)
        types = df_criteria['Type'].str.lower().tolist()

        # Step 1: Normalize
        norm_matrix = np.zeros_like(matrix, dtype=float)
        for j in range(len(criteria)):
            col = matrix[:, j]
            if types[j] == 'benefit':
                norm_matrix[:, j] = (col - col.min()) / (col.max() - col.min())
            elif types[j] == 'cost':
                norm_matrix[:, j] = (col.max() - col) / (col.max() - col.min())

        # Step 2: Weighted normalized matrix
        weighted_matrix = norm_matrix * weights

        # Step 3: Compute S, R
        S = weighted_matrix.sum(axis=1)
        R = weighted_matrix.max(axis=1)

        # Step 4: Compute Q
        v = 0.5
        S_min, S_max = S.min(), S.max()
        R_min, R_max = R.min(), R.max()
        Q = v * (S - S_min) / (S_max - S_min + 1e-9) + (1 - v) * (R - R_min) / (R_max - R_min + 1e-9)

        # Step 5: Ranking
        ranks = np.argsort(Q) + 1

        df_result = pd.DataFrame({
            'Alternative': alternatives,
            'S': S,
            'R': R,
            'Q': Q,
            'Rank': ranks
        })

        st.subheader("Final VIKOR Scores and Rankings")
        st.dataframe(df_result)

        # Plot
        st.subheader("VIKOR Q Values")
        fig, ax = plt.subplots()
        ax.bar(df_result['Alternative'], df_result['Q'], color='skyblue')
        ax.set_ylabel('Q Score (Lower is Better)')
        st.pyplot(fig)

        # Downloadable Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_data.to_excel(writer, sheet_name='1_RawData', index=False)
            pd.DataFrame(norm_matrix, columns=criteria).insert(0, 'Alternative', alternatives)
            pd.DataFrame(norm_matrix, columns=criteria).to_excel(writer, sheet_name='2_Normalized', index=False)
            pd.DataFrame(weighted_matrix, columns=criteria).insert(0, 'Alternative', alternatives)
            pd.DataFrame(weighted_matrix, columns=criteria).to_excel(writer, sheet_name='3_Weighted', index=False)
            df_result.to_excel(writer, sheet_name='4_FinalScores', index=False)
        st.download_button("📥 Download Full Excel Report", data=output.getvalue(), file_name="VIKOR_Results.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
