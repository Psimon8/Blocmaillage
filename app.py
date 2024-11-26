import streamlit as st
import pandas as pd

# Streamlit app title
st.title("Hierarchy Maillage Processor")

# File upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, usecols="A,C,D,E,F")
    df.columns = ['URL', 'N', 'N+1', 'N+2', 'N+3']  # Rename columns

    # Function to create the maillage
    def create_maillage(df):
        result_df = df.copy()
        result_df['Matches'] = ''  # Add a column for matches

        # Iterate over rows
        for i, row in df.iterrows():
            matches = []
            # Find matches where N, N+1, and N+2 are the same
            for j, compare_row in df.iterrows():
                if i != j and (row['N'] == compare_row['N'] and row['N+1'] == compare_row['N+1'] and row['N+2'] == compare_row['N+2']):
                    matches.append(compare_row['N+3'])
            # Add matches to the result column
            result_df.at[i, 'Matches'] = ', '.join(matches)

        return result_df

    # Apply the maillage logic
    result_df = create_maillage(df)

    # Display the results
    st.write("Processed Results:")
    st.dataframe(result_df)

    # Provide a download link
    output_file = "Processed_Maillage.xlsx"
    result_df.to_excel(output_file, index=False)
    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Processed File",
            data=file,
            file_name="Processed_Maillage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
