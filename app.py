import streamlit as st
import pandas as pd

# Streamlit app title
st.title("Hierarchy Maillage Processor")

# File upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, usecols="A,B,C,D,E,F")
    df.columns = ['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']  # Rename columns

    # Function to create the maillage
    def create_maillage(df):
        result_df = df.copy()
        result_df['Matches (N+3)'] = [[] for _ in range(len(result_df))]  # Initialize for N+3 matches
        result_df['Matches (N+2)'] = [[] for _ in range(len(result_df))]  # Initialize for N+2 matches

        # Step 1: Apply matching logic for N+3
        for i, row in df.iterrows():
            matches_n3 = []
            # Find matches where N, N+1, and N+2 are the same
            for j, compare_row in df.iterrows():
                if i != j and (
                    row['N'] == compare_row['N'] and 
                    row['N+1'] == compare_row['N+1'] and 
                    row['N+2'] == compare_row['N+2']
                ):
                    # Create a hyperlink for the match
                    hyperlink_n3 = f'=LIEN_HYPERTEXTE("{compare_row["URL"]}"; "{compare_row["ancre"]}")'
                    matches_n3.append(hyperlink_n3)

            # Remove duplicates and assign matches
            result_df.at[i, 'Matches (N+3)'] = list(set(matches_n3))

        # Step 2: Apply matching logic for N+2
        for i, row in df.iterrows():
            matches_n2 = []
            # Find matches where N and N+1 are the same (ignoring N+3)
            for j, compare_row in df.iterrows():
                if i != j and (
                    row['N'] == compare_row['N'] and 
                    row['N+1'] == compare_row['N+1']
                ):
                    # Create a hyperlink for the match
                    hyperlink_n2 = f'=LIEN_HYPERTEXTE("{compare_row["URL"]}"; "{compare_row["ancre"]}")'
                    matches_n2.append(hyperlink_n2)

            # Remove duplicates and assign matches
            result_df.at[i, 'Matches (N+2)'] = list(set(matches_n2))

        return result_df

    # Apply the maillage logic
    try:
        # Apply the logic to generate matches
        result_df = create_maillage(df)

        # Prepare the result for export
        export_df = result_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']].copy()

        # Unpack matches for N+3 and N+2 into unified "Link" columns
        links = []
        for matches in result_df['Matches (N+3)']:
            links.extend(matches)
        for matches in result_df['Matches (N+2)']:
            links.extend(matches)

        # Create a new DataFrame for links
        unique_links = list(dict.fromkeys(links))  # Remove duplicates
        for i in range(len(link Nites)):
