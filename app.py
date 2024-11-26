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
                    hyperlink_n3 = f'=HYPERLINK("{compare_row["URL"]}"; "{compare_row["ancre"]}")'
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
                    hyperlink_n2 = f'=HYPERLINK("{compare_row["URL"]}"; "{compare_row["ancre"]}")'
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

        # Unpack matches for N+3
        max_matches_n3 = result_df['Matches (N+3)'].apply(len).max()  # Get max matches for N+3
        for match_idx in range(max_matches_n3):
            export_df[f'Match N+3 - {match_idx + 1}'] = result_df['Matches (N+3)'].apply(
                lambda x: x[match_idx] if len(x) > match_idx else ''  # Add matches to separate columns
            )

        # Unpack matches for N+2
        max_matches_n2 = result_df['Matches (N+2)'].apply(len).max()  # Get max matches for N+2
        for match_idx in range(max_matches_n2):
            export_df[f'Match N+2 - {match_idx + 1}'] = result_df['Matches (N+2)'].apply(
                lambda x: x[match_idx] if len(x) > match_idx else ''  # Add matches to separate columns
            )

        # Display the results
        st.write("Processed Results:")
        st.dataframe(export_df)

        # Export the results to an Excel file
        output_file = "Processed_Maillage.xlsx"
        export_df.to_excel(output_file, index=False, engine='openpyxl')

        # Provide a download link
        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Processed File",
                data=file,
                file_name="Processed_Maillage.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
