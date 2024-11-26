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
        result_df['Matches'] = [[] for _ in range(len(result_df))]  # Initialize as lists for storing multiple matches

        # Iterate over rows
        for i, row in df.iterrows():
            matches = []
            # Find matches where N, N+1, and N+2 are the same
            for j, compare_row in df.iterrows():
                if i != j and (
                    row['N'] == compare_row['N'] and 
                    row['N+1'] == compare_row['N+1'] and 
                    row['N+2'] == compare_row['N+2']
                ):
                    # Create a hyperlink for the match
                    hyperlink = f'=HYPERLINK("{compare_row["URL"]}", "{compare_row["ancre"]}")'
                    matches.append(hyperlink)  # Add hyperlink to matches

            # Remove duplicates and assign matches
            result_df.at[i, 'Matches'] = list(set(matches))  # Ensure unique matches

        return result_df

    # Apply the maillage logic
    try:
        # Apply the logic to generate matches
        result_df = create_maillage(df)

        # Prepare the result for export
        export_df = result_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']].copy()

        # Unpack matches into separate columns
        max_matches = result_df['Matches'].apply(len).max()  # Get the maximum number of matches
        for match_idx in range(max_matches):
            export_df[f'Match {match_idx + 1}'] = result_df['Matches'].apply(
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
