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

        # Combine all links from Matches (N+3) and Matches (N+2)
        max_links = 0
        for i, row in result_df.iterrows():
            all_links = row['Matches (N+3)'] + row['Matches (N+2)']  # Combine both lists
            all_links = list(dict.fromkeys(all_links))  # Remove duplicates
            for j, link in enumerate(all_links):
                column_name = f'Link {j + 1}'
                if column_name not in export_df.columns:
                    export_df[column_name] = ''
                export_df.at[i, column_name] = link
            max_links = max(max_links, len(all_links))

        # Rearrange columns: Links to the left, followed by original columns
        link_columns = [f'Link {i + 1}' for i in range(max_links)]
        final_columns = ['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3'] + link_columns
        export_df = export_df[final_columns]

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
