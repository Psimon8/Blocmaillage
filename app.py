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
        result_df['Matches (N+2)'] = [[] for _ in range(len(result_df))]  # Initialize for N+2 matches
        result_df['Matches (N+3)'] = [[] for _ in range(len(result_df))]  # Initialize for N+3 matches

        # Step 1: Apply matching logic for N+2
        for i, row in df.iterrows():
            matches_n2 = []
            for j, compare_row in df.iterrows():
                if i != j and (
                    row['N+1'] == compare_row['N+1'] and  # N+1 must be identical
                    pd.isna(compare_row['N+3']) and  # N+3 must be empty
                    row['N+2'] != compare_row['N+2']  # Ensure distinct N+2 values
                ):
                    matches_n2.append((compare_row["URL"], compare_row["ancre"]))  # Tuple for display and Excel

            # Assign matches
            result_df.at[i, 'Matches (N+2)'] = matches_n2

        # Step 2: Apply matching logic for N+3
        for i, row in df.iterrows():
            matches_n3 = []
            for j, compare_row in df.iterrows():
                if i != j and (
                    row['N+2'] == compare_row['N+2'] and  # N+2 must be identical
                    row['N+3'] != compare_row['N+3']  # Ensure distinct N+3 values
                ):
                    matches_n3.append((compare_row["URL"], compare_row["ancre"]))  # Tuple for display and Excel

            # Assign matches
            result_df.at[i, 'Matches (N+3)'] = matches_n3

        return result_df

    # Apply the maillage logic
    try:
        # Apply the logic to generate matches
        result_df = create_maillage(df)

        # Prepare the result for export and display
        export_df = result_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']].copy()

        # Combine all links from Matches (N+2) and Matches (N+3)
        max_links = 0
        for i, row in result_df.iterrows():
            all_links = row['Matches (N+2)'] + row['Matches (N+3)']  # Combine both lists
            all_links = list(dict.fromkeys(all_links))  # Remove duplicates
            for j, (link_url, link_text) in enumerate(all_links):
                column_name = f'Link {j + 1}'
                if column_name not in export_df.columns:
                    export_df[column_name] = ''
                export_df.at[i, column_name] = f"[{link_text}]({link_url})"  # Markdown for clickable links
            max_links = max(max_links, len(all_links))

        # Rearrange columns: Links to the left, followed by original columns
        link_columns = [f'Link {i + 1}' for i in range(max_links)]
        final_columns = ['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3'] + link_columns
        export_df = export_df[final_columns]

        # Display the processed results in Streamlit as a clickable table
        st.write("Processed Results:")
        st.markdown(
            export_df.to_html(escape=False, index=False), unsafe_allow_html=True
        )  # Render the links as clickable in the table

        # Export the results to an Excel file
        for col in link_columns:
            if col in export_df.columns:
                # Convert Markdown links back to Excel-friendly hyperlinks
                export_df[col] = export_df[col].apply(
                    lambda x: x.replace("[", "").replace("](", ";").replace(")", "").replace("(", "")
                )
                export_df[col] = export_df[col].apply(lambda x: f'=LIEN_HYPERTEXTE{x}')

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
