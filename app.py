import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Streamlit app title
st.title("Hierarchy Maillage Processor")

# File upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Read the Excel file and ensure the necessary columns exist
    try:
        df = pd.read_excel(uploaded_file, usecols="A,B,C,D,E,F")
        df.columns = ['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']  # Rename columns
    except Exception as e:
        st.error("Error reading the file. Ensure it contains columns A, B, C, D, E, F.")
        st.stop()

    # Function to create the maillage
    def create_maillage(df, max_links_per_url=5):  # Ajout du paramètre de limitation
        """
        Create maillage logic:
        1. Matches for N+2 are created where N+1 is identical, and N+3 is empty.
        2. Matches for N+3 are created where N+2 is identical.
        """
        result_df = df.copy()
        result_df['Matches (N+2)'] = [[] for _ in range(len(result_df))]
        result_df['Matches (N+3)'] = [[] for _ in range(len(result_df))]

        # Step 1: Apply matching logic for N+2
        for i, row in df.iterrows():
            matches_n2 = []
            for j, compare_row in df.iterrows():
                if (
                    i != j and
                    row['N+1'] == compare_row['N+1'] and  # N+1 must be identical
                    pd.isna(compare_row['N+3']) and  # N+3 must be empty
                    row['N+2'] != compare_row['N+2']  # Ensure distinct N+2
                ):
                    matches_n2.append(f'=LIEN_HYPERTEXTE("{compare_row["URL"]}"; "{compare_row["ancre"]}")')

            # Assign matches
            result_df.at[i, 'Matches (N+2)'] = matches_n2

        # Step 2: Apply matching logic for N+3
        for i, row in df.iterrows():
            matches_n3 = []
            for j, compare_row in df.iterrows():
                if (
                    i != j and
                    row['N+2'] == compare_row['N+2'] and  # N+2 must be identical
                    row['N+3'] != compare_row['N+3']  # Ensure distinct N+3
                ):
                    matches_n3.append(f'=LIEN_HYPERTEXTE("{compare_row["URL"]}"; "{compare_row["ancre"]}")')

            # Assign matches
            result_df.at[i, 'Matches (N+3)'] = matches_n3

        # Limiter le nombre de liens
        for i, row in result_df.iterrows():
            matches_n2 = result_df.at[i, 'Matches (N+2)'][:max_links_per_url]
            matches_n3 = result_df.at[i, 'Matches (N+3)'][:max_links_per_url]
            result_df.at[i, 'Matches (N+2)'] = matches_n2
            result_df.at[i, 'Matches (N+3)'] = matches_n3

        return result_df

    # Apply the maillage logic
    try:
        # Apply the logic to generate matches
        result_df = create_maillage(df, max_links_per_url=5)

        # Prepare the result for export
        export_df = result_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']].copy()

        # Combine all links from Matches (N+2) and Matches (N+3)
        max_links = 0
        for i, row in result_df.iterrows():
            all_links = row['Matches (N+2)'] + row['Matches (N+3)']  # Combine both lists
            all_links = list(dict.fromkeys(all_links))  # Remove duplicates
            for j, link in enumerate(all_links):
                column_name = f'Link {j + 1}'
                if column_name not in export_df.columns:
                    export_df[column_name] = ''
                export_df.at[i, column_name] = link
            max_links = max(max_links, len(all_links))

        # Rearrange columns: Links at the end of original columns
        link_columns = [f'Link {i + 1}' for i in range(max_links)]
        export_df = export_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3'] + link_columns]

        # Display the processed results in Streamlit
        st.write("Processed Results:")
        st.dataframe(export_df)

        # Ajout de la visualisation
        st.subheader("Distribution des liens")
        link_counts = export_df.filter(like='Link').notna().sum(axis=1)
        fig = plt.figure(figsize=(10, 6))
        plt.hist(link_counts, bins=20)
        plt.xlabel("Nombre de liens par URL")
        plt.ylabel("Fréquence")
        st.pyplot(fig)

        # Affichage avec liens cliquables
        def make_clickable(url):
            return f'<a href="{url}" target="_blank">{url}</a>'

        clickable_df = export_df.copy()
        clickable_df['URL'] = clickable_df['URL'].apply(make_clickable)
        st.write("Processed Results:")
        st.write(clickable_df.to_html(escape=False), unsafe_allow_html=True)

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
        st.error(f"An error occurred while processing: {str(e)}")
