import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from urllib.parse import urlparse

# Streamlit app title
st.title("Descending Maillage Processor")

# File upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, usecols="A,B,C,D,E,F")
        df.columns = ['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']
    except Exception as e:
        st.error("Error reading the file. Ensure it contains columns A, B, C, D, E, F.")
        st.stop()

    # Function to compute breadcrumb level
    def get_breadcrumb_level(url):
        path = urlparse(url).path.strip("/")
        return len(path.split("/")) if path else 0

    # Function to create descending maillage
    def create_descending_maillage(df, max_links_per_url=5):
        """
        Logic:
        - Links are created only for URLs at the same breadcrumb level.
        """
        df['Breadcrumb Level'] = df['URL'].apply(get_breadcrumb_level)
        result_df = df.copy()
        result_df['Matches'] = [[] for _ in range(len(result_df))]

        for i, row in df.iterrows():
            matches = []
            source_level = row['Breadcrumb Level']
            source_path = urlparse(row['URL']).path.strip("/")

            for j, compare_row in df.iterrows():
                if i != j:
                    compare_level = compare_row['Breadcrumb Level']
                    compare_path = urlparse(compare_row['URL']).path.strip("/")
                    
                    # Same level and not the exact same page
                    if (
                        compare_level == source_level and
                        not compare_path.startswith(source_path + "/")
                    ):
                        matches.append(f'=LIEN_HYPERTEXTE("{compare_row["URL"]}"; "{compare_row["ancre"]}")')
            
            # Limit the number of links
            result_df.at[i, 'Matches'] = matches[:max_links_per_url]

        return result_df

    try:
        result_df = create_descending_maillage(df, max_links_per_url=5)

        # Prepare result for export
        export_df = result_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3']].copy()
        max_links = 0

        for i, row in result_df.iterrows():
            all_links = row['Matches']
            for j, link in enumerate(all_links):
                column_name = f'Link {j + 1}'
                if column_name not in export_df.columns:
                    export_df[column_name] = ''
                export_df.at[i, column_name] = link
            max_links = max(max_links, len(all_links))

        link_columns = [f'Link {i + 1}' for i in range(max_links)]
        export_df = export_df[['URL', 'ancre', 'N', 'N+1', 'N+2', 'N+3'] + link_columns]

        # Display results
        st.write("Processed Results:")
        st.dataframe(export_df)

        # Visualization
        st.subheader("Link Distribution")
        link_counts = export_df.filter(like='Link').notna().sum(axis=1)
        fig = plt.figure(figsize=(10, 6))
        plt.hist(link_counts, bins=20)
        plt.xlabel("Number of Links per URL")
        plt.ylabel("Frequency")
        st.pyplot(fig)

        # Download option
        output_file = "Processed_Descending_Maillage.xlsx"
        export_df.to_excel(output_file, index=False, engine='openpyxl')
        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Processed File",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"An error occurred while processing: {str(e)}")
