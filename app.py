import streamlit as st
import pandas as pd
import io
import re

def get_url_segments(url):
    """Extrait les segments de l'URL après le domaine et retire les chiffres"""
    segments = url.split('/')[3:] # Ignorer http:, '', domain
    segments = [re.sub(r'[0-9]', '', segment) + '-' for segment in segments if segment]
    return segments

def get_matching_urls(df, current_index, segment_col):
    """Trouve les URLs qui partagent le même segment"""
    current_segment = df.iloc[current_index, segment_col]
    if pd.isna(current_segment):
        return []
    matches = df[df.iloc[:, segment_col] == current_segment].index.tolist()
    return [idx for idx in matches if idx != current_index]

def fill_empty_rows_with_format_nolimit(df, max_links):
    last_row = len(df)
    if last_row == 0 or df.shape[1] < 9:
        st.error("Le fichier importé ne contient pas suffisamment de données.")
        return df, 0

    cells_completed = 0

    # Maillage
    for i in range(last_row):
        current_url = df.iloc[i, 7]  # URL actuelle
        current_segments = get_url_segments(current_url)
        if len(current_segments) < 3:
            continue

        matched_rows = []
        
        # Recherche séquentielle F->E->D
        for col in [5, 4, 2]:
            if not matched_rows:
                matched_rows = get_matching_urls(df, i, col)

        if not matched_rows:
            continue

        col_index = 0
        for k in matched_rows:
            if col_index >= max_links:
                break
            if col_index < max_links and 8 + col_index < df.shape[1]:
                source_value = df.iloc[k, 7]
                if source_value not in df.iloc[i, 7:8+col_index].values:
                    df.iloc[i, 8 + col_index] = source_value
                    cells_completed += 1
                    col_index += 1

    return df, cells_completed

# Interface Streamlit
st.title('Completer Bloc de Maillage V2 - Maillage inter Nolimit')

max_links = st.slider('Nombre maximum de liens', min_value=5, max_value=50, value=20)

uploaded_file = st.file_uploader("Choisissez un fichier XLSX", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Aperçu du fichier importé:")
    st.write(df)

    if st.button('Exécuter la fonction'):
        # Exécuter les fonctions nécessaires
        for i in range(len(df)):
            segments = get_url_segments(df.iloc[i, 0])
            if len(segments) >= 3:
                df.iloc[i, 2] = '/' + segments[-3] # Colonne C
                df.iloc[i, 3] = '/' + segments[-2] # Colonne D
                df.iloc[i, 4] = '/' + segments[-1] # Colonne E
        
        df, cells_completed = fill_empty_rows_with_format_nolimit(df, max_links)
        st.success(f'{cells_completed} cellules ont été complétées.')

        # Télécharger le fichier modifié
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)
        st.download_button(
            label="Télécharger le fichier XLSX",
            data=buffer,
            file_name="fichier_modifie.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )