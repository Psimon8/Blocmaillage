import streamlit as st
import pandas as pd
import io
from urllib.parse import urlparse

def get_url_segments(url):
    """Extrait les segments de l'URL après le domaine"""
    parsed = urlparse(url)
    path = parsed.path.strip('/')
    return path.split('/')

def is_parent_url(parent_url, child_url):
    """Vérifie si une URL est parent d'une autre"""
    parent_segments = get_url_segments(parent_url)
    child_segments = get_url_segments(child_url)
    return len(parent_segments) == len(child_segments) - 1 and child_segments[:len(parent_segments)] == parent_segments

def fill_empty_rows_with_format_nolimit(df, max_links):
    last_row = len(df)
    if last_row == 0 or df.shape[1] < 9:
        st.error("Le fichier importé ne contient pas suffisamment de données.")
        return df, 0

    cells_completed = 0

    for i in range(last_row):
        current_url = df.iloc[i, 7]  # URL actuelle
        matched_rows = []
        
        # Recherche séquentielle F->E->D
        for col in [5, 4, 2]:  # F, E, D
            if not matched_rows:
                current_value = df.iloc[i, col]
                matched_rows = df[df.iloc[:, col] == current_value].index.tolist()

        if not matched_rows:
            continue

        # Trier les URLs par niveau
        n_minus_one_urls = []
        same_level_urls = []
        
        for k in matched_rows:
            url = df.iloc[k, 7]
            if url == current_url:
                continue
                
            if is_parent_url(url, current_url):
                n_minus_one_urls.append(url)
            elif len(get_url_segments(url)) == len(get_url_segments(current_url)):
                same_level_urls.append(url)

        # Ajouter d'abord les URLs N-1, puis même niveau
        col_index = 0
        all_urls = n_minus_one_urls + same_level_urls
        
        for url in all_urls:
            if col_index >= max_links:
                break
            if 8 + col_index < df.shape[1] and pd.isna(df.iloc[i, 8 + col_index]):
                if url not in df.iloc[i, 7:8+col_index].values:
                    df.iloc[i, 8 + col_index] = url
                    cells_completed += 1
                    col_index += 1

    return df, cells_completed

# Interface Streamlit
st.title('Completer Bloc de Maillage V2 - Maillage inter Nolimit')

# Slider pour le nombre maximum de liens
max_links = st.slider('Nombre maximum de liens', min_value=5, max_value=50, value=20)

uploaded_file = st.file_uploader("Choisissez un fichier XLSX", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Aperçu du fichier importé:")
    st.write(df)

    if st.button('Exécuter la fonction'):
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