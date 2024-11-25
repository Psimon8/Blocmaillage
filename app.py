import streamlit as st
import pandas as pd
import io

# Completer Bloc de Maillage V2 - Maillage inter Nolimit
def fill_empty_rows_with_format_nolimit(df):
    last_row = len(df)  # Utiliser la longueur du DataFrame pour obtenir le dernier rang
    if last_row == 0 or df.shape[1] < 9:
        st.error("Le fichier importé ne contient pas suffisamment de données.")
        return df, 0

    e_values = df.iloc[:, 4]  # Utiliser la colonne E pour la correspondance
    k_values = df.iloc[:, 7]  # Valeurs de la colonne H
    cells_completed = 0  # Compteur pour les cellules complétées

    for i in range(last_row):
        current_e_value = e_values.iloc[i]
        matched_rows = df[df.iloc[:, 4] == current_e_value].index.tolist()

        col_index = 0  # Commencer à la première colonne de la plage étendue
        for k in matched_rows:
            if 8 + col_index < df.shape[1] and pd.isna(df.iloc[i, 8 + col_index]):  # Vérifier si la cellule cible est vide
                source_value = df.iloc[k, 7]  # 8 est la colonne H
                if source_value not in k_values.iloc[i]:  # Vérifier si la valeur est déjà présente dans H
                    df.iloc[i, 8 + col_index] = source_value
                    cells_completed += 1  # Incrémenter le compteur pour chaque cellule complétée
                    col_index += 1  # Passer à la colonne suivante pour la prochaine insertion

    return df, cells_completed

# Interface Streamlit
st.title('Completer Bloc de Maillage V2 - Maillage inter Nolimit')

uploaded_file = st.file_uploader("Choisissez un fichier XLSX", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Aperçu du fichier importé:")
    st.write(df)

    if st.button('Exécuter la fonction'):
        df, cells_completed = fill_empty_rows_with_format_nolimit(df)
        st.success(f'{cells_completed} cellules ont été complétées.')

        # Télécharger le fichier modifié
        st.write("Téléchargez le fichier modifié:")
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