import streamlit as st  # Selected Line 1
import pandas as pd
import io

# Completer Bloc de Maillage V2 - Maillage inter Nolimit
def fill_empty_rows_with_format_nolimit(df):
    last_row = len(df)  # Use the length of the DataFrame to get the last row
    # Check if the DataFrame has fewer than 9 columns, as the function requires at least 9 columns to operate correctly
    if last_row == 0 or df.shape[1] < 9:
        st.error("Le fichier importé ne contient pas suffisamment de données.")
        return df, 0
    e_values = df.iloc[:, 5]  # Use column F for matching
    e_values = df.iloc[:, 5]  # Utiliser la colonne F pour la correspondance
    k_values = df.iloc[:, 7]  # Valeurs de la colonne H
    cells_completed = 0  # Compteur pour les cellules complétées

    # Utiliser des opérations vectorisées pour trouver les correspondances
    for col in [5, 4]:  # Vérifier d'abord la colonne F, puis la colonne E
        matched_rows = df.groupby(df.iloc[:, col]).apply(lambda x: x.index.tolist()).to_dict()
        for i in range(last_row):
            current_e_value = df.iloc[i, col]
            current_url = df.iloc[i, 7]  # URL actuelle
            if current_e_value in matched_rows:
                col_index = 0  # Commencer à la première colonne de la plage étendue
                for k in matched_rows[current_e_value]:
                    if col_index >= 20:  # Limiter le nombre de liens ajoutés à 20
                        break
                    if 8 + col_index < df.shape[1] and pd.isna(df.iloc[i, 8 + col_index]):  # Vérifier si la cellule cible est vide
                        source_value = df.iloc[k, 7]  # 8 est la colonne H
            if source_value != current_url and source_value not in k_values.iloc[i]:  # Vérifier si la valeur est déjà présente dans H et n'est pas l'URL actuelle
                if 8 + col_index < df.shape[1] and pd.isna(df.iloc[i, 8 + col_index]):  # Check if the target cell is empty
                    source_value = df.iloc[k, 7]  # 8 is column H
                    df.iloc[i, 8 + col_index] = source_value
                    cells_completed += 1  # Incrémenter le compteur pour chaque cellule complétée
                    col_index += 1  # Passer à la colonne suivante pour la prochaine insertion

    # If the script is at the last level (column F is empty), link all values in H that share the same values in C and D
    empty_f_rows = df[pd.isna(df.iloc[:, 5])]
    if not empty_f_rows.empty:
        grouped = empty_f_rows.groupby([df.iloc[:, 2], df.iloc[:, 3]])
        for (c_value, d_value), group in grouped:
            matched_rows = df[(df.iloc[:, 2] == c_value) & (df.iloc[:, 3] == d_value)].index.tolist()
            for i in group.index:
                col_index = 0
                for k in matched_rows:
                    if col_index >= 20:  # Limiter le nombre de liens ajoutés à 20
                        break
                    if 8 + col_index < df.shape[1] and pd.isna(df.iloc[i, 8 + col_index]):  # Vérifier si la cellule cible est vide
                        source_value = df.iloc[k, 7]  # 8 est la colonne H
                        if source_value != df.iloc[i, 7] and source_value not in k_values.iloc[i]:  # Vérifier si la valeur est déjà présente dans H et n'est pas l'URL actuelle
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