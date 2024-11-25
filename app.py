import streamlit as st
import pandas as pd


# Completer Bloc de Maillage V2 - Maillage inter Nolimit
def fill_empty_rows_with_format_nolimit(sheet):
    last_row = len(sheet.col_values(6))  # Utiliser la colonne F pour obtenir le dernier rang
    e_values = sheet.col_values(5)[1:last_row]  # Utiliser la colonne E pour la correspondance
    ks_range = sheet.range(f"I2:IV{last_row}")  # Étendre la plage jusqu'à la dernière colonne possible (IV)
    values_to_check = [cell.value for cell in ks_range]  # Valeurs de la nouvelle plage étendue
    k_values = sheet.col_values(8)[1:last_row]  # Valeurs de la colonne H
    cells_completed = 0  # Compteur pour les cellules complétées

    for i in range(len(values_to_check)):
        current_e_value = e_values[i]
        matched_rows = [j + 2 for j in range(len(e_values)) if current_e_value == e_values[j]]

        col_index = 0  # Commencer à la première colonne de la plage étendue
        for k in matched_rows:
            target_cell = sheet.cell(i + 2, 9 + col_index)  # 9 est la colonne I
            if target_cell.value == "" and col_index < len(values_to_check[i]):  # Vérifier si la cellule cible est vide et dans la plage
                source_cell = sheet.cell(k, 8)  # 8 est la colonne H
                rich_text = source_cell.value
                if rich_text not in k_values[i]:  # Vérifier si la valeur est déjà présente dans H
                    sheet.update_cell(i + 2, 9 + col_index, rich_text)
                    cells_completed += 1  # Incrémenter le compteur pour chaque cellule complétée
                    col_index += 1  # Passer à la colonne suivante pour la prochaine insertion

    return cells_completed

# Interface Streamlit
st.title('Completer Bloc de Maillage V2 - Maillage inter Nolimit')

uploaded_file = st.file_uploader("Choisissez un fichier XLSX", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Aperçu du fichier importé:")
    st.write(df)