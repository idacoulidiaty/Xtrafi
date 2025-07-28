
import os
import tempfile
import pandas as pd
import streamlit as st
from config import *

def get_latest_excel_file(folder):
    files = [f for f in os.listdir(folder) if f.endswith((".xlsx", ".xls"))]
    if not files:
        return None
    files.sort(key=lambda x: os.path.getctime(os.path.join(folder, x)), reverse=True)
    return os.path.join(folder, files[0])

def load_excel_data_dynamic_start(file_path):
    preview_df = pd.read_excel(file_path, header=None, engine="calamine")
    start_row_idx = preview_df[0].astype(str).str.contains(
        "Liste des données qui seront utilisées pour", case=False, na=False
    )
    if start_row_idx.any():
        start_index = start_row_idx[start_row_idx].index[0] + 1
        return pd.read_excel(file_path, skiprows=start_index, engine="calamine")
    else:
        raise ValueError("❌ Ligne de début de données introuvable dans le fichier.")





def load_data(uploaded_file, selected_file):
    df_onglet_1 = df_onglet_2 = df_onglet_3 = source = None

    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getbuffer())
            temp_file_path = tmp.name
        try:
            label, date_value = get_date_generation_restitution(temp_file_path)
            df_onglet_1 = load_excel_data_dynamic_start(temp_file_path)
            df_onglet_2 = pd.read_excel(temp_file_path, sheet_name=onglet_2, engine="calamine")
            df_onglet_3 = pd.read_excel(temp_file_path, sheet_name=onglet_3, engine="calamine")
            source = f"🔼 Fichier Uploadé - {label} : {date_value}"
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement du fichier uploadé : {e}")

    elif selected_file:
        file_path = os.path.join(WATCHED_FOLDER, selected_file)
        try:
            label, date_value = get_date_generation_restitution(file_path)
            df_onglet_1 = load_excel_data_dynamic_start(file_path)
            df_onglet_2 = pd.read_excel(file_path, sheet_name=onglet_2, engine="calamine")
            df_onglet_3 = pd.read_excel(file_path, sheet_name=onglet_3, engine="calamine")
            source = f"✅ Fichier sélectionné - {label} : {date_value}"
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement du fichier sélectionné : {e}")

    return df_onglet_1, df_onglet_2, df_onglet_3, source



def get_date_generation_restitution(file_path):
    preview_df = pd.read_excel(file_path,sheet_name=onglet_1, header=None, engine="calamine")
    for row_idx in range(len(preview_df)):
        for col_idx in range(len(preview_df.columns)):
            cell_value = str(preview_df.iat[row_idx, col_idx])
            if "Date génération" in cell_value:
                label = preview_df.iat[row_idx, col_idx]
                value = preview_df.iat[row_idx + 1, col_idx]

    if label and value:
        return label, value
    else:
        raise ValueError("❌ 'Date génération restitution' introuvable.")

