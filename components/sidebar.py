import streamlit as st
import os
from config import WATCHED_FOLDER
from data_loader import get_latest_excel_file

def sidebar_file_selection():

    st.sidebar.header("📁 Importer un fichier")
    uploaded_file = st.sidebar.file_uploader("Téléversez un fichier Excel", type=["xlsx", "xls"])

    watched_files = [f for f in os.listdir(WATCHED_FOLDER) if f.endswith((".xlsx", ".xls"))]
    selected_file = None
    latest_file_path = get_latest_excel_file(WATCHED_FOLDER)
    default_file = os.path.basename(latest_file_path) if latest_file_path else None

    if watched_files:
        selected_file = st.sidebar.selectbox(
            "📂 Ou choisissez un fichier existant dans le dossier surveillé :",
            options=watched_files,
            index=watched_files.index(default_file) if default_file in watched_files else 0
        )



    return uploaded_file, selected_file


