
import pandas as pd
from config import *
import xlrd

# def get_latest_excel_file(folder):
#     """
#     Focntion pour récupérer le fichier Excel le plus récent d'un répertoire donné.
#     """
#     files = [f for f in os.listdir(folder) if f.endswith((".xlsx", ".xls"))]
#     if not files:
#         return None
#     files.sort(key=lambda x: os.path.getctime(os.path.join(folder, x)), reverse=True)
#     return os.path.join(folder, files[0])


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


def load_data(uploaded_file):
    """
    Fonction pour sélectionner un fichier Excel, uploader et lire les onglets,
    en gérant les erreurs liées à l'absence d'onglets ou autres erreurs lors de la lecture.
    """
    import pandas as pd
    import streamlit as st
    import tempfile
    import xlrd

    df_onglet_1 = df_onglet_2 = df_onglet_3 = source = None
    temp_file_path = None

    # ----- Étape 1 : sauvegarde temporaire du fichier uploadé -----
    try:
        if uploaded_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.getbuffer())
                temp_file_path = tmp.name
        else:
            return None, None, None, None
    except Exception:
        st.error("❌ Erreur lors de la copie du fichier. Vérifiez qu’il n’est pas corrompu ou verrouillé.")
        return None, None, None, None

    # ----- Étape 2 : ouverture du fichier Excel -----
    try:
        xls = pd.ExcelFile(temp_file_path, engine="calamine")
        available_sheets = xls.sheet_names
    except Exception:
        st.error(
            "❌ Erreur lors de l’ouverture du fichier sélectionné. Veuillez charger un fichier au bon format et à jour.\n"
            "Causes possibles : "
            "\n- Extension du fichier non prise en charge (uniquement XLS et XLSX)"
            "\n- Fichier vide ou corrompu"
            "\n- Fichier verrouillé"
        )
        return None, None, None, None

    # ----- Étape 3 : récupération des métadonnées -----
    try:
        label, date_value = get_date_generation_restitution(temp_file_path)
        source = f"🔼 Fichier Uploadé - {label} : {date_value}" if uploaded_file else ""
    except Exception:
        source = "ℹ️ Fichier chargé (impossible de récupérer la date)"

    # ----- Étape 4 : lecture onglet par onglet -----
    onglets_attendus = {
        "onglet 1": onglet_1,
        "onglet 2": onglet_2,
        "onglet 3": onglet_3
    }

    def clean_dataframe(df):
        """Nettoie le DataFrame pour éviter les erreurs de type (NaN texte, tirets, etc.)."""
        if df is None or df.empty:
            return df
        # Remplacer les valeurs invalides
        df = df.replace(["NaN", "nan", "-", ""], pd.NA)
        
        for col in df.columns:
            # Conversion numérique si possible
            df[col] = pd.to_numeric(df[col], errors="ignore")
        
        # Remplace pandas.NA par None pour compatibilité openpyxl
        df = df.where(pd.notna(df), None)
        
        return df
    
    for nom, sheet in onglets_attendus.items():
        df = pd.DataFrame()  # par défaut vide
        matched_sheets = [s for s in available_sheets if s.strip() == sheet.strip()]
        if matched_sheets:
            try:
                # Essai calamine
                df = pd.read_excel(temp_file_path, sheet_name=matched_sheets[0], engine="calamine")
            except Exception:
                try:
                    # fallback xlrd
                    workbook = xlrd.open_workbook_xls(temp_file_path, ignore_workbook_corruption=True)
                    df = pd.read_excel(workbook, sheet_name=matched_sheets[0])
                except Exception as e2:
                    st.warning(f"⚠️ Impossible de lire {nom} : {e2}")
                    df = pd.DataFrame()
        else:
            st.warning(f"⚠️ {nom.capitalize()} est absent du fichier source.")

        # Nettoyage systématique
        df = clean_dataframe(df)

        # Assignation
        if nom == "onglet 1":
            try:
                df_onglet_1 = load_excel_data_dynamic_start(temp_file_path)
                df_onglet_1 = clean_dataframe(df_onglet_1)
            except Exception:
                df_onglet_1 = df
        elif nom == "onglet 2":
            df_onglet_2 = df
        elif nom == "onglet 3":
            df_onglet_3 = df

    return df_onglet_1, df_onglet_2, df_onglet_3, source




def get_date_generation_restitution(file_path):
    """
    Focntion pour récupérer la date de restition dans l'onglet 1 des fichiers
    de restitution
    """
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

