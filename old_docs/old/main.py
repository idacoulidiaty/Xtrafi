import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage

# ---------- CONFIG APP---------- #
WATCHED_FOLDER = "data"
st.set_page_config(page_title="Xtrafi Data Viz", layout="wide")


# ---------- LOGO ---------- #
st.image("static/logo-xtrafi.png", width=200)
st.title("📊 Votre Dashboard")


# ---------- FONCTIONS TRANSFORMATION DONNEES---------- #
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

def get_date_generation_restitution(file_path):
    preview_df = pd.read_excel(file_path, header=None, engine="calamine")
    mask = preview_df.astype(str).apply(lambda row: row.str.contains("Date génération", case=False, na=False)).any(axis=1)
    if mask.any():
        row_index = mask[mask].index[0]
        col_index = preview_df.iloc[row_index].astype(str).str.contains("Date génération", case=False, na=False)
        col_idx = col_index[col_index].index[0]
        label = preview_df.iloc[row_index, col_idx]
        date_value = preview_df.iloc[row_index + 1, col_idx]
        return label, date_value
    else:
        raise ValueError("❌ 'Date génération restitution' introuvable.")

def get_latest_excel_file(folder):
    files = [f for f in os.listdir(folder) if f.endswith((".xlsx", ".xls"))]
    if not files:
        return None
    files.sort(key=lambda x: os.path.getctime(os.path.join(folder, x)), reverse=True)
    return os.path.join(folder, files[0])


# ---------- SIDEBAR : Upload ou sélection ---------- #
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


# ---------- CHARGEMENT DES DONNÉES ---------- #
df_onglet_1 = None
df_onglet_2 = None
df_onglet_3 = None
source = None
if uploaded_file:
    try:
        label, date_value = get_date_generation_restitution(uploaded_file)
        df_onglet_1 = load_excel_data_dynamic_start(uploaded_file)
        df_onglet_2 = pd.read_excel(uploaded_file, sheet_name="Restit Brute (Total Répdts)", engine="calamine")
        df_onglet_3 = pd.read_excel(uploaded_file, sheet_name="Restit Rapport (Total Répdts)", engine="calamine")
        source = f"🔼 Fichier Uploadé - {label} : {date_value}"
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier uploadé : {e}")

elif selected_file:
    file_path = os.path.join(WATCHED_FOLDER, selected_file)
    try:
        label, date_value = get_date_generation_restitution(file_path)
        df_onglet_1 = load_excel_data_dynamic_start(file_path)
        df_onglet_2 = pd.read_excel(file_path, sheet_name="Restit Brute (Total Répdts)", engine="calamine")
        df_onglet_3 = pd.read_excel(file_path, sheet_name="Restit Rapport (Total Répdts)", engine="calamine")
        df_onglet_3 = df_onglet_3[['Code\nRAPPORT Ind.\nVIRTUEL',
       'Nom\nRAPPORT Ind.\nVIRTUEL', 'Code\nREPORTING\nInd. VIRTUEL',
       'Nom\nREPORTING Ind.\nVIRTUEL', 'Code\nAPP Indicateur\nVIRTUEL',
       'Nom\nAPP Indicateur\nVIRTUEL',
       "Unité de conversion\n(de l'indicateur virtuel)", 
       'Total\nMontant\nCollecte\nRéelle\nExercice N-1',
       'Total\nValo. Financière\nCollecte Réelle\nExercice N-1',
       'Total\nMontant\nCollecte\nRéelle\nExercice N',
       'Total\nValo. Financière\nCollecte Réelle\nExercice N',
       'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N',
       'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N',
       'Total\nValo. Financière\nCollecte\nO.Strat\nExercice N',
       'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N',
       'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N',
       'Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N']]
        source = f"👀 Fichier sélectionné - {label} : {date_value}"
        
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier sélectionné : {e}")


# ---------- AFFICHAGE DES DONNÉES ---------- #
if df_onglet_1 is not None and df_onglet_2 is not None and df_onglet_3 is not None:
    st.success(source)

    onglet_selectionne = st.radio("🧭 Choisissez l'onglet à afficher :", [
        "📈 Paramètres restitution (onglet 1)",
        "📊 Données brutes (onglet 2)",
        "📋 Rapport consolidé (onglet 3)"
    ])

    df_filtré = None
    if onglet_selectionne == "📈 Paramètres restitution (onglet 1)":
        st.subheader("📈 Aperçu : Paramètres restitution")
        st.dataframe(df_onglet_1, use_container_width=True)

    elif onglet_selectionne == "📊 Données brutes (onglet 2)":
        st.subheader("📊 Aperçu : Données Brutes")
        st.dataframe(df_onglet_2, use_container_width=True)

    elif onglet_selectionne == "📋 Rapport consolidé (onglet 3)":
        # st.subheader("📋 Rapport Consolidé")


        # 🧩 Colonnes sélectionnables
        colonnes_masquables = ['Code\nRAPPORT Ind.\nVIRTUEL',
       'Nom\nRAPPORT Ind.\nVIRTUEL', 'Code\nREPORTING\nInd. VIRTUEL',
       'Nom\nREPORTING Ind.\nVIRTUEL', 'Code\nAPP Indicateur\nVIRTUEL',
       'Nom\nAPP Indicateur\nVIRTUEL',
       "Unité de conversion\n(de l'indicateur virtuel)", 
       'Total\nMontant\nCollecte\nRéelle\nExercice N-1',
       'Total\nValo. Financière\nCollecte Réelle\nExercice N-1',
       'Total\nMontant\nCollecte\nRéelle\nExercice N',
       'Total\nValo. Financière\nCollecte Réelle\nExercice N',
       'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N',
       'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N',
       'Total\nValo. Financière\nCollecte\nO.Strat\nExercice N',
       'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N',
       'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N',
       'Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N']

        # Filtrage
        # colonnes_visibles = st.multiselect("🕶️ Colonnes à afficher :", options=colonnes_masquables, default=colonnes_masquables)

        # Sélection et affichage
        # df_filtré = df_onglet_3[colonnes_visibles] if colonnes_visibles else pd.DataFrame()

        df_filtré = df_onglet_3[colonnes_masquables]
        st.subheader("🔍 Aperçu des données")
        st.dataframe(df_filtré, use_container_width=True)
        
        

    # ---------- AFFICHAGE FINAL ---------- #
    # Onglets
    tab1, tab2 , tab3, tab4= st.tabs(["📋 TdB REEL N vs REEL N-1", "📋 TdB REEL N vs Obj. Op. N","📋 TdB REEL N vs Obj. Strat. N","📊 Visualisations"])
    
    
    if df_filtré is not None:
        # ---------- Onglet 1 : Tableau formaté ---------- #
        with tab1:
            # st.subheader("📋 Données formatées et mises en forme")
            # st.success(f"✅ Données chargées ({source})")



            # # ---------- KPIs ---------- #
            # # Filtrage des colonnes numériques visibles
            # colonnes_numeriques_visibles = df_filtré.select_dtypes(include='number').columns.tolist()
            # st.subheader("📌 Indicateurs clés")
            # numeric_cols = df_filtré[colonnes_numeriques_visibles]
            # if not numeric_cols.empty:
            #     def format_euro(val):
            #         return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ") + "€"

            #     col1, col2, col3 = st.columns(3)
            #     col1.metric("💶 Total Valo Financière Collecte Réelle N", format_euro(numeric_cols['Total\nValo. Financière\nCollecte Réelle\nExercice N'].sum().sum()))
            #     col2.metric("💶 Total Valo Financière Objectifs Stratégiques", format_euro(numeric_cols['Total\nValo. Financière\nCollecte\nO.Strat\nExercice N'].sum().sum()))
            #     col3.metric("💶 Total Valo Financière Objectifs Opérationnels", format_euro(numeric_cols['Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N'].sum().sum()))
            # else:
            #     st.warning("Aucune colonne numérique trouvée pour les indicateurs.")


            # ---------- RENOMMER LES COLONNES ---------- #
            df_style = df_filtré.rename(columns={
    'Code\nREPORTING\nInd. VIRTUEL': 'Code Reporting',
    'Nom\nREPORTING Ind.\nVIRTUEL': 'Nom Reporting',
    'Nom\nAPP Indicateur\nVIRTUEL': 'Nom indicateur Virtuel',
    "Unité de conversion\n(de l'indicateur virtuel)": "Unité de conversion de l'indicateur",
    'Total\nMontant\nCollecte\nRéelle\nExercice N-1': 'Reel N-1',
    'Total\nValo. Financière\nCollecte Réelle\nExercice N-1': 'Valorisation Financière REEL N-1',
    'Total\nMontant\nCollecte\nRéelle\nExercice N': 'Reel N',
    'Total\nValo. Financière\nCollecte Réelle\nExercice N': 'Valorisation Financière REEL N',
    'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N': 'Objectifs Stratégiques PLAFOND période N',
    'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N': 'Objectifs Stratégiques SEUIL période N',
    'Total\nValo. Financière\nCollecte\nO.Strat\nExercice N': 'Valorisation Financière Objectifs Stratégiques N',
    'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N': 'Objectifs Opérationnels PLAFOND période N',
    'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N': 'Objectifs Opérationnels SEUIL période N',
    'Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N': 'Valorisation Financière Objectifs Opérationnels N'
})

                       # ---------- CALCUL DES VARIATIONS ---------- #
            df_style["VARIATION Réel N vs Réel N-1 (%)"] = (
                (df_style["Reel N"] - df_style["Reel N-1"]) / df_style["Reel N-1"]
            ) * 100

            # Variation Réel N vs Objectifs Stratégiques PLAFOND
            df_style["VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)"] = (
                (df_style["Reel N"] - df_style["Objectifs Stratégiques PLAFOND période N"]) / df_style["Objectifs Stratégiques PLAFOND période N"]
            ) * 100

            # Variation Réel N vs Objectifs Stratégiques PLANCHER
            df_style["VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)"] = (
                (df_style["Reel N"] - df_style["Objectifs Stratégiques SEUIL période N"]) / df_style["Objectifs Stratégiques SEUIL période N"]
            ) * 100

            # Variation Réel N vs Objectifs Opérationnels PLAFOND
            df_style["VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)"] = (
                (df_style["Reel N"] - df_style["Objectifs Opérationnels PLAFOND période N"]) / df_style["Objectifs Opérationnels PLAFOND période N"]
            ) * 100

            # Variation Réel N vs Objectifs Opérationnels PLANCHER
            df_style["VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)"] = (
                (df_style["Reel N"] - df_style["Objectifs Opérationnels SEUIL période N"]) / df_style["Objectifs Opérationnels SEUIL période N"]
            ) * 100

            # ---------- AFFICHAGE DU TABLEAU STYLÉ ---------- #
            st.subheader("📋 Tableau de bord REEL N vs REEL N-1")

            cols_to_display = [
                'Code\nRAPPORT Ind.\nVIRTUEL',
                'Nom\nRAPPORT Ind.\nVIRTUEL',
                'Code Reporting',
                'Nom Reporting',
                "Nom indicateur Virtuel",
                "Reel N-1",
                "Reel N",
                "VARIATION Réel N vs Réel N-1 (%)",
                "Unité de conversion de l'indicateur",
                'Valorisation Financière REEL N-1',
                'Valorisation Financière REEL N'
            ]

            df1 = df_style[cols_to_display]

            # Formatage personnalisé
            styler = (
                df1.style
                # Valorisation financière : pas de décimale
                .format({col: "{:,.0f}" for col in df1.columns if "Valorisation Financière" in col})
                # Pourcentages : 1 décimale, sinon 2 décimales (pour "Reel N", "Reel N-1" et colonnes variation)
                .format({
                    col: "{:.1f}%" if "%" in col else "{:.2f}"
                    for col in df1.columns
                    if col.startswith("VARIATION") or col in ["Reel N", "Reel N-1"]
                })
                # Alignement à droite de toutes les cellules
                .set_properties(**{'text-align': 'right'})
                # Styles couleurs comme avant
                .applymap(lambda v: 'background-color: palegreen', subset=["Reel N"])
                .applymap(lambda v: 'background-color: gainsboro', subset=["Reel N-1"])
                .applymap(lambda v: 'background-color: lemonchiffon', subset=["VARIATION Réel N vs Réel N-1 (%)"])
                .applymap(lambda v: 'background-color: paleturquoise', subset=["Valorisation Financière REEL N-1"])
                .applymap(lambda v: 'background-color: tan', subset=["Valorisation Financière REEL N"])
                # Style de la table et des cellules
                .set_table_styles([
                    {
                        'selector': 'th',
                        'props': [
                            ('background-color', '#f2f2f2'),
                            ('white-space', 'pre-wrap'),
                            ('word-wrap', 'break-word'),
                            ('max-width', '150px'),
                            ('font-size', '12px'),
                            ('text-align', 'center')
                        ]
                    },
                    {
                        'selector': 'td',
                        'props': [
                            ('white-space', 'pre-wrap'),
                            ('word-wrap', 'break-word'),
                            ('max-width', '150px'),
                            ('font-size', '12px'),
                            ('text-align', 'right')  # renforcement alignement chiffres à droite
                        ]
                    }
                ])
            )

            st.markdown(styler.to_html(), unsafe_allow_html=True)

            
        with tab2:
            #-------------------- TDB REEL vs OBJ OP --------------------
            st.subheader("\n \n")
            st.subheader("Tableau de Bord REEL N vs Objectifs Opérationnels N")

            st.subheader("\n \n")

            cols_to_display = [
                'Code\nRAPPORT Ind.\nVIRTUEL',
                'Nom\nRAPPORT Ind.\nVIRTUEL',
                'Code Reporting',
                'Nom Reporting',
                "Nom indicateur Virtuel",
                "Reel N",
                'Objectifs Opérationnels SEUIL période N',
                "VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)",
                'Objectifs Opérationnels PLAFOND période N',
                "VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)",
                "Unité de conversion de l'indicateur",
                'Valorisation Financière REEL N',
                'Valorisation Financière Objectifs Opérationnels N'
            ]

            df2 = df_style[cols_to_display]

            styler = (
    df2.style
    .applymap(lambda v: 'background-color: palegreen', subset=["Reel N"])
    .applymap(lambda v: 'background-color: gainsboro', subset=[
        'Objectifs Opérationnels SEUIL période N',
        'Objectifs Opérationnels PLAFOND période N'])
    .applymap(lambda v: 'background-color: lemonchiffon', subset=[
        "VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)",
        "VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)"])
    .applymap(lambda v: 'background-color: paleturquoise', subset=['Valorisation Financière Objectifs Opérationnels N'])
    .applymap(lambda v: 'background-color: tan', subset=["Valorisation Financière REEL N"])
    .set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#f2f2f2'), ('white-space', 'pre-wrap'), ('font-size', '12px')]},
        {'selector': 'td', 'props': [('white-space', 'pre-wrap'), ('text-align', 'right'), ('font-size', '12px')]}
    ])
    .format({
        col: "{:.0f}" for col in df2.columns if "Valorisation Financière" in col
    } | {
        col: "{:.1f}%" if "%" in col else "{:.2f}" for col in df2.columns
        if col not in [
            "Valorisation Financière REEL N",
            "Valorisation Financière Objectifs Opérationnels N"
        ] and df2[col].dtype.kind in "fi"
    })
)

            st.markdown(styler.to_html(), unsafe_allow_html=True)


        with tab3:
            #-------------------- TDB REEL vs OBJ STRAT --------------------
            st.subheader("\n  \n")
            st.subheader("Tableau de Bord REEL N vs Objectifs Stratégiques N")
            cols_to_display = [
                'Code\nRAPPORT Ind.\nVIRTUEL',
                'Nom\nRAPPORT Ind.\nVIRTUEL',
                'Code Reporting',
                'Nom Reporting',
                "Nom indicateur Virtuel",
                "Reel N",
                'Objectifs Stratégiques SEUIL période N',
                "VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)",
                'Objectifs Stratégiques PLAFOND période N',
                "VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)",
                "Unité de conversion de l'indicateur",
                'Valorisation Financière REEL N',
                'Valorisation Financière Objectifs Stratégiques N'
            ]

            df3 = df_style[cols_to_display]

            styler = (
    df3.style
    .applymap(lambda v: 'background-color: palegreen', subset=["Reel N"])
    .applymap(lambda v: 'background-color: gainsboro', subset=[
        'Objectifs Stratégiques SEUIL période N',
        'Objectifs Stratégiques PLAFOND période N'])
    .applymap(lambda v: 'background-color: lemonchiffon', subset=[
        "VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)",
        "VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)"])
    .applymap(lambda v: 'background-color: paleturquoise', subset=['Valorisation Financière Objectifs Stratégiques N'])
    .applymap(lambda v: 'background-color: tan', subset=["Valorisation Financière REEL N"])
    .set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#f2f2f2'), ('white-space', 'pre-wrap'), ('font-size', '12px')]},
        {'selector': 'td', 'props': [('white-space', 'pre-wrap'), ('text-align', 'right'), ('font-size', '12px')]}
    ])
    .format({
        col: "{:.0f}" for col in df3.columns if "Valorisation Financière" in col
    } | {
        col: "{:.1f}%" if "%" in col else "{:.2f}" for col in df3.columns
        if col not in [
            "Valorisation Financière REEL N",
            "Valorisation Financière Objectifs Stratégiques N"
        ] and df3[col].dtype.kind in "fi"
    })
)


            st.markdown(styler.to_html(), unsafe_allow_html=True)

        # ---------- GRAPHIQUES ---------- #
        with tab4: 
                selected_fig = None  # Pour l'export plus bas

                st.subheader("📈 Visualisation")
                st.subheader("\n \n")


                # Graphique 1 - Eau consommée (Barre horizontale avec valeur réelle)
                # eau = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("eau consommée", case=False, na=False)]
                # objectif_eau = 1500000 
                # if not eau.empty and 'Total\nMontant\nCollecte\nRéelle\nExercice N' in eau.columns:
                #     eau_val = float(eau['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0])
                # else:
                #     st.warning("⚠️ Donnée manquante pour le montant de collecte réelle (exercice N)")
                #     eau_val = 0 
                    
                # # st.subheader("💧 Total eau consommée")
                # # Créer une barre horizontale avec Plotly
                # fig1 = go.Figure()
                # fig1.add_trace(go.Bar(
                #     x=[eau_val],
                #     y=["Eau consommée"],
                #     orientation='h',
                #     marker_color='dodgerblue',
                #     name="Réel"
                # ))
                # fig1.add_trace(go.Bar(
                #     x=[max(objectif_eau - eau_val, 0)],
                #     y=["Eau consommée"],
                #     orientation='h',
                #     marker_color='lightgray',
                #     name="Reste à consommer",
                #     hoverinfo='none'
                # ))
                # fig1.update_layout(
                #     title="Avancement de la consommation d’eau",
                #     xaxis=dict(range=[0, objectif_eau], title="LE"),
                #     barmode='stack',
                #     height=120,
                #     margin=dict(l=100, r=20, t=40, b=20),
                #     showlegend=False
                # )
                # st.metric(label= "Eau consommée", value=f"{eau_val:,.0f}".replace(",", " ") + " Litres d'Eau")

                col1,col2 = st.columns(2)
                with col1:

                    with st.container():
                        st.subheader("💧 Total eau consommée N vs N-1")
                    
                        eau = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("eau consommée", case=False, na=False)]
                    
                        if not eau.empty and 'Total\nMontant\nCollecte\nRéelle\nExercice N' in eau.columns:
                            valeur_n1 = eau['Total\nMontant\nCollecte\nRéelle\nExercice N-1']
                            valeur_n = eau['Total\nMontant\nCollecte\nRéelle\nExercice N']
                            plafond_strat = eau['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N']
                            plancher_strat = eau['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N']
                            plafond_oper = eau['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N']
                            plancher_oper = eau['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N']
                    
                            fig1 = go.Figure()
                            fig1.add_trace(go.Bar(
                                name='Exercice N-1',
                                x=eau['Nom Ind.N2\nAPP'],
                                y=valeur_n1,
                                marker_color='lightblue',
                                width=0.1  # ✅ largeur réduite
                            ))
                            fig1.add_trace(go.Bar(
                                name='Exercice N',
                                x=eau['Nom Ind.N2\nAPP'],
                                y=valeur_n,
                                marker_color='blue',
                                width=0.1  # ✅ largeur réduite
                            ))
                    
                            def add_horizontal_line(fig, y_value, label, color, dash):
                                if pd.notna(y_value):
                                    fig.add_shape(
                                        type="line",
                                        x0=-0.5, x1=0.5,
                                        y0=y_value, y1=y_value,
                                        line=dict(color=color, width=2, dash=dash),
                                        xref='x', yref='y'
                                    )
                                    fig.add_annotation(
                                        x=0.5,
                                        y=y_value,
                                        text=label,
                                        showarrow=False,
                                        yanchor="bottom",
                                        font=dict(color=color)
                                    )
                    
                            # ✅ MULTISELECT placé à l’intérieur du même container que le graphique
                            options_obj = st.multiselect(
                                "🎯 Objectifs à afficher sur le graphique",
                                options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
                                default=[],
                                key="multiselect_objectifs_eau"
                            )
                    
                            # Ajout des lignes sélectionnées
                            if "Stratégique - Plafond" in options_obj:
                                add_horizontal_line(fig1, np.sum(plafond_strat), "Obj. Strat. plafond moyen", "green", "dash")
                            if "Stratégique - Plancher" in options_obj:
                                add_horizontal_line(fig1, np.sum(plancher_strat), "Obj. Strat. plancher moyen", "black", "dash")
                            if "Opérationnel - Plafond" in options_obj:
                                add_horizontal_line(fig1, np.sum(plafond_oper), "Obj. Opé. plafond moyen", "orange", "dash")
                            if "Opérationnel - Plancher" in options_obj:
                                add_horizontal_line(fig1, np.sum(plancher_oper), "Obj. Opé. plancher moyen", "yellow", "dash")
                    
                            # ✅ TOUJOURS DANS LE MÊME CONTAINER
                            st.markdown("ℹ️ Sélectionnez un ou plusieurs types d’objectifs pour les afficher.")
                            st.plotly_chart(fig1, use_container_width=True)


            
                with col2:
                    st.subheader("% Eau stockée N-1 vs N")

                    stock = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("eau stockée", case=False, na=False)]

                    if not stock.empty:
                        st.subheader("📊 % d’eau stockée – Exercice N vs N-1")

                        # Récupération des valeurs
                        value_n1 = float(stock['Total\nMontant\nCollecte\nRéelle\nExercice N-1'].values[0])
                        value_n = float(stock['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0])

                        plafond_strat = stock['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'].values[0]
                        plancher_strat = stock['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'].values[0]
                        plafond_oper = stock['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'].values[0]
                        plancher_oper = stock['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'].values[0]

                        # Sélecteur multichoix
                        options_jauge = st.multiselect(
                            "🎯 Objectifs à afficher sur la jauge",
                            options=[
                                "Stratégique - Plafond",
                                "Stratégique - Plancher",
                                "Opérationnel - Plafond",
                                "Opérationnel - Plancher"
                            ],
                            default=[],
                            key="multiselect_objectifs_jauge"
                        )

                        # Création du graphique
                        fig2 = go.Figure()

                        fig2.add_trace(go.Bar(
                            y=["Eau stockée"],
                            x=[value_n1],
                            name='Exercice N-1',
                            orientation='h',
                            marker_color='lightblue',
                            text=f"{value_n1}%",
                            textposition='outside',
                            width=0.3
                        ))

                        fig2.add_trace(go.Bar(
                            y=["Eau stockée"],
                            x=[value_n],
                            name='Exercice N',
                            orientation='h',
                            marker_color='blue',
                            text=f"{value_n}%",
                            textposition='outside',
                            width=0.3
                        ))

                        # Ajout des seuils verticaux (adapté comme dans fig4)
                        def add_vertical_line(fig, x_val, label, color, dash):
                            if pd.notna(x_val):
                                fig.add_shape(
                                    type="line",
                                    x0=x_val, x1=x_val,
                                    y0=-0.5, y1=0.5,
                                    line=dict(color=color, width=2, dash=dash),
                                    xref='x', yref='y'
                                )
                                fig.add_annotation(
                                    x=x_val,
                                    y=0.5,
                                    text=label,
                                    showarrow=False,
                                    yanchor="bottom",
                                    textangle=-90,
                                    font=dict(color=color),
                                    xanchor="left"
                                )

                        if "Stratégique - Plafond" in options_jauge:
                            add_vertical_line(fig2, plafond_strat, "Obj. Strat. plafond", "green", "dash")
                        if "Stratégique - Plancher" in options_jauge:
                            add_vertical_line(fig2, plancher_strat, "Obj. Strat. plancher", "green", "dot")
                        if "Opérationnel - Plafond" in options_jauge:
                            add_vertical_line(fig2, plafond_oper, "Obj. Opé. plafond", "orange", "dash")
                        if "Opérationnel - Plancher" in options_jauge:
                            add_vertical_line(fig2, plancher_oper, "Obj. Opé. plancher", "orange", "dot")

                        # Mise en forme finale
                        fig2.update_layout(
                            xaxis=dict(title="Pourcentage", range=[0, 100]),
                            yaxis=dict(title=""),
                            barmode='group',
                            template='plotly_white',
                            height=350
                        )

                        st.plotly_chart(fig2, use_container_width=True)

                    else:
                        st.warning("⚠️ Donnée manquante pour l’eau stockée.")



                col3,col4 = st.columns(2)
                with col3:
                    # # Graphique 3 - Fréquence réunion RSE
                    # rse = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("réunion RSE", case=False, na=False)]
                    # st.subheader("📅 Fréquence Réunions RSE")
                    # fig3, ax2 = plt.subplots()
                    # if not rse.empty and 'Total\nMontant\nCollecte\nRéelle\nExercice N' in rse.columns:
                    #     val_str = str(rse['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0]).replace('%', '')
                    #     try:
                    #         val = float(val_str)
                    #     except ValueError:
                    #         st.warning("⚠️ Format de valeur incorrect pour le RSE (exercice N)")
                    #         val = 0
                    # else:
                    #     st.warning("⚠️ Donnée manquante pour le RSE (exercice N)")
                    #     val = 0
                    # ax2.pie([val, 100-val], labels=[f"Réunions RSE ({val}%)", f"Autres Réunions ({100-val}%)"], autopct="%1.1f%%", startangle=90)
                    # ax2.set_title("Fréquence Réunions RSE")
                    # st.plotly_chart(fig3, use_container_width=True)

                # #Graphique 4 : Évolution du nombre total de femmes dans l’effectif VS Objectifs opérationnels plafond et plancher
                # #🔍 Filtrage de l’indicateur "femmes"
                # idicateur_femmes = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("femmes", case=False, na=False)]
                # i not indicateur_femmes.empty:
                #    valeur_n1 = indicateur_femmes['Total\nMontant\nCollecte\nRéelle\nExercice N-1'].values[0]
                #    valeur_n = indicateur_femmes['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0]
                #    seuil = indicateur_femmes['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'].values[0]
                #    plafond = indicateur_femmes['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'].values[0]
                #    # objectif = indicateur_femmes['Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N'].values[0]
                #    # 📊 Création du graphique
                #    fig_femmes = go.Figure()
                #    fig_femmes.add_trace(go.Bar(
                #        x=['Exercice N-1', 'Exercice N'],
                #        y=[valeur_n1, valeur_n],
                #        name='Nombre de femmes',
                #        marker_color='blue'
                #    ))
                #    fig_femmes.add_trace(go.Scatter(
                #        x=['Exercice N-1', 'Exercice N'], y=[seuil, seuil],
                #        mode='lines', name='Seuil (Plancher)', line=dict(color='orange', dash='dash')
                #    ))
                #    fig_femmes.add_trace(go.Scatter(
                #        x=['Exercice N-1', 'Exercice N'], y=[plafond, plafond],
                #        mode='lines', name='Plafond', line=dict(color='red', dash='dash')
                #    ))
                #    # fig_femmes.add_trace(go.Scatter(
                #    #     x=['Exercice N-1', 'Exercice N'], y=[objectif, objectif],
                #    #     mode='lines', name='Objectif Opérationnel', line=dict(color='green', dash='dash')
                #    # ))
                #    # Mise en forme
                #    fig_femmes.update_layout(
                #        title="Évolution du nombre total de femmes dans l’effectif VS Objectifs opérationnels plafond et plancher",
                #        xaxis_title="Exercice",
                #        yaxis_title="Nombre de femmes",
                #        barmode='group',
                #        template='plotly_white'
                #    )
                #    # ➕ Affichage dans Streamlit
                #    st.plotly_chart(fig_femmes)


                # with col3:
                    with st.container():
                        # Graphique 4 : Consommation d'eau des sièges, agences, bureaux VS Consommation totale d'eau
                        # 🔍 Sélection des indicateurs eau
                        indicateurs_eau = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("consommation d'eau", case=False, na=False)
                        ]
                        if not indicateurs_eau.empty:
                            st.subheader("📊 Consommation d'eau des sièges, agences, bureaux VS Consommation totale d'eau")
                            noms_indicateurs = []
                            valeurs_n1 = []
                            valeurs_n = []
                            plafond_strat_list = []
                            plancher_strat_list = []
                            plafond_oper_list = []
                            plancher_oper_list = []
                            for _, row in indicateurs_eau.iterrows():
                                nom = row['Nom Ind.N2\nAPP']
                                noms_indicateurs.append(nom)
                                valeur_n1 = row.get('Total\nMontant\nCollecte\nRéelle\nExercice N-1', 0)
                                valeur_n = row.get('Total\nMontant\nCollecte\nRéelle\nExercice N', 0)
                                valeurs_n1.append(valeur_n1)
                                valeurs_n.append(valeur_n)
                                plafond_strat_list.append(row.get('Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'))
                                plancher_strat_list.append(row.get('Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'))
                                plafond_oper_list.append(row.get('Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'))
                                plancher_oper_list.append(row.get('Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'))
                            fig4 = go.Figure()
                            # Barres
                            fig4.add_trace(go.Bar(name='Exercice N-1', x=noms_indicateurs, y=valeurs_n1, marker_color='lightblue', width = 0.2))
                            fig4.add_trace(go.Bar(name='Exercice N', x=noms_indicateurs, y=valeurs_n, marker_color='blue', width = 0.2))
                            # Lignes horizontales pour les plafonds et planchers (on filtre les valeurs non nulles)
                            def add_horizontal_line(fig, y_value, label, color, dash):
                                if pd.notna(y_value):
                                    fig4.add_shape(
                                        type="line",
                                        x0=-0.5, x1=len(noms_indicateurs) - 0.5,
                                        y0=y_value, y1=y_value,
                                        line=dict(color=color, width=2, dash=dash),
                                        xref='x', yref='y'
                                    )
                                    fig4.add_annotation(
                                        x=len(noms_indicateurs) - 1,
                                        y=y_value,
                                        text=label,
                                        showarrow=False,
                                        yanchor="bottom",
                                        font=dict(color=color)
                                    )
                            # Liste déroulante pour choisir les objectifs à afficher
                            options_obj = st.multiselect(
                                "🎯 Objectifs à afficher",
                                options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
                                default=[],  # Possibilité demettre une valeur par défaut ici
                                key="multiselect_objectifs"
                            )
                            # Ajout des lignes si sélectionnées
                            if "Stratégique - Plafond" in options_obj:
                                add_horizontal_line(fig4, np.nanmean(plafond_strat_list), "Obj. Strat. plafond moyen", "green", "dash")
                            if "Stratégique - Plancher" in options_obj:
                                add_horizontal_line(fig4, np.nanmean(plancher_strat_list), "Obj. Strat. plancher moyen", "green", "dot")
                            if "Opérationnel - Plafond" in options_obj:
                                add_horizontal_line(fig4, np.nanmean(plafond_oper_list), "Obj. Opé. plafond moyen", "orange", "dash")
                            if "Opérationnel - Plancher" in options_obj:
                                add_horizontal_line(fig4, np.nanmean(plancher_oper_list), "Obj. Opé. plancher moyen", "orange", "dot")
                            st.markdown("ℹ️ Vous pouvez sélectionner un ou plusieurs types d’objectifs à afficher sur le graphique.")
                            # Layout
                            fig4.update_layout(
                                title="Consommation d'eau des sièges, agences, bureaux VS Consommation totale d'eau",
                                barmode='group',
                                xaxis_title="Indicateur",
                                yaxis_title="Volume (Litres)",
                                template='plotly_white',
                                height=500
                            )
                            st.plotly_chart(fig4, use_container_width=True)



                
                with col4:
                                    #Graphique 5 : Consommations carburant 
                    #  Sélection des indicateurs carburant
                    indicateurs_carburant = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].isin([
                        "Consommation totale de carburant véhicule",
                        "Consommation diesel des véhicules",
                        "Consommation Essence/Super des véhicules"
                    ])]

                    # Préparer les listes de seuils
                    plafond_strat_list = []
                    plancher_strat_list = []
                    plafond_oper_list = []
                    plancher_oper_list = []

                    if not indicateurs_carburant.empty:
                        st.subheader("📊 Consommations carburant véhicules")

                        fig5 = go.Figure()

                        for i, row in indicateurs_carburant.iterrows():
                            nom = row['Nom Ind.N2\nAPP']

                            valeur_n1 = row.get('Total\nMontant\nCollecte\nRéelle\nExercice N-1', 0)
                            valeur_n = row.get('Total\nMontant\nCollecte\nRéelle\nExercice N', 0)

                            fig5.add_trace(go.Bar(
                                name=f"{nom} - N-1", x=[nom], y=[valeur_n1],
                                marker_color='lightblue'
                            ))
                            fig5.add_trace(go.Bar(
                                name=f"{nom} - N", x=[nom], y=[valeur_n],
                                marker_color='blue'
                            ))

                            # Ajout des valeurs dans les listes DANS la boucle
                            plafond_strat_list.append(row.get('Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'))
                            plancher_strat_list.append(row.get('Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'))
                            plafond_oper_list.append(row.get('Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'))
                            plancher_oper_list.append(row.get('Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'))

                        # Liste déroulante pour choisir les objectifs à afficher
                        options_obj = st.multiselect(
                            "🎯 Objectifs à afficher",
                            options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
                            default=[],
                            key="multiselect_objectifs_2"
                        )

                        # Ajout des lignes si sélectionnées
                        if "Stratégique - Plafond" in options_obj:
                            add_horizontal_line(fig5, np.nanmean(plafond_strat_list), "Obj. Strat. plafond moyen", "green", "dash")
                        if "Stratégique - Plancher" in options_obj:
                            add_horizontal_line(fig5, np.nanmean(plancher_strat_list), "Obj. Strat. plancher moyen", "green", "dot")
                        if "Opérationnel - Plafond" in options_obj:
                            add_horizontal_line(fig5, np.nanmean(plafond_oper_list), "Obj. Opé. plafond moyen", "orange", "dash")
                        if "Opérationnel - Plancher" in options_obj:
                            add_horizontal_line(fig5, np.nanmean(plancher_oper_list), "Obj. Opé. plancher moyen", "orange", "dot")

                        st.markdown("ℹ️ Vous pouvez sélectionner un ou plusieurs types d’objectifs à afficher sur le graphique.")

                        fig5.update_layout(
                            title="Consommations carburant véhicules",
                            barmode='group',
                            xaxis_title="Indicateur",
                            yaxis_title="Valeur (Litres)",
                            template='plotly_white',
                            height=500
                        )

                        st.plotly_chart(fig5, use_container_width=True)



                col5,col6 = st.columns(2)
                with col5:
                                    #Graphique 6 : Gazs à effet de serre
                    #  Sélection des indicateurs carburant
                    indicateurs_rgaes = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].isin([
                        "Rejets de gaz à effets de serre (RGAES)",
                        "RGAES hors production éléctrique",
                    ])]

                    if not indicateurs_rgaes.empty:
                        st.subheader("📊 Rejets Gaz à Effet de Serre N-1 vs N")

                        fig6 = go.Figure()

                        # Préparer les listes de seuils
                        plafond_strat_list = []
                        plancher_strat_list = []
                        plafond_oper_list = []
                        plancher_oper_list = []

                        for i, row in indicateurs_rgaes.iterrows():
                            nom = row['Nom Ind.N2\nAPP']

                            valeur_n1 = row.get('Total\nMontant\nCollecte\nRéelle\nExercice N-1', 0)
                            valeur_n = row.get('Total\nMontant\nCollecte\nRéelle\nExercice N', 0)

                            fig6.add_trace(go.Bar(
                                name=f"{nom} - N-1", x=[nom], y=[valeur_n1],
                                marker_color='lightblue'
                            ))
                            fig6.add_trace(go.Bar(
                                name=f"{nom} - N", x=[nom], y=[valeur_n],
                                marker_color='blue'
                            ))

                            # Ajout des valeurs dans les listes
                            plafond_strat_list.append(row.get('Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'))
                            plancher_strat_list.append(row.get('Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'))
                            plafond_oper_list.append(row.get('Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'))
                            plancher_oper_list.append(row.get('Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'))

                        # Fonction pour ajouter une ligne horizontale
                        def add_horizontal_line(fig, y_value, label, color, dash):
                            if pd.notna(y_value):
                                fig.add_shape(
                                    type="line",
                                    x0=-0.5, x1=len(indicateurs_rgaes) - 0.5,
                                    y0=y_value, y1=y_value,
                                    line=dict(color=color, width=2, dash=dash),
                                    xref='x', yref='y'
                                )
                                fig.add_annotation(
                                    x=len(indicateurs_rgaes) - 1,
                                    y=y_value,
                                    text=label,
                                    showarrow=False,
                                    yanchor="bottom",
                                    font=dict(color=color)
                                )

                        # Sélection des objectifs à afficher
                        options_obj = st.multiselect(
                            "🎯 Objectifs à afficher",
                            options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
                            default=[],
                            key="multiselect_objectifs_3"
                        )

                        # Ajout des lignes
                        if "Stratégique - Plafond" in options_obj:
                            add_horizontal_line(fig6, np.nanmean(plafond_strat_list), "Obj. Strat. plafond moyen", "green", "dash")
                        if "Stratégique - Plancher" in options_obj:
                            add_horizontal_line(fig6, np.nanmean(plancher_strat_list), "Obj. Strat. plancher moyen", "green", "dot")
                        if "Opérationnel - Plafond" in options_obj:
                            add_horizontal_line(fig6, np.nanmean(plafond_oper_list), "Obj. Opé. plafond moyen", "orange", "dash")
                        if "Opérationnel - Plancher" in options_obj:
                            add_horizontal_line(fig6, np.nanmean(plancher_oper_list), "Obj. Opé. plancher moyen", "orange", "dot")

                        st.markdown("ℹ️ Vous pouvez sélectionner un ou plusieurs types d’objectifs à afficher sur le graphique.")

                        # Layout
                        fig6.update_layout(
                            title="Rejets Gaz à Effet de Serre N-1 vs N",
                            barmode='group',
                            xaxis_title="Indicateur",
                            yaxis_title="Valeur (Litres)",
                            template='plotly_white',
                            height=500
                        )

                        st.plotly_chart(fig6, use_container_width=True)



                # figures = [fig1, fig2, fig3, fig4,fig5]
                # titres = ["💧 Total eau consommée N vs N-1", "💧 Pourcentage d'eau stockée", "📊 Fréquence Réunions RSE", 
                #           "📊Consommation d'eau des sièges, agences, bureaux VS Consommation totale d'eau","Consommations carburant véhicules"
                #           ]

                # for i in range(0, len(figures), 2):
                #     cols = st.columns(2)
                #     for j in range(2):
                #         if i + j < len(figures):
                #             with cols[j]:
                #                 with st.container():
                #                     st.markdown(
                #                         f"""
                #                         <div style="border:1px solid #ddd; padding:15px; border-radius:10px; box-shadow: 2px 2px 8px rgba(0,0,0,0.1); background-color: '#fafafa';">
                #                             <h4 style="text-align:center;">{titres[i + j]}</h4>
                #                         """,
                #                         unsafe_allow_html=True
                #                     )
                #                     # 🛠 Ajouter un `key` unique ici
                #                     st.plotly_chart(figures[i + j], use_container_width=True, key=f"fig_{i}_{j}")
                #                     st.markdown("</div>", unsafe_allow_html=True)

        # ---------- EXPORT ---------- #
        st.subheader("⬇️ Exporter les données et visualisations")

        if st.button("📤 Exporter en Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # === 1. EXPORT TABLEAUX AVEC STYLE ===
                def style_dataframe(df):
                    return df.style.set_properties(**{
                        'background-color': '#f0f8ff',
                        'color': 'black',
                        'border-color': 'gray',
                        'border-width': '1px',
                        'border-style': 'solid'
                    }).set_table_styles([{
                        'selector': 'th',
                        'props': [
                            ('background-color', '#1f4e79'),
                            ('color', 'white'),
                            ('font-weight', 'bold')
                        ]
                    }])


                style_dataframe(df1).to_excel(writer, sheet_name="Tableau de Bord REEL N vs REEL N-1", index=False)
                style_dataframe(df2).to_excel(writer, sheet_name="Tableau de Bord REEL N vs Objectifs Opérationnels N", index=False)
                style_dataframe(df3).to_excel(writer, sheet_name="Tableau de Bord REEL N vs Objectifs Stratégiques N", index=False)

                # Onglet Graphiques
                workbook = writer.book
                sheet = workbook.create_sheet(title='Graphiques')

                # Graphique 1 - Eau consommée (barre de progression réelle vs objectif)

                # Convertir en image PNG
                # Graphique 1 - Eau consommée (barre de progression réelle vs objectif)
                img_buffer1 = BytesIO()
                fig1.write_image(img_buffer1, format="png")  # ✅ write_image pour Plotly
                img_buffer1.seek(0)
                img1 = XLImage(img_buffer1)
                sheet.add_image(img1, 'B2')


                # Graphique 2 - % d’eau stockée avec jauge

                # Convertir le graphique de la jauge en image PNG et l'ajouter au fichier Excel
                img_buffer2 = BytesIO()
                fig2.write_image(img_buffer2, format="png")
                img_buffer2.seek(0)
                img2 = XLImage(img_buffer2)
                sheet.add_image(img2, 'M2')  # Positionner le graphique de la jauge à une position différente

                # # Graphique 3 - Fréquence Réunion RSE
                # img_buffer3 = BytesIO()
                # fig3.savefig(img_buffer3, format="png")
                # img_buffer3.seek(0)
                # img3 = XLImage(img_buffer3)
                # sheet.add_image(img3, 'B16')  # Positionner le 3e graphique à une position différente


                                # Graphique 4 
                img_buffer4 = BytesIO()
                fig4.write_image(img_buffer4, format="png")
                img_buffer4.seek(0)
                img4 = XLImage(img_buffer4)
                sheet.add_image(img4, 'M31') 



                                                # Graphique 5
                img_buffer5 = BytesIO()
                fig5.write_image(img_buffer5, format="png")
                img_buffer5.seek(0)
                img5 = XLImage(img_buffer5)
                sheet.add_image(img5, 'M48') 




                                # Graphique 6
                img_buffer6 = BytesIO()
                fig6.write_image(img_buffer6, format="png")
                img_buffer6.seek(0)
                img6 = XLImage(img_buffer6)
                sheet.add_image(img6, 'M20') 

                # img_buffer_femmes = BytesIO()
                # fig_femmes.write_image(img_buffer_femmes, format="png")
                # img_buffer_femmes.seek(0)
                # img_femmes = XLImage(img_buffer_femmes)
                # sheet.add_image(img_femmes, 'M28') 

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=output.getvalue(),
                file_name="xtrafi_export_avec_graphiques.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("💡 Téléversez un fichier ou déposez-en un dans le dossier `data/` pour démarrer.")
