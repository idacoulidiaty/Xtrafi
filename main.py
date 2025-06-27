
import streamlit as st

# 🔧 Configuration page
st.set_page_config(
    page_title="Xtrafi Data Viz",
    layout="wide"
)

# 🎨 STYLES CSS PERSONNALISÉS
import streamlit as st

st.markdown("""
    <style>
    /* 🌑 Couleur bleu nuit pour la sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0b1f3a;
    }

    /* 🎨 Texte blanc dans la sidebar */
    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    /* 📁 Bloc de téléversement personnalisé */
    section[data-testid="stSidebar"] .st-emotion-cache-taue2i {
        background-color: #112d50 !important;
        color: white !important;
        border-radius: 10px;
        padding: 1rem;
    }

    /* --- 🚀 Bouton de téléversement --- */
section[data-testid="stSidebar"] .stFileUploader button {
    background-color: #19549d !important;
    color: white !important;
    font-weight: bold;
    border: none;
    border-radius: 6px;
    padding: 0.4rem 1rem;
}

    /* --- 🔽 Selectbox (champ visible) --- */
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #0b1f3a !important;
    color: white !important;
    border-radius: 6px !important;
    font-weight: bold;
}

/* --- 🔽 Menu déroulant selectbox --- */
div[role="listbox"] {
    background-color: #0b1f3a !important;
    color: white !important;
    border-radius: 6px;
}

/* --- 🟦 Chaque option dans menu selectbox --- */
div[role="option"] {
    background-color: #0b1f3a !important;
    color: white !important;
}

/* --- 🖱 Option survolée --- */
div[role="option"]:hover {
    background-color: #133456 !important;
}

/* --- 🧼 Supprimer bordures et marges parasites --- */
.st-emotion-cache-1kyxreq, .st-emotion-cache-16txtl3 {
    background-color: transparent !important;
    padding: 0 !important;}

    /* 🧢 Titre principal plus gros */
    .main > div > div > div > h1 {
        font-size: 2.2rem;
        font-weight: 800;
        color: #0b1f3a;
    }
            



    </style>
""", unsafe_allow_html=True)




import os
import pandas as pd

from config import WATCHED_FOLDER, LOGO_PATH
from data_loader import (
    get_latest_excel_file,
    load_data,
    get_date_generation_restitution
)
from components.sidebar import sidebar_file_selection
from preprocessing import filter_and_rename_columns, compute_variations
from components.tab_dataframes import afficher_onglet_1, afficher_onglet_2, afficher_onglet_3
from components.tab_dashboards import afficher_tableau
from utils.styles import style_kpi
from components.graphs_dashboard import (
    afficher_graphique_eau_stockee,
    afficher_graphique_consommation_eau,
    afficher_graphique_eau,
    afficher_graphique_carburant,
    afficher_graphique_rgaes
)
from components.export_excel import export_excel_with_figures





st.markdown("""
<div style='margin-top: -90px; margin-bottom: 40px;'>
    <h1 style='font-size: 80px; font-weight: bold; color: #0b1f3a;'>Votre Tableau de Bord ESG</h1>
</div>
""", unsafe_allow_html=True)



# Upload et sélection fichier dans sidebar (avec ta fonction dédiée)
uploaded_file, selected_file = sidebar_file_selection()

# Chargement des données via la fonction load_data() qui encapsule la logique
df1, df2, df3, source = load_data(uploaded_file, selected_file)

# Affichage source si ok
if source:
    st.success(source)

# Affichage des onglets si données présentes
if df1 is not None and df2 is not None and df3 is not None:

    onglet = st.radio("🧭 Choisissez l'onglet :", [
        "📈 Paramètres restitution",
        "📊 Données brutes",
        "📋 Rapport consolidé"
    ])

    if onglet == "📈 Paramètres restitution":
        afficher_onglet_1(df1)
    elif onglet == "📊 Données brutes":
        afficher_onglet_2(df2)
    elif onglet == "📋 Rapport consolidé":
        afficher_onglet_3(df3)

    df3 = filter_and_rename_columns(df3)
    df3 = compute_variations(df3)

    st.markdown("---")
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 REEL N vs N-1",
        "📋 REEL N vs Objectifs Opérationnels",
        "📋 REEL N vs Objectifs Stratégiques",
        "📊 Visualisations"
    ])

    with tab1:
        cols_tab1 = [
            'Code\nRAPPORT Ind.\nVIRTUEL', 'Nom\nRAPPORT Ind.\nVIRTUEL',
            'Code Reporting', 'Nom Reporting', 'Nom indicateur Virtuel',
            'Reel N-1', 'Reel N', 'VARIATION Réel N vs Réel N-1 (%)',
            "Unité de conversion de l'indicateur",
            'Valorisation Financière REEL N-1', 'Valorisation Financière REEL N'
        ]
        afficher_tableau("📋 Tableau de bord REEL N vs REEL N-1", df3[cols_tab1], style_kpi)

    with tab2:
        cols_tab2 = [
            'Code\nRAPPORT Ind.\nVIRTUEL', 'Nom\nRAPPORT Ind.\nVIRTUEL',
            'Code Reporting', 'Nom Reporting', 'Nom indicateur Virtuel',
            'Reel N', 'Objectifs Opérationnels SEUIL période N',
            'VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)',
            'Objectifs Opérationnels PLAFOND période N',
            'VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)',
            "Unité de conversion de l'indicateur",
            'Valorisation Financière REEL N',
            'Valorisation Financière Objectifs Opérationnels N'
        ]
        afficher_tableau("📋 REEL N vs Objectifs Opérationnels", df3[cols_tab2], style_kpi)

    with tab3:
        cols_tab3 = [
            'Code\nRAPPORT Ind.\nVIRTUEL', 'Nom\nRAPPORT Ind.\nVIRTUEL',
            'Code Reporting', 'Nom Reporting', 'Nom indicateur Virtuel',
            'Reel N', 'Objectifs Stratégiques SEUIL période N',
            'VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)',
            'Objectifs Stratégiques PLAFOND période N',
            'VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)',
            "Unité de conversion de l'indicateur",
            'Valorisation Financière REEL N',
            'Valorisation Financière Objectifs Stratégiques N'
        ]
        afficher_tableau("📋 REEL N vs Objectifs Stratégiques", df3[cols_tab3], style_kpi)


    with tab4:

        col1,col2 = st.columns(2)  
        
        with col1:
            fig_eau = afficher_graphique_eau(df2)

        with col2:
            fig_eau_stockee = afficher_graphique_eau_stockee(df2)

        col3, col4 = st.columns(2)
        with col3 :
            fig_consommation_eau = afficher_graphique_consommation_eau(df2)

        with col4:
            fig_carburant = afficher_graphique_carburant(df2)

        col5,col6 = st.columns(2)
        with col5:
            fig_rgaes = afficher_graphique_rgaes(df2)




 # ---------- EXPORT ---------- #
st.markdown("---")
st.subheader("⬇️ Exporter les données et visualisations")

if st.button("📤 Exporter en Excel"):
    df_list = [
        (df3[cols_tab1], "REEL N vs REEL N-1"),
        (df3[cols_tab2], "REEL N vs Obj Opérationnels"),
        (df3[cols_tab3], "REEL N vs Obj Stratégiques")
    ]

    fig_list = [
        fig_eau,
        fig_eau_stockee,
        fig_consommation_eau,
        fig_carburant,
        fig_rgaes
    ]


    excel_bytes = export_excel_with_figures(
        df_list=df_list,
        fig_list=fig_list
    )

    st.download_button(
        label="📥 Télécharger le fichier Excel",
        data=excel_bytes,
        file_name="xtrafi_export_avec_graphiques.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
