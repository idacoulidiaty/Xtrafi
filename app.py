import streamlit as st  # ← À mettre en premie
import os
import pandas as pd

# 🔧 Configuration page
st.set_page_config(
    page_title="Xtrafi Data Viz",
    layout="wide"
)

# 🎨 STYLES CSS PERSONNALISÉS
from utils.styles import load_css

# css_path = "static/style.css"
# load_css(css_path)

# -------------------- IMPORTS XTRAFI --------------------
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


import streamlit as st
from streamlit.runtime.scriptrunner import RerunException
def logout(authenticator):
    try:
        authenticator.logout()
    except Exception as e:
        st.warning(f"Erreur lors de la déconnexion : {e}")

    # Nettoyer le session state
    for key in ("authentication_status", "username", "name"):
        st.session_state.pop(key, None)

    # Indiquer qu’on a demandé logout
    st.session_state["logout_triggered"] = True

    # Forcer le rechargement de l’app
    st.rerun()


def run_app(name, username, authenticator=None):
    with st.sidebar:
        # Affiche logo UNE SEULE FOIS ici, tout en haut
        st.image(LOGO_PATH, use_container_width=True)

        # Espace important après logo
        st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)

        # Bienvenue utilisateur
        st.markdown(f"### Bienvenue {name} 👋")

        st.markdown("<div style='height: 100px;'></div>", unsafe_allow_html=True)

        # Appelle la fonction d'upload + liste déroulante SANS LOGO dedans !
        uploaded_file, selected_file = sidebar_file_selection()

        st.markdown("<div style='height: 150px;'></div>", unsafe_allow_html=True)

        # Espace flexible pour pousser le bouton déconnexion en bas
        st.markdown("<div style='flex-grow:1'></div>", unsafe_allow_html=True)

        # Bouton déconnexion
        if st.button("Déconnexion", key="logout_button"):
            st.write("Déconnexion demandée")
            logout(authenticator)


    # Titre principal
    st.markdown("""
    <div style='margin-top: -90px; margin-bottom: 40px;'>
        <h1 style='font-size: 80px; font-weight: bold; color: #0b1f3a;'>Votre Tableau de Bord ESG</h1>
    </div>
    """, unsafe_allow_html=True)

    # Chargement des données (depuis upload ou sélection)
    df1, df2, df3, source = load_data(uploaded_file, selected_file)

    # Affichage source si ok
    if source:
        st.success(source)

    # Si données présentes, afficher les onglets
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
            col1, col2 = st.columns(2)
            with col1:
                fig_eau = afficher_graphique_eau(df2)
            with col2:
                fig_eau_stockee = afficher_graphique_eau_stockee(df2)

            col3, col4 = st.columns(2)
            with col3:
                fig_consommation_eau = afficher_graphique_consommation_eau(df2)
            with col4:
                fig_carburant = afficher_graphique_carburant(df2)

            col5, col6 = st.columns(2)
            with col5:
                fig_rgaes = afficher_graphique_rgaes(df2)

        st.markdown("---")
        st.subheader("⬇️ Exporter les données et visualisations")

        if st.button("📤 Exporter en Excel"):
            df_list = [
                (df3[cols_tab1], "REEL N vs REEL N-1"),
                (df3[cols_tab2], "REEL N vs Obj Opérationnels"),
                (df3[cols_tab3], "REEL N vs Obj Stratégiques")
            ]
            fig_list = [
                fig_eau, fig_eau_stockee, fig_consommation_eau, fig_carburant, fig_rgaes
            ]
            excel_bytes = export_excel_with_figures(df_list=df_list, fig_list=fig_list)
            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=excel_bytes,
                file_name="xtrafi_export_avec_graphiques.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
