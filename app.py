# -------------------- CONFIG PAGE (doit venir TOUT EN PREMIER) --------------------
import streamlit as st
# st.set_page_config(page_title="Xtrafi Data Viz", layout="wide")

# -------------------- AUTRES IMPORTS --------------------
import os
import uuid
import pandas as pd

# -------------------- IMPORTS MAISON -----------------
from config import WATCHED_FOLDER, LOGO_PATH
from data_loader import load_data
from components.sidebar import sidebar_file_selection
from preprocessing import filter_and_rename_columns, compute_variations,plot_variations_detaillees
from components.tab_dataframes import afficher_onglet, filtrer_par_code_axe      # ✅ nouvelle fonction unique
from components.tab_dashboards import afficher_tableau
from components.graphs_dashboard import generate_graphs_par_indicateur_en_colonnes
from components.export_excel import export_excel_with_figures
from utils.styles import style_kpi, load_css
from authentification.auth import load_auth_config, get_org_logo


# Charge le CSS global
load_css("static/style.css")

# -------------------- LOGOUT UTIL --------------------
def logout(authenticator):
    try:
        authenticator.logout()
    except Exception as e:
        st.warning(f"Erreur lors de la déconnexion : {e}")

    # Vide la session
    for k in ("authentication_status", "username", "name"):
        st.session_state.pop(k, None)

    st.session_state["logout_triggered"] = True
    st.rerun()


# -------------------- APP PRINCIPALE -----------------
def run_app(name, authenticator=None):
    # ----- SIDEBAR -----
    with st.sidebar:
        config = load_auth_config()
        org = st.session_state.get("organization")
        logo_path = get_org_logo(org, config) or LOGO_PATH

        st.image(logo_path, use_container_width=True)
        st.markdown(" " * 4, unsafe_allow_html=True)  # petit espace
        st.markdown(f"### Bienvenue {name}")

        uploaded_file, selected_file = sidebar_file_selection()

        st.markdown("<div style='flex-grow:1'></div>", unsafe_allow_html=True)  # pousse le reste vers le bas
        if st.button("Déconnexion", key="logout_button"):
            logout(authenticator)

    # ----- TITRE -----
    st.markdown("""
        <h1 style='margin-top:-40px; font-size:3rem; font-weight:bold; color:#0b1f3a;'>
            Votre Tableau de Bord ESG
        </h1>
    """, unsafe_allow_html=True)

    # ----- CHARGEMENT DONNÉES -----
    df1, df2, df3, source = load_data(uploaded_file, selected_file)
    if source:
        st.success(source)

    if not all([df1 is not None, df2 is not None, df3 is not None]):
        st.info("Importez ou sélectionnez un fichier pour commencer.")
        st.stop()

    df2_apercu = df2.drop(columns=[col for col in df2.columns if col.startswith("Unnamed")]).fillna("-")
 


    # ----- ONGLETS DE DONNÉES BRUTES -----    

    onglet_map = {
        "📈 Paramètres restitution": (df1, 1),
        "📊 Données brutes": (df2_apercu, 2),
        # "📋 Rapport consolidé": (df3, 3),
    }
    choix = st.radio("🧭 Choisissez l'onglet (aperçu des données) :", list(onglet_map.keys()))
    afficher_onglet(*onglet_map[choix])                       # ✅ un seul appel

    cols1 = [
        'Axe',
        # 'Code\nRAPPORT Ind.\nVIRTUEL', 'Nom\nRAPPORT Ind.\nVIRTUEL',
        'Code Reporting', 'Nom Reporting', 'Nom indicateur Virtuel',
        'Reel N-1', 'Reel N', 'VARIATION Réel N vs Réel N-1 (%)',
        "Unité de conversion de l'indicateur",
        'Valorisation Financière REEL N-1', 'Valorisation Financière REEL N',
    ]

    cols2 = [
        'Axe',
        # 'Code\nRAPPORT Ind.\nVIRTUEL', 'Nom\nRAPPORT Ind.\nVIRTUEL',
        'Code Reporting', 'Nom Reporting', 'Nom indicateur Virtuel',
        'Reel N', 'Objectifs Opérationnels SEUIL période N',
        'VARIATION Objectifs Opérationnels SEUIL période N vs Réel N (%)',
        'Objectifs Opérationnels PLAFOND période N',
        'VARIATION Objectifs Opérationnels PLAFOND période N vs Réel N (%)',
        "Unité de conversion de l'indicateur",
        'Valorisation Financière REEL N',
        'Valorisation Financière Objectifs Opérationnels N',
    ]

    cols3 = [
        'Axe',
        # 'Code\nRAPPORT Ind.\nVIRTUEL', 'Nom\nRAPPORT Ind.\nVIRTUEL',
        'Code Reporting', 'Nom Reporting', 'Nom indicateur Virtuel',
        'Reel N', 'Objectifs Stratégiques SEUIL période N',
        'VARIATION Objectifs Stratégiques SEUIL période N vs Réel N (%)',
        'Objectifs Stratégiques PLAFOND période N',
        'VARIATION Objectifs Stratégiques PLAFOND période N vs Réel N (%)',
        "Unité de conversion de l'indicateur",
        'Valorisation Financière REEL N',
        'Valorisation Financière Objectifs Stratégiques N',
    ]

    # ----- TRAITEMENT + KPI -----
    df3 = filter_and_rename_columns(df3)
    df3 = compute_variations(df3, col_reel_n='Reel N', col_reel_n1='Reel N-1')



    # st.markdown("---")
    # tab1, tab2, tab3, tab4 = st.tabs(
    #     ["📋 REEL N vs N-1",
    #      "📋 REEL N vs Obj. Opérationnels",
    #      "📋 REEL N vs Obj. Stratégiques",
    #      "📊 Visualisations"] 
    # ) 

    # --- Tableaux KPI ---  
    st.markdown("---")
    onglets_labels = [
        "📋 REEL N vs N-1",
        "📋 REEL N vs Obj. Opérationnels",
        "📋 REEL N vs Obj. Stratégiques",
        "📊 Visualisations",
        "Vision globale"
    ]

    if "onglet_actif" not in st.session_state:
        st.session_state.onglet_actif = 0

    choix_onglet = st.radio("Sélectionnez l'onglet", onglets_labels, index=st.session_state.onglet_actif, horizontal=True)
    st.session_state.onglet_actif = onglets_labels.index(choix_onglet)

    # --- Tableaux KPI ---
    if st.session_state.onglet_actif == 0:
        df3_filtré = filtrer_par_code_axe(df3, key_prefix="tab1")

        # Stocke le DF filtré dans session_state pour export dynamique
        st.session_state['df_tab1_filtré'] = df3_filtré[cols1]

        afficher_tableau("📋 Tableau REEL N vs N‑1", st.session_state['df_tab1_filtré'], style_kpi)

    elif st.session_state.onglet_actif == 1:
        df3_filtré = filtrer_par_code_axe(df3, key_prefix="tab2")

        st.session_state['df_tab2_filtré'] = df3_filtré[cols2]
        afficher_tableau("📋 REEL N vs Obj. Opérationnels", st.session_state['df_tab2_filtré'], style_kpi)

    elif st.session_state.onglet_actif == 2:
        df3_filtré = filtrer_par_code_axe(df3, key_prefix="tab3")

        st.session_state['df_tab3_filtré'] = df3_filtré[cols3]
        afficher_tableau("📋 REEL N vs Obj. Stratégiques", st.session_state['df_tab3_filtré'], style_kpi)

    elif st.session_state.onglet_actif == 3:
        df2_filtré = filtrer_par_code_axe(df2, key_prefix="tab4")

        st.markdown("---")

        figs, fig_infos = generate_graphs_par_indicateur_en_colonnes(
            df2_filtré,
            col_indicateur="Nom Ind.N2\nAPP",
            col_val_n1="Total\nMontant\nCollecte\nRéelle\nExercice N-1",
            col_val_n="Total\nMontant\nCollecte\nRéelle\nExercice N",
            colonnes_objectifs=[
                'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N',
                'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N',
                'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N',
                'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N',
            ],
            n_cols=2
        )

        st.session_state['last_figs'] = figs
        st.session_state['last_fig_infos'] = fig_infos

    elif st.session_state.onglet_actif == 4:

        fig_glob = plot_variations_detaillees(df3)
        st.plotly_chart(fig_glob, use_container_width=True)
        st.session_state['fig_glob'] = fig_glob

    # ----- EXPORT -----
    
    st.markdown("---")
    if st.button("📤 Exporter Excel + Graphiques"):
        df1_export = st.session_state.get('df_tab1_filtré')
        df2_export = st.session_state.get('df_tab2_filtré')
        df3_export = st.session_state.get('df_tab3_filtré')

        if df1_export is None:
            df1_export = df3[cols1]  # données complètes, non filtrées 
        if df2_export is None: 
            df2_export = df3[cols2]
        if df3_export is None:
            df3_export = df3[cols3]

        df_list = [
            (df1_export, "REEL_N_vs_N-1"),
            (df2_export, "REEL_N_vs_Obj_Op"),
            (df3_export, "REEL_N_vs_Obj_Strat"),
        ]

        figs = st.session_state.get('last_figs', [])
        fig_glob = st.session_state.get('fig_glob', None)


        xls = export_excel_with_figures(df_list, figs, fig_glob=fig_glob)

        st.download_button(
            "📥 Télécharger l'export",
            data=xls,
            file_name="xtrafiBI_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


    # --- Graphiques ---


    # with tab4:
        # # Choix utilisateur entre df2 et df3
        # df_choice = st.radio(
        #     "Choisissez la source des données pour les graphiques",
        #     options=["Données brutes", "Données consolidées"]
        # )

        # # Récupérer la bonne DataFrame selon choix
        # if df_choice == "Données brutes":
        #     df_source = df2
        #     col_indicateur = "Nom Ind.N2\nAPP" 
        #     col_val_n1 = "Total\nMontant\nCollecte\nRéelle\nExercice N-1"
        #     col_val_n = "Total\nMontant\nCollecte\nRéelle\nExercice N"
        # else:
        #     df_source = df3
        #     col_indicateur = "Nom indicateur Virtuel"  
        #     col_val_n1 = "Reel N-1"
        #     col_val_n = "Reel N"
        


        # # Appliquer le filtre ESG sur la DataFrame choisie
        # df_filtré = filtrer_par_code_axe(df_source, key_prefix="tab4")

        # st.markdown("---")
        # st.markdown(f"✅ Vous visualisez les graphiques à partir des **{df_choice.lower()}**")
        # st.markdown("---")

        # figs = generate_graphs_par_indicateur_en_colonnes(
        #     df_filtré,
        #     col_indicateur=col_indicateur,
        #     col_val_n1=col_val_n1,
        #     col_val_n=col_val_n,
        #     n_cols=2
        # )

    # with tab4:
    #     df2_filtré = filtrer_par_code_axe(df2, key_prefix="tab4")

    #     st.markdown("---")

    #     figs, fig_infos = generate_graphs_par_indicateur_en_colonnes(
    #         df2_filtré,
    #         col_indicateur="Nom Ind.N2\nAPP",
    #         col_val_n1="Total\nMontant\nCollecte\nRéelle\nExercice N-1",
    #         col_val_n="Total\nMontant\nCollecte\nRéelle\nExercice N",
    #         colonnes_objectifs=[
    #             'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N',
    #             'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N',
    #             'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N',
    #             'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N',

    #         ],
    #         n_cols=2
    #     )
        
    #         # Stocke graphs pour export
    #     st.session_state['last_figs'] = figs
    #     st.session_state['last_fig_infos'] = fig_infos


    # # ----- EXPORT -----
    # st.markdown("---")
    # if st.button("📤 Exporter Excel + Graphiques"):
    #     df_list = [
    #         (df3[cols1], "REEL N_vs_N‑1"),
    #         (df3[cols2], "REEL N_vs_Obj_Op"),
    #         (df3[cols3], "REEL N_vs_Obj_Strat"),
    #     ]
    #     # fig_list = [fig_eau, fig_eau_stockee, fig_consommation_eau, fig_carburant, fig_rgaes]
    #     xls = export_excel_with_figures(df_list, figs)
    #     st.download_button(
    #         "📥 Télécharger l'export",
    #         data=xls,
    #         file_name="xtrafiBI_export.xlsx",
    #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #     )
