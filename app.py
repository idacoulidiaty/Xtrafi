# -------------------- CONFIG PAGE (doit venir TOUT EN PREMIER) --------------------
import streamlit as st
# st.set_page_config(page_title="Xtrafi Data Viz", layout="wide")

# -------------------- AUTRES IMPORTS --------------------

import pandas as pd

# -------------------- IMPORTS MAISON -----------------
from config import WATCHED_FOLDER, LOGO_PATH
from data_loader import load_data
# from components.sidebar import sidebar_file_selection
from preprocessing import filter_and_rename_columns, compute_variations,plot_variations_detaillees
from components.tab_dataframes import afficher_onglet, filtrer_par_code_axe     
from components.tab_dashboards import afficher_tableau
from components.graphs_dashboard import generate_graphs_par_indicateur_en_colonnes
from components.export_excel import export_excel_with_figures,sanitize_for_excel
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

        # selected_file = sidebar_file_selection()

        st.markdown("<div style='flex-grow:1'></div>", unsafe_allow_html=True)  # pousse le reste vers le bas
        # if st.button("Retour à la Selection d'Organisations"):

        if st.button("Déconnexion", key="logout_button"):
            logout(authenticator)

    # ----- TITRE -----
    st.markdown("""
        <h1 style='margin-top:-40px; font-size:3rem; font-weight:bold; color:#0b1f3a;'>
            Votre Tableau de Bord Durabilité et Donut
        </h1>
    """, unsafe_allow_html=True)


    # ----- AFFICHAGE SEULEMENT SI UN FICHIER EST UPLOADÉ -----
    uploaded_file = st.file_uploader("Téléversez un fichier Excel", type=["xlsx", "xls"])
    if not uploaded_file:
        st.info("Chargez un fichier Excel de Restitution pour visualiser votre tableau de bord") 

    elif uploaded_file:
        # ----- CHARGEMENT DONNÉES -----
        df1, df2, df3, source = load_data(uploaded_file) #, selected_file)

        if source:
            st.success(source)

        # Initialisation des aperçus pour éviter UnboundLocalError
        df1_apercu, df2_apercu, df3_apercu = None, None, None

        # Vérification de présence des onglets
        if df1 is None:
            st.warning("⚠️Impossible d'afficher l'aperçu de l'onglet 1 (manquant).")
        else:
            df1_apercu = df1.copy()

        if df2 is None:
            st.warning("⚠️Impossible d'afficher l'aperçu de l'onglet 2 (manquant).")
            st.stop()
        else:
            # ----- Nettoyage colonnes -----
            df2_clean = df2.drop(columns=[col for col in df2.columns if col.startswith("Unnamed")]).fillna("-")

            # Colonnes de contexte (toujours affichées si présentes)
            colonnes_contexte = [
                "Code Axe RAPPORT",
                "Nom axe RAPPORT",
                "Code axe APP",
                "Nom axe APP",
                "Nom Question\nAPP",
                "Unité\nconversion",
                "Unité\nValorisation\nfinancière",
            ]

            # Règles de sélection dynamiques
            prefixes_selection = [
                "Total\nMontant\nCollecte\nRéelle\nExercice N",
                "Total\nValo. Financière\nCollecte Réelle\nExercice N",
                "Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N",
                "Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N",
                "Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N",
            ]

            # Colonnes dynamiques selon règles
            colonnes_dynamiques = [
                col for col in df2_clean.columns
                if any(col.startswith(prefix) for prefix in prefixes_selection)
            ]

            # Construction de l’aperçu
            colonnes_apercu = [c for c in colonnes_contexte if c in df2_clean.columns] + colonnes_dynamiques
            df2_apercu = df2_clean[colonnes_apercu]

        # ----- TRAITEMENT DF3 + KPI -----
        if df3 is not None:
            df3, missing_cols = filter_and_rename_columns(df3)

            # Afficher un seul warning si des colonnes sont manquantes
            if missing_cols:
                st.warning(
                    f"⚠️ Le fichier est incomplet : il manque les colonnes suivantes :\n- " +
                    "\n- ".join(missing_cols) +
                    "\n\n➡️ Le dashboard va quand même s'afficher, mais certaines données seront manquantes." +
                    "\n\nVeuillez télécharger un fichier de restitution à jour depuis Xtrafi App"
                )
        
            df3 = compute_variations(df3, col_reel_n='Reel N', col_reel_n1='Reel N-1')

        # ----- ONGLETS / AFFICHAGE -----
        onglet_map = {
            "📈 Paramètres restitution": (df1, 1),
            "📊 Données brutes": (df2_apercu, 2),
            # "📋 Rapport consolidé": (df3, 3),
        }

        choix = st.radio(
            "🧭 Choisissez l'onglet (aperçu des données) :", 
            list(onglet_map.keys())
        )

        # Mettre l’aperçu des données dans un expander
        with st.expander(f"Aperçu des données - {choix}", expanded=True):
            afficher_onglet(*onglet_map[choix])


        cols1 = [
            'Axe',
            'Code Rapport Indicateur Virtuel', 'Code App indicateur Virtuel',
            'Nom indicateur Virtuel',
            'Reel N-1', 'Reel N', 'VARIATION Réel N vs Réel N-1 (%)',
            "Unité de conversion de l'indicateur",
            'Valorisation Financière REEL N-1', 'Valorisation Financière REEL N',
            'Unité Valorisation financière'
        ]

        cols2 = [
            'Axe',
            'Code Rapport Indicateur Virtuel', 'Code App indicateur Virtuel',
            'Nom indicateur Virtuel',
            'Reel N', 'Objectifs Opérationnels SEUIL période N',
            'VARIATION Objectifs Opérationnels SEUIL période N vs Réel N (%)',
            'Objectifs Opérationnels PLAFOND période N',
            'VARIATION Objectifs Opérationnels PLAFOND période N vs Réel N (%)',
            "Unité de conversion de l'indicateur",
            'Valorisation Financière REEL N',
            'Valorisation Financière Objectifs Opérationnels N',
            'Unité Valorisation financière'
        ]

        cols3 = [
            'Axe',
            'Code Rapport Indicateur Virtuel', 'Code App indicateur Virtuel',
            'Nom indicateur Virtuel',
            'Reel N', 'Objectifs Stratégiques SEUIL période N',
            'VARIATION Objectifs Stratégiques SEUIL période N vs Réel N (%)',
            'Objectifs Stratégiques PLAFOND période N',
            'VARIATION Objectifs Stratégiques PLAFOND période N vs Réel N (%)',
            "Unité de conversion de l'indicateur",
            'Valorisation Financière REEL N',
            'Valorisation Financière Objectifs Stratégiques N',
            'Unité Valorisation financière'
        ]

        # ----- TRAITEMENT + KPI -----
        # df3 = filter_and_rename_columns(df3)
        # df3 = compute_variations(df3, col_reel_n='Reel N', col_reel_n1='Reel N-1')



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
            available_cols1 = [col for col in cols1 if col in df3_filtré.columns]
            st.session_state['df_tab1_filtré'] = df3_filtré[available_cols1].sort_values(by=['Code App indicateur Virtuel'],ignore_index=True)
            afficher_tableau("📋 Tableau REEL N vs N‑1", st.session_state['df_tab1_filtré'], style_kpi, logo_path)

        elif st.session_state.onglet_actif == 1:
            df3_filtré = filtrer_par_code_axe(df3, key_prefix="tab2")
            available_cols2 = [col for col in cols2 if col in df3_filtré.columns]
            st.session_state['df_tab2_filtré'] = df3_filtré[available_cols2].sort_values(by=['Code App indicateur Virtuel'],ignore_index=True)
            afficher_tableau("📋 REEL N vs Obj. Opérationnels", st.session_state['df_tab2_filtré'], style_kpi, logo_path)

        elif st.session_state.onglet_actif == 2:
            df3_filtré = filtrer_par_code_axe(df3, key_prefix="tab3")
            available_cols3 = [col for col in cols3 if col in df3_filtré.columns]
            st.session_state['df_tab3_filtré'] = df3_filtré[available_cols3].sort_values(by=['Code App indicateur Virtuel'],ignore_index=True)
            afficher_tableau("📋 REEL N vs Obj. Stratégiques", st.session_state['df_tab3_filtré'], style_kpi, logo_path)

        elif st.session_state.onglet_actif == 3:
            col1,col2 = st.columns([2,1])
            with col1:
                st.subheader("📊 Visualisations graphiques des questions")
            with col2:
                st.image(logo_path,width=100)
            # st.markdown("---")
            df2_filtré = filtrer_par_code_axe(df2, key_prefix="tab4")

            figs, fig_infos = generate_graphs_par_indicateur_en_colonnes(
                df2_filtré,
                col_indicateur="Nom Question\nAPP",
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
            col1,col2 = st.columns([2,1])
            with col1:
                st.subheader("📊 Synthèse des évolutions du Réel entre N-1 et N en fonction des Objectifs Opérationnels")
            with col2:
                st.image(logo_path,width=100)
            # st.markdown("---")
            fig_glob = plot_variations_detaillees(df3)
            st.plotly_chart(fig_glob, use_container_width=True)
            st.session_state['fig_glob'] = fig_glob

        # ----- EXPORT -----

        st.markdown("---")

        # Données à exporter
        df1_export = st.session_state.get('df_tab1_filtré')
        df2_export = st.session_state.get('df_tab2_filtré')
        df3_export = st.session_state.get('df_tab3_filtré')

        if df1_export is None:
            available_cols1 = [col for col in cols1 if col in df3.columns]
            df1_export = df3[available_cols1]

        if df2_export is None:
            available_cols2 = [col for col in cols2 if col in df3.columns]
            df2_export = df3[available_cols2]

        if df3_export is None:
            available_cols3 = [col for col in cols3 if col in df3.columns]
            df3_export = df3[available_cols3]

        df_list = [
            (sanitize_for_excel(df1_export), "REEL_N_vs_N-1"),
            (sanitize_for_excel(df2_export), "REEL_N_vs_Obj_Op"),
            (sanitize_for_excel(df3_export), "REEL_N_vs_Obj_Strat"),
        ]

        figs = st.session_state.get('last_figs', [])
        fig_glob = st.session_state.get('fig_glob', None)

        # Génération du fichier
        xls = export_excel_with_figures(df_list, figs, fig_glob=fig_glob)

        # 📥 Bouton unique
        st.download_button(
            "📤 Exporter Excel + Graphiques",
            data=xls,
            file_name="xtrafiBI_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    

   