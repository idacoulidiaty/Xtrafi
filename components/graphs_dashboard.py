import streamlit as st
import numpy as np
import pandas as pd
from components.tab_graphs import add_horizontal_line, bar_comparatif
import plotly.graph_objects as go

def afficher_graphique_eau(df_onglet_2):
    # Filtrer indicateurs eau consommée
    df_eau = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("eau consommée", case=False, na=False)]
    if df_eau.empty:
        st.warning("⚠️ Données eau consommée manquantes")
        return
    
    st.subheader("Total consommation d'eau N-1 vs N")

    # Préparer la figure via bar_comparatif
    fig = bar_comparatif(
        df=df_eau,
        col_x='Nom Ind.N2\nAPP',
        col_y_n1='Total\nMontant\nCollecte\nRéelle\nExercice N-1',
        col_y_n='Total\nMontant\nCollecte\nRéelle\nExercice N',
        label_x="Indicateurs Eau",
        label_y="Volume (Litres)"
    )

    # Moyennes objectifs strat/opér pour lignes horizontales
    plafond_strat = np.nanmean(df_eau['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'])
    plancher_strat = np.nanmean(df_eau['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'])
    plafond_oper = np.nanmean(df_eau['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'])
    plancher_oper = np.nanmean(df_eau['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'])

    options_obj = st.multiselect(
        "🎯 Objectifs à afficher sur le graphique Eau",
        options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
        default=[],
        key="multiselect_obj_eau"
    )

    if "Stratégique - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_strat, "Obj. Strat. plafond moyen", "green", "dash", x_start=-0.5, x_end=len(df_eau)-0.5)
    if "Stratégique - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_strat, "Obj. Strat. plancher moyen", "green", "dot", x_start=-0.5, x_end=len(df_eau)-0.5)
    if "Opérationnel - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_oper, "Obj. Opé. plafond moyen", "orange", "dash", x_start=-0.5, x_end=len(df_eau)-0.5)
    if "Opérationnel - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_oper, "Obj. Opé. plancher moyen", "orange", "dot", x_start=-0.5, x_end=len(df_eau)-0.5)

    fig.update_layout(
        title="Total consommation d'eau N-1 vs N",
        barmode='group',
        xaxis_title="Indicateurs Eau",
        yaxis_title="Volume (Litres)",
        template='plotly_white',
        height=500,
        bargap=0.2

    )  
    st.plotly_chart(fig, use_container_width=True)
    return fig


def afficher_graphique_carburant(df_onglet_2):
    # Filtrer indicateurs carburant
    carburant_list = [
        "Consommation totale de carburant véhicule",
        "Consommation diesel des véhicules",
        "Consommation Essence/Super des véhicules"
    ]
    df_carburant = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].isin(carburant_list)]
    

    if df_carburant.empty:
        st.warning("⚠️ Données carburant manquantes")
        return
    
    st.subheader("Consommations carburant véhicules")
    st.subheader("\n")

    # Préparer la figure
    df_carburant['Nom Ind.N2\nAPP'] = df_carburant['Nom Ind.N2\nAPP'].apply(
    lambda x: x.replace("Consommation totale de carburant véhicule", "Consommation totale<br>carburant véhicule")
             .replace("Consommation diesel des véhicules", "Consommation <br>diesel des véhicules")
             .replace("Consommation Essence/Super des véhicules", "Consommation Essence<br>/Super des véhicules")
)
    fig = bar_comparatif(
        df=df_carburant,
        col_x='Nom Ind.N2\nAPP',
        col_y_n1='Total\nMontant\nCollecte\nRéelle\nExercice N-1',
        col_y_n='Total\nMontant\nCollecte\nRéelle\nExercice N',
        label_x="Indicateurs Carburant",
        label_y="Volume (Litres)"
    )

    # Moyennes objectifs
    plafond_strat = np.nanmean(df_carburant['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'])
    plancher_strat = np.nanmean(df_carburant['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'])
    plafond_oper = np.nanmean(df_carburant['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'])
    plancher_oper = np.nanmean(df_carburant['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'])

    options_obj = st.multiselect(
        "🎯 Objectifs à afficher sur le graphique Carburant",
        options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
        default=[],
        key="multiselect_obj_carburant"
    )

    if "Stratégique - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_strat, "Obj. Strat. plafond moyen", "green", "dash", x_start=-0.5, x_end=len(df_carburant)-0.5)
    if "Stratégique - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_strat, "Obj. Strat. plancher moyen", "green", "dot", x_start=-0.5, x_end=len(df_carburant)-0.5)
    if "Opérationnel - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_oper, "Obj. Opé. plafond moyen", "orange", "dash", x_start=-0.5, x_end=len(df_carburant)-0.5)
    if "Opérationnel - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_oper, "Obj. Opé. plancher moyen", "orange", "dot", x_start=-0.5, x_end=len(df_carburant)-0.5)

    fig.update_layout(
     title = "Consommations carburant véhicules",
        xaxis=dict(
        title="Indicateurs Carburant",
        title_font=dict(size=16),
        tickangle=0, 
        automargin=True),
     template='plotly_white',
     height=500,
     bargap=0.2,
     barmode = 'group',

    )

    st.plotly_chart(fig, use_container_width=True)
    return fig


def afficher_graphique_rgaes(df_onglet_2):
    # Filtrer indicateurs RGAES
    rgaes_list = [
        "Rejets de gaz à effets de serre (RGAES)",
        "RGAES hors production éléctrique"
    ]
    df_rgaes = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].isin(rgaes_list)].copy()

    if df_rgaes.empty:
        st.warning("⚠️ Données RGAES manquantes")
        return

    # Remplacer certains textes par des versions avec <br> pour les sauts de ligne dans les labels
    df_rgaes['Nom Ind.N2\nAPP'] = df_rgaes['Nom Ind.N2\nAPP'].apply(
        lambda x: x.replace("Rejets de gaz à effets de serre (RGAES)", "Rejets de gaz à effets<br>de serre (RGAES)")
                    .replace ("RGAES hors production éléctrique", "RGAES hors production<br>éléctrique")
    )

    st.subheader("Rejets Gaz à Effet de Serre N-1 vs N")

    # Préparer la figure
    fig = bar_comparatif(
        df=df_rgaes,
        col_x='Nom Ind.N2\nAPP',
        col_y_n1='Total\nMontant\nCollecte\nRéelle\nExercice N-1',
        col_y_n='Total\nMontant\nCollecte\nRéelle\nExercice N',
        label_x="Indicateurs RGAES",
        label_y="Valeur"
    )

    # Moyennes objectifs
    plafond_strat = np.nanmean(df_rgaes['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'])
    plancher_strat = np.nanmean(df_rgaes['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'])
    plafond_oper = np.nanmean(df_rgaes['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'])
    plancher_oper = np.nanmean(df_rgaes['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'])

    options_obj = st.multiselect(
        "🎯 Objectifs à afficher sur le graphique RGAES",
        options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
        default=[],
        key="multiselect_obj_rgaes"
    )

    if "Stratégique - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_strat, "Obj. Strat. plafond moyen", "green", "dash", x_start=-0.5, x_end=len(df_rgaes)-0.5)
    if "Stratégique - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_strat, "Obj. Strat. plancher moyen", "green", "dot", x_start=-0.5, x_end=len(df_rgaes)-0.5)
    if "Opérationnel - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_oper, "Obj. Opé. plafond moyen", "orange", "dash", x_start=-0.5, x_end=len(df_rgaes)-0.5)
    if "Opérationnel - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_oper, "Obj. Opé. plancher moyen", "orange", "dot", x_start=-0.5, x_end=len(df_rgaes)-0.5)

    fig.update_layout(
        title="Rejets Gaz à Effet de Serre N-1 vs N",
        xaxis=dict(
            title="Indicateurs Carburant",
            title_font=dict(size=16),
            tickangle=0, 
            automargin=True
        ),
        template='plotly_white',
        height=500,
        bargap=0.2,
        barmode='group'
    )

    st.plotly_chart(fig, use_container_width=True)
    return fig





def afficher_graphique_eau_stockee(df_onglet_2):
    # Filtrer la ligne "eau stockée"
    stock = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("eau stockée", case=False, na=False)]
    if stock.empty:
        st.warning("⚠️ Donnée manquante pour l’eau stockée.")
        return

    st.subheader("% Eau stockée N-1 vs N")

    valeur_n1 = pd.to_numeric(stock['Total\nMontant\nCollecte\nRéelle\nExercice N-1'].values[0], errors="coerce")
    valeur_n = pd.to_numeric(stock['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0], errors="coerce")

    if pd.isnull(valeur_n1) or pd.isnull(valeur_n):
        st.warning("❌ Données non valides pour les montants Exercice N-1 ou N.")
        return

    # Créer un petit DataFrame pour bar_comparatif
    df_temp = pd.DataFrame({
        "Indicateur": ["% Eau stockée"],
        "N-1": [valeur_n1],
        "N": [valeur_n]
    })

    # Générer la figure
    fig = bar_comparatif(
        df=df_temp,
        col_x="Indicateur",
        col_y_n1="N-1",
        col_y_n="N",
        label_x="",
        label_y="Pourcentage"
    )

    # Ajouter les lignes d’objectif
    plafond_strat = pd.to_numeric(stock['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'].values[0], errors="coerce")
    plancher_strat = pd.to_numeric(stock['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'].values[0], errors="coerce")
    plafond_oper = pd.to_numeric(stock['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'].values[0], errors="coerce")
    plancher_oper = pd.to_numeric(stock['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'].values[0], errors="coerce")

    options_obj = st.multiselect(
        "🎯 Objectifs à afficher",
        options=[
            "Stratégique - Plafond",
            "Stratégique - Plancher",
            "Opérationnel - Plafond",
            "Opérationnel - Plancher"
        ],
        default=[],
        key="multiselect_obj_stockee"
    )

    if "Stratégique - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_strat, "Obj. Strat. plafond", "green", "dash", x_start=-0.5, x_end=0.5)
    if "Stratégique - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_strat, "Obj. Strat. plancher", "green", "dot", x_start=-0.5, x_end=0.5)
    if "Opérationnel - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_oper, "Obj. Opé. plafond", "orange", "dash", x_start=-0.5, x_end=0.5)
    if "Opérationnel - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_oper, "Obj. Opé. plancher", "orange", "dot", x_start=-0.5, x_end=0.5)

    fig.update_layout(
        title="% Eau stockée N-1 vs N",
        yaxis=dict(range=[0, 100]),
        height=500,
        bargap=0.2
    )

    st.plotly_chart(fig, use_container_width=True)
    return fig





def afficher_graphique_consommation_eau(df_onglet_2):
    indicateurs_eau = df_onglet_2[df_onglet_2['Nom Ind.N2\nAPP'].str.contains("consommation d'eau", case=False, na=False)]
    if indicateurs_eau.empty:
        st.warning("⚠️ Données consommation d'eau manquantes")
        return
    
    indicateurs_eau['Nom Ind.N2\nAPP'] = indicateurs_eau['Nom Ind.N2\nAPP'].apply(
        lambda x: x.replace("Consommation d'eau des sièges, agences, bureaux", "Consommation d'eau des sièges,<br>agences, bureaux" )
    )
    st.subheader("Consommation d'eau des sièges, agences, bureaux VS Consommation totale d'eau")

    # Appel direct à bar_comparatif
    fig = bar_comparatif(
        df=indicateurs_eau,
        col_x='Nom Ind.N2\nAPP',
        col_y_n1='Total\nMontant\nCollecte\nRéelle\nExercice N-1',
        col_y_n='Total\nMontant\nCollecte\nRéelle\nExercice N',
        label_x="Indicateurs Eau",
        label_y="Volume (Litres)"
    )

    # Moyennes objectifs
    plafond_strat = np.nanmean(indicateurs_eau['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'])
    plancher_strat = np.nanmean(indicateurs_eau['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'])
    plafond_oper = np.nanmean(indicateurs_eau['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'])
    plancher_oper = np.nanmean(indicateurs_eau['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'])

    options_obj = st.multiselect(
        "🎯 Objectifs à afficher",
        options=["Stratégique - Plafond", "Stratégique - Plancher", "Opérationnel - Plafond", "Opérationnel - Plancher"],
        default=[],
        key="multiselect_objectifs_consommation_eau"
    )

    if "Stratégique - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_strat, "Obj. Strat. plafond moyen", "green", "dash")
    if "Stratégique - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_strat, "Obj. Strat. plancher moyen", "green", "dot")
    if "Opérationnel - Plafond" in options_obj:
        add_horizontal_line(fig, plafond_oper, "Obj. Opé. plafond moyen", "orange", "dash")
    if "Opérationnel - Plancher" in options_obj:
        add_horizontal_line(fig, plancher_oper, "Obj. Opé. plancher moyen", "orange", "dot")

    st.markdown("ℹ️ Vous pouvez sélectionner un ou plusieurs types d’objectifs à afficher sur le graphique.")

    fig.update_layout(
        title="Consommation d'eau des sièges, agences, bureaux VS Consommation totale d'eau",
        height=500
    )

    st.plotly_chart(fig, use_container_width=True)
    return fig

