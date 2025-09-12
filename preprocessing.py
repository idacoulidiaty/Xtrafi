import pandas as pd
import streamlit as st
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

import pandas as pd
import streamlit as st

import streamlit as st
import pandas as pd

def filter_and_rename_columns(df):
    if df is None or df.empty:
        return pd.DataFrame(), ["⚠️ Le fichier importé est vide ou n'a pas pu être lu correctement."]

    # Map des renommages
    rename_map = {
        'Code\nREPORTING\nInd. VIRTUEL': 'Code Reporting',
        'Nom\nREPORTING Ind.\nVIRTUEL': 'Nom Reporting',
        'Code\nRAPPORT Ind.\nVIRTUEL': 'Code Rapport Indicateur Virtuel',
        'Nom\nRAPPORT Ind.\nVIRTUEL': 'Nom Rapport Indicateur Virtuel',
        'Code\nAPP Indicateur\nVIRTUEL': 'Code indicateur Virtuel',
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
    }

    # Renommer uniquement si la colonne existe
    rename_map_existing = {k: v for k, v in rename_map.items() if k in df.columns}
    df = df.rename(columns=rename_map_existing)

    # Colonnes attendues après renommage
    required_columns = list(rename_map.values()) + ['Axe', 'Unité Valorisation financière']

    # Colonnes manquantes
    missing_cols = [col for col in required_columns if col not in df.columns]

    # Ne conserver que les colonnes existantes
    available_cols = [col for col in required_columns if col in df.columns]
    df = df[available_cols] if available_cols else pd.DataFrame()

    return df, missing_cols



def rename_columns(df):
 
    df = df.rename(columns={
            
    })

    return df



def compute_ratio(numerateur, denominateur):
    """Calcule la variation (%) avec gestion des divisions par zéro et NaN"""
    result = (numerateur - denominateur) / denominateur
    return result.replace([float("inf"), -float("inf")], pd.NA) * 100


def compute_variations(df, col_reel_n, col_reel_n1):
    """
    Calcule les variations en % entre Réel N et Réel N-1, et entre Réel N et tous les objectifs disponibles.

    Args:
        df (pd.DataFrame) : Données source.
        col_reel_n (str) : Nom de la colonne contenant les valeurs Réel N.
        col_reel_n1 (str) : Nom de la colonne contenant les valeurs Réel N-1.

    Returns:
        pd.DataFrame enrichi des colonnes de variation.
    """
    if col_reel_n not in df.columns or col_reel_n1 not in df.columns:
        return df

    df["VARIATION Réel N vs Réel N-1 (%)"] = compute_ratio(df[col_reel_n], df[col_reel_n1])

    # Heuristique pour repérer les colonnes d’objectifs pertinentes
    objectif_cols = [
        col for col in df.columns
        if any(kw in col for kw in ['O.Strat', 'O.Opér', 'Objectifs Stratégiques', 'Objectifs Opérationnels']) 
        and "N" in col  # on cible bien Exercice N
    ]

    for col in objectif_cols:
        if col in df.columns:
            var_colname = f"VARIATION {col} vs Réel N (%)"
            df[var_colname] = compute_ratio(df[col_reel_n], df[col])

    return df




def plot_variations_detaillees(df):
    # Vérification colonne Axe
    if 'Axe' in df.columns:
        df['Axe'] = df['Axe'].astype(str).str.upper()
    else:
        st.warning("⚠️ La colonne 'Axe' est absente : certaines visualisations globales peuvent ne pas s'afficher correctement.")
        df['Axe'] = "N/A"

    # Vérification colonne Nom indicateur Virtuel
    if 'Nom indicateur Virtuel' in df.columns:
        df['axe_indic'] = df['Axe'] + ' - ' + df['Nom indicateur Virtuel'].astype(str)
    else:
        st.warning("⚠️ La colonne 'Nom indicateur Virtuel' est absente : regroupement par indicateur désactivé.")
        df['axe_indic'] = df['Axe']

    # Ordre des axes
    axes_order = ['EN', 'SO', 'GO']
    order_list = []
    if 'Nom indicateur Virtuel' in df.columns:
        for axe in axes_order:
            indicateurs_dans_axe = df[df['Axe'] == axe]['Nom indicateur Virtuel'].unique()
            order_list.extend([f"{axe} - {ind}" for ind in indicateurs_dans_axe if pd.notna(ind)])
        df['axe_indic'] = pd.Categorical(df['axe_indic'], categories=order_list, ordered=True)

    # Déterminer couleur selon variation + objectifs
    def color_from_objectifs(row):
        plafond = row.get('Objectifs Opérationnels PLAFOND période N', pd.NA)
        plancher = row.get('Objectifs Opérationnels SEUIL période N', pd.NA)
        val = row.get('Reel N', pd.NA)
        variation = row.get('VARIATION Réel N vs Réel N-1 (%)', pd.NA)

        # Si info manquante → gris neutre
        if pd.isna(plancher) or pd.isna(plafond) or pd.isna(val) or pd.isna(variation):
            return '#cccccc'

        if variation < 0:
            return '#ffcc00' if plancher <= val <= plafond else '#d62728'
        if variation > 0:
            return '#2ca02c' if plancher <= val <= plafond else '#ff7f0e'

        return '#cccccc'

    df['color'] = df.apply(color_from_objectifs, axis=1)

    # Vérification colonne variation
    variation_col = 'VARIATION Réel N vs Réel N-1 (%)'
    if variation_col not in df.columns:
        st.error(f"⚠️ La colonne '{variation_col}' est absente : impossible d'afficher la visualisation des variations.")
        return go.Figure()

    # Vérification colonne axe_indic
    if 'axe_indic' not in df.columns:
        st.error("⚠️ Impossible de construire la variable axe_indic (Axe + Nom indicateur Virtuel).")
        return go.Figure()

    # Barres principales
    fig = px.bar(
        df,
        x=variation_col,
        y='axe_indic',
        orientation='h',
        color='color',
        color_discrete_map='identity',
        title="Variations Réel N vs Réel N-1 par indicateur et axe"
    )

    # Ajouter traces factices pour légende
    legend_colors = {
        "Décroissance du Réel entre N-1 et N respectant les objectifs": "#ffcc00",
        "Croissance du Réel entre N-1 et N respectant les objectifs": "#2ca02c",
        "Décroissance du Réel entre N-1 et N ne respectant pas les objectifs": "#d62728",
        "Croissance du Réel entre N-1 et N ne respectant pas les objectifs": "#ff7f0e",
        "Objectifs opérationnels non renseignés": "#cccccc"
    }
    for label, color in legend_colors.items():
        fig.add_trace(
            go.Bar(
                x=[0], y=[None],
                marker_color=color,
                name=label,
                showlegend=True
            )
        )

    # Mise en page
    fig.update_layout(
        xaxis=dict(
            title="Variation (%)",
            showgrid=True,
            gridcolor="lightgrey",
            zeroline=True,
            zerolinecolor="black",
            zerolinewidth=2
        ),
        yaxis=dict(
            title="Indicateur (groupé par Axe)",
            showgrid=True,
            gridcolor="lightgrey"
        ),
        yaxis_categoryorder='array',
        yaxis_categoryarray=order_list,
        height=600,
        width=900,
        bargap=0.2
    )

    return fig
