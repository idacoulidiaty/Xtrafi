import pandas as pd
import re
import streamlit as st

def filter_and_rename_columns(df):
    required_columns = [
        'Axe',
        'Code\nREPORTING\nInd. VIRTUEL',
        'Nom\nREPORTING Ind.\nVIRTUEL',
        'Code\nAPP Indicateur\nVIRTUEL',
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
        'Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N'
    ]

    try:
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(
                f"❌ Le fichier est obsolète : il manque les colonnes suivantes :\n- " +
                "\n- ".join(missing_cols) +
                "\n\nVeuillez importer un fichier à jour généré par Xtrafi App."
                )           
            st.stop()

        df = df[required_columns]

        df = df.rename(columns={
            'Code\nREPORTING\nInd. VIRTUEL': 'Code Reporting',
            'Nom\nREPORTING Ind.\nVIRTUEL': 'Nom Reporting',
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
        })

        return df

    except Exception as e:
        st.error(f"❌ Une erreur est survenue lors du filtrage et renommage des colonnes : {str(e)}")
        st.stop()


def rename_columns(df):
 
    df = df.rename(columns={
            
    })

    return df


import re
import pandas as pd

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




import pandas as pd
import numpy as np
import plotly.express as px

def plot_variations_detaillees(df):
    # Nettoyer la casse de la colonne Axe
    df['Axe'] = df['Axe'].str.upper()

    axes_order = ['EN', 'SO', 'GO']

    # Concaténer Axe et nom indicateur
    df['axe_indic'] = df['Axe'] + ' - ' + df['Nom indicateur Virtuel']

    order_list = []
    for axe in axes_order:
        indicateurs_dans_axe = df[df['Axe'] == axe]['Nom indicateur Virtuel'].unique()
        order_list.extend([f"{axe} - {ind}" for ind in indicateurs_dans_axe])

    df['axe_indic'] = pd.Categorical(df['axe_indic'], categories=order_list, ordered=True)

    # Couleurs vert / rouge selon signe de la variation
    col_pos = '#2ca02c'
    col_neg = '#d62728'
    df['color'] = np.where(df['VARIATION Réel N vs Réel N-1 (%)'] >= 0, col_pos, col_neg)

    fig = px.bar(
        df,
        x='VARIATION Réel N vs Réel N-1 (%)',
        y='axe_indic',
        orientation='h',
        color='color',
        color_discrete_map={col_pos: col_pos, col_neg: col_neg},
        title="Variations Réel N vs Réel N-1 par indicateur et axe"
    )

    fig.update_layout(
        xaxis_title="Variation (%)",
        yaxis_title="Indicateur (groupé par Axe)",
        yaxis_categoryorder='array',
        yaxis_categoryarray=order_list,
        height=600,
        width=900,
        showlegend=False,
        bargap=0.2
    )

    return fig
