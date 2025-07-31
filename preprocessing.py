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
