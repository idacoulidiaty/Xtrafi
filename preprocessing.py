import pandas as pd
import re

def filter_and_rename_columns(df):
    df = df[[
        'Axe',
        # 'Code\nRAPPORT Ind.\nVIRTUEL',
        # 'Nom\nRAPPORT Ind.\nVIRTUEL',
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
    ]]

    df = df.rename(columns={
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

    return df



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
