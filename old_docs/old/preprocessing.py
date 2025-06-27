import pandas as pd
import numpy as np


def filter_dataframe(df, column, values):
    """Filtre un DataFrame selon une liste de valeurs d'une colonne."""
    return df[df[column].isin(values)]


def extract_objectif_lists(df):
    """Extrait les valeurs des différents types de seuils objectifs d'un DataFrame."""
    return (
        df['Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'].tolist(),
        df['Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N'].tolist(),
        df['Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'].tolist(),
        df['Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N'].tolist()
    )


def get_valeurs_reelles(row):
    """Retourne les valeurs réelles pour N-1 et N d'une ligne."""
    return (
        row.get('Total\nMontant\nCollecte\nRéelle\nExercice N-1', 0),
        row.get('Total\nMontant\nCollecte\nRéelle\nExercice N', 0)
    )


def calculate_means(*lists):
    """Calcule la moyenne de plusieurs listes (avec gestion des NaN)."""
    return [np.nanmean(l) if l else None for l in lists]
