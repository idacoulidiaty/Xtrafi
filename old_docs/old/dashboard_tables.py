import pandas as pd

def generate_dashboard_tables(df):
    """
    Crée trois tableaux de bord :
    - df1 : Comparaison Réel N vs N-1
    - df2 : Comparaison Réel N vs Objectif Opérationnel N
    - df3 : Comparaison Réel N vs Objectif Stratégique N
    """
    colonnes_selection = [
        'Nom Ind.N2\nAPP',
        'Total\nMontant\nCollecte\nRéelle\nExercice N-1',
        'Total\nMontant\nCollecte\nRéelle\nExercice N',
        'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N',
        'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N',
        'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N',
        'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'
    ]

    df_filtered = df[colonnes_selection].copy()

    df1 = df_filtered[[
        'Nom Ind.N2\nAPP',
        'Total\nMontant\nCollecte\nRéelle\nExercice N-1',
        'Total\nMontant\nCollecte\nRéelle\nExercice N'
    ]].rename(columns={
        'Total\nMontant\nCollecte\nRéelle\nExercice N-1': 'Réel N-1',
        'Total\nMontant\nCollecte\nRéelle\nExercice N': 'Réel N'
    })

    df2 = df_filtered[[
        'Nom Ind.N2\nAPP',
        'Total\nMontant\nCollecte\nRéelle\nExercice N',
        'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N',
        'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N'
    ]].rename(columns={
        'Total\nMontant\nCollecte\nRéelle\nExercice N': 'Réel N',
        'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N': 'Obj. Opér. Plancher',
        'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N': 'Obj. Opér. Plafond'
    })

    df3 = df_filtered[[
        'Nom Ind.N2\nAPP',
        'Total\nMontant\nCollecte\nRéelle\nExercice N',
        'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N',
        'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N'
    ]].rename(columns={
        'Total\nMontant\nCollecte\nRéelle\nExercice N': 'Réel N',
        'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N': 'Obj. Strat. Plancher',
        'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N': 'Obj. Strat. Plafond'
    })

    return df1, df2, df3



