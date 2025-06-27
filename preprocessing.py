import pandas as pd

def filter_and_rename_columns(df):
    df = df[[
        'Code\nRAPPORT Ind.\nVIRTUEL',
        'Nom\nRAPPORT Ind.\nVIRTUEL',
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

def compute_variations(df):
    df["VARIATION Réel N vs Réel N-1 (%)"] = (
        (df["Reel N"] - df["Reel N-1"]) / df["Reel N-1"]
    ) * 100

    df["VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)"] = (
        (df["Reel N"] - df["Objectifs Stratégiques PLAFOND période N"]) / df["Objectifs Stratégiques PLAFOND période N"]
    ) * 100

    df["VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)"] = (
        (df["Reel N"] - df["Objectifs Stratégiques SEUIL période N"]) / df["Objectifs Stratégiques SEUIL période N"]
    ) * 100

    df["VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)"] = (
        (df["Reel N"] - df["Objectifs Opérationnels PLAFOND période N"]) / df["Objectifs Opérationnels PLAFOND période N"]
    ) * 100

    df["VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)"] = (
        (df["Reel N"] - df["Objectifs Opérationnels SEUIL période N"]) / df["Objectifs Opérationnels SEUIL période N"]
    ) * 100

    return df
