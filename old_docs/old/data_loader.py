import pandas as pd

def load_all_data(file_path):
    """
    Charge les 3 dataframes principales depuis le fichier Excel.
    """
    df1 = pd.read_excel(file_path, sheet_name="Tableau de Bord REEL N vs REEL N-1")
    df2 = pd.read_excel(file_path, sheet_name="Tableau de Bord REEL N vs Objectifs Opérationnels N")
    df3 = pd.read_excel(file_path, sheet_name="Tableau de Bord REEL N vs Objectifs Stratégiques N")
    return df1, df2, df3
