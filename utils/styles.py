import pandas as pd

def style_kpi(df: pd.DataFrame):
    df = df.copy()

    formatted_values = []
    for idx, row in df.iterrows():
        formatted_row = []
        unite = str(row.get("Unité de conversion de l'indicateur", "")).strip()
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                formatted_row.append("-")
                continue
            if "VARIATION" in col:
                try:
                    formatted_row.append(f"{val:.1f}%")
                except:
                    formatted_row.append(str(val))
                continue
            if "Valorisation Financière" in col:
                try:
                    s = f"{val:,.2f}".replace(",", " ")
                    formatted_row.append(s)
                except:
                    formatted_row.append(str(val))
                continue
            try:
                if unite == "%":
                    formatted_row.append(f"{val:.1f}")
                else:
                    s = f"{val:,.2f}".replace(",", " ")
                    formatted_row.append(s)
            except:
                formatted_row.append(str(val))
        formatted_values.append(formatted_row)

    formatted_df = pd.DataFrame(formatted_values, columns=df.columns, index=df.index)

    style = formatted_df.style.set_properties(**{'text-align': 'right'})

    # Création d'un DataFrame de classes CSS (même taille)
    classes_df = pd.DataFrame("", index=df.index, columns=df.columns)

    # On attribue la classe 'nowrap-col' uniquement aux colonnes ciblées
    for col in ["Reel N", "Reel N-1"]:
        if col in classes_df.columns:
            classes_df[col] = 'nowrap-col'

    style = style.set_td_classes(classes_df)

    # On ajoute les styles CSS
    style = style.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#f2f2f2'), ('white-space', 'pre-wrap'), ('font-size', '12px')]},
        {'selector': 'td', 'props': [('white-space', 'pre-wrap'), ('text-align', 'right'), ('font-size', '12px')]},
        {'selector': 'td.nowrap-col', 'props': [('white-space', 'nowrap')]},  # cible la classe nowrap-col
    ])

    # Couleurs conditionnelles corrigées
    if "Reel N" in df.columns:
        style = style.applymap(lambda v: 'background-color: aliceblue', subset=["Reel N"])
    if "Reel N-1" in df.columns:
        style = style.applymap(lambda v: 'background-color: whitesmoke', subset=["Reel N-1"])
    if "VARIATION Réel N vs Réel N-1 (%)" in df.columns:
        style = style.applymap(lambda v: 'background-color: lemonchiffon', subset=["VARIATION Réel N vs Réel N-1 (%)"])
    if "Valorisation Financière REEL N-1" in df.columns:
        style = style.applymap(lambda v: 'background-color: whitesmoke', subset=["Valorisation Financière REEL N-1"])
    if "Valorisation Financière REEL N" in df.columns:
        style = style.applymap(lambda v: 'background-color: aliceblue', subset=["Valorisation Financière REEL N"])
    
    # if 'Objectifs Opérationnels SEUIL période N' in df.columns:
    #     style = style.applymap(lambda v: 'background-color: lavender', subset=['Objectifs Opérationnels SEUIL période N'])
    if "VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)" in df.columns:
        style = style.applymap(lambda v: 'background-color: lemonchiffon', subset=["VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)"])
    # if 'Objectifs Opérationnels PLAFOND période N' in df.columns:
    #     style = style.applymap(lambda v: 'background-color: lavender', subset=['Objectifs Opérationnels PLAFOND période N'])
    if "VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)" in df.columns:
        style = style.applymap(lambda v: 'background-color: lemonchiffon', subset=["VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)"])
    
    # if 'Objectifs Stratégiques SEUIL période N' in df.columns:
        # style = style.applymap(lambda v: 'background-color: lavender', subset=['Objectifs Stratégiques SEUIL période N'])
    if "VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)" in df.columns:
        style = style.applymap(lambda v: 'background-color: lemonchiffon', subset=["VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)"])
    # if 'Objectifs Stratégiques PLAFOND période N' in df.columns:
        # style = style.applymap(lambda v: 'background-color: aliceblue', subset=['Objectifs Stratégiques PLAFOND période N'])
    if "VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)" in df.columns:
        style = style.applymap(lambda v: 'background-color: lemonchiffon', subset=["VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)"])

    return style
