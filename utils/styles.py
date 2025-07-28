import pandas as pd
import streamlit as st

def style_kpi(df: pd.DataFrame):
    df = df.copy()

    # Formatage des valeurs
    formatted_values = []
    for idx, row in df.iterrows():
        formatted_row = []
        unite = str(row.get("Unité de conversion de l'indicateur", "")).strip()
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                formatted_row.append("-")
                continue
            try:
                if "VARIATION" in col:
                    formatted_row.append(f"{val:.1f}%")
                elif "Valorisation Financière" in col:
                    s = f"{val:,.2f}".replace(",", " ")
                    formatted_row.append(s)
                elif unite == "%":
                    formatted_row.append(f"{val:.1f}")
                else:
                    s = f"{val:,.2f}".replace(",", " ")
                    formatted_row.append(s)
            except:
                formatted_row.append(str(val))
        formatted_values.append(formatted_row)

    formatted_df = pd.DataFrame(formatted_values, columns=df.columns, index=df.index)

    # Styling de base
    style = formatted_df.style.set_properties(**{'text-align': 'right'})

    # Classe CSS "nowrap-col"
    classes_df = pd.DataFrame("", index=df.index, columns=df.columns)
    for col in ["Reel N", "Reel N-1"]:
        if col in df.columns:
            classes_df[col] = 'nowrap-col'
    style = style.set_td_classes(classes_df)

    # Table styles généraux
    style = style.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#f2f2f2'), ('white-space', 'pre-wrap'), ('font-size', '12px')]},
        {'selector': 'td', 'props': [('white-space', 'pre-wrap'), ('text-align', 'right'), ('font-size', '12px')]},
        {'selector': 'td.nowrap-col', 'props': [('white-space', 'nowrap')]},
    ])

    # ✅ Mapping des couleurs conditionnelles
    col_color_map = {
        "Reel N": "aliceblue",
        "Reel N-1": "whitesmoke",
        "VARIATION Réel N vs Réel N-1 (%)": "lemonchiffon",
        "Valorisation Financière REEL N-1": "whitesmoke",
        "Valorisation Financière REEL N": "aliceblue",
        "VARIATION Objectifs Opérationnels SEUIL N vs Réel N (%)": "lemonchiffon",
        "VARIATION Objectifs Opérationnels PLAFOND N vs Réel N (%)": "lemonchiffon",
        "VARIATION Objectifs Stratégiques SEUIL N vs Réel N (%)": "lemonchiffon",
        "VARIATION Objectifs Stratégiques PLAFOND N vs Réel N (%)": "lemonchiffon"
    }

    # 🧠 Utilisation  de .map() colonne par colonne
    for col, color in col_color_map.items():
        if col in df.columns:
            style = style.map(lambda col: f"background-color: {color}", subset=[col])

    return style


def load_css(file_path):
    with open(file_path) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)




style_objectifs = {
    'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N': {
        "line_color": "orange", "line_dash": "dot"
    },
    'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N': {
        "line_color": "orange", "line_dash": "dot"
    },
    'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N': {
        "line_color": "green", "line_dash": "dash"
    },
    'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N': {
        "line_color": "green", "line_dash": "dash"
    },
}

