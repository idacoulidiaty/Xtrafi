import streamlit as st
import plotly.graph_objects as go
import pandas as pd
from utils.styles import style_objectifs  
from preprocessing import filter_and_rename_columns, compute_variations


import re

def safe_key(name):
    return re.sub(r'\W+', '_', name)

def format_montant(valeur, monnaie):
    return f"{int(valeur):,}".replace(",", " ") + f" {monnaie}"

def custom_metric(label, value, fontsize=20):
    st.markdown(f"""
        <div style="padding: 0.5em 1em; border: 1px solid #ccc; border-radius: 8px; background-color: #f9f9f9;">
            <div style="font-size: 0.9em; color: #888;">{label}</div>
            <div style="font-size: {fontsize}px; font-weight: bold;">{value}</div>
        </div>
    """, unsafe_allow_html=True)


def generate_graphs_par_indicateur_en_colonnes(
    df,
    col_indicateur,
    col_val_n1,
    col_val_n,
    colonnes_objectifs=None,
    n_cols=2
):
    import streamlit as st
    import pandas as pd
    import plotly.graph_objects as go
    from utils.styles import style_objectifs

    noms_objectifs_lisibles = {
        'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N': 'Obj. Strat. PLAFOND N',
        'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N': 'Obj. Strat. SEUIL N',
        'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N': 'Obj. Opé. PLAFOND N',
        'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N': 'Obj. Opé. SEUIL N'
    }

    indicateurs = df[col_indicateur].dropna().unique()
    cols = st.columns(n_cols)
    figs = []
    fig_infos = []

    for i, indicateur in enumerate(indicateurs):
        df_indic = df[df[col_indicateur] == indicateur]
        df_indic = compute_variations(df_indic, col_reel_n=col_val_n, col_reel_n1=col_val_n1)

        val_n1 = df_indic[col_val_n1].values[0] if col_val_n1 in df_indic else None
        val_n = df_indic[col_val_n].values[0] if col_val_n in df_indic else None

        unite = df_indic["Unité\nconversion"].values[0] if "Unité\nconversion" in df_indic else ""
        monnaie = df_indic["Unité\nValorisation\nfinancière"].values[0] if "Unité\nValorisation\nfinancière" in df_indic else ""
        valo_fi_n1 = df_indic['Total\nValo. Financière\nCollecte Réelle\nExercice N-1'].values[0] if 'Total\nValo. Financière\nCollecte Réelle\nExercice N-1' in df_indic else None
        valo_fi_n = df_indic['Total\nValo. Financière\nCollecte Réelle\nExercice N'].values[0] if 'Total\nValo. Financière\nCollecte Réelle\nExercice N' in df_indic else None

        objectifs_disponibles = []
        for col in (colonnes_objectifs or []):
            if col in df.columns:
                serie = df_indic[col]
                if not serie.dropna().empty:
                    objectifs_disponibles.append(col)

        col_ui = cols[i % n_cols]


        labels_disponibles = [noms_objectifs_lisibles.get(col, col) for col in objectifs_disponibles]

        selected_labels = col_ui.multiselect(
            f"🎯 Afficher les objectifs pour « {indicateur} »",
            options=labels_disponibles
        )


        selected_objectifs = [
            col for col, label in noms_objectifs_lisibles.items()
            if label in selected_labels
        ]

        with col_ui.expander(f"📊 **{indicateur} (en {unite})**", expanded=True):
            fig = go.Figure()
            fig.add_bar(x=[0], y=[val_n1], name="Exercice N-1", marker_color="palegreen", width=0.07, hovertemplate=f"Exercice N-1 : {float(val_n1):,.2f} {unite}<extra></extra>".replace(",", " "))
            fig.add_bar(x=[0.10], y=[val_n], name="Exercice N", marker_color="lightblue", width=0.07, hovertemplate=f"Exercice N : {float(val_n):,.2f} {unite}<extra></extra>".replace(",", " "))
            fig.update_layout(
                title=f'''{indicateur}''',
                barmode="group",
                bargap=0.01,
                height=500,
                xaxis=dict(
                    tickmode="array",
                    tickvals=[0, 0.10],
                    ticktext=["N-1", "N"],
                    showticklabels=True,
                    range=[-0.1, 0.4]
                )
            )

            with col_ui.expander(f"📊 **{indicateur} (en {unite})**", expanded=True):
                fig = go.Figure()
                fig.add_bar(
                    x=[0], y=[val_n1],
                    name="Exercice N-1",
                    marker_color="palegreen",
                    width=0.07,
                    hovertemplate=f"Exercice N-1 : {float(val_n1):,.2f} {unite}<extra></extra>".replace(",", " ")
                )
                fig.add_bar(
                    x=[0.10], y=[val_n],
                    name="Exercice N",
                    marker_color="lightblue",
                    width=0.07,
                    hovertemplate=f"Exercice N : {float(val_n):,.2f} {unite}<extra></extra>".replace(",", " ")
                )
                fig.update_layout(
                    title=f"{indicateur}",
                    barmode="group",
                    bargap=0.01,
                    height=500,
                    xaxis=dict(
                        tickmode="array",
                        tickvals=[0, 0.10],
                        ticktext=["N-1", "N"],
                        showticklabels=True,
                        range=[-0.1, 0.4]
                    )
                )

            for obj_col in selected_objectifs:
                val_obj = df_indic[obj_col].values[0]
                if pd.notna(val_obj):
                    style = style_objectifs.get(obj_col, {"line_color": "red", "line_dash": "dot"})
                    fig.add_hline(
                        y=val_obj,
                        line_dash=style["line_dash"],
                        line_color=style["line_color"],
                        annotation_text=f"{noms_objectifs_lisibles.get(obj_col, obj_col)} = {val_obj:,.2f}",
                        annotation_position="top right"
                    )

            text_lines = []

            if pd.notna(valo_fi_n1):
                text_lines.append(f"Valo. N-1 : {int(valo_fi_n1):,} {monnaie}".replace(",", " "))
            if pd.notna(valo_fi_n):
                text_lines.append(f"Valo. N : {int(valo_fi_n):,} {monnaie}".replace(",", " "))

            for obj_col in selected_objectifs:
                var_colname = f"VARIATION {obj_col} vs Réel N (%)"
                if var_colname in df_indic.columns:
                    variation = df_indic[var_colname].values[0]
                    val_obj = df_indic[obj_col].values[0]  
                    if pd.notna(variation) and pd.notna(val_obj):
                        if "Plancher" in obj_col:
                            couleur = "🔴" if val_n < val_obj else "🟢"
                        elif "Plafond" in obj_col:
                            couleur = "🔴" if val_n > val_obj else "🟢"
                        else:
                            couleur = ""  
                        label = noms_objectifs_lisibles.get(obj_col, obj_col)
                        text_lines.append(f"{label} : {variation:.2f}% {couleur}")


            if text_lines:
                fig.add_annotation(
                    text="<br>".join(text_lines),
                    xref="paper", yref="paper",
                    x=1.15, y=0.4,
                    showarrow=False,
                    align="left",
                    font=dict(size=13, color="black"),
                    bgcolor="#f5f5f5",
                    bordercolor="gray",
                    borderwidth=1,
                    borderpad=6
                )

            st.plotly_chart(fig, use_container_width=True, key=f"plot_{safe_key(indicateur)}")

        figs.append(fig)
        fig_infos.append({
            "indicateur": indicateur,
            "val_n1": val_n1,
            "val_n": val_n,
            "objectifs": {
                obj: df_indic[obj].values[0]
                for obj in selected_objectifs
                if obj in df_indic.columns and pd.notna(df_indic[obj].values[0])
            }
        })

    return figs, fig_infos



