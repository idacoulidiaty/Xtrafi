import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd

# 🎨 Configuration graphique commune
# GRAPH_CONFIG = {
#     "bar_width": 0.3,
#     "height": 500,
#     "template": "plotly_white",
#     "bargap": 0.2,          # espace entre les groupes de barres
#     "margin": dict(l=40, r=40, t=40, b=40),  # marges du graphique
# }
def add_horizontal_line(fig, y_value, label, color, dash, x_start=-0.5, x_end=0.5):
    if pd.notna(y_value):
        fig.add_shape(
            type="line",
            x0=x_start, x1=x_end,
            y0=y_value, y1=y_value,
            line=dict(color=color, width=2, dash=dash),
            xref='x', yref='y'
        )
        fig.add_annotation(
            x=x_end,
            y=y_value,
            text=label,
            showarrow=False,
            yanchor="bottom",
            font=dict(color=color)
        )

def uniformiser_largeur_barres(df, col_x, min_nb_categories=6):
    if len(df) < min_nb_categories:
        n_to_add = min_nb_categories - len(df)
        df_padding = pd.DataFrame({
            col_x: [f""] * n_to_add,  # catégories vides
            # Ajouter les colonnes nécessaires avec 0
            'Total\nMontant\nCollecte\nRéelle\nExercice N-1': [0]*n_to_add,
            'Total\nMontant\nCollecte\nRéelle\nExercice N': [0]*n_to_add
        })
        df = pd.concat([df, df_padding], ignore_index=True)
    return df


def bar_comparatif(df, col_x, col_y_n1, col_y_n, label_x="Indicateur", label_y="Valeur"):
    df = uniformiser_largeur_barres(df, col_x)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Exercice N-1",
        x=df[col_x],
        y=df[col_y_n1],
        marker_color="lightblue",
        width=0.3
    ))
    fig.add_trace(go.Bar(
        name="Exercice N",
        x=df[col_x],
        y=df[col_y_n],
        marker_color="midnightblue",
        width=0.3
    ))
    fig.update_layout(
        barmode="group",
        xaxis_title=label_x,
        yaxis_title=label_y,
        template='plotly_white',
        height=500,
        bargap= 0.3,
        bargroupgap = 0.8
    )
    return fig


def export_plot_button(fig, filename="graphique.png", label="📥 Exporter"):
    if st.button(label):
        fig.write_image(filename)
        st.success(f"{filename} exporté avec succès.")
