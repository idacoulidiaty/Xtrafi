



def generate_graphs_par_indicateur_en_colonnes(df, col_indicateur, col_val_n1, col_val_n, n_cols=2):
    indicateurs = df[col_indicateur].dropna().unique()
    cols = st.columns(n_cols)
    for i, indicateur in enumerate(indicateurs):
        df_indic = df[df[col_indicateur] == indicateur]
        fig = bar_comparatif(
            df_indic,
            col_x=col_indicateur,
            col_y_n1=col_val_n1,
            col_y_n=col_val_n,
            label_x="Indicateur",
            label_y="Valeur"
        )
        # Affiche dans la colonne i % n_cols
        cols[i % n_cols].plotly_chart(fig, use_container_width=True)

        # S'il y a plusieurs lignes pour un indicateur, elles sont toutes affichées
