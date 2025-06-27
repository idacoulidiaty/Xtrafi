import streamlit as st
from data_loader import load_all_data
from visualization import get_fig1, get_fig2, get_fig4, get_fig5, get_fig6
from export_utils import export_to_excel

def main():
    st.title("Dashboard Xtrafi")

    # Chargement des données
    df1, df2, df3 = load_all_data("data/*.xls*")

    # Création des graphiques
    fig1 = get_fig1(df1)
    fig2 = get_fig2(df2)
    fig4 = get_fig4(df2)
    fig5 = get_fig5(df3)
    fig6 = get_fig6(df3)

    # Affichage dans Streamlit (tableaux, graphiques, etc.) ici

    st.subheader("⬇️ Exporter les données et visualisations")
    if st.button("📤 Exporter en Excel"):
        excel_data = export_to_excel(df1, df2, df3, fig1, fig2, fig4, fig5, fig6)
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=excel_data,
            file_name="xtrafi_export_avec_graphiques.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("💡 Téléversez un fichier ou déposez-en un dans le dossier `data/` pour démarrer.")

if __name__ == "__main__":
    main()
