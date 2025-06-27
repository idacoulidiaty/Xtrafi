import streamlit as st
from io import BytesIO
from openpyxl.drawing.image import Image as XLImage
import numpy as np
import pandas as pd


def style_dataframe(df):
    return df.style.set_properties(**{
        'background-color': '#f0f8ff',
        'color': 'black',
        'border-color': 'gray',
        'border-width': '1px',
        'border-style': 'solid'
    }).set_table_styles([{
        'selector': 'th',
        'props': [
            ('background-color', '#1f4e79'),
            ('color', 'white'),
            ('font-weight', 'bold')
        ]
    }])


def export_to_excel(df1, df2, df3, figures):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Export tableaux avec style
        style_dataframe(df1).to_excel(writer, sheet_name="Tableau de Bord REEL N vs REEL N-1", index=False)
        style_dataframe(df2).to_excel(writer, sheet_name="Tableau de Bord REEL N vs Objectifs Opérationnels N", index=False)
        style_dataframe(df3).to_excel(writer, sheet_name="Tableau de Bord REEL N vs Objectifs Stratégiques N", index=False)

        # Onglet Graphiques
        workbook = writer.book
        sheet = workbook.create_sheet(title='Graphiques')

        # Positions d'insertion pour chaque figure
        positions = ['B2', 'M2', 'B16', 'M31', 'M48', 'M20']

        for fig, position in zip(figures, positions):
            if fig:
                buffer = BytesIO()
                fig.write_image(buffer, format="png")
                buffer.seek(0)
                img = XLImage(buffer)
                sheet.add_image(img, position)

    return output


def render_export_section(df1, df2, df3, figures):
    st.subheader("⬇️ Exporter les données et visualisations")

    if st.button("📤 Exporter en Excel"):
        output = export_to_excel(df1, df2, df3, figures)

        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=output.getvalue(),
            file_name="xtrafi_export_avec_graphiques.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("💡 Téléversez un fichier ou déposez-en un dans le dossier `data/` pour démarrer.")
