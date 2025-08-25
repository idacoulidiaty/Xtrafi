from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows
import plotly.graph_objects as go
import plotly.io as pio

# Nécessaire pour fig.write_image
pio.kaleido.scope.default_format = "png"

def export_excel_with_figures(df_list, fig_list, fig_glob=None):
    output = BytesIO()
    wb = Workbook()
    del wb["Sheet"]  # Supprimer la feuille par défaut

    # === 1. EXPORT TABLEAUX AVEC STYLE STREAMLIT ===
    for df, sheet_name in df_list:
        ws = wb.create_sheet(title=sheet_name[:31])

        # Ajouter les lignes du DataFrame
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            ws.append(row)

        # Style de l'en-tête
        header_fill = PatternFill(start_color="1f77b4", end_color="1f77b4", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", name="Calibri")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Bordures et alignement des cellules
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

        # Auto-ajustement des colonnes
        for column_cells in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = max_length + 2

    # === 2. INSERTION DES GRAPHIQUES EN DEUX COLONNES ===
    if fig_list:
        ws_graph = wb.create_sheet(title="Graphiques")
        col_positions = ["B", "M"]  # Colonnes pour affichage double
        x_index = 0
        y_cursor = [2, 2]  # Position verticale pour chaque colonne

        for fig in fig_list:
            if fig is None or not isinstance(fig, go.Figure):
                continue  # Skip les figures invalides

            # Titre et layout : assure-toi que ça a été défini en amont lors de la génération
            fig.update_layout(
                margin=dict(t=60, b=40),
                title=dict(x=0.5, font=dict(size=14)),
                showlegend=True
            )

            # Export en image
            try:
                img_buffer = BytesIO()
                fig.write_image(img_buffer, format="png")
                img_buffer.seek(0)
                img = XLImage(img_buffer)
            except Exception as e:
                print(f"Erreur d'export image : {e}")
                continue

            # Placement dans feuille Excel
            col_letter = col_positions[x_index % 2]
            row_number = y_cursor[x_index % 2]
            cell_location = f"{col_letter}{row_number}"

            ws_graph.add_image(img, cell_location)
            y_cursor[x_index % 2] += 25  # Espace vertical entre les graphs
            x_index += 1

    # === 3. INSERTION DU GRAPHIQUE fig_glob DANS UN ONGLET DISTINCT ===
    if fig_glob and isinstance(fig_glob, go.Figure):
        ws_glob = wb.create_sheet(title="Graphique Variations")

        try:
            # Récupérer tous les labels y présents dans la figure (robuste à plusieurs traces)
            y_labels = []
            for tr in fig_glob.data:
                if hasattr(tr, "y") and tr.y is not None:
                    # tr.y peut être tuple/list/np.array
                    y_labels.extend([str(v) for v in tr.y])

            # préserver l'ordre et garder uniques
            if y_labels:
                seen = set()
                labels = [x for x in y_labels if not (x in seen or seen.add(x))]
            else:
                labels = []

            n_bars = len(labels)

            # calculer marge gauche selon la longueur max d'étiquette
            max_label_len = max((len(l) for l in labels), default=0)
            left_margin = int(max(120, max_label_len * 7))  # ajuster le facteur si nécessaire

            # hauteur adaptée au nombre de barres
            height = max(600, 30 * n_bars + 200)  # 30 px par barre + marge
            # largeur = marge gauche + espace utile pour barres
            width = max(1200, left_margin + 1000)

            # diminuer la taille des ticks si trop long
            tick_font_size = 11
            if max_label_len > 40:
                tick_font_size = 9
            if max_label_len > 80:
                tick_font_size = 8

            # Appliquer marges / automargin pour s'assurer que tout rentre
            fig_glob.update_layout(
                margin=dict(l=left_margin, r=50, t=60, b=50),
                height=height,
                width=width
            )
            fig_glob.update_yaxes(automargin=True, tickfont=dict(size=tick_font_size))

            # Export image haute résolution
            img_buffer = BytesIO()
            fig_glob.write_image(img_buffer, format="png", width=width, height=height, scale=2)
            img_buffer.seek(0)
            img = XLImage(img_buffer)
            ws_glob.add_image(img, "A1")

        except Exception as e:
            print(f"Erreur d'export image fig_glob : {e}")


    # Finalisation
    wb.save(output)
    output.seek(0)
    return output.getvalue()
