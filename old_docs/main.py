import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage

# ---------- CONFIG APP---------- #
WATCHED_FOLDER = "data"
st.set_page_config(page_title="Xtrafi Data Viz", layout="wide")


# ---------- LOGO ---------- #
st.image("static/logo-xtrafi.png", width=200)
st.title("📊 Votre Dashboard")


# ---------- FONCTIONS TRANSFORMATION DONNEES---------- #
def load_excel_data_dynamic_start(file_path):
    preview_df = pd.read_excel(file_path, header=None, engine="calamine")
    start_row_idx = preview_df[0].astype(str).str.contains(
        "Liste des données qui seront utilisées pour", case=False, na=False
    )
    if start_row_idx.any():
        start_index = start_row_idx[start_row_idx].index[0] + 1
        return pd.read_excel(file_path, skiprows=start_index, engine="calamine")
    else:
        raise ValueError("❌ Ligne de début de données introuvable dans le fichier.")

def get_date_generation_restitution(file_path):
    preview_df = pd.read_excel(file_path, header=None, engine="calamine")
    mask = preview_df.astype(str).apply(lambda row: row.str.contains("Date génération", case=False, na=False)).any(axis=1)
    if mask.any():
        row_index = mask[mask].index[0]
        col_index = preview_df.iloc[row_index].astype(str).str.contains("Date génération", case=False, na=False)
        col_idx = col_index[col_index].index[0]
        label = preview_df.iloc[row_index, col_idx]
        date_value = preview_df.iloc[row_index + 1, col_idx]
        return label, date_value
    else:
        raise ValueError("❌ 'Date génération restitution' introuvable.")

def get_latest_excel_file(folder):
    files = [f for f in os.listdir(folder) if f.endswith((".xlsx", ".xls"))]
    if not files:
        return None
    files.sort(key=lambda x: os.path.getctime(os.path.join(folder, x)), reverse=True)
    return os.path.join(folder, files[0])


# ---------- SIDEBAR : Upload ou sélection ---------- #
st.sidebar.header("📁 Importer un fichier")
uploaded_file = st.sidebar.file_uploader("Téléversez un fichier Excel", type=["xlsx", "xls"])

watched_files = [f for f in os.listdir(WATCHED_FOLDER) if f.endswith((".xlsx", ".xls"))]
selected_file = None
latest_file_path = get_latest_excel_file(WATCHED_FOLDER)
default_file = os.path.basename(latest_file_path) if latest_file_path else None

if watched_files:
    selected_file = st.sidebar.selectbox(
        "📂 Ou choisissez un fichier existant dans le dossier surveillé :",
        options=watched_files,
        index=watched_files.index(default_file) if default_file in watched_files else 0
    )


# ---------- CHARGEMENT DES DONNÉES ---------- #
df_onglet_1 = None
df_onglet_2 = None
df_onglet_3 = None
source = None
if uploaded_file:
    try:
        label, date_value = get_date_generation_restitution(uploaded_file)
        df_onglet_1 = load_excel_data_dynamic_start(uploaded_file)
        df_onglet_2 = pd.read_excel(uploaded_file, sheet_name="Restit Brute (Total Répdts)", engine="calamine")
        df_onglet_3 = pd.read_excel(uploaded_file, sheet_name="Restit Rapport (Total Répdts)", engine="calamine")
        source = f"🔼 Fichier Uploadé - {label} : {date_value}"
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier uploadé : {e}")

elif selected_file:
    file_path = os.path.join(WATCHED_FOLDER, selected_file)
    try:
        label, date_value = get_date_generation_restitution(file_path)
        df_onglet_1 = load_excel_data_dynamic_start(file_path)
        df_onglet_2 = pd.read_excel(file_path, sheet_name="Restit Brute (Total Répdts)", engine="calamine")
        df_onglet_3 = pd.read_excel(file_path, sheet_name="Restit Rapport (Total Répdts)", engine="calamine")
        df_onglet_3 = df_onglet_3[['Code\nRAPPORT Ind.\nVIRTUEL',
       'Nom\nRAPPORT Ind.\nVIRTUEL', 'Code\nREPORTING\nInd. VIRTUEL',
       'Nom\nREPORTING Ind.\nVIRTUEL', 'Code\nAPP Indicateur\nVIRTUEL',
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
       'Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N']]
        source = f"👀 Fichier sélectionné - {label} : {date_value}"
        
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier sélectionné : {e}")


# ---------- AFFICHAGE DES DONNÉES ---------- #
if df_onglet_1 is not None and df_onglet_2 is not None and df_onglet_3 is not None:
    st.success(source)

    onglet_selectionne = st.radio("🧭 Choisissez l'onglet à afficher :", [
        "📈 Paramètres restitution (onglet 1)",
        "📊 Données brutes (onglet 2)",
        "📋 Rapport consolidé (onglet 3)"
    ])

    df_filtré = None
    if onglet_selectionne == "📈 Paramètres restitution (onglet 1)":
        st.subheader("📈 Aperçu : Paramètres restitution")
        st.dataframe(df_onglet_1, use_container_width=True)

    elif onglet_selectionne == "📊 Données brutes (onglet 2)":
        st.subheader("📊 Aperçu : Données Brutes")
        st.dataframe(df_onglet_2, use_container_width=True)

    elif onglet_selectionne == "📋 Rapport consolidé (onglet 3)":
        st.subheader("📋 Rapport Consolidé : Masquez/Démasquez des colonnes")

        # 🧩 Colonnes sélectionnables
        colonnes_masquables = ['Code\nRAPPORT Ind.\nVIRTUEL',
       'Nom\nRAPPORT Ind.\nVIRTUEL', 'Code\nREPORTING\nInd. VIRTUEL',
       'Nom\nREPORTING Ind.\nVIRTUEL', 'Code\nAPP Indicateur\nVIRTUEL',
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
       'Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N']

        # Filtrage
        colonnes_visibles = st.multiselect("🕶️ Colonnes à afficher :", options=colonnes_masquables, default=colonnes_masquables)

        # Sélection et affichage
        df_filtré = df_onglet_3[colonnes_visibles] if colonnes_visibles else pd.DataFrame()
        
        

    # ---------- AFFICHAGE FINAL ---------- #
    # Onglets
    tab1, tab2 = st.tabs(["📋 Données formatées", "📊 Visualisations"])
    
    
    if df_filtré is not None:
        # ---------- Onglet 1 : Tableau formaté ---------- #
        with tab1:
            st.subheader("📋 Données formatées et mises en forme")
            st.success(f"✅ Données chargées ({source})")

            st.subheader("🔍 Aperçu des données")
            st.dataframe(df_filtré, use_container_width=True)


            # ---------- KPIs ---------- #
            # Filtrage des colonnes numériques visibles
            colonnes_numeriques_visibles = df_filtré.select_dtypes(include='number').columns.tolist()
            st.subheader("📌 Indicateurs clés")
            numeric_cols = df_filtré[colonnes_numeriques_visibles]
            if not numeric_cols.empty:
                def format_euro(val):
                    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ") + "€"

                col1, col2, col3 = st.columns(3)
                col1.metric("💶 Total Valo Financière Collecte Réelle N", format_euro(numeric_cols['Total\nValo. Financière\nCollecte Réelle\nExercice N'].sum().sum()))
                col2.metric("💶 Total Valo Financière Objectifs Stratégiques", format_euro(numeric_cols['Total\nValo. Financière\nCollecte\nO.Strat\nExercice N'].sum().sum()))
                col3.metric("💶 Total Valo Financière Objectifs Opérationnels", format_euro(numeric_cols['Total\nValo. Financière\nCollecte\nO.Opér.\nExercice N'].sum().sum()))
            else:
                st.warning("Aucune colonne numérique trouvée pour les indicateurs.")


            # ---------- RENOMMER LES COLONNES ---------- #
            df_style = df_filtré.rename(columns={
                'Nom\nAPP Indicateur\nVIRTUEL': 'Nom indicateur abrégé',
                "Unité de conversion\n(de l'indicateur virtuel)": "Unité de conversion de l'indicateur",
                'Total\nMontant\nCollecte\nRéelle\nExercice N-1': 'Reel N-1',
                'Total\nMontant\nCollecte\nRéelle\nExercice N': 'Reel N',
                'Total\nMontant\nCollecte\nO.Strat Plafond\nExercice N': 'Objectifs Stratégiques PLAFOND période N',
                'Total\nMontant\nCollecte\nO.Strat Plancher\nExercice N': 'Objectifs Stratégiques SEUIL  période N',
                'Total\nMontant\nCollecte\nO.Opér.Plafond\nExercice N': 'Objectifs Opérationnels PLAFOND période N',
                'Total\nMontant\nCollecte\nO.Opér.Plancher\nExercice N': 'Objectifs Opérationnels SEUIL  période N',
            })

            # ---------- CALCUL DES VARIATIONS ---------- #
            df_style["VARIATION Objectifs Stratégiques période N  VS  Réel N"] = (
                df_style["Objectifs Stratégiques PLAFOND période N"] - df_style["Reel N"]
            )
            df_style["VARIATION Objectifs Opérationnels période N  VS  Réel N"] = (
                df_style["Objectifs Opérationnels PLAFOND période N"] - df_style["Reel N"]
            )

            # ---------- AFFICHAGE DU TABLEAU STYLÉ ---------- #
            st.subheader("📋 Données formatées et mises en forme")

            cols_to_display = [
                "Nom indicateur abrégé",
                "Unité de conversion de l'indicateur",
                "Reel N-1",
                "Reel N",
                "Objectifs Stratégiques SEUIL  période N",
                "Objectifs Stratégiques PLAFOND période N",
                "VARIATION Objectifs Stratégiques période N  VS  Réel N",
                "Objectifs Opérationnels SEUIL  période N",
                "Objectifs Opérationnels PLAFOND période N",
                "VARIATION Objectifs Opérationnels période N  VS  Réel N"
            ]

            # Fonctions de style
            def style_variation(val):
                try:
                    if pd.isna(val):
                        return ''
                    color = 'green' if val >= 0 else 'red'
                    return f'background-color: {color}; color: white; font-weight: bold'
                except:
                    return ''

            styler = (
                df_style[cols_to_display].style
                .applymap(style_variation, subset=["VARIATION Objectifs Stratégiques période N  VS  Réel N"])
                .applymap(style_variation, subset=["VARIATION Objectifs Opérationnels période N  VS  Réel N"])
                .applymap(lambda v: 'font-weight: bold' if isinstance(v, str) and "Territoire" in v else '', subset=["Nom indicateur abrégé"])
                .applymap(lambda v: 'background-color: #FFF2CC', subset=["Objectifs Stratégiques PLAFOND période N"])
                .applymap(lambda v: 'background-color: #D9EAD3', subset=["Reel N"])
                .applymap(lambda v: 'background-color: #F4CCCC', subset=[
                    "VARIATION Objectifs Stratégiques période N  VS  Réel N",
                    "VARIATION Objectifs Opérationnels période N  VS  Réel N"
                ])
                .applymap(lambda v: 'background-color: #EAD1DC', subset=["Objectifs Opérationnels PLAFOND période N"])
                .set_table_styles([
                    {
                        'selector': 'th',  # headers
                        'props': [
                            ('white-space', 'pre-wrap'),
                            ('word-wrap', 'break-word'),
                            ('max-width', '150px'),
                            ('font-size', '12px')
                        ]
                    },
                    {
                        'selector': 'td',  # data cells
                        'props': [
                            ('white-space', 'pre-wrap'),
                            ('word-wrap', 'break-word'),
                            ('max-width', '150px'),
                            ('font-size', '12px')
                        ]
                    }
                ])
            )


            st.markdown(
                styler.to_html(),
                unsafe_allow_html=True
            )


        # ---------- COURBE ---------- #
        with tab2: 
            selected_fig = None  # Pour l'export plus bas

            st.subheader("📈 Visualisation")

            # Graphique 1 - Eau consommée (Barre horizontale avec valeur réelle)
            eau = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("eau consommée", case=False, na=False)]
            eau_val = float(eau['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0])
            objectif_eau = 1500000  

            st.subheader("💧 Total eau consommée")

            # Créer une barre horizontale avec Plotly
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=[eau_val],
                y=["Eau consommée"],
                orientation='h',
                marker_color='dodgerblue',
                name="Réel"
            ))

            fig1.add_trace(go.Bar(
                x=[max(objectif_eau - eau_val, 0)],
                y=["Eau consommée"],
                orientation='h',
                marker_color='lightgray',
                name="Reste à consommer",
                hoverinfo='none'
            ))

            fig1.update_layout(
                title="Avancement de la consommation d’eau",
                xaxis=dict(range=[0, objectif_eau], title="LE"),
                barmode='stack',
                height=120,
                margin=dict(l=100, r=20, t=40, b=20),
                showlegend=False
            )

            st.metric(label="Eau consommée", value=f"{eau_val:,.0f}".replace(",", " ") + " Litres d'Eau")
            st.plotly_chart(fig1, use_container_width=True)



            # Graphique 2 - % d'eau stockée avec une jauge
            stock = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("eau stockée", case=False, na=False)]
            st.subheader("📊 % d’eau stockée")
            value = float(stock['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0])

            # Affichage du pourcentage sous forme de texte
            st.metric(label="Eau stockée", value=f"{value:.1f} %")

            # Création de la jauge
            fig2 = go.Figure(go.Indicator(
            mode="gauge+number",
            value=value,
            number={'suffix': " %"},  # <-- C’est ici qu’on ajoute le %
            title={"text": "Pourcentage d'eau stockée"},
            gauge={
                'axis': {'range': [0, 100]},
                'bar': {'color': "lightblue"},
                'steps': [
                    {'range': [0, 50], 'color': "orange"},
                    {'range': [50, 75], 'color': "yellow"},
                    {'range': [75, 100], 'color': "green"}
                ],
                'threshold': {
                    'line': {'color': "black", 'width': 4},
                    'thickness': 0.75,
                    'value': value
                }
            }
            ))

            # Affichage de la jauge
            st.plotly_chart(fig2)

            # Graphique 3 - Fréquence réunion RSE
            rse = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("réunion RSE", case=False, na=False)]
            st.subheader("📅 Fréquence Réunions RSE")
            fig2, ax2 = plt.subplots()
            val = float(str(rse['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0]).replace('%', ''))
            ax2.pie([val, 100-val], labels=[f"Réunions RSE ({val}%)", f"Autres Réunions ({100-val}%)"], autopct="%1.1f%%", startangle=90)
            ax2.set_title("Fréquence Réunions RSE")
            st.pyplot(fig2)


        # ---------- EXPORT ---------- #
        st.subheader("⬇️ Exporter les données et visualisations")

        if st.button("📤 Exporter en Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Onglet 1 : Données
                df_style.to_excel(writer, sheet_name='Données', index=False)

                # Onglet 2 : Graphiques
                workbook = writer.book
                sheet = workbook.create_sheet(title='Graphiques')

                # Graphique 1 - Eau consommée (barre de progression réelle vs objectif)
                eau = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("eau consommée", case=False, na=False)]
                eau_val = float(eau['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0])
                objectif_eau = 1500000  # Valeur cible

                fig1, ax1 = plt.subplots(figsize=(6, 1.5))
                ax1.barh([0], [objectif_eau], color="#e0e0e0", edgecolor='gray')  # Barre objectif en fond
                ax1.barh([0], [min(eau_val, objectif_eau)], color="#2196F3")  # Eau consommée
                ax1.set_xlim(0, objectif_eau)
                ax1.set_yticks([])
                ax1.set_xticks([0, objectif_eau / 2, objectif_eau])
                ax1.set_xticklabels([f"0", f"{int(objectif_eau/2):,}", f"{int(objectif_eau):,}"])
                ax1.set_title("Avancement de l’eau consommée (LE)", fontsize=12)
                ax1.spines[['top', 'right', 'left']].set_visible(False)
                ax1.set_xlabel("LE consommés")

                # Convertir en image PNG
                img_buffer1 = BytesIO()
                fig1.savefig(img_buffer1, format="png", bbox_inches="tight")
                img_buffer1.seek(0)
                img1 = XLImage(img_buffer1)
                sheet.add_image(img1, 'B2')


                # Graphique 2 - % d’eau stockée avec jauge
                stock = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("eau stockée", case=False, na=False)]
                value = float(stock['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0])

                # Création de la jauge avec Plotly

                fig2 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=value,
                    title={"text": "Pourcentage d'eau stockée"},
                    gauge={
                        'axis': {'range': [0, 100]},
                        'bar': {'color': "lightblue"},
                        'steps': [
                            {'range': [0, 50], 'color': "orange"},
                            {'range': [50, 75], 'color': "yellow"},
                            {'range': [75, 100], 'color': "green"}
                        ],
                        'threshold': {
                            'line': {'color': "black", 'width': 4},
                            'thickness': 0.75,
                            'value': value
                        }
                    }
                ))

                # Convertir le graphique de la jauge en image PNG et l'ajouter au fichier Excel
                img_buffer2 = BytesIO()
                fig2.write_image(img_buffer2, format="png")
                img_buffer2.seek(0)
                img2 = XLImage(img_buffer2)
                sheet.add_image(img2, 'M2')  # Positionner le graphique de la jauge à une position différente

                # Graphique 3 - Fréquence Réunion RSE
                rse = df_filtré[df_filtré['Nom\nAPP Indicateur\nVIRTUEL'].str.contains("réunion RSE", case=False, na=False)]
                fig3, ax3 = plt.subplots()
                val = float(str(rse['Total\nMontant\nCollecte\nRéelle\nExercice N'].values[0]).replace('%', ''))
                ax3.pie([val, 100-val], labels=[f"Réunions RSE ({val}%)", f"Autres Réunions ({100-val}%)"], autopct="%1.1f%%", startangle=90)
                ax3.set_title("Fréquence Réunions RSE")

                # Convertir le graphique en image PNG et l'ajouter au fichier Excel
                img_buffer3 = BytesIO()
                fig3.savefig(img_buffer3, format="png")
                img_buffer3.seek(0)
                img3 = XLImage(img_buffer3)
                sheet.add_image(img3, 'B16')  # Positionner le 3e graphique à une position différente

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=output.getvalue(),
                file_name="xtrafi_export_avec_graphiques.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("💡 Téléversez un fichier ou déposez-en un dans le dossier `data/` pour démarrer.")
