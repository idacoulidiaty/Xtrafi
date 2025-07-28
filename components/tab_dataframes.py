import streamlit as st

def afficher_onglet(df, no_onglet=1):
    if no_onglet == 1:
        subheader = "📈 Aperçu : Paramètres restitution"
    elif no_onglet == 2:
        subheader = "📊 Aperçu : Données Brutes"
    elif no_onglet == 3:
        subheader = "📋 Rapport consolidé"
    else: 
        raise ValueError(f"no_onglet should be 1, 2 ou 3. You provided {no_onglet}")
    
    st.subheader(subheader)
    st.dataframe(df, use_container_width=True)





# Dictionnaire de correspondance code ↔ nom complet
CODE_AXE_MAPPING = {
    'EN': 'Environnement',
    'SO': 'Social',
    'GO': 'Gouvernance'
}

def filtrer_par_code_axe(df, key_prefix="default"):
    possible_cols = ["Code axe\nAPP", "Code \naxeAPP", "Axe"]
    colonne_code = next((col for col in possible_cols if col in df.columns), None)

    if colonne_code is None:
        return df  # Pas de filtre possible, on retourne df inchangé



    valeurs_disponibles = df[colonne_code].dropna().unique()
    valeurs_affichees = [CODE_AXE_MAPPING.get(v, v) for v in valeurs_disponibles]

    choix_utilisateur = st.multiselect(
        "Filtrer par axe ESG",
        options=valeurs_affichees,
        default=valeurs_affichees,
        key=f"multiselect_{key_prefix}"
    )

    if not choix_utilisateur:
        return df  # Aucun filtre appliqué

    codes_selectionnes = [k for k, v in CODE_AXE_MAPPING.items() if v in choix_utilisateur]

    return df[df[colonne_code].isin(codes_selectionnes)]
