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





def filtrer_par_code_axe(df, key_prefix="default"):
    possible_cols = ["Code axe\nAPP", "Code \naxeAPP", "Axe"]
    colonne_code = next((col for col in possible_cols if col in df.columns), None)

    if colonne_code is None:
        return df  # Pas de filtre possible, on retourne df inchangé

    # On récupère les valeurs uniques présentes dans la colonne
    valeurs_disponibles = (
        df[colonne_code]
        .dropna()
        .astype(str)  # sécurité
        .unique()
    )

    # On trie pour avoir un affichage propre
    valeurs_disponibles = sorted(valeurs_disponibles)

    # Streamlit multiselect dynamique
    choix_utilisateur = st.multiselect(
        "Filtrer par axe :",
        options=valeurs_disponibles,
        default=valeurs_disponibles,  # par défaut : tout est sélectionné
        key=f"multiselect_{key_prefix}"
    )

    if not choix_utilisateur:
        return df  # aucun filtre → retour inchangé

    # Filtrer en fonction de la sélection
    return df[df[colonne_code].isin(choix_utilisateur)]
