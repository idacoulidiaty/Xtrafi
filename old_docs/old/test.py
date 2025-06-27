import streamlit as st
import pandas as pd
import os

# --- Simuler une base de comptes utilisateurs avec plusieurs orgas ---
users = {
    "utilisateur1": {"password": "motdepasse1", "orgas": ["OrgaX", "OrgaZ"]},
    "ida": {"password": "123456", "orgas": ["Xtrafi","Xtrafi_learning"]}
}

st.title("🔐 Connexion à l'application")

# --- Champs d'identification ---
username = st.text_input("👤 Nom d'utilisateur")
password = st.text_input("🔑 Mot de passe", type="password")

# --- Authentification ---
if username in users and password == users[username]["password"]:
    # Récupère les orgas possibles
    orgas_possibles = users[username]["orgas"]

    # Sélection de l'organisation
    if len(orgas_possibles) > 1:
        orga = st.selectbox("🏢 Sélectionnez votre organisation", orgas_possibles)
    else:
        orga = orgas_possibles[0]
        st.info(f"Organisation associée automatiquement : **{orga}**")

    st.success(f"Bienvenue **{username}** de l'organisation **{orga}**")

    # --- Construction du chemin d'accès au fichier Excel ---
    base_dir = os.path.expanduser("~/Documents") 
    dossier_utilisateur = os.path.join(base_dir, orga, "restitution", username)

    st.write("📁 Dossier recherché :", dossier_utilisateur)

    # --- Recherche d’un fichier Excel dans le dossier ---
    if os.path.isdir(dossier_utilisateur):
        fichiers_excel = [
            f for f in os.listdir(dossier_utilisateur)
            if f.endswith(".xls") or f.endswith(".xlsx")
        ]
        if fichiers_excel:
            st.success(f"📄 Fichiers trouvés :")
            st.write(fichiers_excel)  # Affiche la liste sous forme de tableau
        else:
            st.warning("⚠️ Aucun fichier Excel trouvé dans ce dossier.")
    else:
        st.error("🚫 Dossier non trouvé.")
