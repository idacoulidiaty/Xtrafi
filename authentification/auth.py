import yaml
import uuid
import streamlit_authenticator as stauth
import streamlit as st
from config import LOGO_PATH
from utils.styles import load_css
from passlib.hash import bcrypt
from pathlib import Path

# ------------------ 🎨 STYLES CSS ------------------
css_path = "static/style.css"
load_css(css_path)

# ------------------ 📄 CHEMIN DU FICHIER DE CONFIG ------------------
CONFIG_FILE = "authentification/auth_config.yaml"

# ------------------ ⚙️ UTILITAIRES ------------------

def load_auth_config():
    with open(CONFIG_FILE, "r") as file:
        return yaml.safe_load(file)

def save_auth_config(config):
    with open(CONFIG_FILE, "w") as file:
        yaml.safe_dump(config, file, default_flow_style=False)

# ------------------ 🔐 INITIALISATION AUTH ------------------

def init_authenticator(show_logo=True):
    # Affiche le logo dans la sidebar
    if show_logo:
        st.sidebar.image(LOGO_PATH, use_container_width=True)

    if "login_key" not in st.session_state:
        st.session_state.login_key = "login_form_default"

    config = load_auth_config()
    config["cookie"]["key"] = st.session_state.login_key

    authenticator = stauth.Authenticate(
        credentials=config["credentials"],
        cookie_name=config["cookie"]["name"],
        cookie_key=config["cookie"]["key"],
        cookie_expiry_days=config["cookie"]["expiry_days"],
    )
    return authenticator, config

# ------------------ 👤 INTERFACE DE CONNEXION ------------------

def login_interface():
    st.markdown("## Connexion à Xtrafi BI")

    authenticator, config = init_authenticator()
    authenticator.login(location="main", max_login_attempts=3, key=st.session_state.login_key)

    authentication_status = st.session_state.get("authentication_status")
    username = st.session_state.get("username")

    if authentication_status:
        users = config["credentials"]["usernames"]
        user_data = users.get(username, {})

        orgs = user_data.get("organizations", [])
        if len(orgs) > 1:
            selected_org = st.selectbox("📂 Sélectionnez une organisation :", orgs, key="org_selector")
            st.session_state.organization = selected_org
        elif len(orgs) == 1:
            st.session_state.organization = orgs[0]
        else:
            st.session_state.organization = None

        if st.button("➡️ Accéder à la plateforme"):
            st.session_state.page = "app"
            st.rerun()

    elif authentication_status is False:
        st.error("Nom d'utilisateur ou mot de passe incorrect.")
    elif authentication_status is None:
        st.info("Veuillez entrer vos identifiants.")

    st.markdown("---")
    if st.button("🔐 Modifier mon mot de passe"):
        st.session_state.page = "change_password"
        st.rerun()

# ------------------ 🔒 CHANGEMENT DE MOT DE PASSE ------------------

def changer_mot_de_passe_utilisateur():
    st.markdown("## 🔐 Modifier mon mot de passe")
    config = load_auth_config()
    users = config["credentials"]["usernames"]

    # 🔑 On récupère le nom d'utilisateur depuis un champ, pas depuis la session
    username = st.text_input("Nom d'utilisateur")

    current_pwd = st.text_input("Mot de passe actuel", type="password")
    new_pwd = st.text_input("Nouveau mot de passe", type="password")
    confirm_pwd = st.text_input("Confirmer le nouveau mot de passe", type="password")

    if st.button("Enregistrer"):
        if username not in users:
            st.error("❌ Utilisateur non trouvé.")
            return

        stored_pwd_hash = users[username]["password"]
        if not bcrypt.verify(current_pwd, stored_pwd_hash):
            st.error("❌ Mot de passe actuel incorrect.")
        elif new_pwd != confirm_pwd:
            st.error("❌ Les nouveaux mots de passe ne correspondent pas.")
        elif len(new_pwd) < 6:
            st.warning("⚠️ Mot de passe trop court (6 caractères minimum).")
        else:
            users[username]["password"] = bcrypt.hash(new_pwd)
            save_auth_config(config)
            st.success("✅ Mot de passe modifié avec succès.")
            st.session_state.page = "login"
            st.rerun()
