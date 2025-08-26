import yaml
import uuid
import streamlit_authenticator as stauth
import streamlit as st
from config import LOGO_PATH
from utils.styles import load_css
from passlib.hash import bcrypt
from pathlib import Path
from authentification.reset_pswd import send_reset_email


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


def is_super_admin(username, config):
    users = config.get("credentials", {}).get("usernames", {})
    user = users.get(username, {})
    roles = user.get("roles", [])
    return "super_admin" in roles

def is_admin(username, config):
    users = config.get("credentials", {}).get("usernames", {})
    user = users.get(username, {})
    roles = user.get("roles", [])
    return "admin" in roles or is_super_admin(username, config)

def get_user_organizations(username, config):
    users = config.get("credentials", {}).get("usernames", {})
    user = users.get(username, {})
    return user.get("organizations", [])

def get_org_logo(org_name, config):
    if org_name and config:
        orgs = config.get("organizations", {})
        org_data = orgs.get(org_name, {})
        logo_path = org_data.get("logo", None)
        if logo_path:
            return logo_path
    return None



# ------------------ 🔐 INITIALISATION AUTH ------------------

def init_authenticator(show_logo=True):
    config = load_auth_config()
    # Affiche le logo de l'organisation sélectionnée si possible, sinon logo générique
    logo_path = None
    if show_logo:
        username = st.session_state.get("username")
        organization = st.session_state.get("organization")
        if username and organization:
            logo_path = get_org_logo(organization, config)
        if not logo_path:
            from config import LOGO_PATH
            logo_path = LOGO_PATH
        st.sidebar.image(logo_path, use_container_width=True)

    if "login_key" not in st.session_state:
        st.session_state.login_key = "login_form_default"

    config["cookie"]["key"] = st.session_state.login_key

    authenticator = stauth.Authenticate(
        credentials=config["credentials"],
        cookie_name=config["cookie"]["name"],
        cookie_key=config["cookie"]["key"],
        cookie_expiry_days=config["cookie"]["expiry_days"],
    )
    return authenticator, config

# ------------------ 👤 INTERFACE DE CONNEXION ------------------

def login_interface(authenticator, config):
    st.markdown("## Connexion à Xtrafi BI")

    # authenticator, config = init_authenticator()
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
    col1,col2 = st.columns(2)
    with col1:
        if st.button("🔐 Modifier mon mot de passe"):
            st.session_state.page = "change_password"
            st.rerun()
    with col2:
        if st.button("🔑 Mot de passe oublié ?"):
            st.session_state.page = "forgot_password"
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

def forgot_password_page():
    st.markdown("## 🔑 Mot de passe oublié ?")

    user_input = st.text_input("Entrez votre email ou nom d'utilisateur")
    if st.button("Envoyer le lien de réinitialisation"):

        config = load_auth_config()
        users = config["credentials"]["usernames"]

        # Rechercher l'utilisateur correspondant soit par username soit par email
        matched_user = None
        for username, data in users.items():
            if user_input == username or user_input == data.get("email"):
                matched_user = username
                break

        if not matched_user:
            st.error("Utilisateur introuvable.")
            return

        # Générer un token temporaire (UUID)
        token = str(uuid.uuid4())
        users[matched_user]["reset_token"] = token
        save_auth_config(config)

        # Créer le lien de réinitialisation
        from urllib.parse import quote

        reset_link = f"https://xtrafibi-v1.streamlit.app/?reset_token={quote(token)}&user={quote(matched_user)}"

        # reset_link = f"http://localhost:8501/?reset_token={token}&user={matched_user}"

        # Récupérer l'email pour l'envoi
        recipient_email = users[matched_user].get("email")
        if not recipient_email:
            st.error("Cet utilisateur n'a pas d'adresse email configurée.")
            return

        # Envoi de l'email
        sender_email = "xtrafibi@gmail.com"
        app_password = "uzlcebsviunylziy"  # mot de passe d'application Gmail
        success = send_reset_email(recipient_email, reset_link, sender_email, app_password)

        if success:
            st.success("Lien de réinitialisation envoyé par email !")
        else:
            st.error("Erreur lors de l'envoi de l'email.")




def reset_password_page():
    config = load_auth_config()
    users = config["credentials"]["usernames"]

    token = st.session_state.get("reset_token")
    username = st.session_state.get("reset_user")

    # DEBUG
    st.write("DEBUG - username:", username)
    st.write("DEBUG - token:", token)
    st.write("DEBUG - token YAML:", users.get(username, {}).get("reset_token"))

    if username not in users or users[username].get("reset_token") != token:
        st.error("Lien invalide ou expiré.")
        return

    new_pwd = st.text_input("Nouveau mot de passe", type="password")
    confirm_pwd = st.text_input("Confirmer le mot de passe", type="password")

    if st.button("Réinitialiser le mot de passe"):
        if new_pwd != confirm_pwd:
            st.error("Les mots de passe ne correspondent pas.")
        elif len(new_pwd) < 6:
            st.warning("Mot de passe trop court (6 caractères minimum).")
        else:
            users[username]["password"] = bcrypt.hash(new_pwd)
            users[username].pop("reset_token", None)
            save_auth_config(config)
            st.success("Mot de passe réinitialisé avec succès !")
            st.session_state.page = "login"
            st.rerun()
