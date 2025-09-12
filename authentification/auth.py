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
def init_authenticator():
    if "authenticator" in st.session_state:
        return st.session_state.authenticator, st.session_state.auth_config

    config = load_auth_config()

    if "login_key" not in st.session_state:
        st.session_state.login_key = f"login_form_{uuid.uuid4()}"

    config["cookie"]["key"] = st.session_state.login_key

    authenticator = stauth.Authenticate(
        credentials=config["credentials"],
        cookie_name=config["cookie"]["name"],
        cookie_key=config["cookie"]["key"],
        cookie_expiry_days=config["cookie"]["expiry_days"],
    )

    # Container unique pour le logo
    if "sidebar_logo_container" not in st.session_state:
        st.session_state.sidebar_logo_container = st.sidebar.empty()

    st.session_state.sidebar_logo_container.image(LOGO_PATH, use_container_width=True)

    # Stocker dans la session
    st.session_state.authenticator = authenticator
    st.session_state.auth_config = config

    return authenticator, config


# ------------------ 👤 INTERFACE DE CONNEXION ------------------

def login_interface(authenticator, config):
    st.markdown("## Connexion à Xtrafi BI")

    # --- Reset uniquement à la première ouverture ---
    if "reset_done_once" not in st.session_state:
        if "authentication_status" not in st.session_state:
            st.session_state["authentication_status"] = None
        st.session_state.reset_done_once = True
        st.session_state.pop("selected_org_temp", None)

    # --- Authentification (on laisse la lib gérer les valeurs dans st.session_state) ---
    authenticator.login(
        location="main",
        max_login_attempts=10,
        key=st.session_state.get("login_key", "login_form_default"),
        fields={
            "Form name": "Bienvenue",
            "Username": "Nom d'utilisateur",
            "Password": "Mot de passe",
            "Login": "Se connecter"
        }
    )

    authentication_status = st.session_state.get("authentication_status")
    username = st.session_state.get("username")

    # Container unique pour le logo (crée si absent)
    if "sidebar_logo_container" not in st.session_state:
        st.session_state.sidebar_logo_container = st.sidebar.empty()

    # Afficher logo Xtrafi uniquement si aucune organisation choisie
    if not st.session_state.get("organization"):
        st.session_state.sidebar_logo_container.image(LOGO_PATH, use_container_width=True)

    # --- Si pas authentifié -> boutons changement mdp / mot de passe oublié ---
    if not authentication_status:
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col2:
            if st.button("🔑 Mot de passe oublié ?", key="forgot_pwd_btn"):
                st.session_state.page = "forgot_password"
                st.rerun()

    # --- Après authentification réussie ---
    if authentication_status is True:
        users = config.get("credentials", {}).get("usernames", {})
        user_data = users.get(username, {})
        orgs = user_data.get("organizations", []) or []

        # --- Détermination organisation sélectionnée (écrire toujours en session) ---
        selected_org_temp = None
        if len(orgs) > 1:
            # init session var si inexistante ou si valeur non valide
            if "selected_org_temp" not in st.session_state or st.session_state.selected_org_temp not in orgs:
                st.session_state.selected_org_temp = orgs[0]
            # selectbox with unique key - we assign result back to session var
            st.session_state.selected_org_temp = st.selectbox(
                "📂 Sélectionnez une organisation :",
                orgs,
                index=orgs.index(st.session_state.selected_org_temp) if st.session_state.selected_org_temp in orgs else 0,
                key="selected_org_temp_select"
            )
            selected_org_temp = st.session_state.selected_org_temp
        elif len(orgs) == 1:
            selected_org_temp = orgs[0]
            st.session_state.selected_org_temp = selected_org_temp
            st.markdown(f"**Organisation :** {selected_org_temp}")

        # --- Boutons toujours visibles (avec key unique) ---
        col1, col2,col3 = st.columns(3)
        with col1:
            if st.button("⬅️ Retour à la connexion"):
                st.session_state.page = "login"
                st.session_state.organization = None
                st.session_state.selected_org_temp = None
                st.session_state.authentication_status = None
                st.rerun()
        with col2:
            if st.button("🔐 Modifier mon mot de passe", key="change_pwd_btn"):
                st.session_state.page = "change_password"
                st.rerun()
        with col3:
            # Si utilisateur n'a pas d'organisation mais est super-admin, on met une valeur par défaut
            if not st.session_state.get("organization") and is_super_admin(username, config):
                st.session_state.organization = "all_organizations"
            if st.button("➡️ Accéder à la plateforme"):
                st.session_state.organization = selected_org_temp or st.session_state.organization
                st.session_state.page = "app"
                st.rerun()

        # Afficher logo si orga déjà choisie
        if st.session_state.get("organization"):
            logo_path = get_org_logo(st.session_state.organization, config) or LOGO_PATH
            st.session_state.sidebar_logo_container.image(logo_path, use_container_width=True)

    # --- Gestion erreurs / info login ---
    elif authentication_status is False:
        st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")
        # on laisse la possibilité de retenter la saisie
    elif authentication_status is None:
        st.info("Veuillez entrer vos identifiants.")



# ------------------ 🔒 CHANGEMENT DE MOT DE PASSE ------------------

def changer_mot_de_passe_utilisateur():
    with st.sidebar:
        st.image(LOGO_PATH)
    st.markdown("## 🔐 Modifier mon mot de passe")
    config = load_auth_config()
    users = config["credentials"]["usernames"]

    # 🔑 On récupère le nom d'utilisateur depuis un champ, pas depuis la session
    username = st.text_input("Nom d'utilisateur")

    current_pwd = st.text_input("Mot de passe actuel", type="password")
    new_pwd = st.text_input("Nouveau mot de passe", type="password")
    confirm_pwd = st.text_input("Confirmer le nouveau mot de passe", type="password")
    col1,col2=st.columns(2)
    with col1:
        # --- Bouton retour à la connexion ---
        if st.button("⬅️ Retour à la connexion"):
            st.session_state.page = "login"
            st.rerun()
    with col2:
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

    with st.sidebar:
        st.image(LOGO_PATH)

    user_input = st.text_input("Entrez votre email ou nom d'utilisateur")
    col1,col2=st.columns(2)
    with col1:
        # --- Bouton retour à la connexion ---
        if st.button("⬅️ Retour à la connexion"):
            st.session_state.page = "login"
            st.rerun()
    with col2:
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

            # reset_link = f"http://localhost:8501/?reset_token={quote(token)}&user={quote(matched_user)}"

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
                st.success("Lien de réinitialisation envoyé.Veuillez consulter vos emails")
            else:
                st.error("Erreur lors de l'envoi de l'email.")





def reset_password_page(reset_token, reset_user):
    with st.sidebar:
        st.image(LOGO_PATH)
    config = load_auth_config()
    users = config["credentials"]["usernames"]

    # Si le reset est déjà fait, afficher un message direct
    import time

    if st.session_state.get("reset_done"):
        st.success("✅ Mot de passe réinitialisé avec succès ! Redirection en cours...")
        st.query_params.clear()
        st.session_state.page = "login"
        time.sleep(3)
        st.rerun()
        return



    # Vérification du lien
    if reset_user not in users or users[reset_user].get("reset_token") != reset_token:
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
            users[reset_user]["password"] = bcrypt.hash(new_pwd)
            users[reset_user].pop("reset_token", None)
            save_auth_config(config)

            st.session_state.reset_done = True
            st.rerun()

