import streamlit as st

# ------------------ ⚙️ CONFIG PAGE ------------------
st.set_page_config(page_title="Xtrafi Data Viz", layout="wide")

from app import run_app
import uuid
from authentification.auth import *
from utils.styles import load_css

# ------------------ 🎨 STYLES CSS ------------------
css_path = "static/style.css"
load_css(css_path)

# ------------------ 🔑 INITIALISATION DES VARIABLES DE SESSION ------------------
for key in ["page", "authentication_status", "username", "name", "organization", "selected_org_temp", "login_key"]:
    if key not in st.session_state or st.session_state[key] is None:
        if key == "login_key":
            st.session_state[key] = f"login_form_{uuid.uuid4()}"
        else:
            st.session_state[key] = None



# ------------------ 🔄 DÉCONNEXION GÉRÉE ------------------
if st.session_state.get("logout_triggered"):
    st.session_state.clear()
    st.session_state.login_key = f"login_form_{uuid.uuid4()}"
    st.rerun()

# ------------------ 🌐 ROUTAGE DES PAGES ------------------
if "page" not in st.session_state:
    st.session_state.page = "login"

# ------------------ 🔄 ROUTAGE POUR RÉINITIALISATION ------------------
from urllib.parse import unquote

query_params = st.query_params

reset_token = unquote(query_params.get("reset_token", ""))
reset_user = unquote(query_params.get("user", ""))

if reset_token and reset_user:
    from authentification.auth import reset_password_page
    reset_password_page(reset_token, reset_user)  # utilise st.session_state si nécessaire
    st.stop()
elif st.session_state.get("page") == "forgot_password":
    from authentification.auth import forgot_password_page
    forgot_password_page()
    st.stop()
elif st.session_state.get("page") == "change_password":
    from authentification.auth import changer_mot_de_passe_utilisateur
    changer_mot_de_passe_utilisateur()
    st.stop()

# ------------------ 🔐 AUTHENTICATEUR ------------------
authenticator, config = init_authenticator()

# ------------------ 🔍 INFOS DE SESSION ------------------
authenticator, config = st.session_state.get("authenticator"), st.session_state.get("auth_config")
if authenticator is None:
    authenticator, config = init_authenticator()

authentication_status = st.session_state.get("authentication_status")

# tant que l'utilisateur n'est pas authentifié (None ou False) -> afficher login
if authentication_status is not True:
    login_interface(authenticator, config)
    st.stop()

# ici : authentication_status is True
username = st.session_state.get("username")
name = st.session_state.get("name")


user_data = config.get("credentials", {}).get("usernames", {}).get(username, {})
roles = user_data.get("roles", [])

# ------------------ 🔀 ROUTAGE SELON AUTH ------------------
if authentication_status is False:
    st.error("Nom d'utilisateur ou mot de passe incorrect.")

elif authentication_status is True:
    # --- Étape intermédiaire : choix organisation ---
    if not st.session_state.get("organization"):
        login_interface(authenticator, config)   # c'est là que le 2e formulaire apparaît
        st.stop()

    # --- Une fois l'organisation choisie, on route par rôle ---
    if "super_admin" in roles:
        nav = st.sidebar.selectbox("🔐 Menu Super Admin", ["Dashboard", "Super Admin"], key="navigation")
        if nav == "Super Admin":
            from authentification.admin_tools import run_admin_panel
            run_admin_panel(username, authenticator, name)
        else:
            run_app(name, authenticator)
    elif "admin" in roles:
        nav = st.sidebar.selectbox("🔐 Menu administrateur", ["Dashboard", "Admin"], key="navigation")
        if nav == "Admin":
            from authentification.admin_tools import run_admin_panel
            run_admin_panel(username, authenticator, name)
        else:
            run_app(name, authenticator)
    else:
        run_app(name, authenticator)

