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

# Ne remplir que si ce n'est pas déjà défini
if "reset_token" not in st.session_state or not st.session_state.reset_token:
    st.session_state.reset_token = unquote(query_params.get("reset_token", ""))

if "reset_user" not in st.session_state or not st.session_state.reset_user:
    st.session_state.reset_user = unquote(query_params.get("user", ""))

# Debug
st.write("DEBUG - session token:", st.session_state.reset_token)
st.write("DEBUG - session user:", st.session_state.reset_user)

# Si les deux valeurs sont présentes, passer sur la page reset_password
if st.session_state.reset_token and st.session_state.reset_user:
    st.session_state.page = "reset_password"

# Routage des pages
if st.session_state.get("page") == "reset_password":
    from authentification.auth import reset_password_page
    reset_password_page()
    st.stop()
elif st.session_state.get("page") == "forgot_password":
    from authentification.auth import forgot_password_page
    forgot_password_page()
    st.stop()




# ------------------ 🔐 AUTHENTICATEUR ------------------
authenticator, config = init_authenticator(show_logo=False)

# ------------------ 🔍 INFOS DE SESSION ------------------
authentication_status = st.session_state.get("authentication_status")
username = st.session_state.get("username")
name = st.session_state.get("name")
organization = st.session_state.get("organization")
user_data = config.get("credentials", {}).get("usernames", {}).get(username, {})
roles = user_data.get("roles", [])

# ------------------ 🔀 ROUTAGE SELON ROLE ------------------
if authentication_status is None:
    login_interface(authenticator, config)
    st.stop()
elif authentication_status is False:
    st.error("Nom d'utilisateur ou mot de passe incorrect.")
elif authentication_status is True:
    if "super_admin" in roles:
        nav = st.sidebar.selectbox("🔐 Menu Super Admin", ["Dashboard", "Admin"], key="navigation")
        if nav == "Admin":
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
