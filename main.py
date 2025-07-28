import streamlit as st
# ------------------ ⚙️ CONFIG PAGE ------------------
from app import run_app
st.set_page_config(page_title="Xtrafi Data Viz", layout="wide")


import uuid
from authentification.auth import *
from utils.styles import load_css




# ------------------ 🎨 STYLES CSS ------------------
css_path = "static/style.css"
load_css(css_path)


# ------------------ 🔄 DÉCONNEXION GÉRÉE PROPREMENT ------------------
if st.session_state.get("logout_triggered"):
    st.session_state.clear()
    st.session_state.login_key = f"login_form_{uuid.uuid4()}"
    st.rerun()

# ------------------ 🌐 ROUTAGE DES PAGES ------------------
if "page" not in st.session_state:
    st.session_state.page = "login"

# Redirection vers la page de changement de mot de passe
if st.session_state.page == "change_password":
    changer_mot_de_passe_utilisateur()
    st.stop()

if st.session_state.page == "login":
    login_interface()
    st.stop()

elif st.session_state.page == "change_password":
    changer_mot_de_passe_utilisateur()
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
    st.warning("Veuillez vous connecter.")
elif authentication_status is False:
    st.error("Nom d'utilisateur ou mot de passe incorrect.")
elif authentication_status is True:
    if "admin" in roles:
        nav = st.sidebar.selectbox("🔐 Menu administrateur", ["Dashboard", "Admin"], key="navigation")
        if nav == "Admin":
            from authentification.admin_tools import run_admin_panel
            run_admin_panel(username, authenticator, name)
        else:
            run_app(name, authenticator)
    else:
        run_app(name, authenticator)
