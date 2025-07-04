import streamlit as st
import yaml
import bcrypt
from pathlib import Path
import pandas as pd
from config import LOGO_PATH
from utils.styles import load_css
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

CONFIG_FILE = 'authentification/auth_config.yaml'

def load_config():
    if not Path(CONFIG_FILE).exists():
        st.error(f"Le fichier de configuration {CONFIG_FILE} est introuvable.")
        return None
    with open(CONFIG_FILE, "r") as f:
        return yaml.safe_load(f)

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        yaml.safe_dump(config, f, default_flow_style=False)

def hash_password(password: str) -> str:
    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt())
    return hashed.decode()

def verify_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode(), hashed.encode())

def check_admin_access(username, config):
    if config is None:
        return False
    users = config.get("credentials", {}).get("usernames", {})
    user = users.get(username, {})
    roles = user.get("roles", [])  
    return "admin" in roles

def run_admin_panel(username, authenticator, name):
    config = load_config()

    if not check_admin_access(username, config):
        st.error("Accès refusé. Seuls les administrateurs peuvent accéder à cette page.")
        st.stop()
    admin_interface(username, config, authenticator, name)

def admin_interface(current_username, config, authenticator, name):
    if config is None:
        st.stop()

    st.sidebar.image(LOGO_PATH, use_container_width=True)
    st.sidebar.markdown(f"### Bienvenue, {name}")
    if st.sidebar.button("Déconnexion", key="logout_admin"):
        authenticator.logout("main")
        st.session_state.logout_triggered = True
        st.rerun()
    st.title("Admin - Gestion des utilisateurs")

    users = config.get("credentials", {}).get("usernames", {})
    user_list = sorted(users.keys())

    # Création d'un DataFrame pour le tableau informatif
    df_users = pd.DataFrame([
        {
            "Nom d'utilisateur": u,
            "Nom complet": data.get("name", ""),
            "Rôles": ", ".join(data.get("roles", [])),
            "Organisations": ", ".join(data.get("organizations", []))
        }
        for u, data in users.items()
    ]).sort_values("Nom d'utilisateur")

    # Disposition en colonnes : 1/3 à gauche pour le form, 2/3 à droite pour tableau
    col1, col2 = st.columns([1, 2])

    with col1:
        choix = ["Nouvel utilisateur"] + user_list
        selected_choice = st.selectbox("Sélectionnez un utilisateur ou 'Nouvel utilisateur'", choix)

        is_new = selected_choice == "Nouvel utilisateur"

        if is_new:
            user_data = {"name": "", "password": "", "roles": ["user"], "organizations": []}
        else:
            user_data = users[selected_choice].copy()
            user_data.setdefault("roles", ["user"])
            user_data.setdefault("organizations", [])

        with st.form("user_form"):
            username_input = st.text_input("Nom d'utilisateur", value=selected_choice if not is_new else "")
            name_input = st.text_input("Nom complet", value=user_data.get("name", ""))
            password_input = st.text_input("Mot de passe (laisser vide pour ne pas changer)", type="password")
            role_input = st.multiselect("Rôles", ["admin", "user", "viewer"], default=user_data.get("roles", []))
            orgs_input = st.text_area("Organisations (séparées par virgule)", value=",".join(user_data.get("organizations", [])))

            submitted = st.form_submit_button("Sauvegarder")

        if submitted:
            if not username_input.strip():
                st.error("Le nom d'utilisateur ne peut pas être vide.")
            elif is_new and username_input in users:
                st.error("Ce nom d'utilisateur existe déjà.")
            else:
                user_entry = users.get(username_input, {})
                user_entry["name"] = name_input.strip()
                user_entry["roles"] = role_input
                user_entry["organizations"] = [o.strip() for o in orgs_input.split(",") if o.strip()]

                if password_input:
                    user_entry["password"] = hash_password(password_input)
                else:
                    if is_new:
                        st.error("Un mot de passe est requis pour un nouvel utilisateur.")
                        st.stop()
                    else:
                        user_entry["password"] = users[selected_choice]["password"]

                if not is_new and username_input != selected_choice:
                    del users[selected_choice]

                users[username_input] = user_entry
                config["credentials"]["usernames"] = users
                save_config(config)
                st.success(f"Utilisateur '{username_input}' sauvegardé.")
                st.rerun()

        if not is_new:
            st.write("---")  # Séparateur visuel
            st.write("### Supprimer cet utilisateur")
            confirm = st.checkbox(f"Je confirme la suppression de l'utilisateur '{selected_choice}'", key="confirm_delete")

            if st.button("Supprimer", key="delete_user_btn"):
                if confirm:
                    del users[selected_choice]
                    config["credentials"]["usernames"] = users
                    save_config(config)
                    st.success(f"Utilisateur '{selected_choice}' supprimé.")
                    st.rerun()
                else:
                    st.warning("Veuillez cocher la case de confirmation avant de supprimer.")

    with col2:
        st.write("### Tableau des utilisateurs (lecture seule)")
        st.dataframe(df_users, use_container_width=True)
