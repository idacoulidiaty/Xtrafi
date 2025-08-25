import streamlit as st
import yaml
import bcrypt
from pathlib import Path
import pandas as pd
from config import LOGO_PATH
from utils.styles import load_css
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from authentification.auth import get_org_logo

CONFIG_FILE = 'authentification/auth_config.yaml'

st.title("Gestion des Utilsateurs et des Organisations \n \n \n")

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

def filter_users_by_admin(current_username, config):
    """
    Retourne la liste des utilisateurs visibles par l'admin.
    Super admin voit tout, admin ne voit que users des orga où il est admin.
    """
    users = config.get("credentials", {}).get("usernames", {})
    current_user = users.get(current_username, {})
    current_roles = current_user.get("roles", [])
    
    if "super_admin" in current_roles:
        # super_admin voit tous les users
        return list(users.keys())

    if "admin" in current_roles:
        admin_orgs = current_user.get("organizations", [])
        visible_users = []
        for u, data in users.items():
            user_orgs = data.get("organizations", [])
            if set(admin_orgs).intersection(user_orgs):
                visible_users.append(u)
        return visible_users

    # autre roles : pas de visibilité
    return []


def check_admin_access(username, config):
    if config is None:
        return False
    users = config.get("credentials", {}).get("usernames", {})
    user = users.get(username, {})
    roles = user.get("roles", [])
    return "admin" in roles or "super_admin" in roles


def run_admin_panel(username, authenticator, name):
    config = load_config()

    if not check_admin_access(username, config):
        st.error("Accès refusé. Seuls les administrateurs peuvent accéder à cette page.")
        st.stop()
    admin_interface(username, config, authenticator, name)

def admin_interface(current_username, config, authenticator, name):
    if config is None:
        st.stop()

    st.sidebar.image(get_org_logo(st.session_state.get("organization"), config) or LOGO_PATH, use_container_width=True)
    st.sidebar.markdown(f"### Bienvenue, {name}")
    if st.sidebar.button("Déconnexion", key="logout_admin"):
        authenticator.logout("main")
        st.session_state.logout_triggered = True
        st.rerun()


    users = config.get("credentials", {}).get("usernames", {})
    current_user = users.get(current_username, {})
    current_roles = current_user.get("roles", [])
    current_orgs = current_user.get("organizations", [])

    # Filtrer users selon droit admin
    visible_users = filter_users_by_admin(current_username, config)

    # Préparer DataFrame avec uniquement users visibles
    df_users = pd.DataFrame([
        {
            "Nom d'utilisateur": u,
            "Nom complet": users[u].get("name", ""),
            "Rôles": ", ".join(users[u].get("roles", [])),
            "Organisations": ", ".join(users[u].get("organizations", []))
        }
        for u in visible_users
    ])
    if not df_users.empty:
        df_users = df_users.sort_values("Nom d'utilisateur")
    else:
        df_users = pd.DataFrame(columns=["Nom d'utilisateur", "Nom complet", "Rôles", "Organisations"])

    # --- Onglet Utilisateurs et Organisations Optimisé ---
    def save_and_rerun(config, message="Modifications enregistrées"):
        save_config(config)
        st.session_state["flash_message"] = message
        st.session_state["flash_rerun"] = True
        st.rerun()

    def upload_logo(file, org_name):
        if file:
            path = f"static/logos/{org_name.replace(' ', '_').lower()}.png"
            with open(path, "wb") as f:
                f.write(file.getbuffer())
            return path
        return None

    def delete_user(users, username, config, confirm_delete):
        if confirm_delete:
            if username in users:
                del users[username]
                config["credentials"]["usernames"] = users
                save_and_rerun(config, f"L'utilisateur '{username}' a été supprimé.")
        else:
            st.warning("Veuillez cocher la case de confirmation avant de supprimer.")

    def delete_organization(users, org_name, config, confirm_delete, current_username):
        if confirm_delete:
            # Vérifier que l'organisation à supprimer n'est pas la seule orga pour certains users
            blocked_users = []
            for u, data in users.items():
                if org_name in data.get("organizations", []):
                    if len(data.get("organizations", [])) == 1 and "super_admin" not in data.get("roles", []):
                        blocked_users.append(u)

            if blocked_users:
                st.warning(
                    f"Impossible de supprimer '{org_name}'. Les utilisateurs suivants y sont uniquement affectés : "
                    + ", ".join(blocked_users)
                    + ". Veuillez les affecter à une autre organisation avant."
                )
                return

            # Suppression possible
            if org_name in config["organizations"]:
                del config["organizations"][org_name]
                for u, data in users.items():
                    if org_name in data.get("organizations", []):
                        data["organizations"].remove(org_name)
                config["credentials"]["usernames"] = users
                save_and_rerun(config, f"L'organisation '{org_name}' a été supprimée.")
        else:
            st.warning("Veuillez cocher la case de confirmation avant de supprimer.")


    if st.session_state.get("flash_rerun"):
        st.success(st.session_state.get("flash_message", ""))
        st.session_state["flash_rerun"] = False
        
    # Onglets
    tab_users, tab_orgas = st.tabs(["Utilisateurs", "Organisations"])

    # --- Onglet Utilisateurs ---
    with tab_users:
        st.title("Admin - Gestion des utilisateurs")
        col1, col2 = st.columns([1, 2])

        with col1:
            choix = ["Nouvel utilisateur"] + visible_users
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
                email_input = st.text_input("Email", value=user_data.get("email", ""))
                password_input = st.text_input("Mot de passe (laisser vide pour ne pas changer)", type="password")

                # Rôles autorisés
                available_roles = ["user", "viewer"]
                if "super_admin" in current_roles:
                    available_roles += ["admin", "super_admin"]
                elif "admin" in current_roles:
                    available_roles += ["admin"]

                role_input = st.multiselect("Rôles", available_roles, default=user_data.get("roles", []))
                if "super_admin" in current_roles:
                    visible_orgs_for_user = list(config.get("organizations", {}).keys())
                else:
                    # admin classique : seules les orga dont il est membre
                    visible_orgs_for_user = [o for o in current_orgs if o in config.get("organizations", {})]

                orgs_input = st.multiselect(
                    "Organisations",
                    options=visible_orgs_for_user,
                    default=user_data.get("organizations", [])
                )


                submitted = st.form_submit_button("Sauvegarder")

            if submitted:
                if not username_input.strip():
                    st.error("Le nom d'utilisateur ne peut pas être vide.")
                elif is_new and username_input in users:
                    st.error("Ce nom d'utilisateur existe déjà.")
                else:
                    user_entry = users.get(username_input, {})
                    user_entry["name"] = name_input.strip()
                    user_entry["email"] = email_input.strip()
                    user_entry["roles"] = role_input
                    user_entry["organizations"] = orgs_input

                    if password_input:
                        user_entry["password"] = hash_password(password_input)
                    elif is_new:
                        st.error("Un mot de passe est requis pour un nouvel utilisateur.")
                        st.stop()
                    else:
                        user_entry["password"] = users[selected_choice]["password"]

                    if not is_new and username_input != selected_choice:
                        del users[selected_choice]

                    users[username_input] = user_entry
                    config["credentials"]["usernames"] = users
                    if not orgs_input and "super_admin" not in role_input:
                        st.error("Au moins une organisation doit être affectée à cet utilisateur.")
                        st.stop()
                    save_and_rerun(config)

            if not is_new:
                st.write("---")
                st.write("### Supprimer cet utilisateur")
                confirm = st.checkbox(f"Je confirme la suppression de l'utilisateur '{selected_choice}'", key="confirm_delete")
                if st.button("Supprimer", key="delete_user_btn"):
                    delete_user(users, selected_choice, config, confirm)


        with col2:
            st.write("### Tableau des utilisateurs (lecture seule)")
            df_users = pd.DataFrame([
                {
                    "Nom d'utilisateur": u,
                    "Nom complet": users[u].get("name", ""),
                    "Email": users[u].get("email", ""),
                    "Rôles": ", ".join(users[u].get("roles", [])),
                    "Organisations": ", ".join(users[u].get("organizations", []))
                }
                for u in visible_users
            ])
            if df_users.empty:
                df_users = pd.DataFrame(columns=["Nom d'utilisateur", "Nom complet", "Rôles", "Organisations"])
            else:
                df_users = df_users.sort_values("Nom d'utilisateur")
            st.dataframe(df_users, use_container_width=True)


     # --- Onglet Organisations ---
    with tab_orgas:
        st.title("Gestion des Organisations")
        col1, col2 = st.columns([2, 1])
    
        # Déterminer les organisations visibles selon le rôle
        if "super_admin" in current_roles:
            visible_orgs = list(config.get("organizations", {}).keys())
            can_edit_all_orgs = True
        elif "admin" in current_roles:
            visible_orgs = [o for o in current_orgs if o in config.get("organizations", {})]
            can_edit_all_orgs = False
        else:
            st.warning("Vous n'avez pas accès à la gestion des organisations.")
            st.stop()
    
        with col1:
            # Sélecteur d'organisation
            selectable_orgs = visible_orgs + (["Nouvelle organisation"] if can_edit_all_orgs else [])
            selected_orga = st.selectbox("Sélectionnez une organisation", selectable_orgs)
    
            users_list = list(users.keys())
    
            # Création nouvelle organisation
            if selected_orga == "Nouvelle organisation" and can_edit_all_orgs:
                new_org_name = st.text_input("Nom de la nouvelle organisation", "")
                uploaded_logo = st.file_uploader("Uploader un logo (png/jpg)", type=["png", "jpg", "jpeg"])
                selected_users = st.multiselect("Affecter des utilisateurs à cette organisation", users_list)
    
                if st.button("Créer organisation"):
                    if not new_org_name.strip():
                        st.error("Le nom de l'organisation est requis.")
                    elif new_org_name in config.get("organizations", {}):
                        st.error("Cette organisation existe déjà.")
                    else:
                        logo_path = upload_logo(uploaded_logo, new_org_name)
                        config["organizations"][new_org_name] = {"logo": logo_path or ""}
    
                        for u in selected_users:
                            if "organizations" not in users[u]:
                                users[u]["organizations"] = []
                            users[u]["organizations"].append(new_org_name)
    
                        config["credentials"]["usernames"] = users
                        save_and_rerun(config)
    
            # Modification / suppression organisation existante
            elif selected_orga in visible_orgs:
                org_data = config["organizations"][selected_orga]
                logo_path = org_data.get("logo", "")
                if logo_path:
                    st.image(logo_path, width=150)
                else:
                    st.info("Aucun logo disponible pour cette organisation.")
    
                users_in_orga = [u for u, data in users.items() if selected_orga in data.get("organizations", [])]
                st.write("### Utilisateurs affectés")
                st.write(", ".join(users_in_orga) if users_in_orga else "Aucun utilisateur")
    
                selected_users = st.multiselect("Modifier les utilisateurs affectés", users_list, default=users_in_orga)
                if st.button("Mettre à jour les utilisateurs"):
                    for u in users:
                        if selected_orga in users[u].get("organizations", []):
                            users[u]["organizations"].remove(selected_orga)
                    for u in selected_users:
                        if "organizations" not in users[u]:
                            users[u]["organizations"] = []
                        users[u]["organizations"].append(selected_orga)
                    config["credentials"]["usernames"] = users
                    save_and_rerun(config)
    
                # Upload nouveau logo
                new_logo = st.file_uploader("Uploader un nouveau logo (laisser vide pour garder actuel)", type=["png", "jpg", "jpeg"], key="upload_logo")
                if st.button("Mettre à jour le logo") and new_logo is not None:
                    logo_path = upload_logo(new_logo, selected_orga)
                    config["organizations"][selected_orga]["logo"] = logo_path
                    save_and_rerun(config)
    
                # Suppression organisation
                if "super_admin" in current_roles:
                    st.write("---")
                    st.write(f"### Supprimer l'organisation '{selected_orga}'")
                    confirm_delete = st.checkbox(f"Je confirme la suppression de l'organisation '{selected_orga}'", key="confirm_delete_orga")
                    if st.button("Supprimer l'organisation", key="delete_orga_btn"):
                        delete_organization(users, selected_orga, config, confirm_delete, current_username)
                else:
                    st.warning("Vous n'avez pas accès à la suppression des organisations. Veuillez contacter votre administrateur")
                    st.stop()
    
        with col2:
            # Tableau des organisations visibles
            df_orgs = pd.DataFrame([
                {
                    "Organisation": o
                }
                for o in visible_orgs
            ])
            if df_orgs.empty:
                df_orgs = pd.DataFrame(columns=["Organisation", "Logo"])
            else:
                df_orgs = df_orgs.sort_values("Organisation")
            st.write("### Organisations accessibles (lecture seule)")
            st.dataframe(df_orgs, use_container_width=True)
    