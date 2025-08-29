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



def delete_organization(users, org_name, config, confirm_delete, current_username):
    if not confirm_delete:
        return
    # Retirer l'organisation de tous les users
    for u in users:
        if org_name in users[u].get("organizations", []):
            users[u]["organizations"].remove(org_name)
    # Supprimer l'organisation du config
    if "organizations" in config and org_name in config["organizations"]:
        del config["organizations"][org_name]
    save_config(config)
    st.success(f"Organisation '{org_name}' supprimée avec succès.")



# def afficher_tableau_organisations(config, current_roles, current_orgs):
#     st.write("### Tableau des organisations")
#     orgs = config.get("organizations", {})
#     users = config.get("credentials", {}).get("usernames", {})

#     data = []
#     for org_name, org_data in orgs.items():
#         if "super_admin" in current_roles or org_name in current_orgs:
#             users_in_org = [u for u, d in users.items() if org_name in d.get("organizations", [])]
#             data.append({
#                 "Organisation": org_name,
#                 "Nombre d'utilisateurs": len(users_in_org),
#                 "Logo": org_data.get("logo", "")
#             })

#     df_orgs = pd.DataFrame(data)
#     if df_orgs.empty:
#         df_orgs = pd.DataFrame(columns=["Organisation", "Nombre d'utilisateurs", "Logo"])
#     gb = GridOptionsBuilder.from_dataframe(df_orgs)
#     gb.configure_pagination()
#     grid_options = gb.build()
#     AgGrid(df_orgs, gridOptions=grid_options, update_mode=GridUpdateMode.NO_UPDATE)



# ------------------ 🎨 Préparations ------------------
def load_admin_helpers(config, current_username):
    users = config.get("credentials", {}).get("usernames", {})
    current_user = users.get(current_username, {})
    current_roles = current_user.get("roles", [])
    current_orgs = current_user.get("organizations", [])

    if "selected_orga" not in st.session_state:
        st.session_state.selected_orga = None
    if "selected_users_in_orga" not in st.session_state:
        st.session_state.selected_users_in_orga = []
    if "onglet_admin_actif" not in st.session_state:
        st.session_state.onglet_admin_actif = "Utilisateurs"

    # Flash messages après rerun
    if st.session_state.get("flash_rerun"):
        st.success(st.session_state.get("flash_message", ""))
        st.session_state["flash_rerun"] = False

    return users, current_roles, current_orgs


# ------------------ 💾 Sauvegarde et fichiers ------------------
def save_and_rerun(config, message="Modifications enregistrées", selected_orga=None, selected_users=None):
    save_config(config)
    st.session_state["flash_message"] = message
    st.session_state["flash_rerun"] = True
    if selected_orga is not None:
        st.session_state.selected_orga = selected_orga
    if selected_users is not None:
        st.session_state.selected_users_in_orga = selected_users
    st.rerun()

def upload_logo(file, org_name):
    if file:
        path = f"static/logos/{org_name.replace(' ', '_').lower()}.png"
        with open(path, "wb") as f:
            f.write(file.getbuffer())
        return path
    return None



# ------------------ 👤 Gestion utilisateurs ------------------
def user_management(config, current_roles, current_orgs, visible_users):
    users = config["credentials"]["usernames"]

    # Filtrer les utilisateurs visibles selon l'admin courant
    visible_users = filter_users_by_admin(st.session_state.get("username"), config)

    if "selected_choice" not in st.session_state:
    # Si des utilisateurs sont visibles → par défaut le premier utilisateur
        st.session_state.selected_choice = visible_users[0] if visible_users else "Nouvel utilisateur"


    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("### Formulaire utilisateur")

        # --- Bouton unique "Nouvel utilisateur" ---
        if st.button("➕ Nouvel utilisateur", key="btn_new_user_1"):
            st.session_state.selected_choice = "Nouvel utilisateur"
            # Force le rerun pour réinitialiser le formulaire
            st.rerun()

        # Déterminer si on est en mode création ou édition
        is_new = st.session_state.selected_choice == "Nouvel utilisateur"

        # Afficher le selectbox seulement si on n'est PAS en mode création
        if not is_new:
            choix = visible_users
            if st.session_state.selected_choice not in choix:
                st.session_state.selected_choice = choix[0] if choix else "Nouvel utilisateur"

            selected_choice = st.selectbox(
                "Sélectionnez un utilisateur à administrer",
                choix,
                index=choix.index(st.session_state.selected_choice),
                key="selectbox_user_choice"
            )
            st.session_state.selected_choice = selected_choice
        else:
            selected_choice = "Nouvel utilisateur"


        # Rôles disponibles selon le rôle de l'admin
        available_roles = ["user", "viewer"]
        if "super_admin" in current_roles:
            available_roles += ["admin", "super_admin"]
        elif "admin" in current_roles:
            available_roles += ["admin"]

        # Organisations visibles selon le rôle
        visible_orgs_for_user = (
            list(config.get("organizations", {}).keys())
            if "super_admin" in current_roles
            else [o for o in current_orgs if o in config.get("organizations", {})]
        )

        # Valeurs par défaut selon utilisateur sélectionné
        if is_new:
            st.write("### Création d'un nouvel utilisateur")
            if st.button("⬅ Retour"):
                # Réinitialiser à un utilisateur par défaut (le 1er visible si dispo)
                st.session_state.selected_choice = visible_users[0] if visible_users else "Nouvel utilisateur"
                st.rerun()
            default_name, default_email, default_roles, default_orgs = "", "", ["user"], []
        else:
            udata = users[selected_choice]
            default_name = udata.get("name", "")
            default_email = udata.get("email", "")
            default_roles = udata.get("roles", ["user"])
            default_orgs = [o for o in udata.get("organizations", []) if o in visible_orgs_for_user]

        # --- Formulaire utilisateur ---
        with st.form("user_form", clear_on_submit=False):
            username_input = st.text_input("Nom d'utilisateur", value="" if is_new else selected_choice)
            name_input = st.text_input("Nom complet", value=default_name)
            email_input = st.text_input("Email", value=default_email)
            password_input = st.text_input("Mot de passe (laisser vide pour ne pas changer)", type="password")
            role_input = st.multiselect("Rôles", options=available_roles, default=default_roles)
            orgs_input = st.multiselect("Organisations", options=visible_orgs_for_user, default=default_orgs)
            submitted = st.form_submit_button("Sauvegarder")

        if submitted:
            # --- Vérifications obligatoires ---
            if not username_input.strip():
                st.error("Le nom d'utilisateur ne peut pas être vide.")
            elif is_new and username_input in users:
                st.error("Ce nom d'utilisateur existe déjà.")
            elif not email_input.strip():
                st.error("L'email est obligatoire pour créer un utilisateur.")
                st.stop()
            elif "@" not in email_input:
                st.error("L'email doit contenir un '@'.")
                st.stop()
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

                # Si le nom a changé pour un utilisateur existant
                if not is_new and username_input != selected_choice:
                    del users[selected_choice]
                users[username_input] = user_entry
                config["credentials"]["usernames"] = users

                if not orgs_input and "super_admin" not in role_input:
                    st.error("Au moins une organisation doit être affectée à cet utilisateur.")
                    st.stop()

                save_config(config)
                st.success(f"Utilisateur '{username_input}' sauvegardé.")

                # Mise à jour du selectbox uniquement si c'est un nouvel utilisateur
                if is_new:
                    st.session_state.selected_choice = "Nouvel utilisateur"
                else:
                    st.session_state.selected_choice = username_input  # reste sur l'utilisateur modifié

                st.rerun()

        # --- Suppression utilisateur ---
        if not is_new:
            st.write("---")
            st.write(f"### Supprimer '{selected_choice}'")
            confirm = st.checkbox(
                f"Je confirme la suppression de l'utilisateur '{selected_choice}'",
                key=f"confirm_delete_{selected_choice}"
            )
            if st.button("Supprimer", key=f"delete_user_btn_{selected_choice}"):
                if confirm:
                    del users[selected_choice]
                    config["credentials"]["usernames"] = users
                    save_config(config)
                    st.session_state["flash_message"] = f"Utilisateur '{selected_choice}' supprimé."
                    st.session_state["flash_rerun"] = True


                    # ✅ Redirection logique après suppression
                    remaining_users = filter_users_by_admin(st.session_state.get("username"), config)
                    if remaining_users:
                        st.session_state.selected_choice = remaining_users[0]
                    else:
                        st.session_state.selected_choice = "Nouvel utilisateur"

                    st.rerun()
                else:
                    st.warning("Veuillez cocher la case de confirmation avant de supprimer.")


    # --- Colonne 2 : Liste lecture seule ---
    with col2:
        st.write("### Liste des utilisateurs (Lecture seule)")
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
        st.dataframe(df_users, use_container_width=True, height=600)



# ------------------ 🏢 Gestion organisations ------------------
def organization_management(config, current_roles, current_orgs):
    users = config["credentials"]["usernames"]
    col1, col2 = st.columns(2)

    # Détermination des orga visibles et droit édition
    if "super_admin" in current_roles:
        visible_orgs = list(config.get("organizations", {}).keys())
        can_edit_all_orgs = True
    elif "admin" in current_roles:
        visible_orgs = [o for o in current_orgs if o in config.get("organizations", {})]
        can_edit_all_orgs = False
    else:
        st.warning("Vous n'avez pas accès à la gestion des organisations.")
        st.stop()

    # ⚠️ Message pour les admins classiques concernant la suppression
    if "admin" in current_roles and "super_admin" not in current_roles:
        st.info("ℹ️ Vous n'avez pas accès à la suppression des organisations. "
                "Merci de contacter votre administrateur si nécessaire.")

    # --- Colonne 1 : Gestion des organisations ---
    with col1:
        st.markdown("### Gestion des Organisations")

        # Initialisation de la sélection
        if "selected_orga" not in st.session_state:
            st.session_state.selected_orga = visible_orgs[0] if visible_orgs else None

        # Bouton unique "Nouvelle organisation"
        if can_edit_all_orgs:
            if st.button("➕ Nouvelle organisation", key="btn_new_org_1"):
                st.session_state.selected_orga = "Nouvelle organisation"

        # Selectbox pour choisir organisation
        selectable_orgs = visible_orgs + (["Nouvelle organisation"] if can_edit_all_orgs else [])
        if st.session_state.selected_orga not in selectable_orgs:
            st.session_state.selected_orga = selectable_orgs[0] if selectable_orgs else None
        selected_orga = st.selectbox(
            "Sélectionnez une organisation",
            options=selectable_orgs,
            index=selectable_orgs.index(st.session_state.selected_orga),
            key="selectbox_orgas"
        )
        st.session_state.selected_orga = selected_orga

        # Liste users affectables
        users_list = []
        if "super_admin" in current_roles:
            users_list = list(users.keys())
        elif "admin" in current_roles:
            for u, data in users.items():
                if set(data.get("organizations", [])) & set(current_orgs):
                    if any(r in ["user", "viewer", "admin"] for r in data.get("roles", [])):
                        users_list.append(u)

        # --- Création nouvelle organisation ---
        if selected_orga == "Nouvelle organisation" and can_edit_all_orgs:
            new_org_name = st.text_input("Nom de la nouvelle organisation", "")
            uploaded_logo = st.file_uploader("Uploader un logo (png/jpg)", type=["png", "jpg", "jpeg"])
            selected_users = st.multiselect("Modifier les utilisateurs affectés",
                                            options=users_list,
                                            default=[])

            if st.button("Créer organisation", key="create_new_org"):
                if not new_org_name.strip():
                    st.error("Le nom de l'organisation est requis.")
                elif new_org_name in config.get("organizations", {}):
                    st.error("Cette organisation existe déjà.")
                else:
                    logo_path = upload_logo(uploaded_logo, new_org_name)
                    config["organizations"][new_org_name] = {"logo": logo_path or ""}
                    for u in selected_users:
                        users[u].setdefault("organizations", []).append(new_org_name)
                    config["credentials"]["usernames"] = users
                    st.session_state.selected_orga = new_org_name
                    st.success(f"L'organisation '{new_org_name}' a été créée !")
                    save_and_rerun(config)

        # --- Modification organisation existante ---
        elif selected_orga in visible_orgs:
            org_data = config["organizations"][selected_orga]
            logo_path = org_data.get("logo", "")
            if logo_path:
                st.image(logo_path, width=150)
            else:
                st.info("Aucun logo disponible pour cette organisation.")

            # Modifier nom et users
            org_name_input = st.text_input("Modifier le nom de l'organisation", value=selected_orga)
            users_in_orga = [u for u, data in users.items() if selected_orga in data.get("organizations", [])]
            selected_users = st.multiselect("Modifier les utilisateurs affectés",
                                            options=users_list,
                                            default=users_in_orga)

            if st.button("Mettre à jour l'organisation", key="update_org"):
                # Changement de nom
                if org_name_input != selected_orga:
                    if org_name_input in config["organizations"]:
                        st.error("Une organisation avec ce nom existe déjà.")
                        st.stop()
                    config["organizations"][org_name_input] = config["organizations"].pop(selected_orga)
                    for u in users:
                        if selected_orga in users[u].get("organizations", []):
                            users[u]["organizations"].remove(selected_orga)
                            users[u]["organizations"].append(org_name_input)
                    st.session_state.selected_orga = org_name_input
                    selected_orga = org_name_input
                    visible_orgs = list(config.get("organizations", {}).keys())

                # Mise à jour users
                for u in users:
                    if selected_orga in users[u].get("organizations", []):
                        users[u]["organizations"].remove(selected_orga)
                for u in selected_users:
                    users[u].setdefault("organizations", []).append(selected_orga)

                config["credentials"]["usernames"] = users
                save_config(config)
                st.success("Organisation mise à jour !")
                st.session_state.selected_orga = selected_orga
                st.rerun()

            # --- Modifier logo ---
            new_logo = st.file_uploader("Uploader un nouveau logo (laisser vide pour garder actuel)",
                                        type=["png", "jpg", "jpeg"], key="upload_logo")
            if st.button("Mettre à jour le logo", key="update_logo") and new_logo is not None:
                logo_path = upload_logo(new_logo, selected_orga)
                config["organizations"][selected_orga]["logo"] = logo_path
                save_config(config)
                st.success("Logo mis à jour !")
                st.session_state.selected_orga = selected_orga
                st.rerun()

            # --- Suppression sécurisée ---
            if "super_admin" in current_roles:
                st.write("---")
                st.write(f"### Supprimer l'organisation '{selected_orga}'")
                blocked_users = [
                    u for u, data in users.items()
                    if selected_orga in data.get("organizations", []) and len(data.get("organizations", [])) == 1
                    and "super_admin" not in data.get("roles", [])
                ]
                delete_allowed = not blocked_users
                if blocked_users:
                    st.warning(f"⚠️ Impossible de supprimer '{selected_orga}'. "
                               f"Les utilisateurs suivants y sont uniquement affectés : {', '.join(blocked_users)}. "
                               "Veuillez les affecter à une autre organisation avant de supprimer.")
                confirm_delete = st.checkbox(f"Je confirme la suppression de l'organisation '{selected_orga}'",
                                            key="confirm_delete_orga")
                if st.button("Supprimer l'organisation", key="delete_orga_btn"):
                    if not delete_allowed:
                        st.error("Suppression impossible : certains utilisateurs seraient laissés sans organisation.")
                    elif not confirm_delete:
                        st.warning("Veuillez cocher la case de confirmation avant de supprimer.")
                    else:
                        delete_organization(users, selected_orga, config, confirm_delete, st.session_state.get("username"))
                        st.session_state.selected_orga = None
                        st.session_state.selected_users_in_orga = []

    # --- Colonne 2 : Tableau lecture seule ---
    with col2:
        st.write("### Organisations et utilisateurs affectés (Lecture seule)")
        org_list = []
        for org_name, org_data in config.get("organizations", {}).items():
            users_in_org = [u for u, data in users.items() if org_name in data.get("organizations", [])]
            org_list.append({
                "Organisation": org_name,
                "Utilisateurs affectés": ", ".join(users_in_org),
                "Nombre d'utilisateurs": len(users_in_org)
            })

        df_orgs = pd.DataFrame(org_list)
        if df_orgs.empty:
            df_orgs = pd.DataFrame(columns=["Organisation", "Utilisateurs affectés", "Nombre d'utilisateurs"])
        st.dataframe(df_orgs, use_container_width=True)



def admin_interface(current_username, config, authenticator, name):
    st.title("Administration des Utilisateurs et des Organisations")
    load_css("static/style.css")
    if not config:
        st.stop()

    st.sidebar.image(
        get_org_logo(st.session_state.get("organization"), config) or LOGO_PATH,
        use_container_width=True
    )
    st.sidebar.markdown(f"### Bienvenue, {name}")
    if st.sidebar.button("Déconnexion", key="logout_admin"):
        authenticator.logout("main")
        st.session_state.logout_triggered = True
        st.rerun()

    # --- Récupération utilisateur courant ---
    users = config.get("credentials", {}).get("usernames", {})
    current_user = users.get(current_username, {})
    current_roles = current_user.get("roles", [])
    current_orgs = current_user.get("organizations", [])

    # --- Utilisateurs visibles selon rôle ---
    visible_users = filter_users_by_admin(current_username, config)

    # --- Flash messages après rerun ---
    if st.session_state.get("flash_rerun"):
        st.success(st.session_state.get("flash_message", ""))
        st.session_state["flash_rerun"] = False

    # --- Initialisation onglet actif ---
    if "onglet_admin_actif" not in st.session_state:
        st.session_state.onglet_admin_actif = "Utilisateurs"

    # --- Sélection onglet ---
    onglets_labels = ["Utilisateurs", "Organisations"]
    choix_onglet = st.radio(
        "Sélectionnez l'onglet",
        onglets_labels,
        index=onglets_labels.index(st.session_state.onglet_admin_actif),
        horizontal=True
    )
    st.session_state.onglet_admin_actif = choix_onglet

    # --- Affichage selon onglet ---
    if choix_onglet == "Utilisateurs":
        user_management(config, current_roles, current_orgs, visible_users)
    else:
        organization_management(config, current_roles, current_orgs)

    