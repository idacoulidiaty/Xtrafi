# Xtrafi BI Dashboard

Une application Streamlit interactive permettant de visualiser, analyser et exporter des indicateurs extra-financiers (ESG : Environnement, Social, Gouvernance).
Ce projet a été développé dans le cadre de la mise en place de la brique datavisualistaion des services proposés par Xtrafi (Valence, France).



# Fonctionnalités principales

Chargement dynamique des données :

Fichiers Excel (.xls, .xlsx) issus de Xtrafi App.

Lecture robuste des onglets avec gestion des fichiers corrompus/partiels.

Conversion automatique en .xlsx si nécessaire.

Navigation claire via la sidebar :

Paramètres restitution (fichier, métadonnées, date de génération).

Données brutes et tableaux comparatifs :

Réel N vs N-1

Réel N vs Objectifs Opérationnels

Réel N vs Objectifs Stratégiques

Visualisations interactives : graphiques Plotly, filtres ESG, vision globale.

Filtres dynamiques :

Filtrage automatique par Axe ESG (Environnement, Social, Gouvernance).

Options générées directement à partir des données importées.

Tableaux enrichis :

Mise en forme stylisée.

Sélection automatique des indicateurs ESG.

Export Excel :

Génération d’un fichier Excel enrichi et bien mis en page.

Téléchargement direct depuis l’interface Streamlit.

Authentification & rôles :

Gestion sécurisée des utilisateurs avec streamlit_authenticator.

Interface d’administration (Ajout/modification d’utilisateurs).




# Technologies utilisées

Streamlit
 – Framework de dashboard interactif.

Pandas
 – Manipulation des données.

Plotly
 – Graphiques interactifs.

OpenPyXL
 & Calamine
 – Lecture Excel robuste.

AgGrid
 – Tableaux interactifs.

streamlit-authenticator
 – Authentification utilisateur.



# Installation & lancement

1. Créer un environnement virtuel
python -m venv venv
source venv/bin/activate   # Linux/Mac
venv\Scripts\activate      # Windows

2. Installer les dépendances
pip install -r requirements.txt

3. Lancer l’application
streamlit run main.py



# Exemple d’utilisation

Charger un fichier de restitution exporté depuis Xtrafi App.

Explorer les données via les différentes vues (N vs N-1, Objectifs, Vision Globale).

Filtrer par Axe ESG grâce au multiselect dynamique.

Générer et télécharger un rapport Excel enrichi.

🔒 Sécurité & authentification

Connexion via login/mot de passe.

Sélection de l’organisation associée après authentification.

Rôle admin → accès à l’interface de gestion des utilisateurs.



# Points forts

Résilient aux fichiers partiels ou mal formatés.

Interface simple et claire pour des utilisateurs non techniques.

Exports automatisés prêts à être utilisés pour le reporting.

Extensible : nouveaux filtres ou axes ESG peuvent être ajoutés facilement.
