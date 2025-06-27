import streamlit as st

def afficher_onglet_1(df):
    st.subheader("📈 Aperçu : Paramètres restitution")
    st.dataframe(df, use_container_width=True)

def afficher_onglet_2(df):
    st.subheader("📊 Aperçu : Données Brutes")
    st.dataframe(df, use_container_width=True)

def afficher_onglet_3(df):
    st.subheader("📋 Rapport consolidé")
    st.dataframe(df, use_container_width=True)
