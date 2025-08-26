import streamlit as st
from urllib.parse import unquote

st.set_page_config(page_title="Debug Query Params", layout="wide")

st.markdown("## 🐛 Debug des query params")

# Affichage brut de query_params
st.write("### query_params (brut) :")
st.write(st.query_params)

# Extraction directe
reset_token = st.query_params.get("reset_token", "")
reset_user = st.query_params.get("user", "")

st.write("### reset_token (get direct):", reset_token, type(reset_token))
st.write("### reset_user (get direct):", reset_user, type(reset_user))

# Extraction avec unquote
st.write("### reset_token (unquote):", unquote(reset_token))
st.write("### reset_user (unquote):", unquote(reset_user))
