import streamlit as st

def afficher_tableau(title, df, style_fn):
    st.subheader(title)
    styled_df = style_fn(df)
    st.markdown(styled_df.to_html(), unsafe_allow_html=True)
