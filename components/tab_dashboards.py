import streamlit as st

def afficher_tableau(title, df, style_fn,logo):
    col1,col2 = st.columns([2,1])
    with col1:
        st.subheader(title)
    with col2:
        st.image(logo,width=100)
    styled_df = style_fn(df)
    st.markdown(styled_df.to_html(), unsafe_allow_html=True)
