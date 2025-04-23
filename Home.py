# Home.py – Página inicial com controle de login

import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

if not st.session_state.get("acesso_liberado"):
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
    st.stop()

st.image("logo-mmr.png", width=120)
st.title("Bem-vindo ao Portal de Relatórios")
st.markdown(f"✅ Olá, **{st.session_state.get('empresa')}**. Use o menu lateral para acessar os relatórios disponíveis.")
