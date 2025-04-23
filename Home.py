# Home.py – Página inicial com controle de login e logo

import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# Bloqueia acesso se não estiver logado
if not st.session_state.get("acesso_liberado"):
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
    st.stop()

# Logo da MMR direto do GitHub Pages
st.image("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success("✅ Acesso liberado! Escolha um relatório no menu lateral.")
