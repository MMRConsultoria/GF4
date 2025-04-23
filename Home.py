# Home.py – Página inicial com verificação de login e logo remoto

import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# Logo remoto hospedado no GitHub Pages
st.image("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")

# Verifica se o login foi feito
if not st.session_state.get("acesso_liberado"):
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
else:
    st.success("✅ Acesso liberado! Escolha um relatório no menu lateral.")
