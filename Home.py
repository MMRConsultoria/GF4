# Home.py – Exibe mensagem apenas se não estiver autenticado

import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

st.image("logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")

if not st.session_state.get("acesso_liberado"):
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
else:
    st.success("✅ Acesso liberado! Escolha um relatório no menu lateral.")
