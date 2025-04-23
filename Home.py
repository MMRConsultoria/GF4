# pages/home.py – Página inicial com logo e redirecionamento para login

import streamlit as st

st.set_page_config(page_title="MMR Consultoria", layout="centered")

st.markdown("""
    <div style='display: flex; align-items: center; justify-content: center;'>
        <img src='https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png' alt='MMR Consultoria' style='height: 160px;'/>
    </div>
    <h1 style='text-align: center;'>Bem-vindo ao Portal de Relatórios</h1>
    <p style='text-align: center;'>Para acessar os relatórios, vá até a aba <strong>Login</strong> no menu lateral.</p>
""", unsafe_allow_html=True)

st.info("🔐 Acesse o menu à esquerda e clique em 'Login' para continuar.")
