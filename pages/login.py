# pages/login.py – Tela de login com controle de sessão (sem exibir alerta após login)

import streamlit as st

st.set_page_config(page_title="Login | MMR Consultoria")
st.title("🔐 Acesso Restrito")

# Empresas cadastradas com suas respectivas senhas
empresas = {
    "Empresa A": "1234",
    "Empresa B": "abcd",
    "Empresa C": "senha2024"
}

st.markdown("Escolha sua empresa e digite a senha para acessar os relatórios.")

empresa = st.selectbox("Empresa:", list(empresas.keys()))
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    if senha == empresas[empresa]:
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = empresa
        st.success(f"✅ Acesso liberado para {empresa}!")
    else:
        st.error("❌ Senha incorreta. Tente novamente.")

# Exibe aviso apenas se não estiver autenticado
if "acesso_liberado" not in st.session_state:
    st.warning("⚠️ Você precisa fazer login para visualizar os relatórios.")
