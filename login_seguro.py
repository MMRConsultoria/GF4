# login_seguro.py
import streamlit as st
import socket
import requests

st.set_page_config(page_title="Login | MMR Consultoria")

# Lista de IPs autorizados
IPS_AUTORIZADOS = ["189.54.123.45", "201.10.22.33"]  # atualize conforme necessário

# Usuários cadastrados
USUARIOS = {
    "1825": {"email": "mari@mmr.com", "senha": "nova_senha_segura"},
    # adicione mais se quiser
}

# 🔍 Descobrir IP externo do usuário
@st.cache_data(ttl=600)
def get_ip():
    try:
        return requests.get("https://api.ipify.org").text
    except:
        return "0.0.0.0"

ip_usuario = get_ip()

# Se já estiver logado, redireciona
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")

st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, e-mail e senha.")

codigo = st.text_input("Código da Empresa:")
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    usuario = USUARIOS.get(codigo)
    if usuario and usuario["senha"] == senha and usuario["email"] == email:
        if ip_usuario in IPS_AUTORIZADOS:
            st.session_state["acesso_liberado"] = True
            st.session_state["empresa"] = codigo
            st.switch_page("Home.py")
        else:
            st.error("❌ IP não autorizado.")
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
