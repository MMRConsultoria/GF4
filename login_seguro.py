import streamlit as st
import requests

st.set_page_config(page_title="Login | MMR Consultoria")

# ================================
# 🔒 CONFIGURAÇÕES
# ================================
# Lista de IPs permitidos
IPS_AUTORIZADOS = [
    "187.65.101.50",   # Exemplo IP de casa
    "201.6.230.101",   # Exemplo IP do escritório
    # Adicione aqui outros IPs confiáveis
]

# Dados de login válidos
credenciais_validas = {
    "1825": {
        "email": "mari@mmr.com",
        "senha": "novaSenhaSegura123"
    },
    # Você pode adicionar mais empresas aqui
}

# ================================
# 🚫 VERIFICAÇÃO DE IP
# ================================
@st.cache_data(ttl=300)
def obter_ip_usuario():
    try:
        response = requests.get("https://api.ipify.org?format=json")
        return response.json()["ip"]
    except Exception:
        return None

ip_usuario = obter_ip_usuario()
if ip_usuario and ip_usuario not in IPS_AUTORIZADOS:
    st.error(f"❌ Acesso negado para este IP: {ip_usuario}")
    st.stop()

# ================================
# ✅ REDIRECIONAMENTO SE JÁ ESTIVER LOGADO
# ================================
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")

# ================================
# 🧾 FORMULÁRIO DE LOGIN
# ================================
st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, seu e-mail e senha para acessar os relatórios.")

codigo_empresa = st.text_input("Código da Empresa:")
email = st.text_input("E-mail:")
senha = st.text_input("Senha", type="password")

if st.button("Entrar"):
    cred = credenciais_validas.get(codigo_empresa)
    if cred and email == cred["email"] and senha == cred["senha"]:
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo_empresa
        st.success("✅ Login realizado com sucesso!")
        st.switch_page("Home.py")
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
