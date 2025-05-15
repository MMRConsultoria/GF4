# login_seguro.py
import streamlit as st
import hashlib
import socket

st.set_page_config(page_title="Login | MMR Consultoria")

# ========================
# Lista de usuários válidos
# ========================
usuarios = [
    {
        "codigo": "1825",
        "email": "mari@mmr.com",
        "senha_hash": hashlib.sha256("1234".encode()).hexdigest(),
        "ip_autorizado": "177.45.12.88"  # EXEMPLO DE IP
    },
    {
        "codigo": "2024",
        "email": "cliente@exemplo.com",
        "senha_hash": hashlib.sha256("cliente2024".encode()).hexdigest(),
        "ip_autorizado": "177.45.12.99"  # OUTRO IP
    }
]

# ========================
# Função para validar credenciais e IP
# ========================
def validar_login(codigo, email, senha):
    senha_hash = hashlib.sha256(senha.encode()).hexdigest()
    ip_cliente = st.experimental_get_query_params().get("ip", ["127.0.0.1"])[0]

    for usuario in usuarios:
        if (
            usuario["codigo"] == codigo and
            usuario["email"] == email and
            usuario["senha_hash"] == senha_hash
        ):
            if usuario["ip_autorizado"] == ip_cliente:
                return True
            else:
                st.error(f"❌ IP não autorizado: {ip_cliente}")
                return False
    return False

# ========================
# Interface de login
# ========================
if st.session_state.get("acesso_liberado"):
    st.switch_page("pages/OperacionalFaturamentoReal.py")

st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, seu e-mail e senha para acessar os relatórios.")

codigo = st.text_input("Código da Empresa")
email = st.text_input("E-mail")
senha = st.text_input("Senha", type="password")

if st.button("Entrar"):
    if validar_login(codigo, email, senha):
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo
        st.switch_page("pages/OperacionalFaturamentoReal.py")
    else:
        st.error("❌ Credenciais inválidas ou IP não autorizado.")
