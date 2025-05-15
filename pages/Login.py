import streamlit as st
import requests
st.set_page_config(page_title="Login | MMR Consultoria")

query_params = st.experimental_get_query_params()
codigo_param = query_params.get("codigo", [None])[0]
empresa_param = query_params.get("empresa", [None])[0]

# ⚠️ Normaliza texto (boa prática)
if empresa_param:
    empresa_param = empresa_param.lower()

# Bloquear se vier direto, sem passar pela página HTML
if not codigo_param or not empresa_param:
    st.error("🚫 Acesso negado. Você deve acessar por meio do portal oficial da MMR Consultoria.")
    st.stop()
st.set_page_config(page_title="Login | MMR Consultoria")

# 🔍 Descobrir IP externo do usuário
@st.cache_data(ttl=600)
def get_ip():
    try:
        return requests.get("https://api.ipify.org").text
    except:
        return "0.0.0.0"

# Lista de IPs autorizados
IPS_AUTORIZADOS = ["35.197.92.111", "201.10.22.33"]  # atualize conforme necessário

# 👉 Captura o IP corretamente
ip_usuario = get_ip()
st.write(f"🛠️ Seu IP: {ip_usuario}")  # Pode remover depois do debug

# ❌ Bloqueia se IP não estiver na lista
if ip_usuario not in IPS_AUTORIZADOS:
    st.markdown("## 🔐 IP não autorizado")
    st.code(ip_usuario, language="text")
    st.info("Copie este IP e envie para a equipe da MMR Consultoria para liberar o acesso.")
    st.stop()

# ✅ Lista de usuários (permite múltiplos com o mesmo código)
USUARIOS = [
    {"codigo": "1825", "email": "maricelisrossi@gmail.com", "senha": "1825"},
    {"codigo": "1825", "email": "andre.machado@grupofit.com.br", "senha": "Sala1825"},
    # {"codigo": "3377", "email": "joao@empresa.com", "senha": "joao123"},
    # {"codigo": "0041", "email": "ana@consultoria.com", "senha": "ana456"}
]

# ✅ Redireciona se já estiver logado
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")

# 🧾 Tela de login
st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, e-mail e senha.")

codigo = st.text_input("Código da Empresa:")
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

# ✅ Botão de login
if st.button("Entrar"):
    usuario_encontrado = next(
        (u for u in USUARIOS if u["codigo"] == codigo and u["email"] == email and u["senha"] == senha),
        None
    )

    if usuario_encontrado:
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo
        st.switch_page("Home.py")
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
