import streamlit as st

st.set_page_config(page_title="Login | MMR Consultoria")

# Redireciona se já estiver logado
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")  # Redireciona para a página principal

st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, seu e-mail e senha para acessar os relatórios.")

codigo_empresa = st.text_input("Código da Empresa:")
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    if codigo_empresa == "1825" and senha == "1234":
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo_empresa
        st.switch_page("Home.py")
    else:
        st.error("❌ Código ou senha incorretos. Tente novamente.")
