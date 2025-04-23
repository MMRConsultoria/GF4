import streamlit as st

st.set_page_config(page_title="Login | MMR Consultoria")

# Se já estiver logado, volta ao início
if st.session_state.get("acesso_liberado"):
    st.experimental_rerun()

st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, seu e-mail e senha para acessar os relatórios.")

codigo_empresa = st.text_input("Código da Empresa:")
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    if codigo_empresa == "1825" and senha == "1234":
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo_empresa

        # Truque: forçar ir para Home "manualmente" depois do login
        st.success("Login feito com sucesso!")
        st.experimental_rerun()
    else:
        st.error("❌ Código ou senha incorretos. Tente novamente.")
