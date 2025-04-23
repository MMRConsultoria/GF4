import streamlit as st

st.set_page_config(page_title="Login | MMR Consultoria")

st.title("🔐 Acesso Restrito")
st.subheader("Escolha sua empresa e digite a senha para acessar os relatórios.")

empresa = st.text_input("Empresa:")
email = st.text_input("Email:")
senha = st.text_input("Senha", type="password")

if st.button("Entrar"):
    if empresa == "1825" and senha == "suasenha":  # troque pela lógica correta
        st.session_state.logado = True
        st.success("Login feito com sucesso!")
        st.experimental_rerun()  # recarrega a página para limpar a tela de login
    else:
        st.error("Credenciais inválidas.")
