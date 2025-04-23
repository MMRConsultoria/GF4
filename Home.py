import streamlit as st

# Inicializa o estado de login na primeira visita
if "logado" not in st.session_state:
    st.session_state.logado = False

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria", layout="wide")

# Conteúdo principal
st.title("Portal de Relatórios 📊")

if not st.session_state.logado:
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
else:
    st.success("✅ Login realizado com sucesso. Escolha um relatório no menu à esquerda.")
