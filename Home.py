import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# Checa acesso liberado
if "acesso_liberado" not in st.session_state:
    st.session_state.acesso_liberado = False

# Se não tiver acesso, mostra aviso e para
if not st.session_state.acesso_liberado:
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
    st.stop()

# Se tiver acesso, mostra conteúdo normalmente
st.image("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {st.session_state.get('empresa')}!")
