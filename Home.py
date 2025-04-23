import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# Verifica se login foi feito
if not st.session_state.get("acesso_liberado"):
    st.warning("⚠️ Acesse o menu à esquerda e clique em 'Login' para continuar.")
    st.stop()

st.image("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {st.session_state.get('empresa')}!")
