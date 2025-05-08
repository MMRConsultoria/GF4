
import streamlit as st
from logo_sidebar import mostrar_logo_cliente

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")


st.sidebar.markdown(
    """
    <div style="text-align: center; padding: 10px 0 30px 0;">
        <img src="https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_cliente.png" width="100">
    </div>
    """,
    unsafe_allow_html=True
)
# Garante que a chave exista antes de usar
if "acesso_liberado" not in st.session_state:
    st.session_state["acesso_liberado"] = False

# Se não estiver logado, para o app silenciosamente (sem mensagem)
if not st.session_state["acesso_liberado"]:
    st.stop()

# Conteúdo visível apenas após login
st.image("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {st.session_state.get('empresa')}!")
