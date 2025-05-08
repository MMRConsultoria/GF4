
import streamlit as st

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# Logo discreto no canto superior direito
st.markdown(
    """
    <div style="display: flex; justify-content: flex-end; align-items: center; margin-bottom: -30px;">
        <img src="https://raw.githubusercontent.com/seu_usuario/seu_repositorio/main/logo_fit.png" alt="Logo" height="40">
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
