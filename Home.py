# Home.py
import streamlit as st
#from logo_sidebar import mostrar_logo_cliente

# ✅ Configuração da página
st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# ✅ Verificação de login ANTES de exibir o conteúdo
if not st.session_state.get("acesso_liberado"):
    st.switch_page("pages/Login.py")  # Caminho corrigido
    st.stop()

# ✅ Exibe o logo do cliente na sidebar
#mostrar_logo_cliente()

# ✅ Logo fixo na sidebar
st.sidebar.markdown(
    """
    <div style="text-align: center; padding: 10px 0 30px 0;">
        <img src="https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_cliente.png" width="100">
    </div>
    """,
    unsafe_allow_html=True
)

# ✅ Conteúdo visível apenas após login autorizado
st.image("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")

st.success(f"✅ Acesso liberado para o código {st.session_state.get('empresa')}!")
