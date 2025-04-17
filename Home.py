import streamlit as st

st.set_page_config(page_title="Portal de Relatórios", layout="centered")

st.title("🔐 Portal de Relatórios")

# Login simples
nome_usuario = st.text_input("👤 Digite seu nome para continuar:")

if nome_usuario:
    st.success(f"Bem-vindo(a), {nome_usuario}!")

    # Seleção de relatório
    st.subheader("📑 Escolha um relatório para processar:")
    relatorio = st.selectbox("Relatórios disponíveis:", [
        "Relatório de Sangria"
        # Aqui depois você pode adicionar outros relatórios
    ])

    if st.button("📥 Processar"):
        if relatorio == "Relatório de Sangria":
            st.session_state["usuario_logado"] = nome_usuario
            st.session_state["relatorio_escolhido"] = relatorio
            st.switch_page("pages/relatorio_sangria.py")
