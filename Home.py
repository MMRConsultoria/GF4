# pages/home.py – Página inicial com notícias e cotação

import streamlit as st

st.set_page_config(page_title="Início | MMR Consultoria", layout="wide")

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/news.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Painel MMR Consultoria</h1>
    </div>
""", unsafe_allow_html=True)

st.markdown("""
Explore abaixo as principais notícias do mercado financeiro e internacional, além da cotação atualizada do dólar. Faça login para acessar os relatórios exclusivos.
""")

st.subheader("📰 Últimas notícias")
noticias = [
    "🕊️ Papa Francisco falece aos 88 anos após AVC e parada cardíaca",
    "🇨🇳 China ameaça retaliar acordos comerciais que favoreçam os EUA",
    "🇺🇦 Zelensky se declara disposto a negociar com a Rússia em caso de cessar-fogo"
]
for noticia in noticias:
    st.markdown(f"- {noticia}")

st.subheader("💵 Cotação do Dólar Hoje")
st.markdown("- **Dólar comercial:** R$ 5,72")
st.markdown("- **Variação do dia:** -1,45%")

st.info("🔐 Para acessar os relatórios financeiros e operacionais, use o menu lateral e faça login na plataforma.")
