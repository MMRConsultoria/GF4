# pages/PainelResultados.py

import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Painel de Indicadores", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] {
        margin-top: 20px;
    }
    button[data-baseweb="tab"] {
        font-size: 16px;
        font-weight: 600;
        padding: 10px 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/combo-chart--v1.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.2rem;'>📊 Painel de Indicadores</h1>
    </div>
""", unsafe_allow_html=True)

# Criar as 4 abas
aba1, aba2, aba3, aba4 = st.tabs([
    "📈 Visão Geral",
    "📆 Análise Mensal",
    "📉 Comparativo Anual",
    "📋 Relatório Detalhado"
])

# Aba 1: Visão Geral
with aba1:
    st.subheader("📈 Visão Geral")
    st.info("Aqui você pode adicionar um gráfico resumo, KPIs principais ou destaques estratégicos.")

# Aba 2: Análise Mensal
with aba2:
    st.subheader("📆 Análise Mensal")
    st.info("Coloque aqui um gráfico de barras ou linhas mês a mês.")

# Aba 3: Comparativo Anual
with aba3:
    st.subheader("📉 Comparativo Anual")
    st.info("Ideal para mostrar evolução por ano ou por trimestre.")

# Aba 4: Relatório Detalhado
with aba4:
    st.subheader("📋 Relatório Detalhado")
    st.info("Você pode colocar tabelas detalhadas e botões de download aqui.")
