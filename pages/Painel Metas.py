# pages/Painel Metas.py
import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# ================================
# 2. Dados necessários
# ================================
# 🏢 Tabela Empresa (com De/Para)
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# 🎯 Metas
df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())

# 📊 Realizado
df_anos = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
df_anos["Data"] = pd.to_datetime(df_anos["Data"], errors="coerce", dayfirst=True)

# ================================
# 3. Layout
# ================================
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] { margin-top: 20px; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Metas Mensais</h1>
    </div>
""", unsafe_allow_html=True)

aba1, aba2 = st.tabs([
    "📈 Analise Metas",
    "📊 Auditoria Metas"
])

# ================================
# Aba 1: Comparativo Metas vs. Realizado
# ================================
with aba1:
    st.subheader("📊 Comparativo Metas vs. Realizado por Loja (Fat.Total)")

    # Função para converter valores R$ em float
    def parse_valor(val):
        if isinstance(val, str):
            return float(val.replace("R$", "").replace(".", "").replace(",", ".").strip())
        return float(val or 0)

    # ----------------------------
    # 🧼 Limpeza e normalização
    # ----------------------------
    df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
    df_metas["Loja"] = df_metas["Loja"].str.strip()

    df_depara = df_empresa[["Loja", "De Para Metas"]].drop_duplicates()
    df_depara.columns = ["LojaOriginal", "LojaFinal"]

    # Aplica De/Para em metas
    df_metas = df_metas.merge(df_depara, left_on="Loja", right_on="LojaOriginal", how="left")
    df_metas["Loja Final"] = df_metas["LojaFinal"].fillna(df_metas["Loja"])

    # Aplica De/Para em realizado
    df_anos["Loja"] = df_anos["Loja"].str.strip()
    df_anos = df_anos.merge(df_depara, left_on="Loja", right_on="LojaOriginal", how="left")
    df_anos["Loja Final"] = df_anos["LojaFinal"].fillna(df_anos["Loja"])

    df_anos["Mês"] = df_anos["Data"].dt.strftime("%b")  # Jan, Fev, ...
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Fat.Total"] = df_anos["Fat.Total"].apply(parse_valor)

    # ----------------------------
    # 🎯 Agrupamentos
    # ----------------------------
    metas_grouped = df_metas.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index()
    metas_grouped = metas_grouped.rename(columns={"Fat.Total": "Meta"})

    realizado_grouped = df_anos.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index()
    realizado_grouped = realizado_grouped.rename(columns={"Fat.Total": "Realizado"})

    # ----------------------------
    # 🔗 Merge e cálculo
    # ----------------------------
    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja Final"], how="outer").fillna(0)
    comparativo["% Atingido"] = comparativo["Realizado"] / comparativo["Meta"].replace(0, np.nan)
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]

    # Ordenação de mês correta
    ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)

    comparativo = comparativo.sort_values(["Ano", "Loja Final", "Mês"])

    # ----------------------------
    # 📊 Exibição
    # ----------------------------
    st.dataframe(
        comparativo.style.format({
            "Meta": "R$ {:,.2f}",
            "Realizado": "R$ {:,.2f}",
            "Diferença": "R$ {:,.2f}",
            "% Atingido": "{:.2%}"
        }),
        use_container_width=True
    )

# ================================
# Aba 2: Em construção
# ================================
with aba2:
    st.info("🔧 Esta aba está em desenvolvimento.")
