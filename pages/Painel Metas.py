# pages/Painel Metas.py

import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime

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
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# ================================
# 2. Estilo e layout
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

# ================================
# Função auxiliar para converter valores
# ================================
def parse_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
    except:
        return 0.0

# 🚩 ESSA É A FUNÇÃO BLINDAGEM FINAL
def garantir_escalar(x):
    if isinstance(x, list):
        if len(x) == 1:
            return x[0]
        return str(x)
    return x

# ================================
# Abas
# ================================
aba1, aba2 = st.tabs(["📈 Analise Metas", "📊 Auditoria Metas"])

# ================================
# Aba 1: Análise
# ================================
with aba1:

    # --- Metas ---
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
    df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
    df_metas["Loja"] = df_metas["Loja"].str.strip()

    df_depara = df_empresa[["Loja", "De Para Metas"]].drop_duplicates()
    df_depara.columns = ["LojaOriginal", "LojaFinal"]

    df_metas = df_metas.merge(df_depara, left_on="Loja", right_on="LojaOriginal", how="left")
    df_metas["Loja Final"] = df_metas["LojaFinal"].fillna(df_metas["Loja"])

    # --- Realizado ---
    df_anos = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
    df_anos.columns = df_anos.columns.str.strip()
    df_anos["Loja"] = df_anos["Loja"].str.strip()
    df_anos = df_anos.merge(df_depara, left_on="Loja", right_on="LojaOriginal", how="left")
    df_anos["Loja Final"] = df_anos["LojaFinal"].fillna(df_anos["Loja"])
    df_anos["Mês"] = df_anos["Data"].apply(lambda x: pd.to_datetime(x).strftime("%b"))
    df_anos["Ano"] = df_anos["Data"].apply(lambda x: pd.to_datetime(x).year)
    df_anos["Fat.Total"] = df_anos["Fat.Total"].apply(parse_valor)

    # 🔢 Ajuste dos filtros
    mes_atual = datetime.now().strftime("%b")
    ano_atual = datetime.now().year

    ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

    anos_disponiveis = sorted(df_anos["Ano"].unique())
    ano_selecionado = st.selectbox("Selecione o Ano:", anos_disponiveis, index=anos_disponiveis.index(ano_atual) if ano_atual in anos_disponiveis else 0)
    mes_selecionado = st.selectbox("Selecione o Mês:", ordem_meses, index=ordem_meses.index(mes_atual) if mes_atual in ordem_meses else 0)

    df_anos_filtrado = df_anos[(df_anos["Ano"] == ano_selecionado) & (df_anos["Mês"] == mes_selecionado)].copy()
    df_metas_filtrado = df_metas[(df_metas["Ano"] == ano_selecionado) & (df_metas["Mês"] == mes_selecionado)].copy()

    # 🚩 BLINDAGEM antes do groupby:
    for col in ["Ano", "Mês", "Loja Final"]:
        df_metas_filtrado[col] = df_metas_filtrado[col].apply(garantir_escalar)
        df_anos_filtrado[col] = df_anos_filtrado[col].apply(garantir_escalar)

    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Meta"})
    realizado_grouped = df_anos_filtrado.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja Final"], how="outer").fillna(0)
    comparativo["% Atingido"] = np.where(comparativo["Meta"] == 0, np.nan, comparativo["Realizado"] / comparativo["Meta"])
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]

    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)
    comparativo = comparativo.sort_values(["Ano", "Loja Final", "Mês"])

    st.dataframe(
        comparativo.style.format({
            "Meta": "R$ {:,.2f}",
            "Realizado": "R$ {:,.2f}",
            "Diferença": "R$ {:,.2f}",
            "% Atingido": "{:.2%}"
        }),
        use_container_width=True
    )
