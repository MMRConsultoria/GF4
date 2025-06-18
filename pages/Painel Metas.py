# pages/Painel Metas.py

import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
import unicodedata
import re

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
# 2. Carrega Tabela Empresa
# ================================
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_empresa["Loja"] = df_empresa["Loja"].str.strip()
df_empresa["De Para Metas"] = df_empresa["De Para Metas"].str.strip()

# ================================
# 3. Função auxiliar para normalizar textos
# ================================
def normalizar_texto(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'\s+', ' ', texto)
    return texto

# ================================
# 4. Estilo e layout
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
# 5. Abas
# ================================
aba1, aba2 = st.tabs(["📈 Análise Metas", "📊 Auditoria Metas"])

# ================================
# Função auxiliar para tratar valores
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

# ================================
# Aba 1: Análise Metas
# ================================
with aba1:

    # ---- Carrega Metas ----
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
    df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
    df_metas["Loja"] = df_metas["Loja"].str.strip()

    # ---- Normaliza textos ----
    df_metas["Loja_norm"] = df_metas["Loja"].apply(normalizar_texto)
    df_empresa["DePara_norm"] = df_empresa["De Para Metas"].apply(normalizar_texto)

    # ---- Merge com de/para com nomes normalizados ----
    df_metas = df_metas.merge(
        df_empresa[["Loja", "DePara_norm"]].rename(columns={"Loja": "Loja_Padronizada"}),
        left_on="Loja_norm",
        right_on="DePara_norm",
        how="left"
    )

    # Se encontrou correspondência usa o nome oficial da Tabela Empresa, senão mantém nome original da Metas
    df_metas["Loja Final"] = np.where(df_metas["Loja_Padronizada"].notna(), df_metas["Loja_Padronizada"], df_metas["Loja"])

    df_metas.drop(columns=["DePara_norm", "Loja_Padronizada", "Loja_norm"], inplace=True)
    df_metas.rename(columns={"Loja Final": "Loja"}, inplace=True)

    # Ajusta as colunas de Ano e Mês (garante consistência)
    df_metas["Mês"] = df_metas["Mês"].str.strip()
    df_metas["Ano"] = pd.to_numeric(df_metas["Ano"], errors='coerce').fillna(0).astype(int)

    # ---- Carrega Realizado ----
    df_anos = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
    df_anos.columns = df_anos.columns.str.strip()
    df_anos["Loja"] = df_anos["Loja"].str.strip()
    df_anos["Mês"] = df_anos["Data"].apply(lambda x: pd.to_datetime(x).strftime("%b"))
    df_anos["Ano"] = df_anos["Data"].apply(lambda x: pd.to_datetime(x).year)
    df_anos["Fat.Total"] = df_anos["Fat.Total"].apply(parse_valor)

    # 🔢 Filtros de período
    mes_atual = datetime.now().strftime("%b")
    ano_atual = datetime.now().year

    ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

    anos_disponiveis = sorted(df_anos["Ano"].unique())
    ano_selecionado = st.selectbox("Selecione o Ano:", anos_disponiveis, index=anos_disponiveis.index(ano_atual) if ano_atual in anos_disponiveis else 0)
    mes_selecionado = st.selectbox("Selecione o Mês:", ordem_meses, index=ordem_meses.index(mes_atual) if mes_atual in ordem_meses else 0)

    df_metas_filtrado = df_metas[(df_metas["Ano"] == ano_selecionado) & (df_metas["Mês"] == mes_selecionado)]
    df_anos_filtrado = df_anos[(df_anos["Ano"] == ano_selecionado) & (df_anos["Mês"] == mes_selecionado)]

    # Agrupamento
    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Meta"})
    realizado_grouped = df_anos_filtrado.groupby(["Ano", "Mês", "Loja"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja"], how="outer").fillna(0)
    comparativo["% Atingido"] = comparativo["Realizado"] / comparativo["Meta"].replace(0, np.nan)
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]

    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)
    comparativo = comparativo.sort_values(["Ano", "Loja", "Mês"])

    # Exibição
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
# Aba 2: Auditoria (em desenvolvimento)
# ================================
with aba2:
    st.info("Em desenvolvimento.")
