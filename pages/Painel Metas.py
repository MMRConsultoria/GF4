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
# Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# ================================
# Funções auxiliares
# ================================
def normalizar_texto(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def parse_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
    except:
        return 0.0

# ESSA FUNÇÃO É A CHAVE FINAL
def garantir_escalar(x):
    if isinstance(x, list):
        if len(x) == 1:
            return x[0]
        return str(x)
    return x

# ================================
# Carrega Tabela Empresa
# ================================
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_empresa["Loja"] = df_empresa["Loja"].str.strip()
df_empresa["De Para Metas"] = df_empresa["De Para Metas"].str.strip()
df_empresa["Loja_norm"] = df_empresa["De Para Metas"].apply(normalizar_texto)

# ================================
# Carrega Metas
# ================================
df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
df_metas["Loja"] = df_metas["Loja"].str.strip()
df_metas["Loja_norm"] = df_metas["Loja"].apply(normalizar_texto)
df_metas["Ano"] = pd.to_numeric(df_metas["Ano"], errors='coerce').fillna(0).astype(int)
df_metas["Mês"] = df_metas["Mês"].astype(str).str.strip().str.capitalize()

# Merge com de/para
df_metas = df_metas.merge(
    df_empresa[["Loja", "Loja_norm"]].rename(columns={"Loja": "Loja_Padronizada"}),
    on="Loja_norm", how="left"
)
df_metas["Loja Final"] = np.where(df_metas["Loja_Padronizada"].notna(), df_metas["Loja_Padronizada"], df_metas["Loja"])
df_metas.drop(columns=["Loja_norm", "Loja_Padronizada"], inplace=True)
df_metas.rename(columns={"Loja Final": "Loja"}, inplace=True)

# ================================
# Carrega Realizado
# ================================
df_realizado = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
df_realizado.columns = df_realizado.columns.str.strip()
df_realizado["Loja"] = df_realizado["Loja"].str.strip()
df_realizado["Mês"] = df_realizado["Data"].apply(lambda x: pd.to_datetime(x).strftime("%b"))
df_realizado["Ano"] = df_realizado["Data"].apply(lambda x: pd.to_datetime(x).year)
df_realizado["Fat.Total"] = df_realizado["Fat.Total"].apply(parse_valor)

# ================================
# Início da tela
# ================================
mes_atual = datetime.now().strftime("%b")
ano_atual = datetime.now().year
ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
anos_disponiveis = sorted(df_realizado["Ano"].unique())

aba1, aba2 = st.tabs(["📈 Análise Metas", "📊 Auditoria Metas"])

with aba1:
    ano_selecionado = st.selectbox("Selecione o Ano:", anos_disponiveis, index=anos_disponiveis.index(ano_atual) if ano_atual in anos_disponiveis else 0)
    mes_selecionado = st.selectbox("Selecione o Mês:", ordem_meses, index=ordem_meses.index(mes_atual) if mes_atual in ordem_meses else 0)

    # Filtra já no início
    df_metas_filtrado = df_metas[(df_metas["Ano"] == ano_selecionado) & (df_metas["Mês"] == mes_selecionado)].copy()
    df_realizado_filtrado = df_realizado[(df_realizado["Ano"] == ano_selecionado) & (df_realizado["Mês"] == mes_selecionado)].copy()

    # 🚩 AQUI GARANTIMOS ESCALAR PURO ANTES DO GROUPBY:
    for col in ["Ano", "Mês", "Loja"]:
        df_metas_filtrado[col] = df_metas_filtrado[col].apply(garantir_escalar)
        df_realizado_filtrado[col] = df_realizado_filtrado[col].apply(garantir_escalar)

    # Agrupamento
    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Meta"})
    realizado_grouped = df_realizado_filtrado.groupby(["Ano", "Mês", "Loja"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    # Merge final loja por loja
    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja"], how="outer").fillna(0)

    comparativo["% Atingido"] = np.where(comparativo["Meta"] == 0, np.nan, comparativo["Realizado"] / comparativo["Meta"])
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]

    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)
    comparativo = comparativo.sort_values(["Loja"])

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

with aba2:
    st.info("Em desenvolvimento.")
