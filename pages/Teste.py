# pages/Teste.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder

# Configuracao do app
st.set_page_config(page_title="Vendas Diarias", layout="wide")

# Bloqueia o acesso caso o usuário não esteja logado
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
# 2. Layout e titulo
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
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Vendas Diarias</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 3. Carrega dados
# ================================
aba_vendas = "Fat Sistema Externo"
df_vendas = pd.DataFrame(planilha_empresa.worksheet(aba_vendas).get_all_records())
df_vendas.columns = df_vendas.columns.str.strip()
df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()
df_vendas["Grupo"] = df_vendas["Grupo"].astype(str).str.strip()

# Converte Fat.Total com seguranca
df_vendas["Fat.Total"] = (
    df_vendas["Fat.Total"]
    .astype(str)
    .str.replace("R$", "", regex=False)
    .str.replace("(", "-", regex=False)
    .str.replace(")", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
)
df_vendas["Fat.Total"] = pd.to_numeric(df_vendas["Fat.Total"], errors="coerce")

# ================================
# 4. Seleciona periodo
# ================================
data_min = df_vendas["Data"].min()
data_max = df_vendas["Data"].max()

col1, col2 = st.columns(2)
with col1:
    data_inicio = st.date_input("Data Inicial", value=data_max, min_value=data_min, max_value=data_max)
with col2:
    data_fim = st.date_input("Data Final", value=data_max, min_value=data_min, max_value=data_max)

# ================================
# 5. Filtro e pivoteamento
# ================================
df_filtrado = df_vendas[(df_vendas["Data"] >= pd.to_datetime(data_inicio)) & (df_vendas["Data"] <= pd.to_datetime(data_fim))]

df_agrupado = df_filtrado.groupby(["Data", "Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
df_pivot = df_agrupado.pivot_table(
    index=["Grupo", "Loja"], columns="Data", values="Fat.Total", aggfunc="sum", fill_value=0
).reset_index()

# Renomeia colunas com a data
df_pivot.columns = [
    col if isinstance(col, str) else f"Fat Total ({col.strftime('%d/%m/%Y')})"
    for col in df_pivot.columns
]

# ================================
# 6. Reordena colunas e adiciona total
# ================================
colunas_existentes = df_pivot.columns.tolist()
colunas_chave = [col for col in ["Grupo", "Loja"] if col in colunas_existentes]
colunas_restantes = [col for col in colunas_existentes if col not in colunas_chave]
df_final = df_pivot[colunas_chave + colunas_restantes]

# Total geral
colunas_para_remover = [col for col in ["Grupo", "Loja"] if col in df_final.columns]
total_geral = df_final.drop(columns=colunas_para_remover).sum(numeric_only=True)
total_geral["Grupo"] = "TOTAL"
total_geral["Loja"] = ""
df_final = pd.concat([pd.DataFrame([total_geral]), df_final], ignore_index=True)

# ================================
# 7. Exibicao final
# ================================
st.markdown("### 📊 Resumo por Loja - Coluna por Dia")
st.dataframe(
    df_final.style.format({col: "R$ {:,.2f}" for col in df_final.columns if "Fat Total" in col}),
    use_container_width=True,
    height=600
)
