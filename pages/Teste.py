# pages/Teste.py

import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta

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
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# ================================
# 2. Estilo visual
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
# 3. Carrega dados de vendas
# ================================

aba_vendas = "Fat Sistema Externo"
df_vendas = pd.DataFrame(planilha_empresa.worksheet(aba_vendas).get_all_records())

# 🔧 Limpa os nomes das colunas
df_vendas.columns = df_vendas.columns.str.strip()

df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()

df_vendas["Fat.Total"] = (
    df_vendas["Fat.Total"]
    .astype(str)
    .str.replace("R$", "", regex=False)
    .str.replace("(", "-")
    .str.replace(")", "")
    .str.replace(" ", "")
    .str.replace(".", "")
    .str.replace(",", ".")
)

# ✅ Converte com segurança, transformando erros em NaN
df_vendas["Fat.Total"] = pd.to_numeric(df_vendas["Fat.Total"], errors="coerce")
# ================================
# 4. Seleção de intervalo de datas
# ================================
datas_disponiveis = df_vendas["Data"].dropna().sort_values().unique()
data_inicial_default = datas_disponiveis[-1] - timedelta(days=6)
data_final_default = datas_disponiveis[-1]

data_inicio, data_fim = st.date_input(
    "Selecione o intervalo de datas:",
    value=(data_inicial_default, data_final_default),
    min_value=min(datas_disponiveis),
    max_value=max(datas_disponiveis)
)

# ================================
# 5. Gera relatório pivô
# ================================
df_filtrado = df_vendas[
    (df_vendas["Data"] >= pd.to_datetime(data_inicio)) & 
    (df_vendas["Data"] <= pd.to_datetime(data_fim))
]

df_pivot = (
    df_filtrado
    .groupby(["Data", "Loja"], as_index=False)["Fat.Total"].sum()
    .pivot(index="Loja", columns="Data", values="Fat.Total")
)

# Renomeia colunas com formato desejado
df_pivot.columns = [f"Fat Total ({d.strftime('%d/%m/%Y')})" for d in df_pivot.columns]

# Junta com Grupo
df_final = df_pivot.reset_index().merge(df_empresa[["Loja", "Grupo"]], on="Loja", how="left")

# Garante ordem: Grupo, Loja, [valores]
cols_ordenadas = ["Grupo", "Loja"] + [col for col in df_final.columns if col not in ["Grupo", "Loja"]]
df_final = df_final[cols_ordenadas]

# Calcula total geral (somando apenas colunas de valores)
total_geral = df_final.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
total_geral["Grupo"] = "TOTAL"
total_geral["Loja"] = ""
df_final = pd.concat([pd.DataFrame([total_geral]), df_final], ignore_index=True)
df_final = df_final[cols_ordenadas]  # Garante que continue na ordem certa


# ================================
# 6. Exibição
# ================================
st.markdown("### 📊 Resumo por Loja - Coluna por Dia")
st.dataframe(
    df_final.style.format({col: "R$ {:,.2f}" for col in df_final.columns if "Fat Total" in col}),
    use_container_width=True,
    height=600
)
