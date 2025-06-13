# pages/Painel Metas.py
import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")  # ✅ Escolha um título só

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px
import io
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime, date
from datetime import datetime, date, timedelta

#st.set_page_config(page_title="Painel Agrupado", layout="wide")
#st.set_page_config(page_title="Vendas Diarias", layout="wide")
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
# 2. Configuração inicial do app
# ================================


# 🎨 Estilizar abas
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

# Cabeçalho bonito
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Metas Mensais</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 3. Separação em ABAS
# ================================
aba1, aba2 = st.tabs([
    "📈 Analise Metas",
    "📊 Auditoria Metas"
   
])

# ================================
# Aba 1: Graficos Trimestrais
# ================================
with aba1:
    #st.header("📊 Comparativo Metas vs Realizado")

    # Carrega aba de metas
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())

    # Aplica De Para Metas
    df_depara = df_empresa[["Loja", "De Para Metas"]].dropna()
    df_metas = df_metas.merge(df_depara, left_on="Loja", right_on="De Para Metas", how="left")
    df_metas["Loja Final"] = df_metas["Loja_y"].fillna(df_metas["Loja_x"])

    # Agrupa metas
    df_metas_grouped = df_metas.groupby(["Ano", "Mês", "Loja Final"])["Meta"].sum().reset_index()

    # Agrupa realizado (substitua df_fat por seu DataFrame real de vendas)
    df_realizado_grouped = df_fat_sistema_externo.groupby(["Ano", "Mês", "Loja"])["Fat.Total"].sum().reset_index()

    # Aplica De Para também no realizado (se necessário)
    df_realizado_grouped = df_realizado_grouped.merge(df_depara, left_on="Loja", right_on="Loja", how="left")
    df_realizado_grouped["Loja Final"] = df_realizado_grouped["De Para Metas"].fillna(df_realizado_grouped["Loja"])

    df_realizado_grouped = df_realizado_grouped.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index()

    # Junta metas e realizado
    df_comp = df_metas_grouped.merge(df_realizado_grouped, on=["Ano", "Mês", "Loja Final"], how="outer")
    df_comp = df_comp.fillna(0)
    df_comp["% Atingimento"] = np.where(
        df_comp["Meta"] > 0,
        df_comp["Fat.Total"] / df_comp["Meta"],
        np.nan
    )

    # Exibe tabela
    st.dataframe(df_comp.style.format({
        "Meta": "R$ {:,.2f}",
        "Fat.Total": "R$ {:,.2f}",
        "% Atingimento": "{:.2%}"
    }))

# ================================
# Aba 2: Relatorio Analitico
# ================================
with aba2:
    st.info("em desenvolvimento.")

