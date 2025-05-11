# pages/Relatorios Gereciais.py

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

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela_Empresa").get_all_records())

planilha_dados = gc.open("Faturamento Sistema Externo")
aba_dados = planilha_dados.worksheet("Fat Sistema Externo")
df = pd.DataFrame(aba_dados.get_all_records())

st.set_page_config(page_title="Relatórios Gerenciais", layout="wide")
st.title("📊 Relatórios Gerenciais")

aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
    "📊 Gráfico Anual Comparativo",
    "🗓️ Relatório Mensal Detalhado",
    "📌 Análise Extra 1",
    "📌 Análise Extra 2",
    "📌 Análise Extra 3",
    "📌 Análise Extra 4",
    "📌 Análise Extra 5"
])

with aba1:
    st.subheader("📊 Gráfico Anual Comparativo")
    anos_disponiveis = sorted(df["Ano"].dropna().unique())

    anos_comparacao = st.multiselect(
        "📊 Anos para gráficos de comparação",
        options=anos_disponiveis,
        default=anos_disponiveis
    )

    def limpar_valor(x):
        try:
            if isinstance(x, str):
                return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
            elif isinstance(x, (int, float)):
                return x
        except:
            return None

    for coluna in ["Fat.Total", "Serv/Tx", "Fat.Real"]:
        if coluna in df.columns:
            df[coluna] = df[coluna].apply(limpar_valor)
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    df["Ano"] = df["Data"].dt.year
    df["Mês"] = df["Data"].dt.month

    meses_portugues = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    df["Nome Mês"] = df["Mês"].map(meses_portugues)
    df_anos_filtrado = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"])
    df_anos = df_anos_filtrado.copy()

    fat_mensal = df_anos.groupby(["Nome Mês", "Ano"])["Fat.Real"].sum().reset_index()
    meses = {
        "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
        "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
    }

    fat_mensal["MesNum"] = fat_mensal["Nome Mês"].str[:3].str.lower().map(meses)
    fat_mensal["Ano"] = fat_mensal["Ano"].astype(str)
    fat_mensal["MesAno"] = fat_mensal["Nome Mês"].str[:3].str.capitalize() + "/" + fat_mensal["Ano"].str[-2:]
    fat_mensal = fat_mensal.sort_values(["MesNum", "Ano"])

    color_map = {"2024": "#1f77b4", "2025": "#ff7f0e"}

    fig = px.bar(
        fat_mensal,
        x="Nome Mês",
        y="Fat.Real",
        color="Ano",
        barmode="group",
        text_auto=".2s",
        custom_data=["MesAno"],
        color_discrete_map=color_map
    )

    fig.update_traces(textposition="outside")
    fig.update_layout(
        xaxis_title=None,
        yaxis_title=None,
        xaxis_tickangle=-45,
        showlegend=False,
        yaxis=dict(showticklabels=False, showgrid=False, zeroline=False)
    )
    st.plotly_chart(fig, use_container_width=True)
