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
# 2. Carrega tabelas base
# ================================
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
df_anos = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())

# ================================
# 3. Normalização e Preparo
# ================================
# 📌 Função para tratar valores monetários
def parse_valor(val):
    try:
        if isinstance(val, str):
            return float(val.replace("R$", "").replace(".", "").replace(",", ".").strip())
        return float(val or 0)
    except:
        return 0

# 🧹 Trata valores na aba Metas
df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
df_metas["Loja"] = df_metas["Loja"].str.strip()

# 📘 De/Para de lojas
df_depara = df_empresa[["Loja", "De Para Metas"]].drop_duplicates()
df_depara.columns = ["LojaOriginal", "LojaFinal"]

df_metas = df_metas.merge(df_depara, left_on="Loja", right_on="LojaOriginal", how="left")
df_metas["Loja Final"] = df_metas["LojaFinal"].fillna(df_metas["Loja"])

# 🎯 Agrupa metas por loja/mês
metas_grouped = df_metas.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index()
metas_grouped = metas_grouped.rename(columns={"Fat.Total": "Meta"})

# ================================
# 4. Realizado: Fat Sistema Externo
# ================================
df_anos["Loja"] = df_anos["Loja"].str.strip()
df_anos = df_anos.merge(df_depara, left_on="Loja", right_on="LojaOriginal", how="left")
df_anos["Loja Final"] = df_anos["LojaFinal"].fillna(df_anos["Loja"])
df_anos["Mês"] = df_anos["Mês"].str.strip().str[:3].str.title()  # Certifica padrão tipo "Jan"
df_anos["Ano"] = df_anos["Ano"].astype(int)
df_anos["Fat.Total"] = df_anos[" Fat.Total "].apply(parse_valor)

realizado_grouped = df_anos.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index()
realizado_grouped = realizado_grouped.rename(columns={"Fat.Total": "Realizado"})

# ================================
# 5. Junta Metas vs Realizado
# ================================
comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja Final"], how="outer").fillna(0)
comparativo["% Atingido"] = comparativo["Realizado"] / comparativo["Meta"].replace(0, np.nan)
comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]

# ➕ Adiciona Grupo
df_grupo = df_empresa[["Loja", "Grupo"]].drop_duplicates()
df_grupo.columns = ["Loja Final", "Grupo"]
comparativo = comparativo.merge(df_grupo, on="Loja Final", how="left")

# 🗂️ Ordena colunas corretamente
ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)
comparativo = comparativo.sort_values(["Ano", "Grupo", "Loja Final", "Mês"])

# ================================
# 6. Interface Streamlit
# ================================
# 🎨 Estilo visual
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

# 📌 Cabeçalho
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Metas Mensais</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 7. ABAS
# ================================
aba1, aba2 = st.tabs([
    "📈 Analise Metas",
    "📊 Auditoria Metas"
])

# ================================
# Aba 1: Comparativo
# ================================
with aba1:
    st.subheader("📊 Comparativo Metas vs. Realizado por Loja (Fat.Total)")

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
# Aba 2: Em desenvolvimento
# ================================
with aba2:
    st.info("Em desenvolvimento.")
