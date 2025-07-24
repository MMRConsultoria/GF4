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

# Configuração do app
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
# 2. Layout e título
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
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Vendas Diárias</h1>
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

# Converte Fat.Total com segurança
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
# 4. Seleciona período
# ================================
data_min = df_vendas["Data"].min()
data_max = df_vendas["Data"].max()

col1, col2 = st.columns([2, 3])

with col1:
    data_inicio, data_fim = st.date_input(
        "Selecione o intervalo de datas:",
        value=(data_max, data_max),
        min_value=data_min,
        max_value=data_max
    )

with col2:
    st.write("🔜 Aqui virão os filtros: Loja, Grupo, etc")

# ================================
# 5 e 6. Filtro, colunas diárias e acumulado
# ================================

# Converte datas
data_inicio_dt = pd.to_datetime(data_inicio)
data_fim_dt = pd.to_datetime(data_fim)
primeiro_dia_mes = data_fim_dt.replace(day=1)

# ----------- Filtra para colunas diárias ----------
df_filtrado = df_vendas[
    (df_vendas["Data"] >= data_inicio_dt) &
    (df_vendas["Data"] <= data_fim_dt)
]

df_agrupado_dias = df_filtrado.groupby(["Data", "Loja", "Grupo"], as_index=False)["Fat.Total"].sum()

df_pivot = df_agrupado_dias.pivot_table(
    index=["Grupo", "Loja"],
    columns="Data",
    values="Fat.Total",
    aggfunc="sum",
    fill_value=0
).reset_index()

# Renomeia colunas de data com prefixo "Fat Total"
df_pivot.columns = [
    col if isinstance(col, str) else f"Fat Total {col.strftime('%d/%m/%Y')}"
    for col in df_pivot.columns
]

# ----------- Calcula acumulado do mês até data final ----------
df_mes = df_vendas[
    (df_vendas["Data"] >= primeiro_dia_mes) &
    (df_vendas["Data"] <= data_fim_dt)
]

df_acumulado = df_mes.groupby(["Grupo", "Loja"], as_index=False)["Fat.Total"].sum()
nome_coluna_acumulado = f"Acumulado Mês (01/{data_fim_dt.strftime('%m')} até {data_fim_dt.strftime('%d/%m')})"
df_acumulado = df_acumulado.rename(columns={"Fat.Total": nome_coluna_acumulado})

# ----------- Junta diário + acumulado ----------
df_final = df_pivot.merge(df_acumulado, on=["Grupo", "Loja"], how="left")

# Ordena colunas: Grupo, Loja, dias..., acumulado
colunas_chave = ["Grupo", "Loja"]
# Extrai data da string "Fat Total dd/mm/aaaa" para ordenar corretamente
def extrair_data(col):
    return datetime.strptime(col.replace("Fat Total ", ""), "%d/%m/%Y")

colunas_dias = sorted(
    [col for col in df_pivot.columns if col not in colunas_chave],
    key=extrair_data
)
colunas_finais = colunas_chave + colunas_dias + [nome_coluna_acumulado]
df_final = df_final[colunas_finais]

# ----------- Total geral e subtotal por grupo -----------

# Calcula total geral e salva linha
total_geral_dict = {
    "Grupo": "TOTAL",
    "Loja": "",
}
total_geral_dict.update(df_final.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True).to_dict())
linha_total = pd.DataFrame([total_geral_dict])

# Remove total para inserirmos depois
df_sem_total = df_final.copy()

# Ordena por grupo e loja
df_sem_total = df_sem_total.sort_values(by=["Grupo", "Loja"])

# Monta blocos com subtotais
linhas_com_subtotais = []

for grupo, grupo_df in df_sem_total.groupby("Grupo"):
    linhas_com_subtotais.append(grupo_df)

    subtotal = grupo_df.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
    subtotal["Grupo"] = grupo
    subtotal["Loja"] = "SUBTOTAL"

    linhas_com_subtotais.append(pd.DataFrame([subtotal]))

# Junta subtotais e reinsere total geral no topo
df_final_com_subtotal = pd.concat(linhas_com_subtotais, ignore_index=True)
df_final_com_subtotal = pd.concat([linha_total, df_final_com_subtotal], ignore_index=True)


# ================================
# 7. Exibição final
# ================================
st.markdown("### 📊 Resumo por Loja - Coluna por Dia + Acumulado do Mês")
st.dataframe(
    df_final.style.format({col: "R$ {:,.2f}" for col in df_final.columns if col not in ["Grupo", "Loja"]}),
    use_container_width=True,
    height=600
)
