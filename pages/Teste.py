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

# ================================
# Configuração e acesso
# ================================
st.set_page_config(page_title="Vendas Diárias", layout="wide")
if not st.session_state.get("acesso_liberado"):
    st.stop()

# Conexão com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# ================================
# Layout e título
# ================================
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
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
        <h1 style='margin: 0; font-size: 2.4rem;'>Relatório Vendas Diárias</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# Carrega dados
# ================================
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_vendas = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
df_empresa["Loja"] = df_empresa["Loja"].str.strip().str.upper()
df_empresa["Grupo"] = df_empresa["Grupo"].str.strip()

df_vendas.columns = df_vendas.columns.str.strip()
df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()
df_vendas["Grupo"] = df_vendas["Grupo"].astype(str).str.strip()
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
# Seleção de período
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
    st.write("🔜 Filtros adicionais em breve")

data_inicio_dt = pd.to_datetime(data_inicio)
data_fim_dt = pd.to_datetime(data_fim)
primeiro_dia_mes = data_fim_dt.replace(day=1)
datas_periodo = pd.date_range(start=data_inicio_dt, end=data_fim_dt)

# ================================
# Prepara base com todas as lojas
# ================================
df_lojas_grupos = df_empresa[["Loja", "Grupo"]].drop_duplicates()

df_base_completa = pd.MultiIndex.from_product(
    [df_lojas_grupos["Loja"], datas_periodo],
    names=["Loja", "Data"]
).to_frame(index=False)
df_base_completa = df_base_completa.merge(df_lojas_grupos, on="Loja", how="left")

df_filtro_dias = df_vendas[
    (df_vendas["Data"] >= data_inicio_dt) & (df_vendas["Data"] <= data_fim_dt)
]
df_agrupado_dias = df_filtro_dias.groupby(["Data", "Loja", "Grupo"], as_index=False)["Fat.Total"].sum()

df_completo = df_base_completa.merge(df_agrupado_dias, on=["Data", "Loja", "Grupo"], how="left")
df_completo["Fat.Total"] = df_completo["Fat.Total"].fillna(0)

# ================================
# Pivot diário
# ================================
df_pivot = df_completo.pivot_table(
    index=["Grupo", "Loja"],
    columns="Data",
    values="Fat.Total",
    aggfunc="sum",
    fill_value=0
).reset_index()
df_pivot.columns = [
    col if isinstance(col, str) else f"Fat Total {col.strftime('%d/%m/%Y')}"
    for col in df_pivot.columns
]

# ================================
# Acumulado do mês
# ================================
df_mes = df_vendas[
    (df_vendas["Data"] >= primeiro_dia_mes) & (df_vendas["Data"] <= data_fim_dt)
]
df_acumulado = df_mes.groupby(["Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
df_acumulado = df_lojas_grupos.merge(df_acumulado, on=["Loja", "Grupo"], how="left")
df_acumulado["Fat.Total"] = df_acumulado["Fat.Total"].fillna(0)
nome_col_acumulado = f"Acumulado Mês (01/{data_fim_dt.strftime('%m')} até {data_fim_dt.strftime('%d/%m')})"
df_acumulado = df_acumulado.rename(columns={"Fat.Total": nome_col_acumulado})

df_base = df_pivot.merge(df_acumulado, on=["Grupo", "Loja"], how="left")

# ================================
# Subtotais e ordenação
# ================================
col_acumulado = nome_col_acumulado
colunas_valores = [col for col in df_base.columns if col not in ["Grupo", "Loja"]]

# Total geral
linha_total = df_base[colunas_valores].sum(numeric_only=True)
linha_total["Grupo"] = "TOTAL"
linha_total["Loja"] = ""

# Subtotais
blocos = []
grupos_info = []
for grupo, df_grp in df_base.groupby("Grupo"):
    total_grupo = df_grp[col_acumulado].sum()
    grupos_info.append((grupo, total_grupo, df_grp))

grupos_info.sort(key=lambda x: x[1], reverse=True)

for grupo, _, df_grp in grupos_info:
    df_grp_ord = df_grp.sort_values(by=col_acumulado, ascending=False)
    subtotal = df_grp_ord.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
    subtotal["Grupo"] = f"SUBTOTAL {grupo}"
    subtotal["Loja"] = ""
    blocos.append(df_grp_ord)
    blocos.append(pd.DataFrame([subtotal]))

df_final = pd.concat([pd.DataFrame([linha_total])] + blocos, ignore_index=True)

# ================================
# Estilo para exibição
# ================================
def formatar_brasileiro(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

df_formatado = df_final.copy()
df_formatado[colunas_valores] = df_formatado[colunas_valores].applymap(formatar_brasileiro)

cores_grupos = ["#dce6f1", "#d9ead3"]
estilos = []
cor_idx = -1
grupo_atual = None

for _, row in df_final.iterrows():
    grupo = row["Grupo"]
    loja = row["Loja"]
    if grupo == "TOTAL":
        estilos.append(["background-color: #eeeeee; font-weight: bold"] * len(row))
    elif isinstance(grupo, str) and grupo.startswith("SUBTOTAL"):
        estilos.append(["background-color: #ffe599; font-weight: bold"] * len(row))
        grupo_atual = None
    elif loja == "":
        estilos.append(["background-color: #f9f9f9"] * len(row))
    else:
        if grupo != grupo_atual:
            cor_idx = (cor_idx + 1) % len(cores_grupos)
            grupo_atual = grupo
        cor = cores_grupos[cor_idx]
        estilos.append([f"background-color: {cor}"] * len(row))

# ================================
# Exibe
# ================================
st.markdown("### 📊 Relatório Final com Estilo")
st.dataframe(
    df_formatado.style.apply(lambda _: estilos, axis=None),
    use_container_width=True,
    height=750
)
