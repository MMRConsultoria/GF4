import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime

# Configuração
st.set_page_config(page_title="Vendas Diárias", layout="wide")
if not st.session_state.get("acesso_liberado"):
    st.stop()

# Conexão com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# Carrega dados
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_vendas = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
# ================================
# Carrega e trata a aba de Metas
# ================================
df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())

# Padroniza colunas da aba Metas
df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
df_metas["Mês"] = df_metas["Mês"].astype(str).str.zfill(2)
df_metas["Ano"] = df_metas["Ano"].astype(str).str.strip()

# Trata a coluna de valor da Meta
df_metas["Meta"] = (
    df_metas["Meta"]
    .astype(str)
    .str.replace("R$", "", regex=False)
    .str.replace("(", "-", regex=False)
    .str.replace(")", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
)
df_metas["Meta"] = pd.to_numeric(df_metas["Meta"], errors="coerce").fillna(0)


# Filtros
data_min = df_vendas["Data"].min()
data_max = df_vendas["Data"].max()
col1, col2, col3 = st.columns([2, 2, 2])
with col1:
    data_inicio, data_fim = st.date_input("📅 Intervalo de datas:", (data_max, data_max), data_min, data_max)
with col2:
    modo_exibicao = st.selectbox("🧭 Ver por:", ["Loja", "Grupo"])
with col3:
    st.write(" ")

data_inicio_dt = pd.to_datetime(data_inicio)
data_fim_dt = pd.to_datetime(data_fim)
primeiro_dia_mes = data_fim_dt.replace(day=1)
datas_periodo = pd.date_range(start=data_inicio_dt, end=data_fim_dt)

# Nome da coluna de meta
mes_extenso = data_fim_dt.strftime("%B").capitalize()
coluna_meta = f"Meta {mes_extenso}"
if coluna_meta in df_metas.columns:
    df_metas[coluna_meta] = (
        df_metas[coluna_meta]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
    )
else:
    df_metas[coluna_meta] = np.nan

# Base combinada
df_lojas_grupos = df_empresa[["Loja", "Grupo"]].drop_duplicates()
df_base_completa = pd.MultiIndex.from_product(
    [df_lojas_grupos["Loja"], datas_periodo], names=["Loja", "Data"]
).to_frame(index=False)
df_base_completa = df_base_completa.merge(df_lojas_grupos, on="Loja", how="left")
df_filtro_dias = df_vendas[(df_vendas["Data"] >= data_inicio_dt) & (df_vendas["Data"] <= data_fim_dt)]
df_agrupado_dias = df_filtro_dias.groupby(["Data", "Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
df_completo = df_base_completa.merge(df_agrupado_dias, on=["Data", "Loja", "Grupo"], how="left")
df_completo["Fat.Total"] = df_completo["Fat.Total"].fillna(0)

df_pivot = df_completo.pivot_table(
    index=["Grupo", "Loja"], columns="Data", values="Fat.Total", aggfunc="sum", fill_value=0
).reset_index()
df_pivot.columns = [
    col if isinstance(col, str) else f"Fat Total {col.strftime('%d/%m/%Y')}"
    for col in df_pivot.columns
]

df_mes = df_vendas[(df_vendas["Data"] >= primeiro_dia_mes) & (df_vendas["Data"] <= data_fim_dt)]
df_acumulado = df_mes.groupby(["Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
df_acumulado = df_lojas_grupos.merge(df_acumulado, on=["Loja", "Grupo"], how="left")
df_acumulado["Fat.Total"] = df_acumulado["Fat.Total"].fillna(0)
nome_col_acumulado = f"Acumulado Mês (01/{data_fim_dt.strftime('%m')} até {data_fim_dt.strftime('%d/%m')})"
df_acumulado = df_acumulado.rename(columns={"Fat.Total": nome_col_acumulado})
df_base = df_pivot.merge(df_acumulado, on=["Grupo", "Loja"], how="left")
df_base = df_base[df_base[nome_col_acumulado] != 0]

# Subtotais e ordenação
col_acumulado = nome_col_acumulado
colunas_valores = [col for col in df_base.columns if col not in ["Grupo", "Loja"]]
linha_total = df_base[colunas_valores].sum(numeric_only=True)
linha_total["Grupo"] = "TOTAL"
linha_total["Loja"] = f"Lojas: {df_base['Loja'].nunique():02d}"

blocos = []
grupos_info = []
for grupo, df_grp in df_base.groupby("Grupo"):
    total_grupo = df_grp[col_acumulado].sum()
    grupos_info.append((grupo, total_grupo, df_grp))

grupos_info.sort(key=lambda x: x[1], reverse=True)

for grupo, _, df_grp in grupos_info:
    df_grp_ord = df_grp.sort_values(by=col_acumulado, ascending=False)
    subtotal = df_grp_ord.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
    subtotal["Grupo"] = f"{'SUBTOTAL ' if modo_exibicao == 'Loja' else ''}{grupo}"
    subtotal["Loja"] = f"Lojas: {df_grp_ord['Loja'].nunique():02d}"
    if modo_exibicao == "Loja":
        blocos.append(df_grp_ord)
    blocos.append(pd.DataFrame([subtotal]))

# Concatena tudo
df_final = pd.concat([pd.DataFrame([linha_total])] + blocos, ignore_index=True)

# ================================
# Pega mês e ano do filtro
# ================================
mes_filtro = data_fim_dt.strftime("%m")
ano_filtro = data_fim_dt.strftime("%Y")

# Filtra metas do mês/ano
df_metas_filtrado = df_metas[(df_metas["Mês"] == mes_filtro) & (df_metas["Ano"] == ano_filtro)].copy()

# Padroniza df_final['Loja'] antes do merge
df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.upper()

# Junta metas no df_final
df_final = df_final.merge(
    df_metas_filtrado[["Loja", "Meta"]],
    on="Loja",
    how="left"
)

df_final["Meta"] = df_final["Meta"].fillna(0)


# ================================
# Carrega a aba de Metas
# ================================
df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
df_metas["Mês"] = df_metas["Mês"].astype(str).str.zfill(2)
df_metas["Ano"] = df_metas["Ano"].astype(str)

# Pega mês e ano do filtro
mes_filtro = data_fim_dt.strftime("%m")
ano_filtro = data_fim_dt.strftime("%Y")

# Filtra metas do mês/ano
df_metas_filtrado = df_metas[(df_metas["Mês"] == mes_filtro) & (df_metas["Ano"] == ano_filtro)].copy()

# Trata os valores da meta
df_metas_filtrado["Meta"] = (
    df_metas_filtrado["Meta"]
    .astype(str)
    .str.replace("R$", "", regex=False)
    .str.replace("(", "-", regex=False)
    .str.replace(")", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
)
df_metas_filtrado["Meta"] = pd.to_numeric(df_metas_filtrado["Meta"], errors="coerce").fillna(0)

# Junta com df_final
df_final = df_final.merge(
    df_metas_filtrado[["Loja", "Meta"]],
    on="Loja",
    how="left"
)
df_final["Meta"] = df_final["Meta"].fillna(0)

# Percentuais
df_final["%LojaXGrupo"] = np.nan
df_final["%Grupo"] = np.nan

filtro_lojas = (
    (df_final["Loja"] != "") &
    (~df_final["Grupo"].str.startswith("SUBTOTAL")) &
    (df_final["Grupo"] != "TOTAL")
)

df_lojas_reais = df_final[filtro_lojas].copy()
soma_por_grupo = df_lojas_reais.groupby("Grupo")[col_acumulado].transform("sum")
soma_total_geral = df_lojas_reais[col_acumulado].sum()

if modo_exibicao != "Grupo":
    df_final.loc[filtro_lojas, "%LojaXGrupo"] = (
        df_lojas_reais[col_acumulado].values / soma_por_grupo.values
    ).round(4)

if modo_exibicao == "Grupo":
    filtro_grupos = df_final["Loja"].astype(str).str.startswith("Lojas:")
else:
    filtro_grupos = df_final["Grupo"].astype(str).str.startswith("SUBTOTAL")

df_final.loc[filtro_grupos, "%Grupo"] = (
    df_final.loc[filtro_grupos, col_acumulado] / soma_total_geral
).round(4)

# Posiciona Meta ao lado do Acumulado
colunas_chave = ["Grupo", "Loja"]
colunas_valores = [col for col in df_final.columns if col not in colunas_chave]
if "Meta" in colunas_valores and col_acumulado in colunas_valores:
    colunas_valores.remove("Meta")
    idx = colunas_valores.index(col_acumulado)
    colunas_valores.insert(idx + 1, "Meta")
df_final = df_final[colunas_chave + colunas_valores]


# Formatação
colunas_percentuais = ["%LojaXGrupo", "%Grupo"]

def formatar_brasileiro_com_coluna(valor, coluna):
    try:
        if pd.isna(valor):
            return ""
        if coluna in colunas_percentuais:
            return f"{valor:.2%}".replace(".", ",") if valor >= 0 else ""
        else:
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return ""

df_formatado = df_final.copy()
for col in colunas_valores:
    df_formatado[col] = df_formatado[col].apply(lambda x: formatar_brasileiro_com_coluna(x, col))

# Estilo
cores_alternadas = ["#dce6f1", "#d9ead3"]
estilos = []

if modo_exibicao == "Grupo":
    for i, row in df_final.iterrows():
        if row["Grupo"] == "TOTAL":
            estilos.append(["background-color: #eeeeee; font-weight: bold"] * len(row))
        else:
            cor = cores_alternadas[i % 2]
            estilos.append([f"background-color: {cor}; font-weight: 600"] * len(row))
else:
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
                cor_idx = (cor_idx + 1) % len(cores_alternadas)
                grupo_atual = grupo
            cor = cores_alternadas[cor_idx]
            estilos.append([f"background-color: {cor}"] * len(row))

# Exibição
def aplicar_estilo_final(df, estilos_linha):
    def apply_row_style(row):
        return estilos_linha[row.name]
    return df.style.apply(apply_row_style, axis=1)

st.dataframe(
    aplicar_estilo_final(df_formatado, estilos),
    use_container_width=True,
    height=750
)
