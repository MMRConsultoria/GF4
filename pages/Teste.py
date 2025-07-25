import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
from calendar import monthrange

# Configuração
st.set_page_config(page_title="Vendas Diárias", layout="wide")
if not st.session_state.get("acesso_liberado"):
    st.stop()

# Autenticação
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# Carrega dados
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

# Base combinada com 0s
df_lojas_grupos = df_empresa[["Loja", "Grupo"]].drop_duplicates()
df_base_completa = pd.MultiIndex.from_product(
    [df_lojas_grupos["Loja"], datas_periodo], names=["Loja", "Data"]
).to_frame(index=False)
df_base_completa = df_base_completa.merge(df_lojas_grupos, on="Loja", how="left")
df_filtro_dias = df_vendas[(df_vendas["Data"] >= data_inicio_dt) & (df_vendas["Data"] <= data_fim_dt)]
df_agrupado_dias = df_filtro_dias.groupby(["Data", "Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
df_completo = df_base_completa.merge(df_agrupado_dias, on=["Data", "Loja", "Grupo"], how="left")
df_completo["Fat.Total"] = df_completo["Fat.Total"].fillna(0)

# Pivot com datas
df_pivot = df_completo.pivot_table(
    index=["Grupo", "Loja"], columns="Data", values="Fat.Total", aggfunc="sum", fill_value=0
).reset_index()
df_pivot.columns = [
    col if isinstance(col, str) else f"Fat Total {col.strftime('%d/%m/%Y')}"
    for col in df_pivot.columns
]

# Acumulado do mês
df_mes = df_vendas[(df_vendas["Data"] >= primeiro_dia_mes) & (df_vendas["Data"] <= data_fim_dt)]
df_acumulado = df_mes.groupby(["Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
df_acumulado = df_lojas_grupos.merge(df_acumulado, on=["Loja", "Grupo"], how="left")
df_acumulado["Fat.Total"] = df_acumulado["Fat.Total"].fillna(0)
col_acumulado = f"Acumulado Mês (01/{data_fim_dt.strftime('%m')} até {data_fim_dt.strftime('%d/%m')})"
df_acumulado = df_acumulado.rename(columns={"Fat.Total": col_acumulado})
df_base = df_pivot.merge(df_acumulado, on=["Grupo", "Loja"], how="left")
df_base = df_base[df_base[col_acumulado] != 0]

# Adiciona coluna de Meta
df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
mapa_meses = {
    "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04", "MAI": "05", "JUN": "06",
    "JUL": "07", "AGO": "08", "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12"
}
df_metas["Mês"] = df_metas["Mês"].astype(str).str.strip().str.upper().map(mapa_meses)
df_metas["Ano"] = df_metas["Ano"].astype(str).str.strip()
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
mes_filtro = data_fim_dt.strftime("%m")
ano_filtro = data_fim_dt.strftime("%Y")
df_metas_filtrado = df_metas[(df_metas["Mês"] == mes_filtro) & (df_metas["Ano"] == ano_filtro)].copy()
df_base["Loja"] = df_base["Loja"].astype(str).str.strip().str.upper()
df_base = df_base.merge(df_metas_filtrado[["Loja", "Meta"]], on="Loja", how="left")
df_base["Meta"] = df_base["Meta"].fillna(0)

# %Atingido
df_base["%Atingido"] = df_base[col_acumulado] / df_base["Meta"]
df_base["%Atingido"] = df_base["%Atingido"].replace([np.inf, -np.inf], np.nan).fillna(0).round(4)

# Reordena colunas
colunas_base = ["Grupo", "Loja"]
col_diarias = sorted([col for col in df_base.columns if col.startswith("Fat Total")])
colunas_finais = colunas_base + col_diarias + [col_acumulado, "Meta", "%Atingido", "%LojaXGrupo", "%Grupo"]
for col in ["%LojaXGrupo", "%Grupo"]:
    if col not in df_base.columns:
        df_base[col] = np.nan
df_base = df_base[colunas_finais]

# Subtotais e totais
linha_total = df_base.drop(columns=colunas_base).sum(numeric_only=True)
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
df_final = pd.concat([pd.DataFrame([linha_total])] + blocos, ignore_index=True)

# Percentuais
filtro_lojas = (
    (df_final["Loja"] != "") &
    (~df_final["Grupo"].astype(str).str.startswith("SUBTOTAL")) &
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

# %Atingido final
if modo_exibicao == "Loja":
    mascara_subtotal = df_final["Grupo"].astype(str).str.startswith("SUBTOTAL")
    mascara_total = df_final["Grupo"] == "TOTAL"
    df_final.loc[mascara_subtotal | mascara_total, "%Atingido"] = (
        df_final.loc[mascara_subtotal | mascara_total, col_acumulado] /
        df_final.loc[mascara_subtotal | mascara_total, "Meta"]
    ).replace([np.inf, -np.inf], np.nan).fillna(0).round(4)
else:
    mascara_total = df_final["Grupo"] == "TOTAL"
    mascara_subtotal = df_final["Loja"].astype(str).str.startswith("Lojas:")
    df_final.loc[mascara_total | mascara_subtotal, "%Atingido"] = (
        df_final.loc[mascara_total | mascara_subtotal, col_acumulado] /
        df_final.loc[mascara_total | mascara_subtotal, "Meta"]
    ).replace([np.inf, -np.inf], np.nan).fillna(0).round(4)
    df_final.loc[~(mascara_total | mascara_subtotal), "%Atingido"] = ""

# Oculta coluna %LojaXGrupo se for modo Grupo
if modo_exibicao == "Grupo":
    colunas_exibir = [c for c in colunas_finais if c != "%LojaXGrupo"]
else:
    colunas_exibir = colunas_finais
df_final = df_final[colunas_exibir]
df_formatado = df_formatado[colunas_exibir]

# Formata valores
colunas_percentuais = ["%LojaXGrupo", "%Grupo", "%Atingido"]
def formatar(valor, col):
    try:
        if pd.isna(valor) or valor == "":
            return ""
        return f"{valor:.2%}".replace(".", ",") if col in colunas_percentuais else f"R$ {valor:,.2f}".replace(".", "X").replace(",", ".").replace("X", ",")
    except:
        return ""
df_formatado = df_final.copy()
for col in colunas_exibir:
    if col not in ["Grupo", "Loja"]:
        df_formatado[col] = df_formatado[col].apply(lambda x: formatar(x, col))

# Faturamento desejável
dia_hoje = data_fim_dt.day
dias_mes = monthrange(data_fim_dt.year, data_fim_dt.month)[1]
perc_desejavel = dia_hoje / dias_mes
meta_total = df_final.loc[df_final["Grupo"] == "TOTAL", "Meta"].values[0]
linha_desejavel = pd.DataFrame([{
    "Grupo": "",
    "Loja": f"FATURAMENTO DESEJÁVEL ATÉ {data_fim_dt.strftime('%d/%m')}",
    "%Atingido": formatar(perc_desejavel, "%Atingido")
} | {col: "" for col in df_formatado.columns if col not in ["Grupo", "Loja", "%Atingido"]}])

# Estilo visual
def aplicar_estilo_final(df, estilos_linha):
    def apply_row_style(row):
        return estilos_linha[row.name]
    return df.style.apply(apply_row_style, axis=1)

cores_alternadas = ["#dce6f1", "#d9ead3"]
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
    elif loja == "":
        estilos.append(["background-color: #f9f9f9"] * len(row))
    else:
        if grupo != grupo_atual:
            cor_idx = (cor_idx + 1) % len(cores_alternadas)
            grupo_atual = grupo
        cor = cores_alternadas[cor_idx]
        estilos.append([f"background-color: {cor}"] * len(row))
estilos_final = [["background-color: #dddddd; font-weight: bold"] * len(df_formatado.columns)] + estilos
df_exibir = pd.concat([linha_desejavel, df_formatado], ignore_index=True)

# Exibe na tela
st.dataframe(
    aplicar_estilo_final(df_exibir, estilos_final),
    use_container_width=True,
    height=750
)
