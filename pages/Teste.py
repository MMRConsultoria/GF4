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
ultimo_dia_disponivel = df_vendas["Data"].max()
data_min = df_vendas["Data"].min()
data_max = df_vendas["Data"].max()

with col1:
    data_inicio, data_fim = st.date_input("📅 Intervalo de datas:", (data_max, data_max), data_min, data_max)
with col2:
    modo_exibicao = st.selectbox("🧭 Ver por:", ["Loja", "Grupo"])
with col3:
    filtro_meta = st.selectbox("🎯 Mostrar:", ["Meta", "Sem Meta"])

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
from datetime import datetime

col_diarias = [
    col for col in df_base.columns if col.startswith("Fat Total")
]

# Extrai a data do nome da coluna e ordena corretamente
col_diarias.sort(key=lambda x: datetime.strptime(x.replace("Fat Total ", ""), "%d/%m/%Y"))
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
# Define colunas com base no filtro "Meta" ou "Sem Meta"
colunas_visiveis = ["Grupo", "Loja"] + col_diarias + [col_acumulado]

if filtro_meta == "Meta":
    colunas_visiveis += ["Meta", "%Atingido"]
elif filtro_meta == "Sem Meta":
    colunas_visiveis += ["%LojaXGrupo", "%Grupo"]

df_final = df_final[colunas_visiveis]

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
for col in colunas_visiveis:  # ✅ CORRETO
    if col not in ["Grupo", "Loja"]:
        df_formatado[col] = df_formatado[col].apply(lambda x: formatar(x, col))

# Calcula o percentual desejável até o dia selecionado
dia_hoje = data_fim_dt.day
dias_mes = monthrange(data_fim_dt.year, data_fim_dt.month)[1]
perc_desejavel = dia_hoje / dias_mes


# Faturamento desejável (com ordem correta das colunas)
linha_desejavel_dict = {}
for col in colunas_visiveis:  # ✅ CORRETO
    if col == "Grupo":
        linha_desejavel_dict[col] = ""
    elif col == "Loja":
        linha_desejavel_dict[col] = f"FATURAMENTO DESEJÁVEL ATÉ {data_fim_dt.strftime('%d/%m')}"
    elif col == "%Atingido":
        linha_desejavel_dict[col] = formatar(perc_desejavel, "%Atingido")
    else:
        linha_desejavel_dict[col] = ""
linha_desejavel = pd.DataFrame([linha_desejavel_dict])

# Estilo visual
def aplicar_estilo_final(df, estilos_linha):
    def apply_row_style(row):
        base_style = estilos_linha[row.name].copy()
        if "%Atingido" in df.columns and row.name > 0:
            try:
                valor = row["%Atingido"]
                if isinstance(valor, str) and "%" in valor:
                    valor_float = float(valor.replace("%", "").replace(",", ".")) / 100
                else:
                    valor_float = float(valor)
                if not pd.isna(valor_float):
                    if valor_float >= perc_desejavel:
                        idx = df.columns.get_loc("%Atingido")
                        base_style[idx] = base_style[idx] + "; color: green; font-weight: bold"
                    else:
                        idx = df.columns.get_loc("%Atingido")
                        base_style[idx] = base_style[idx] + "; color: red; font-weight: bold"
            except:
                pass
        return base_style
    return df.style.apply(apply_row_style, axis=1)


cores_alternadas = ["#eef4fa", "#f5fbf3"]  # azul e verde bem suaves
estilos = []
cor_idx = -1
grupo_atual = None

# RESUMO POR TIPO (antes de aplicar estilos)
df_tipo = df_empresa[["Loja", "Tipo"]].drop_duplicates()
df_tmp = df_final.copy()
df_tmp = df_tmp.merge(df_tipo, on="Loja", how="left")
df_tmp = df_tmp[df_tmp["Grupo"].astype(str).str.startswith("SUBTOTAL")].copy()

# Agrupamento
df_resumo_tipo = (
# Colunas adicionais
df_resumo_tipo["Grupo"] = df_resumo_tipo["Tipo"]
df_resumo_tipo["Loja"] = df_tmp.groupby("Tipo")["Loja"].nunique().apply(lambda x: f"Lojas: {x:02d}").values

# Preenche colunas faltantes com ""
for col in df_formatado.columns:
    if col not in df_resumo_tipo.columns:
        df_resumo_tipo[col] = ""

# Ordena colunas na mesma ordem do relatório
df_resumo_tipo = df_resumo_tipo[df_formatado.columns]

# Formata
df_resumo_tipo_formatado = df_resumo_tipo.copy()
for col in df_resumo_tipo.columns:
    if col not in ["Grupo", "Loja"]:
        df_resumo_tipo_formatado[col] = df_resumo_tipo_formatado[col].apply(lambda x: formatar(x, col))


df_linhas_visiveis = pd.concat([df_resumo_tipo_formatado, df_formatado], ignore_index=True)

for _, row in df_linhas_visiveis.iterrows():
    grupo = row["Grupo"]
    loja = row["Loja"]

    if isinstance(grupo, str) and grupo.startswith("Tipo:"):
        estilos.append(["background-color: #e6f2e6; font-weight: bold"] * len(row))  # verde pastel
    elif grupo == "TOTAL":
        estilos.append(["background-color: #f2f2f2; font-weight: bold"] * len(row))
    elif isinstance(grupo, str) and grupo.startswith("SUBTOTAL"):
        estilos.append(["background-color: #fff8dc; font-weight: bold"] * len(row))
    elif loja == "":
        estilos.append(["background-color: #fdfdfd"] * len(row))
    else:
        if grupo != grupo_atual:
            cor_idx = (cor_idx + 1) % len(cores_alternadas)
            grupo_atual = grupo
        cor = cores_alternadas[cor_idx]
        estilos.append([f"background-color: {cor}"] * len(row))




estilos_final = [["background-color: #dddddd; font-weight: bold"] * len(df_formatado.columns)] + estilos

# ================================
# ➕ Subtotais por Tipo (linha extra no topo)
# ================================
df_tipo = df_empresa[["Loja", "Tipo"]].drop_duplicates()
df_merge_tipo = df_final.merge(df_tipo, on="Loja", how="left")
df_merge_tipo = df_merge_tipo[df_merge_tipo["Loja"].astype(str).str.startswith("Lojas:") == False]
df_merge_tipo = df_merge_tipo[df_merge_tipo["Grupo"] != "TOTAL"]
df_merge_tipo = df_merge_tipo[df_merge_tipo["Grupo"].astype(str).str.startswith("SUBTOTAL") == False]

# Acumula por tipo
linhas_tipos = []
tipos_ordenados = df_merge_tipo.groupby("Tipo")[col_acumulado].sum().sort_values(ascending=False).index.tolist()

for tipo in tipos_ordenados:
    df_tipo_filtro = df_merge_tipo[df_merge_tipo["Tipo"] == tipo]
    if df_tipo_filtro.empty:
        continue

    linha = {}
    linha["Grupo"] = tipo
    linha["Loja"] = f"Lojas: {df_tipo_filtro['Loja'].nunique()}"
    for col in df_tipo_filtro.columns:
        if col in ["Grupo", "Loja", "Tipo"]:
            continue
        if pd.api.types.is_numeric_dtype(df_tipo_filtro[col]):
            linha[col] = df_tipo_filtro[col].sum()
        else:
            linha[col] = ""

    linhas_tipos.append(linha)

df_resumo_tipo = pd.DataFrame(linhas_tipos)
df_resumo_tipo_formatado = df_resumo_tipo.copy()
for col in df_resumo_tipo.columns:
    if col not in ["Grupo", "Loja"]:
        df_resumo_tipo_formatado[col] = df_resumo_tipo_formatado[col].apply(lambda x: formatar(x, col))




# 🔸 Cria estilos adicionais para as linhas extras
estilo_desejavel = [["background-color: #dddddd; font-weight: bold"] * len(df_formatado.columns)]
estilo_resumo_tipo = [["background-color: #f0f8ff; font-weight: bold"] * len(df_formatado.columns)] * len(df_resumo_tipo_formatado)

# 🔸 Recria df_exibir com as 3 partes
df_exibir = pd.concat([linha_desejavel, df_resumo_tipo_formatado, df_formatado], ignore_index=True)

# 🔸 Junta todos os estilos: desejável + resumo tipo + estilos das linhas finais
estilos_final = estilo_desejavel + estilo_resumo_tipo + estilos

# 🔸 Exibe com estilo correto
st.dataframe(
    aplicar_estilo_final(df_exibir, estilos_final),
    use_container_width=True,
    height=750
)
