# pages/PainelResultados.py
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
from calendar import monthrange


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
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Rateio</h1>
    </div>
""", unsafe_allow_html=True)

  # 🎨 Remove fundo e borda do multiselect, deixa só texto preto
st.markdown("""
    <style>
    /* Itens selecionados */
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] {
        background-color: transparent !important; /* sem fundo */
        border: none !important;                   /* sem borda */
        color: black !important;                   /* texto preto */
    }
    /* Texto e ícone dentro do item selecionado */
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] * {
        color: black !important;
        fill: black !important;                    /* ícone X */
    }
    /* Área interna do multiselect sem fundo */
    div[data-testid="stMultiSelect"] > div {
        background-color: transparent !important;
    }
    </style>
""", unsafe_allow_html=True)
# ================================
# 3. Separação em ABAS
# ================================


# ================================
# Aba 3: Relatórios Vendas
# ================================


import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Carrega dados
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_vendas = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())

# Normalização
df_empresa["Loja"] = df_empresa["Loja"].str.strip().str.upper()
df_empresa["Grupo"] = df_empresa["Grupo"].str.strip()
df_vendas.columns = df_vendas.columns.str.strip()
df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()
df_vendas["Grupo"] = df_vendas["Grupo"].astype(str).str.strip()

# Merge com Tipo
df_vendas = df_vendas.merge(df_empresa[["Loja", "Tipo"]], on="Loja", how="left")

# Ajusta Fat.Total
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


# ==== Filtros lado a lado ====
col1, col2 = st.columns([1, 2])

with col1:
    tipos_disponiveis = sorted(df_vendas["Tipo"].dropna().unique())
    tipos_disponiveis.insert(0, "Todos")
    tipo_selecionado = st.selectbox("🏪 Tipo:", options=tipos_disponiveis, index=0)

with col2:
    df_vendas["Mes/Ano"] = df_vendas["Data"].dt.strftime("%m/%Y")
    meses_disponiveis = sorted(
        df_vendas["Mes/Ano"].unique(),
        key=lambda x: datetime.strptime("01/" + x, "%d/%m/%Y")
    )
    meses_selecionados = st.multiselect(
        "🗓️ Selecione os meses:",
        options=meses_disponiveis,
        default=[datetime.today().strftime("%m/%Y")]
    )

df_filtrado = df_vendas[df_vendas["Mes/Ano"].isin(meses_selecionados)]
df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%m/%Y")

if tipo_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_selecionado]

# ==== Agrupamento por Tipo + Grupo ====
chaves = ["Tipo", "Grupo"]
df_agrupado = df_filtrado.groupby(chaves + ["Período"], as_index=False)["Fat.Total"].sum()

# Agrupa direto sem criar coluna de mês/ano
df_final = df_agrupado.groupby(["Tipo", "Grupo"], as_index=False)["Fat.Total"].sum()

# Renomeia para "Total"
df_final.rename(columns={"Fat.Total": "Total"}, inplace=True)

# 🔹 Garante que a coluna Rateio comece zerada
df_final["Rateio"] = 0.0


# Calcula % Total
total_geral = df_final["Total"].sum()
df_final["% Total"] = df_final["Total"] / total_geral

# ==== Ordenação ====

subtotais_tipo = df_final.groupby("Tipo")["Total"].sum().reset_index()
subtotais_tipo = subtotais_tipo.sort_values(by="Total", ascending=False)
ordem_tipos = subtotais_tipo["Tipo"].tolist()

df_final["ord_tipo"] = df_final["Tipo"].apply(lambda x: ordem_tipos.index(x) if x in ordem_tipos else 999)
df_final = df_final.sort_values(by=["ord_tipo", "Total"], ascending=[True, False]).drop(columns="ord_tipo")
# ==== Monta subtotais por Tipo ====
linhas_com_subtotal = []
for tipo in ordem_tipos:
    bloco_tipo = df_final[df_final["Tipo"] == tipo].copy()
    linhas_com_subtotal.append(bloco_tipo)
    subtotal = bloco_tipo.drop(columns=["Tipo", "Grupo"]).sum(numeric_only=True)
    subtotal["Tipo"] = tipo
    subtotal["Grupo"] = f"Subtotal {tipo}"
    linhas_com_subtotal.append(pd.DataFrame([subtotal]))
df_final = pd.concat(linhas_com_subtotal, ignore_index=True)

# ==== Linha TOTAL no topo ====
apenas_grupos = df_final[~df_final["Grupo"].str.startswith("Subtotal", na=False)]
linha_total = apenas_grupos.drop(columns=["Tipo", "Grupo"]).sum(numeric_only=True)
linha_total["Tipo"] = ""
linha_total["Grupo"] = "TOTAL"
df_final = pd.concat([pd.DataFrame([linha_total]), df_final], ignore_index=True)


# ==== Percentual apenas para grupos ====
total_geral = df_final.loc[~df_final["Grupo"].str.startswith("Subtotal", na=False) &
                           (df_final["Grupo"] != "TOTAL"), "Total"].sum()
df_final["% Total"] = df_final.apply(
    lambda row: f"{(row['Total']/total_geral):.2%}" 
    if (not row["Grupo"].startswith("Subtotal") and row["Grupo"] != "TOTAL" and total_geral > 0) else "",
    axis=1
)

# ==== Inputs de Rateio por Tipo (lado a lado) ====
# usa a ordem já calculada; se não existir, cria a partir do df_final
tipos_base = ordem_tipos if 'ordem_tipos' in locals() else \
    [t for t in df_final["Tipo"].dropna().unique() if str(t).strip() != ""]

# evita itens vazios e mantém só Tipos válidos
tipos_unicos = [t for t in tipos_base if str(t).strip() != ""]

# quantos campos por linha (ajuste para 2, 3, 4...)
COLS_POR_LINHA = 3

# dicionário com os valores digitados (vamos usar na próxima etapa)
valores_rateio_por_tipo = {}

for i in range(0, len(tipos_unicos), COLS_POR_LINHA):
    linha = tipos_unicos[i:i+COLS_POR_LINHA]
    cols = st.columns(len(linha))
    for c, tipo in zip(cols, linha):
        with c:
            valores_rateio_por_tipo[tipo] = st.number_input(
                f"💰 Rateio — {tipo}",
                min_value=0.0, step=1000.0, format="%.2f",
                key=f"rateio_{tipo}"
            )

# ==== Preenche a coluna Rateio proporcional ao % Total (por Tipo) ====
# ==== Rateio = valor digitado do Tipo × % Total da própria linha ====

# 1) Garante uma versão numérica do % (ex.: "42,15%" -> 0.4215)
df_final["perc_num"] = df_final["% Total"].apply(
    lambda x: pd.to_numeric(str(x).replace("%", "").replace(",", "."), errors="coerce") / 100
)

# 2) Começa com Rateio zerado
df_final["Rateio"] = 0.0

# 3) Só calcula nas linhas "normais" (ignora TOTAL e Subtotal)
mask_validas = (~df_final["Grupo"].str.startswith("Subtotal", na=False)) & (df_final["Grupo"] != "TOTAL")

# 4) Para cada Tipo, aplica: Rateio_linha = valor_digitado_do_tipo × perc_num_da_linha
for tipo, valor_digitado in valores_rateio_por_tipo.items():
    if valor_digitado and valor_digitado > 0:
        mask = (df_final["Tipo"] == tipo) & mask_validas
        df_final.loc[mask, "Rateio"] = valor_digitado * df_final.loc[mask, "perc_num"]

# 5) Mantém TOTAL/Subtotal com 0
df_final.loc[~mask_validas, "Rateio"] = 0.0

# 6) Limpa a auxiliar
df_final.drop(columns=["perc_num"], inplace=True)


# ==== Reordenar colunas ====
# ==== Reordenar colunas (Rateio no fim) ====
colunas_finais = ["Tipo", "Grupo", "Total", "% Total", "Rateio"]
df_final = df_final[colunas_finais]
# ==== Função de formatação ====

# ==== Cópia para exibição com formatação ====
df_view = df_final.copy()

def formatar(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

# Formata todas as colunas numéricas (inclusive Rateio)
for col in ["Total", "Rateio"]:
    if col in df_view.columns:
        df_view[col] = df_view[col].apply(
            lambda x: formatar(x) if pd.notnull(x) and x != "" else x
        )
# ==== Estilo ====
def aplicar_estilo(df):
    def estilo_linha(row):
        if row["Grupo"] == "TOTAL":
            return ["background-color: #f4b084; font-weight: bold"] * len(row)
        elif "Subtotal" in str(row["Grupo"]):
            return ["background-color: #d9d9d9; font-weight: bold"] * len(row)
        else:
            return ["" for _ in row]
    return df.style.apply(estilo_linha, axis=1)

# Exibe a cópia formatada
st.dataframe(aplicar_estilo(df_view), use_container_width=True, height=700)
# ==== Exporta Excel ====
df_exportar = df_final.copy()
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_exportar.to_excel(writer, index=False, sheet_name="Relatório")
output.seek(0)

wb = load_workbook(output)
ws = wb["Relatório"]

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="305496")
center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

for cell in ws[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_alignment
    cell.border = border

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
    grupo_valor = row[1].value
    estilo_fundo = None
    if isinstance(grupo_valor, str):
        if grupo_valor.strip().upper() == "TOTAL":
            estilo_fundo = PatternFill("solid", fgColor="F4B084")
        elif "SUBTOTAL" in grupo_valor.strip().upper():
            estilo_fundo = PatternFill("solid", fgColor="D9D9D9")
    for cell in row:
        cell.border = border
        cell.alignment = center_alignment
        if estilo_fundo:
            cell.fill = estilo_fundo
        col_name = ws.cell(row=1, column=cell.column).value
        if isinstance(cell.value, (int, float)):
            if col_name == "% Total":
                cell.number_format = "0.000%"
            else:
                cell.number_format = '"R$" #,##0.00'

for i, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
    max_length = 0
    for cell in col_cells:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[get_column_letter(i)].width = max_length + 2

for col_nome in ["Tipo", "Grupo"]:
    if col_nome in df_exportar.columns:
        col_idx = df_exportar.columns.get_loc(col_nome) + 1
        for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for c in cell:
                c.alignment = Alignment(horizontal="left")

output_final = BytesIO()
wb.save(output_final)
output_final.seek(0)

st.download_button(
    label="📥 Baixar Excel",
    data=output_final,
    file_name="Resumo_Grupos_Mensal.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
