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
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatórios</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 3. Separação em ABAS
# ================================
(tab_rateio,) = st.tabs(["📋 Rateio"])  # ✅ desembrulha a única aba

# ================================
# Aba 3: Relatórios Vendas
# ================================
with tab_rateio:
    import pandas as pd
    import numpy as np
    from datetime import datetime

    # 🎨 Estilo de selectbox/multiselect
    st.markdown("""
    <style>
    .stMultiSelect [data-baseweb="tag"] {
        background-color: transparent !important;
        color: black !important;
        font-weight: 500 !important;
        border: none !important;
        box-shadow: none !important;
        padding: 2px 6px !important;
        margin: 2px 4px !important;
    }
    .stMultiSelect > div {
        background-color: transparent !important;
        border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

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
    df_vendas = df_vendas.merge(
        df_empresa[["Loja", "Tipo"]],
        on="Loja",
        how="left"
    )    

    # Tratamento do Fat.Total
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

    # ==== Filtro de Tipo ====
    tipos_disponiveis = sorted(df_vendas["Tipo"].dropna().unique())
    tipos_disponiveis.insert(0, "Todos")
    tipo_selecionado = st.selectbox("🏪 Tipo:", options=tipos_disponiveis, index=0)

    # ==== Período fixo: Mensal ====
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

    # ==== Agrupamento fixo ====
    chaves = ["Tipo", "Grupo", "Loja"]
    df_agrupado = df_filtrado.groupby(chaves + ["Período"], as_index=False)["Fat.Total"].sum()

    # ==== Pivot ====
    df_pivot = df_agrupado.pivot_table(
        index=chaves,
        columns="Período",
        values="Fat.Total",
        fill_value=0
    ).reset_index()

    # ==== Ordenação das colunas de período ====
    def ordenar_datas(col):
        return datetime.strptime("01/" + col, "%d/%m/%Y")

    colunas_periodo = sorted(
        [c for c in df_pivot.columns if c not in chaves],
        key=ordenar_datas
    )

    # ==== Ordem final das colunas ====
    df_final = df_pivot[chaves + colunas_periodo].copy()

    # ==== Cálculo % Total ====
    soma_total_geral = df_final[colunas_periodo].sum(numeric_only=True).sum()
    df_final["__soma_linha"] = df_final[colunas_periodo].sum(axis=1, numeric_only=True)
    df_final["% Total"] = df_final["__soma_linha"] / soma_total_geral
    df_final = df_final.drop(columns=["__soma_linha"])

    # ==== Linha TOTAL ====
    linha_total = df_final.drop(columns=["Tipo", "Grupo", "Loja"]).sum(numeric_only=True)
    linha_total["Tipo"] = ""
    linha_total["Grupo"] = "TOTAL"
    linha_total["Loja"] = ""
    df_final = pd.concat([pd.DataFrame([linha_total]), df_final], ignore_index=True)

    # ==== Formatação visual ====
    def formatar(valor):
        try:
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return valor

    df_formatado = df_final.copy()
    for col in colunas_periodo:
        if col in df_formatado.columns:
            df_formatado[col] = df_formatado[col].apply(lambda x: formatar(x) if pd.notnull(x) else "")

    df_formatado["% Total"] = pd.to_numeric(df_final["% Total"], errors="coerce").apply(
        lambda x: f"{x:.2%}" if pd.notnull(x) else ""
    )

    # ==== Estilo TOTAL ====
    def aplicar_estilo(df):
        def estilo_linha(row):
            if row["Grupo"] == "TOTAL":
                return ["background-color: #f0f0f0; font-weight: bold"] * len(row)
            else:
                return ["" for _ in row]
        return df.style.apply(estilo_linha, axis=1)

    tabela_final = aplicar_estilo(df_formatado)

    st.dataframe(
        tabela_final,
        use_container_width=True,
        height=700
    )

    # ==== Exportação para Excel ====
    from io import BytesIO
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

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
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # Cabeçalho
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border

    # Estilo linhas
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        grupo_valor = row[1].value
        estilo_fundo = None
        if isinstance(grupo_valor, str) and grupo_valor.strip().upper() == "TOTAL":
            estilo_fundo = PatternFill("solid", fgColor="F4B084")
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

    # Ajusta largura das colunas
    for i, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        max_length = 0
        for cell in col_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(i)].width = max_length + 2

    # Alinha à esquerda colunas de texto
    colunas_df = list(df_exportar.columns)
    colunas_esquerda = ["Tipo", "Grupo", "Loja"]
    for col_nome in colunas_esquerda:
        if col_nome in colunas_df:
            col_idx = colunas_df.index(col_nome) + 1
            for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for c in cell:
                    c.alignment = Alignment(horizontal="left")

    # Salva final
    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)

    st.download_button(
        label="📥 Baixar Excel",
        data=output_final,
        file_name="Relatorio_Vendas_Mensal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
