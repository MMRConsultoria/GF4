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
    import streamlit as st
    from datetime import datetime

    st.markdown("""
    <style>
    /* Remove aparência de "caixa" do chip selecionado */
    .stMultiSelect [data-baseweb="tag"] {
        background-color: transparent !important;
        color: black !important;
        font-weight: 500 !important;
        border: none !important;
        box-shadow: none !important;
        padding: 2px 6px !important;
        margin: 2px 4px !important;
    }
    
    /* Remove o fundo da área de seleção */
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

    # ==== Filtros principais ====
    # ==== Filtros principais ====
    data_min = df_vendas["Data"].min()
    data_max = df_vendas["Data"].max()
    
    col1, col2, col3, col4 = st.columns([2, 2, 2, 2])
    
    with col1:
        modo_exibicao = st.selectbox("🔀 Ver por:", ["Loja", "Grupo"], key="modo_exibicao_relatorio")
    
    # Filtro de Loja ou Grupo (apenas define, será aplicado depois)
    todos = "Todas"
    opcoes_filtro = sorted(df_vendas[modo_exibicao].dropna().unique())
    opcoes_filtro.insert(0, todos)
    
    with col2:
        selecao = st.selectbox(
            f"🎯 Selecione {modo_exibicao.lower()} (opcional):",
            options=opcoes_filtro,
            index=0,
            key="filtro_unico_loja_grupo"
        )
    
    with col3:
        modo_periodo = st.selectbox("🕒 Período:", ["Diário", "Mensal", "Anual"], key="modo_periodo_relatorio")

    with col4:
        tipos_disponiveis = sorted(df_vendas["Tipo"].dropna().unique())
        tipos_disponiveis.insert(0, "Todos")
        tipo_selecionado = st.selectbox("🏪 Tipo:", options=tipos_disponiveis, index=0)

    
    
    # ==== Filtro por período ====
    # ==== Filtro por período ====
    if modo_periodo == "Diário":
        # 📅 Intervalo com validação
            datas_selecionadas = st.date_input(
                "📅 Intervalo de datas:",
                value=(data_max, data_max),
                min_value=data_min,
                max_value=data_max,
                key="data_vendas_relatorio"
            )
        
            if isinstance(datas_selecionadas, (tuple, list)) and len(datas_selecionadas) == 2:
                data_inicio, data_fim = datas_selecionadas
            else:
                st.warning("⚠️ Por favor, selecione um intervalo com **duas datas** (início e fim).")
                st.stop()
        
            df_filtrado = df_vendas[
                (df_vendas["Data"] >= pd.to_datetime(data_inicio)) &
                (df_vendas["Data"] <= pd.to_datetime(data_fim))
            ]
            # ==== Aplica filtro de Tipo ====
            if tipo_selecionado != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_selecionado]

            df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%d/%m/%Y")

    
    elif modo_periodo == "Mensal":
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

    
    elif modo_periodo == "Anual":
        df_vendas["Ano"] = df_vendas["Data"].dt.strftime("%Y")
        anos_disponiveis = sorted(df_vendas["Ano"].unique())
        anos_selecionados = st.multiselect(
            "📅 Selecione os anos:",
            options=anos_disponiveis,
            default=[datetime.today().strftime("%Y")]
        )
        df_filtrado = df_vendas[df_vendas["Ano"].isin(anos_selecionados)]
        df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%Y")

        if tipo_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_selecionado]

    
    # ==== Aplica filtro de Loja ou Grupo ====
    if selecao != todos:
        df_filtrado = df_filtrado[df_filtrado[modo_exibicao] == selecao]

    
        

    # Agrupamento
    chaves = ["Tipo", "Loja", "Grupo"] if modo_exibicao == "Loja" else ["Tipo", "Grupo"]
    df_agrupado = df_filtrado.groupby(chaves + ["Período"], as_index=False)["Fat.Total"].sum()

    # Pivot
    df_pivot = df_agrupado.pivot_table(index=chaves, columns="Período", values="Fat.Total", fill_value=0).reset_index()

    # Ordenação das datas
    def ordenar_datas(col):
        try:
            return datetime.strptime(col, "%d/%m/%Y")
        except:
            try:
                return datetime.strptime("01/" + col, "%d/%m/%Y")
            except:
                return datetime.strptime("01/01/" + col, "%d/%m/%Y")

    colunas_periodo = sorted(
        [c for c in df_pivot.columns if c not in ["Tipo", "Loja", "Grupo"]],
        key=ordenar_datas
    )

    # Garante que colunas existam
    if "Loja" not in df_pivot.columns:
        df_pivot["Loja"] = ""
    if "Grupo" not in df_pivot.columns:
        df_pivot["Grupo"] = ""

    # Define ordem final: Grupo, Loja, depois períodos
    colunas_finais = ["Tipo", "Grupo", "Loja"] + colunas_periodo
    df_final = df_pivot[colunas_finais].copy()

    # Total acumulado (última coluna)
    ultima_coluna_valor = colunas_periodo[-1]
    df_final["__ordem"] = df_final[ultima_coluna_valor]

    # Ordena do maior para o menor
    df_final = df_final.sort_values(by="__ordem", ascending=False).drop(columns="__ordem").reset_index(drop=True)

    # === Linha de Lojas Ativas (quantas lojas venderam algo em cada período) ===
    # === Linha de Lojas Ativas (quantas lojas venderam por período) ===
    df_lojas_por_periodo = df_filtrado.groupby("Período")["Loja"].nunique()

    # 🔢 Soma total de todas as colunas de valor
    colunas_valor = colunas_periodo  # já está definido corretamente
    soma_total_geral = df_final[colunas_valor].sum(numeric_only=True).sum()
    
    # ➕ Cria uma coluna auxiliar com o total por linha
    df_final["__soma_linha"] = df_final[colunas_valor].sum(axis=1, numeric_only=True)
    
    # 🎯 Calcula o percentual total com base na soma de todas as colunas
    df_final["% Total"] = df_final["__soma_linha"] / soma_total_geral
    
    # 🧹 Remove a auxiliar
    df_final = df_final.drop(columns=["__soma_linha"])


    
    # Monta linha com estrutura do df_final, já convertendo os valores para str (evita float!)
    linha_lojas = {col: "" for col in df_final.columns}
    linha_lojas["Grupo"] = "Lojas Ativas"
    linha_lojas["Loja"] = ""
    linha_lojas["Tipo"] = ""
        
    for periodo in df_lojas_por_periodo.index:
        if periodo in linha_lojas:
            linha_lojas[periodo] = str(int(df_lojas_por_periodo[periodo]))  # 👈 força int e depois str
    
    # === Linha TOTAL ===
    linha_total = df_final.drop(columns=["Grupo", "Loja", "Tipo"]).sum(numeric_only=True)
    linha_total["Grupo"] = "TOTAL"
    linha_total["Loja"] = ""
    linha_total["Tipo"] = ""
    
    # Junta tudo: Lojas Ativas → TOTAL → Dados
    df_final = pd.concat([
        pd.DataFrame([linha_lojas]),
        pd.DataFrame([linha_total]),
        df_final
    ], ignore_index=True)

    # Formatação
    def formatar(valor):
        try:
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return valor

    # === Formatação final ===
    df_formatado = df_final.copy()
    
    for col in colunas_periodo:
        if col in df_formatado.columns:
            # Só aplica formatação em valores monetários (não em 'Lojas Ativas')
            df_formatado[col] = df_formatado.apply(
                lambda row: formatar(row[col]) if row["Grupo"] not in ["Lojas Ativas"] else row[col],
                axis=1
            )
    df_formatado = df_formatado[["Tipo", "Grupo", "Loja"] + colunas_periodo + ["% Total"]]
    #df_formatado = df_formatado[["Tipo", "Grupo", "Loja"] + colunas_periodo]
    # Formata a nova coluna de percentual
    df_percentual = pd.to_numeric(df_final["% Total"], errors="coerce")
    df_formatado["% Total"] = df_percentual.apply(lambda x: f"{x:.2%}" if pd.notnull(x) else "")

    # Estilo para destacar TOTAL
    def aplicar_estilo(df):
        def estilo_linha(row):
            if row["Grupo"] == "TOTAL":
                return ["background-color: #f0f0f0; font-weight: bold"] * len(row)
            elif row["Grupo"] == "Lojas Ativas":
                return ["background-color: #eeeeee; font-style: italic"] * len(row)  # ⛔ aqui está o destaque
            else:
                return ["" for _ in row]
        return df.style.apply(estilo_linha, axis=1)

    # Se for exibição por Grupo, remove a coluna "Loja" antes de aplicar estilo
    df_visivel = df_formatado.drop(columns=["Loja"]) if modo_exibicao == "Grupo" else df_formatado
    
    # Aplica estilo na nova tabela visual
    tabela_final = aplicar_estilo(df_visivel)
    
    # Exibe
    st.dataframe(
        tabela_final,
        use_container_width=True,
        height=700
    )
    from io import BytesIO
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    # ➤ Usa a mesma lógica da visualização: remove "Loja" se for modo Grupo
    df_exportar = df_final.drop(columns=["Loja"]) if modo_exibicao == "Grupo" else df_final.copy()
    
    # ➤ Exporta para Excel com BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exportar.to_excel(writer, index=False, sheet_name="Relatório")
    output.seek(0)
    
    # ➤ Reabre com openpyxl para aplicar formatação
    wb = load_workbook(output)
    ws = wb["Relatório"]
    
    # === Estilos ===
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")  # Azul escuro
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    
    # ➤ Aplica estilo no cabeçalho
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border
    
    # ➤ Detecta nome da última coluna (geralmente "% Total")
    ultima_coluna_nome = df_exportar.columns[-1]
    
    # ➤ Aplica formatação e estilo nas células
    # ➤ Aplica formatação e estilo nas células
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        grupo_valor = row[1].value  # Coluna "Grupo" (segunda coluna)
        estilo_fundo = None
    
        if isinstance(grupo_valor, str):
            if grupo_valor.strip().upper() == "TOTAL":
                estilo_fundo = PatternFill("solid", fgColor="F4B084")  # Laranja escuro
            elif grupo_valor.strip().upper() == "LOJAS ATIVAS":
                estilo_fundo = PatternFill("solid", fgColor="D9D9D9")  # Cinza claro
    
        for cell in row:
            cell.border = border
            cell.alignment = center_alignment
            if estilo_fundo:
                cell.fill = estilo_fundo
    
            col_name = ws.cell(row=1, column=cell.column).value  # Nome da coluna
    
            if isinstance(cell.value, (int, float)):
                if col_name == "% Total":
                    cell.number_format = "0.000%"
                else:
                    cell.number_format = '"R$" #,##0.00'

    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
    # ➤ Ajusta a largura das colunas automaticamente
    from openpyxl.utils import get_column_letter

    # ➤ Ajusta a largura das colunas com base no valor formatado
    for i, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    cell_str = str(cell.value)
                    max_length = max(max_length, len(cell_str))
            except:
                pass
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = max_length + 2




    # ➤ Alinha à esquerda as colunas "Tipo", "Grupo" e "Loja"
    colunas_df = list(df_exportar.columns)
    colunas_esquerda = ["Tipo", "Grupo", "Loja"]
    for col_nome in colunas_esquerda:
        if col_nome in colunas_df:
            col_idx = colunas_df.index(col_nome) + 1
            for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for c in cell:
                    c.alignment = Alignment(horizontal="left")
    
    # ➤ Define largura fixa da coluna "Loja" para 22
    if "Loja" in colunas_df:
        col_idx_loja = colunas_df.index("Loja")
        col_letra_loja = chr(ord("A") + col_idx_loja)
        ws.column_dimensions[col_letra_loja].width = 22


    
    # ➤ Salva para download
    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    
    # ➤ Botão de download
    st.download_button(
        label="📥 Baixar Excel",
        data=output_final,
        file_name="Relatorio_Vendas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
