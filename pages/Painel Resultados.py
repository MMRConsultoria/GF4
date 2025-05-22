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
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Vendas Diarias</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 3. Separação em ABAS
# ================================
aba1, aba2, aba3, aba4 = st.tabs([
    "📈 Graficos Anuais",
    "📊 Graficos Trimestrais",
    "📆 Relatório Analitico",
    "📋 Analise Lojas"
])

# ================================
# Aba 1: Graficos Anuais
# ================================
with aba1:
    planilha = gc.open("Vendas diarias")
    aba = planilha.worksheet("Fat Sistema Externo")
    dados = aba.get_all_records()
    df = pd.DataFrame(dados)
    
    # ✅ Limpa espaços invisíveis nos nomes das colunas
    df.columns = df.columns.str.strip()
    
    #st.write("🧪 Colunas carregadas:", df.columns.tolist())
    
    def limpar_valor(x):
        try:
            if isinstance(x, str):
                return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
            return float(x)
        except:
            return None

    for col in ["Fat.Total", "Serv/Tx", "Fat.Real"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_valor)

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    df["Ano"] = df["Data"].dt.year
    df["Mês"] = df["Data"].dt.month
    meses_portugues = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    df["Nome Mês"] = df["Mês"].map(meses_portugues)

    anos_disponiveis = sorted(df["Ano"].dropna().unique())
    anos_comparacao = st.multiselect(" ", options=anos_disponiveis, default=anos_disponiveis)


    if "Data" in df.columns and "Fat.Real" in df.columns and "Ano" in df.columns:
        df_anos = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"]).copy()
    else:
        st.error("❌ A aba 'Fat Sistema Externo' não contém as colunas necessárias: 'Data', 'Ano' ou 'Fat.Real'.")
        st.stop()

    
    #df_anos = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"]).copy()
    # Normalizar nomes das lojas para evitar duplicações por acento, espaço ou caixa
    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.lower()

    # Calcular a quantidade de lojas únicas por ano (com base em loja + ano únicos)
    df_lojas = df_anos.drop_duplicates(subset=["Ano", "Loja"])
    df_lojas = df_lojas.groupby("Ano")["Loja"].nunique().reset_index()
    df_lojas.columns = ["Ano", "Qtd_Lojas"]


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

    df_total = fat_mensal.groupby("Ano")["Fat.Real"].sum().reset_index()
    df_total["Ano"] = df_total["Ano"].astype(int)
    df_lojas["Ano"] = df_lojas["Ano"].astype(int)
    df_total = df_total.merge(df_lojas, on="Ano", how="left")
    df_total["AnoTexto"] = df_total.apply(
        lambda row: f"{int(row['Ano'])}       R$ {row['Fat.Real']/1_000_000:,.1f} Mi".replace(",", "."), axis=1
    )
    df_total["Ano"] = df_total["Ano"].astype(int)

    # ORDEM CORRETA dos anos de cima para baixo (mais antigo no topo)
    anos_ordenados = sorted(df_total["Ano"].unique())  # ex: [2023, 2024, 2025]
    anos_ordenados_str = [str(ano) for ano in anos_ordenados]

    # Converter a coluna "Ano" para string e categoria ordenada
    df_total["Ano"] = df_total["Ano"].astype(str)
    df_total["Ano"] = pd.Categorical(df_total["Ano"], categories=anos_ordenados_str, ordered=True)

    # Reordenar o dataframe com base na ordem correta
    df_total = df_total.sort_values("Ano", ascending=True)
    
    fig_total = px.bar(
        df_total,
        x="Fat.Real",
        y="Ano",
        orientation="h",
        color="Ano",
        text="AnoTexto",
        color_discrete_map=color_map
    )
    fig_total.update_traces(
        textposition="inside",
        textfont=dict(size=16, color="white"),
        insidetextanchor="start",
        showlegend=False
    )
    fig_total.update_traces(
        textposition="outside",
        textfont=dict(size=16),
        showlegend=False
    )
    for i, row in df_total.iterrows():
        fig_total.add_annotation(
            x=0.1,
            y=row["Ano"],
            text=row["AnoTexto"],
            showarrow=False,
            xanchor="left",
            yanchor="middle",
            font=dict(color="white", size=16),
            xref="x",
            yref="y"
        )
        fig_total.add_annotation(
            x=row["Fat.Real"],
            y=row["Ano"],
            showarrow=False,
            text=f"{int(row['Qtd_Lojas'])} Lojas",
            xanchor="left",
            yanchor="bottom",
            yshift=-8,
            font=dict(color="red", size=16, weight="bold"),
            xref="x",
            yref="y"
        )
    fig_total.update_layout(
        height=130,
        margin=dict(t=0, b=0, l=0, r=0),
        title=None,
        xaxis=dict(visible=False),
        yaxis=dict(
            categoryorder="array",
            categoryarray=anos_ordenados_str,  # ordem natural: 2023 em cima, 2025 embaixo
            showticklabels=False,
            showgrid=False,
            zeroline=False
        ),
        yaxis_title=None,
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.subheader("Faturamento Anual")
    st.plotly_chart(fig_total, use_container_width=True)
    st.markdown("---")
    st.subheader("Faturamento Mensal")
    st.plotly_chart(fig, use_container_width=True)

# ================================
# Aba 2: Graficos Trimestrais
# ================================
with aba2:
    st.info("em desenvolvimento.")

# ================================
# Aba 3: Relatorio Analitico
# ================================
with aba3:
    st.info("em desenvolvimento.")

# ================================
# Aba 4: Analise Lojas
# ================================
with aba4:
    from datetime import datetime, date
    st.markdown("""
    <style>
    /* Zera espaçamentos verticais padrão */
    div[data-testid="stVerticalBlock"] {
        gap: 0.4rem !important;
        row-gap: 0.4rem !important;
    }

    /* Alinha chips (multiselect) com campo de data */
    .stMultiSelect {
        margin-bottom: -0.6rem !important;
    }
    
    /* Estiliza os chips selecionados nos multiselects */
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #cccccc !important;  /* cinza médio */
        color: black !important;
        font-weight: 600;
        border-radius: 6px;
        padding: 4px 10px;
    }
    
    /* Reduz espaço do campo de data com os rádios */
    div[data-testid="stDateInput"] {
        margin-top: -0.4rem !important;
        margin-bottom: -0.4rem !important;
        padding-bottom: 0rem !important;
    }

    /* Elimina margens entre rádios */
    .stRadio {
        margin-top: -0.5rem !important;
        margin-bottom: -0.5rem !important;
    }

    /* Refina ainda mais os blocos invisíveis */
    section > div > div > div > div {
        margin-top: 0rem !important;
        margin-bottom: 0rem !important;
    }

    /* Zera padding entre colunas internas (radio) */
    [data-baseweb="radio"] {
        margin: 0rem !important;
    }

    /* Padding geral da página */
    .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
    }
    </style>
    """, unsafe_allow_html=True)


    # Normaliza dados
    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.lower().str.title()
    df_anos["Fat.Total"] = pd.to_numeric(df_anos["Fat.Total"], errors="coerce")
    df_anos["Fat.Real"] = pd.to_numeric(df_anos["Fat.Real"], errors="coerce")
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Mês Num"] = df_anos["Data"].dt.month
    df_anos["Mês Nome"] = df_anos["Data"].dt.strftime('%B')
    df_anos["Mês"] = df_anos["Data"].dt.strftime('%m/%Y')
    df_anos["Dia"] = df_anos["Data"].dt.strftime('%d/%m/%Y')

    # === FILTROS ===
    #anos_disponiveis = sorted(df_anos["Ano"].unique(), reverse=True)
    #ano_opcao = st.multiselect("📅 Selecione ano/mês(s):", options=anos_disponiveis, default=anos_disponiveis, key="ano_aba3")
    anos_disponiveis = sorted(df_anos["Ano"].unique(), reverse=True)
    ultimo_ano = anos_disponiveis[0] if anos_disponiveis else datetime.today().year
    ano_opcao = st.multiselect("📅 Selecione ano/mês(s):", options=anos_disponiveis, default=[ultimo_ano], key="ano_aba3")
    
    
    df_filtrado = df_anos[df_anos["Ano"].isin(ano_opcao)]

    meses_dict = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
                  7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}

    meses_disponiveis = sorted(df_filtrado["Mês Num"].unique())
    #meses_nomes_disponiveis = [meses_dict[m] for m in meses_disponiveis]
    #meses_selecionados = st.multiselect("", options=meses_nomes_disponiveis, default=meses_nomes_disponiveis, key="meses_aba3")
    meses_nomes_disponiveis = [meses_dict[m] for m in meses_disponiveis]
    mes_atual_nome = meses_dict[datetime.today().month]
    meses_selecionados = st.multiselect("", options=meses_nomes_disponiveis, default=[mes_atual_nome], key="meses_aba3")
    
    meses_numeros = [k for k, v in meses_dict.items() if v in meses_selecionados]
    df_filtrado = df_filtrado[df_filtrado["Mês Num"].isin(meses_numeros)]

   # Garantir que "hoje" seja do tipo date
    hoje = date.today()

    # Verifica se df_filtrado tem dados válidos para datas
    if not df_filtrado.empty and pd.to_datetime(df_filtrado["Data"], errors="coerce").notna().any():
        data_minima = pd.to_datetime(df_filtrado["Data"], errors="coerce").min().date()
        data_maxima = pd.to_datetime(df_filtrado["Data"], errors="coerce").max().date()

        # Garante que "hoje" esteja dentro do intervalo
        if hoje < data_minima:
            hoje = data_minima
        elif hoje > data_maxima:
            hoje = data_maxima
    else:
        data_minima = hoje
        data_maxima = hoje

    # Campo de data seguro
    data_inicio, data_fim = st.date_input(
        "",
        value=[hoje, hoje],
        min_value=data_minima,
        max_value=data_maxima
    )
    df_filtrado = df_filtrado[(df_filtrado["Data"] >= pd.to_datetime(data_inicio)) & (df_filtrado["Data"] <= pd.to_datetime(data_fim))].copy()


    # Filtros laterais lado a lado
    col1, col2, col3, col4 = st.columns([1.2, 2, 2, 2])  # col1 levemente mais estreita

    with col1:
        st.write("")  # Garante altura igual às outras colunas com título
        exibir_total = st.radio(
            " ", 
            options=[True, False],
            format_func=lambda x: "Total Sim" if x else "Total Não",
            index=0,
            horizontal=True
        )
    with col2:
        modo_visao = st.radio(" ", ["Por Loja", "Por Grupo"], horizontal=True, key="visao_aba4")

    with col3:
        tipo_metrica = st.radio(" ", ["Bruto", "Real", "Ambos"], horizontal=True, key="metrica_aba4")
    
    with col4:
         agrupamento = st.radio(" ", ["Ano", "Mês", "Dia"], horizontal=True, key="agrup_aba4")

    # Filtro para exibir ou não a coluna Total
    #exibir_total_opcao = st.radio("📊 Coluna Total:", ["Sim", "Não"], index=0, horizontal=True)
    #exibir_total = exibir_total_opcao == "Sim"

    # Criação do agrupador e ordem com base na escolha
    if agrupamento == "Ano":
        df_filtrado["Agrupador"] = df_filtrado["Ano"].astype(str)
        df_filtrado["Ordem"] = df_filtrado["Data"].dt.year

    elif agrupamento == "Mês":
        df_filtrado["Agrupador"] = df_filtrado["Data"].dt.strftime("%m/%Y")
        df_filtrado["Ordem"] = df_filtrado["Data"].dt.to_period("M").dt.to_timestamp()

    elif agrupamento == "Dia":
        df_filtrado["Agrupador"] = df_filtrado["Data"].dt.strftime("%d/%m/%Y")
        df_filtrado["Ordem"] = df_filtrado["Data"]

    # Garante a ordem correta
    ordem = (
        df_filtrado[["Agrupador", "Ordem"]]
        .drop_duplicates()
        .dropna()
        .sort_values("Ordem", ascending=False)
    )["Agrupador"].tolist()
   


    ordem = df_filtrado[["Agrupador", "Ordem"]].drop_duplicates().sort_values("Ordem", ascending=False)["Agrupador"].tolist()

    if modo_visao == "Por Grupo":
        df_grouped = df_filtrado.groupby(["Grupo", "Agrupador"]).agg(
            Bruto=("Fat.Total", "sum"),
            Real=("Fat.Real", "sum")
        ).reset_index()

        if tipo_metrica == "Bruto":
            tabela = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Bruto").fillna(0)
        elif tipo_metrica == "Real":
            tabela = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Real").fillna(0)
        else:
            tab_b = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Bruto").fillna(0)
            tab_r = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Real").fillna(0)
            tab_b.columns = [f"{c} (Bruto)" for c in tab_b.columns]
            tab_r.columns = [f"{c} (Real)" for c in tab_r.columns]
            tabela = pd.concat([tab_b, tab_r], axis=1)
            colunas_intercaladas = []
            for col in ordem:
                colunas_intercaladas.append(f"{col} (Bruto)")
                colunas_intercaladas.append(f"{col} (Real)")
            tabela = tabela[[c for c in colunas_intercaladas if c in tabela.columns]]
    else:
        tab_b = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Total", aggfunc="sum", fill_value=0)
        tab_r = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Real", aggfunc="sum", fill_value=0)
        if tipo_metrica == "Bruto":
            tabela = tab_b
        elif tipo_metrica == "Real":
            tabela = tab_r
        else:
            tab_b.columns = [f"{c} (Bruto)" for c in tab_b.columns]
            tab_r.columns = [f"{c} (Real)" for c in tab_r.columns]
            tabela = pd.concat([tab_b, tab_r], axis=1)
            colunas_intercaladas = []
            for col in ordem:
                colunas_intercaladas.append(f"{col} (Bruto)")
                colunas_intercaladas.append(f"{col} (Real)")
            tabela = tabela[[c for c in colunas_intercaladas if c in tabela.columns]]

    colunas_ordenadas = [col for col in ordem if col in tabela.columns or f"{col} (Bruto)" in tabela.columns or f"{col} (Real)" in tabela.columns]
    todas_colunas = []
    for col in colunas_ordenadas:
        if tipo_metrica == "Ambos":
            if f"{col} (Bruto)" in tabela.columns: todas_colunas.append(f"{col} (Bruto)")
            if f"{col} (Real)" in tabela.columns: todas_colunas.append(f"{col} (Real)")
        else:
            todas_colunas.append(col)
    tabela = tabela[todas_colunas]
    if tipo_metrica == "Ambos":
        cols_bruto = [col for col in tabela.columns if "(Bruto)" in col]
        cols_real = [col for col in tabela.columns if "(Real)" in col]

        if exibir_total:
            tabela["Total Bruto"] = tabela[cols_bruto].sum(axis=1)
            tabela["Total Real"] = tabela[cols_real].sum(axis=1)
            colunas_finais = ["Total Bruto", "Total Real"] + [col for col in tabela.columns if col not in ["Total Bruto", "Total Real"]]
            tabela = tabela[colunas_finais]

    
        total_row = pd.DataFrame(tabela.sum(numeric_only=True)).T
        total_row.index = ["Total Geral"]
        tabela_final = pd.concat([total_row, tabela])
    else:
        cols_validas = [col for col in tabela.columns if col != "Total"]
        if exibir_total:
            tabela["Total"] = tabela[cols_validas].sum(axis=1)
            tabela = tabela[["Total"] + cols_validas]
        else:
            tabela = tabela[cols_validas]

        
        total_geral = pd.DataFrame(tabela.sum(numeric_only=True)).T
        total_geral.index = ["Total Geral"]
        tabela_final = pd.concat([total_geral, tabela])

    quantidade = tabela.shape[0]
    nome = "Grupos" if modo_visao == "Por Grupo" else "Lojas"
    st.markdown(f"**🔢 Total de {nome}: {quantidade}**")

    # Detecta a coluna de data mais recente
    colunas_validas = [col for col in tabela_final.columns if "/" in col or (col.isdigit() and len(col) == 4)]

    def parse_col(col):
        try:
            if "/" in col:
                return pd.to_datetime(f"01/{col}", dayfirst=True, errors="coerce")
            elif col.isdigit() and len(col) == 4:
                return pd.to_datetime(f"01/01/{col}", dayfirst=True)
        except:
            return pd.NaT
        return pd.NaT

    datas_convertidas = [(col, parse_col(col)) for col in colunas_validas if pd.notnull(parse_col(col))]

    if datas_convertidas:
        col_mais_recente = max(datas_convertidas, key=lambda x: x[1])[0]
    
        # Ordena pela coluna mais recente (exceto a linha Total Geral)
        tem_total = "Total Geral" in tabela_final.index
        if tem_total:
            total_row = tabela_final.loc[["Total Geral"]]
            corpo_ordenado = tabela_final.drop(index="Total Geral").sort_values(by=col_mais_recente, ascending=False)
            tabela_final = pd.concat([total_row, corpo_ordenado])
        else:
            tabela_final = tabela_final.sort_values(by=col_mais_recente, ascending=False)

    # 🔥 Ordenação da tabela na TELA: pela coluna (Bruto) mais recente, se não tiver, pela (Real)
    colunas_bruto = [col for col in tabela_final.columns if '(Bruto)' in col]
    colunas_real = [col for col in tabela_final.columns if '(Real)' in col]

    # 📅 Ordena as colunas com base na data do nome
    def extrair_data(col):
        try:
            parte = col.split(' ')[0]
            if '/' in parte:
                return pd.to_datetime(f"01/{parte}", dayfirst=True)
            elif parte.isdigit() and len(parte) == 4:
                return pd.to_datetime(f"01/01/{parte}")
        except:
            return pd.NaT
        return pd.NaT

    # 🔍 Busca a coluna (Bruto) mais recente
    if colunas_bruto:
        colunas_bruto_ordenadas = sorted(colunas_bruto, key=extrair_data, reverse=True)
        coluna_ordenacao = colunas_bruto_ordenadas[0]
    elif colunas_real:
        colunas_real_ordenadas = sorted(colunas_real, key=extrair_data, reverse=True)
        coluna_ordenacao = colunas_real_ordenadas[0]
    else:
        coluna_ordenacao = None

    # 🔥 Faz a ordenação, mantendo "Total Geral" no topo
    if coluna_ordenacao:
        tem_total = "Total Geral" in tabela_final.index
        if tem_total:
            total_row = tabela_final.loc[["Total Geral"]]
            corpo_ordenado = tabela_final.drop(index="Total Geral").sort_values(by=coluna_ordenacao, ascending=False)
            tabela_final = pd.concat([total_row, corpo_ordenado])
        else:
            tabela_final = tabela_final.sort_values(by=coluna_ordenacao, ascending=False)

    
    tabela_formatada = tabela_final.applymap(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(x, (float, int)) else x
    )
    st.dataframe(tabela_formatada, use_container_width=True)

    


   


import io
import itertools

buffer = io.BytesIO()
df_acumulado_tipo = pd.DataFrame(columns=["Tipo", "Acumulado no Mês Tipo"])
# 🔥 Padroniza nomes
df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.lower().str.title()
df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower().str.title()

# 🔗 Prepara o dataframe para exportação
if modo_visao == "Por Loja":
    tabela_exportar = tabela_final.reset_index()
    tabela_exportar = tabela_exportar.rename(columns={tabela_exportar.columns[0]: "Loja"})

    tabela_exportar = tabela_exportar.merge(
        df_empresa[["Loja", "Grupo", "Tipo"]],
        on="Loja",
        how="left"
    )

    cols = ["Grupo", "Loja", "Tipo"] + [col for col in tabela_exportar.columns if col not in ["Grupo", "Loja", "Tipo"]]
    tabela_exportar = tabela_exportar[cols]

else:
    tabela_exportar = tabela_final.reset_index()
    tabela_exportar = tabela_exportar.rename(columns={tabela_exportar.columns[0]: "Grupo"})

# 🔥 Limpeza de Grupo
tabela_exportar["Grupo"] = tabela_exportar["Grupo"].astype(str).str.strip()
tabela_exportar = tabela_exportar[~tabela_exportar["Grupo"].isin(["", "nan", "NaN", "None"])]

# 🔥 Cópia sem a coluna Tipo
tabela_exportar_sem_tipo = tabela_exportar.drop(columns="Tipo", errors="ignore")

# 🔥 Cria relação Loja → Grupo → Tipo
relacao_loja = df_empresa[["Loja", "Grupo", "Tipo"]].drop_duplicates()

# 🔥 Cria dataframe acumulado SOMENTE quando agrupamento for Dia
df_acumulado_tipo = pd.DataFrame()

if agrupamento == "Dia":
    try:
        data_maxima = pd.to_datetime(data_fim)
        ano = data_maxima.year
        mes = data_maxima.month
        dia = data_maxima.day

        df_acumulado = df_anos[
            (df_anos["Data"].dt.year == ano) &
            (df_anos["Data"].dt.month == mes) &
            (df_anos["Data"].dt.day <= dia)
        ].copy()

        df_acumulado["Loja"] = df_acumulado["Loja"].astype(str).str.strip().str.lower().str.title()

        df_acumulado = df_acumulado.merge(relacao_loja, on="Loja", how="left")
        df_acumulado = df_acumulado[~df_acumulado["Tipo"].isna()]

        # 🔥 Acumulado por Loja
        if modo_visao == "Por Loja":
            df_agrupado = df_acumulado.groupby("Loja")["Fat.Real"].sum().reset_index()
            df_agrupado.rename(columns={"Fat.Real": "Acumulado no Mês"}, inplace=True)
            tabela_exportar_sem_tipo = tabela_exportar_sem_tipo.merge(
                df_agrupado, on="Loja", how="left"
            )

        # 🔥 Acumulado por Grupo
        elif modo_visao == "Por Grupo":
            df_agrupado = df_acumulado.groupby("Grupo")["Fat.Real"].sum().reset_index()
            df_agrupado.rename(columns={"Fat.Real": "Acumulado no Mês"}, inplace=True)
            tabela_exportar_sem_tipo = tabela_exportar_sem_tipo.merge(
                df_agrupado, on="Grupo", how="left"
            )

        # 🔥 Acumulado por Tipo (🚩 GARANTIDO AQUI 🚩)
        df_acumulado_tipo = df_acumulado.groupby("Tipo")["Fat.Real"].sum().reset_index()
        df_acumulado_tipo.rename(columns={"Fat.Real": "Acumulado no Mês Tipo"}, inplace=True)

        tabela_exportar_sem_tipo = tabela_exportar_sem_tipo.merge(
            df_acumulado_tipo, on="Tipo", how="left"
        )

        # 🔥 Organiza as colunas
        cols_atuais = [col for col in tabela_exportar_sem_tipo.columns if col not in ["Acumulado no Mês", "Acumulado no Mês Tipo"]]
        tabela_exportar_sem_tipo = tabela_exportar_sem_tipo[cols_atuais + ["Acumulado no Mês", "Acumulado no Mês Tipo"]]

    except Exception as e:
        st.warning(f"⚠️ Erro no cálculo do acumulado do mês: {e}")



# 🔥 Criação do Excel
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    tabela_exportar_sem_tipo.to_excel(writer, sheet_name="Faturamento", index=False, header=True, startrow=0)

    workbook = writer.book
    worksheet = writer.sheets["Faturamento"]

    cores_grupo = itertools.cycle(["#D9EAD3", "#CFE2F3"])

    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    subtotal_format = workbook.add_format({
        'bold': True, 'bg_color': '#FFE599', 'border': 1, 'num_format': 'R$ #,##0.00'
    })

    totalgeral_format = workbook.add_format({
        'bold': True, 'bg_color': '#A9D08E', 'border': 1, 'num_format': 'R$ #,##0.00'
    })

    for col_num, header in enumerate(tabela_exportar_sem_tipo.columns):
        worksheet.write(0, col_num, header, header_format)

    grupos_info = []

    for grupo_atual in tabela_exportar_sem_tipo["Grupo"].unique():
        grupo_linhas = tabela_exportar_sem_tipo[tabela_exportar_sem_tipo["Grupo"] == grupo_atual]
        coluna_valor = [col for col in tabela_exportar_sem_tipo.columns if "Total" in col or col.isnumeric()]
        coluna_valor = coluna_valor[0] if coluna_valor else tabela_exportar_sem_tipo.columns[2]
        subtotal_grupo = grupo_linhas[coluna_valor].sum()

        grupos_info.append({
            "grupo": grupo_atual,
            "linhas": grupo_linhas,
            "subtotal": subtotal_grupo,
            "qtd_lojas": grupo_linhas["Loja"].nunique() if "Loja" in grupo_linhas.columns else 0
        })

    grupos_info = sorted(grupos_info, key=lambda x: x["subtotal"], reverse=True)

    # ✅ Calcula subtotais por Tipo (corrigido)
    tipos_info = []
    if "Tipo" in tabela_exportar.columns:
        colunas_soma = [col for col in tabela_exportar_sem_tipo.columns if col not in ["Grupo", "Loja", "Tipo"]]

        for tipo_atual in tabela_exportar["Tipo"].dropna().unique():
            tipo_linhas = tabela_exportar[tabela_exportar["Tipo"] == tipo_atual]
            qtd_lojas_tipo = tipo_linhas["Loja"].nunique() if "Loja" in tipo_linhas.columns else 0

            soma_colunas = []
            for col in colunas_soma:
                if col in tipo_linhas.columns:
                    soma = tipo_linhas[col].sum()
                    soma_colunas.append(soma)
                else:
                    soma_colunas.append(0)

            # 🔥 Pega o acumulado correto
            acumulado_tipo = df_acumulado_tipo[df_acumulado_tipo["Tipo"] == tipo_atual]["Acumulado no Mês Tipo"]
            acumulado_tipo = acumulado_tipo.values[0] if not acumulado_tipo.empty else 0

            soma_colunas.append(acumulado_tipo)

            tipos_info.append({
                "tipo": tipo_atual,
                "qtd_lojas": qtd_lojas_tipo,
                "somas": soma_colunas
            })






            
            if agrupamento == "Dia":
                data_maxima = pd.to_datetime(data_fim)
                ano = data_maxima.year
                mes = data_maxima.month
                dia = data_maxima.day

                df_acumulado_tipo = df_anos.merge(
                    df_empresa[["Loja", "Grupo", "Tipo"]],
                    on="Loja",
                    how="left"
                )

                df_acumulado_tipo = df_acumulado_tipo[
                    (df_acumulado_tipo["Data"].dt.year == ano) &
                    (df_acumulado_tipo["Data"].dt.month == mes) &
                    (df_acumulado_tipo["Data"].dt.day <= dia) &
                    (df_acumulado_tipo["Tipo"] == tipo_atual)
                ]

                acumulado_tipo = df_acumulado_tipo["Fat.Real"].sum()

                soma_colunas.append(acumulado_tipo)


       
            tipos_info.append({
                "tipo": tipo_atual,
                "qtd_lojas": qtd_lojas_tipo,
                "somas": soma_colunas
            })

   
    linha = 1

    for tipo in tipos_info:
        # 🔥 Prepara a linha do Tipo com alinhamento correto nas colunas
        linha_tipo = []
        for idx, col in enumerate(tabela_exportar_sem_tipo.columns):
            if idx == 0:
                linha_tipo.append(f"Tipo: {tipo['tipo']}")
            elif idx == 1:
                linha_tipo.append(f"Lojas: {tipo['qtd_lojas']}")
            else:
                pos = idx - 2
                if pos < len(tipo["somas"]):
                    linha_tipo.append(tipo["somas"][pos])
                else:
                    linha_tipo.append("")

        # 🔥 Escreve a linha no Excel
        for col_num, val in enumerate(linha_tipo):
            if isinstance(val, (int, float)) and not pd.isna(val):
                worksheet.write_number(linha, col_num, val, subtotal_format)
            else:
                worksheet.write(linha, col_num, str(val), subtotal_format)

        linha += 1
   
    # 🔥 Total Geral
    linha_totalgeral = ["Total Geral", ""]

    for col in tabela_exportar_sem_tipo.columns[2:]:
        soma = sum(g["linhas"][col].sum() for g in grupos_info)
        linha_totalgeral.append(soma)

    for col_num, val in enumerate(linha_totalgeral):
        if isinstance(val, (int, float)) and not pd.isna(val):
            worksheet.write_number(linha, col_num, val, totalgeral_format)
        else:
            worksheet.write(linha, col_num, str(val), totalgeral_format)

    linha += 1

    # 🔥 Escreve os dados dos grupos
    row_num = linha

    for grupo, group_color in zip(grupos_info, cores_grupo):
        grupo_atual = grupo["grupo"]
        grupo_linhas = grupo["linhas"]
        qtd_lojas = grupo["qtd_lojas"]

        group_row_format = workbook.add_format({
            'bg_color': group_color, 'border': 1, 'num_format': 'R$ #,##0.00'
        })

        for _, row in grupo_linhas.iterrows():
            for col_num, val in enumerate(row):
                if isinstance(val, (int, float)) and not pd.isna(val):
                    worksheet.write_number(row_num, col_num, val, group_row_format)
                else:
                    worksheet.write(row_num, col_num, str(val) if not pd.isna(val) else "", group_row_format)
            row_num += 1

        linha_subtotal = [f"Subtotal {grupo_atual}", f"Lojas: {qtd_lojas}"]

        for col in tabela_exportar_sem_tipo.columns[2:]:
            soma = grupo_linhas[col].sum()
            linha_subtotal.append(soma)

        for col_num, val in enumerate(linha_subtotal):
            if isinstance(val, (int, float)) and not pd.isna(val):
                worksheet.write_number(row_num, col_num, val, subtotal_format)
            else:
                worksheet.write(row_num, col_num, str(val), subtotal_format)

        row_num += 1

    worksheet.set_column(0, len(tabela_exportar_sem_tipo.columns), 18)
    worksheet.hide_gridlines(option=2)

# 🔽 Botão download
st.download_button(
    label="📥 Baixar Excel",
    data=buffer.getvalue(),
    file_name="faturamento_visual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_excel_visual_painel"
)
