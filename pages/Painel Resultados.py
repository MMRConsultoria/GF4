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
    st.info("Ideal para mostrar evolução por ano ou por trimestre.")

# ================================
# Aba 3: Relatorio Analitico
# ================================
# ================================
# Aba 3: Análise por Período
# ================================
# ================================
# Aba 3: Análise por Ano, Mês e Dia
# ================================
with aba3:
  
    st.markdown("""
        <style>
        /* Cor dos chips de seleção (anos e meses) */
        .stMultiSelect [data-baseweb="tag"] {
            background-color: #bbbbbb !important;  /* cinza claro */
            color: black !important;
        }

        .stMultiSelect [data-baseweb="tag"] span {
            font-weight: 600;
        }
        </style>
    """, unsafe_allow_html=True)

    # depois vem seus filtros
    # st.multiselect("Ano")
    # st.multiselect("Mês")

    
    #st.markdown("### 📊 Análise de Faturamento por Período")
    
    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.lower().str.title()
    df_anos["Fat.Total"] = pd.to_numeric(df_anos["Fat.Total"], errors="coerce")
    df_anos["Fat.Real"] = pd.to_numeric(df_anos["Fat.Real"], errors="coerce")
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Mês Num"] = df_anos["Data"].dt.month
    df_anos["Mês Nome"] = df_anos["Data"].dt.strftime('%B')
    df_anos["Mês"] = df_anos["Data"].dt.strftime('%m/%Y')
    df_anos["Dia"] = df_anos["Data"].dt.strftime('%d/%m/%Y')

    # === SELEÇÃO DE INTERVALO DE DATAS (como na auditoria Everest) ===
    data_min = df_anos["Data"].min()
    data_max = df_anos["Data"].max()

    #data_inicio, data_fim = st.date_input(
    #    "📆 Selecione o intervalo de datas:",
    #    value=[data_min, data_max],
    #    min_value=data_min,
    #    max_value=data_max
    #)

    #df_filtrado = df_anos[(df_anos["Data"] >= pd.to_datetime(data_inicio)) & (df_anos["Data"] <= pd.to_datetime(data_fim))]



    # === ANO ===
    anos_disponiveis = sorted(df_anos["Ano"].unique(), reverse=True)
    ano_opcao = st.multiselect("📅 Selecione o(s) ano(s):", options=anos_disponiveis, default=anos_disponiveis)

    df_filtrado = df_anos[df_anos["Ano"].isin(ano_opcao)]

    # === MÊS ===
    meses_dict = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    meses_disponiveis = sorted(df_filtrado["Mês Num"].unique())
    meses_nomes_disponiveis = [meses_dict[m] for m in meses_disponiveis]

    meses_selecionados = st.multiselect(
        "🗓️ Selecione o(s) mês(es):", options=meses_nomes_disponiveis, default=meses_nomes_disponiveis
    )

    meses_numeros = [k for k, v in meses_dict.items() if v in meses_selecionados]
    df_filtrado = df_filtrado[df_filtrado["Mês Num"].isin(meses_numeros)]

    # === DIA COM CALENDÁRIO ===
    data_inicio, data_fim = st.date_input(
        "📆 Selecione o intervalo de dias:",
        value=[df_filtrado["Data"].min(), df_filtrado["Data"].max()],
        min_value=df_filtrado["Data"].min(),
        max_value=df_filtrado["Data"].max()
    )

    df_filtrado = df_filtrado[
        (df_filtrado["Data"] >= pd.to_datetime(data_inicio)) &
        (df_filtrado["Data"] <= pd.to_datetime(data_fim))
    ].copy()

    # === AGRUPAMENTO ===
    agrupamento = st.radio("📂 Agrupar por:", ["Ano", "Mês", "Dia"], horizontal=True)

    if agrupamento == "Ano":
        df_filtrado["Agrupador"] = df_filtrado["Ano"].astype(str)
        df_filtrado["Ordem"] = pd.to_datetime(df_filtrado["Ano"], format="%Y")

    elif agrupamento == "Mês":
        df_filtrado["Agrupador"] = df_filtrado["Data"].dt.strftime("%m/%Y")
        df_filtrado["Ordem"] = df_filtrado["Data"].dt.to_period("M").dt.to_timestamp()

    elif agrupamento == "Dia":
        df_filtrado["Agrupador"] = df_filtrado["Data"].dt.strftime("%d/%m/%Y")
        df_filtrado["Ordem"] = df_filtrado["Data"]


    # === MÉTRICA ===
    tipo_metrica = st.radio("💰 Selecione a métrica:", ["Bruto", "Real", "Ambos"], horizontal=True)

    # === Tabela conforme métrica ===

    if tipo_metrica == "Bruto":
        tabela = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Total", aggfunc="sum", fill_value=0)

        # Ordenar colunas em ordem decrescente
        ordem_agrupador = df_filtrado[["Agrupador", "Ordem"]].drop_duplicates().sort_values("Ordem", ascending=False)
        colunas_ordenadas = list(ordem_agrupador["Agrupador"])
        tabela = tabela.reindex(columns=colunas_ordenadas)

        tabela.insert(0, "Total Bruto", tabela.sum(axis=1))
        linha_total = pd.DataFrame(tabela.sum()).T
        linha_total.index = ["Total Geral"]
        tabela_final = pd.concat([linha_total, tabela])
    
    elif tipo_metrica == "Real":
        tabela = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Real", aggfunc="sum", fill_value=0)

        # Ordenar colunas em ordem decrescente
        ordem_agrupador = df_filtrado[["Agrupador", "Ordem"]].drop_duplicates().sort_values("Ordem", ascending=False)
        colunas_ordenadas = list(ordem_agrupador["Agrupador"])
        tabela = tabela.reindex(columns=colunas_ordenadas)

        tabela.insert(0, "Total Real", tabela.sum(axis=1))
        linha_total = pd.DataFrame(tabela.sum()).T
        linha_total.index = ["Total Geral"]
        tabela_final = pd.concat([linha_total, tabela])
        
        # Ordenar colunas em ordem decrescente usando a data real (Ordem)
        ordem_agrupador = df_filtrado[["Agrupador", "Ordem"]].drop_duplicates()
        ordem_agrupador = ordem_agrupador.sort_values("Ordem", ascending=False)
        colunas_ordenadas = list(ordem_agrupador["Agrupador"])
        tabela = tabela.reindex(columns=colunas_ordenadas)
    else:  # Ambos
        tab_bruto = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Total", aggfunc="sum", fill_value=0)
        tab_real = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Real", aggfunc="sum", fill_value=0)

        # ✅ Ordenação decrescente
        ordem = (
            df_filtrado[["Agrupador", "Ordem"]]
            .drop_duplicates()
            .sort_values("Ordem", ascending=False)
            .Agrupador.tolist()
        )

        tab_bruto = tab_bruto.reindex(columns=ordem)
        tab_real = tab_real.reindex(columns=ordem)

        # ✅ Renomear colunas
        tab_bruto.columns = [f"{col} (Bruto)" for col in tab_bruto.columns]
        tab_real.columns = [f"{col} (Real)" for col in tab_real.columns]

        # ✅ Intercalar colunas
        colunas_intercaladas = []
        for col in ordem:
            colunas_intercaladas.append(f"{col} (Bruto)")
            colunas_intercaladas.append(f"{col} (Real)")

        tabela = pd.concat([tab_bruto, tab_real], axis=1)
        tabela = tabela[[c for c in colunas_intercaladas if c in tabela.columns]]

        # Totais
        tabela.insert(0, "Total Real", tab_real.sum(axis=1))
        tabela.insert(0, "Total Bruto", tab_bruto.sum(axis=1))
        linha_total = pd.DataFrame(tabela.sum()).T
        linha_total.index = ["Total Geral"]
        tabela_final = pd.concat([linha_total, tabela])



    
    # === Exibir e exportar ===
    tabela_formatada = tabela_final.applymap(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    st.markdown("---")
    st.dataframe(tabela_formatada, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        tabela_final.to_excel(writer, sheet_name="Faturamento Detalhado", index=True)

    st.download_button(
        label="📥 Baixar Excel com Totais",
        data=buffer.getvalue(),
        file_name="faturamento_detalhado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ================================
# Aba 4: Analise Lojas
# ================================
with aba4:

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
    anos_disponiveis = sorted(df_anos["Ano"].unique(), reverse=True)
    ano_opcao = st.multiselect("📅 Selecione ano/mês(s):", options=anos_disponiveis, default=anos_disponiveis, key="ano_aba3")
    df_filtrado = df_anos[df_anos["Ano"].isin(ano_opcao)]

    meses_dict = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
                  7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}

    meses_disponiveis = sorted(df_filtrado["Mês Num"].unique())
    meses_nomes_disponiveis = [meses_dict[m] for m in meses_disponiveis]
    meses_selecionados = st.multiselect("", options=meses_nomes_disponiveis, default=meses_nomes_disponiveis, key="meses_aba3")
    meses_numeros = [k for k, v in meses_dict.items() if v in meses_selecionados]
    df_filtrado = df_filtrado[df_filtrado["Mês Num"].isin(meses_numeros)]

    data_inicio, data_fim = st.date_input("", value=[df_filtrado["Data"].min(), df_filtrado["Data"].max()])
    df_filtrado = df_filtrado[(df_filtrado["Data"] >= pd.to_datetime(data_inicio)) & (df_filtrado["Data"] <= pd.to_datetime(data_fim))].copy()


    # Filtros laterais lado a lado
    col1, col2, col3 = st.columns(3)

    with col1:
        agrupamento = st.radio(" ", ["Ano", "Mês", "Dia"], horizontal=True, key="agrup_aba4")

    with col2:
        modo_visao = st.radio(" ", ["Por Loja", "Por Grupo"], horizontal=True, key="visao_aba4")

    with col3:
        tipo_metrica = st.radio(" ", ["Bruto", "Real", "Ambos"], horizontal=True, key="metrica_aba4")


    # Filtro para exibir ou não a coluna Total
    exibir_total_opcao = st.radio("📊 Coluna Total:", ["Sim", "Não"], index=0, horizontal=True)
    exibir_total = exibir_total_opcao == "Sim"

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

            # Adiciona colunas separadas de Grupo e Loja
            tabela = tabela.reset_index()  # "Loja" vira coluna
            tabela = tabela.merge(df_filtrado[["Loja", "Grupo"]].drop_duplicates(), on="Loja", how="left")
            tabela = tabela[["Grupo", "Loja"] + [col for col in tabela.columns if col not in ["Grupo", "Loja"]]]

               # ✅ Adiciona linha Total Geral com Grupo vazio e Loja = "Total Geral"
            total_row = pd.DataFrame(tabela.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)).T
            total_row.insert(0, "Loja", "Total Geral")
            total_row.insert(0, "Grupo", "")
            tabela_final = pd.concat([total_row, tabela], ignore_index=True) 


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

    tabela_formatada = tabela_final.applymap(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(x, (float, int)) else x
    )
    st.dataframe(tabela_formatada, use_container_width=True)

buffer = io.BytesIO()
tabela_exportar = tabela_final.copy()

with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    tabela_exportar.to_excel(writer, sheet_name="Faturamento", index=False, startrow=1)

    workbook = writer.book
    worksheet = writer.sheets["Faturamento"]

    # Cabeçalhos formatados
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    even_row_format = workbook.add_format({
        'bg_color': '#DCE6F1', 'border': 1, 'num_format': 'R$ #,##0.00'
    })
    odd_row_format = workbook.add_format({
        'bg_color': '#FFFFFF', 'border': 1, 'num_format': 'R$ #,##0.00'
    })
    bold_row_format = workbook.add_format({
        'bold': True, 'border': 1, 'num_format': 'R$ #,##0.00'
    })

    # Cabeçalho manual
    for col_num, header in enumerate(tabela_exportar.columns):
        worksheet.write(0, col_num, header, header_format)

    # Dados com zebrinha e negrito no Total Geral
    for row_num, row in enumerate(tabela_exportar.itertuples(index=False), start=1):
        is_total = getattr(row, "Loja") == "Total Geral"
        row_format = bold_row_format if is_total else (even_row_format if row_num % 2 == 0 else odd_row_format)

        for col_num, val in enumerate(row):
            if isinstance(val, (float, int)):
                worksheet.write_number(row_num, col_num, val, row_format)
            else:
                worksheet.write(row_num, col_num, val, row_format)

    worksheet.set_column(0, len(tabela_exportar.columns), 18)
    worksheet.hide_gridlines(option=2)


# Botão download
st.download_button(
    label="📥 Baixar Excel",
    data=buffer.getvalue(),
    file_name="faturamento_visual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_excel_visual"
)

