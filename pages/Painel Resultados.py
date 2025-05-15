# pages/PainelResultados.py


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

st.set_page_config(page_title="Vendas Diarias", layout="wide")
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
# PainelResultados.py (trecho para a Aba 3)

with aba3:
    st.markdown("### 📋 Relatório Analítico Detalhado")

    # 🧩 Escolha do nível de agrupamento
    nivel_agrupamento = st.selectbox("🔽 Nível de análise", ["Ano", "Mês", "Dia"])

    # Agrupamentos baseados em datas
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Mês"] = df_anos["Data"].dt.strftime("%m/%Y")
    df_anos["Dia"] = df_anos["Data"].dt.strftime("%d/%m/%Y")
    df_anos["Chave"] = df_anos[nivel_agrupamento]

    # ✅ Opções de métrica
    tipo_metrica = st.radio("💰 Métrica:", ["Bruto", "Real", "Ambos"], horizontal=True)

    # 🔄 Gerar tabelas conforme a métrica
    if tipo_metrica == "Bruto":
        tabela = df_anos.pivot_table(index="Loja", columns="Chave", values="Fat.Total", aggfunc="sum", fill_value=0)
        coluna_total = tabela.sum(axis=1)
        tabela.insert(0, "Total Geral", coluna_total)
        linha_total = pd.DataFrame(tabela.sum()).T
        linha_total.index = ["Total Geral"]
        tabela_final = pd.concat([linha_total, tabela])

    elif tipo_metrica == "Real":
        tabela = df_anos.pivot_table(index="Loja", columns="Chave", values="Fat.Real", aggfunc="sum", fill_value=0)
        coluna_total = tabela.sum(axis=1)
        tabela.insert(0, "Total Geral", coluna_total)
        linha_total = pd.DataFrame(tabela.sum()).T
        linha_total.index = ["Total Geral"]
        tabela_final = pd.concat([linha_total, tabela])

    else:  # Ambos
        tabela_bruto = df_anos.pivot_table(index="Loja", columns="Chave", values="Fat.Total", aggfunc="sum", fill_value=0)
        tabela_bruto.columns = [f"{col} (Bruto)" for col in tabela_bruto.columns]

        tabela_real = df_anos.pivot_table(index="Loja", columns="Chave", values="Fat.Real", aggfunc="sum", fill_value=0)
        tabela_real.columns = [f"{col} (Real)" for col in tabela_real.columns]

        tabela = pd.concat([tabela_bruto, tabela_real], axis=1)

        # Calcular os totais separados
        colunas_bruto = [c for c in tabela.columns if "(Bruto)" in c]
        colunas_real = [c for c in tabela.columns if "(Real)" in c]

        tabela["Total Geral (Bruto)"] = tabela[colunas_bruto].sum(axis=1)
        tabela["Total Geral (Real)"] = tabela[colunas_real].sum(axis=1)

        # Reorganizar colunas
        cols_ordenadas = ["Total Geral (Bruto)", "Total Geral (Real)"] + \
                         sorted([c for c in tabela.columns if c not in ["Total Geral (Bruto)", "Total Geral (Real)"]], reverse=True)
        tabela = tabela[cols_ordenadas]

        linha_total = pd.DataFrame(tabela.sum()).T
        linha_total.index = ["Total Geral"]
        tabela_final = pd.concat([linha_total, tabela])

    # 💄 Formatar valores monetários
    tabela_formatada = tabela_final.applymap(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.dataframe(tabela_formatada)

    # 📥 Exportar para Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        tabela_final.to_excel(writer, sheet_name=f"Relatorio_{nivel_agrupamento}", index=True)

    st.download_button(
        label="📥 Baixar Excel",
        data=buffer.getvalue(),
        file_name=f"relatorio_{nivel_agrupamento.lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ================================
# Aba 4: Analise Lojas
# ================================
with aba4:
    st.info("Você pode colocar tabelas detalhadas e botões de download aqui.")
