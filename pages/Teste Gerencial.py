# pages/RelatorioGerenciais.py

import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px

st.set_page_config(page_title="Relatorios", layout="wide")

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

planilha = gc.open("Faturamento Sistema Externo")
aba = planilha.worksheet("Fat Sistema Externo")
dados = aba.get_all_records()
df = pd.DataFrame(dados)

# ================================
# 2. Limpeza e preparação dos dados
# ================================
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
df_anos = df[df["Data"].notna() & df["Fat.Real"].notna()].copy()

# ================================
# 3. Layout e estilo
# ================================
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

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatórios Gerenciais</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 4. Abas
# ================================
aba1, aba2  = st.tabs([
    "📊 Graficos Anuais - Grupo",
    "📥 Relatório Analitico"
  
])

# ================================
# 📊 Aba 1 - Gráfico Anual
# ================================
with aba1:
    st.subheader("Faturamento Anual")

    fat_mensal = df_anos.groupby(["Mês", "Ano"])["Fat.Real"].sum().reset_index()
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    fat_mensal["Nome Mês"] = fat_mensal["Mês"].map(meses)
    fat_mensal["Ano"] = fat_mensal["Ano"].astype(str)
    fat_mensal = fat_mensal.sort_values(["Mês", "Ano"])

    color_map = {"2024": "#1f77b4", "2025": "#ff7f0e"}

    fig = px.bar(
        fat_mensal,
        x="Nome Mês",
        y="Fat.Real",
        color="Ano",
        barmode="group",
        text_auto=".2s",
        custom_data=["Ano"],
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
    st.plotly_chart(fig, use_container_width=True)



# ================================
# 📥 Aba 2 - Relatório Analítico
# ================================
with aba2:
    st.subheader("📥 Relatório Analítico")
    anos = sorted(df_anos["Ano"].dropna().unique(), reverse=True)
    anos_selecionados = st.multiselect("🗓️ Selecione os anos", anos, default=anos)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for ano in anos_selecionados:
            df_fat = df_anos[df_anos["Ano"] == ano].copy()
            df_fat["Loja"] = df_fat["Loja"].astype(str).str.title()
            df_fat["Mês"] = df_fat["Data"].dt.strftime("%m - %B")

            tabela = df_fat.pivot_table(index="Loja", columns="Mês", values="Fat.Real", aggfunc="sum", fill_value=0)
            linha_total = tabela.sum().to_frame().T
            linha_total.index = ["Total Geral"]
            coluna_total = tabela.sum(axis=1)
            tabela.insert(0, "Total Geral", coluna_total)
            linha_total.insert(0, "Total Geral", coluna_total.sum())
            tabela_final = pd.concat([linha_total, tabela])

            tabela_formatada = tabela_final.applymap(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            st.markdown(f"### 📋 Faturamento - {ano}")
            st.dataframe(tabela_formatada)
            tabela_final.to_excel(writer, sheet_name=f"Faturamento_{ano}")

    st.download_button(
        label="📥 Baixar Excel com Totais",
        data=buffer.getvalue(),
        file_name="faturamento_real_totais.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
