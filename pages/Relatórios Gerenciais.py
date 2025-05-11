# pages/Relatorio Gerencial.py



# pages/FaturamentoServico.py

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

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela_Empresa").get_all_records())

# Configuração da página
st.set_page_config(page_title="Relatórios Gerenciais", layout="wide")

# Título da página
st.title("📊 Relatórios Gerenciais")

# ================================
# 3. Abas internas
# ================================
aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
    "📊 Gráfico Anual Comparativo",
    "🗓️ Relatório Mensal Detalhado",
    "📌 Análise Extra 1",
    "📌 Análise Extra 2",
    "📌 Análise Extra 3",
    "📌 Análise Extra 4",
    "📌 Análise Extra 5"
])

# ================================
# 4. Aba 1 - Gráfico Anual Comparativo
# ================================
with aba1:
    st.subheader("📊 Gráfico Anual Comparativo")

    def limpar_valor(x):
        try:
            if isinstance(x, str):
                return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
            elif isinstance(x, (int, float)):
                return x
        except:
            return None

    for coluna in ["Fat.Total", "Serv/Tx", "Fat.Real"]:
        if coluna in df_empresa.columns:
            df_empresa[coluna] = df_empresa[coluna].apply(limpar_valor)
            df_empresa[coluna] = pd.to_numeric(df_empresa[coluna], errors="coerce")

    df_empresa["Data"] = pd.to_datetime(df_empresa["Data"], errors="coerce", dayfirst=True)
    df_empresa["Ano"] = df_empresa["Data"].dt.year
    df_empresa["Mês"] = df_empresa["Data"].dt.month

    meses_portugues = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    df_empresa["Nome Mês"] = df_empresa["Mês"].map(meses_portugues)

    anos_disponiveis = sorted(df_empresa["Ano"].dropna().unique())
    anos_comparacao = st.multiselect("📊 Anos para gráficos de comparação", options=anos_disponiveis, default=anos_disponiveis)

    df = df_empresa[df_empresa["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"]).copy()
    df_lojas = df.groupby("Ano")["Loja"].nunique().reset_index().rename(columns={"Loja": "Qtd_Lojas"})
    fat_mensal = df.groupby(["Nome Mês", "Ano"])["Fat.Real"].sum().reset_index()

    meses = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
             "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
    fat_mensal["MesNum"] = fat_mensal["Nome Mês"].str[:3].str.lower().map(meses)
    fat_mensal["Ano"] = fat_mensal["Ano"].astype(str)
    fat_mensal["MesAno"] = fat_mensal["Nome Mês"].str[:3].str.capitalize() + "/" + fat_mensal["Ano"].str[-2:]
    fat_mensal = fat_mensal.sort_values(["MesNum", "Ano"])

    color_map = {"2024": "#1f77b4", "2025": "#ff7f0e"}

    fig = px.bar(fat_mensal, x="Nome Mês", y="Fat.Real", color="Ano", barmode="group",
                 text_auto=".2s", custom_data=["MesAno"], color_discrete_map=color_map)
    fig.update_traces(textposition="outside")
    fig.update_layout(xaxis_title=None, yaxis_title=None, xaxis_tickangle=-45, showlegend=False,
                      yaxis=dict(showticklabels=False, showgrid=False, zeroline=False))

    df_total = fat_mensal.groupby("Ano")["Fat.Real"].sum().reset_index()
    df_total = df_total.merge(df_lojas, on="Ano", how="left")
    df_total["AnoTexto"] = df_total.apply(
        lambda row: f"{int(row['Ano'])}         R$ {row['Fat.Real']/1_000_000:,.1f} Mi".replace(",", "."), axis=1
    )

    fig_total = px.bar(df_total, x="Fat.Real", y="Ano", orientation="h", color="Ano",
                       text="AnoTexto", color_discrete_map=color_map)
    fig_total.update_traces(textposition="outside", textfont=dict(size=16), showlegend=False)
    for i, row in df_total.iterrows():
        fig_total.add_annotation(x=0.1, y=row["Ano"], text=row["AnoTexto"], showarrow=False,
                                 xanchor="left", yanchor="middle", font=dict(color="white", size=16),
                                 xref="x", yref="y")
        fig_total.add_annotation(x=row["Fat.Real"], y=row["Ano"],
                                 text=f"{int(row['Qtd_Lojas'])} Lojas", xanchor="left", yanchor="bottom",
                                 yshift=-8, font=dict(color="red", size=16, weight="bold"),
                                 showarrow=False, xref="x", yref="y")
    fig_total.update_layout(height=130, margin=dict(t=0, b=0, l=0, r=0), title=None,
                            xaxis=dict(visible=False), yaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
                            yaxis_title=None, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")

    st.plotly_chart(fig, use_container_width=True)
    st.plotly_chart(fig_total, use_container_width=True)

# ================================
# 5. Aba 2 - Relatório Mensal Detalhado
# ================================
with aba2:
    st.subheader("Faturamento Anual")
    st.plotly_chart(fig_total, use_container_width=True)

    st.markdown("---")
    st.subheader("Faturamento Mensal")
    st.plotly_chart(fig, use_container_width=True)

    df["Ano"] = df["Data"].dt.year
    anos_disponiveis = sorted(df["Ano"].dropna().unique())
    anos_selecionados = st.multiselect("🗓️ Selecione os anos que deseja exibir", options=anos_disponiveis, default=anos_disponiveis)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for ano in anos_selecionados:
            df_fat = df[df["Ano"] == ano].copy()
            df_fat["Loja"] = df_fat["Loja"].astype(str).str.strip().str.lower().str.title()
            df_fat["Fat.Real"] = pd.to_numeric(df_fat["Fat.Real"], errors="coerce")

            meses_pt = {
                "January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril",
                "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto",
                "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"
            }
            df_fat["Mês"] = df_fat["Data"].dt.strftime("%m - %B")
            df_fat["Mês"] = df_fat["Mês"].apply(lambda x: f"{x[:6]}{meses_pt.get(x[6:], x[6:])}")

            tabela_fat_real = df_fat.pivot_table(index="Loja", columns="Mês", values="Fat.Real", aggfunc="sum", fill_value=0)

            linha_total = tabela_fat_real.sum().to_frame().T
            linha_total.index = ["Total Geral"]
            coluna_total = tabela_fat_real.sum(axis=1)
            tabela_fat_real.insert(0, "Total Geral", coluna_total)
            linha_total.insert(0, "Total Geral", coluna_total.sum())
            tabela_com_total = pd.concat([linha_total, tabela_fat_real])

            tabela_formatada = tabela_com_total.applymap(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )

            st.markdown("---")
            st.subheader(f"📋 Faturamento Real por Loja e Mês - {ano}")
            st.dataframe(tabela_formatada)

            tabela_com_total.to_excel(writer, sheet_name=f"Faturamento_{ano}")

    st.markdown("---")
    st.download_button(
        label="📥 Baixar Excel com Totais por Ano",
        data=buffer.getvalue(),
        file_name="faturamento_real_totais_por_ano.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ================================
# 6. Abas futuras (placeholders)
# ================================
with aba3:
    st.subheader("📌 Análise Extra 1")
    st.info("Aba reservada para conteúdo futuro.")

with aba4:
    st.subheader("📌 Análise Extra 2")
    st.info("Aba reservada para conteúdo futuro.")

with aba5:
    st.subheader("📌 Análise Extra 3")
    st.info("Aba reservada para conteúdo futuro.")

with aba6:
    st.subheader("📌 Análise Extra 4")
    st.info("Aba reservada para conteúdo futuro.")

with aba7:
    st.subheader("📌 Análise Extra 5")
    st.info("Aba reservada para conteúdo futuro.")
