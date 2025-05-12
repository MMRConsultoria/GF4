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

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela_Empresa").get_all_records())

# ================================
# 2. Configuração inicial do app
# ================================
st.set_page_config(page_title="Faturamento por Serviço", layout="wide")
#st.title("📋 Relatório de Faturamento por Serviço")

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


# Cabeçalho bonito (depois do estilo)
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento por Serviço</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# 3. Separação em ABAS
# ================================

aba1, aba2, aba3, aba4 = st.tabs([
    "📈 Graficos Anuais",
    "📊 Graficos Trimestrais",
    "📆 Analise Mensal",
    "📋 Analise Lojas"
])



# ================================
# Aba 1: Graficos Anuais
# ================================

with aba1:
 
# ================================
# 📈 Relatórios Gerenciais (Painel Interativo)
# ================================


	
    # Conectar ao Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)

    # Carregar dados
    planilha = gc.open("Faturamento Sistema Externo")
    aba = planilha.worksheet("Fat Sistema Externo")
    dados = aba.get_all_records()
    df = pd.DataFrame(dados)

   

# =========================
# 🧹 Tratamento dos dados
# =========================
# 🔽 Interface para selecionar anos no comparativo (após tratamento dos dados)
    anos_disponiveis = sorted(df["Ano"].dropna().unique())

    # Para os gráficos
    anos_comparacao = st.multiselect(
    	"📊 Anos para gráficos de comparação",
    	options=anos_disponiveis,
    	default=anos_disponiveis
    )

   	
    # Filtrar os dados com base na seleção
    #df_anos_filtrado = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"])
    def limpar_valor(x):
        try:
            if isinstance(x, str):
                return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
            elif isinstance(x, (int, float)):
                return x
        except:
            return None
        return None

    for coluna in ["Fat.Total", "Serv/Tx", "Fat.Real"]:
        if coluna in df.columns:
            df[coluna] = df[coluna].apply(limpar_valor)
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    # Conversão da coluna Data
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)

    # Criar colunas auxiliares
    df["Ano"] = df["Data"].dt.year
    df["Mês"] = df["Data"].dt.month
    meses_portugues = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    df["Nome Mês"] = df["Mês"].map(meses_portugues)

    # Só depois faz o filtro
    df_anos_filtrado = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"])
     
	
   # Filtro de anos
    #df_anos = df[df["Ano"].isin([2024, 2025])].dropna(subset=["Data", "Fat.Real"])
    # Filtro de anos com base no multiselect
    anos_disponiveis = sorted(df["Ano"].dropna().unique())
    #anos_comparacao = st.multiselect("📊 Escolha os anos para comparação nos gráficos", options=anos_disponiveis, default=anos_disponiveis)
    df_anos = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"]).copy()
#NOVO
    # Calcular a quantidade de lojas únicas por ano
    df_lojas = df_anos.groupby("Ano")["Loja"].nunique().reset_index()
    df_lojas.columns = ["Ano", "Qtd_Lojas"]	




    # Calcular a quantidade de lojas únicas por ano
    df_lojas = df_anos.groupby("Ano")["Loja"].nunique().reset_index()
    df_lojas.columns = ["Ano", "Qtd_Lojas"]



	
   # Agrupamento por mês e ano
    #fat_mensal = df_anos.groupby(["Nome Mês", "Ano"])["Fat.Real"].sum().reset_index()
    fat_mensal = df_anos_filtrado.groupby(["Nome Mês", "Ano"])["Fat.Real"].sum().reset_index()


# ==============================
# ➕ Barras Mensais (ajuste para o DataFrame com Totais)
# ==============================

# Converter mês para número (sem erro no "Total")
meses = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
}
fat_mensal["MesNum"] = fat_mensal["Nome Mês"].str[:3].str.lower().map(meses)

# Cria coluna MesAno
fat_mensal["Ano"] = fat_mensal["Ano"].astype(str)
fat_mensal["MesAno"] = fat_mensal["Nome Mês"].str[:3].str.capitalize() + "/" + fat_mensal["Ano"].str[-2:]

# Ordenar corretamente por mês e ano
fat_mensal = fat_mensal.sort_values(["MesNum", "Ano"])

# =========================
# 📊 Visualização
# =========================


# Filtrar os dados com base na seleção
#df_anos_filtrado = df[df["Ano"].isin(anos_comparacao)].dropna(subset=["Data", "Fat.Real"])



#st.subheader("📊 Faturamento Anual")

color_map = {
    "2024": "#1f77b4",  # Azul (igual ao gráfico mensal)
    "2025": "#ff7f0e",  # Laranja (igual ao gráfico mensal)
}

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

# Posicionar o valor no topo da barra
fig.update_traces(textposition="outside")



# Layout limpo e estilizado
fig.update_layout(
    xaxis_title=None,
    yaxis_title=None,
    xaxis_tickangle=-45,
    showlegend=False,
    yaxis=dict(
        showticklabels=False,
        showgrid=False,
        zeroline=False
    )
)
# ==============================
# 📉 Gráfico horizontal: Total Anual 2024 vs 2025
# ==============================
# 📉 Gráfico horizontal minimalista com total anual (valores visíveis e cores mantidas)
#df_total = fat_mensal.groupby("Ano")["Fat.Real"].sum().reset_index()

#NOVO

# Total de faturamento por ano
#df_total = fat_mensal.groupby("Ano")["Fat.Real"].sum().reset_index()
df_total = fat_mensal.groupby("Ano")["Fat.Real"].sum().reset_index()

# Calcular quantidade de lojas
df_lojas = df_anos.groupby("Ano")["Loja"].nunique().reset_index()
df_lojas.columns = ["Ano", "Qtd_Lojas"]

# Garantir que ambas as colunas 'Ano' são do mesmo tipo
df_total["Ano"] = df_total["Ano"].astype(int)
df_lojas["Ano"] = df_lojas["Ano"].astype(int)


# Junta com quantidade de lojas
df_total = df_total.merge(df_lojas, on="Ano", how="left")
#df_total["AnoTexto"] = df_total.apply(
 #   lambda row: f"{row['Ano']}  R$ {row['Fat.Real']/1_000_000:,.1f} Mi".replace(",", "."), axis=1
#)

df_total["AnoTexto"] = df_total.apply(
    lambda row: f"{int(row['Ano'])}                                                      R$ {row['Fat.Real']/1_000_000:,.1f} Mi".replace(",", "."), axis=1
)

df_total["Ano"] = df_total["Ano"].astype(str)
fig_total = px.bar(
    df_total,
    x="Fat.Real",
    y="Ano",
    #title=None,	
    orientation="h",
    color="Ano",  # Mantém as cores iguais ao gráfico mensal
    text="AnoTexto",  # 👈 usa a nova coluna,
    color_discrete_map=color_map
	
)
# 🔥 Remove título de eixos e legenda
fig_total.update_traces(
    textposition="inside",
    textfont=dict(size=16, color="white"),
    insidetextanchor="start",
    showlegend=False
)

# Estilo da barra
fig_total.update_traces(
    textposition="outside",  # Valor do lado de fora
    insidetextanchor="start",
    textfont=dict(size=16),
    showlegend=False      # reforço	
)
# Ano dentro da barra (ex: 2025)
for i, row in df_total.iterrows():
    fig_total.add_annotation(
        x=0.1,
        y=row["Ano"],
       # text=f"<b>{int(row['Ano'])}</b>",  # remove o .0
        text=row["AnoTexto"],  # texto com valor formatado
	showarrow=False,
        xanchor="left",
        yanchor="middle",
        font=dict(color="white", size=16),
        xref="x",
        yref="y"
    )
# Qtd de lojas ao final da barra (ex: 10 lojas)
for i, row in df_total.iterrows():
    fig_total.add_annotation(
        x=row["Fat.Real"],
        y=row["Ano"],
	showarrow=False,  
        text=f"{int(row['Qtd_Lojas'])} Lojas",  # remove o .0
        xanchor="left",
        yanchor="bottom",
	yshift=-8,     
        font=dict(color="red", size=16,weight="bold"),
        xref="x",
        yref="y"
    )

fig_total.update_layout(
    height=130,
    margin=dict(t=0, b=0, l=0, r=0),
    title=None,
    xaxis=dict(visible=False),
    yaxis=dict(
        showticklabels=False,
        showgrid=False,
        zeroline=False
    ),
    yaxis_title=None,     # ✅ remove "Ano" da lateral
    showlegend=False,     # ✅ remove a legenda de cores
    plot_bgcolor="rgba(0,0,0,0)"
)
# Exibir no Streamlit

st.subheader("Faturamento Anual")
st.plotly_chart(fig_total, use_container_width=True)

st.markdown("---")
st.subheader("Faturamento Mensal")
st.plotly_chart(fig, use_container_width=True)
# =========================
# 📋 Faturamento Real por Loja e Mês (com totais e exportação)
# =========================
import io

# 1. Prepara os dados com todos os anos disponíveis
df_anos["Ano"] = df_anos["Data"].dt.year
anos_disponiveis = sorted(df_anos["Ano"].dropna().unique())

# 2. Permitir seleção dos anos
anos_selecionados = st.multiselect("🗓️ Selecione os anos que deseja exibir", options=anos_disponiveis, default=anos_disponiveis)

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    for ano in anos_selecionados:
        df_fat = df_anos[df_anos["Ano"] == ano].copy()
        df_fat["Loja"] = df_fat["Loja"].astype(str).str.strip().str.lower().str.title()
        df_fat["Fat.Real"] = pd.to_numeric(df_fat["Fat.Real"], errors="coerce")

        # 3. Traduzir meses para português
        meses_pt = {
            "January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril",
            "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto",
            "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"
        }
        df_fat["Mês"] = df_fat["Data"].dt.strftime("%m - %B")
        df_fat["Mês"] = df_fat["Mês"].apply(lambda x: f"{x[:6]}{meses_pt.get(x[6:], x[6:])}")

        # 4. Tabela dinâmica
        tabela_fat_real = df_fat.pivot_table(
            index="Loja",
            columns="Mês",
            values="Fat.Real",
            aggfunc="sum",
            fill_value=0
        )

        # 5. Totais
        linha_total = tabela_fat_real.sum().to_frame().T
        linha_total.index = ["Total Geral"]
        coluna_total = tabela_fat_real.sum(axis=1)
        tabela_fat_real.insert(0, "Total Geral", coluna_total)
        linha_total.insert(0, "Total Geral", coluna_total.sum())
        tabela_com_total = pd.concat([linha_total, tabela_fat_real])

        # 6. Formatação brasileira
        tabela_formatada = tabela_com_total.applymap(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        # 7. Mostrar no app
        st.markdown("---")
        st.subheader(f"📋 Faturamento Real por Loja e Mês - {ano}")
        st.dataframe(tabela_formatada)

        # 8. Gravar no Excel
        tabela_com_total.to_excel(writer, sheet_name=f"Faturamento_{ano}")

# 9. Botão de download final
st.markdown("---")
st.download_button(
    label="📥 Baixar Excel com Totais por Ano",
    data=buffer.getvalue(),
    file_name="faturamento_real_totais_por_ano.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# ================================
# Aba 2: Graficos Trimestrais
# ================================

with aba2:
   
    st.info("Coloque aqui um gráfico de barras ou linhas mês a mês.")
    
# ================================
# Aba 3: Analise Mensal
# ================================

with aba3:
   
    st.info("Ideal para mostrar evolução por ano ou por trimestre.")

# ================================
# Aba 4: Analise Lojas
# ================================

with aba4:
   
    st.info("Você pode colocar tabelas detalhadas e botões de download aqui.")
