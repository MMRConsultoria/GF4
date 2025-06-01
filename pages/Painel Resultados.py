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
  

    # 🔧 Padroniza nomes de Loja e Grupo
    df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.title()
    df_empresa["Grupo"] = df_empresa["Grupo"].astype(str).str.strip().str.upper()

    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.title()
    df_anos["Grupo"] = df_anos["Grupo"].astype(str).str.strip().str.upper()



    # Normaliza dados
    
    df_anos["Grupo"] = df_anos["Grupo"].str.split("-").str[0].str.strip()
    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.lower().str.title()
    df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower().str.title()
    df_anos["Fat.Total"] = pd.to_numeric(df_anos["Fat.Total"], errors="coerce")
    df_anos["Fat.Real"] = pd.to_numeric(df_anos["Fat.Real"], errors="coerce")
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Mês Num"] = df_anos["Data"].dt.month
    df_anos["Mês Nome"] = df_anos["Data"].dt.strftime('%B')
    df_anos["Mês"] = df_anos["Data"].dt.strftime('%m/%Y')
    df_anos["Dia"] = df_anos["Data"].dt.strftime('%d/%m/%Y')
    
    # 🔗 Lojas ativas (logo após normalizar os dados)
    todas_lojas = df_empresa[
        df_empresa["Lojas Ativas"].astype(str).str.strip().str.lower() == "ativa"
    ][["Loja", "Grupo", "Tipo"]].drop_duplicates()
 

    
    # === FILTROS ===
    #anos_disponiveis = sorted(df_anos["Ano"].unique(), reverse=True)
    #ano_opcao = st.multiselect("📅 Selecione ano/mês(s):", options=anos_disponiveis, default=anos_disponiveis, key="ano_aba3")
    anos_disponiveis = sorted(df_anos["Ano"].unique(), reverse=True)
    ultimo_ano = anos_disponiveis[0] if anos_disponiveis else datetime.today().year
    ano_opcao = st.multiselect("📅 Selecione ano/mês(s):", options=anos_disponiveis, default=[ultimo_ano], key="ano_aba3")
    
   
    df_filtrado = df_anos[df_anos["Ano"].isin(ano_opcao)]
    df_filtrado["Grupo"] = df_filtrado["Grupo"].astype(str).str.strip().str.upper()
    df_filtrado["Loja"] = df_filtrado["Loja"].astype(str).str.strip().str.title()
   
    meses_dict = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
                  7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}

    meses_disponiveis = sorted(df_filtrado["Mês Num"].unique())
    #meses_nomes_disponiveis = [meses_dict[m] for m in meses_disponiveis]
    #meses_selecionados = st.multiselect("", options=meses_nomes_disponiveis, default=meses_nomes_disponiveis, key="meses_aba3")
    meses_nomes_disponiveis = [meses_dict[m] for m in meses_disponiveis]
    mes_atual_nome = meses_dict[datetime.today().month]
    meses_selecionados = st.multiselect("", options=meses_nomes_disponiveis, default=[mes_atual_nome], key="meses_aba3")
    
   
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


    hoje = date.today()
    with col4:
        agrupamento = st.radio(" ", ["Ano", "Mês", "Dia"], horizontal=True, key="agrup_aba4")

    meses_numeros = [k for k, v in meses_dict.items() if v in meses_selecionados]

    

    # 📌 Define data_fim como a data mais recente do DataFrame
    datas_disponiveis = sorted(pd.to_datetime(df_anos["Data"].dropna().unique()))
    data_fim = datas_disponiveis[-1].date() if datas_disponiveis else date.today()


    # 🔢 Filtra só as lojas ativas
    lojas_ativas = df_empresa[
        df_empresa["Lojas Ativas"].astype(str).str.strip().str.lower() == "ativa"
     ][["Loja", "Grupo", "Tipo"]].drop_duplicates()

   


        # 🔄 Aplica o filtro principal com base no período
    if agrupamento == "Dia" and modo_visao == "Por Grupo":
        data_selecionada = pd.to_datetime(data_fim)

       
        

        

        # 🧹 Trata colunas duplicadas de Grupo
        if "Grupo_x" in df_filtrado.columns:
            df_filtrado["Grupo"] = df_filtrado["Grupo_x"]
        if "Grupo_y" in df_filtrado.columns:
            df_filtrado["Grupo"] = df_filtrado["Grupo"].combine_first(df_filtrado["Grupo_y"])
        df_filtrado.drop(columns=[col for col in df_filtrado.columns if col.startswith("Grupo_")], inplace=True)

        # ✅ Preenche dados nulos
        for col in ["Fat.Total", "Fat.Real", "Serv/Tx", "Ticket"]:
            if col in df_filtrado.columns:
                df_filtrado[col] = df_filtrado[col].fillna(0)

        # 🧮 Preenche colunas de data para manter compatibilidade
        df_filtrado["Ano"] = data_selecionada.year
        df_filtrado["Mês Num"] = data_selecionada.month
        df_filtrado["Mês Nome"] = data_selecionada.strftime('%B')
        df_filtrado["Mês"] = data_selecionada.strftime('%m/%Y')
        df_filtrado["Dia"] = data_selecionada.strftime('%d/%m/%Y')
        df_filtrado["Agrupador"] = data_selecionada.strftime('%d/%m/%Y')
        df_filtrado["Ordem"] = data_selecionada

    



        # Filtro de datas (Dia)
    if agrupamento == "Dia":
        df_filtrado = df_filtrado[
            (df_filtrado["Data"] >= pd.to_datetime(data_inicio)) &
            (df_filtrado["Data"] <= pd.to_datetime(data_fim))
        ]



        # 🔥 Garante que grupos ativos SEM movimento apareçam com 0
        grupos_ativos = df_empresa[
            df_empresa["Grupo Ativo"].astype(str).str.strip().str.lower() == "ativo"
        ]["Grupo"].dropna().unique()

        grupos_presentes = df_filtrado["Grupo"].dropna().unique()
        grupos_faltando = list(set(grupos_ativos) - set(grupos_presentes))

        if grupos_faltando:
            df_faltando = pd.DataFrame({
                "Grupo": grupos_faltando,
                "Loja": [f"Grupo_{g}_sem_loja" for g in grupos_faltando],
                "Tipo": None,
                "Data": data_selecionada,
                "Fat.Total": 0,
                "Fat.Real": 0,
                "Serv/Tx": 0,
                "Ticket": 0,
                "Ano": data_selecionada.year,
                "Mês Num": data_selecionada.month,
                "Mês Nome": data_selecionada.strftime('%B'),
                "Mês": data_selecionada.strftime('%m/%Y'),
                "Dia": data_selecionada.strftime('%d/%m/%Y'),
                "Agrupador": data_selecionada.strftime('%d/%m/%Y'),
                "Ordem": data_selecionada
            })

            df_filtrado = pd.concat([df_filtrado, df_faltando], ignore_index=True)



        # Preenche colunas numéricas com 0 para lojas sem movimento
        for col in ["Fat.Total", "Fat.Real", "Serv/Tx", "Ticket"]:
            if col in df_filtrado.columns:
                df_filtrado[col] = df_filtrado[col].fillna(0)

        
        



    # ✅ Só aplica o filtro de mês quando o agrupamento for "Mês" ou "Dia"
    if agrupamento in ["Mês", "Dia"]:
        df_filtrado = df_filtrado[df_filtrado["Mês Num"].isin(meses_numeros)]
    else:
        # Ignora filtro de mês — mantém todos os meses dos anos selecionados
        pass
    
    # 🧠 Garante seleção válida
    anos_validos = [a for a in ano_opcao if isinstance(a, int)]

    # 📅 Define intervalo padrão com base no agrupamento
    if agrupamento == "Ano" and anos_validos:
        data_inicio_padrao = date(min(anos_validos), 1, 1)
        data_fim_padrao = date(max(anos_validos), 12, 31)
    elif agrupamento == "Mês":
        data_inicio_padrao = datetime(data_minima.year, data_minima.month, 1).date()
        data_fim_padrao = data_maxima
    else:  # Dia
        ontem = date.today() - timedelta(days=1)
        data_inicio_padrao = max(data_minima, min(ontem, data_maxima))
        data_fim_padrao = data_inicio_padrao

    # ✅ Campo de data com valores padrão corretos
    data_inicio, data_fim = st.date_input(
        "",
        value=[data_inicio_padrao, data_fim_padrao],
        min_value=data_minima,
        max_value=data_maxima
    )

    if agrupamento == "Ano" and ano_opcao:
        df_filtrado = df_filtrado[df_filtrado["Ano"].isin(ano_opcao)]

    
    elif agrupamento == "Mês":
        data_inicio_padrao = datetime(data_minima.year, data_minima.month, 1).date()
        data_fim_padrao = data_maxima
    else:  # Dia
        data_inicio_padrao = hoje

        
        data_fim_padrao = hoje
   
   
   
    # ✅ Aplica o filtro de datas corretamente conforme o agrupamento
    if agrupamento == "Dia":
        df_filtrado = df_filtrado[
            (df_filtrado["Data"] >= pd.to_datetime(data_inicio)) &
            (df_filtrado["Data"] <= pd.to_datetime(data_fim))
        ]
    elif agrupamento == "Mês":
        periodo_inicio = pd.to_datetime(data_inicio).to_period("M")
        periodo_fim = pd.to_datetime(data_fim).to_period("M")
        df_filtrado = df_filtrado[
            (df_filtrado["Data"].dt.to_period("M") >= periodo_inicio) &
            (df_filtrado["Data"].dt.to_period("M") <= periodo_fim)
        ]
   
    df_filtrado["Grupo"] = df_filtrado["Grupo"].astype(str).str.strip().str.upper()
    df_filtrado["Loja"] = df_filtrado["Loja"].astype(str).str.strip().str.title()

    
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

    # 🔥 Remove agrupadores inválidos (None, nan, vazio)
    df_filtrado = df_filtrado[~df_filtrado["Agrupador"].isin([None, "None", "nan", "NaN", ""])]


  # Garante a ordem correta
    ordem = (
        df_filtrado[["Agrupador", "Ordem"]]
        .drop_duplicates()
        .dropna()
        .sort_values("Ordem", ascending=False)
    )["Agrupador"].tolist()
   
    ordem = [c for c in ordem if pd.notnull(c) and str(c).strip().lower() not in ["none", "nan", ""]]

    




    if modo_visao == "Por Loja":
        lojas_com_venda = df_filtrado["Loja"].unique()
        lojas_sem_venda = todas_lojas[~todas_lojas["Loja"].isin(lojas_com_venda)]

        if not lojas_sem_venda.empty:
            df_sem_venda = pd.DataFrame({
                "Loja": lojas_sem_venda["Loja"],
                "Grupo": lojas_sem_venda["Grupo"],
                "Tipo": lojas_sem_venda["Tipo"],
                "Fat.Total": 0,
                "Fat.Real": 0,
                "Data": pd.NaT,
                "Ano": None,
                "Mês Num": None,
                "Mês Nome": None,
                "Mês": None,
                "Dia": None,
                "Agrupador": None,
                "Ordem": None
            })

            df_filtrado = pd.concat([df_filtrado, df_sem_venda], ignore_index=True)


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
            tabela.columns = [
                col if pd.notnull(col) and str(col).strip().lower() not in ["none", "nan", ""] else "Excluir"
                for col in tabela.columns
            ]
            tabela = tabela.loc[:, tabela.columns != "Excluir"]


            
            colunas_intercaladas = []
            for col in ordem:
                colunas_intercaladas.append(f"{col} (Bruto)")
                colunas_intercaladas.append(f"{col} (Real)")

         # 🔥 Aqui já remove qualquer None no processo de montar a lista
            colunas_intercaladas = [c for c in colunas_intercaladas if pd.notnull(c) and str(c).strip().lower() not in ["none", "nan", ""]]
            tabela = tabela[[c for c in colunas_intercaladas if c in tabela.columns]]

     


    tabela = tabela[[col for col in tabela.columns if pd.notnull(col) and str(col).strip().lower() not in ["none", "nan", ""]]]

    if modo_visao == "Por Loja":
        tabela = tabela.fillna(0)
        tabela.index.name = "Loja"

        if exibir_total:
            tabela["Total"] = tabela.sum(axis=1)
            colunas_finais = ["Total"] + [col for col in tabela.columns if col != "Total"]
            tabela = tabela[colunas_finais]

        total_geral = pd.DataFrame(tabela.sum(numeric_only=True)).T
        total_geral.index = ["Total Geral"]

        tabela_final = pd.concat([total_geral, tabela])

    elif modo_visao == "Por Grupo":
        tabela = tabela.fillna(0)
        tabela.index.name = "Grupo"

        if exibir_total:
            tabela["Total"] = tabela.sum(axis=1)
            colunas_finais = ["Total"] + [col for col in tabela.columns if col != "Total"]
            tabela = tabela[colunas_finais]

        total_geral = pd.DataFrame(tabela.sum(numeric_only=True)).T
        total_geral.index = ["Total Geral"]

        tabela_final = pd.concat([total_geral, tabela])

        # 🔍 Ordena a tabela pelo valor mais recente
        colunas_numericas = tabela_final.select_dtypes(include='number').columns.tolist()
        if colunas_numericas:
            col_mais_recente = colunas_numericas[0]
            if "Total" in col_mais_recente and len(colunas_numericas) > 1:
                col_mais_recente = colunas_numericas[1]
            tabela_final = tabela_final.sort_values(by=col_mais_recente, ascending=False)

        tabela_formatada = tabela_final.applymap(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if isinstance(x, (float, int)) else x
        )


    
  
    # ✅ 🔥 Limpeza imediata e universal após criar tabela_final
    tabela_final.columns.name = None  # Remove nome do eixo das colunas
    tabela_final = tabela_final.loc[:, ~tabela_final.columns.isnull()]  # Remove colunas None (NaN real)
    tabela_final = tabela_final.drop(columns=[None, 'None', 'nan'], errors='ignore')

    if modo_visao == "Por Loja":
        lojas_ativas = todas_lojas["Loja"].tolist()
        lojas_existentes = tabela.index.tolist()

        lojas_faltando = list(set(lojas_ativas) - set(lojas_existentes))

        if lojas_faltando:
            df_sem_venda = pd.DataFrame(0, index=lojas_faltando, columns=tabela.columns)
            tabela = pd.concat([tabela, df_sem_venda])

            tabela = tabela.sort_index()

        colunas_ordenadas = [col for col in ordem if col in tabela.columns or f"{col} (Bruto)" in tabela.columns or f"{col} (Real)" in tabela.columns]
        todas_colunas = []
        for col in colunas_ordenadas:
            if tipo_metrica == "Ambos":
                if f"{col} (Bruto)" in tabela.columns: todas_colunas.append(f"{col} (Bruto)")
                if f"{col} (Real)" in tabela.columns: todas_colunas.append(f"{col} (Real)")
            else:
                todas_colunas.append(col)


        # 🔥 Limpeza definitiva de colunas inválidas
        tabela = tabela[todas_colunas]
        tabela = tabela.loc[:, ~tabela.columns.isnull()]  # Remove colunas None (NaN real)
  
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

  
    
    
    if 'tabela_formatada' in locals():
        st.subheader("")
        st.dataframe(tabela_formatada, use_container_width=True)
    else:
        st.warning("")
   

import io
import itertools
import pandas as pd

buffer = io.BytesIO()

# 🔥 Limpeza da Tabela Empresa
df_empresa = df_empresa.dropna(how='all')
df_empresa = df_empresa[df_empresa["Loja"].notna() & (df_empresa["Loja"].astype(str).str.strip() != "")]

df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower().str.title()
df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.lower().str.title()

# 🔗 Pega a lista de lojas ativas
todas_lojas = df_empresa[
    df_empresa["Lojas Ativas"].astype(str).str.strip().str.lower() == "Ativa"
][["Loja", "Grupo", "Tipo"]].drop_duplicates()



# 🔥 Verifica se deve calcular acumulado
data_max = pd.to_datetime(data_fim)
hoje = pd.to_datetime(pd.Timestamp.now()).tz_localize(None)

mostrar_acumulado = (
    agrupamento == "Dia" and
    data_max.year == hoje.year and
    data_max.month == hoje.month and
    data_inicio.year == hoje.year and
    data_inicio.month == hoje.month
)

# 🔥 Inicializa acumulados vazios
acumulado_por_tipo = pd.DataFrame(columns=["Tipo", "Acumulado no Mês Tipo"])
acumulado_por_grupo = pd.DataFrame(columns=["Grupo", "Acumulado no Mês"])
acumulado_por_loja = pd.DataFrame(columns=["Loja", "Acumulado no Mês"])

# 🔥 Criação da tabela_exportar
if modo_visao == "Por Loja":
    tabela_final.index.name = "Loja"
    tabela_exportar = tabela_final.reset_index()

    tabela_exportar = tabela_exportar.merge(
        df_empresa[["Loja", "Grupo", "Tipo"]],
        on="Loja", how="left"
    )

    tabela_exportar = tabela_exportar[["Grupo", "Loja", "Tipo"] + 
                                      [col for col in tabela_exportar.columns if col not in ["Grupo", "Loja", "Tipo"]]]

elif modo_visao == "Por Grupo":
    tabela_final.index.name = "Grupo"
    tabela_exportar = tabela_final.reset_index()

    if modo_visao == "Por Grupo":
        tabela_exportar["Tipo"] = df_empresa.groupby("Grupo")["Tipo"].agg(lambda x: x.mode().iloc[0] if not x.mode().empty else None).reindex(tabela_exportar["Grupo"]).values
 
    #st.write("🚧 Debug Grupo", tabela_exportar)
    #st.write("📄 df_empresa", df_empresa)

# 🔥 Define qual coluna usar para o acumulado (Bruto ou Real)
coluna_acumulado = "Fat.Total"  # 🔥 Para Acumulado no mês pelo Bruto (com gorjeta)
# Se quiser mudar para Real, basta trocar para:
# coluna_acumulado = "Fat.Real"

# 🔥 Cálculo do Acumulado no Mês
if mostrar_acumulado:
    primeiro_dia_mes = data_max.replace(day=1)
    df_acumulado = df_anos[
        (df_anos["Data"] >= primeiro_dia_mes) &
        (df_anos["Data"] <= data_max)
    ].copy()

    df_acumulado = df_acumulado.merge(
        df_empresa[["Loja", "Grupo", "Tipo"]].drop_duplicates(),
        on="Loja", how="left", suffixes=('', '_drop')
    )
    df_acumulado = df_acumulado.loc[:, ~df_acumulado.columns.str.endswith('_drop')]

    acumulado_por_tipo = df_acumulado.groupby("Tipo")[coluna_acumulado].sum().reset_index().rename(
        columns={coluna_acumulado: "Acumulado no Mês Tipo"}
    )
    acumulado_por_grupo = df_acumulado.groupby("Grupo")[coluna_acumulado].sum().reset_index().rename(
        columns={coluna_acumulado: "Acumulado no Mês"}
    )
    acumulado_por_loja = df_acumulado.groupby("Loja")[coluna_acumulado].sum().reset_index().rename(
        columns={coluna_acumulado: "Acumulado no Mês"}
    )




# 🔥 Merge dos acumulados
if mostrar_acumulado:
    if modo_visao == "Por Loja":
        tabela_exportar = tabela_exportar.merge(acumulado_por_loja, on="Loja", how="left")
    if modo_visao == "Por Grupo":
        tabela_exportar = tabela_exportar.merge(acumulado_por_grupo, on="Grupo", how="left")
    tabela_exportar = tabela_exportar.merge(acumulado_por_tipo, on="Tipo", how="left")
else:
    if modo_visao == "Por Loja":
        tabela_exportar["Acumulado no Mês"] = None
    if modo_visao == "Por Grupo":
        tabela_exportar["Acumulado no Mês"] = None
    tabela_exportar["Acumulado no Mês Tipo"] = None

# 🔥 Remove a coluna "Acumulado no Mês Tipo" do corpo
tabela_exportar_sem_tipo = tabela_exportar.drop(columns=["Acumulado no Mês Tipo","Tipo"], errors="ignore")

# 🔍 Ordenação pela data mais recente
colunas_data = [col for col in tabela_exportar_sem_tipo.columns if "/" in col]

def extrair_data(col):
    try:
        col_limpo = col.split(" ")[0].strip()
        return pd.to_datetime(col_limpo, format="%d/%m/%Y", dayfirst=True, errors="coerce")
    except:
        return pd.NaT

colunas_validas = [col for col in colunas_data if not pd.isna(extrair_data(col))]
coluna_mais_recente = max(colunas_validas, key=lambda x: extrair_data(x)) if colunas_validas else None

if coluna_mais_recente:
    tabela_exportar_sem_tipo = tabela_exportar_sem_tipo.sort_values(by=coluna_mais_recente, ascending=False)

# 🔥 Remove colunas 100% vazias
tabela_exportar_sem_tipo = tabela_exportar_sem_tipo.dropna(axis=1, how="all")
# ✅ Aqui faz a substituição de NaN na coluna Acumulado no Mês
if "Acumulado no Mês" in tabela_exportar_sem_tipo.columns:
    tabela_exportar_sem_tipo["Acumulado no Mês"] = tabela_exportar_sem_tipo["Acumulado no Mês"].fillna(0)
tabela_exportar_sem_tipo = tabela_exportar_sem_tipo.rename(columns=lambda x: x.replace('Bruto', 'Bruto- Com Gorjeta').replace('Real', 'Real-Sem Gorjeta'))


if modo_visao == "Por Loja":
    lojas_existentes = tabela_final.index.tolist()
    lojas_ativas = todas_lojas["Loja"].tolist()

    lojas_faltando = list(set(lojas_ativas) - set(lojas_existentes))

    if lojas_faltando:
        # 🔥 Cria dataframe das lojas sem venda
        df_sem_venda = pd.DataFrame(index=lojas_faltando)

        for col in tabela_final.columns:
            df_sem_venda[col] = 0

        tabela_final = pd.concat([tabela_final, df_sem_venda])

    tabela_final = tabela_final.sort_index()



# ✅ 🔥 REMOVE coluna None
tabela_final.columns.name = None  # Remove nome do eixo das colunas
tabela_final = tabela_final.loc[:, ~tabela_final.columns.isnull()]
tabela_final = tabela_final.drop(columns=[None, 'None', 'nan'], errors='ignore')








# 🔥 Geração do Excel
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    tabela_exportar_sem_tipo.to_excel(writer, sheet_name="Faturamento", index=False, startrow=0)

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

    linha = 1
    num_colunas = len(tabela_exportar_sem_tipo.columns)




    # 🔥 Determina a coluna de identificação (Loja ou Grupo)
    coluna_id = "Loja" if "Loja" in tabela_exportar_sem_tipo.columns else "Grupo"

    
    # 🔥 Subtotal por Tipo (Sempre aparece)
    for tipo_atual in sorted(tabela_exportar["Tipo"].dropna().unique()):
        linhas_tipo = tabela_exportar_sem_tipo[
            (tabela_exportar_sem_tipo["Grupo"].isin(
                df_empresa[df_empresa["Tipo"] == tipo_atual]["Grupo"].unique()
            )) &
      # ~tabela_exportar_sem_tipo[coluna_id].astype(str).str.contains("Subtotal|Total", case=False, na=False)
             ~tabela_exportar_sem_tipo[coluna_id].astype(str).str.contains("Subtotal|Total", case=False, na=False)
        ]
        
        qtd_lojas_tipo = df_empresa[
            (df_empresa["Tipo"] == tipo_atual) &
            (df_empresa["Lojas Ativas"].astype(str).str.strip().str.lower() == "ativa")
        ]["Loja"].nunique()
        soma_colunas = linhas_tipo.select_dtypes(include='number').sum()

        linha_tipo = [f"Tipo: {tipo_atual}", f"Lojas: {qtd_lojas_tipo}"]
        linha_tipo += [soma_colunas.get(col, "") for col in tabela_exportar_sem_tipo.columns[2:]]

        for col_num, val in enumerate(linha_tipo):
            if isinstance(val, (int, float)) and not pd.isna(val):
                worksheet.write_number(linha, col_num, val, subtotal_format)
            else:
                worksheet.write(linha, col_num, str(val), subtotal_format)
        linha += 1

    # 🔢 Filtra só as lojas ativas
    lojas_ativas = df_empresa[
       df_empresa["Lojas Ativas"].astype(str).str.strip().str.lower() == "ativa"
    ][["Loja", "Grupo", "Tipo"]].drop_duplicates()
    
    
    # 🔝 Total Geral
    linhas_validas = ~tabela_exportar_sem_tipo[coluna_id].astype(str).str.contains("Total|Subtotal", case=False, na=False)


    df_para_total = tabela_exportar_sem_tipo[linhas_validas]

    soma_total = df_para_total.select_dtypes(include='number').sum()
    # 🔢 Conta todas as lojas ativas (sem duplicar)
    total_lojas_ativas = lojas_ativas["Loja"].nunique()
    linha_total = [f"Total Geral", f"Lojas: {total_lojas_ativas}"]
    linha_total += [soma_total.get(col, "") for col in tabela_exportar_sem_tipo.columns[2:]]

    for col_num, val in enumerate(linha_total):
        if isinstance(val, (int, float)) and not pd.isna(val):
            worksheet.write_number(linha, col_num, val, totalgeral_format)
        else:
            worksheet.write(linha, col_num, str(val), totalgeral_format)
    linha += 1

    # 🔢 Subtotal por Grupo
   

    # 🔢 Filtra a base para considerar apenas as lojas ativas

    # 🔒 Garante que só aplica o filtro se "Loja" existir
    if "Loja" in tabela_exportar_sem_tipo.columns and "Loja" in lojas_ativas.columns:
        df_ativos = tabela_exportar_sem_tipo[
            tabela_exportar_sem_tipo["Loja"].isin(lojas_ativas["Loja"])
        ].copy()
    else:
        df_ativos = tabela_exportar_sem_tipo.copy()

    
   
    # 🔢 Calcula subtotais por grupo (soma de todas as colunas numéricas)
    df_numerico = df_ativos.select_dtypes(include='number')
    df_numerico["Grupo"] = df_ativos["Grupo"]

    subtotais = df_numerico.groupby("Grupo").sum().sum(axis=1).reset_index()
    subtotais.columns = ["Grupo", "Subtotal"]

    # 🔢 Junta com o Tipo e mantém somente grupos que aparecem nos dados ativos
    grupos_com_dados = df_ativos["Grupo"].dropna().unique().tolist()

    grupos_tipo = (
        lojas_ativas[["Grupo", "Tipo"]]
        .dropna()
        .drop_duplicates()
        .merge(subtotais, on="Grupo", how="left")
        .query("Grupo in @grupos_com_dados")
        .sort_values(by=["Tipo", "Subtotal"], ascending=[True, False])
    )

    grupos_ordenados = grupos_tipo["Grupo"].tolist()
    
    # ✅ Adiciona acumulado dos grupos ativos, se visão por Grupo e agrupado por Dia
    if modo_visao == "Por Grupo" and agrupamento == "Dia":
        grupos_ativos = df_empresa[
            df_empresa["Grupo Ativo"].astype(str).str.strip().str.lower() == "ativo"
        ]["Grupo"].dropna().unique()

        df_acumulado_grupo = (
            df_filtrado[df_filtrado["Grupo"].isin(grupos_ativos)]
            .groupby("Grupo")[["Fat.Total", "Fat.Real", "Serv/Tx", "Ticket"]]
            .sum()
            .reset_index()
        )

        df_acumulado_grupo["Loja"] = "ACUMULADO GRUPO ATIVO"
        df_acumulado_grupo["Tipo"] = None
        df_acumulado_grupo["Data"] = None
        df_acumulado_grupo["Ano"] = None
        df_acumulado_grupo["Mês Num"] = None
        df_acumulado_grupo["Mês Nome"] = None
        df_acumulado_grupo["Mês"] = None
        df_acumulado_grupo["Dia"] = None
        df_acumulado_grupo["Agrupador"] = "ACUMULADO"
        df_acumulado_grupo["Ordem"] = 99999999

        df_filtrado = pd.concat([df_filtrado, df_acumulado_grupo], ignore_index=True)

    for grupo_atual, cor in zip(grupos_ordenados, cores_grupo):






    
        linhas_grupo = tabela_exportar_sem_tipo[
            (tabela_exportar_sem_tipo["Grupo"] == grupo_atual) &
            ~tabela_exportar_sem_tipo[coluna_id].astype(str).str.contains("Subtotal|Total", case=False, na=False)
        ]

        qtd_lojas_tipo = lojas_ativas[lojas_ativas["Tipo"] == tipo_atual]["Loja"].nunique()
        grupo_format = workbook.add_format({
            'bg_color': cor, 'border': 1, 'num_format': 'R$ #,##0.00'
        })

        for _, row in linhas_grupo.iterrows():
            row = row.copy()






            if modo_visao == "Por Grupo":
                # 🟡 Insere quantidade de lojas ativas ao lado do grupo
                qtd_lojas = lojas_ativas[lojas_ativas["Grupo"] == grupo_atual]["Loja"].nunique()
                row.iloc[0] = f"{grupo_atual} - Loja: {qtd_lojas}"

            for col_num, val in enumerate(row):
                if isinstance(val, (int, float)) and not pd.isna(val):
                    worksheet.write_number(linha, col_num, val, grupo_format)
                else:
                    worksheet.write(linha, col_num, str(val), grupo_format)
            linha += 1

        
        # ✅ Subtotal por grupo apenas no modo "Por Loja"
        if modo_visao == "Por Loja":
            soma_grupo = linhas_grupo.select_dtypes(include='number').sum()
            # 🔧 Calcula somente as lojas ativas no grupo
            qtd_lojas = lojas_ativas[lojas_ativas["Grupo"] == grupo_atual]["Loja"].nunique()
            
            linha_grupo = [f"Subtotal {grupo_atual}", f"Lojas: {qtd_lojas}"]
            linha_grupo += [soma_grupo.get(col, "") for col in tabela_exportar_sem_tipo.columns[2:]]

            for col_num, val in enumerate(linha_grupo):
                if isinstance(val, (int, float)) and not pd.isna(val):
                    worksheet.write_number(linha, col_num, val, subtotal_format)
                else:
                    worksheet.write(linha, col_num, str(val), subtotal_format)
            linha += 1

       

# 🔧 Ajustes visuais finais
worksheet.set_column(0, num_colunas, 18)
# Atualiza o cabeçalho para incluir a coluna % Participação
for col_num, header in enumerate(tabela_exportar_sem_tipo.columns):
    worksheet.write(0, col_num, header, header_format)
# 🔥 Adiciona o cabeçalho da coluna de participação
worksheet.write(0, num_colunas, "% Participação", header_format)

worksheet.hide_gridlines(option=2)

# 🔽 Botão Download
st.download_button(
    label="📥 Baixar Excel",
    data=buffer.getvalue(),
    file_name="faturamento_visual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
