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
aba1, aba2, aba3, aba4, aba5 = st.tabs([
    "📈 Gráficos Anuais",
    "📊 Gráficos Trimestrais",
    "📆 Relatórios Vendas",
    "📋 Relatório Vendas/Metas",
    "📋 Relatórios Financeiros"
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
# Aba 3: Relatórios Vendas
# ================================
with aba3:
    import pandas as pd
    import numpy as np
    import streamlit as st
    from datetime import datetime

    st.markdown("""
        <style>
        .stMultiSelect > div {
            background-color: #f5f5f5 !important;
            color: black !important;
            border: 1px solid #ccc !important;
            border-radius: 6px;
        }
    
        .stMultiSelect [data-baseweb="tag"] {
            background-color: #cccccc !important;
            color: black !important;
            font-weight: 600;
            border-radius: 6px;
            padding: 4px 10px;
        }
    
        .stSelectbox > div > div {
            background-color: #f5f5f5 !important;
            color: black !important;
            border: 1px solid #ccc !important;
        }
    
        .stDateInput > div > div {
            background-color: #f5f5f5 !important;
            color: black !important;
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

    # Filtros principais
    col1, col2, col3 = st.columns([2, 2, 2])
    with col1:
        modo_exibicao = st.selectbox("🔀 Ver por:", ["Loja", "Grupo"], key="modo_exibicao_relatorio")
    with col2:
        modo_periodo = st.selectbox("🕒 Período:", ["Diário", "Mensal", "Anual"], key="modo_periodo_relatorio")
    
    # Container para o filtro de datas ou períodos, sempre no mesmo lugar
    with col3:
        if modo_periodo == "Diário":
            data_min = df_vendas["Data"].min()
            data_max = df_vendas["Data"].max()
            data_inicio, data_fim = st.date_input(
                "📅 Intervalo de datas:",
                (data_max, data_max),
                data_min,
                data_max,
                key="data_vendas_relatorio"
            )
            df_filtrado = df_vendas[
                (df_vendas["Data"] >= pd.to_datetime(data_inicio)) &
                (df_vendas["Data"] <= pd.to_datetime(data_fim))
            ]
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
                default=[datetime.today().strftime("%m/%Y")],
                key="meses_relatorio"
            )
            df_filtrado = df_vendas[df_vendas["Mes/Ano"].isin(meses_selecionados)]
            df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%m/%Y")
    
        elif modo_periodo == "Anual":
            df_vendas["Ano"] = df_vendas["Data"].dt.strftime("%Y")
            anos_disponiveis = sorted(df_vendas["Ano"].unique())
            anos_selecionados = st.multiselect(
                "📅 Selecione os anos:",
                options=anos_disponiveis,
                default=[datetime.today().strftime("%Y")],
                key="anos_relatorio"
            )
            df_filtrado = df_vendas[df_vendas["Ano"].isin(anos_selecionados)]
            df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%Y")
    


    

    # Agrupamento
    chaves = ["Loja", "Grupo"] if modo_exibicao == "Loja" else ["Grupo"]
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

    colunas_periodo = sorted([c for c in df_pivot.columns if c not in ["Loja", "Grupo"]], key=ordenar_datas)

    # Garante que colunas existam
    if "Loja" not in df_pivot.columns:
        df_pivot["Loja"] = ""
    if "Grupo" not in df_pivot.columns:
        df_pivot["Grupo"] = ""

    # Define ordem final: Grupo, Loja, depois períodos
    colunas_finais = ["Grupo", "Loja"] + colunas_periodo
    df_final = df_pivot[colunas_finais].copy()

    # Total acumulado (última coluna)
    ultima_coluna_valor = colunas_periodo[-1]
    df_final["__ordem"] = df_final[ultima_coluna_valor]

    # Ordena do maior para o menor
    df_final = df_final.sort_values(by="__ordem", ascending=False).drop(columns="__ordem").reset_index(drop=True)

    # === Linha de Lojas Ativas (quantas lojas venderam algo em cada período) ===
    df_lojas_por_periodo = df_filtrado.groupby("Período")["Loja"].nunique()
    
    # Monta linha com mesma estrutura do df_final
    linha_lojas = {col: "" for col in df_final.columns}
    linha_lojas["Grupo"] = "Lojas Ativas"
    linha_lojas["Loja"] = ""
    for periodo in df_lojas_por_periodo.index:
        if periodo in linha_lojas:
            linha_lojas[periodo] = int(df_lojas_por_periodo[periodo])  # 👈 força inteiro
    
    # === Linha TOTAL ===
    linha_total = df_final.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
    linha_total["Grupo"] = "TOTAL"
    linha_total["Loja"] = ""
    
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
    
    df_formatado = df_formatado[["Grupo", "Loja"] + colunas_periodo]

    # Estilo para destacar TOTAL
    def aplicar_estilo(df):
        def estilo_linha(row):
            if row["Grupo"] == "TOTAL":
                return ["background-color: #f0f0f0; font-weight: bold"] * len(row)
            else:
                return ["" for _ in row]
        return df.style.apply(estilo_linha, axis=1)
        
    # Exibição
    st.dataframe(
        aplicar_estilo(df_formatado),
        use_container_width=True,
        height=750
    )



# ================================
# Aba 4: Relatório Vendas/Metas
# ================================
with aba4:
    # Carrega dados
    df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
    df_vendas = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
    df_empresa["Loja"] = df_empresa["Loja"].str.strip().str.upper()
    df_empresa["Grupo"] = df_empresa["Grupo"].str.strip()
    df_vendas.columns = df_vendas.columns.str.strip()
    df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
    df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()
    df_vendas["Grupo"] = df_vendas["Grupo"].astype(str).str.strip()
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
    
    
    
    # Filtros
    data_min = df_vendas["Data"].min()
    data_max = df_vendas["Data"].max()
    col1, col2, col3 = st.columns([2, 2, 2])
    ultimo_dia_disponivel = df_vendas["Data"].max()
    data_min = df_vendas["Data"].min()
    data_max = df_vendas["Data"].max()
    
    with col1:
        data_inicio, data_fim = st.date_input("📅 Intervalo de datas:", (data_max, data_max), data_min, data_max)
    with col2:
        modo_exibicao = st.selectbox("🧭 Ver por:", ["Loja", "Grupo"])
    with col3:
        filtro_meta = st.selectbox("🎯 Mostrar:", ["Meta", "Sem Meta"])
    
    data_inicio_dt = pd.to_datetime(data_inicio)
    data_fim_dt = pd.to_datetime(data_fim)
    primeiro_dia_mes = data_fim_dt.replace(day=1)
    datas_periodo = pd.date_range(start=data_inicio_dt, end=data_fim_dt)
    
    # Base combinada com 0s
    df_lojas_grupos = df_empresa[["Loja", "Grupo"]].drop_duplicates()
    df_base_completa = pd.MultiIndex.from_product(
        [df_lojas_grupos["Loja"], datas_periodo], names=["Loja", "Data"]
    ).to_frame(index=False)
    df_base_completa = df_base_completa.merge(df_lojas_grupos, on="Loja", how="left")
    df_filtro_dias = df_vendas[(df_vendas["Data"] >= data_inicio_dt) & (df_vendas["Data"] <= data_fim_dt)]
    df_agrupado_dias = df_filtro_dias.groupby(["Data", "Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
    df_completo = df_base_completa.merge(df_agrupado_dias, on=["Data", "Loja", "Grupo"], how="left")
    df_completo["Fat.Total"] = df_completo["Fat.Total"].fillna(0)
    
    # Pivot com datas
    df_pivot = df_completo.pivot_table(
        index=["Grupo", "Loja"], columns="Data", values="Fat.Total", aggfunc="sum", fill_value=0
    ).reset_index()
    df_pivot.columns = [
        col if isinstance(col, str) else f"Fat Total {col.strftime('%d/%m/%Y')}"
        for col in df_pivot.columns
    ]
    
    # Acumulado do mês
    df_mes = df_vendas[(df_vendas["Data"] >= primeiro_dia_mes) & (df_vendas["Data"] <= data_fim_dt)]
    df_acumulado = df_mes.groupby(["Loja", "Grupo"], as_index=False)["Fat.Total"].sum()
    df_acumulado = df_lojas_grupos.merge(df_acumulado, on=["Loja", "Grupo"], how="left")
    df_acumulado["Fat.Total"] = df_acumulado["Fat.Total"].fillna(0)
    col_acumulado = f"Acumulado Mês (01/{data_fim_dt.strftime('%m')} até {data_fim_dt.strftime('%d/%m')})"
    df_acumulado = df_acumulado.rename(columns={"Fat.Total": col_acumulado})
    df_base = df_pivot.merge(df_acumulado, on=["Grupo", "Loja"], how="left")
    df_base = df_base[df_base[col_acumulado] != 0]
    
    # Adiciona coluna de Meta
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
    df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
    mapa_meses = {
        "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04", "MAI": "05", "JUN": "06",
        "JUL": "07", "AGO": "08", "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12"
    }
    df_metas["Mês"] = df_metas["Mês"].astype(str).str.strip().str.upper().map(mapa_meses)
    df_metas["Ano"] = df_metas["Ano"].astype(str).str.strip()
    df_metas["Meta"] = (
        df_metas["Meta"]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df_metas["Meta"] = pd.to_numeric(df_metas["Meta"], errors="coerce").fillna(0)
    mes_filtro = data_fim_dt.strftime("%m")
    ano_filtro = data_fim_dt.strftime("%Y")
    df_metas_filtrado = df_metas[(df_metas["Mês"] == mes_filtro) & (df_metas["Ano"] == ano_filtro)].copy()
    df_base["Loja"] = df_base["Loja"].astype(str).str.strip().str.upper()
    # Adiciona coluna Meta
    df_base = df_base.merge(df_metas_filtrado[["Loja", "Meta"]], on="Loja", how="left")
    
    # Adiciona coluna Tipo (vindo de Tabela Empresa)
    # Merge da coluna Tipo
    df_base = df_base.merge(
        df_empresa[["Loja", "Tipo"]].drop_duplicates(), 
        on="Loja", 
        how="left", 
        validate="many_to_one"
    )
    
    
    df_base["Meta"] = df_base["Meta"].fillna(0)
    
    
    # %Atingido
    df_base["%Atingido"] = df_base[col_acumulado] / df_base["Meta"]
    df_base["%Atingido"] = df_base["%Atingido"].replace([np.inf, -np.inf], np.nan).fillna(0).round(4)
    
    # Reordena colunas
    colunas_base = ["Grupo", "Loja", "Tipo"]
    from datetime import datetime
    
    col_diarias = [
        col for col in df_base.columns if col.startswith("Fat Total")
    ]
    
    # Extrai a data do nome da coluna e ordena corretamente
    col_diarias.sort(key=lambda x: datetime.strptime(x.replace("Fat Total ", ""), "%d/%m/%Y"))
    colunas_finais = colunas_base + col_diarias + [col_acumulado, "Meta", "%Atingido", "%LojaXGrupo", "%Grupo"]
    for col in ["%LojaXGrupo", "%Grupo"]:
        if col not in df_base.columns:
            df_base[col] = np.nan
    # Garante que Tipo está presente antes de selecionar colunas finais
    if "Tipo" not in df_base.columns:
        df_base = df_base.merge(df_empresa[["Loja", "Tipo"]].drop_duplicates(), on="Loja", how="left")
    
    # 🔧 Define colunas visíveis antes de qualquer concatenação
    df_base = df_base[colunas_finais]
    colunas_visiveis = colunas_finais.copy()
    
    # 🔢 Linha total
    linha_total = df_base.drop(columns=["Grupo", "Loja", "Tipo"]).sum(numeric_only=True)
    linha_total["Grupo"] = "TOTAL"
    linha_total["Loja"] = f"Lojas: {df_base['Loja'].nunique():02d}"
    linha_total["Tipo"] = ""
    
    # 🧱 Agrupa por grupo
    # 🧱 Agrupa por grupo e define o Tipo do grupo
    ordem_tipos = ["Airports", "Airports - Kopp", "On-Premise"]
    ordem_tipo_dict = {tipo: i for i, tipo in enumerate(ordem_tipos)}
    
    grupos_info = []
    for grupo, df_grp in df_base.groupby("Grupo"):
        df_grp = df_grp.copy()
        
        # Detecta o tipo mais comum do grupo (ou NA se indefinido)
        tipo_dominante = df_grp["Tipo"].dropna().mode().iloc[0] if not df_grp["Tipo"].dropna().empty else "—"
        tipo_ordenado = ordem_tipo_dict.get(tipo_dominante, 999)
        
        total_grupo = df_grp[col_acumulado].sum()
        grupos_info.append((tipo_ordenado, grupo, total_grupo, df_grp, tipo_dominante))
    
    # 📊 Ordena primeiro por Tipo, depois por acumulado (decrescente)
    grupos_info.sort(key=lambda x: (x[0], -x[2]))
    
    
    # 🔁 Monta blocos
    blocos = []
    for _, grupo, _, df_grp, tipo_dominante in grupos_info:
        df_grp_ord = df_grp.sort_values(by=col_acumulado, ascending=False)
    
        # Subtotal
        tipo_valor = tipo_dominante
    
        subtotal = df_grp_ord.drop(columns=["Grupo", "Loja", "Tipo"]).sum(numeric_only=True)
        subtotal["Grupo"] = f"{'SUBTOTAL ' if modo_exibicao == 'Loja' else ''}{grupo}"
        subtotal["Loja"] = f"Lojas: {df_grp_ord['Loja'].nunique():02d}"
        subtotal["Tipo"] = tipo_valor
    
        # ✅ Garante todas as colunas
        for col in colunas_visiveis:
            if col not in subtotal:
                subtotal[col] = np.nan
        subtotal = subtotal[colunas_visiveis]
    
        # 🟦 Lojas
        if modo_exibicao == "Loja":
            blocos.append(df_grp_ord[colunas_visiveis])
    
        # 🟨 Subtotal
        blocos.append(pd.DataFrame([subtotal], columns=colunas_visiveis))
    
    # 🔚 Junta tudo
    linha_total = pd.DataFrame([linha_total], columns=colunas_visiveis)
    df_final = pd.concat([linha_total] + blocos, ignore_index=True)
    
    #st.write("🔍 Diagnóstico: Linhas de loja sem Tipo", df_final[(df_final["Tipo"].isna()) & (~df_final["Loja"].str.startswith("Lojas:"))])
    # Percentuais
    filtro_lojas = (
        (df_final["Loja"] != "") &
        (~df_final["Grupo"].astype(str).str.startswith("SUBTOTAL")) &
        (df_final["Grupo"] != "TOTAL")
    )
    df_lojas_reais = df_final[filtro_lojas].copy()
    soma_por_grupo = df_lojas_reais.groupby("Grupo")[col_acumulado].transform("sum")
    soma_total_geral = df_lojas_reais[col_acumulado].sum()
    if modo_exibicao != "Grupo":
        df_final.loc[filtro_lojas, "%LojaXGrupo"] = (
            df_lojas_reais[col_acumulado].values / soma_por_grupo.values
        ).round(4)
    if modo_exibicao == "Grupo":
        filtro_grupos = df_final["Loja"].astype(str).str.startswith("Lojas:")
    else:
        filtro_grupos = df_final["Grupo"].astype(str).str.startswith("SUBTOTAL")
    df_final.loc[filtro_grupos, "%Grupo"] = (
        df_final.loc[filtro_grupos, col_acumulado] / soma_total_geral
    ).round(4)
    
    # %Atingido final
    if modo_exibicao == "Loja":
        mascara_subtotal = df_final["Grupo"].astype(str).str.startswith("SUBTOTAL")
        mascara_total = df_final["Grupo"] == "TOTAL"
        df_final.loc[mascara_subtotal | mascara_total, "%Atingido"] = (
            df_final.loc[mascara_subtotal | mascara_total, col_acumulado] /
            df_final.loc[mascara_subtotal | mascara_total, "Meta"]
        ).replace([np.inf, -np.inf], np.nan).fillna(0).round(4)
    else:
        mascara_total = df_final["Grupo"] == "TOTAL"
        mascara_subtotal = df_final["Loja"].astype(str).str.startswith("Lojas:")
        df_final.loc[mascara_total | mascara_subtotal, "%Atingido"] = (
            df_final.loc[mascara_total | mascara_subtotal, col_acumulado] /
            df_final.loc[mascara_total | mascara_subtotal, "Meta"]
        ).replace([np.inf, -np.inf], np.nan).fillna(0).round(4)
        df_final.loc[~(mascara_total | mascara_subtotal), "%Atingido"] = ""
    
    # Oculta coluna %LojaXGrupo se for modo Grupo
    # Define colunas com base no filtro "Meta" ou "Sem Meta"
    colunas_visiveis = ["Grupo", "Loja", "Tipo"] + col_diarias + [col_acumulado]
    
    if filtro_meta == "Meta":
        colunas_visiveis += ["Meta", "%Atingido"]
    elif filtro_meta == "Sem Meta":
        if modo_exibicao == "Loja":
            colunas_visiveis += ["%LojaXGrupo", "%Grupo"]
        else:  # Grupo
            colunas_visiveis += ["%Grupo"]
    if "Tipo" in colunas_visiveis:
        colunas_visiveis.remove("Tipo")
    df_final = df_final[colunas_visiveis]
    
    # Formata valores
    colunas_percentuais = ["%LojaXGrupo", "%Grupo", "%Atingido"]
    def formatar(valor, col):
        try:
            if pd.isna(valor) or valor == "":
                return ""
            return f"{valor:.2%}".replace(".", ",") if col in colunas_percentuais else f"R$ {valor:,.2f}".replace(".", "X").replace(",", ".").replace("X", ",")
        except:
            return ""
    # Soma total da meta do mês filtrado
    total_meta_mes = df_metas_filtrado["Meta"].sum()
    
    df_formatado = df_final.copy()
    for col in colunas_visiveis:
        if col in colunas_percentuais:
            df_formatado[col] = df_formatado[col].apply(lambda x: formatar(x, col))
        elif col not in ["Grupo", "Loja", "Tipo"]:
            df_formatado[col] = df_formatado[col].apply(lambda x: formatar(x, col))
        else:
            df_formatado[col] = df_formatado[col].fillna("")  # 👈 tipo e loja não numéricos
    
    # ================================
    # ➕ Linhas de resumo por Tipo
    # ================================
    df_base_tipo = df_base.copy()
    
    # Ignora lojas sem tipo
    df_base_tipo = df_base_tipo[~df_base_tipo["Tipo"].isna()]
    
    linhas_resumo_tipo = []
    tipos_ordenados = df_base_tipo.groupby("Tipo")[col_acumulado].sum().sort_values(ascending=False).index.tolist()
    
    for tipo in tipos_ordenados:
        df_tipo_filtro = df_base_tipo[df_base_tipo["Tipo"] == tipo]
        if df_tipo_filtro.empty:
            continue
    
        linha = {}
        linha["Grupo"] = tipo
        linha["Loja"] = f"Lojas: {df_tipo_filtro['Loja'].nunique():02d}"
        linha["Tipo"] = tipo if pd.notna(tipo) and tipo != "" else "—"  # 👈 aqui
    
        # Somatórios
        for col in col_diarias:
            linha[col] = df_tipo_filtro[col].sum()
    
        linha[col_acumulado] = df_tipo_filtro[col_acumulado].sum()
    
        if filtro_meta == "Meta":
            meta_total = df_tipo_filtro["Meta"].sum()
            linha["Meta"] = meta_total
            linha["%Atingido"] = linha[col_acumulado] / meta_total if meta_total > 0 else 0
    
        elif filtro_meta == "Sem Meta":
            if modo_exibicao == "Loja":
                soma_grupo = df_lojas_reais[col_acumulado].sum()
                linha["%LojaXGrupo"] = linha[col_acumulado] / soma_grupo if soma_grupo > 0 else 0
            linha["%Grupo"] = linha[col_acumulado] / soma_total_geral if soma_total_geral > 0 else 0
    
        linhas_resumo_tipo.append(linha)
    
    
    df_resumo_tipo = pd.DataFrame(linhas_resumo_tipo)
    
    # Formata
    df_resumo_tipo_formatado = df_resumo_tipo.copy()
    for col in df_resumo_tipo.columns:
        if col not in ["Grupo", "Loja"]:
            df_resumo_tipo_formatado[col] = df_resumo_tipo[col].apply(lambda x: formatar(x, col))
    
    
    
    
    
    
    
    
    
    
    
    # Calcula o percentual desejável até o dia selecionado
    dia_hoje = data_fim_dt.day
    dias_mes = monthrange(data_fim_dt.year, data_fim_dt.month)[1]
    perc_desejavel = dia_hoje / dias_mes
    
    # Cálculo proporcional da meta até o dia
    meta_total_mes = df_metas_filtrado["Meta"].sum()
    meta_desejada_ate_hoje = (meta_total_mes / dias_mes) * dia_hoje
    
    # Faturamento desejável (com ordem correta das colunas)
    # Faturamento desejável (com ordem correta das colunas)
    linha_desejavel_dict = {}
    for col in colunas_visiveis:
        if col == "Grupo":
            linha_desejavel_dict[col] = ""
        elif col == "Loja":
            linha_desejavel_dict[col] = f"FATURAMENTO DESEJÁVEL ATÉ {data_fim_dt.strftime('%d/%m')}"
        elif col == "%Atingido":
            linha_desejavel_dict[col] = formatar(perc_desejavel, "%Atingido")
        elif col == col_acumulado:
            linha_desejavel_dict[col] = formatar(meta_desejada_ate_hoje, "Meta")
        else:
            linha_desejavel_dict[col] = ""
    
    linha_desejavel = pd.DataFrame([linha_desejavel_dict])
    
    # 🔽 Remove "Tipo" da visualização final
    for df_temp in [df_resumo_tipo_formatado, df_formatado, linha_desejavel]:
        if "Tipo" in df_temp.columns:
            df_temp.drop(columns=["Tipo"], inplace=True)
    
    # 🔁 Junta tudo para exibir
    df_linhas_visiveis = pd.concat([df_resumo_tipo_formatado, df_formatado], ignore_index=True)
    df_exibir = pd.concat([linha_desejavel, df_linhas_visiveis], ignore_index=True)
    
    # 🎨 Define função de estilo
    def aplicar_estilo_final(df, estilos_linha):
        def apply_row_style(row):
            base_style = estilos_linha[row.name].copy()
            if "%Atingido" in df.columns and row.name > 0:
                try:
                    valor = row["%Atingido"]
                    if isinstance(valor, str) and "%" in valor:
                        valor_float = float(valor.replace("%", "").replace(",", ".")) / 100
                    else:
                        valor_float = float(valor)
                    if not pd.isna(valor_float):
                        idx = df.columns.get_loc("%Atingido")
                        if valor_float >= perc_desejavel:
                            base_style[idx] += "; color: green; font-weight: bold"
                        else:
                            base_style[idx] += "; color: red; font-weight: bold"
                except:
                    pass
            return base_style
        return df.style.apply(apply_row_style, axis=1)
    
    # 🎨 Estilo visual por linha
    cores_alternadas = ["#eef4fa", "#f5fbf3"]
    estilos_linha = []
    cor_idx = -1
    grupo_atual = None
    
    tem_grupo_resumo = (
        'df_resumo_tipo_formatado' in locals()
        and not df_resumo_tipo_formatado.empty
        and "Grupo" in df_resumo_tipo_formatado.columns
    )
    
    for _, row in df_linhas_visiveis.iterrows():
        grupo = row["Grupo"]
        loja = row["Loja"]
    
        if isinstance(grupo, str) and tem_grupo_resumo and grupo in df_resumo_tipo_formatado["Grupo"].values:
            estilos_linha.append(["background-color: #fffbea; font-weight: bold"] * len(row))
        elif grupo == "TOTAL":
            estilos_linha.append(["background-color: #f2f2f2; font-weight: bold"] * len(row))
        elif isinstance(grupo, str) and grupo.startswith("SUBTOTAL"):
            estilos_linha.append(["background-color: #fff8dc; font-weight: bold"] * len(row))
        elif loja == "":
            estilos_linha.append(["background-color: #fdfdfd"] * len(row))
        else:
            if grupo != grupo_atual:
                cor_idx = (cor_idx + 1) % len(cores_alternadas)
                grupo_atual = grupo
            cor = cores_alternadas[cor_idx]
            estilos_linha.append([f"background-color: {cor}"] * len(row))
    
    # ➕ Linha desejável no topo
    estilos_final = [["background-color: #dddddd; font-weight: bold"] * len(df_linhas_visiveis.columns)]
    estilos_final += estilos_linha
    
    # 📊 Exibe resultado final
    st.dataframe(
        aplicar_estilo_final(df_exibir, estilos_final),
        use_container_width=True,
        height=750
    )
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from io import BytesIO
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Border, Side
    
    # Gera o Excel já na memória
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendas"
    
    # Define bordas
    border_padrao = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    
    border_grossa = Border(
        left=Side(style="medium"),
        right=Side(style="medium"),
        top=Side(style="medium"),
        bottom=Side(style="medium"),
    )
    
    # Cabeçalho com azul escuro
    # Cabeçalho estilizado
    for col_idx, col in enumerate(df_exibir.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col)
        cell.fill = PatternFill("solid", fgColor="1F4E78")  # azul escuro
        cell.font = Font(bold=True, color="FFFFFF")         # texto branco
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="medium"),
            right=Side(style="medium"),
            top=Side(style="medium"),
            bottom=Side(style="medium")
        )
    
    ws.row_dimensions[1].height = 30
    
    # Preenche os dados na planilha
    # Dentro do loop de preenchimento de dados
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    
    # Estilos de borda
    thin = Side(border_style="thin", color="000000")
    thick = Side(border_style="medium", color="000000")
    border_padrao = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_grossa = Border(left=thick, right=thick, top=thick, bottom=thick)
    
    # Dados
    for row_idx, (i, row) in enumerate(df_exibir.iterrows(), start=2):
        estilo_linha = estilos_final[row_idx - 2]  # -2 porque o cabeçalho está na linha 1
    
        # ✅ Detecta se a linha é SUBTOTAL ou TOTAL
        grupo = row.get("Grupo", "")
        is_subtotal = isinstance(grupo, str) and grupo.startswith("SUBTOTAL")
        is_total = grupo == "TOTAL"
        usar_borda_grossa = is_subtotal or is_total
    
        for col_idx, (col, valor) in enumerate(row.items(), start=1):
            # Aplica valor e formatação numérica
            if isinstance(valor, str) and "%" in valor:
                try:
                    valor_float = float(valor.replace("%", "").replace(",", ".")) / 100
                    cell = ws.cell(row=row_idx, column=col_idx, value=valor_float)
                    cell.number_format = '0.00%'
                except:
                    cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            elif isinstance(valor, str) and "R$" in valor:
                try:
                    valor_float = float(valor.replace("R$", "").replace(".", "").replace(",", "."))
                    cell = ws.cell(row=row_idx, column=col_idx, value=valor_float)
                    cell.number_format = 'R$ #,##0.00'
                except:
                    cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            else:
                cell = ws.cell(row=row_idx, column=col_idx, value=valor)
    
            # Estilo de fundo
            estilo = estilo_linha[col_idx - 1]
            if "background-color" in estilo:
                cor = estilo.split("background-color: ")[1].split(";")[0].replace("#", "")
                cell.fill = PatternFill("solid", fgColor=cor)
    
            # Estilo de negrito
            if "font-weight: bold" in estilo:
                cell.font = Font(bold=True)
    
            # Alinhamento
            cell.alignment = Alignment(horizontal="left" if col in ["Grupo", "Loja"] else "right")
    
            # ✅ Borda (grossa se subtotal ou total)
            cell.border = border_grossa if usar_borda_grossa else border_padrao
    
            # ✅ Cor verde/vermelha no %Atingido
            if col == "%Atingido":
                try:
                    if isinstance(valor, str) and "%" in valor:
                        valor_float = float(valor.replace("%", "").replace(",", ".")) / 100
                    elif isinstance(valor, (int, float)):
                        valor_float = float(valor)
                    else:
                        valor_float = None
    
                    if valor_float is not None:
                        if valor_float >= perc_desejavel:
                            cell.font = Font(color="006400", bold=True)  # Verde escuro
                        else:
                            cell.font = Font(color="B22222", bold=True)  # Vermelho escuro
                except:
                    pass
    
    
    # ⬇️ Ajusta automaticamente a largura das colunas
    # Ajuste refinado de largura das colunas
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length = 0
        for cell in column_cells:
            try:
                if cell.number_format == 'R$ #,##0.00' and isinstance(cell.value, (float, int)):
                    cell_str = f'R$ {cell.value:,.2f}'.replace(",", "X").replace(".", ",").replace("X", ".")
                    length = len(cell_str)
                elif cell.number_format == '0.00%' and isinstance(cell.value, (float, int)):
                    cell_str = f'{cell.value:.2%}'.replace(".", ",")
                    length = len(cell_str)
                else:
                    cell_str = str(cell.value) if cell.value is not None else ""
                    length = len(cell_str)
                max_length = max(max_length, length)
            except:
                pass
    
        adjusted_width = max_length + 2  # margem extra
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = adjusted_width
    
    # Salva em memória
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    # Botão direto de download
    st.download_button(
        label="📥 Baixar Excel",
        data=buffer,
        file_name="vendas_formatado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ======================
# 📝 Relatórios Financeiros
# ======================

with aba5:
    try:
        import pandas as pd
        pd.set_option('display.max_colwidth', 20)
        pd.set_option('display.width', 1000)

        aba_relatorio = planilha.worksheet("Faturamento Meio Pagamento")
        df_relatorio = pd.DataFrame(aba_relatorio.get_all_records())
        df_relatorio.columns = df_relatorio.columns.str.strip()

        aba_meio_pagamento = planilha.worksheet("Tabela Meio Pagamento")
        df_meio_pagamento = pd.DataFrame(aba_meio_pagamento.get_all_records())
        df_meio_pagamento.columns = df_meio_pagamento.columns.str.strip()

        # Corrige valores
        df_relatorio["Valor (R$)"] = (
            df_relatorio["Valor (R$)"]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace("(", "-")
            .str.replace(")", "")
            .str.replace(" ", "")
            .str.replace(".", "")
            .str.replace(",", ".")
            .astype(float)
        )

        df_relatorio["Data"] = pd.to_datetime(df_relatorio["Data"], dayfirst=True, errors="coerce")
        df_relatorio = df_relatorio[df_relatorio["Data"].notna()]

        from unidecode import unidecode
        for col in ["Loja", "Grupo", "Meio de Pagamento"]:
            df_relatorio[col] = df_relatorio[col].astype(str).str.strip().str.upper().map(unidecode)
            if col in df_meio_pagamento.columns:
                df_meio_pagamento[col] = df_meio_pagamento[col].astype(str).str.strip().str.upper().map(unidecode)

        min_data = df_relatorio["Data"].min().date()
        max_data = df_relatorio["Data"].max().date()
        
        
        
        
        
        col1, col2, col3 = st.columns(3)

        with col1:
            data_inicio, data_fim = st.date_input(
                "Período:",
                value=(max_data, max_data),
                min_value=min_data,
                max_value=max_data
            )
        
        with col2:
            modo_relatorio = st.selectbox(
                "Tipo de análise:",
                ["Vendas", "Financeiro", "Vendas + Prazo e Taxas"]
            )
        
        with col3:
            if modo_relatorio == "Vendas":
                tipo_relatorio = st.selectbox(
                    "Relatório:",
                    ["Meio de Pagamento", "Loja", "Grupo"]
                )
            else:
                tipo_relatorio = None
        if data_inicio > data_fim:
            st.warning("🚫 A data inicial não pode ser maior que a data final.")
        else:
            df_filtrado = df_relatorio[
                (df_relatorio["Data"].dt.date >= data_inicio) &
                (df_relatorio["Data"].dt.date <= data_fim)
            ]

            if df_filtrado.empty:
                st.info("🔍 Não há dados para o período selecionado.")
            else:
                if modo_relatorio == "Vendas":
                    
                    if tipo_relatorio == "Meio de Pagamento":
                        index_cols = ["Meio de Pagamento"]
                    elif tipo_relatorio == "Loja":
                        index_cols = ["Loja", "Grupo", "Meio de Pagamento"]
                    elif tipo_relatorio == "Grupo":
                        index_cols = ["Grupo", "Meio de Pagamento"]

                    df_pivot = pd.pivot_table(
                        df_filtrado,
                        index=index_cols,
                        columns=df_filtrado["Data"].dt.strftime("%d/%m/%Y"),
                        values="Valor (R$)",
                        aggfunc="sum",
                        fill_value=0
                    ).reset_index()

                    novo_nome_datas = {col: f"Vendas - {col}" for col in df_pivot.columns if "/" in str(col)}
                    df_pivot.rename(columns=novo_nome_datas, inplace=True)

                    df_pivot["Total Vendas"] = df_pivot[[c for c in df_pivot.columns if "Vendas -" in str(c)]].sum(axis=1)

                    linha_total_dict = {df_pivot.columns[0]: "TOTAL GERAL"}
                    for col in df_pivot.columns[1:]:
                        if "Vendas -" in str(col) or col == "Total Vendas":
                            linha_total_dict[col] = df_pivot[col].sum()
                        else:
                            linha_total_dict[col] = np.nan
                    linha_total = pd.DataFrame([linha_total_dict])

                    df_pivot_total = pd.concat([linha_total, df_pivot], ignore_index=True)

                    df_pivot_exibe = df_pivot_total.copy()
                    for col in df_pivot_exibe.select_dtypes(include=[np.number]).columns:
                        df_pivot_exibe[col] = df_pivot_exibe[col].map(
                            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            if pd.notna(x) else ""
                        )

                    st.dataframe(df_pivot_exibe, use_container_width=True)

                elif modo_relatorio == "Financeiro":
                    df_completo = df_filtrado.merge(
                        df_meio_pagamento[["Meio de Pagamento", "Prazo", "Antecipa S/N"]],
                        on="Meio de Pagamento",
                        how="left"
                    )
                    df_completo["Prazo"] = pd.to_numeric(df_completo["Prazo"], errors="coerce").fillna(0).astype(int)
                    df_completo["Antecipa S/N"] = df_completo["Antecipa S/N"].astype(str).str.strip().str.upper()

                    from pandas.tseries.offsets import BDay
                    df_completo["Data Recebimento"] = df_completo.apply(
                        lambda row: row["Data"] + BDay(1) if row["Antecipa S/N"] == "SIM" else row["Data"] + BDay(row["Prazo"]),
                        axis=1
                    )

                    df_financeiro = df_completo.groupby(df_completo["Data Recebimento"].dt.date)["Valor (R$)"].sum().reset_index()
                    df_financeiro = df_financeiro.rename(columns={"Data Recebimento": "Data"}).sort_values("Data")

                    total_geral = df_financeiro["Valor (R$)"].sum()
                    linha_total = pd.DataFrame([["TOTAL GERAL", total_geral]], columns=df_financeiro.columns)
                    df_financeiro_total = pd.concat([linha_total, df_financeiro], ignore_index=True)

                    df_financeiro_total["Valor (R$)"] = df_financeiro_total["Valor (R$)"].map(
                        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        if pd.notna(x) else ""
                    )

                    st.dataframe(df_financeiro_total, use_container_width=True)

                elif modo_relatorio == "Vendas + Prazo e Taxas":
                    df_completo = df_filtrado.merge(
                        df_meio_pagamento[["Meio de Pagamento", "Prazo", "Antecipa S/N", "Taxa Bandeira", "Taxa Antecipação"]],
                        on="Meio de Pagamento",
                        how="left"
                    )

                    df_pivot = pd.pivot_table(
                        df_completo,
                        index=["Meio de Pagamento", "Prazo", "Antecipa S/N", "Taxa Bandeira", "Taxa Antecipação"],
                        columns=df_completo["Data"].dt.strftime("%d/%m/%Y"),
                        values="Valor (R$)",
                        aggfunc="sum",
                        fill_value=0
                    ).reset_index()

                    colunas_datas = [col for col in df_pivot.columns if "/" in col]
                    novo_nome_datas = {col: f"Vendas - {col}" for col in colunas_datas}
                    df_pivot.rename(columns=novo_nome_datas, inplace=True)
                    df_pivot.rename(columns={"Vendas - Antecipa S/N": "Antecipa S/N"}, inplace=True)

                    colunas_vendas = [col for col in df_pivot.columns if "Vendas" in col]
                    cols_fixas = ["Meio de Pagamento", "Prazo", "Antecipa S/N", "Taxa Bandeira", "Taxa Antecipação"]
                    novas_cols = []

                    for col_vendas in colunas_vendas:
                        data_col = col_vendas.split(" - ")[1]
                        col_taxa_bandeira = f"Vlr Taxa Bandeira - {data_col}"
                        taxa_bandeira = (
                            pd.to_numeric(df_pivot["Taxa Bandeira"].astype(str)
                                          .str.replace("%","")
                                          .str.replace(",","."),
                                          errors="coerce").fillna(0) / 100
                        )
                        df_pivot[col_taxa_bandeira] = df_pivot[col_vendas] * taxa_bandeira

                        col_taxa_antecipacao = f"Vlr Taxa Antecipação - {data_col}"
                        taxa_antecipacao = (
                            pd.to_numeric(df_pivot["Taxa Antecipação"].astype(str)
                                          .str.replace("%","")
                                          .str.replace(",","."),
                                          errors="coerce").fillna(0) / 100
                        )
                        df_pivot[col_taxa_antecipacao] = df_pivot[col_vendas] * taxa_antecipacao

                        novas_cols.extend([col_vendas, col_taxa_bandeira, col_taxa_antecipacao])

                    df_pivot = df_pivot[cols_fixas + novas_cols]

                    df_pivot["Total Vendas"] = df_pivot[colunas_vendas].sum(axis=1)
                    df_pivot["Total Tx Bandeira"] = df_pivot[[col for col in df_pivot.columns if "Vlr Taxa Bandeira" in col]].sum(axis=1)
                    df_pivot["Total Tx Antecipação"] = df_pivot[[col for col in df_pivot.columns if "Vlr Taxa Antecipação" in col]].sum(axis=1)
                    df_pivot["Total a Receber"] = df_pivot["Total Vendas"] - df_pivot["Total Tx Bandeira"] - df_pivot["Total Tx Antecipação"]

                    linha_total_dict = {col: "" for col in df_pivot.columns}
                    linha_total_dict["Meio de Pagamento"] = "TOTAL GERAL"
                    for col in df_pivot.columns:
                        if "Vendas" in col or "Vlr Taxa Bandeira" in col or "Vlr Taxa Antecipação" in col \
                            or "Total Tx" in col or col in ["Total Vendas", "Total a Receber"]:
                            linha_total_dict[col] = df_pivot[col].sum()

                    linha_total = pd.DataFrame([linha_total_dict])
                    df_pivot_total = pd.concat([linha_total, df_pivot], ignore_index=True)

                    df_pivot_exibe = df_pivot_total.copy()
                    for col in [c for c in df_pivot_exibe.columns if "Vendas" in c or "Vlr Taxa Bandeira" in c 
                                or "Vlr Taxa Antecipação" in c or "Total Tx" in c or c in ["Total Vendas", "Total a Receber"]]:
                        df_pivot_exibe[col] = df_pivot_exibe[col].map(
                            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        )

                    st.dataframe(df_pivot_exibe, use_container_width=True)

                    from openpyxl import load_workbook

                    output = BytesIO()
                    df_exportar = df_pivot_total.copy()
                    df_exportar["Taxa Bandeira"] = (
                        pd.to_numeric(df_exportar["Taxa Bandeira"].astype(str)
                                      .str.replace("%", "")
                                      .str.replace(",", "."),
                                      errors="coerce") / 100
                    )
                    df_exportar["Taxa Antecipação"] = (
                        pd.to_numeric(df_exportar["Taxa Antecipação"].astype(str)
                                      .str.replace("%", "")
                                      .str.replace(",", "."),
                                      errors="coerce") / 100
                    )

                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_exportar.to_excel(writer, index=False, sheet_name="PrazoTaxas")
                    output.seek(0)

                    wb = load_workbook(output)
                    ws = wb["PrazoTaxas"]
                    header = [cell.value for cell in ws[1]]

                    for col_name in ["Taxa Bandeira", "Taxa Antecipação"]:
                        if col_name in header:
                            col_idx = header.index(col_name) + 1
                            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                                for cell in row:
                                    cell.number_format = "0.00%"

                    for col_name in header:
                        if ("Vendas" in col_name or "Vlr Taxa Bandeira" in col_name 
                            or "Vlr Taxa Antecipação" in col_name or "Total Tx" in col_name 
                            or col_name in ["Total Vendas", "Total a Receber"]):
                            col_idx = header.index(col_name) + 1
                            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                                for cell in row:
                                    cell.number_format = '"R$" #,##0.00'

                    output_final = BytesIO()
                    wb.save(output_final)
                    output_final.seek(0)

                    st.download_button(
                        "📥 Baixar Excel",
                        data=output_final,
                        file_name=f"Vendas_Prazo_Taxas_{data_inicio.strftime('%d-%m-%Y')}_a_{data_fim.strftime('%d-%m-%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ Erro ao acessar Google Sheets: {e}")
