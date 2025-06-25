# pages/Painel Metas.py

import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
import io

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

# Padronizar Tabela Empresa
for col in ["Loja", "Grupo", "Tipo"]:
    df_empresa[col] = df_empresa[col].astype(str).str.strip().str.upper()

# ================================
# 2. Estilo e layout
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
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Metas Mensais</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# Função auxiliar para converter valores
# ================================
def parse_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
    except:
        return 0.0

# Função auxiliar para tratar datas misturadas

def tratar_data(val):
    try:
        val_str = str(val).strip()
        if val_str.replace('.', '', 1).isdigit():
            return pd.Timestamp('1899-12-30') + pd.to_timedelta(float(val), unit='D')
        else:
            return pd.to_datetime(val_str, dayfirst=True, errors='coerce')
    except:
        return pd.NaT

def garantir_escalar(x):
    if isinstance(x, list):
        if len(x) == 1:
            return x[0]
        return str(x)
    return x

def formatar_moeda_br(val):
    try:
        if val in [None, np.nan, "None", "nan", "NaN", ""]:
            return ""
        val_float = float(val)
        return f"R$ {val_float:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    except:
        return ""



# ================================
# Abas
# ================================
aba1, aba2 = st.tabs(["📈 Analise Metas", "📊 Auditoria Metas"])

with aba1:

    # --- Metas ---
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
    df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
    df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
    df_metas["Grupo"] = df_metas["Grupo"].astype(str).str.strip().str.upper()
    df_metas = df_metas[df_metas["Loja"] != ""]

    # --- Realizado ---
    df_anos = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
    df_anos.columns = df_anos.columns.str.strip()
    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.upper()
    df_anos["Grupo"] = df_anos["Grupo"].astype(str).str.strip().str.upper()
    df_anos["Data"] = df_anos["Data"].apply(tratar_data)
    df_anos = df_anos.dropna(subset=["Data"])

  

    meses_map = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
    df_anos["Mês"] = df_anos["Data"].dt.month.map(meses_map)
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Fat.Total"] = df_anos["Fat.Total"].apply(parse_valor)

    df_anos = df_anos.merge(df_empresa[["Loja", "Grupo", "Tipo"]], on=["Loja", "Grupo"], how="left")

    mes_atual = datetime.now().strftime("%b")
    ano_atual = datetime.now().year
    ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

    anos_disponiveis = sorted(df_anos["Ano"].unique())
    ano_selecionado = st.selectbox("Selecione o Ano:", anos_disponiveis, index=anos_disponiveis.index(ano_atual) if ano_atual in anos_disponiveis else 0)
    mes_selecionado = st.selectbox("Selecione o Mês:", ordem_meses, index=ordem_meses.index(mes_atual) if mes_atual in ordem_meses else 0)

    df_anos_filtrado = df_anos[(df_anos["Ano"] == ano_selecionado) & (df_anos["Mês"] == mes_selecionado)].copy()
    df_metas_filtrado = df_metas[(df_metas["Ano"] == ano_selecionado) & (df_metas["Mês"] == mes_selecionado)].copy()

    import calendar
    
    # Agora sim pega a última data correta já do mês filtrado
    ultima_data_realizado_dt = df_anos_filtrado["Data"].max()
    ultima_data_realizado = ultima_data_realizado_dt.strftime("%d/%m/%Y")
    
    # Calcula o percentual desejável até a última data
    dias_do_mes = calendar.monthrange(ano_selecionado, ordem_meses.index(mes_selecionado) + 1)[1]
    dias_corridos = ultima_data_realizado_dt.day
    percentual_meta_desejavel = dias_corridos / dias_do_mes





    

    for col in ["Ano", "Mês", "Loja", "Grupo"]:
        df_metas_filtrado[col] = df_metas_filtrado[col].apply(garantir_escalar)
        df_anos_filtrado[col] = df_anos_filtrado[col].apply(garantir_escalar)

    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Meta"})
    realizado_grouped = df_anos_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo", "Tipo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja", "Grupo"], how="outer").fillna(0)
    comparativo["% Atingido"] = np.where(comparativo["Meta"] == 0, 0, comparativo["Realizado"] / comparativo["Meta"])
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]
    comparativo["% Falta Atingir"] = np.maximum(0, 1 - comparativo["% Atingido"])
    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)

    tipo_subtotais = []
    for tipo, dados_tipo in comparativo.groupby("Tipo"):
        soma_meta_tipo = dados_tipo["Meta"].sum()
        soma_realizado_tipo = dados_tipo["Realizado"].sum()
        soma_diferenca_tipo = dados_tipo["Diferença"].sum()
        perc_atingido_tipo = soma_realizado_tipo / soma_meta_tipo if soma_meta_tipo != 0 else 0
        perc_falta_tipo = max(0, 1 - perc_atingido_tipo)
        qtde_lojas_tipo = dados_tipo["Loja"].nunique()

        linha_tipo = pd.DataFrame({
            "Ano": [""], "Mês": [""], "Grupo": [""], "Loja": [f"Tipo: {tipo} - Lojas: {qtde_lojas_tipo:02}"],
            "Meta": [soma_meta_tipo], "Realizado": [soma_realizado_tipo], "% Atingido": [perc_atingido_tipo], "% Falta Atingir": [perc_falta_tipo], "Diferença": [soma_diferenca_tipo], "Tipo": [tipo]
        })
        tipo_subtotais.append(linha_tipo)

    # (até aqui seu código é igual, vamos direto no ponto de alteração)
    
    # ✅ Aqui começa o bloco do resultado_final com a ordenação por Tipo
    resultado_final = []
    total_lojas_geral = comparativo["Loja"].nunique()
    
    # Primeiro, criamos uma lista auxiliar com os subtotais incluindo o tipo já capturado
    subtotais_aux = []
    for grupo, dados_grupo in comparativo.groupby("Grupo"):
        soma_meta_grupo = dados_grupo["Meta"].sum()
        soma_realizado_grupo = dados_grupo["Realizado"].sum()
        soma_diferenca_grupo = dados_grupo["Diferença"].sum()
        perc_atingido_grupo = soma_realizado_grupo / soma_meta_grupo if soma_meta_grupo != 0 else 0
        perc_falta_grupo = max(0, 1 - perc_atingido_grupo)
        qtde_lojas_grupo = dados_grupo["Loja"].nunique()
        
        tipo_grupo = df_empresa[df_empresa["Grupo"] == grupo]["Tipo"].dropna().unique()
        tipo_str = tipo_grupo[0] if len(tipo_grupo) == 1 else "OUTROS"
    
        subtotais_aux.append({
            "grupo": grupo,
            "tipo": tipo_str,
            "qtde_lojas": qtde_lojas_grupo,
            "meta": soma_meta_grupo,
            "realizado": soma_realizado_grupo,
            "diferenca": soma_diferenca_grupo,
            "perc_atingido": perc_atingido_grupo,
            "perc_falta": perc_falta_grupo
        })
    
    # Definimos a ordem de prioridade dos tipos
    ordem_tipo = {"AIRPORTS": 1, "AIRPORTS - KOPP": 2, "ON-PREMISE": 3, "OUTROS": 4}


    # Normaliza o tipo (garantir comparação segura)
    for item in subtotais_aux:
        item["tipo"] = str(item["tipo"]).strip().upper()

    
    # Ordenamos os grupos com base no tipo
    subtotais_aux = sorted(subtotais_aux, key=lambda x: (ordem_tipo.get(x["tipo"], 99), x["grupo"]))
    
    # Agora, com os grupos já ordenados, montamos o resultado final
    for subtotal in subtotais_aux:
        grupo = subtotal["grupo"]
        tipo_str = subtotal["tipo"]
        qtde_lojas_grupo = subtotal["qtde_lojas"]
    
        dados_grupo = comparativo[comparativo["Grupo"] == grupo]
        resultado_final.append(dados_grupo)
    
        linha_subtotal = pd.DataFrame({
            "Ano": [""], "Mês": [""], "Grupo": [grupo],
            "Loja": [f"{grupo} - Lojas: {qtde_lojas_grupo:02}"],
            "Meta": [subtotal["meta"]], "Realizado": [subtotal["realizado"]],
            "% Atingido": [subtotal["perc_atingido"]], "% Falta Atingir": [subtotal["perc_falta"]],
            "Diferença": [subtotal["diferenca"]], "Tipo": [tipo_str]
        })
        resultado_final.append(linha_subtotal)
    
    # ✅ Total Geral continua exatamente igual ao seu
    total_meta = comparativo["Meta"].sum()
    total_realizado = comparativo["Realizado"].sum()
    total_diferenca = comparativo["Diferença"].sum()
    percentual_total = total_realizado / total_meta if total_meta != 0 else 0
    percentual_falta_total = max(0, 1 - percentual_total)
    
    linha_total = pd.DataFrame({
        "Ano": [""], "Mês": [""], "Grupo": [""], "Loja": [f"TOTAL GERAL - Lojas: {total_lojas_geral:02}"],
        "Meta": [total_meta], "Realizado": [total_realizado],
        "% Atingido": [percentual_total], "% Falta Atingir": [percentual_falta_total],
        "Diferença": [total_diferenca], "Tipo": [""]
    })

    # Cria a linha única da Meta Desejável
    linha_meta_desejavel = pd.DataFrame({
        "Ano": [""], 
        "Mês": [""], 
        "Grupo": [""],
        "Loja": [f"🎯 Meta Desejável até {ultima_data_realizado}"],
        "Meta": [np.nan],
        "Realizado": [np.nan],
        "% Atingido": [percentual_meta_desejavel],
        "% Falta Atingir": [np.nan],
        "Diferença": [np.nan],
        "Tipo": [""]
    })











    
    # ✅ Monta o comparativo final preservando o seu restante
    comparativo_final = pd.concat([linha_meta_desejavel] + tipo_subtotais + [linha_total] + resultado_final, ignore_index=True)
    # ✅ Ajusta o nome da coluna "Realizado"
    comparativo_final.rename(columns={"Realizado": f"Realizado até {ultima_data_realizado}"}, inplace=True)

    # Ajusta os tipos para garantir float e evitar None no Styler
    for col in ["Meta", f"Realizado até {ultima_data_realizado}", "Diferença", "% Atingido", "% Falta Atingir"]:
        comparativo_final[col] = pd.to_numeric(comparativo_final[col], errors='coerce')
       
    def formatar_linha(row):
        estilo = []
        for idx, val in enumerate(row):
            if "Meta Desejável" in row["Loja"]:
                if val == 'None':
                    estilo.append('background-color: #FF6666; color: #FF6666')
                else:
                    estilo.append('background-color: #FF6666; color: white')
            elif "TOTAL GERAL" in row["Loja"]:
                estilo.append('background-color: #0366d6; color: white')
            elif "Tipo:" in row["Loja"]:
                estilo.append('background-color: #FFE699')
            elif "Lojas:" in row["Loja"]:
                estilo.append('background-color: #d0e6f7')
            else:
                estilo.append('')
        return estilo

       


        
            

    # ✅ Exibe a data de realizado antes da tabela
    st.markdown(f"**Última data realizada:** {ultima_data_realizado}")
    
    
    
    st.dataframe(
        comparativo_final.style
            .format({
                "Meta": formatar_moeda_br, 
                f"Realizado até {ultima_data_realizado}": formatar_moeda_br, 
                "Diferença": formatar_moeda_br, 
                "% Atingido": "{:.2%}", 
                "% Falta Atingir": "{:.2%}"
            }, na_rep="")
            .apply(formatar_linha, axis=1),
        use_container_width=True
    )
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        comparativo_final.to_excel(writer, index=False, sheet_name='Metas')
    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel",
        data=output,
        file_name=f"Relatorio_Metas_{ano_selecionado}_{mes_selecionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ================================
with aba2:
# ================================    
    st.header("Painel de Gráficos Semanais")

    # ================================
    # Prepara dados semanais
    # ================================
    df_anos['Semana'] = df_anos['Data'].dt.isocalendar().week

    # Filtros básicos
    grupos_disponiveis = df_anos['Grupo'].unique().tolist()
    grupos_disponiveis.sort()
    grupo_selecionado = st.selectbox("Selecione o Grupo:", grupos_disponiveis)

    operacoes_disponiveis = df_empresa['Tipo'].unique().tolist()
    operacoes_disponiveis.sort()
    operacao_selecionada = st.selectbox("Selecione a Operação:", operacoes_disponiveis)

    meses_disponiveis = ordem_meses
    mes_selecionado_grafico = st.selectbox("Selecione o Mês:", meses_disponiveis)

    # Filtro aplicado
    df_filtrado = df_anos[
        (df_anos['Grupo'] == grupo_selecionado) & 
        (df_anos['Tipo'] == operacao_selecionada) &
        (df_anos['Mês'] == mes_selecionado_grafico)
    ]

    # Agrupa por semana
    vendas_semana = df_filtrado.groupby('Semana')['Fat.Total'].sum().reset_index()

    # ================================
    # Gráfico de barras (Vendas por Semana)
    # ================================
    import plotly.graph_objects as go

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=vendas_semana['Semana'],
        y=vendas_semana['Fat.Total'],
        marker_color='green',
        name='Faturamento'
    ))

    fig.update_layout(title="Vendas por Semana",
                      xaxis_title="Semana do Ano",
                      yaxis_title="Faturamento",
                      height=400)
    
    st.plotly_chart(fig, use_container_width=True)

    # ================================
    # Ritmo Ideal (Velocímetro)
    # ================================
    ritmo_real = vendas_semana['Fat.Total'].sum()
    meta_total = df_metas[df_metas['Grupo'] == grupo_selecionado]['Fat.Total'].sum()

    percentual = (ritmo_real / meta_total) * 100 if meta_total > 0 else 0

    fig_gauge = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = percentual,
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': "Ritmo Ideal (%)"},
        gauge = {
            'axis': {'range': [0, 100]},
            'bar': {'color': "darkblue"},
            'steps' : [
                {'range': [0, 50], 'color': "red"},
                {'range': [50, 75], 'color': "yellow"},
                {'range': [75, 100], 'color': "green"}
            ]
        }
    ))

    st.plotly_chart(fig_gauge, use_container_width=True)

    # ================================
    # Índice de Crescimento (Barras horizontais)
    # ================================
    # Exemplo básico de crescimento por loja (apenas estrutura inicial)
    crescimento_df = df_filtrado.groupby('Loja')['Fat.Total'].sum().reset_index()
    crescimento_df['Crescimento'] = (crescimento_df['Fat.Total'] / crescimento_df['Fat.Total'].mean()) * 100

    fig_crescimento = go.Figure(go.Bar(
        x=crescimento_df['Crescimento'],
        y=crescimento_df['Loja'],
        orientation='h',
        marker_color='green'
    ))

    fig_crescimento.update_layout(title="Índice de Crescimento",
                                   xaxis_title="% Crescimento",
                                   height=400)

    st.plotly_chart(fig_crescimento, use_container_width=True)
