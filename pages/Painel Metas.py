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


# Função auxiliar para tratar datas misturadas (string e número serial):
def tratar_data(val):
    try:
        val_str = str(val).strip()
        if val_str.replace('.', '', 1).isdigit():
            return pd.Timestamp('1899-12-30') + pd.to_timedelta(float(val), unit='D')
        else:
            return pd.to_datetime(val_str, dayfirst=True, errors='coerce')
    except:
        return pd.NaT


# Função auxiliar de blindagem
def garantir_escalar(x):
    if isinstance(x, list):
        if len(x) == 1:
            return x[0]
        return str(x)
    return x

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
   # Já com data tratada e blindada:
    df_anos["Data"] = df_anos["Data"].apply(tratar_data)
    df_anos = df_anos.dropna(subset=["Data"])
    
    # Agora o mapeamento correto do mês em português:
    meses_map = {
        1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr',
        5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago',
        9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
    }
    df_anos["Mês"] = df_anos["Data"].dt.month.map(meses_map)
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Fat.Total"] = df_anos["Fat.Total"].apply(parse_valor)

    # Adiciona Tipo no df_anos via merge
    df_anos = df_anos.merge(df_empresa[["Loja", "Grupo", "Tipo"]], on=["Loja", "Grupo"], how="left")

    mes_atual = datetime.now().strftime("%b")
    ano_atual = datetime.now().year
    ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

    anos_disponiveis = sorted(df_anos["Ano"].unique())
    ano_selecionado = st.selectbox("Selecione o Ano:", anos_disponiveis, index=anos_disponiveis.index(ano_atual) if ano_atual in anos_disponiveis else 0)
    mes_selecionado = st.selectbox("Selecione o Mês:", ordem_meses, index=ordem_meses.index(mes_atual) if mes_atual in ordem_meses else 0)

    df_anos_filtrado = df_anos[(df_anos["Ano"] == ano_selecionado) & (df_anos["Mês"] == mes_selecionado)].copy()
    df_metas_filtrado = df_metas[(df_metas["Ano"] == ano_selecionado) & (df_metas["Mês"] == mes_selecionado)].copy()

    for col in ["Ano", "Mês", "Loja", "Grupo"]:
        df_metas_filtrado[col] = df_metas_filtrado[col].apply(garantir_escalar)
        df_anos_filtrado[col] = df_anos_filtrado[col].apply(garantir_escalar)

    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Meta"})
    realizado_grouped = df_anos_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo", "Tipo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja", "Grupo"], how="outer").fillna(0)
    comparativo["% Atingido"] = np.where(comparativo["Meta"] == 0, 0, comparativo["Realizado"] / comparativo["Meta"])
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]
    comparativo["% Falta Atingir"] = np.maximum(0, 1 - comparativo["% Atingido"])
    
    # Cria o peso de ordenação do Tipo
    ordem_tipo = {
        'AIRPORTS': 1,
        'ONPRIMESSE': 2,
        'BARES': 3
    }
    
    # Adiciona coluna auxiliar de ordenação
    comparativo['Ordem_Tipo'] = comparativo['Tipo'].map(ordem_tipo)
    
    # Ordena as lojas
    comparativo = comparativo.sort_values(
        by=['Grupo', 'Ordem_Tipo', 'Realizado'],
        ascending=[True, True, False]
    )
    
    
    
    
    
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
            "Ano": [""], "Mês": [""], "Grupo": [""],
            "Loja": [f"Tipo: {tipo} - Lojas: {qtde_lojas_tipo:02}"],
            "Meta": [soma_meta_tipo], "Realizado": [soma_realizado_tipo],
            "% Atingido": [perc_atingido_tipo], "% Falta Atingir": [perc_falta_tipo],
            "Diferença": [soma_diferenca_tipo]
        })
        tipo_subtotais.append(linha_tipo)

    resultado_final = []
    total_lojas_geral = comparativo["Loja"].nunique()

    for grupo, dados_grupo in comparativo.groupby("Grupo"):
        resultado_final.append(dados_grupo)
        soma_meta_grupo = dados_grupo["Meta"].sum()
        soma_realizado_grupo = dados_grupo["Realizado"].sum()
        soma_diferenca_grupo = dados_grupo["Diferença"].sum()
        perc_atingido_grupo = soma_realizado_grupo / soma_meta_grupo if soma_meta_grupo != 0 else 0
        perc_falta_grupo = max(0, 1 - perc_atingido_grupo)
        qtde_lojas_grupo = dados_grupo["Loja"].nunique()

        linha_subtotal = pd.DataFrame({
            "Ano": [""], "Mês": [""], "Grupo": [grupo],
            "Loja": [f"{grupo} - Lojas: {qtde_lojas_grupo:02}"],
            "Meta": [soma_meta_grupo], "Realizado": [soma_realizado_grupo],
            "% Atingido": [perc_atingido_grupo], "% Falta Atingir": [perc_falta_grupo],
            "Diferença": [soma_diferenca_grupo]
        })
        resultado_final.append(linha_subtotal)

    total_meta = comparativo["Meta"].sum()
    total_realizado = comparativo["Realizado"].sum()
    total_diferenca = comparativo["Diferença"].sum()
    percentual_total = total_realizado / total_meta if total_meta != 0 else 0
    percentual_falta_total = max(0, 1 - percentual_total)

    linha_total = pd.DataFrame({
        "Ano": [""], "Mês": [""], "Grupo": [""],
        "Loja": [f"TOTAL GERAL - Lojas: {total_lojas_geral:02}"],
        "Meta": [total_meta], "Realizado": [total_realizado],
        "% Atingido": [percentual_total], "% Falta Atingir": [percentual_falta_total],
        "Diferença": [total_diferenca]
    })

    comparativo_final = pd.concat(tipo_subtotais + [linha_total] + resultado_final, ignore_index=True)

    # Separa as linhas de lojas normais (que não são subtotal nem total)
    linhas_lojas = comparativo_final[~comparativo_final['Loja'].str.contains('Lojas:|TOTAL GERAL', na=False)].copy()
    
    # Cria o peso de ordenação do Tipo
    ordem_tipo = {
        'AIRPORTS': 1,
        'ONPRIMESSE': 2,
        'BARES': 3
    }
    
    # Adiciona coluna auxiliar de ordenação
    linhas_lojas['Ordem_Tipo'] = linhas_lojas['Tipo'].map(ordem_tipo)
    
    # Ordena por Grupo, depois Tipo, depois Realizado decrescente
    linhas_lojas = linhas_lojas.sort_values(
        by=['Grupo', 'Ordem_Tipo', 'Realizado'],
        ascending=[True, True, False]
    )
    
    # Agora junta com os subtotais e total
    outras_linhas = comparativo_final[comparativo_final['Loja'].str.contains('Lojas:|TOTAL GERAL', na=False)]
    comparativo_final_ordenado = pd.concat([outras_linhas, linhas_lojas], ignore_index=True)



    

    def formatar_linha(row):
        if "TOTAL GERAL" in row["Loja"]:
            return ['background-color: #0366d6; color: white'] * len(row)
        elif "Tipo:" in row["Loja"]:
            return ['background-color: #FFE699'] * len(row)
        elif "Lojas:" in row["Loja"]:
            return ['background-color: #d0e6f7'] * len(row)
        else:
            return [''] * len(row)

    st.dataframe(
        comparativo_final_ordenado.style
            .format({
                "Meta": "R$ {:,.2f}",
                "Realizado": "R$ {:,.2f}",
                "Diferença": "R$ {:,.2f}",
                "% Atingido": "{:.2%}",
                "% Falta Atingir": "{:.2%}"
            })
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
