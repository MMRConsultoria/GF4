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
import calendar

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
# Funções auxiliares
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
    if pd.isna(val):
        return ""
    return f"R$ {val:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

# ================================
# Abas
# ================================
aba1, aba2 = st.tabs(["📈 Analise Metas", "📊 Auditoria Metas"])

with aba1:

    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
    df_metas["Fat.Total"] = df_metas["Fat.Total"].apply(parse_valor)
    df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
    df_metas["Grupo"] = df_metas["Grupo"].astype(str).str.strip().str.upper()
    df_metas = df_metas[df_metas["Loja"] != ""]

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

    # Agora sim pega a última data correta filtrada
    ultima_data_realizado_dt = df_anos_filtrado["Data"].max()
    ultima_data_realizado = ultima_data_realizado_dt.strftime("%d/%m/%Y")

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
        linha_tipo = pd.DataFrame({"Ano": [""], "Mês": [""], "Grupo": [""], "Loja": [f"Tipo: {tipo} - Lojas: {qtde_lojas_tipo:02}"], "Meta": [soma_meta_tipo], "Realizado": [soma_realizado_tipo], "% Atingido": [perc_atingido_tipo], "% Falta Atingir": [perc_falta_tipo], "Diferença": [soma_diferenca_tipo], "Tipo": [tipo]})
        tipo_subtotais.append(linha_tipo)

    linha_meta_desejavel = pd.DataFrame({"Ano": [""], "Mês": [""], "Grupo": [""], "Loja": [f"🎯 Meta Desejável até {ultima_data_realizado}"], "Meta": [""], "Realizado": [""], "% Atingido": [percentual_meta_desejavel], "% Falta Atingir": [""], "Diferença": [""], "Tipo": [""]})

    total_meta = comparativo["Meta"].sum()
    total_realizado = comparativo["Realizado"].sum()
    total_diferenca = comparativo["Diferença"].sum()
    percentual_total = total_realizado / total_meta if total_meta != 0 else 0
    percentual_falta_total = max(0, 1 - percentual_total)

    linha_total = pd.DataFrame({"Ano": [""], "Mês": [""], "Grupo": [""], "Loja": ["TOTAL GERAL"], "Meta": [total_meta], "Realizado": [total_realizado], "% Atingido": [percentual_total], "% Falta Atingir": [percentual_falta_total], "Diferença": [total_diferenca], "Tipo": [""]})

    comparativo_final = pd.concat([linha_meta_desejavel] + tipo_subtotais + [linha_total], ignore_index=True)
    comparativo_final.rename(columns={"Realizado": f"Realizado até {ultima_data_realizado}"}, inplace=True)

    def formatar_linha(row):
        if "Meta Desejável" in row["Loja"]:
            return ['background-color: #FF6666; color: white'] * len(row)
        elif "TOTAL GERAL" in row["Loja"]:
            return ['background-color: #0366d6; color: white'] * len(row)
        elif "Tipo:" in row["Loja"]:
            return ['background-color: #FFE699'] * len(row)
        else:
            return [''] * len(row)

    st.dataframe(
        comparativo_final.style
            .format({
                "Meta": formatar_moeda_br,
                f"Realizado até {ultima_data_realizado}": formatar_moeda_br,
                "Diferença": formatar_moeda_br,
                "% Atingido": "{:.2%}", "% Falta Atingir": "{:.2%}"
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
