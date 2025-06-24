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

# 1. Conexão com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# Padronizar Tabela Empresa
for col in ["Loja", "Grupo", "Tipo"]:
    df_empresa[col] = df_empresa[col].astype(str).str.strip().str.upper()

# Função auxiliar para converter valores
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

# Abas
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

    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Meta"})
    realizado_grouped = df_anos_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo", "Tipo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja", "Grupo"], how="outer").fillna(0)
    comparativo["% Atingido"] = np.where(comparativo["Meta"] == 0, 0, comparativo["Realizado"] / comparativo["Meta"])
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]
    comparativo["% Falta Atingir"] = np.maximum(0, 1 - comparativo["% Atingido"])
    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)

    # Prepara subtotais
    resultado_final = []

    # Monta subtotais de grupos
    subtotais = []
    for grupo, dados_grupo in comparativo.groupby("Grupo"):
        soma_meta = dados_grupo["Meta"].sum()
        soma_realizado = dados_grupo["Realizado"].sum()
        soma_diferenca = dados_grupo["Diferença"].sum()
        perc_atingido = soma_realizado / soma_meta if soma_meta != 0 else 0
        perc_falta = max(0, 1 - perc_atingido)
        qtde_lojas = dados_grupo["Loja"].nunique()

        tipo_subtotal = dados_grupo["Tipo"].dropna().unique()
        if len(tipo_subtotal) == 1:
            tipo_str = tipo_subtotal[0]
        else:
            tipo_str = "OUTROS"

        subtotais.append({
            "Grupo": grupo, "Tipo": tipo_str, "QtdeLojas": qtde_lojas,
            "Meta": soma_meta, "Realizado": soma_realizado, "Diferença": soma_diferenca,
            "% Atingido": perc_atingido, "% Falta Atingir": perc_falta
        })

    # Ordenação desejada
    ordem_tipo = {"AIRPORTS": 1, "ONPRIMESSE": 2, "OUTROS": 3}
    subtotais_ordenado = sorted(subtotais, key=lambda x: (ordem_tipo.get(x["Tipo"], 3), x["Grupo"]))

    # Gera resultado final
    for subtotal in subtotais_ordenado:
        grupo = subtotal["Grupo"]
        tipo_str = subtotal["Tipo"]
        qtde_lojas = subtotal["QtdeLojas"]
        dados_grupo = comparativo[comparativo["Grupo"] == grupo]
        resultado_final.append(dados_grupo)

        linha_subtotal = pd.DataFrame({
            "Ano": [""], "Mês": [""], "Grupo": [grupo],
            "Loja": [f"{grupo} - {tipo_str} - Lojas: {qtde_lojas:02}"],
            "Meta": [subtotal["Meta"]], "Realizado": [subtotal["Realizado"]],
            "% Atingido": [subtotal["% Atingido"]],
            "% Falta Atingir": [subtotal["% Falta Atingir"]],
            "Diferença": [subtotal["Diferença"]],
            "Tipo": [tipo_str]
        })
        resultado_final.append(linha_subtotal)

    # Total geral
    total_meta = comparativo["Meta"].sum()
    total_realizado = comparativo["Realizado"].sum()
    total_diferenca = comparativo["Diferença"].sum()
    percentual_total = total_realizado / total_meta if total_meta != 0 else 0
    percentual_falta_total = max(0, 1 - percentual_total)
    total_lojas_geral = comparativo["Loja"].nunique()

    linha_total = pd.DataFrame({
        "Ano": [""], "Mês": [""], "Grupo": [""],
        "Loja": [f"TOTAL GERAL - Lojas: {total_lojas_geral:02}"],
        "Meta": [total_meta], "Realizado": [total_realizado],
        "% Atingido": [percentual_total], "% Falta Atingir": [percentual_falta_total],
        "Diferença": [total_diferenca],
        "Tipo": [""]
    })

    comparativo_final = pd.concat([linha_total] + resultado_final, ignore_index=True)

    def formatar_linha(row):
        if "TOTAL GERAL" in row["Loja"]:
            return ['background-color: #0366d6; color: white'] * len(row)
        elif "Lojas:" in row["Loja"]:
            return ['background-color: #d0e6f7'] * len(row)
        else:
            return [''] * len(row)

    st.dataframe(
        comparativo_final.style
            .format({
                "Meta": "R$ {:,.2f}", "Realizado": "R$ {:,.2f}",
                "Diferença": "R$ {:,.2f}", "% Atingido": "{:.2%}",
                "% Falta Atingir": "{:.2%}"
            }).apply(formatar_linha, axis=1),
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
