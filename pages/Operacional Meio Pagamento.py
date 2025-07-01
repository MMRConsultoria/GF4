import streamlit as st
import pandas as pd
import numpy as np
import re
import json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Meio de Pagamento", layout="wide")

# 🔥 CSS para estilizar as abas igual ao seu modelo
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

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# 🔥 Título com ícone
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Vendas Diarias</h1>
    </div>
""", unsafe_allow_html=True)

# Conexão Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Vendas diarias")
df_empresa = pd.DataFrame(planilha.worksheet("Tabela Empresa").get_all_records())

# ========================
# Abas estilizadas com st.tabs
# ========================
tab1, tab2, tab3 = st.tabs([
    "📥 Upload e Processamento",
    "🔄 Atualizar Google Sheets",
    "📝 Desenvolvimento"
])

# ======================
# 📥 Upload e processamento
# ======================
with tab1:
    uploaded_file = st.file_uploader(
        label="📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento",
        type=["xlsx", "xlsm"],
        help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
    )

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            abas_disponiveis = xls.sheet_names

            if len(abas_disponiveis) == 1:
                aba_escolhida = abas_disponiveis[0]
            else:
                aba_escolhida = st.selectbox("Escolha a aba para processar", abas_disponiveis)

            df_raw = pd.read_excel(xls, sheet_name=aba_escolhida, header=None)
            df_raw = df_raw[~df_raw.iloc[:, 1].astype(str).str.lower().str.contains("total|subtotal", na=False)]

        except Exception as e:
            st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes: {e}")
        else:
            if str(df_raw.iloc[0, 1]).strip().lower() != "faturamento diário por meio de pagamento":
                st.error("❌ A célula B1 deve conter 'Faturamento diário por meio de pagamento'.")
                st.stop()

            linha_inicio_dados = 5
            blocos = []
            col = 3
            loja_atual = None

            while col < df_raw.shape[1]:
                valor_linha4 = str(df_raw.iloc[3, col]).strip()
                match = re.match(r"^\d+\s*-\s*(.+)$", valor_linha4)
                if match:
                    loja_atual = match.group(1).strip().lower()

                meio_pgto = str(df_raw.iloc[4, col]).strip()
                if not loja_atual or not meio_pgto or meio_pgto.lower() in ["nan", ""]:
                    col += 1
                    continue

                linha3 = str(df_raw.iloc[2, col]).strip().lower()
                linha5 = meio_pgto.lower()

                if any(palavra in texto for texto in [linha3, valor_linha4.lower(), linha5] for palavra in ["total", "serv/tx", "total real"]):
                    col += 1
                    continue

                try:
                    df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
                    df_temp.columns = ["Data", "Valor (R$)"]
                    df_temp = df_temp[~df_temp["Data"].astype(str).str.lower().str.contains("total|subtotal")]
                    df_temp.insert(1, "Meio de Pagamento", meio_pgto)
                    df_temp.insert(2, "Loja", loja_atual)
                    blocos.append(df_temp)
                except Exception as e:
                    st.warning(f"⚠️ Erro ao processar coluna {col}: {e}")

                col += 1

            if not blocos:
                st.error("❌ Nenhum dado válido encontrado na planilha.")
            else:
                df = pd.concat(blocos, ignore_index=True).dropna(how="any")
                df = df[~df["Data"].astype(str).str.lower().str.contains("total|subtotal")]
                df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
                df = df[df["Data"].notna()]

                dias_semana = {'Monday': 'segunda-feira','Tuesday': 'terça-feira','Wednesday': 'quarta-feira','Thursday': 'quinta-feira',
                               'Friday': 'sexta-feira','Saturday': 'sábado','Sunday': 'domingo'}
                df["Dia da Semana"] = df["Data"].dt.day_name().map(dias_semana)
                df = df.sort_values(by=["Data", "Loja"])
                periodo_min = df["Data"].min().strftime("%d/%m/%Y")
                periodo_max = df["Data"].max().strftime("%d/%m/%Y")
                df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

                df["Loja"] = df["Loja"].astype(str).str.strip().str.replace(r"^\d+\s*-\s*", "", regex=True).str.lower()
                df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
                df = pd.merge(df, df_empresa, on="Loja", how="left")

                for col in ["Código Everest", "Grupo", "Código Grupo Everest"]:
                    if col not in df.columns:
                        df[col] = np.nan

                df["Mês"] = pd.to_datetime(df["Data"], dayfirst=True).dt.month.map({
                    1: 'jan',2: 'fev',3: 'mar',4: 'abr',5: 'mai',6: 'jun',7: 'jul',8: 'ago',9: 'set',10: 'out',11: 'nov',12: 'dez'})
                df["Ano"] = pd.to_datetime(df["Data"], dayfirst=True).dt.year
                df = df[["Data", "Dia da Semana", "Meio de Pagamento", "Loja","Código Everest", "Grupo", "Código Grupo Everest","Valor (R$)", "Mês", "Ano"]]

                df_meios_pgto = pd.DataFrame(planilha.worksheet("Tabela Meio Pagamento").get_all_records())
                meios_cadastrados = df_meios_pgto["Meio de Pagamento"].astype(str).str.strip().str.lower().unique()
                meios_usados = df["Meio de Pagamento"].astype(str).str.strip().str.lower().unique()
                meios_nao_cadastrados = [m for m in meios_usados if m not in meios_cadastrados]

                lojas_sem_codigo = df[df["Código Everest"].isna()]["Loja"].unique()

                col1, col2 = st.columns(2)
                col1.markdown(f"<div style='font-size:1.2rem;'>📅 <strong>Período processado</strong><br>{periodo_min} até {periodo_max}</div>", unsafe_allow_html=True)

                tem_erros = False

                if len(meios_nao_cadastrados) > 0:
                    tem_erros = True
                    lista_meios = "<br>".join([f"- {m}" for m in meios_nao_cadastrados])
                    col1.markdown(f"""
                        <div style='color:#856404; font-size:0.95rem; margin-top:5px;'>
                        ⚠️ {len(meios_nao_cadastrados)} meio(s) de pagamento não localizado(s):<br>{lista_meios}<br>
                        Cadastre na Tabela Meio Pagamento e reprocessar!
                        </div>
                    """, unsafe_allow_html=True)

                valor_total_formatado = f"R$ {df['Valor (R$)'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                col2.markdown(f"<div style='font-size:1.2rem;'>💰 <strong>Valor total</strong><br><span style='color:green;'>{valor_total_formatado}</span></div>", unsafe_allow_html=True)

                if len(lojas_sem_codigo) > 0:
                    tem_erros = True
                    st.markdown(f"<div style='color:#856404; font-size:0.95rem; margin-top:5px;'>⚠️ Lojas sem código Everest cadastrado: {', '.join(lojas_sem_codigo)}<br>🔗 <a href='https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU/edit' target='_blank' style='color:#0d6efd;'>Atualize os dados na planilha de empresas</a></div>", unsafe_allow_html=True)

                if not tem_erros:
                    st.success("✅ Relatório de faturamento por meio de pagamento gerado com sucesso!")
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="FaturamentoPorMeio")
                    output.seek(0)
                    st.download_button("📥 Baixar relatório", data=output, file_name="FaturamentoPorMeio_transformado.xlsx")

# ======================
# 🔄 Atualizar Google Sheets
# ======================
with tab2:
    st.info("🚀 Aqui ficará sua funcionalidade para atualizar dados no Google Sheets.")

# ======================
# 📝 Auditar integração Everest
# ======================
with tab3:
    st.info("🔍 Desenvolvimento")
