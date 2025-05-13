# pages/OperacionalFaturamentoReal.py



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
aba1, aba2, aba3, aba4 = st.tabs(["📄 Upload e Processamento", "📥 Download Excel", "🔄 Atualizar Google Sheets","📊 Comparativo Everest"])

# ================================
# 📄 Aba 1 - Upload e Processamento
# ================================
with aba1:
    uploaded_file = st.file_uploader("📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento", type=["xlsx"])

    if uploaded_file:
        try:
            # 🔹 Carregar o arquivo
            xls = pd.ExcelFile(uploaded_file)
            df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

            # 🔹 Validar B1
            texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
            if texto_b1 != "faturamento diário sintético multi-loja":
                st.error(f"❌ A célula B1 está com '{texto_b1}'. Corrija para 'Faturamento diário sintético multi-loja'.")
                st.stop()

            df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
            df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], dayfirst=True, errors='coerce')

            # 🔹 Processamento dos registros
            registros = []
            col = 3
            while col < df.shape[1]:
                nome_loja = str(df_raw.iloc[3, col]).strip()
                if re.match(r"^\d+\s*-?\s*", nome_loja):
                    nome_loja = nome_loja.split("-", 1)[-1].strip()

                    header_col = str(df.iloc[0, col]).strip().lower()
                    if "fat.total" in header_col:
                        for i in range(1, df.shape[0]):
                            linha = df.iloc[i]
                            valor_data = df.iloc[i, 2]
                            valor_check = str(df.iloc[i, 1]).strip().lower()

                            if pd.isna(valor_data) or valor_check in ["total", "subtotal"]:
                                continue

                            data = valor_data
                            valores = linha[col:col+5].values

                            if pd.isna(valores).all():
                                continue

                            registros.append([
                                data,
                                nome_loja,
                                *valores,
                                data.strftime("%b"),
                                data.year
                            ])
                    col += 5
                else:
                    col += 1

            if len(registros) == 0:
                st.warning("⚠️ Nenhum registro encontrado.")

            # 🔹 Montar df_final
            df_final = pd.DataFrame(registros, columns=[
                "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ])

            df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
            df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
            df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")

            # 🔹 Ajustar dados
            dias_traducao = {
                "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
                "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
            }
            df_final.insert(1, "Dia da Semana", pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.day_name().map(dias_traducao))
            df_final["Data"] = pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.strftime("%d/%m/%Y")

            for col_val in ["Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas"]:
                df_final[col_val] = pd.to_numeric(df_final[col_val], errors="coerce").round(2)

            meses = {
                "jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"
            }
            df_final["Mês"] = df_final["Mês"].str.lower().map(meses)

            df_final["Data_Ordenada"] = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce')
            df_final = df_final.sort_values(by=["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")

            colunas_finais = [
                "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
                "Ticket", "Mês", "Ano"
            ]
            df_final = df_final[colunas_finais]

            st.session_state.df_final = df_final
            st.session_state.atualizou_google = False

            # 🔥 Agora exibir:
            # 📄 Nome do Arquivo
           # st.markdown(f"""
              #  <div style='font-size:15px; font-weight: bold; margin-bottom:10px;'>
                   # 📄 Arquivo selecionado: {uploaded_file.name}
              #  </div>
            #""", unsafe_allow_html=True)

            # 📅 e 💰 Período e Valor Total
            datas_validas = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce').dropna()

            if not datas_validas.empty:
                data_inicial = datas_validas.min().strftime("%d/%m/%Y")
                data_final = datas_validas.max().strftime("%d/%m/%Y")
                
                valor_total = df_final["Fat.Total"].sum().round(2)
                valor_total_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown(f"""
                        <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>📅 Período processado</div>
                        <div style='font-size:30px; color:#000;'>{data_inicial} até {data_final}</div>
                    """, unsafe_allow_html=True)

                with col2:
                    st.markdown(f"""
                        <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>💰 Valor total</div>
                        <div style='font-size:30px; color:green;'>{valor_total_formatado}</div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ Não foi possível identificar o período de datas.")

           # 🔎 Empresas não localizadas
            empresas_nao_localizadas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()

            if len(empresas_nao_localizadas) > 0:
                # Listar as empresas não localizadas
                empresas_nao_localizadas_str = "<br>".join(empresas_nao_localizadas)
                
                # Construir a mensagem com o link direto
                mensagem = f"""
                ⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s): 
                <br>{empresas_nao_localizadas_str}
                <br>
                ✏️ Atualize a tabela clicando 
                <a href='https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link' target='_blank'><strong>aqui</strong></a>.
                """

                st.markdown(mensagem, unsafe_allow_html=True)


            else:
                st.success("✅ Todas as empresas foram localizadas na Tabela_Empresa!")

           # 🔗 Links úteis
           # st.markdown("""
#🔗 [Link **Tabela_Empresa**](https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link)

#""")

        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {e}")

# ================================
# 📥 Aba 2 - Download Excel
# ================================
with aba2:
    st.header("📥 Download Relatório Excel")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
            output.seek(0)
            return output

        excel_data = to_excel(df_final)

        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=excel_data,
            file_name="faturamento_servico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo na aba anterior.")
# =======================================
# Atualizar Google Sheets (Evitar duplicação)
# =======================================

with aba3:

     
    #st.header("📤 Atualizar Banco de Dados (Evitar duplicação usando coluna M)")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final.copy()
        
        # 🔗 Links úteis
        st.markdown("""
          🔗 [Link  **Faturamento Sistema Externo**](https://docs.google.com/spreadsheets/d/1_3uX7dlvKefaGDBUhWhyDSLbfXzAsw8bKRVvfiIz8ic/edit?usp=sharing)
        """, unsafe_allow_html=True)

        # Criar a coluna "M" com a concatenação de "Data", "Fat.Total" e "Loja" como string para verificação de duplicação
        df_final['M'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d') + \
                        df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)

        #df_final['M'] = df_final['Data'] + df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)


        # Não converter para string, apenas utilizar "M" para verificação de duplicação
        df_final['M'] = df_final['M'].apply(str)

        # Converter o restante do DataFrame para string, mas mantendo as colunas numéricas com seu formato correto
        df_final = df_final.applymap(str)
        

      

        #TIRAR ASPAS DOS VALORES, DATA E NUMEROS

        
        
        # Formatando os valores monetários (não convertendo para string, mantendo como numérico)
        df_final['Fat.Total'] = df_final['Fat.Total'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
        df_final['Serv/Tx'] = df_final['Serv/Tx'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
        df_final['Fat.Real'] = df_final['Fat.Real'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
        df_final['Ticket'] = df_final['Ticket'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)

        # Garantir datetime sem aspas
        df_final['Data'] = pd.to_datetime(df_final['Data'].astype(str).str.replace("'", "").str.strip(), dayfirst=True)

        # Converter para número serial (dias desde 1899-12-30, padrão do Excel/Sheets)
        df_final['Data'] = (df_final['Data'] - pd.Timestamp("1899-12-30")).dt.days
       
        # Corrigir coluna Ano: remover aspas, espaços e garantir que seja inteiro
        df_final['Ano'] = df_final['Ano'].apply(
        lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
        )

        # Corrigir colunas D e F (Código Everest e Código Grupo Everest): remover aspas e converter para número inteiro
        df_final['Código Everest'] = df_final['Código Everest'].apply(
        lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
        )

        df_final['Código Grupo Everest'] = df_final['Código Grupo Everest'].apply(
        lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
        )
        
        
        # Conectar ao Google Sheets
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
        gc = gspread.authorize(credentials)

        planilha_destino = gc.open("Faturamento Sistema Externo")
        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

        # Obter dados já existentes na aba
        valores_existentes = aba_destino.get_all_values()

        # Criar um conjunto de linhas existentes na coluna M (usada para verificar duplicação)
        dados_existentes = set([linha[12] for linha in valores_existentes[1:]])  # Ignorando cabeçalho, coluna M é a 13ª (índice 12)

        novos_dados = []
        duplicados = []  # Armazenar os registros duplicados
        rows = df_final.fillna("").values.tolist()

       

        
        # Verificar duplicação somente na coluna "M"
        for linha in rows:
            chave_m = linha[-1]  # A chave da coluna M (última coluna)
            if chave_m not in dados_existentes:
                novos_dados.append(linha)
                dados_existentes.add(chave_m)  # Adiciona a chave da linha para não enviar novamente
            else:
                duplicados.append(linha)  # Adiciona a linha duplicada à lista

        # Adicionar o botão de atualização do Google Sheets
        if st.button("📥 Enviar dados para o Google Sheets"):
            with st.spinner("🔄 Atualizando o Google Sheets..."):
                try:
                    if novos_dados:
                        # Manter a primeira linha vazia para começar a inserção
                        primeira_linha_vazia = len(valores_existentes) + 1
                        
                        # Enviar os novos dados para o Google Sheets
                        aba_destino.update(f"A{primeira_linha_vazia}", novos_dados)

# ASPAS RESOLVIDO
                        
                        # 🔧 Aplicar formatação de data na coluna A (Data) - prbblema de aspas resolvido
                        from gspread_formatting import CellFormat, NumberFormat, format_cell_range

                        data_format = CellFormat(
                            numberFormat=NumberFormat(type='DATE', pattern='dd/mm/yyyy')
                        )

                        # 🔢 Formato para coluna Ano como número sem aspas
                        numero_format = CellFormat(
                        numberFormat=NumberFormat(type='NUMBER', pattern='0')
                        )
                      
                        
                        # Considerando que a coluna A é onde está a data
                        format_cell_range(aba_destino, f"A2:A{primeira_linha_vazia + len(novos_dados)}", data_format)
                        format_cell_range(aba_destino, f"L2:L{primeira_linha_vazia + len(novos_dados)}", numero_format)  
                        format_cell_range(aba_destino, f"D2:D{primeira_linha_vazia + len(novos_dados)}", numero_format)
                        format_cell_range(aba_destino, f"F2:F{primeira_linha_vazia + len(novos_dados)}", numero_format)




                        
                        st.success(f"✅ {len(novos_dados)} novo(s) registro(s) enviado(s) com sucesso para o Google Sheets!")

                    if duplicados:
                        st.warning(f"⚠️ {len(duplicados)} registro(s) foram duplicados e não foram enviados para o Google Sheets.")
                        # Exibir as linhas duplicadas para o usuário
                   #     st.write("Registros Duplicados:", duplicados)

                   # else:
                    #    st.info("✅ Dados atualizados google sheets.")
                except Exception as e:
                    st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")

    else:
        st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")


# =======================================
# Aba 4 - Comparativo Everest (independente do upload)
# =======================================

with aba4:
    st.header("📊 Comparativo Everest (via Google Sheets)")

    try:
        planilha = gc.open("Faturamento Sistema Externo")
        aba_everest = planilha.worksheet("Fat Everest")
        aba_externo = planilha.worksheet("Fat Sistema Externo")

        df_everest = pd.DataFrame(aba_everest.get_all_values()[1:])  # Ignora cabeçalho
        df_externo = pd.DataFrame(aba_externo.get_all_values()[1:])

        df_everest.columns = [f"col{i}" for i in range(df_everest.shape[1])]
        df_externo.columns = [f"col{i}" for i in range(df_externo.shape[1])]

        # Converter col0 em datetime
        df_everest["col0"] = pd.to_datetime(df_everest["col0"], dayfirst=True, errors="coerce")
        df_externo["col0"] = pd.to_datetime(df_externo["col0"], dayfirst=True, errors="coerce")

        # Filtro só aparece se houver dados válidos
        datas_validas = df_everest["col0"].dropna()
        if not datas_validas.empty:
            min_data = datas_validas.min().date()
            max_data = datas_validas.max().date()

            data_range = st.date_input(
                "Selecione o intervalo de datas:",
                value=(min_data, max_data),  # fornece um intervalo válido por padrão
                min_value=min_data,
                max_value=max_data
        )

# Verifica se o usuário realmente selecionou dois valores
if isinstance(data_range, tuple) and len(data_range) == 2:
    data_inicio, data_fim = data_range
else:
    data_inicio, data_fim = None, None


            # Só executa se o usuário escolher um intervalo
            if data_inicio and data_fim:
                # Traduzir data para português (visualmente)
                meses_pt = {
                    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
                    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
                    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
                }
                dias_pt = {
                    0: "segunda-feira", 1: "terça-feira", 2: "quarta-feira",
                    3: "quinta-feira", 4: "sexta-feira", 5: "sábado", 6: "domingo"
                }

                def data_pt(data):
                    return f"{dias_pt[data.weekday()]}, {data.day} de {meses_pt[data.month]} de {data.year}"

                st.markdown(f"📅 Período selecionado: **{data_pt(data_inicio)}** até **{data_pt(data_fim)}**")

                # Filtrar os dataframes
                ev = df_everest[(df_everest["col0"].dt.date >= data_inicio) & (df_everest["col0"].dt.date <= data_fim)].reset_index(drop=True)
                ex = df_externo[(df_externo["col0"].dt.date >= data_inicio) & (df_externo["col0"].dt.date <= data_fim)].reset_index(drop=True)

                tam = min(len(ev), len(ex))
                diferencas = []

                for i in range(tam):
                    linha_ev = ev.loc[i]
                    linha_ex = ex.loc[i]

                    def tratar_valor(valor):
                        return float(str(valor).replace("R$", "").replace(".", "").replace(",", ".").strip())

                    try:
                        val_ev = tratar_valor(linha_ev["col7"])  # H
                        val_ex = tratar_valor(linha_ex["col6"])  # G
                        val_diff_ext = tratar_valor(linha_ex["col8"])  # I
                    except:
                        continue

                    if (
                        linha_ev["col1"] != linha_ex["col3"] or  # B vs D
                        round(val_ev, 2) != round(val_ex, 2) or
                        round(val_ev - val_ex, 2) != round(val_diff_ext, 2)
                    ):
                        diferencas.append({
                            "Data": linha_ev["col0"].strftime("%d/%m/%Y") if pd.notnull(linha_ev["col0"]) else "",
                            "Loja (Everest - B)": linha_ev["col1"],
                            "Loja (Externo - D)": linha_ex["col3"],
                            "Nome Loja (Externo - col2)": linha_ex["col2"],
                            "Valor H (Everest)": linha_ev["col7"],
                            "Valor G (Externo)": linha_ex["col6"],
                            "H - G (calculado)": round(val_ev - val_ex, 2),
                            "Valor I (Externo)": linha_ex["col8"]
                        })

                if diferencas:
                    st.warning(f"⚠️ {len(diferencas)} diferença(s) encontrada(s):")
                    st.dataframe(pd.DataFrame(diferencas))
                else:
                    st.success("✅ Nenhuma diferença encontrada no período selecionado.")
            else:
                st.info("👆 Selecione o intervalo de datas para realizar a análise.")
        else:
            st.warning("⚠️ Nenhuma data válida encontrada nas abas do Google Sheets.")

    except Exception as e:
        st.error(f"❌ Erro ao carregar ou comparar dados: {e}")
