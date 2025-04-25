import pandas as pd
import re

# Supondo que df_raw já está carregado da planilha com header=None
linha_inicio_dados = 6
blocos = []
col = 3
loja_atual = None

while col < df_raw.shape[1]:
    valor_linha4 = str(df_raw.iloc[3, col]).strip()

    # Verifica se é uma nova loja (ex: "30 - Central fase 1")
    match = re.match(r"^\d+\s*-\s*(.+)$", valor_linha4)
    if match:
        loja_atual = match.group(1).strip().lower()

    meio_pgto = str(df_raw.iloc[4, col]).strip()
    if not loja_atual or not meio_pgto or meio_pgto.lower() in ["nan", ""]:
        col += 1
        continue

    # Verifica palavras proibidas nas linhas 3, 4 e 5
    linha3 = str(df_raw.iloc[2, col]).strip().lower()
    linha5 = meio_pgto.lower()

    if any(palavra in texto for texto in [linha3, valor_linha4.lower(), linha5] for palavra in ["total", "serv/tx", "total real"]):
        col += 1
        continue

    # Bloco válido
    df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
    df_temp.columns = ["Data", "Valor (R$)"]
    df_temp.insert(1, "Meio de Pagamento", meio_pgto)
    df_temp.insert(2, "Loja", loja_atual)
    blocos.append(df_temp)

    col += 1

# Junta tudo ao final
if blocos:
    df_final = pd.concat(blocos, ignore_index=True)
    st.dataframe(df_final.head())
else:
    st.warning("Nenhum dado válido foi identificado.")
