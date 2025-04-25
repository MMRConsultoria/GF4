linha_inicio_dados = 6
blocos = []
col = 3
loja_atual = None

while col < df_raw.shape[1]:
    # Linha 4 pode conter uma nova loja
    valor_linha4 = str(df_raw.iloc[3, col]).strip()

    # Verifica se a célula tem o padrão "número - nome da loja"
    match = re.match(r"^\d+\s*-\s*(.+)$", valor_linha4)
    if match:
        loja_atual = match.group(1).strip().lower()

    # Se não temos uma loja atual definida, ou a linha 5 está vazia, pula
    meio_pgto = str(df_raw.iloc[4, col]).strip()
    if not loja_atual or not meio_pgto or meio_pgto.lower() in ["nan", ""]:
        col += 1
        continue

    # Verifica se alguma das 3 primeiras linhas tem palavras proibidas
    linha3 = str(df_raw.iloc[2, col]).strip().lower()
    linha5 = meio_pgto.lower()

    if any(palavra in texto for texto in [linha3, valor_linha4.lower(), linha5] for palavra in ["total", "serv/tx", "total real"]):
        col += 1
        continue

    # Extrai os dados da coluna válida
    df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
    df_temp.columns = ["Data", "Valor (R$)"]
    df_temp.insert(1, "Meio de Pagamento", meio_pgto)
    df_temp.insert(2, "Loja", loja_atual)
    blocos.append(df_temp)

    col += 1
