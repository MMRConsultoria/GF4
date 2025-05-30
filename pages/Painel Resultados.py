

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


# ✅ Adiciona grupos ativos sem movimento no df_filtrado (zerado)
        grupos_presentes = df_filtrado["Grupo"].dropna().unique()
        grupos_sem_movimento = list(set(grupos_ativos) - set(grupos_presentes))

        if grupos_sem_movimento:
            df_sem_mov = pd.DataFrame({
                "Grupo": grupos_sem_movimento,
                "Loja": [f"{g} - Loja: 0" for g in grupos_sem_movimento],
                "Fat.Total": 0,
                "Fat.Real": 0,
                "Serv/Tx": 0,
                "Ticket": 0,
                "Data": None,
                "Ano": None,
                "Mês Num": None,
                "Mês Nome": None,
                "Mês": None,
                "Dia": None,
                "Agrupador": "ACUMULADO",
                "Ordem": 99999999
            })

            df_filtrado = pd.concat([df_filtrado, df_sem_mov], ignore_index=True)
            tabela_exportar_sem_tipo = pd.concat([tabela_exportar_sem_tipo, df_sem_mov], ignore_index=True)

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
