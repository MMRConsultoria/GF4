# 🔥 Criação da tabela, dependendo do modo de visão
if modo_visao == "Por Loja":
    # 🔗 Agrupamento e pivotagem por Loja
    tab_b = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Total", aggfunc="sum", fill_value=0)
    tab_r = df_filtrado.pivot_table(index="Loja", columns="Agrupador", values="Fat.Real", aggfunc="sum", fill_value=0)

    if tipo_metrica == "Bruto":
        tabela = tab_b
    elif tipo_metrica == "Real":
        tabela = tab_r
    else:
        tab_b.columns = [f"{c} (Bruto)" for c in tab_b.columns]
        tab_r.columns = [f"{c} (Real)" for c in tab_r.columns]
        tabela = pd.concat([tab_b, tab_r], axis=1)

        colunas_intercaladas = []
        for col in ordem:
            colunas_intercaladas.append(f"{col} (Bruto)")
            colunas_intercaladas.append(f"{col} (Real)")

        colunas_intercaladas = [c for c in colunas_intercaladas if c in tabela.columns]
        tabela = tabela[colunas_intercaladas]

elif modo_visao == "Por Grupo":
    df_grouped = df_filtrado.groupby(["Grupo", "Agrupador"]).agg(
        Bruto=("Fat.Total", "sum"),
        Real=("Fat.Real", "sum")
    ).reset_index()

    if df_grouped.empty:
        st.warning("🚨 Não há dados para os filtros selecionados no modo 'Por Grupo'.")
        st.stop()

    if tipo_metrica == "Bruto":
        tabela = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Bruto").fillna(0)
    elif tipo_metrica == "Real":
        tabela = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Real").fillna(0)
    else:
        tab_b = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Bruto").fillna(0)
        tab_r = df_grouped.pivot(index="Grupo", columns="Agrupador", values="Real").fillna(0)

        tab_b.columns = [f"{c} (Bruto)" for c in tab_b.columns]
        tab_r.columns = [f"{c} (Real)" for c in tab_r.columns]

        tabela = pd.concat([tab_b, tab_r], axis=1)

        colunas_intercaladas = []
        for col in ordem:
            colunas_intercaladas.append(f"{col} (Bruto)")
            colunas_intercaladas.append(f"{col} (Real)")

        colunas_intercaladas = [c for c in colunas_intercaladas if c in tabela.columns]
        tabela = tabela[colunas_intercaladas]
