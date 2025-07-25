blocos = []
grupos_info = []

for grupo, df_grp in df_base.groupby("Grupo"):
    total_grupo = df_grp[col_acumulado].sum()
    grupos_info.append((grupo, total_grupo, df_grp))

# Ordena grupos
grupos_info.sort(key=lambda x: x[1], reverse=True)

for grupo, _, df_grp in grupos_info:
    df_grp_ord = df_grp.sort_values(by=col_acumulado, ascending=False)

    # Subtotal
    subtotal_values = df_grp_ord.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
    subtotal_df = pd.DataFrame([subtotal_values])
    subtotal_df["Grupo"] = f"SUBTOTAL {grupo}"
    subtotal_df["Loja"] = f"Lojas: {df_grp_ord['Loja'].nunique():02d}"
    
    tipo_unico = df_grp_ord["Tipo"].dropna().unique()
    if len(tipo_unico) == 1:
        subtotal_df["Tipo"] = tipo_unico[0]
    else:
        subtotal_df["Tipo"] = ""

    if modo_exibicao == "Loja":
        blocos.append(df_grp_ord)

    blocos.append(subtotal_df)

# Total final
linha_total = df_base.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
linha_total["Grupo"] = "TOTAL"
linha_total["Loja"] = f"Lojas: {df_base['Loja'].nunique():02d}"
linha_total["Tipo"] = "TOTAL"

# Agora monta o df_final corretamente
df_final = pd.concat([pd.DataFrame([linha_total])] + blocos, ignore_index=True)
