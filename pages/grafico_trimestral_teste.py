import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Teste Gráfico Trimestral", layout="wide")

st.title("📊 Teste de Gráfico Trimestral")

# 🔧 Dados simulados
dados = {
    "Data": pd.date_range(start="2024-01-01", periods=180, freq="7D"),
    "Fat.Real": [1000 + (i % 4) * 500 for i in range(180)]
}

df = pd.DataFrame(dados)
df["Ano"] = df["Data"].dt.year
df["Trimestre"] = df["Data"].dt.quarter
df["Nome Trimestre"] = "T" + df["Trimestre"].astype(str)

# 📈 Agrupamento
fat_trimestral = df.groupby(["Nome Trimestre", "Ano"])["Fat.Real"].sum().reset_index()
fat_trimestral["TrimestreNum"] = fat_trimestral["Nome Trimestre"].str.extract(r'(\d)').astype(int)
fat_trimestral["Ano"] = fat_trimestral["Ano"].astype(str)
fat_trimestral = fat_trimestral.sort_values(["TrimestreNum", "Ano"])

st.write("✅ Dados agrupados:")
st.dataframe(fat_trimestral)

# 🎨 Gráfico
fig = px.bar(
    fat_trimestral,
    x="Nome Trimestre",
    y="Fat.Real",
    color="Ano",
    barmode="group",
    text="Fat.Real",
    color_discrete_map={"2024": "#1f77b4", "2025": "#ff7f0e"}
)

fig.update_traces(textposition="outside")
fig.update_layout(
    xaxis_title=None,
    yaxis_title=None,
    xaxis_tickangle=-45,
    showlegend=False,
    yaxis=dict(showticklabels=False, showgrid=False, zeroline=False)
)

st.plotly_chart(fig, use_container_width=True)
