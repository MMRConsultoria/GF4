# ================== Imports ==================
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import pytz
import io
import pandas as pd
import streamlit as st

# ================== Helpers de mês/ano ==================
MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
MAPA_MES_NOME2NUM = {m.lower(): i+1 for i, m in enumerate(MESES_PT)}

def mes_ano_label(ano: int, mes: int) -> str:
    return f"{MESES_PT[mes-1]}/{ano}"

def juntar_labels(labels: list[str]) -> str:
    """Transforma lista em 'A, B e C'."""
    if not labels:
        return ""
    if len(labels) == 1:
        return labels[0]
    if len(labels) == 2:
        return f"{labels[0]} e {labels[1]}"
    return f"{', '.join(labels[:-1])} e {labels[-1]}"

def extrair_meses_disponiveis(df: pd.DataFrame):
    """
    Retorna:
      - lista ordenada desc de (ano, mes) únicos
      - lista de labels 'Mês/AAAA' correspondente
    Aceita:
      - coluna 'Data' (dd/mm/aaaa ou datetime)
      - ou colunas 'Ano' e 'Mês'/'Mes' (número ou nome PT-BR)
    """
    df_cols = {c.strip().lower(): c for c in df.columns}
    pares = set()

    # 1) A partir de 'Data'
    if "data" in df_cols:
        datas = pd.to_datetime(df[df_cols["data"]], dayfirst=True, errors="coerce").dropna()
        if not datas.empty:
            for d in datas.dt.to_pydatetime():
                pares.add((d.year, d.month))

    # 2) A partir de 'Ano' e 'Mês/Mes' (se não trouxe pela 'Data')
    if not pares and "ano" in df_cols and (("mês" in df_cols) or ("mes" in df_cols)):
        col_ano = df_cols["ano"]
        col_mes = df_cols.get("mês", df_cols.get("mes"))
        tmp = df[[col_ano, col_mes]].dropna()

        def mes_to_num(v):
            s = str(v).strip().lower()
            if s.isdigit():
                n = int(float(s))
                return n if 1 <= n <= 12 else None
            return MAPA_MES_NOME2NUM.get(s)

        tmp["_MesNum"] = tmp[col_mes].apply(mes_to_num)
        tmp = tmp.dropna(subset=["_MesNum"])
        for a, m in zip(tmp[col_ano], tmp["_MesNum"]):
            try:
                pares.add((int(a), int(m)))
            except:
                pass

    # Fallback: mês/ano corrente
    if not pares:
        now = datetime.now(pytz.timezone("America/Sao_Paulo"))
        pares.add((now.year, now.month))

    # Ordena do mais recente para o mais antigo
    ordenado = sorted(pares, key=lambda t: (t[0], t[1]), reverse=True)
    labels = [mes_ano_label(a, m) for (a, m) in ordenado]
    return ordenado, labels

# ================== PDF ==================
def gerar_pdf(df, mes_rateio, usuario):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=30,
        bottomMargin=30,
        leftMargin=20,
        rightMargin=20
    )

    elementos = []
    estilos = getSampleStyleSheet()
    estilo_normal = estilos["Normal"]
    estilo_titulo = estilos["Heading1"]

    # Logo
    try:
        logo_url = "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_grupofit.png"
        img = Image(logo_url, width=100, height=40)
        elementos.append(img)
    except:
        pass

    # Título
    elementos.append(Paragraph(f"<b>Rateio - {mes_rateio}</b>", estilo_titulo))

    # Data no fuso correto
    fuso_brasilia = pytz.timezone("America/Sao_Paulo")
    data_geracao = datetime.now(fuso_brasilia).strftime("%d/%m/%Y %H:%M")

    elementos.append(Paragraph(f"<b>Usuário:</b> {usuario}", estilo_normal))
    elementos.append(Paragraph(f"<b>Data de Geração:</b> {data_geracao}", estilo_normal))
    elementos.append(Spacer(1, 12))

    # Converte DataFrame para lista (mantém a ordem das colunas do df)
    dados_tabela = [df.columns.tolist()] + df.values.tolist()

    # Tabela
    tabela = Table(dados_tabela, repeatRows=1)

    # Estilo inicial
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),  # Cabeçalho azul escuro
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ]))

    # Cores de linhas (subtotal/total na 2ª coluna; ajuste o índice se necessário)
    for i in range(1, len(dados_tabela)):
        linha_texto = str(dados_tabela[i][1]).strip().lower()
        if "subtotal" in linha_texto or "total" == linha_texto:
            tabela.setStyle(TableStyle([
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#BFBFBF")),
                ("FONTNAME", (0, i), (-1, i), "Helvetica-Bold")
            ]))
        else:
            tabela.setStyle(TableStyle([
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2"))
            ]))

    elementos.append(tabela)

    doc.build(elementos)
    pdf_value = buffer.getvalue()
    buffer.close()
    return pdf_value

# ================== UI (mês fechado, múltiplos) ==================
# Observação: espera-se que 'df_view' já exista no seu app antes deste arquivo/trecho.
usuario_logado = st.session_state.get("usuario_logado", "Usuário Desconhecido")

# Extrai meses disponíveis do df_view e cria opções
_, labels_opcoes = extrair_meses_disponiveis(df_view)

# Multiselect (padrão: mês mais recente)
meses_escolhidos = st.multiselect(
    "Selecione o(s) mês(es)",
    options=labels_opcoes,
    default=labels_opcoes[:1] if labels_opcoes else []
)

# Se o usuário desmarcar tudo, força o mês mais recente (se existir)
if not meses_escolhidos and labels_opcoes:
    meses_escolhidos = labels_opcoes[:1]

# Monta o título a partir da(s) escolha(s)
mes_rateio = juntar_labels(meses_escolhidos)

# (Opcional) Caso queira filtrar o df_view apenas para os meses escolhidos,
# descomente e ajuste de acordo com as colunas do seu df:
# ----------------------------------------------------------------
# def filtrar_df_por_labels(df, labels):
#     df_cols = {c.strip().lower(): c for c in df.columns}
#     # Preferência por coluna 'Data'
#     if "data" in df_cols:
#         dt = pd.to_datetime(df[df_cols["data"]], dayfirst=True, errors="coerce")
#         mask = pd.Series(False, index=df.index)
#         for lb in labels:
#             mes_nome, ano_str = lb.split("/")
#             ano = int(ano_str)
#             mes = MAPA_MES_NOME2NUM[mes_nome.lower()]
#             mask = mask | ((dt.dt.year == ano) & (dt.dt.month == mes))
#         return df.loc[mask].copy()
#     # Alternativa: colunas 'Ano' e 'Mês/Mes'
#     if "ano" in df_cols and (("mês" in df_cols) or ("mes" in df_cols)):
#         col_ano = df_cols["ano"]
#         col_mes = df_cols.get("mês", df_cols.get("mes"))
#         tmp = df.copy()
#         def mes_to_num(v):
#             s = str(v).strip().lower()
#             if s.isdigit():
#                 n = int(float(s))
#                 return n if 1 <= n <= 12 else None
#             return MAPA_MES_NOME2NUM.get(s)
#         tmp["_MesNum"] = tmp[col_mes].apply(mes_to_num)
#         mask = pd.Series(False, index=tmp.index)
#         for lb in labels:
#             mes_nome, ano_str = lb.split("/")
#             ano = int(ano_str)
#             mes = MAPA_MES_NOME2NUM[mes_nome.lower()]
#             mask = mask | ((tmp[col_ano].astype(int) == ano) & (tmp["_MesNum"] == mes))
#         tmp = tmp.loc[mask].drop(columns=["_MesNum"])
#         return tmp
#     return df
#
# df_para_pdf = filtrar_df_por_labels(df_view, meses_escolhidos)
# ----------------------------------------------------------------

# Sem filtrar as linhas (mantém df_view como está)
df_para_pdf = df_view

# Gera o PDF com o título dinâmico
pdf_bytes = gerar_pdf(df_para_pdf, mes_rateio=mes_rateio, usuario=usuario_logado)

# Botão de download
st.download_button(
    label="📄 Baixar PDF",
    data=pdf_bytes,
    file_name=f"Rateio_{datetime.now().strftime('%Y%m%d')}.pdf",
    mime="application/pdf"
)
