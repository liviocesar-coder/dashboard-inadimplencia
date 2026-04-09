import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(
    page_title="Dashboard de Recebimentos | EXBR",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO = "Recebimentos Online - Fortel e Wirelink - TRATADA.xlsx"
LOGO = "logoexbr.png"   # troque se seu arquivo tiver outro nome

# Paleta inspirada na EXBR
COR_PRIMARIA = "#0F5B78"
COR_SECUNDARIA = "#21D59B"
COR_FUNDO = "#F3F6F8"
COR_CARD = "#FFFFFF"
COR_TEXTO = "#1F2937"
COR_MUTED = "#6B7280"
COR_BORDA = "#D9E2E8"


def formatar_moeda(valor):
    return f"R$ {valor:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def aplicar_estilo():
    st.markdown(
        f"""
        <style>
        /* Fundo geral */
        .stApp {{
            background-color: {COR_FUNDO};
            color: {COR_TEXTO};
        }}

        /* Esconde a barra preta superior do Streamlit */
        [data-testid="stHeader"] {{
            display: none;
        }}

        /* Esconde menu flutuante superior */
        [data-testid="stToolbar"] {{
            display: none;
        }}

        /* Ajusta espaçamento do conteúdo */
        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 96rem;
        }}

        /* Sidebar */
        section[data-testid="stSidebar"] {{
            background-color: #FFFFFF;
            border-right: 1px solid {COR_BORDA};
        }}

        /* Textos */
        h1, h2, h3, h4, h5, h6, p, label, div, span {{
            color: {COR_TEXTO};
        }}

        /* Título da página */
        .titulo-pagina {{
            font-size: 2.2rem;
            font-weight: 800;
            color: {COR_PRIMARIA};
            margin-bottom: 0.15rem;
        }}

        .subtitulo-pagina {{
            font-size: 1rem;
            color: {COR_MUTED};
            margin-bottom: 1.1rem;
        }}

        /* Cards KPI */
        .kpi-card {{
            background: {COR_CARD};
            border: 1px solid {COR_BORDA};
            border-radius: 18px;
            padding: 16px 18px;
            box-shadow: 0 2px 8px rgba(15, 91, 120, 0.06);
            min-height: 118px;
            overflow: hidden;
        }}

        .kpi-label {{
            font-size: 0.95rem;
            font-weight: 700;
            color: {COR_MUTED};
            margin-bottom: 12px;
        }}

        .kpi-value {{
            font-size: clamp(1.15rem, 1.25vw, 1.9rem);
            line-height: 1.05;
            font-weight: 800;
            color: {COR_PRIMARIA};
            white-space: nowrap;
            overflow: hidden;
            text-overflow: clip;
            letter-spacing: -0.5px;
        }}

        /* Títulos de seção */
        .secao-titulo {{
            font-size: 1.2rem;
            font-weight: 800;
            color: {COR_PRIMARIA};
            margin-top: 0.2rem;
            margin-bottom: 0.7rem;
        }}

        /* Inputs */
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div {{
            background-color: #FFFFFF !important;
        }}

        .stRadio > div {{
            background-color: transparent !important;
        }}

        .stAlert {{
            border-radius: 14px;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )


def card_kpi(label, valor):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{valor}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


@st.cache_data
def carregar_dados():
    online = pd.read_excel(ARQUIVO, sheet_name="ONLINE", header=1)
    htm = pd.read_excel(ARQUIVO, sheet_name="HTM", header=1)

    online["EMPRESA"] = "ONLINE"
    htm["EMPRESA"] = "HTM"

    online.columns = [str(col).strip() for col in online.columns]
    htm.columns = [str(col).strip() for col in htm.columns]

    df = pd.concat([online, htm], ignore_index=True)
    df = df.dropna(how="all")

    colunas_numericas = ["VALOR", "SALDO", "LIQUIDAÇÃO", "PAGO OU RECEB", "RECEB PERÍODO"]
    for col in colunas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    colunas_data = ["DATA", "EMISSÃO", "VENCIMENTO", "DATA LIQUIDAÇÃO"]
    for col in colunas_data:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "SALDO" in df.columns:
        df["STATUS"] = df["SALDO"].apply(lambda x: "Pago" if x <= 0 else "Não Pago")
    elif "RECEB PERÍODO" in df.columns:
        df["STATUS"] = df["RECEB PERÍODO"].apply(lambda x: "Pago" if x > 0 else "Não Pago")
    else:
        df["STATUS"] = "Não definido"

    if "VENCIMENTO" in df.columns and "SALDO" in df.columns:
        hoje = pd.Timestamp.today().normalize()
        df["VENCIDO"] = df.apply(
            lambda row: "Vencido"
            if pd.notnull(row["VENCIMENTO"]) and row["VENCIMENTO"] < hoje and row["SALDO"] > 0
            else "Em Dia",
            axis=1
        )
    else:
        df["VENCIDO"] = "Sem vencimento"

    return df


def montar_serie_temporal(df_base, visao="Mensal"):
    if "EMISSÃO" in df_base.columns:
        base_faturado = df_base.dropna(subset=["EMISSÃO"]).copy()
        base_faturado["DATA_REF"] = base_faturado["EMISSÃO"]
    elif "DATA" in df_base.columns:
        base_faturado = df_base.dropna(subset=["DATA"]).copy()
        base_faturado["DATA_REF"] = base_faturado["DATA"]
    else:
        base_faturado = pd.DataFrame(columns=["DATA_REF", "VALOR"])

    if "VENCIMENTO" in df_base.columns and "SALDO" in df_base.columns:
        base_inad = df_base[df_base["SALDO"] > 0].dropna(subset=["VENCIMENTO"]).copy()
        base_inad["DATA_REF"] = base_inad["VENCIMENTO"]
    else:
        base_inad = pd.DataFrame(columns=["DATA_REF", "SALDO"])

    if visao == "Mensal":
        if not base_faturado.empty:
            faturado = (
                base_faturado.groupby(base_faturado["DATA_REF"].dt.to_period("M"))["VALOR"]
                .sum()
                .reset_index()
            )
            faturado["PERIODO"] = faturado["DATA_REF"].astype(str)
            faturado["TIPO"] = "Faturado"
            faturado["VALOR_GRAFICO"] = faturado["VALOR"]
            faturado = faturado[["PERIODO", "TIPO", "VALOR_GRAFICO"]]
        else:
            faturado = pd.DataFrame(columns=["PERIODO", "TIPO", "VALOR_GRAFICO"])

        if not base_inad.empty:
            inad = (
                base_inad.groupby(base_inad["DATA_REF"].dt.to_period("M"))["SALDO"]
                .sum()
                .reset_index()
            )
            inad["PERIODO"] = inad["DATA_REF"].astype(str)
            inad["TIPO"] = "Inadimplente"
            inad["VALOR_GRAFICO"] = inad["SALDO"]
            inad = inad[["PERIODO", "TIPO", "VALOR_GRAFICO"]]
        else:
            inad = pd.DataFrame(columns=["PERIODO", "TIPO", "VALOR_GRAFICO"])
    else:
        if not base_faturado.empty:
            faturado = (
                base_faturado.groupby(base_faturado["DATA_REF"].dt.year)["VALOR"]
                .sum()
                .reset_index()
            )
            faturado["PERIODO"] = faturado["DATA_REF"].astype(str)
            faturado["TIPO"] = "Faturado"
            faturado["VALOR_GRAFICO"] = faturado["VALOR"]
            faturado = faturado[["PERIODO", "TIPO", "VALOR_GRAFICO"]]
        else:
            faturado = pd.DataFrame(columns=["PERIODO", "TIPO", "VALOR_GRAFICO"])

        if not base_inad.empty:
            inad = (
                base_inad.groupby(base_inad["DATA_REF"].dt.year)["SALDO"]
                .sum()
                .reset_index()
            )
            inad["PERIODO"] = inad["DATA_REF"].astype(str)
            inad["TIPO"] = "Inadimplente"
            inad["VALOR_GRAFICO"] = inad["SALDO"]
            inad = inad[["PERIODO", "TIPO", "VALOR_GRAFICO"]]
        else:
            inad = pd.DataFrame(columns=["PERIODO", "TIPO", "VALOR_GRAFICO"])

    return pd.concat([faturado, inad], ignore_index=True)


aplicar_estilo()
df = carregar_dados()

# SIDEBAR
if Path(LOGO).exists():
    try:
        st.sidebar.image(LOGO, use_container_width=True)
    except Exception:
        st.sidebar.warning("A logo foi encontrada, mas não pôde ser lida como imagem.")
else:
    st.sidebar.warning("Logo não encontrada na pasta do projeto.")

st.sidebar.markdown("## Filtros")

empresas = st.sidebar.multiselect(
    "Empresa",
    options=sorted(df["EMPRESA"].dropna().unique()),
    default=sorted(df["EMPRESA"].dropna().unique())
)

status = st.sidebar.multiselect(
    "Status",
    options=sorted(df["STATUS"].dropna().unique()),
    default=sorted(df["STATUS"].dropna().unique())
)

if "PESSOA" in df.columns:
    pessoas = st.sidebar.multiselect(
        "Cliente / Pessoa",
        options=sorted(df["PESSOA"].dropna().astype(str).unique()),
        default=[]
    )
else:
    pessoas = []

vencido = st.sidebar.multiselect(
    "Situação de Vencimento",
    options=sorted(df["VENCIDO"].dropna().unique()),
    default=sorted(df["VENCIDO"].dropna().unique())
)

visao_tempo = st.sidebar.radio(
    "Visualização da série temporal",
    options=["Mensal", "Anual"],
    index=0
)

# FILTROS
df_filtrado = df[df["EMPRESA"].isin(empresas)]
df_filtrado = df_filtrado[df_filtrado["STATUS"].isin(status)]
df_filtrado = df_filtrado[df_filtrado["VENCIDO"].isin(vencido)]

if pessoas:
    df_filtrado = df_filtrado[df_filtrado["PESSOA"].astype(str).isin(pessoas)]

# TOPO
st.markdown('<div class="titulo-pagina">Dashboard de Recebimentos</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo-pagina">Visão consolidada de ONLINE e HTM</div>', unsafe_allow_html=True)

# KPIs
total_registros = len(df_filtrado)
total_valor = df_filtrado["VALOR"].sum() if "VALOR" in df_filtrado.columns else 0
total_saldo = df_filtrado["SALDO"].sum() if "SALDO" in df_filtrado.columns else 0
total_pago = (
    df_filtrado.loc[df_filtrado["STATUS"] == "Pago", "VALOR"].sum()
    if "VALOR" in df_filtrado.columns else 0
)
inadimplencia = (total_saldo / total_valor) * 100 if total_valor > 0 else 0

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    card_kpi("Registros", f"{total_registros:,}".replace(",", "."))
with c2:
    card_kpi("Valor Total", formatar_moeda(total_valor))
with c3:
    card_kpi("Saldo em Aberto", formatar_moeda(total_saldo))
with c4:
    card_kpi("Valor Pago", formatar_moeda(total_pago))
with c5:
    card_kpi("Inadimplência", f"{inadimplencia:.2f}%")

st.divider()

# GRÁFICO DE LINHAS
st.markdown('<div class="secao-titulo">Faturamento x Inadimplência</div>', unsafe_allow_html=True)

serie_tempo = montar_serie_temporal(df_filtrado, visao_tempo)

if not serie_tempo.empty:
    grafico_linhas = (
        alt.Chart(serie_tempo)
        .mark_line(point=True, strokeWidth=3)
        .encode(
            x=alt.X("PERIODO:N", title="Período", sort=None),
            y=alt.Y("VALOR_GRAFICO:Q", title="Valor"),
            color=alt.Color(
                "TIPO:N",
                scale=alt.Scale(
                    domain=["Faturado", "Inadimplente"],
                    range=[COR_PRIMARIA, COR_SECUNDARIA]
                ),
                title="Série"
            ),
            tooltip=[
                alt.Tooltip("PERIODO:N", title="Período"),
                alt.Tooltip("TIPO:N", title="Tipo"),
                alt.Tooltip("VALOR_GRAFICO:Q", title="Valor", format=",.0f")
            ]
        )
        .properties(height=380)
    )
    st.altair_chart(grafico_linhas, use_container_width=True)
else:
    st.info("Não há dados suficientes para montar a série temporal.")

st.divider()

# RESUMOS
g1, g2 = st.columns(2)

with g1:
    st.markdown('<div class="secao-titulo">Status</div>', unsafe_allow_html=True)
    resumo_status = df_filtrado["STATUS"].value_counts().reset_index()
    resumo_status.columns = ["STATUS", "QUANTIDADE"]

    grafico_status = (
        alt.Chart(resumo_status)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
        .encode(
            x=alt.X("STATUS:N", title="Status"),
            y=alt.Y("QUANTIDADE:Q", title="Quantidade"),
            color=alt.Color(
                "STATUS:N",
                scale=alt.Scale(
                    domain=["Pago", "Não Pago"],
                    range=[COR_SECUNDARIA, COR_PRIMARIA]
                ),
                legend=None
            ),
            tooltip=["STATUS", "QUANTIDADE"]
        )
        .properties(height=320)
    )
    st.altair_chart(grafico_status, use_container_width=True)

with g2:
    if "EMPRESA" in df_filtrado.columns and "SALDO" in df_filtrado.columns:
        st.markdown('<div class="secao-titulo">Saldo por Empresa</div>', unsafe_allow_html=True)
        saldo_empresa = df_filtrado.groupby("EMPRESA", as_index=False)["SALDO"].sum()

        grafico_empresa = (
            alt.Chart(saldo_empresa)
            .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, color=COR_PRIMARIA)
            .encode(
                x=alt.X("EMPRESA:N", title="Empresa"),
                y=alt.Y("SALDO:Q", title="Saldo"),
                tooltip=[
                    alt.Tooltip("EMPRESA:N", title="Empresa"),
                    alt.Tooltip("SALDO:Q", title="Saldo", format=",.0f")
                ]
            )
            .properties(height=320)
        )
        st.altair_chart(grafico_empresa, use_container_width=True)

st.divider()

if "PESSOA" in df_filtrado.columns and "SALDO" in df_filtrado.columns:
    st.markdown('<div class="secao-titulo">Top devedores</div>', unsafe_allow_html=True)
    top_devedores = (
        df_filtrado.groupby("PESSOA", as_index=False)["SALDO"]
        .sum()
        .sort_values("SALDO", ascending=False)
        .head(15)
    )
    top_devedores["SALDO"] = top_devedores["SALDO"].apply(formatar_moeda)
    st.dataframe(top_devedores, use_container_width=True, hide_index=True)

st.markdown('<div class="secao-titulo">Insight</div>', unsafe_allow_html=True)
if total_saldo > 0:
    st.warning(
        f"Você possui {formatar_moeda(total_saldo)} em aberto. "
        f"Priorize a cobrança dos maiores devedores."
    )
else:
    st.success("Nenhum valor em aberto. Excelente controle financeiro.")

st.markdown('<div class="secao-titulo">Tabela detalhada</div>', unsafe_allow_html=True)
st.dataframe(df_filtrado, use_container_width=True)

csv = df_filtrado.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "Baixar dados filtrados em CSV",
    data=csv,
    file_name="recebimentos_filtrados.csv",
    mime="text/csv"
)