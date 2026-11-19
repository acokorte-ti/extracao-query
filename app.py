import datetime
import os
import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
import oracledb
import pandas as pd
from io import BytesIO

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLSX_MAX_LINHAS = 1_048_575  # limite do Excel (1 linha é o cabeçalho)


# ------------------------------------------------------------
# Datas sem hora. Toda coluna DATE do Oracle chega aqui como
# data+hora e, sem tratamento, aparece como "2026-11-18 00:00:00".
# As colunas listadas em "datas_sem_hora" do relatório (config.yaml)
# saem só como DD/MM/AAAA na tela, no CSV e no XLSX.
# ------------------------------------------------------------
DATA_CURTA_CSV = "%d/%m/%Y"
DATA_CURTA_XLSX = "DD/MM/YYYY"
DATA_CURTA_TELA = "DD/MM/YYYY"


def _e_data(valor):
    """True para data/data+hora preenchida (None e NaT não contam)."""
    return isinstance(valor, datetime.date) and not pd.isna(valor)


def colunas_data_curta(dataframe, nomes):
    """Das colunas pedidas em "datas_sem_hora", devolve as que existem no
    DataFrame e são mesmo de data. Nome que não existe ou coluna que não é
    data é ignorado, em vez de quebrar a página."""
    achadas = []
    for c in nomes or []:
        if c not in dataframe.columns:
            continue
        serie = dataframe[c]
        if pd.api.types.is_datetime64_any_dtype(serie):
            achadas.append(c)
        elif serie.dtype == object:
            # Basta uma data fora do alcance do pandas (ex.: entrega
            # digitada com ano 7026 no Mega) para a coluna inteira chegar
            # como "object", com datetime do Python em vez de datetime64.
            preenchidos = serie.dropna()
            if len(preenchidos) and preenchidos.map(_e_data).all():
                achadas.append(c)
    return achadas


def tela_config_datas(dataframe, nomes):
    """column_config do st.dataframe: mostra essas colunas só com a data.
    O valor continua sendo data (ordenar pela coluna segue funcionando)."""
    return {
        c: st.column_config.DateColumn(format=DATA_CURTA_TELA)
        for c in colunas_data_curta(dataframe, nomes)
    }


def df_to_csv_bytes(dataframe, date_only_cols=None):
    """Converte o DataFrame em bytes de um CSV (separador ;, UTF-8 com BOM).
    As colunas de date_only_cols saem como texto DD/MM/AAAA."""
    cols = colunas_data_curta(dataframe, date_only_cols)
    if cols:
        dataframe = dataframe.copy()
        for c in cols:
            # map (e não .dt.strftime) para valer também quando a coluna
            # chega como "object" (ver colunas_data_curta).
            dataframe[c] = dataframe[c].map(
                lambda v: v.strftime(DATA_CURTA_CSV) if _e_data(v) else None
            )
    return dataframe.to_csv(index=False, sep=";").encode("utf-8-sig")


def df_to_xlsx_bytes(dataframe, table_name=None, date_only_cols=None):
    """Converte o DataFrame em bytes de um arquivo .xlsx (requer openpyxl).
    Se table_name for informado (e houver linhas), os dados viram uma
    Tabela do Excel com esse nome (filtros no cabeçalho).
    As colunas de date_only_cols continuam sendo data de verdade no Excel,
    só que exibidas como DD/MM/AAAA (as outras colunas de data não mudam)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Dados")
        # O pandas grava toda data com o formato "YYYY-MM-DD HH:MM:SS" e,
        # com o engine openpyxl, ignora o date_format/datetime_format do
        # ExcelWriter (testado no pandas 2.3.3 da VM). Por isso o formato
        # é trocado direto nas células das colunas pedidas.
        planilha = writer.sheets["Dados"]
        for c in colunas_data_curta(dataframe, date_only_cols):
            n_col = dataframe.columns.get_loc(c) + 1
            for (celula,) in planilha.iter_rows(
                min_row=2,
                max_row=len(dataframe) + 1,
                min_col=n_col,
                max_col=n_col,
            ):
                if celula.value not in (None, ""):
                    celula.number_format = DATA_CURTA_XLSX
        if table_name and len(dataframe) > 0 and len(dataframe.columns) > 0:
            from openpyxl.worksheet.table import Table, TableStyleInfo
            from openpyxl.utils import get_column_letter

            ws = writer.sheets["Dados"]
            ref = f"A1:{get_column_letter(len(dataframe.columns))}{len(dataframe) + 1}"
            tabela = Table(displayName=table_name, ref=ref)
            tabela.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showRowStripes=True
            )
            ws.add_table(tabela)
    return buf.getvalue()

# ------------------------------------------------------------
# Instant Client (modo thick, exigido pelo Native Network
# Encryption do banco). Ajuste o caminho se mudar de VM.
# ------------------------------------------------------------
oracledb.init_oracle_client(lib_dir="/home/ubuntu/oracle_instant_client/instantclient_23_26")

# ------------------------------------------------------------
# Imagens da marca AçoKorte: 2 arquivos simples, na MESMA pasta
# do app.py na VM (~/apps/portal-dados/), sem subpastas:
#   icon.png -> ícone da aba do navegador
#   logo.png -> logo no cabeçalho
# ------------------------------------------------------------
st.set_page_config(
    page_title="Portal de Dados - AçoKorte",
    page_icon="icon.png",
    layout="wide",
)

# ------------------------------------------------------------
# Ícones vetoriais (SVG) embutidos como texto simples — sem
# depender de arquivos externos nem de base64 de imagem.
# ------------------------------------------------------------
def svg(paths, size=18, color="currentColor", stroke_width=2):
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="{stroke_width}" '
        f'stroke-linecap="round" stroke-linejoin="round">{paths}</svg>'
    )

ICON_BARCHART = svg('<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>')
ICON_HOME = svg('<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M9 22V12h6v10"/>', size=14)
ICON_USER = svg('<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>', size=18, color="#c7cbd1")
ICON_FILTER = svg('<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>', size=20, color="#e0132a")
ICON_BARCHART_WHITE = svg('<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>', size=18, color="#ffffff")
ICON_GEAR = svg('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>', size=18, color="#4a5568")
ICON_SHIELD_WHITE = svg('<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>', size=18, color="#ffffff")
ICON_TREND = svg('<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>', size=18, color="#4a5568")
ICON_LOGO_BADGE = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="46" height="46" viewBox="0 0 46 46">'
    '<rect x="1" y="1" width="44" height="44" rx="22" fill="#e0132a"/>'
    '<path d="M14 23h10m0 0-4-4m4 4-4 4" stroke="#fff" stroke-width="2.4" '
    'stroke-linecap="round" stroke-linejoin="round" fill="none"/>'
    '<path d="M32 23H22m0 0 4-4m-4 4 4 4" stroke="#fff" stroke-width="2.4" '
    'stroke-linecap="round" stroke-linejoin="round" fill="none" opacity="0.55"/>'
    '</svg>'
)
ICON_ARROW_RIGHT = svg('<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>', size=18, color="#fff")
ICON_ARROW_BIG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="220" height="220" viewBox="0 0 100 100">'
    '<path d="M85 50 L45 15 L45 35 L15 35 L15 65 L45 65 L45 85 Z" '
    'fill="none" stroke="#e0132a" stroke-width="3"/>'
    '</svg>'
)

# ------------------------------------------------------------
# CSS de marca AçoKorte: tema escuro, vermelho, cards com
# borda, campos de filtro compactos.
# ------------------------------------------------------------
st.markdown(
    f"""
    <style>
        :root {{
            --ak-red: #e0132a;
            --ak-bg: #f4f5f7;
            --ak-card: #ffffff;
            --ak-border: #e5e8ec;
            --ak-muted: #8a94a3;
            --ak-text: #1f2430;
        }}
        .stApp {{
            background-color: var(--ak-bg);
        }}
        .block-container {{
            padding-top: 1rem;
            max-width: 100%;
        }}
        h1, h2, h3 {{
            color: var(--ak-text) !important;
        }}
        /* esconde a barra padrão do Streamlit (fica sobrepondo o
           nosso cabeçalho quando ele usa margem negativa no topo) */
        header[data-testid="stHeader"] {{
            height: 0px;
            visibility: hidden;
        }}
        /* Força esquema claro. Sem isso, se o navegador/SO do usuário
           estiver em modo escuro, vários componentes internos do
           Streamlit (BaseWeb) caem no visual escuro padrão dele (fundo
           preto, borda preta grossa) mesmo com o CSS de marca por cima,
           porque nosso CSS não alcança todo elemento interno deles. */
        html, body, .stApp, [data-testid="stAppViewContainer"] {{
            color-scheme: light !important;
        }}
        /* ---------- barra superior (logo + tagline + usuário) ---------- */
        .st-key-ak_topbar {{
            background: #ffffff;
            border: 1px solid var(--ak-border);
            border-radius: 10px;
            padding: 10px 18px;
            margin: 0 0 16px 0;
        }}
        .ak-brand-title {{ font-weight: 800; font-size: 1.05rem; color: var(--ak-text); letter-spacing: 0.5px; }}
        .ak-brand-sub {{ font-weight: 600; font-size: 0.72rem; color: var(--ak-muted); letter-spacing: 0.3px; }}
        .st-key-ak_topbar div[data-testid="stPopover"] button {{
            background: #fff !important;
            border: 1px solid var(--ak-border) !important;
            color: var(--ak-text) !important;
            border-radius: 8px !important;
        }}
        /* ---------- banner "Relatórios" ---------- */
        .ak-header {{
            display: flex;
            align-items: center;
            justify-content: flex-start;
            background: linear-gradient(90deg, #ffffff 0%, #fdf1f2 100%);
            border: 1px solid var(--ak-border);
            border-left: 4px solid var(--ak-red);
            padding: 16px 20px;
            margin: 0 0 18px 0;
            border-radius: 10px;
            overflow: visible;
        }}
        .ak-nav {{ display: flex; align-items: center; gap: 12px; }}
        .ak-nav-icon {{ background: var(--ak-red); border-radius: 10px; padding: 9px; display: flex; }}
        .ak-nav-title {{ font-weight: 800; color: var(--ak-text); font-size: 1.02rem; }}
        .ak-nav-sub {{ color: var(--ak-muted); font-size: 0.8rem; }}
        /* ---------- abas de relatório (st.radio horizontal) ---------- */
        div[data-testid="stRadio"] > div[role="radiogroup"] {{
            gap: 30px;
            border-bottom: 1px solid var(--ak-border);
            margin-bottom: 4px;
        }}
        div[data-testid="stRadio"] label {{
            padding: 4px 2px 12px 2px !important;
            margin-bottom: -1px !important;
            border-bottom: 3px solid transparent;
        }}
        div[data-testid="stRadio"] label p {{
            color: var(--ak-muted) !important;
            font-weight: 600 !important;
            font-size: 0.92rem !important;
        }}
        div[data-testid="stRadio"] label:has(input:checked) {{
            border-bottom: 3px solid var(--ak-red);
        }}
        div[data-testid="stRadio"] label:has(input:checked) p {{
            color: var(--ak-red) !important;
        }}
        /* ---------- breadcrumb + título ---------- */
        .ak-title-row {{ display: flex; align-items: center; gap: 12px; margin: 10px 0 2px 0; }}
        .ak-title-icon {{ background: var(--ak-red); border-radius: 10px; padding: 10px; display: flex; }}
        .ak-title-text {{ font-size: 1.6rem; font-weight: 800; color: var(--ak-text); }}
        .ak-breadcrumb {{ display: flex; align-items: center; gap: 6px; color: var(--ak-muted); font-size: 0.82rem; margin: 0 0 18px 54px; }}
        .ak-breadcrumb .current {{ color: var(--ak-red); font-weight: 600; }}
        /* ---------- botões ---------- */
        div[data-testid="stButton"] button {{
            border-radius: 8px !important;
        }}
        div[data-testid="stButton"] button[kind="secondary"] {{
            background-color: #ffffff !important;
            border: 1px solid var(--ak-border) !important;
            color: #4a5568 !important;
        }}
        div[data-testid="stButton"] button[kind="primary"] {{
            background-color: var(--ak-red) !important;
            border: 1px solid var(--ak-red) !important;
            color: #ffffff !important;
            box-shadow: none !important;
        }}
        /* ---------- card de filtros ---------- */
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background: var(--ak-card) !important;
            border: 1px solid var(--ak-border) !important;
            border-radius: 12px !important;
        }}
        .ak-filter-head {{ display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; }}
        .ak-filter-head-left {{ display: flex; align-items: center; gap: 10px; }}
        .ak-filter-title {{ font-weight: 700; color: var(--ak-text); font-size: 1.02rem; }}
        .ak-filter-sub {{ color: var(--ak-muted); font-size: 0.8rem; }}
        /* campos de texto/data: confirmado via DOM ao vivo (chrome-devtools,
           2026-09-11) que nesta versão do Streamlit o elemento que REALMENTE
           desenha a borda/fundo do campo de texto é
           div[data-testid="stTextInputRootElement"] (não existe mais
           data-baseweb="base-input" nessa versão!), e do campo de data é
           div[data-testid="stDateInputField"]. Os campos "número" do
           catálogo (QT, VRUNITARIO etc.) são, na prática, st.text_input
           também (ver kind=="number" no código abaixo), então caem na
           mesma regra de texto. */
        div[data-testid="stTextInputRootElement"],
        div[data-testid="stDateInputField"] {{
            background-color: #f8f9fb !important;
            border: 1px solid var(--ak-border) !important;
            border-radius: 6px !important;
            box-shadow: none !important;
        }}
        div[data-testid="stTextInputRootElement"] input {{
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
            color: var(--ak-text) !important;
            padding: 0.25rem 0.5rem;
            font-size: 0.82rem;
        }}
        /* selectbox: confirmado via DOM que o "controle" visível é o <div>
           filho direto de div.react-aria-ComboBox (não existe mais
           data-baseweb="select" nessa versão). */
        div[data-testid="stSelectbox"] div.react-aria-ComboBox > div {{
            background-color: #f8f9fb !important;
            border: 1px solid var(--ak-border) !important;
            border-radius: 6px !important;
            box-shadow: none !important;
        }}
        div[data-testid="stSelectbox"] input,
        div[data-testid="stSelectbox"] * {{
            color: var(--ak-text) !important;
        }}
        div[data-testid="stTextInput"] label,
        div[data-testid="stDateInput"] label,
        div[data-testid="stSelectbox"] label {{
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--ak-muted) !important;
            text-transform: uppercase;
            margin-bottom: 0.15rem;
        }}
        div[data-testid="stVerticalBlock"] {{ gap: 0.4rem; }}
        /* ---------- rodapé ---------- */
        .ak-footer {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin: 28px 0 4px 0;
            padding-top: 16px;
            border-top: 1px solid var(--ak-border);
            color: var(--ak-muted);
            font-size: 0.75rem;
        }}
        .ak-footer strong {{ color: var(--ak-text); margin-right: 8px; }}
        .ak-footer-right {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 700;
            letter-spacing: 1px;
            color: var(--ak-muted);
            font-size: 0.7rem;
        }}
        .ak-footer-right i {{ width: 26px; height: 2px; background: var(--ak-red); display: inline-block; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# Carrega config.yaml (credenciais de LOGIN DO PORTAL + mapa
# de permissoes + catalogo de relatorios). Isso e diferente da
# credencial do Oracle, que fica fixa mais abaixo.
# ------------------------------------------------------------
with open("config.yaml") as f:
    config = yaml.load(f, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
)

auth_status = st.session_state.get("authentication_status")

# ------------------------------------------------------------
# Tela de login — tema claro AçoKorte
#
# O formulário continua sendo gerado pelo streamlit-authenticator.
# Aqui alteramos apenas o layout externo e a aparência dos widgets
# nativos do Streamlit, preservando a autenticação existente.
# ------------------------------------------------------------
if not auth_status:
    st.markdown(
        """
        <style>
        /* ===== LOGIN — TEMA CLARO ===== */
        .stApp {
            background:
                radial-gradient(circle at 8% 12%, rgba(224,19,42,.055), transparent 28%),
                linear-gradient(135deg, #f8fafc 0%, #eef2f6 100%) !important;
        }

        .block-container {
            max-width: 1180px !important;
            padding-top: 3.5rem !important;
            padding-bottom: 2rem !important;
        }

        /* Remove elementos que poluem a tela de login */
        header[data-testid="stHeader"] {
            height: 0 !important;
            visibility: hidden !important;
        }

        /* Painel esquerdo */
        .ak-login-brand {
            min-height: 620px;
            padding: 55px 35px 45px 45px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            position: relative;
            overflow: hidden;
            border-radius: 24px 0 0 24px;
            background:
                linear-gradient(145deg, #ffffff 0%, #f5f7fa 72%, #edf1f5 100%);
            border: 1px solid #e1e6ec;
            border-right: 0;
        }

        .ak-login-brand::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            width: 6px;
            height: 100%;
            background: #e0132a;
        }

        .ak-login-brand::after {
            content: "";
            position: absolute;
            right: -100px;
            bottom: -120px;
            width: 330px;
            height: 330px;
            border-radius: 50%;
            background: rgba(224,19,42,.045);
        }

        /* seta translúcida decorativa no canto inferior direito */
        .ak-arrow-decor {
            position: absolute;
            right: -30px;
            bottom: 30px;
            width: 220px;
            height: 220px;
            opacity: .5;
            z-index: 0;
            pointer-events: none;
        }

        /* pilha de "chapas de aço" decorativa, feita só com CSS */
        .ak-steel-decor {
            position: relative;
            z-index: 2;
            display: flex;
            flex-direction: column;
            gap: 4px;
            width: 190px;
            margin-bottom: 14px;
        }

        .ak-steel-plate {
            height: 9px;
            border-radius: 3px;
            background: linear-gradient(90deg, #c7ccd3 0%, #eef0f3 45%, #b9bfc7 100%);
            border: 1px solid #aeb4bc;
        }

        .ak-steel-plate:nth-child(1) { width: 100%; }
        .ak-steel-plate:nth-child(2) { width: 92%; }
        .ak-steel-plate:nth-child(3) { width: 100%; }
        .ak-steel-plate:nth-child(4) { width: 85%; }

        .ak-logo-row {
            position: relative;
            z-index: 2;
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .ak-logo-badge {
            flex: 0 0 46px;
            display: flex;
        }

        .ak-logo-text {
            font-size: 1.4rem;
            line-height: 1.1;
            font-weight: 800;
            letter-spacing: .5px;
            color: #111827;
        }

        .ak-logo-text span {
            color: #111827;
        }

        .ak-logo-sub {
            margin-top: 2px;
            font-size: .68rem;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: #8a94a3;
        }

        .ak-login-copy {
            position: relative;
            z-index: 2;
            max-width: 440px;
        }

        .ak-eyebrow {
            display: inline-block;
            margin-bottom: 14px;
            color: #e0132a;
            font-size: .68rem;
            font-weight: 800;
            letter-spacing: 2px;
        }

        .ak-login-copy h1 {
            margin: 0 0 14px;
            color: #151b26 !important;
            font-size: clamp(2.2rem, 4vw, 3.5rem) !important;
            line-height: 1.02;
            letter-spacing: -1.8px;
        }

        .ak-login-copy p {
            margin: 0;
            max-width: 420px;
            color: #667085;
            font-size: .96rem;
            line-height: 1.75;
        }

        .ak-benefits {
            position: relative;
            z-index: 2;
            display: grid;
            gap: 13px;
            margin-top: 34px;
        }

        .ak-benefit {
            display: flex;
            align-items: center;
            gap: 13px;
        }

        .ak-benefit-icon {
            width: 38px;
            height: 38px;
            flex: 0 0 38px;
            display: grid;
            place-items: center;
            border-radius: 10px;
            background: #f1f3f6;
            border: 1px solid #e4e8ed;
        }

        .ak-benefit-icon.ak-benefit-icon--dark {
            background: #1c2230;
            border-color: #1c2230;
        }

        .ak-benefit-icon svg {
            display: block;
        }

        .ak-benefit strong {
            display: block;
            color: #202938;
            font-size: .82rem;
            margin-bottom: 2px;
        }

        .ak-benefit span {
            color: #8a94a3;
            font-size: .72rem;
        }

        .ak-login-bottom {
            position: relative;
            z-index: 2;
            display: flex;
            align-items: center;
            gap: 10px;
            color: #9aa3af;
            font-size: .63rem;
            font-weight: 700;
            letter-spacing: 1.5px;
        }

        .ak-login-bottom i {
            width: 34px;
            height: 2px;
            background: #e0132a;
            display: inline-block;
        }

        /* Painel direito */
        .st-key-ak_login_panel {
            min-height: 620px;
            padding: 50px 60px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            border-radius: 0 24px 24px 0;
            background: #ffffff;
            border: 1px solid #e1e6ec;
            box-shadow: 18px 18px 50px rgba(16,24,40,.07);
        }

        .ak-login-heading {
            margin: 0 auto 28px;
            width: min(100%, 390px);
        }

        .ak-login-heading h2 {
            margin: 0 0 7px;
            color: #151b26 !important;
            font-size: 1.75rem !important;
            font-weight: 800;
            letter-spacing: -.6px;
        }

        .ak-login-heading p {
            margin: 0;
            color: #8993a2;
            font-size: .8rem;
        }

        /* O streamlit-authenticator usa widgets nativos do Streamlit.
           Estes seletores alteram somente a aparência deles. */
        .st-key-ak_login_panel div[data-testid="stTextInput"],
        .st-key-ak_login_panel div[data-testid="stTextInput"] > div {
            width: 100%;
        }

        .st-key-ak_login_panel div[data-testid="stTextInput"] label {
            color: #344054 !important;
            font-size: .73rem !important;
            font-weight: 700 !important;
            margin-bottom: 7px !important;
        }

        .st-key-ak_login_panel div[data-testid="stTextInput"] input {
            height: 48px !important;
            padding: 0 15px !important;
            color: #1d2939 !important;
            background: #ffffff !important;
            border: 1px solid #d9e0e7 !important;
            border-radius: 10px !important;
            font-size: .84rem !important;
            box-shadow: 0 1px 2px rgba(16,24,40,.03) !important;
            transition: border-color .18s ease, box-shadow .18s ease;
        }

        .st-key-ak_login_panel div[data-testid="stTextInput"] input:focus {
            border-color: #e0132a !important;
            box-shadow: 0 0 0 3px rgba(224,19,42,.09) !important;
        }

        .st-key-ak_login_panel div[data-testid="stTextInput"] input::placeholder {
            color: #a6afbb !important;
        }

        /* Botão gerado pelo authenticator */
        .st-key-ak_login_panel div[data-testid="stFormSubmitButton"] button,
        .st-key-ak_login_panel button[kind="primary"] {
            width: 100% !important;
            min-height: 48px !important;
            margin-top: 10px !important;
            border-radius: 10px !important;
            border: 1px solid #e0132a !important;
            background: #e0132a !important;
            color: #ffffff !important;
            font-size: .84rem !important;
            font-weight: 700 !important;
            box-shadow: 0 7px 18px rgba(224,19,42,.18) !important;
            transition: transform .15s ease, box-shadow .15s ease, background .15s ease;
        }

        .st-key-ak_login_panel div[data-testid="stFormSubmitButton"] button:hover,
        .st-key-ak_login_panel button[kind="primary"]:hover {
            background: #c91025 !important;
            border-color: #c91025 !important;
            transform: translateY(-1px);
            box-shadow: 0 10px 22px rgba(224,19,42,.22) !important;
        }

        .st-key-ak_login_panel div[data-testid="stFormSubmitButton"] button:active,
        .st-key-ak_login_panel button[kind="primary"]:active {
            transform: translateY(0);
        }

        /* Esconde detalhes visuais desnecessários dentro do formulário */
        .st-key-ak_login_panel div[data-testid="stForm"] {
            width: min(100%, 390px);
            margin: 0 auto !important;
            padding: 0 !important;
            border: 0 !important;
            background: transparent !important;
            box-shadow: none !important;
        }

        .ak-login-error {
            width: min(100%, 390px);
            margin: 12px auto 0;
        }

        .st-key-ak_login_panel div[data-testid="stCheckbox"] {
            width: min(100%, 390px);
            margin: 4px auto 0;
        }

        .st-key-ak_login_panel div[data-testid="stCheckbox"] label p {
            color: #344054 !important;
            font-size: .78rem !important;
        }

        .ak-login-footer {
            width: min(100%, 390px);
            margin: 28px auto 0;
            padding-top: 20px;
            border-top: 1px solid #edf0f3;
            text-align: center;
            color: #98a2b3;
            font-size: .65rem;
        }

        .ak-login-footer strong {
            display: block;
            margin-bottom: 3px;
            color: #667085;
            font-size: .68rem;
        }

        /* Remove excesso de espaços entre widgets do login */
        .st-key-ak_login_panel [data-testid="stVerticalBlock"] {
            gap: .55rem;
        }

        @media (max-width: 900px) {
            .block-container {
                padding: 1.5rem 1rem 2rem !important;
            }

            .ak-login-brand {
                min-height: 360px;
                padding: 35px 30px;
                border-right: 1px solid #e1e6ec;
                border-radius: 20px 20px 0 0;
            }

            .ak-login-copy h1 {
                font-size: 2.5rem !important;
            }

            .st-key-ak_login_panel {
                min-height: auto;
                padding: 42px 30px;
                border-radius: 0 0 20px 20px;
                box-shadow: 0 18px 45px rgba(16,24,40,.07);
            }
        }

        @media (max-width: 620px) {
            .ak-login-brand {
                min-height: 300px;
                padding: 30px 24px;
            }

            .ak-benefits {
                display: none;
            }

            .ak-login-copy h1 {
                font-size: 2.2rem !important;
            }

            .st-key-ak_login_panel {
                padding: 36px 22px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Layout externo. O formulário continua sendo 100% do authenticator.
    brand_col, form_col = st.columns([1.05, 0.95], gap="small")

    with brand_col:
        st.markdown(
            '<div class="ak-login-brand">'
            f'<div class="ak-arrow-decor">{ICON_ARROW_BIG}</div>'
            '<div class="ak-logo-row">'
            f'<div class="ak-logo-badge">{ICON_LOGO_BADGE}</div>'
            '<div>'
            '<div class="ak-logo-text">AÇOKORTE</div>'
            '<div class="ak-logo-sub">Força em cada corte</div>'
            '</div>'
            '</div>'
            '<div class="ak-login-copy">'
            '<h1>Bem-vindo!</h1>'
            '<p>Acesse o sistema e gerencie seus processos '
            'com mais agilidade e eficiência.</p>'
            '<div class="ak-benefits">'
            '<div class="ak-benefit">'
            f'<div class="ak-benefit-icon">{ICON_GEAR}</div>'
            '<div><strong>Mais controle</strong>'
            '<span>Processos organizados</span></div>'
            '</div>'
            '<div class="ak-benefit">'
            f'<div class="ak-benefit-icon">{ICON_TREND}</div>'
            '<div><strong>Mais produtividade</strong>'
            '<span>Informações em tempo real</span></div>'
            '</div>'
            '<div class="ak-benefit">'
            f'<div class="ak-benefit-icon ak-benefit-icon--dark">{ICON_SHIELD_WHITE}</div>'
            '<div><strong>Mais segurança</strong>'
            '<span>Seus dados protegidos</span></div>'
            '</div>'
            '</div>'
            '</div>'
            '<div>'
            '<div class="ak-steel-decor">'
            '<div class="ak-steel-plate"></div>'
            '<div class="ak-steel-plate"></div>'
            '<div class="ak-steel-plate"></div>'
            '<div class="ak-steel-plate"></div>'
            '</div>'
            '<div class="ak-login-bottom">AÇO QUE MOVE RESULTADOS<i></i></div>'
            '</div>'
            '</div>',
            unsafe_allow_html=True,
        )

    with form_col:
        # Usamos um st.container(key=...) de verdade (em vez de tentar abrir
        # e fechar uma <div> em chamadas separadas de st.markdown — isso NÃO
        # funciona no Streamlit, cada chamada isola o próprio HTML). O
        # container real gera uma div com a classe "st-key-ak_login_panel"
        # que dá pra estilizar via CSS de forma confiável, contendo todos os
        # elementos (título, campos, checkbox, botão, rodapé) de verdade.
        with st.container(key="ak_login_panel"):
            st.markdown(
                '<div class="ak-login-heading">'
                '<h2>Login</h2>'
                '<p>Entre com suas credenciais para continuar</p>'
                '</div>',
                unsafe_allow_html=True,
            )

            # Campos em português: o streamlit-authenticator (0.4.x) aceita
            # um dict "fields" pra traduzir os rótulos e o texto do botão.
            # Se essa versão instalada na VM não aceitar (TypeError), cai no
            # login() padrão em inglês em vez de quebrar a página.
            try:
                authenticator.login(
                    location="main",
                    fields={
                        "Form name": "",
                        "Username": "Usuário",
                        "Password": "Senha",
                        "Login": "Entrar  →",
                    },
                )
            except TypeError:
                authenticator.login(location="main")

            auth_status = st.session_state.get("authentication_status")

            st.checkbox("Lembrar de mim", key="lembrar_login")

            if auth_status is False:
                st.error("Usuário ou senha incorretos.")

            st.markdown(
                '<div class="ak-login-footer">'
                '<strong>AçoKORTE Sistemas</strong>'
                "Todos os direitos reservados"
                "</div>",
                unsafe_allow_html=True,
            )

    st.stop()

name = st.session_state.get("name")
username = st.session_state.get("username")

# ------------------------------------------------------------
# Barra superior: logo + tagline à esquerda, menu do usuário
# (popover com o botão Sair dentro) à direita — tudo dentro de
# um único st.container(key=...) pra formar uma barra branca real
# (ver GOTCHA 2 no histórico: não dá pra "abrir/fechar" uma <div>
# em chamadas separadas de st.markdown).
# ------------------------------------------------------------
with st.container(key="ak_topbar"):
    logo_col, tagline_col, menu_col = st.columns([1, 3, 1])
    with logo_col:
        st.image("logo.png", width=140)
    with tagline_col:
        st.markdown(
            '<div style="padding-top:10px;" class="ak-brand-sub">Força em cada corte</div>',
            unsafe_allow_html=True,
        )
    with menu_col:
        with st.popover(f"👤 {name}", use_container_width=True):
            st.write(f"Conectado como **{name}**")
            authenticator.logout(button_name="🚪 Sair", location="main")

st.markdown(
    f"""
    <div class="ak-header">
        <div class="ak-nav">
            <div class="ak-nav-icon">{ICON_BARCHART_WHITE}</div>
            <div>
                <div class="ak-nav-title">Relatórios</div>
                <div class="ak-nav-sub">Análises e consultas do sistema</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

allowed_report_keys = config.get("user_reports", {}).get(username, [])
reports_catalog = config.get("reports", {})

if not allowed_report_keys:
    st.warning("Você não tem nenhum relatório liberado. Fale com o administrador do portal.")
    st.stop()

report_options = {
    key: reports_catalog[key]["label"]
    for key in allowed_report_keys
    if key in reports_catalog
}

if len(report_options) > 1:
    selected_key = st.radio(
        "Relatório",
        options=list(report_options.keys()),
        format_func=lambda k: report_options[k],
        horizontal=True,
        label_visibility="collapsed",
    )
else:
    selected_key = list(report_options.keys())[0]

view_name = reports_catalog[selected_key]["view"]
report_label = report_options[selected_key]
# Nome-base dos arquivos baixados. Vem de "arquivo" no config.yaml;
# se o relatório não tiver, usa a própria chave (comportamento antigo).
file_base = reports_catalog[selected_key].get("arquivo") or selected_key
# Nome da Tabela do Excel dentro do XLSX ("tabela" no config.yaml).
# Se não houver, o XLSX sai com os dados simples, sem Tabela.
xlsx_table_name = reports_catalog[selected_key].get("tabela") or None
# Colunas de data que saem só como DD/MM/AAAA, sem hora
# ("datas_sem_hora" no config.yaml). Sem a chave, nada muda.
date_only_cols = reports_catalog[selected_key].get("datas_sem_hora") or []

# ------------------------------------------------------------
# Título com ícone + breadcrumb (decorativo, dinâmico).
# ------------------------------------------------------------
st.markdown(
    f"""
    <div class="ak-title-row">
        <div class="ak-title-icon">{ICON_BARCHART_WHITE}</div>
        <div class="ak-title-text">{report_label}</div>
    </div>
    <div class="ak-breadcrumb">
        {ICON_HOME} Relatórios &nbsp;›&nbsp; <span class="current">{report_label}</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# Conexão Oracle - credencial FIXA da aplicação (não é o login
# do portal). Lida de variáveis de ambiente (ORACLE_USER e
# ORACLE_PASSWORD), definidas em ~/apps/portal-dados/.env na VM
# (arquivo separado, nunca sobrescrito por atualizações do app.py).
# ------------------------------------------------------------
ORACLE_USER = os.environ.get("ORACLE_USER")
ORACLE_PASSWORD = os.environ.get("ORACLE_PASSWORD")

if not ORACLE_USER or not ORACLE_PASSWORD:
    st.error(
        "Credenciais do Oracle não encontradas nas variáveis de ambiente "
        "ORACLE_USER / ORACLE_PASSWORD. Verifique o arquivo .env na VM "
        "(ou o EnvironmentFile do serviço systemd)."
    )
    st.stop()

@st.cache_resource
def get_connection():
    return oracledb.connect(
        user=ORACLE_USER,
        password=ORACLE_PASSWORD,
        dsn="10.0.0.235:1521/dbmegaxt_pdb1.sub08032113590.vcnmegaxt.oraclevcn.com",
    )

conn = get_connection()
cursor = conn.cursor()
cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY'")

# ------------------------------------------------------------
# Descobre as colunas da view selecionada e classifica cada uma
# como texto / data / numero, para gerar os filtros automaticamente.
# ------------------------------------------------------------
@st.cache_data(ttl=3600)
def get_columns(_conn, view):
    c = _conn.cursor()
    c.execute(f"SELECT * FROM {view} WHERE 1=0")
    cols = []
    for col in c.description:
        col_name = col[0]
        type_code = col[1]
        if type_code in (oracledb.DB_TYPE_DATE, oracledb.DB_TYPE_TIMESTAMP):
            kind = "date"
        elif type_code in (
            oracledb.DB_TYPE_NUMBER,
            oracledb.DB_TYPE_BINARY_INTEGER,
            oracledb.DB_TYPE_BINARY_DOUBLE,
            oracledb.DB_TYPE_BINARY_FLOAT,
        ):
            kind = "number"
        else:
            kind = "text"
        cols.append((col_name, kind))
    return cols

columns = get_columns(conn, view_name)

FILTER_PREFIXES = ("txt_", "dt_", "num_")

# Contador de "geração" dos campos de filtro, por relatório. Ele entra
# na key de cada campo (txt_/dt_/num_...). Apagar o valor do session_state
# NÃO é suficiente para limpar visualmente um widget já montado no
# navegador (o Streamlit mantém o valor exibido). A forma confiável é
# mudar a key do widget, forçando ele a ser recriado do zero (vazio).
if "filtro_gen" not in st.session_state:
    st.session_state["filtro_gen"] = {}

def limpar_filtros_da_view(view_key):
    """Callback usado no on_click do botão Limpar filtros. Incrementa a
    geração dos campos dessa view (o que troca a key de todos eles no
    próximo rerun) e limpa o resultado da última busca."""
    st.session_state["filtro_gen"][view_key] = st.session_state["filtro_gen"].get(view_key, 0) + 1
    prefixes = tuple(f"{p}{view_key}_" for p in FILTER_PREFIXES)
    for k in list(st.session_state.keys()):
        if k.startswith(prefixes):
            del st.session_state[k]
    for k in ("last_view", "last_where", "last_params", "last_qtd"):
        st.session_state.pop(k, None)

gen = st.session_state["filtro_gen"].get(selected_key, 0)

# ------------------------------------------------------------
# Card de filtros (usa st.container(border=True) para ter a
# borda/fundo do card automaticamente, e tudo dentro dele fica
# visualmente contido, incluindo o cabeçalho, a grade de campos
# e a linha de "Linhas na tela" + Buscar).
# ------------------------------------------------------------
with st.container(border=True):
    head_left, head_right = st.columns([5, 1])
    with head_left:
        st.markdown(
            f"""
            <div class="ak-filter-head-left">
                {ICON_FILTER}
                <div>
                    <div class="ak-filter-title">Filtros de Pesquisa</div>
                    <div class="ak-filter-sub">Preencha os campos desejados para buscar as informações</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with head_right:
        st.write("")
        st.button(
            "🗑️ Limpar filtros",
            key=f"limpar_{selected_key}",
            on_click=limpar_filtros_da_view,
            args=(selected_key,),
        )
    filtros = {}
    cols_per_row = 5
    rows = [columns[i : i + cols_per_row] for i in range(0, len(columns), cols_per_row)]

    for row in rows:
        st_cols = st.columns(len(row))
        for st_col, (col_name, kind) in zip(st_cols, row):
            with st_col:
                if kind == "text":
                    val = st.text_input(col_name, key=f"txt_{selected_key}_{col_name}_{gen}")
                    if val:
                        filtros[col_name] = ("text", val)
                elif kind == "date":
                    sub1, sub2 = st.columns(2)
                    d_ini = sub1.date_input(
                        f"{col_name} de", value=None, key=f"dt_{selected_key}_{col_name}_ini_{gen}"
                    )
                    d_fim = sub2.date_input(
                        f"{col_name} até", value=None, key=f"dt_{selected_key}_{col_name}_fim_{gen}"
                    )
                    if d_ini or d_fim:
                        filtros[col_name] = ("date", d_ini, d_fim)
                elif kind == "number":
                    val = st.text_input(col_name, key=f"num_{selected_key}_{col_name}_{gen}")
                    if val:
                        filtros[col_name] = ("number", val)

    st.write("")
    opcoes_linhas = [10, 50, 100, 500, 1000, 2000, 5000]
    col_qtd, col_info = st.columns([1, 3])
    with col_qtd:
        qtd_linhas = st.selectbox("Linhas na tela", opcoes_linhas, index=2)
        buscar = st.button(
            "🔍 Buscar", type="primary", key=f"buscar_{selected_key}", use_container_width=True
        )
    with col_info:
        st.write("")
        if not filtros:
            st.caption("Preencha os filtros desejados (opcional) e clique em Buscar.")

if buscar:
    condicoes = []
    params = {}
    idx = 0

    for col_name, info in filtros.items():
        kind = info[0]

        if kind == "text":
            idx += 1
            p = f"p{idx}"
            condicoes.append(f"UPPER({col_name}) LIKE UPPER(:{p})")
            params[p] = f"%{info[1]}%"

        elif kind == "date":
            _, d_ini, d_fim = info
            if d_ini:
                idx += 1
                p = f"p{idx}"
                condicoes.append(f"{col_name} >= :{p}")
                params[p] = d_ini
            if d_fim:
                idx += 1
                p = f"p{idx}"
                condicoes.append(f"{col_name} <= :{p}")
                params[p] = d_fim

        elif kind == "number":
            _, n_val = info
            try:
                idx += 1
                p = f"p{idx}"
                condicoes.append(f"{col_name} = :{p}")
                params[p] = float(n_val)
            except ValueError:
                st.error(f"Valor inválido no filtro de '{col_name}'. Use apenas números.")
                st.stop()

    where_clause = " WHERE " + " AND ".join(condicoes) if condicoes else ""

    st.session_state["last_view"] = view_name
    st.session_state["last_where"] = where_clause
    st.session_state["last_params"] = params
    st.session_state["last_qtd"] = qtd_linhas

# ------------------------------------------------------------
# Resultados (preview limitado à quantidade escolhida) sempre
# que houver uma busca feita para a view atualmente selecionada.
# ------------------------------------------------------------
if st.session_state.get("last_view") == view_name:
    where_clause = st.session_state["last_where"]
    params = st.session_state["last_params"]
    qtd = st.session_state.get("last_qtd", 100)

    # Sem ORDER BY o Oracle devolve as linhas em qualquer ordem interna
    # dele, então limitar "as primeiras N linhas" podia trazer registros
    # antigos ao acaso. Ordenamos pelo primeiro campo de data da view
    # (do mais recente pro mais antigo) pra sempre priorizar os dados
    # mais atuais quando há limite de linhas na tela.
    # NULLS LAST: no Oracle a ordem decrescente põe os vazios na frente.
    # Sem isso, as linhas sem data (ex.: OPs sem pedido no PCP) ocupavam
    # as primeiras posições e escondiam os registros recentes.
    date_cols = [c for c, kind in columns if kind == "date"]
    order_by = f" ORDER BY {date_cols[0]} DESC NULLS LAST" if date_cols else ""

    query = f"SELECT * FROM {view_name}{where_clause}{order_by} FETCH FIRST {qtd} ROWS ONLY"
    df = pd.read_sql(query, conn, params=params)
    st.success(f"{len(df)} linha(s) encontrada(s) (exibindo no máximo {qtd}).")
    st.dataframe(
        df,
        use_container_width=True,
        column_config=tela_config_datas(df, date_only_cols),
    )

    csv = df_to_csv_bytes(df, date_only_cols)
    dl_csv, dl_xlsx, _dl_fill = st.columns([1, 1, 3])
    with dl_csv:
        st.download_button(
            f"Exportar CSV (até {qtd} linhas)",
            csv,
            f"{file_base}.csv",
            "text/csv",
            use_container_width=True,
        )
    with dl_xlsx:
        st.download_button(
            f"Exportar XLSX (até {qtd} linhas)",
            df_to_xlsx_bytes(df, xlsx_table_name, date_only_cols),
            f"{file_base}.xlsx",
            XLSX_MIME,
            use_container_width=True,
        )

    st.divider()
    st.caption(
        "Precisa de todas as linhas, sem limite? O botão abaixo busca o "
        "resultado completo — pode demorar mais dependendo da quantidade "
        "de linhas."
    )
    if st.button("Gerar arquivo completo (sem limite)"):
        full_query = f"SELECT * FROM {view_name}{where_clause}{order_by}"
        with st.spinner("Buscando todas as linhas, isso pode levar um tempo..."):
            df_full = pd.read_sql(full_query, conn, params=params)
        st.success(f"{len(df_full)} linha(s) no total.")
        csv_full = df_to_csv_bytes(df_full, date_only_cols)
        full_csv_col, full_xlsx_col, _full_fill = st.columns([1, 1, 3])
        with full_csv_col:
            st.download_button(
                "Baixar CSV completo",
                csv_full,
                f"{file_base}_completo.csv",
                "text/csv",
                use_container_width=True,
            )
        with full_xlsx_col:
            if len(df_full) > XLSX_MAX_LINHAS:
                st.warning(
                    f"Resultado com {len(df_full)} linhas excede o limite do Excel "
                    f"({XLSX_MAX_LINHAS}). Use o CSV ou refine os filtros."
                )
            else:
                with st.spinner("Montando o XLSX..."):
                    xlsx_full = df_to_xlsx_bytes(
                        df_full, xlsx_table_name, date_only_cols
                    )
                st.download_button(
                    "Baixar XLSX completo",
                    xlsx_full,
                    f"{file_base}_completo.xlsx",
                    XLSX_MIME,
                    use_container_width=True,
                )
else:
    st.info("Preencha os filtros desejados (opcional) e clique em Buscar.")

# ------------------------------------------------------------
# Rodapé (decorativo).
# ------------------------------------------------------------
st.markdown(
    """
    <div class="ak-footer">
        <div><strong>AçoKORTE Sistemas</strong>Todos os direitos reservados</div>
        <div class="ak-footer-right">AÇO QUE MOVE RESULTADOS<i></i></div>
    </div>
    """,
    unsafe_allow_html=True,
)
