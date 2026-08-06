import io
import math
import os
from urllib.parse import quote

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

LOCAL_EXCEL_FALLBACK = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..",
    "Controle de Relatório dos Instrumentos", "01_ARQUIVO_ATUAL",
    "CONTROLE_DOCUMENTAL_INSTRUMENTACAO_ATUAL.xlsx",
)
SUPABASE_BUCKET = "gplan-data"
SUPABASE_FILE_PATH = "CONTROLE_DOCUMENTAL_INSTRUMENTACAO_ATUAL.xlsx"
PAGE_SIZE = 100
# O Render roda em UTC; sem converter, o cabecalho mostrava 3h a mais.
BR_TZ = "America/Sao_Paulo"


def render_html(html: str):
    st.markdown("\n".join(line.strip() for line in html.strip().split("\n")), unsafe_allow_html=True)


STATUS_DISPLAY_MAP = {
    "NAO POSTADO": "Não postado",
    "Sem Comentários": "Sem comentários",
    "Com Comentários": "Com comentários",
    "Para Construção": "Para construção",
    "Recusado": "Recusado",
    "Em Análise": "Em análise",
    "Cancelado": "Cancelado",
    "Certificado": "Certificado",
    "Pendente Certificação": "Pendente certificação",
    "Sem Workflow": "Sem workflow",
    "Emitido para Comentários": "Emitido para comentários",
    "Conforme Construído": "Conforme construído",
    "Aceito Com Comentários": "Aceito com comentários",
    "Em Workflow": "Em workflow",
    "Para Informação": "Para informação",
    "Para Compra": "Para compra",
}

STATUS_COLOR_MAP = {
    "Não postado": "#7c8aa8",
    "Sem comentários": "#2dd4bf",
    "Com comentários": "#5b8def",
    "Para construção": "#fbbf24",
    "Recusado": "#f87171",
    "Em análise": "#9d6bff",
    "Cancelado": "#4b5468",
}
DEFAULT_STATUS_COLORS = ["#5b8def", "#2dd4bf", "#fbbf24", "#f87171", "#9d6bff", "#7c8aa8", "#4b5468", "#34d399"]

REPORT_ROWS = [
    ("RIR instrumentos", "RIR", "BASE: RIR obrigatorio para todos os TAGs", False),
    ("RIR cabos", "RIR", "CONDICIONAL: RIR de cabo por TAG", False),
    ("CCP", "CCP", None, False),
    ("RTFCJI", "RTFCJI", None, False),
    ("RIMITPI", "RIMITPI", None, False),
    ("RIFMI", "RIFMI", None, False),
    ("RIMTU", "RIMTU", None, False),
    ("RILTCI", "RILTCI", None, False),
    ("RIMJBI", "RIMJBI", None, False),
    ("CTECRI", "CTECRI", None, False),
    ("RIMII eletroduto", "RIMII", "CONDICIONAL: infra eletroduto por planta", True),
    ("RIMII bandeja", "RIMII", "CONDICIONAL: infra bandeja por planta", True),
    ("RIMSI suporte", "RIMSI", "CONDICIONAL: infra suporte por planta", True),
]


def sentence_case(value: object) -> str:
    text = str(value) if value is not None else ""
    if not text:
        return text
    mapped = STATUS_DISPLAY_MAP.get(text)
    if mapped:
        return mapped
    return text[0].upper() + text[1:].lower() if len(text) > 1 else text.upper()


def format_missing(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    return str(value)


def format_date(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    if isinstance(value, str):
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    else:
        parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return str(value)
    return parsed.strftime("%d/%m/%Y")


def format_date_column(series: pd.Series) -> pd.Series:
    """Vectorized equivalent of format_date, ~100x faster on large columns
    than calling pd.to_datetime() once per row via .apply()."""
    parsed = pd.to_datetime(series, dayfirst=True, errors="coerce")
    formatted = parsed.dt.strftime("%d/%m/%Y")
    return formatted.where(parsed.notna(), "—")


def search_any_column(df: pd.DataFrame, text: str) -> pd.Series:
    """Column-wise (vectorized) search across every column, instead of
    row-wise .apply() which is orders of magnitude slower on large tables."""
    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        mask = mask | df[col].astype(str).str.contains(text, case=False, na=False)
    return mask


def get_supabase_client():
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_KEY")
    if not url or not key:
        try:
            url = url or st.secrets.get("SUPABASE_URL")
            key = key or st.secrets.get("SUPABASE_KEY")
        except Exception:
            pass
    if not url or not key:
        return None
    from supabase import create_client
    return create_client(url, key)


@st.cache_data(ttl=60)
def get_source_cache_key() -> str:
    client = get_supabase_client()
    if client is not None:
        files = client.storage.from_(SUPABASE_BUCKET).list()
        for f in files:
            if f["name"] == SUPABASE_FILE_PATH:
                return f.get("updated_at", "") or f.get("id", "")
        return "missing"
    return str(os.path.getmtime(LOCAL_EXCEL_FALLBACK)) if os.path.exists(LOCAL_EXCEL_FALLBACK) else "missing"


@st.cache_data(show_spinner="Carregando planilha...")
def load_data(cache_key: str):
    client = get_supabase_client()
    if client is not None:
        file_bytes = client.storage.from_(SUPABASE_BUCKET).download(SUPABASE_FILE_PATH)
        source = io.BytesIO(file_bytes)
    elif os.path.exists(LOCAL_EXCEL_FALLBACK):
        source = LOCAL_EXCEL_FALLBACK
    else:
        st.error(
            "Não encontrei a planilha. Configure SUPABASE_URL e SUPABASE_KEY em "
            ".streamlit/secrets.toml, ou rode localmente com a base de dados presente."
        )
        st.stop()

    excel_file = pd.ExcelFile(source)
    tags = pd.read_excel(excel_file, sheet_name="01_BASE_TAGS")
    cabos = pd.read_excel(excel_file, sheet_name="02_BASE_CABOS")
    tubing = pd.read_excel(excel_file, sheet_name="03_BASE_TUBING")
    sigem = pd.read_excel(excel_file, sheet_name="04_BASE_SIGEM")
    resumo = pd.read_excel(excel_file, sheet_name="07_TAG_RESUMO")
    esperados = pd.read_excel(excel_file, sheet_name="08_RELATORIOS_ESPERADOS")
    return tags, cabos, tubing, sigem, resumo, esperados


def paginate(df: pd.DataFrame, key: str, search_signature: str) -> pd.DataFrame:
    total_rows = len(df)
    total_pages = max(1, math.ceil(total_rows / PAGE_SIZE))

    page_key = f"gplan_page_{key}"
    sig_key = f"gplan_page_sig_{key}"
    if st.session_state.get(sig_key) != search_signature:
        st.session_state[sig_key] = search_signature
        st.session_state[page_key] = 1
    current_page = min(max(1, st.session_state.get(page_key, 1)), total_pages)

    start = (current_page - 1) * PAGE_SIZE
    end = min(start + PAGE_SIZE, total_rows)

    col_prev, col_mid, col_next = st.columns([1, 3, 1])
    with col_prev:
        if st.button("← Anterior", key=f"prev_{key}", disabled=current_page <= 1, use_container_width=True):
            st.session_state[page_key] = current_page - 1
            st.rerun()
    with col_mid:
        range_txt = f"{start + 1}–{end}" if total_rows else "0"
        render_html(
            f'<div style="text-align:center; color:#9aa4bc; font-size:12.5px; padding-top:9px;">'
            f"Exibindo {range_txt} de {total_rows:,} · Página {current_page} de {total_pages}"
            f"</div>".replace(",", ".")
        )
    with col_next:
        if st.button("Próxima →", key=f"next_{key}", disabled=current_page >= total_pages, use_container_width=True):
            st.session_state[page_key] = current_page + 1
            st.rerun()

    return df.iloc[start:end]


REPORT_LABELS = [r[0] for r in REPORT_ROWS]
REPORT_BY_LABEL = {r[0]: (r[1], r[2]) for r in REPORT_ROWS}


def consume_url_filters(esperados: pd.DataFrame):
    """Aplica os filtros que chegam do Dashboard via ?rel= / ?status=.

    Usa um token do proprio query string para aplicar UMA vez: sem isso, todo
    rerun sobrescreveria o que o usuario acabou de escolher nos multiselects.
    Nao limpa o query param porque escrever nele dispara outro rerun.
    """
    rel = [v for v in st.query_params.get_all("rel") if v in REPORT_BY_LABEL]
    valid_status = {sentence_case(s) for s in esperados["STATUS_SIGEM"].dropna().unique()}
    sts = [v for v in st.query_params.get_all("status") if v in valid_status]
    if not rel and not sts:
        return

    token = f"{sorted(rel)}|{sorted(sts)}"
    if st.session_state.get("_flt_token") == token:
        return
    st.session_state["_flt_token"] = token
    st.session_state["flt_rel"] = rel
    st.session_state["flt_status"] = sts


def filter_by_report_labels(df: pd.DataFrame, labels: list) -> pd.DataFrame:
    """Filtra por rotulo de relatorio respeitando a mesma logica do count_rows:
    'RIR instrumentos' e 'RIR cabos' compartilham RELATORIO='RIR' e so se
    distinguem pela ORIGEM_REGRA."""
    if not labels:
        return df
    mask = pd.Series(False, index=df.index)
    for label in labels:
        if label not in REPORT_BY_LABEL:
            continue
        report, origin = REPORT_BY_LABEL[label]
        m = df["RELATORIO"] == report
        if origin is not None:
            m &= df["ORIGEM_REGRA"] == origin
        mask |= m
    return df[mask]


def count_rows(df: pd.DataFrame, report: str, origin: str | None, emitted: bool, unique_doc: bool) -> int:
    subset = df[df["RELATORIO"] == report]
    if origin is not None:
        subset = subset[subset["ORIGEM_REGRA"] == origin]
    if emitted:
        subset = subset[subset["EXISTE_NO_SIGEM"] == "SIM"]
    if unique_doc:
        return int(subset["DOCUMENTO_ESPERADO"].nunique())
    return int(len(subset))


def inject_css():
    render_html(
        """
        <style>
        :root {
          --dark-bg: #0a0e1a;
          --dark-card: #12172a;
          --dark-card-2: #171d33;
          --border-color: rgba(255,255,255,0.06);
          --border-strong: rgba(255,255,255,0.12);
          --accent-blue: #5b8def;
          --accent-purple: #9d6bff;
          --accent-green: #34d399;
          --accent-red: #f87171;
          --accent-amber: #fbbf24;
          --accent-teal: #2dd4bf;
          --text-1: #f4f6fb;
          --text-2: #9aa4bc;
          --text-3: #6b7590;
        }

        .stApp { background: var(--dark-bg); }
        /* Streamlit usa 96px topo / 80px lateral / 160px rodape por padrao,
           o que deixa uma faixa vazia enorme no fim da pagina e desperdica largura. */
        [data-testid="stMainBlockContainer"], .block-container {
          padding: 32px 44px 48px !important;
          max-width: 1600px !important;
        }
        [data-testid="stHeader"] { background: transparent !important; }
        section[data-testid="stSidebar"] {
          background: linear-gradient(180deg, #0d1224 0%, #0a0e1a 100%);
          border-right: 1px solid var(--border-color);
        }
        /* O st.navigation sempre injeta o menu antes do conteudo do usuario,
           entao a marca caia embaixo. Reordena via flex para o topo. */
        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
          display: flex !important; flex-direction: column;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarHeader"] { order: 0; }
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
          order: 1; padding-top: 0 !important; padding-bottom: 0 !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] { order: 2; }
        section[data-testid="stSidebar"] [data-testid="stSidebarNavSeparator"] { display: none; }

        .gplan-brand {
          display:flex; align-items:center; gap:11px; padding: 0 4px 32px;
          border-bottom: 1px solid var(--border-color);
        }
        .gplan-brand-mark { width:34px; height:34px; flex-shrink:0; }
        .gplan-brand-name { font-size:16px; font-weight:800; color:var(--text-1); letter-spacing:-0.3px; line-height:1.1; }
        .gplan-brand-sub { font-size:11px; color:var(--text-3); margin-top:3px; }

        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a,
        section[data-testid="stSidebar"] nav a {
          border-radius: 10px !important;
          color: var(--text-2) !important;
          font-size: 13.5px !important;
          font-weight: 500 !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover,
        section[data-testid="stSidebar"] nav a:hover {
          background: rgba(255,255,255,0.05) !important;
          color: var(--text-1) !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"],
        section[data-testid="stSidebar"] nav a[aria-current="page"] {
          background: linear-gradient(135deg, rgba(91,141,239,0.18), rgba(157,107,255,0.12)) !important;
          color: #ffffff !important;
          box-shadow: inset 0 0 0 1px rgba(91,141,239,0.3);
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] span,
        section[data-testid="stSidebar"] nav a[aria-current="page"] span {
          color: #ffffff !important;
        }

        .gplan-header { display:flex; justify-content:space-between; align-items:center; margin-bottom: 20px; }
        .gplan-header h1 { font-size: 24px; font-weight: 700; color: var(--text-1); margin: 0; letter-spacing: -0.4px; }
        .gplan-updated {
          font-size: 12.5px; color: var(--text-2); display:flex; align-items:center; gap: 8px;
          background: var(--dark-card); border: 1px solid var(--border-color); padding: 8px 14px; border-radius: 100px;
        }
        .gplan-updated .dot { width:6px; height:6px; border-radius:50%; background: var(--accent-green); box-shadow: 0 0 8px var(--accent-green); }
        .gplan-count-pill {
          font-size: 12.5px; color: var(--text-2); font-weight: 600;
          background: var(--dark-card); border: 1px solid var(--border-color); padding: 8px 14px; border-radius: 100px;
        }
        .gplan-count-pill strong { color: var(--text-1); }

        .kpi-card {
          background: var(--dark-card); border: 1px solid var(--border-color); border-radius: 16px;
          padding: 20px 22px; position: relative; overflow: hidden;
        }
        .kpi-card::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; }
        .kpi-top { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom: 12px; }
        .kpi-icon { width:32px; height:32px; border-radius:9px; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
        .kpi-icon svg { width:16px; height:16px; }
        .kpi-label { font-size: 11.5px; color: var(--text-2); font-weight:600; letter-spacing:0.2px; }
        .kpi-value { font-size: 30px; font-weight: 800; color: var(--text-1); letter-spacing:-0.8px; margin-bottom:4px; }
        .kpi-sub { font-size: 12px; color: var(--text-3); margin-bottom: 14px; }
        .kpi-progress-row { display:flex; align-items:center; gap:10px; }
        .kpi-track { flex:1; height:5px; background: rgba(255,255,255,0.06); border-radius:3px; overflow:hidden; }
        .kpi-fill { height:100%; border-radius:3px; }
        .kpi-pct { font-size:12px; font-weight:700; color: var(--text-1); min-width:38px; text-align:right; }

        .gplan-panel { background: var(--dark-card); border: 1px solid var(--border-color); border-radius: 16px; padding: 24px; height: 100%; }
        .gplan-panel-title { font-size: 14.5px; font-weight: 700; color: var(--text-1); margin-bottom: 18px; }

        /* Streamlit aplica bordas em toda celula de <table>; o mockup so tem
           linhas horizontais, entao zeramos e recolocamos apenas o border-bottom. */
        .gtbl, .gtbl th, .gtbl td { border:none !important; background:transparent !important; }
        .gtbl { width:100%; border-collapse:collapse; }
        .gtbl th {
          text-align:left; font-size:10.5px; font-weight:600; color:var(--text-3) !important;
          text-transform:uppercase; letter-spacing:0.6px; padding:0 14px 12px; white-space:nowrap;
          border-bottom:1px solid var(--border-color) !important;
        }
        /* nowrap para cada registro ocupar uma linha so; quem estoura a largura
           rola no container .gtbl-scroll, em vez de quebrar e inchar a altura. */
        .gtbl td { padding:12px 14px; border-bottom:1px solid rgba(255,255,255,0.04) !important; font-size:13px; color:var(--text-1); white-space:nowrap; }
        .gtbl tbody tr:last-child td { border-bottom:none !important; }
        .gtbl tbody tr:hover td { background:rgba(255,255,255,0.025) !important; }
        /* Centralizado (e nao a direita) para o numero ficar sob o proprio
           cabecalho, em vez de encostar na coluna seguinte.
           Precisa de "th.gtbl-num"/"td.gtbl-num": so ".gtbl-num" perde em
           especificidade para ".gtbl th { text-align:left }" e o cabecalho
           acabava a esquerda enquanto o valor ia pro centro. */
        .gtbl th.gtbl-num, .gtbl td.gtbl-num { text-align:center; }
        .gtbl-num { font-variant-numeric:tabular-nums; }
        .gtbl-muted { color:var(--text-2); }
        .gtbl-mono { font-size:12px; color:var(--text-2); }
        .gtbl-strong { font-weight:600; }
        .gtbl-scroll { overflow-x:auto; }
        .gtbl-tag {
          display:inline-block; font-size:12px; font-weight:600; color:#a9c5ff; white-space:nowrap;
          background:rgba(91,141,239,0.14); border:1px solid rgba(91,141,239,0.22);
          padding:3px 9px; border-radius:6px; letter-spacing:0.2px;
        }
        a.gtbl-link, a.gtbl-link:hover, a.gtbl-link:visited { text-decoration:none !important; color:#a9c5ff !important; }
        a.gtbl-link:hover { background:rgba(91,141,239,0.26) !important; border-color:rgba(91,141,239,0.45) !important; }
        .gtbl-badge {
          display:inline-block; min-width:26px; text-align:center; font-size:11.5px; font-weight:600;
          padding:3px 9px; border-radius:6px; font-variant-numeric:tabular-nums; white-space:nowrap;
        }
        .gtbl-badge.crit { color:#fca5a5; background:rgba(248,113,113,0.16); border:1px solid rgba(248,113,113,0.28); }
        .gtbl-badge.warn { color:#fcd34d; background:rgba(251,191,36,0.14); border:1px solid rgba(251,191,36,0.26); }
        .gtbl-badge.ok   { color:#6ee7d0; background:rgba(45,212,191,0.13); border:1px solid rgba(45,212,191,0.24); }
        .gtbl-empty { padding:34px 4px; text-align:center; color:var(--text-3); font-size:13px; }
        .prg-trilha { font-size:12.5px; color:var(--text-2); margin-bottom:18px; }
        .prg-sep { color:var(--text-3); margin:0 2px; }
        a.prg-link { color:#a9c5ff !important; text-decoration:none !important; font-weight:600; }
        a.prg-link:hover { text-decoration:underline !important; }
        .prg-nome { color:var(--text-2); }
        .prg-tot {
          display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:20px;
          background:var(--dark-card-2); border:1px solid var(--border-color);
          border-radius:14px; padding:22px 26px; margin-bottom:36px;
        }
        .prg-tot > div { display:flex; flex-direction:column; gap:4px; }
        .prg-tot-lbl { font-size:10.5px; text-transform:uppercase; letter-spacing:0.5px; color:var(--text-3); font-weight:600; }
        .prg-tot-val { font-size:19px; font-weight:800; color:var(--text-1); letter-spacing:-0.3px; }
        .prg-tot-sub { font-size:10.5px; color:var(--text-3); }
        .prg-espaco { height:12px; }
        /* centra o toggle na mesma linha do campo dos selects ao lado */
        .prg-toggle-espaco { height:34px; }

        /* arvore SOP > SSOP > TAGs em <details>: expande sem ida ao servidor */
        .arvore { display:flex; flex-direction:column; gap:8px; }
        .arv-no { border:1px solid var(--border-color); border-radius:11px; overflow:hidden;
                  background:var(--dark-card-2); }
        .arv-no[open] { border-color:var(--border-strong); }
        .arv-no > summary {
          display:flex; align-items:center; gap:14px; padding:13px 16px;
          cursor:pointer; list-style:none; user-select:none; transition:background 120ms;
        }
        .arv-no > summary::-webkit-details-marker { display:none; }
        .arv-no > summary:hover { background:rgba(255,255,255,0.035); }
        .arv-no[open] > summary { background:rgba(255,255,255,0.03);
                                  border-bottom:1px solid var(--border-color); }
        .arv-seta {
          width:17px; height:17px; flex-shrink:0; position:relative;
          border:1px solid var(--border-strong); border-radius:5px;
        }
        .arv-seta::before, .arv-seta::after {
          content:''; position:absolute; background:var(--text-2);
          left:50%; top:50%; transform:translate(-50%,-50%);
        }
        .arv-seta::before { width:9px; height:1.5px; }
        .arv-seta::after  { width:1.5px; height:9px; transition:opacity 140ms; }
        .arv-no[open] > summary .arv-seta::after { opacity:0; }
        .arv-no[open] > summary .arv-seta { border-color:var(--accent-teal); }
        .arv-no[open] > summary .arv-seta::before { background:var(--accent-teal); }
        .arv-nome { font-size:13px; font-weight:700; color:var(--text-1); min-width:150px; }
        .arv-num { font-size:11.5px; color:var(--text-3); white-space:nowrap; font-variant-numeric:tabular-nums; }
        .arv-sub { color:var(--text-2); font-weight:600; }
        .arv-val { color:var(--text-2); }
        .arv-avanco { display:flex; align-items:center; gap:9px; margin-left:auto; flex-shrink:0; }
        .arv-track { width:88px; height:6px; background:rgba(255,255,255,0.07); border-radius:4px; overflow:hidden; }
        .arv-fill { display:block; height:100%; border-radius:4px; }
        .arv-fill.ok { background:var(--accent-teal); }
        .arv-fill.warn { background:var(--accent-amber); }
        .arv-fill.crit { background:var(--accent-red); }
        .arv-pct { font-size:12.5px; font-weight:700; color:var(--text-1);
                   min-width:52px; text-align:right; font-variant-numeric:tabular-nums; }
        .arv-corpo { padding:12px 14px 14px; }
        .arv-n2 { background:var(--dark-card); margin-bottom:7px; }
        .arv-n2:last-child { margin-bottom:0; }
        .arv-n2 > summary { padding:11px 14px; }
        .arv-n2 .arv-nome { font-weight:600; font-size:12.5px; }
        .arv-dica { font-size:12px; color:var(--text-3); padding:10px 4px; }
        .arv-aviso {
          font-size:12.5px; color:var(--accent-amber); background:rgba(251,191,36,0.08);
          border:1px solid rgba(251,191,36,0.22); border-radius:10px;
          padding:12px 16px; margin-bottom:18px;
        }

        /* graficos de "mais avancados" em grid 2x2: as linhas do grid tem
           altura uniforme, entao as colunas nunca desalinham. */
        /* align-items:start deixa cada card com a altura do seu conteudo; com
           stretch, um card de 1 barra ficava do tamanho do de 5 e sobrava um
           vao de 200px+ dentro dele. */
        .gr-grid {
          display:grid; grid-template-columns:repeat(auto-fit, minmax(420px, 1fr));
          gap:28px; margin-bottom:8px; align-items:start;
        }
        /* height:auto anula o height:100% de .gplan-panel, que esticava o card
           de 1 barra ate a altura do de 5 (245px de vao interno). */
        .gr-panel { padding:26px 28px; margin-bottom:0 !important; height:auto !important; }
        .gr-panel .gplan-panel-title { margin-bottom:20px; }
        .gr-row { margin-bottom:16px; }
        .gr-row:last-child { margin-bottom:0; }
        .gr-top { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:6px; gap:12px; }
        .gr-nome { font-size:12.5px; font-weight:600; color:var(--text-1); white-space:nowrap;
                   overflow:hidden; text-overflow:ellipsis; }
        .gr-pct { font-size:13px; font-weight:800; color:var(--accent-teal); font-variant-numeric:tabular-nums; }
        .gr-track { height:8px; background:rgba(255,255,255,0.05); border-radius:5px; overflow:hidden; }
        .gr-fill { height:100%; border-radius:5px;
                   background:linear-gradient(90deg, var(--accent-teal), #22c1b0); transition:width 400ms ease; }
        .gr-sub { font-size:10.5px; color:var(--text-3); margin-top:5px; }

        /* barra de avanco dentro da celula da tabela */
        .cel-avanco { display:flex; align-items:center; gap:9px; justify-content:flex-end; }
        .cel-track { width:62px; height:6px; background:rgba(255,255,255,0.06); border-radius:4px; overflow:hidden; flex-shrink:0; }
        .cel-fill { height:100%; border-radius:4px; }
        .cel-fill.ok { background:var(--accent-teal); }
        .cel-fill.warn { background:var(--accent-amber); }
        .cel-fill.crit { background:var(--accent-red); }
        .cel-pct { font-size:12px; font-weight:600; min-width:46px; text-align:right; font-variant-numeric:tabular-nums; }

        .flt-summary { font-size:12.5px; color:var(--text-2); padding:2px 2px 0; }
        .flt-summary strong { color:var(--text-1); }

        .sg-chart-wrap { position:relative; width:100%; max-width:250px; margin: 4px auto 24px; }
        .sg-donut { width:100%; height:auto; display:block; }
        .sg-center {
          position:absolute; inset:0; display:flex; flex-direction:column;
          align-items:center; justify-content:center; pointer-events:none;
        }
        .sg-center-value { font-size:30px; font-weight:800; color:var(--text-1); letter-spacing:-1px; line-height:1; }
        .sg-center-label { font-size:11.5px; color:var(--text-3); margin-top:5px; }
        .sg-legend { display:flex; flex-direction:column; }
        .sg-leg-row {
          display:flex; align-items:center; gap:10px; padding:9px 8px;
          margin: 0 -8px; border-radius:7px;
          border-bottom:1px solid rgba(255,255,255,0.04);
        }
        .sg-leg-row:last-child { border-bottom:none; }
        .sg-leg-dot { width:9px; height:9px; border-radius:3px; flex-shrink:0; }
        .sg-leg-name { font-size:12.5px; color:var(--text-2); flex:1; }
        .sg-leg-val { font-size:12.5px; font-weight:700; color:var(--text-1); font-variant-numeric:tabular-nums; }

        /* As linhas e a legenda viraram <a> (drill-down para Relatorios
           filtrado); precisam voltar a se comportar como bloco e perder o
           estilo de link que o Streamlit aplica. */
        a.rep-row, a.sg-leg-row { text-decoration: none !important; cursor: pointer; }
        a.rep-row { display: block; border-radius: 8px; padding: 6px 8px; margin: 0 -8px 8px; transition: background 120ms; }
        a.rep-row:hover { background: rgba(255,255,255,0.035); }
        a.rep-row:last-of-type { margin-bottom: 0; }
        /* cor explicita: "inherit" herdaria o azul de link do Streamlit */
        a.rep-row .rep-name { color: var(--text-1) !important; }
        a.rep-row .rep-stat { color: var(--text-3) !important; }
        a.sg-leg-row .sg-leg-name { color: var(--text-2) !important; }
        a.sg-leg-row .sg-leg-val { color: var(--text-1) !important; }
        a.sg-leg-row:hover .sg-leg-name { color: var(--text-1) !important; }
        .sg-seg { transition: opacity 120ms; }
        .sg-donut:hover .sg-seg { opacity: 0.45; }
        .sg-donut .sg-seg:hover { opacity: 1; }
        a.sg-leg-row:hover { background: rgba(255,255,255,0.035); }
        a.sg-leg-row:hover .sg-leg-name { color: var(--text-1); }

        .rep-row { margin-bottom: 14px; }
        .rep-row:last-child { margin-bottom: 0; }
        .rep-label { display:flex; justify-content:space-between; margin-bottom:6px; font-size:12.5px; }
        .rep-name { font-weight:600; color: var(--text-1); }
        .rep-stat { color: var(--text-3); font-variant-numeric: tabular-nums; }
        .rep-track { height:7px; background: rgba(255,255,255,0.05); border-radius:4px; overflow:hidden; display:flex; }
        .rep-done { background: linear-gradient(90deg, var(--accent-teal), #22c1b0); height:100%; }
        .rep-pending { background: rgba(255,255,255,0.05); height:100%; }
        .doc-tag { font-size:9px; font-weight:600; color: var(--text-3); background: rgba(255,255,255,0.06); padding:1px 6px; border-radius:4px; text-transform:uppercase; letter-spacing:0.3px; margin-left:6px; }

        .group-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap:16px; }
        .group-card-v2 { background: var(--dark-card-2); border: 1px solid var(--border-color); border-radius: 12px; padding: 18px; }
        .group-card-top { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:12px; }
        .group-card-name { font-size:13px; font-weight:700; color: var(--text-1); }
        .group-card-pct { font-size:13px; font-weight:800; color: var(--accent-teal); }
        .group-card-value { font-size:24px; font-weight:800; color: var(--text-1); letter-spacing:-0.4px; margin-bottom:12px; }
        .group-card-unit { font-size:11px; font-weight:500; color: var(--text-3); }
        .group-card-track { height:6px; background: rgba(255,255,255,0.06); border-radius:3px; overflow:hidden; margin-bottom:12px; }
        .group-card-fill { height:100%; border-radius:3px; background: linear-gradient(90deg, var(--accent-teal), #22c1b0); }
        .group-card-nums { display:flex; justify-content:space-between; font-size:11px; color: var(--text-3); }

        .tag-detail-card { background: var(--dark-card); border: 1px solid var(--border-strong); border-radius: 16px; padding: 26px 28px; }
        .detail-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap:12px; margin-bottom:26px; }
        .detail-grid .detail-item { margin-bottom: 0; }
        .ficha-head { margin-bottom: 20px; }
        .ficha-tag { font-size:22px; font-weight:800; color:var(--text-1); letter-spacing:-0.5px; }
        .ficha-desc { font-size:13px; color:var(--text-2); margin-top:3px; }
        .ficha-sub { font-size:13.5px; font-weight:700; color:var(--text-1); margin-bottom:14px; }

        /* Modal da ficha resolvido por :target -- abre sem ida ao servidor. */
        /* acima da sidebar do Streamlit, que usa z-index 999991 */
        .fmodal { display:none; position:fixed; inset:0; z-index:1000000; }
        .fmodal:target { display:flex; align-items:center; justify-content:center; padding:32px 20px; }
        .fmodal-bg {
          position:absolute; inset:0; background:rgba(6,9,18,0.68);
          backdrop-filter:blur(6px); -webkit-backdrop-filter:blur(6px);
        }
        .fmodal-box {
          position:relative; width:min(1080px, 100%); max-height:88vh; overflow:auto;
          background:var(--dark-card); border:1px solid var(--border-strong);
          border-radius:16px; box-shadow:0 24px 70px rgba(0,0,0,0.6);
          animation: fmodal-in 140ms ease-out;
        }
        @keyframes fmodal-in { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
        .fmodal-head {
          display:flex; align-items:center; justify-content:space-between;
          padding:22px 28px 0; position:sticky; top:0; background:var(--dark-card); z-index:1;
        }
        .fmodal-title { font-size:20px; font-weight:800; color:var(--text-1); letter-spacing:-0.4px; }
        .fmodal-x {
          font-size:26px; line-height:1; color:var(--text-3) !important; text-decoration:none !important;
          padding:0 6px; border-radius:8px;
        }
        .fmodal-x:hover { color:var(--text-1) !important; background:rgba(255,255,255,0.06); }
        .fmodal-body { padding:16px 28px 28px; }
        .fmodal-body .ficha-sub { margin-top:4px; }

        /* Fundo desfocado atras do modal da ficha. */
        div[data-testid="stDialog"] > div:first-child,
        div[data-baseweb="modal"] > div:first-child {
          backdrop-filter: blur(6px) !important;
          -webkit-backdrop-filter: blur(6px) !important;
          background: rgba(6, 9, 18, 0.68) !important;
        }
        /* o painel do dialog e um <section>, nao um <div> */
        div[data-testid="stDialog"] [role="dialog"] {
          background: var(--dark-card) !important;
          border: 1px solid var(--border-strong) !important;
          border-radius: 16px !important;
          box-shadow: 0 24px 70px rgba(0,0,0,0.6) !important;
        }
        div[data-testid="stDialog"] [role="dialog"] h2 {
          font-size: 20px !important; font-weight: 800 !important;
          letter-spacing: -0.4px; padding-bottom: 4px;
        }

        .detail-item { background: var(--dark-card-2); border: 1px solid var(--border-color); border-radius: 12px; padding: 14px 16px; margin-bottom: 12px; }
        .detail-label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-3); font-weight: 600; margin-bottom: 6px; }
        .detail-value { font-size: 15px; font-weight: 700; color: var(--text-1); }

        div[data-testid="stMetric"] { background: var(--dark-card); border: 1px solid var(--border-color); border-radius: 12px; padding: 12px 16px; }
        </style>
        """
    )


def data_atualizacao(cache_key: str) -> str:
    """Quando a PLANILHA foi atualizada, no horario de Brasilia.

    Antes exibia pd.Timestamp.now(), que era a hora de renderizar a pagina:
    mudava a cada clique e, no Render (que roda em UTC), aparecia 3h adiantada.
    O cache_key ja carrega o updated_at do arquivo no Supabase; localmente e o
    mtime do xlsx.
    """
    try:
        if cache_key and cache_key[:4].isdigit():          # ISO-8601 do Supabase
            ts = pd.Timestamp(cache_key)
            ts = ts.tz_localize("UTC") if ts.tzinfo is None else ts.tz_convert("UTC")
        else:                                              # mtime local (epoch)
            ts = pd.Timestamp(float(cache_key), unit="s", tz="UTC")
        return ts.tz_convert(BR_TZ).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return "—"


def render_header(title: str, extra_pill: str | None = None):
    now = st.session_state.get("gplan_atualizado_em", "—")
    extra_html = f'<div class="gplan-count-pill">{extra_pill}</div>' if extra_pill else ""
    render_html(
        f"""
        <div class="gplan-header">
          <h1>{title}</h1>
          <div style="display:flex; gap:12px;">
            {extra_html}
            <div class="gplan-updated"><span class="dot"></span>Atualizado {now}</div>
          </div>
        </div>
        """
    )


KPI_ICONS = {
    "shield": '<path d="M12 2l7 4v6c0 5-3.4 8.7-7 10-3.6-1.3-7-5-7-10V6l7-4z"/>',
    "check": '<path d="M20 6L9 17l-5-5"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v6l4 2"/>',
    "archive": '<path d="M4 4h16v16H4z"/><path d="M4 9h16"/>',
    "trend": '<path d="M3 17l6-6 4 4 8-8"/><path d="M17 7h4v4"/>',
}


def br_num(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def status_panel(labels: list, values: list, colors: list, total: int) -> str:
    """Painel 'Status SIGEM' inteiro como um unico bloco HTML: donut em SVG puro
    + legenda com os valores alinhados a direita. Precisa ser um bloco so, porque
    o Streamlit fecha qualquer <div> deixado aberto entre chamadas de markdown."""
    radius, stroke = 68.0, 26.0
    circ = 2 * math.pi * radius
    total_val = sum(values) or 1

    segments, offset = "", 0.0
    for label, value, color in zip(labels, values, colors):
        length = value / total_val * circ
        # Cada fatia e um link para Relatorios ja filtrado por aquele status.
        segments += (
            f'<a href="/relatorios?status={quote(label)}" target="_self">'
            f'<title>{esc(label)} · {br_num(value)}</title>'
            f'<circle class="sg-seg" cx="88" cy="88" r="{radius}" fill="none" stroke="{color}" '
            f'stroke-width="{stroke}" stroke-dasharray="{length:.2f} {circ - length:.2f}" '
            f'stroke-dashoffset="{-offset:.2f}" transform="rotate(-90 88 88)"></circle></a>'
        )
        offset += length

    legend = ""
    for label, value, color in zip(labels, values, colors):
        legend += (
            f'<a class="sg-leg-row" href="/relatorios?status={quote(label)}" target="_self" '
            f'title="Ver {esc(label)} em Relatórios">'
            f'<span class="sg-leg-dot" style="background:{color};"></span>'
            f'<span class="sg-leg-name">{esc(label)}</span>'
            f'<span class="sg-leg-val">{br_num(value)}</span>'
            f"</a>"
        )

    return f"""
        <div class="gplan-panel">
          <div class="gplan-panel-title">Status SIGEM</div>
          <div class="sg-chart-wrap">
            <svg class="sg-donut" viewBox="0 0 176 176">{segments}</svg>
            <div class="sg-center">
              <div class="sg-center-value">{br_num(total)}</div>
              <div class="sg-center-label">Relatórios</div>
            </div>
          </div>
          <div class="sg-legend">{legend}</div>
        </div>
    """


def esc(value: object) -> str:
    text = "" if value is None else str(value)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def tag_pill(value: object) -> str:
    return f'<span class="gtbl-tag">{esc(value)}</span>'


def status_badge(status: str) -> str:
    """Badge com a cor do proprio status, reaproveitando o mapa usado no donut."""
    label = sentence_case(status)
    color = STATUS_COLOR_MAP.get(label, "#7c8aa8")
    return (
        f'<span class="gtbl-badge" style="color:{color}; background:{color}1f; '
        f'border:1px solid {color}3d;">{esc(label)}</span>'
    )


def yes_no_badge(value: object) -> str:
    yes = str(value).strip().upper() in {"SIM", "SIM.", "S", "YES", "TRUE"}
    return f'<span class="gtbl-badge {"ok" if yes else "crit"}">{"Sim" if yes else "Não"}</span>'


def html_table(headers: list, rows_html: str, empty_msg: str = "Nenhum registro encontrado.") -> str:
    """Tabela HTML estilizada, substituindo o st.dataframe (grid do Streamlit),
    que e um canvas fechado e nao aceita estilizacao."""
    if not rows_html:
        return f'<div class="gtbl-empty">{empty_msg}</div>'
    head = "".join(
        f'<th class="gtbl-num">{h[1:]}</th>' if h.startswith("#") else f"<th>{h}</th>"
        for h in headers
    )
    return (
        f'<div class="gtbl-scroll"><table class="gtbl">'
        f"<thead><tr>{head}</tr></thead><tbody>{rows_html}</tbody></table></div>"
    )


def ficha_anchor(tag: object) -> str:
    """Id de ancora da ficha. So letras/digitos para ser um id CSS valido."""
    return "f-" + "".join(c if c.isalnum() else "-" for c in str(tag))


def tag_link(value: object) -> str:
    """Pill da tag que abre a ficha em modal.

    Usa ancora (#) e nao query param: o modal e resolvido por CSS :target,
    sem ida ao servidor. Com ?ficha= o Streamlit refazia a navegacao inteira
    (reconexao + re-execucao do script) e o modal demorava a aparecer.
    """
    return (
        f'<a class="gtbl-tag gtbl-link" href="#{ficha_anchor(value)}" '
        f'title="Ver ficha de {esc(value)}">{esc(value)}</a>'
    )


def fichas_modais_html(tags_ids, resumo: pd.DataFrame, esperados: pd.DataFrame,
                       tags: pd.DataFrame) -> str:
    """Modais das tags visiveis na pagina, abertos/fechados via CSS :target."""
    blocos = ""
    for tag_id in dict.fromkeys(tags_ids):  # unicas, preservando a ordem
        corpo = tag_ficha_html(tag_id, resumo, esperados, tags, com_cabecalho=False)
        if corpo is None:
            continue
        blocos += f"""
            <div class="fmodal" id="{ficha_anchor(tag_id)}">
              <a class="fmodal-bg" href="#" aria-label="Fechar"></a>
              <div class="fmodal-box">
                <div class="fmodal-head">
                  <div class="fmodal-title">{esc(tag_id)}</div>
                  <a class="fmodal-x" href="#" aria-label="Fechar">&times;</a>
                </div>
                <div class="fmodal-body">{corpo}</div>
              </div>
            </div>
        """
    return blocos


def top10_panel(top10: pd.DataFrame) -> str:
    rows = ""
    for _, r in top10.iterrows():
        pendentes = int(r["RELATORIOS_PENDENTES"])
        esperados = int(r["RELATORIOS_ESPERADOS"])
        emitidos = int(r["RELATORIOS_POSTADOS"])
        conclusao = f"{r['AVANCO_DOCUMENTAL'] * 100:.1f}%".replace(".", ",").replace(",0%", "%")
        ratio = (pendentes / esperados) if esperados else 0
        tone = "crit" if ratio >= 0.8 else ("warn" if ratio >= 0.4 else "ok")
        rows += f"""
            <tr>
              <td>{tag_link(r['TAG'])}</td>
              <td class="gtbl-muted">{esc(str(r['GRUPO_REGRA']).title())}</td>
              <td class="gtbl-num">{esperados}</td>
              <td class="gtbl-num">{emitidos}</td>
              <td class="gtbl-num"><span class="gtbl-badge {tone}">{pendentes}</span></td>
              <td class="gtbl-num gtbl-strong">{conclusao}</td>
            </tr>
        """
    table = html_table(
        ["Tag", "Tipo", "#Esperados", "#Emitidos", "#Pendentes", "#Conclusão"], rows
    )
    return (
        '<div class="gplan-panel">'
        '<div class="gplan-panel-title">Top 10 tags com mais pendências</div>'
        f"{table}</div>"
    )


def kpi_card(label: str, value: str, sub: str, pct: float, color: str, icon: str) -> str:
    pct_display = f"{pct * 100:.1f}%".replace(".", ",")
    width = max(0.0, min(pct, 1.0)) * 100
    icon_svg = KPI_ICONS.get(icon, "")
    return f"""
        <div class="kpi-card" style="--kpi-accent:{color};">
          <div class="kpi-top">
            <div class="kpi-label">{label}</div>
            <div class="kpi-icon" style="background:{color}22; color:{color};">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{icon_svg}</svg>
            </div>
          </div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-sub">{sub}</div>
          <div class="kpi-progress-row">
            <div class="kpi-track"><div class="kpi-fill" style="width:{width:.1f}%; background:{color};"></div></div>
            <div class="kpi-pct">{pct_display}</div>
          </div>
        </div>
    """


def render_dashboard(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame):
    render_header("Dashboard")

    total_tags = len(resumo)
    total_esperados = int(resumo["RELATORIOS_ESPERADOS"].sum())
    total_emitidos = int(resumo["RELATORIOS_POSTADOS"].sum())
    total_pendentes = total_esperados - total_emitidos
    avanco_geral = (total_emitidos / total_esperados) if total_esperados else 0
    tags_completas = int((resumo["STATUS_DOCUMENTAL"] == "PRONTA DOCUMENTAL").sum())

    cols = st.columns(5)
    kpis = [
        ("Total de tags", f"{total_tags:,}".replace(",", "."), "Base principal", 1.0, "#5b8def", "shield"),
        ("Tags completas", f"{tags_completas:,}".replace(",", "."), "Pronta documental",
         (tags_completas / total_tags) if total_tags else 0, "#34d399", "check"),
        ("Pendentes", f"{total_pendentes:,}".replace(",", "."), "Aguardando entrega",
         (total_pendentes / total_esperados) if total_esperados else 0, "#f87171", "clock"),
        ("Emitidos SIGEM", f"{total_emitidos:,}".replace(",", "."), "Com status localizado",
         avanco_geral, "#fbbf24", "archive"),
        ("Avanço geral", f"{avanco_geral * 100:.1f}%".replace(".", ","), "Progressão do projeto",
         avanco_geral, "#9d6bff", "trend"),
    ]
    for col, (label, value, sub, pct, color, icon) in zip(cols, kpis):
        with col:
            render_html(kpi_card(label, value, sub, pct, color, icon))

    st.write("")
    col_left, col_right = st.columns([1.6, 1])

    with col_left:
        # Ordena do maior quantitativo esperado para o menor: barras em escada
        # leem melhor do que na ordem fixa em que as regras foram declaradas.
        barras = []
        for label, report, origin, unique_doc in REPORT_ROWS:
            barras.append((
                count_rows(esperados, report, origin, False, unique_doc),
                label, report, origin, unique_doc,
            ))
        barras.sort(key=lambda b: b[0], reverse=True)

        rows_html = ""
        for esperado, label, report, origin, unique_doc in barras:
            emitido = count_rows(esperados, report, origin, True, unique_doc)
            pct = (emitido / esperado * 100) if esperado else 0
            tag = '<span class="doc-tag">doc/planta</span>' if unique_doc else ""
            href = f"/relatorios?rel={quote(label)}"
            rows_html += f"""
                <a class="rep-row" href="{href}" target="_self" title="Ver {label} em Relatórios">
                  <div class="rep-label"><span class="rep-name">{label}{tag}</span><span class="rep-stat">{emitido:,}/{esperado:,} · {pct:.1f}%</span></div>
                  <div class="rep-track"><div class="rep-done" style="width:{pct:.1f}%;"></div><div class="rep-pending" style="width:{100-pct:.1f}%;"></div></div>
                </a>
            """.replace(",", ".")
        render_html(f'<div class="gplan-panel"><div class="gplan-panel-title">Esperado × emitido por relatório</div>{rows_html}</div>')

    with col_right:
        status_counts = esperados["STATUS_SIGEM"].value_counts()
        labels = [sentence_case(s) for s in status_counts.index]
        values = status_counts.values.tolist()
        colors = [STATUS_COLOR_MAP.get(label, DEFAULT_STATUS_COLORS[i % len(DEFAULT_STATUS_COLORS)]) for i, label in enumerate(labels)]
        render_html(status_panel(labels, values, colors, len(esperados)))

    st.write("")
    grouped = resumo.groupby("GRUPO_REGRA").agg(
        tags=("TAG", "count"),
        esperados=("RELATORIOS_ESPERADOS", "sum"),
        emitidos=("RELATORIOS_POSTADOS", "sum"),
    ).reset_index()
    grouped["avanco"] = (grouped["emitidos"] / grouped["esperados"]).fillna(0) * 100
    grouped = grouped.sort_values("tags", ascending=False)

    cards_html = ""
    for _, row in grouped.iterrows():
        cards_html += f"""
            <div class="group-card-v2">
              <div class="group-card-top"><span class="group-card-name">{row['GRUPO_REGRA'].title()}</span><span class="group-card-pct">{row['avanco']:.1f}%</span></div>
              <div class="group-card-value">{br_num(int(row['tags']))} <span class="group-card-unit">tags</span></div>
              <div class="group-card-track"><div class="group-card-fill" style="width:{row['avanco']:.1f}%;"></div></div>
              <div class="group-card-nums"><span>{br_num(int(row['emitidos']))} emitidos</span><span>{br_num(int(row['esperados']))} esperados</span></div>
            </div>
        """
    render_html(
        f'<div class="gplan-panel"><div class="gplan-panel-title">Resumo por grupo</div>'
        f'<div class="group-grid">{cards_html}</div></div>'
    )

    st.write("")
    top10 = resumo.sort_values("RELATORIOS_PENDENTES", ascending=False).head(10)
    render_html(
        top10_panel(top10)
        + fichas_modais_html(top10["TAG"].tolist(), resumo, esperados, tags)
    )


def render_relatorios(esperados: pd.DataFrame, resumo: pd.DataFrame, tags: pd.DataFrame):
    render_header("Relatórios previstos")

    # Mantem ORIGEM_REGRA no dataframe (sem exibir): e ela que separa
    # "RIR instrumentos" de "RIR cabos" no filtro por tipo de relatorio.
    df = esperados.copy()
    df["REVISAO_SIGEM"] = df["REVISAO_SIGEM"].fillna("—")
    df["DATA_SIGEM"] = format_date_column(df["DATA_SIGEM"])

    consume_url_filters(esperados)

    search = st.text_input("Pesquisar", placeholder="Pesquisar por tag, descrição, relatório, documento, status...", label_visibility="collapsed")

    status_options = sorted({sentence_case(s) for s in esperados["STATUS_SIGEM"].dropna().unique()})
    col_rel, col_sts = st.columns(2)
    with col_rel:
        sel_rel = st.multiselect("Tipo de relatório", REPORT_LABELS, key="flt_rel",
                                 placeholder="Todos os relatórios")
    with col_sts:
        sel_sts = st.multiselect("Status SIGEM", status_options, key="flt_status",
                                 placeholder="Todos os status")

    if sel_rel:
        df = filter_by_report_labels(df, sel_rel)
    if sel_sts:
        df = df[df["STATUS_SIGEM"].map(sentence_case).isin(sel_sts)]
    if search:
        # so nas colunas visiveis: varrer o dataframe inteiro faria a busca
        # casar com colunas internas como ORIGEM_REGRA.
        visiveis = ["TAG", "DESCRICAO", "GRUPO", "RELATORIO", "REFERENCIA",
                    "DOCUMENTO_ESPERADO", "STATUS_SIGEM", "REVISAO_SIGEM", "DATA_SIGEM"]
        df = df[search_any_column(df[visiveis], search)]

    if sel_rel or sel_sts or search:
        render_html(
            '<div class="flt-summary">Exibindo <strong>'
            + br_num(len(df))
            + "</strong> de "
            + br_num(len(esperados))
            + " relatórios</div>"
        )

    df_page = paginate(df, "relatorios", f"{search}|{sel_rel}|{sel_sts}")

    rows = ""
    for _, r in df_page.iterrows():
        rows += f"""
            <tr>
              <td>{tag_link(r['TAG'])}</td>
              <td>{esc(r['DESCRICAO'])}</td>
              <td class="gtbl-muted">{esc(str(r['GRUPO']).title())}</td>
              <td class="gtbl-strong">{esc(r['RELATORIO'])}</td>
              <td class="gtbl-muted">{esc(r['REFERENCIA'])}</td>
              <td class="gtbl-mono">{esc(r['DOCUMENTO_ESPERADO'])}</td>
              <td class="gtbl-num">{yes_no_badge(r['EXISTE_NO_SIGEM'])}</td>
              <td>{status_badge(r['STATUS_SIGEM'])}</td>
              <td class="gtbl-num gtbl-muted">{esc(r['REVISAO_SIGEM'])}</td>
              <td class="gtbl-num gtbl-muted">{esc(r['DATA_SIGEM'])}</td>
            </tr>
        """
    render_html(
        '<div class="gplan-panel">'
        + html_table(
            ["Tag", "Descrição", "Grupo", "Relatório", "Referência", "Documento esperado",
             "#Existe no SIGEM", "Status SIGEM", "#Revisão", "#Data"],
            rows,
            "Nenhum relatório encontrado para essa busca.",
        )
        + "</div>"
        + fichas_modais_html(df_page["TAG"].tolist(), resumo, esperados, tags)
    )


def tag_ficha_html(tag_id: str, resumo: pd.DataFrame, esperados: pd.DataFrame,
                   tags: pd.DataFrame, com_cabecalho: bool = True) -> str | None:
    """Ficha completa da tag como um bloco HTML unico, para servir tanto a aba
    Pesquisa tag quanto o modal aberto pelo Dashboard/Relatorios."""
    resumo_row = resumo[resumo["TAG"] == tag_id]
    if resumo_row.empty:
        return None

    r = resumo_row.iloc[0]
    tags_row = tags[tags["TAG"] == tag_id]
    comunicacao = tags_row.iloc[0]["COMUNICACAO"] if not tags_row.empty else "—"

    qtd_cabos = int(r["QTD_CABOS"])
    cabo_txt = f"Sim ({qtd_cabos} cabo{'s' if qtd_cabos != 1 else ''})" if qtd_cabos > 0 else "Não"

    # dados que vivem so na base de TAGs (hierarquia, criterio e preco)
    t = tags_row.iloc[0] if not tags_row.empty else {}

    def da_base(campo, default="—"):
        v = t.get(campo, default) if hasattr(t, "get") else default
        return default if v is None or vazio(v) else v

    preco = pd.to_numeric(pd.Series([t.get("PRECO_UNITARIO")]), errors="coerce").fillna(0).iloc[0] if hasattr(t, "get") else 0

    detalhes = [
        ("Tipo", str(r["GRUPO_REGRA"]).title()),
        ("Item PPU", r["ITEM_PPU"]),
        ("Comunicação", comunicacao),
        ("Tem cabo", cabo_txt),
        ("Tem tubing", "Sim" if str(r["TEM_TUBING"]).upper() == "SIM" else "Não"),
        ("Relatórios esperados", int(r["RELATORIOS_ESPERADOS"])),
        ("Relatórios entregues", int(r["RELATORIOS_POSTADOS"])),
        ("Relatórios pendentes", int(r["RELATORIOS_PENDENTES"])),
        ("Avanço", f"{r['AVANCO_DOCUMENTAL'] * 100:.1f}%".replace(".", ",")),
        ("SOP", da_base("SOP")),
        ("SSOP", da_base("SSOP")),
        ("Segmento", da_base("SEGMENTO")),
        ("Malha", da_base("MALHA")),
        ("Critério de medição", da_base("CRITERIO_MEDICAO")),
        ("Preço unitário", br_moeda(float(preco))),
    ]
    cards = "".join(
        f'<div class="detail-item"><div class="detail-label">{esc(lbl)}</div>'
        f'<div class="detail-value">{esc(val)}</div></div>'
        for lbl, val in detalhes
    )

    rows = ""
    for _, tr in esperados[esperados["TAG"] == tag_id].iterrows():
        rows += f"""
            <tr>
              <td class="gtbl-strong">{esc(tr['RELATORIO'])}</td>
              <td class="gtbl-muted">{esc(tr['REFERENCIA'])}</td>
              <td class="gtbl-mono">{esc(tr['DOCUMENTO_ESPERADO'])}</td>
              <td>{status_badge(tr['STATUS_SIGEM'])}</td>
              <td class="gtbl-num gtbl-muted">{esc(format_missing(tr['REVISAO_SIGEM']))}</td>
            </tr>
        """
    tabela = html_table(
        ["Relatório", "Referência", "Documento esperado", "Status SIGEM", "#Revisão"], rows
    )

    cabecalho = (
        f'<div class="ficha-head"><div class="ficha-tag">{esc(r["TAG"])}</div>'
        f'<div class="ficha-desc">{esc(r["DESCRICAO"])}</div></div>'
        if com_cabecalho else ""
    )
    return (
        f'<div class="ficha">{cabecalho}'
        f'<div class="detail-grid">{cards}</div>'
        f'<div class="ficha-sub">Relatórios da tag</div>{tabela}</div>'
    )


def render_pesquisa_tag(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame):
    render_header("Pesquisa tag")

    if "gplan_selected_tag" not in st.session_state:
        st.session_state.gplan_selected_tag = None

    # A lista virou tabela HTML, entao a selecao de linha vem por query param
    # (?tag=...) em vez do on_select do st.dataframe.
    tag_from_url = st.query_params.get("tag")
    if tag_from_url and tag_from_url != st.session_state.gplan_selected_tag:
        st.session_state.gplan_selected_tag = tag_from_url

    if st.session_state.gplan_selected_tag:
        tag_id = st.session_state.gplan_selected_tag
        if resumo[resumo["TAG"] == tag_id].empty:
            st.session_state.gplan_selected_tag = None
            st.query_params.clear()
            st.rerun()

        _, col_back = st.columns([4, 1])
        with col_back:
            if st.button("← Voltar à lista", use_container_width=True):
                st.session_state.gplan_selected_tag = None
                st.query_params.clear()
                st.rerun()

        render_html('<div class="gplan-panel">'
                    + tag_ficha_html(tag_id, resumo, esperados, tags)
                    + "</div>")
        return

    search = st.text_input("Pesquisar", placeholder="Digite a tag para ver a ficha completa (ex: AIT-120005)...", label_visibility="collapsed")

    list_df = resumo[["TAG", "DESCRICAO", "GRUPO_REGRA", "ITEM_PPU", "RELATORIOS_ESPERADOS", "AVANCO_DOCUMENTAL"]].copy()
    if search:
        mask = list_df["TAG"].astype(str).str.contains(search, case=False, na=False) | list_df["DESCRICAO"].astype(str).str.contains(search, case=False, na=False)
        list_df = list_df[mask]

    list_df["AVANCO_DOCUMENTAL"] = (list_df["AVANCO_DOCUMENTAL"] * 100).round(1)
    list_df_page = paginate(list_df, "pesquisa", search)

    rows = ""
    for _, r in list_df_page.iterrows():
        avanco = r["AVANCO_DOCUMENTAL"]
        tone = "ok" if avanco >= 70 else ("warn" if avanco >= 30 else "crit")
        href = f"?tag={quote(str(r['TAG']))}"
        rows += f"""
            <tr>
              <td><a class="gtbl-tag gtbl-link" href="{href}" target="_self">{esc(r['TAG'])}</a></td>
              <td>{esc(r['DESCRICAO'])}</td>
              <td class="gtbl-muted">{esc(str(r['GRUPO_REGRA']).title())}</td>
              <td class="gtbl-num gtbl-muted">{esc(r['ITEM_PPU'])}</td>
              <td class="gtbl-num">{int(r['RELATORIOS_ESPERADOS'])}</td>
              <td class="gtbl-num"><span class="gtbl-badge {tone}">{f"{avanco:.1f}".replace(".", ",")}%</span></td>
            </tr>
        """
    render_html(
        '<div class="gplan-panel">'
        + html_table(
            ["Tag", "Descrição", "Tipo", "#PPU", "#Relatórios", "#Avanço"],
            rows,
            "Nenhuma tag encontrada para essa busca.",
        )
        + "</div>"
    )


def render_sigem(sigem: pd.DataFrame):
    render_header("Base SIGEM", extra_pill=f"<strong>{len(sigem):,}</strong> documentos".replace(",", "."))

    status_options = ["Todos"] + sorted(sigem["STATUS"].dropna().unique().tolist())
    col1, col2 = st.columns([1, 3])
    with col1:
        status_filter = st.selectbox("Status", status_options)
    with col2:
        text_search = st.text_input("Pesquisa de texto", placeholder="Buscar em qualquer campo do documento...")

    df = sigem.copy()
    if status_filter != "Todos":
        df = df[df["STATUS"] == status_filter]
    if text_search:
        mask = search_any_column(df, text_search)
        df = df[mask]

    df = df.copy()
    df["_REVISAO_SORT"] = df["REVISAO"].astype(str)
    df["_DOCUMENTO_SORT"] = df["DOCUMENTO"].astype(str)
    df = df.sort_values(["_REVISAO_SORT", "_DOCUMENTO_SORT"]).drop(columns=["_REVISAO_SORT", "_DOCUMENTO_SORT"])

    df["DATA"] = format_date_column(df["DATA"])
    df_page = paginate(df, "sigem", f"{status_filter}|{text_search}")

    rows = ""
    for _, r in df_page.iterrows():
        rows += f"""
            <tr>
              <td class="gtbl-mono">{esc(r['DOCUMENTO'])}</td>
              <td>{status_badge(r['STATUS'])}</td>
              <td class="gtbl-num gtbl-muted">{esc(r['REVISAO'])}</td>
              <td class="gtbl-num gtbl-muted">{esc(r['DATA'])}</td>
              <td>{esc(r['TITULO'])}</td>
            </tr>
        """
    render_html(
        '<div class="gplan-panel">'
        + html_table(
            ["Documento", "Status", "#Revisão", "#Data", "Título"],
            rows,
            "Nenhum documento encontrado para esses filtros.",
        )
        + "</div>"
    )


SEM_VALOR = {"-", "", "NAN", "NONE"}


def vazio(v: object) -> bool:
    """A base marca ausencia com '-', nao com celula vazia."""
    return str(v).strip().upper() in SEM_VALOR


def br_moeda(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "~").replace(".", ",").replace("~", ".")


def progresso_base(resumo: pd.DataFrame, tags: pd.DataFrame) -> pd.DataFrame:
    """Une o avanco documental (07_TAG_RESUMO) com a hierarquia e o preco
    (01_BASE_TAGS). O avanco por TAG e o mesmo usado nas demais abas."""
    cols = ["TAG", "SOP", "SSOP", "SUBGRUPO_PRIORIDADE", "SEGMENTO", "MALHA",
            "CRITERIO_MEDICAO", "PRECO_UNITARIO",
            "STATUS_LOCALIZACAO", "STATUS_CALIBRACAO", "STATUS_MONTAGEM", "STATUS_FINAL"]
    disponiveis = [c for c in cols if c in tags.columns]
    df = resumo.merge(tags[disponiveis], on="TAG", how="left")
    for c in ("SOP", "SSOP", "SEGMENTO", "MALHA", "SUBGRUPO_PRIORIDADE"):
        if c not in df.columns:
            df[c] = "-"
        df[c] = df[c].fillna("-").astype(str).str.strip()
    if "PRECO_UNITARIO" not in df.columns:
        df["PRECO_UNITARIO"] = 0.0
    df["PRECO_UNITARIO"] = pd.to_numeric(df["PRECO_UNITARIO"], errors="coerce").fillna(0.0)
    # O valor so avanca quando a TAG fecha 100% documental -- nao ha medicao
    # parcial: uma TAG a 90% vale zero.
    df["COMPLETA"] = df["AVANCO_DOCUMENTAL"].fillna(0) >= 1.0
    df["VALOR_AVANCADO"] = df["PRECO_UNITARIO"].where(df["COMPLETA"], 0.0)
    return df


def ordem_prioridade(v: object) -> tuple:
    """'6.1' -> (6,1). Ordena por grupo e subgrupo numericos, senao '3.12'
    viria antes de '3.2' na ordenacao alfabetica. Sem prioridade vai pro fim."""
    s = str(v).strip()
    if vazio(s):
        return (9999, 9999)
    try:
        partes = [int(p) for p in s.replace(",", ".").split(".")]
        return tuple(partes + [0] * (2 - len(partes)))[:2]
    except ValueError:
        return (9998, 9998)


def prioridade_do_grupo(serie: pd.Series) -> str:
    """A prioridade mais alta (menor numero) entre as TAGs do grupo."""
    reais = [v for v in serie if not vazio(v)]
    return min(reais, key=ordem_prioridade) if reais else "-"


def agrega_nivel(df: pd.DataFrame, coluna: str, subnivel: str | None = None) -> pd.DataFrame:
    """Avanco = soma(emitidos)/soma(esperados): um relatorio pesa o mesmo em
    qualquer nivel, coerente com o avanco do Dashboard e com medicao. A media
    dos subniveis daria outro numero (num SOP chegou a 14pp de diferenca),
    porque faria um SSOP de 1 TAG pesar como um de 173."""
    agg = dict(
        tags=("TAG", "count"),
        esperados=("RELATORIOS_ESPERADOS", "sum"),
        emitidos=("RELATORIOS_POSTADOS", "sum"),
        completas=("COMPLETA", "sum"),
        valor=("PRECO_UNITARIO", "sum"),
        valor_avancado=("VALOR_AVANCADO", "sum"),
        prioridade=("SUBGRUPO_PRIORIDADE", prioridade_do_grupo),
    )
    if subnivel:
        agg["subniveis"] = (subnivel, "nunique")
    g = df.groupby(coluna).agg(**agg).reset_index()
    g["avanco"] = (g["emitidos"] / g["esperados"]).fillna(0) * 100
    return g.sort_values("avanco", ascending=False)


def pill_prioridade(v: object) -> str:
    if vazio(v):
        return '<span class="gtbl-muted">—</span>'
    grupo = ordem_prioridade(v)[0]
    tom = "crit" if grupo <= 3 else ("warn" if grupo <= 6 else "ok")
    return f'<span class="gtbl-badge {tom}">{esc(v)}</span>'


MIN_TAGS_GRAFICO = 5


def grafico_avanco(titulo: str, g: pd.DataFrame, coluna: str,
                   rotulo_sub: str = "", limite: int = 5) -> str:
    """Barras dos mais avancados, para indicar por onde comecar.

    Exige MIN_TAGS_GRAFICO: sem isso o ranking enche de pacotes de 1 TAG, que
    sobem por acaso e nao representam trabalho relevante.
    """
    validos = g[~g[coluna].apply(vazio) & (g["tags"] >= MIN_TAGS_GRAFICO)]
    dados = validos.nlargest(limite, "avanco")
    if dados.empty:
        return (f'<div class="gplan-panel gr-panel"><div class="gplan-panel-title">{esc(titulo)}</div>'
                f'<div class="gtbl-empty">Nenhum grupo com {MIN_TAGS_GRAFICO}+ tags.</div></div>')
    linhas = ""
    for _, r in dados.iterrows():
        pct = r["avanco"]
        if rotulo_sub and "subniveis" in r:
            quant = f"{br_num(int(r['subniveis']))} {rotulo_sub} · {br_num(int(r['tags']))} tags"
        else:
            quant = f"{br_num(int(r['tags']))} tags"
        linhas += f"""
            <div class="gr-row">
              <div class="gr-top">
                <span class="gr-nome">{esc(r[coluna])}</span>
                <span class="gr-pct">{f"{pct:.1f}".replace(".", ",")}%</span>
              </div>
              <div class="gr-track"><div class="gr-fill" style="width:{max(pct, 1.5):.1f}%;"></div></div>
              <div class="gr-sub">{quant} · prioridade {esc(r['prioridade'])}</div>
            </div>
        """
    return (
        '<div class="gplan-panel gr-panel">'
        f'<div class="gplan-panel-title">{esc(titulo)}</div>'
        f"{linhas}</div>"
    )


def status_pill(v: object) -> str:
    """Status de campo com cor por natureza (bom / atencao / ruim)."""
    if vazio(v):
        return '<span class="gtbl-muted">—</span>'
    t = str(v).strip()
    n = t.upper()
    bons = {"LOCALIZADO", "APROVADO", "MONTADO", "CALIBRADO", "SIM", "APTO"}
    ruins = {"NAO LOCALIZADO", "NÃO LOCALIZADO", "REPROVADO", "NAO MONTADO",
             "NÃO MONTADO", "CANCELADO", "NAO APTO", "NÃO APTO"}
    tom = "ok" if n in bons else ("crit" if n in ruins else "warn")
    return f'<span class="gtbl-badge {tom}">{esc(t)}</span>'


def render_progresso(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame):
    render_header("Progresso")

    df = progresso_base(resumo, tags)

    # filtros independentes da hierarquia
    segs = ["Todos"] + sorted({s for s in df.SEGMENTO if not vazio(s)})
    malhas = ["Todas"] + sorted({m for m in df.MALHA if not vazio(m)})
    c1, c2, c3 = st.columns([2, 2, 1.4])
    with c1:
        f_seg = st.selectbox("Segmento", segs, key="prg_seg")
    with c2:
        f_malha = st.selectbox("Malha", malhas, key="prg_malha")
    with c3:
        # rotulo invisivel so para alinhar a base do toggle com a dos selects
        st.markdown('<div class="prg-toggle-espaco"></div>', unsafe_allow_html=True)
        detalhar = st.toggle("Mostrar instrumentos", value=False, key="prg_detalhe",
                             help="Abre o terceiro nível: as TAGs dentro de cada SSOP")
    if f_seg != "Todos":
        df = df[df.SEGMENTO == f_seg]
    if f_malha != "Todas":
        df = df[df.MALHA == f_malha]

    render_html(_painel_totais(df, "Geral"))
    _graficos(df)
    _arvore(df, detalhar, resumo, esperados, tags)


def _barra(pct: float) -> str:
    tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
    return (
        '<span class="arv-avanco">'
        f'<span class="arv-track"><span class="arv-fill {tom}" style="width:{max(pct,1.5):.1f}%;"></span></span>'
        f'<span class="arv-pct">{f"{pct:.1f}".replace(".", ",")}%</span></span>'
    )


def _resumo_no(r) -> str:
    """Numeros que aparecem na linha do SOP/SSOP, ao lado do nome."""
    return (
        f'<span class="arv-num">{br_num(int(r["tags"]))} tags</span>'
        f'<span class="arv-num">{br_num(int(r["emitidos"]))}/{br_num(int(r["esperados"]))} relat.</span>'
        f'<span class="arv-num arv-val">{br_moeda(r["valor"])}</span>'
        f"{_barra(r['avanco'])}"
    )


def _tabela_tags(sub: pd.DataFrame, com_modal: bool = True) -> str:
    linhas = ""
    for _, t in sub.sort_values("AVANCO_DOCUMENTAL", ascending=False).iterrows():
        pct = (t["AVANCO_DOCUMENTAL"] or 0) * 100
        tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
        # sem modal, a pill leva para a ficha na aba Pesquisa tag
        alvo = tag_link(t["TAG"]) if com_modal else (
            f'<a class="gtbl-tag gtbl-link" href="/pesquisa?tag={quote(str(t["TAG"]))}" '
            f'target="_self" title="Abrir ficha de {esc(t["TAG"])}">{esc(t["TAG"])}</a>')
        linhas += f"""
            <tr>
              <td>{alvo}</td>
              <td>{esc(t['DESCRICAO'])}</td>
              <td class="gtbl-num">{pill_prioridade(t.get('SUBGRUPO_PRIORIDADE'))}</td>
              <td class="gtbl-num">{int(t['RELATORIOS_POSTADOS'])}/{int(t['RELATORIOS_ESPERADOS'])}</td>
              <td class="gtbl-num"><span class="gtbl-badge {tom}">{f"{pct:.1f}".replace(".", ",")}%</span></td>
              <td class="gtbl-num">{status_pill(t.get('STATUS_LOCALIZACAO'))}</td>
              <td class="gtbl-num">{status_pill(t.get('STATUS_CALIBRACAO'))}</td>
              <td class="gtbl-num">{status_pill(t.get('STATUS_MONTAGEM'))}</td>
              <td class="gtbl-num gtbl-muted">{'—' if vazio(t.get('STATUS_FINAL')) else esc(t.get('STATUS_FINAL'))}</td>
              <td class="gtbl-num gtbl-muted">{br_moeda(t['PRECO_UNITARIO'])}</td>
            </tr>
        """
    return html_table(
        ["Tag", "Descrição", "#Prioridade", "#Emit./Esp.", "#Avanço",
         "#Localização", "#Calibração", "#Montagem", "#Status final", "#Preço unit."],
        linhas, "Nenhuma TAG.",
    )


# Medido na base real: a tabela custa ~800 bytes por TAG, mas cada modal de
# ficha custa ~4,8 KB. Listar as 5.097 daria 3,9 MB de tabela e 23,3 MB de
# modais. Dai dois limites: a lista aguenta bem mais do que as fichas.
MAX_TAGS_DETALHE = 2000
MAX_TAGS_MODAL = 400


def _arvore(df: pd.DataFrame, detalhar: bool, resumo=None, esperados=None, tags=None):
    """SOP > SSOP > TAGs em <details>: expande no lugar, sem ida ao servidor,
    e permite abrir varios ao mesmo tempo.

    O detalhamento so e montado ate MAX_TAGS_DETALHE, e as fichas em modal ate
    MAX_TAGS_MODAL -- listar as 5.097 daria 3,9 MB de tabela mais 23,3 MB de
    modais, o que travaria a pagina no plano gratuito.
    """
    excedeu = detalhar and len(df) > MAX_TAGS_DETALHE
    detalhar = detalhar and not excedeu
    if excedeu:
        render_html(
            '<div class="arv-aviso"><strong>' + br_num(len(df)) + " TAGs</strong> no recorte "
            "atual, acima do limite de " + br_num(MAX_TAGS_DETALHE) + " para listar de uma vez. "
            "Escolha um <strong>segmento</strong> ou uma <strong>malha</strong> nos filtros "
            "acima que os instrumentos aparecem.</div>"
        )

    # a dica dentro do SSOP tem de refletir o motivo real de nao listar
    if excedeu:
        dica = ('<div class="arv-dica">Filtre um segmento ou malha acima para ver '
                'as TAGs deste subpacote.</div>')
    else:
        dica = ('<div class="arv-dica">Ligue “Mostrar instrumentos” acima para listar '
                'as TAGs deste subpacote.</div>')

    # fichas em modal so ate MAX_TAGS_MODAL; acima disso a pill abre a aba
    # Pesquisa tag, que mostra a mesma ficha sem custar 4,8 KB por TAG
    com_modal = detalhar and len(df) <= MAX_TAGS_MODAL

    g_sop = agrega_nivel(df, "SOP", subnivel="SSOP")
    blocos = ""
    for _, s in g_sop.iterrows():
        nome_sop = s["SOP"]
        rotulo = "Sem SOP" if vazio(nome_sop) else esc(nome_sop)
        sub_sop = df[df.SOP == nome_sop]

        ssops = ""
        for _, ss in agrega_nivel(sub_sop, "SSOP").iterrows():
            nome_ssop = ss["SSOP"]
            corpo = (_tabela_tags(sub_sop[sub_sop.SSOP == nome_ssop], com_modal)
                     if detalhar else dica)
            ssops += f"""
                <details class="arv-no arv-n2">
                  <summary>
                    <span class="arv-seta"></span>
                    <span class="arv-nome">{'Sem SSOP' if vazio(nome_ssop) else esc(nome_ssop)}</span>
                    {pill_prioridade(ss['prioridade'])}
                    {_resumo_no(ss)}
                  </summary>
                  <div class="arv-corpo">{corpo}</div>
                </details>
            """

        blocos += f"""
            <details class="arv-no arv-n1">
              <summary>
                <span class="arv-seta"></span>
                <span class="arv-nome">{rotulo}</span>
                {pill_prioridade(s['prioridade'])}
                <span class="arv-num arv-sub">{br_num(int(s['subniveis']))} SSOP</span>
                {_resumo_no(s)}
              </summary>
              <div class="arv-corpo">{ssops}</div>
            </details>
        """

    modais = (fichas_modais_html(df["TAG"].tolist(), resumo, esperados, tags)
              if com_modal and resumo is not None else "")
    render_html(
        '<div class="prg-espaco"></div><div class="gplan-panel">'
        '<div class="gplan-panel-title">SOP · clique para expandir</div>'
        f'<div class="arvore">{blocos or "<div class=\'gtbl-empty\'>Nada para esses filtros.</div>"}</div>'
        "</div>" + modais
    )


def _painel_totais(df: pd.DataFrame, titulo: str) -> str:
    esp = int(df["RELATORIOS_ESPERADOS"].sum())
    emi = int(df["RELATORIOS_POSTADOS"].sum())
    pct = (emi / esp * 100) if esp else 0
    completas = int(df["COMPLETA"].sum())
    total_v = float(df["PRECO_UNITARIO"].sum())
    avanc_v = float(df["VALOR_AVANCADO"].sum())
    pct_v = (avanc_v / total_v * 100) if total_v else 0
    return (
        '<div class="prg-tot">'
        f'<div><span class="prg-tot-lbl">TAGs</span><span class="prg-tot-val">{br_num(len(df))}</span></div>'
        f'<div><span class="prg-tot-lbl">TAGs completas</span>'
        f'<span class="prg-tot-val">{br_num(completas)}</span>'
        f'<span class="prg-tot-sub">100% documental</span></div>'
        f'<div><span class="prg-tot-lbl">Avanço documental</span>'
        f'<span class="prg-tot-val">{f"{pct:.1f}".replace(".", ",")}%</span>'
        f'<span class="prg-tot-sub">{br_num(emi)} de {br_num(esp)} relatórios</span></div>'
        f'<div><span class="prg-tot-lbl">Valor total</span>'
        f'<span class="prg-tot-val">{br_moeda(total_v)}</span></div>'
        f'<div><span class="prg-tot-lbl">Valor avançado</span>'
        f'<span class="prg-tot-val">{br_moeda(avanc_v)}</span>'
        f'<span class="prg-tot-sub">{f"{pct_v:.1f}".replace(".", ",")}% · só TAGs completas</span></div>'
        "</div>"
    )


def _graficos(df: pd.DataFrame):
    """Os quatro recortes mais avancados: onde ha mais chance de fechar rapido.

    Um grid CSS unico, e nao st.columns: cada coluna do Streamlit empilha de
    forma independente, entao com poucos itens as colunas ficavam com alturas
    diferentes (chegou a 218px de desequilibrio) e abria um vao no meio.
    """
    blocos = "".join([
        grafico_avanco("SOP mais avançados", agrega_nivel(df, "SOP", subnivel="SSOP"),
                       "SOP", rotulo_sub="SSOP"),
        grafico_avanco("SSOP mais avançados", agrega_nivel(df, "SSOP"), "SSOP"),
        grafico_avanco("Segmentos mais avançados",
                       agrega_nivel(df, "SEGMENTO", subnivel="MALHA"), "SEGMENTO",
                       rotulo_sub="malhas"),
        grafico_avanco("Malhas mais avançadas", agrega_nivel(df, "MALHA"), "MALHA"),
    ])
    render_html(f'<div class="gr-grid">{blocos}</div>')


def main():
    favicon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "favicon.png")
    st.set_page_config(
        page_title="Gplan",
        page_icon=favicon if os.path.exists(favicon) else "📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()

    cache_key = get_source_cache_key()
    if cache_key == "missing":
        st.error(
            "Não encontrei a planilha no Supabase Storage nem localmente. "
            "Verifique se o arquivo foi enviado ao bucket 'gplan-data'."
        )
        st.stop()

    st.session_state["gplan_atualizado_em"] = data_atualizacao(cache_key)
    tags, cabos, tubing, sigem, resumo, esperados = load_data(cache_key)

    with st.sidebar:
        render_html(
            '<div class="gplan-brand">'
            '<svg class="gplan-brand-mark" viewBox="0 0 48 48" fill="none">'
            '<circle cx="24" cy="24" r="19" stroke="#232a44" stroke-width="5"/>'
            '<path d="M24 5a19 19 0 0 1 15.6 29.8" stroke="url(#gplanArc)" stroke-width="5" stroke-linecap="round"/>'
            '<circle cx="24" cy="24" r="5.5" fill="#2dd4bf"/>'
            '<path d="M24 24L33 15" stroke="#f4f6fb" stroke-width="3" stroke-linecap="round"/>'
            '<defs><linearGradient id="gplanArc" x1="24" y1="5" x2="40" y2="35" gradientUnits="userSpaceOnUse">'
            '<stop stop-color="#5b8def"/><stop offset="1" stop-color="#2dd4bf"/>'
            "</linearGradient></defs></svg>"
            '<div class="gplan-brand-text">'
            '<div class="gplan-brand-name">Gplan</div>'
            '<div class="gplan-brand-sub">Instrumentação · U-12</div>'
            "</div></div>"
        )

    dashboard_page = st.Page(lambda: render_dashboard(resumo, esperados, tags), title="Dashboard", icon=":material/dashboard:", url_path="dashboard", default=True)
    relatorios_page = st.Page(lambda: render_relatorios(esperados, resumo, tags), title="Relatórios", icon=":material/description:", url_path="relatorios")
    progresso_page = st.Page(lambda: render_progresso(resumo, esperados, tags), title="Progresso", icon=":material/insights:", url_path="progresso")
    pesquisa_page = st.Page(lambda: render_pesquisa_tag(resumo, esperados, tags), title="Pesquisa tag", icon=":material/search:", url_path="pesquisa")
    sigem_page = st.Page(lambda: render_sigem(sigem), title="Base SIGEM", icon=":material/database:", url_path="sigem")

    nav = st.navigation([dashboard_page, progresso_page, relatorios_page, pesquisa_page, sigem_page], position="sidebar")
    nav.run()


if __name__ == "__main__":
    main()
