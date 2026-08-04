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
          display:flex; align-items:center; gap:10px; padding:9px 0;
          border-bottom:1px solid rgba(255,255,255,0.04);
        }
        .sg-leg-row:last-child { border-bottom:none; }
        .sg-leg-dot { width:9px; height:9px; border-radius:3px; flex-shrink:0; }
        .sg-leg-name { font-size:12.5px; color:var(--text-2); flex:1; }
        .sg-leg-val { font-size:12.5px; font-weight:700; color:var(--text-1); font-variant-numeric:tabular-nums; }

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
        .detail-item { background: var(--dark-card-2); border: 1px solid var(--border-color); border-radius: 12px; padding: 14px 16px; margin-bottom: 12px; }
        .detail-label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-3); font-weight: 600; margin-bottom: 6px; }
        .detail-value { font-size: 15px; font-weight: 700; color: var(--text-1); }

        div[data-testid="stMetric"] { background: var(--dark-card); border: 1px solid var(--border-color); border-radius: 12px; padding: 12px 16px; }
        </style>
        """
    )


def render_header(title: str, extra_pill: str | None = None):
    now = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
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
    for value, color in zip(values, colors):
        length = value / total_val * circ
        segments += (
            f'<circle cx="88" cy="88" r="{radius}" fill="none" stroke="{color}" '
            f'stroke-width="{stroke}" stroke-dasharray="{length:.2f} {circ - length:.2f}" '
            f'stroke-dashoffset="{-offset:.2f}" transform="rotate(-90 88 88)"></circle>'
        )
        offset += length

    legend = ""
    for label, value, color in zip(labels, values, colors):
        legend += (
            f'<div class="sg-leg-row">'
            f'<span class="sg-leg-dot" style="background:{color};"></span>'
            f'<span class="sg-leg-name">{label}</span>'
            f'<span class="sg-leg-val">{br_num(value)}</span>'
            f"</div>"
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
              <td>{tag_pill(r['TAG'])}</td>
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


def render_dashboard(resumo: pd.DataFrame, esperados: pd.DataFrame):
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
        rows_html = ""
        for label, report, origin, unique_doc in REPORT_ROWS:
            esperado = count_rows(esperados, report, origin, False, unique_doc)
            emitido = count_rows(esperados, report, origin, True, unique_doc)
            pct = (emitido / esperado * 100) if esperado else 0
            tag = '<span class="doc-tag">doc/planta</span>' if unique_doc else ""
            rows_html += f"""
                <div class="rep-row">
                  <div class="rep-label"><span class="rep-name">{label}{tag}</span><span class="rep-stat">{emitido:,}/{esperado:,} · {pct:.1f}%</span></div>
                  <div class="rep-track"><div class="rep-done" style="width:{pct:.1f}%;"></div><div class="rep-pending" style="width:{100-pct:.1f}%;"></div></div>
                </div>
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
    render_html(top10_panel(top10))


def render_relatorios(esperados: pd.DataFrame):
    render_header("Relatórios previstos")

    display_cols = ["TAG", "DESCRICAO", "GRUPO", "RELATORIO", "REFERENCIA", "DOCUMENTO_ESPERADO",
                     "EXISTE_NO_SIGEM", "STATUS_SIGEM", "REVISAO_SIGEM", "DATA_SIGEM"]
    df = esperados[display_cols].copy()
    df["REVISAO_SIGEM"] = df["REVISAO_SIGEM"].fillna("—")
    df["DATA_SIGEM"] = format_date_column(df["DATA_SIGEM"])

    search = st.text_input("Pesquisar", placeholder="Pesquisar por tag, descrição, relatório, documento, status...", label_visibility="collapsed")
    if search:
        df = df[search_any_column(df, search)]

    df_page = paginate(df, "relatorios", search)

    rows = ""
    for _, r in df_page.iterrows():
        rows += f"""
            <tr>
              <td>{tag_pill(r['TAG'])}</td>
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
        resumo_row = resumo[resumo["TAG"] == tag_id]
        tags_row = tags[tags["TAG"] == tag_id]
        if resumo_row.empty:
            st.session_state.gplan_selected_tag = None
            st.query_params.clear()
            st.rerun()

        r = resumo_row.iloc[0]
        comunicacao = tags_row.iloc[0]["COMUNICACAO"] if not tags_row.empty else "—"

        col_title, col_back = st.columns([4, 1])
        with col_title:
            st.markdown(f"### {r['TAG']}")
            st.caption(r["DESCRICAO"])
        with col_back:
            if st.button("← Voltar à lista", use_container_width=True):
                st.session_state.gplan_selected_tag = None
                st.query_params.clear()
                st.rerun()

        qtd_cabos = int(r["QTD_CABOS"])
        cabo_txt = f"Sim ({qtd_cabos} cabo{'s' if qtd_cabos != 1 else ''})" if qtd_cabos > 0 else "Não"
        tubing_txt = "Sim" if str(r["TEM_TUBING"]).upper() == "SIM" else "Não"

        detail_cols = st.columns(3)
        details = [
            ("Tipo", r["GRUPO_REGRA"].title()),
            ("Item PPU", r["ITEM_PPU"]),
            ("Comunicação", comunicacao),
            ("Tem cabo", cabo_txt),
            ("Tem tubing", tubing_txt),
            ("Relatórios esperados", int(r["RELATORIOS_ESPERADOS"])),
            ("Relatórios entregues", int(r["RELATORIOS_POSTADOS"])),
            ("Relatórios pendentes", int(r["RELATORIOS_PENDENTES"])),
            ("Avanço", f"{r['AVANCO_DOCUMENTAL'] * 100:.1f}%".replace(".", ",")),
        ]
        for i, (label, value) in enumerate(details):
            with detail_cols[i % 3]:
                render_html(f'<div class="detail-item"><div class="detail-label">{label}</div><div class="detail-value">{value}</div></div>')

        tag_reports = esperados[esperados["TAG"] == tag_id][
            ["RELATORIO", "REFERENCIA", "DOCUMENTO_ESPERADO", "STATUS_SIGEM", "REVISAO_SIGEM"]
        ].copy()
        rows = ""
        for _, tr in tag_reports.iterrows():
            rows += f"""
                <tr>
                  <td class="gtbl-strong">{esc(tr['RELATORIO'])}</td>
                  <td class="gtbl-muted">{esc(tr['REFERENCIA'])}</td>
                  <td class="gtbl-mono">{esc(tr['DOCUMENTO_ESPERADO'])}</td>
                  <td>{status_badge(tr['STATUS_SIGEM'])}</td>
                  <td class="gtbl-num gtbl-muted">{esc(format_missing(tr['REVISAO_SIGEM']))}</td>
                </tr>
            """
        render_html(
            '<div class="gplan-panel">'
            '<div class="gplan-panel-title">Relatórios da tag</div>'
            + html_table(
                ["Relatório", "Referência", "Documento esperado", "Status SIGEM", "#Revisão"], rows
            )
            + "</div>"
        )
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

    dashboard_page = st.Page(lambda: render_dashboard(resumo, esperados), title="Dashboard", icon=":material/dashboard:", url_path="dashboard", default=True)
    relatorios_page = st.Page(lambda: render_relatorios(esperados), title="Relatórios", icon=":material/description:", url_path="relatorios")
    pesquisa_page = st.Page(lambda: render_pesquisa_tag(resumo, esperados, tags), title="Pesquisa tag", icon=":material/search:", url_path="pesquisa")
    sigem_page = st.Page(lambda: render_sigem(sigem), title="Base SIGEM", icon=":material/database:", url_path="sigem")

    nav = st.navigation([dashboard_page, relatorios_page, pesquisa_page, sigem_page], position="sidebar")
    nav.run()


if __name__ == "__main__":
    main()
