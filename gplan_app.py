import io
import math
import os

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
        section[data-testid="stSidebar"] {
          background: linear-gradient(180deg, #0d1224 0%, #0a0e1a 100%);
          border-right: 1px solid var(--border-color);
        }
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
        .kpi-label { font-size: 11.5px; color: var(--text-2); font-weight:600; letter-spacing:0.2px; margin-bottom: 12px; }
        .kpi-value { font-size: 30px; font-weight: 800; color: var(--text-1); letter-spacing:-0.8px; margin-bottom:4px; }
        .kpi-sub { font-size: 12px; color: var(--text-3); margin-bottom: 14px; }
        .kpi-progress-row { display:flex; align-items:center; gap:10px; }
        .kpi-track { flex:1; height:5px; background: rgba(255,255,255,0.06); border-radius:3px; overflow:hidden; }
        .kpi-fill { height:100%; border-radius:3px; }
        .kpi-pct { font-size:12px; font-weight:700; color: var(--text-1); min-width:38px; text-align:right; }

        .gplan-panel { background: var(--dark-card); border: 1px solid var(--border-color); border-radius: 16px; padding: 24px; height: 100%; }
        .gplan-panel-title { font-size: 14.5px; font-weight: 700; color: var(--text-1); margin-bottom: 18px; }

        .rep-row { margin-bottom: 14px; }
        .rep-row:last-child { margin-bottom: 0; }
        .rep-label { display:flex; justify-content:space-between; margin-bottom:6px; font-size:12.5px; }
        .rep-name { font-weight:600; color: var(--text-1); }
        .rep-stat { color: var(--text-3); font-variant-numeric: tabular-nums; }
        .rep-track { height:7px; background: rgba(255,255,255,0.05); border-radius:4px; overflow:hidden; display:flex; }
        .rep-done { background: linear-gradient(90deg, var(--accent-teal), #22c1b0); height:100%; }
        .rep-pending { background: rgba(255,255,255,0.05); height:100%; }
        .doc-tag { font-size:9px; font-weight:600; color: var(--text-3); background: rgba(255,255,255,0.06); padding:1px 6px; border-radius:4px; text-transform:uppercase; letter-spacing:0.3px; margin-left:6px; }

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


def kpi_card(label: str, value: str, sub: str, pct: float, color: str) -> str:
    pct_display = f"{pct * 100:.1f}%".replace(".", ",")
    width = max(0.0, min(pct, 1.0)) * 100
    return f"""
        <div class="kpi-card" style="--kpi-accent:{color};">
          <div class="kpi-label">{label}</div>
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
        ("Total de tags", f"{total_tags:,}".replace(",", "."), "Base principal", 1.0, "#5b8def"),
        ("Tags completas", f"{tags_completas:,}".replace(",", "."), "Pronta documental",
         (tags_completas / total_tags) if total_tags else 0, "#34d399"),
        ("Pendentes", f"{total_pendentes:,}".replace(",", "."), "Aguardando entrega",
         (total_pendentes / total_esperados) if total_esperados else 0, "#f87171"),
        ("Emitidos SIGEM", f"{total_emitidos:,}".replace(",", "."), "Com status localizado",
         avanco_geral, "#fbbf24"),
        ("Avanço geral", f"{avanco_geral * 100:.1f}%".replace(".", ","), "Progressão do projeto",
         avanco_geral, "#9d6bff"),
    ]
    for col, (label, value, sub, pct, color) in zip(cols, kpis):
        with col:
            render_html(kpi_card(label, value, sub, pct, color))

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

        fig = go.Figure(
            data=[go.Pie(labels=labels, values=values, hole=0.62, marker=dict(colors=colors, line=dict(color="#12172a", width=2)), textinfo="none")]
        )
        fig.update_layout(
            showlegend=True,
            legend=dict(orientation="v", font=dict(color="#9aa4bc", size=11), bgcolor="rgba(0,0,0,0)"),
            margin=dict(l=0, r=0, t=0, b=0),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=280,
            annotations=[dict(text=f"{len(esperados):,}".replace(",", ".") + "<br><span style='font-size:11px;color:#6b7590'>Relatórios</span>",
                               x=0.5, y=0.5, font=dict(size=22, color="#f4f6fb"), showarrow=False)],
        )
        render_html('<div class="gplan-panel"><div class="gplan-panel-title">Status SIGEM</div>')
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        render_html("</div>")

    st.write("")
    grouped = resumo.groupby("GRUPO_REGRA").agg(
        tags=("TAG", "count"),
        esperados=("RELATORIOS_ESPERADOS", "sum"),
        emitidos=("RELATORIOS_POSTADOS", "sum"),
    ).reset_index()
    grouped["avanco"] = (grouped["emitidos"] / grouped["esperados"]).fillna(0) * 100
    grouped = grouped.sort_values("tags", ascending=False)

    render_html('<div class="gplan-panel"><div class="gplan-panel-title">Resumo por grupo</div>')
    group_cols = st.columns(len(grouped))
    for col, (_, row) in zip(group_cols, grouped.iterrows()):
        with col:
            render_html(
                f"""
                <div class="group-card-v2">
                  <div class="group-card-top"><span class="group-card-name">{row['GRUPO_REGRA'].title()}</span><span class="group-card-pct">{row['avanco']:.1f}%</span></div>
                  <div class="group-card-value">{int(row['tags']):,} <span class="group-card-unit">tags</span></div>
                  <div class="group-card-track"><div class="group-card-fill" style="width:{row['avanco']:.1f}%;"></div></div>
                  <div class="group-card-nums"><span>{int(row['emitidos']):,} emitidos</span><span>{int(row['esperados']):,} esperados</span></div>
                </div>
                """.replace(",", ".")
            )
    render_html("</div>")

    st.write("")
    top10 = resumo.sort_values("RELATORIOS_PENDENTES", ascending=False).head(10)
    top10_display = top10[["TAG", "DESCRICAO", "GRUPO_REGRA", "RELATORIOS_ESPERADOS", "RELATORIOS_POSTADOS", "RELATORIOS_PENDENTES", "AVANCO_DOCUMENTAL"]].copy()
    top10_display["AVANCO_DOCUMENTAL"] = (top10_display["AVANCO_DOCUMENTAL"] * 100).round(1)
    top10_display.columns = ["Tag", "Descrição", "Grupo", "Esperados", "Emitidos", "Pendentes", "Avanço (%)"]

    render_html('<div class="gplan-panel"><div class="gplan-panel-title">Top 10 tags com mais pendências</div>')
    st.dataframe(top10_display, hide_index=True, use_container_width=True)
    render_html("</div>")


def render_relatorios(esperados: pd.DataFrame):
    render_header("Relatórios previstos")

    display_cols = ["TAG", "DESCRICAO", "GRUPO", "RELATORIO", "REFERENCIA", "DOCUMENTO_ESPERADO",
                     "EXISTE_NO_SIGEM", "STATUS_SIGEM", "REVISAO_SIGEM", "DATA_SIGEM"]
    df = esperados[display_cols].copy()
    df["STATUS_SIGEM"] = df["STATUS_SIGEM"].map(sentence_case)
    df["EXISTE_NO_SIGEM"] = df["EXISTE_NO_SIGEM"].map({"SIM": "Sim", "NAO": "Não"}).fillna(df["EXISTE_NO_SIGEM"])
    df["REVISAO_SIGEM"] = df["REVISAO_SIGEM"].apply(format_missing)
    df["DATA_SIGEM"] = df["DATA_SIGEM"].apply(format_date)
    df.columns = ["Tag", "Descrição", "Grupo", "Relatório", "Referência", "Documento esperado",
                  "Existe no SIGEM", "Status SIGEM", "Revisão", "Data"]

    search = st.text_input("Pesquisar", placeholder="Pesquisar por tag, descrição, relatório, documento, status...", label_visibility="collapsed")
    if search:
        mask = df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)
        df = df[mask]

    df_page = paginate(df, "relatorios", search)
    st.dataframe(df_page, hide_index=True, use_container_width=True, height=560)


def render_pesquisa_tag(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame):
    render_header("Pesquisa tag")

    if "gplan_selected_tag" not in st.session_state:
        st.session_state.gplan_selected_tag = None

    if st.session_state.gplan_selected_tag:
        tag_id = st.session_state.gplan_selected_tag
        resumo_row = resumo[resumo["TAG"] == tag_id]
        tags_row = tags[tags["TAG"] == tag_id]
        if resumo_row.empty:
            st.session_state.gplan_selected_tag = None
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

        st.markdown("**Relatórios da tag**")
        tag_reports = esperados[esperados["TAG"] == tag_id][
            ["RELATORIO", "REFERENCIA", "DOCUMENTO_ESPERADO", "STATUS_SIGEM", "REVISAO_SIGEM"]
        ].copy()
        tag_reports["STATUS_SIGEM"] = tag_reports["STATUS_SIGEM"].map(sentence_case)
        tag_reports["REVISAO_SIGEM"] = tag_reports["REVISAO_SIGEM"].apply(format_missing)
        tag_reports.columns = ["Relatório", "Referência", "Documento esperado", "Status SIGEM", "Revisão"]
        st.dataframe(tag_reports, hide_index=True, use_container_width=True)
        return

    search = st.text_input("Pesquisar", placeholder="Digite a tag para ver a ficha completa (ex: AIT-120005)...", label_visibility="collapsed")

    list_df = resumo[["TAG", "DESCRICAO", "GRUPO_REGRA", "ITEM_PPU", "RELATORIOS_ESPERADOS", "AVANCO_DOCUMENTAL"]].copy()
    if search:
        mask = list_df["TAG"].astype(str).str.contains(search, case=False, na=False) | list_df["DESCRICAO"].astype(str).str.contains(search, case=False, na=False)
        list_df = list_df[mask]

    list_df["AVANCO_DOCUMENTAL"] = (list_df["AVANCO_DOCUMENTAL"] * 100).round(1)
    list_df.columns = ["Tag", "Descrição", "Tipo", "PPU", "Relatórios", "Avanço (%)"]

    list_df_page = paginate(list_df, "pesquisa", search)
    event = st.dataframe(
        list_df_page, hide_index=True, use_container_width=True, height=560,
        on_select="rerun", selection_mode="single-row",
    )
    if event.selection.rows:
        selected_row = list_df_page.iloc[event.selection.rows[0]]
        st.session_state.gplan_selected_tag = selected_row["Tag"]
        st.rerun()


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
        mask = df.apply(lambda row: row.astype(str).str.contains(text_search, case=False, na=False).any(), axis=1)
        df = df[mask]

    df = df.copy()
    df["_REVISAO_SORT"] = df["REVISAO"].astype(str)
    df["_DOCUMENTO_SORT"] = df["DOCUMENTO"].astype(str)
    df = df.sort_values(["_REVISAO_SORT", "_DOCUMENTO_SORT"]).drop(columns=["_REVISAO_SORT", "_DOCUMENTO_SORT"])

    df["DATA"] = df["DATA"].apply(format_date)
    df.columns = ["Documento", "Status", "Revisão", "Data", "Título"]

    df_page = paginate(df, "sigem", f"{status_filter}|{text_search}")
    st.dataframe(df_page, hide_index=True, use_container_width=True, height=560)


def main():
    st.set_page_config(page_title="Gplan", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
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
            '<div style="padding: 0 4px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); margin-bottom: 8px;">'
            '<div style="font-size:15px; font-weight:700; color:#f4f6fb;">Gplan</div>'
            '<div style="font-size:11.5px; color:#6b7590; margin-top:2px;">Instrumentação · U-12</div>'
            "</div>"
        )

    dashboard_page = st.Page(lambda: render_dashboard(resumo, esperados), title="Dashboard", url_path="dashboard", default=True)
    relatorios_page = st.Page(lambda: render_relatorios(esperados), title="Relatórios", url_path="relatorios")
    pesquisa_page = st.Page(lambda: render_pesquisa_tag(resumo, esperados, tags), title="Pesquisa tag", url_path="pesquisa")
    sigem_page = st.Page(lambda: render_sigem(sigem), title="Base SIGEM", url_path="sigem")

    nav = st.navigation([dashboard_page, relatorios_page, pesquisa_page, sigem_page], position="sidebar")
    nav.run()


if __name__ == "__main__":
    main()
