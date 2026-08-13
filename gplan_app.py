import base64
import io
import json
import math
import os
import time
from contextlib import contextmanager
from urllib.parse import quote

import numpy as np
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


# A marca segue o tema pelas variaveis: o anel de fundo e o ponteiro sao os
# que sumiriam -- anel escuro sobre fundo claro, ponteiro claro sobre card
# branco. O arco e o miolo ficam nas cores da marca nos dois temas.
def _logo_svg(sufixo: str = "") -> str:
    grad = f"gpArc{sufixo}"
    return (
        '<svg viewBox="0 0 48 48" fill="none">'
        '<circle cx="24" cy="24" r="19" stroke="var(--neutro)" stroke-width="5"/>'
        f'<path d="M24 5a19 19 0 0 1 15.6 29.8" stroke="url(#{grad})" stroke-width="5" stroke-linecap="round"/>'
        '<circle cx="24" cy="24" r="5.5" fill="var(--accent-teal)"/>'
        '<path d="M24 24L33 15" stroke="var(--text-1)" stroke-width="3" stroke-linecap="round"/>'
        f'<defs><linearGradient id="{grad}" x1="24" y1="5" x2="40" y2="35" gradientUnits="userSpaceOnUse">'
        '<stop stop-color="var(--accent-blue)"/><stop offset="1" stop-color="var(--accent-teal)"/>'
        "</linearGradient></defs></svg>"
    )


LOGO_SVG = _logo_svg()


# A transicao inteira leva isso, mesmo quando o trabalho acaba antes. Sem o
# minimo a tela pisca e some no mesmo movimento, o que se le como falha e nao
# como carregamento -- e as abas em cache ficam prontas em menos de 100 ms.
CARGA_MINIMA = 0.85


def tela_carregando(texto: str, pct: int | None = None, coberta: bool = True,
                    vidro: bool = False, saindo: bool = False) -> str:
    """Carregando com a marca. Sem pct a barra corre sozinha, indeterminada.

    So passar numero quando ele significar alguma coisa: uma porcentagem
    inventada mente sobre quanto falta, e a aba Progresso demora o bastante
    para isso irritar.

    `vidro` desfoca o que esta atras em vez de tapar com cor solida: serve
    para filtro e troca de aba, onde ja existe conteudo na tela e sumir com
    ele daria a sensacao de recomecar do zero. Fundo solido fica para a
    primeira abertura, quando nao ha nada atras mesmo.
    """
    if pct is None:
        barra = '<div class="gpl-fill gpl-indet"></div>'
    else:
        barra = f'<div class="gpl-fill" style="width:{pct}%;"></div>'
    classes = "gpl"
    if coberta:
        classes += " gpl-vidro" if vidro else " gpl-cheia"
    if saindo:
        classes += " gpl-saindo"
    return (
        f'<div class="{classes}">'
        '<div class="gpl-corpo">'
        f'<div class="gpl-mark">{LOGO_SVG}</div>'
        '<div class="gpl-nome">Gplan</div>'
        f'<div class="gpl-txt">{esc(texto)}</div>'
        f'<div class="gpl-track">{barra}</div></div></div>'
    )


def lembrado(widget, chave: str, *args, **kwargs):
    """Filtro que sobrevive a troca de aba.

    O Streamlit descarta o estado de um widget quando a pagina que o criou sai
    de cena: filtrar na Pesquisa tag, ir para a Base SIGEM e voltar devolvia o
    campo vazio. Aqui o valor tambem fica numa chave propria, que a navegacao
    nao limpa, e o widget e reconstruido a partir dela.
    """
    guardado = st.session_state.get(f"mem_{chave}")
    if guardado is not None:
        # args = (rotulo, opcoes): a lista e o segundo, nao o primeiro
        opcoes = list(args[1]) if len(args) > 1 else list(kwargs.get("options", []))
        if widget is st.selectbox:
            if guardado in opcoes:
                kwargs["index"] = opcoes.index(guardado)
        elif widget is st.multiselect:
            kwargs["default"] = [v for v in guardado if v in opcoes]
        else:
            kwargs["value"] = guardado
    valor = widget(*args, key=chave, **kwargs)
    st.session_state[f"mem_{chave}"] = valor
    return valor


def render_html_pesado(html: str):
    """Para os blocos grandes: o st.markdown passa a string inteira por um
    parser de Markdown antes de virar HTML, e com os 24,8 MB da arvore da aba
    Progresso a pagina simplesmente nunca terminava de abrir. O st.html insere
    direto -- 13 s no lugar de mais de 10 minutos.

    So serve para bloco sem <style> e sem <svg>: o st.html descarta os dois.
    O donut do Status SIGEM e os icones dos cards, por exemplo, precisam
    continuar no render_html.
    """
    st.html(html)


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

# O status vira classe, e a cor sai do tema. Guardar o hex e pinta-lo no
# atributo style prendia o badge ao tema escuro: no claro o teal #2dd4bf sobre
# a propria lavagem clara dava 1,7:1, e "Sem comentarios" sumia da tabela.
STATUS_CLASSE = {
    "Não postado": "mudo", "Sem comentários": "ok", "Com comentários": "andamento",
    "Para construção": "warn", "Recusado": "crit", "Em análise": "roxo",
    "Cancelado": "cinza",
}


def classe_status(label: str) -> str:
    return STATUS_CLASSE.get(label, "mudo")

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
        # regex=False: o usuario digita codigo de documento, e um "." ali e um
        # ponto literal. Sem isso, buscar C1N_..._3.1.1.1_... casava com 13
        # linhas em vez de uma, porque cada ponto virava "qualquer caractere".
        mask = mask | df[col].astype(str).str.contains(text, case=False, na=False, regex=False)
    return mask


def get_supabase_client():
    # GPLAN_LOCAL=1 le a planilha do disco e ignora o Supabase, sem precisar
    # esconder o secrets.toml. Serve para testar local: e mais rapido e nao
    # esbarra no proxy da rede, que derruba o SSL da chamada ao Supabase.
    if os.environ.get("GPLAN_LOCAL") == "1":
        return None
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
    # A medicao de campo so existe em planilha gerada pelo pipeline novo; a
    # antiga que estiver no Supabase continua abrindo, com a aba vazia.
    if "06_BASE_GITEC" in excel_file.sheet_names:
        gitec = pd.read_excel(excel_file, sheet_name="06_BASE_GITEC")
    else:
        gitec = pd.DataFrame(columns=["TAG", "ITEM_PPU_GITEC", "FASE", "AGRUPAMENTO",
                                      "ETAPA", "STATUS", "VALOR", "DATA_EXECUCAO"])
    # A area de cada TAG e o mapa area<->desenho so existem em planilha gerada
    # pelo pipeline novo. Sem eles a aba Planta abre explicando o que falta, em
    # vez de quebrar.
    locacao = (pd.read_excel(excel_file, sheet_name="05_BASE_LOCAÇÃO")
               if "05_BASE_LOCAÇÃO" in excel_file.sheet_names
               else pd.DataFrame(columns=["TAG", "AREA"]))
    if "AREA" not in locacao.columns:
        locacao["AREA"] = pd.NA
    aux_areas = (pd.read_excel(excel_file, sheet_name="05_AUX_AREAS")
                 if "05_AUX_AREAS" in excel_file.sheet_names
                 else pd.DataFrame(columns=["AREA", "NOME_AREA", "DESENHO",
                                            "DESENHO_GERAL"]))
    resumo = aplicar_regra_aprovados(resumo, esperados)
    return tags, cabos, tubing, sigem, resumo, esperados, gitec, locacao, aux_areas


# Um relatorio so conta como avanco depois de aprovado pela fiscalizacao.
# Postado nao basta: recusado, em analise e cancelado seguem sendo pendencia.
STATUS_APROVADOS = {"SEM COMENTÁRIOS", "COM COMENTÁRIOS", "PARA CONSTRUÇÃO"}

# Entra na chave do cache. O HTML da arvore e das fichas fica guardado por
# (planilha, filtros), e o Streamlit nao percebe mudanca numa funcao chamada
# por dentro -- mudei a regra de avanco e a arvore continuou servindo o numero
# velho. Subir esse numero ao mexer em como o avanco e calculado.
REGRA_VERSAO = 9
# O cache do Streamlit hasheia o corpo da funcao cacheada, nao o das funcoes que
# ela chama. As fichas sao montadas dentro de funcoes cacheadas, entao mudar o
# desenho delas nao invalida nada: a aba Progresso continuava servindo o HTML
# antigo, e a economia do sprite simplesmente nao aparecia. Subir este numero e
# o que diz ao cache que o desenho mudou.
VISUAL_VERSAO = 8


def aprovado(serie: pd.Series) -> pd.Series:
    return serie.astype(str).str.strip().str.upper().isin(STATUS_APROVADOS)


def aplicar_regra_aprovados(resumo: pd.DataFrame, esperados: pd.DataFrame) -> pd.DataFrame:
    """Recalcula o avanco por TAG contando so relatorio aprovado.

    A conta sai da 08_RELATORIOS_ESPERADOS, que tem o status de cada relatorio,
    e nao da coluna ja pronta da 07_TAG_RESUMO: assim o app nao pode divergir
    da base, mesmo se a planilha for gerada por um pipeline mais antigo.

    RELATORIOS_POSTADOS continua sendo o que existe no SIGEM em qualquer
    status. A diferenca entre postados e aprovados e justamente a fila de
    pendencia -- 722 recusados, 510 em analise, 49 cancelados na base de hoje.
    """
    e = esperados.copy()
    e["_postado"] = e["EXISTE_NO_SIGEM"].astype(str).str.strip().str.upper().eq("SIM")
    e["_aprovado"] = aprovado(e["STATUS_SIGEM"])
    por_tag = e.groupby("TAG").agg(
        _esperados=("TAG", "size"),
        _postados=("_postado", "sum"),
        _aprovados=("_aprovado", "sum"),
    )

    df = resumo.merge(por_tag, left_on="TAG", right_index=True, how="left")
    for c in ("_esperados", "_postados", "_aprovados"):
        df[c] = df[c].fillna(0).astype(int)
    df["RELATORIOS_ESPERADOS"] = df["_esperados"]
    df["RELATORIOS_POSTADOS"] = df["_postados"]
    df["RELATORIOS_APROVADOS"] = df["_aprovados"]
    df["RELATORIOS_PENDENTES"] = df["_esperados"] - df["_aprovados"]
    df["AVANCO_DOCUMENTAL"] = (df["_aprovados"] / df["_esperados"]).fillna(0.0)
    return df.drop(columns=["_esperados", "_postados", "_aprovados"])


def faixa_resumo(itens: list) -> str:
    """Resumo do recorte, acima da tabela.

    As abas de lista eram filtro e tabela, nada mais: chegava-se a 25.100
    linhas sem ideia da proporcao entre elas. A faixa responde antes de
    filtrar. `itens` sao (rotulo, valor, tom) -- tom pinta so o que pede
    atencao, senao vira semaforo e nada se destaca.
    """
    celulas = "".join(
        f'<div class="fx-item"><span class="fx-lbl">{esc(l)}</span>'
        f'<span class="fx-val {t or ""}">{esc(v)}</span></div>'
        for l, v, t in itens
    )
    return f'<div class="gplan-panel fx-faixa">{celulas}</div>'


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
            f'<div class="gtbl-pag">'
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
        # aprovado, nao apenas postado: a barra do Dashboard tem que contar a
        # mesma coisa que o avanco da TAG, senao as duas telas se contradizem
        subset = subset[aprovado(subset["STATUS_SIGEM"])]
    if unique_doc:
        return int(subset["DOCUMENTO_ESPERADO"].nunique())
    return int(len(subset))


# ==========================================================================
#  Tema: um único lugar decide toda cor da interface
# ==========================================================================
# Nada no CSS abaixo escreve uma cor literal. Tudo consome uma destas
# variáveis, e trocar de tema é trocar este dicionário -- inclusive as
# lavagens (rgba sobre a superfície) e os tons de texto sobre chip colorido,
# que são justamente os que somem quando alguém só inverte fundo e letra.
#
# O tema claro não é o escuro invertido: o fundo é um cinza levemente azulado,
# os cartões são brancos por cima dele, e as cores semânticas foram escurecidas
# até terem contraste de leitura sobre branco -- o âmbar #fbbf24 sobre branco
# dá 1,7:1, ilegível, e vira #a16207.
TEMAS = {
    "escuro": {
        "bg": "#0a0e1a", "card": "#12172a", "card2": "#171d33", "fundo3": "#080c16",
        "borda": "rgba(255,255,255,0.06)", "borda_forte": "rgba(255,255,255,0.12)",
        # componente das lavagens: no escuro clareia a superfície, no claro escurece
        "rgb_tinta": "255,255,255",
        "texto1": "#f4f6fb", "texto2": "#a6b0c6", "texto3": "#828da8",
        "neutro": "#3a4a68", "neutro2": "#4b5468",
        "overlay": "rgba(6,9,18,0.68)", "sombra": "rgba(0,0,0,0.6)",
        "sombra_leve": "rgba(0,0,0,0.35)",
        "azul": "#5b8def", "roxo": "#9d6bff", "verde": "#34d399",
        "vermelho": "#f87171", "ambar": "#fbbf24", "teal": "#2dd4bf",
        "rgb_azul": "91,141,239", "rgb_roxo": "157,107,255", "rgb_verde": "52,211,153",
        "rgb_vermelho": "248,113,113", "rgb_ambar": "251,191,36", "rgb_teal": "45,212,191",
        # texto sobre chip da própria cor -- clareado no escuro, escurecido no claro
        "txt_azul": "#a9c5ff", "txt_roxo": "#c9b6ff", "txt_verde": "#6ee7d0",
        "txt_vermelho": "#fca5a5", "txt_ambar": "#fcd34d", "txt_teal": "#6ee7d0",
        "sobre_cor": "#0a0e1a", "teal2": "#22c1b0",
        # o desenho da planta vem preto sobre branco: invertido, o traço fica
        # claro sobre o escuro e a prancha para de ser um retângulo branco
        "rgb_chapa": "8,12,22",
        "planta_filtro": "invert(1) brightness(.86) contrast(1.22)",
        "planta_opacidade": ".52",
        "esquema": "dark",
    },
    "claro": {
        "bg": "#eef1f7", "card": "#ffffff", "card2": "#f3f5fa", "fundo3": "#e4e8f1",
        "borda": "rgba(17,26,48,0.10)", "borda_forte": "rgba(17,26,48,0.20)",
        "rgb_tinta": "17,26,48",
        "texto1": "#111a30", "texto2": "#43506c", "texto3": "#5c6884",
        "neutro": "#aeb9cc", "neutro2": "#95a1b8",
        "overlay": "rgba(23,32,54,0.42)", "sombra": "rgba(23,32,54,0.18)",
        "sombra_leve": "rgba(23,32,54,0.10)",
        "azul": "#2f63d4", "roxo": "#7440cf", "verde": "#0f8f5f",
        "vermelho": "#cf3b3b", "ambar": "#a16207", "teal": "#0b8478",
        "rgb_azul": "47,99,212", "rgb_roxo": "116,64,207", "rgb_verde": "15,143,95",
        "rgb_vermelho": "207,59,59", "rgb_ambar": "161,98,7", "rgb_teal": "11,132,120",
        "txt_azul": "#1e4aa8", "txt_roxo": "#5b2fa8", "txt_verde": "#0a6c48",
        "txt_vermelho": "#a72c2c", "txt_ambar": "#7c4a05", "txt_teal": "#08655c",
        "sobre_cor": "#ffffff", "teal2": "#0a6f65",
        # no claro o desenho já é escuro sobre branco: inverter deixaria a
        # prancha preta no meio de uma tela clara
        "rgb_chapa": "255,255,255",
        "planta_filtro": "grayscale(1) contrast(1.35) brightness(.82)",
        "planta_opacidade": ".55",
        "esquema": "light",
    },
}
TEMA_PADRAO = "escuro"


def tema_ativo() -> str:
    """O tema escolhido, lido antes de qualquer CSS sair.

    A preferência vive na URL, o mesmo mecanismo que já guarda os filtros. É o
    que sobrevive a fechar e reabrir, e é lido no primeiro instante da execução
    -- ler depois faria a tela nascer num tema e trocar no meio do carregamento.
    """
    escolhido = st.session_state.get("gplan_tema")
    if escolhido not in TEMAS:
        escolhido = st.query_params.get("tema")
    if escolhido not in TEMAS:
        escolhido = TEMA_PADRAO
    st.session_state["gplan_tema"] = escolhido
    return escolhido


def tokens_css(tema: str) -> str:
    t = TEMAS.get(tema, TEMAS[TEMA_PADRAO])
    return f"""
          color-scheme: {t['esquema']};
          --dark-bg: {t['bg']};
          --dark-card: {t['card']};
          --dark-card-2: {t['card2']};
          --fundo-3: {t['fundo3']};
          --border-color: {t['borda']};
          --border-strong: {t['borda_forte']};
          --rgb-tinta: {t['rgb_tinta']};
          --overlay: {t['overlay']};
          --sombra: {t['sombra']};
          --sombra-leve: {t['sombra_leve']};
          --text-1: {t['texto1']};
          --text-2: {t['texto2']};
          --text-3: {t['texto3']};
          --neutro: {t['neutro']};
          --neutro-2: {t['neutro2']};
          --accent-blue: {t['azul']};
          --accent-purple: {t['roxo']};
          --accent-green: {t['verde']};
          --accent-red: {t['vermelho']};
          --accent-amber: {t['ambar']};
          --accent-teal: {t['teal']};
          --rgb-azul: {t['rgb_azul']};
          --rgb-roxo: {t['rgb_roxo']};
          --rgb-verde: {t['rgb_verde']};
          --rgb-vermelho: {t['rgb_vermelho']};
          --rgb-ambar: {t['rgb_ambar']};
          --rgb-teal: {t['rgb_teal']};
          --txt-azul: {t['txt_azul']};
          --txt-roxo: {t['txt_roxo']};
          --txt-verde: {t['txt_verde']};
          --txt-vermelho: {t['txt_vermelho']};
          --txt-ambar: {t['txt_ambar']};
          --txt-teal: {t['txt_teal']};
          --sobre-cor: {t['sobre_cor']};
          --teal-2: {t['teal2']};
          --rgb-chapa: {t['rgb_chapa']};
          --planta-filtro: {t['planta_filtro']};
          --planta-opacidade: {t['planta_opacidade']};"""


def trocar_tema():
    """Guarda a escolha na sessao e na URL, e deixa o Streamlit redesenhar.

    Nao ha recarga de pagina: o Streamlit reexecuta o script, o inject_css sai
    com o outro :root e a arvore inteira -- cartoes, tabelas, graficos, icones,
    modais -- muda junto, porque todos leem as mesmas variaveis. A URL guarda a
    preferencia para a proxima abertura.
    """
    novo = "claro" if st.session_state.get("gplan_tema") == "escuro" else "escuro"
    st.session_state["gplan_tema"] = novo
    st.query_params["tema"] = novo


def seletor_tema():
    """O botao de tema: so o icone, preso no alto a direita.

    Fica fora do fluxo por CSS. Precisa ser um st.button de verdade -- um link
    no cabecalho trocaria o tema recarregando a pagina inteira, e a planilha
    leva segundos para voltar. Assim o Streamlit so reexecuta o script.
    """
    claro = tema_ativo() == "claro"
    st.button("Tema", key="gplan_btn_tema", on_click=trocar_tema,
              icon=":material/dark_mode:" if claro else ":material/light_mode:",
              help="Mudar para o tema escuro" if claro else "Mudar para o tema claro")


def inject_css():
    render_html(
        """
        <style>
        :root {__TOKENS__
        }

        .stApp { background: var(--dark-bg); }
        /* Streamlit usa 96px topo / 80px lateral / 160px rodape por padrao,
           o que deixa uma faixa vazia enorme no fim da pagina e desperdica largura. */
        [data-testid="stMainBlockContainer"], .block-container {
          padding: 32px 44px 48px !important;
          max-width: 1600px !important;
        }
        [data-testid="stHeader"] { background: transparent !important; }
        /* O indicador de execucao do Streamlit -- "Running", "Stop", "Rerun" --
           nasce com a cor do config e some no tema claro. Fica sobre o
           conteudo, entao leva superficie propria para nao se perder nele. */
        [data-testid="stStatusWidget"] { background: var(--dark-card) !important;
          border: 1px solid var(--border-strong) !important; border-radius: 10px;
          box-shadow: 0 4px 14px var(--sombra-leve); }
        [data-testid="stStatusWidget"] *, [data-testid="stStatusWidget"] button {
          color: var(--text-1) !important; }
        [data-testid="stStatusWidget"] svg, [data-testid="stStatusWidget"] svg * {
          fill: var(--text-1) !important; stroke: var(--text-1) !important; }
        section[data-testid="stSidebar"] {
          background: linear-gradient(180deg, var(--fundo-3) 0%, var(--dark-bg) 100%);
          border-right: 1px solid var(--border-color);
          width: 212px !important; min-width: 212px !important;
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
        /* Menu abaixo dos filtros. O dropdown do selectbox e portalado no
           body com position:fixed e nao vira para cima quando falta espaco:
           com os cinco itens de menu antes, o ultimo filtro comecava a 617px
           numa janela de 650 e a lista abria atras da barra de tarefas. Com
           os filtros no topo o ultimo termina por volta de 330px, e os 300px
           da lista cabem. */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] { order: 4; }
        section[data-testid="stSidebar"] [data-testid="stSidebarNavSeparator"] { display: none; }

        .gplan-brand {
          display:flex; align-items:center; gap:11px; padding: 0 4px 32px;
          border-bottom: 1px solid var(--border-color);
        }
        .gplan-brand-mark { width:34px; height:34px; flex-shrink:0; }
        .gplan-brand-name { font-size:16px; font-weight:800; color:var(--text-1); letter-spacing:-0.3px; line-height:1.1; }
        .gplan-brand-sub { font-size:10.5px; color:var(--text-3); margin-top:3px;
                           white-space:nowrap; }
        .gplan-brand { padding-bottom: 11px !important; }
        .gplan-brand-mark { width:30px !important; height:30px !important; }

        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a,
        section[data-testid="stSidebar"] nav a {
          border-radius: 10px !important;
          color: var(--text-2) !important;
          font-size: 13.5px !important;
          font-weight: 500 !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover,
        section[data-testid="stSidebar"] nav a:hover {
          background: rgba(var(--rgb-tinta),0.05) !important;
          color: var(--text-1) !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"],
        section[data-testid="stSidebar"] nav a[aria-current="page"] {
          background: linear-gradient(135deg, rgba(var(--rgb-azul),0.18), rgba(var(--rgb-roxo),0.12)) !important;
          color: var(--text-1) !important;
          box-shadow: inset 0 0 0 1px rgba(var(--rgb-azul),0.3);
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] span,
        section[data-testid="stSidebar"] nav a[aria-current="page"] span {
          color: var(--text-1) !important;
        }


        /* =========================================================
           Dashboard em tela unica. So vale nesta pagina: o :has()
           casa apenas quando .du-tela esta no DOM, entao as outras
           abas continuam com o padding e a rolagem de sempre.
           ========================================================= */
        .stApp:has(.du-tela) [data-testid="stMain"] { overflow: hidden !important; }
        .stApp:has(.du-tela) [data-testid="stHeader"] { height: 2.1rem !important; }
        .stApp:has(.du-tela) [data-testid="stMainBlockContainer"],
        .stApp:has(.du-tela) .block-container {
          box-sizing: border-box; height: 100dvh;
          display: flex !important; flex-direction: column;
          padding: 2.4rem 18px 12px !important; max-width: none !important;
        }
        .stApp:has(.du-tela) [data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] {
          flex: 1; min-height: 0;
        }
        /* O bloco de CSS injetado vira um container vazio de 16px + gap no meio
           da coluna. Some com ele -- e so <style>, nao tem nada para mostrar. */
        .stApp:has(.du-tela) [data-testid="stElementContainer"]:has([data-testid="stMarkdownContainer"] > style) {
          display: none !important;
        }
        .stApp:has(.du-tela) [data-testid="stElementContainer"]:has(.du-tela) {
          flex: 1; min-height: 0; display: flex; flex-direction: column;
        }
        /* Todo ancestral entre o container e a tela, e nao so os que tem
           data-testid: existe um <div> sem nome no meio, e era ele que ficava
           com min-height:auto e segurava a altura em 950px numa janela de 900. */
        .stApp:has(.du-tela) [data-testid="stElementContainer"]:has(.du-tela) *:has(.du-tela) {
          flex: 1; min-height: 0; display: flex; flex-direction: column;
          margin-bottom: 0 !important;
        }
        .du-tela {
          flex: 1; min-height: 0;
          display: grid; grid-template-rows: auto auto 1fr auto;
          gap: clamp(8px, 1.15vh, 13px);
          font-variant-numeric: tabular-nums;
        }
        .du-cab { display:flex; align-items:flex-end; justify-content:space-between; gap:14px; }
        .du-cab .du-h1 { font-size: clamp(17px,2.3vh,23px); font-weight:700; letter-spacing:-.3px;
                     margin:0; line-height:1.1; color:var(--text-1); }
        /* O markdown do Streamlit sublinha todo <a> e pinta de azul; aqui cada
           link e um cartao, uma barra ou uma linha de tabela. */
        .du-tela a, .du-tela a:hover, .du-tela a:visited {
          text-decoration: none !important; color: inherit; }
        .du-cab p { font-size:11.5px; color:var(--text-3); margin:3px 0 0; }
        .du-acoes { display:flex; align-items:center; gap:8px; }
        .du-selo { display:flex; align-items:center; gap:7px; background:var(--dark-card);
                   border:1px solid var(--border-color); border-radius:100px; padding:6px 13px;
                   font-size:11px; color:var(--text-2); white-space:nowrap; }
        .du-selo i { width:6px; height:6px; border-radius:50%; background:var(--accent-green);
                     box-shadow:0 0 8px var(--accent-green); }
        .du-selo.filtro { border-color:rgba(var(--rgb-ambar),.45); color:var(--text-1); }
        .du-selo.filtro i { background:var(--accent-amber); box-shadow:0 0 8px var(--accent-amber); }

        .du-kpis { display:grid; grid-template-columns:repeat(5,1fr); gap:clamp(8px,.85vw,13px); }
        .du-kpi { background:var(--dark-card); border:1px solid var(--border-color); border-radius:13px;
                  padding:clamp(9px,1.35vh,14px) clamp(11px,.85vw,15px);
                  display:flex; flex-direction:column; gap:6px; text-decoration:none; }
        a.du-kpi:hover { border-color:rgba(var(--rgb-azul),.45); }
        .du-kpi .topo { display:flex; align-items:center; gap:9px; }
        .du-kpi .tile { width:clamp(26px,3.2vh,33px); height:clamp(26px,3.2vh,33px); border-radius:9px;
                        flex:none; display:flex; align-items:center; justify-content:center; }
        .du-kpi .tile svg { width:15px; height:15px; }
        .du-kpi .rot { font-size:11px; color:var(--text-2); font-weight:600; }
        .du-kpi .val { font-size:clamp(21px,3.4vh,31px); font-weight:800; letter-spacing:-.8px;
                       line-height:1; color:var(--text-1); }
        .du-kpi .sub { font-size:10px; color:var(--text-3); }
        .du-trilho { height:3px; border-radius:99px; background:rgba(var(--rgb-tinta),.07);
                     overflow:hidden; color:var(--accent-green); }
        .du-trilho i { display:block; height:100%; border-radius:99px; background:currentColor; }
        /* o quadradinho do icone: cor pela classe, fundo na lavagem dela */
        .tile.fxc-azul  { color:var(--accent-blue);   background:rgba(var(--rgb-azul),.13); }
        .tile.fxc-teal  { color:var(--accent-teal);   background:rgba(var(--rgb-teal),.13); }
        .tile.fxc-roxo  { color:var(--accent-purple); background:rgba(var(--rgb-roxo),.13); }
        .tile.fxc-ambar { color:var(--accent-amber);  background:rgba(var(--rgb-ambar),.13); }
        .tile.fxc-rubi  { color:var(--accent-red);    background:rgba(var(--rgb-vermelho),.13); }
        .tile.fxc-verde { color:var(--accent-green);  background:rgba(var(--rgb-verde),.13); }
        .tile.fxc-mudo  { color:var(--text-2);        background:rgba(var(--rgb-tinta),.10); }
        .tile.fxc-cinza { color:var(--neutro);        background:rgba(var(--rgb-tinta),.08); }

        .du-meio { display:grid; grid-template-columns:1.58fr 1fr; gap:clamp(8px,.85vw,13px); min-height:0; }
        .du-col { display:grid; grid-template-rows:auto 1fr; gap:clamp(8px,1.15vh,13px); min-height:0; }
        .du-pn { background:var(--dark-card); border:1px solid var(--border-color); border-radius:13px;
                 display:flex; flex-direction:column; min-height:0; overflow:hidden; }
        .du-pn > .du-t { font-size:12px; font-weight:700; color:var(--text-1); flex:none;
                      padding:clamp(9px,1.2vh,13px) clamp(11px,.9vw,15px) clamp(6px,.85vh,9px); }
        .du-miolo { flex:1; min-height:0; display:flex; flex-direction:column;
                    padding:0 clamp(11px,.9vw,15px) clamp(9px,1.2vh,13px); }
        .du-rodape { flex:none; border-top:1px solid var(--border-color); padding:clamp(6px,.9vh,9px);
                     text-align:center; font-size:10.5px; color:var(--text-2); text-decoration:none;
                     display:block; }
        .du-rodape:hover { color:var(--text-1); background:rgba(var(--rgb-tinta),.04); }

        .du-grupos { display:grid; grid-template-columns:repeat(auto-fit,minmax(0,1fr));
                     gap:clamp(6px,.6vw,10px); min-height:0; }
        .du-gp { background:var(--dark-card-2); border:1px solid var(--border-color); border-radius:10px;
                 padding:clamp(7px,1vh,11px) clamp(8px,.6vw,12px);
                 display:flex; flex-direction:column; gap:5px; text-decoration:none; }
        a.du-gp:hover { border-color:rgba(var(--rgb-teal),.45); }
        .du-gp .lin { display:flex; justify-content:space-between; align-items:baseline; }
        .du-gp .nm { font-size:10.5px; color:var(--text-2); }
        .du-tela .du-gp .pc { font-size:10.5px; font-weight:700; color:var(--txt-verde); }
        .du-gp .qt { font-size:clamp(15px,2.15vh,21px); font-weight:800; letter-spacing:-.5px;
                     line-height:1; color:var(--text-1); }
        .du-gp .qt em { font-style:normal; font-size:9.5px; color:var(--text-3);
                        font-weight:400; margin-left:4px; }
        .du-gp .pe { font-size:9px; color:var(--text-3); display:flex; justify-content:space-between; }

        .du-barras { flex:1; min-height:0; display:grid;
                     grid-auto-rows:minmax(min-content,1fr); gap:1px; overflow-y:auto; }
        .du-br { display:grid; align-items:center; min-height:0; line-height:1.15; text-decoration:none;
                 grid-template-columns:clamp(80px,7.8vw,124px) 1fr clamp(74px,6.4vw,96px) clamp(38px,3.2vw,50px);
                 gap:clamp(7px,.7vw,11px); }
        .du-br:hover .nm { color:var(--text-1); }
        .du-br .nm { font-size:10.5px; color:var(--text-2); white-space:nowrap;
                     overflow:hidden; text-overflow:ellipsis; }
        .du-br .tr { height:6px; border-radius:99px; background:rgba(var(--rgb-tinta),.06); overflow:hidden; }
        .du-br .tr i { display:block; height:100%; border-radius:99px;
                       background:linear-gradient(90deg,var(--teal-2),var(--accent-green)); }
        .du-br .fr { font-size:10px; color:var(--text-3); text-align:right; }
        .du-br .pc { font-size:10.5px; font-weight:700; color:var(--text-1); text-align:right; }

        .du-sigem { flex:1; min-height:0; display:grid; grid-template-columns:auto 1fr;
                    align-items:center; gap:clamp(9px,.9vw,15px); }
        .du-rosca { position:relative; width:clamp(84px,11.5vh,118px); aspect-ratio:1; flex:none; }
        .du-rosca svg { width:100%; height:100%; }
        .du-rosca .centro { position:absolute; inset:0; display:grid; place-content:center;
                            text-align:center; pointer-events:none; }
        .du-rosca .centro b { font-size:clamp(14px,2vh,19px); font-weight:800; letter-spacing:-.5px;
                              display:block; line-height:1; color:var(--text-1); }
        .du-rosca .centro span { font-size:8.5px; color:var(--text-3); }
        .du-leg { display:grid; grid-auto-rows:minmax(min-content,1fr);
                  height:100%; min-height:0; gap:1px; overflow-y:auto; }
        .du-lg { display:grid; grid-template-columns:9px 1fr auto auto; align-items:center;
                 gap:clamp(5px,.5vw,9px); font-size:10.5px; text-decoration:none;
                 min-height:0; line-height:1.15; }
        .du-lg:hover .nm { color:var(--text-1); }
        .du-lg i { width:7px; height:7px; border-radius:50%; }
        .du-lg .nm { color:var(--text-2); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        .du-lg .n { font-weight:700; color:var(--text-1); }
        .du-lg .p { color:var(--text-3); font-size:9.5px; min-width:32px; text-align:right; }

        .du-tab { flex:1; min-height:0; display:flex; flex-direction:column; }
        .du-tab .cabtab, .du-tab .lin { display:grid; align-items:center; gap:6px;
                                        grid-template-columns:1.35fr 1.1fr .7fr .7fr .78fr; }
        .du-tab .cabtab { font-size:8.6px; letter-spacing:.7px; text-transform:uppercase;
                          color:var(--text-3); padding-bottom:5px;
                          border-bottom:1px solid var(--border-color); flex:none; }
        .du-tab .corpo { flex:1; min-height:0; display:grid;
                         grid-auto-rows:minmax(min-content,1fr); overflow-y:auto; }
        .du-tab .lin { border-bottom:1px solid rgba(var(--rgb-tinta),.045); font-size:10.5px;
                       min-height:0; line-height:1.2; }
        .du-tab .lin:last-child { border-bottom:0; }
        .du-tela .du-tab .tp { color:var(--text-3); }
        .du-tab .num { text-align:right; color:var(--text-2); }
        .du-tab .pnd { text-align:right; }
        .du-tab .cc { text-align:right; color:var(--text-2); font-size:10px; }
        /* a pill da tag e a de pendencia vem com o corpo das tabelas grandes;
           aqui elas sozinhas levavam a linha de 17px para 23px, e as 10 do
           Top 10 deixavam de caber por 48px */
        .du-tab .gtbl-badge { padding:0 6px !important; font-size:9.5px !important;
                              line-height:1.55 !important; }
        .du-tab .gtbl-tag { padding:1px 7px !important; font-size:10.5px !important;
                            line-height:1.35 !important; }
        .du-barras::-webkit-scrollbar, .du-leg::-webkit-scrollbar,
        .du-tab .corpo::-webkit-scrollbar { width:5px; }
        .du-barras::-webkit-scrollbar-thumb, .du-leg::-webkit-scrollbar-thumb,
        .du-tab .corpo::-webkit-scrollbar-thumb {
          background:rgba(var(--rgb-tinta),.13); border-radius:99px; }

        .du-pe { display:grid; grid-template-columns:repeat(5,1fr); gap:clamp(8px,.85vw,13px); }
        .du-mini { background:var(--dark-card); border:1px solid var(--border-color); border-radius:13px;
                   padding:clamp(8px,1.15vh,12px) clamp(11px,.85vw,15px);
                   display:flex; align-items:center; gap:11px; text-decoration:none; }
        a.du-mini:hover { border-color:rgba(var(--rgb-azul),.45); }
        .du-mini .tile { width:clamp(24px,3vh,31px); height:clamp(24px,3vh,31px); border-radius:9px;
                         flex:none; display:flex; align-items:center; justify-content:center; }
        .du-mini .tile svg { width:14px; height:14px; }
        .du-mini .rot { font-size:10px; color:var(--text-3); line-height:1.3; }
        .du-mini .val { font-size:clamp(14px,2.1vh,20px); font-weight:800; letter-spacing:-.5px;
                        line-height:1.2; color:var(--text-1); }
        .du-mini .val.pq { font-size:clamp(11.5px,1.65vh,15px); letter-spacing:-.2px; }
        .du-vazio { grid-column:1/-1; display:grid; place-content:center; text-align:center;
                    color:var(--text-2); font-size:13px; gap:6px; }

        /* Abaixo de 700px de altura nao da para manter os numeros legiveis numa
           tela so. Em vez de espremer ate ninguem conseguir ler, devolvo a
           rolagem: e o unico ponto em que a tela unica cede. */
        @media (max-height: 700px), (max-width: 1150px) {
          .stApp:has(.du-tela) [data-testid="stMain"] { overflow: auto !important; }
          .stApp:has(.du-tela) [data-testid="stMainBlockContainer"],
          .stApp:has(.du-tela) .block-container { height:auto; display:block !important; }
          .du-tela { min-height:640px; grid-template-rows:auto auto auto auto; }
          .du-meio { min-height:560px; }
        }
        @media (max-width: 1150px) {
          .du-kpis, .du-pe { grid-template-columns:repeat(3,1fr); }
          .du-meio { grid-template-columns:1fr; }
        }

        /* --------- filtros da lateral: marca em cima, filtros embaixo -------
           O st.navigation injeta o menu como irmao do conteudo do usuario, e
           tudo que mando para a lateral cai num container so -- entao marca e
           filtros ficariam juntos, os dois acima ou os dois abaixo do menu.
           Com display:contents esse container some do layout e cada bloco vira
           filho direto da coluna flex, o que permite intercalar. */
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] > div,
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] > div > [data-testid="stVerticalBlock"] {
          display: contents !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] .stElementContainer,
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stElementContainer"] { order: 3; }
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] .stElementContainer:first-child { order: 1; }
        /* O display:contents tirou o bloco vertical do layout, e com ele o
           gap de 16px que compensava a margem negativa que o Streamlit poe em
           todo markdown -- por isso o titulo dos filtros subia por cima do
           primeiro rotulo. */
        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] { gap: 5px; }
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
        /* Os filtros ficam logo abaixo do menu, e nao colados no rodape.
           Empurrado para o fim da coluna, o bloco passava da dobra em tela
           com barra de tarefas: o ultimo seletor sumia atras dela e o
           dropdown abria para baixo, fora do monitor. */
        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
          overflow-y: auto; padding-bottom: 14px;
        }
        /* A lista do selectbox e portalada no body com position:fixed e
           altura fixa de 300px, e nao vira para cima quando falta espaco --
           as opcoes de baixo ficavam atras da barra de tarefas, inalcancaveis.
           Como os filtros sao a primeira coisa da lateral, a lista sempre
           abre por volta de 430px do topo; limitar a altura ao que resta
           abaixo disso faz ela rolar por dentro em vez de sumir. */
        @media (max-height: 900px) {
          div[data-testid="stSelectboxVirtualDropdown"],
          div[data-testid="stSelectboxVirtualDropdown"] > div,
          div[data-testid="stSelectboxVirtualDropdown"] ul {
            max-height: max(100px, calc(100dvh - 470px)) !important;
          }
        }
        .flt-topo { padding:2px 4px 0; display:flex; align-items:baseline;
                    justify-content:space-between; gap:8px; }
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] {
          border-top: 1px solid var(--border-color); padding-top: 8px; margin-top: 10px; }
        .flt-titulo { font-size:9.5px; letter-spacing:.6px; text-transform:uppercase;
                      color:var(--text-3); font-weight:700; white-space:nowrap; }
        .flt-conta { font-size:10px; color:var(--text-3); white-space:nowrap; }
        .flt-conta b { color:var(--txt-verde); }
        section[data-testid="stSidebar"] .stSelectbox label { font-size:9.5px !important;
          color:var(--text-3) !important; margin-bottom:0 !important;
          min-height:0 !important; line-height:1.25 !important; }
        section[data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.stSelectbox) {
          margin-bottom:-9px; }
        section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div {
          background:var(--dark-card-2) !important; border-color:var(--border-color) !important;
          font-size:11.5px !important; min-height:30px !important; }
        section[data-testid="stSidebar"] .stButton button { width:100%; font-size:11px !important;
          padding:4px 10px !important; border-radius:8px !important; }
        /* ------- botao de tema: so o icone, preso no alto a direita -------
           Fora do fluxo, ao lado do menu do proprio Streamlit. O rotulo existe
           para leitor de tela, mas nao ocupa espaco. */
        .st-key-gplan_btn_tema { position: fixed !important; top: 11px; right: 58px;
          z-index: 999990; width: auto !important; }
        .st-key-gplan_btn_tema button {
          width: 38px !important; height: 38px !important; min-height: 0 !important;
          padding: 0 !important; border-radius: 50% !important;
          display: inline-flex !important; align-items: center; justify-content: center;
          background: var(--dark-card) !important;
          border: 1px solid var(--border-strong) !important;
          box-shadow: 0 4px 14px var(--sombra-leve); }
        .st-key-gplan_btn_tema button:hover {
          background: rgba(var(--rgb-azul),0.16) !important;
          border-color: rgba(var(--rgb-azul),0.5) !important;
          transform: translateY(-1px); }
        .st-key-gplan_btn_tema button p,
        .st-key-gplan_btn_tema button [data-testid="stMarkdownContainer"] {
          display: none !important; }
        .st-key-gplan_btn_tema button [data-testid="stIconMaterial"] {
          margin: 0 !important; font-size: 20px !important; color: var(--text-1) !important; }
        @media (max-width: 700px) { .st-key-gplan_btn_tema { right: 50px; top: 8px; } }

        /* ---------------------------------------------------------------
           Chrome do proprio Streamlit. O config.toml fixa base="dark", entao
           no tema claro os widgets nasceriam com texto claro sobre fundo
           claro -- rotulo, valor selecionado, lista do dropdown e placeholder
           sumiriam. Tudo aqui le os mesmos tokens, e serve aos dois temas.
           --------------------------------------------------------------- */
        /* A cor base desce por heranca. Listar span/div/p aqui daria a estas
           regras especificidade maior que a de qualquer classe do app -- a
           etiqueta da zona, por exemplo, perdia o proprio --sobre-cor e voltava
           a ser texto claro em cima de amarelo. */
        .stApp, .stMarkdown, [data-testid="stMarkdownContainer"] { color: var(--text-1); }
        label, .stSelectbox label, .stMultiSelect label, .stTextInput label,
        [data-testid="stWidgetLabel"] p { color: var(--text-2) !important; }

        /* O controle do selectbox nesta versao e um div[role=group] com as cores
           do config.toml queimadas dentro -- fundo #0a0e1a e texto claro. No
           tema claro isso vira caixa preta com letra branca no meio da tela
           branca. Vale para selectbox, multiselect e campo de texto. */
        /* O invólucro de cada widget leva o secondaryBackgroundColor do
           config.toml -- #12172a queimado -- e cada tipo usa um testid
           diferente: o selectbox é div[role=group], o campo de texto é
           stTextInputRootElement. Faltando um deles, sobra uma caixa escura no
           meio da tela clara. */
        [data-testid="stSelectbox"] div[role="group"],
        [data-testid="stMultiSelect"] div[role="group"],
        [data-testid="stTextInput"] div[role="group"],
        [data-testid="stTextInputRootElement"],
        [data-testid="stNumberInputContainer"],
        [data-testid="stTextArea"] textarea,
        [data-testid="stDateInput"] div[role="group"],
        [data-baseweb="select"] > div, [data-baseweb="input"],
        .stTextInput input, .stNumberInput input, .stTextArea textarea {
          background: var(--dark-card-2) !important; color: var(--text-1) !important;
          border-color: var(--border-color) !important; }
        /* tudo que esta dentro do controle -- valor escolhido, pilulas,
           contador -- e nao so o input: o texto ali nasce com a cor do
           config.toml e no tema claro fica branco sobre branco */
        [data-testid="stSelectbox"] div[role="group"] *,
        [data-testid="stMultiSelect"] div[role="group"] *,
        [data-testid="stTextInput"] div[role="group"] *,
        [data-testid="stSelectbox"] input, [data-testid="stMultiSelect"] input,
        [data-testid="stTextInput"] input, [data-testid="stSelectbox"] button,
        [data-testid="stMultiSelect"] button, [role="combobox"] {
          color: var(--text-1) !important; background: transparent !important; }
        /* O Streamlit pendura um icone de ancora em todo <h1> do markdown, com
           a cor do config a 60% -- no tema claro some contra o fundo. */
        [data-testid="stHeaderActionElements"] svg,
        [data-testid="stHeaderActionElements"] svg * {
          stroke: var(--text-3) !important; color: var(--text-3) !important; }
        /* o texto de espera do multiselect fica fora do role=group */
        [data-testid="stMultiSelect"] div, [data-testid="stMultiSelect"] span {
          color: var(--text-1); }
        [data-testid="stMultiSelect"] [class*="placeholder"],
        [data-testid="stMultiSelect"] div[aria-live] { color: var(--text-3) !important; }
        /* o icone do menu e uma ligadura de fonte, nao um svg: sem herdar a
           cor do link ele fica com a do config e some no tema claro */
        section[data-testid="stSidebar"] nav a span,
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a span {
          color: inherit !important; }
        [data-testid="stSelectbox"] svg, [data-testid="stMultiSelect"] svg,
        [data-testid="stTextInput"] svg, [data-baseweb="select"] svg {
          fill: var(--text-2) !important; color: var(--text-2) !important; }
        input::placeholder, textarea::placeholder {
          color: var(--text-3) !important; opacity: 1 !important; }

        /* a lista de opcoes e portalada para fora do widget: sem regra propria
           ela nasce com o fundo do config e o texto some */
        [role="listbox"], [role="dialog"] [role="listbox"], [data-baseweb="menu"],
        div[data-testid="stSelectboxVirtualDropdown"],
        div[data-testid="stSelectboxVirtualDropdown"] ul {
          background: var(--dark-card) !important; color: var(--text-1) !important;
          border: 1px solid var(--border-color) !important;
          box-shadow: 0 18px 46px var(--sombra) !important; }
        [role="option"], [data-baseweb="menu"] li,
        div[data-testid="stSelectboxVirtualDropdown"] li {
          color: var(--text-1) !important; background: transparent !important; }
        [role="option"]:hover, [role="option"][data-focused],
        [role="option"][aria-selected="true"], [data-baseweb="menu"] li:hover,
        div[data-testid="stSelectboxVirtualDropdown"] li:hover {
          background: rgba(var(--rgb-azul),0.16) !important; color: var(--text-1) !important; }

        /* O body guarda as cores do config.toml, e e nele que os portais
           nascem: sem isto o menu suspenso abre escuro no tema claro. */
        body { background: var(--dark-bg) !important; color: var(--text-1) !important; }
        /* O botao Deploy e afordancia do Streamlit Cloud e nao serve aqui --
           sem tira-lo, o botao de tema cai em cima dele no canto. */
        [data-testid="stAppDeployButton"] { display: none !important; }
        /* Icones do proprio Streamlit -- recolher a lateral, menu do canto,
           setas de expander. Nascem com a cor do config.toml: quase brancos, o
           que no tema claro os deixa invisiveis contra o fundo. */
        [data-testid="stHeader"] svg, [data-testid="stToolbar"] svg,
        [data-testid="stSidebarCollapseButton"] svg,
        [data-testid="stExpandSidebarButton"] svg,
        [data-testid="stSidebarCollapsedControl"] svg,
        [data-testid="stMainMenu"] svg, [data-testid="stExpander"] svg {
          fill: var(--text-2) !important; color: var(--text-2) !important; }
        [data-testid="stHeader"] { background: transparent !important; }

        .stButton button, [data-testid="stBaseButton-secondary"] {
          background: var(--dark-card-2) !important; color: var(--text-1) !important;
          border: 1px solid var(--border-color) !important; }
        .stButton button:hover, [data-testid="stBaseButton-secondary"]:hover {
          background: rgba(var(--rgb-azul),0.14) !important;
          border-color: rgba(var(--rgb-azul),0.45) !important;
          color: var(--text-1) !important; }
        .stButton button:focus-visible, a:focus-visible, button:focus-visible {
          outline: 2px solid var(--accent-blue) !important; outline-offset: 2px; }
        .stButton button svg, [data-testid="stBaseButton-secondary"] svg {
          fill: currentColor !important; }

        /* pilulas de multiselect e chips */
        [data-baseweb="tag"] { background: rgba(var(--rgb-azul),0.18) !important;
          color: var(--text-1) !important; border-color: rgba(var(--rgb-azul),0.35) !important; }
        [data-baseweb="tag"] svg { fill: var(--text-1) !important; }

        /* barra de rolagem: a padrao do Chrome no claro fica quase invisivel */
        * { scrollbar-color: var(--neutro) transparent; }
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-thumb { background: var(--neutro); border-radius: 8px;
          border: 2px solid transparent; background-clip: padding-box; }
        ::-webkit-scrollbar-track { background: transparent; }


        /* =========================================================
           Fichas: tag, relatorio e nivel. As tres usam os mesmos
           blocos -- trilha, tiles, KPIs, paineis -- para que quem
           aprendeu a ler uma leia as outras sem reaprender nada.
           ========================================================= */
        /* célula que corta no fim em vez de empurrar a tabela para fora */
        .gtbl td.gt-corta { max-width: 260px; overflow: hidden; text-overflow: ellipsis;
          white-space: nowrap; }
        /* Lista longa rola por dentro do painel em vez de paginar: o cabecalho
           fica grudado no topo para nao se perder a coluna no meio da rolagem. */
        .fx-rolagem { max-height: min(52vh, 620px); overflow-y: auto; }
        .fx-rolagem table.gtbl thead th { position: sticky; top: 0; z-index: 2;
          background: var(--dark-card-2); }
        .fx-rolagem::-webkit-scrollbar { width: 8px; }
        .fx-rolagem::-webkit-scrollbar-thumb { background: rgba(var(--rgb-tinta),.13);
          border-radius: 99px; }
        .fx-rolagem::-webkit-scrollbar-track { background: transparent; }
        /* Icone: <span> vazio pintado por mascara. O desenho vem de fx_css_icones. */
        .fxi { display:inline-block; width:1em; height:1em; flex:none;
               background-color:currentColor;
               -webkit-mask:var(--fxi) center/contain no-repeat;
               mask:var(--fxi) center/contain no-repeat; }
        __ICONES__
        .fx-tile .ic .fxi, .fx-kpi .ic .fxi, .fx-pn-t .ic .fxi,
        .fx-acao .ic .fxi, .fx-com .cab .ic .fxi, .fx-cab .marca .fxi
          { width:100%; height:100%; }
        .fx-folha { width:12px; height:12px; vertical-align:-1px; margin-right:7px;
                    color:var(--text-2); }

        /* Donut sem SVG: anel de conic-gradient com furo de mascara. */
        .fx-rosca { color:var(--accent-teal); }
        .fx-rosca .anel { position:absolute; inset:0; border-radius:50%;
          background:conic-gradient(currentColor calc(var(--p,0) * 1%),
                                    rgba(var(--rgb-tinta),.07) 0);
          /* as paradas de cor de um gradiente radial sao relativas a linha do
             gradiente, nao a caixa: com calc(50% - 13px) o furo saia com 15px
             em vez de 43 e o donut virava uma pizza cheia. Em fracao do raio
             o furo fica certo em qualquer tamanho. */
          -webkit-mask:radial-gradient(farthest-side, #0000 75%, #000 76%);
          mask:radial-gradient(farthest-side, #0000 75%, #000 76%); }
        /* Cores dos blocos como classe -- ver FX_COR. */
        .fxc-azul  { color:var(--accent-blue); } .fx-tile .ic.fxc-azul  { background:rgba(var(--rgb-azul),.11); }
        .fxc-teal  { color:var(--accent-teal); } .fx-tile .ic.fxc-teal  { background:rgba(var(--rgb-teal),.11); }
        .fxc-roxo  { color:var(--accent-purple); } .fx-tile .ic.fxc-roxo  { background:rgba(var(--rgb-roxo),.11); }
        .fxc-ambar { color:var(--accent-amber); } .fx-tile .ic.fxc-ambar { background:rgba(var(--rgb-ambar),.11); }
        .fxc-rubi  { color:var(--accent-red); } .fx-tile .ic.fxc-rubi  { background:rgba(var(--rgb-vermelho),.11); }
        .fxc-mudo  { color:var(--text-2); } .fx-tile .ic.fxc-mudo  { background:rgba(var(--rgb-tinta),.13); }
        .fxc-verde { color:var(--accent-green); } .fx-tile .ic.fxc-verde { background:rgba(var(--rgb-verde),.11); }
        .fxc-cinza { color:var(--neutro); }
        .fx-trilho.fxc-azul i  { background:var(--accent-blue); }
        .fx-trilho.fxc-teal i  { background:var(--accent-teal); }
        .fx-trilho.fxc-roxo i  { background:var(--accent-purple); }
        .fx-trilho.fxc-ambar i { background:var(--accent-amber); }
        .fx-trilho.fxc-rubi i  { background:var(--accent-red); }
        .fx-trilho.fxc-mudo i  { background:var(--text-2); }
        .fx-trilho.fxc-verde i { background:var(--accent-green); }
        .fx-lg i.fxc-azul, .du-lg i.fxc-azul  { background:var(--accent-blue); }
        .fx-lg i.fxc-teal, .du-lg i.fxc-teal  { background:var(--accent-teal); }
        .fx-lg i.fxc-roxo, .du-lg i.fxc-roxo  { background:var(--accent-purple); }
        .fx-lg i.fxc-ambar, .du-lg i.fxc-ambar { background:var(--accent-amber); }
        .fx-lg i.fxc-rubi, .du-lg i.fxc-rubi  { background:var(--accent-red); }
        .fx-lg i.fxc-mudo, .du-lg i.fxc-mudo  { background:var(--text-2); }
        .fx-lg i.fxc-verde, .du-lg i.fxc-verde { background:var(--accent-green); }
        .fx-lg i.fxc-cinza, .du-lg i.fxc-cinza { background:var(--neutro); }
        .fx { display:flex; flex-direction:column; gap:14px; }
        .fx svg { width:100%; height:100%; }

        /* trilha: a cadeia de onde a coisa pendura */
        .fx-trilha { display:flex; align-items:center; gap:7px; flex-wrap:wrap;
                     font-size:11.5px; color:var(--text-3); }
        /* o caminho de volta e clicavel, mas sem risco embaixo: sublinhado
           aqui vira ruido, sao quatro links seguidos */
        .fx-trilha a { color:var(--text-2) !important; text-decoration:none !important;
                       border-radius:5px; padding:1px 5px; margin:0 -2px; }
        .fx-trilha a:hover { color:var(--text-1) !important;
                             background:rgba(var(--rgb-tinta),.06); }
        .fx-trilha .sep { color:var(--text-3); }
        .fx-trilha .aqui { color:var(--text-1); font-weight:650; }

        /* tiles do topo */
        .fx-tiles { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
                    gap:10px; }
        .fx-tile { background:var(--dark-card-2); border:1px solid var(--border-color);
                   border-radius:12px; padding:11px 12px; display:flex; align-items:center;
                   gap:10px; min-width:0; text-decoration:none !important; }
        a.fx-tile:hover { border-color:rgba(var(--rgb-azul),.5); }
        .fx-tile .ic { width:31px; height:31px; flex:none; border-radius:9px; padding:7px;
                       display:grid; place-items:center; }
        .fx-tile .cp { min-width:0; display:flex; flex-direction:column; gap:2px; }
        .fx-tile .rot { font-size:9.5px; letter-spacing:.7px; text-transform:uppercase;
                        color:var(--text-3); font-weight:700; }
        .fx-tile .val { font-size:13.5px; font-weight:700; color:var(--text-1); line-height:1.25;
                        overflow-wrap:anywhere; display:-webkit-box; -webkit-line-clamp:2;
                        -webkit-box-orient:vertical; overflow:hidden; }
        .fx-tile .sub { font-size:10.5px; color:var(--text-3); }

        /* duas colunas */
        .fx-corpo { display:grid; grid-template-columns:1fr 288px; gap:13px; align-items:start; }
        .fx-col { display:flex; flex-direction:column; gap:13px; min-width:0; }

        /* paineis */
        .fx-pn { background:var(--dark-card-2); border:1px solid var(--border-color);
                 border-radius:13px; overflow:hidden; }
        .fx-pn-t { display:flex; align-items:center; gap:8px; font-size:12.5px; font-weight:700;
                   color:var(--text-1); padding:12px 14px 10px; border-bottom:1px solid var(--border-color); }
        .fx-pn-t .ic { width:15px; height:15px; flex:none; color:var(--text-3); }
        .fx-pn-t .conta { margin-left:auto; font-size:10.5px; font-weight:500; color:var(--text-3); }
        .fx-pn-c { padding:13px 14px; }
        .fx-pn-c.zero { padding:0; }
        .fx-pn-c.centro { display:flex; flex-direction:column; align-items:center; gap:11px; }

        /* KPIs */
        .fx-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(120px,1fr)); gap:10px; }
        .fx-kpi { background:var(--dark-card); border:1px solid var(--border-color);
                  border-radius:11px; padding:12px; display:flex; flex-direction:column; gap:5px; }
        .fx-kpi .top { display:flex; align-items:flex-start; justify-content:space-between; gap:7px; }
        .fx-kpi .rot { font-size:9.5px; letter-spacing:.7px; text-transform:uppercase;
                       color:var(--text-3); font-weight:700; }
        .fx-kpi .ic { width:16px; height:16px; flex:none; }
        .fx-kpi .val { font-size:25px; font-weight:800; letter-spacing:-.9px; line-height:1;
                       color:var(--text-1); }
        .fx-kpi .sub { font-size:10px; color:var(--text-3); min-height:24px; }
        .fx-trilho { height:3px; border-radius:99px; background:rgba(var(--rgb-tinta),.07); overflow:hidden; }
        .fx-trilho i { display:block; height:100%; border-radius:99px; }

        /* dados em par rotulo/valor */
        .fx-dados { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr));
                    gap:9px; margin-top:11px; }
        .fx-dado { background:var(--dark-card); border:1px solid var(--border-color);
                   border-radius:10px; padding:9px 11px; min-width:0; }
        .fx-dado .rot { font-size:9.5px; letter-spacing:.7px; text-transform:uppercase;
                        color:var(--text-3); font-weight:700; }
        .fx-dado .val { font-size:14px; font-weight:700; color:var(--text-1); margin-top:4px;
                        overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }

        /* rosca */
        .fx-rosca { position:relative; width:112px; aspect-ratio:1; }
        .fx-rosca .centro { position:absolute; inset:0; display:grid; place-content:center;
                            text-align:center; }
        .fx-rosca .centro b { font-size:20px; font-weight:800; letter-spacing:-.7px;
                              display:block; line-height:1; color:var(--text-1); }
        .fx-rosca .centro span { font-size:9.5px; color:var(--text-3); }
        .fx-leg { display:flex; flex-direction:column; gap:6px; width:100%; }
        .fx-lg { display:grid; grid-template-columns:9px 1fr auto auto; align-items:center;
                 gap:8px; font-size:11.5px; text-decoration:none !important; }
        .fx-lg i { width:8px; height:8px; border-radius:2px; }
        .fx-lg .nm { color:var(--text-2); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        .fx-lg b { font-weight:700; color:var(--text-1); }
        .fx-lg em { font-style:normal; font-size:10px; color:var(--text-3); min-width:34px; text-align:right; }
        .fx-lg.total { border-top:1px solid var(--border-color); padding-top:6px; }
        .fx-nota { font-size:10.5px; color:var(--text-3); text-align:center; line-height:1.5; }

        /* linhas de info */
        .fx-linha { display:flex; align-items:center; justify-content:space-between; gap:10px;
                    font-size:12px; padding:7px 0; border-bottom:1px solid rgba(var(--rgb-tinta),.045); }
        .fx-linha:last-child { border-bottom:0; }
        .fx-linha > span:first-child { color:var(--text-2); }
        .fx-linha b { font-weight:700; font-size:12.5px; color:var(--text-1); }
        .fx-linha a { text-decoration:none !important; }

        /* acoes */
        .fx-acoes { display:flex; flex-direction:column; gap:7px; }
        .fx-acao { display:flex; align-items:center; gap:9px; font-size:12px;
                   color:var(--text-1) !important; text-decoration:none !important;
                   background:var(--dark-card); border:1px solid var(--border-color);
                   border-radius:9px; padding:9px 11px; }
        .fx-acao .ic { width:15px; height:15px; flex:none; color:var(--text-3); }
        .fx-acao:hover { border-color:rgba(var(--rgb-azul),.5); }


        /* revisoes do relatorio */
        .fx-rev-atual td { background:rgba(var(--rgb-ambar),.05); }
        .fx-atual { font-size:8.5px; letter-spacing:.6px; text-transform:uppercase;
                    color:var(--txt-ambar); background:rgba(var(--rgb-ambar),.14);
                    border:1px solid rgba(var(--rgb-ambar),.3); border-radius:5px;
                    padding:1px 5px; margin-left:6px; font-weight:700; vertical-align:1px; }
        .fx-com { border-radius:10px; padding:10px 12px; border:1px solid; margin:0 0 4px; }
        .fx-com.rec { background:rgba(var(--rgb-vermelho),.075); border-color:rgba(var(--rgb-vermelho),.3); }
        .fx-com.obs { background:rgba(var(--rgb-azul),.06); border-color:rgba(var(--rgb-azul),.26); }
        .fx-com .cab { display:flex; align-items:center; gap:7px; font-size:10px; letter-spacing:.6px;
                       text-transform:uppercase; font-weight:700; margin-bottom:6px; }
        .fx-com.rec .cab { color:var(--txt-vermelho); }
        .fx-com.obs .cab { color:var(--txt-azul); }
        .fx-com .cab .ic { width:13px; height:13px; flex:none; }
        .fx-com p { white-space:pre-wrap; word-break:break-word; font-size:11.5px;
                    line-height:1.55; margin:0; }
        .fx-com.rec p { color:var(--txt-vermelho); }
        .fx-com.obs p { color:var(--txt-azul); }
        .gtbl td.fx-com-cel { padding:0 14px 10px !important; }

        /* cabecalho da ficha fora do modal (aba Pesquisa tag) */
        .fx-cab { display:flex; align-items:center; gap:12px; }
        .fx-cab .marca { width:40px; height:40px; flex:none; border-radius:12px; padding:10px;
                         display:grid; place-items:center; background:rgba(var(--rgb-azul),.15);
                         color:var(--accent-blue); }
        .fx-cab h2 { font-size:21px; font-weight:800; letter-spacing:-.5px; margin:0;
                     color:var(--text-1); line-height:1.15; }
        .fx-cab p { font-size:12.5px; color:var(--text-2); margin:3px 0 0; }

        @media (max-width: 900px) {
          .fx-corpo { grid-template-columns:1fr; }
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
        .kpi-value { font-size: 30px; font-weight: 800; color: var(--text-1); letter-spacing:-0.8px; margin-bottom:14px; }
        .kpi-progress-row { display:flex; align-items:center; gap:10px; }
        .kpi-track { flex:1; height:5px; background: rgba(var(--rgb-tinta),0.06); border-radius:3px; overflow:hidden; }
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
        .gtbl td { padding:12px 14px; border-bottom:1px solid rgba(var(--rgb-tinta),0.04) !important; font-size:13px; color:var(--text-1); white-space:nowrap; }
        .gtbl tbody tr:last-child td { border-bottom:none !important; }
        .gtbl tbody tr:hover td { background:rgba(var(--rgb-tinta),0.025) !important; }
        /* Centralizado (e nao a direita) para o numero ficar sob o proprio
           cabecalho, em vez de encostar na coluna seguinte.
           Precisa de "th.gtbl-num"/"td.gtbl-num": so ".gtbl-num" perde em
           especificidade para ".gtbl th { text-align:left }" e o cabecalho
           acabava a esquerda enquanto o valor ia pro centro. */
        .gtbl th.gtbl-num, .gtbl td.gtbl-num { text-align:center; }
        .gtbl-num { font-variant-numeric:tabular-nums; }
        /* A arvore ja traz no HTML as TAGs de todas as malhas, para o "+" abrir
           na hora. Sao ~5.100 linhas: repetir class="gtbl-num" custaria uns
           150 bytes por linha, entao aqui o alinhamento sai pela posicao. */
        .gtbl-tags td:nth-child(n+3) { text-align:center; font-variant-numeric:tabular-nums; }
        .gtbl-muted { color:var(--text-2); }
        .gtbl-mono { font-size:12px; color:var(--text-2); }
        .gtbl-strong { font-weight:600; }
        .gtbl-scroll { overflow-x:auto; }
        .gtbl-tag {
          display:inline-block; font-size:12px; font-weight:600; color:var(--txt-azul); white-space:nowrap;
          background:rgba(var(--rgb-azul),0.14); border:1px solid rgba(var(--rgb-azul),0.22);
          padding:3px 9px; border-radius:6px; letter-spacing:0.2px;
        }
        a.gtbl-link, a.gtbl-link:hover, a.gtbl-link:visited { text-decoration:none !important; color:var(--txt-azul) !important; }
        a.gtbl-link:hover { background:rgba(var(--rgb-azul),0.26) !important; border-color:rgba(var(--rgb-azul),0.45) !important; }
        .gtbl-badge {
          display:inline-block; min-width:26px; text-align:center; font-size:11.5px; font-weight:600;
          padding:3px 9px; border-radius:6px; font-variant-numeric:tabular-nums; white-space:nowrap;
        }
        .gtbl-badge.crit { color:var(--txt-vermelho); background:rgba(var(--rgb-vermelho),0.16); border:1px solid rgba(var(--rgb-vermelho),0.28); }
        .gtbl-badge.warn { color:var(--txt-ambar); background:rgba(var(--rgb-ambar),0.14); border:1px solid rgba(var(--rgb-ambar),0.26); }
        .gtbl-badge.ok   { color:var(--txt-teal); background:rgba(var(--rgb-teal),0.13); border:1px solid rgba(var(--rgb-teal),0.24); }
        /* etapa do caminho, nao alerta: azul, a mesma familia da pill de TAG */
        .gtbl-badge.andamento { color:var(--txt-azul); background:rgba(var(--rgb-azul),0.14); border:1px solid rgba(var(--rgb-azul),0.26); }
        .gtbl-badge.roxo { color:var(--txt-roxo); background:rgba(var(--rgb-roxo),0.14); border:1px solid rgba(var(--rgb-roxo),0.26); }
        .gtbl-badge.mudo { color:var(--text-2); background:rgba(var(--rgb-tinta),0.07); border:1px solid rgba(var(--rgb-tinta),0.13); }
        .gtbl-badge.cinza { color:var(--text-3); background:rgba(var(--rgb-tinta),0.05); border:1px solid rgba(var(--rgb-tinta),0.10); }
        .gtbl-pag { text-align:center; color:var(--text-2); font-size:12.5px; padding-top:9px; }
        /* as fatias do donut do SIGEM: a cor vem da classe, o traco a segue */
        .du-rosca svg .fatia { stroke:currentColor; }
        .du-rosca svg .trilho { stroke:rgba(var(--rgb-tinta),.07); }
        .gtbl-empty { padding:34px 4px; text-align:center; color:var(--text-3); font-size:13px; }
        /* Recusados ha mais tempo: linha inteira clicavel, leva para a aba
           Relatorios ja buscando a TAG. */
        /* As duas colunas do Dashboard num grid so, para terminarem juntas.
           O painel de recusados estica para fechar o bloco. */
        .dash-linha { display:grid; grid-template-columns:1.55fr 1fr; gap:22px;
                      align-items:start; margin-bottom:22px; }
        .dash-dir { display:flex; flex-direction:column; gap:22px; }
        .dash-linha > .gplan-panel, .dash-dir > .gplan-panel { margin-bottom:0 !important; }
        @media (max-width:1100px) { .dash-linha { grid-template-columns:1fr; } }
        /* faixa de resumo das abas de lista */
        .fx-faixa { display:grid; grid-auto-flow:column; grid-auto-columns:1fr;
                    gap:10px; padding:16px 22px !important; margin-bottom:18px !important; }
        .fx-item { display:flex; flex-direction:column; gap:5px; }
        .fx-lbl { font-size:10.5px; font-weight:600; letter-spacing:0.6px;
                  text-transform:uppercase; color:var(--text-3); }
        .fx-val { font-size:19px; font-weight:800; color:var(--text-1);
                  letter-spacing:-0.4px; font-variant-numeric:tabular-nums; }
        .fx-val.ruim { color:var(--txt-vermelho); }
        .fx-val.bom { color:var(--txt-teal); }
        @media (max-width:900px) { .fx-faixa { grid-auto-flow:row; grid-auto-columns:auto;
                                               grid-template-columns:repeat(2,1fr); } }
        .rec-resumo { font-size:12px; color:var(--text-3); margin-bottom:10px; }
        a.rec-linha {
          display:grid; grid-template-columns:1fr auto auto; align-items:center; gap:12px;
          padding:7px 10px; margin-bottom:4px; border-radius:8px;
          background:var(--dark-card-2); border:1px solid var(--border-color);
          text-decoration:none !important; transition:border-color 120ms, background 120ms;
        }
        a.rec-linha:last-child { margin-bottom:0; }
        a.rec-linha:hover { background:rgba(var(--rgb-tinta),0.04);
                            border-color:rgba(var(--rgb-vermelho),0.4); }
        .rec-tag { font-size:12.5px; font-weight:600; color:var(--text-1); }
        .rec-rel { font-size:11px; color:var(--text-3); }
        .rec-dias { font-size:11.5px; font-weight:600; color:var(--txt-vermelho);
                    font-variant-numeric:tabular-nums; white-space:nowrap; }
        /* O motivo da recusa e o que precisa ser tratado: vermelho e legivel,
           nao pill -- o texto da fiscalizacao pode ser longo.
           Precisa de ".gtbl td.rel-com": so ".rel-com" perde em especificidade
           para ".gtbl td { color:var(--text-1) }" e o texto saia branco. */
        .gtbl td.rel-com { color:var(--txt-vermelho); font-size:12.5px; white-space:pre-wrap;
                           max-width:520px; line-height:1.45; }
        .rel-titulo { font-size:15px; font-family:ui-monospace,SFMono-Regular,Menlo,monospace;
                      letter-spacing:-0.2px; word-break:break-all; }
        .rel-sit { flex-shrink:0; align-self:center; }
        a.rel-sigem {
          display:inline-block; font-size:11.5px; font-weight:600; white-space:nowrap;
          color:var(--txt-azul) !important; background:rgba(var(--rgb-azul),0.14);
          border:1px solid rgba(var(--rgb-azul),0.26); border-radius:6px; padding:3px 10px;
          text-decoration:none !important; transition:background 120ms;
        }
        a.rel-sigem:hover { background:rgba(var(--rgb-azul),0.26); border-color:rgba(var(--rgb-azul),0.45); }
        /* Abre a ficha do relatorio. Fica em coluna propria porque o endereco
           do documento e longo demais para virar area de clique. */
        a.btn-detalhes {
          display:inline-block; font-size:11.5px; font-weight:600; white-space:nowrap;
          color:var(--txt-azul) !important; background:rgba(var(--rgb-tinta),0.06);
          border:1px solid var(--border-strong); border-radius:7px; padding:4px 12px;
          text-decoration:none !important; transition:background 120ms, border-color 120ms;
        }
        a.btn-detalhes:hover { background:rgba(var(--rgb-azul),0.2);
                               border-color:rgba(var(--rgb-azul),0.45); color:var(--txt-azul) !important; }
        .prg-trilha { font-size:12.5px; color:var(--text-2); margin-bottom:18px; }
        .prg-sep { color:var(--text-3); margin:0 2px; }
        a.prg-link { color:var(--txt-azul) !important; text-decoration:none !important; font-weight:600; }
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

        /* ficha do nivel atual (SOP / SSOP / segmento / malha) */
        .fn-panel { padding:26px 30px; margin-bottom:24px !important; }
        .fn-head { display:flex; justify-content:space-between; align-items:flex-start;
                   gap:24px; margin-bottom:22px; flex-wrap:wrap; }
        .fn-tipo { font-size:10.5px; text-transform:uppercase; letter-spacing:0.8px;
                   color:var(--text-3); font-weight:700; margin-bottom:5px; }
        .fn-titulo { font-size:22px; font-weight:800; color:var(--text-1); letter-spacing:-0.5px; }
        .fn-avanco { display:flex; align-items:center; gap:12px; }
        .fn-track { width:160px; height:9px; background:rgba(var(--rgb-tinta),0.07);
                    border-radius:5px; overflow:hidden; }
        .fn-fill { height:100%; border-radius:5px; }
        .fn-fill.ok { background:var(--accent-teal); }
        .fn-fill.warn { background:var(--accent-amber); }
        .fn-fill.crit { background:var(--accent-red); }
        .fn-pct { font-size:20px; font-weight:800; color:var(--text-1);
                  font-variant-numeric:tabular-nums; min-width:66px; text-align:right; }
        .fn-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(132px,1fr)); gap:16px; }
        .fn-item { background:var(--dark-card-2); border:1px solid var(--border-color);
                   border-radius:10px; padding:13px 15px; }
        .fn-lbl { font-size:10px; text-transform:uppercase; letter-spacing:0.5px;
                  color:var(--text-3); font-weight:600; margin-bottom:5px; }
        .fn-val { font-size:15px; font-weight:700; color:var(--text-1); }

        /* atalho para ver os instrumentos do nivel atual */
        a.prg-atalho {
          display:inline-block; margin-bottom:16px; padding:10px 18px;
          font-size:12.5px; font-weight:600; text-decoration:none !important;
          color:var(--txt-azul) !important; background:rgba(var(--rgb-azul),0.12);
          border:1px solid rgba(var(--rgb-azul),0.28); border-radius:9px; transition:background 120ms;
        }
        a.prg-atalho:hover { background:rgba(var(--rgb-azul),0.22); }

        /* arvore SOP > SSOP > TAGs em <details>: expande sem ida ao servidor */
        .arvore { display:flex; flex-direction:column; gap:8px; }
        .arv-no { border:1px solid var(--border-color); border-radius:11px; overflow:hidden;
                  background:var(--dark-card-2); }
        .arv-no[open] { border-color:var(--border-strong); }
        /* Grid, e nao flex: no flex o conteudo se amontoa a esquerda e a barra
           vai sozinha para a direita com margin-left:auto, deixando uns 40% da
           linha vazios no meio. Com colunas fixas cada campo cai sempre no
           mesmo lugar e a barra fica com o espaco que sobrava. */
        .arv-no > summary, .arv-folha > .arv-linha {
          display:grid; align-items:center; gap:14px; padding:13px 16px;
          grid-template-columns:17px minmax(140px,1fr) 46px 92px 74px 116px 132px minmax(150px,1.6fr);
          cursor:pointer; list-style:none; user-select:none; transition:background 120ms;
        }
        .arv-no > summary::-webkit-details-marker { display:none; }
        .arv-no > summary:hover { background:rgba(var(--rgb-tinta),0.035); }
        .arv-no[open] > summary { background:rgba(var(--rgb-tinta),0.03);
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
        /* o codigo e link para a ficha; o + fica a cargo do <summary>.
           sem sublinhado: so muda de cor ao passar o mouse. */
        a.arv-ficha { text-decoration:none !important; color:var(--text-1) !important;
                      transition:color 120ms; }
        a.arv-ficha:hover { color:var(--txt-azul) !important; }
        .arv-vazio { width:17px; }
        .arv-num { font-size:11.5px; color:var(--text-3); white-space:nowrap; font-variant-numeric:tabular-nums; }
        .arv-sub { color:var(--text-2); font-weight:600; }
        .arv-val { color:var(--text-2); }
        /* a barra ocupa a coluna inteira; a porcentagem tem largura propria
           para os digitos ficarem alinhados de uma linha para outra */
        .arv-avanco { display:grid; grid-template-columns:1fr 52px; align-items:center; gap:11px; }
        .arv-track { height:8px; background:rgba(var(--rgb-tinta),0.07); border-radius:5px; overflow:hidden; }
        .arv-fill { display:block; height:100%; border-radius:4px; }
        .arv-fill.ok { background:var(--accent-teal); }
        .arv-fill.warn { background:var(--accent-amber); }
        .arv-fill.crit { background:var(--accent-red); }
        .arv-pct { font-size:12.5px; font-weight:700; color:var(--text-1);
                   text-align:right; font-variant-numeric:tabular-nums; }
        .arv-fill { border-radius:5px; }
        .arv-corpo { padding:12px 14px 14px; }
        .arv-n2 { background:var(--dark-card); margin-bottom:7px; }
        .arv-n2:last-child { margin-bottom:0; }
        .arv-n2 > summary { padding:11px 14px; }
        .arv-n2 .arv-nome { font-weight:600; font-size:12.5px; }
        .arv-n3 { background:var(--dark-card-2); margin-bottom:6px; }
        .arv-n3:last-child { margin-bottom:0; }
        .arv-n3 .arv-nome { font-weight:600; font-size:12.5px; }
        .arv-n3 > summary { padding:10px 14px; }
        .arv-n4 { background:var(--dark-card); margin-bottom:6px; }
        .arv-n4:last-child { margin-bottom:0; }
        .arv-n4 .arv-nome { font-weight:500; font-size:12px; }
        .arv-n4 > .arv-linha { padding:10px 14px; }
        .arv-n4 > summary { padding:10px 14px; }
        .arv-tags { padding:2px 0 4px; }

        /* Carregando. Acima do modal (1000000) e da sidebar (999991), senao a
           tela de carga aparece por baixo da navegacao. */
        .gpl { display:flex; flex-direction:column; align-items:center; justify-content:center;
               gap:14px; padding:54px 20px; }
        .gpl-cheia, .gpl-vidro {
          position:fixed; inset:0; z-index:1000001;
          animation:gpl-entra 220ms ease-out;
        }
        .gpl-cheia { background:var(--dark-bg); }
        /* desfoca em vez de tapar: a tela anterior continua ali atras, entao a
           troca parece continuacao e nao um recomeco do zero */
        .gpl-vidro { background:var(--overlay); backdrop-filter:blur(7px);
                     -webkit-backdrop-filter:blur(7px); }
        @keyframes gpl-entra { from { opacity:0; } to { opacity:1; } }
        /* Sai dissolvendo. Remover o elemento o apagaria num quadro so, que e
           exatamente o corte seco. Fica invisivel e sem captar clique. */
        .gpl-saindo { animation:gpl-sai 420ms ease-in forwards; pointer-events:none; }
        @keyframes gpl-sai { from { opacity:1; } to { opacity:0; visibility:hidden; } }
        .gpl-corpo { display:flex; flex-direction:column; align-items:center; gap:14px;
                     opacity:0; animation:gpl-surge 300ms ease-out 130ms forwards; }
        @keyframes gpl-surge { to { opacity:1; } }
        .gpl-mark { width:58px; height:58px; animation:gpl-bate 1.5s ease-in-out infinite; }
        .gpl-mark svg { width:100%; height:100%; }
        @keyframes gpl-bate { 0%,100% { opacity:0.5; transform:scale(0.93); }
                              50%     { opacity:1;   transform:scale(1); } }
        .gpl-nome { font-size:18px; font-weight:800; color:var(--text-1); letter-spacing:-0.4px; }
        .gpl-txt { font-size:12.5px; color:var(--text-2); min-height:16px; }
        .gpl-track { width:250px; height:6px; border-radius:4px; overflow:hidden;
                     background:rgba(var(--rgb-tinta),0.08); }
        .gpl-fill { height:100%; border-radius:4px; transition:width 260ms ease;
                    background:linear-gradient(90deg,var(--accent-blue),var(--accent-teal)); }
        /* sem etapa conhecida a barra corre de ponta a ponta, sem fingir % */
        .gpl-indet { width:40%; animation:gpl-corre 1.15s ease-in-out infinite; }
        @keyframes gpl-corre { 0% { margin-left:-40%; } 100% { margin-left:100%; } }

        /* graficos de "mais avancados" em grid 2x2: as linhas do grid tem
           altura uniforme, entao as colunas nunca desalinham. */
        /* align-items:start deixa cada card com a altura do seu conteudo; com
           stretch, um card de 1 barra ficava do tamanho do de 5 e sobrava um
           vao de 200px+ dentro dele. */
        /* os quatro lado a lado; abaixo de ~1250px cai para duas colunas */
        .gr-grid {
          display:grid; grid-template-columns:repeat(4, 1fr);
          gap:16px; margin-bottom:8px; align-items:start;
        }
        @media (max-width: 1250px) { .gr-grid { grid-template-columns:repeat(2, 1fr); } }
        @media (max-width: 700px)  { .gr-grid { grid-template-columns:1fr; } }
        /* height:auto anula o height:100% de .gplan-panel, que esticava o card
           de 1 barra ate a altura do de 5 (245px de vao interno). */
        /* compacto: quatro cards por linha exigem menos folga interna */
        .gr-panel { padding:18px 18px; margin-bottom:0 !important; height:auto !important; }
        .gr-panel .gplan-panel-title { margin-bottom:16px; font-size:13px; }
        .gr-row { margin-bottom:13px; }
        .gr-row:last-child { margin-bottom:0; }
        .gr-top { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:5px; gap:8px; }
        .gr-nome { font-size:11.5px; font-weight:600; color:var(--text-1); white-space:nowrap;
                   overflow:hidden; text-overflow:ellipsis; }
        .gr-pct { font-size:12px; font-weight:800; color:var(--accent-teal);
                  font-variant-numeric:tabular-nums; flex-shrink:0; }
        .gr-track { height:6px; background:rgba(var(--rgb-tinta),0.05); border-radius:4px; overflow:hidden; }
        .gr-fill { height:100%; border-radius:4px;
                   background:linear-gradient(90deg, var(--accent-teal), var(--teal-2)); transition:width 400ms ease; }
        .gr-sub { font-size:9.5px; color:var(--text-3); margin-top:4px; white-space:nowrap;
                  overflow:hidden; text-overflow:ellipsis; }

        /* barra de avanco dentro da celula da tabela */
        .cel-avanco { display:flex; align-items:center; gap:9px; justify-content:flex-end; }
        .cel-track { width:62px; height:6px; background:rgba(var(--rgb-tinta),0.06); border-radius:4px; overflow:hidden; flex-shrink:0; }
        .cel-fill { height:100%; border-radius:4px; }
        .cel-fill.ok { background:var(--accent-teal); }
        .cel-fill.warn { background:var(--accent-amber); }
        .cel-fill.crit { background:var(--accent-red); }
        .cel-pct { font-size:12px; font-weight:600; min-width:46px; text-align:right; font-variant-numeric:tabular-nums; }

        .flt-summary { font-size:12.5px; color:var(--text-2); padding:2px 2px 0; }
        .flt-summary strong { color:var(--text-1); }

        .sg-corpo { display:flex; align-items:center; gap:20px; }
        .sg-chart-wrap { position:relative; width:148px; flex-shrink:0; margin:2px 0; }
        .sg-legend { flex:1; min-width:0; }
        @media (max-width:780px) { .sg-corpo { flex-direction:column; }
                                   .sg-chart-wrap { width:180px; } }
        .sg-donut { width:100%; height:auto; display:block; }
        .sg-center {
          position:absolute; inset:0; display:flex; flex-direction:column;
          align-items:center; justify-content:center; pointer-events:none;
        }
        .sg-center-value { font-size:20px; font-weight:800; color:var(--text-1); letter-spacing:-0.6px; line-height:1; }
        .sg-center-label { font-size:10px; color:var(--text-3); margin-top:3px; }
        .sg-legend { display:flex; flex-direction:column; }
        .sg-leg-row {
          display:flex; align-items:center; gap:10px; padding:7px 8px;
          margin: 0 -8px; border-radius:7px;
          border-bottom:1px solid rgba(var(--rgb-tinta),0.04);
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
        a.rep-row:hover { background: rgba(var(--rgb-tinta),0.035); }
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
        a.sg-leg-row:hover { background: rgba(var(--rgb-tinta),0.035); }
        a.sg-leg-row:hover .sg-leg-name { color: var(--text-1); }

        .rep-row { margin-bottom: 14px; }
        .rep-row:last-child { margin-bottom: 0; }
        .rep-label { display:flex; justify-content:space-between; margin-bottom:6px; font-size:12.5px; }
        .rep-name { font-weight:600; color: var(--text-1); }
        .rep-stat { color: var(--text-3); font-variant-numeric: tabular-nums; }
        .rep-track { height:7px; background: rgba(var(--rgb-tinta),0.05); border-radius:4px; overflow:hidden; display:flex; }
        .rep-done { background: linear-gradient(90deg, var(--accent-teal), var(--teal-2)); height:100%; }
        .rep-pending { background: rgba(var(--rgb-tinta),0.05); height:100%; }
        .doc-tag { font-size:9px; font-weight:600; color: var(--text-3); background: rgba(var(--rgb-tinta),0.06); padding:1px 6px; border-radius:4px; text-transform:uppercase; letter-spacing:0.3px; margin-left:6px; }

        .group-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap:16px; }
        .group-card-v2 { background: var(--dark-card-2); border: 1px solid var(--border-color); border-radius: 12px; padding: 18px; }
        .group-card-top { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:12px; }
        .group-card-name { font-size:13px; font-weight:700; color: var(--text-1); }
        .group-card-pct { font-size:13px; font-weight:800; color: var(--accent-teal); }
        .group-card-value { font-size:24px; font-weight:800; color: var(--text-1); letter-spacing:-0.4px; margin-bottom:12px; }
        .group-card-unit { font-size:11px; font-weight:500; color: var(--text-3); }
        .group-card-track { height:6px; background: rgba(var(--rgb-tinta),0.06); border-radius:3px; overflow:hidden; margin-bottom:12px; }
        .group-card-fill { height:100%; border-radius:3px; background: linear-gradient(90deg, var(--accent-teal), var(--teal-2)); }
        .group-card-nums { display:flex; justify-content:space-between; font-size:11px; color: var(--text-3); }

        .tag-detail-card { background: var(--dark-card); border: 1px solid var(--border-strong); border-radius: 16px; padding: 26px 28px; }
        .detail-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap:12px; margin-bottom:26px; }
        .detail-grid .detail-item { margin-bottom: 0; }
        /* Os 16 campos da ficha se repetem em 5.098 TAGs. Escrever as tres
           classes em cada um custava 1,4 KB por ficha, uns 7 MB no total, e a
           pagina inteira precisa caber no navegador. Aqui as classes saem e o
           estilo vem pela posicao -- nada muda visualmente. */
        .detail-grid > div { background:var(--dark-card-2); border:1px solid var(--border-color);
                             border-radius:12px; padding:14px 16px; }
        .detail-grid > div > span:first-child { display:block; font-size:10.5px;
          text-transform:uppercase; letter-spacing:0.5px; color:var(--text-3);
          font-weight:600; margin-bottom:6px; }
        .detail-grid > div > span:last-child { display:block; font-size:15px;
          font-weight:700; color:var(--text-1); }
        /* mesma coisa na tabela de relatorios: sao 25.095 linhas somadas */
        .gtbl-rel td:nth-child(2) { color:var(--text-2); }
        .gtbl-rel td:nth-child(3) { font-size:12px; color:var(--text-2); }
        .gtbl-rel td:nth-child(5) { text-align:center; color:var(--text-2);
                                    font-variant-numeric:tabular-nums; }
        .ficha-head { margin-bottom: 20px; }
        .ficha-tag { font-size:22px; font-weight:800; color:var(--text-1); letter-spacing:-0.5px; }
        .ficha-desc { font-size:13px; color:var(--text-2); margin-top:3px; }
        .ficha-sub { font-size:13.5px; font-weight:700; color:var(--text-1); margin-bottom:14px; }

        /* Modal da ficha resolvido por :target -- abre sem ida ao servidor. */
        /* acima da sidebar do Streamlit, que usa z-index 999991 */
        .fmodal { display:none; position:fixed; inset:0; z-index:1000000; }
        .fmodal:target,
        .fmodal-on { display:flex; align-items:center; justify-content:center; padding:32px 20px; }
        .fmodal-bg {
          position:absolute; inset:0; background:var(--overlay);
          backdrop-filter:blur(6px); -webkit-backdrop-filter:blur(6px);
        }
        .fmodal-box {
          position:relative; width:min(1080px, 100%); max-height:88vh; overflow:auto;
          background:var(--dark-card); border:1px solid var(--border-strong);
          border-radius:16px; box-shadow:0 24px 70px var(--sombra);
          animation: fmodal-in 140ms ease-out;
        }
        @keyframes fmodal-in { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
        .fmodal-head {
          display:flex; align-items:center; gap:22px;
          padding:22px 28px 0; position:sticky; top:0; background:var(--dark-card); z-index:1;
        }
        .fmodal-head > div:first-child { flex:1; }
        .fmodal-head .fn-avanco { flex-shrink:0; }
        .fmodal-title { font-size:20px; font-weight:800; color:var(--text-1); letter-spacing:-0.4px; }
        .fmodal-x {
          font-size:26px; line-height:1; color:var(--text-3) !important; text-decoration:none !important;
          padding:0 6px; border-radius:8px;
        }
        .fmodal-x:hover { color:var(--text-1) !important; background:rgba(var(--rgb-tinta),0.06); }
        .fmodal-body { padding:16px 28px 28px; }
        .fmodal-body .ficha-sub { margin-top:4px; }

        /* Fundo desfocado atras do modal da ficha. */
        div[data-testid="stDialog"] > div:first-child,
        div[data-baseweb="modal"] > div:first-child {
          backdrop-filter: blur(6px) !important;
          -webkit-backdrop-filter: blur(6px) !important;
          background: var(--overlay) !important;
        }
        /* o painel do dialog e um <section>, nao um <div> */
        div[data-testid="stDialog"] [role="dialog"] {
          background: var(--dark-card) !important;
          border: 1px solid var(--border-strong) !important;
          border-radius: 16px !important;
          box-shadow: 0 24px 70px var(--sombra) !important;
        }
        div[data-testid="stDialog"] [role="dialog"] h2 {
          font-size: 20px !important; font-weight: 800 !important;
          letter-spacing: -0.4px; padding-bottom: 4px;
        }

        .detail-item { background: var(--dark-card-2); border: 1px solid var(--border-color); border-radius: 12px; padding: 14px 16px; margin-bottom: 12px; }
        .detail-label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-3); font-weight: 600; margin-bottom: 6px; }
        .detail-value { font-size: 15px; font-weight: 700; color: var(--text-1); }

        div[data-testid="stMetric"] { background: var(--dark-card); border: 1px solid var(--border-color); border-radius: 12px; padding: 12px 16px; }

        /* ---------------------------------------------------- aba Planta */
        .pl-kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:16px; }
        .pl-kpi { background:var(--dark-card); border:1px solid var(--border-color);
                  border-radius:14px; padding:15px 18px; }
        .pl-kpi .r { font-size:10px; letter-spacing:.75px; text-transform:uppercase;
                     color:var(--text-3); font-weight:700; }
        .pl-kpi .v { font-size:29px; font-weight:800; letter-spacing:-1.2px;
                     line-height:1.1; color:var(--text-1); white-space:nowrap;
                     margin-top:7px; }
        /* o valor por extenso e longo: encolhe com a coluna em vez de estourar */
        .pl-kpi .v.dinheiro { font-size:clamp(15px, 1.42vw, 23px); letter-spacing:-.6px; }
        .pl-kpi .v.andando { color:var(--accent-amber); }
        .pl-kpi .v.feito { color:var(--accent-teal); }
        .pl-kpi .v.parado { color:var(--accent-red); }
        .pl-kpi .s { font-size:11px; color:var(--text-3); margin-top:6px; }
        .pl-barra { height:8px; border-radius:99px; margin-top:10px;
                    background:rgba(var(--rgb-tinta),.07); overflow:hidden; }
        .pl-barra i { display:block; height:100%; border-radius:99px; }
        .pl-barra i.feito { background:var(--accent-teal); }
        .pl-barra i.andando { background:var(--accent-amber); }
        .pl-barra i.parado { background:var(--accent-red); }

        .pl-pn { padding:16px 16px 14px; margin-bottom:14px; }
        .pl-pn .gplan-panel-title { display:flex; align-items:baseline; gap:12px;
                                    margin-bottom:12px; font-size:13px; }
        .pl-res { margin-left:auto; font-size:11px; font-weight:500; color:var(--text-3); }
        .pl-res b { font-weight:800; font-size:13px; }
        .pl-res .feito { color:var(--accent-teal); }
        .pl-res .andando { color:var(--accent-amber); }
        .pl-res .parado { color:var(--accent-red); }

        .pl-tela { position:relative; width:100%; height:0;
                   border:1px solid var(--border-color); border-radius:11px;
                   overflow:hidden; background:var(--fundo-3); }
        /* O desenho vem preto sobre branco. Invertido, o traco fica claro sobre
           o escuro e a prancha deixa de ser um retangulo branco no meio de uma
           tela escura -- e as zonas passam a ler por cima dele. */
        .pl-tela img { position:absolute; inset:0; width:100%; height:100%;
                       object-fit:fill; filter:var(--planta-filtro);
                       opacity:var(--planta-opacidade); }

        /* a zona virou link para abrir a ficha: sem isto o navegador sublinha
           codigo, percentual e contagem */
        .pl-zona { position:absolute; display:grid; place-items:center; border:1.5px solid;
                   border-radius:8px; overflow:hidden; text-decoration:none !important;
                   transition:filter .13s, box-shadow .13s; }
        .pl-zona:hover, .pl-zona:visited { text-decoration:none !important; }
        .pl-zona.rec { border-radius:3px; place-items:end start; padding:0 0 8px 8px; }
        .pl-zona.feito { border-color:var(--accent-teal); background:rgba(var(--rgb-teal),.10);
                         color:var(--accent-teal); }
        .pl-zona.andando { border-color:var(--accent-amber); background:rgba(var(--rgb-ambar),.10);
                           color:var(--accent-amber); }
        .pl-zona.parado { border-color:var(--accent-red); background:rgba(var(--rgb-vermelho),.09);
                          color:var(--accent-red); }
        /* O preenchimento sobe com o percentual: a zona e o proprio grafico. */
        .pl-zona::before { content:""; position:absolute; left:0; right:0; bottom:0;
                           height:calc(var(--p) * 1%); background:currentColor; opacity:.22; }
        .pl-zona:hover { filter:brightness(1.35); box-shadow:0 0 0 2px currentColor; z-index:9; }
        .pl-mio { position:relative; max-width:100%; display:flex; flex-direction:column;
                  align-items:center; text-align:center; line-height:1.15;
                  padding:5px 8px; border-radius:8px; background:rgba(var(--rgb-chapa),.72); }
        .pl-zona.rec .pl-mio { align-items:flex-start; text-align:left;
                               max-width:calc(100% - 18px); }
        .pl-mio b { font-size:var(--fs,11px); font-weight:750; color:var(--text-1);
                    letter-spacing:-.1px; }
        .pl-mio i { font-style:normal; font-size:calc(var(--fs,11px) * 1.4);
                    font-weight:800; letter-spacing:-.4px; line-height:1.1; }
        .pl-mio u { text-decoration:none; font-size:calc(var(--fs,11px) * .82);
                    color:var(--text-2); font-weight:600; }
        /* A etiqueta da area: fundo na cor do status, texto no fundo da tela.
           Nao da para usar currentColor nos dois -- definir a cor do texto
           redefine o currentColor da etiqueta, e ela some contra o proprio fundo. */
        .pl-ar { position:absolute; top:0; left:0; font-size:9px; font-weight:800;
                 letter-spacing:.4px; color:var(--sobre-cor); padding:1px 6px;
                 border-radius:0 0 7px 0; line-height:1.5; }
        .pl-zona.feito .pl-ar { background:var(--accent-teal); }
        .pl-zona.andando .pl-ar { background:var(--accent-amber); }
        .pl-zona.parado .pl-ar { background:var(--accent-red); }

        /* A lista de instrumentos rola dentro do painel. Paginar aqui obrigaria
           a fechar a ficha para trocar de pagina -- o modal e :target puro, nao
           guarda estado -- e a area 140 tem 749 instrumentos. O cabecalho fica
           preso no topo, senao some na primeira rolada. */
        .pl-rol { max-height:min(46vh, 430px); overflow-y:auto; overflow-x:auto; }
        .pl-rol .gtbl-scroll { overflow:visible; }
        .pl-rol thead th { position:sticky; top:0; z-index:2;
                           background:var(--dark-card-2); }
        .pl-rol td.desc { color:var(--text-2); max-width:200px; overflow:hidden;
                          text-overflow:ellipsis; white-space:nowrap; }
        .pl-sim { display:inline-block; font-size:10.5px; font-weight:700;
                  color:var(--txt-verde); background:rgba(var(--rgb-verde),.13);
                  border:1px solid rgba(var(--rgb-verde),.28); border-radius:6px; padding:1px 7px; }
        .pl-nao { color:var(--text-3); }

        .pl-lg { padding:16px; }
        .pl-ch-c { display:flex; flex-direction:column; gap:9px; }
        .pl-ch { display:flex; align-items:center; gap:11px; background:var(--dark-card-2);
                 border:1px solid var(--border-color); border-radius:11px; padding:10px 12px; }
        .pl-ch .sw { width:14px; height:28px; border-radius:5px; border:1.5px solid currentColor;
                     flex:none; background:linear-gradient(180deg,transparent 45%,currentColor 45%); }
        .pl-ch.feito { color:var(--accent-teal); }
        .pl-ch.andando { color:var(--accent-amber); }
        .pl-ch.parado { color:var(--accent-red); }
        .pl-ch .tx b { display:block; font-size:12px; font-weight:750; color:var(--text-1); line-height:1.3; }
        .pl-ch .tx em { font-style:normal; font-size:10.5px; color:var(--text-3); }
        .pl-ch .qt { margin-left:auto; text-align:right; }
        .pl-ch .qt b { display:block; font-size:16px; font-weight:800; line-height:1.2; }
        .pl-ch .qt em { font-style:normal; font-size:10px; color:var(--text-3); }
        @media (max-width:1250px) { .pl-kpis { grid-template-columns:repeat(2,1fr); } }
        @media (max-width:620px) { .pl-kpis { grid-template-columns:1fr; } }
        </style>
        """.replace("__ICONES__", fx_css_icones())
            .replace("__TOKENS__", tokens_css(tema_ativo()))
    )


def data_atualizacao(cache_key: str) -> str:
    """Quando a PLANILHA foi atualizada, no horario de Brasilia.

    Antes exibia pd.Timestamp.now(), que era a hora de renderizar a pagina:
    mudava a cada clique e, no Render (que roda em UTC), aparecia 3h adiantada.
    O cache_key ja carrega o updated_at do arquivo no Supabase; localmente e o
    mtime do xlsx.

    Sao dois formatos: o Supabase devolve ISO-8601 e o disco devolve epoch. A
    diferenca e que epoch converte para float e ISO nao -- distinguir pelos
    quatro primeiros digitos nao serve, porque "1786369744" tambem comeca com
    quatro digitos e caia no ramo errado, deixando o cabecalho em "—".
    """
    cache_key = str(cache_key or "").split("|")[0]      # tolera a chave de cache
    try:
        ts = pd.Timestamp(float(cache_key), unit="s", tz="UTC")   # epoch do disco
    except (TypeError, ValueError):
        try:
            ts = pd.Timestamp(cache_key)                          # ISO do Supabase
            ts = ts.tz_localize("UTC") if ts.tzinfo is None else ts.tz_convert("UTC")
        except Exception:
            return "—"
    return ts.tz_convert(BR_TZ).strftime("%d/%m/%Y %H:%M")


@contextmanager
def carregando(texto: str):
    """Cobre a tela enquanto a aba monta, e descobre no fim.

    O Streamlit desenha de cima para baixo, entao sem isso a pagina aparece em
    pedacos por varios segundos. A cobertura sai no finally: se a montagem
    quebrar, o erro tem que ficar visivel em vez de uma tela de carga eterna.

    Guardada em session_state, e nao numa global: o processo atende varias
    sessoes ao mesmo tempo e uma nao pode escrever na tela da outra.
    """
    # primeira abertura da sessao tapa com cor; dai em diante desfoca o que ja
    # esta na tela, para a troca parecer continuacao e nao recomeco
    vidro = st.session_state.get("_ja_abriu", False)
    st.session_state["_ja_abriu"] = True

    capa = st.empty()
    capa.markdown(tela_carregando(texto, vidro=vidro), unsafe_allow_html=True)
    st.session_state["_capa"] = (capa, texto, vidro, time.monotonic())
    try:
        yield
    finally:
        descobrir()


def descobrir():
    """Dissolve a tela de carga. Pode ser chamada mais de uma vez.

    Nao usa empty(): remover o elemento o faz sumir num quadro, que e o corte
    seco. Aqui ele e trocado por uma copia que se apaga sozinha e fica
    invisivel e sem captar clique.
    """
    guardado = st.session_state.pop("_capa", None)
    if guardado is None:
        return
    capa, texto, vidro, inicio = guardado
    falta = CARGA_MINIMA - (time.monotonic() - inicio)
    if falta > 0:
        time.sleep(falta)
    capa.markdown(tela_carregando(texto, vidro=vidro, saindo=True),
                  unsafe_allow_html=True)


def _sob_carga(texto: str, montar):
    """Roda a montagem da aba por tras da tela de carga."""
    with carregando(texto):
        montar()


def render_header(title: str, extra_pill: str | None = None):
    now = st.session_state.get("gplan_atualizado_em", "—")
    extra_html = f'<div class="gplan-count-pill">{extra_pill}</div>' if extra_pill else ""
    # O filtro da lateral vale para o app inteiro, entao ele tem que aparecer
    # em toda aba. Sem isso alguem abre Relatorios, ve 464 tags em vez de 5.098
    # e nao tem como saber que esta olhando um recorte.
    filtro = st.session_state.get("_flt_selo", "")
    render_html(
        f"""
        <div class="gplan-header">
          <h1>{title}</h1>
          <div style="display:flex; gap:12px; align-items:center;">
            {filtro}
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
    "calendario": ('<rect x="3" y="5" width="18" height="16" rx="2"/>'
                   '<path d="M3 10h18M8 3v4M16 3v4"/>'
                   '<rect x="7" y="13" width="3" height="3" rx="0.5" fill="currentColor"/>'),
}


def br_num(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def br_pct(value: float, casas: int = 1) -> str:
    """Porcentagem com virgula. So para texto lido na tela -- largura de barra
    em CSS continua com ponto, senao o navegador descarta a regra."""
    return f"{value:.{casas}f}%".replace(".", ",")


def com_filtros(href: str) -> str:
    """Cola os filtros ativos num link.

    O menu da lateral troca de aba sem recarregar e o session_state segue
    vivo. Estes links nao: sao navegacao de verdade, o navegador recarrega e
    a sessao comeca do zero. Sem os filtros na URL, filtrar por uma fase e
    clicar numa barra do Dashboard devolvia Relatorios com as 5.098 tags de
    volta -- e sem nada na tela dizendo que o filtro caiu. De quebra, o
    endereco passa a poder ser copiado e colado com o filtro junto.
    """
    # Nao repete o que o proprio link ja traz: a trilha da ficha manda
    # /progresso?fase=X, e colar o fase= atual em cima criaria dois valores
    # para a mesma chave -- o Streamlit fica com um deles e o outro some sem
    # aviso, o que daria um filtro diferente do que foi clicado.
    ja_tem = href.split("?", 1)[1] if "?" in href else ""
    partes = [f"{nome}={quote(valor)}"
              for chave, _rot, padrao, _f, _c in FILTROS
              if (nome := URL_DO_FILTRO.get(chave, chave)) and f"{nome}=" not in ja_tem
              and (valor := st.session_state.get(f"gf_{chave}", padrao)) and valor != padrao]
    if not partes:
        return href
    return href + ("&" if "?" in href else "?") + "&".join(partes)


def esc(value: object) -> str:
    text = "" if value is None else str(value)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def tag_pill(value: object) -> str:
    return f'<span class="gtbl-tag">{esc(value)}</span>'


def status_badge(status: str) -> str:
    """Badge com a cor do proprio status, reaproveitando o mapa usado no donut."""
    label = sentence_case(status)
    return (f'<span class="gtbl-badge {classe_status(label)}">{esc(label)}</span>')


def yes_no_badge(value: object) -> str:
    yes = str(value).strip().upper() in {"SIM", "SIM.", "S", "YES", "TRUE"}
    return f'<span class="gtbl-badge {"ok" if yes else "crit"}">{"Sim" if yes else "Não"}</span>'


def html_table(headers: list, rows_html: str, empty_msg: str = "Nenhum registro encontrado.",
               classe: str = "gtbl") -> str:
    """Tabela HTML estilizada, substituindo o st.dataframe (grid do Streamlit),
    que e um canvas fechado e nao aceita estilizacao."""
    if not rows_html:
        return f'<div class="gtbl-empty">{empty_msg}</div>'
    head = "".join(
        f'<th class="gtbl-num">{h[1:]}</th>' if h.startswith("#") else f"<th>{h}</th>"
        for h in headers
    )
    return (
        f'<div class="gtbl-scroll"><table class="{classe}">'
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
                       tags: pd.DataFrame, espera_por_doc: dict | None = None,
                       niveis_na_pagina: bool = False) -> str:
    """Modais das tags visiveis na pagina, abertos/fechados via CSS :target."""
    ids = list(dict.fromkeys(tags_ids))  # unicas, preservando a ordem
    # Indexar uma vez. Antes cada ficha varria resumo, esperados e tags
    # inteiros: com as 1.209 TAGs de uma pagina davam 42 milhoes de comparacoes
    # por carregamento, e era isso que levava a aba a 115 s no Render.
    def por_tag(df):
        return {k: v for k, v in df[df["TAG"].isin(ids)].groupby("TAG")}

    g_resumo, g_esp, g_tags = por_tag(resumo), por_tag(esperados), por_tag(tags)
    vazio_esp, vazio_tags = esperados.iloc[0:0], tags.iloc[0:0]

    blocos = ""
    for tag_id in ids:
        if tag_id not in g_resumo:
            continue
        corpo = tag_ficha_html(tag_id, g_resumo[tag_id], g_esp.get(tag_id, vazio_esp),
                               g_tags.get(tag_id, vazio_tags), com_cabecalho=False,
                               espera_por_doc=espera_por_doc,
                               niveis_na_pagina=niveis_na_pagina)
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


TOPO_RECUSADOS = 7


def painel_recusados(esperados: pd.DataFrame, sigem: pd.DataFrame) -> str:
    """O que esta parado ha mais tempo por recusa da fiscalizacao.

    Cada recusa e de uma TAG diferente -- 722 recusados em 722 TAGs, nenhuma
    com duas -- entao agrupar por TAG daria a mesma lista. A linha traz os
    dois: a TAG e o relatorio dela que voltou.
    """
    rec = esperados[esperados["STATUS_SIGEM"].astype(str).str.strip().str.upper() == "RECUSADO"]
    if rec.empty:
        return ('<div class="gplan-panel"><div class="gplan-panel-title">Recusados há mais tempo</div>'
                '<div class="gtbl-empty">Nenhum relatório recusado.</div></div>')

    dt = pd.to_datetime(sigem["DATA"], dayfirst=True, errors="coerce")
    ultima = pd.DataFrame({"doc": sigem["DOCUMENTO"], "dt": dt}).dropna() \
        .sort_values("dt").groupby("doc")["dt"].last()
    hoje = pd.Timestamp.now(tz=BR_TZ).tz_localize(None).normalize()
    rec = rec.assign(_dt=rec["DOCUMENTO_ESPERADO"].map(ultima))
    rec = rec.assign(_dias=(hoje - rec["_dt"]).dt.days)
    piores = rec.dropna(subset=["_dias"]).sort_values("_dias", ascending=False).head(TOPO_RECUSADOS)

    linhas = "".join(
        f'<a class="rec-linha" href="{com_filtros("/relatorios?busca=" + quote(str(t)))}" target="_self" '
        f'title="Ver os relatórios de {esc(t)}">'
        f'<span class="rec-tag">{esc(t)}</span>'
        f'<span class="rec-rel">{esc(r)}</span>'
        f'<span class="rec-dias">{br_num(int(d))} dias</span></a>'
        for t, r, d in zip(piores["TAG"], piores["RELATORIO"], piores["_dias"])
    )
    media = int(rec["_dias"].mean())
    return (
        '<div class="gplan-panel">'
        '<div class="gplan-panel-title">Recusados há mais tempo</div>'
        f'<div class="rec-resumo">{br_num(len(rec))} relatórios recusados · '
        f'{br_num(media)} dias parados em média</div>'
        f'<div class="rec-lista">{linhas}</div></div>' 
    )


def fichas_completas(ids, resumo: pd.DataFrame, esperados: pd.DataFrame,
                     tags: pd.DataFrame, sigem: pd.DataFrame, cache_key: str = "") -> str:
    """Fichas das TAGs pedidas MAIS as dos relatorios que elas listam.

    As duas andam juntas: a ficha da TAG mostra todos os relatorios dela, e
    cada um tem um botao Detalhes. Gerar so as fichas dos documentos que
    aparecem na tabela da pagina deixa esses botoes apontando para ancoras que
    nao existem, e o clique nao faz nada -- aconteceu no Dashboard e em
    Relatorios, com 43 dos 83 botoes sem destino. Chamar isto no lugar de
    fichas_modais_html impede que os dois voltem a divergir.
    """
    alvo = set(ids)
    meus = esperados[esperados["TAG"].isin(alvo)]
    docs = meus[meus["STATUS_SIGEM"].map(POSTADO)]["DOCUMENTO_ESPERADO"].tolist()
    historico = _revisoes_por_doc(cache_key, sigem)
    espera = espera_por_documento(historico)
    return (fichas_modais_html(ids, resumo, esperados, tags, espera)
            + fichas_relatorios_html(docs, esperados, historico))


def du_tile(cor: str, icone: str) -> str:
    """O quadradinho do icone. A cor entra como classe, nunca no style.

    Pintar o hex no atributo prendia o cartao ao tema escuro: no claro o
    #fbbf24 do icone sobre a propria lavagem clara dava 1,5:1.
    """
    svg = KPI_ICONS.get(icone, "")
    return (f'<span class="tile {fx_classe_cor(cor)}">'
            f'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round">{svg}</svg></span>')


def du_kpi(rot: str, val: str, sub: str, pct: float, cor: str, icone: str, href: str = "") -> str:
    corpo = (f'<div class="topo">{du_tile(cor, icone)}<div class="rot">{rot}</div></div>'
             f'<div class="val">{val}</div><div class="sub">{sub}</div>'
             f'<div class="du-trilho {fx_classe_cor(cor)}">'
             f'<i style="width:{max(0.0, min(pct, 1.0)) * 100:.1f}%;"></i></div>')
    if href:
        return f'<a class="du-kpi" href="{com_filtros(href)}" target="_self">{corpo}</a>'
    return f'<div class="du-kpi">{corpo}</div>'


def du_grupos(resumo: pd.DataFrame) -> str:
    g = resumo.groupby("GRUPO_REGRA").agg(
        tags=("TAG", "count"), esperados=("RELATORIOS_ESPERADOS", "sum"),
        aprovados=("RELATORIOS_APROVADOS", "sum"),
    ).reset_index()
    g["avanco"] = (g["aprovados"] / g["esperados"]).fillna(0) * 100
    g = g.sort_values("tags", ascending=False)
    cartoes = "".join(
        f'<a class="du-gp" href="{com_filtros("/?grupo=" + quote(sentence_case(r["GRUPO_REGRA"])))}" target="_self" '
        f'title="Filtrar tudo por {esc(sentence_case(r["GRUPO_REGRA"]))}">'
        f'<span class="lin"><span class="nm">{esc(sentence_case(r["GRUPO_REGRA"]))}</span>'
        f'<span class="pc">{br_pct(r["avanco"])}</span></span>'
        f'<span class="qt">{br_num(int(r["tags"]))}<em>tags</em></span>'
        f'<span class="du-trilho"><i style="width:{r["avanco"]:.1f}%;"></i></span>'
        f'<span class="pe"><span>{br_num(int(r["aprovados"]))} aprovados</span>'
        f'<span>{br_num(int(r["esperados"]))} esp.</span></span></a>'
        for _, r in g.iterrows())
    return ('<div class="du-pn"><div class="du-t">Resumo por grupo de instrumento</div>'
            f'<div class="du-miolo"><div class="du-grupos">{cartoes}</div></div></div>')


def du_barras(esperados: pd.DataFrame) -> str:
    # Maior avanco em cima: barras em escada leem melhor do que na ordem em
    # que as regras foram declaradas.
    linhas = []
    for label, report, origin, unico in REPORT_ROWS:
        esperado = count_rows(esperados, report, origin, False, unico)
        aprov = count_rows(esperados, report, origin, True, unico)
        pct = (aprov / esperado * 100) if esperado else 0.0
        linhas.append((pct, label, aprov, esperado))
    linhas.sort(key=lambda x: -x[0])
    html = "".join(
        f'<a class="du-br" href="{com_filtros("/relatorios?rel=" + quote(label))}" target="_self" '
        f'title="Ver {esc(label)} em Relatórios"><span class="nm">{esc(label)}</span>'
        f'<span class="tr"><i style="width:{max(pct, 0.6):.1f}%;"></i></span>'
        f'<span class="fr">{br_num(aprov)} / {br_num(esperado)}</span>'
        f'<span class="pc">{br_pct(pct)}</span></a>'
        for pct, label, aprov, esperado in linhas)
    return ('<div class="du-pn"><div class="du-t">Esperado × aprovado por tipo de relatório</div>'
            f'<div class="du-miolo"><div class="du-barras">{html}</div></div>'
            f'<a class="du-rodape" href="{com_filtros("/relatorios")}" target="_self">'
            'Ver todos os relatórios →</a></div>')


def du_status(esperados: pd.DataFrame) -> str:
    contagem = esperados["STATUS_SIGEM"].value_counts()
    itens = [(sentence_case(k), int(v)) for k, v in contagem.items()]
    total = int(len(esperados)) or 1
    raio, largura = 46.0, 13.0
    circ = 2 * math.pi * raio
    fatias, giro = "", 0.0
    for i, (rotulo, valor) in enumerate(itens):
        cor = STATUS_COLOR_MAP.get(rotulo, DEFAULT_STATUS_COLORS[i % len(DEFAULT_STATUS_COLORS)])
        comp = valor / total * circ
        fatias += (f'<a href="{com_filtros("/relatorios?status=" + quote(rotulo))}" target="_self">'
                   f'<title>{esc(rotulo)} · {br_num(valor)}</title>'
                   f'<circle class="fatia {fx_classe_cor(cor)}" cx="60" cy="60" r="{raio}" fill="none" '
                   f'stroke-width="{largura}" stroke-dasharray="{comp:.2f} {circ - comp:.2f}" '
                   f'stroke-dashoffset="{-giro:.2f}" transform="rotate(-90 60 60)"></circle></a>')
        giro += comp
    legenda = ""
    for i, (rotulo, valor) in enumerate(itens):
        cor = STATUS_COLOR_MAP.get(rotulo, DEFAULT_STATUS_COLORS[i % len(DEFAULT_STATUS_COLORS)])
        legenda += (f'<a class="du-lg" href="{com_filtros("/relatorios?status=" + quote(rotulo))}" target="_self" '
                    f'title="Ver {esc(rotulo)} em Relatórios">'
                    f'<i class="{fx_classe_cor(cor)}"></i><span class="nm">{esc(rotulo)}</span>'
                    f'<span class="n">{br_num(valor)}</span>'
                    f'<span class="p">{br_pct(valor / total * 100)}</span></a>')
    return ('<div class="du-pn"><div class="du-t">Status SIGEM</div><div class="du-miolo"><div class="du-sigem">'
            f'<div class="du-rosca"><svg viewBox="0 0 120 120">'
            f'<circle class="trilho" cx="60" cy="60" r="{raio}" fill="none" '
            f'stroke-width="{largura}"></circle>{fatias}</svg>'
            f'<div class="centro"><b>{br_num(int(len(esperados)))}</b><span>relatórios</span></div></div>'
            f'<div class="du-leg">{legenda}</div></div></div></div>')


def du_top10(resumo: pd.DataFrame) -> str:
    top = resumo.sort_values(["RELATORIOS_PENDENTES", "RELATORIOS_ESPERADOS"],
                             ascending=False).head(10)
    linhas = ""
    for _, r in top.iterrows():
        pend, espe = int(r["RELATORIOS_PENDENTES"]), int(r["RELATORIOS_ESPERADOS"])
        razao = (pend / espe) if espe else 0
        tom = "crit" if razao >= 0.8 else ("warn" if razao >= 0.4 else "ok")
        linhas += (f'<div class="lin">{tag_link(r["TAG"])}'
                   f'<span class="tp">{esc(str(r["GRUPO_REGRA"]).title())}</span>'
                   f'<span class="num">{espe}</span>'
                   f'<span class="pnd"><span class="gtbl-badge {tom}">{pend}</span></span>'
                   f'<span class="cc">{br_pct(r["AVANCO_DOCUMENTAL"] * 100)}</span></div>')
    if not linhas:
        linhas = '<div class="lin"><span class="tp">Nenhuma tag no filtro.</span></div>'
    return ('<div class="du-pn"><div class="du-t">Top 10 tags com mais pendências</div><div class="du-miolo">'
            '<div class="du-tab"><div class="cabtab"><span>Tag</span><span>Tipo</span>'
            '<span class="num">Esper.</span><span class="pnd">Pend.</span>'
            '<span class="cc">Concl.</span></div>'
            f'<div class="corpo">{linhas}</div></div></div>'
            f'<a class="du-rodape" href="{com_filtros("/pesquisa")}" target="_self">'
            'Ver todas as tags pendentes →</a></div>')


def du_mini(rot: str, val: str, sub: str, cor: str, icone: str, href: str = "") -> str:
    classe = "val pq" if len(val) > 9 else "val"
    corpo = (f'{du_tile(cor, icone)}<span><span class="rot">{rot}</span>'
             f'<span class="{classe}" style="display:block;">{val}</span>'
             f'<span class="rot">{sub}</span></span>')
    if href:
        return f'<a class="du-mini" href="{com_filtros(href) if href.startswith("/") else href}" target="_self">{corpo}</a>'
    return f'<div class="du-mini">{corpo}</div>'


def du_modal_recusados(esperados: pd.DataFrame, sigem: pd.DataFrame) -> str:
    """O painel de recusados vira modal do cartao 'Recusados'.

    Ele existia solto na coluna direita e nao cabe numa tela sem rolagem. Em
    vez de perder a informacao, ela passa a abrir por :target -- mesmo custo
    zero de servidor das fichas.
    """
    return ('<div class="fmodal" id="du-recusados">'
            '<a class="fmodal-bg" href="#" aria-label="Fechar"></a>'
            '<div class="fmodal-box"><div class="fmodal-head">'
            '<div class="fmodal-title">Recusados há mais tempo</div>'
            '<a class="fmodal-x" href="#" aria-label="Fechar">&times;</a></div>'
            f'<div class="fmodal-body">{painel_recusados(esperados, sigem)}</div></div></div>')


def render_dashboard(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame,
                     sigem: pd.DataFrame, cache_key: str = ""):
    """Dashboard inteiro numa tela so, sem rolagem.

    Sai como UM bloco de HTML de proposito. Com st.columns cada coluna empilha
    por conta propria e as alturas nunca fecham; num grid unico o cabecalho,
    os KPIs e o rodape pegam so o que precisam e a faixa do meio fica com o
    resto -- e as listas de dentro se distribuem no que sobrou em vez de ter
    altura fixa. E por isso que a tela comprime junto com a janela em vez de
    estourar numa barra de rolagem.
    """
    agora = st.session_state.get("gplan_atualizado_em", "—")
    selo_filtro = st.session_state.get("_flt_selo", "")

    if resumo.empty:
        render_html('<div class="du-tela"><div class="du-vazio">'
                    "<div><strong>Nenhuma tag atende aos filtros.</strong></div>"
                    "<div>Ajuste ou limpe os filtros na barra lateral.</div></div></div>")
        return

    total_tags = len(resumo)
    total_esperados = int(resumo["RELATORIOS_ESPERADOS"].sum())
    total_postados = int(resumo["RELATORIOS_POSTADOS"].sum())
    total_aprovados = int(resumo["RELATORIOS_APROVADOS"].sum())
    # pendente inclui o que foi postado mas nao passou: recusado, em analise e
    # cancelado voltam para a fila em vez de contarem como entregue
    total_pendentes = total_esperados - total_aprovados
    avanco = (total_aprovados / total_esperados) if total_esperados else 0.0
    completas = int((resumo["AVANCO_DOCUMENTAL"] >= 1.0).sum())
    st_norm = esperados["STATUS_SIGEM"].astype(str).str.strip().str.upper()

    kpis = (
        du_kpi("Total de tags", br_num(total_tags), "", 1.0,
               "#5b8def", "shield", "/pesquisa")
        + du_kpi("Tags completas", br_num(completas),
                 f"{br_pct(completas / total_tags * 100)} do total",
                 completas / total_tags, "#34d399", "check")
        + du_kpi("Pendentes", br_num(total_pendentes),
                 f"{br_pct(total_pendentes / total_esperados * 100) if total_esperados else '—'} dos esperados",
                 (total_pendentes / total_esperados) if total_esperados else 0, "#f87171", "clock")
        + du_kpi("Emitidos SIGEM", br_num(total_postados),
                 f"{br_pct(total_postados / total_esperados * 100) if total_esperados else '—'} dos esperados",
                 (total_postados / total_esperados) if total_esperados else 0, "#fbbf24", "archive")
        + du_kpi("Avanço geral", br_pct(avanco * 100), f"{br_num(total_aprovados)} aprovados",
                 avanco, "#9d6bff", "trend")
    )

    # A medicao de campo entra pelo GITEC. Planilha antiga nao tem essas
    # colunas, entao tudo aqui degrada para zero em vez de estourar.
    # Estar no GITEC nao e estar medido: o evento vai para a fiscalizacao e so
    # vira medicao quando ela aprova. VALOR_GITEC ja traz so o aprovado.
    tem_gitec = "VALOR_GITEC" in resumo.columns
    medido_apr = float(resumo["VALOR_GITEC"].fillna(0).sum()) if tem_gitec else 0.0
    em_verif = (float(resumo["VALOR_GITEC_VERIF"].fillna(0).sum())
                if "VALOR_GITEC_VERIF" in resumo.columns else 0.0)
    montagem = (tags["STATUS_MONTAGEM"].astype(str).str.strip().str.upper()
                if "STATUS_MONTAGEM" in tags.columns else pd.Series(dtype=str))
    montados = int(montagem.eq("MONTADO").sum())

    # Previsto de medicao: o que ja fechou a documentacao e o GITEC ainda nao
    # mediu. Tag no SIGEM e no GITEC ja foi medido; tag so no GITEC tambem.
    # Sobra o que esta 100% documental e fora do GITEC -- e o que vai ser
    # medido. Dar R$ 0,00 com tag montada nao e falha de conta: quer dizer que
    # nenhuma das montadas fechou a documentacao.
    preco_tag = pd.to_numeric(tags.set_index("TAG")["PRECO_UNITARIO"],
                              errors="coerce").fillna(0.0)
    prontas = resumo[resumo["AVANCO_DOCUMENTAL"] >= 1.0]
    if tem_gitec:
        prontas = prontas[prontas["MEDIDO_GITEC"].astype(str).str.upper() != "SIM"]
    previsto = float(preco_tag.reindex(prontas["TAG"]).fillna(0).sum())

    minis = (
        du_mini("Tags montadas", br_num(montados),
                f"{br_pct(montados / total_tags * 100)} do total" if total_tags else "",
                "#5b8def", "shield")
        + du_mini("Previsto de medição", br_moeda(previsto), "", "#2dd4bf", "check")
        + du_mini("Valor total", br_moeda(float(preco_tag.sum())), "", "#9d6bff", "trend")
        # medido de verdade e o aprovado; o que esta em verificacao ainda pode
        # voltar, entao fica embaixo em vez de somar no numero grande
        + du_mini("Medido no GITEC", br_moeda(medido_apr),
                  f"{br_moeda(em_verif)} aguardando aprovação" if em_verif else "",
                  "#fbbf24", "archive", "/gitec")
        + du_mini("Última atualização", agora, "Sincronizado", "#5b8def", "calendario")
    )

    top = resumo.sort_values(["RELATORIOS_PENDENTES", "RELATORIOS_ESPERADOS"],
                             ascending=False).head(10)
    render_html(
        '<div class="du-tela">'
        '<header class="du-cab"><div><div class="du-h1">Dashboard</div>'
        "<p>Visão geral do andamento de relatórios e integração SIGEM</p></div>"
        f'<div class="du-acoes">{selo_filtro}</div></header>'
        f'<section class="du-kpis">{kpis}</section>'
        '<section class="du-meio">'
        f'<div class="du-col">{du_grupos(resumo)}{du_barras(esperados)}</div>'
        f'<div class="du-col">{du_status(esperados)}{du_top10(resumo)}</div>'
        "</section>"
        f'<section class="du-pe">{minis}</section>'
        "</div>"
        + du_modal_recusados(esperados, sigem)
        + fichas_completas(top["TAG"].tolist(), resumo, esperados, tags, sigem, cache_key)
    )


def render_relatorios(esperados: pd.DataFrame, resumo: pd.DataFrame, tags: pd.DataFrame,
                      sigem: pd.DataFrame, cache_key: str = ""):
    # Com filtro ativo o cabecalho diz quanto sobrou: sem isso a aba mostra um
    # numero menor e nada na tela explica por que.
    render_header("Relatórios previstos",
                  extra_pill=f"<strong>{br_num(len(esperados))}</strong> relatórios"
                  if st.session_state.get("_flt_selo") else None)

    # Mantem ORIGEM_REGRA no dataframe (sem exibir): e ela que separa
    # "RIR instrumentos" de "RIR cabos" no filtro por tipo de relatorio.
    df = esperados.copy()
    df["REVISAO_SIGEM"] = df["REVISAO_SIGEM"].fillna("—")
    df["DATA_SIGEM"] = format_date_column(df["DATA_SIGEM"])

    consume_url_filters(esperados)

    # ?busca= chega da ficha da TAG na aba Progresso. Aplicado uma vez, senao
    # todo rerun sobrescreveria o que o usuario digitou depois.
    pedido = st.query_params.get("busca")
    if pedido and st.session_state.get("_busca_token") != pedido:
        st.session_state["_busca_token"] = pedido
        st.session_state["mem_rel_busca"] = pedido
        st.session_state.pop("rel_busca", None)

    search = lembrado(st.text_input, "rel_busca", "Pesquisar", placeholder="Pesquisar por tag, descrição, relatório, documento, status...", label_visibility="collapsed")

    status_options = sorted({sentence_case(s) for s in esperados["STATUS_SIGEM"].dropna().unique()})
    col_rel, col_sts = st.columns(2)
    with col_rel:
        sel_rel = lembrado(st.multiselect, "flt_rel", "Tipo de relatório", REPORT_LABELS,
                           placeholder="Todos os relatórios")
    with col_sts:
        sel_sts = lembrado(st.multiselect, "flt_status", "Status SIGEM", status_options,
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

    st_norm = df["STATUS_SIGEM"].astype(str).str.strip().str.upper()
    render_html(faixa_resumo([
        ("No recorte", br_num(len(df)), None),
        ("Aprovados", br_num(int(st_norm.isin(STATUS_APROVADOS).sum())), "bom"),
        ("Recusados", br_num(int((st_norm == "RECUSADO").sum())), "ruim"),
        ("Em análise", br_num(int((st_norm == "EM ANÁLISE").sum())), None),
        ("Não postados", br_num(int((st_norm == "NAO POSTADO").sum())), None),
        ("Cancelados", br_num(int((st_norm == "CANCELADO").sum())), None),
    ]))

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
              <td class="gtbl-num">{botao_detalhes(r['DOCUMENTO_ESPERADO'], POSTADO(r['STATUS_SIGEM']))}</td>
            </tr>
        """
    render_html(
        '<div class="gplan-panel">'
        + html_table(
            ["Tag", "Descrição", "Grupo", "Relatório", "Referência", "Documento esperado",
             "#Existe no SIGEM", "Status SIGEM", "#Revisão", "#Data", "#Detalhes"],
            rows,
            "Nenhum relatório encontrado para essa busca.",
        )
        + "</div>"
        + fichas_completas(df_page["TAG"].tolist(), resumo, esperados, tags,
                           sigem, cache_key)
    )


# Nome da coluna de comentario da fiscalizacao na 04_BASE_SIGEM. Ainda nao
# existe na base; quando existir, e so acrescentar o nome aqui que a ficha
# passa a mostrar. Enquanto nao houver, a coluna simplesmente nao aparece.
COLUNAS_COMENTARIO = ("COMENTARIO", "COMENTÁRIO", "COMENTARIO_RECUSA",
                      "MOTIVO_RECUSA", "OBSERVACAO", "OBSERVAÇÃO")


def coluna_comentario(sigem: pd.DataFrame) -> str | None:
    for c in COLUNAS_COMENTARIO:
        if c in sigem.columns:
            return c
    return None


def POSTADO(status: object) -> bool:
    """Tem historico no SIGEM, logo tem ficha para abrir."""
    return str(status).strip().upper() not in ("NAO POSTADO", "NÃO POSTADO", "", "NAN")


# Os desenhos dos icones, so o miolo do <path>. Um dicionario e nao SVGs
# prontos porque o mesmo icone aparece em tamanho e cor diferentes.
FX_ICO = {
    "folha": '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
    "caixa": '<path d="M4 4h16v16H4z"/><path d="M4 9h16"/>',
    "onda": '<path d="M2 12h3l3-8 4 16 3-8h7"/>',
    "chip": '<rect x="6" y="6" width="12" height="12" rx="2"/><path d="M9 2v4M15 2v4M9 18v4M15 18v4M2 9h4M2 15h4M18 9h4M18 15h4"/>',
    "cabo": '<path d="M7 3v6a5 5 0 0 0 10 0V3"/><path d="M12 14v7"/>',
    "gota": '<path d="M12 3s6 6.4 6 10a6 6 0 0 1-12 0c0-3.6 6-10 6-10z"/>',
    "ok": '<path d="M20 6L9 17l-5-5"/>',
    "relogio": '<circle cx="12" cy="12" r="9"/><path d="M12 7v6l4 2"/>',
    "nuvem": '<path d="M17 18a4 4 0 0 0 .5-8 6 6 0 0 0-11.6 1.5A3.5 3.5 0 0 0 6.5 18z"/>',
    "regua": '<path d="M3 8h18v8H3z"/><path d="M7 8v3M11 8v3M15 8v3M19 8v3"/>',
    "tag": '<path d="M3 12l9-9 9 9-9 9z"/><circle cx="12" cy="12" r="2"/>',
    "livro": '<path d="M4 4h7v16H4z"/><path d="M13 4h7v16h-7z"/>',
    "grade": '<path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z"/>',
    "seta": '<path d="M3 17l6-6 4 4 8-8"/><path d="M17 7h4v4"/>',
    "link": '<path d="M10 13a5 5 0 0 0 7 0l3-3a5 5 0 0 0-7-7l-1 1"/><path d="M14 11a5 5 0 0 0-7 0l-3 3a5 5 0 0 0 7 7l1-1"/>',
    "alerta": '<path d="M12 3l9 16H3z"/><path d="M12 9v4M12 17h.01"/>',
    "pessoas": '<circle cx="9" cy="8" r="3"/><path d="M2 20a7 7 0 0 1 14 0"/><path d="M17 8a3 3 0 0 1 0 6"/>',
    "moeda": '<circle cx="12" cy="12" r="9"/><path d="M12 7v10M9.5 9.5h4a1.8 1.8 0 0 1 0 3.6h-3a1.8 1.8 0 0 0 0 3.6h4"/>',
}


# As cores dos blocos viram classe. O mesmo style inline repetido em dezenas de
# milhares de tiles custava mais que a folha de estilo inteira.
FX_COR = {"#5b8def": "azul", "#2dd4bf": "teal", "#9d6bff": "roxo", "#fbbf24": "ambar",
          "#f87171": "rubi", "#7c8aa8": "mudo", "#34d399": "verde", "#3a4a68": "cinza"}


def fx_css_icones() -> str:
    # Os icones como mascara CSS, um por classe. O desenho fica aqui, uma vez,
    # e no HTML sobra um <span> vazio. Alem de cortar uns 10 MB na aba
    # Progresso, e o que faz o icone sobreviver ao st.html -- que descarta
    # <svg> e e justamente o que aquela aba usa para conseguir abrir.
    regras = []
    for nome, desenho in FX_ICO.items():
        svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" '
               'stroke="black" stroke-width="1.9" stroke-linecap="round" '
               f'stroke-linejoin="round">{desenho}</svg>')
        # aspas simples dentro do data: URI, para nao ter de escapar as duplas
        uri = quote(svg.replace('"', "'"), safe="/:='<>() ")
        regras.append(f'.fxi-{nome}{{--fxi:url("data:image/svg+xml,{uri}");}}')
    return "\n        ".join(regras)


def fx_svg(nome: str, classe: str = "") -> str:
    # Nome antigo de proposito: os pontos de uso continuam pedindo "um icone".
    extra = f" {classe}" if classe else ""
    return f'<span class="fxi fxi-{nome}{extra}"></span>'


def fx_classe_cor(cor: str) -> str:
    return "fxc-" + FX_COR.get(cor, "azul")


def fx_trilha(itens: list) -> str:
    """A cadeia de onde a coisa pendura, do topo ate ela.

    Cada item e (rotulo, href). Fase e SOP viram link que filtra o app inteiro
    naquele nivel; SSOP e malha ficam como texto porque nao ha filtro global
    para eles -- link que nao leva a lugar nenhum e pior que texto.
    """
    itens = [(r, h) for r, h in itens if r and not str(r).rstrip().endswith("—")]
    partes = []
    for i, (rotulo, href) in enumerate(itens):
        if i:
            partes.append('<span class="sep">&rsaquo;</span>')
        if href:
            partes.append(f'<a href="{href}" target="_self">{esc(rotulo)}</a>')
        elif i == len(itens) - 1:
            partes.append(f'<span class="aqui">{esc(rotulo)}</span>')
        else:
            partes.append(f"<span>{esc(rotulo)}</span>")
    return f'<div class="fx-trilha">{"".join(partes)}</div>'


def fx_tile(rotulo: str, valor: object, icone: str, cor: str, sub: str = "",
            href: str = "") -> str:
    corpo = (f'<span class="ic {fx_classe_cor(cor)}">{fx_svg(icone)}</span>'
             f'<span class="cp"><span class="rot">{esc(rotulo)}</span>'
             f'<span class="val">{esc(valor)}</span>'
             + (f'<span class="sub">{esc(sub)}</span>' if sub else "") + "</span>")
    if href:
        return f'<a class="fx-tile" href="{href}" target="_self">{corpo}</a>'
    return f'<div class="fx-tile">{corpo}</div>'


def fx_kpi(rotulo: str, valor: object, sub: str, pct: float, cor: str, icone: str) -> str:
    largura = max(0.0, min(pct, 100.0))
    return (f'<div class="fx-kpi"><div class="top"><span class="rot">{esc(rotulo)}</span>'
            f'<span class="ic {fx_classe_cor(cor)}">{fx_svg(icone)}</span></div>'
            f'<div class="val">{esc(valor)}</div><div class="sub">{esc(sub)}</div>'
            f'<div class="fx-trilho {fx_classe_cor(cor)}"><i style="width:{largura:.1f}%;"></i></div></div>')


def fx_dado(rotulo: str, valor: object) -> str:
    return (f'<div class="fx-dado"><div class="rot">{esc(rotulo)}</div>'
            f'<div class="val" title="{esc(valor)}">{esc(valor)}</div></div>')


def fx_painel(titulo: str, icone: str, corpo: str, conta: str = "",
              classe_corpo: str = "") -> str:
    extra = f'<span class="conta">{esc(conta)}</span>' if conta else ""
    return (f'<div class="fx-pn"><div class="fx-pn-t"><span class="ic">{fx_svg(icone)}</span>'
            f"{esc(titulo)}{extra}</div>"
            f'<div class="fx-pn-c {classe_corpo}">{corpo}</div></div>')


def fx_rosca(feito: int, total: int, cor: str = "#2dd4bf",
             rotulo: str = "aprovado") -> str:
    # Donut em conic-gradient com um furo de mascara no meio: mesmo desenho do
    # SVG, sem uma tag que o st.html va descartar.
    pct = (feito / total * 100) if total else 0.0
    return (f'<div class="fx-rosca {fx_classe_cor(cor)}" style="--p:{pct:.1f};">'
            '<div class="anel"></div>'
            f'<div class="centro"><b>{br_pct(pct)}</b><span>{esc(rotulo)}</span>'
            "</div></div>")


def fx_lg(rotulo: str, valor: object, pct: str, cor: str, total: bool = False) -> str:
    return (f'<div class="fx-lg{" total" if total else ""}">'
            f'<i class="{fx_classe_cor(cor)}"></i><span class="nm">{esc(rotulo)}</span>'
            f'<b>{esc(valor)}</b><em>{esc(pct)}</em></div>')


def fx_linha(rotulo: str, valor: str) -> str:
    return f'<div class="fx-linha"><span>{esc(rotulo)}</span><b>{valor}</b></div>'


def fx_acao(rotulo: str, icone: str, href: str, nova_aba: bool = False) -> str:
    alvo = 'target="_blank" rel="noopener"' if nova_aba else 'target="_self"'
    return (f'<a class="fx-acao" href="{href}" {alvo}>'
            f'<span class="ic">{fx_svg(icone)}</span><span>{esc(rotulo)}</span></a>')


def doc_ancora(doc: object) -> str:
    """Id do modal de um relatorio."""
    return "rel-" + "".join(c if c.isalnum() else "-" for c in str(doc))


def botao_detalhes(doc: object, tem_ficha: bool = True) -> str:
    """Botao que abre a ficha do relatorio, em coluna propria.

    O endereco do documento fica como texto: ele e longo e cheio de ponto e
    underline, e sublinhado de link no meio disso atrapalha a leitura em vez
    de convidar ao clique. O botao diz o que faz.

    Sem ficha nao ha botao: os 15.203 relatorios nunca postados nao tem
    historico para mostrar, e um botao que abre o nada e pior que nada. O
    status deles ja aparece na propria linha.
    """
    if not tem_ficha:
        return '<span class="gtbl-muted">—</span>'
    return (f'<a class="btn-detalhes" href="#{doc_ancora(doc)}" '
            f'title="Ver revisões e comentários">Detalhes</a>')


@st.cache_data(show_spinner=False, max_entries=3)
def _revisoes_por_doc(cache_key: str, _sigem: pd.DataFrame) -> dict:
    """Historico de revisoes de cada documento, ordenado da mais nova.

    Na base do SIGEM o endereco do documento vem separado da revisao, entao
    cada postagem e uma linha -- da para reconstruir o historico inteiro.
    """
    col_com = coluna_comentario(_sigem)
    s = _sigem.copy()
    if col_com is None:
        s["_com"] = None
    else:
        s["_com"] = s[col_com]
    # ATALHO e o link direto do SIGEM para aquela revisao; base antiga nao tem
    s["_url"] = s["ATALHO"] if "ATALHO" in s.columns else None
    s["_dt"] = pd.to_datetime(s["DATA"], dayfirst=True, errors="coerce")
    # DATA_PARECER e o "Modificado em" da Sheet2: para um recusado, o dia em
    # que a fiscalizacao disse por que recusou. E dali que o relogio de parado
    # deve correr -- da postagem ele conta tambem o tempo que o documento
    # passou na fila esperando ser analisado, que nao e espera por providencia.
    s["_parecer"] = (pd.to_datetime(s["DATA_PARECER"], dayfirst=True, errors="coerce")
                     if "DATA_PARECER" in s.columns else pd.NaT)
    s = s.sort_values("_dt", ascending=False)
    hist: dict[str, list] = {}
    for doc, rev, status, data, com, url, parecer in zip(
            s["DOCUMENTO"].values, s["REVISAO"].values, s["STATUS"].values,
            s["_dt"].values, s["_com"].values, s["_url"].values, s["_parecer"].values):
        hist.setdefault(str(doc), []).append(
            (rev, status, pd.Timestamp(data) if pd.notna(data) else None, com, url,
             pd.Timestamp(parecer) if pd.notna(parecer) else None))
    return hist


def data_de_espera(data_postagem, data_parecer):
    """A data a partir da qual o documento esta esperando providencia.

    Vale o parecer quando existe. O "Modificado em" da Sheet2 e sempre igual ou
    posterior a postagem, e para um recusado e o dia em que a fiscalizacao
    disse por que recusou -- e dali que a bola esta com quem emitiu. Contar da
    postagem somaria tambem os dias que o documento passou na fila esperando
    analise, que nao sao espera por providencia: nos 5.111 recusados de hoje
    isso inflava a conta em 34 dias, em media.
    """
    return data_parecer if data_parecer is not None else data_postagem


def espera_por_documento(historico: dict) -> dict:
    """Data em que cada documento passou a esperar providencia.

    Sai da revisao mais nova de cada um. E a mesma conta em toda tela: duas
    paginas mostrando dias diferentes para o mesmo documento e pior que
    nenhuma das duas mostrar.
    """
    return {doc: data_de_espera(h[0][2], h[0][5])
            for doc, h in historico.items() if h and len(h[0]) > 5}


def ficha_relatorio_html(doc: str, linhas_esperadas: pd.DataFrame, historico: list) -> str:
    """Ficha de um relatorio: o que se espera dele e como ele andou no SIGEM.

    O miolo e o historico: cada revisao com o parecer da inspecao logo abaixo,
    em vermelho quando foi recusa. E dai que sai o "parado ha N dias" -- o
    numero que diz se o documento esta andando ou encalhado.
    """
    primeira = linhas_esperadas.iloc[0]
    tags_doc = sorted({str(t) for t in linhas_esperadas["TAG"]})
    atual = historico[0] if historico else None
    status_atual = str(atual[1]).strip() if atual else "Não postado"
    aprovada = status_atual.upper() in STATUS_APROVADOS
    tom = "ok" if aprovada else ("crit" if status_atual.upper() in
                                 {"RECUSADO", "CANCELADO"} else "andamento")
    hoje = pd.Timestamp.now(tz=BR_TZ).tz_localize(None).normalize()
    n_apr = sum(1 for h in historico if str(h[1]).strip().upper() in STATUS_APROVADOS)

    # -------------------------------------------------------------- trilha
    trilha = []
    if len(tags_doc) == 1:
        trilha.append((f"Tag {tags_doc[0]}", f"#{ficha_anchor(tags_doc[0])}"))
    else:
        trilha.append((f"{br_num(len(tags_doc))} instrumentos", ""))
    trilha.append((f'{primeira["RELATORIO"]} · {primeira["REFERENCIA"]}', ""))

    # --------------------------------------------------------------- tiles
    tiles = (
        fx_tile("Relatório", str(primeira["RELATORIO"]), "folha", "#5b8def",
                str(primeira["GRUPO"]).title())
        + fx_tile("Referência", str(primeira["REFERENCIA"]), "tag", "#9d6bff")
        + fx_tile("Revisão atual", str(atual[0]) if atual else "—", "chip", "#fbbf24",
                  f"{br_num(len(historico))} emitida{'s' if len(historico) != 1 else ''}"
                  if historico else "nunca postado")
        + fx_tile("Situação", sentence_case(status_atual), "alerta",
                  "#34d399" if aprovada else "#f87171")
    )
    espera = data_de_espera(atual[2], atual[5]) if atual and len(atual) > 5 else (
        atual[2] if atual else None)
    if espera is not None and not aprovada:
        dias = (hoje - espera.normalize()).days
        desde = ("desde o parecer de " if atual[5] is not None else "desde ") + f"{espera:%d/%m/%Y}"
        tiles += fx_tile("Parado há", f"{br_num(dias)} dia{'s' if dias != 1 else ''}",
                         "relogio", "#f87171", desde)
    elif espera is not None:
        tiles += fx_tile("Liberado em", f"{espera:%d/%m/%Y}", "ok", "#34d399")

    # ------------------------------------------------------------ revisoes
    if historico:
        linhas = []
        # da mais antiga para a mais nova: a ficha conta uma historia, e
        # historia se le do comeco
        for rev, status, data, com, url, parecer in reversed(historico):
            txt = str(status).strip()
            t = ("ok" if txt.upper() in STATUS_APROVADOS else
                 "crit" if txt.upper() in {"RECUSADO", "CANCELADO"} else "andamento")
            eh_atual = (rev, status, data, com, url, parecer) == historico[0]
            marca = '<span class="fx-atual">atual</span>' if eh_atual else ""
            cel_data = (f'<td class="gtbl-num">{data:%d/%m/%Y}</td>' if data is not None
                        else '<td class="gtbl-num gtbl-muted">—</td>')
            # o relogio corre do parecer, nao da postagem: e de la que a bola
            # esta com quem emitiu o documento
            marco = data_de_espera(data, parecer)
            cel_dias = ('<td class="gtbl-num gtbl-muted">—</td>' if marco is None else
                        f'<td class="gtbl-num">{br_num((hoje - marco.normalize()).days)} dias</td>')
            cel_url = (f'<td class="gtbl-num"><a class="rel-sigem" href="{esc(url)}" '
                       f'target="_blank" rel="noopener">SIGEM</a></td>'
                       if url is not None and not vazio(url)
                       else '<td class="gtbl-num gtbl-muted">—</td>')
            linhas.append(
                f'<tr{" class=fx-rev-atual" if eh_atual else ""}>'
                f'<td class="gtbl-strong">{esc(rev)}{marca}</td>'
                f'<td><span class="gtbl-badge {t}">{esc(sentence_case(txt))}</span></td>'
                f"{cel_data}{cel_dias}{cel_url}</tr>")
            if com is not None and not vazio(com):
                classe = "rec" if t == "crit" else "obs"
                linhas.append(
                    f'<tr><td class="fx-com-cel" colspan="5">'
                    f'<div class="fx-com {classe}"><div class="cab">'
                    f'<span class="ic">{fx_svg("alerta")}</span>'
                    f"Parecer da inspeção · revisão {esc(rev)}"
                    + (f" · {parecer:%d/%m/%Y}" if parecer is not None else "")
                    + "</div>"
                    f"<p>{esc(com).strip()}</p></div></td></tr>")
        corpo = html_table(["Revisão", "Status", "#Data", "#Há", "#Link"], "".join(linhas))
        conta = (f"{br_num(len(historico))} emissões · "
                 + (f"{br_num(n_apr)} liberada{'s' if n_apr != 1 else ''}" if n_apr
                    else "nenhuma liberada"))
    else:
        corpo = ('<div class="gtbl-empty">Ainda não postado no SIGEM. '
                 "Nenhuma revisão registrada.</div>")
        conta = "sem histórico"

    # Fechar volta para a ficha da TAG de onde se veio, e nao para a arvore:
    # quem abre um relatorio quase sempre quer conferir os outros da mesma TAG
    # em seguida. So da para saber a origem quando o documento pertence a uma
    # TAG so -- 161 deles sao compartilhados, ate por 700 TAGs, e nesses o
    # fechar volta para a arvore mesmo.
    volta = f"#{ficha_anchor(tags_doc[0])}" if len(tags_doc) == 1 else "#fechado"
    rotulo_volta = "Voltar para a TAG" if len(tags_doc) == 1 else "Fechar"
    return (
        f'<div class="fmodal" id="{doc_ancora(doc)}">'
        f'<a class="fmodal-bg" href="{volta}" aria-label="{rotulo_volta}"></a>'
        '<div class="fmodal-box"><div class="fmodal-head"><div>'
        f'<div class="fn-tipo">Relatório · {esc(primeira["RELATORIO"])}</div>'
        f'<div class="fmodal-title rel-titulo">{esc(doc)}</div></div>'
        f'<span class="gtbl-badge {tom} rel-sit">{esc(sentence_case(status_atual))}</span>'
        f'<a class="fmodal-x" href="{volta}" aria-label="{rotulo_volta}" '
        f'title="{rotulo_volta}">&times;</a></div>'
        f'<div class="fmodal-body"><div class="fx">{fx_trilha(trilha)}'
        f'<div class="fx-tiles">{tiles}</div>'
        + fx_painel("Todas as revisões", "relogio", corpo, conta=conta, classe_corpo="zero")
        + "</div></div></div></div>"
    )


def fichas_relatorios_html(docs, esperados: pd.DataFrame, historico: dict) -> str:
    """Fichas dos relatorios visiveis na pagina."""
    unicos = list(dict.fromkeys(str(d) for d in docs))
    por_doc = {k: v for k, v in
               esperados[esperados["DOCUMENTO_ESPERADO"].isin(unicos)].groupby("DOCUMENTO_ESPERADO")}
    return "".join(
        ficha_relatorio_html(d, por_doc[d], historico.get(d, []))
        for d in unicos if d in por_doc
    )


def tag_ficha_html(tag_id: str, resumo: pd.DataFrame, esperados: pd.DataFrame,
                   tags: pd.DataFrame, com_cabecalho: bool = True,
                   espera_por_doc: dict | None = None,
                   niveis_na_pagina: bool = False) -> str | None:
    """Ficha completa da tag, como um bloco HTML unico.

    Serve tanto a aba Pesquisa tag quanto o modal aberto pelo Dashboard, pelos
    Relatorios e pela arvore. Tudo aqui se encadeia: a trilha diz de onde a tag
    pendura e filtra o app naquele nivel, a tabela leva a ficha de cada
    relatorio, e a movimentacao aponta o documento que esta segurando o avanco.
    """
    resumo_row = resumo[resumo["TAG"] == tag_id]
    if resumo_row.empty:
        return None

    r = resumo_row.iloc[0]
    tags_row = tags[tags["TAG"] == tag_id]
    t = tags_row.iloc[0] if not tags_row.empty else {}

    def da_base(campo, default="—"):
        v = t.get(campo, default) if hasattr(t, "get") else default
        return default if v is None or vazio(v) else v

    preco = (pd.to_numeric(pd.Series([t.get("PRECO_UNITARIO")]), errors="coerce")
             .fillna(0).iloc[0]) if hasattr(t, "get") else 0
    qtd_cabos = int(r["QTD_CABOS"])
    qtd_tubing = int(r["QTD_TUBING"]) if not vazio(r.get("QTD_TUBING")) else 0

    esp = int(r["RELATORIOS_ESPERADOS"])
    apr = int(r["RELATORIOS_APROVADOS"])
    pos = int(r["RELATORIOS_POSTADOS"])
    pen = int(r["RELATORIOS_PENDENTES"])
    meus = esperados[esperados["TAG"] == tag_id]
    st_norm = meus["STATUS_SIGEM"].astype(str).str.strip().str.upper()
    recusados = int(st_norm.eq("RECUSADO").sum())

    # ------------------------------------------------------------- trilha
    # O caminho de volta: cada degrau abre a ficha daquele nivel. Na aba
    # Progresso as fichas de nivel estao na propria pagina, entao basta a
    # ancora e nada recarrega. Nas outras abas elas nao existem -- ali o link
    # leva para a Progresso ja apontando para a ficha, em vez de virar uma
    # ancora sem destino, que e clique que nao faz nada.
    def degrau(tipo: str, valor: str) -> str:
        alvo = f"#{_ancora(tipo, valor)}"
        return alvo if niveis_na_pagina else com_filtros("/progresso") + alvo

    trilha = []
    for rotulo, campo, tipo in (("Fase", "FASE", "FASE"), ("SOP", "SOP", "SOP"),
                                ("SSOP", "SSOP", "SSOP"), ("Malha", "MALHA", "MALHA")):
        v = da_base(campo)
        if v != "—":
            trilha.append((f"{rotulo} {v}", degrau(tipo, str(v))))
    trilha.append((str(tag_id), ""))

    # -------------------------------------------------------------- tiles
    tiles = (
        fx_tile("Tipo", str(r["GRUPO_REGRA"]).title(), "caixa", "#5b8def")
        + fx_tile("Comunicação", da_base("COMUNICACAO"), "onda", "#2dd4bf")
        + fx_tile("Item PPU", r["ITEM_PPU"], "chip", "#9d6bff")
        + fx_tile("Cabo", "Sim" if qtd_cabos else "Não", "cabo", "#5b8def",
                  f"{qtd_cabos} cabo{'s' if qtd_cabos != 1 else ''}" if qtd_cabos else "sem cabo")
        + fx_tile("Tubing", "Sim" if qtd_tubing else "Não", "gota", "#7c8aa8",
                  f"{qtd_tubing} linhas" if qtd_tubing else "sem tubing")
        + fx_tile("Status documental", sentence_case(r["STATUS_DOCUMENTAL"]), "alerta",
                  "#34d399" if str(r["STATUS_DOCUMENTAL"]).upper().startswith("LIB") else "#f87171")
    )

    # --------------------------------------------------------------- KPIs
    # TAG cancelada tem zero relatorio esperado -- e a ficha divide tudo por
    # ele. Sem esta guarda a primeira cancelada derruba a aba Progresso, que
    # monta a ficha de todas as 5.098.
    def pct(n: int) -> float:
        return (n / esp * 100) if esp else 0.0

    kpis = (
        fx_kpi("Aprovados", br_num(apr), f"{br_pct(pct(apr))} dos esperados" if esp else "",
               pct(apr), "#2dd4bf", "ok")
        + fx_kpi("Postados no SIGEM", br_num(pos), "", pct(pos), "#fbbf24", "nuvem")
        + fx_kpi("Pendentes", br_num(pen), f"{br_pct(pct(pen))} dos esperados" if esp else "",
                 pct(pen), "#f87171", "relogio")
        + fx_kpi("Recusados", br_num(recusados), "", pct(recusados), "#9d6bff", "alerta")
    )
    dados = (
        fx_dado("Critério de medição", da_base("CRITERIO_MEDICAO"))
        + fx_dado("Preço unitário", br_moeda(float(preco)))
        + fx_dado("SOP", da_base("SOP"))
        + fx_dado("SSOP", da_base("SSOP"))
        + fx_dado("Segmento", da_base("SEGMENTO"))
        + fx_dado("Malha", da_base("MALHA"))
    )

    # ---------------------------------------------------- tabela dos relatorios
    linhas = []
    for rel, ref, doc, stat, rev in zip(
            meus["RELATORIO"].values, meus["REFERENCIA"].values,
            meus["DOCUMENTO_ESPERADO"].values, meus["STATUS_SIGEM"].values,
            meus["REVISAO_SIGEM"].values):
        linhas.append(
            f'<tr><td class="gtbl-strong">{fx_svg("folha", "fx-folha")}{esc(rel)}</td>'
            f"<td>{esc(ref)}</td><td>{esc(doc)}</td><td>{status_badge(stat)}</td>"
            f"<td>{esc(format_missing(rev))}</td>"
            f'<td class="gtbl-num">{botao_detalhes(doc, POSTADO(stat))}</td></tr>'
        )
    tabela = html_table(
        ["Relatório", "Referência", "Documento esperado", "Status SIGEM", "#Revisão",
         "#Detalhes"], "".join(linhas), classe="gtbl gtbl-rel")

    # ------------------------------------------------- coluna da direita
    legenda = (fx_lg("Aprovados", br_num(apr), br_pct(pct(apr)), "#2dd4bf")
               + fx_lg("Pendentes", br_num(pen), br_pct(pct(pen)), "#f87171")
               + fx_lg("Esperados", br_num(esp), "", "#3a4a68", total=True))
    nota = ""
    if esp:
        avanco = fx_painel("Avanço documental", "seta",
                           fx_rosca(apr, esp) + f'<div class="fx-leg">{legenda}</div>' + nota,
                           classe_corpo="centro")
    else:
        # sem regra documental nao ha avanco: dizer por que, em vez de um
        # donut em 0% que parece atraso
        avanco = fx_painel(
            "Avanço documental", "seta",
            '<p class="fx-nota">Tag cancelada: não gera relatório nem pendência. '
            "Continua na base para rastreio.</p>"
            if str(r["STATUS_DOCUMENTAL"]).strip().upper() == "CANCELADA"
            else '<p class="fx-nota">Nenhum relatório previsto para esta tag.</p>',
            classe_corpo="centro")

    campo = [(rot, da_base(col)) for rot, col in
             (("Localização", "STATUS_LOCALIZACAO"), ("Calibração", "STATUS_CALIBRACAO"),
              ("Montagem", "STATUS_MONTAGEM"), ("Status final", "STATUS_FINAL"))]
    campo_html = "".join(fx_linha(rot, status_pill(v)) for rot, v in campo
                         if v != "—")
    bloco_campo = (fx_painel("Situação em campo", "pessoas", campo_html)
                   if campo_html else "")
    # Card proprio, logo depois da situacao em campo -- que e onde o leitor
    # acabou de ver "On demand" e vai perguntar quando chega. So existe para
    # quem esta em compra: numa tag calibrada nao quer dizer nada.
    bloco_fornecimento = painel_previsao_fornecimento(
        da_base("STATUS_FINAL"),
        t.get("PREVISAO_FORNECIMENTO") if hasattr(t, "get") else None)

    # o que esta segurando o avanco: o documento parado ha mais tempo leva
    # direto para a ficha dele, com o parecer da inspecao
    postados = meus[meus["STATUS_SIGEM"].map(POSTADO)]
    mov = ""
    if not postados.empty:
        dts = pd.to_datetime(postados["DATA_SIGEM"], dayfirst=True, errors="coerce")
        if dts.notna().any():
            hoje = pd.Timestamp.now(tz=BR_TZ).tz_localize(None).normalize()
            mov += fx_linha("Última no SIGEM", f"{dts.max():%d/%m/%Y}")
            travados = postados[~postados["STATUS_SIGEM"].map(
                lambda x: str(x).strip().upper() in STATUS_APROVADOS)]
            if not travados.empty:
                # a mesma conta da ficha do relatorio: do parecer, quando ha.
                # Duas telas mostrando dias diferentes para o mesmo documento e
                # pior que nao mostrar nenhuma.
                espera = espera_por_doc or {}
                postagem = pd.to_datetime(travados["DATA_SIGEM"], dayfirst=True,
                                          errors="coerce")
                d2 = pd.Series(
                    [espera.get(str(doc), dt)
                     for doc, dt in zip(travados["DOCUMENTO_ESPERADO"], postagem)],
                    index=travados.index)
                if d2.notna().any():
                    i = d2.idxmin()
                    dias = (hoje - pd.Timestamp(d2.loc[i]).normalize()).days
                    mov += fx_linha(
                        "Parado há mais tempo",
                        f'<a class="gtbl-link" href="#{doc_ancora(travados.loc[i, "DOCUMENTO_ESPERADO"])}">'
                        f'{esc(travados.loc[i, "RELATORIO"])} · {br_num(dias)} dia'
                        f'{"s" if dias != 1 else ""}</a>')
    if mov:
        mov = fx_painel("Movimentação", "relogio", mov)

    acoes = fx_painel("Ações", "link", '<div class="fx-acoes">'
                      + fx_acao("Ver em Relatórios", "folha",
                                com_filtros("/relatorios?busca=" + quote(str(tag_id))))
                      + fx_acao("Ver na árvore", "grade",
                                com_filtros("/progresso?fase=" + quote(str(da_base("FASE"))))
                                if da_base("FASE") != "—" else com_filtros("/progresso"))
                      + "</div>")

    cabecalho = (
        f'<div class="fx-cab"><span class="marca">{fx_svg("tag")}</span>'
        f'<div><h2>{esc(r["TAG"])}</h2><p>{esc(r["DESCRICAO"])}</p></div></div>'
        if com_cabecalho else ""
    )
    return (
        f'<div class="fx">{cabecalho}{fx_trilha(trilha)}'
        f'<div class="fx-tiles">{tiles}</div>'
        '<div class="fx-corpo"><div class="fx-col">'
        + fx_painel("Resumo documental", "grade",
                    f'<div class="fx-kpis">{kpis}</div><div class="fx-dados">{dados}</div>')
        + fx_painel("Relatórios da tag", "folha", tabela,
                    conta=f"{br_num(esp)} previstos", classe_corpo="zero")
        + f'</div><div class="fx-col">{avanco}{bloco_campo}{bloco_fornecimento}{mov}{acoes}</div></div></div>'
    )


def render_pesquisa_tag(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame,
                        sigem: pd.DataFrame, cache_key: str = ""):
    render_header("Pesquisa tag",
                  extra_pill=f"<strong>{br_num(len(resumo))}</strong> tags"
                  if st.session_state.get("_flt_selo") else None)

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

        meus = esperados[esperados["TAG"] == tag_id]
        postados = meus[meus["STATUS_SIGEM"].map(POSTADO)]["DOCUMENTO_ESPERADO"].tolist()
        render_html(f'<div class="gplan-panel" id="{ficha_anchor(tag_id)}">'
                    + tag_ficha_html(tag_id, resumo, esperados, tags,
                                     espera_por_doc=espera_por_documento(
                                         _revisoes_por_doc(cache_key, sigem)))
                    + "</div>"
                    + fichas_relatorios_html(postados, esperados,
                                             _revisoes_por_doc(cache_key, sigem)))
        return

    search = lembrado(st.text_input, "pesq_busca", "Pesquisar", placeholder="Digite a tag para ver a ficha completa (ex: AIT-120005)...", label_visibility="collapsed")

    list_df = resumo[["TAG", "DESCRICAO", "GRUPO_REGRA", "ITEM_PPU", "RELATORIOS_ESPERADOS", "AVANCO_DOCUMENTAL"]].copy()
    # o status de campo mora na 01_BASE_TAGS, nao no resumo. Planilha antiga
    # nao tem essas colunas: sem o merge a pill so mostra tracinho.
    campo = [c for c in ("STATUS_LOCALIZACAO", "STATUS_CALIBRACAO", "STATUS_MONTAGEM",
                         "STATUS_FINAL") if c in tags.columns]
    if campo:
        list_df = list_df.merge(tags[["TAG"] + campo], on="TAG", how="left")
    if search:
        mask = (list_df["TAG"].astype(str).str.contains(search, case=False, na=False, regex=False)
                | list_df["DESCRICAO"].astype(str).str.contains(search, case=False, na=False, regex=False))
        list_df = list_df[mask]

    list_df["AVANCO_DOCUMENTAL"] = (list_df["AVANCO_DOCUMENTAL"] * 100).round(1)
    list_df_page = paginate(list_df, "pesquisa", search)

    rows = ""
    for _, r in list_df_page.iterrows():
        avanco = r["AVANCO_DOCUMENTAL"]
        tone = "ok" if avanco >= 70 else ("warn" if avanco >= 30 else "crit")
        href = com_filtros(f"?tag={quote(str(r['TAG']))}")
        rows += f"""
            <tr>
              <td><a class="gtbl-tag gtbl-link" href="{href}" target="_self">{esc(r['TAG'])}</a></td>
              <td>{esc(r['DESCRICAO'])}</td>
              <td class="gtbl-muted">{esc(str(r['GRUPO_REGRA']).title())}</td>
              <td class="gtbl-num gtbl-muted">{esc(r['ITEM_PPU'])}</td>
              <td class="gtbl-num">{int(r['RELATORIOS_ESPERADOS'])}</td>
              <td class="gtbl-num"><span class="gtbl-badge {tone}">{br_pct(avanco)}</span></td>
              <td class="gtbl-num">{status_pill(r.get('STATUS_LOCALIZACAO'))}</td>
              <td class="gtbl-num">{status_pill(r.get('STATUS_CALIBRACAO'))}</td>
              <td class="gtbl-num">{status_pill(r.get('STATUS_MONTAGEM'))}</td>
              <td class="gtbl-num">{status_pill(r.get('STATUS_FINAL'))}</td>
            </tr>
        """
    render_html(
        '<div class="gplan-panel">'
        + html_table(
            ["Tag", "Descrição", "Tipo", "#PPU", "#Relatórios", "#Avanço",
             "#Localização", "#Calibração", "#Montagem", "#Status final"],
            rows,
            "Nenhuma tag encontrada para essa busca.",
        )
        + "</div>"
    )


def render_sigem(sigem: pd.DataFrame, esperados: pd.DataFrame | None = None,
                 filtrado: bool = False):
    """A base crua do SIGEM, recortada pelo filtro da lateral quando ha um.

    Com filtro ativo sobram os documentos das tags escolhidas. Isso deixa de
    fora o que nao casa com tag nenhuma -- justamente o que se costuma vir
    procurar aqui --, entao a contagem diz quantos ficaram de fora, em vez de
    o numero simplesmente encolher sem explicacao.
    """
    total = len(sigem)
    if filtrado and esperados is not None:
        sigem = sigem[sigem["DOCUMENTO"].isin(set(esperados["DOCUMENTO_ESPERADO"]))]
    pill = f"<strong>{br_num(len(sigem))}</strong> documentos"
    if filtrado and len(sigem) != total:
        pill += f" · {br_num(total - len(sigem))} fora do filtro"
    render_header("Base SIGEM", extra_pill=pill)

    status_options = ["Todos"] + sorted(sigem["STATUS"].dropna().unique().tolist())
    col1, col2 = st.columns([1, 3])
    with col1:
        status_filter = lembrado(st.selectbox, "sig_status", "Status", status_options)
    with col2:
        text_search = lembrado(st.text_input, "sig_busca", "Pesquisa de texto", placeholder="Buscar em qualquer campo do documento...")

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


def servico_do_agrupamento(valor: object) -> str:
    """Só o serviço do agrupamento do GITEC.

    Vem como "B_4.3.8.1.2_Instalação de analisadores incluindo materiais...".
    O código já tem coluna própria ao lado, e repetir aqui empurrava a tabela
    5px além do painel -- o bastante para cortar a última coluna.
    """
    texto = str(valor or "")
    partes = texto.split("_", 2)
    return partes[2] if len(partes) > 2 else texto


def render_gitec(gitec: pd.DataFrame, resumo: pd.DataFrame, tags: pd.DataFrame,
                 esperados: pd.DataFrame, sigem: pd.DataFrame, cache_key: str = ""):
    """A medição de campo: o que o GITEC já mediu dos instrumentos da base.

    É o outro lado da conta do controle documental. A documentação diz o que
    pode ser medido; esta aba diz o que foi. O cruzamento das duas é o que
    mostra instrumento medido antes de a documentação fechar, e instrumento
    pronto que ainda não virou medição.
    """
    render_header("Gitec")

    if gitec.empty:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "Nenhuma medição do GITEC nesta planilha. Rode o pipeline com a "
                    "06_BASE_GITEC.xlsx na pasta das bases.</div></div>")
        return

    g = gitec.copy()
    g["VALOR"] = pd.to_numeric(g["VALOR"], errors="coerce").fillna(0.0)
    g["_dt"] = pd.to_datetime(g["DATA_EXECUCAO"], errors="coerce", dayfirst=True)
    # medido e so o aprovado; o resto esta com a fiscalizacao e pode nao virar
    # medicao nenhuma
    aprovado_m = g["STATUS"].astype(str).str.strip().str.upper().str.startswith("APROVADO")
    medido_apr = float(g.loc[aprovado_m, "VALOR"].sum())
    em_verif = float(g.loc[~aprovado_m, "VALOR"].sum())
    tags_medidas = set(g.loc[aprovado_m, "TAG"])
    tags_fila = set(g.loc[~aprovado_m, "TAG"]) - tags_medidas

    preco_tag = pd.to_numeric(tags.set_index("TAG")["PRECO_UNITARIO"],
                              errors="coerce").fillna(0.0)
    medidas = tags_medidas
    # pronta e ainda nao aprovada continua prevista, mesmo ja estando na fila
    prontas = resumo[(resumo["AVANCO_DOCUMENTAL"] >= 1.0) & (~resumo["TAG"].isin(medidas))]
    previsto = float(preco_tag.reindex(prontas["TAG"]).fillna(0).sum())
    total_obra = float(preco_tag.sum())
    montagem = (tags["STATUS_MONTAGEM"].astype(str).str.strip().str.upper()
                if "STATUS_MONTAGEM" in tags.columns else pd.Series(dtype=str))
    montadas = set(tags.loc[montagem.eq("MONTADO"), "TAG"]) if len(montagem) else set()

    tiles = (
        fx_tile("Instrumentos medidos", br_num(len(tags_medidas)), "tag", "#2dd4bf",
                f"{br_pct(len(tags_medidas) / len(tags) * 100)} da base")
        + fx_tile("Aguardando aprovação", br_num(len(tags_fila)), "relogio", "#fbbf24")
        + fx_tile("Montadas sem medição", br_num(len(montadas - medidas)), "alerta", "#f87171")
        + fx_tile("Última medição", f"{g['_dt'].max():%d/%m/%Y}" if g["_dt"].notna().any()
                  else "—", "relogio", "#5b8def")
        + fx_tile("Medição do contrato", br_pct(medido_apr / total_obra * 100)
                  if total_obra else "—", "seta", "#34d399")
    )

    kpis = (
        fx_kpi("Medido e aprovado", br_moeda(medido_apr),
               f"{br_pct(medido_apr / total_obra * 100)} do valor da obra" if total_obra else "",
               (medido_apr / total_obra * 100) if total_obra else 0, "#34d399", "ok")
        + fx_kpi("Aguardando aprovação", br_moeda(em_verif), "ainda não é medição",
                 (em_verif / total_obra * 100) if total_obra else 0, "#fbbf24", "relogio")
        + fx_kpi("Previsto de medição", br_moeda(previsto),
                 f"{br_num(len(prontas))} prontas e fora do GITEC",
                 (previsto / total_obra * 100) if total_obra else 0, "#2dd4bf", "nuvem")
        + fx_kpi("Valor da obra", br_moeda(total_obra), "", 100, "#9d6bff", "moeda")
    )

    # por item de PPU: é a chave que liga os dois contratos. Medido e
    # aguardando andam em colunas separadas porque um item é medido muitas
    # vezes, em tags diferentes e em datas diferentes.
    g["_apr"] = aprovado_m
    por_item = g.groupby("ITEM_PPU_GITEC").agg(
        tags_=("TAG", "nunique"),
        medidos=("_apr", "sum"),
        eventos=("TAG", "size"),
        valor=("VALOR", "sum"),
        valor_apr=("VALOR", lambda v: float(v[g.loc[v.index, "_apr"]].sum())),
        ultima=("_dt", "max"),
    ).reset_index()
    por_item["aguardando"] = por_item["eventos"] - por_item["medidos"]
    nome_item = (tags.dropna(subset=["ITEM_PPU"])
                 .groupby(tags["ITEM_PPU"].astype(str).str.strip())["TIPO_ORIGEM"]
                 .agg(lambda x: sentence_case(x.mode().iloc[0]) if len(x.mode()) else "—")
                 .to_dict())
    linhas_item = ""
    for _, r in por_item.sort_values("valor_apr", ascending=False).iterrows():
        aguard = int(r["aguardando"])
        medidos = int(r["medidos"])
        tom = "ok" if not aguard else ("warn" if medidos else "crit")
        rotulo = (f"{br_num(medidos)} medido{'s' if medidos != 1 else ''}"
                  + (f" · {br_num(aguard)} aguardando" if aguard else ""))
        data = f"{r['ultima']:%d/%m/%Y}" if pd.notna(r["ultima"]) else "—"
        linhas_item += (
            f'<tr><td class="gtbl-strong">{esc(r["ITEM_PPU_GITEC"])}</td>'
            f'<td class="gtbl-muted">{esc(nome_item.get(str(r["ITEM_PPU_GITEC"]).strip(), "—"))}</td>'
            f'<td class="gtbl-num">{br_num(int(r["tags_"]))}</td>'
            f'<td><span class="gtbl-badge {tom}">{rotulo}</span></td>'
            f'<td class="gtbl-num">{data}</td>'
            f'<td class="gtbl-num gtbl-strong">{br_moeda(r["valor_apr"])}</td></tr>')
    painel_item = fx_painel(
        "Medição por item de PPU", "chip",
        html_table(["Item", "Tipo", "#Instrumentos", "Status", "#Última medição",
                    "#Valor medido"], linhas_item),
        conta=f"{len(por_item)} itens", classe_corpo="zero")

    # movimentação: quanto foi medido em cada mês
    mov = g.dropna(subset=["_dt"]).copy()
    painel_mov = ""
    if not mov.empty:
        mov["mes"] = mov["_dt"].dt.to_period("M")
        por_mes = mov.groupby("mes").agg(n=("TAG", "nunique"), valor=("VALOR", "sum")).reset_index()
        teto = float(por_mes["valor"].max()) or 1.0
        barras = "".join(
            f'<div class="du-br"><span class="nm">{r["mes"].strftime("%m/%Y")}</span>'
            f'<span class="tr"><i style="width:{r["valor"] / teto * 100:.1f}%;"></i></span>'
            f'<span class="fr">{br_num(int(r["n"]))} tags</span>'
            f'<span class="pc">{br_moeda(r["valor"])}</span></div>'
            for _, r in por_mes.sort_values("mes").iterrows())
        painel_mov = fx_painel("Movimentação por mês", "relogio",
                               f'<div class="du-barras">{barras}</div>',
                               conta=f"{len(por_mes)} meses")

    # a tabela das medições, com o cruzamento documental de cada tag
    avanco = dict(zip(resumo["TAG"], resumo["AVANCO_DOCUMENTAL"]))
    # Sem paginacao: 119 eventos cabem inteiros e a rolagem propria do painel
    # resolve. Paginar aqui custava uma ida ao servidor -- e os controles ainda
    # apareciam la em cima, longe da tabela que eles paginavam.
    g = g.sort_values("_dt", ascending=False)
    pagina = g
    linhas = ""
    for _, r in pagina.iterrows():
        pct = float(avanco.get(r["TAG"], 0)) * 100
        tom = "ok" if pct >= 100 else ("warn" if pct >= 50 else "crit")
        ap = str(r["STATUS"]).strip().upper().startswith("APROVADO")
        data = (f"{r['_dt']:%d/%m/%Y}" if pd.notna(r["_dt"]) else "—")
        linhas += (
            f'<tr><td>{tag_link(r["TAG"])}</td>'
            f'<td class="gtbl-muted">{esc(r["ITEM_PPU_GITEC"])}</td>'
            f'<td class="gtbl-muted gt-corta" title="{esc(r["AGRUPAMENTO"])}">'
            f'{esc(servico_do_agrupamento(r["AGRUPAMENTO"]))}</td>'
            f'<td><span class="gtbl-badge {"ok" if ap else "andamento"}">'
            f'{esc(sentence_case(r["STATUS"]))}</span></td>'
            f'<td class="gtbl-num">{data}</td>'
            f'<td class="gtbl-num gtbl-strong">{br_moeda(r["VALOR"])}</td>'
            f'<td class="gtbl-num"><span class="gtbl-badge {tom}">{br_pct(pct)}</span></td></tr>')
    tabela = fx_painel(
        "Medições", "folha",
        '<div class="fx-rolagem">'
        + html_table(["Tag", "Item", "Agrupamento", "Status", "#Data", "#Valor",
                      "#Avanço documental"], linhas)
        + "</div>",
        conta=f"{br_num(len(g))} eventos", classe_corpo="zero")

    render_html(
        '<div class="fx">'
        f'<div class="fx-tiles">{tiles}</div>'
        '<div class="fx-corpo"><div class="fx-col">'
        + fx_painel("Resumo da medição", "grade", f'<div class="fx-kpis">{kpis}</div>')
        + painel_item + tabela
        + '</div><div class="fx-col">'
        + fx_painel("Gitec x Status", "seta",
                    fx_rosca(int(medido_apr), int(medido_apr + em_verif) or 1)
                    + '<div class="fx-leg">'
                    + fx_lg("Aprovado", br_moeda(medido_apr), "", "#34d399")
                    + fx_lg("Aguardando aprovação", br_moeda(em_verif), "", "#fbbf24")
                    + "</div>", classe_corpo="centro")
        + painel_mov
        + fx_painel("Cruzamento com a documentação", "grade",
                    fx_linha("Aguardando aprovação", br_num(len(tags_fila)))
                    + fx_linha("Medidas e prontas",
                             br_num(len(medidas & set(resumo.loc[resumo.AVANCO_DOCUMENTAL >= 1.0, "TAG"]))))
                    + fx_linha("Medidas sem documentação pronta",
                               br_num(len(medidas) - len(medidas & set(
                                   resumo.loc[resumo.AVANCO_DOCUMENTAL >= 1.0, "TAG"]))))
                    + fx_linha("Prontas e ainda não medidas", br_num(len(prontas)))
                    + fx_linha("Montadas sem medição", br_num(len(montadas - medidas))))
        + "</div></div></div>"
        + fichas_completas(pagina["TAG"].tolist(), resumo, esperados, tags, sigem, cache_key)
    )


# ==========================================================================
#  Aba Planta: o avanco de montagem desenhado sobre o arranjo da unidade
# ==========================================================================

MAPA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "mapa")


def carregar_mapa() -> dict:
    """As pranchas e o contorno das zonas, como o gerar_mapa_assets.py deixou."""
    caminho = os.path.join(MAPA_DIR, "zonas.json")
    if not os.path.exists(caminho):
        return {"pranchas": []}
    # A chave e a data do arquivo: rodou o gerador de novo, o mapa recarrega
    # sozinho. Um numero de versao fixo aqui exigiria lembrar de subi-lo.
    return _mapa_do_disco(os.path.getmtime(caminho))


@st.cache_data(show_spinner=False)
def _mapa_do_disco(mtime: float) -> dict:
    """A imagem entra como data URI: o Streamlit nao serve arquivo estatico por
    caminho relativo dentro do HTML que a gente injeta."""
    caminho = os.path.join(MAPA_DIR, "zonas.json")
    with open(caminho, encoding="utf-8") as f:
        mapa = json.load(f)
    for p in mapa["pranchas"]:
        png = os.path.join(MAPA_DIR, p["arquivo"])
        if os.path.exists(png):
            with open(png, "rb") as img:
                p["uri"] = "data:image/png;base64," + base64.b64encode(img.read()).decode()
        else:
            p["uri"] = ""
    return mapa


def classe_avanco(pct: float) -> str:
    return "feito" if pct >= 99.5 else "andando" if pct > 0 else "parado"


def dados_por_area(tags: pd.DataFrame, resumo: pd.DataFrame,
                   locacao: pd.DataFrame,
                   aux: pd.DataFrame) -> tuple[dict, dict, pd.DataFrame]:
    """Quanto cada area ja montou, e que desenho pertence a que area.

    A TAG e localizada por area, nao por desenho -- a base nao tem coluna de
    desenho por instrumento. Por isso o numero e sempre da area: duas zonas da
    mesma area mostram o mesmo percentual, e isso e o dado, nao um arredondamento.

    Os desenhos marcados como gerais (CHZ-113 e CHZ-302, que atravessam a
    unidade) ficam de fora do mapa desenho->area: eles nao delimitam zona.
    """
    if locacao.empty or "AREA" not in locacao.columns:
        return {}, {}, pd.DataFrame()

    area_da_tag = (locacao.dropna(subset=["AREA"])
                   .assign(AREA=lambda d: d["AREA"].astype(str).str.strip())
                   .set_index("TAG")["AREA"].to_dict())

    nome, area_do_desenho = {}, {}
    if not aux.empty:
        for _, r in aux.iterrows():
            a = str(r["AREA"]).strip()
            nome.setdefault(a, str(r.get("NOME_AREA") or "").strip() or "—")
            if not bool(r.get("DESENHO_GERAL")):
                area_do_desenho[str(r["DESENHO"]).strip()] = a

    t = tags.copy()
    # TAG cancelada nao gera pendencia nem entra em avanco, aqui como no resto
    if "STATUS_FINAL" in t.columns:
        t = t[t["STATUS_FINAL"].astype(str).str.strip().str.upper() != "CANCELADO"]
    t["_area"] = t["TAG"].map(area_da_tag)
    t = t[t["_area"].notna()].copy()
    if t.empty:
        return {}, area_do_desenho, pd.DataFrame()

    t["_montado"] = (t.get("STATUS_MONTAGEM", pd.Series("", index=t.index))
                     .astype(str).str.strip().str.upper() == "MONTADO")
    t["_preco"] = pd.to_numeric(t.get("PRECO_UNITARIO"), errors="coerce").fillna(0.0)
    # O que a ficha da planta mostra por instrumento vem do resumo: avanco
    # documental, se o GITEC ja mediu e quanto.
    r = resumo.set_index("TAG")
    doc = pd.to_numeric(r.get("AVANCO_DOCUMENTAL"), errors="coerce").fillna(0.0)
    t["_docfrac"] = t["TAG"].map(doc).fillna(0.0)
    t["_medido"] = (t["TAG"].map(r.get("MEDIDO_GITEC", pd.Series(dtype=object)))
                    .astype(str).str.strip().str.upper())
    t["_valorgitec"] = t["TAG"].map(
        pd.to_numeric(r.get("VALOR_GITEC"), errors="coerce")).fillna(0.0)
    if "GRUPO_REGRA" in r.columns:
        t["GRUPO_REGRA"] = t["TAG"].map(r["GRUPO_REGRA"])

    areas = {}
    for a, sub in t.groupby("_area"):
        qtd = len(sub)
        mont = int(sub["_montado"].sum())
        # O que ja foi montado se divide em duas partes que somam o todo: o que
        # o GITEC ja aprovou e o que ainda nao virou medicao. Estar no GITEC nao
        # basta -- so o aprovado conta, que e o que MEDIDO_GITEC ja diz.
        montado = sub["_montado"]
        medido = montado & (sub["_medido"] == "SIM")
        areas[a] = {
            "area": a, "nome": nome.get(a, "—"), "tags": qtd, "montados": mont,
            "pct": round(mont / qtd * 100, 1) if qtd else 0.0,
            "doc": round(sub["_docfrac"].mean() * 100, 1),
            "valor": float(sub["_preco"].sum()),
            "valor_montado": float(sub.loc[montado, "_preco"].sum()),
            "valor_medido": float(sub.loc[medido, "_preco"].sum()),
            "montados_medidos": int(medido.sum()),
        }
    return areas, area_do_desenho, t


def zona_area(zona: dict, area_do_desenho: dict) -> str | None:
    """A area da zona, se os desenhos dela apontarem todos para a mesma."""
    alvo = {area_do_desenho.get(d) for d in zona.get("desenhos", [])} - {None}
    return alvo.pop() if len(alvo) == 1 else None


def codigos_da_zona(zona: dict) -> str:
    """CHZ-325, ou CHZ-316/317, ou JEI-001 a 007 quando a lista e longa."""
    c = [d.replace("800-", "") for d in zona.get("desenhos", [])]
    if not c:
        return "—"
    if len(c) > 3:
        return f"{c[0]} a {c[-1][-3:]}"
    return c[0] + "".join("/" + x[-3:] for x in c[1:])


# Largura util de cada prancha na tela, so para dimensionar o rotulo dentro da
# zona: a coluna do Streamlit tem 1.600 px menos o respiro lateral.
PLANTA_LARGURA = {"principal": 1500, "piperack": 1160, "se1200": 270}


def planta_zonas_html(prancha: dict, areas: dict, area_do_desenho: dict) -> str:
    largura = PLANTA_LARGURA.get(prancha["id"], 1500)
    altura = largura * prancha["prop"] / 100
    partes = []
    for z in prancha["zonas"]:
        a = areas.get(zona_area(z, area_do_desenho) or "")
        if not a:
            continue
        rotulo = codigos_da_zona(z)
        px, py = z["l"] / 100 * largura, z["a"] / 100 * altura
        fs = max(8.5, min(13.0, (px - 14) / (len(rotulo) * 0.58)))
        estilo = (f"left:{z['x']:.2f}%; top:{z['y']:.2f}%; width:{z['l']:.2f}%;"
                  f" height:{z['a']:.2f}%; --p:{a['pct']:.1f}; --fs:{fs:.1f}px")
        if z.get("recorte"):
            estilo += f"; clip-path:{z['recorte']}"
        # Numa zona pequena a etiqueta da area brigaria com o codigo pelo mesmo
        # espaco; ali ela sai, e a area continua no resumo e no title.
        etiqueta = (f'<span class="pl-ar">{esc(a["area"])}</span>'
                    if px > 58 and py > 46 else "")
        extra = (f'<u>{br_num(a["montados"])} de {br_num(a["tags"])}</u>'
                 if py > 78 and px > 104 else "")
        titulo = (f'{rotulo} · área {a["area"]} — {a["nome"]}\n'
                  f'{br_num(a["montados"])} de {br_num(a["tags"])} montados '
                  f'({br_pct(a["pct"])}) — clique para abrir a ficha')
        partes.append(
            f'<a class="pl-zona {classe_avanco(a["pct"])}'
            f'{" rec" if z.get("recorte") else ""}" style="{estilo}"'
            f' href="#{_ancora("PLANTA", rotulo)}" title="{esc(titulo)}">{etiqueta}'
            f'<span class="pl-mio"><b>{esc(rotulo)}</b>'
            f'<i>{br_pct(a["pct"])}</i>{extra}</span></a>')
    return "".join(partes)


def planta_prancha_html(prancha: dict, areas: dict, area_do_desenho: dict) -> str:
    vistas = {z for z in (zona_area(z, area_do_desenho) for z in prancha["zonas"])
              if z and z in areas}
    qtd = sum(areas[a]["tags"] for a in vistas)
    mont = sum(areas[a]["montados"] for a in vistas)
    pct = mont / qtd * 100 if qtd else 0.0
    return (
        '<div class="gplan-panel pl-pn">'
        f'<div class="gplan-panel-title">{esc(prancha["rotulo"])}'
        f'<span class="pl-res">{len(vistas)} área{"s" if len(vistas) > 1 else ""}'
        f' · {br_num(qtd)} instrumentos · '
        f'<b class="{classe_avanco(pct)}">{br_pct(pct)}</b></span></div>'
        f'<div class="pl-tela" style="padding-top:{prancha["prop"]:.3f}%">'
        f'<img src="{prancha["uri"]}" alt="Planta — {esc(prancha["rotulo"])}">'
        + planta_zonas_html(prancha, areas, area_do_desenho)
        + "</div></div>")


def render_planta(tags: pd.DataFrame, resumo: pd.DataFrame, locacao: pd.DataFrame,
                  aux: pd.DataFrame):
    """O avanco de montagem por area, desenhado sobre o arranjo da unidade.

    A aba e so o visual e o resumo: serve para bater o olho e ver onde a
    montagem anda e onde parou. O detalhe por TAG continua na Progresso.
    """
    render_header("Planta")

    mapa = carregar_mapa()
    areas, area_do_desenho, base = dados_por_area(tags, resumo, locacao, aux)

    if not mapa["pranchas"]:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "A marcação das plantas não está no repositório. Rode "
                    "<code>gerar_mapa_assets.py</code> apontando para o PPTX do mapa "
                    "de infraestrutura.</div></div>")
        return
    if not areas:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "Esta planilha ainda não tem a área de cada TAG. Rode o pipeline "
                    "com a 05_BASE_LOCAÇÃO.xlsx atualizada — ela traz a coluna ÁREA e "
                    "a aba AUX, que é o que liga o instrumento ao desenho."
                    "</div></div>")
        return

    total = sum(a["tags"] for a in areas.values())
    mont = sum(a["montados"] for a in areas.values())
    val_m = sum(a["valor_montado"] for a in areas.values())
    val_t = sum(a["valor"] for a in areas.values())
    faixas = [("feito", "Concluído", "100% montado", lambda a: a["pct"] >= 99.5),
              ("andando", "Em andamento", "1% a 99%", lambda a: 0 < a["pct"] < 99.5),
              ("parado", "Não iniciado", "nenhum montado", lambda a: a["pct"] == 0)]
    contagem = {c: [a for a in areas.values() if f(a)] for c, _t, _f, f in faixas}

    val_med = sum(a["valor_medido"] for a in areas.values())
    val_sem = val_m - val_med
    n_med = sum(a["montados_medidos"] for a in areas.values())

    pct_mont = mont / total * 100 if total else 0.0
    pct_valor = val_m / val_t * 100 if val_t else 0.0
    pct_sem = val_sem / val_m * 100 if val_m else 0.0
    pct_med = val_med / val_m * 100 if val_m else 0.0
    # O valor sai por extenso: o que se leva para reuniao de medicao e o numero
    # cheio, e "R$ 535k" nao serve para conferir com o GITEC. Os dois ultimos
    # cartoes sao a divisao do segundo -- somados, dao o valor montado.
    render_html(f"""
      <div class="pl-kpis">
        <div class="pl-kpi"><div class="r">Montagem nas áreas</div>
          <div class="v andando">{br_pct(pct_mont)}</div>
          <div class="s">{br_num(mont)} de {br_num(total)} instrumentos montados</div>
          <div class="pl-barra"><i class="andando" style="width:{pct_mont:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="r">Valor montado</div>
          <div class="v dinheiro">{br_moeda(val_m)}</div>
          <div class="s">{br_pct(pct_valor)} de {br_moeda(val_t)} nas áreas mapeadas</div>
          <div class="pl-barra"><i class="feito" style="width:{pct_valor:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="r">Montado sem medir</div>
          <div class="v dinheiro parado">{br_moeda(val_sem)}</div>
          <div class="s">{br_num(mont - n_med)} instrumentos ·
            {br_pct(pct_sem)} do montado</div>
          <div class="pl-barra"><i class="parado" style="width:{pct_sem:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="r">Montado e medido no GITEC</div>
          <div class="v dinheiro feito">{br_moeda(val_med)}</div>
          <div class="s">{br_num(n_med)} instrumentos ·
            {br_pct(pct_med)} do montado</div>
          <div class="pl-barra"><i class="feito" style="width:{pct_med:.1f}%"></i></div></div>
      </div>""")

    por_id = {p["id"]: p for p in mapa["pranchas"]}
    if "principal" in por_id:
        render_html_pesado(planta_prancha_html(por_id["principal"], areas, area_do_desenho))

    # As pranchas deitadas ficam embaixo da principal, na largura toda; a
    # subestacao e um desenho em pe -- numa faixa larga ela viraria uma torre
    # de mil pixels, entao vai para uma coluna estreita ao lado, com a legenda
    # ocupando a altura que sobra do lado largo.
    resto = [p for p in mapa["pranchas"] if p["id"] != "principal"]
    largas = [p for p in resto if p["prop"] < 100]
    altas = [p for p in resto if p["prop"] >= 100]
    if altas:
        col_a, col_b = st.columns([4, 1], gap="medium")
    else:
        col_a, col_b = st.container(), None
    with col_a:
        for p in largas:
            render_html_pesado(planta_prancha_html(p, areas, area_do_desenho))
        render_html(planta_legenda_html(faixas, contagem))
    if col_b is not None:
        with col_b:
            for p in altas:
                render_html_pesado(planta_prancha_html(p, areas, area_do_desenho))

    # As fichas ficam no fim da pagina, fechadas: o :target so abre a que a
    # zona clicada aponta. Sem todas geradas, o clique cairia no vazio.
    render_html_pesado(planta_fichas_html(mapa, areas, area_do_desenho, base))


def planta_ficha_html(desenho: str, area: dict, sub: pd.DataFrame,
                      irmaos: list[str]) -> str:
    """A ficha da planta: tudo que esta cadastrado nos instrumentos dela.

    A lista de TAGs rola dentro do painel em vez de paginar. Paginar aqui
    obrigaria a fechar a ficha para trocar de pagina -- o modal e :target puro,
    nao tem estado proprio -- e a area 140 tem 749 instrumentos.
    """
    n = len(sub)
    montados = int(sub["_montado"].sum())
    pct = montados / n * 100 if n else 0.0
    completos = int((sub["_docfrac"] >= 1.0).sum())
    medidos = int((sub["_medido"] == "SIM").sum())
    doc = float(sub["_docfrac"].mean() * 100) if n else 0.0
    valor = float(sub["_preco"].sum())
    valor_mont = float(sub.loc[sub["_montado"], "_preco"].sum())
    medido_valor = float(sub["_valorgitec"].sum())
    sem_medicao = int(((sub["_montado"]) & (sub["_medido"] != "SIM")).sum())
    tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")

    trilha = [(f'Área {area["area"]}', ""), (esc(area["nome"]), ""), (desenho, "")]

    tiles = (
        fx_tile("Instrumentos", br_num(n), "tag", "#2dd4bf",
                f"{br_num(completos)} com documentação completa")
        + fx_tile("Montados", br_num(montados), "ok", "#34d399", br_pct(pct))
        + fx_tile("Avanço documental", br_pct(doc), "livro", "#5b8def")
        + fx_tile("Valor total", br_moeda(valor), "moeda", "#9d6bff")
        + fx_tile("Medido no GITEC", br_moeda(medido_valor), "seta", "#fbbf24",
                  f"{br_num(medidos)} instrumentos")
    )

    kpis = (
        fx_kpi("Montados em campo", br_num(montados), f"{br_pct(pct)} da área",
               pct, "#34d399", "ok")
        + fx_kpi("Documentação completa", br_num(completos),
                 f"de {br_num(n)} instrumentos",
                 completos / n * 100 if n else 0, "#5b8def", "livro")
        + fx_kpi("Medidos pelo GITEC", br_num(medidos), br_moeda(medido_valor),
                 medidos / n * 100 if n else 0, "#fbbf24", "moeda")
        + fx_kpi("Montados sem medição", br_num(sem_medicao),
                 f"de {br_num(montados)} montados",
                 sem_medicao / montados * 100 if montados else 0, "#f87171", "relogio")
    )
    grupos = (sub["GRUPO_REGRA"].dropna().astype(str).str.title().value_counts()
              if "GRUPO_REGRA" in sub.columns else pd.Series(dtype=int))
    dados = (
        fx_dado("Área", f'{area["area"]} · {area["nome"]}')
        + fx_dado("Outras plantas da área",
                  ", ".join(d.replace("800-", "") for d in irmaos) or "—")
        + fx_dado("Tipos de instrumento",
                  ", ".join(f"{k} ({v})" for k, v in grupos.head(4).items()) or "—")
        + fx_dado("Valor montado", f"{br_moeda(valor_mont)} de {br_moeda(valor)}")
    )

    # ------------------------------------------------- a lista, com rolagem
    linhas = []
    for _, r in sub.sort_values(["_montado", "TAG"], ascending=[False, True]).iterrows():
        d = float(r["_docfrac"]) * 100
        linhas.append(
            "<tr>"
            f'<td>{tag_pill(r["TAG"])}</td>'
            f'<td class="desc">{esc(r.get("DESCRICAO", ""))}</td>'
            f'<td>{"<span class=\'pl-sim\'>Montado</span>" if r["_montado"] else "<span class=\'pl-nao\'>—</span>"}</td>'
            f'<td class="num">{br_pct(d)}</td>'
            f'<td>{"<span class=\'pl-sim\'>Medido</span>" if r["_medido"] == "SIM" else "<span class=\'pl-nao\'>—</span>"}</td>'
            f'<td class="num">{br_moeda(float(r["_preco"]))}</td>'
            "</tr>")
    tabela = html_table(
        ["TAG", "Descrição", "Montagem", "#Doc.", "Medição", "#Valor"],
        "".join(linhas), "Nenhum instrumento nesta planta.",
        classe="gtbl gtbl-tags")

    direita = fx_painel(
        "Montagem da planta", "seta",
        fx_rosca(montados, n, "#34d399", "montado")
        + '<div class="fx-leg">'
        + fx_lg("Montados", br_num(montados), br_pct(pct), "#34d399")
        + fx_lg("A montar", br_num(n - montados),
                br_pct(100 - pct if n else 0), "#f87171")
        + fx_lg("Instrumentos", br_num(n), "", "#3a4a68", total=True)
        + "</div>",
        classe_corpo="centro")

    return (
        f'<div class="fmodal" id="{_ancora("PLANTA", desenho)}">'
        '<a class="fmodal-bg" href="#" aria-label="Fechar"></a>'
        '<div class="fmodal-box"><div class="fmodal-head">'
        '<div><div class="fn-tipo">Planta</div>'
        f'<div class="fmodal-title">{esc(desenho)}</div></div>'
        '<div class="fn-avanco">'
        f'<div class="fn-track"><div class="fn-fill {tom}" style="width:{max(pct, 1.5):.1f}%;"></div></div>'
        f'<div class="fn-pct">{br_pct(pct)}</div></div>'
        '<a class="fmodal-x" href="#" aria-label="Fechar">&times;</a></div>'
        f'<div class="fmodal-body"><div class="fx">{fx_trilha(trilha)}'
        f'<div class="fx-tiles">{tiles}</div>'
        '<div class="fx-corpo"><div class="fx-col">'
        + fx_painel("Resumo da planta", "grade",
                    f'<div class="fx-kpis">{kpis}</div><div class="fx-dados">{dados}</div>')
        + fx_painel("Instrumentos", "tag", f'<div class="pl-rol">{tabela}</div>',
                    conta=f"{br_num(n)} tags · role para ver todas",
                    classe_corpo="zero")
        + f'</div><div class="fx-col">{direita}</div></div></div></div></div></div>'
    )


def planta_fichas_html(mapa: dict, areas: dict, area_do_desenho: dict,
                       base: pd.DataFrame) -> str:
    """Uma ficha por zona marcada, todas fechadas.

    Por zona, e nao por desenho: a zona e o que se clica, e a area 140 aparece
    em duas delas com sete codigos JEI numa -- gerar uma ficha por codigo
    repetiria oito vezes a mesma lista de 749 instrumentos.

    Sem gerar todas de uma vez, o clique cairia numa ancora inexistente e
    simplesmente nao faria nada.
    """
    zonas = [z for p in mapa["pranchas"] for z in p["zonas"]]
    por_area: dict[str, list[str]] = {}
    for z in zonas:
        a = zona_area(z, area_do_desenho)
        if a:
            por_area.setdefault(a, []).extend(z["desenhos"])

    partes = []
    for z in zonas:
        a = zona_area(z, area_do_desenho)
        if not a or a not in areas:
            continue
        meus = set(z["desenhos"])
        irmaos = [x.replace("800-", "") for x in dict.fromkeys(por_area.get(a, []))
                  if x not in meus]
        partes.append(planta_ficha_html(codigos_da_zona(z), areas[a],
                                        base[base["_area"] == a], irmaos))
    return "".join(partes)


def planta_legenda_html(faixas: list, contagem: dict) -> str:
    """A chave de leitura do mapa. Sem repetir area por area: isso e a Progresso."""
    itens = "".join(
        f'<div class="pl-ch {c}"><span class="sw"></span>'
        f'<div class="tx"><b>{t}</b><em>{f}</em></div>'
        f'<div class="qt"><b>{len(contagem[c])}</b>'
        f'<em>{br_num(sum(a["tags"] for a in contagem[c]))} inst.</em></div></div>'
        for c, t, f, _fn in faixas)
    return ('<div class="gplan-panel pl-lg"><div class="gplan-panel-title">Legenda</div>'
            f'<div class="pl-ch-c">{itens}</div></div>')


def vazio(v: object) -> bool:
    """A base marca ausencia com '-', nao com celula vazia."""
    return str(v).strip().upper() in SEM_VALOR


def br_moeda(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "~").replace(".", ",").replace("~", ".")


def progresso_base(resumo: pd.DataFrame, tags: pd.DataFrame) -> pd.DataFrame:
    """Une o avanco documental (07_TAG_RESUMO) com a hierarquia e o preco
    (01_BASE_TAGS). O avanco por TAG e o mesmo usado nas demais abas."""
    cols = ["TAG", "FASE", "SOP", "SSOP", "SUBGRUPO_PRIORIDADE", "SEGMENTO", "MALHA",
            "CRITERIO_MEDICAO", "PRECO_UNITARIO",
            "STATUS_LOCALIZACAO", "STATUS_CALIBRACAO", "STATUS_MONTAGEM", "STATUS_FINAL"]
    disponiveis = [c for c in cols if c in tags.columns]
    df = resumo.merge(tags[disponiveis], on="TAG", how="left")
    for c in ("FASE", "SOP", "SSOP", "SEGMENTO", "MALHA", "SUBGRUPO_PRIORIDADE"):
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
        emitidos=("RELATORIOS_APROVADOS", "sum"),
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
                <span class="gr-pct">{br_pct(pct)}</span>
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


# Cadeia de expansao. Segmento ficou de fora: 92 dos 142 segmentos atravessam
# varios SSOP, entao ele nao e um sub-nivel -- vira apenas coluna na ficha.
CADEIA = [("SOP", "SOP", "SSOP"), ("SSOP", "SSOP", "malhas"), ("MALHA", "Malha", "tags")]


def assinatura_filtros() -> str:
    """Os filtros ativos como texto, para entrar na chave de cache.

    A arvore, as fichas de nivel e as das TAGs sao cacheadas, e a chave levava
    so a planilha, o segmento, a malha e a pagina. Com um SOP filtrado na
    lateral, os graficos e os totais mudavam -- eles sao calculados na hora --
    mas a arvore vinha do cache, inteira, do jeito que estava antes do filtro.
    A tela contava duas historias ao mesmo tempo.
    """
    return "|".join(f"{c}={st.session_state.get(f'gf_{c}', p)}"
                    for c, _rot, p, _f, _col in FILTROS)


def _ancora(tipo: str, valor: str) -> str:
    """Id do modal de um nivel. So alfanumerico, para ser id CSS valido."""
    limpo = "".join(c if c.isalnum() else "-" for c in str(valor))
    return f"n-{tipo}-{limpo}"


# max_entries baixo de proposito: cada entrada guarda ate ~5 MB de HTML e o
# plano do Render tem 512 MB. Com 8 por funcao dava 83 MB so de cache.
@st.cache_data(show_spinner=False, max_entries=3)
def arvore_html(cache_key: str, filtro: str, f_seg: str, f_malha: str, pag: int,
                _df: pd.DataFrame) -> str:
    """A arvore inteira, ja em HTML e sempre fechada.

    Sao 224 agregacoes e 1.883 tabelas de TAG: 2,7 s aqui, uns 20 s na CPU do
    Render. Sem cache isso se repetia a cada clique. A chave e o par
    (planilha, filtros, pagina) -- o _df nao entra no hash, so acompanha. A
    pagina precisa estar na chave: sem ela a pagina 2 servia a arvore da 1.
    """
    def recorte(base, coluna, valor):
        """Linhas de um nivel. Vazio agrupa junto: 1.210 TAGs nao tem fase."""
        return base[base[coluna] == valor] if not vazio(valor) \
            else base[base[coluna].apply(vazio)]

    blocos = []
    for _, fase in agrega_nivel(_df, "FASE", subnivel="SOP").iterrows():
        d_fase = recorte(_df, "FASE", fase["FASE"])
        sops = []
        for _, sop in agrega_nivel(d_fase, "SOP", subnivel="SSOP").iterrows():
            d_sop = recorte(d_fase, "SOP", sop["SOP"])
            ssops = []
            for _, ss in agrega_nivel(d_sop, "SSOP", subnivel="MALHA").iterrows():
                d_ssop = recorte(d_sop, "SSOP", ss["SSOP"])
                malhas = []
                for _, ml in agrega_nivel(d_ssop, "MALHA").iterrows():
                    d_malha = recorte(d_ssop, "MALHA", ml["MALHA"])
                    # a malha sempre tem "+": as TAGs dela ja vem no HTML
                    corpo = (f'<div class="arv-tags">'
                             f'{_tabela_tags(d_malha, com_modal=True)}</div>')
                    malhas.append(_no("MALHA", ml["MALHA"], ml, nivel=4, filhos=corpo))
                ssops.append(_no("SSOP", ss["SSOP"], ss, nivel=3, filhos="".join(malhas)))
            sops.append(_no("SOP", sop["SOP"], sop, nivel=2, filhos="".join(ssops)))
        blocos.append(_no("FASE", fase["FASE"], fase, nivel=1, filhos="".join(sops)))
    return "".join(blocos)


def render_progresso(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame,
                     sigem: pd.DataFrame, cache_key: str = ""):
    render_header("Progresso")
    df = progresso_base(resumo, tags)

    segs = ["Todos"] + sorted({s for s in df.SEGMENTO if not vazio(s)})
    malhas = ["Todas"] + sorted({m for m in df.MALHA if not vazio(m)})
    c1, c2 = st.columns(2)
    with c1:
        f_seg = lembrado(st.selectbox, "prg_seg", "Segmento", segs)
    with c2:
        f_malha = lembrado(st.selectbox, "prg_malha", "Malha", malhas)
    if f_seg != "Todos":
        df = df[df.SEGMENTO == f_seg]
    if f_malha != "Todas":
        df = df[df.MALHA == f_malha]

    render_html(_totais(df))
    _graficos(df)

    # Filtros, totais e graficos ja estao na tela: a cobertura sai aqui e a
    # espera pela arvore passa a ser mostrada no lugar dela, la embaixo.
    descobrir()

    # Os 66 SOPs numa pagina so, com todas as fichas junto: e o que permite
    # abrir SOP, SSOP, malha e TAG sobre a pagina, sem navegar.
    #
    # Isso exige RAM que o plano gratuito nao tem. Medido: 880 MB para um
    # usuario, 1.275 MB para dois, contra os 512 MB da instancia -- o processo
    # e morto no meio da resposta e o navegador recebe 502. Ligado assim por
    # decisao do Daniel, que vai contratar o plano; ate la a aba fica fora do ar
    # e as outras quatro seguem funcionando.
    pag, df_pag = 0, df

    # A barra fica no lugar onde a arvore vai nascer, e cada etapa e uma etapa
    # de verdade -- nada de porcentagem inventada. No fim a propria arvore
    # substitui a barra, entao a pagina nao pisca nem salta.
    lugar = st.empty()
    n = br_num(len(df_pag))

    lugar.markdown(tela_carregando(f"Montando a árvore de {n} instrumentos", 20,
                                   coberta=False), unsafe_allow_html=True)
    filtro = assinatura_filtros()
    arvore = arvore_html(cache_key, filtro, f_seg, f_malha, pag, df_pag)

    lugar.markdown(tela_carregando("Preparando as fichas de SOP, SSOP e malha", 55,
                                   coberta=False), unsafe_allow_html=True)
    niveis = fichas_niveis_html(cache_key, filtro, f_seg, f_malha, pag, df_pag)

    lugar.markdown(tela_carregando(f"Preparando as fichas das {n} TAGs", 80,
                                   coberta=False), unsafe_allow_html=True)
    fichas = fichas_tags_html(cache_key, filtro, f_seg, f_malha, pag,
                              df_pag["TAG"].tolist(), resumo, esperados, tags, sigem)

    lugar.markdown(tela_carregando("Preparando as fichas dos relatórios", 92,
                                   coberta=False), unsafe_allow_html=True)
    fichas += fichas_relatorios_pagina(cache_key, filtro, f_seg, f_malha, pag,
                                       df_pag["TAG"].tolist(), esperados, sigem)

    lugar.html(
        '<div class="gplan-panel">'
        '<div class="gplan-panel-title">FASE · SOP · SSOP · MALHA · TAG</div>'
        f'<div class="arvore">{arvore}</div></div>' + niveis + fichas
    )


# max_entries baixo de proposito: cada entrada guarda muitos MB de HTML e o
# plano do Render tem 512 MB.
@st.cache_data(show_spinner=False, max_entries=3)
def fichas_niveis_html(cache_key: str, filtro: str, f_seg: str, f_malha: str, pag: int,
                       _df: pd.DataFrame) -> str:
    """As fichas de SOP, SSOP e malha, todas fechadas."""
    partes = []
    for tipo, rotulo in (("FASE", "Fase"), ("SOP", "SOP"), ("SSOP", "SSOP"),
                         ("MALHA", "Malha")):
        for valor, sub in _df.groupby(tipo):
            partes.append(_modal_nivel(tipo, valor, sub, rotulo))
    return "".join(partes)


@st.cache_data(show_spinner=False, max_entries=3)
def fichas_relatorios_pagina(cache_key: str, filtro: str, f_seg: str, f_malha: str, pag: int,
                             _tags_ids: list, _esperados: pd.DataFrame,
                             _sigem: pd.DataFrame) -> str:
    """Fichas dos relatorios das TAGs da pagina, so os ja postados.

    Os 15.203 nunca postados custariam 20 MB para nao dizer nada alem do
    "Nao postado" que a linha ja mostra. Os 4.885 com historico custam 6 MB e
    sao os unicos com revisao, data e motivo de recusa para exibir.
    """
    meus = _esperados[_esperados["TAG"].isin(set(_tags_ids))]
    postados = meus[meus["STATUS_SIGEM"].map(POSTADO)]
    return fichas_relatorios_html(postados["DOCUMENTO_ESPERADO"].tolist(), _esperados,
                                  _revisoes_por_doc(cache_key, _sigem))


@st.cache_data(show_spinner=False, max_entries=3)
def fichas_tags_html(cache_key: str, filtro: str, f_seg: str, f_malha: str, pag: int,
                     _ids: list, _resumo: pd.DataFrame, _esperados: pd.DataFrame,
                     _tags: pd.DataFrame, _sigem: pd.DataFrame | None = None) -> str:
    """As fichas das 5.098 TAGs. Cada uma varre a 08_RELATORIOS_ESPERADOS, por
    isso o cache -- sem ele isso se repetia a cada clique."""
    espera = (espera_por_documento(_revisoes_por_doc(cache_key, _sigem))
              if _sigem is not None else None)
    return fichas_modais_html(_ids, _resumo, _esperados, _tags, espera,
                              niveis_na_pagina=True)


def _totais(df: pd.DataFrame) -> str:
    esp = int(df["RELATORIOS_ESPERADOS"].sum())
    emi = int(df["RELATORIOS_APROVADOS"].sum())
    pct = (emi / esp * 100) if esp else 0
    return (
        '<div class="prg-tot">'
        f'<div><span class="prg-tot-lbl">Instrumentos</span><span class="prg-tot-val">{br_num(len(df))}</span></div>'
        f'<div><span class="prg-tot-lbl">Completos</span><span class="prg-tot-val">{br_num(int(df["COMPLETA"].sum()))}</span></div>'
        f'<div><span class="prg-tot-lbl">Avanço</span><span class="prg-tot-val">{br_pct(pct)}</span></div>'
        f'<div><span class="prg-tot-lbl">Valor total</span><span class="prg-tot-val">{br_moeda(df["PRECO_UNITARIO"].sum())}</span></div>'
        f'<div><span class="prg-tot-lbl">Valor avançado</span><span class="prg-tot-val">{br_moeda(df["VALOR_AVANCADO"].sum())}</span></div>'
        "</div>"
    )


def _no(tipo: str, nome: object, r, nivel: int, filhos: str = "",
        aberto: bool = False) -> str:
    """Uma linha da arvore: o + expande os filhos, o codigo abre a ficha."""
    rotulo = f"Sem {tipo.lower()}" if vazio(nome) else esc(nome)
    pct = r["avanco"]
    tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
    valor = "(sem)" if vazio(nome) else str(nome)
    # ancora, nao query param: a ficha ja esta na pagina e abre por :target,
    # sem recarregar
    link = (f'<a class="arv-nome arv-ficha" href="#{_ancora(tipo, valor)}" '
            f'title="Abrir ficha">{rotulo}</a>')
    # cada nivel resume o nivel imediatamente abaixo: SOP conta SSOPs,
    # SSOP conta malhas, malha ja e o ultimo agrupamento antes das TAGs.
    # A malha nao tem sub-nivel, mas a celula vem vazia mesmo assim: sem ela
    # as colunas do grid escorregariam uma casa nessa linha.
    sub = ""
    if "subniveis" in r and pd.notna(r["subniveis"]):
        n = int(r["subniveis"])
        nome_sub = {"FASE": ("SOP", "SOPs"), "SOP": ("SSOP", "SSOP"),
                    "SSOP": ("malha", "malhas")}.get(tipo)
        if nome_sub:
            sub = f"{br_num(n)} {nome_sub[0] if n == 1 else nome_sub[1]}"
    n_tags = int(r["tags"])
    resumo = (
        f'<span class="arv-num arv-sub">{sub}</span>'
        f'<span class="arv-num">{br_num(n_tags)} {"tag" if n_tags == 1 else "tags"}</span>'
        f'<span class="arv-num">{br_num(int(r["emitidos"]))}/{br_num(int(r["esperados"]))} aprov.</span>'
        f'<span class="arv-num arv-val">{br_moeda(r["valor"])}</span>'
        '<span class="arv-avanco">'
        f'<span class="arv-track"><span class="arv-fill {tom}" style="width:{max(pct,1.5):.1f}%;"></span></span>'
        f'<span class="arv-pct">{br_pct(pct)}</span></span>'
    )
    # sem filhos nao ha o que expandir: a linha fica sem o "+"
    if not filhos:
        return (f'<div class="arv-no arv-n{nivel} arv-folha">'
                f'<div class="arv-linha"><span class="arv-vazio"></span>{link}'
                f'{pill_prioridade(r["prioridade"])}{resumo}</div></div>')
    # sem id aqui: a ancora _ancora() pertence a ficha, e um id repetido no no
    # da arvore fazia o :target casar com o no em vez da janela
    return (f'<details class="arv-no arv-n{nivel}"{" open" if aberto else ""}>'
            f'<summary><span class="arv-seta"></span>{link}'
            f'{pill_prioridade(r["prioridade"])}{resumo}</summary>'
            f'<div class="arv-corpo">{filhos}</div></details>')


def _distintos(serie: pd.Series, limite: int = 3) -> str:
    """Valores distintos de uma coluna, resumidos quando sao muitos."""
    vals = sorted({str(v).strip() for v in serie if not vazio(v)})
    if not vals:
        return "—"
    if len(vals) <= limite:
        return ", ".join(vals)
    return f"{', '.join(vals[:limite])} +{br_num(len(vals) - limite)}"


# A escada da hierarquia. Cada ficha mostra o degrau imediatamente abaixo e
# nunca pula: o SOP lista SSOPs, o SSOP lista malhas, e so a malha chega nas
# TAGs. Pular um degrau quebra a sequencia que o usuario segue na arvore, e
# ainda repete a mesma lista de TAGs em tres niveis -- eram 15.294 linhas
# somadas contra 7.138.
FX_ABAIXO = {"FASE": "SOP", "SOP": "SSOP", "SSOP": "MALHA"}
FX_ACIMA = {"SOP": ["FASE"], "SSOP": ["FASE", "SOP"], "MALHA": ["FASE", "SOP", "SSOP"]}
FX_ROTULO_NIVEL = {"FASE": "Fase", "SOP": "SOP", "SSOP": "SSOP", "MALHA": "Malha"}


def _modal_nivel(tipo: str, nome: object, sub: pd.DataFrame, rotulo_tipo: str) -> str:
    """Ficha da fase/SOP/SSOP/malha: o que esta cadastrado nele e o degrau
    seguinte da hierarquia."""
    valor_ancora = "(sem)" if vazio(nome) else str(nome)
    titulo = f"Sem {rotulo_tipo.lower()}" if vazio(nome) else str(nome)
    esp = int(sub["RELATORIOS_ESPERADOS"].sum())
    apr = int(sub["RELATORIOS_APROVADOS"].sum())
    pos = int(sub["RELATORIOS_POSTADOS"].sum()) if "RELATORIOS_POSTADOS" in sub.columns else 0
    pen = esp - apr
    pct = (apr / esp * 100) if esp else 0
    tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
    completas = int(sub["COMPLETA"].sum())

    # ------------------------------------------- trilha: a cadeia ate aqui
    # A ficha de nivel so existe na aba Progresso, e todas as outras estao na
    # mesma pagina: a trilha sobe por ancora, sem recarregar nada. So vale para
    # o pai unico -- quando o nivel atravessa dois pais, "SOP-1 e mais 1" nao
    # aponta para ficha nenhuma, entao fica como texto.
    trilha = []
    for pai in FX_ACIMA.get(tipo, []):
        v = _distintos(sub[pai], 1)
        unico = sub[pai].nunique() == 1 and v not in ("—", "-", "")
        trilha.append((f"{FX_ROTULO_NIVEL[pai]} {v}",
                       f"#{_ancora(pai, v)}" if unico else ""))
    trilha.append((f"{FX_ROTULO_NIVEL[tipo]} {titulo}", ""))

    # --------------------------------------------------------------- tiles
    n_mal = sub[~sub.MALHA.apply(vazio)].MALHA.nunique()
    tiles = fx_tile("Instrumentos", br_num(len(sub)), "tag", "#2dd4bf",
                    f"{br_num(completas)} completos")
    filho = FX_ABAIXO.get(tipo)
    if filho:
        quantos = n_mal if filho == "MALHA" else sub[filho].nunique()
        tiles += fx_tile({"SOP": "SOPs", "SSOP": "SSOPs", "MALHA": "Malhas"}[filho],
                         br_num(quantos), "livro", "#5b8def")
    tiles += (
        fx_tile("Prioridade", prioridade_do_grupo(sub["SUBGRUPO_PRIORIDADE"]),
                "alerta", "#fbbf24")
        + fx_tile("Valor total", br_moeda(sub["PRECO_UNITARIO"].sum()), "moeda", "#9d6bff")
        + fx_tile("Valor avançado", br_moeda(sub["VALOR_AVANCADO"].sum()), "ok", "#34d399")
    )

    # --------------------------------------------------------------- KPIs
    kpis = (
        fx_kpi("Aprovados", br_num(apr), f"{br_pct(pct)} dos esperados", pct, "#2dd4bf", "ok")
        + fx_kpi("Postados no SIGEM", br_num(pos), "",
                 (pos / esp * 100) if esp else 0, "#fbbf24", "nuvem")
        + fx_kpi("Pendentes", br_num(pen), f"{br_pct((pen / esp * 100) if esp else 0)} dos esperados",
                 (pen / esp * 100) if esp else 0, "#f87171", "relogio")
        + fx_kpi("Instrumentos completos", br_num(completas),
                 f"de {br_num(len(sub))} no nível",
                 (completas / len(sub) * 100) if len(sub) else 0, "#5b8def", "tag")
    )
    dados = (
        fx_dado("Segmento" if tipo == "MALHA" else "Segmentos",
                _distintos(sub.SEGMENTO) if tipo == "MALHA"
                else br_num(sub[~sub.SEGMENTO.apply(vazio)].SEGMENTO.nunique()))
        + fx_dado("Tipos de instrumento", _distintos(sub.GRUPO_REGRA.str.title())
                  if "GRUPO_REGRA" in sub.columns else "—")
        + fx_dado("Critério de medição", _distintos(sub.CRITERIO_MEDICAO, 2)
                  if "CRITERIO_MEDICAO" in sub.columns else "—")
    )

    # ----------------------------------- o degrau seguinte, sem pular nenhum
    if tipo == "MALHA":
        rotulo_sub, corpo = "Instrumentos", _tabela_tags(sub, com_modal=True)
        conta = f"{br_num(len(sub))} tags"
    else:
        neto = {"FASE": "SSOP", "SOP": "MALHA"}.get(tipo)
        rotulo_sub = {"SOP": "SOPs", "SSOP": "SSOPs", "MALHA": "Malhas"}[filho]
        corpo = _tabela_niveis(sub, filho, neto)
        quantos = n_mal if filho == "MALHA" else sub[filho].nunique()
        conta = f"{br_num(quantos)} no nível abaixo"

    direita = fx_painel(
        f"Avanço do {FX_ROTULO_NIVEL[tipo].lower()}", "seta",
        fx_rosca(apr, esp)
        + '<div class="fx-leg">'
        + fx_lg("Aprovados", br_num(apr), br_pct(pct), "#2dd4bf")
        + fx_lg("Pendentes", br_num(pen), br_pct((pen / esp * 100) if esp else 0), "#f87171")
        + fx_lg("Esperados", br_num(esp), "", "#3a4a68", total=True)
        + "</div>"
        + f'<p class="fx-nota">{br_num(completas)} de {br_num(len(sub))} instrumentos '
          f"fecharam 100% documental.</p>",
        classe_corpo="centro")

    return (
        f'<div class="fmodal" id="{_ancora(tipo, valor_ancora)}">'
        '<a class="fmodal-bg" href="#" aria-label="Fechar"></a>'
        '<div class="fmodal-box"><div class="fmodal-head">'
        f'<div><div class="fn-tipo">{esc(rotulo_tipo)}</div>'
        f'<div class="fmodal-title">{esc(titulo)}</div></div>'
        '<div class="fn-avanco">'
        f'<div class="fn-track"><div class="fn-fill {tom}" style="width:{max(pct, 1.5):.1f}%;"></div></div>'
        f'<div class="fn-pct">{br_pct(pct)}</div></div>'
        '<a class="fmodal-x" href="#" aria-label="Fechar">&times;</a></div>'
        f'<div class="fmodal-body"><div class="fx">{fx_trilha(trilha)}'
        f'<div class="fx-tiles">{tiles}</div>'
        '<div class="fx-corpo"><div class="fx-col">'
        + fx_painel("Resumo documental", "grade",
                    f'<div class="fx-kpis">{kpis}</div><div class="fx-dados">{dados}</div>')
        + fx_painel(rotulo_sub, "livro" if tipo != "MALHA" else "tag", corpo,
                    conta=conta, classe_corpo="zero")
        + f'</div><div class="fx-col">{direita}</div></div></div></div></div></div>'
    )


def _tabela_niveis(sub: pd.DataFrame, coluna: str, subnivel: str = None) -> str:
    """Os filhos diretos de um nivel: o que o SOP mostra dos SSOPs dele.

    Cada linha leva a propria ficha, entao dentro da ficha do SOP da para
    descer para a do SSOP sem fechar nada.
    """
    rotulo = {"SOP": "SOP", "SSOP": "SSOP", "MALHA": "Malha"}[coluna]
    conta_sub = {"SSOP": "#SSOPs", "MALHA": "#Malhas"}.get(subnivel)
    linhas = []
    for _, r in agrega_nivel(sub, coluna, subnivel=subnivel).iterrows():
        nome = r[coluna]
        alvo = "(sem)" if vazio(nome) else str(nome)
        rot = f"Sem {rotulo.lower()}" if vazio(nome) else esc(nome)
        pct = r["avanco"]
        tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
        n_sub = ""
        if conta_sub and "subniveis" in r and pd.notna(r["subniveis"]):
            n_sub = f'<td class="gtbl-num">{br_num(int(r["subniveis"]))}</td>'
        linhas.append(
            f'<tr><td><a class="gtbl-tag gtbl-link" href="#{_ancora(coluna, alvo)}">{rot}</a></td>'
            f'<td class="gtbl-num">{pill_prioridade(r["prioridade"])}</td>{n_sub}'
            f'<td class="gtbl-num">{br_num(int(r["tags"]))}</td>'
            f'<td class="gtbl-num">{br_num(int(r["emitidos"]))}/{br_num(int(r["esperados"]))}</td>'
            f'<td class="gtbl-num gtbl-muted">{br_moeda(r["valor"])}</td>'
            f'<td class="gtbl-num"><span class="gtbl-badge {tom}">{br_pct(pct)}</span></td></tr>'
        )
    cab = [rotulo, "#Prioridade"] + ([conta_sub] if conta_sub else []) + \
          ["#TAGs", "#Aprov./Esp.", "#Valor", "#Avanço"]
    return html_table(cab, "".join(linhas), f"Nenhum {rotulo.lower()}.")


# Os estados de campo, classificados pelo que significam. Antes so existiam
# "bom" e "ruim", e todo o resto caia em ambar -- o que punha "EM COMPRA",
# "EM REPARO" e os 4.637 "Nao Programado" com cara de alerta. Nenhum deles e
# problema: sao etapas do caminho.
STATUS_BONS = {"LOCALIZADO", "APROVADO", "MONTADO", "CALIBRADO", "SIM", "APTO"}
STATUS_ANDAMENTO = {"ON DEMAND", "EM COMPRA", "EM REPARO", "CALIBRAR",
                    "EM PROGRAMAÇÃO", "EM PROGRAMACAO", "NÃO PROGRAMADO",
                    "NAO PROGRAMADO", "ALMOXARIFADO CONSAG"}
STATUS_RUINS = {"NAO LOCALIZADO", "NÃO LOCALIZADO", "REPROVADO", "NAO MONTADO",
                "NÃO MONTADO", "CANCELADO", "REMOVER", "NAO APTO", "NÃO APTO"}


# A previsao de fornecimento so existe para o que ainda esta sendo comprado.
# Nos outros estados a coluna vem vazia na base, e mostrar a linha assim mesmo
# sugeriria que falta preencher algo -- quando na verdade nao se aplica.
EM_COMPRA = {"ON DEMAND", "EM COMPRA"}


def painel_previsao_fornecimento(status_final: object, previsao: object) -> str:
    """O card "Previsão de fornecimento" da ficha da TAG.

    Só existe para ON DEMAND e EM COMPRA -- nos outros estados a coluna vem
    vazia na base, e mostrar o card assim mesmo sugeriria que falta preencher
    algo, quando na verdade não se aplica.

    Com data, diz se atrasou e de quanto, ou em quantos dias chega. Sem data,
    diz que não há definição: é diferente de não se aplicar, e por isso o card
    aparece do mesmo jeito, com o traço.
    """
    if str(status_final).strip().upper() not in EM_COMPRA:
        return ""

    d = pd.to_datetime(previsao, dayfirst=True, errors="coerce")
    if pd.isna(d):
        corpo = (fx_linha("Data prevista", '<span class="gtbl-muted">—</span>')
                 + fx_linha("Situação",
                            '<span class="gtbl-badge mudo">Sem definição</span>'))
        return fx_painel("Previsão de fornecimento", "caixa", corpo)

    hoje = pd.Timestamp.now(tz=BR_TZ).tz_localize(None).normalize()
    dias = (pd.Timestamp(d).normalize() - hoje).days
    if dias < 0:
        tom, texto = "crit", f"Atrasado há {br_num(-dias)} dia{'s' if dias != -1 else ''}"
    elif dias == 0:
        tom, texto = "warn", "Recebimento hoje"
    else:
        tom, texto = "ok", f"No prazo · em {br_num(dias)} dia{'s' if dias != 1 else ''}"
    corpo = (fx_linha("Data prevista", f"<b>{d:%d/%m/%Y}</b>")
             + fx_linha("Situação", f'<span class="gtbl-badge {tom}">{texto}</span>'))
    return fx_painel("Previsão de fornecimento", "caixa", corpo)


def status_pill(v: object) -> str:
    """Status de campo colorido pelo que significa.

    O ambar ficou reservado para valor que nao esta em nenhuma das listas: se
    a base ganhar um estado novo, ele aparece em destaque em vez de se
    confundir com os que ja foram classificados.
    """
    if vazio(v):
        return '<span class="gtbl-muted">—</span>'
    t = str(v).strip()
    n = t.upper()
    if n in STATUS_BONS:
        tom = "ok"
    elif n in STATUS_ANDAMENTO:
        tom = "andamento"
    elif n in STATUS_RUINS:
        tom = "crit"
    else:
        tom = "warn"
    return f'<span class="gtbl-badge {tom}">{esc(t)}</span>'


# A cadeia de navegacao. Segmento e malha nao sao sub-hierarquias reais de
# SOP/SSOP (92 dos 142 segmentos atravessam varios SSOP, e 77% das TAGs nao
# tem segmento), entao cada nivel oferece tambem o grupo "sem <nivel>" e a
# opcao de ver os instrumentos ali mesmo -- assim nenhuma TAG fica inalcancavel.
SEM = "(sem)"


















CABECALHO_TAGS = ["Tag", "Descrição", "#Prioridade", "#Aprov./Esp.", "#Avanço",
                  "#Localização", "#Calibração", "#Montagem", "#Status final",
                  "#Preço unit."]

_COLS_TAGS = ("TAG", "DESCRICAO", "SUBGRUPO_PRIORIDADE", "RELATORIOS_APROVADOS",
              "RELATORIOS_ESPERADOS", "AVANCO_DOCUMENTAL", "STATUS_LOCALIZACAO",
              "STATUS_CALIBRACAO", "STATUS_MONTAGEM", "STATUS_FINAL", "PRECO_UNITARIO")


def linhas_tags(sub: pd.DataFrame, com_modal: bool = True) -> pd.Series:
    """O <tr> de cada TAG, indexado como o dataframe que entrou.

    Percorre arrays com zip em vez de iterrows: o iterrows monta uma Series
    por linha, e sao 5.097 delas na arvore inteira. So essa troca derruba a
    montagem de 1,9 s para 0,4 s.

    Sem indentacao e sem class por celula: cada byte aqui vira ~5 KB de pagina.
    """
    if sub.empty:
        return pd.Series(dtype=object)
    sub = sub.sort_values("AVANCO_DOCUMENTAL", ascending=False)
    vazia = np.full(len(sub), None)
    col = {c: (sub[c].values if c in sub.columns else vazia) for c in _COLS_TAGS}

    linhas = []
    for tag, desc, prio, emi, esp, av, loc, cal, mon, fim, preco in zip(
            *(col[c] for c in _COLS_TAGS)):
        pct = (av or 0) * 100
        tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
        # sem modal, a pill leva para a ficha na aba Pesquisa tag
        alvo = tag_link(tag) if com_modal else (
            f'<a class="gtbl-tag gtbl-link" href="{com_filtros("/pesquisa?tag=" + quote(str(tag)))}" '
            f'target="_self">{esc(tag)}</a>')
        linhas.append(
            f"<tr><td>{alvo}</td><td>{esc(desc)}</td>"
            f"<td>{pill_prioridade(prio)}</td>"
            f"<td>{int(emi)}/{int(esp)}</td>"
            f'<td><span class="gtbl-badge {tom}">{br_pct(pct)}</span></td>'
            f"<td>{status_pill(loc)}</td><td>{status_pill(cal)}</td><td>{status_pill(mon)}</td>"
            f"<td>{status_pill(fim)}</td>"
            f'<td class="gtbl-muted">{br_moeda(preco)}</td></tr>'
        )
    return pd.Series(linhas, index=sub.index)


def _tabela_tags(sub: pd.DataFrame, com_modal: bool = True) -> str:
    return html_table(CABECALHO_TAGS, "".join(linhas_tags(sub, com_modal)),
                      "Nenhuma TAG.", classe="gtbl gtbl-tags")


# Medido na base real com as 1.883 malhas abertas: a arvore inteira da 4,5 MB
# e custa ~2 s a mais que a aba Relatorios. Cabe. O que nao cabe e a ficha
# completa da TAG (~4,8 KB cada, 23 MB no total) -- por isso a pill leva para a
# aba Pesquisa em vez de montar o modal aqui.






def _graficos(df: pd.DataFrame):
    """Os quatro recortes mais avancados: onde ha mais chance de fechar rapido.

    Um grid CSS unico, e nao st.columns: cada coluna do Streamlit empilha de
    forma independente, entao com poucos itens as colunas ficavam com alturas
    diferentes (chegou a 218px de desequilibrio) e abria um vao no meio.
    """
    blocos = "".join([
        grafico_avanco("Fases mais avançadas", agrega_nivel(df, "FASE", subnivel="SOP"),
                       "FASE", rotulo_sub="SOP"),
        grafico_avanco("SOP mais avançados", agrega_nivel(df, "SOP", subnivel="SSOP"),
                       "SOP", rotulo_sub="SSOP"),
        grafico_avanco("SSOP mais avançados", agrega_nivel(df, "SSOP", subnivel="MALHA"),
                       "SSOP", rotulo_sub="malhas"),
        grafico_avanco("Malhas mais avançadas", agrega_nivel(df, "MALHA"), "MALHA"),
    ])
    render_html(f'<div class="gr-grid">{blocos}</div>')


# ---------------------------------------------------------------- filtros
# grupo e status sao guardados com o rotulo normalizado (Instrumento, Em
# analise); fase e sop guardam o codigo cru, que ja vem legivel da planilha.
NORMALIZA = {"grupo", "status"}
# O ?status= ja era do multiselect da aba Relatorios. Se o filtro global usasse
# o mesmo nome, clicar numa fatia do donut -- que so quer ver aqueles
# relatorios -- passava a recortar o app inteiro para as tags que tem aquele
# status. Sao intencoes diferentes, entao sao parametros diferentes.
URL_DO_FILTRO = {"status": "st_tag"}
FILTROS = [
    ("fase", "Fase", "Todas", "tags", "FASE"),
    ("sop", "SOP", "Todos", "tags", "SOP"),
    ("grupo", "Grupo de instrumento", "Todos", "resumo", "GRUPO_REGRA"),
    ("status", "Status SIGEM", "Todos", "esperados", "STATUS_SIGEM"),
]


def _valores(serie: pd.Series) -> list:
    limpo = serie.dropna().astype(str).str.strip()
    return sorted({v for v in limpo if v and v.lower() not in ("nan", "-", "none")})


def _tags_do_filtro(chave: str, valor: str, tags, resumo, esperados) -> set:
    """TAGs que atendem a UM filtro. O status e por tag, nao por relatorio:
    'Recusado' devolve as tags que tem ao menos um relatorio recusado. Filtrar
    a lista de relatorios em si mudaria o denominador do avanco e as contas da
    tela passariam a se contradizer."""
    if chave == "fase":
        return set(tags.loc[tags["FASE"].astype(str).str.strip() == valor, "TAG"])
    if chave == "sop":
        return set(tags.loc[tags["SOP"].astype(str).str.strip() == valor, "TAG"])
    if chave == "grupo":
        return set(resumo.loc[resumo["GRUPO_REGRA"].map(sentence_case) == valor, "TAG"])
    if chave == "status":
        return set(esperados.loc[esperados["STATUS_SIGEM"].map(sentence_case) == valor, "TAG"])
    return set()


def _universo(escolhas: dict, tags, resumo, esperados, pular: str = "") -> set:
    ids = set(tags["TAG"])
    for chave, valor in escolhas.items():
        if valor and chave != pular:
            ids &= _tags_do_filtro(chave, valor, tags, resumo, esperados)
    return ids


def _limpar_filtros():
    """Callback do botao. Precisa ser callback: mexer na chave de um widget
    depois que ele ja foi criado no mesmo run levanta excecao no Streamlit --
    os callbacks rodam antes do rerun, quando ainda e permitido."""
    for chave, _, padrao, _, _ in FILTROS:
        st.session_state[f"gf_{chave}"] = padrao


def consumir_filtros_url(tags: pd.DataFrame, resumo: pd.DataFrame,
                         esperados: pd.DataFrame) -> None:
    """Aplica ?fase=/?sop=/?grupo= que chegam de um clique dentro da propria
    tela. Um token evita reaplicar a cada rerun, senao a URL sobrescreveria
    para sempre o que o usuario escolher depois no seletor."""
    vindo = {c: st.query_params.get(URL_DO_FILTRO.get(c, c), "")
             for c, *_ in FILTROS if st.query_params.get(URL_DO_FILTRO.get(c, c))}
    if not vindo:
        return
    token = "|".join(f"{c}={v}" for c, v in sorted(vindo.items()))
    if st.session_state.get("_flt_url") == token:
        return
    st.session_state["_flt_url"] = token
    base = {"tags": tags, "resumo": resumo, "esperados": esperados}
    for chave, _, padrao, fonte, coluna in FILTROS:
        if chave not in vindo:
            continue
        serie = base[fonte][coluna]
        validos = _valores(serie.map(sentence_case) if chave in NORMALIZA else serie)
        if vindo[chave] in validos:
            st.session_state[f"gf_{chave}"] = vindo[chave]


def sidebar_filtros(tags: pd.DataFrame, resumo: pd.DataFrame,
                    esperados: pd.DataFrame) -> dict:
    """Filtros da lateral, em cascata e validos para o app inteiro.

    Cada campo so oferece o que ainda sobra depois dos outros: escolher uma
    FASE reduz a lista de SOPs aos daquela fase, e escolher um SOP reduz os
    grupos. Sem isso da para montar uma combinacao que nao devolve nada e
    parece defeito. Quando o valor guardado sai da lista -- porque outro
    filtro mudou -- ele volta ao padrao em vez de estourar.

    A chave do session_state E a chave do widget, de proposito: com duas
    chaves o valor vindo da URL era escrito numa e o selectbox continuava
    lendo a outra, e o seletor nao mexia.
    """
    consumir_filtros_url(tags, resumo, esperados)

    def escolhido(chave: str, padrao: str) -> str:
        v = st.session_state.get(f"gf_{chave}", padrao)
        return "" if v == padrao else v

    with st.sidebar:
        # O contador so existe depois que os quatro seletores forem lidos, mas
        # ele mora na mesma linha do titulo, acima deles. Reservar o espaco com
        # st.empty() e preencher no fim economiza a linha inteira -- e era ela
        # que empurrava o ultimo item do menu para tras da barra de tarefas.
        topo = st.empty()
        for chave, rotulo, padrao, fonte, coluna in FILTROS:
            escolhas = {c: escolhido(c, p) for c, _, p, _, _ in FILTROS}
            base = {"tags": tags, "resumo": resumo, "esperados": esperados}[fonte]
            sub = base[base["TAG"].isin(
                _universo(escolhas, tags, resumo, esperados, pular=chave))]
            serie = sub[coluna].map(sentence_case) if chave in NORMALIZA else sub[coluna]
            opcoes = [padrao] + _valores(serie)
            if st.session_state.get(f"gf_{chave}", padrao) not in opcoes:
                st.session_state[f"gf_{chave}"] = padrao
            st.selectbox(rotulo, opcoes, key=f"gf_{chave}")

        escolhas = {c: escolhido(c, p) for c, _, p, _, _ in FILTROS}
        ativos = {c: v for c, v in escolhas.items() if v}
        alvo = _universo(escolhas, tags, resumo, esperados)
        topo.markdown(
            '<div class="flt-topo"><span class="flt-titulo">Filtros rápidos</span>'
            f'<span class="flt-conta"><b>{br_num(len(alvo))}</b> de {br_num(len(tags))}</span>'
            "</div>", unsafe_allow_html=True)
        if ativos:
            st.button("Limpar filtros", key="gf_limpar", on_click=_limpar_filtros)

    rotulos = {c: r for c, r, *_ in FILTROS}
    st.session_state["_flt_selo"] = (
        '<div class="du-selo filtro"><i></i>'
        + " · ".join(f"{rotulos[c]}: {esc(v)}" for c, v in ativos.items())
        + f" — {br_num(len(alvo))} de {br_num(len(tags))} tags</div>"
    ) if ativos else ""
    return escolhas


def aplicar_filtros(escolhas: dict, tags: pd.DataFrame, resumo: pd.DataFrame,
                    esperados: pd.DataFrame):
    """Recorta as tres bases para o mesmo conjunto de TAGs.

    As tres tem que sair juntas, senao o Dashboard conta 300 tags e soma os
    25.100 relatorios de todas elas. A base SIGEM crua fica fora: ela e a
    fonte, e recortar a fonte esconderia documento que nao casa com tag
    nenhuma -- justamente o que se quer enxergar la."""
    if not any(escolhas.values()):
        return tags, resumo, esperados
    ids = _universo(escolhas, tags, resumo, esperados)
    return (tags[tags["TAG"].isin(ids)].copy(),
            resumo[resumo["TAG"].isin(ids)].copy(),
            esperados[esperados["TAG"].isin(ids)].copy())


def main():
    favicon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "favicon.png")
    st.set_page_config(
        page_title="Gplan",
        page_icon=favicon if os.path.exists(favicon) else "📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()
    seletor_tema()

    # O carimbo da planilha e a chave de cache sao coisas diferentes: a chave
    # leva a versao da regra colada no fim, e quem le a data precisa do valor
    # cru. Passar a chave para data_atualizacao deixava o cabecalho em "—".
    fonte = get_source_cache_key()
    cache_key = f"{fonte}|r{REGRA_VERSAO}|v{VISUAL_VERSAO}"
    if cache_key == "missing":
        st.error(
            "Não encontrei a planilha no Supabase Storage nem localmente. "
            "Verifique se o arquivo foi enviado ao bucket 'gplan-data'."
        )
        st.stop()

    st.session_state["gplan_atualizado_em"] = data_atualizacao(fonte)
    (tags, cabos, tubing, sigem, resumo, esperados,
     gitec, locacao, aux_areas) = load_data(cache_key)

    with st.sidebar:
        render_html(
            '<div class="gplan-brand">'
            + _logo_svg("Lat").replace("<svg ", '<svg class="gplan-brand-mark" ')
            + '<div class="gplan-brand-text">'
            '<div class="gplan-brand-name">Gplan</div>'
            '<div class="gplan-brand-sub">Instrumentação · U-12</div>'
            "</div></div>"
        )

    escolhas = sidebar_filtros(tags, resumo, esperados)
    tags, resumo, esperados = aplicar_filtros(escolhas, tags, resumo, esperados)
    # a medicao segue as tags: filtrou a fase, a aba Gitec mostra so o que foi
    # medido nela
    gitec_f = gitec[gitec["TAG"].isin(set(tags["TAG"]))] if not gitec.empty else gitec

    dashboard_page = st.Page(lambda: _sob_carga("Carregando o painel", lambda: render_dashboard(resumo, esperados, tags, sigem, cache_key)), title="Dashboard", icon=":material/dashboard:", url_path="dashboard", default=True)
    relatorios_page = st.Page(lambda: _sob_carga("Carregando os relatórios", lambda: render_relatorios(esperados, resumo, tags, sigem, cache_key)), title="Relatórios", icon=":material/description:", url_path="relatorios")
    progresso_page = st.Page(lambda: _sob_carga("Abrindo o Progresso", lambda: render_progresso(resumo, esperados, tags, sigem, cache_key)), title="Progresso", icon=":material/insights:", url_path="progresso")
    pesquisa_page = st.Page(lambda: _sob_carga("Carregando as tags", lambda: render_pesquisa_tag(resumo, esperados, tags, sigem, cache_key)), title="Pesquisa tag", icon=":material/search:", url_path="pesquisa")
    sigem_page = st.Page(lambda: _sob_carga("Carregando a base SIGEM", lambda: render_sigem(sigem, esperados, any(escolhas.values()))), title="Base SIGEM", icon=":material/database:", url_path="sigem")
    gitec_page = st.Page(lambda: _sob_carga("Carregando a medição de campo", lambda: render_gitec(gitec_f, resumo, tags, esperados, sigem, cache_key)), title="Gitec", icon=":material/engineering:", url_path="gitec")
    planta_page = st.Page(lambda: _sob_carga("Desenhando o avanço na planta", lambda: render_planta(tags, resumo, locacao, aux_areas)), title="Planta", icon=":material/map:", url_path="planta")

    nav = st.navigation([dashboard_page, progresso_page, relatorios_page, pesquisa_page,
                         sigem_page, gitec_page, planta_page], position="sidebar")
    nav.run()


if __name__ == "__main__":
    main()
