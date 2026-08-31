import base64
import collections
import io
import json
import math
import os
import re
import time
from contextlib import contextmanager
from urllib.parse import quote

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import acesso

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
def _logo_svg(sufixo: str = "", px: int = 48) -> str:
    """A marca.

    O width/height no próprio SVG não é decoração: um SVG que só tem viewBox
    ocupa TODA a largura disponível até o CSS chegar, e no instante entre o
    HTML e a folha de estilo a marca enchia a tela. O CSS continua mandando
    onde há regra; isto é só o tamanho de partida, para não haver instante
    nenhum sem tamanho.
    """
    grad = f"gpArc{sufixo}"
    return (
        f'<svg width="{px}" height="{px}" viewBox="0 0 48 48" fill="none">'
        '<circle cx="24" cy="24" r="19" stroke="var(--text-3)" stroke-width="5"/>'
        f'<path d="M24 5a19 19 0 0 1 15.6 29.8" stroke="url(#{grad})" stroke-width="5" stroke-linecap="round"/>'
        '<circle cx="24" cy="24" r="5.5" fill="var(--accent-teal)"/>'
        '<path d="M24 24L33 15" stroke="var(--text-1)" stroke-width="3" stroke-linecap="round"/>'
        f'<defs><linearGradient id="{grad}" x1="24" y1="5" x2="40" y2="35" gradientUnits="userSpaceOnUse">'
        '<stop stop-color="var(--accent-blue)"/><stop offset="1" stop-color="var(--accent-teal)"/>'
        "</linearGradient></defs></svg>"
    )


LOGO_SVG = _logo_svg(px=58)


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

# As barras do Dashboard tem que somar o KPI: mesma pergunta, mesma resposta.
# Duas coisas quebravam isso e deixavam 598 documentos fora da conta.
#
# 1. Duas familias inteiras nao tinham barra nenhuma -- o RIMSI do pedestal
#    (629 documentos) e o RILM de placa de orificio (71). O RIMSI so aparecia
#    pela regra de infra por planta, que sao outros 20 documentos; os 629 do
#    pedestal simplesmente nao existiam na tela.
# 2. Todas as linhas contavam LINHA, e a linha e o par (TAG, documento). Um
#    RILTCI de cabo serve as duas pontas e um CTECRI serve mais de uma TAG:
#    781 linhas para 722 documentos, 1.311 para 1.268. O KPI conta documento
#    unico, entao as barras infladas nunca fechavam com ele.
#
# Por isso toda linha aqui conta documento unico (o ultimo campo): e a mesma
# unidade de totais_por_documento. Onde nao ha compartilhamento o numero nao
# muda, e onde houver amanha a tela nao volta a divergir sozinha.
REPORT_ROWS = [
    ("RIR instrumentos", "RIR", "BASE: RIR obrigatorio para todos os TAGs", True),
    ("RIR cabos", "RIR", "CONDICIONAL: RIR de cabo por TAG", True),
    ("CCP", "CCP", None, True),
    ("RTFCJI", "RTFCJI", None, True),
    ("RIMITPI", "RIMITPI", None, True),
    ("RIFMI", "RIFMI", None, True),
    ("RIMTU", "RIMTU", None, True),
    ("RILTCI", "RILTCI", None, True),
    ("RIMJBI", "RIMJBI", None, True),
    ("CTECRI", "CTECRI", None, True),
    ("RILM placa de orifício", "RILM", "BASE: descricao contem Placa de Orificio", True),
    ("RIMSI pedestal", "RIMSI", "BASE: RIMSI do pedestal (08_BASE_PEDESTAL)", True),
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
    # O controle de lancamento de circuitos e o de-para de TAG so existem em
    # planilha gerada pelo pipeline novo. Sem eles a aba Certificacao abre
    # explicando o que rodar, em vez de quebrar.
    lancamento = (pd.read_excel(excel_file, sheet_name="02_CABOS_LANCAMENTO")
                  if "02_CABOS_LANCAMENTO" in excel_file.sheet_names
                  else pd.DataFrame(columns=["CIRCUITO", "ORIGEM", "DESTINO", "DISCIPLINA",
                                             "TIPO", "STATUS", "PCT", "METROS"]))
    depara = (pd.read_excel(excel_file, sheet_name="02_CABOS_DEPARA")
              if "02_CABOS_DEPARA" in excel_file.sheet_names
              else pd.DataFrame(columns=["PONTA", "TAG", "COMO"]))
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
    # O que mudou de uma atualizacao para a outra. A base de TAGs e a de cabos
    # nao guardam historico -- quem apura a diferenca e o pipeline, no momento
    # em que grava a planilha, e deixa pronto nesta aba.
    movimentacoes = (pd.read_excel(excel_file, sheet_name="14_MOVIMENTACOES")
                     if "14_MOVIMENTACOES" in excel_file.sheet_names
                     else pd.DataFrame(columns=["DATA", "TIPO", "OBJETO", "CAMPO",
                                                "DE", "PARA", "QTD_TAGS"]))
    resumo = aplicar_regra_aprovados(resumo, esperados)
    return (tags, cabos, tubing, sigem, resumo, esperados, gitec, locacao,
            aux_areas, lancamento, depara, movimentacoes)


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


def totais_por_documento(esperados_sub: pd.DataFrame) -> tuple[int, int, int, int]:
    """Esperados/postados/aprovados/pendentes por DOCUMENTO unico, nao por par
    (TAG, documento).

    Um documento pode ser esperado por varias TAGs -- o RIMII/RIMSI de infra
    por planta passa de 700. RELATORIOS_ESPERADOS por TAG (em aplicar_regra_
    aprovados) conta certo dentro de cada TAG, mas somar essa coluna entre
    varias TAGs conta o mesmo documento compartilhado uma vez por TAG
    pendurada nele -- infla esperados e pendentes bem acima do numero real de
    relatorios distintos. Aqui o documento entra uma unica vez, do mesmo jeito
    que du_status ja faz para o donut de Status SIGEM.
    """
    unicos = esperados_sub.drop_duplicates(subset=["DOCUMENTO_ESPERADO"])
    esp = len(unicos)
    pos = int(unicos["EXISTE_NO_SIGEM"].astype(str).str.strip().str.upper().eq("SIM").sum())
    apr = int(aprovado(unicos["STATUS_SIGEM"]).sum())
    pen = esp - apr
    return esp, pos, apr, pen


def totais_por_documento_agrupado(df: pd.DataFrame, esperados: pd.DataFrame,
                                  coluna: str) -> pd.DataFrame:
    """Como totais_por_documento, mas por grupo (FASE/SOP/SSOP/MALHA/GRUPO_
    REGRA/...) em vez do total geral -- usado por toda a arvore de Progresso e
    pelo Resumo por grupo do Dashboard.

    Um documento compartilhado entre TAGs de grupos diferentes (RIMII/RIMSI de
    infra por planta atravessa FASE/SOP/SSOP e ate tipo de instrumento) conta
    uma vez DENTRO de cada grupo que ele afeta -- por isso os grupos, somados,
    podem passar do total geral: cada um responde "quanto desse grupo esta
    aprovado", nao uma fatia de uma torta unica.
    """
    grupo_por_tag = df.set_index("TAG")[coluna]
    e = esperados[esperados["TAG"].isin(set(df["TAG"]))][
        ["TAG", "DOCUMENTO_ESPERADO", "STATUS_SIGEM"]].copy()
    e["_grp"] = e["TAG"].map(grupo_por_tag)
    unicos = e.drop_duplicates(subset=["DOCUMENTO_ESPERADO", "_grp"])
    return unicos.groupby("_grp").agg(
        esperados=("DOCUMENTO_ESPERADO", "size"),
        emitidos=("STATUS_SIGEM", lambda s: int(aprovado(s).sum())),
    )


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
        # a pastilha da etiqueta da zona na Planta usa isto -- ver o comentario
        # em .pl-mio
        "rgb_chapa": "8,12,22",
        # o metal do equipamento no desenho da Certificação
        "metal": "#2a3350", "metal2": "#212942", "metal3": "#39456b",
        # o desenho da planta vem preto sobre branco: invertido, o traço fica
        # claro sobre o escuro e a prancha para de ser um retângulo branco
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
        "rgb_chapa": "255,255,255",
        "metal": "#c4cdde", "metal2": "#b2bcd2", "metal3": "#dbe1ee",
        # no claro o desenho já é escuro sobre branco: inverter deixaria a
        # prancha preta no meio de uma tela clara
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
        # O cookie e a memoria entre visitas. Chega junto com o pedido da
        # pagina, entao da para ler aqui, antes de qualquer CSS sair -- o
        # localStorage so responderia depois, e a tela nasceria num tema e
        # trocaria no meio do carregamento.
        try:
            escolhido = st.context.cookies.get("gplan_tema")
        except Exception:
            escolhido = None
    if escolhido not in TEMAS:
        escolhido = TEMA_PADRAO
    st.session_state["gplan_tema"] = escolhido
    # Nao repor "tema" na query aqui. st.query_params[...] = x reconstroi a URL
    # no navegador (para refletir o novo parametro) e essa reconstrucao nao
    # preserva o #fragmento -- um link como /progresso#n-FASE-X chegava com o
    # hash, e sumia 2s depois quando este trecho rodava, no meio da propria
    # navegacao para a ficha de nivel. O cookie (lembrar_tema, mais abaixo) ja
    # basta para a preferencia sobreviver a navegar e recarregar; a URL nao
    # precisa carregar "tema" tambem.
    return escolhido


def lembrar_tema(tema: str) -> None:
    """Grava a escolha no navegador, para a proxima vez que a pessoa abrir.

    Cookie, e nao localStorage, porque este valor precisa existir no servidor
    na primeira execucao -- e o st.context.cookies quem o le. Um ano de prazo:
    a preferencia de tema nao envelhece.
    """
    st.components.v1.html(
        "<script>try{parent.document.cookie="
        f"'gplan_tema={tema};path=/;max-age=31536000;samesite=lax'}}catch(e){{}}</script>",
        height=0)


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
    # o cookie e reescrito no proximo desenho, por lembrar_tema


# ===================================================================== #
# Acesso: quem entrou decide o que a tela mostra                         #
# ===================================================================== #
# A pergunta que originou isto era "como esconder valor numa reunião". Um
# botão de modo apresentação resolveria a tela e não o controle: qualquer um
# clica de volta, inclusive na frente do cliente. Aqui o que aparece é
# consequência do login, e não uma preferência que se desliga.

COOKIE_SESSAO = "gplan_sessao"


def lembrar_sessao(token: str) -> None:
    """Guarda a sessão no navegador.

    Mesmo mecanismo do tema: cookie lido pelo servidor no primeiro instante.
    Sem ele, recarregar a página derrubaria o login -- o session_state não
    sobrevive a um F5. O que vai no cookie é o refresh_token do Supabase, e
    não o access_token: o access vence em uma hora e a sessão cairia no meio
    do expediente.
    """
    if not token:
        return
    st.components.v1.html(
        "<script>try{parent.document.cookie="
        f"'{COOKIE_SESSAO}={token};path=/;max-age=2592000;samesite=lax'}}"
        "catch(e){}</script>", height=0)


def esquecer_sessao() -> None:
    st.components.v1.html(
        "<script>try{parent.document.cookie="
        f"'{COOKIE_SESSAO}=;path=/;max-age=0;samesite=lax'}}catch(e){{}}</script>",
        height=0)


def usuario_atual() -> dict | None:
    """Quem está logado: da sessão, ou do cookie.

    Quem valida é o Supabase -- este código não confere senha nem assinatura.
    Conta desativada ou apagada deixa de retomar sozinha, porque a validação
    acontece lá e não numa cópia local.
    """
    if st.session_state.get("gplan_usuario"):
        return st.session_state["gplan_usuario"]
    try:
        token = st.context.cookies.get(COOKIE_SESSAO)
    except Exception:
        token = None
    if not token:
        return None
    retomado = acesso.retomar(token)
    if not retomado:
        return None
    usuario, novo_token = retomado
    st.session_state["gplan_usuario"] = usuario
    st.session_state["gplan_token"] = novo_token
    return usuario


def pode(permissao: str) -> bool:
    return acesso.pode(st.session_state.get("gplan_usuario"), permissao)


def recarregar_usuario() -> None:
    """Relê o perfil depois de mexer nele, para a tela não ficar com o
    anterior -- trocar a própria foto e continuar vendo a antiga parece que
    não salvou."""
    st.session_state.pop("gplan_usuario", None)
    usuario_atual()


def avisar_erro(acao: str, erro: Exception) -> None:
    """Erro de Supabase com o nome do que falhou.

    Sem isto o que aparece é um traceback de biblioteca, e ninguém liga a
    causa à configuração do projeto -- foi assim que "Email logins are
    disabled" passou por erro de senha.
    """
    conselho = acesso.explicar(erro)
    st.error(f"Não consegui {acao}."
             + (f"\n\n{conselho}" if conselho else "")
             + f"\n\nDetalhe do Supabase: {type(erro).__name__}: {erro}")


# A entrada é a primeira tela que o cliente vê, e às vezes a única antes da
# reunião começar. O movimento aqui é de chegada -- a marca aparece, o cartão
# sobe -- e para depois: animação que fica repetindo em tela de login vira
# ruído de fundo na hora de digitar a senha. O halo é a única coisa que
# continua, devagar, e some inteiro em prefers-reduced-motion.
ENTRADA_CSS = """
<style>
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display:none; }
.block-container { padding-top:11vh !important; max-width:none !important; }

.lg-halo { position:fixed; inset:0; z-index:0; pointer-events:none; overflow:hidden; }
.lg-halo::before, .lg-halo::after {
  content:""; position:absolute; width:52vw; height:52vw; border-radius:50%;
  filter:blur(90px); opacity:.20;
}
.lg-halo::before { background:var(--accent-blue); top:-18vw; left:-12vw;
  animation:lg-vaga 22s ease-in-out infinite alternate; }
.lg-halo::after { background:var(--accent-teal); bottom:-22vw; right:-10vw;
  animation:lg-vaga 27s ease-in-out infinite alternate-reverse; }
@keyframes lg-vaga { to { transform:translate3d(6vw, 4vh, 0) scale(1.12); } }

.lg-caixa { position:relative; z-index:1; width:min(392px, 92vw); margin:0 auto 14px;
  text-align:center; animation:lg-sobe .55s cubic-bezier(.2,.7,.3,1) both; }
.lg-marca { display:flex; align-items:center; justify-content:center; gap:13px;
  margin-bottom:24px; }
svg.lg-mark { width:44px; height:44px; flex:none;
  filter:drop-shadow(0 4px 16px rgba(var(--rgb-teal),.45));
  animation:lg-marca .7s cubic-bezier(.2,.7,.3,1) both; }
.lg-txt { text-align:left; }
.lg-nome { font-size:27px; font-weight:800; letter-spacing:-.7px; color:var(--text-1);
  line-height:1.05; }
.lg-sub { font-size:11px; color:var(--text-3); letter-spacing:.5px;
  text-transform:uppercase; font-weight:600; margin-top:3px; }
/* o Entrar respira um pouco mais, agora que e o unico botao do formulario */
.st-key-lg_entrar { margin-top:6px; }

/* o cartão é o próprio formulário do Streamlit */
[data-testid="stForm"] { position:relative; z-index:1; width:min(392px, 92vw);
  margin:0 auto; background:var(--dark-card); border:1px solid var(--border-color);
  border-radius:16px; padding:20px 20px 6px;
  box-shadow:0 22px 60px var(--sombra), 0 1px 0 rgba(var(--rgb-tinta),.05) inset;
  animation:lg-sobe .55s .07s cubic-bezier(.2,.7,.3,1) both; }
[data-testid="stForm"] label { font-size:12px !important; font-weight:600 !important;
  color:var(--text-2) !important; }
[data-testid="stForm"] input { border-radius:10px !important; }
/* O Streamlit embrulha o botão num div próprio: sem esticar o embrulho, o
   width:100% do botão mede o embrulho encolhido e não muda nada. E a regra
   precisa ser do botão de ENVIO, não de "button" solto -- dentro do form há
   também o olho de mostrar a senha, que esticado ocupa a linha inteira. */
/* a largura estava presa três níveis acima: o stElementContainer do botão
   nasce do tamanho do texto, e width:100% dentro dele media esses 63 px */
[data-testid="stForm"] [data-testid="stElementContainer"] { width:100% !important; }
[data-testid="stFormSubmitButton"] { width:100%; margin-top:4px; }
[data-testid="stFormSubmitButton"] button { width:100%; border-radius:10px !important;
  font-weight:700 !important; letter-spacing:.2px; transition:transform .12s ease,
  box-shadow .12s ease; }
[data-testid="stFormSubmitButton"] button:hover { transform:translateY(-1px);
  box-shadow:0 8px 22px rgba(var(--rgb-azul),.32); }
[data-testid="stFormSubmitButton"] button:active { transform:translateY(0); }

@keyframes lg-sobe { from { opacity:0; transform:translateY(14px); } }
@keyframes lg-marca { from { opacity:0; transform:translateY(6px) rotate(-14deg) scale(.82); } }

@media (prefers-reduced-motion:reduce) {
  .lg-caixa, [data-testid="stForm"], svg.lg-mark { animation:none; }
  .lg-halo::before, .lg-halo::after { animation:none; }
}
</style>
"""


def traduzir_campos_de_senha() -> None:
    """O botão de mostrar a senha vem do Streamlit, e vem em inglês.

    Não há parâmetro para trocar esse rótulo, e ele é recriado a cada
    redesenho -- por isso um observador, e não uma passada única. Roda dentro
    do iframe de components e alcança o documento de fora por parent.
    """
    st.components.v1.html("""
<script>
try {
  const doc = parent.document;
  const NOMES = {"Show password": "Mostrar a senha",
                 "Hide password": "Ocultar a senha"};
  const traduz = () => {
    doc.querySelectorAll('button[aria-label]').forEach(b => {
      const novo = NOMES[b.getAttribute('aria-label')];
      if (novo) { b.setAttribute('aria-label', novo); b.title = novo; }
    });
  };
  traduz();
  new MutationObserver(traduz).observe(doc.body, {childList: true, subtree: true});
} catch (e) {}
</script>""", height=0)


def exigir_login() -> dict:
    """Sem login não desenha nada. Devolve o usuário ou para a execução.

    Quem confere a senha é o Auth do Supabase. Não há mais "primeiro acesso"
    aqui: a primeira conta nasce pelo painel do Supabase, e criar login virou
    trabalho do administrador, na aba Acessos. Tela de login que aceita criar
    conta é tela que qualquer um usa para entrar.
    """
    u = usuario_atual()
    if u:
        return u

    # Bypass so-para-verificacao-automatizada-local: uma variavel separada
    # do GPLAN_LOCAL que o Daniel ja usa (esse continua exigindo login
    # normal, pelo Supabase). So liga com as DUAS: GPLAN_LOCAL=1 (dados do
    # disco) e GPLAN_DEV_SKIP_LOGIN=1 (usuario fake, so nesta sessao). Nunca
    # ativa no Render -- as duas exigem env var explicita, nenhuma tem
    # default ligado.
    if (os.environ.get("GPLAN_LOCAL") == "1"
            and os.environ.get("GPLAN_DEV_SKIP_LOGIN") == "1"):
        st.session_state["gplan_usuario"] = {
            "id": "dev-local", "nome": "Dev Local", "email": "dev@local.test",
            "permissoes": acesso.TODAS, "ativo": True, "papel": "Administrador",
        }
        return st.session_state["gplan_usuario"]

    # Tudo da entrada mora num container proprio para poder ser APAGADO de
    # uma vez quando o login der certo. Antes daqui vinha um st.rerun(), e o
    # que se via era o formulario esmaecendo enquanto o sistema surgia por
    # tras dele, os dois na tela ao mesmo tempo. Apagar e seguir troca isso
    # por um corte limpo -- e ainda evita autenticar duas vezes.
    tela = st.empty()
    entrada = tela.container()
    with entrada:
        render_html(ENTRADA_CSS)
        traduzir_campos_de_senha()
        render_html(
        '<div class="lg-halo"></div>'
        '<div class="lg-caixa"><div class="lg-marca">'
        + _logo_svg("Lat").replace("<svg ", '<svg class="lg-mark" ')
        + '<div class="lg-txt"><div class="lg-nome">Gplan</div>'
        '<div class="lg-sub">Instrumentação · U-12</div></div></div></div>')

        # Três linhas e um botão: e-mail, senha, entrar. Não há mais nada
        # aqui de propósito -- quem repõe senha é o administrador, pela aba
        # Acessos.
        with st.form("entrar"):
            email = st.text_input("E-mail").strip().lower()
            senha = st.text_input("Senha", type="password")
            entrar = st.form_submit_button("Entrar", type="primary",
                                           key="lg_entrar")

    # Não há "esqueci a senha" aqui: por decisão do Daniel, quem repõe senha é
    # o administrador, pela aba Acessos. Oferecer o link e não ter SMTP seria
    # pior que não oferecer -- a pessoa clicaria, veria "enviado" e esperaria
    # um e-mail que não vem.
    entrou = None
    with entrada:
        if entrar:
            try:
                entrou = acesso.entrar(email, senha)
            except acesso.SemSupabase as erro:
                st.error(str(erro))
                st.stop()
            except Exception as erro:
                avisar_erro("falar com o Supabase para autenticar", erro)
                st.stop()
            if not entrou:
                # a mesma mensagem para conta inexistente, senha errada e
                # conta desativada: dizer qual dos três entrega quem existe
                st.error("E-mail ou senha inválidos.")
            else:
                usuario, token = entrou

    if entrar and entrou:
        tela.empty()          # a tela de login sai inteira, de uma vez
        st.session_state["gplan_usuario"] = usuario
        st.session_state["gplan_token"] = token
        lembrar_sessao(token)
        return usuario
    st.stop()


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
        }
        /* A largura so vale com a lateral aberta. Fixar sem essa condicao
           vencia o recolhimento do Streamlit -- a setinha mudava o estado e a
           barra continuava ocupando os 212 px, com o conteudo parado do lado.
           Fechada, ela vai a zero e o conteudo toma o espaco sozinho, porque
           esta no fluxo normal. */
        section[data-testid="stSidebar"][aria-expanded="true"] {
          width: 212px !important; min-width: 212px !important;
        }
        section[data-testid="stSidebar"][aria-expanded="false"] {
          width: 0 !important; min-width: 0 !important; border-right: 0;
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
        /* O menu fica entre a marca e o filtro rapido (que voltou, order 3
           abaixo, antes do perfil -- ver sidebar_filtros e o comentario no
           main()). Historico: os seletores ja ficaram no topo, antes do
           menu, mas o dropdown do selectbox e portalado no body e abria
           atras da barra de tarefas quando comecava baixo demais -- por
           isso o menu tomou o lugar de cima e o filtro passou para depois
           dele. */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] { order: 2; }
        section[data-testid="stSidebar"] [data-testid="stSidebarNavSeparator"] { display: none; }
        /* O Streamlit desenha esse cabecalho com a cor fixa do tema PROPRIO
           dele (claro pro navegador, nao o claro/escuro que este app
           implementa por conta -- ver sistema de temas). Sem sobrescrever,
           ficava sempre com a cor de texto do escuro, ilegivel no fundo
           claro deste app. */
        section[data-testid="stSidebar"] header[data-testid="stNavSectionHeader"] {
          display:flex; align-items:center; gap:7px;
          color:var(--text-3) !important; }
        section[data-testid="stSidebar"] header[data-testid="stNavSectionHeader"] p {
          color:inherit !important; }
        section[data-testid="stSidebar"] header[data-testid="stNavSectionHeader"]::before {
          content:""; display:inline-block; width:13px; height:13px; flex:none;
          background-color:currentColor;
          -webkit-mask:var(--fxi-secao) center/contain no-repeat;
          mask:var(--fxi-secao) center/contain no-repeat; }
        __ICONES_SECAO__

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
        /* Resumo so das TAGs prioritarias, ao lado do "sub" na mesma linha
           -- a mesma pergunta do cartao, restrita as prioritarias, sem abrir
           outra tela. Cor um pouco mais forte que o sub simples, pra dar pra
           notar que e um recorte, nao so mais texto de legenda. */
        .du-kpi-linha, span.du-kpi-linha { display:flex; align-items:baseline; justify-content:space-between; gap:8px; }
        .du-kpi-prio { font-size:10px; color:var(--txt-ambar); font-weight:650; white-space:nowrap; flex:none; }
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
        /* A ultima coluna da .du-br foi dimensionada para percentual ("90,0%"),
           e na Previsao Medicao ela carrega moeda -- "R$ 2.896.094,34" nao
           cabe em 50px e quebrava em quatro linhas. Esta variante troca so as
           larguras; o resto da barra continua igual. */
        .du-br.moeda { grid-template-columns:clamp(92px,8vw,126px) 1fr
                       clamp(42px,3.4vw,56px) clamp(96px,9vw,132px); }
        .du-br.moeda .pc { font-size:10px; white-space:nowrap; }
        .du-br:hover .nm { color:var(--text-1); }
        .du-br .nm { font-size:10.5px; color:var(--text-2); white-space:nowrap;
                     overflow:hidden; text-overflow:ellipsis; }
        .du-br .tr { height:6px; border-radius:99px; background:rgba(var(--rgb-tinta),.06); overflow:hidden; }
        .du-br .tr i { display:block; height:100%; border-radius:99px;
                       background:linear-gradient(90deg,var(--teal-2),var(--accent-green)); }
        /* variantes semanticas da mesma barra, para a fila da Previsao
           Medicao: verde quem esta pronto, ambar quem esta a uma etapa,
           vermelho o resto. Sem isto tudo saia com o mesmo verde do gradiente
           e a barra deixava de dizer o que dizia. */
        .du-br .tr i.feito { background:var(--accent-teal); }
        .du-br .tr i.andando { background:var(--accent-amber); }
        .du-br .tr i.parado { background:var(--accent-red); }
        /* a barrinha que acompanha "3/4" na tabela: mesma pintura da
           .pl-barra, sem a margem que a de cartao tem */
        .med-barra { height:6px; width:52px; border-radius:99px; display:inline-block;
                     margin-left:6px; vertical-align:middle; overflow:hidden;
                     background:rgba(var(--rgb-tinta),.07); }
        .med-barra i { display:block; height:100%; border-radius:99px; }
        .med-barra i.feito { background:var(--accent-teal); }
        .med-barra i.andando { background:var(--accent-amber); }
        .med-barra i.parado { background:var(--accent-red); }
        /* a coluna de avanco da Previsao Medicao: numero, barra e percentual
           na mesma linha, alinhados, sem a coluna extra que empurrava a
           tabela para o lado */
        .med-av { white-space:nowrap; }
        .med-pct { margin-left:8px; color:var(--text-3); font-size:11px;
                   font-variant-numeric:tabular-nums; }
        /* A coluna de avanco tem tres pecas na mesma linha (fracao, barra,
           percentual) e precisa de largura propria, senao a barra encolhe
           conforme o numero de selos da linha ao lado. */
        .med-tab td.med-av { width:150px; white-space:nowrap; }
        .med-frac { font-family:ui-monospace,Consolas,monospace; font-size:12px;
                    font-weight:700; color:var(--text-1); }
        .med-frac .med-de { font-weight:400; color:var(--text-3); }
        /* Os selos sao o que a linha tem de mais denso: quebram em duas
           fileiras em vez de esticar a tabela para fora da tela. */
        .med-tab td.med-selos { line-height:2; }
        .med-tab td.med-selos .gtbl-badge { margin:1px 2px 1px 0; }
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
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stElementContainer"],
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stLayoutWrapper"] { order: 3; }
        /* Cada bloco pelo conteudo, nao pela posicao no DOM -- ":first-child"
           deveria pegar a marca, mas na pratica quem vinha primeiro variava
           (chegou a ser o proprio filtro), e a marca ficava solta na ordem
           padrao, sem bater com o menu nem com o perfil. Marca sempre
           primeiro; filtro (botao + campos, quando abertos) cai no order:3
           generico logo acima; perfil sempre por ultimo -- direto pelo que
           cada um tem dentro, sem depender de qual o Streamlit desenhou
           primeiro.
           O perfil mora dentro de um st.container (pra dar posicionamento
           proprio ao botao invisivel) -- e isso faz o Streamlit envolve-lo
           num stLayoutWrapper, nao no stElementContainer que os widgets
           simples (marca, seletores) usam. Sem cobrir os dois tipos, o
           perfil nao batia com o order:3 generico NEM com o :has() de
           order:4, ficava sem nenhuma ordem definida e pulava pra frente de
           tudo -- inclusive da marca. */
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stElementContainer"]:has(.gplan-brand),
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stLayoutWrapper"]:has(.gplan-brand) {
          order: 1 !important; }
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stElementContainer"]:has(.sb-perfil),
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stLayoutWrapper"]:has(.sb-perfil),
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stElementContainer"]:has(.st-key-abrir_perfil),
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stLayoutWrapper"]:has(.st-key-abrir_perfil) {
          order: 4 !important; }
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
        /* Filtro recolhido por padrao, pra nao empurrar o perfil pra fora da
           tela: com os 6 campos sempre abertos, o menu + filtro passava de
           800px e obrigava rolagem dentro da lateral. Duas tentativas
           anteriores nao vingaram: um "cartao" com fundo/borda em cada campo
           ficou pesado sem resolver o espaco, e o st.expander nativo nao
           respeitava a mesma reordenacao por flex que marca e perfil
           respeitam (aparecia antes da marca, nao depois do menu). Um botao
           comum alternando um session_state, do mesmo tipo do "Abrir o
           perfil" logo abaixo, se comporta certo. */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] {
          border-top: 1px solid var(--border-color); padding-top: 8px; margin-top: 10px; }
        section[data-testid="stSidebar"] .st-key-flt_toggle button {
          background:var(--dark-card-2) !important; border-color:var(--border-color) !important;
          color:var(--text-2) !important; font-weight:600 !important;
          justify-content:flex-start !important; }
        section[data-testid="stSidebar"] .st-key-flt_toggle button p {
          font-size:12px !important; white-space:nowrap !important; overflow:hidden !important;
          text-overflow:ellipsis !important; }
        section[data-testid="stSidebar"] .stSelectbox label { font-size:10px !important;
          color:var(--text-3) !important; margin-bottom:0 !important;
          min-height:0 !important; line-height:1.25 !important; }
        section[data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.stSelectbox) {
          margin-bottom:-9px; }
        section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div {
          background:var(--dark-card) !important; border-color:var(--border-color) !important;
          font-size:11.5px !important; min-height:30px !important;
          transition:border-color 120ms; }
        section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"]:hover > div {
          border-color:rgba(var(--rgb-azul),.5) !important; }
        /* O botao de limpar e uma acao leve, nao um segundo botao do tamanho
           do menu -- fundo transparente e cor de aviso so aparecem no hover. */
        section[data-testid="stSidebar"] .stButton button { width:100%; font-size:11px !important;
          padding:4px 10px !important; border-radius:8px !important; }
        section[data-testid="stSidebar"] .st-key-gf_limpar button {
          background:transparent !important; border:1px dashed var(--border-color) !important;
          color:var(--text-3) !important; box-shadow:none !important; }
        section[data-testid="stSidebar"] .st-key-gf_limpar button:hover {
          color:var(--txt-vermelho) !important; border-color:rgba(var(--rgb-vermelho),.4) !important;
          background:rgba(var(--rgb-vermelho),.08) !important; }
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
        .fx-acao .ic .fxi, .fx-com .cab .ic .fxi, .fx-cab .marca .fxi,
        .pl-kpi .ic .fxi
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
        .pl-kpi .ic.fxc-azul  { background:rgba(var(--rgb-azul),.11); }
        .pl-kpi .ic.fxc-teal  { background:rgba(var(--rgb-teal),.11); }
        .pl-kpi .ic.fxc-roxo  { background:rgba(var(--rgb-roxo),.11); }
        .pl-kpi .ic.fxc-ambar { background:rgba(var(--rgb-ambar),.11); }
        .pl-kpi .ic.fxc-rubi  { background:rgba(var(--rgb-vermelho),.11); }
        .pl-kpi .ic.fxc-verde { background:rgba(var(--rgb-verde),.11); }
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
        .gpl-mark svg { width:100%; height:100%; display:block; }
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

        /* ---- Última atualização ----
           Um diário de obra, não um quadro de cartões: o dia manda na coluna da
           esquerda e cada linha começa pela TAG, que é o objeto deste controle.
           À direita, o estado de hoje nas três frentes -- documento, cabo,
           montagem -- para que ver o movimento e ver onde ele parou sejam o
           mesmo olhar. */
        .ua-log { display:flex; flex-direction:column; }
        .ua-dia { display:grid; grid-template-columns:74px 1fr; gap:16px;
          padding:14px 0; border-top:1px solid var(--border-color); }
        .ua-dia:first-child { border-top:0; padding-top:2px; }
        .ua-sel { position:sticky; top:0; align-self:start; text-align:right;
          padding-top:3px; }
        .ua-num { font-size:23px; font-weight:800; line-height:1; color:var(--text-1);
          font-variant-numeric:tabular-nums; }
        .ua-mes { font-size:9.5px; letter-spacing:1.2px; text-transform:uppercase;
          font-weight:800; color:var(--text-3); margin-top:5px; }
        .ua-qtd { font-size:10px; color:var(--text-3); margin-top:7px; }
        .ua-movs { display:flex; flex-direction:column; gap:8px; min-width:0; }
        .ua-mov { display:grid; grid-template-columns:1fr auto; gap:16px;
          align-items:start; padding:11px 14px; border-radius:11px;
          background:var(--dark-card-2); border:1px solid var(--border-color);
          border-left:3px solid var(--border-strong); }
        .ua-mov.ok { border-left-color:var(--accent-teal); }
        .ua-mov.warn { border-left-color:var(--accent-amber); }
        .ua-mov.crit { border-left-color:var(--accent-red); }
        .ua-mov.azul { border-left-color:var(--accent-blue); }
        .ua-mov.neutro { border-left-color:var(--border-strong); }
        .ua-cp { min-width:0; }
        .ua-tag { font-family:ui-monospace,Consolas,monospace; font-size:13px;
          font-weight:700; color:var(--text-1); text-decoration:none !important; }
        /* O Streamlit pinta <a> com a cor primária e ganha da regra de classe:
           no tema claro isso dava azul sobre branco a 2,63:1. */
        a.ua-tag, a.ua-tag:visited { color:var(--text-1) !important; }
        a.ua-tag:hover { color:var(--txt-azul) !important; }
        .ua-frente { font-size:9px; letter-spacing:.7px; text-transform:uppercase;
          font-weight:800; padding:2px 7px; border-radius:5px; margin-left:9px;
          vertical-align:1.5px; }
        .ua-frente.documento { color:var(--txt-azul); background:rgba(var(--rgb-azul),0.15); }
        .ua-frente.cabo { color:var(--txt-teal); background:rgba(var(--rgb-teal),0.15); }
        .ua-frente.montagem { color:var(--txt-roxo); background:rgba(var(--rgb-roxo),0.15); }
        .ua-frente.campo { color:var(--txt-ambar); background:rgba(var(--rgb-ambar),0.15); }
        .ua-frente.cadastro { color:var(--text-3); background:rgba(var(--rgb-chapa),0.15); }
        /* O trajeto: de onde saiu, em cinza, para onde foi, na cor do destino.
           É a frase inteira da movimentação numa linha. */
        .ua-de { color:var(--text-3); }
        .ua-seta { color:var(--text-3); padding:0 7px; }
        .ua-para { font-weight:700; color:var(--text-1); }
        .ua-para.ok { color:var(--txt-teal); }
        .ua-para.warn { color:var(--txt-ambar); }
        .ua-para.crit { color:var(--txt-vermelho); }
        .ua-para.azul { color:var(--txt-azul); }
        /* Quem se moveu junto: as 70 TAGs que foram para "Montado" de uma vez. */
        .ua-alvos { display:flex; flex-wrap:wrap; gap:5px; margin-top:8px; }
        .ua-alvos .ua-tag { font-size:11px; font-weight:600; padding:2px 7px;
          border-radius:5px; background:var(--dark-card); border:1px solid var(--border-color); }
        .ua-mais { font-size:10.5px; color:var(--text-3); align-self:center; padding-left:3px; }
        /* Quantas TAGs dependem daquele documento, colado no nome dele. */
        .ua-quantas { font-style:normal; font-size:9.5px; font-weight:800; margin-left:6px;
          padding:1px 5px; border-radius:4px; color:var(--txt-azul);
          background:rgba(var(--rgb-azul),0.16); }
        .ua-fato { font-size:12.5px; color:var(--text-2); margin-top:5px; line-height:1.5; }
        .ua-fato b { color:var(--text-1); font-weight:650;
          font-family:ui-monospace,Consolas,monospace; font-size:12px; }
        .ua-parecer { font-weight:700; }
        .ua-parecer.ok { color:var(--txt-teal); }
        .ua-parecer.warn { color:var(--txt-ambar); }
        .ua-parecer.crit { color:var(--txt-vermelho); }
        .ua-parecer.azul { color:var(--txt-azul); }
        .ua-onde { font-size:10.5px; color:var(--text-3); margin-top:4px; }
        .ua-dir { text-align:right; }
        .ua-quando { font-size:11px; color:var(--text-3); white-space:nowrap; }
        .ua-est { display:flex; gap:3px; justify-content:flex-end; margin-top:8px;
          cursor:help; }
        .ua-est i { width:20px; height:5px; border-radius:3px;
          background:var(--border-strong); }
        .ua-est i.ok { background:var(--accent-teal); }
        .ua-est i.meio { background:var(--accent-amber); }
        .ua-est i.nao { background:var(--accent-red); }
        .ua-rolo { max-height:640px; overflow-y:auto; padding-right:10px; }
        .ua-vazio { color:var(--text-3); font-size:12.5px; padding:20px 4px; }
        /* ---- Certificação ---- */
        .ct-leg { display:flex; gap:16px; flex-wrap:wrap; font-size:11px;
          color:var(--text-3); padding:10px 14px; background:var(--dark-card-2);
          border-radius:10px; margin-top:12px; }
        .ct-leg span { display:flex; align-items:center; gap:7px; }
        .ct-leg i { width:22px; height:3px; border-radius:2px; display:block; }
        .ct-leg b { width:11px; height:11px; border-radius:50%; display:block; }
        /* a TAG que esta no desenho fica marcada na tabela: sem isso, com 1.700
           linhas, nao da para saber qual delas o desenho esta mostrando */
        .ct-lin.sel td { background:rgba(var(--rgb-azul),0.12); }
        .ct-lin.sel td:first-child { box-shadow:inset 3px 0 0 var(--accent-blue); }
        /* rolagem propria, com o cabecalho preso: com 1.500 linhas, rolar sem
           ele e perder de vista qual coluna e qual */
        .ct-painel { margin-top:20px; padding:20px 22px 22px; }
        .ct-painel .gplan-panel-title { display:flex; align-items:baseline; gap:12px;
          margin-bottom:14px; }
        .ct-painel .gplan-panel-title span { margin-left:auto; font-size:11.5px; }
        .ct-rolo { max-height:440px; overflow-y:auto; border:1px solid var(--border-color);
          border-radius:12px; background:var(--dark-card-2); }
        .ct-rolo table { width:100%; }
        .ct-rolo td, .ct-rolo th { padding:9px 14px; }
        .ct-rolo table { margin:0; }
        .ct-rolo thead th { position:sticky; top:0; z-index:2;
          background:var(--dark-card-2); }
        /* O segmented control nasce com a paleta do config.toml, que e escura, e
           o texto dentro dele herda a cor do tema -- no claro dava escuro sobre
           escuro, 1,11:1. Mesma armadilha do campo de busca: o widget precisa
           ser vestido por testid, um por um. */
        /* O "?" de ajuda desenha com o tom do config.toml: no claro sobrava um
           traco quase branco sobre fundo claro, 1,03:1. */
        [data-testid="stTooltipHoverTarget"] { color: var(--text-3) !important; }
        [data-testid="stTooltipHoverTarget"] svg,
        [data-testid="stTooltipHoverTarget"] svg * { stroke: currentColor !important; }
        [data-testid="stTooltipHoverTarget"]:hover { color: var(--text-1) !important; }
        [data-testid="stButtonGroup"] button {
          background: var(--dark-card-2) !important;
          border: 1px solid var(--border-strong) !important;
        }
        [data-testid="stButtonGroup"] button,
        [data-testid="stButtonGroup"] button * { color: var(--text-2) !important; }
        [data-testid="stButtonGroup"] button:hover { border-color: var(--accent-blue) !important; }
        [data-testid="stButtonGroup"] button[aria-checked="true"] {
          background: rgba(var(--rgb-azul),0.14) !important;
          border-color: var(--accent-blue) !important;
        }
        [data-testid="stButtonGroup"] button[aria-checked="true"],
        [data-testid="stButtonGroup"] button[aria-checked="true"] * {
          color: var(--txt-azul) !important; font-weight: 650 !important;
        }
        .pl-kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:16px; }
        .pl-kpis.cinco { grid-template-columns:repeat(5,1fr); }
        .pl-kpis.cinco .pl-kpi .v { font-size:25px; }
        .pl-kpis.tres { grid-template-columns:repeat(3,1fr); }

        /* O perfil mora no rodape da lateral. O conteudo da sidebar vira uma
           coluna e o ultimo bloco e empurrado para baixo -- assim ele fica
           colado no pe da tela por mais curto que o menu seja, sem position
           fixed, que brigaria com a rolagem do proprio menu. */
        /* O perfil desce para o pe do menu. Quem faz o empurrao e o
           margin-top:auto no CONTAINER da linha, e nao na linha: o
           stSidebarUserContent e o stVerticalBlock sao display:contents, entao
           quem vira filho do flex da barra e o stElementContainer -- ou o
           stLayoutWrapper, que e o que o perfil usa por morar dentro de um
           st.container (ver render_perfil_lateral). Mirar so no
           stElementContainer nao empurrava nada pra ele. */
        [data-testid="stSidebarContent"] [data-testid="stElementContainer"]:has(.sb-perfil),
        [data-testid="stSidebarContent"] [data-testid="stLayoutWrapper"]:has(.sb-perfil) {
            margin-top:auto; }

        /* Uma linha so, sem cartao: foto, nome e e-mail, e a seta que diz que
           dali se entra. A altura e fixa porque o botao invisivel por cima
           precisa cobrir exatamente esta area. */
        /* A barra tem ~150 px uteis: cada pixel gasto na foto ou na seta sai
           do nome. Com 32 px de foto o nome ja saia cortado em "Daniel Bo...". */
        .sb-perfil { display:flex; align-items:center; gap:8px; height:64px;
                     padding:0 4px; margin-top:10px;
                     border-top:1px solid var(--border-color); }
        .sb-ini, .sb-foto { width:28px; height:28px; flex:none; border-radius:50%; }
        .sb-ini { display:grid; place-items:center; font-size:10.5px; font-weight:800;
                  color:var(--sobre-cor); background:var(--accent-teal); }
        .sb-foto { object-fit:cover; border:1px solid var(--border-strong); }
        .sb-txt { min-width:0; flex:1; display:flex; flex-direction:column;
                  line-height:1.28; }
        .sb-txt b { font-size:11.5px; font-weight:700; color:var(--text-1);
                    overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
        .sb-txt i { font-style:normal; font-size:9.5px; color:var(--text-3);
                    overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
        .sb-seta { flex:none; width:14px; height:14px; color:var(--text-3);
                   transition:transform .13s ease, color .13s ease; }
        .sb-seta svg { width:100%; height:100%; display:block; }
        .sb-perfil:hover .sb-seta { color:var(--text-1); transform:translateX(2px); }
        .sb-perfil:hover .sb-txt b { color:var(--accent-blue); }

        /* O botao real, transparente e cobrindo o container do perfil
           inteiro (position:absolute; inset:0) -- nao mais uma margem
           negativa fixa em pixel. Aquele calculo supunha a linha do perfil
           como vizinha imediata do botao no fluxo normal, e quebrava (cobria
           o fim do menu e o botao do filtro com um clique fantasma) assim
           que outra coisa passou a viver no mesmo bloco reordenado por
           flex. Com os dois dentro do MESMO st.container (key=perfil_wrap,
           position:relative, sem display:contents), o botao so precisa
           preencher esse container -- o tamanho dele quem da e a propria
           linha, que continua no fluxo normal. */
        .st-key-perfil_wrap { position:relative; }
        .st-key-perfil_wrap .st-key-abrir_perfil {
          position:absolute; inset:0; margin:0 !important; height:auto !important; z-index:2; }
        .st-key-perfil_wrap .st-key-abrir_perfil button {
            height:100%; width:100%; opacity:0; border:none; background:transparent;
            padding:0; cursor:pointer; }

        /* dialogo do perfil */
        .pf-topo { display:flex; align-items:center; gap:13px; margin-bottom:14px; }
        .pf-ini, .pf-foto { width:52px; height:52px; flex:none; border-radius:50%; }
        .pf-ini { display:grid; place-items:center; font-size:18px; font-weight:800;
                  color:var(--sobre-cor); background:var(--accent-teal); }
        .pf-foto { object-fit:cover; border:1px solid var(--border-strong); }
        .pf-nome { font-size:17px; font-weight:800; color:var(--text-1);
                   letter-spacing:-.3px; }
        .pf-login { font-size:12px; color:var(--text-3); margin-top:2px; }
        .pf-perms { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:16px; }
        .pf-tag { font-size:10.5px; font-weight:650; color:var(--txt-teal);
                  background:rgba(var(--rgb-teal),.12); border-radius:6px;
                  padding:3px 8px; }
        .pf-tag.vazio { color:var(--text-3); background:rgba(var(--rgb-tinta),.07); }

        /* Aba Acessos: o panorama em cima, os cartoes no meio, a edicao
           embaixo -- na ordem em que a pergunta aparece. */
        .ac-topo { display:flex; flex-wrap:wrap; gap:9px; margin:2px 0 16px; }
        .ac-kpi { display:flex; flex-direction:column; gap:1px; padding:9px 14px;
                  border-radius:11px; background:var(--dark-card-2);
                  border:1px solid var(--border-color); min-width:88px; }
        .ac-kpi .n { font-size:19px; font-weight:800; color:var(--text-1);
                     line-height:1.1; font-variant-numeric:tabular-nums; }
        .ac-kpi .r { font-size:9.5px; font-weight:700; letter-spacing:.5px;
                     text-transform:uppercase; color:var(--text-3); }
        .ac-kpi.roxo .n { color:var(--accent-purple); }
        .ac-kpi.teal .n { color:var(--accent-teal); }
        .ac-kpi.azul .n { color:var(--accent-blue); }
        .ac-kpi.ambar .n { color:var(--accent-amber); }
        .ac-kpi.vermelho .n { color:var(--accent-red); }

        .ac-card { background:var(--dark-card); border:1px solid var(--border-color);
                   border-radius:13px; padding:12px 14px; margin-bottom:9px; }
        .ac-card.off { opacity:.62; }
        .ac-cab { display:flex; align-items:center; gap:10px; }
        .ac-ini, .ac-foto { width:36px; height:36px; flex:none; border-radius:50%; }
        .ac-ini { display:grid; place-items:center; font-size:12.5px; font-weight:800;
                  color:var(--sobre-cor); background:var(--accent-teal); }
        .ac-foto { object-fit:cover; border:1px solid var(--border-strong); }
        .ac-id { min-width:0; flex:1; display:flex; flex-direction:column;
                 line-height:1.3; }
        .ac-id b { font-size:13px; font-weight:750; color:var(--text-1);
                   overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
        .ac-id i { font-style:normal; font-size:10.5px; color:var(--text-3);
                   overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
        .ac-eu { font-style:normal; font-size:9px; font-weight:700; margin-left:6px;
                 color:var(--txt-azul); background:rgba(var(--rgb-azul),.15);
                 border-radius:4px; padding:1px 5px; letter-spacing:.3px; }
        .ac-papel { flex:none; font-size:9px; font-weight:800; letter-spacing:.5px;
                    text-transform:uppercase; border-radius:5px; padding:2px 7px; }
        .ac-papel.roxo { color:var(--txt-roxo); background:rgba(var(--rgb-roxo),.15); }
        .ac-papel.teal { color:var(--txt-teal); background:rgba(var(--rgb-teal),.15); }
        .ac-papel.azul { color:var(--txt-azul); background:rgba(var(--rgb-azul),.15); }
        .ac-papel.ambar { color:var(--txt-ambar); background:rgba(var(--rgb-ambar),.15); }
        .ac-off { font-size:9.5px; font-weight:700; letter-spacing:.4px;
                  text-transform:uppercase; color:var(--accent-red);
                  background:rgba(var(--rgb-vermelho),.12); border-radius:5px;
                  padding:2px 7px; display:inline-block; }
        .ac-perms { display:flex; flex-wrap:wrap; gap:5px; margin-top:11px; }
        .ac-perms span { font-size:10px; color:var(--text-2);
                         background:rgba(var(--rgb-tinta),.06); border-radius:5px;
                         padding:2px 7px; }
        .ac-perms span.vazio { color:var(--text-3); font-style:italic; }
        /* o que o tipo escolhido libera, mostrado na hora de escolher */
        .ac-resumo { display:flex; flex-wrap:wrap; gap:5px; margin:-6px 0 10px; }
        .ac-resumo span { font-size:10px; color:var(--txt-teal);
                          background:rgba(var(--rgb-teal),.12); border-radius:5px;
                          padding:2px 8px; }

        /* o selo do papel: mesma familia de cor nos dois lugares onde aparece */
        .sb-papel, .pf-papel { font-size:9.5px; font-weight:800; letter-spacing:.5px;
                               text-transform:uppercase; border-radius:5px;
                               padding:2px 7px; white-space:nowrap; }
        /* na lateral o papel e a terceira linha, alinhado a esquerda e do
           tamanho do texto -- esticado ocuparia a linha inteira a toa */
        .sb-papel { font-style:normal; align-self:flex-start; margin-top:3px;
                    font-size:8.5px; padding:1px 5px; }
        .pf-papel { margin-left:9px; vertical-align:middle; }
        .sb-papel.roxo,  .pf-papel.roxo  { color:var(--txt-roxo);
            background:rgba(var(--rgb-roxo),.15); }
        .sb-papel.teal,  .pf-papel.teal  { color:var(--txt-teal);
            background:rgba(var(--rgb-teal),.15); }
        .sb-papel.azul,  .pf-papel.azul  { color:var(--txt-azul);
            background:rgba(var(--rgb-azul),.15); }
        .sb-papel.ambar, .pf-papel.ambar { color:var(--txt-ambar);
            background:rgba(var(--rgb-ambar),.15); }

        /* Avanco por segmento: uma barra por segmento H1, a pior em cima.
           Trilho de largura fixa e o metro ao lado -- com 142 segmentos a
           lista e um ranking, e o numero e quem da a escala. */
        .cs-lista { max-height:330px; overflow-y:auto; padding-right:4px; }
        .cs-lin { display:grid; grid-template-columns:104px 1fr 150px;
                  gap:12px; align-items:center; padding:5px 0; }
        .cs-lin + .cs-lin { border-top:1px solid var(--border-color); }
        .cs-seg { font-size:12px; font-weight:700; color:var(--text-2);
                  font-family:ui-monospace,'Cascadia Mono',Consolas,monospace; }
        .cs-trilho { height:14px; border-radius:4px; overflow:hidden;
                     background:rgba(var(--rgb-tinta),.07); }
        /* A cor vai no background, e nao no color: as classes feito/andando/
           parado do projeto pintam texto, e a barra herdava um color que
           nunca aparecia -- ela saia sem preenchimento nenhum. Mesmo par de
           regras da .pl-barra, que e a barra dos cartoes. */
        .cs-trilho i { display:block; height:100%; border-radius:4px;
                       background:var(--neutro); }
        .cs-trilho i.feito { background:var(--accent-teal); }
        .cs-trilho i.andando { background:var(--accent-amber); }
        .cs-trilho i.parado { background:var(--accent-red); }
        .cs-num { font-size:11px; color:var(--text-3); text-align:right;
                  font-variant-numeric:tabular-nums; white-space:nowrap; }
        .cs-num b { font-size:12.5px; font-weight:800; }
        .pl-kpi { background:var(--dark-card); border:1px solid var(--border-color);
                  border-radius:14px; padding:15px 18px; }
        /* O icone e opcional -- so a Planta usa, alinhado com o mapa de infra
           de instrumentacao (visual padronizado pra apresentacao). A
           Certificacao continua com o cartao sem .top, sem mudar nada nela. */
        .pl-kpi .top { display:flex; align-items:flex-start; justify-content:space-between;
                       gap:8px; }
        .pl-kpi .top .r { margin-top:2px; }
        .pl-kpi .ic { width:28px; height:28px; flex:none; border-radius:8px; padding:6px; }
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
        .pl-res .pl1 { color:#E8974E; }
        .pl-res .pl2 { color:#D4B106; }
        .pl-res .pl3 { color:#A8A020; }
        .pl-res .pl4 { color:#8BC34A; }
        .pl-res .pl5 { color:#6BAF3E; }
        .pl-res .pl6 { color:#4F9130; }
        .pl-res .pl7 { color:#2F7D32; }

        /* Cabo e avanço geral lado a lado: mesma pergunta, dois pesos
           diferentes. Cada .pl-pn dentro mantém seu próprio card -- só a
           largura é que divide ao meio. */
        .cs-duas { display:grid; grid-template-columns:1fr 1fr; gap:14px; align-items:stretch; }
        @media (max-width:900px) { .cs-duas { grid-template-columns:1fr; } }

        .pl-tela { position:relative; width:100%; height:0;
                   border:1px solid var(--border-color); border-radius:11px;
                   overflow:hidden; background:var(--fundo-3); }
        /* O desenho vem preto sobre branco. Invertido, o traco fica claro sobre
           o escuro e a prancha deixa de ser um retangulo branco no meio de uma
           tela escura -- e as zonas passam a ler por cima dele. image-rendering
           evita o desfoque do navegador ao redimensionar o desenho tecnico. */
        .pl-tela img { position:absolute; inset:0; width:100%; height:100%;
                       object-fit:fill; filter:var(--planta-filtro);
                       opacity:var(--planta-opacidade);
                       image-rendering:-webkit-optimize-contrast;
                       image-rendering:crisp-edges; }

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
        /* Escala de avanco da Planta: sete faixas, laranja a verde escuro --
           tabela que o Daniel definiu. So vale nesta aba: o resto do projeto
           (Certificacao, Progresso) continua no feito/andando/parado. */
        .pl-zona.pl1 { border-color:#E8974E; background:rgba(232,151,78,.34); color:#E8974E; }
        .pl-zona.pl2 { border-color:#D4B106; background:rgba(212,177,6,.34); color:#D4B106; }
        .pl-zona.pl3 { border-color:#A8A020; background:rgba(168,160,32,.34); color:#A8A020; }
        .pl-zona.pl4 { border-color:#8BC34A; background:rgba(139,195,74,.34); color:#8BC34A; }
        .pl-zona.pl5 { border-color:#6BAF3E; background:rgba(107,175,62,.34); color:#6BAF3E; }
        .pl-zona.pl6 { border-color:#4F9130; background:rgba(79,145,48,.34); color:#4F9130; }
        .pl-zona.pl7 { border-color:#2F7D32; background:rgba(47,125,50,.34); color:#2F7D32; }
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
        .pl-zona.pl1 .pl-ar { background:#E8974E; }
        .pl-zona.pl2 .pl-ar { background:#D4B106; }
        .pl-zona.pl3 .pl-ar { background:#A8A020; }
        .pl-zona.pl4 .pl-ar { background:#8BC34A; }
        .pl-zona.pl5 .pl-ar { background:#6BAF3E; }
        .pl-zona.pl6 .pl-ar { background:#4F9130; }
        .pl-zona.pl7 .pl-ar { background:#2F7D32; }

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
        /* Em linha, nao empilhada -- do jeito que ficou no mapa de infra de
           instrumentacao, pra apresentacao com visual padronizado. */
        .pl-ch-c { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr));
                   gap:9px; }
        .pl-ch { display:flex; align-items:center; gap:11px; background:var(--dark-card-2);
                 border:1px solid var(--border-color); border-radius:11px; padding:10px 12px; }
        .pl-ch .sw { width:14px; height:28px; border-radius:5px; border:1.5px solid currentColor;
                     flex:none; background:linear-gradient(180deg,transparent 45%,currentColor 45%); }
        .pl-ch.feito { color:var(--accent-teal); }
        .pl-ch.andando { color:var(--accent-amber); }
        .pl-ch.parado { color:var(--accent-red); }
        .pl-ch.pl1 { color:#E8974E; }
        .pl-ch.pl2 { color:#D4B106; }
        .pl-ch.pl3 { color:#A8A020; }
        .pl-ch.pl4 { color:#8BC34A; }
        .pl-ch.pl5 { color:#6BAF3E; }
        .pl-ch.pl6 { color:#4F9130; }
        .pl-ch.pl7 { color:#2F7D32; }
        .pl-ch .tx b { display:block; font-size:12px; font-weight:750; color:var(--text-1); line-height:1.3; }
        .pl-ch .tx em { font-style:normal; font-size:10.5px; color:var(--text-3); }
        .pl-ch .qt { margin-left:auto; text-align:right; }
        .pl-ch .qt b { display:block; font-size:16px; font-weight:800; line-height:1.2; }
        .pl-ch .qt em { font-style:normal; font-size:10px; color:var(--text-3); }
        @media (max-width:1250px) { .pl-kpis { grid-template-columns:repeat(2,1fr); } }
        @media (max-width:620px) { .pl-kpis { grid-template-columns:1fr; } }
        </style>
        """.replace("__ICONES__", fx_css_icones())
            .replace("__ICONES_SECAO__", secao_css_icones())
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
                       niveis_na_pagina: bool = False,
                       volta_por_tag: dict | None = None,
                       com_relatorio: bool = True) -> str:
    """Modais das tags visiveis na pagina, abertos/fechados via CSS :target.

    volta_por_tag diz para onde o X de cada TAG leva. Na Planta a TAG e aberta
    de dentro da ficha da planta, e fechar tem que devolver para ela -- fechar
    tudo obrigava a achar a planta de novo no desenho. Sem o mapa, o X fecha.
    """
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
                               niveis_na_pagina=niveis_na_pagina,
                               com_relatorio=com_relatorio)
        if corpo is None:
            continue
        volta = (volta_por_tag or {}).get(tag_id, "#fechado")
        rotulo = "Voltar" if volta != "#fechado" else "Fechar"
        blocos += f"""
            <div class="fmodal" id="{ficha_anchor(tag_id)}">
              <a class="fmodal-bg" href="{volta}" aria-label="{rotulo}"></a>
              <div class="fmodal-box">
                <div class="fmodal-head">
                  <div class="fmodal-title">{esc(tag_id)}</div>
                  <a class="fmodal-x" href="{volta}" aria-label="{rotulo}"
                     title="{rotulo}">&times;</a>
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
                     tags: pd.DataFrame, sigem: pd.DataFrame, cache_key: str = "",
                     origem: str = "fc") -> str:
    """Fichas das TAGs pedidas MAIS as dos relatorios e dos niveis que elas
    citam -- usada por Dashboard, Relatorios e Gitec, entao consertar aqui
    vale para as tres de uma vez.

    As tres familias andam juntas: a ficha da TAG mostra os relatorios dela
    (com um botao Detalhes cada) e a trilha de Fase/SOP/SSOP/Malha (com um
    degrau cada). Gerar so a ficha da TAG deixa esses cliques apontando para
    ancoras que nao existem -- aconteceu no Dashboard e em Relatorios, com 43
    dos 83 botoes de Detalhes sem destino, e a varredura completa achou o
    mesmo buraco no degrau de nivel, nestas mesmas tres abas. Chamar isto no
    lugar de fichas_modais_html direto impede que as paginas voltem a divergir.

    origem precisa ser diferente por chamador. fichas_niveis_html e cacheada
    pelo Streamlit e o parametro com o DataFrame (_df) nao entra no calculo do
    cache -- e a convencao do "_" na frente. As tres paginas chamando com o
    mesmo "fc" faziam a primeira que rodasse (o Dashboard) gravar o cache, e
    Relatorios e Gitec receberem de volta os niveis do Dashboard, nao os
    deles: a Malha MI-YST-121100 do Gitec nunca aparecia, porque quem
    respondia era sempre a mesma resposta guardada para "fc".
    """
    alvo = set(ids)
    meus = esperados[esperados["TAG"].isin(alvo)]
    docs = meus[meus["STATUS_SIGEM"].map(POSTADO)]["DOCUMENTO_ESPERADO"].tolist()
    historico = _revisoes_por_doc(cache_key, sigem)
    espera = espera_por_documento(historico)
    base_niveis = progresso_base(resumo, tags)
    base_niveis = base_niveis[base_niveis["TAG"].isin(alvo)]
    return (fichas_niveis_html(cache_key, f"{origem}:{assinatura_tags(alvo)}", "", "",
                               0, base_niveis, esperados)
            + fichas_modais_html(ids, resumo, esperados, tags, espera,
                                 niveis_na_pagina=True)
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


def du_kpi(rot: str, val: str, sub: str, pct: float, cor: str, icone: str, href: str = "",
          prioridade: str = "") -> str:
    # "prioridade" e o recorte so das TAGs prioritarias (SSOP_PRIORITARIO),
    # ao lado do "sub" na mesma linha -- a mesma pergunta do cartao, so que
    # restrita as prioritarias, sem abrir outra tela pra ver isso.
    linha_sub = (f'<div class="du-kpi-linha"><div class="sub">{sub}</div>'
                 f'<div class="du-kpi-prio">{prioridade}</div></div>'
                 if prioridade else f'<div class="sub">{sub}</div>')
    corpo = (f'<div class="topo">{du_tile(cor, icone)}<div class="rot">{rot}</div></div>'
             f'<div class="val">{val}</div>{linha_sub}'
             f'<div class="du-trilho {fx_classe_cor(cor)}">'
             f'<i style="width:{max(0.0, min(pct, 1.0)) * 100:.1f}%;"></i></div>')
    if href:
        return f'<a class="du-kpi" href="{com_filtros(href)}" target="_self">{corpo}</a>'
    return f'<div class="du-kpi">{corpo}</div>'


def du_grupos(resumo: pd.DataFrame, esperados: pd.DataFrame) -> str:
    # Um documento compartilhado pode cair em mais de um grupo (o RIMII/RIMSI
    # de infra por planta cobre tags de Instrumento, Valvula, Caixa, Painel e
    # Analisador juntas) -- ver totais_por_documento_agrupado.
    doc_g = totais_por_documento_agrupado(resumo, esperados, "GRUPO_REGRA")
    g = doc_g.rename(columns={"emitidos": "aprovados"}).reset_index(names="GRUPO_REGRA")
    tags_por_grupo = resumo.groupby("GRUPO_REGRA")["TAG"].size()
    g["tags"] = g["GRUPO_REGRA"].map(tags_por_grupo).fillna(0).astype(int)
    g["avanco"] = (g["aprovados"] / g["esperados"]).fillna(0) * 100
    g = g.sort_values("tags", ascending=False)
    # O cartao so vira link enquanto existir o filtro que consome o ?grupo=.
    # Sem o filtro na lateral o clique levava de volta ao Dashboard sem
    # recortar nada -- um link morto e pior que texto.
    clicavel = any(c == "grupo" for c, *_ in FILTROS)
    abre = ("a" if clicavel else "div")
    cartoes = "".join(
        f'<{abre} class="du-gp"'
        + (f' href="{com_filtros("/?grupo=" + quote(sentence_case(r["GRUPO_REGRA"])))}"'
           f' target="_self" title="Filtrar tudo por '
           f'{esc(sentence_case(r["GRUPO_REGRA"]))}"' if clicavel else "")
        + ">"
        f'<span class="lin"><span class="nm">{esc(sentence_case(r["GRUPO_REGRA"]))}</span>'
        f'<span class="pc">{br_pct(r["avanco"])}</span></span>'
        f'<span class="qt">{br_num(int(r["tags"]))}<em>tags</em></span>'
        f'<span class="du-trilho"><i style="width:{r["avanco"]:.1f}%;"></i></span>'
        f'<span class="pe"><span>{br_num(int(r["aprovados"]))} aprovados</span>'
        f'<span>{br_num(int(r["esperados"]))} esp.</span></span></{abre}>'
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
    # Um documento pode ser esperado por mais de uma TAG -- o RIMII/RIMSI de
    # infra por planta e o exemplo, ate 806 TAGs num so. A linha e por par
    # (TAG, documento); contar linha contava o mesmo parecer uma vez por TAG
    # penduradas nele. O documento e unico, entao so ele entra na conta.
    unicos = esperados.drop_duplicates(subset=["DOCUMENTO_ESPERADO"])
    contagem = unicos["STATUS_SIGEM"].value_counts()
    itens = [(sentence_case(k), int(v)) for k, v in contagem.items()]
    total = int(len(unicos)) or 1
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
            f'<div class="centro"><b>{br_num(total if itens else 0)}</b><span>relatórios</span></div></div>'
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


def du_mini(rot: str, val: str, sub: str, cor: str, icone: str, href: str = "",
           prioridade: str = "") -> str:
    classe = "val pq" if len(val) > 9 else "val"
    linha_sub = (f'<span class="du-kpi-linha"><span class="rot">{sub}</span>'
                 f'<span class="du-kpi-prio">{prioridade}</span></span>'
                 if prioridade else f'<span class="rot">{sub}</span>')
    corpo = (f'{du_tile(cor, icone)}<span><span class="rot">{rot}</span>'
             f'<span class="{classe}" style="display:block;">{val}</span>'
             f'{linha_sub}</span>')
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
            '<a class="fmodal-bg" href="#fechado" aria-label="Fechar"></a>'
            '<div class="fmodal-box"><div class="fmodal-head">'
            '<div class="fmodal-title">Recusados há mais tempo</div>'
            '<a class="fmodal-x" href="#fechado" aria-label="Fechar">&times;</a></div>'
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
    # Por documento unico, nao por par (TAG, documento) -- ver totais_por_
    # documento. Somar RELATORIOS_ESPERADOS/APROVADOS por TAG inflava esses
    # totais toda vez que um documento era compartilhado por varias TAGs.
    total_esperados, total_postados, total_aprovados, total_pendentes = (
        totais_por_documento(esperados))
    # pendente inclui o que foi postado mas nao passou: recusado, em analise e
    # cancelado voltam para a fila em vez de contarem como entregue
    avanco = (total_aprovados / total_esperados) if total_esperados else 0.0
    completas = int((resumo["AVANCO_DOCUMENTAL"] >= 1.0).sum())
    st_norm = esperados["STATUS_SIGEM"].astype(str).str.strip().str.upper()

    # Recorte so das TAGs prioritarias (SSOP_PRIORITARIO=SIM), pra mostrar ao
    # lado de cada cartao quanto daquele numero e delas -- a mesma pergunta
    # do cartao, sem abrir outra tela. Planilha sem a coluna (pipeline
    # antigo) degrada pra conjunto vazio, sem quebrar.
    prioritarias = (set(tags.loc[tags["SSOP_PRIORITARIO"].astype(str).str.strip().str.upper() == "SIM", "TAG"])
                    if "SSOP_PRIORITARIO" in tags.columns else set())
    resumo_prio = resumo[resumo["TAG"].isin(prioritarias)]
    total_tags_prio = len(resumo_prio)
    completas_prio = int((resumo_prio["AVANCO_DOCUMENTAL"] >= 1.0).sum())
    # Mesma logica de totais_por_documento: recorta esperados pelas TAGs
    # prioritarias e conta cada documento uma unica vez, senao um documento
    # compartilhado entre varias TAGs prioritarias (ou entre uma prioritaria
    # e outras que nao sao) inflava a fatia igual ao bug do total geral.
    if total_tags_prio:
        esperados_prio_df = esperados[esperados["TAG"].isin(prioritarias)]
        _, postados_prio, aprovados_prio, pendentes_prio = totais_por_documento(esperados_prio_df)
    else:
        postados_prio = aprovados_prio = pendentes_prio = 0
    # Nao e o avanco calculado so entre as prioritarias (outro denominador,
    # outra pergunta) -- e quantos PONTOS PERCENTUAIS do avanco geral vieram
    # de aprovacao de documento ligado a tag prioritaria. Mesmo denominador
    # (total_esperados) dos dois, entao os dois pedacos somam o total.
    pct_prio_do_avanco = (aprovados_prio / total_esperados * 100) if total_esperados else 0.0

    def prio(rotulo_valor: str) -> str:
        """"" quando nao ha nenhuma prioritaria no recorte -- o cartao fica
        limpo em vez de anunciar "0 prioritárias" em todo lugar."""
        return f"⚑ {rotulo_valor}" if total_tags_prio else ""

    kpis = (
        du_kpi("Total de tags", br_num(total_tags), "", 1.0,
               "#5b8def", "shield", "/pesquisa",
               prioridade=prio(f"{br_num(total_tags_prio)} prioritárias"))
        + du_kpi("Tags completas", br_num(completas),
                 f"{br_pct(completas / total_tags * 100)} do total",
                 completas / total_tags, "#34d399", "check",
                 prioridade=prio(f"{br_num(completas_prio)} prioritárias"))
        + du_kpi("Pendentes", br_num(total_pendentes),
                 f"{br_pct(total_pendentes / total_esperados * 100) if total_esperados else '—'} dos esperados",
                 (total_pendentes / total_esperados) if total_esperados else 0, "#f87171", "clock",
                 prioridade=prio(f"{br_num(pendentes_prio)} de prioritárias"))
        + du_kpi("Emitidos SIGEM", br_num(total_postados),
                 f"{br_pct(total_postados / total_esperados * 100) if total_esperados else '—'} dos esperados",
                 (total_postados / total_esperados) if total_esperados else 0, "#fbbf24", "archive",
                 prioridade=prio(f"{br_num(postados_prio)} de prioritárias"))
        + du_kpi("Avanço geral", br_pct(avanco * 100), f"{br_num(total_aprovados)} aprovados",
                 avanco, "#9d6bff", "trend",
                 prioridade=prio(f"{br_pct(pct_prio_do_avanco)} p.p. de prioritárias"))
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
    montados_prio = int((montagem.eq("MONTADO") & tags["TAG"].isin(prioritarias)).sum())

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
                "#5b8def", "shield",
                prioridade=prio(f"{br_num(montados_prio)} prioritárias"))
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
        f'<div class="du-col">{du_grupos(resumo, esperados)}{du_barras(esperados)}</div>'
        f'<div class="du-col">{du_status(esperados)}{du_top10(resumo)}</div>'
        "</section>"
        f'<section class="du-pe">{minis}</section>'
        "</div>"
        + du_modal_recusados(esperados, sigem)
        + fichas_completas(top["TAG"].tolist(), resumo, esperados, tags, sigem, cache_key,
                           origem="dash")
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

    # Esta faixa responde "quantos relatorios estao em cada status" -- e o
    # relatorio e o documento, nao a linha. Um RIMII/RIMSI de infra por planta
    # e uma linha por TAG (ate 806 pra um so), e contar linha contava o mesmo
    # parecer uma vez por TAG pendurada nele. A tabela de baixo continua por
    # linha de proposito -- ali a pergunta e outra, "o que falta para cada
    # TAG", e cada linha e a pendencia de uma TAG de verdade.
    doc_unico = df.drop_duplicates(subset=["DOCUMENTO_ESPERADO"])
    st_norm = doc_unico["STATUS_SIGEM"].astype(str).str.strip().str.upper()
    render_html(faixa_resumo([
        ("No recorte", br_num(len(doc_unico)), None),
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
                           sigem, cache_key, origem="rel")
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


# Icone do titulo de cada secao do menu lateral (Visao geral, Documentacao,
# Avanco, Administracao -- nessa ordem, ver secoes em main()). O atalho
# ":material/nome:" quebra a secao INTEIRA quando usado no titulo (a propria
# <section data-testid="stSidebar"> some do DOM -- confirmado ao vivo), e o
# titulo e so texto puro, sem HTML aceito, entao o icone so pode entrar por
# CSS -- daqui, via nth-of-type na mesma ordem das secoes.
SECAO_ICO = [
    '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/>'
    '<rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
    '<path d="M3 6a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    '<path d="M3 17l6-6 4 4 8-8"/><path d="M17 7h4v4"/>',
    '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
]


def secao_css_icones() -> str:
    regras = []
    for i, desenho in enumerate(SECAO_ICO, start=1):
        svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" '
               'stroke="black" stroke-width="1.9" stroke-linecap="round" '
               f'stroke-linejoin="round">{desenho}</svg>')
        uri = quote(svg.replace('"', "'"), safe="/:='<>() ")
        # nth-of-type no proprio <header> nao funciona: cada secao tem o seu
        # dentro de uma <div> propria, entao TODO header e "o primeiro" da
        # sua propria div-mae -- os 4 batiam sempre com :nth-of-type(1). Quem
        # e irmao de verdade sao essas divs, dentro de <ul data-testid=
        # "stSidebarNavItems">: e nelas que o nth-of-type conta certo.
        regras.append(
            f'[data-testid="stSidebarNavItems"] > div:nth-of-type({i}) '
            'header[data-testid="stNavSectionHeader"]::before'
            f'{{--fxi-secao:url("data:image/svg+xml,{uri}");}}')
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
                   niveis_na_pagina: bool = False,
                   com_relatorio: bool = True) -> str | None:
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
            f'<td class="gtbl-num">{botao_detalhes(doc, com_relatorio and POSTADO(stat))}</td></tr>'
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
    # SSOP_PRIORITARIO diz SE a tag pede atencao; SUBGRUPO_PRIORIDADE diz
    # QUANTO -- as duas vem da mesma base (01_BASE_TAGS) e sao perguntas
    # diferentes, entao o card mostra as duas linhas juntas.
    prioridade_html = ""
    ssop_prio, subgrupo_prio = da_base("SSOP_PRIORITARIO"), da_base("SUBGRUPO_PRIORIDADE")
    if ssop_prio != "—":
        prioridade_html += fx_linha("SSOP Prioritário", pill_ssop_prioritario(ssop_prio))
    if subgrupo_prio != "—":
        prioridade_html += fx_linha("Subgrupo de prioridade", pill_prioridade(subgrupo_prio))
    bloco_prioridade = (fx_painel("Prioridade", "alerta", prioridade_html)
                        if prioridade_html else "")
    # Card proprio, logo depois da situacao em campo -- que e onde o leitor
    # acabou de ver "On demand" e vai perguntar quando chega. So existe para
    # quem esta em compra: numa tag calibrada nao quer dizer nada.
    bloco_fornecimento = painel_previsao_fornecimento(
        da_base("STATUS_FINAL"),
        t.get("PREVISAO_FORNECIMENTO") if hasattr(t, "get") else None,
        t.get("STATUS_FORNECIMENTO") if hasattr(t, "get") else None)

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
        + f'</div><div class="fx-col">{avanco}{bloco_campo}{bloco_prioridade}{bloco_fornecimento}{mov}{acoes}</div></div></div>'
    )


def _pesq_faixa_nada():
    """Callback vazio: so existe pra habilitar default={"faixa": ...} no
    componente -- a API exige um on_<evento>_change registrado pra cada
    chave do default. O valor de verdade sai de resultado.faixa, lido direto
    depois de montar; nao precisa reagir a nada aqui."""


# Componente bidirecional de verdade (st.components.v2), nao mais o iframe
# com truque de alcancar o number_input escondido no documento pai: aquilo
# so atualizava o valor visual, o Streamlit nunca recebia o "input" sintetico
# como disparo de on_change (aparenta exigir evento confiavel/genuino, que
# JS nao produz). Aqui a bolha e o proprio DOM do componente (dentro de uma
# shadow root, sem fronteira de iframe) e setStateValue() e o canal oficial
# de volta pro Python -- testado isoladamente antes de entrar aqui.
#
# Registrado uma unica vez, no nivel do modulo: registrar de novo a cada
# render (dentro de render_pesquisa_tag) e o que o proprio Streamlit despreza
# ("evite juntar definicao e montagem").
PESQ_FAIXA_JS = r"""
export default function (component) {
  const { data, setStateValue, parentElement } = component;
  parentElement.innerHTML = `
    <style>
      *{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
      .faixa{position:relative;height:48px;
        font:400 12px ui-sans-serif,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
        font-variant-numeric:tabular-nums}
      .trilho{position:absolute;top:32px;left:16px;right:16px;height:7px;border-radius:999px;
        background:linear-gradient(90deg,${data.nao},${data.and} 50%,${data.ok});
        box-shadow:inset 0 1px 3px rgba(0,0,0,.22)}
      .alca{position:absolute;top:0;height:26px;transform:translateX(-50%);
        display:flex;align-items:flex-end;justify-content:center;
        cursor:grab;touch-action:none;user-select:none;z-index:1}
      .alca.ativa{z-index:2}
      .alca:active{cursor:grabbing}
      .bolha{background:${data.card};color:${data.t1};border:2px solid ${data.ok};
        border-radius:999px;padding:2px 0;font:inherit;font-weight:700;font-size:12px;
        width:32px;text-align:center;box-shadow:0 3px 10px ${data.sombra};outline:none;
        transition:transform .15s ease,box-shadow .15s ease;cursor:text}
      .alca.ativa .bolha,.bolha:focus{transform:scale(1.14);
        box-shadow:0 0 0 3px ${data.okA3},0 4px 14px ${data.okA45}}
    </style>
    <div class="faixa" id="faixa">
      <div class="trilho" id="trilho"></div>
      <div class="alca" id="hMin"><input class="bolha" id="bMin" inputmode="numeric" maxlength="3"></div>
      <div class="alca" id="hMax"><input class="bolha" id="bMax" inputmode="numeric" maxlength="3"></div>
    </div>`;

  const faixa = parentElement.querySelector('#faixa');
  const trilho = parentElement.querySelector('#trilho');
  const hMin = parentElement.querySelector('#hMin'), hMax = parentElement.querySelector('#hMax');
  const bMin = parentElement.querySelector('#bMin'), bMax = parentElement.querySelector('#bMax');
  let vMin = data.min, vMax = data.max;

  function coloca() {
    const r = trilho.getBoundingClientRect(), f = faixa.getBoundingClientRect();
    const largura = Math.max(0, r.width);
    const base = r.left - f.left;
    hMin.style.left = (base + (vMin / 100) * largura) + 'px';
    hMax.style.left = (base + (vMax / 100) * largura) + 'px';
    bMin.value = vMin;
    bMax.value = vMax;
  }
  coloca();
  window.addEventListener('resize', coloca);

  function commit() { setStateValue('faixa', { min: vMin, max: vMax }); }

  function valorDoPonteiro(clientX) {
    const r = trilho.getBoundingClientRect();
    const largura = Math.max(1, r.width);
    const pct = (clientX - r.left) / largura;
    return Math.round(Math.max(0, Math.min(1, pct)) * 100);
  }

  function arrasta(alvo) {
    const alca = alvo === 'min' ? hMin : hMax;
    alca.classList.add('ativa');
    function mover(e) {
      const x = e.touches ? e.touches[0].clientX : e.clientX;
      const v = valorDoPonteiro(x);
      if (alvo === 'min') vMin = Math.min(v, vMax); else vMax = Math.max(v, vMin);
      coloca();
    }
    function soltar() {
      alca.classList.remove('ativa');
      document.removeEventListener('pointermove', mover);
      document.removeEventListener('pointerup', soltar);
      commit();
    }
    document.addEventListener('pointermove', mover);
    document.addEventListener('pointerup', soltar);
  }
  hMin.addEventListener('pointerdown', e => {
    if (e.target === bMin) return; e.preventDefault(); arrasta('min'); });
  hMax.addEventListener('pointerdown', e => {
    if (e.target === bMax) return; e.preventDefault(); arrasta('max'); });

  function ligaBolha(input, alvo) {
    input.addEventListener('focus', () => input.select());
    input.addEventListener('input', () => {
      input.value = input.value.replace(/[^0-9]/g, '').slice(0, 3); });
    // Aplica direto no keydown do Enter (nao so via blur): dentro de uma
    // shadow root sem delegatesFocus, .blur() programatico as vezes nao
    // dispara o evento 'blur' de verdade, entao depender so dele deixava o
    // Enter sem efeito. O blur continua de reserva pra quando o usuario
    // clica fora sem apertar Enter.
    function aplicar() {
      let v = parseInt(input.value, 10);
      if (isNaN(v)) v = alvo === 'min' ? vMin : vMax;
      v = Math.max(0, Math.min(100, v));
      if (alvo === 'min') vMin = Math.min(v, vMax); else vMax = Math.max(v, vMin);
      coloca();
      commit();
    }
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); aplicar(); input.blur(); } });
    input.addEventListener('blur', aplicar);
  }
  ligaBolha(bMin, 'min');
  ligaBolha(bMax, 'max');
}
"""
PESQ_FAIXA_COMP = st.components.v2.component("pesq_faixa", js=PESQ_FAIXA_JS)


def render_pesquisa_tag(resumo: pd.DataFrame, esperados: pd.DataFrame, tags: pd.DataFrame,
                        sigem: pd.DataFrame, cache_key: str = "",
                        lancamento: pd.DataFrame | None = None,
                        depara: pd.DataFrame | None = None):
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

    # O painel sai da planilha de cabos, e nem toda planilha a tem: sem ela o
    # campo não aparece, em vez de aparecer vazio sem explicação.
    painel_de = (cert_painel_por_tag(lancamento, depara, cache_key)
                 if lancamento is not None and not lancamento.empty
                 and depara is not None else {})
    if painel_de:
        col_busca, col_painel = st.columns([3, 2], gap="medium")
    else:
        col_busca, col_painel = st.container(), None
    with col_busca:
        search = lembrado(st.text_input, "pesq_busca", "Pesquisar", placeholder="Digite a tag para ver a ficha completa (ex: AIT-120005)...", label_visibility="collapsed")

    # Duas pontas: uma so ponta (min ou max fixo) responderia "acima de X" ou
    # "abaixo de X" -- as duas juntas fecham uma faixa fechada, entao 90-100
    # mostra so quem esta nesse intervalo, nao tudo acima de 90.
    #
    # mem_ sobrevive a troca de aba (mesmo raciocinio do lembrado()); dentro
    # da MESMA aba quem manda e o proprio componente, via key= estavel.
    lo0 = st.session_state.get("mem_pesq_avanco_min", 0)
    hi0 = st.session_state.get("mem_pesq_avanco_max", 100)
    t = TEMAS.get(tema_ativo(), TEMAS[TEMA_PADRAO])
    resultado_faixa = PESQ_FAIXA_COMP(
        data={"min": lo0, "max": hi0,
             "nao": t["vermelho"], "and": t["ambar"], "ok": t["teal"],
             "card": t["card"], "t1": t["texto1"], "sombra": t["sombra"],
             "okA3": f"rgba({t['rgb_teal']},.3)", "okA45": f"rgba({t['rgb_teal']},.45)"},
        default={"faixa": {"min": lo0, "max": hi0}},
        key="pesq_faixa", height=48,
        on_faixa_change=_pesq_faixa_nada)
    faixa_bruta = resultado_faixa.get("faixa") or {"min": lo0, "max": hi0}
    faixa_avanco = (int(faixa_bruta["min"]), int(faixa_bruta["max"]))
    st.session_state["mem_pesq_avanco_min"] = faixa_avanco[0]
    st.session_state["mem_pesq_avanco_max"] = faixa_avanco[1]

    list_df = resumo[["TAG", "DESCRICAO", "GRUPO_REGRA", "ITEM_PPU", "RELATORIOS_ESPERADOS", "AVANCO_DOCUMENTAL"]].copy()
    if painel_de:
        # Só entra no filtro o painel que tem segmentação: instrumento pendurado
        # em caixa de junção. Painel de Elétrica que alimenta um circuito direto
        # não é segmento -- e eram dezenas deles enchendo a lista com uma TAG.
        do_painel: dict[str, list] = {}
        for t in resumo["TAG"].astype(str):
            pa = painel_de.get(t)
            if pa:
                do_painel.setdefault(pa, []).append(t)
        # Segmentação é painel → caixa → instrumento: os três. Só a caixa não
        # basta -- o PN-12-220(CIU) tem uma, e nada chega nela.
        do_painel = {pa: v for pa, v in do_painel.items()
                     if any(t.startswith(CAIXA_PREFIXOS) for t in v)
                     and any(not t.startswith(CAIXA_PREFIXOS) for t in v)}
        with col_painel:
            escolha_painel = st.selectbox(
                "Painel", ["Todos os painéis"]
                + [f"{p}  ·  {len(v)} tags" for p, v in sorted(do_painel.items())],
                key="pesq_painel", label_visibility="collapsed",
                help="O painel vem da planilha de cabos: mostra o segmento inteiro "
                     "que depende dele.")
        if escolha_painel != "Todos os painéis":
            alvo_p = escolha_painel.split("  ·  ")[0]
            list_df = list_df[list_df["TAG"].isin(do_painel.get(alvo_p, []))]
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
    if faixa_avanco != (0, 100):
        list_df = list_df[list_df["AVANCO_DOCUMENTAL"].between(faixa_avanco[0], faixa_avanco[1])]
    list_df_page = paginate(list_df, "pesquisa", f"{search}|{faixa_avanco}")

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
        + fichas_completas(pagina["TAG"].tolist(), resumo, esperados, tags, sigem, cache_key,
                           origem="gitec")
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


# A escala de sete degraus e so da Planta -- laranja a verde escuro, do jeito
# que o Daniel definiu. O resto do projeto (Certificacao, Progresso) continua
# no feito/andando/parado de classe_avanco: sao paginas diferentes, com
# pergunta diferente ("terminou ou nao" vs "quanto do mapa ja avancou").
PLANTA_FAIXAS_PCT = [20, 40, 60, 70, 80, 90]


def classe_avanco_planta(pct: float) -> str:
    for i, teto in enumerate(PLANTA_FAIXAS_PCT, start=1):
        if pct < teto:
            return f"pl{i}"
    return "pl7"


def dados_por_area(tags: pd.DataFrame, resumo: pd.DataFrame,
                   locacao: pd.DataFrame,
                   aux: pd.DataFrame) -> tuple[dict, dict, pd.DataFrame]:
    """Quanto cada area e cada planta ja montaram.

    A coluna LOCACAO da 05_BASE_LOCACAO diz em QUE DESENHO a TAG esta -- sao
    2.113 instrumentos apontando para 20 plantas. Enquanto isso passou
    despercebido, o numero mostrado era o da AREA, e as ate nove plantas de
    uma mesma area repetiam o mesmo previsto e o mesmo executado, como se
    fossem uma so.

    Agora cada planta tem o proprio numero, e a area continua existindo para
    as TAGs sem LOCACAO preenchida e para a ficha da area.

    Os desenhos marcados como gerais (CHZ-113 e CHZ-302, que atravessam a
    unidade) ficam de fora do mapa desenho->area: eles nao delimitam zona.
    """
    if locacao.empty or "AREA" not in locacao.columns:
        return {}, {}, pd.DataFrame()

    area_da_tag = (locacao.dropna(subset=["AREA"])
                   .assign(AREA=lambda d: d["AREA"].astype(str).str.strip())
                   .set_index("TAG")["AREA"].to_dict())
    # LOCACAO ja vem no formato do desenho ("800-CHZ-322"), o mesmo da
    # 05_AUX_AREAS -- da para cruzar direto, sem normalizar nada
    desenho_da_tag = {}
    if "LOCACAO" in locacao.columns:
        for tag, d in zip(locacao["TAG"], locacao["LOCACAO"]):
            texto = str(d).strip()
            if texto and texto not in ("-", "nan"):
                desenho_da_tag[tag] = texto

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

    t["_desenho"] = t["TAG"].map(desenho_da_tag)

    def resumir(sub, rotulo, nome_longo):
        qtd = len(sub)
        mont = int(sub["_montado"].sum())
        # O que ja foi montado se divide em duas partes que somam o todo: o que
        # o GITEC ja aprovou e o que ainda nao virou medicao. Estar no GITEC nao
        # basta -- so o aprovado conta, que e o que MEDIDO_GITEC ja diz.
        montado = sub["_montado"]
        medido = montado & (sub["_medido"] == "SIM")
        return {
            "area": rotulo, "nome": nome_longo, "tags": qtd, "montados": mont,
            "pct": round(mont / qtd * 100, 1) if qtd else 0.0,
            "doc": round(sub["_docfrac"].mean() * 100, 1),
            "valor": float(sub["_preco"].sum()),
            "valor_montado": float(sub.loc[montado, "_preco"].sum()),
            "valor_medido": float(sub.loc[medido, "_preco"].sum()),
            "montados_medidos": int(medido.sum()),
        }

    areas = {a: resumir(sub, a, nome.get(a, "—"))
             for a, sub in t.groupby("_area")}
    # a planta herda o nome da area a que pertence: e o rotulo que a ficha usa
    plantas = {d: resumir(sub, d, nome.get(area_do_desenho.get(d, ""), "—"))
               for d, sub in t.groupby("_desenho")}
    return areas, plantas, area_do_desenho, t


def dados_da_zona(zona: dict, areas: dict, plantas: dict,
                  area_do_desenho: dict) -> dict | None:
    """O numero que a zona mostra.

    Vem das plantas da zona, somadas. So cai para o numero da area quando
    nenhuma planta dela tem TAG com LOCACAO preenchida -- e nesse caso as
    zonas irmas repetem o mesmo valor, porque e o unico que existe.
    """
    proprias = [plantas[d] for d in zona.get("desenhos", []) if d in plantas]
    a = zona_area(zona, area_do_desenho)
    if proprias:
        somado = {c: sum(p[c] for p in proprias)
                  for c in ("tags", "montados", "montados_medidos",
                            "valor", "valor_montado", "valor_medido")}
        qtd = somado["tags"]
        return {**proprias[0], **somado,
                # o "area" de proprias[0] e o desenho (plantas[d] usa a chave
                # "area" para o proprio nome) -- sem isto a etiqueta da zona e
                # o link "abrir ficha da area" mostravam "800-CHZ-327" em vez
                # do codigo real, e o link nem batia com ficha nenhuma.
                "area": a or proprias[0]["area"],
                "pct": round(somado["montados"] / qtd * 100, 1) if qtd else 0.0,
                # media ponderada pelo tamanho de cada planta, e nao media das
                # medias: planta de 3 TAGs pesaria igual a de 300
                "doc": round(sum(p["doc"] * p["tags"] for p in proprias) / qtd, 1)
                if qtd else 0.0,
                "por_planta": True}
    dados = areas.get(a or "")
    return {**dados, "por_planta": False} if dados else None


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


def planta_zonas_html(prancha: dict, areas: dict, plantas: dict,
                      area_do_desenho: dict) -> str:
    largura = PLANTA_LARGURA.get(prancha["id"], 1500)
    altura = largura * prancha["prop"] / 100
    partes = []
    for z in prancha["zonas"]:
        a = dados_da_zona(z, areas, plantas, area_do_desenho)
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
            f'<a class="pl-zona {classe_avanco_planta(a["pct"])}'
            f'{" rec" if z.get("recorte") else ""}" style="{estilo}"'
            f' href="#{_ancora("PLANTA", rotulo)}" title="{esc(titulo)}">{etiqueta}'
            f'<span class="pl-mio"><b>{esc(rotulo)}</b>'
            f'<i>{br_pct(a["pct"])}</i>{extra}</span></a>')
    return "".join(partes)


def planta_prancha_html(prancha: dict, areas: dict, plantas: dict,
                        area_do_desenho: dict) -> str:
    # O total da prancha soma o que cada zona mostra. Zona que caiu no numero
    # da area entra uma vez so: duas zonas irmas sem planta propria repetem o
    # mesmo numero, e soma-lo duas vezes inventaria instrumento.
    qtd = mont = 0
    areas_ja_contadas = set()
    for z in prancha["zonas"]:
        d = dados_da_zona(z, areas, plantas, area_do_desenho)
        if not d:
            continue
        if not d["por_planta"]:
            if d["area"] in areas_ja_contadas:
                continue
            areas_ja_contadas.add(d["area"])
        qtd += d["tags"]
        mont += d["montados"]
    vistas = {a for a in (zona_area(z, area_do_desenho) for z in prancha["zonas"])
              if a and a in areas}
    pct = mont / qtd * 100 if qtd else 0.0
    return (
        '<div class="gplan-panel pl-pn">'
        f'<div class="gplan-panel-title">{esc(prancha["rotulo"])}'
        f'<span class="pl-res">{len(vistas)} área{"s" if len(vistas) > 1 else ""}'
        f' · {br_num(qtd)} instrumentos · '
        f'<b class="{classe_avanco_planta(pct)}">{br_pct(pct)}</b></span></div>'
        f'<div class="pl-tela" style="padding-top:{prancha["prop"]:.3f}%">'
        f'<img src="{prancha["uri"]}" alt="Planta — {esc(prancha["rotulo"])}">'
        + planta_zonas_html(prancha, areas, plantas, area_do_desenho)
        + "</div></div>")




# ===================================================================== #
# Certificação: a cadeia física de um instrumento                        #
# ===================================================================== #
# Certificar uma TAG é assinar que tudo antes dela está pronto. O que vem
# antes é físico e está em duas bases: o circuito, na 02_CABOS_LANCAMENTO, e
# a montagem, na 01_BASE_TAGS. A aba junta as duas e responde uma pergunta
# só -- esta TAG pode ser certificada, e se não, onde para.

PAINEL_PREFIXOS = ("PN", "PL", "PCC")
CAIXA_PREFIXOS = ("CJ", "CFF")


def cert_metro_real(linha) -> float:
    """O metro lançado: o que o campo mediu, ou o proporcional do total.

    A coluna COMPR.(M) CAMPO REALIZADO só existe na base nova. Sem ela -- e nas
    linhas em que ela está vazia -- o proporcional do percentual é a melhor
    conta que há. Zerar seria dizer que nada foi lançado, o que é falso.

    O medido é capado no previsto. Em 1.130 dos 4.375 circuitos o campo mediu
    mais metro do que a estimativa previa, e sem teto o avanço passava de
    100% -- 300% no pior segmento. Avanço é quanto do previsto já foi feito; o
    metro que sobra é erro de estimativa, não obra a mais. Circuito sem
    previsto não tem teto a aplicar: ali o medido é tudo o que se sabe.
    """
    previsto = cert_num(linha.get("METROS"))
    real = cert_num(linha.get("METROS_REAL"))
    if real > 0:
        return min(real, previsto) if previsto > 0 else real
    return round(previsto * cert_num(linha.get("PCT")) / 100, 1)


# O que o pipeline escreve em DOC_REF nas ligações que ELE acrescentou para
# fechar a topologia do desenho (_aplicar_correcoes_conhecidas). É o único
# marcador confiável: duas dessas linhas reaproveitam de propósito o nome de
# circuito do próprio desenho (C-YST-121166, C-YST-121188), então olhar o
# nome não separa remendo de circuito de verdade.
CERT_MARCA_REMENDO = "ausente na base de cabos"

# Largura real da moldura do desenho nesta aba, medida no navegador. O SVG
# escala pela largura, então é ela que decide a altura útil: com uma
# estimativa maior que a real, a moldura ficava mais alta que o desenho e
# sobrava um vão escuro embaixo.
CERT_MOLDURA_PX = 970


def cert_so_potencia(linhas_do_no: list, destino: str) -> bool:
    """Se o único cabo entre este nó e esse destino é de potência.

    O instrumento costuma ter dois cabos: o de sinal, que define a qual
    painel ele pertence, e o de potência (-P), que vai ao painel de
    alimentação. Só o primeiro responde "de quem é este instrumento".
    """
    irmaos = [x for x in linhas_do_no if str(x["DESTINO"]).strip() == destino]
    return bool(irmaos) and all(
        re.search(r"-P\d*$", str(x["CIRCUITO"]).strip(), re.I) for x in irmaos)


def cert_metro_medido(linha) -> float:
    """O metro que o campo mediu, SEM teto no previsto.

    O cert_metro_real capa no previsto porque o avanço geral é "quanto do
    previsto já foi feito" -- ali passar de 100% não quer dizer nada. Mas na
    leitura de um cabo específico o metro a mais é fato: a estimativa é que
    estava curta, e esconder isso apaga metade do que há para ver.
    """
    real = cert_num(linha.get("METROS_REAL"))
    if real > 0:
        return real
    return round(cert_num(linha.get("METROS")) * cert_num(linha.get("PCT")) / 100, 1)


def cert_status_conjunto(circuitos: list) -> str:
    """O status de um conjunto de circuitos, tirado do que a base informa.

    Status e metragem são fatos independentes, e os dois importam:

    * um cabo pode estar **Concluído com menos metro que o previsto** -- foi
      previsto mais do que a obra precisou, e isso não o torna inacabado;
    * um cabo pode ter **mais metro que o previsto e seguir Em Andamento** --
      a estimativa é que estava curta, e o serviço não acabou.

    Deduzir o status do percentual de metragem (o que esta aba fazia) apagava
    os dois casos: o primeiro virava "Em Andamento" por engano, e o segundo
    virava "Concluído" antes da hora. Aqui o status vem da coluna STATUS, e a
    metragem é mostrada ao lado, sem um decidir pelo outro.

    Entre vários circuitos vale o mais atrasado: a TAG não está pronta porque
    um dos cabos dela ficou pronto.
    """
    if not circuitos:
        return "sem circuito"
    vistos = [str(c["STATUS"]).strip() for c in circuitos]
    for pior in ("Não Iniciado", "Em Andamento"):
        if pior in vistos:
            return pior
    return vistos[0] or "sem circuito"


def cert_remendo(linha) -> bool:
    """Se a linha é uma ligação que nós acrescentamos, não um cabo medido.

    Ela existe só para o grafo fechar onde a planilha ainda não tem o
    circuito, e nasce sempre "Não Iniciado / 0 m". Contada como cabo pendente
    de verdade, ela reprova o laço inteiro -- era o que fazia 18 TAGs com o
    cabo real 100% lançado aparecerem como "Bloqueado".
    """
    return CERT_MARCA_REMENDO in str(linha.get("DOC_REF") or "")


def cert_txt(serie: pd.Series) -> pd.Series:
    """Coluna de texto sem nulo, em qualquer versão do pandas.

    Até o pandas 2, `astype(str)` transformava NaN na string "nan". No pandas 3
    o nulo sobrevive à conversão, e a linha seguinte quebra com
    "'float' object has no attribute 'startswith'" -- que foi exatamente o que
    derrubou esta aba em produção, rodando num Python mais novo que o daqui.
    O fillna antes tira a diferença entre as duas versões.
    """
    return serie.fillna("").astype(str).str.strip()


def cert_nivel(ponta: str) -> int:
    """0 campo, 1 caixa, 2 painel. Decide o sentido de leitura do trecho."""
    p = str(ponta).upper()
    if p.startswith(PAINEL_PREFIXOS):
        return 2
    if p.startswith(CAIXA_PREFIXOS):
        return 1
    return 0


@st.cache_data(show_spinner=False, max_entries=3)
def cert_montagem(tags: pd.DataFrame, depara: pd.DataFrame, cache_key: str) -> dict:
    """Status de montagem por ponta da planilha de cabos.

    Passa pelo de-para porque as duas bases escrevem a mesma TAG de formas
    diferentes -- ZSH/L-120001 aqui, ZSH-120001 e ZSL-120001 lá. Quando uma
    ponta cobre duas TAGs vale a pior: não dá para dar por montada uma ponta
    em que só metade do par está.
    """
    if tags.empty or depara.empty:
        return {}
    por_tag = dict(zip(cert_txt(tags["TAG"]),
                       cert_txt(tags.get("STATUS_MONTAGEM", pd.Series(dtype=str)))))
    ordem = ["Montado", "Em Programação", "Não Programado", "Não Montado"]
    saida: dict[str, dict] = {}
    for ponta, grupo in depara.groupby(cert_txt(depara["PONTA"])):
        if not ponta or ponta == "nan":
            continue
        alvos = [t for t in cert_txt(grupo["TAG"]) if t and t != "nan"]
        estados = [por_tag.get(t, "") for t in alvos]
        pior = sorted(estados, key=lambda x: ordem.index(x) if x in ordem else 9)
        saida[ponta] = {"mont": pior[-1] if pior else "",
                        "como": str(grupo["COMO"].iloc[0]),
                        "tags": alvos}
    return saida


def cert_num(valor) -> float:
    """Número de célula, sem nulo. `valor or 0` não serve: NaN é verdadeiro, e
    passaria adiante -- viraria NaN no JSON do desenho, que o navegador recusa."""
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if v != v else v


def _cert_circuito(linha, mont: dict) -> dict:
    ponta = str(linha["ORIGEM"]).strip()
    m = mont.get(ponta, {})
    return {"id": str(linha["CIRCUITO"]).strip(), "disc": str(linha["DISCIPLINA"]).strip(),
            "status": str(linha["STATUS"]).strip(),
            "pct": round(cert_num(linha["PCT"]), 1),
            "m": cert_num(linha["METROS"]),
            "m_real": cert_metro_real(linha),
            # o circuito de potência traz -P no fim do código; a coluna TIPO não
            # separa os dois -- ela diz o sistema, não a função do cabo
            "pot": bool(re.search(r"-P\d*$", str(linha["CIRCUITO"]).strip(), re.I)),
            "org": ponta,
            "dst": str(linha["DESTINO"]).strip(),
            "mont": m.get("mont", ""), "como": m.get("como", ""), "tags": m.get("tags", [])}


@st.cache_data(show_spinner=False, max_entries=3)
def cert_alvos(lanc: pd.DataFrame, cache_key: str) -> list:
    """O que dá para abrir: cada caixa de junção e cada tronco de fieldbus.

    O instrumento não é alvo -- ele é folha. Escolher um instrumento abre a
    cadeia da caixa dele, que é onde a resposta mora.
    """
    if lanc.empty:
        return []
    pontas = cert_txt(pd.concat([lanc["ORIGEM"], lanc["DESTINO"]]))
    caixas = {p for p in pontas if p.startswith("CJ")}
    troncos = {re.sub(r"[A-Z]$", "", p) for p in pontas if p.startswith("CFF")}
    return sorted(caixas | troncos)


@st.cache_data(show_spinner=False, max_entries=3)
def _cert_traduz(depara: pd.DataFrame) -> dict[str, list]:
    """Ponta da planilha de cabos -> lista de TAGs reais da 01_BASE_TAGS.

    A mesma tradução que cert_painel_por_tag e cert_montagem já usam, agora
    reaproveitada: a direção varia por tipo de instrumento -- ZSH/L-120001 (uma
    ponta) vira ZSH-120001 e ZSL-120001 (duas TAGs separadas na base), enquanto
    HS-120610A1 e HS-120610A2 (duas pontas) viram a mesma HS-120610A1/A2 (uma
    TAG só, combinada, na base). Sem essa tradução, contagens que cruzam a
    planilha de cabos com a 01_BASE_TAGS por igualdade de texto nunca casam
    essas mais de 400 TAGs -- elas têm cabo lançado, só que escrito diferente.
    """
    if depara.empty:
        return {}
    traduz: dict[str, list] = {}
    for ponta, grupo in depara.groupby(cert_txt(depara["PONTA"])):
        if not ponta or ponta == "nan":
            continue
        traduz[ponta] = [t for t in cert_txt(grupo["TAG"]) if t and t != "nan"]
    return traduz


def cert_indice(lanc: pd.DataFrame, cache_key: str) -> dict:
    """De cada instrumento para a cadeia dele. É o que a busca consulta.

    Nem todo instrumento liga direto numa caixa de junção -- detecção de
    fumaça/gás (YST) e varios outros tipos correm em loop, instrumento a
    instrumento (YST-121101 -> YST-121100 -> YST-121165 -> ... -> painel),
    sem caixa nenhuma no meio; outros ligam direto no painel, pulando a caixa.
    Andar so um salto (o antigo comportamento) nunca achava caixa nenhuma
    nesses casos, e o instrumento inteiro sumia da busca -- mesmo com cabo
    lancado e concluido. Aqui anda pela cadeia ate achar a primeira caixa OU,
    na falta dela, o painel; 15 saltos cobrem com folga a maior cadeia real
    da base (10, medido em 2026-08-27).

    A caminhada tenta TODOS os ramos, não só o primeiro registrado na
    planilha -- mesmo ajuste que o cert_paineis já faz. Um nó com duas
    saídas (o cruzamento entre dois loops YST, por exemplo: YST-121156 liga
    tanto no resto do loop dele quanto, à parte, em YST-121144, de outro
    loop) tinha o ramo errado vencendo só por vir primeiro na planilha, e
    toda a cadeia dali pra trás (5 TAGs do LOOP9, neste caso) nunca achava
    caixa nem painel -- sumindo da busca por TAG mesmo com cabo lançado.
    """
    if lanc.empty:
        return {}
    adj: dict[str, list] = {}
    for _, r in lanc.iterrows():
        org = str(r["ORIGEM"]).strip()
        if org and org != "nan":
            adj.setdefault(org, []).append(str(r["DESTINO"]).strip())

    def alcanca(no):
        pilha, vistos = [(no, 0)], {no}
        while pilha:
            atual, prof = pilha.pop()
            if cert_nivel(atual) >= 1:
                return atual
            if prof >= 15:
                continue
            for prox in adj.get(atual, []):
                if prox not in vistos:
                    vistos.add(prox)
                    pilha.append((prox, prof + 1))
        return None

    saida: dict[str, str] = {}
    for _, r in lanc.iterrows():
        org, dst = str(r["ORIGEM"]).strip(), str(r["DESTINO"]).strip()
        # ponta em branco vira um instrumento fantasma na busca; ela ja aparece
        # na aba Correções, que e onde alguem trata
        if org in ("", "nan") or cert_nivel(org) != 0:
            continue
        alvo = alcanca(dst)
        if alvo:
            saida[org] = re.sub(r"[A-Z]$", "", alvo) if alvo.startswith("CFF") else alvo

    # Ponta que só aparece como DESTINO nunca acha caixa/painel andando pela
    # própria origem -- ela não tem uma (fim de loop instrumento-a-instrumento,
    # ou alimentada direto por um painel). Repete a busca olhando quem a
    # alimenta: se o alimentador já achou um alvo, essa TAG usa o mesmo (é a
    # mesma cadeia física); se o alimentador já é o próprio painel/caixa, o
    # alvo é ele mesmo.
    for _, r in lanc.iterrows():
        org, dst = str(r["ORIGEM"]).strip(), str(r["DESTINO"]).strip()
        if dst in ("", "nan") or cert_nivel(dst) != 0 or dst in adj or dst in saida:
            continue
        if cert_nivel(org) >= 1:
            saida[dst] = re.sub(r"[A-Z]$", "", org) if org.startswith("CFF") else org
        elif org in saida:
            saida[dst] = saida[org]

    # Os loops YST de CERT_LOOPS_YST têm o painel conferido contra o desenho
    # de interligação -- vale mais que a caminhada por grafo. Alguns desses
    # loops têm uma ligação de cruzamento com OUTRO loop (YST-121156, por
    # exemplo, também liga em YST-121144, de um loop de outro painel); a
    # caminhada pode achar esse painel vizinho antes do painel de verdade,
    # dependendo só da ordem em que as linhas aparecem na planilha -- e
    # nenhuma ordem é confiável o bastante pra decidir isso. Aqui, o painel
    # de cada TAG de um loop conhecido é sempre o do loop, sem ambiguidade.
    for nome_loop, (painel_loop, tags_loop) in CERT_LOOPS_YST.items():
        for t in tags_loop:
            saida[t] = painel_loop
    return saida


@st.cache_data(show_spinner=False, max_entries=3)
def cert_painel_por_tag(lanc: pd.DataFrame, depara: pd.DataFrame, cache_key: str) -> dict:
    """De cada TAG do segmento para o painel que a alimenta.

    Vale para o instrumento e para a caixa de junção: a caixa também é TAG, tem
    montagem e documento próprios, e deixá-la de fora fazia o filtro mostrar 20
    das 30 TAGs do PN-12-204.

    O nome da ponta passa pelo de-para antes de sair: a planilha de cabos e a
    base de TAGs escrevem a mesma TAG de formas diferentes, e sem a tradução um
    terço do segmento não casaria com a lista.

    Quando a subida não chega em painel nenhum -- caixa sem circuito de saída --
    a TAG fica de fora, em vez de receber um palpite.
    """
    if lanc.empty:
        return {}
    saida: dict[str, list] = {}
    for r in lanc.to_dict("records"):
        org = str(r["ORIGEM"]).strip()
        if org not in ("", "nan"):
            saida.setdefault(org, []).append(str(r["DESTINO"]).strip())

    traduz: dict[str, list] = {}
    if not depara.empty:
        for ponta, grupo in depara.groupby(cert_txt(depara["PONTA"])):
            traduz[ponta] = [t for t in cert_txt(grupo["TAG"]) if t and t != "nan"]

    pontas = {v for c in ("ORIGEM", "DESTINO") for v in cert_txt(lanc[c])}
    por_tag: dict[str, str] = {}
    for ponta in pontas:
        if ponta in ("", "nan") or cert_nivel(ponta) == 2:
            continue
        atual = ponta
        # 15 saltos, mesmo teto do cert_panorama: cadeia instrumento-a-instrumento
        # (sem caixa) pode ser mais longa que a caixa-a-caixa que o "seis" original
        # media
        for _ in range(15):
            if cert_nivel(atual) == 2:
                for nome in traduz.get(ponta, [ponta]):
                    por_tag[nome] = atual
                break
            prox = saida.get(atual)
            if not prox:
                break
            atual = prox[0]
    return por_tag


@st.cache_data(show_spinner=False, max_entries=3)
def cert_circuitos_por_tag(lanc: pd.DataFrame, cache_key: str) -> dict:
    """Todos os circuitos de cada instrumento, esteja onde estiver o destino.

    A cadeia de uma caixa só enxerga o que chega nela. O circuito de potência
    de uma TAG costuma ir direto ao painel, e ficaria de fora -- mas ele
    também precisa fechar para a TAG ser certificada.
    """
    if lanc.empty:
        return {}
    saida: dict[str, list] = {}
    for r in lanc.to_dict("records"):
        org = str(r["ORIGEM"]).strip()
        if org in ("", "nan") or cert_nivel(org) != 0:
            continue
        saida.setdefault(org, []).append({
            "id": str(r["CIRCUITO"]).strip(), "dst": str(r["DESTINO"]).strip(),
            "status": str(r["STATUS"]).strip(), "pct": round(cert_num(r["PCT"]), 1),
            "m": cert_num(r["METROS"]), "m_real": cert_metro_real(r),
            "pot": bool(re.search(r"-P\d*$", str(r["CIRCUITO"]).strip(), re.I)),
            "fibra": str(r["CIRCUITO"]).strip().upper().startswith("CFO")})
    for cs in saida.values():
        cs.sort(key=lambda c: c["id"])
    return saida


@st.cache_data(show_spinner=False, max_entries=3)
def cert_circuitos_por_ponta(lanc: pd.DataFrame, depara: pd.DataFrame,
                             cache_key: str) -> tuple[dict, dict]:
    """Os circuitos que saem de cada ponta, e a metragem de cada circuito.

    Devolve os IDS dos circuitos, não metros já somados, porque o total de um
    conjunto NÃO é a soma dos totais de cada um: o cabo do ZSH/L-120112 sai
    uma vez e serve duas TAGs da base (ZSH-120112 e ZSL-120112). Somando por
    TAG, os 129 m dele entravam duas vezes -- são 197 circuitos assim, 6.879 m
    contados a mais no cartão do topo. Com os ids, quem soma faz a união antes
    e cada cabo entra uma vez só.

    A ponta entra pelo nome cru da planilha de cabos E pelo nome da
    01_BASE_TAGS quando o de-para sabe traduzir. Sem isso 564 TAGs apareciam
    em "TAGs mapeadas" e somavam zero metro, porque a chave aqui era o nome
    cru e o universo lá em cima usa o nome da base: 24.907 m ficavam fora da
    conta -- o cartão dizia 50.027 m onde havia 67.893.
    """
    if lanc.empty:
        return {}, {}
    metros: dict[int, tuple[float, float]] = {}
    da_ponta: dict[str, set] = {}
    # A chave é a LINHA, não o nome do circuito: três nomes se repetem em
    # linhas diferentes (C-YST-121157 sai do 121147 com 36 m e do 121156 com
    # 0 m; C-PN-12-238-03 e C-YST-121166 idem). Chaveando pelo nome, uma
    # linha apagava a outra e sumiam 162 m -- 141 deles no MB-RTU-03 sozinho.
    # Isso não atrapalha a união: o cabo compartilhado do ZSH/L é UMA linha,
    # e as duas TAGs apontam para ela.
    for i, r in enumerate(lanc.to_dict("records")):
        org = str(r["ORIGEM"]).strip()
        if org in ("", "nan"):
            continue
        # já vem capado no previsto pelo cert_metro_real
        metros[i] = (cert_num(r["METROS"]), cert_metro_real(r))
        da_ponta.setdefault(org, set()).add(i)
    for cru, alvos in _cert_traduz(depara).items():
        if cru in da_ponta:
            for alvo in alvos:
                da_ponta.setdefault(alvo, set()).update(da_ponta[cru])
    return da_ponta, metros


def cert_metros_uniao(da_ponta: dict, metros: dict, nomes) -> tuple[float, float]:
    """Previsto e lançado de um conjunto de pontas, cada circuito uma vez."""
    ids: set = set()
    for n in nomes:
        ids |= da_ponta.get(n, set())
    return (sum(metros[c][0] for c in ids), sum(metros[c][1] for c in ids))


# Os quatro campos da 01_BASE_TAGS que recortam a Certificação. A ordem é a
# da cadeia física, do painel para a ponta: painel -> caixa -> segmento H1 ->
# malha. CFF e PAINEL só existem na planilha depois do pipeline que os
# importa; sem eles o filtro correspondente simplesmente não aparece, em vez
# de a aba quebrar numa planilha antiga.
CERT_FILTROS = [("PAINEL", "Painel"), ("CFF", "Caixa (CFF)"),
                ("SEGMENTO", "Segmento"), ("MALHA", "Malha")]


@st.cache_data(show_spinner=False, max_entries=3)
def cert_atributos(tags: pd.DataFrame, cache_key: str) -> dict:
    """TAG -> painel, caixa, segmento e malha, direto da base.

    A base escreve "-" onde não há valor, e são 3.776 das 5.098 TAGs sem
    fieldbus: normalizar para vazio aqui evita que "-" vire uma opção de
    filtro que não quer dizer nada.
    """
    campos = [c for c, _ in CERT_FILTROS if c in tags.columns]
    if not campos:
        return {}
    fatia = tags[["TAG"] + campos]
    return {str(r["TAG"]).strip():
            {c: ("" if vazio(r[c]) else str(r[c]).strip()) for c in campos}
            for r in fatia.to_dict("records") if str(r["TAG"]).strip()}


@st.cache_data(show_spinner=False, max_entries=3)
def cert_paineis(lanc: pd.DataFrame, cache_key: str) -> dict:
    """Cada painel e as caixas que saem dele, com o que pendura em cada uma.

    Só a caixa de primeiro nível vira cartão. O que vem pendurado nela conta na
    subárvore dela: desenhar tudo faria da tela um mapa de fios, e o detalhe já
    é o que a vista da caixa mostra.
    """
    if lanc.empty:
        return {}
    saida: dict[str, list] = {}
    entrada: dict[str, list] = {}
    # Índice por nome de circuito: o cabo de um instrumento leva o nome dele
    # (C-YST-X), e é por aqui que ele é achado -- as colunas ORIGEM/DESTINO
    # gravam a partir da ponta de campo e não servem para isso.
    por_circ: dict[str, list] = {}
    for r in lanc.to_dict("records"):
        org = str(r["ORIGEM"]).strip()
        dst = str(r["DESTINO"]).strip()
        if org not in ("", "nan"):
            saida.setdefault(org, []).append(r)
        if dst not in ("", "nan"):
            entrada.setdefault(dst, []).append(r)
        por_circ.setdefault(str(r["CIRCUITO"]).strip(), []).append(r)

    def sobe(no):
        """O caminho de um nó até o painel, do primeiro passo ao último.

        Quando o nó tem mais de uma saída (loop de instrumento a instrumento
        com duas pontas, por exemplo), seguir sempre a primeira registrada
        podia cair num ramo que nunca chega no painel e perder o instrumento
        inteiro do desenho -- mesmo com outro ramo que fecharia certinho.
        Por isso tenta os ramos por busca, não só o primeiro da lista.
        """
        pilha = [(no, [])]
        vistos = {no}
        while pilha:
            atual, caminho = pilha.pop()
            if len(caminho) >= 15:
                continue
            for prox_r in saida.get(atual, []):
                destino = str(prox_r["DESTINO"]).strip()
                novo = caminho + [(atual, prox_r)]
                if cert_nivel(destino) == 2:
                    return novo, destino
                if destino not in vistos:
                    vistos.add(destino)
                    pilha.append((destino, novo))
        return [], None

    paineis: dict[str, dict] = {}
    for ponta in saida:
        if cert_nivel(ponta) != 1:
            continue
        caminho, painel = sobe(ponta)
        if not painel:
            continue
        raiz, circ = caminho[-1]      # a caixa que fala direto com o painel
        p = paineis.setdefault(painel, {})
        b = p.setdefault(raiz, {"nome": raiz, "tronco": _cert_circuito(circ, {}),
                                "caixas": set(), "inst": 0, "cabo_ok": 0, "tags": [],
                                "direto": False})
        b["caixas"].add(ponta)
    # cada instrumento entra na conta da caixa de primeiro nível que o alimenta
    for ponta, circuitos in saida.items():
        if cert_nivel(ponta) != 0:
            continue
        caminho, painel = sobe(ponta)
        if not painel:
            continue
        raiz = caminho[-1][0]
        p = paineis.setdefault(painel, {})
        b = p.get(raiz)
        if b is None:
            # sem caixa nenhuma no meio -- loop de instrumento a instrumento
            # (deteccao de fumaca/gas, por exemplo) ou instrumento ligado
            # direto no painel. "raiz" aqui e o ultimo instrumento antes do
            # painel, o mesmo pra todo mundo que faz parte dessa cadeia --
            # e o que agrupa o loop inteiro num bloco so, do jeito que uma
            # caixa agruparia quem pendura nela. Sem isso o instrumento
            # simplesmente sumia do desenho do painel.
            b = p.setdefault(raiz, {"nome": raiz, "tronco": _cert_circuito(caminho[-1][1], {}),
                                    "caixas": set(), "inst": 0, "cabo_ok": 0, "tags": [],
                                    "direto": True})
        b["inst"] += 1
        # o cartao mostra o cabo DO INSTRUMENTO (o circuito proprio dele na
        # planilha), nao a cadeia inteira ate a alimentacao do painel: pintar
        # de vermelho um instrumento cujo cabo esta Concluido, porque falta o
        # cabo do painel muitos saltos acima, e dizer o oposto do que a
        # planilha de cabos diz. A cadeia continua no "cabo_cadeia", e e ela
        # que a coluna Certificacao e o cert_panorama usam.
        # o cabo do instrumento e o circuito que leva o NOME dele (C-YST-X),
        # quando existe: e assim que o desenho identifica o cabo de cada
        # detector, e e a leitura que bate com o campo
        proprios = [c for c in circuitos if not cert_remendo(c)] or circuitos
        # procurado na base INTEIRA, não só nas linhas onde esta TAG é ponta:
        # o último ponto do laço só recebe cabo, e limitar a busca às pontas
        # dele devolvia metragem zero para um cabo que está lançado
        nomeado = [c for c in por_circ.get(f"C-YST-{ponta.split('-')[-1]}", [])
                   if not cert_remendo(c)]
        if nomeado:
            proprios = nomeado
        pronto_tag = all(cert_num(c["PCT"]) >= 99.5 for c in proprios)
        pronto_cadeia = all(cert_num(c["PCT"]) >= 99.5 for _, c in caminho)
        if pronto_tag:
            b["cabo_ok"] += 1
        m = sum(cert_num(c["METROS"]) for c in proprios)
        m_real = sum(cert_metro_medido(c) for c in proprios)
        b["tags"].append({"org": ponta, "cabo": pronto_tag,
                          "cabo_cadeia": pronto_cadeia,
                          # status da base, metragem à parte -- ver
                          # cert_status_conjunto
                          "status": cert_status_conjunto(proprios),
                          "pct": round(m_real / m * 100, 1) if m else 0.0,
                          "m": m, "m_real": m_real,
                          # distancia ate o painel/caixa -- e o que da a ordem
                          # fisica do laco (ver ordem_tags mais abaixo)
                          "prof": len(caminho)})
    def ordem_tags(v):
        if v.get("ordem_fixa"):
            # loop com identidade conferida no diagrama (CERT_LOOPS_YST) --
            # a ordem ja e a fisica de verdade, nao a estimada pela
            # profundidade da caminhada
            return v["tags"]
        if v["direto"]:
            # num loop de instrumento a instrumento a ordem que importa e a
            # da fiacao fisica: quem esta mais longe do painel vem primeiro,
            # quem fecha nele vem por ultimo -- a mesma sequencia que o
            # diagrama de interligacao desenha (instrumento A -> B -> ... ->
            # painel). Numa caixa comum os instrumentos sao ramais
            # independentes, nao um laco, e a ordem alfabetica continua.
            return sorted(v["tags"], key=lambda t: -t["prof"])
        return sorted(v["tags"], key=lambda t: t["org"])

    # Os loops de deteccao de fumaca/gas (YST) tem identidade fixa, conferida
    # contra o diagrama de interligacao DE-5290.00-2111-800-CHZ-135 -- nao
    # dependem mais da caminhada por grafo pra dizer QUEM e do mesmo loop.
    # A caminhada quebrava exatamente quando um instrumento tem ligacao
    # direta com o painel ALEM de continuar o loop (ex.: YST-121149 liga em
    # YST-121150 E direto no painel): sobe() achava o atalho primeiro e
    # isolava 149 num bloco de 1, longe do resto do loop dele. Aqui o grupo,
    # a ordem e o nome vem do desenho; o status de cada TAG (cabo pronto,
    # pct) continua vindo da caminhada de verdade, sem mudar a matematica.
    for nome_loop, (painel_loop, ordem) in CERT_LOOPS_YST.items():
        p = paineis.get(painel_loop)
        if p is None:
            continue
        achados: dict[str, dict] = {}
        tronco_loop = None
        for raiz, b in list(p.items()):
            restantes = []
            for t in b["tags"]:
                if t["org"] in ordem:
                    achados[t["org"]] = t
                    if raiz == t["org"] and tronco_loop is None:
                        tronco_loop = b["tronco"]
                else:
                    restantes.append(t)
            if len(restantes) != len(b["tags"]):
                if restantes:
                    b["tags"] = restantes
                    b["inst"] = len(restantes)
                    b["cabo_ok"] = sum(1 for x in restantes if x["cabo"])
                else:
                    del p[raiz]
        # Um membro do loop pode nunca ter sido visitado pela caminhada: ela
        # so parte de quem TEM saida propria (organizada por ORIGEM), e um
        # instrumento que so recebe cabo -- ultimo do trecho antes do painel,
        # sem nada saindo dele -- nunca vira "ponta" pra caminhada nenhuma.
        # Sem isto o loop ficava incompleto no desenho (YST-121103/111/113/116
        # sumiam de LOOP2/3/4 inteiros, mesmo com cabo cadastrado): aqui o
        # circuito que chega nele -- a propria convencao da base, que nomeia
        # o circuito pelo destino -- vira a fonte do status desse instrumento.
        for t_org in ordem:
            if t_org in achados:
                continue
            linhas = saida.get(t_org) or entrada.get(t_org) or []
            # O remendo de topologia e um preenchimento so pra fechar o
            # grafo -- nunca o status de verdade da TAG. Sem isto,
            # YST-121111 (que so recebe: o circuito de painel direto
            # C-YST-121111, Concluido, E o remendo 112->111, sempre "Nao
            # Iniciado" por definicao) tinha o remendo decidindo o status,
            # trocando "Concluido" por "pendente" so por causa de qual das
            # duas linhas veio primeiro na planilha.
            reais = [rr for rr in linhas if not cert_remendo(rr)]
            if reais:
                linhas = reais
            # O cabo do instrumento e o circuito com o NOME dele (C-YST-X), e
            # ele e procurado na base INTEIRA -- nao so entre as linhas onde
            # esta TAG aparece como ponta. O ultimo ponto do laco so recebe
            # cabo, entao procurar pelas pontas dele nao achava nada e a
            # tabela mostrava "Concluido, 0 de 0 m": status certo com
            # metragem zerada, que e pior que erro, e contradicao na mesma
            # linha.
            proprio = [rr for rr in por_circ.get(f"C-YST-{t_org.split('-')[-1]}", [])
                       if not cert_remendo(rr)]
            if proprio:
                linhas = proprio
            pronto_t = bool(linhas) and all(cert_num(rr["PCT"]) >= 99.5 for rr in linhas)
            m_t = sum(cert_num(rr["METROS"]) for rr in linhas)
            m_real_t = sum(cert_metro_medido(rr) for rr in linhas)
            achados[t_org] = {
                "org": t_org,
                "cabo": pronto_t,
                "cabo_cadeia": pronto_t,
                "status": cert_status_conjunto(linhas),
                "pct": round(m_real_t / m_t * 100, 1) if m_t else 0.0,
                "m": m_t, "m_real": m_real_t,
                "prof": 0,
            }
        if not achados:
            continue
        tags_loop = [achados[t] for t in ordem if t in achados]
        if tronco_loop is None:
            tronco_loop = {"id": "", "disc": "", "status": "—", "pct": 0, "m": 0,
                           "m_real": 0, "pot": False, "org": "", "dst": ""}
        p[nome_loop] = {
            "nome": nome_loop, "tronco": tronco_loop, "caixas": set(),
            "inst": len(tags_loop), "cabo_ok": sum(1 for t in tags_loop if t["cabo"]),
            "tags": tags_loop, "direto": True, "ordem_fixa": True,
        }

    def circ_entre(a, b):
        """O circuito real entre duas pontas, em qualquer das duas direções --
        a planilha grava a partir da ponta de campo, que varia por linha."""
        for r in saida.get(a, []):
            if str(r["DESTINO"]).strip() == b:
                return r
        for r in saida.get(b, []):
            if str(r["DESTINO"]).strip() == a:
                return r
        return None

    def circuito_do_trecho(a, b, usados):
        """O cabo de um trecho quando a linha da base traz a outra ponta.

        A base nomeia o circuito pela ponta de DESTINO -- conferido em 88 de
        88 circuitos C-YST-*. Então o cabo do trecho A–B chama-se
        C-YST-<A> ou C-YST-<B>. Em 9 trechos dos laços a linha existe,
        está Concluída, e só a ORIGEM aponta o vizinho errado: o cabo está
        lançado e o desenho mostrava "não iniciado" por causa do remendo.

        Só entra circuito LIVRE -- que nenhum outro trecho deste laço já usa.
        Sem essa trava, o cabo do trecho vizinho seria contado duas vezes:
        em LOOP1 o C-YST-121178 é o cabo de 121177–121178, e aceitá-lo
        também em 121102–121178 somaria o mesmo metro nos dois.
        """
        for cand in (f"C-YST-{b.split('-')[-1]}", f"C-YST-{a.split('-')[-1]}"):
            if cand in usados:
                continue
            achadas = por_circ.get(cand)
            if achadas:
                return achadas[0]
        return None

    def com_fiacao(v, painel):
        """A ordem final, com o fio real (da base de cabos) até o próximo
        instrumento do laço -- e até o painel, no último. Sem isto o desenho
        ligava um cartão no outro com um traço decorativo, sem metragem nem
        status: exatamente a informação "como de costume" que falta na busca
        por TAG de um laço YST.
        """
        ordenados = ordem_tags(v)
        if not (v.get("direto") and len(ordenados) > 1):
            return ordenados, None
        ordenados = [dict(t) for t in ordenados]
        fixa = bool(v.get("ordem_fixa"))
        if fixa:
            # REGRA DO DESENHO: o cabo de um instrumento leva o NOME DELE --
            # o cabo que sai do YST-121159 é o C-YST-121159, e é ele que vai
            # ao próximo ponto do laço (ou ao painel, no último).
            #
            # As colunas ORIGEM/DESTINO não montam o laço: a base grava a
            # partir da ponta de campo, então o sentido inverte de linha para
            # linha e a ponta gravada nem sempre é a vizinha do desenho. Ler
            # o laço por elas foi o que fez o trecho até o painel ficar sem
            # metragem e sem avanço, e o que me levou a criar remendos em 0%
            # por cima de cabo lançado. Pelo nome do circuito, as 80 TAGs dos
            # laços têm cabo próprio e todas estão Concluídas -- que é o que
            # o campo confirma.
            for t in ordenados:
                achadas = por_circ.get(f"C-YST-{t['org'].split('-')[-1]}")
                t["circ_prox"] = _cert_circuito(achadas[0], {}) if achadas else None
            ult = por_circ.get(f"C-YST-{ordenados[-1]['org'].split('-')[-1]}")
            return ordenados, (_cert_circuito(ult[0], {}) if ult else None)
        # duas passadas: a primeira anota o que a base já resolve e reserva
        # esses circuitos; só depois os vãos procuram cabo livre
        trechos, usados = [], set()
        for i, t in enumerate(ordenados):
            prox = ordenados[i + 1]["org"] if i + 1 < len(ordenados) else painel
            r = circ_entre(t["org"], prox)
            if r is not None and not cert_remendo(r):
                usados.add(str(r["CIRCUITO"]).strip())
            trechos.append((t, prox, r))
        for t, prox, r in trechos:
            if fixa and (r is None or cert_remendo(r)):
                alt = circuito_do_trecho(t["org"], prox, usados)
                if alt is not None:
                    usados.add(str(alt["CIRCUITO"]).strip())
                    r = alt
            t["circ_prox"] = _cert_circuito(r, {}) if r is not None else None
        r_saida = circ_entre(ordenados[-1]["org"], painel)
        if fixa and (r_saida is None or cert_remendo(r_saida)):
            alt = circuito_do_trecho(ordenados[-1]["org"], painel, usados)
            r_saida = alt if alt is not None else r_saida
        return ordenados, (_cert_circuito(r_saida, {}) if r_saida is not None else None)

    saida_final: dict[str, list] = {}
    for painel, cs in paineis.items():
        blocos = []
        for v in cs.values():
            tags_fio, circ_saida = com_fiacao(v, painel)
            blocos.append(dict(v, caixas=len(v["caixas"]), tags=tags_fio,
                               circ_saida=circ_saida))
        saida_final[painel] = sorted(blocos, key=lambda x: x["nome"])
    return saida_final


# Ordem fisica real de cada loop de deteccao de fumaca/gas, do diagrama de
# interligacao DE-5290.00-2111-800-CHZ-135 (posicao das caixas + a ligacao de
# margem entre a 1a caixa de cada fileira, conferida em 2026-08-28). So os 8
# loops com ordem limpa e ja confirmada entram aqui -- LOOP7 tem topologia em
# Y (duas pontas saindo do painel e convergindo num beco sem saida) e LOOP8
# nao deu pra ler com confianca no desenho (pagina em zigue-zague), entao
# os dois continuam pela caminhada automatica ate serem conferidos direito.
CERT_LOOPS_YST: dict[str, tuple[str, list[str]]] = {
    "LOOP1": ("PN-12-236", ["YST-121102", "YST-121178", "YST-121177", "YST-121101",
                           "YST-121100", "YST-121165", "YST-121166"]),
    "LOOP2": ("PN-12-236", ["YST-121179", "YST-121109", "YST-121110", "YST-121198",
                           "YST-121108", "YST-121104", "YST-121103"]),
    "LOOP3": ("PN-12-236", ["YST-121113", "YST-121114", "YST-121115", "YST-121116",
                           "YST-121117", "YST-121118", "YST-121119", "YST-121120",
                           "YST-121121", "YST-121122"]),
    "LOOP4": ("PN-12-236", ["YST-121111", "YST-121112", "YST-121125", "YST-121126",
                           "YST-121127", "YST-121128", "YST-121167", "YST-121168",
                           "YST-121169", "YST-121170"]),
    "LOOP5": ("PN-12-237", ["YST-121171", "YST-121172", "YST-121173", "YST-121174",
                           "YST-121175", "YST-121176", "YST-121124", "YST-121123"]),
    "LOOP6": ("PN-12-237", ["YST-121129", "YST-121130", "YST-121131", "YST-121132",
                           "YST-121133", "YST-121134", "YST-121135", "YST-121136",
                           "YST-121137", "YST-121138"]),
    # LOOP7 sai e volta ao PN-12-237 como os outros, e a sequência inteira
    # vem do encadeamento dos cabos na base: C-YST-121140 encosta em 121139,
    # C-YST-121192 em 121140, e assim por diante. Ficava de fora por um elo
    # só -- entre o YST-121141 e o YST-121163 não há circuito cadastrado na
    # planilha de cabos, só o remendo CORRECAO-YST-121163-YST-121141 -- e sem
    # fechar o laço nenhum destes 8 achava painel: sumiam do desenho, e com
    # eles 8 das 26 TAGs do segmento MB-RTU-02. Declarado aqui, os 8 aparecem
    # e o elo que falta continua vermelho, que é o que ele é.
    "LOOP7": ("PN-12-237", ["YST-121139", "YST-121140", "YST-121192", "YST-121141",
                           "YST-121163", "YST-121164", "YST-121162", "YST-121161"]),
    # LOOP8 tem formato proprio: passa por um dispositivo (CBZ-12-001, o
    # BZ-500 do desenho) no meio do percurso, entre YST-121160 e YST-121145.
    # Fica aqui pelo mesmo motivo dos outros -- a base cruza este laco com o
    # LOOP9 (121147 -> 121157 e 121156 -> 121144), e sem a identidade fixa os
    # dois apareciam grudados num bloco so. Laco nenhum e junto com outro.
    "LOOP8": ("PN-12-238", ["YST-121189", "YST-121190", "YST-121191", "YST-121160",
                           "YST-121145", "YST-121146", "YST-121147", "YST-121144"]),
    "LOOP9": ("PN-12-238", ["YST-121149", "YST-121150", "YST-121151", "YST-121152",
                           "YST-121153", "YST-121155", "YST-121156", "YST-121157",
                           "YST-121158", "YST-121159"]),
    "LOOP10": ("PN-12-238", ["YST-121148", "YST-121180", "YST-121181", "YST-121182",
                            "YST-121183", "YST-121184", "YST-121185", "YST-121186",
                            "YST-121187", "YST-121188"]),
}


@st.cache_data(show_spinner=False, max_entries=3)
def cert_alimentacao_painel(lanc: pd.DataFrame, painel: str, cache_key: str) -> list:
    """Os circuitos que alimentam um painel, na mesma regra do cert_panorama.

    Primeiro a ELÉTRICA que chega nele; na falta dela, o circuito "-P" que
    sai dele para outro painel ou dispositivo -- em PN-12-236/237 a
    alimentação vem assim, em INSTRUMENTAÇÃO, marcada só pelo sufixo.
    """
    if lanc.empty or not painel:
        return []
    linhas = lanc.to_dict("records")
    alim = [r for r in linhas
            if str(r["DESTINO"]).strip() == painel
            and str(r["DISCIPLINA"]).strip() == "ELÉTRICA"]
    if not alim:
        no, vistos = painel, set()
        for _ in range(15):
            if no in vistos:
                break
            vistos.add(no)
            pot = [r for r in linhas if str(r["ORIGEM"]).strip() == no
                   and re.search(r"-P\d*$", str(r["CIRCUITO"]).strip(), re.I)]
            if not pot:
                break
            alim = pot
            no = str(pot[0]["DESTINO"]).strip()
    return [_cert_circuito(r, {}) for r in alim]


def cert_cadeia_painel(painel: str, blocos: list, mont: dict,
                       alimentacao: list | None = None) -> dict:
    """A cena do painel: ele, a alimentação dele e as caixas de primeiro nível.

    A alimentação entra no desenho porque é ela que costuma travar tudo: no
    laço YST todos os trechos estão lançados e mesmo assim a certificação
    reprova, porque o cabo que alimenta o painel (C-PN-12-236-P → PN-12102)
    está em 0%. Sem ele desenhado, a tela mostrava só trecho verde e escrevia
    "Bloqueado" do lado, sem nada vermelho que explicasse -- o desenho dizia
    uma coisa e a tabela outra.
    """
    return {"tipo": "painel", "caixa": painel, "painel": painel,
            "painel_indef": False, "mont": mont.get(painel, {}).get("mont", ""),
            "eletrica": alimentacao or [], "tronco": [], "segmentos": [],
            "ligacoes": [],
            "ramais": [], "blocos": [
                # o laço não é TAG: não se monta, e por isso não recebe
                # montagem nenhuma. Só a caixa de junção, que é TAG de
                # verdade e tem instalação própria, é consultada aqui.
                {**b, "mont": ("" if b.get("ordem_fixa")
                               else mont.get(b["nome"], {}).get("mont", "")),
                 "tags": [{**t, "mont": mont.get(t["org"], {}).get("mont", "")}
                          for t in b["tags"]]}
                for b in blocos]}


def cert_cadeia(alvo: str, lanc: pd.DataFrame, mont: dict) -> dict:
    """A cadeia de um alvo, nas duas topologias que a planilha guarda.

    convencional  painel — caixa — instrumentos
    fieldbus      painel — CFF-A — CFF-B — CFF-C, com instrumentos em cada caixa

    A diferença não é enfeite: no fieldbus, quem deriva da caixa A não espera o
    tronco B→C, e é isso que decide se a TAG pode ser certificada.
    """
    d = lanc
    org, dst = cert_txt(d["ORIGEM"]), cert_txt(d["DESTINO"])

    def eletrica(painel):
        if not painel:
            return []
        e = d[(cert_txt(d["DISCIPLINA"]) == "ELÉTRICA") & (dst == painel)]
        return [_cert_circuito(r, mont) for _, r in e.iterrows()]

    chegada: dict[str, list] = {}
    for _, r in d.iterrows():
        ds = str(r["DESTINO"]).strip()
        if ds and ds != "nan":
            chegada.setdefault(ds, []).append(r)

    def upstream(destinos, excluir_prefixo):
        """As linhas que chegam nos destinos dados, inclusive quem vem por
        trás de outro instrumento sem caixa no meio -- o mesmo loop de
        instrumento a instrumento (detecção de fumaça/gás, por exemplo) que o
        cert_panorama já atravessa. Sem isso só o último instrumento antes da
        caixa aparecia no desenho, e o resto da cadeia ficava escondido.

        Cada linha sai marcada com a "raiz": qual dos destinos originais essa
        cadeia alcança, pra quem chama saber de qual caixa/segmento aquele
        instrumento -- mesmo vários saltos atrás -- pendura.
        """
        vistos = set(destinos)
        coletadas: list[tuple] = []
        fronteira = [(no, no) for no in destinos]
        for _ in range(15):
            proxima = []
            for no, raiz in fronteira:
                for r in chegada.get(no, []):
                    o = str(r["ORIGEM"]).strip()
                    if o.startswith(excluir_prefixo):
                        continue
                    coletadas.append((r, raiz))
                    if cert_nivel(o) == 0 and o not in vistos:
                        vistos.add(o)
                        proxima.append((o, raiz))
            if not proxima:
                break
            fronteira = proxima
        return sorted(coletadas, key=lambda x: str(x[0]["ORIGEM"]).strip())

    segs = sorted({v for v in pd.concat([org, dst])
                   if re.fullmatch(re.escape(alvo) + r"[A-Z]", v)})
    if segs:
        ao_painel = d[org.isin(segs) & dst.str.startswith(PAINEL_PREFIXOS)]
        painel = str(ao_painel["DESTINO"].iloc[0]).strip() if len(ao_painel) else None
        entre = d[org.isin(segs) & dst.isin(segs)]
        elos = {tuple(sorted([str(r["ORIGEM"]).strip(), str(r["DESTINO"]).strip()])):
                _cert_circuito(r, mont) for _, r in entre.iterrows()}
        ligacoes = []
        for a, b in zip(segs, segs[1:]):
            c = elos.get(tuple(sorted([a, b])))
            if c:
                ligacoes.append({**c, "de": a, "para": b})
        # a ponta da cadeia e a caixa que nenhum tronco tem como destino: e nela
        # que o terminador fecha o segmento
        destinos = {str(v).strip() for v in entre["DESTINO"]}
        fim = [x for x in segs if x not in destinos]
        ins = upstream(segs, "CFF")
        return {
            "tipo": "cff", "caixa": alvo, "painel": painel, "painel_indef": painel is None,
            "terminador": fim[-1] if fim else segs[-1],
            "mont": mont.get(alvo, {}).get("mont", ""),
            "eletrica": eletrica(painel),
            "tronco": [_cert_circuito(r, mont) for _, r in ao_painel.iterrows()],
            "segmentos": [{"nome": x, "ordem": i,
                           "mont": mont.get(x, {}).get("mont", "")}
                          for i, x in enumerate(segs)],
            "ligacoes": ligacoes,
            "ramais": cert_agrupar(
                [{**_cert_circuito(r, mont), "seg": raiz} for r, raiz in ins]),
        }

    sai = d[org == alvo]
    # caixa sem circuito de saida: nao da para saber a que painel ela chega. O
    # painel nao some do desenho por isso -- aparece sem tag, com a observacao,
    # e a etapa fica inconclusiva.
    painel = str(sai["DESTINO"].iloc[0]).strip() if len(sai) else None
    ent = upstream([alvo], "CJ")
    return {
        "tipo": "caixa", "caixa": alvo, "painel": painel, "painel_indef": painel is None,
        "mont": mont.get(alvo, {}).get("mont", ""),
        "eletrica": eletrica(painel),
        "tronco": [_cert_circuito(r, mont) for _, r in sai.iterrows()],
        "segmentos": [], "ligacoes": [],
        "ramais": cert_agrupar([_cert_circuito(r, mont) for r, _ in ent]),
    }


def cert_agrupar(circuitos: list) -> list:
    """Junta os circuitos de cada TAG num cartão só.

    O estado do conjunto é o do pior circuito: a TAG não está pronta porque
    metade dos cabos dela chegou. O percentual vem do metro lançado sobre o
    total, que é a conta que o campo reconhece.
    """
    por_tag: dict[str, list] = {}
    for c in circuitos:
        por_tag.setdefault(c["org"], []).append(c)
    saida = []
    for org, cs in por_tag.items():
        if len(cs) == 1:
            saida.append({**cs[0], "circuitos": cs})
            continue
        total = sum(c["m"] for c in cs)
        real = sum(c["m_real"] for c in cs)
        pct = round(real / total * 100, 1) if total else 0.0
        if all(c["pct"] >= 99.5 for c in cs):
            status = "Concluído"
        elif any(c["pct"] > 0 for c in cs):
            status = "Em Andamento"
        else:
            status = cs[0]["status"]
        saida.append({**cs[0], "id": f"{len(cs)} circuitos", "status": status,
                      "pct": pct, "m": total, "m_real": real, "circuitos": cs})
    return sorted(saida, key=lambda c: c["org"])


CERT_ROTULO = {"ok": "Apto", "warn": "Predecessora em andamento",
               "crit": "Bloqueado", "desc": "Não dá para afirmar"}
CERT_CLASSE = {"ok": "ok", "warn": "warn", "crit": "crit", "desc": "roxo"}


CERT_JS = r"""
const D = __DADOS__;
let zoom = 1;
// quais caixas estao abertas na vista do painel. Vive aqui dentro: abrir uma
// caixa nao e pergunta para o servidor. Uma cena de bloco so (a busca por
// uma TAG de laco YST cai nisso, ver render_certificacao) ja nasce aberta --
// pedir um clique a mais pra ver a unica coisa na tela nao ajuda ninguem.
const abertos = new Set(D.cadeia.tipo === 'painel' && D.cadeia.blocos.length === 1
  ? [D.cadeia.blocos[0].nome] : []);

const br = (n, c) => Number(n).toLocaleString('pt-BR',
  {minimumFractionDigits: c || 0, maximumFractionDigits: c || 0});
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
const feito = c => c.pct >= 99.5;
const cls = s => s === 'Concluído' ? 'ok' : s === 'Em Andamento' ? 'and' : 'nao';
const corMont = m => m === 'Montado' ? 'ok' : m === 'Em Programação' ? 'and'
                   : m ? 'nao' : 'roxo';
const nomeMont = m => m || 'sem correspondente na base de TAGs';

// A dica viaja num atributo do próprio elemento: um ouvinte só, no documento,
// em vez de um por cartão.
const dd = h => `data-d="${h.replace(/"/g, '&quot;')}"`;
const dl = (k, v, tom) => `<div class='l'><i>${k}</i><b${
  tom ? ` style='color:var(--${tom})'` : ''}>${v}</b></div>`;
const cab = (h, n) => `<div class='h'>${h}</div><div class='n'>${esc(n)}</div>`;

// A planilha grava ORIGEM na ponta de campo e DESTINO no lado de montante: o
// tronco entra lá como CFF-12-0024A → PN-12-201, que é o inverso de como o
// cabo corre. A dica mostra o percurso físico e repete as colunas cruas.
const nivel = v => /^(PN|PL|PCC)/.test(v) ? 2 : /^(CJ|CFF)/.test(v) ? 1 : 0;
const percurso = c => c.de && c.para ? [c.de, c.para]
  : nivel(c.org) < nivel(c.dst) ? [c.dst, c.org] : [c.org, c.dst];

function dicaCabo(c, papel) {
  const t = cls(c.status), [a, b] = percurso(c);
  return cab(papel || 'cabo', c.id) + dl('situação', c.status, t) +
    (c.pct > 0 && c.pct < 100 ? dl('lançado', br(c.pct, 1) + '%', t) : '') +
    dl('lançado', br(c.m_real === undefined ? c.m * c.pct / 100 : c.m_real) +
       ' de ' + br(c.m) + ' m') + dl('disciplina', c.disc) +
    dl('percurso', esc(a) + ' → ' + esc(b)) +
    (a !== c.org ? `<div class='obs'>na planilha: <b>ORIGEM</b> ${esc(c.org)} ·
      <b>DESTINO</b> ${esc(c.dst)} — a base grava a partir da ponta de campo,
      que é o inverso de como o cabo corre</div>` : '');
}

function dicaTag(r) {
  const cs = r.circuitos || [r];
  // com mais de um circuito, a linha do cabo é o conjunto e a lista abre a
  // conta: é a diferença entre "em andamento" e saber qual dos dois falta
  const papel = c => c.fibra ? 'fibra' : c.pot ? 'potência' : 'sinal';
  const lista = `<div class='obs'>
    <div style='color:var(--t2);font-weight:700;margin-bottom:5px'>${cs.length}
      ${cs.length === 1 ? 'circuito' : 'circuitos'} desta TAG</div>` + cs.map(c => `
    <div class='l' style='padding:3px 0'>
      <i>${esc(c.id)} <b style='color:var(--t3);font-weight:600'>${papel(c)}</b>
        <b style='color:var(--t3);font-weight:400;display:block;font-size:9.5px'>
          → ${esc(c.dst || '')}</b></i>
      <b style='color:var(--${cls(c.status)})'>${c.status}<b
        style='color:var(--t2);font-weight:600;display:block;font-size:10px'>
        ${br(c.m_real)} de ${br(c.m)} m</b></b></div>`).join('') + '</div>';
  return cab('instrumento', r.org) +
    dl('certificação', r.rot, r.tom === 'ok' ? 'ok' : r.tom === 'warn' ? 'and'
       : r.tom === 'crit' ? 'nao' : 'roxo') +
    (r.tom === 'ok' ? '' : dl('trava em', r.onde)) +
    dl('montagem', nomeMont(r.mont), corMont(r.mont)) +
    dl('cabo', r.status + (r.pct > 0 && r.pct < 100 ? ' · ' + br(r.pct, 1) + '%' : ''),
       cls(r.status)) +
    dl('lançado', br(r.m_real) + ' de ' + br(r.m) + ' m') +
    (r.seg ? dl('caixa', esc(r.seg)) : '') + lista +
    (r.ancora ? "<div class='solta'>clique para abrir a ficha da TAG</div>" : '') +
    (r.como && r.como !== 'exato'
      ? `<div class='obs'>de-para ${r.como}: a base de TAGs grava
         <b>${esc((r.tags || []).join(' + '))}</b></div>` : '');
}

const dicaCaixa = (nome, mont, extra) => cab('caixa de junção', nome) +
  dl('montagem', nomeMont(mont), corMont(mont)) + (extra || '');

// A calha não é um circuito da planilha, mas carrega os ramais que descem
// dela -- então a cor é o consolidado deles. Cinza diria "não se sabe".
function tomGrupo(l) {
  if (!l.length) return 'roxo';
  if (l.every(feito)) return 'ok';
  if (l.some(r => r.pct > 0)) return 'and';
  return 'nao';
}
function calha(d, l, w) {
  const n = l.filter(feito).length, t = tomGrupo(l);
  const dica = `<div class='h'>calha de derivação</div>
    <div class='n'>${n} de ${l.length} lançados</div>
    <div class='obs'>a calha não é um circuito da planilha: a cor dela é o
      consolidado dos ramais que descem daqui</div>`;
  return `<path d="${d}" stroke="var(--${t})" stroke-width="${w}" fill="none"
      stroke-opacity=".62" stroke-linecap="round"/>
    <path class="hit" d="${d}" stroke-width="${w + 11}" ${dd(dica)}/>`;
}

function bloco(x, y, w, h, p, extra, tom) {
  const dx = p * 0.62, dy = p * 0.4;
  const topo = `${x},${y} ${x + dx},${y - dy} ${x + w + dx},${y - dy} ${x + w},${y}`;
  const lado = `${x + w},${y} ${x + w + dx},${y - dy} ${x + w + dx},${y + h - dy} ${x + w},${y + h}`;
  // a cor entra como véu sobre as três faces, e só: quem define o estado é o
  // corpo da caixa. Cada face leva uma dose diferente para o relevo sobreviver.
  const veu = tom ? `<g fill="var(--${tom})">
      <polygon points="${topo}" opacity=".5"/><polygon points="${lado}" opacity=".26"/>
      <rect x="${x}" y="${y}" width="${w}" height="${h}" opacity=".38"/></g>` : '';
  return `<g><polygon class="eq-topo" points="${topo}"/>
    <polygon class="eq-lado" points="${lado}"/>
    <rect class="eq-face" x="${x}" y="${y}" width="${w}" height="${h}"/>${veu}
    <rect class="eq-l" x="${x}" y="${y}" width="${w}" height="${h}"/>${extra || ''}</g>`;
}

function painel(x, y, nome, indef) {
  const w = 96, h = 150, p = 26;
  let det = `<line class="eq-l" x1="${x + w / 2}" y1="${y}" x2="${x + w / 2}" y2="${y + h}"/>`;
  for (let i = 0; i < 5; i++)
    det += `<line class="eq-l" x1="${x + 8}" y1="${y + 16 + i * 7}" x2="${x + w / 2 - 8}" y2="${y + 16 + i * 7}"/>`;
  det += `<rect class="eq-l" x="${x + w / 2 + 9}" y="${y + 18}" width="${w / 2 - 18}" height="26" rx="2"/>`;
  det += `<circle cx="${x + w / 2 - 6}" cy="${y + h / 2}" r="2.5" fill="var(--t3)"/>`;
  det += `<circle cx="${x + w / 2 + 6}" cy="${y + h / 2}" r="2.5" fill="var(--t3)"/>`;
  det += `<rect class="eq-lado" x="${x - 4}" y="${y + h}" width="${w + 8}" height="7"/>`;
  const dica = indef
    ? cab('painel', 'sem tag') + dl('origem', 'desconhecida', 'roxo') +
      `<div class='obs'>nenhum circuito sai da caixa: não dá para saber a que painel
       ela chega. O painel continua no desenho, e a etapa fica inconclusiva.</div>`
    : cab('painel', nome) + dl('montagem', 'sem esse controle') +
      `<div class='obs'>painel não está na base de TAGs — é o controle que não
       existe, não um dado faltando</div>`;
  const corpo = `<g class="${indef ? 'indef' : ''}" ${dd(dica)}>${bloco(x, y, w, h, p, det)}</g>`;
  const rot = `<text class="eq-rot" x="${x + w / 2}" y="${y - 20}" text-anchor="middle">PAINEL</text>`;
  if (!indef)
    return corpo + rot +
      `<text class="eq-nm" x="${x + w / 2}" y="${y + h + 26}" text-anchor="middle">${esc(nome)}</text>`;
  return corpo + rot +
    `<text class="eq-int" x="${x + w / 2}" y="${y + h / 2 + 8}" text-anchor="middle">?</text>` +
    `<text class="eq-nm" x="${x + w / 2}" y="${y + h + 26}" text-anchor="middle"
      fill="var(--roxo)">sem tag</text>` +
    `<text class="eq-obs" x="${x + w / 2}" y="${y + h + 42}" text-anchor="middle">
      nenhum circuito sai da caixa</text>`;
}

// A caixa de junção de um segmento fieldbus: o tronco entra de um lado e sai
// do outro, e a última caixa da fila leva o terminador que fecha o segmento.
function caixaFF(x, y, nome, rot, fim, mont) {
  const w = 96, h = 66, p = 22;
  let det = `<rect class="eq-l" x="${x + 8}" y="${y + 8}" width="${w - 16}" height="${h - 24}" rx="2"/>`;
  det += `<circle cx="${x + 6}" cy="${y + 6}" r="1.6" fill="var(--t3)"/>
          <circle cx="${x + w - 6}" cy="${y + 6}" r="1.6" fill="var(--t3)"/>`;
  for (let i = 0; i < 4; i++)
    det += `<rect class="eq-lado" x="${x + 12 + i * 20}" y="${y + h}" width="8" height="6" rx="2"/>`;
  const t = fim ? `<g>
      <path d="M${x + w} ${y + 17.5} H${x + w + 13}" stroke="var(--ok)" stroke-width="2.4"/>
      <rect x="${x + w + 13}" y="${y + 6}" width="21" height="23" rx="3"
        fill="rgba(45,212,191,.15)" stroke="var(--ok)" stroke-width="1.5"/>
      <path d="M${x + w + 18} ${y + 12} h11 M${x + w + 18} ${y + 17.5} h11
               M${x + w + 20} ${y + 23} h7" stroke="var(--ok)" stroke-width="1.6"/>
      <text x="${x + w + 23}" y="${y + 41}" text-anchor="middle" fill="var(--ok)"
        font-size="6.5" font-weight="800">TERMINADOR</text></g>` : '';
  const dica = dicaCaixa(nome, mont,
    dl('papel', fim ? 'ponta do segmento, com terminador' : 'passagem do tronco'));
  return `<g ${dd(dica)}>${bloco(x, y, w, h, p, det, corMont(mont))}${t}</g>` +
    `<text class="eq-rot" x="${x + w / 2}" y="${y - 16}" text-anchor="middle">${rot}</text>` +
    `<text class="eq-nm" x="${x + w / 2}" y="${y + h + 24}" text-anchor="middle">${esc(nome)}</text>`;
}

function caixaJ(x, y, nome, rot, mont) {
  const w = 76, h = 62, p = 20;
  let det = `<rect class="eq-l" x="${x + 7}" y="${y + 8}" width="${w - 14}" height="${h - 24}" rx="2"/>`;
  for (let i = 0; i < 4; i++)
    det += `<rect class="eq-lado" x="${x + 9 + i * 16}" y="${y + h}" width="8" height="6" rx="2"/>`;
  det += `<circle cx="${x + 5}" cy="${y + 5}" r="1.6" fill="var(--t3)"/>
          <circle cx="${x + w - 5}" cy="${y + 5}" r="1.6" fill="var(--t3)"/>`;
  return `<g ${dd(dicaCaixa(nome, mont))}>${bloco(x, y, w, h, p, det, corMont(mont))}</g>` +
    `<text class="eq-rot" x="${x + w / 2}" y="${y - 16}" text-anchor="middle">${rot}</text>` +
    `<text class="eq-nm" x="${x + w / 2}" y="${y + h + 24}" text-anchor="middle">${esc(nome)}</text>`;
}

// anel e miolo na mesma cor: verde montado, vermelho não. Pintar só o miolo
// deixava a informação pequena demais numa grade de 60 símbolos.
function transmissor(x, y, mont) {
  const c = corMont(mont);
  return `<g><rect class="eq-lado" x="${x - 2}" y="${y + 11}" width="4" height="10"/>
    <circle cx="${x}" cy="${y}" r="10" fill="var(--${c})" fill-opacity=".2"
      stroke="var(--${c})" stroke-width="3"/>
    ${mont ? `<circle cx="${x}" cy="${y}" r="4.8" fill="var(--${c})"/>`
           : `<text x="${x}" y="${y + 3.8}" text-anchor="middle" font-size="10"
              font-weight="800" fill="var(--roxo)">?</text>`}
    <rect x="${x - 5}" y="${y - 14.5}" width="10" height="4.5" rx="1.5" fill="var(--metal3)"/></g>`;
}

// O cartão de uma caixa na vista do painel: o corpo pintado pela montagem
// dela, e embaixo quanto do que pendura já tem cabo pronto.
function cartaoCaixa(x, y, b, aberto) {
  const w = 150, h = 74;
  const pct = b.inst ? b.cabo_ok / b.inst * 100 : 0;
  const tom = !b.inst ? 'roxo' : pct >= 99.5 ? 'ok' : pct > 0 ? 'and' : 'nao';
  // O laco NAO e uma TAG: nao tem montagem, nao esta na base de TAGs e nao
  // se instala. Tratado como caixa, ele saia roxo com "sem correspondente na
  // base de TAGs" -- cobrando a montagem de uma coisa que nao se monta. A
  // borda dele segue o cabo, que e o que o laco tem de proprio.
  const eLaco = !!b.ordem_fixa;
  const corBorda = eLaco ? tom : corMont(b.mont);
  const dica = cab(b.direto ? 'sem caixa · loop direto no painel' : 'caixa de junção', b.nome) +
    (eLaco ? '' : dl('montagem', nomeMont(b.mont), corMont(b.mont))) +
    dl('instrumentos', b.inst) +
    dl('com cabo pronto', b.inst ? `${b.cabo_ok} · ${br(pct, 0)}%` : '—', tom) +
    (b.caixas > 1 ? dl('caixas em série', b.caixas) : '') +
    dl('tronco', b.tronco.status + ' · ' + br(b.tronco.m) + ' m', cls(b.tronco.status)) +
    (b.inst ? `<div class='solta'>${aberto ? 'clique no − para fechar'
      : 'clique no + para ver os instrumentos'}</div>` : '');
  // o + fica no proprio cartao: e nele que a pergunta "o que tem aqui dentro"
  // nasce, e nao numa lista ao lado
  const sinal = b.inst ? `<g class="abre" data-abre="${esc(b.nome)}">
      <rect x="${x + w - 30}" y="${y + 12}" width="22" height="22" rx="7"
        fill="rgba(var(--rgb-tinta),.08)" stroke="var(--linha2)"/>
      <path d="M${x + w - 24} ${y + 23} h10 ${aberto ? '' : `M${x + w - 19} ${y + 18} v10`}"
        stroke="var(--t2)" stroke-width="2" stroke-linecap="round"/></g>` : '';
  return `<g class="bloco" ${dd(dica)}>
    <rect x="${x}" y="${y}" width="${w}" height="${h}" rx="10"
      fill="var(--metal2)" stroke="var(--${corBorda})" stroke-width="1.6"
      stroke-opacity="${aberto ? 1 : .7}"/>
    <rect x="${x}" y="${y}" width="${w}" height="4" rx="2" fill="var(--${corBorda})"/>
    <text x="${x + 11}" y="${y + 26}" fill="var(--t1)" font-size="11.5" font-weight="700"
      font-family="ui-monospace,Consolas,monospace">${esc(b.nome).slice(0, 16)}</text>
    <text x="${x + 11}" y="${y + 44}" fill="var(--t3)" font-size="9.5">${b.inst}
      ${b.inst === 1 ? 'instrumento' : 'instrumentos'}${b.caixas > 1 ? ' · ' + b.caixas + ' caixas' : ''}</text>
    <rect x="${x + 11}" y="${y + 52}" width="${w - 22}" height="5" rx="2.5"
      fill="rgba(var(--rgb-tinta),.10)"/>
    <rect x="${x + 11}" y="${y + 52}" width="${(w - 22) * pct / 100}" height="5" rx="2.5"
      fill="var(--${tom})"/>
    <text x="${x + 11}" y="${y + 68}" fill="var(--${tom})" font-size="9"
      font-weight="700">${b.cabo_ok} de ${b.inst} com cabo pronto</text>${sinal}</g>`;
}

// Na vista do painel o instrumento é um selo: a cor da montagem, a do cabo e a
// tag. O detalhe é da vista da caixa -- aqui o que importa é enxergar o conjunto.
function cartaoMini(x, y, t) {
  const cm = corMont(t.mont), cc = t.cabo ? 'ok' : t.pct > 0 ? 'and' : 'nao';
  // a cor e o cabo DO INSTRUMENTO; a cadeia inteira (o que a certificacao
  // exige) entra como linha propria, pra nao dizer que o cabo esta faltando
  // quando o que falta e a alimentacao do painel la em cima
  const dica = cab('instrumento', t.org) +
    dl('cabo do instrumento',
       t.status + (t.pct > 0 && t.pct < 100 ? ' · ' + br(t.pct, 1) + '%' : ''), cc) +
    (t.m ? dl('metragem', br(t.m_real) + ' de ' + br(t.m) + ' m') : '') +
    (t.cabo && t.cabo_cadeia === false
      ? dl('certificação', 'presa antes deste instrumento', 'and') : '') +
    dl('montagem', nomeMont(t.mont), cm) +
    "<div class='solta'>clique para abrir a ficha da TAG</div>";
  return `<g class="inst" data-ficha="${esc(t.ancora || '')}" ${dd(dica)}>
    <rect x="${x - 34}" y="${y - 20}" width="68" height="40" rx="7"
      fill="var(--metal2)" stroke="rgba(var(--rgb-tinta),.12)"/>
    <rect x="${x - 34}" y="${y - 20}" width="68" height="3.5" rx="1.75" fill="var(--${cc})"/>
    <circle cx="${x - 22}" cy="${y + 3}" r="5" fill="var(--${cm})" fill-opacity=".35"
      stroke="var(--${cm})" stroke-width="2"/>
    <text x="${x - 12}" y="${y + 6.5}" fill="var(--t2)" font-size="7"
      font-family="ui-monospace,Consolas,monospace">${esc(t.base || t.org).slice(0, 12)}</text></g>`;
}

// O segmento vai no rodape do cartao. So cresce quando ha o que escrever:
// TAG sem fieldbus continua com o cartao curto de antes.
function cartao(xi, yi, r) {
  const sel = (D.tag === r.org || D.tag === r.base) ? ' sel' : '';
  const alt = 52 + (r.fseg ? 11 : 0);
  const extra = r.fseg ? `<text x="${xi}" y="${yi + 26}" text-anchor="middle"
      class="chip-sub">${esc(r.fseg).slice(0, 14)}</text>` : '';
  return `<g class="inst${sel}" data-ficha="${esc(r.ancora || '')}" ${dd(dicaTag(r))}>
    <rect class="chip-bg" x="${xi - 36}" y="${yi - 30}" width="72" height="${alt}" rx="9"/>
    ${transmissor(xi, yi - 8, r.mont)}
    <text x="${xi}" y="${yi + 16}" text-anchor="middle" fill="var(--t2)" font-size="7.5"
      font-family="ui-monospace,Consolas,monospace">${esc(r.base || r.org).slice(0, 13)}</text>
    ${extra}</g>`;
}

// A malha corre ao lado do proprio cabo, deitada e bem pequena. E o que
// distingue um ramal do outro quando o filtro traz varias malhas no mesmo
// segmento: no cartao ela competiria com a TAG, no cabo ela nomeia o cabo.
function malhaDoCabo(x, y1, y2, r) {
  if (!r.malha) return '';
  // 15 px e nao 7: a onda do cabo tem amplitude de ate 9 px, e o texto colado
  // caia bem dentro do tracado. O proximo cabo so comeca 73 px adiante, entao
  // sobra folga -- e o texto, deitado, ocupa a altura da fonte na horizontal.
  const ym = (y1 + y2) / 2, xr = x + 15;
  return `<text class="rot-malha" transform="rotate(-90 ${xr} ${ym})"
    x="${xr}" y="${ym}" text-anchor="middle">${esc(r.malha).slice(0, 18)}</text>`;
}

// O cabo não corre reto: sobra comprimento e ele acompanha a bandeja. A onda
// sai perpendicular ao percurso -- deslocar só o Y sumia no trecho vertical.
function zigue(x1, y1, x2, y2, amp) {
  const dx = x2 - x1, dy = y2 - y1, len = Math.hypot(dx, dy) || 1;
  const n = Math.max(2, Math.min(7, Math.round(len / 70)));
  const a = amp === undefined ? Math.min(14, len / 12) : amp;
  const px = -dy / len, py = dx / len;
  let d = `M${x1} ${y1}`;
  for (let i = 1; i <= n; i++) {
    const t = i / n, tm = t - 0.5 / n, o = (i % 2 ? -a : a);
    d += ` Q${x1 + dx * tm + px * o} ${y1 + dy * tm + py * o} ${x1 + dx * t} ${y1 + dy * t}`;
  }
  return d;
}

// Todo trecho desenha o percurso inteiro apagado e por cima o estado: o não
// iniciado fica vermelho, e não cinza -- cinza lia como "não existe".
function fio(d, c, w, papel) {
  const alvo = `<path class="hit" d="${d}" stroke-width="${Math.max(w + 12, 15)}"
    ${dd(dicaCabo(c, papel))}/>`;
  const fundo = `<path d="${d}" stroke="rgba(var(--rgb-tinta),.10)" stroke-width="${w}"
    fill="none" stroke-dasharray="3 6" stroke-linecap="round"/>`;
  if (c.pct <= 0) return fundo + `<path d="${d}" stroke="var(--nao)" stroke-width="${w}"
    fill="none" stroke-dasharray="9 7" stroke-opacity=".8" stroke-linecap="round"/>` + alvo;
  const anda = c.status === 'Em Andamento';
  return fundo + `<path d="${d}" stroke="var(--${cls(c.status)})" stroke-width="${w}"
    fill="none" stroke-linecap="round" pathLength="100"
    stroke-dasharray="${anda ? '7 5' : c.pct + ' 100'}"
    class="${anda ? 'andando' : ''}"/>` + alvo;
}

function cena(c) {
  const R = c.ramais, T = c.tronco, E = c.eletrica, S = c.segmentos;
  const grupos = R.reduce((a, r) => {
    const k = r.seg || c.caixa; (a[k] = a[k] || []).push(r); return a; }, {});
  const maiorGrupo = Math.max(1, ...Object.values(grupos).map(l => l.length));
  const ff = c.tipo === 'cff';
  const colsMax = ff ? 3 : Math.min(12, Math.max(2, Math.ceil(Math.sqrt(maiorGrupo * 1.9))));
  const filMax = Math.ceil(maiorGrupo / colsMax);
  // O ramal cresce quando sobra tela e encolhe quando não sobra: numa caixa de
  // três instrumentos o cabo é o que há para olhar; numa de sessenta, cada
  // pixel gasto na descida vira fileira a mais de rolagem.
  const queda = maiorGrupo <= 2 ? 138 : maiorGrupo <= 4 ? 124
              : maiorGrupo <= 8 ? 110 : maiorGrupo <= 16 ? 94 : 78;
  const passoX = maiorGrupo <= 4 ? 104 : maiorGrupo <= 8 ? 94 : 82;
  const fila = queda + 40, yRail = 146;
  // o cartao cresce 11 px quando tem o rodape do segmento; a malha corre ao
  // lado do cabo e nao ocupa altura nenhuma
  const rot = R.some(r => r.fseg) ? 11 : 0;
  const ALT = ff ? 150 + yRail + queda + (filMax - 1) * fila + 76 + rot
                 : Math.max(460, 190 + 28 + queda + (filMax - 1) * fila + 76 + rot);
  const base = ff ? 150 : 190;
  const p = [], nos = [], xPain = 150, xPrim = 400;
  let xFim = xPrim + 200;

  if (c.tipo === 'painel') {
    // uma faixa por caixa: o cartao dela e, na mesma linha, os instrumentos que
    // penduram nela. Em grade, descobrir de qual caixa um instrumento vinha
    // exigia seguir o fio com o olho.
    const B = c.blocos, xCx = xPain + 240, xInst = xCx + 200, passo = 84;
    const porLinha = 12;
    // a faixa da alimentacao ocupa o topo; sem a folga, o primeiro bloco
    // subia por cima dos fios dela
    const topo = 56 + Math.max(0, (E.length - 1)) * 14;
    let y = topo;
    const espinha = [];
    B.forEach(b => {
      const aberto = abertos.has(b.nome);
      // Um loop YST com identidade fixa (CERT_LOOPS_YST) sempre desenha em
      // DUAS fileiras no diagrama de interligacao de verdade, dobrando bem
      // no meio -- nao numa linha so que so cresce pro lado. Uma caixa comum
      // continua na grade generica (porLinha), onde a ordem das colunas nao
      // significa nada.
      const fixa = b.ordem_fixa && b.tags.length > 3;
      const metadeA = fixa ? Math.floor(b.tags.length / 2) : b.tags.length;
      const nB = b.tags.length - metadeA;
      const fil = aberto ? (fixa ? 2 : Math.max(1, Math.ceil(b.tags.length / porLinha))) : 0;
      // SERPENTINA, como no diagrama: a 1a fileira corre para a direita e a
      // 2a volta para a esquerda. Com as duas no mesmo sentido, a dobra tinha
      // de atravessar o desenho inteiro na diagonal, cortando os cartoes --
      // com a volta, ela vira um trecho curto na ponta, que e o que o desenho
      // mostra.
      const colDe = k => !fixa ? k % porLinha
                       : k < metadeA ? k : (nB - 1) - (k - metadeA);
      const filDe = k => fixa ? (k < metadeA ? 0 : 1) : Math.floor(k / porLinha);
      // a fileira de baixo corre ao contrario: a seta e o fio saem pela
      // esquerda do cartao, nao pela direita
      const invertida = k => fixa && k >= metadeA;
      // No laco o cabo entre instrumentos E a informacao -- no diagrama ele
      // corre solto, com o nome do circuito escrito em cima. Com o passo de
      // 84 os cartoes (68 de largura) ficavam a 16 px um do outro e o cabo
      // virava um toco de 9 px, sem espaco para o nome: amontoado e mudo.
      const pas = fixa ? 168 : passo;
      const meio = y + 37;
      espinha.push(meio);
      p.push(fio(zigue(xPain + 106, meio, xCx, meio, 6), b.tronco, 2.6,
                 'tronco painel → caixa'));
      p.push(cartaoCaixa(xCx, y, b, aberto));
      if (aberto) {
        // a calha corre ACIMA dos cartões e cada TAG desce dela. Correndo na
        // altura deles, o cabo de um atravessava todos os anteriores -- é o
        // mesmo defeito que a vista da caixa já tinha tido.
        const yCalha = y + 22, passoFil = 76;
        // No laco YST a ligacao e TAG -> TAG, como no desenho de
        // interligacao: a linha sai de um instrumento e entra no seguinte, na
        // mesma altura. A calha por cima com uma descida para cada cartao e
        // uma leitura de barramento -- certa para uma caixa de juncao, onde
        // os ramais sao independentes, e errada para o laco, onde o cabo
        // passa de um instrumento ao outro em serie.
        if (!fixa) {
          for (let f = 0; f < fil; f++) {
            const fatia = b.tags.slice(f * porLinha, (f + 1) * porLinha);
            p.push(calha(`M${xInst - 14} ${yCalha + f * passoFil}
              H${xInst + (fatia.length - 1) * pas + 36}`, fatia, 2.4));
          }
          if (fil > 1)
            p.push(calha(`M${xInst - 14} ${yCalha} V${yCalha + (fil - 1) * passoFil}`,
                         b.tags, 2.4));
          p.push(calha(`M${xCx + 150} ${meio} H${xInst - 14} V${yCalha}`, b.tags, 2.4));
        } else {
          // a entrada do painel chega na primeira TAG da fileira de cima,
          // num percurso continuo, sem barramento no meio
          const xPri = xInst + colDe(0) * pas + 36, yPri = yCalha + 44;
          p.push(`<path d="M${xCx + 150} ${meio} H${xInst - 30} V${yPri} H${xPri - 34}"
            stroke="var(--${b.tags[0].cabo ? 'ok' : b.tags[0].pct > 0 ? 'and' : 'nao'})"
            stroke-width="2.2" fill="none" stroke-linejoin="round"/>`);
        }
        b.tags.forEach((t, k) => {
          const cl = colDe(k), fl = filDe(k);
          const xi = xInst + cl * pas + 36, yr = yCalha + fl * passoFil;
          const yi = yr + 44;
          if (!fixa)
            p.push(`<path d="M${xi} ${yr} V${yi - 20}" stroke="var(--${
              t.cabo ? 'ok' : t.pct > 0 ? 'and' : 'nao'})" stroke-width="2"
              fill="none" stroke-linecap="round"/>`);
          p.push(cartaoMini(xi, yi, t));
          xFim = Math.max(xFim, xi + 46);
          // num loop de instrumento a instrumento (sem caixa), o fio entre um
          // cartao e o seguinte e o circuito real da base de cabos -- mesma
          // metragem, status e percentual que a busca por caixa mostra, so
          // que na ordem fisica do diagrama de interligacao (A -> B -> ...),
          // nao alfabetica.
          //
          // A ligacao e so a linha, como no desenho tecnico: sai da borda de
          // um instrumento e entra na borda do seguinte. Sem seta -- o
          // desenho de interligacao nao tem nenhuma, e a ordem ja esta na
          // sequencia dos cartoes.
          if (b.direto && k < b.tags.length - 1) {
            const flProx = filDe(k + 1);
            const xProx = xInst + colDe(k + 1) * pas + 36;
            if (flProx === fl) {
              // mesma fileira: a linha liga as duas bordas, na mesma altura
              const dir = xProx > xi ? 1 : -1;
              const de = xi + 34 * dir, ate = xProx - 34 * dir;
              if (t.circ_prox)
                p.push(fio(`M${de} ${yi} H${ate}`, t.circ_prox, 2.2,
                           'cabo · ' + esc(t.org) + ' → próximo do laço'));
              else
                p.push(`<path d="M${de} ${yi} H${ate}" stroke="var(--t3)"
                  stroke-width="1.4" fill="none" opacity=".6"/>`);
              // o nome do circuito em cima do trecho, como no diagrama de
              // interligacao -- e o que permite conferir a tela contra a
              // planilha sem passar o mouse cabo por cabo
              if (fixa && t.circ_prox && t.circ_prox.id)
                p.push(`<text x="${(de + ate) / 2}" y="${yi - 8}" text-anchor="middle"
                  class="rot-circ">${esc(t.circ_prox.id).slice(0, 16)}</text>`);
            } else {
              // A DOBRA: fim da fileira de cima, comeco da de baixo. Como a
              // de baixo volta no sentido contrario, as duas pontas ficam na
              // mesma coluna -- a dobra e um trecho continuo descendo pela
              // lateral, igual ao desenho, e nao uma diagonal cruzando tudo.
              const yProx = yCalha + flProx * passoFil + 44;
              const xLado = Math.max(xi, xProx) + 52;
              const dDobra = `M${xi + 34} ${yi} H${xLado} V${yProx} H${xProx + 34}`;
              if (t.circ_prox)
                p.push(fio(dDobra, t.circ_prox, 2.2,
                           'cabo · ' + esc(t.org) + ' → próximo do laço (dobra)'));
              else
                p.push(`<path d="${dDobra}" stroke="var(--t3)" stroke-width="1.4"
                  fill="none" opacity=".6"/>`);
            }
          }
        });
        // um loop de verdade fecha nas duas pontas do painel -- nao e uma
        // corrente que termina no vazio. A volta pontilhada do ultimo
        // instrumento ate o proprio cartao do loop e so pra deixar isso
        // visivel, do jeito que o diagrama de interligacao desenha (a central
        // tem OUT e IN do mesmo loop, nao so uma saida).
        if (b.direto && b.tags.length > 1) {
          const ultimoI = b.tags.length - 1;
          const clU = colDe(ultimoI), flU = filDe(ultimoI);
          const xU = xInst + clU * pas + 36;
          const yU = yCalha + flU * passoFil + 44;
          // a volta vai ate o proprio eixo do painel (mesmo x da espinha
          // que junta as entradas), nao so ate o cartao do loop -- e o que
          // deixa visivel que o laco ENTRA e SAI do painel, os dois lados
          // que o desenho de interligacao mostra (terminais OUT e IN da
          // mesma central).
          // no laco a volta sai pela BORDA do ultimo instrumento e corre por
          // baixo ate o eixo do painel: percurso continuo, do mesmo jeito que
          // o desenho fecha o laco. Sem seta -- o desenho tecnico nao tem.
          const yVolta = yCalha + (fil - 1) * passoFil + 78;
          const dVolta = fixa
            ? `M${xU - 34} ${yU} H${xInst - 30} V${yVolta} H${xPain + 106} V${meio}`
            : `M${xU} ${yU + 20} V${yVolta} H${xPain + 106} V${meio}`;
          if (b.circ_saida)
            p.push(fio(dVolta, b.circ_saida, 2, 'cabo · laço → painel'));
          else
            p.push(`<path d="${dVolta}" stroke="var(--t3)" stroke-width="1.5" fill="none"
              stroke-dasharray="1 5" stroke-linecap="round" opacity=".55"/>`);
          if (!fixa)
            p.push(`<polygon points="${xPain + 106},${meio} ${xPain + 100},${meio - 6} ${xPain + 112},${meio - 6}"
              fill="var(--t3)" opacity=".55"/>`);
        }
      }
      y += aberto ? Math.max(96, 22 + fil * 76 + 34 + (b.direto && b.tags.length > 1 ? 14 : 0)) : 96;
    });
    p.push(painel(xPain, topo, c.painel, false));
    if (espinha.length > 1)
      p.push(`<path d="M${xPain + 106} ${espinha[0]} V${espinha[espinha.length - 1]}"
        stroke="rgba(var(--rgb-tinta),.16)" stroke-width="3" fill="none"/>`);
    // A alimentacao do painel entra no desenho: e ela que trava o laco YST
    // inteiro (todos os trechos lancados e mesmo assim "Bloqueado"). Sem
    // desenhar, o vermelho ficava so na palavra, sem trecho vermelho nenhum.
    if (E.length) {
      E.forEach((e, i) => {
        const yA = 26 + i * 14;
        p.push(fio(`M40 ${yA} H${xPain + 48} V${topo}`, e, 2.4, 'alimentação do painel'));
      });
      const travada = E.some(e => e.pct < 99.5);
      p.push(`<text x="40" y="16" fill="var(--${travada ? 'nao' : 'ok'})" font-size="9"
        font-weight="800" letter-spacing=".7">ALIMENTAÇÃO DO PAINEL${
        travada ? ' · PENDENTE' : ''}</text>`);
    } else {
      p.push(`<path d="M40 26 H${xPain + 48} V${topo}" stroke="var(--roxo)"
        stroke-width="2" fill="none" stroke-dasharray="2 6" opacity=".55"/>`);
      p.push(`<text x="40" y="16" fill="var(--roxo)" font-size="9" font-weight="800"
        letter-spacing=".7">ALIMENTAÇÃO · SEM CIRCUITO CADASTRADO</text>`);
    }
    return `<svg class="cena" viewBox="0 0 ${Math.max(xFim + 40, 1100)} ${y + 30}" role="img"
      aria-label="Segmento do painel ${esc(String(c.painel))}: ${B.length} bloco${B.length === 1 ? '' : 's'} e
      ${B.reduce((a, b) => a + b.inst, 0)} instrumentos">${p.join('')}</svg>`;
  }

  if (E.length) {
    E.forEach((e, i) => p.push(fio(`M40 ${34 + i * 16} H${xPain + 48} V${base - 34}`,
                                   e, 2.4, 'alimentação da Elétrica')));
    p.push(`<text x="40" y="22" fill="var(--and)" font-size="9" font-weight="800"
      letter-spacing=".7">ELÉTRICA · ALIMENTAÇÃO</text>`);
  } else {
    // sem circuito cadastrado não há o que desenhar, e calar seria pior: a
    // etapa existe no campo mesmo faltando na planilha
    p.push(`<path d="M40 40 H${xPain + 48} V${base - 34}" stroke="var(--roxo)"
      stroke-width="2" fill="none" stroke-dasharray="2 6" opacity=".55"/>`);
    p.push(`<text x="40" y="28" fill="var(--roxo)" font-size="9" font-weight="800"
      letter-spacing=".7">ELÉTRICA · SEM CIRCUITO CADASTRADO</text>`);
  }
  p.push(painel(xPain, base - 34, c.painel, c.painel_indef));

  if (ff) {
    // UM segmento, e as caixas de junção dele em série: o tronco sai do cartão
    // H1 no painel, atravessa cada caixa e morre no terminador.
    const passo = 430, largura = 96, yTk = base + 17;
    p.push(`<text x="${xPain + 48}" y="${base + 130}" text-anchor="middle"
      class="rot-tk" font-size="7.5">CARTÃO H1 · SEGMENTO ${esc(c.caixa)}</text>`);
    S.forEach((sg, i) => {
      const x = xPrim + i * passo;
      p.push(caixaFF(x, base, sg.nome, 'CAIXA ' + sg.nome.slice(-1),
                     sg.nome === c.terminador, sg.mont));
      const de = i === 0 ? xPain + 96 : xPrim + (i - 1) * passo + largura + 4;
      const cabo = i === 0 ? T[0] : c.ligacoes[i - 1];
      if (cabo) {
        p.push(fio(zigue(de, yTk, x, yTk), cabo, 5,
                   i === 0 ? 'tronco · painel → caixa A' : 'tronco entre caixas'));
        p.push(`<text x="${(de + x) / 2}" y="${base - 4}" text-anchor="middle"
          class="rot-tk" font-size="7.5">TRONCO · ${br(cabo.m)} m</text>`);
      }
      const lista = grupos[sg.nome] || [];
      if (!lista.length) return;
      const cols = Math.min(3, lista.length), fil = Math.ceil(lista.length / cols);
      const px = passoX + 6, yBar = base + yRail, xDesce = x + 48;
      // a grade nasce centrada sob a caixa: com uma coluna só, o cartão fica
      // embaixo da descida em vez de 44 px ao lado dela
      const xIni = xDesce - (cols - 1) * px / 2;
      p.push(calha(`M${xDesce} ${base + 98} V${yBar}`, lista, 3.4));
      if (fil > 1) p.push(calha(`M${xIni} ${yBar} V${yBar + (fil - 1) * fila}`, lista, 2.8));
      for (let f = 0; f < fil; f++) {
        const fatia = lista.slice(f * cols, (f + 1) * cols);
        // a calha vai da descida ate a ultima coluna: quando os dois coincidem
        // ela some, e era esse o buraco entre a caixa e o cabo
        const a = Math.min(xIni, xDesce), z2 = Math.max(xIni + (fatia.length - 1) * px, xDesce);
        p.push(calha(`M${a} ${yBar + f * fila} H${z2}`, fatia, 2.8));
      }
      lista.forEach((r, k) => {
        const fl = Math.floor(k / cols);
        const xi = xIni + (k % cols) * px, yi = yBar + fl * fila + queda;
        p.push(fio(zigue(xi, yBar + fl * fila, xi, yi - 25, Math.min(9, queda / 12)),
                   r, 2, 'ramal até o instrumento'));
        p.push(malhaDoCabo(xi, yBar + fl * fila, yi - 25, r));
        p.push(cartao(xi, yi, r));
        xFim = Math.max(xFim, xi + 60);
      });
    });
    xFim = Math.max(xFim, xPrim + (S.length - 1) * 430 + 195);
  } else {
    const x = xPrim + 120;
    p.push(caixaJ(x, base, c.caixa, 'CAIXA DE JUNÇÃO', c.mont));
    nos.push({nome: c.caixa, x: x + 38, y: base});
    if (!T.length) {
      p.push(`<path d="M${xPain + 96} ${base + 30} H${x}" stroke="var(--roxo)"
        stroke-width="4" fill="none" stroke-dasharray="2 7" opacity=".6"/>`);
      p.push(`<text x="${(xPain + 96 + x) / 2}" y="${base + 18}" text-anchor="middle"
        class="eq-obs">tronco não cadastrado</text>`);
    } else {
      // vários cabos no mesmo tronco: a onda tem de caber no vão entre eles,
      // senão os traços se cruzam e o feixe vira uma escada
      const passo = Math.min(17, 68 / Math.max(T.length, 1));
      T.forEach((t, i) => {
        const y = base + 30 - (T.length - 1) * passo / 2 + i * passo;
        p.push(fio(zigue(xPain + 96, y, x, y, Math.min(7, passo * 0.32)),
                   t, T.length > 6 ? 3 : 5, 'tronco · painel → caixa'));
      });
    }
    nos.forEach(no => {
      const lista = grupos[no.nome] || [];
      if (!lista.length) return;
      const colunas = Math.min(12, Math.max(2, Math.ceil(Math.sqrt(lista.length * 1.9))));
      const fileiras = Math.ceil(lista.length / colunas);
      const x0 = no.x + 96, y0 = no.y + 28, xBar = x0 - 40;
      // a bandeja sai da caixa, desce por trás das fileiras e serve uma calha
      // para cada uma. O ramal colorido é só o trecho do instrumento.
      p.push(calha(`M${no.x + 40} ${no.y + 28} H${xBar}`, lista, 3.4));
      p.push(calha(`M${xBar} ${y0} V${y0 + (fileiras - 1) * fila}`, lista, 2.8));
      for (let f = 0; f < fileiras; f++) {
        const fatia = lista.slice(f * colunas, (f + 1) * colunas);
        p.push(calha(`M${xBar} ${y0 + f * fila}
                      H${Math.max(x0 + (fatia.length - 1) * passoX + 26, xBar + 30)}`,
                     fatia, 2.8));
      }
      lista.forEach((r, i) => {
        const xi = x0 + (i % colunas) * passoX + 26;
        const yi = y0 + Math.floor(i / colunas) * fila + queda;
        p.push(fio(zigue(xi, y0 + Math.floor(i / colunas) * fila, xi, yi - 25,
                         Math.min(9, queda / 12)), r, 2, 'ramal até o instrumento'));
        p.push(malhaDoCabo(xi, y0 + Math.floor(i / colunas) * fila, yi - 25, r));
        p.push(cartao(xi, yi, r));
        xFim = Math.max(xFim, xi + 50);
      });
    });
  }
  return `<svg class="cena" viewBox="0 0 ${Math.max(xFim + 40, 1100)} ${ALT}" role="img"
    aria-label="Trajeto físico: a elétrica alimenta o painel ${esc(String(c.painel))},
    o tronco segue até ${esc(c.caixa)} e os instrumentos derivam das caixas">
    ${p.join('')}</svg>`;
}

function aplicarZoom() {
  const sv = document.querySelector('#cena svg');
  if (sv) sv.style.width = (zoom * 100) + '%';
  document.getElementById('zn').textContent = Math.round(zoom * 100) + '%';
}
function zoomar(f) { zoom = Math.min(6, Math.max(0.4, zoom * f)); aplicarZoom(); }
function zoomAjuste() { zoom = 1; aplicarZoom(); document.getElementById('zona').scrollTo(0, 0); }

// Levar a TAG escolhida para o centro da tela: é o que faz "pesquisei, achei"
// virar "estou olhando de perto", sem procurar o cartão no meio de sessenta.
function focar(fator) {
  const alvo = document.querySelector('.inst.sel');
  const z = document.getElementById('zona');
  if (!alvo) { zoomar(fator); return; }
  const antes = alvo.getBoundingClientRect(), zr = z.getBoundingClientRect();
  const cx = z.scrollLeft + antes.left - zr.left + antes.width / 2;
  const cy = z.scrollTop + antes.top - zr.top + antes.height / 2;
  const z0 = zoom;
  zoomar(fator);
  const k = zoom / z0;
  z.scrollTo({left: cx * k - z.clientWidth / 2, top: cy * k - z.clientHeight / 2,
              behavior: 'smooth'});
}

function redesenhar() {
  document.getElementById('cena').innerHTML = cena(D.cadeia);
  aplicarZoom();
}
redesenhar();
// Aproximar na TAG faz sentido numa caixa de sessenta instrumentos, onde
// achar o cartao no meio da grade e o problema. No laco YST, que ja vem
// sozinho na tela, os 160% cortavam o desenho: aparecia meia fileira e o
// resto ficava fora da moldura. Ali a vista inteira e o que interessa --
// o laco so diz alguma coisa visto por completo.
if (D.tag && !(D.cadeia.tipo === 'painel' && D.cadeia.blocos.length === 1))
  requestAnimationFrame(() => focar(1.6));

// A ficha da TAG e um modal da pagina de fora. Trocar a ancora seria o caminho
// natural, mas o iframe do Streamlit e sandbox sem allow-top-navigation: mexer
// no location do pai levanta SecurityError. O que ele permite e allow-same-origin
// -- entao a ficha abre pela classe que o CSS ja reconhece ao lado do :target.
// O que fecha a ficha e registrado assim que a moldura carrega, e nao na
// primeira ficha aberta: fechar nao pode depender de por onde ela abriu.
//
// Registra de novo a cada moldura, sem flag de "so uma vez": ouvinte
// duplicado nao faz mal (fechar() so tira uma classe que ja pode nao estar
// la), mas UMA tentativa que falhe silenciosamente -- por qualquer folego de
// timing no carregamento -- travava fechar ate o F5, ja que a flag global
// impedia qualquer moldura seguinte de tentar de novo pelo resto da sessao.
function armarFechamento() {
  try {
    const d = parent.document;
    const fechar = () => d.querySelectorAll('.fmodal-on')
      .forEach(m => m.classList.remove('fmodal-on'));
    // o X e o fundo sao <a href="#">: com :target isso bastava, com a classe
    // e preciso ouvir o clique
    d.addEventListener('click', ev => {
      if (ev.target.closest && ev.target.closest('.fmodal-bg, .fmodal-x')) fechar();
    }, true);
    d.addEventListener('keydown', ev => { if (ev.key === 'Escape') fechar(); });
    // Uma ficha aberta por ancora e outra aberta por classe empilham: as duas
    // sao fixed e ocupam a tela inteira, e a de cima engole o clique de fechar
    // da de baixo. Foi o que travava ao descer da TAG para a malha. Quando a
    // ancora muda, a que abriu por classe sai de cena.
    parent.addEventListener('hashchange', fechar);
  } catch (_) { /* sem acesso ao pai, a ficha nao abre por aqui */ }
}
armarFechamento();

function abrirFicha(id) {
  try {
    const d = parent.document;
    const alvo = d.getElementById(id);
    if (!alvo) return false;
    d.querySelectorAll('.fmodal-on').forEach(m => m.classList.remove('fmodal-on'));
    alvo.classList.add('fmodal-on');
    return true;
  } catch (_) {
    return false;   // sem acesso ao pai, o clique volta a prender a dica
  }
}

(function () {
  const b = document.getElementById('dica'), z = document.getElementById('zona');
  let fixado = null;

  function posicionar(cx, cy) {
    // a dica vira de lado quando encosta na borda, em vez de sair da tela
    const x = cx + 18 + b.offsetWidth > innerWidth - 10
            ? Math.max(8, cx - b.offsetWidth - 18) : cx + 18;
    const y = cy + 18 + b.offsetHeight > innerHeight - 10
            ? Math.max(8, cy - b.offsetHeight - 18) : cy + 18;
    b.style.transform = `translate(${x}px,${y}px)`;
  }
  function mostrar(alvo, cx, cy, preso) {
    b.innerHTML = alvo.getAttribute('data-d')
      + (preso ? "<div class='solta'>clique fora para soltar</div>" : '');
    b.classList.add('ver');
    b.classList.toggle('fixo', !!preso);
    posicionar(cx, cy);
  }
  function soltar() {
    if (fixado) fixado.classList.remove('fixo');
    fixado = null;
    b.classList.remove('ver', 'fixo');
  }

  addEventListener('mousemove', e => {
    if (fixado) return;               // com a dica presa, ela não persegue o mouse
    const alvo = e.target.closest && e.target.closest('[data-d]');
    if (!alvo) { b.classList.remove('ver'); return; }
    mostrar(alvo, e.clientX, e.clientY, false);
  }, {passive: true});
  addEventListener('mouseleave', () => { if (!fixado) b.classList.remove('ver'); });
  // o Esc precisa valer dos dois lados: depois de clicar num cartao o foco fica
  // aqui dentro, e o ouvinte que registrei na pagina de fora nao recebe a tecla
  addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    soltar();
    try {
      parent.document.querySelectorAll('.fmodal-on')
        .forEach(m => m.classList.remove('fmodal-on'));
    } catch (_) { /* sem acesso ao pai, so a dica solta */ }
  });
  // Ctrl+roda amplia; sem o Ctrl a roda continua rolando, que é o que se espera
  z.addEventListener('wheel', e => {
    if (!e.ctrlKey && !e.metaKey) return;
    e.preventDefault();
    zoomar(e.deltaY < 0 ? 1.14 : 1 / 1.14);
  }, {passive: false});
  let a = null;
  z.addEventListener('pointerdown', e => {
    // o alvo tem de ser lido agora: com setPointerCapture, o pointerup passa a
    // chegar na moldura e o cartao sob o dedo se perde
    a = {x: e.clientX, y: e.clientY, l: z.scrollLeft, t: z.scrollTop, andou: false,
         alvo: e.target.closest
               && (e.target.closest('[data-abre]') || e.target.closest('[data-d]'))};
    z.classList.add('arrasta'); z.setPointerCapture(e.pointerId);
  });
  z.addEventListener('pointermove', e => {
    if (!a) return;
    // 4 px de folga: sem isso o tremor da mão vira arrasto e engole o clique
    if (Math.abs(e.clientX - a.x) > 4 || Math.abs(e.clientY - a.y) > 4) a.andou = true;
    z.scrollLeft = a.l - (e.clientX - a.x);
    z.scrollTop = a.t - (e.clientY - a.y);
  });
  z.addEventListener('pointerup', e => {
    const clicou = a && !a.andou, alvo = a && a.alvo;
    a = null; z.classList.remove('arrasta');
    if (!clicou) return;
    if (!alvo) { soltar(); return; }
    const abre = alvo.getAttribute && alvo.getAttribute('data-abre');
    if (abre) {
      abertos.has(abre) ? abertos.delete(abre) : abertos.add(abre);
      soltar(); redesenhar(); return;
    }
    const ficha = alvo.getAttribute('data-ficha');
    if (ficha && abrirFicha(ficha)) { soltar(); return; }
    if (alvo === fixado) { soltar(); return; }
    soltar();
    fixado = alvo;
    mostrar(alvo, e.clientX, e.clientY, true);
  });
  z.addEventListener('pointercancel', () => { a = null; z.classList.remove('arrasta'); });
})();
"""


def cert_cena_html(cadeia: dict, tag: str, tema: str, altura: int) -> str:
    """O documento do iframe: tokens do tema, o desenho e os controles.

    O iframe não herda o CSS do app, então os tokens entram copiados do mesmo
    dicionário TEMAS -- uma fonte só, duas superfícies.
    """
    t = TEMAS.get(tema, TEMAS[TEMA_PADRAO])
    dados = json.dumps({"cadeia": cadeia, "tag": tag}, ensure_ascii=False)
    return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{--bg:{t['card']};--card2:{t['card2']};--fundo3:{t['fundo3']};
  --linha:{t['borda']};--linha2:{t['borda_forte']};--rgb-tinta:{t['rgb_tinta']};
  --t1:{t['texto1']};--t2:{t['texto2']};--t3:{t['texto3']};
  --ok:{t['teal']};--and:{t['ambar']};--nao:{t['vermelho']};
  --azul:{t['azul']};--roxo:{t['roxo']};
  --metal:{t['metal']};--metal2:{t['metal2']};--metal3:{t['metal3']};
  color-scheme:{t['esquema']}}}
body{{background:var(--bg);color:var(--t1);overflow:hidden;
  font:400 13px/1.5 ui-sans-serif,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
  font-variant-numeric:tabular-nums;-webkit-font-smoothing:antialiased}}
.barra{{display:flex;align-items:center;gap:6px;padding:8px 12px;
  border-bottom:1px solid var(--linha)}}
.tt{{font-size:11.5px;color:var(--t2);font-weight:600;margin-right:auto}}
.tt b{{color:var(--t1);font-family:ui-monospace,Consolas,monospace}}
.zb{{height:26px;min-width:27px;padding:0 7px;border-radius:8px;background:var(--card2);
  border:1px solid var(--linha2);color:var(--t2);font:inherit;font-size:13px;
  font-weight:700;cursor:pointer;display:grid;place-items:center;line-height:1}}
.zb:hover{{border-color:var(--azul);color:var(--t1)}}
.zb.tx{{font-size:10px;letter-spacing:.4px}}
.zn{{font-size:11px;color:var(--t3);min-width:42px;text-align:center;font-weight:600}}
/* a cena rola dentro da moldura; sem isso o zoom empurraria a página inteira */
/* o desenho fica centrado na moldura: quando a cena e mais baixa que o
   espaco (laco largo e raso, por exemplo), a sobra se divide em cima e
   embaixo em vez de virar um vao escuro so no pe. "safe" para que, quando a
   cena for MAIOR que a moldura, ela continue rolando do inicio em vez de
   ficar com o topo cortado. */
.zona{{overflow:auto;height:calc(100vh - 43px);cursor:grab;overscroll-behavior:contain;
  display:flex;flex-direction:column;justify-content:safe center}}
.zona.arrasta{{cursor:grabbing}}
svg.cena{{width:100%;height:auto;display:block}}
.eq-face{{fill:var(--metal)}}.eq-topo{{fill:var(--metal3)}}.eq-lado{{fill:var(--metal2)}}
.indef .eq-face,.indef .eq-topo,.indef .eq-lado{{opacity:.42}}
.indef .eq-l{{stroke:rgba(var(--rgb-roxo,157,107,255),.5);stroke-dasharray:4 4}}
.eq-obs{{fill:var(--roxo);font-size:9px;font-weight:700}}
.eq-int{{fill:var(--roxo);font-size:20px;font-weight:800}}
.eq-l{{stroke:rgba(var(--rgb-tinta),.18);stroke-width:1;fill:none}}
.eq-nm{{fill:var(--t1);font-size:11.5px;font-weight:700;
  font-family:ui-monospace,Consolas,monospace}}
.eq-rot{{fill:var(--t3);font-size:8.5px;letter-spacing:.7px;font-weight:700}}
.rot-tk{{fill:var(--and);font-weight:800;letter-spacing:.5px}}
.chip-bg{{fill:var(--metal2);stroke:rgba(var(--rgb-tinta),.14)}}
/* Segmento e malha usam a mesma fonte: sao a mesma classe de informacao, e
   duas escalas de letra no mesmo desenho leem como hierarquia que nao existe. */
.chip-sub,.rot-malha{{fill:var(--t3);font-size:6.4px;letter-spacing:.2px;
  font-family:ui-monospace,Consolas,monospace}}
/* o nome do circuito sobre o trecho do laco: o mesmo papel que ele tem no
   diagrama de interligacao, identificar o cabo daquele trecho */
.rot-circ{{fill:var(--t3);font-size:7px;letter-spacing:.2px;opacity:.85;
  font-family:ui-monospace,Consolas,monospace}}
/* So a malha leva contorno: ela e escrita sobre o desenho, e o cabo ondula
   justamente por onde ela passa. O contorno na cor do fundo abre um vao em
   volta da letra -- sem ele o texto some dentro do tracado. paint-order manda
   pintar o contorno ANTES do preenchimento, senao a borda come a letra por
   dentro. O segmento nao precisa: mora no cartao, que ja e chapado. */
.rot-malha{{paint-order:stroke;stroke:var(--bg);stroke-width:2.4px;
  stroke-linejoin:round}}
.inst.sel .chip-bg{{stroke:var(--azul);stroke-width:2.5}}
.inst:hover .chip-bg{{stroke:var(--t2)}}
.hit{{stroke:transparent;fill:none}}
.inst,.abre{{cursor:pointer}}
.abre:hover rect{{stroke:var(--azul)}}
.inst.fixo .chip-bg{{stroke:var(--t2);stroke-width:2;stroke-dasharray:4 3}}
.dica.fixo{{border-color:var(--t3)}}
.dica .solta{{font-size:9.5px;color:var(--t3);margin-top:7px;padding-top:6px;
  border-top:1px solid var(--linha);letter-spacing:.3px}}
.andando{{animation:corre 1.05s linear infinite}}
@keyframes corre{{to{{stroke-dashoffset:-16}}}}
@media (prefers-reduced-motion:reduce){{.andando{{animation:none}}}}
.dica{{position:fixed;z-index:60;pointer-events:none;opacity:0;transition:opacity .09s;
  background:var(--fundo3);border:1px solid var(--linha2);border-radius:11px;
  padding:10px 13px;max-width:310px;box-shadow:0 12px 34px rgba(0,0,0,.4);left:0;top:0}}
.dica.ver{{opacity:1}}
.dica .h{{font-size:9px;letter-spacing:.8px;text-transform:uppercase;font-weight:800;
  color:var(--t3)}}
.dica .n{{font-family:ui-monospace,Consolas,monospace;font-size:13px;font-weight:700;
  margin:3px 0 7px;color:var(--t1)}}
.dica .l{{display:flex;gap:14px;justify-content:space-between;font-size:11.5px;padding:2.5px 0}}
.dica .l i{{font-style:normal;color:var(--t3)}}
.dica .l b{{font-weight:700;text-align:right}}
.dica .obs{{font-size:10.5px;color:var(--t3);margin-top:6px;padding-top:6px;
  border-top:1px solid var(--linha);line-height:1.45}}
</style></head><body>
<div class="dica" id="dica"></div>
<div class="barra">
  <span class="tt">passe o mouse para ver a situação · <b>+</b> abre a caixa ·
    clique na TAG para abrir a ficha · arraste para deslocar · Ctrl + roda amplia</span>
  <button class="zb" onclick="focar(1/1.3)" title="Afastar">−</button>
  <span class="zn" id="zn">100%</span>
  <button class="zb" onclick="focar(1.3)" title="Aproximar">+</button>
  <button class="zb tx" onclick="focar(2.2)" title="Aproximar na TAG selecionada">NA TAG</button>
  <button class="zb tx" onclick="zoomAjuste()" title="Voltar à largura da tela">AJUSTAR</button>
</div>
<div class="zona" id="zona"><div id="cena"></div></div>
<script>{CERT_JS.replace('__DADOS__', dados)}</script>
</body></html>"""


def cert_onde(trava, aberto: bool) -> str:
    """Onde a cadeia para, escrito do mesmo jeito que a lista de predecessoras.

    O texto vem do papel do trecho, não do nome da coluna: quem lê quer saber
    se falta a alimentação, o tronco ou o ramal.
    """
    if trava is None:
        return "predecessora sem circuito cadastrado"
    org, dst = str(trava["ORIGEM"]).strip(), str(trava["DESTINO"]).strip()
    # o circuito de alimentação de um painel para outro (ou pra um
    # dispositivo como um CBZ) às vezes vem em INSTRUMENTAÇÃO, não em
    # ELÉTRICA -- o "-P" no nome é quem marca isso, mesmo critério do campo
    # "pot" de _cert_circuito. Só conta quando quem envia já é painel: um
    # circuito de potência entre instrumentos não é alimentação de painel.
    eh_alimentacao = (str(trava["DISCIPLINA"]).strip() == "ELÉTRICA"
                      or (cert_nivel(org) == 2
                          and re.search(r"-P\d*$", str(trava["CIRCUITO"]).strip(), re.I)))
    if eh_alimentacao:
        return f"alimentação do painel {dst}"
    if cert_nivel(org) == 0:
        return "ramal até o instrumento"
    a, b = (dst, org) if cert_nivel(org) < cert_nivel(dst) else (org, dst)
    return f"tronco {a} → {b}".lower()


@st.cache_data(show_spinner=False, max_entries=3)
def cert_panorama(lanc: pd.DataFrame, mont: dict, cache_key: str) -> dict:
    """Quantas TAGs estão aptas, sem montar cadeia por cadeia.

    Percorrer as 480 caixas para responder um número do topo custaria mais que
    desenhar a tela. Aqui a leitura é de uma passada só: cada instrumento sobe
    pela própria origem até achar o painel, somando o que encontra pelo
    caminho, e a alimentação do painel entra no fim.
    """
    vazio = {"circuitos": 0, "metros": 0.0, "lancado": 0.0, "tags": 0, "caixas": 0,
             "cabo_apto": 0, "tag_apta": 0, "montadas_travadas": 0, "por_tag": {}}
    if lanc.empty:
        return vazio
    linhas = lanc.to_dict("records")
    saida: dict[str, list] = {}
    chegada: dict[str, list] = {}
    eletrica: dict[str, list] = {}
    por_circuito: dict[str, list] = {}
    for r in linhas:
        org, dst = str(r["ORIGEM"]).strip(), str(r["DESTINO"]).strip()
        saida.setdefault(org, []).append(r)
        if dst not in ("", "nan"):
            chegada.setdefault(dst, []).append(r)
        por_circuito.setdefault(str(r["CIRCUITO"]).strip(), []).append(r)
        if str(r["DISCIPLINA"]).strip() == "ELÉTRICA":
            eletrica.setdefault(dst, []).append(r)

    # A conta é por TAG, não por circuito: um instrumento com dois ramais
    # aparecia duas vezes no cartão do topo e uma só na tabela, e os dois
    # números brigavam na tela. Com dois ramais, vale a pior das duas cadeias --
    # certificar exige que todas fechem, não uma delas.
    PIOR = {"ok": 0, "desc": 1, "warn": 2, "crit": 3}

    # O painel que o diagrama de interligacao confirma para cada TAG de loop
    # YST. A planilha as vezes tem DUAS saidas para paineis diferentes na
    # ultima ponta do laco -- YST-121170 vai para PN-12-236 (o do desenho,
    # Concluido) e tambem para PN-12-239 (Nao Iniciado, painel que nem
    # alimentacao tem na base). Sem um criterio, quem decidia era a ordem das
    # linhas na planilha, e o laco inteiro travava no circuito errado.
    painel_conferido: dict[str, str] = {}
    for _painel_loop, _tags_loop in CERT_LOOPS_YST.values():
        for _t in _tags_loop:
            painel_conferido[_t] = _painel_loop

    def trilha(inicio: str, preferido: str | None):
        """Os circuitos do caminho que leva ao painel -- so os dele.

        Somar TODAS as saidas de cada no (o que esta funcao fazia antes)
        puxava para dentro da cadeia circuito que nem passa por ela: o LOOP4
        inteiro reprovava por causa de C-PN-12-239-01, um cabo de outro
        painel que nao esta no caminho de ninguem do laco.

        Cabo paralelo continua entrando inteiro -- varias linhas entre as
        MESMAS duas pontas (292 pares na base) sao cabos irmaos do mesmo
        trecho, e todos precisam fechar para o trecho fechar.
        """
        pilha = [(inicio, [], {inicio})]
        alternativo, passos = None, 0
        while pilha and passos < 300:
            passos += 1
            no, trecho, vistos = pilha.pop()
            if len(trecho) >= 15:
                continue
            # O SINAL decide a qual painel o instrumento pertence. Quase todo
            # instrumento tem dois cabos saindo dele: o de sinal, que vai ao
            # painel dele por uma caixa de junção/derivação (CJA/CJD), e o de
            # potência (-P), que vai ao painel de alimentação por uma caixa de
            # potência (CJP). Tratados como caminhos equivalentes, quem
            # escolhia o painel era a ordem das linhas na planilha -- em 47
            # TAGs (BSL, AST, AIT) o veredito mudava conforme o cabo que a
            # caminhada pegasse primeiro. A potência continua contando na
            # certificação; ela só não decide de quem o instrumento é.
            saltos = sorted({str(x["DESTINO"]).strip() for x in saida.get(no, [])})
            saltos.sort(key=lambda d: cert_so_potencia(saida.get(no, []), d))
            for prox_no in saltos:
                irmaos = [x for x in saida.get(no, [])
                          if str(x["DESTINO"]).strip() == prox_no]
                novo = trecho + irmaos
                if cert_nivel(prox_no) == 2:
                    so_pot = cert_so_potencia(saida.get(no, []), prox_no)
                    if (preferido is None or prox_no == preferido) and not so_pot:
                        return novo, prox_no
                    if alternativo is None:
                        alternativo = (novo, prox_no)
                    continue
                if prox_no not in vistos:
                    pilha.append((prox_no, novo, vistos | {prox_no}))
        return alternativo if alternativo else ([], None)

    def cabo_da_tag(tag: str) -> dict:
        """O cabo DO INSTRUMENTO, sem nada do que vem antes dele.

        É outra pergunta que a da certificação: "o cabo desta TAG foi
        lançado?" não é "tudo que alimenta esta TAG está pronto?". Misturar
        as duas numa coluna só fazia a tela dizer "faltando lançamento" para
        instrumento cujo cabo a planilha dá como Concluído -- o cabo estava
        lançado, o que faltava era a alimentação do painel, muitos saltos
        acima. Cada uma tem sua coluna agora.
        """
        # O cabo do instrumento e o circuito que leva o NOME dele: o cabo do
        # YST-121188 e o C-YST-121188. Procurar so entre as linhas em que ele
        # aparece como ORIGEM/DESTINO deixava de fora justamente o ultimo
        # ponto do laco, que so recebe cabo -- e a coluna dizia "sem
        # circuito" para um detector cujo cabo esta lancado.
        nomeado = [c for c in por_circuito.get(f"C-YST-{tag.split('-')[-1]}", [])
                   if not cert_remendo(c)]
        proprios = nomeado or [c for c in (saida.get(tag) or chegada.get(tag) or [])
                               if not cert_remendo(c)]
        if not proprios:
            return {"cabo_tag": False, "pct_tag": 0.0, "status_tag": "sem circuito",
                    "m_tag": 0.0, "m_real_tag": 0.0}
        m = sum(cert_num(c["METROS"]) for c in proprios)
        # medido sem teto: o metro a mais que o previsto e fato para ler
        m_real = sum(cert_metro_medido(c) for c in proprios)
        pct = round(m_real / m * 100, 1) if m else 0.0
        pronto = all(cert_num(c["PCT"]) >= 99.5 for c in proprios)
        return {"cabo_tag": pronto, "pct_tag": pct,
                # O status é o que a base informa, nunca deduzido da
                # metragem: um cabo Concluído com menos metro que o previsto
                # é normal (previu-se mais do que precisou), e um cabo com
                # mais metro que o previsto pode seguir Em Andamento. Deduzir
                # o status do percentual apagava esses dois casos, que são
                # justamente os que interessam de olhar.
                "status_tag": cert_status_conjunto(proprios),
                "m_tag": m, "m_real_tag": m_real}

    por_tag: dict[str, dict] = {}
    for r in linhas:
        org, dst = str(r["ORIGEM"]).strip(), str(r["DESTINO"]).strip()
        if org in ("", "nan") or cert_nivel(org) != 0:
            continue
        cadeia, aberto = [r], False
        if cert_nivel(dst) == 2:
            # instrumento ligado direto no painel: a cadeia e o proprio ramal
            atual = dst
        else:
            trecho, painel_achado = trilha(dst, painel_conferido.get(org))
            cadeia.extend(trecho)
            if painel_achado:
                atual = painel_achado
            else:
                # nenhum caminho chega em painel: beco sem saida ou laco que
                # gira em circulo (ex: YST-121125 <-> YST-121112)
                aberto = True
                atual = str(trecho[-1]["DESTINO"]).strip() if trecho else dst
        painel_da_cadeia = atual
        if cert_nivel(atual) == 2:
            alim = eletrica.get(atual)
            if not alim:
                # eletrica() só acha linha ELÉTRICA->painel; alguns painéis
                # (PN-12-236/237 -> PN-12102, PN-12-238 -> CBZ-12-001, achado
                # no diagrama YST em 2026-08-27) sobem pra outro painel ou
                # dispositivo com o circuito em INSTRUMENTAÇÃO, marcado só
                # pelo "-P" no nome -- o mesmo criterio que o campo "pot" de
                # _cert_circuito já usa. Sem isso a alimentação ficava sempre
                # "desconhecida" mesmo quando o circuito real existe na base.
                no = atual
                for _ in range(15):
                    pot = [r for r in saida.get(no, [])
                           if re.search(r"-P\d*$", str(r["CIRCUITO"]).strip(), re.I)]
                    if not pot:
                        break
                    alim = pot
                    no = str(pot[0]["DESTINO"]).strip()
            if alim:
                cadeia.extend(alim)
            else:
                aberto = True
        # A correcao de topologia (CORRECAO-*) nao e cabo pendente: e um
        # remendo nosso, para o grafo fechar onde a planilha ainda nao tem a
        # linha. Ela nasce sempre "Nao Iniciado / 0%", entao contada como
        # cabo de verdade ela reprovava o laco inteiro -- 18 TAGs com o cabo
        # real 100% lancado apareciam "Bloqueado" por causa dela. O que ela
        # diz de honesto e "falta o registro deste trecho", e isso ja tem
        # tom proprio: "desc" (nao da para afirmar), nunca "crit".
        reais = [c for c in cadeia if not cert_remendo(c)]
        # remendo que herdou o avanço do circuito real NÃO é falta de
        # registro: o cabo daquele trecho existe e está lançado, e só a
        # ligação é que o pipeline acrescentou. Contá-lo como buraco fazia o
        # laço inteiro cair para "não dá para afirmar" com todo o cabo pronto.
        sem_registro = any(cert_remendo(c) and cert_num(c["PCT"]) < 99.5
                           for c in cadeia)
        trava = next((c for c in reais if cert_num(c["PCT"]) < 99.5), None)
        if trava is None:
            tom = "desc" if (aberto or sem_registro) else "ok"
        elif str(trava["STATUS"]).strip() == "Em Andamento":
            tom = "warn"
        else:
            tom = "crit"
        # quando a caminhada para antes do painel, o ultimo no e uma caixa --
        # chamar isso de "alimentação do painel CJD-12-0702" era mentira, e a
        # tabela repetia a mentira 1.500 vezes
        if tom == "ok":
            onde = "—"
        elif trava is not None:
            onde = cert_onde(trava, aberto)
        elif sem_registro:
            onde = "trecho do laço sem circuito cadastrado na planilha de cabos"
        elif cert_nivel(painel_da_cadeia) == 2:
            onde = f"alimentação do painel {painel_da_cadeia}"
        else:
            onde = f"nenhum circuito sai de {painel_da_cadeia}: o painel é desconhecido"
        antes = por_tag.get(org)
        if antes is None or PIOR[tom] > PIOR[antes["tom"]]:
            montada = mont.get(org, {}).get("mont", "") == "Montado"
            por_tag[org] = {"tom": tom, "caixa": dst, "onde": onde,
                            "montada": montada,
                            # cabo apto e a cadeia inteira lancada; TAG apta e
                            # isso mais o instrumento no lugar
                            "cabo": tom == "ok", "apta": tom == "ok" and montada,
                            # a cadeia inteira, nao so o trecho travado -- e o
                            # que a exportacao de pendencias de lancamento usa
                            "cadeia": cadeia, **cabo_da_tag(org)}

    # Ponta que só aparece como DESTINO (nunca ORIGEM) não entra no loop acima
    # -- é o fim de um loop instrumento-a-instrumento (detecção de fumaça/gás,
    # por exemplo) ou um instrumento alimentado direto por um painel, sem
    # caixa no meio, na direção contrária do que o loop principal varre. Sem
    # isso a TAG some da Certificação mesmo tendo cabo lançado e cadastro
    # próprio -- foi o caso de 10 dos 88 TAGs do diagrama YST conferido em
    # 2026-08-27.
    for r in linhas:
        org, dst = str(r["ORIGEM"]).strip(), str(r["DESTINO"]).strip()
        if dst in ("", "nan") or cert_nivel(dst) != 0 or dst in saida:
            continue
        if cert_nivel(org) == 0:
            base = por_tag.get(org)
            if base is None:
                continue
            tom, onde, cabo = base["tom"], base["onde"], base["cabo"]
            cadeia = base.get("cadeia", [])
        elif cert_nivel(org) == 2:
            cadeia, aberto = [r], False
            alim = eletrica.get(org)
            if alim:
                cadeia.extend(alim)
            else:
                aberto = True
            trava = next((c for c in cadeia if cert_num(c["PCT"]) < 99.5), None)
            if trava is None:
                tom = "desc" if aberto else "ok"
            elif str(trava["STATUS"]).strip() == "Em Andamento":
                tom = "warn"
            else:
                tom = "crit"
            onde = ("—" if tom == "ok" else
                    cert_onde(trava, aberto) if trava is not None else
                    f"alimentação do painel {org}")
            cabo = tom == "ok"
        else:
            continue
        antes = por_tag.get(dst)
        if antes is None or PIOR[tom] > PIOR[antes["tom"]]:
            montada = mont.get(dst, {}).get("mont", "") == "Montado"
            por_tag[dst] = {"tom": tom, "caixa": org, "onde": onde,
                            "montada": montada,
                            "cabo": cabo, "apta": tom == "ok" and montada,
                            "cadeia": cadeia, **cabo_da_tag(dst)}

    # ------------------------------------------------------------------
    # Conclusão só aparece quando a lógica CONSEGUE concluir.
    #
    # 128 TAGs alcançam mais de um painel na base. Em 49 delas o veredito
    # muda conforme o caminho que se pega (BSL-120700A dá "apto" pelo
    # PN-12-264 e "em andamento" pelo PN-12102). Escolher um lado -- o que o
    # código fazia, pela ordem das linhas na planilha -- é inventar
    # engenharia: o número saía bonito ou feio por acaso.
    #
    # Onde não há conclusão, a tela não afirma nada: fica "não dá para
    # afirmar", dizendo qual é a ambiguidade. É assim que a falta vira
    # sinal de que há o que corrigir na base, em vez de virar um número
    # errado que ninguém questiona.
    def paineis_alcancaveis(tag, limite=12):
        """Os painéis que o SINAL deste instrumento alcança.

        A potência fica de fora: ela vai para o painel de alimentação, que
        quase nunca é o painel do instrumento, e contá-la aqui inventava
        ambiguidade em 47 TAGs que na verdade não têm nenhuma.
        """
        achados, pilha, vistos = set(), [(tag, 0)], {tag}
        while pilha:
            no, d = pilha.pop()
            if d > limite:
                continue
            for r in saida.get(no, []):
                dst = str(r["DESTINO"]).strip()
                if cert_so_potencia(saida.get(no, []), dst):
                    continue
                if cert_nivel(dst) == 2:
                    achados.add(dst)
                elif dst not in vistos:
                    vistos.add(dst)
                    pilha.append((dst, d + 1))
        return achados

    def _trilha_ate(inicio, alvo, limite=15):
        pilha = [(inicio, [], {inicio})]
        while pilha:
            no, tr, vis = pilha.pop()
            if len(tr) >= limite:
                continue
            for prox in sorted({str(x["DESTINO"]).strip() for x in saida.get(no, [])}):
                irmaos = [x for x in saida.get(no, [])
                          if str(x["DESTINO"]).strip() == prox]
                novo = tr + irmaos
                if prox == alvo:
                    return novo
                if cert_nivel(prox) != 2 and prox not in vis:
                    pilha.append((prox, novo, vis | {prox}))
        return None

    def _tom_por_painel(tag, painel):
        cadeia_p = None
        for r in saida.get(tag, []):
            dst = str(r["DESTINO"]).strip()
            if dst == painel:
                cadeia_p = [r]
                break
            tr = _trilha_ate(dst, painel)
            if tr is not None:
                cadeia_p = [r] + tr
                break
        if cadeia_p is None:
            return None
        alim = eletrica.get(painel)
        if not alim:
            no, vis = painel, set()
            for _ in range(15):
                if no in vis:
                    break
                vis.add(no)
                pot = [x for x in saida.get(no, [])
                       if re.search(r"-P\d*$", str(x["CIRCUITO"]).strip(), re.I)]
                if not pot:
                    break
                alim = pot
                no = str(pot[0]["DESTINO"]).strip()
        if alim:
            cadeia_p = cadeia_p + list(alim)
        reais_p = [c for c in cadeia_p if not cert_remendo(c)]
        trava_p = next((c for c in reais_p if cert_num(c["PCT"]) < 99.5), None)
        if trava_p is None:
            return "ok" if alim else "desc"
        return "warn" if str(trava_p["STATUS"]).strip() == "Em Andamento" else "crit"

    for tag, v in por_tag.items():
        # o loop YST tem o painel conferido no diagrama: ali a ambiguidade
        # da planilha já está resolvida, e a conclusão vale
        if tag in painel_conferido:
            continue
        alcancaveis = paineis_alcancaveis(tag)
        if len(alcancaveis) < 2:
            continue
        tons = {p: _tom_por_painel(tag, p) for p in sorted(alcancaveis)}
        distintos = {t for t in tons.values() if t}
        if len(distintos) < 2:
            continue
        v["tom"] = "desc"
        v["cabo"] = False
        v["apta"] = False
        v["ambiguo"] = sorted(p for p, t in tons.items() if t)
        v["onde"] = ("caminho indefinido na base: chega em "
                     + " e ".join(v["ambiguo"])
                     + " com resultados diferentes")

    previsto = lanc["METROS"].fillna(0)
    metros = float(previsto.sum())
    # o metro que o campo mediu, e não o proporcional do percentual: quando a
    # coluna existe ela é o número que vale
    proporcional = previsto * lanc["PCT"].fillna(0) / 100
    real = (lanc["METROS_REAL"].fillna(0) if "METROS_REAL" in lanc
            else pd.Series(0.0, index=lanc.index))
    # mesmo teto do cert_metro_real, aplicado circuito a circuito: o medido não
    # passa do previsto. Somar primeiro e capar no fim deixaria um circuito
    # medido a mais cobrir o atraso de outro. Sem previsto não há teto -- ali o
    # teto vira o próprio valor, e o clip não faz nada.
    valor = real.where(real > 0, proporcional)
    lancado = float(valor.clip(upper=previsto.where(previsto > 0, valor)).sum())
    pontas = cert_txt(pd.concat([lanc["ORIGEM"], lanc["DESTINO"]]))
    caixas = {p for p in pontas if p.startswith(CAIXA_PREFIXOS)}
    return {"circuitos": int(len(lanc)), "metros": metros, "lancado": lancado,
            "tags": len(por_tag), "caixas": len(caixas), "por_tag": por_tag,
            "cabo_apto": sum(1 for v in por_tag.values() if v["cabo"]),
            "tag_apta": sum(1 for v in por_tag.values() if v["apta"]),
            # o que espera cabo estando montado: e a fila que a obra consegue
            # destravar, ao contrario da que espera montagem
            "montadas_travadas": sum(1 for v in por_tag.values()
                                     if v["montada"] and not v["cabo"])}


def cert_altura(cad: dict, largura_px: int = 1320) -> int:
    """A altura da moldura, tirada da própria cena.

    O iframe tem altura fixa: com um valor único, uma cadeia de 3 instrumentos
    deixava meia tela de vazio e uma de 60 nascia cortada. As contas aqui são
    as mesmas do desenho -- se mudarem lá, mudam aqui, e é por isso que os
    números estão nomeados dos dois lados.
    """
    grupos: dict[str, list] = {}
    for r in cad["ramais"]:
        grupos.setdefault(r.get("seg") or cad["caixa"], []).append(r)
    maior = max((len(v) for v in grupos.values()), default=1)
    ff = cad["tipo"] == "cff"
    cols = 3 if ff else min(12, max(2, math.ceil(math.sqrt(maior * 1.9))))
    fileiras = math.ceil(maior / cols)
    queda = (138 if maior <= 2 else 124 if maior <= 4 else 110 if maior <= 8
             else 94 if maior <= 16 else 78)
    passo_x = 104 if maior <= 4 else 94 if maior <= 8 else 82
    fila = queda + 40
    # mesmo acréscimo do desenho: 11 px pelo rodapé do segmento no cartão
    rot = 11 if any(r.get("fseg") for r in cad["ramais"]) else 0
    alt = ((150 + 146 + queda + (fileiras - 1) * fila + 76 + rot) if ff
           else max(460, 190 + 28 + queda + (fileiras - 1) * fila + 76 + rot))
    if ff:
        larg = max(400 + (len(cad["segmentos"]) - 1) * 430 + 195,
                   400 + (min(3, maior) - 1) * (passo_x + 6) + 64, 1100)
    else:
        larg = max(400 + 120 + 38 + 96 + (cols - 1) * passo_x + 76, 1100)
    # 1.320 px é a largura medida do painel numa janela de 1.600 com a lateral
    # aberta. Em tela menor a cena encolhe e sobra folga; em tela muito maior
    # ela passa do quadro e a moldura rola, que é o defeito mais barato dos dois.
    return int(min(860, max(360, largura_px * alt / larg + 46)))


def render_certificacao(tags: pd.DataFrame, lanc: pd.DataFrame, depara: pd.DataFrame,
                        resumo: pd.DataFrame, esperados: pd.DataFrame, sigem: pd.DataFrame,
                        cache_key: str):
    """Se esta TAG pode ser certificada, e onde a cadeia dela para.

    A aba é visual: o resumo em cima, o trajeto no meio e a lista de
    predecessoras ao lado. Certificar é assinar que tudo antes está pronto --
    então o que a tela precisa mostrar é o "tudo antes", inteiro, de uma vez.
    """
    render_header("Certificação")

    if lanc.empty:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "Esta planilha ainda não tem o controle de lançamento de circuitos. "
                    "Rode o pipeline com a <code>07_BASE_CABOS_COMPLETO.xlsx</code> na "
                    "pasta de bases — ela traz status, percentual e metragem de cada "
                    "circuito, que é o que responde a certificação.</div></div>")
        return

    mont = cert_montagem(tags, depara, cache_key)
    pan = cert_panorama(lanc, mont, cache_key)
    alvos = cert_alvos(lanc, cache_key)
    indice = cert_indice(lanc, cache_key)
    if not alvos:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "Nenhuma caixa de junção nesta base de circuitos.</div></div>")
        return

    # Duas perguntas diferentes, dois controles: o status recorta o universo,
    # a busca escolhe dentro dele. Num campo só, escolher "TAG apta" e depois
    # digitar uma travada devolveria lista vazia sem dizer por quê.
    por_tag = pan["por_tag"]

    # indice e por_tag nascem com a chave da PONTA da planilha de cabos, que
    # escreve a mesma TAG de forma diferente da 01_BASE_TAGS -- ZSH/L-120001
    # lá vira ZSH-120001 e ZSL-120001 aqui; HS-120610A1 e HS-120610A2 lá viram
    # a mesma HS-120610A1/A2 aqui. Sem isso, mais de 400 TAGs com cabo lançado
    # nunca cruzavam com a base e sumiam de "TAGs mapeadas" e da busca. A
    # chave crua continua no dicionário (a ficha da cadeia ainda busca por
    # ela) -- isto só ACRESCENTA a TAG real como sinônimo, apontando pro
    # mesmo valor.
    traduz = _cert_traduz(depara)
    pior_tom = {"ok": 0, "desc": 1, "warn": 2, "crit": 3}

    def nome_de_base(ponta: str) -> str:
        """O nome da TAG como a 01_BASE_TAGS escreve.

        A planilha de cabos as vezes usa menos zeros -- TJT-12-042 la,
        TJT-12-0042 aqui -- e o de-para ja resolve isso (e o "zero" dele). Sem
        aplicar a traducao no desenho, o cartao saia com o nome curto: sem
        ficha para abrir, sem segmento nem malha (que sao indexados pelo nome
        da base) e sem se reconhecer na linha da tabela ao lado, que usa o
        nome longo. Era a mesma TAG escrita de dois jeitos na mesma tela.
        """
        if ponta in atrib:
            return ponta
        certos = [t for t in traduz.get(ponta, ()) if t in atrib]
        return certos[0] if len(certos) == 1 else ponta
    for cru, valor in list(indice.items()):
        for alvo in traduz.get(cru, ()):
            indice.setdefault(alvo, valor)
    for cru, valor in list(por_tag.items()):
        for alvo in traduz.get(cru, ()):
            atual = por_tag.get(alvo)
            if atual is None or pior_tom[valor["tom"]] > pior_tom[atual["tom"]]:
                por_tag[alvo] = valor

    # Recorte pela cadeia física, direto da 01_BASE_TAGS. Vem antes do status
    # porque muda o universo: a contagem de cada recorte tem que ser a do
    # painel escolhido, não a da obra inteira.
    atrib = cert_atributos(tags, cache_key)
    campos = [(c, rot) for c, rot in CERT_FILTROS
              if any(c in v for v in atrib.values())]
    # indice e por_tag vem do lancamento de cabos (lanc), que nao passa pelo
    # filtro geral da lateral -- so a 01_BASE_TAGS passa. Sem cruzar aqui, a
    # Certificacao continuava mostrando a obra inteira mesmo com um filtro
    # ativo, a unica aba que escapava do recorte.
    #
    # universo NAO exige mais "t in indice": indice so registra um alvo
    # quando a cadeia acha caixa ou painel, e uma cadeia genuinamente aberta
    # (loop de instrumento que nunca fecha, por exemplo) nao acha nenhum --
    # exigir os dois escondia a TAG de "TAGs mapeadas" mesmo com por_tag já
    # sabendo que ela existe e que o cabo esta incompleto. A busca (linha
    # abaixo, "if nome in indice") ja cai de volta pro nome puro quando falta
    # indice, entao soltar essa exigencia aqui nao quebra ela.
    tags_no_filtro_geral = set(tags["TAG"])
    universo = [t for t in por_tag if t in tags_no_filtro_geral]

    # Tudo em TAG. Metro e circuito são a unidade da planilha de cabos, não a
    # do controle: o que se certifica é o instrumento. Os cartões usam so o
    # recorte do filtro geral (universo), nao o do painel/segmento local mais
    # abaixo -- esses dois continuam cada um com seu proprio escopo.
    circ_da_ponta, metros_circ = cert_circuitos_por_ponta(lanc, depara, cache_key)
    total = len(universo)
    metros_filtro, lancado_filtro = cert_metros_uniao(
        circ_da_ponta, metros_circ, universo)
    pct_lanc = lancado_filtro / metros_filtro * 100 if metros_filtro else 0.0
    montadas_travadas = sum(1 for t in universo
                            if por_tag[t]["montada"] and not por_tag[t]["cabo"])
    render_html(f"""
      <div class="pl-kpis tres">
        <div class="pl-kpi"><div class="r">TAGs mapeadas</div>
          <div class="v">{br_num(total)}</div>
          <div class="s">com cadeia de cabo na base</div></div>
        <div class="pl-kpi"><div class="r">Avanço do cabo</div>
          <div class="v andando">{br_pct(pct_lanc)}</div>
          <div class="s">{br_num(int(lancado_filtro))} de {br_num(int(metros_filtro))} m</div>
          <div class="pl-barra"><i class="andando" style="width:{pct_lanc:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="r">Montadas travadas</div>
          <div class="v">{br_num(montadas_travadas)}</div>
          <div class="s">montadas, esperando cabo</div></div>
      </div>""")

    # Exportar pendencia de LANCAMENTO DE CABO por prefixo de TAG (ex.: AST,
    # OST) -- pra montar programacao de campo sem abrir TAG por TAG na tela.
    # E uma pergunta diferente dos filtros abaixo ("o que falta lancar pra
    # estes tipos", nao um recorte da tela) e funciona independente deles,
    # e independente de estar olhando um painel/caixa especifico ou nao.
    with st.expander("Exportar pendências de cabo"):
        prefixos_txt = st.text_input(
            "Prefixos de TAG, separados por vírgula", value="AST,OST",
            key="cert_export_prefixos")
        prefixos = tuple(p.strip().upper() for p in prefixos_txt.split(",") if p.strip())
        alvo_export = sorted(t for t in universo if prefixos and t.upper().startswith(prefixos))
        circuitos_export: dict[str, dict] = {}
        for tag in alvo_export:
            for c in por_tag[tag].get("cadeia", []):
                # so o que ainda falta lancar -- montagem do instrumento e
                # outra pendencia, essa exportacao e so sobre o cabo
                if cert_num(c["PCT"]) >= 99.5:
                    continue
                cid = str(c["CIRCUITO"]).strip()
                linha = circuitos_export.setdefault(cid, {
                    "CIRCUITO": cid, "ORIGEM": str(c["ORIGEM"]).strip(),
                    "DESTINO": str(c["DESTINO"]).strip(),
                    "DISCIPLINA": str(c["DISCIPLINA"]).strip(),
                    "STATUS": str(c["STATUS"]).strip(),
                    "PCT": round(cert_num(c["PCT"]), 1),
                    "METROS": cert_num(c["METROS"]),
                    "METROS_REAL": cert_metro_real(c),
                    "TAGS": set(),
                })
                linha["TAGS"].add(tag)
        linhas_export = sorted(
            ({**v, "TAGS": ", ".join(sorted(v["TAGS"]))} for v in circuitos_export.values()),
            key=lambda v: (v["PCT"], v["CIRCUITO"]))
        if not prefixos:
            st.caption("Digite ao menos um prefixo (ex.: AST, OST).")
        elif not alvo_export:
            st.caption("Nenhum TAG mapeado com esse prefixo.")
        elif not linhas_export:
            st.caption(f"{len(alvo_export)} TAGs com esse prefixo · "
                      "nenhum circuito de cabo pendente, já está tudo lançado.")
        else:
            st.caption(f"{len(alvo_export)} TAGs com esse prefixo · "
                      f"{len(linhas_export)} circuitos de cabo pendentes")
            df_export = pd.DataFrame(linhas_export)
            csv = df_export.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
            st.download_button(
                "Baixar pendências de cabo (CSV)", csv,
                file_name=f"pendencias_cabo_{'_'.join(prefixos)}.csv",
                mime="text/csv", key="cert_export_download")
            st.dataframe(df_export, hide_index=True, width="stretch")

    escolhido = {c: st.session_state.get(f"cert_f_{c}", "Todos") for c, _ in campos}

    def combina(tag, exceto=""):
        v = atrib.get(tag, {})
        return all(escolhido[c] == "Todos" or v.get(c, "") == escolhido[c]
                   for c, _ in campos if c != exceto)

    if campos:
        for (campo, rotulo), col in zip(campos, st.columns(len(campos))):
            # As opções saem do universo já recortado pelos OUTROS filtros:
            # sem isso dá para escolher um painel e um segmento que não se
            # cruzam, e a tela volta vazia sem dizer por quê.
            opcoes = ["Todos"] + sorted({atrib[t][campo] for t in universo
                                         if atrib.get(t, {}).get(campo)
                                         and combina(t, campo)})
            chave = f"cert_f_{campo}"
            # o valor guardado pode ter saído das opções depois de mexer em
            # outro filtro -- sem isto o Streamlit levanta erro na hora
            if st.session_state.get(chave) not in opcoes:
                st.session_state[chave] = "Todos"
            with col:
                escolhido[campo] = st.selectbox(rotulo, opcoes, key=chave)

    universo_f = [t for t in universo if combina(t)]
    ativos = [f"{rot.lower()} {escolhido[c]}" for c, rot in campos
              if escolhido[c] != "Todos"]

    def conta(prova):
        return sum(1 for t in universo_f if prova(por_tag[t]))

    RECORTES = {
        "Todos": (lambda v: True, len(universo_f)),
        "Circuitos aptos": (lambda v: v["cabo"], conta(lambda v: v["cabo"])),
        "TAG apta": (lambda v: v["apta"], conta(lambda v: v["apta"])),
        "Montadas travadas": (lambda v: v["montada"] and not v["cabo"],
                              conta(lambda v: v["montada"] and not v["cabo"])),
        "Inaptas": (lambda v: not v["cabo"], conta(lambda v: not v["cabo"])),
        # a base não deixa concluir: a TAG chega em mais de um painel com
        # resultados diferentes. Fica num recorte próprio porque é fila de
        # correção da base, não de campo -- ninguém vai lançar cabo para
        # resolver, alguém vai acertar o cadastro.
        "Corrigir na base": (lambda v: bool(v.get("ambiguo")),
                             conta(lambda v: bool(v.get("ambiguo")))),
    }
    alvo_recorte = st.segmented_control(
        "Status de certificação", list(RECORTES),
        format_func=lambda x: f"{x} · {br_num(RECORTES[x][1])}",
        default="Todos", key="cert_recorte") or "Todos"
    cabe = RECORTES[alvo_recorte][0]

    tags_no_filtro = sorted(t for t in universo_f if cabe(por_tag[t]))

    # Avanço por segmento. Fica depois dos filtros de propósito: escolhido um
    # painel, o gráfico compara os segmentos DELE -- é a pergunta que a aba
    # não respondia, e que obrigava a abrir caixa por caixa.
    if any(c == "SEGMENTO" for c, _ in campos):
        # O segmento é uma coisa física inteira: o tronco que sai do painel, os
        # troncos entre as caixas e os ramais até os instrumentos. A caixa não
        # tem coluna de segmento, mas a caixa É do segmento -- quem diz isso é
        # a coluna CFF da 01_BASE_TAGS, e é dela que sai este mapa.
        #
        # Sem os troncos o H1-0057 aparecia 100% com o cabo da CFF-12-0057C
        # para a CFF-12-0057B por lançar. São 290 troncos e 48.846 m que
        # ficavam de fora, 152 deles não concluídos.
        # circ_da_ponta/metros_circ ja vieram la de cima, dos cartoes do topo.
        caixa_seg = {}
        for v in atrib.values():
            cx, sg = v.get("CFF", ""), v.get("SEGMENTO", "")
            # exige que a caixa realmente origine circuito: descarta o "N.A."
            # que a base usa como preenchimento em 6 segmentos
            if cx and sg and cx in circ_da_ponta:
                caixa_seg[cx] = sg

        # Só entra o que está mapeado na lógica: TAG com cadeia de cabo até uma
        # caixa. TAG que a base dá como pronta e que não aparece aqui é defeito
        # de base a corrigir, e somá-la escondia exatamente o defeito -- por
        # isso o universo, e não a 01_BASE_TAGS inteira.
        #
        # O filtro escolhe QUAIS segmentos aparecem; o número de cada um sai do
        # segmento inteiro. Recortar o segmento pelo filtro daria um tronco
        # inteiro dividido por meia dúzia de ramais, que não é o avanço de nada.
        visiveis = {atrib[t].get("SEGMENTO", "") for t in universo_f
                    if atrib.get(t, {}).get("SEGMENTO")}
        segs: dict[str, dict] = {}
        for tag in universo:
            nome_seg = atrib.get(tag, {}).get("SEGMENTO", "")
            if nome_seg not in visiveis:
                continue
            s = segs.setdefault(nome_seg,
                                {"m": 0.0, "real": 0.0, "tags": 0, "caixas": set(),
                                 "circ": set(), "montadas": 0})
            s["tags"] += 1
            if por_tag[tag]["montada"]:
                s["montadas"] += 1
            # junta os IDS, soma depois: um segmento com ZSH-x e ZSL-x tem as
            # duas penduradas no mesmo cabo, e somar por TAG contava o trecho
            # duas vezes dentro do proprio segmento
            s["circ"] |= circ_da_ponta.get(tag, set())
        for cx, sg in caixa_seg.items():
            if sg in segs:
                segs[sg]["caixas"].add(cx)
        for s in segs.values():
            for cx in s["caixas"]:
                s["circ"] |= circ_da_ponta.get(cx, set())
            s["m"] = sum(metros_circ[c][0] for c in s["circ"])
            s["real"] = sum(metros_circ[c][1] for c in s["circ"])
        if segs:
            # Do mais adiantado para o menos, sem separar concluído de não
            # iniciado: é o ranking que Daniel pediu, cabo e geral na mesma
            # ordem. Empate no percentual desempata pelo maior previsto -- é
            # o tamanho que diz quanto aquele número pesa na obra.
            def por_avanco(kv):
                s = kv[1]
                p = s["real"] / s["m"] * 100 if s["m"] else 0.0
                return (-p, -s["m"])

            ordem = sorted(segs.items(), key=por_avanco)
            andando = sum(1 for _, s in ordem
                          if s["m"] and 0 < s["real"] / s["m"] * 100 < 99.5)
            prontos = sum(1 for _, s in ordem
                          if s["m"] and s["real"] / s["m"] * 100 >= 99.5)
            # O metro já vem capado no previsto pelo cert_metro_real, que é a
            # mesma fonte do cartão "Avanço do cabo" -- então este percentual
            # não passa de 100 e não é uma segunda resposta para a pergunta
            # que o cartão já responde.
            linhas_seg = []
            for nome_seg, s in ordem:
                p = s["real"] / s["m"] * 100 if s["m"] else 0.0
                linhas_seg.append(
                    f'<div class="cs-lin"><span class="cs-seg">{esc(nome_seg)}</span>'
                    f'<div class="cs-trilho"><i class="{classe_avanco(p)}" '
                    f'style="width:{min(max(p, 0), 100):.1f}%"></i></div>'
                    f'<span class="cs-num"><b class="{classe_avanco(p)}">{br_pct(p)}</b>'
                    f' · {br_num(int(s["real"]))}/{br_num(int(s["m"]))} m'
                    f' · {br_num(len(s["caixas"]))} cx'
                    f' · {br_num(s["tags"])} TAG{"s" if s["tags"] > 1 else ""}</span></div>')
            painel_cabo = (
                '<div class="gplan-panel pl-pn"><div class="gplan-panel-title">'
                'Avanço do cabo por segmento'
                f'<span class="pl-res">{br_num(len(ordem))} '
                f'segmento{"s" if len(ordem) > 1 else ""} · '
                f'<b class="andando">{br_num(andando)}</b> em andamento · '
                f'<b class="feito">{br_num(prontos)}</b> com cabo pronto'
                f'</span></div><div class="cs-lista">{"".join(linhas_seg)}</div></div>')

            # Avanço geral: o cabo do segmento pesa 1 unidade (a própria
            # porcentagem dele) e cada TAG pesa outra (100 se montada, 0 se
            # não) -- a média simples das duas é quanto do segmento inteiro,
            # cabo e instrumento junto, já fechou.
            def peso_geral(s):
                p_cabo = s["real"] / s["m"] * 100 if s["m"] else 0.0
                unidades = 1 + s["tags"]
                return (p_cabo + 100 * s["montadas"]) / unidades

            def por_avanco_geral(kv):
                s = kv[1]
                p = peso_geral(s)
                unidades = 1 + s["tags"]
                return (-p, -unidades)

            ordem_geral = sorted(segs.items(), key=por_avanco_geral)
            andando_geral = sum(1 for _, s in ordem_geral
                                if 0 < peso_geral(s) < 99.5)
            prontos_geral = sum(1 for _, s in ordem_geral
                                if peso_geral(s) >= 99.5)
            linhas_geral = []
            for nome_seg, s in ordem_geral:
                p = peso_geral(s)
                linhas_geral.append(
                    f'<div class="cs-lin"><span class="cs-seg">{esc(nome_seg)}</span>'
                    f'<div class="cs-trilho"><i class="{classe_avanco(p)}" '
                    f'style="width:{min(max(p, 0), 100):.1f}%"></i></div>'
                    f'<span class="cs-num"><b class="{classe_avanco(p)}">{br_pct(p)}</b>'
                    f' · {br_num(s["montadas"])}/{br_num(s["tags"])}'
                    f' TAG{"s" if s["tags"] != 1 else ""} montada'
                    f'{"s" if s["montadas"] != 1 else ""}</span></div>')
            painel_geral = (
                '<div class="gplan-panel pl-pn"><div class="gplan-panel-title">'
                'Avanço geral por segmento'
                f'<span class="pl-res">{br_num(len(ordem_geral))} '
                f'segmento{"s" if len(ordem_geral) > 1 else ""} · '
                f'<b class="andando">{br_num(andando_geral)}</b> em andamento · '
                f'<b class="feito">{br_num(prontos_geral)}</b> completo'
                f'{"s" if prontos_geral > 1 else ""}'
                f'</span></div><div class="cs-lista">{"".join(linhas_geral)}</div></div>')

            render_html(f'<div class="cs-duas">{painel_cabo}{painel_geral}</div>')
    # Um recorte vazio não pode apagar a tela: "TAG apta" hoje tem zero, e a
    # aba inteira sumia junto. A busca cai de volta para todas, e quem explica
    # o vazio é a tabela, no lugar dela.
    base_busca = tags_no_filtro or sorted(indice)
    # O painel entra na busca como alvo: e a unica forma de ver o segmento
    # inteiro sem abrir uma caixa por vez.
    paineis = cert_paineis(lanc, cache_key)
    opcoes = ([f"{p}  ·  painel · {len(b)} bloco{'s' if len(b) != 1 else ''}"
               for p, b in sorted(paineis.items())]
              + [f"{t}  ·  TAG" for t in base_busca]
              + ([f"{a}  ·  {'fieldbus' if a.startswith('CFF') else 'caixa'}"
                  for a in alvos] if not tags_no_filtro or alvo_recorte == "Todos" else []))
    # O pouso e uma TAG, nao o primeiro painel: PN-12-201 tem 72 caixas e leva
    # 14 s para desenhar, e abrir a aba nao pode custar isso. O painel continua
    # na lista, a um "PN" digitado de distancia.
    padrao = st.session_state.get("cert_escolha")
    if padrao in opcoes:
        idx = opcoes.index(padrao)
    else:
        idx = next((i for i, o in enumerate(opcoes) if o.endswith("·  TAG")), 0)
    escolha = st.selectbox("Pesquisar TAG, caixa ou tronco de fieldbus", opcoes,
                           index=idx, key="cert_escolha",
                           help="Digite para filtrar. A TAG escolhida abre destacada.")
    nome = escolha.split("  ·  ")[0]
    if nome in indice:
        alvo, tag_sel = indice[nome], nome
    else:
        alvo, tag_sel = nome, ""

    # Com filtro da base ativo quem manda no desenho é o filtro, não a busca.
    # Trocar o filtro tira a escolha anterior das opções, e a busca repousa
    # sozinha na primeira TAG do recorte -- uma TAG que ninguém pediu, que
    # arrastava o desenho para o laço dela: o segmento MB-RTU-03 tem 28
    # instrumentos em três laços e a tela mostrava os 8 de um só. Aqui o alvo
    # volta para onde as TAGs do recorte realmente moram.
    if ativos:
        onde_mora = collections.Counter(indice[t] for t in universo_f if t in indice)
        if onde_mora and not onde_mora.get(alvo):
            alvo = onde_mora.most_common(1)[0][0]

    com_ficha = set(resumo["TAG"].astype(str))
    def render_fichas(ids):
        """As fichas ficam no fim da página, fechadas: o clique no desenho abre
        só a que ele pedir, e nenhuma ida ao servidor acontece.

        As fichas de nível vêm junto para que os degraus da trilha -- Fase,
        SOP, SSOP, Malha -- abram aqui mesmo. Sem elas na página, o degrau
        virava um link para a aba Progresso, que é sair da tela para ver algo
        que cabia nela.
        """
        ids = list(dict.fromkeys(ids))
        if not ids:
            return
        espera = espera_por_documento(_revisoes_por_doc(cache_key, sigem))
        base_niveis = progresso_base(resumo, tags)
        base_niveis = base_niveis[base_niveis["TAG"].isin(ids)]
        # As tres familias juntas: sem a do relatorio, o "ver detalhe" dentro da
        # ficha da TAG apontava para uma ancora inexistente e o :target fechava
        # tudo. O degrau de nivel abre aqui mesmo -- fechar aquela ficha volta
        # para o pai dela na trilha (Malha -> SSOP -> SOP -> Fase -> fechado),
        # a mesma logica de qualquer outra ficha do projeto.
        render_html(fichas_niveis_html(cache_key, f"cert:{assinatura_tags(ids)}", "", "",
                                       0, base_niveis, esperados)
                    + fichas_modais_html(ids, resumo, esperados, tags, espera,
                                         niveis_na_pagina=True)
                    + fichas_relatorios_pagina(cache_key, "cert", "", "", 0,
                                               ids, esperados, sigem))

    def render_tabela(do_alvo=None, rotulo=""):
        """A tabela acompanha o desenho: as TAGs do alvo aberto, recortadas
        pelo status. Mostrar o universo enquanto o desenho mostra um segmento
        era pedir para comparar duas coisas diferentes lado a lado.

        Com filtro da base ativo é o contrário: aí o pedido é ver o recorte
        inteiro. O desenho segue mostrando a caixa aberta, e a tabela passa a
        listar todas as TAGs do filtro -- filtrar por um segmento e continuar
        vendo só uma caixa esconderia justamente o que foi pedido.
        """
        if ativos:
            do_alvo, rotulo = None, " · ".join(ativos)
        no_alvo = set(do_alvo) if do_alvo else None
        linhas = []
        for tag in (tags_no_filtro if no_alvo is None
                    else [t for t in tags_no_filtro if t in no_alvo]):
            v = por_tag[tag]
            marca = CERT_CLASSE[v["tom"]]
            mt = "ok" if v["montada"] else "crit"
            atual = " sel" if tag == tag_sel else ""
            # Duas perguntas diferentes, duas colunas. "Cabo do instrumento"
            # é o circuito da própria TAG na planilha de cabos; "Certificação"
            # é a cadeia inteira até a alimentação do painel. Numa coluna só,
            # instrumento com o cabo Concluído aparecia como "Bloqueado" --
            # verdade sobre a cadeia, mentira sobre o cabo dele.
            ct = "ok" if v.get("cabo_tag") else (
                "warn" if v.get("pct_tag", 0) > 0 else "crit")
            rot_cabo = v.get("status_tag", "sem circuito")
            if v.get("pct_tag", 0) > 0 and not v.get("cabo_tag"):
                rot_cabo += f' · {br_pct(v["pct_tag"])}'
            linhas.append(
                f'<tr class="ct-lin{atual}"><td class="gtbl-mono">{tag}</td>'
                f'<td class="gtbl-mono gtbl-muted">{v["caixa"]}</td>'
                f'<td><span class="gtbl-badge {ct}">{rot_cabo}</span></td>'
                f'<td><span class="gtbl-badge {"ok" if v["cabo"] else marca}">'
                f'{"apto" if v["cabo"] else CERT_ROTULO[v["tom"]]}</span></td>'
                f'<td><span class="gtbl-badge {mt}">'
                f'{"montada" if v["montada"] else "não montada"}</span></td>'
                f'<td class="gtbl-muted" style="font-size:11px">{v["onde"]}</td></tr>')
        onde = f" · {rotulo}" if rotulo else ""
        corpo = (f'<div class="gtbl-empty">Nenhuma TAG deste {rotulo or "recorte"} '
                 f'em {alvo_recorte.lower()}.</div>' if not linhas else
                 f'<div class="ct-rolo"><table class="gtbl"><thead><tr><th>TAG</th>'
                 f'<th>Caixa</th><th>Cabo do instrumento</th><th>Certificação</th>'
                 f'<th>Montagem</th><th>Trava em</th></tr>'
                 f'</thead><tbody>{"".join(linhas)}</tbody></table></div>')
        render_html(
            f'<div class="gplan-panel ct-painel">'
            f'<div class="gplan-panel-title">TAGs do desenho'
            f'<span class="gtbl-muted" style="font-weight:500">{br_num(len(linhas))} '
            f'{"TAG" if len(linhas) == 1 else "TAGs"}{onde} · {alvo_recorte.lower()}'
            f'</span></div>{corpo}</div>')

    def desenha_painel(alvo, escopo):
        """Desenha um painel e devolve os blocos que desenhou.

        É função porque um recorte pode morar em mais de um painel: o
        segmento MB-RTU-02 tem 26 instrumentos repartidos entre o PN-12-236 e
        o PN-12-237, e um desenho só mostrava 18 deles.
        """
        blocos_alvo = paineis[alvo]
        # Pesquisar uma TAG que mora num laço sem caixa (deteccao de
        # fumaca/gas, YST) resolvia pro PAINEL inteiro -- indice() nao tem
        # como apontar pra caixa nenhuma nesses casos -- e a tela abria TODOS
        # os lacos do painel, nao so o da TAG procurada. Aqui, com uma TAG
        # especifica escolhida na busca, o recorte cai pro laco dela; buscar
        # o PAINEL (sem TAG) continua abrindo todos, do jeito que ja era.
        bloco_tag = None
        if escopo:
            # Com filtro ativo o desenho é o recorte inteiro: ficam os blocos
            # que têm alguma TAG do filtro. É a mesma regra que a tabela já
            # seguia -- filtrar um segmento e continuar vendo um laço só
            # escondia justamente o que foi pedido. Recorte que cai num laço
            # só continua abrindo esse laço, com a metragem trecho a trecho:
            # filtrar não pode custar detalhe.
            no_escopo = [b for b in blocos_alvo
                         if any(t["org"] in escopo for t in b["tags"])]
            if no_escopo:
                blocos_alvo = no_escopo
            if len(blocos_alvo) == 1:
                bloco_tag = blocos_alvo[0]
        elif tag_sel:
            bloco_tag = next((b for b in blocos_alvo
                              if any(t["org"] == tag_sel for t in b["tags"])), None)
            if bloco_tag is not None:
                blocos_alvo = [bloco_tag]
        # O laço aberto é o da busca só quando a TAG procurada mora nele --
        # com filtro ativo ele pode ter vindo do recorte, sem TAG nenhuma.
        laco_da_busca = bloco_tag is not None and any(
            t["org"] == tag_sel for t in bloco_tag["tags"])
        cad = cert_cadeia_painel(alvo, blocos_alvo, mont,
                                 cert_alimentacao_painel(lanc, alvo, cache_key))
        n_caixas = sum(1 for b in cad["blocos"] if not b["direto"])
        n_diretos = sum(1 for b in cad["blocos"] if b["direto"])
        partes = [f"{n_caixas} caixas"] if n_caixas else []
        if n_diretos:
            partes.append(f"{n_diretos} loop{'s' if n_diretos > 1 else ''} sem caixa")
        if bloco_tag is not None:
            titulo_painel = "Laço da TAG" if laco_da_busca else "Laço do recorte"
            sub = ((f"{tag_sel} → " if laco_da_busca else "")
                   + f"laço {bloco_tag['nome']} → {bloco_tag['inst']} instrumentos")
        else:
            titulo_painel, sub = "Segmento do painel", (
                f"{alvo} → {' + '.join(partes) or '0 caixas'} → "
                f"{sum(b['inst'] for b in cad['blocos'])} instrumentos")
        render_html(f'<div class="gplan-panel-title" style="margin:4px 0 10px">'
                    f'{titulo_painel} <span class="gtbl-muted" '
                    f'style="font-weight:500">{sub}</span></div>')
        for bl in cad["blocos"]:
            for t in bl["tags"]:
                t["base"] = nome_de_base(t["org"])
                t["ancora"] = (ficha_anchor(t["base"])
                               if t["base"] in com_ficha else "")
        # A moldura nasce do tamanho da vista fechada -- uma faixa de 96 px por
        # caixa. Abrir cresce por dentro, e a moldura rola: dimensionar pelo
        # pior caso deixaria meia tela vazia enquanto tudo estivesse fechado.
        # O laço de uma TAG já nasce aberto (ver "abertos" no JS), então esse
        # bloco sozinho entra com a altura de aberto, não a de fechado --
        # senão a moldura vinha baixa e cortava o próprio laço que é o motivo
        # da busca.
        # topo_cena espelha o "topo" do JS: a faixa da alimentação empurra
        # tudo para baixo quando o painel tem mais de um cabo alimentando.
        topo_cena = 56 + max(0, len(cad["eletrica"]) - 1) * 14
        if bloco_tag is not None:
            n_tags = len(bloco_tag["tags"])
            # mesma regra do JS (cert_cena_html): um loop YST com identidade
            # fixa sempre dobra em duas fileiras a partir de 4 tags -- sem
            # espelhar isso aqui, a moldura vinha baixa pela metade e cortava
            # a segunda fileira do proprio laço que é o motivo da busca.
            fil = 2 if (bloco_tag.get("ordem_fixa") and n_tags > 3) else max(1, -(-n_tags // 12))
            alt_cena = topo_cena + max(96, 22 + fil * 76 + 34 + (14 if n_tags > 1 else 0)) + 30
        else:
            alt_cena = topo_cena + 96 * len(cad["blocos"]) + 30
        # A largura tem de ser a REAL da cena, não uma estimativa fixa: o SVG
        # é escalado para caber na largura da moldura, então a altura útil sai
        # de largura_da_moldura × altura_da_cena ÷ largura_da_cena. Com a
        # estimativa antiga (12 colunas de 84 px) o laço, que é bem mais
        # largo, era encolhido e sobrava meia moldura vazia embaixo.
        # A largura da CENA muda conforme o que está desenhado, e é ela que
        # decide a altura útil: o SVG escala pela largura, então a altura na
        # tela é largura_da_moldura × altura_da_cena ÷ largura_da_cena.
        #
        # Com os blocos FECHADOS (a vista do painel inteiro) não há cartão de
        # instrumento nenhum, e a cena tem os 1.100 mínimos -- usar aqui a
        # largura de 12 colunas de instrumento, que era o que estava, dava uma
        # cena "mais larga" do que a real, e a moldura saía baixa e cortava os
        # blocos. Com o laço aberto a largura é a real dos cartões.
        fixa_cena = bool(bloco_tag and bloco_tag.get("ordem_fixa")
                         and len(bloco_tag["tags"]) > 3)
        if bloco_tag is not None:
            n_t = len(bloco_tag["tags"])
            cols_cena = max(n_t // 2, n_t - n_t // 2) if fixa_cena else min(12, n_t)
            pas_cena = 168 if fixa_cena else 84
            larg = max(1100, 150 + 240 + 200 + (cols_cena - 1) * pas_cena + 36 + 46 + 40)
        else:
            larg = 1100
        # folga de sobra: apertado, o desenho não deixa ler nem arrastar. É a
        # área de trabalho da aba, não um selo.
        altura_p = int(min(900, max(420, CERT_MOLDURA_PX * alt_cena / larg + 90)))
        st.components.v1.html(cert_cena_html(cad, tag_sel or "", tema_ativo(), altura_p),
                              height=altura_p, scrolling=False)
        # Só a linha que diz o recorte: a faixa que explicava a leitura do
        # desenho repetia o que a própria cena já mostra e roubava altura de
        # quem interessa, que é o desenho.
        if laco_da_busca:
            render_html('<div class="ct-leg"><span style="color:var(--text-3)">'
                        f'mostrando só o laço desta TAG · pesquise o painel {alvo} '
                        'para ver todos os laços dele</span></div>')
        # No laco (sem caixa), o desenho ja mostra o fio real entre os
        # instrumentos -- mas so no hover. A tabela repete isso em texto: a
        # metragem e o avanco de cada trecho, vindos da base completa de
        # cabos, e nao so o resumo (cabo pronto/nao) que o cartao mostra.
        if bloco_tag is not None:
            # m/m_real vem do proprio cert_paineis, nao de uma nova consulta:
            # cert_circuitos_por_tag so enxerga circuito de ORIGEM, e um
            # instrumento que so recebe cabo (fim de trecho, sem nada saindo
            # dele) fica de fora dela -- exatamente os casos que o
            # cert_paineis ja resolveu (inclusive ignorando a correcao de
            # topologia, que nao e o status de verdade da TAG).
            linhas_loop = []
            for t in bloco_tag["tags"]:
                m, m_real = t.get("m", 0.0), t.get("m_real", 0.0)
                # Status e metragem lado a lado, sem um decidir pelo outro:
                # Concluído com menos metro que o previsto é normal, e mais
                # metro que o previsto pode seguir Em Andamento. Antes o
                # status saía do percentual e esses dois casos sumiam.
                status_t = t.get("status") or "sem circuito"
                marca = ("ok" if status_t == "Concluído" else
                         "warn" if status_t == "Em Andamento" else "crit")
                # a diferença entre executado e previsto é fato para olhar,
                # não erro para corrigir -- fica marcada, sem mudar o status
                dif = ""
                if m and m_real > m:
                    dif = (f' <span class="gtbl-muted">· {br_num(int(m_real - m))} m '
                           "acima do previsto</span>")
                elif m and status_t == "Concluído" and m_real < m:
                    dif = (f' <span class="gtbl-muted">· {br_num(int(m - m_real))} m '
                           "abaixo do previsto</span>")
                atual = " sel" if t["org"] == tag_sel else ""
                linhas_loop.append(
                    f'<tr class="ct-lin{atual}"><td class="gtbl-mono">{t["org"]}</td>'
                    f'<td><span class="gtbl-badge {marca}">{status_t}</span></td>'
                    f'<td class="gtbl-mono">{br_pct(t.get("pct", 0.0))}</td>'
                    f'<td class="gtbl-mono">{br_num(int(m_real))} de {br_num(int(m))} m'
                    f'{dif}</td></tr>')
            render_html(
                '<div class="gplan-panel ct-painel">'
                f'<div class="gplan-panel-title">Cabo do laço {bloco_tag["nome"]}'
                '<span class="gtbl-muted" style="font-weight:500">circuito próprio de '
                'cada instrumento até o próximo do laço, da base completa de cabos'
                '</span></div><div class="ct-rolo"><table class="gtbl"><thead><tr>'
                '<th>TAG</th><th>Cabo deste trecho</th><th>%</th><th>Metragem</th>'
                f'</tr></thead><tbody>{"".join(linhas_loop)}</tbody></table></div></div>')
        return cad["blocos"]

    if alvo in paineis:
        # Um desenho por painel do recorte, na ordem de quantas TAGs cada um
        # segura. Sem filtro continua sendo um só, o painel pesquisado.
        escopo = set(universo_f) if ativos else set()
        if ativos:
            quantas = collections.Counter(
                indice[t] for t in universo_f
                if indice.get(t) in paineis)
            ordem_paineis = [p for p, _ in quantas.most_common()] or [alvo]
        else:
            ordem_paineis = [alvo]
        # Teto de fôlego: um recorte largo não pode virar dez desenhos e
        # travar a aba. O que passa do teto fica dito, e a tabela abaixo
        # continua listando o recorte inteiro de qualquer jeito.
        sobrando, ordem_paineis = ordem_paineis[4:], ordem_paineis[:4]
        blocos_vistos = []
        for alvo_p in ordem_paineis:
            blocos_vistos += desenha_painel(alvo_p, escopo)
        if sobrando:
            render_html('<div class="ct-leg"><span style="color:var(--text-3)">'
                        f'mais {br_num(len(sobrando))} '
                        f'{"painel" if len(sobrando) == 1 else "painéis"} neste '
                        f'recorte ({", ".join(sobrando)}) · a tabela abaixo traz '
                        'todas as TAGs</span></div>')
        # o nome da caixa entra na tabela porque a caixa é TAG (tem montagem e
        # documento próprios); o nome do laço não, porque laço não é TAG
        render_tabela([t["org"] for bl in blocos_vistos for t in bl["tags"]]
                      + [bl["nome"] for bl in blocos_vistos
                         if not bl.get("ordem_fixa")],
                      f"painel {' + '.join(ordem_paineis)}")
        render_fichas([t["org"] for bl in blocos_vistos for t in bl["tags"]
                       if t["org"] in com_ficha])
        return

    def desenha_caixa(alvo):
        """Desenha uma caixa e devolve os ramais dela.

        Mesmo motivo do desenha_painel: um segmento de fieldbus pode estar
        repartido em mais de uma caixa, e desenhar uma só deixava de fora
        instrumentos que o filtro tinha pedido.
        """
        nonlocal tag_sel
        cad = cert_cadeia(alvo, lanc, mont)
        if not cad["ramais"]:
            render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                        f"Nenhum instrumento chega em {alvo} nesta base.</div></div>")
            return []
        if not tag_sel:
            tag_sel = cad["ramais"][0]["org"]

        # A situação vem do panorama, que é quem enxerga todos os circuitos da TAG
        # -- inclusive os que não chegam nesta caixa. Recalcular aqui pela cadeia
        # daria uma segunda resposta para a mesma pergunta, na mesma tela.
        circuitos = cert_circuitos_por_tag(lanc, cache_key)
        for r in cad["ramais"]:
            v = por_tag.get(r["org"], {})
            r["tom"] = v.get("tom", "desc")
            r["rot"] = CERT_ROTULO[r["tom"]]
            r["onde"] = v.get("onde", "—")
            cs = circuitos.get(r["org"], r.get("circuitos") or [r])
            r["circuitos"] = cs
            r["m"] = sum(c["m"] for c in cs)
            r["m_real"] = sum(c["m_real"] for c in cs)
            r["pct"] = round(r["m_real"] / r["m"] * 100, 1) if r["m"] else 0.0
            r["status"] = ("Concluído" if all(c["pct"] >= 99.5 for c in cs)
                           else "Em Andamento" if any(c["pct"] > 0 for c in cs)
                           else cs[0]["status"])
            # a âncora só vale para TAG que existe no controle documental: sem ficha
            # para abrir, o cartão não deve parecer clicável
            r["base"] = nome_de_base(r["org"])
            r["ancora"] = ficha_anchor(r["base"]) if r["base"] in com_ficha else ""
            # segmento e malha no cartão: filtrado um segmento, o desenho traz
            # várias malhas juntas e não havia como saber qual cabo é de qual.
            # "fseg" e não "seg": seg já é a caixa de onde o instrumento pendura.
            do_tag = atrib.get(r["base"], {})
            r["fseg"] = do_tag.get("SEGMENTO", "")
            r["malha"] = do_tag.get("MALHA", "")

        titulo = (f"{cad['painel'] or 'painel indefinido'} → {cad['caixa']}"
                  + (f" ({len(cad['segmentos'])} caixas em série)" if cad["segmentos"] else "")
                  + f" → {len(cad['ramais'])} instrumentos")
        render_html(f'<div class="gplan-panel-title" style="margin:4px 0 10px">'
                    f'Trajeto físico <span class="gtbl-muted" '
                    f'style="font-weight:500">{titulo}</span></div>')

        altura = cert_altura(cad)
        st.components.v1.html(cert_cena_html(cad, tag_sel, tema_ativo(), altura),
                              height=altura, scrolling=False)
        return cad["ramais"]

    # Um desenho por caixa do recorte. Sem filtro continua sendo uma só, a
    # caixa pesquisada.
    if ativos:
        quantas_cx = collections.Counter(
            indice[t] for t in universo_f
            if t in indice and indice[t] not in paineis)
        ordem_caixas = [c for c, _ in quantas_cx.most_common()] or [alvo]
    else:
        ordem_caixas = [alvo]
    sobrando_cx, ordem_caixas = ordem_caixas[4:], ordem_caixas[:4]
    ramais_vistos = []
    for alvo_c in ordem_caixas:
        ramais_vistos += desenha_caixa(alvo_c)
    if not ramais_vistos:
        render_tabela()
        return
    if sobrando_cx:
        render_html('<div class="ct-leg"><span style="color:var(--text-3)">'
                    f'mais {br_num(len(sobrando_cx))} '
                    f'caixa{"" if len(sobrando_cx) == 1 else "s"} neste recorte '
                    f'({", ".join(sobrando_cx)}) · a tabela abaixo traz todas as '
                    'TAGs</span></div>')

    render_html("""
      <div class="ct-leg">
        <span style="font-weight:800;color:var(--text-2);letter-spacing:.3px">CABO</span>
        <span><i style="background:var(--accent-teal)"></i>lançado</span>
        <span><i style="background:var(--accent-amber)"></i>em lançamento</span>
        <span><i style="background:repeating-linear-gradient(90deg,
          var(--accent-red) 0 6px,transparent 6px 11px)"></i>não iniciado</span>
        <span><i style="background:var(--accent-purple)"></i>sem circuito cadastrado</span>
        <span style="width:100%;height:1px;background:var(--border-color);margin:2px 0"></span>
        <span style="font-weight:800;color:var(--text-2);letter-spacing:.3px">MONTAGEM</span>
        <span><b style="background:var(--accent-teal)"></b>montado</span>
        <span><b style="background:var(--accent-amber)"></b>em programação</span>
        <span><b style="background:var(--accent-red)"></b>não montado ou não programado</span>
        <span><b style="background:var(--accent-purple)"></b>sem par na base de TAGs</span>
      </div>""")

    render_tabela([r["org"] for r in ramais_vistos],
                  f"caixa {' + '.join(ordem_caixas)}")
    render_fichas([r["org"] for r in ramais_vistos if r["org"] in com_ficha])


# ===================================================================== #
# Última atualização: o que mudou, com a data que o campo registrou      #
# ===================================================================== #
# Não é um diff entre versões da planilha -- é o evento datado que as
# próprias bases já carregam. Parecer do SIGEM, conexão de ponta e teste de
# cabo, medição do Gitec: cada um traz a data em que aconteceu, e é isso que
# permite dizer "há dois dias" sem inventar histórico.

UA_TOM = {"SEM COMENTÁRIOS": "ok", "PARA CONSTRUÇÃO": "ok", "CERTIFICADO": "ok",
          "COM COMENTÁRIOS": "warn", "PENDENTE CERTIFICAÇÃO": "warn",
          "EM ANÁLISE": "azul", "EMITIDO PARA COMENTÁRIOS": "azul", "PARA COMPRA": "azul",
          "RECUSADO": "crit", "CANCELADO": "crit"}

# O tom de quem chegou. Vale para montagem e circuito, que é onde "para onde
# foi" tem leitura de bom ou ruim.
UA_CHEGADA = {"MONTADO": "ok", "CONCLUÍDO": "ok", "CONCLUIDO": "ok", "LOCALIZADO": "ok",
              "EM PROGRAMAÇÃO": "warn", "EM ANDAMENTO": "warn", "PROGRAMADO": "warn",
              "NÃO MONTADO": "crit", "NÃO LOCALIZADO": "crit", "NÃO INICIADO": "crit",
              "NÃO PROGRAMADO": "crit", "REMOVER": "crit"}

UA_MESES = ["jan", "fev", "mar", "abr", "mai", "jun",
            "jul", "ago", "set", "out", "nov", "dez"]

# A frente de cada tipo de movimentação, na ordem em que aparece na tela.
UA_FRENTES = {"documento": "Documento", "montagem": "Montagem", "cabo": "Cabo",
              "campo": "Campo", "cadastro": "Cadastro"}

# Quantos dias a aba mostra. Quem usou o sistema ontem ou na semana passada
# abre e vê o que mudou de lá para cá, sem escolher nada.
UA_JANELAS = {"Hoje": 1, "3 dias": 3, "7 dias": 7, "30 dias": 30}
UA_JANELA_PADRAO = "7 dias"

# Como chamar o que se moveu junto. A TAG é o objeto do controle, mas quem se
# move na frente de cabo é o circuito, e na de documento é o relatório.
UA_PLURAL = {"cabo": ("circuito", "circuitos"), "documento": ("documento", "documentos")}


def ua_quando(quando: pd.Timestamp, agora: pd.Timestamp) -> str:
    """Quanto tempo faz, na medida que a pessoa usa para falar disso."""
    seg = (agora - quando).total_seconds()
    if seg < 0:
        return "agora"
    if seg < 3600:
        return f"há {max(1, int(seg // 60))} min"
    if seg < 86400:
        return f"há {int(seg // 3600)} h"
    dias = int(seg // 86400)
    if dias == 1:
        return "ontem"
    if dias < 30:
        return f"há {dias} dias"
    if dias < 365:
        meses = dias // 30
        return f"há {meses} {'meses' if meses > 1 else 'mês'}"
    anos = dias // 365
    return f"há {anos} {'anos' if anos > 1 else 'ano'}"


def ua_doc_curto(documento: str) -> str:
    """O nome do relatório sem o cabeçalho do contrato.

    Os 25.931 documentos esperados vêm todos como
    C1N_RNEST_U12_<item>_INS_<TIPO>_<alvo>. O que muda de linha para linha é o
    tipo e o alvo; o nome inteiro fica no title, para procurar no SIGEM.
    """
    tipo, _, alvo = str(documento).split("_INS_")[-1].partition("_")
    return f"{tipo} · {alvo}" if alvo else tipo


@st.cache_data(show_spinner=False, max_entries=3)
def ua_movimentos_documento(cache_key: str, _sigem: pd.DataFrame,
                            _esperados: pd.DataFrame) -> list:
    """A mudança de parecer de cada documento deste controle.

    Sai do próprio SIGEM: cada revisão é uma linha, então a revisão anterior diz
    de onde o documento veio. Sem isso, "Recusado" é um estado, não um
    movimento -- e o que interessa é ter saído de Em Análise para Recusado.

    O documento é o sujeito, não a TAG. O RIMII do CHZ-350 é esperado por 272
    TAGs: pendurar a recusa numa delas seria dizer errado, e era o que a versão
    anterior fazia.
    """
    if _sigem.empty or "DATA_PARECER" not in _sigem or _esperados.empty:
        return []
    quantas = (_esperados.groupby(cert_txt(_esperados["DOCUMENTO_ESPERADO"]))["TAG"]
               .nunique().to_dict())

    d = _sigem.copy()
    d["_doc"] = cert_txt(d["DOCUMENTO"])
    d = d[d["_doc"].isin(quantas)]
    if d.empty:
        return []
    # a data do parecer é quando a fiscalização mexeu; sem parecer, vale a
    # emissão, que é quando a revisão passou a existir
    parecer = pd.to_datetime(d["DATA_PARECER"], errors="coerce", dayfirst=True)
    emissao = pd.to_datetime(d["DATA"], errors="coerce", dayfirst=True)
    d["_q"] = parecer.fillna(emissao)
    d = d[d["_q"].notna()].sort_values("_q")

    movs = []
    for doc, grupo in d.groupby("_doc", sort=False):
        anterior = None
        for _, r in grupo.iterrows():
            atual = str(r.get("STATUS", "") or "").strip()
            rev = str(r.get("REVISAO", "") or "").strip()
            # A primeira revisão entra sem "de": a recusa do RIMII do CHZ-350 é
            # da revisão 0, não veio de lugar nenhum, e sumir com ela por isso
            # seria esconder o parecer que mais importa. Depois disso, só entra
            # quando o parecer muda -- revisão nova com o mesmo status não é
            # movimento.
            if atual and (anterior is None or atual != anterior[0]):
                movs.append({
                    "quando": r["_q"], "tipo": "documento", "objeto": doc,
                    "rotulo": ua_doc_curto(doc), "campo": "Parecer",
                    "de": anterior[0] if anterior else "", "para": atual,
                    "tom": UA_TOM.get(atual.upper(), "azul"),
                    "detalhe": (f"revisão {esc(anterior[1])} &rarr; {esc(rev)}"
                                if anterior and rev and rev != anterior[1] else
                                (f"revisão {esc(rev)}" if rev else "")),
                    "tags": int(quantas.get(doc, 0))})
            anterior = (atual, rev)
    return movs


@st.cache_data(show_spinner=False, max_entries=3)
def ua_movimentos_planilha(cache_key: str, _movs: pd.DataFrame) -> list:
    """As movimentações que o pipeline apurou entre uma atualização e outra."""
    if _movs.empty or "DATA" not in _movs.columns:
        return []
    d = _movs.copy()
    d["_q"] = pd.to_datetime(d["DATA"], errors="coerce", dayfirst=True)
    d = d[d["_q"].notna()]
    saida = []
    for _, r in d.iterrows():
        para = cert_txt(pd.Series([r["PARA"]])).iloc[0]
        de = cert_txt(pd.Series([r["DE"]])).iloc[0]
        tipo = str(r["TIPO"]).strip() or "cadastro"
        saida.append({
            "quando": r["_q"], "tipo": tipo, "objeto": str(r["OBJETO"]).strip(),
            "rotulo": str(r["OBJETO"]).strip(), "campo": str(r["CAMPO"]).strip(),
            "de": de, "para": para,
            "tom": UA_CHEGADA.get(para.upper(), "azul" if tipo != "cadastro" else "neutro"),
            "detalhe": "", "tags": int(cert_num(r.get("QTD_TAGS", 1))) or 1})
    return saida


def ua_movimentos(cache_key: str, movs: pd.DataFrame, sigem: pd.DataFrame,
                  esperados: pd.DataFrame) -> list:
    """Tudo que mudou no controle, do mais recente para o mais antigo."""
    tudo = (ua_movimentos_planilha(cache_key, movs)
            + ua_movimentos_documento(cache_key, sigem, esperados))
    tudo.sort(key=lambda m: m["quando"], reverse=True)
    return tudo


def render_atualizacao(movs: pd.DataFrame, sigem: pd.DataFrame, esperados: pd.DataFrame,
                       lanc: pd.DataFrame, tags: pd.DataFrame, resumo: pd.DataFrame,
                       cache_key: str):
    """O que mudou no controle, e de qual valor para qual.

    As outras abas mostram o retrato de agora. Esta mostra o movimento: quem
    abriu ontem ou na semana passada vê aqui o que andou de lá para cá.
    """
    render_header("Última atualização")
    movimentos = ua_movimentos(cache_key, movs, sigem, esperados)
    # ua_movimentos_planilha nao recebe tags/resumo (so o cache_key entra na
    # chave dela), entao uma TAG fora do filtro geral da lateral ainda
    # aparecia aqui. As de "documento" ja vem certas -- quem decide quais
    # documentos entram e o proprio ua_movimentos_documento, a partir do
    # esperados ja recortado. As de "cabo" seguem sem recorte: o objeto
    # delas e o CIRCUITO (ex. "PCC-101-07L"), que nao e o mesmo espaco de
    # nomes da TAG nem da PONTA do de-para, e resolver isso exigiria a
    # cadeia ORIGEM/DESTINO inteira -- o mesmo problema que a Certificacao
    # resolve com a cadeia de caixas, e que nao cabe aqui de graca.
    tags_no_filtro_geral = set(resumo["TAG"])
    movimentos = [m for m in movimentos
                  if m["tipo"] in ("documento", "cabo") or m["objeto"] in tags_no_filtro_geral]
    if not movimentos:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "Nenhuma movimentação registrada nesta planilha. A aba se enche "
                    "a cada atualização de base pelo ATUALIZAR_CONTROLE."
                    "</div></div>")
        return

    agora = pd.Timestamp.now(tz=BR_TZ).tz_localize(None)
    dentro = {rot: [m for m in movimentos if (agora - m["quando"]).days < dias]
              for rot, dias in UA_JANELAS.items()}

    janela = st.segmented_control(
        "Período", list(UA_JANELAS),
        format_func=lambda x: f"{x} · {br_num(len(dentro[x]))}",
        default=UA_JANELA_PADRAO, key="ua_janela") or UA_JANELA_PADRAO
    recorte = dentro[janela]

    conta = collections.Counter(m["tipo"] for m in recorte)
    frentes = ["Tudo"] + [rot for chave, rot in UA_FRENTES.items() if conta.get(chave)]
    escolha = st.segmented_control(
        "Frente", frentes,
        format_func=lambda x: (f"{x} · {br_num(len(recorte))}" if x == "Tudo" else
                               f"{x} · {br_num(conta.get(_ua_chave(x), 0))}"),
        default="Tudo", key="ua_frente") or "Tudo"
    if escolha != "Tudo":
        recorte = [m for m in recorte if m["tipo"] == _ua_chave(escolha)]

    objetos = len({m["objeto"] for m in recorte})
    campo = conta.get("montagem", 0) + conta.get("cabo", 0) + conta.get("campo", 0)
    render_html(f"""
      <div class="pl-kpis">
        <div class="pl-kpi"><div class="r">Movimentações</div>
          <div class="v">{br_num(len(dentro[janela]))}</div>
          <div class="s">nos últimos {UA_JANELAS[janela]} dias</div></div>
        <div class="pl-kpi"><div class="r">Itens que mudaram</div>
          <div class="v">{br_num(objetos)}</div>
          <div class="s">TAGs, circuitos e documentos</div></div>
        <div class="pl-kpi"><div class="r">Campo</div>
          <div class="v andando">{br_num(campo)}</div>
          <div class="s">montagem, cabo e localização</div></div>
        <div class="pl-kpi"><div class="r">Documento</div>
          <div class="v">{br_num(conta.get('documento', 0))}</div>
          <div class="s">parecer da fiscalização</div></div>
      </div>""")

    # Sem corte por movimentação: 800 delas num dia viram 25 linhas depois do
    # agrupamento, e cortar em 600 apagava os dias anteriores da janela. Quem
    # limita é o número de linhas na tela.
    corpo = ua_linha_do_tempo(recorte, agora, set(cert_txt(resumo["TAG"])))
    render_html(f'<div class="gplan-panel ct-painel">'
                f'<div class="gplan-panel-title">O que mudou'
                f'<span class="gtbl-muted" style="font-weight:500">'
                f'de qual situação para qual · clique na TAG para abrir a ficha'
                f'</span></div><div class="ua-rolo"><div class="ua-log">{corpo}</div>'
                f'</div></div>')

    ids = [m["objeto"] for m in recorte if m["tipo"] != "documento"]
    ids = [t for t in dict.fromkeys(ids) if t in set(cert_txt(resumo["TAG"]))][:250]
    if ids:
        base = progresso_base(resumo, tags)
        render_html(fichas_niveis_html(cache_key, f"ua:{assinatura_tags(ids)}", "", "", 0,
                                       base[base["TAG"].isin(ids)], esperados)
                    + fichas_modais_html(ids, resumo, esperados, tags,
                                         espera_por_documento(
                                             _revisoes_por_doc(cache_key, sigem)),
                                         niveis_na_pagina=True)
                    + fichas_relatorios_pagina(cache_key, "ua", "", "", 0,
                                               ids, esperados, sigem))


def _ua_chave(rotulo: str) -> str:
    for chave, rot in UA_FRENTES.items():
        if rot == rotulo:
            return chave
    return rotulo.lower()


def ua_linha_do_tempo(movimentos: list, agora: pd.Timestamp, com_ficha: set) -> str:
    """O diário: o dia à esquerda, e cada movimentação como uma linha só.

    Uma subida de base move 70 TAGs para "Montado" de uma vez. Em linhas soltas
    isso é a mesma frase 70 vezes; agrupado pelo trajeto, é uma linha que diz
    "70 TAGs · Em Programação para Montado" com as TAGs dentro.
    """
    dias = []
    for m in movimentos:
        dia = m["quando"].date()
        if not dias or dias[-1][0] != dia:
            dias.append((dia, {}))
        dias[-1][1].setdefault((m["tipo"], m["campo"], m["de"], m["para"]), []).append(m)

    blocos, teto = [], 400
    for dia, grupos in dias:
        cartoes, quantos = [], 0
        for (tipo, campo, de, para), itens in grupos.items():
            quantos += len(itens)
            if len(cartoes) < teto:
                cartoes.append(ua_cartao(tipo, campo, de, para, itens, agora, com_ficha))
        blocos.append(
            f'<div class="ua-dia"><div class="ua-sel">'
            f'<div class="ua-num">{dia.day:02d}</div>'
            f'<div class="ua-mes">{UA_MESES[dia.month - 1]} {dia.year % 100:02d}</div>'
            f'<div class="ua-qtd">{br_num(quantos)} mov.</div></div>'
            f'<div class="ua-movs">{"".join(cartoes)}</div></div>')
    return "".join(blocos) or '<div class="ua-vazio">Nada mudou neste período.</div>'


def ua_cartao(tipo: str, campo: str, de: str, para: str, itens: list,
              agora: pd.Timestamp, com_ficha: set) -> str:
    """Uma movimentação: o que é, de onde saiu, para onde foi, e quem mexeu."""
    def alvo(m):
        nome, rot = m["objeto"], m["rotulo"]
        if m["tipo"] == "documento":
            # Quantas TAGs dependem deste documento vai junto do nome: a planta
            # do CHZ-350 é esperada por 272, e um cartão que junta vários
            # documentos não teria onde dizer isso de outro jeito.
            quantas = (f'<i class="ua-quantas">{br_num(m["tags"])}</i>'
                       if m.get("tags", 0) > 1 else "")
            return (f'<span class="ua-tag" title="{esc(nome)}">{esc(rot)}'
                    f'{quantas}</span>')
        if nome in com_ficha:
            return f'<a class="ua-tag" href="#{ficha_anchor(nome)}">{esc(nome)}</a>'
        return f'<span class="ua-tag">{esc(nome)}</span>'

    primeiro = itens[0]
    if len(itens) == 1:
        cabeca = alvo(primeiro)
        lista = ""
    else:
        _, plural = UA_PLURAL.get(tipo, ("TAG", "TAGs"))
        cabeca = f'<span class="ua-tag">{br_num(len(itens))} {plural}</span>'
        mostra = itens[:14]
        lista = ('<div class="ua-alvos">'
                 + "".join(alvo(m) for m in mostra)
                 + (f'<span class="ua-mais">e mais {br_num(len(itens) - 14)}</span>'
                    if len(itens) > 14 else "")
                 + "</div>")

    trajeto = (f'<span class="ua-de">{esc(de)}</span>'
               f'<span class="ua-seta">&rarr;</span>'
               f'<span class="ua-para {primeiro["tom"]}">{esc(para)}</span>'
               if de else
               f'<span class="ua-para {primeiro["tom"]}">{esc(para)}</span>')
    detalhe = primeiro.get("detalhe") or ""
    # Só no cartão de um documento só: quando são vários, cada um leva a sua
    # contagem colada, e repetir a do primeiro embaixo diria errado.
    afeta = ""
    if tipo == "documento" and len(itens) == 1 and primeiro.get("tags"):
        afeta = (f'<div class="ua-onde">afeta {br_num(primeiro["tags"])} '
                 f'{"TAGs" if primeiro["tags"] > 1 else "TAG"}</div>')
    return (f'<div class="ua-mov {primeiro["tom"]}"><div class="ua-cp">'
            f'<div>{cabeca}<span class="ua-frente {tipo}">'
            f'{esc(UA_FRENTES.get(tipo, tipo))}</span></div>'
            f'<div class="ua-fato"><b>{esc(campo)}</b> {trajeto}'
            + (f' <span class="ua-onde" style="display:inline">· {detalhe}</span>'
               if detalhe else "")
            + f'</div>{lista}{afeta}</div>'
            f'<div class="ua-dir"><div class="ua-quando">'
            f'{ua_quando(primeiro["quando"], agora)}</div></div></div>')


def render_planta(tags: pd.DataFrame, resumo: pd.DataFrame, locacao: pd.DataFrame,
                  aux: pd.DataFrame, esperados: pd.DataFrame, sigem: pd.DataFrame,
                  cache_key: str):
    """O avanco de montagem por area, desenhado sobre o arranjo da unidade.

    A aba e so o visual e o resumo: serve para bater o olho e ver onde a
    montagem anda e onde parou. O detalhe por TAG continua na Progresso.
    """
    render_header("Planta")

    mapa = carregar_mapa()
    areas, plantas, area_do_desenho, base = dados_por_area(tags, resumo, locacao, aux)

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
    # Sete faixas, na mesma tabela que o Daniel definiu -- laranja a verde
    # escuro, do menos ao mais avançado.
    faixas = [("pl1", "<20%", "avanço da área", lambda a: a["pct"] < 20),
              ("pl2", ">=20% e <40%", "avanço da área", lambda a: 20 <= a["pct"] < 40),
              ("pl3", ">=40% e <60%", "avanço da área", lambda a: 40 <= a["pct"] < 60),
              ("pl4", ">=60% e <70%", "avanço da área", lambda a: 60 <= a["pct"] < 70),
              ("pl5", ">=70% e <80%", "avanço da área", lambda a: 70 <= a["pct"] < 80),
              ("pl6", ">=80% e <90%", "avanço da área", lambda a: 80 <= a["pct"] < 90),
              ("pl7", ">=90%", "avanço da área", lambda a: a["pct"] >= 90)]
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
        <div class="pl-kpi"><div class="top"><div class="r">Montagem nas áreas</div>
          <span class="ic fxc-ambar">{fx_svg("seta")}</span></div>
          <div class="v andando">{br_pct(pct_mont)}</div>
          <div class="s">{br_num(mont)} de {br_num(total)} instrumentos montados</div>
          <div class="pl-barra"><i class="andando" style="width:{pct_mont:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="top"><div class="r">Valor montado</div>
          <span class="ic fxc-roxo">{fx_svg("moeda")}</span></div>
          <div class="v dinheiro">{br_moeda(val_m)}</div>
          <div class="s">{br_pct(pct_valor)} de {br_moeda(val_t)} nas áreas mapeadas</div>
          <div class="pl-barra"><i class="feito" style="width:{pct_valor:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="top"><div class="r">Montado sem medir</div>
          <span class="ic fxc-rubi">{fx_svg("relogio")}</span></div>
          <div class="v dinheiro parado">{br_moeda(val_sem)}</div>
          <div class="s">{br_num(mont - n_med)} instrumentos ·
            {br_pct(pct_sem)} do montado</div>
          <div class="pl-barra"><i class="parado" style="width:{pct_sem:.1f}%"></i></div></div>
        <div class="pl-kpi"><div class="top"><div class="r">Montado e medido no GITEC</div>
          <span class="ic fxc-verde">{fx_svg("ok")}</span></div>
          <div class="v dinheiro feito">{br_moeda(val_med)}</div>
          <div class="s">{br_num(n_med)} instrumentos ·
            {br_pct(pct_med)} do montado</div>
          <div class="pl-barra"><i class="feito" style="width:{pct_med:.1f}%"></i></div></div>
      </div>""")

    por_id = {p["id"]: p for p in mapa["pranchas"]}

    # A subestacao e um desenho em pe -- numa faixa larga ela viraria uma torre
    # de mil pixels, entao vai para uma coluna estreita ao lado. Essa coluna
    # acompanha a principal desde o topo (a principal entra dentro de col_a,
    # e nao antes dela) -- igual ficou no mapa de infra de instrumentacao, com
    # a subestacao ao lado do mapa inteiro, nao so da faixa de baixo.
    resto = [p for p in mapa["pranchas"] if p["id"] != "principal"]
    largas = [p for p in resto if p["prop"] < 100]
    altas = [p for p in resto if p["prop"] >= 100]
    if altas:
        col_a, col_b = st.columns([4, 1], gap="medium")
    else:
        col_a, col_b = st.container(), None
    with col_a:
        if "principal" in por_id:
            render_html_pesado(planta_prancha_html(por_id["principal"], areas, plantas, area_do_desenho))
        for p in largas:
            render_html_pesado(planta_prancha_html(p, areas, plantas, area_do_desenho))
        render_html(planta_legenda_html(faixas, contagem))
    if col_b is not None:
        with col_b:
            for p in altas:
                render_html_pesado(planta_prancha_html(p, areas, plantas, area_do_desenho))

    # As fichas ficam no fim da pagina, fechadas: o :target so abre a que a
    # zona clicada aponta. Sem todas geradas, o clique cairia no vazio.
    render_html_pesado(planta_fichas_html(mapa, areas, plantas, area_do_desenho, base))
    # E as fichas das TAGs listadas dentro delas: sem estas na pagina, clicar
    # num instrumento da planta nao teria para onde abrir e viraria ida para
    # outra aba.
    ids = [t for t in dict.fromkeys(base["TAG"].astype(str))
           if t in set(resumo["TAG"].astype(str))]
    if ids:
        # Fechar a ficha do instrumento devolve para a ficha da planta de onde
        # ele foi aberto, em vez de fechar tudo. A TAG pertence a uma area so,
        # e a area a uma zona -- entao da para dizer de onde ela veio.
        zonas_area = planta_zonas_por_area(mapa, area_do_desenho)
        volta = {}
        for tag, area_tag in zip(base["TAG"].astype(str), base["_area"]):
            zonas = zonas_area.get(area_tag)
            if zonas:
                volta[tag] = f'#{_ancora("PLANTA", zonas[0])}'
        # Sem fichas de relatorio aqui: sao 2.772 para as 2.098 TAGs da
        # unidade, e a pagina saia de 8 s para minutos. Sem elas o botao
        # Detalhes fica inerte -- com_relatorio=False --, melhor que um clique
        # que nao leva a lugar nenhum. O detalhe do relatorio esta na aba
        # Relatorios, que e onde ele cabe.
        # O degrau de nivel abre aqui mesmo, como em toda ficha do projeto:
        # fechar aquela ficha volta para o pai dela na trilha (a fichas de
        # nivel sao poucas -- uma por Fase/SOP/SSOP/Malha distinta entre os
        # instrumentos visiveis -- entao nao pesa como as de relatorio pesavam.
        base_niveis = progresso_base(resumo, tags)
        render_html_pesado(
            fichas_niveis_html(cache_key, f"planta:{assinatura_tags(ids)}", "", "", 0,
                               base_niveis[base_niveis["TAG"].isin(ids)], esperados)
            + fichas_modais_html(ids, resumo, esperados, tags,
                                 niveis_na_pagina=True, volta_por_tag=volta,
                                 com_relatorio=False))


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

    # os dois primeiros degraus levam a ficha da area; o terceiro e onde estamos
    alvo_area = f'#{_ancora("AREA", area["area"])}'
    trilha = [(f'Área {area["area"]}', alvo_area), (area["nome"], alvo_area),
              (desenho, "")]

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
            f'<td>{tag_link(r["TAG"])}</td>'
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

    # Fechar aqui nao pode ir sempre para "nada": esta ficha tambem e
    # alcancada de dentro da ficha de Area (lista "Plantas da area"), e
    # fechar tem que desfazer so esse ultimo passo, nao a navegacao inteira.
    return (
        f'<div class="fmodal" id="{_ancora("PLANTA", desenho)}">'
        f'<a class="fmodal-bg" href="{alvo_area}" aria-label="Voltar"></a>'
        '<div class="fmodal-box"><div class="fmodal-head">'
        '<div><div class="fn-tipo">Planta</div>'
        f'<div class="fmodal-title">{esc(desenho)}</div></div>'
        '<div class="fn-avanco">'
        f'<div class="fn-track"><div class="fn-fill {tom}" style="width:{max(pct, 1.5):.1f}%;"></div></div>'
        f'<div class="fn-pct">{br_pct(pct)}</div></div>'
        f'<a class="fmodal-x" href="{alvo_area}" aria-label="Voltar">&times;</a></div>'
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


def planta_zonas_por_area(mapa: dict, area_do_desenho: dict) -> dict:
    """As zonas marcadas de cada area, pelo codigo que a ficha da planta usa."""
    por_area: dict[str, list[str]] = {}
    for prancha in mapa["pranchas"]:
        for z in prancha["zonas"]:
            a = zona_area(z, area_do_desenho)
            if a:
                cod = codigos_da_zona(z)
                if cod not in por_area.setdefault(a, []):
                    por_area[a].append(cod)
    return por_area


def planta_area_ficha_html(area: dict, sub: pd.DataFrame, zonas: list[str]) -> str:
    """A ficha da area: o degrau acima da planta, que a trilha ja anunciava.

    A trilha da planta dizia "Area 120 > Filtros e permutadores de cru > CHZ-323"
    com os dois primeiros em texto morto. Agora levam aqui, como o degrau da
    Progresso leva a ficha do SOP.
    """
    n = len(sub)
    montados = int(sub["_montado"].sum())
    pct = montados / n * 100 if n else 0.0
    completos = int((sub["_docfrac"] >= 1.0).sum())
    medidos = int((sub["_medido"] == "SIM").sum())
    doc = float(sub["_docfrac"].mean() * 100) if n else 0.0
    valor = float(sub["_preco"].sum())
    valor_mont = float(sub.loc[sub["_montado"], "_preco"].sum())
    tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")

    tiles = (
        fx_tile("Instrumentos", br_num(n), "tag", "#2dd4bf",
                f"{br_num(completos)} com documentação completa")
        + fx_tile("Montados", br_num(montados), "seta", "#34d399", br_pct(pct))
        + fx_tile("Medidos no Gitec", br_num(medidos), "check", "#60a5fa",
                  f"de {br_num(montados)} montados")
        + fx_tile("Plantas", br_num(len(zonas)), "grade", "#a78bfa",
                  "desenhos desta área"))

    # So os nomes, sem contagem por planta: a base localiza a TAG pela area,
    # nao pelo desenho, entao qualquer numero aqui seria o total da area
    # repetido em cada linha -- foi o que a primeira versao mostrou, 74 em cada
    # uma das duas plantas de uma area de 74 instrumentos.
    tabela = ('<div class="ua-alvos">'
              + "".join(f'<a class="ua-tag gtbl-link" href="#{_ancora("PLANTA", c)}">'
                        f"{esc(c)}</a>" for c in zonas)
              + "</div>") if zonas else (
        '<div class="gtbl-empty">Nenhuma planta marcada nesta área.</div>')

    dados = (fx_dado("Área", f'{area["area"]} · {area["nome"]}')
             + fx_dado("Plantas", ", ".join(zonas) or "—")
             + fx_dado("Valor montado", f"{br_moeda(valor_mont)} de {br_moeda(valor)}")
             + fx_dado("Avanço documental", br_pct(doc)))

    direita = fx_painel(
        "Montagem da área", "seta",
        fx_rosca(montados, n, "#34d399", "montado")
        + '<div class="fx-leg">'
        + fx_lg("Montados", br_num(montados), br_pct(pct), "#34d399")
        + fx_lg("A montar", br_num(n - montados), br_pct(100 - pct if n else 0), "#f87171")
        + fx_lg("Instrumentos", br_num(n), "", "#3a4a68", total=True)
        + "</div>", classe_corpo="centro")

    # Reachavel de qualquer uma das plantas da area: fechar tem que voltar
    # para a que estava aberta, nao apagar a navegacao inteira.
    return (
        f'<div class="fmodal" id="{_ancora("AREA", area["area"])}">'
        '<a class="fmodal-bg" href="#fechado" aria-label="Fechar"></a>'
        '<div class="fmodal-box"><div class="fmodal-head">'
        '<div><div class="fn-tipo">Área</div>'
        f'<div class="fmodal-title">{esc(area["area"])} · {esc(area["nome"])}</div></div>'
        '<div class="fn-avanco">'
        f'<div class="fn-track"><div class="fn-fill {tom}" style="width:{max(pct, 1.5):.1f}%;"></div></div>'
        f'<div class="fn-pct">{br_pct(pct)}</div></div>'
        '<a class="fmodal-x" href="#fechado" aria-label="Fechar">&times;</a></div>'
        f'<div class="fmodal-body"><div class="fx">'
        f'<div class="fx-tiles">{tiles}</div>'
        '<div class="fx-corpo"><div class="fx-col">'
        + fx_painel("Resumo da área", "grade", f'<div class="fx-dados">{dados}</div>')
        + fx_painel("Plantas da área", "tag", tabela,
                    conta=f"{br_num(len(zonas))} plantas")
        + f'</div><div class="fx-col">{direita}</div></div></div></div></div></div>'
    )


def planta_fichas_html(mapa: dict, areas: dict, plantas: dict, area_do_desenho: dict,
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
    # e a ficha de cada area, que e o degrau acima na trilha da planta
    zonas_area = planta_zonas_por_area(mapa, area_do_desenho)
    for a, zonas in zonas_area.items():
        if a in areas:
            partes.append(planta_area_ficha_html(areas[a], base[base["_area"] == a],
                                                 zonas))
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
    """Todo valor em reais da tela passa por aqui -- são 32 lugares, e é por
    isso que a máscara mora nesta função e não em cada um deles: valor novo
    que alguém escrever amanhã já nasce coberto.

    Quem não tem a permissão vê um traço. O número não é arredondado nem
    aproximado: ele não é escrito. Ver ficha_valores_ocultos() para o motivo
    de a permissão também entrar na chave de cache.
    """
    if not pode("ver_valores"):
        return "—"
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


def agrega_nivel(df: pd.DataFrame, esperados: pd.DataFrame, coluna: str,
                 subnivel: str | None = None) -> pd.DataFrame:
    """Avanco = soma(emitidos)/soma(esperados): um relatorio pesa o mesmo em
    qualquer nivel, coerente com o avanco do Dashboard e com medicao. A media
    dos subniveis daria outro numero (num SOP chegou a 14pp de diferenca),
    porque faria um SSOP de 1 TAG pesar como um de 173.

    esperados/emitidos saem de totais_por_documento_agrupado, por documento
    unico -- um documento pendurado em TAGs de mais de um FASE/SOP/SSOP/MALHA
    (o RIMII/RIMSI de infra por planta e o exemplo extremo) nao pode ser
    somado por TAG entre os grupos, senao infla o grupo inteiro.
    """
    agg = dict(
        tags=("TAG", "count"),
        completas=("COMPLETA", "sum"),
        valor=("PRECO_UNITARIO", "sum"),
        valor_avancado=("VALOR_AVANCADO", "sum"),
        prioridade=("SUBGRUPO_PRIORIDADE", prioridade_do_grupo),
    )
    if subnivel:
        agg["subniveis"] = (subnivel, "nunique")
    g = df.groupby(coluna).agg(**agg).reset_index()
    doc_g = totais_por_documento_agrupado(df, esperados, coluna)
    g["esperados"] = g[coluna].map(doc_g["esperados"]).fillna(0).astype(int)
    g["emitidos"] = g[coluna].map(doc_g["emitidos"]).fillna(0).astype(int)
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


def assinatura_tags(ids) -> str:
    """As TAGs pedidas, como texto, para entrar na chave de cache.

    fichas_niveis_html e cacheada por (cache_key, filtro, f_seg, f_malha,
    pag) -- o DataFrame em si (_df) nao entra no calculo, pela convencao do
    "_" na frente. Certificacao, Ultima atualizacao, Planta e a
    fichas_completas (Dashboard/Relatorios/Gitec) chamavam sempre com o mesmo
    filtro fixo (por exemplo "cert"), e a mesma TAG do cache_key inteiro:
    trocar de TAG na busca da Certificacao, ou abrir o Gitec depois do
    Dashboard, devolvia os niveis de quem tinha rodado primeiro, nao os do
    pedido atual. A Malha MI-YST-121100 do Gitec nunca aparecia por isso -- o
    Dashboard tinha gravado o cache primeiro, com TAGs diferentes.
    """
    return ",".join(sorted(str(i) for i in ids))


def _ancora(tipo: str, valor: str) -> str:
    """Id do modal de um nivel. So alfanumerico, para ser id CSS valido."""
    limpo = "".join(c if c.isalnum() else "-" for c in str(valor))
    return f"n-{tipo}-{limpo}"


# max_entries baixo de proposito: cada entrada guarda ate ~5 MB de HTML e o
# plano do Render tem 512 MB. Com 8 por funcao dava 83 MB so de cache.
@st.cache_data(show_spinner=False, max_entries=3)
def arvore_html(cache_key: str, filtro: str, f_seg: str, f_malha: str, pag: int,
                _df: pd.DataFrame, _esperados: pd.DataFrame) -> str:
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
    for _, fase in agrega_nivel(_df, _esperados, "FASE", subnivel="SOP").iterrows():
        d_fase = recorte(_df, "FASE", fase["FASE"])
        sops = []
        for _, sop in agrega_nivel(d_fase, _esperados, "SOP", subnivel="SSOP").iterrows():
            d_sop = recorte(d_fase, "SOP", sop["SOP"])
            ssops = []
            for _, ss in agrega_nivel(d_sop, _esperados, "SSOP", subnivel="MALHA").iterrows():
                d_ssop = recorte(d_sop, "SSOP", ss["SSOP"])
                malhas = []
                for _, ml in agrega_nivel(d_ssop, _esperados, "MALHA").iterrows():
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

    render_html(_totais(df, esperados))
    _graficos(df, esperados)

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
    arvore = arvore_html(cache_key, filtro, f_seg, f_malha, pag, df_pag, esperados)

    lugar.markdown(tela_carregando("Preparando as fichas de SOP, SSOP e malha", 55,
                                   coberta=False), unsafe_allow_html=True)
    niveis = fichas_niveis_html(cache_key, filtro, f_seg, f_malha, pag, df_pag, esperados)

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
                       _df: pd.DataFrame, _esperados: pd.DataFrame) -> str:
    """As fichas de SOP, SSOP e malha, todas fechadas."""
    partes = []
    for tipo, rotulo in (("FASE", "Fase"), ("SOP", "SOP"), ("SSOP", "SSOP"),
                         ("MALHA", "Malha")):
        for valor, sub in _df.groupby(tipo):
            partes.append(_modal_nivel(tipo, valor, sub, rotulo, _esperados))
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


def _totais(df: pd.DataFrame, esperados: pd.DataFrame) -> str:
    esp, _, emi, _ = totais_por_documento(esperados[esperados["TAG"].isin(set(df["TAG"]))])
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


def _modal_nivel(tipo: str, nome: object, sub: pd.DataFrame, rotulo_tipo: str,
                 esperados: pd.DataFrame) -> str:
    """Ficha da fase/SOP/SSOP/malha: o que esta cadastrado nele e o degrau
    seguinte da hierarquia."""
    valor_ancora = "(sem)" if vazio(nome) else str(nome)
    titulo = f"Sem {rotulo_tipo.lower()}" if vazio(nome) else str(nome)
    esp, pos, apr, pen = totais_por_documento(esperados[esperados["TAG"].isin(set(sub["TAG"]))])
    pct = (apr / esp * 100) if esp else 0
    tom = "ok" if pct >= 70 else ("warn" if pct >= 30 else "crit")
    completas = int(sub["COMPLETA"].sum())

    # ------------------------------------------- trilha: a cadeia ate aqui
    # A trilha sobe por ancora, sem recarregar nada -- vale em qualquer aba que
    # gere as fichas de nivel junto com a da TAG. So vale para o pai unico --
    # quando o nivel atravessa dois pais, "SOP-1 e mais 1" nao aponta para
    # ficha nenhuma, entao fica como texto.
    trilha = []
    for pai in FX_ACIMA.get(tipo, []):
        v = _distintos(sub[pai], 1)
        unico = sub[pai].nunique() == 1 and v not in ("—", "-", "")
        trilha.append((f"{FX_ROTULO_NIVEL[pai]} {v}",
                       f"#{_ancora(pai, v)}" if unico else ""))
    trilha.append((f"{FX_ROTULO_NIVEL[tipo]} {titulo}", ""))
    # Fechar sobe um degrau na mesma trilha -- Malha volta para a SSOP dela,
    # SSOP para o SOP, SOP para a Fase. E a mesma logica de toda outra ficha do
    # projeto: fechar volta para o pai, nao para "nada". So cai no fechado de
    # verdade quando nao ha um pai unico para voltar (a Fase, que e o topo, ou
    # um nivel com mais de um pai).
    pais_com_link = [href for _, href in trilha[:-1] if href]
    volta_nivel = pais_com_link[-1] if pais_com_link else "#fechado"

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
        corpo = _tabela_niveis(sub, esperados, filho, neto)
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

    rotulo_volta = "Fechar" if volta_nivel == "#fechado" else "Voltar"
    return (
        f'<div class="fmodal" id="{_ancora(tipo, valor_ancora)}">'
        f'<a class="fmodal-bg" href="{volta_nivel}" aria-label="{rotulo_volta}"></a>'
        '<div class="fmodal-box"><div class="fmodal-head">'
        f'<div><div class="fn-tipo">{esc(rotulo_tipo)}</div>'
        f'<div class="fmodal-title">{esc(titulo)}</div></div>'
        '<div class="fn-avanco">'
        f'<div class="fn-track"><div class="fn-fill {tom}" style="width:{max(pct, 1.5):.1f}%;"></div></div>'
        f'<div class="fn-pct">{br_pct(pct)}</div></div>'
        f'<a class="fmodal-x" href="{volta_nivel}" aria-label="{rotulo_volta}">&times;</a></div>'
        f'<div class="fmodal-body"><div class="fx">{fx_trilha(trilha)}'
        f'<div class="fx-tiles">{tiles}</div>'
        '<div class="fx-corpo"><div class="fx-col">'
        + fx_painel("Resumo documental", "grade",
                    f'<div class="fx-kpis">{kpis}</div><div class="fx-dados">{dados}</div>')
        + fx_painel(rotulo_sub, "livro" if tipo != "MALHA" else "tag", corpo,
                    conta=conta, classe_corpo="zero")
        + f'</div><div class="fx-col">{direita}</div></div></div></div></div></div>'
    )


def _tabela_niveis(sub: pd.DataFrame, esperados: pd.DataFrame, coluna: str,
                   subnivel: str = None) -> str:
    """Os filhos diretos de um nivel: o que o SOP mostra dos SSOPs dele.

    Cada linha leva a propria ficha, entao dentro da ficha do SOP da para
    descer para a do SSOP sem fechar nada.
    """
    rotulo = {"SOP": "SOP", "SSOP": "SSOP", "MALHA": "Malha"}[coluna]
    conta_sub = {"SSOP": "#SSOPs", "MALHA": "#Malhas"}.get(subnivel)
    linhas = []
    for _, r in agrega_nivel(sub, esperados, coluna, subnivel=subnivel).iterrows():
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


def painel_previsao_fornecimento(status_final: object, previsao: object,
                                 status_fornecimento: object = None) -> str:
    """O card "Previsão de fornecimento" da ficha da TAG.

    Só existe para ON DEMAND e EM COMPRA -- nos outros estados a coluna vem
    vazia na base, e mostrar o card assim mesmo sugeriria que falta preencher
    algo, quando na verdade não se aplica.

    Concluído (status_fornecimento) é uma pergunta diferente de atrasado: o
    material já chegou, e a data é so informativa -- mesmo no passado, não e
    atraso. So quando NAO esta concluido e que uma data no passado quer dizer
    atraso de verdade. Sem essa distincao, um material que chegou no prazo
    (mas com Real no passado, que e a natureza de "ja aconteceu") aparecia
    como atrasado igual a um que realmente esta parado.
    """
    if str(status_final).strip().upper() not in EM_COMPRA:
        return ""

    d = pd.to_datetime(previsao, dayfirst=True, errors="coerce")
    concluido = str(status_fornecimento or "").strip().upper() == "CONCLUIDO"

    if concluido:
        corpo = (fx_linha("Situação", '<span class="gtbl-badge ok">Recebido</span>')
                 + fx_linha("Data", f"<b>{d:%d/%m/%Y}</b>" if not pd.isna(d)
                            else '<span class="gtbl-muted">—</span>'))
        return fx_painel("Previsão de fornecimento", "caixa", corpo)

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


def pill_ssop_prioritario(v: object) -> str:
    """SIM/NÃO/CANCELADO da coluna SSOP_PRIORITARIO. Só o SIM pede atenção
    (âmbar); os outros dois ficam neutros -- não são um problema, só não
    entram no recorte de prioridade."""
    if vazio(v):
        return '<span class="gtbl-muted">—</span>'
    t = str(v).strip()
    tom = "warn" if t.upper() == "SIM" else "mudo"
    return f'<span class="gtbl-badge {tom}">{esc(t)}</span>'


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






def _graficos(df: pd.DataFrame, esperados: pd.DataFrame):
    """Os quatro recortes mais avancados: onde ha mais chance de fechar rapido.

    Um grid CSS unico, e nao st.columns: cada coluna do Streamlit empilha de
    forma independente, entao com poucos itens as colunas ficavam com alturas
    diferentes (chegou a 218px de desequilibrio) e abria um vao no meio.
    """
    blocos = "".join([
        grafico_avanco("Fases mais avançadas", agrega_nivel(df, esperados, "FASE", subnivel="SOP"),
                       "FASE", rotulo_sub="SOP"),
        grafico_avanco("SOP mais avançados", agrega_nivel(df, esperados, "SOP", subnivel="SSOP"),
                       "SOP", rotulo_sub="SSOP"),
        grafico_avanco("SSOP mais avançados", agrega_nivel(df, esperados, "SSOP", subnivel="MALHA"),
                       "SSOP", rotulo_sub="malhas"),
        grafico_avanco("Malhas mais avançadas", agrega_nivel(df, esperados, "MALHA"), "MALHA"),
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
# Os quatro que Daniel quer na lateral por enquanto. Os outros continuam
# implementados em _tags_do_filtro -- e so devolver a linha aqui para eles
# voltarem, sem mexer em mais nada:
#     ("fase", "Fase", "Todas", "tags", "FASE"),
#     ("subgrupo_prioridade", "Subgrupo de Prioridade", "Todos", "tags", "SUBGRUPO_PRIORIDADE"),
#     ("grupo", "Grupo de instrumento", "Todos", "resumo", "GRUPO_REGRA"),
#     ("status", "Status SIGEM", "Todos", "esperados", "STATUS_SIGEM"),
FILTROS = [
    ("sop", "SOP", "Todos", "tags", "SOP"),
    ("ssop_prioritario", "SSOP Prioritário", "Todos", "tags", "SSOP_PRIORITARIO"),
    ("skid", "SKID", "Todos", "tags", "SKID"),
    ("cff", "CFF", "Todos", "tags", "CFF"),
]


def _valores(serie: pd.Series) -> list:
    # "0" entra na lista de descarte por causa do SKID: a base escreve 0 nas
    # 363 TAGs que nao pertencem a skid nenhum, do mesmo jeito que escreve "-"
    # no CFF de quem nao e fieldbus. Sem isso o filtro oferecia um "0" que nao
    # quer dizer nada. Nenhuma das outras colunas filtraveis usa 0 como valor.
    limpo = serie.dropna().astype(str).str.strip()
    return sorted({v for v in limpo
                   if v and v.lower() not in ("nan", "-", "none", "0")})


def _tags_do_filtro(chave: str, valor: str, tags, resumo, esperados) -> set:
    """TAGs que atendem a UM filtro. O status e por tag, nao por relatorio:
    'Recusado' devolve as tags que tem ao menos um relatorio recusado. Filtrar
    a lista de relatorios em si mudaria o denominador do avanco e as contas da
    tela passariam a se contradizer."""
    if chave == "fase":
        return set(tags.loc[tags["FASE"].astype(str).str.strip() == valor, "TAG"])
    if chave == "sop":
        return set(tags.loc[tags["SOP"].astype(str).str.strip() == valor, "TAG"])
    if chave == "ssop_prioritario":
        return set(tags.loc[tags["SSOP_PRIORITARIO"].astype(str).str.strip() == valor, "TAG"])
    if chave == "subgrupo_prioridade":
        return set(tags.loc[tags["SUBGRUPO_PRIORIDADE"].astype(str).str.strip() == valor, "TAG"])
    if chave == "skid":
        return set(tags.loc[tags["SKID"].astype(str).str.strip() == valor, "TAG"])
    if chave == "cff":
        return set(tags.loc[tags["CFF"].astype(str).str.strip() == valor, "TAG"])
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


def _limpar_filtros():
    """Callback do botão. Precisa ser callback: mexer na chave de um widget
    depois que ele já foi criado no mesmo run levanta exceção no Streamlit --
    os callbacks rodam antes do rerun, quando ainda é permitido."""
    for chave, _, padrao, _, _ in FILTROS:
        st.session_state[f"gf_{chave}"] = padrao


def sidebar_filtros(tags: pd.DataFrame, resumo: pd.DataFrame,
                    esperados: pd.DataFrame) -> dict:
    """Filtro rápido e avançado da lateral, em cascata e válido para o app
    inteiro -- todas as abas passam a ver só o recorte escolhido, como se a
    base fosse só aquilo, até o filtro ser desfeito.

    "Avançado" é a combinação: cada campo só oferece o que ainda sobra depois
    dos OUTROS já escolhidos -- marcar SSOP Prioritário = Sim reduz os
    Subgrupos de Prioridade aos que aparecem só nessas tags, e vice-versa.
    Sem isso dava para montar uma combinação que não devolve nada e parece
    defeito. Quando o valor guardado sai da lista -- porque outro filtro
    mudou -- ele volta ao padrão em vez de estourar.

    A chave do session_state É a chave do widget, de propósito: com duas
    chaves o valor vindo da URL era escrito numa e o selectbox continuava
    lendo a outra, e o seletor não mexia. Os links que a ficha gera -- "ver
    na Progresso" com ?fase= -- continuam funcionando por essa mesma porta,
    via consumir_filtros_url.
    """
    consumir_filtros_url(tags, resumo, esperados)

    def escolhido(chave: str, padrao: str) -> str:
        v = st.session_state.get(f"gf_{chave}", padrao)
        return "" if v == padrao else v

    # O rotulo do expansor precisa do recorte ANTES de desenhar os seletores
    # -- ele usa o valor que cada widget ja tem guardado do run anterior, o
    # mesmo que os proprios seletores vao ler duas linhas abaixo.
    escolhas_agora = {c: escolhido(c, p) for c, _, p, _, _ in FILTROS}
    ativos_agora = {c: v for c, v in escolhas_agora.items() if v}
    alvo_agora = _universo(escolhas_agora, tags, resumo, esperados)
    rotulo_expansor = f"Filtro rápido · {br_num(len(alvo_agora))} de {br_num(len(tags))}"

    # Fechado por padrao -- com os 6 campos sempre abertos, menu + filtro
    # passava de 800px de altura e forcava rolagem dentro da lateral. Abre
    # sozinho quando ja tem filtro escolhido, pra nao esconder o que esta
    # ativo atras de um clique. Nao usa st.expander: ele nao respeita a
    # mesma reordenacao por flex que marca e perfil respeitam (o filtro
    # aparecia antes da marca, nao depois do menu) -- um botao comum, do
    # mesmo tipo do "Abrir o perfil" logo abaixo, se comporta certo.
    # So forca aberto na TRANSICAO de nenhum filtro pra algum (por exemplo,
    # um link "ver na Progresso" que chega com ?fase=): olhar so "esta
    # ativo agora" reabria sozinho a cada rerun, e fechar manualmente com um
    # filtro ja escolhido nunca pegava -- a caixa voltava a abrir na mesma
    # hora.
    tinha_antes = st.session_state.get("_flt_tinha_ativo", False)
    if "_flt_aberto" not in st.session_state:
        st.session_state["_flt_aberto"] = bool(ativos_agora)
    elif ativos_agora and not tinha_antes:
        st.session_state["_flt_aberto"] = True
    st.session_state["_flt_tinha_ativo"] = bool(ativos_agora)

    def _alternar_filtro():
        st.session_state["_flt_aberto"] = not st.session_state["_flt_aberto"]

    with st.sidebar:
        seta = "▾" if st.session_state["_flt_aberto"] else "▸"
        st.button(f"{seta}  {rotulo_expansor}", key="flt_toggle",
                 on_click=_alternar_filtro, use_container_width=True)

        if st.session_state["_flt_aberto"]:
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
            if ativos:
                st.button("Limpar filtros", key="gf_limpar", on_click=_limpar_filtros)
        else:
            escolhas = escolhas_agora
            ativos = ativos_agora

    alvo = _universo(escolhas, tags, resumo, esperados)
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


@st.dialog("Meu perfil")
def dialogo_perfil():
    """Foto, e-mail, senha e a saída -- tudo do próprio dono.

    Fica num diálogo, e não numa página, porque perfil não é destino: a pessoa
    troca a foto e volta para onde estava, sem perder a aba aberta.
    """
    eu = st.session_state.get("gplan_usuario") or {}
    if not eu:
        st.error("Sessão expirada. Entre de novo.")
        return

    papel = acesso.papel_de(eu)
    render_html(
        '<div class="pf-topo">'
        + (f'<img class="pf-foto" src="{eu["foto"]}" alt="">' if eu.get("foto")
           else f'<span class="pf-ini">'
                f'{esc(acesso.iniciais(eu.get("nome", ""), eu.get("email", "")))}</span>')
        + f'<div><div class="pf-nome">{esc(eu.get("nome") or eu.get("email"))}'
        f'<span class="pf-papel {acesso.COR_PAPEL.get(papel, "teal")}">'
        f'{esc(papel)}</span></div>'
        f'<div class="pf-login">{esc(eu.get("email", ""))}</div></div></div>')

    permitidas = [acesso.PERMISSOES[p] for p in eu.get("permissoes", [])
                  if p in acesso.PERMISSOES]
    render_html('<div class="pf-perms">'
                + "".join(f'<span class="pf-tag">{esc(x)}</span>' for x in permitidas)
                + ("" if permitidas else
                   '<span class="pf-tag vazio">sem permissões</span>')
                + "</div>")

    with st.expander("Dados e foto"):
        nome = st.text_input("Nome", value=eu.get("nome", ""), key="pf_nome")
        foto = st.file_uploader("Foto", type=["png", "jpg", "jpeg"], key="pf_foto")
        if st.button("Salvar dados", type="primary", key="pf_salvar"):
            campos = {"nome": nome.strip()}
            if foto is not None:
                bruto = foto.getvalue()
                # a foto viaja em toda leitura de perfil: grande, deixaria o
                # login lento para quem só quer entrar
                if len(bruto) > 400_000:
                    st.error("A foto precisa ter menos de 400 KB.")
                    st.stop()
                campos["foto"] = ("data:" + (foto.type or "image/png") + ";base64,"
                                  + base64.b64encode(bruto).decode())
            try:
                acesso.salvar_perfil(eu["id"], **campos)
                recarregar_usuario()
                st.rerun()
            except Exception as erro:
                avisar_erro("salvar o perfil", erro)

    with st.expander("Trocar a senha"):
        atual = st.text_input("Senha atual", type="password", key="pf_atual")
        nova = st.text_input("Nova senha", type="password", key="pf_nova",
                             help=acesso.REGRA_SENHA)
        repete = st.text_input("Repita a nova senha", type="password", key="pf_rep")
        st.caption(acesso.REGRA_SENHA)
        if st.button("Trocar a senha", type="primary", key="pf_trocar"):
            if nova != repete:
                st.error("As senhas não conferem.")
            elif (falta := acesso.problema_na_senha(nova)):
                st.error(falta)
            else:
                try:
                    # conferir a atual é o que impede que uma sessão esquecida
                    # numa máquina destrancada vire troca de senha por quem passar
                    if acesso.trocar_minha_senha(eu["email"], atual, nova):
                        st.success("Senha trocada.")
                    else:
                        st.error("A senha atual não confere.")
                except Exception as erro:
                    avisar_erro("trocar a senha", erro)

    if st.button("Sair da conta", key="pf_sair", use_container_width=True):
        acesso.esquecer(st.session_state.get("gplan_token", ""))
        esquecer_sessao()
        st.session_state.pop("gplan_usuario", None)
        st.session_state.pop("gplan_token", None)
        st.rerun()


# ======================================================================= #
#  Previsão Medição -- o que a obra já fez, antes de o papel acompanhar
# ======================================================================= #
# Cada relatório esperado tem um FATO por trás, e é o fato que autoriza medir.
# TODO relatório da TAG entra na conta -- uma TAG de 10 relatórios só está
# pronta com os 10; o que estiver feito vira percentual de avanço.
#
# O mapa é por (RELATÓRIO, REGRA DE ORIGEM), não só pelo nome do relatório: o
# RIR aparece duas vezes com sentidos diferentes (o do instrumento é
# localização, o do cabo é lançamento), e o RIMSI também (suporte por planta é
# documento, pedestal tem base própria com avanço).
#
#   "calibração"  -> STATUS_CALIBRACAO == Aprovado
#   "localização" -> STATUS_LOCALIZACAO == LOCALIZADO
#   "montagem"    -> STATUS_MONTAGEM == Montado
#   "cabo"        -> todos os circuitos da TAG Concluído
#   "pedestal"    -> avanço 100% na 08_BASE_PEDESTAL
#   "documento"   -> o próprio relatório aprovado no SIGEM (não há fato de
#                    campo separado para ler; o documento É a evidência)
# A chave é (RELATÓRIO, trecho da regra de origem). Trecho, e não o texto
# inteiro: a regra do CCP se chama "BASE: CCP obrigatorio exceto Caixa de
# Juncao Fieldbus e PAINEL", e casar pelo texto completo faz a etapa cair
# calada no padrão quando o pipeline reescrever a frase. A ordem importa --
# a primeira que casar vale, e por isso o RIR de cabo vem antes do RIR seco.
MEDICAO_REGRA = [
    ("CCP", "", "calibração"),
    ("RIR", "cabo", "cabo"),
    ("RIR", "", "localização"),
    ("RIFMI", "", "montagem"),
    ("RILTCI", "", "cabo"),
    ("CTECRI", "", "cabo"),
    ("RIMTU", "", "documento"),
    ("RIMII", "", "documento"),
    ("RIMSI", "pedestal", "pedestal"),
    ("RIMSI", "", "documento"),
]


# As quatro famílias que o mapa não nomeia (RILM, RIMITPI, RIMJBI, RTFCJI)
# caem em "documento": sem fato de campo declarado para elas, quem responde é
# o próprio relatório aprovado. É o padrão conservador -- nunca dá uma TAG por
# pronta sem prova.
MEDICAO_PADRAO = "documento"


def medicao_etapa(relatorio: str, origem: str) -> str:
    """Que fato de campo aquele relatório prova, para aquela TAG."""
    alvo = origem.lower()
    for rel, trecho, etapa in MEDICAO_REGRA:
        if rel == relatorio and (not trecho or trecho in alvo):
            return etapa
    return MEDICAO_PADRAO


MEDICAO_ORDEM = ["calibração", "localização", "montagem", "cabo", "pedestal",
                 "documento"]


def vazio_para_traco(valor) -> str:
    """"nan", "-" e "0" da base viram vazio: os três querem dizer 'não tem'."""
    t = str(valor).strip()
    return "" if t.lower() in ("nan", "-", "none", "0") else t


@st.cache_data(show_spinner=False, max_entries=2)
def medicao_pedestal(cache_key: str) -> dict:
    """TAG -> avanço do pedestal dela, de 0 a 1.

    Vem da aba 08_BASE_PEDESTAL já normalizada pelo pipeline: DOCUMENTO, as
    doze TAGINSTR e o AVANCO da linha. O avanço é ponderado na origem --
    montagem do pedestal pesa 0,4, aterramento 0,25, grauteamento 0,2, pintura
    0,1 e abertura de TAG 0,05, somando 1.

    Uma linha de pedestal serve até doze instrumentos, e o avanço dela vale
    para todos. Uma TAG que apareça em mais de um pedestal fica com o PIOR:
    certificar exige que todos fechem, não um deles.

    Sem a coluna (planilha gerada por pipeline antigo) devolve vazio, e o
    critério do pedestal passa a reprovar em vez de derrubar a tela -- o que
    não se sabe não autoriza medir.
    """
    client = get_supabase_client()
    if client is not None:
        fonte = io.BytesIO(client.storage.from_(SUPABASE_BUCKET)
                           .download(SUPABASE_FILE_PATH))
    elif os.path.exists(LOCAL_EXCEL_FALLBACK):
        fonte = LOCAL_EXCEL_FALLBACK
    else:
        return {}
    try:
        base = pd.read_excel(fonte, sheet_name="08_BASE_PEDESTAL")
    except Exception:
        return {}
    if base.empty or "AVANCO" not in base.columns:
        return {}
    colunas = [c for c in base.columns if str(c).upper().startswith("TAGINSTR")]
    saida: dict[str, float] = {}
    for r in base.to_dict("records"):
        try:
            a = float(r.get("AVANCO") or 0)
        except (TypeError, ValueError):
            a = 0.0
        for c in colunas:
            t = vazio_para_traco(r.get(c))
            if t:
                saida[t] = min(saida.get(t, 1.0), a)
    return saida


@st.cache_data(show_spinner=False, max_entries=3)
def medicao_prontidao(tags: pd.DataFrame, resumo: pd.DataFrame,
                      esperados: pd.DataFrame, lanc: pd.DataFrame,
                      depara: pd.DataFrame, cache_key: str) -> list:
    """Uma linha por TAG: quanto das etapas dela o campo já fechou.

    O portão é POR TAG e sobre TODOS os relatórios que ela espera -- uma TAG
    de dez relatórios só está pronta com os dez. Cada relatório é avaliado
    pela regra de origem DELE, não pelo nome da família: o mesmo RIR é
    localização numa linha e cabo na outra.

    A etapa guarda "quantos passaram de quantos", não um sim/não. Uma TAG com
    quatro relatórios de documento e três aprovados tinha a etapa contada como
    feita e como pendente ao mesmo tempo, e a tela pintava de verde -- porque
    olhava a lista de feitas primeiro. Agora verde exige o total.
    """
    if tags.empty:
        return []
    texto = lambda v: str(v).strip()
    circ_da_ponta, _ = cert_circuitos_por_ponta(lanc, depara, cache_key)
    status_circ = {i: texto(r["STATUS"])
                   for i, r in enumerate(lanc.to_dict("records"))}
    pedestal = medicao_pedestal(cache_key)

    medidas, valor_medido = set(), {}
    if "MEDIDO_GITEC" in resumo.columns:
        for r in resumo.to_dict("records"):
            if texto(r.get("MEDIDO_GITEC")).upper() == "SIM":
                t = texto(r["TAG"])
                medidas.add(t)
                valor_medido[t] = cert_num(r.get("VALOR_GITEC"))

    def cabo_ok(tag: str) -> bool:
        # sem circuito cadastrado não é "pronto": é desconhecido, e
        # desconhecido não autoriza medição
        ids = circ_da_ponta.get(tag, set())
        return bool(ids) and all(status_circ.get(i) == "Concluído" for i in ids)

    por_tag: dict[str, list] = {}
    for r in esperados.to_dict("records"):
        por_tag.setdefault(texto(r["TAG"]), []).append(r)

    info = {texto(r["TAG"]): r for r in tags.to_dict("records")}
    linhas = []
    for tag, docs in por_tag.items():
        t = info.get(tag)
        if t is None:
            continue
        etapas: dict[str, list] = {}
        vistos = set()
        for d in docs:
            doc = texto(d["DOCUMENTO_ESPERADO"])
            if doc in vistos:          # o mesmo documento não conta duas vezes
                continue
            vistos.add(doc)
            etapa = medicao_etapa(texto(d["RELATORIO"]),
                                  texto(d["ORIGEM_REGRA"]))
            if etapa == "calibração":
                passou = texto(t["STATUS_CALIBRACAO"]) == "Aprovado"
            elif etapa == "localização":
                passou = texto(t["STATUS_LOCALIZACAO"]) == "LOCALIZADO"
            elif etapa == "montagem":
                passou = texto(t["STATUS_MONTAGEM"]) == "Montado"
            elif etapa == "cabo":
                passou = cabo_ok(tag)
            elif etapa == "pedestal":
                passou = pedestal.get(tag, 0.0) >= 0.999
            else:
                passou = bool(aprovado(pd.Series([d["STATUS_SIGEM"]])).iloc[0])
            marca = etapas.setdefault(etapa, [0, 0])
            marca[1] += 1
            marca[0] += 1 if passou else 0
        if not etapas:
            continue
        feitos = sum(v[0] for v in etapas.values())
        total = sum(v[1] for v in etapas.values())
        linhas.append({
            "tag": tag,
            "descricao": " ".join(texto(t.get("DESCRICAO", "")).split()),
            "etapas": etapas,
            "n_feito": feitos,
            "total": total,
            "falta": total - feitos,
            "pct": feitos / total * 100,
            "fechadas": [e for e, v in etapas.items() if v[0] == v[1]],
            "abertas": [e for e, v in etapas.items() if v[0] < v[1]],
            "medida": tag in medidas,
            "valor_medido": valor_medido.get(tag, 0.0),
            "preco": cert_num(t.get("PRECO_UNITARIO")),
            "sop": vazio_para_traco(t.get("SOP")),
            "skid": vazio_para_traco(t.get("SKID")),
        })
    return linhas


def render_previsao_medicao(tags: pd.DataFrame, resumo: pd.DataFrame,
                            esperados: pd.DataFrame, lanc: pd.DataFrame,
                            depara: pd.DataFrame, sigem: pd.DataFrame,
                            cache_key: str = ""):
    render_header("Previsão Medição")
    linhas = medicao_prontidao(tags, resumo, esperados, lanc, depara, cache_key)
    if not linhas:
        render_html('<div class="gplan-panel"><div class="gtbl-empty">'
                    "Nenhuma TAG com relatório esperado neste recorte."
                    "</div></div>")
        return

    # O olho é só desta aba: tira os reais da vista sem mexer na permissão de
    # quem está logado. Quem não tem ver_valores continua vendo traço com o
    # olho aberto -- ele esconde, não libera.
    olho = st.toggle("Mostrar valores", value=True, key="med_olho",
                     help="Oculta os valores em reais desta tela")

    def moeda(v: float) -> str:
        return br_moeda(v) if olho else "•••"

    # A aba é PREVISÃO: o que já foi medido sai da fila e vira um recorte à
    # parte, senão o topo do ranking fica ocupado por quem não tem mais nada a
    # fazer.
    medidas = [l for l in linhas if l["medida"]]
    fila = [l for l in linhas if not l["medida"]]
    prontas = [l for l in fila if l["falta"] == 0]
    uma = [l for l in fila if l["falta"] == 1]
    feitas_tot = sum(l["n_feito"] for l in fila)
    etapas_tot = sum(l["total"] for l in fila)
    avanco = feitas_tot / max(etapas_tot, 1) * 100

    render_html(f"""
      <div class="pl-kpis">
        <div class="pl-kpi"><div class="r">Prontas para medir</div>
          <div class="v feito">{br_num(len(prontas))}</div>
          <div class="s">{moeda(sum(l["preco"] for l in prontas))} · todas as etapas fechadas</div></div>
        <div class="pl-kpi"><div class="r">A uma etapa</div>
          <div class="v andando">{br_num(len(uma))}</div>
          <div class="s">{moeda(sum(l["preco"] for l in uma))} · logo atrás</div></div>
        <div class="pl-kpi"><div class="r">Avanço das etapas</div>
          <div class="v">{br_pct(avanco)}</div>
          <div class="s">{br_num(feitas_tot)} de {br_num(etapas_tot)} etapas</div>
          <div class="pl-barra"><i class="andando" style="width:{avanco:.1f}%"></i></div>
        </div>
        <div class="pl-kpi"><div class="r">Já medido</div>
          <div class="v">{br_num(len(medidas))}</div>
          <div class="s">{moeda(sum(l["valor_medido"] for l in medidas))} · fora da fila</div></div>
      </div>""")

    # --- os dois gráficos ---
    por_falta = collections.Counter(min(l["falta"], 5) for l in fila)
    maior = max(por_falta.values()) if por_falta else 1
    barras_fila = ""
    for k in sorted(por_falta):
        n = por_falta[k]
        rot = ("Pronta" if k == 0 else
               "Falta 1 etapa" if k == 1 else
               f"Faltam {k} etapas" if k < 5 else "Faltam 5 ou mais")
        cor = "feito" if k == 0 else ("andando" if k == 1 else "parado")
        val = sum(l["preco"] for l in fila if min(l["falta"], 5) == k)
        barras_fila += (
            f'<div class="du-br moeda"><span class="nm">{rot}</span>'
            f'<span class="tr"><i class="{cor}" style="width:'
            f'{max(n / maior * 100, 0.8):.1f}%"></i></span>'
            f'<span class="fr">{br_num(n)}</span>'
            f'<span class="pc">{moeda(val)}</span></div>')

    trava = collections.Counter(e for l in fila for e in l["abertas"])
    maior_t = max(trava.values()) if trava else 1
    barras_trava = ""
    for etapa, n in trava.most_common():
        barras_trava += (
            f'<div class="du-br"><span class="nm">{esc(etapa.capitalize())}</span>'
            f'<span class="tr"><i class="parado" style="width:'
            f'{max(n / maior_t * 100, 0.8):.1f}%"></i></span>'
            f'<span class="fr">{br_num(n)}</span>'
            f'<span class="pc">{br_pct(n / max(len(fila), 1) * 100)}</span></div>')

    render_html(
        '<div class="cs-duas">'
        '<div class="du-pn"><div class="du-t">Quanto falta para cada TAG</div>'
        f'<div class="du-miolo"><div class="du-barras">{barras_fila}</div></div></div>'
        '<div class="du-pn"><div class="du-t">Etapa que mais segura</div>'
        f'<div class="du-miolo"><div class="du-barras">{barras_trava}</div></div></div>'
        '</div>')

    # --- um grupo por vez: a lista misturada era o que não deixava ler ---
    GRUPOS = {
        "Prontas": prontas,
        "Falta 1 etapa": uma,
        "Faltam 2": [l for l in fila if l["falta"] == 2],
        "Faltam 3 ou mais": [l for l in fila if l["falta"] >= 3],
        "Já medidas": medidas,
    }
    escolha = st.segmented_control(
        "Recorte", list(GRUPOS),
        format_func=lambda k: f"{k} · {br_num(len(GRUPOS[k]))}",
        default="Prontas", key="med_grupo") or "Prontas"
    grupo = GRUPOS[escolha]

    TETO = 150
    if escolha == "Já medidas":
        ordem = sorted(grupo, key=lambda l: (-l["valor_medido"], l["tag"]))
    else:
        ordem = sorted(grupo, key=lambda l: (l["falta"], -l["n_feito"],
                                             -l["preco"], l["tag"]))
    mostradas = ordem[:TETO]

    corpo = ""
    for l in mostradas:
        cor = ("feito" if l["falta"] == 0 else
               "andando" if l["falta"] == 1 else "parado")
        selos = ""
        for e in MEDICAO_ORDEM:
            if e not in l["etapas"]:
                continue
            ok, tot = l["etapas"][e]
            classe = "ok" if ok == tot else ("warn" if ok else "crit")
            quanto = "" if tot == 1 else f" {ok}/{tot}"
            selos += (f'<span class="gtbl-badge {classe}">'
                      f'{esc(e.capitalize())}{quanto}</span> ')
        valor = (moeda(l["valor_medido"]) if escolha == "Já medidas"
                 else moeda(l["preco"]))
        corpo += (
            f'<tr><td>{tag_link(l["tag"])}</td>'
            f'<td class="med-av"><span class="med-frac">{l["n_feito"]}'
            f'<span class="med-de">/{l["total"]}</span></span>'
            f'<span class="med-barra">'
            f'<i class="{cor}" style="width:{l["pct"]:.0f}%"></i></span>'
            f'<span class="med-pct">{br_pct(l["pct"], 0)}</span></td>'
            f'<td class="med-selos">{selos}</td>'
            f'<td class="gtbl-mono">{valor}</td></tr>')

    rodape = ""
    if len(ordem) > TETO:
        rodape = (f' · mostrando as {TETO} primeiras de {br_num(len(ordem))}')
    cabecalho_valor = "Valor medido" if escolha == "Já medidas" else "Valor"
    render_html(
        '<div class="gplan-panel ct-painel"><div class="gplan-panel-title">'
        f'{esc(escolha)}'
        f'<span class="gtbl-muted" style="font-weight:500">'
        f'{br_num(len(ordem))} TAGs{rodape} · clique na TAG para abrir a ficha'
        '</span></div><div class="ct-rolo">'
        '<table class="gtbl med-tab"><thead><tr><th>TAG</th><th>Avanço</th>'
        f'<th>Etapas</th><th>{cabecalho_valor}</th></tr></thead>'
        f'<tbody>{corpo}</tbody></table></div></div>')

    # As fichas das TAGs da tabela. Só das mostradas: gerar as 4.939 custaria
    # dezenas de MB de HTML para modais que ninguém vai abrir.
    render_html(fichas_completas([l["tag"] for l in mostradas], resumo,
                                 esperados, tags, sigem, cache_key,
                                 origem="medicao"))

    exportar = pd.DataFrame([{
        "TAG": l["tag"], "DESCRICAO": l["descricao"],
        "ETAPAS_FEITAS": l["n_feito"], "ETAPAS_TOTAL": l["total"],
        "AVANCO_PCT": round(l["pct"], 1),
        "FECHADAS": " + ".join(sorted(l["fechadas"], key=MEDICAO_ORDEM.index)),
        "ABERTAS": " + ".join(sorted(l["abertas"], key=MEDICAO_ORDEM.index)),
        "MEDIDA_NO_GITEC": "sim" if l["medida"] else "nao",
        "VALOR_MEDIDO": l["valor_medido"],
        "PRECO_UNITARIO": l["preco"], "SOP": l["sop"], "SKID": l["skid"],
    } for l in sorted(linhas, key=lambda x: (x["medida"], x["falta"],
                                             -x["preco"], x["tag"]))])
    st.download_button(
        "Baixar planilha completa",
        exportar.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
        file_name="previsao_medicao.csv", mime="text/csv", key="medicao_csv")


def render_perfil_lateral():
    """O perfil no rodapé da lateral: foto, nome e o caminho para o resto.

    Sair saiu da lateral. Era um botão do tamanho do menu para uma ação que se
    usa uma vez por dia, competindo com as abas -- agora mora dentro do
    perfil, junto com foto e senha, que é onde se procura.
    """
    eu = st.session_state.get("gplan_usuario") or {}
    papel = acesso.papel_de(eu)
    dica = "" if pode("ver_valores") else " · este login não vê valores em reais"
    # A linha e o botao vivem dentro do MESMO st.container, que fica com
    # posicionamento proprio (position:relative) em vez de ser espalhado
    # pela reordenacao por flex da lateral. O botao cobre esse container
    # inteiro (position:absolute; inset:0), do tamanho que a linha ditar --
    # sem isso, um calculo de margem negativa fixa (que soma a puxada certa
    # so quando a linha e o vizinho imediato de verdade) passou a cobrir o
    # fim do menu e o botao do filtro assim que outra coisa entrou no meio.
    with st.container(key="perfil_wrap"):
        render_html(
            f'<div class="sb-perfil" title="{esc(papel)}{esc(dica)}">'
            + (f'<img class="sb-foto" src="{eu["foto"]}" alt="">' if eu.get("foto")
               else f'<span class="sb-ini">'
                    f'{esc(acesso.iniciais(eu.get("nome", ""), eu.get("email", "")))}</span>')
            + '<span class="sb-txt">'
            f'<b>{esc(eu.get("nome") or eu.get("email", ""))}</b>'
            f'<i>{esc(eu.get("email", ""))}</i>'
            f'<em class="sb-papel {acesso.COR_PAPEL.get(papel, "teal")}">{esc(papel)}</em>'
            "</span>"
            '<span class="sb-seta" aria-hidden="true">'
            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"'
            ' stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg>'
            "</span></div>")
        # O botão do Streamlit não aceita HTML no rótulo, então ele vem depois da
        # linha e é puxado por cima dela, transparente: o que se vê é a linha, o
        # que se clica é o botão.
        if st.button("Abrir o perfil", key="abrir_perfil", use_container_width=True):
            dialogo_perfil()


def campos_do_papel(prefixo: str, papel_inicial: str, marcadas: list[str]):
    """A caixa que define o usuário, e as permissões que vêm com ela.

    O papel é a escolha que a pessoa faz; a permissão é o detalhe que quase
    nunca se mexe. Por isso o papel vem primeiro e em destaque, e as
    permissões ficam recolhidas -- abrir só quem precisa fugir do padrão.
    """
    papel = st.selectbox(
        "Tipo de usuário", list(acesso.PAPEIS),
        index=list(acesso.PAPEIS).index(papel_inicial), key=f"{prefixo}_papel")
    render_html(
        '<div class="ac-resumo">'
        + "".join(f'<span>{esc(acesso.PERMISSOES[p])}</span>'
                  for p in acesso.PAPEIS[papel])
        + "</div>")
    trocou = papel != papel_inicial
    with st.expander("Ajustar permissões uma a uma"):
        if trocou:
            st.caption(f"Ao salvar, valem as permissões de {papel}. "
                       "Mexer aqui só faz efeito se você mantiver o tipo.")
        perms = st.multiselect(
            "Permissões", acesso.TODAS,
            default=list(acesso.PAPEIS[papel]) if trocou else marcadas,
            format_func=lambda p: acesso.PERMISSOES[p], key=f"{prefixo}_perms")
    return papel, (list(acesso.PAPEIS[papel]) if trocou else perms)


@st.dialog("Novo login", width="large")
def dialogo_novo_login():
    """Cadastrar é coisa de vez em quando: não pode ocupar a tela que serve
    para olhar quem já existe."""
    e1, e2 = st.columns(2)
    email = e1.text_input("E-mail", key="nv_email")
    nome = e2.text_input("Nome", key="nv_nome")
    senha = st.text_input("Senha", type="password", key="nv_senha",
                          help=acesso.REGRA_SENHA)
    st.caption(acesso.REGRA_SENHA)
    papel, perms = campos_do_papel("nv", acesso.PAPEL_PADRAO,
                                   list(acesso.PAPEIS[acesso.PAPEL_PADRAO]))
    if st.button("Criar login", type="primary", use_container_width=True,
                 key="nv_criar"):
        if not email.strip() or not senha:
            st.error("E-mail e senha são obrigatórios.")
        elif (falta := acesso.problema_na_senha(senha)):
            st.error(falta)
        else:
            try:
                acesso.criar(email, senha, nome, papel, perms)
                st.rerun()
            except Exception as erro:
                avisar_erro("criar o login", erro)


@st.dialog("Editar login", width="large")
def dialogo_editar_login(uid: str):
    eu = st.session_state.get("gplan_usuario") or {}
    try:
        usuarios = acesso.listar()
    except Exception as erro:
        avisar_erro("ler os logins", erro)
        return
    u = next((x for x in usuarios if x["id"] == uid), None)
    if not u:
        st.error("Esse login não existe mais.")
        return
    sou_eu = u["id"] == eu.get("id")
    papel_atual = acesso.papel_de(u)

    render_html(
        '<div class="pf-topo">'
        + (f'<img class="pf-foto" src="{u["foto"]}" alt="">' if u["foto"]
           else f'<span class="pf-ini">'
                f'{esc(acesso.iniciais(u["nome"], u["email"]))}</span>')
        + f'<div><div class="pf-nome">{esc(u["nome"] or u["email"])}</div>'
        f'<div class="pf-login">{esc(u["email"])}</div></div></div>')

    papel, perms = campos_do_papel(f"ed{uid}", papel_atual, u["permissoes"])
    ativo = st.checkbox("Login ativo", value=u["ativo"], key=f"ed{uid}_ativo",
                        help="Desativado não entra, e a sessão aberta dele cai.")
    nova = st.text_input("Trocar a senha (deixe vazio para manter)",
                         type="password", key=f"ed{uid}_senha",
                         help=acesso.REGRA_SENHA)

    b1, b2 = st.columns([3, 1])
    if b1.button("Salvar", type="primary", use_container_width=True,
                 key=f"ed{uid}_salvar"):
        # o administrador não pode se trancar do lado de fora: sem ninguém com
        # "administrar" não há como voltar a esta tela
        outros = any("administrar" in o["permissoes"] and o["ativo"]
                     and o["id"] != uid for o in usuarios)
        if sou_eu and not outros and ("administrar" not in perms or not ativo):
            st.error("Você é o único administrador ativo. Crie outro antes de "
                     "tirar a sua permissão ou se desativar.")
        elif nova and (falta := acesso.problema_na_senha(nova)):
            st.error(falta)
        else:
            try:
                acesso.salvar_perfil(uid, papel=papel, permissoes=perms,
                                     ativo=ativo)
                if nova:
                    acesso.trocar_senha(uid, nova)
                if sou_eu:
                    recarregar_usuario()
                st.rerun()
            except Exception as erro:
                avisar_erro("salvar o login", erro)

    if sou_eu:
        b2.button("Remover", disabled=True, use_container_width=True,
                  key=f"ed{uid}_rm", help="Não dá para remover o próprio login.")
    elif b2.button("Remover", use_container_width=True, key=f"ed{uid}_rm"):
        try:
            acesso.remover(uid)
            st.rerun()
        except Exception as erro:
            avisar_erro("remover o login", erro)


def render_acessos():
    """Quem entra e o que cada um vê. Só o administrador chega aqui.

    A página serve para OLHAR: o panorama, o filtro e os cartões de todos.
    Cadastrar e editar viram janela -- eram dois formulários compridos que
    empurravam os cartões para fora da tela, e quem abria a aba para conferir
    um acesso tinha de rolar por eles antes de ver qualquer coisa.

    Nenhuma senha aparece aqui: quem guarda é o Auth do Supabase, e de lá não
    se lê senha de volta.
    """
    render_header("Acessos")
    eu = st.session_state.get("gplan_usuario") or {}

    try:
        usuarios = acesso.listar()
    except Exception as erro:
        avisar_erro("listar os logins", erro)
        return

    por_papel = collections.Counter(acesso.papel_de(u) for u in usuarios)
    inativos = sum(1 for u in usuarios if not u["ativo"])
    render_html(
        '<div class="ac-topo">'
        f'<div class="ac-kpi"><span class="n">{len(usuarios)}</span>'
        f'<span class="r">logins</span></div>'
        + "".join(
            f'<div class="ac-kpi {acesso.COR_PAPEL.get(p, "teal")}">'
            f'<span class="n">{por_papel.get(p, 0)}</span>'
            f'<span class="r">{esc(p.lower())}</span></div>'
            for p in acesso.PAPEIS)
        + (f'<div class="ac-kpi vermelho"><span class="n">{inativos}</span>'
           f'<span class="r">inativos</span></div>' if inativos else "")
        + "</div>")

    c1, c2, c3 = st.columns([3, 2, 1.4])
    busca = c1.text_input("Procurar por nome ou e-mail",
                          key="ac_busca").strip().lower()
    filtro = c2.selectbox("Tipo de usuário", ["Todos"] + list(acesso.PAPEIS),
                          key="ac_papel")
    c3.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if c3.button("Novo login", type="primary", use_container_width=True,
                 key="ac_novo"):
        dialogo_novo_login()

    def cabe(u):
        if filtro != "Todos" and acesso.papel_de(u) != filtro:
            return False
        return not busca or busca in (u["nome"] + " " + u["email"]).lower()

    visiveis = [u for u in usuarios if cabe(u)]
    if not visiveis:
        render_html('<div class="gtbl-empty">Nenhum login com esse recorte.</div>')
        return

    # Um cartão por login, e ao lado de cada um o botão que abre a janela de
    # edição. O botão é do Streamlit porque precisa disparar o diálogo; o
    # cartão é HTML porque precisa da foto e dos selos.
    for u in visiveis:
        papel = acesso.papel_de(u)
        perms = [acesso.PERMISSOES[p] for p in u["permissoes"]
                 if p in acesso.PERMISSOES]
        col_card, col_bt = st.columns([9, 1.3])
        with col_card:
            render_html(
                f'<div class="ac-card{"" if u["ativo"] else " off"}">'
                '<div class="ac-cab">'
                + (f'<img class="ac-foto" src="{u["foto"]}" alt="">' if u["foto"]
                   else f'<span class="ac-ini">'
                        f'{esc(acesso.iniciais(u["nome"], u["email"]))}</span>')
                + f'<div class="ac-id"><b>{esc(u["nome"] or u["email"])}'
                + ('<em class="ac-eu">você</em>' if u["id"] == eu.get("id") else "")
                + f'</b><i>{esc(u["email"])}</i></div>'
                + ('' if u["ativo"] else '<span class="ac-off">desativado</span>')
                + f'<span class="ac-papel {acesso.COR_PAPEL.get(papel, "teal")}">'
                f'{esc(papel)}</span></div>'
                '<div class="ac-perms">'
                + "".join(f'<span>{esc(x)}</span>' for x in perms)
                + ('<span class="vazio">sem permissões</span>' if not perms else "")
                + "</div></div>")
        with col_bt:
            st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
            if st.button("Editar", key=f"ed_{u['id']}", use_container_width=True):
                dialogo_editar_login(u["id"])


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
    lembrar_tema(tema_ativo())
    # antes de qualquer dado: sem login não se carrega planilha nem se desenha
    exigir_login()
    # O cookie é reescrito a CADA desenho, e não só no momento de entrar.
    # Escrever só no login não funcionava: o st.rerun() logo em seguida aborta
    # a execução antes de o componente que roda o script chegar a existir, e o
    # cookie nunca era gravado -- o F5 seguinte caía na tela de login. É o
    # mesmo caminho que o tema já usava, pelo mesmo motivo. De quebra, o
    # refresh_token do Supabase gira a cada retomada, e reescrever sempre é o
    # que mantém no navegador o token que ainda vale.
    lembrar_sessao(st.session_state.get("gplan_token", ""))

    # O carimbo da planilha e a chave de cache sao coisas diferentes: a chave
    # leva a versao da regra colada no fim, e quem le a data precisa do valor
    # cru. Passar a chave para data_atualizacao deixava o cabecalho em "—".
    fonte = get_source_cache_key()
    # A permissão entra na chave porque metade das telas é montada dentro de
    # função @st.cache_data, e o HTML delas já traz o valor escrito. Sem isto,
    # o primeiro a abrir grava a versão dele no cache e o próximo recebe a
    # mesma -- quem não pode ver valores veria os do anterior, e vice-versa.
    cache_key = (f"{fonte}|r{REGRA_VERSAO}|v{VISUAL_VERSAO}"
                 f"|{'com' if pode('ver_valores') else 'sem'}-valores")
    if cache_key == "missing":
        st.error(
            "Não encontrei a planilha no Supabase Storage nem localmente. "
            "Verifique se o arquivo foi enviado ao bucket 'gplan-data'."
        )
        st.stop()

    st.session_state["gplan_atualizado_em"] = data_atualizacao(fonte)
    (tags, cabos, tubing, sigem, resumo, esperados,
     gitec, locacao, aux_areas, lancamento, depara, movimentacoes) = load_data(cache_key)

    with st.sidebar:
        render_html(
            '<div class="gplan-brand">'
            + _logo_svg("Lat").replace("<svg ", '<svg class="gplan-brand-mark" ')
            + '<div class="gplan-brand-text">'
            '<div class="gplan-brand-name">Gplan</div>'
            '<div class="gplan-brand-sub">Instrumentação · U-12</div>'
            "</div></div>"
        )

    # O filtro entra aqui, entre a marca e o perfil: a marca e escrita antes
    # do menu de navegacao (que o Streamlit sempre desenha por cima de
    # qualquer coisa que a lateral ja tenha), e tudo escrito depois cai
    # abaixo dele -- entao chamar o filtro antes do perfil e o que garante a
    # ordem visual marca -> menu (com Acessos) -> filtro -> perfil.
    escolhas = sidebar_filtros(tags, resumo, esperados)
    tags, resumo, esperados = aplicar_filtros(escolhas, tags, resumo, esperados)
    # a medicao segue as tags: filtrou a fase, a aba Gitec mostra so o que foi
    # medido nela
    gitec_f = gitec[gitec["TAG"].isin(set(tags["TAG"]))] if not gitec.empty else gitec

    with st.sidebar:
        render_perfil_lateral()

    dashboard_page = st.Page(lambda: _sob_carga("Carregando o painel", lambda: render_dashboard(resumo, esperados, tags, sigem, cache_key)), title="Dashboard", icon=":material/dashboard:", url_path="dashboard", default=True)
    relatorios_page = st.Page(lambda: _sob_carga("Carregando os relatórios", lambda: render_relatorios(esperados, resumo, tags, sigem, cache_key)), title="Relatórios", icon=":material/description:", url_path="relatorios")
    progresso_page = st.Page(lambda: _sob_carga("Abrindo o Progresso", lambda: render_progresso(resumo, esperados, tags, sigem, cache_key)), title="Progresso", icon=":material/insights:", url_path="progresso")
    pesquisa_page = st.Page(lambda: _sob_carga("Carregando as tags", lambda: render_pesquisa_tag(resumo, esperados, tags, sigem, cache_key, lancamento, depara)), title="Pesquisa tag", icon=":material/search:", url_path="pesquisa")
    sigem_page = st.Page(lambda: _sob_carga("Carregando a base SIGEM", lambda: render_sigem(sigem, esperados, any(escolhas.values()))), title="Base SIGEM", icon=":material/database:", url_path="sigem")
    gitec_page = st.Page(lambda: _sob_carga("Carregando a medição de campo", lambda: render_gitec(gitec_f, resumo, tags, esperados, sigem, cache_key)), title="Gitec", icon=":material/engineering:", url_path="gitec")
    planta_page = st.Page(lambda: _sob_carga("Desenhando o avanço na planta", lambda: render_planta(tags, resumo, locacao, aux_areas, esperados, sigem, cache_key)), title="Planta", icon=":material/map:", url_path="planta")
    certificacao_page = st.Page(lambda: _sob_carga("Montando a cadeia de certificação", lambda: render_certificacao(tags, lancamento, depara, resumo, esperados, sigem, cache_key)), title="Certificação", icon=":material/fact_check:", url_path="certificacao")
    atualizacao_page = st.Page(lambda: _sob_carga("Lendo as movimentações", lambda: render_atualizacao(movimentacoes, sigem, esperados, lancamento, tags, resumo, cache_key)), title="Última atualização", icon=":material/history:", url_path="atualizacao")

    admin_page = st.Page(render_acessos, title="Acessos",
                         icon=":material/manage_accounts:", url_path="acessos")
    medicao_page = st.Page(
        lambda: _sob_carga("Conferindo o que o campo já fechou",
                           lambda: render_previsao_medicao(
                               tags, resumo, esperados, lancamento, depara,
                               sigem, cache_key)),
        title="Previsão Medição", icon=":material/paid:",
        url_path="previsao-medicao")

    # A aba que a permissão não cobre não entra no menu, em vez de entrar e
    # avisar que é proibida: a Gitec sem valores viraria uma tela de traços, e
    # anunciar o que existe e não se pode ver é o que a reunião não precisa.
    #
    # Agrupadas por intenção -- a pergunta que cada uma responde -- em vez de
    # uma lista só de nove abas: Visão geral (retrato de agora), Documentação
    # (o que falta aprovar/emitir) e Avanço (o que já foi feito na obra). O
    # st.navigation faz isso nativamente com um dict {seção: [páginas]}, sem
    # precisar de CSS por cima do menu do próprio Streamlit.
    # ":material/nome:" quebra a barra INTEIRA quando usado no titulo de uma
    # secao (confirmado ao vivo: com ele, a secao inteira
    # <section data-testid="stSidebar"> some do DOM). O icone de cada secao
    # sai por CSS (nth-of-type, ver .stSidebarNav header abaixo) -- o titulo
    # aqui fica so texto puro.
    secoes: dict[str, list] = {
        "Visão geral": [dashboard_page, progresso_page, pesquisa_page],
        "Documentação": [relatorios_page, sigem_page, atualizacao_page],
    }
    avanco = []
    if pode("ver_gitec"):
        avanco.append(gitec_page)
    if pode("ver_planta"):
        avanco.append(planta_page)
    if pode("ver_certificacao"):
        avanco.append(certificacao_page)
    if avanco:
        secoes["Avanço"] = avanco
    if pode("administrar"):
        secoes["Administração"] = [medicao_page, admin_page]

    nav = st.navigation(secoes, position="sidebar")
    nav.run()


if __name__ == "__main__":
    main()
