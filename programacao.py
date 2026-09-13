"""Programação da semana: quem pode ser montado, e o que trava cada um.

A conta que o Daniel fazia à mão toda semana (2026-09-12): pegar a meta da
curva, filtrar as TAGs aprovadas na calibração, conferir material no
almoxarifado, procurar impedimento de suprimentos e olhar se o campo está
pronto -- suporte, bandeja, eletroduto, pedestal e a infraestrutura da
planta. Aqui isso vira um veredito por TAG, com o motivo escrito.

Cada critério LIGA E DESLIGA (escolha dele, 2026-09-13): o que está
desligado continua aparecendo na ficha da TAG, mas não bloqueia ninguém.

Sem Streamlit e sem Supabase de propósito: este módulo recebe os dados
prontos e devolve dados. Quem desenha e quem guarda é o gplan_app.
"""
from __future__ import annotations

import io
import json
from dataclasses import dataclass, field

import pandas as pd

# --------------------------------------------------------------- critérios
# (chave, rótulo, o que significa, ligado por padrão)
CRITERIOS = [
    ("calibracao", "Calibração aprovada",
     "STATUS_CALIBRACAO da 01_BASE_TAGS", True),
    ("estoque", "Material na obra",
     "saldo no almoxarifado, pelo Tag Number", True),
    ("suprimentos", "Sem material pendente",
     "Mapa de Suprimentos, a mesma ligação da aba Suprimentos", True),
    ("suporte", "Suporte, bandeja e eletroduto",
     "05_BASE_LOCAÇÃO", True),
    ("pedestal", "Pedestal concluído",
     "08_BASE_PEDESTAL, avanço por TAG", True),
    ("infra", "Infraestrutura da planta",
     "11_BASE_INFRAESTRUTURA, avanço por desenho", False),
    ("localizacao", "Instrumento localizado",
     "STATUS_LOCALIZACAO", False),
]
LIGADOS_PADRAO = [c for c, _, _, padrao in CRITERIOS if padrao]

# Abaixo disso é ressalva (âmbar), não bloqueio: o campo está quase pronto e
# costuma fechar dentro da semana.
PEDESTAL_RESSALVA = 0.80
INFRA_RESSALVA = 0.80

# ---------------------------------------------------------------- veredito
OK, RESSALVA, TRAVA, SEM_DADO = "ok", "ressalva", "trava", "sem_dado"

# Pela letra do instrumento (convenção ISA). Serve só para agrupar a tela:
# o que não estiver aqui aparece como "Outros", sem prejuízo nenhum.
FAMILIAS = {
    "CJ": "Caixa de junção", "CJA": "Caixa de junção", "CJD": "Caixa de junção",
    "CJP": "Caixa de junção", "CJS": "Caixa de junção", "CFF": "Caixa de junção",
    "PIT": "Transmissor", "FIT": "Transmissor", "TIT": "Transmissor",
    "LIT": "Transmissor", "PDIT": "Transmissor", "AT": "Transmissor",
    "VT": "Transmissor de vibração", "AIT": "Analisador",
    "AST": "Chave", "LSH": "Chave", "LSL": "Chave", "PSH": "Chave",
    "PSL": "Chave", "XSH": "Chave", "TSH": "Chave", "FSL": "Chave",
    "HS": "Chave manual", "ZSH": "Chave de posição", "ZSL": "Chave de posição",
    "PI": "Indicador", "TI": "Indicador", "LI": "Indicador", "FI": "Indicador",
    "PDI": "Indicador", "LG": "Visor de nível", "FG": "Visor de fluxo",
    "TE": "Termoelemento", "FE": "Elemento de vazão", "TW": "Poço termométrico",
    "VE": "Sensor de vibração", "XV": "Válvula on-off", "FV": "Válvula de controle",
    "PV": "Válvula de controle", "TV": "Válvula de controle", "LV": "Válvula de controle",
    "PSV": "Válvula de segurança", "XY": "Solenoide", "PY": "Conversor",
    "FY": "Conversor", "TY": "Conversor", "LY": "Conversor",
}


@dataclass
class Fonte:
    """Tudo o que a programação precisa, já lido pelo gplan_app."""
    tags: pd.DataFrame
    locacao: pd.DataFrame
    estoque: pd.DataFrame
    itens: pd.DataFrame
    por_tag: dict = field(default_factory=dict)      # sup_por_tag
    indice_titulo: dict = field(default_factory=dict)  # sup_indice_titulo
    pedestal: dict = field(default_factory=dict)     # TAG -> avanço 0..1
    infra: dict = field(default_factory=dict)        # sigla da planta -> 0..1


def familia(tag: str) -> str:
    prefixo = "".join(c for c in str(tag).split("-")[0] if c.isalpha()).upper()
    return FAMILIAS.get(prefixo, "Outros")


def planta_do_desenho(desenho: object) -> str:
    """A sigla da planta que mora no fim do endereço do desenho.

    "DE-5290.00-22111-800-CHZ-314_B" -> "CHZ-314". É a mesma leitura que a
    Previsão Medição faz para casar desenho com infraestrutura.
    """
    texto = str(desenho or "").strip().upper()
    if not texto:
        return ""
    pedacos = texto.replace("_", "-").split("-")
    for i in range(len(pedacos) - 1, 0, -1):
        if pedacos[i].isdigit() and pedacos[i - 1].isalpha() and len(pedacos[i - 1]) == 3:
            return f"{pedacos[i - 1]}-{pedacos[i]}"
    return ""


def _texto(v) -> str:
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "-") else s


# ------------------------------------------------------- um critério por vez
def _c_calibracao(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    valor = _texto(linha.get("STATUS_CALIBRACAO"))
    if not valor:
        # o "-" da coluna quer dizer "ainda não passou pela calibração", e
        # não "não sei": instrumento assim não vai para campo
        return TRAVA, "ainda não calibrada"
    if "APROV" in valor.upper():
        return OK, "calibração aprovada"
    return TRAVA, f"calibração: {valor.lower()}"


def _c_estoque(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    saldo = f._estoque_por_tag.get(tag)
    if saldo is None:
        return TRAVA, "material não está na obra"
    qtd, onde = saldo
    if qtd <= 0:
        return TRAVA, "sem saldo em estoque"
    return OK, f"estoque {qtd:g}" + (f" · {onde}" if onde else "")


def _c_suprimentos(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    """A mesma ligação da aba Suprimentos: material com a TAG na coluna, ou
    kit/acessório que cita a TAG no título (ver sup_por_tag e
    sup_indice_titulo, no gplan_app)."""
    estados = []
    info = f.por_tag.get(tag)
    if info:
        estados.append(info.get("geral") or "")
    for idx in f.indice_titulo.get(tag, []):
        try:
            chave = _texto(f.itens.at[idx, "TAG"]) or _texto(f.itens.at[idx, "IDENT_CODE"])
        except KeyError:
            continue
        rel = f.por_tag.get(chave)
        if rel:
            estados.append(rel.get("geral") or "")
    if not estados:
        return SEM_DADO, "sem material no Mapa de Suprimentos"
    ruins = [e for e in estados if e and e not in ("100% recebido", "Cancelado")]
    if ruins:
        return TRAVA, f"material {ruins[0].lower()}"
    return OK, "material recebido"


def _c_suporte(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    loc = f._locacao_por_tag.get(tag)
    if loc is None:
        return SEM_DADO, "sem locação cadastrada"
    faltam = [nome.lower() for nome in ("SUPORTE", "BANDEJA", "ELETRODUTO")
              if _texto(loc.get(nome)).upper() == "NÃO"]
    if faltam:
        return TRAVA, f"falta {' e '.join(faltam)}"
    tem = [nome.lower() for nome in ("SUPORTE", "BANDEJA", "ELETRODUTO")
           if _texto(loc.get(nome)).upper() == "SIM"]
    if not tem:
        return SEM_DADO, "suporte e bandeja sem informação"
    return OK, " + ".join(tem)


def _c_pedestal(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    avanco = f.pedestal.get(tag)
    if avanco is None:
        return SEM_DADO, "TAG sem pedestal na base"
    if avanco >= 1:
        return OK, "pedestal 100%"
    if avanco >= PEDESTAL_RESSALVA:
        return RESSALVA, f"pedestal {avanco * 100:.0f}%"
    return TRAVA, f"pedestal {avanco * 100:.0f}%"


def _c_infra(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    loc = f._locacao_por_tag.get(tag) or {}
    planta = planta_do_desenho(loc.get("LOCACAO"))
    if not planta or planta not in f.infra:
        return SEM_DADO, "planta sem avanço de infraestrutura"
    avanco = f.infra[planta]
    if avanco >= 1:
        return OK, f"infra {planta} 100%"
    if avanco >= INFRA_RESSALVA:
        return RESSALVA, f"infra {planta} {avanco * 100:.0f}%"
    return TRAVA, f"infra {planta} {avanco * 100:.0f}%"


def _c_localizacao(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    valor = _texto(linha.get("STATUS_LOCALIZACAO"))
    if not valor:
        return SEM_DADO, "localização sem informação"
    if "NÃO LOCALIZADO" in valor.upper():
        return TRAVA, "não localizada"
    return OK, valor.lower()


CONTAS = {
    "calibracao": _c_calibracao, "estoque": _c_estoque, "suprimentos": _c_suprimentos,
    "suporte": _c_suporte, "pedestal": _c_pedestal, "infra": _c_infra,
    "localizacao": _c_localizacao,
}


# -------------------------------------------------------------- o veredito
def preparar(f: Fonte) -> Fonte:
    """Índices que as contas usam -- feitos uma vez, não por TAG."""
    qtd = pd.to_numeric(f.estoque.get("Quantidade Estoque"), errors="coerce").fillna(0)
    onde = f.estoque.get("Localização", pd.Series("", index=f.estoque.index)).astype(str)
    por_tag_estoque: dict[str, tuple[float, str]] = {}
    for tag, q, local in zip(f.estoque.get("Tag Number", pd.Series(dtype=str)).astype(str).str.strip(),
                             qtd, onde):
        if not tag or tag == "nan":
            continue
        antes = por_tag_estoque.get(tag, (0.0, ""))
        por_tag_estoque[tag] = (antes[0] + float(q), local if q > 0 else antes[1])
    f._estoque_por_tag = por_tag_estoque
    f._locacao_por_tag = {str(r["TAG"]).strip(): r
                          for r in f.locacao.to_dict("records") if _texto(r.get("TAG"))}
    return f


def candidatas(tags: pd.DataFrame) -> pd.DataFrame:
    """Quem pode entrar numa programação: não montada e sem semana.

    Montada não se programa de novo, e quem já está em outra semana fica de
    fora para a mesma TAG não ser programada duas vezes.
    """
    mont = tags["STATUS_MONTAGEM"].astype(str).str.strip()
    semana = tags["SEMANA_PROGRAMADA"].astype(str).str.strip()
    livre = ~mont.eq("Montado") & semana.isin(("Não Programado", "", "nan"))
    return tags[livre].copy()


def avaliar(f: Fonte, ligados: list[str] | None = None) -> pd.DataFrame:
    """Uma linha por TAG candidata, com o veredito e o porquê.

    Colunas: TAG, DESCRICAO, FAMILIA, AREA, PLANTA, PRIORITARIA, SITUACAO
    ("livre", "ressalva" ou "travada"), MOTIVOS (lista de (chave, situação,
    texto)) e BLOQUEIOS (só o que pesou).
    """
    f = preparar(f)
    ligados = list(LIGADOS_PADRAO if ligados is None else ligados)
    linhas = []
    for linha in candidatas(f.tags).to_dict("records"):
        tag = str(linha.get("TAG") or "").strip()
        if not tag:
            continue
        motivos, bloqueios, ressalvas = [], [], []
        for chave, _rotulo, _fonte, _padrao in CRITERIOS:
            situacao, texto = CONTAS[chave](tag, linha, f)
            motivos.append((chave, situacao, texto))
            if chave not in ligados:
                continue
            if situacao == TRAVA:
                bloqueios.append(texto)
            elif situacao == RESSALVA:
                ressalvas.append(texto)
        loc = f._locacao_por_tag.get(tag) or {}
        linhas.append({
            "TAG": tag,
            "DESCRICAO": _texto(linha.get("DESCRICAO")),
            "FAMILIA": familia(tag),
            "AREA": _texto(loc.get("AREA")) or "sem área",
            "PLANTA": planta_do_desenho(loc.get("LOCACAO")) or "sem planta",
            "MALHA": _texto(linha.get("MALHA")),
            "PRIORITARIA": _texto(linha.get("SSOP_PRIORITARIO")).upper() == "SIM",
            "SITUACAO": "travada" if bloqueios else ("ressalva" if ressalvas else "livre"),
            "MOTIVOS": motivos,
            "BLOQUEIOS": bloqueios,
            "RESSALVAS": ressalvas,
        })
    return pd.DataFrame(linhas)


def resumo(aval: pd.DataFrame) -> dict:
    if aval.empty:
        return {"livre": 0, "ressalva": 0, "travada": 0, "total": 0}
    contagem = aval["SITUACAO"].value_counts().to_dict()
    return {"livre": contagem.get("livre", 0), "ressalva": contagem.get("ressalva", 0),
            "travada": contagem.get("travada", 0), "total": len(aval)}


def por_grupo(aval: pd.DataFrame, coluna: str = "PLANTA") -> list[dict]:
    """As oportunidades agrupadas -- por planta é como o campo trabalha."""
    if aval.empty:
        return []
    grupos = []
    for nome, bloco in aval.groupby(coluna):
        livres = bloco[bloco["SITUACAO"] != "travada"]
        grupos.append({
            "nome": str(nome), "total": len(bloco), "livres": len(livres),
            "prioritarias": int(bloco["PRIORITARIA"].sum()),
            "areas": sorted({a for a in bloco["AREA"] if a}),
            "tags": bloco.sort_values(["SITUACAO", "TAG"]),
        })
    return sorted(grupos, key=lambda g: (-g["livres"], g["nome"]))


# ------------------------------------------------------ guardar e exportar
def para_excel(aval: pd.DataFrame, escolhidas: list[str], semana: str) -> bytes:
    """A programação em .xlsx, pronta para lançar no sistema da empresa."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    escolha = aval[aval["TAG"].isin(escolhidas)].copy()
    escolha["ORDEM"] = escolha["TAG"].map({t: i for i, t in enumerate(escolhidas)})
    escolha = escolha.sort_values("ORDEM")
    wb = Workbook()
    ws = wb.active
    ws.title = semana.replace(" ", "_")[:28] or "Programacao"
    cabecalho = ["TAG", "DESCRIÇÃO", "TIPO", "ÁREA", "PLANTA", "MALHA",
                 "PRIORITÁRIA", "SEMANA", "SITUAÇÃO", "OBSERVAÇÃO"]
    ws.append(cabecalho)
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="0F766E")
        celula.alignment = Alignment(vertical="center")
    for r in escolha.to_dict("records"):
        ws.append([r["TAG"], r["DESCRICAO"], r["FAMILIA"], r["AREA"], r["PLANTA"],
                   r["MALHA"], "SIM" if r["PRIORITARIA"] else "NÃO", semana,
                   r["SITUACAO"], "; ".join(r["RESSALVAS"] + r["BLOQUEIOS"])])
    larguras = (16, 44, 18, 10, 12, 14, 12, 12, 12, 46)
    for i, largura in enumerate(larguras, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = largura
    ws.freeze_panes = "A2"
    memoria = io.BytesIO()
    wb.save(memoria)
    return memoria.getvalue()


def para_guardar(semana: str, escolhidas: list[str], aval: pd.DataFrame,
                 usuario: str, meta: float | None) -> bytes:
    """A programação como JSON, para o Gplan guardar e comparar depois."""
    escolha = aval[aval["TAG"].isin(escolhidas)]
    corpo = {
        "semana": semana,
        "salvo_em": pd.Timestamp.now().isoformat(timespec="seconds"),
        "por": usuario,
        "meta": meta,
        "quantidade": len(escolhidas),
        "tags": [
            {"tag": r["TAG"], "tipo": r["FAMILIA"], "area": r["AREA"],
             "planta": r["PLANTA"], "prioritaria": bool(r["PRIORITARIA"]),
             "situacao": r["SITUACAO"],
             "ressalvas": r["RESSALVAS"], "bloqueios": r["BLOQUEIOS"]}
            for r in escolha.to_dict("records")
        ],
    }
    return json.dumps(corpo, ensure_ascii=False, indent=1).encode("utf-8")


def ler_guardada(dados: bytes) -> dict:
    try:
        return json.loads(dados.decode("utf-8"))
    except (UnicodeDecodeError, ValueError):
        return {}
