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
# Três modos, e quem decide é quem programa: BLOQUEIA tira a TAG da lista,
# AVISA deixa passar em âmbar, IGNORA nem mostra.
#
# O padrão saiu da prova real -- as 719 TAGs JÁ MONTADAS, em 13/09/2026:
# calibração, almoxarifado, suprimentos e localização reprovariam de 0% a 2%
# delas; bandeja reprovaria 19% e pedestal 13%. Critério que reprova um
# quinto do que já foi para campo avisa, não bloqueia (usuário: "se o
# pedestal tiver faltando abertura de tag... não é impeditivo de montagem").
BLOQUEIA, AVISA, IGNORA = "bloqueia", "avisa", "ignora"
MODOS = (BLOQUEIA, AVISA, IGNORA)

# (chave, rótulo, o que significa, modo padrão)
CRITERIOS = [
    ("calibracao", "Calibração aprovada",
     "STATUS_CALIBRACAO da 01_BASE_TAGS", BLOQUEIA),
    ("estoque", "Material na obra",
     "almoxarifado: reservado para a TAG, em estoque ou já retirado", BLOQUEIA),
    ("suprimentos", "Sem material pendente",
     "Mapa de Suprimentos, a mesma ligação da aba Suprimentos", BLOQUEIA),
    ("localizacao", "Instrumento localizado",
     "STATUS_LOCALIZACAO -- ele usa em 100% do que programa", BLOQUEIA),
    ("suporte", "Suporte, bandeja e eletroduto",
     "05_BASE_LOCAÇÃO -- pendência de trecho não impede montar", AVISA),
    ("pedestal", "Pedestal concluído",
     "08_BASE_PEDESTAL -- falta de abertura não impede montar", AVISA),
    ("infra", "Infraestrutura da planta",
     "11_BASE_INFRAESTRUTURA, avanço por desenho", AVISA),
]
MODOS_PADRAO = {c: modo for c, _r, _f, modo in CRITERIOS}
LIGADOS_PADRAO = [c for c, _r, _f, modo in CRITERIOS if modo != IGNORA]

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
    """O material chegou à obra?

    Reservado é o sinal mais forte: o almoxarifado separou aquele material
    para a TAG, e ao separar ele SAI do saldo. Foi o que a programação das
    semanas 63 e 64 mostrou -- 87 TAGs reservadas contra 3 com saldo. Já
    retirado (emitido) também conta: o material foi para o campo.
    """
    saldo = f._estoque_por_tag.get(tag)
    if saldo is None:
        return TRAVA, "material não está na obra"
    if saldo["reservada"] > 0:
        return OK, f"reservado {saldo['reservada']:g}" + (
            f" · {saldo['onde']}" if saldo["onde"] else "")
    if saldo["estoque"] > 0:
        return OK, f"em estoque {saldo['estoque']:g}" + (
            f" · {saldo['onde']}" if saldo["onde"] else "")
    if saldo["emitida"] > 0:
        return OK, f"já retirado {saldo['emitida']:g}"
    if saldo["recebida"] > 0:
        return RESSALVA, "recebido, mas sem saldo nem reserva agora"
    return TRAVA, "material não está na obra"


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
        # ressalva, e não trava: das 719 TAGs já montadas, 134 estão sem
        # bandeja na base -- ela vem depois, ou em paralelo com o instrumento
        return RESSALVA, f"falta {' e '.join(faltam)}"
    tem = [nome.lower() for nome in ("SUPORTE", "BANDEJA", "ELETRODUTO")
           if _texto(loc.get(nome)).upper() == "SIM"]
    if not tem:
        return SEM_DADO, "suporte e bandeja sem informação"
    return OK, " + ".join(tem)


def _c_pedestal(tag: str, linha: pd.Series, f: Fonte) -> tuple[str, str]:
    """Pedestal incompleto é aviso, nunca trava por conta própria: a base
    marca o pedestal inteiro, e o que falta pode ser a abertura de outra TAG
    (usuário, 13/09/2026). Quem quiser que bloqueie escolhe na tela."""
    avanco = f.pedestal.get(tag)
    if avanco is None:
        return SEM_DADO, "TAG sem pedestal na base"
    if avanco >= 1:
        return OK, "pedestal 100%"
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
    def numero(coluna):
        return pd.to_numeric(f.estoque.get(coluna), errors="coerce").fillna(0)

    onde = f.estoque.get("Localização", pd.Series("", index=f.estoque.index)).astype(str)
    por_tag_estoque: dict[str, dict] = {}
    for tag, rec, est, res, emi, local in zip(
            f.estoque.get("Tag Number", pd.Series(dtype=str)).astype(str).str.strip(),
            numero("Quantidade Recebida"), numero("Quantidade Estoque"),
            numero("Quantidade Reservada"), numero("Quantidade Emitida"), onde):
        if not tag or tag == "nan":
            continue
        antes = por_tag_estoque.setdefault(
            tag, {"recebida": 0.0, "estoque": 0.0, "reservada": 0.0, "emitida": 0.0, "onde": ""})
        antes["recebida"] += float(rec)
        antes["estoque"] += float(est)
        antes["reservada"] += float(res)
        antes["emitida"] += float(emi)
        if not antes["onde"] and (res > 0 or est > 0):
            antes["onde"] = local
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
    # os dois lados: o status e a semana. Uma TAG com status "Em Programação"
    # sem semana marcada não pode reaparecer como disponível.
    livre = (~mont.isin(("Montado", "Em Programação"))
             & semana.isin(("Não Programado", "", "nan")))
    return tags[livre].copy()


def programadas(tags: pd.DataFrame, semana: str = "") -> pd.DataFrame:
    """Quem já tem semana marcada na base -- de uma semana, ou de todas."""
    coluna = tags["SEMANA_PROGRAMADA"].astype(str).str.strip()
    if semana:
        return tags[coluna.eq(semana)].copy()
    return tags[~coluna.isin(("Não Programado", "", "nan"))].copy()


def estado(tags: pd.DataFrame) -> dict:
    """O retrato da montagem: o que está montado, programado e disponível."""
    mont = tags["STATUS_MONTAGEM"].astype(str).str.strip()
    semana = tags["SEMANA_PROGRAMADA"].astype(str).str.strip()
    com_semana = ~semana.isin(("Não Programado", "", "nan"))
    por_semana = semana[com_semana].value_counts()
    por_semana = {s: int(n) for s, n in sorted(
        por_semana.items(), key=lambda x: _numero_da_semana(x[0]))}
    return {
        "total": len(tags),
        "montadas": int(mont.eq("Montado").sum()),
        "em_programacao": int(mont.eq("Em Programação").sum()),
        "nao_montado": int(mont.eq("Não Montado").sum()),
        "programadas": int(com_semana.sum()),
        "disponiveis": len(candidatas(tags)),
        "por_semana": por_semana,
    }


def _semana_teste(valor: object) -> str:
    """A semana prevista do teste de malha ("SEM87"). "0" quer dizer sem
    previsão, mesma convenção do SKID -- e vira vazio."""
    texto = _texto(valor).upper().replace(" ", "")
    return "" if texto in ("", "0", "SEM0") else texto


def urgencia(linha: dict) -> tuple:
    """A ordem em que as oportunidades aparecem.

    Primeiro o teste de malha mais próximo (a montagem tem de acontecer
    antes dele), depois prioritária, depois o nível do subgrupo, e por fim
    a TAG -- para a lista não dançar entre um desenho e outro.
    """
    teste = linha.get("TESTE_MALHA") or ""
    numero = int("".join(c for c in teste if c.isdigit()) or 9999)
    nivel = linha.get("NIVEL") or "99"
    try:
        nivel_num = float(nivel)
    except ValueError:
        nivel_num = 99.0
    return (numero, 0 if linha.get("PRIORITARIA") else 1, nivel_num, linha.get("TAG", ""))


def _numero_da_semana(rotulo: object) -> int:
    digitos = "".join(c for c in str(rotulo) if c.isdigit())
    return int(digitos) if digitos else 0


def _modos(escolha) -> dict:
    """Aceita o dicionário de modos, uma lista de critérios ligados (que
    então bloqueiam) ou None, que vale o padrão medido."""
    if escolha is None:
        return dict(MODOS_PADRAO)
    if isinstance(escolha, dict):
        return {c: escolha.get(c, IGNORA) for c, _r, _f, _p in CRITERIOS}
    ligados = set(escolha)
    return {c: (MODOS_PADRAO[c] if c in ligados else IGNORA)
            for c, _r, _f, _p in CRITERIOS}


def avaliar(f: Fonte, ligados: list[str] | None = None) -> pd.DataFrame:
    """Uma linha por TAG candidata, com o veredito e o porquê.

    Colunas: TAG, DESCRICAO, FAMILIA, AREA, PLANTA, PRIORITARIA, SITUACAO
    ("livre", "ressalva" ou "travada"), MOTIVOS (lista de (chave, situação,
    texto)) e BLOQUEIOS (só o que pesou).
    """
    return _avaliar(f, candidatas(f.tags), ligados)


def conferir(f: Fonte, semana: str, ligados: list[str] | None = None) -> pd.DataFrame:
    """O que JÁ está programado para a semana continua apto?

    A programação envelhece: material que não chegou, calibração que
    reprovou. Passar a mesma régua no que já foi programado é o que avisa a
    tempo de trocar -- em 13/09/2026, 28 das 75 TAGs da Semana 64 estavam
    sem material registrado na obra.
    """
    return _avaliar(f, programadas(f.tags, semana), ligados)


def _avaliar(f: Fonte, base: pd.DataFrame, modos) -> pd.DataFrame:
    f = preparar(f)
    modos = _modos(modos)
    linhas = []
    for linha in base.to_dict("records"):
        tag = str(linha.get("TAG") or "").strip()
        if not tag:
            continue
        motivos, bloqueios, ressalvas = [], [], []
        for chave, _rotulo, _fonte, _padrao in CRITERIOS:
            situacao, texto = CONTAS[chave](tag, linha, f)
            motivos.append((chave, situacao, texto))
            modo = modos.get(chave, IGNORA)
            if modo == IGNORA or situacao in (OK, SEM_DADO):
                continue
            if modo == BLOQUEIA and situacao == TRAVA:
                bloqueios.append(texto)
            else:
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
            "SEMANA": _texto(linha.get("SEMANA_PROGRAMADA")),
            "MONTAGEM": _texto(linha.get("STATUS_MONTAGEM")),
            "NIVEL": _texto(linha.get("SUBGRUPO_PRIORIDADE")),
            "FASE": _texto(linha.get("FASE")),
            "SISTEMA": _texto(linha.get("SOP")),
            "SUBSISTEMA": _texto(linha.get("SSOP")),
            "TESTE_MALHA": _semana_teste(linha.get("PREVISAO_TESTE_MALHA")),
            "SITUACAO": "travada" if bloqueios else ("ressalva" if ressalvas else "livre"),
            "MOTIVOS": motivos,
            "BLOQUEIOS": bloqueios,
            "RESSALVAS": ressalvas,
        })
    if not linhas:
        return pd.DataFrame(columns=["TAG", "DESCRICAO", "FAMILIA", "AREA", "PLANTA",
                                     "MALHA", "PRIORITARIA", "NIVEL", "FASE", "SISTEMA",
                                     "SUBSISTEMA", "TESTE_MALHA", "SEMANA", "MONTAGEM",
                                     "SITUACAO", "MOTIVOS", "BLOQUEIOS", "RESSALVAS"])
    return pd.DataFrame(sorted(linhas, key=urgencia))


def resumo(aval: pd.DataFrame) -> dict:
    if aval.empty:
        return {"livre": 0, "ressalva": 0, "travada": 0, "total": 0}
    contagem = aval["SITUACAO"].value_counts().to_dict()
    return {"livre": contagem.get("livre", 0), "ressalva": contagem.get("ressalva", 0),
            "travada": contagem.get("travada", 0), "total": len(aval)}


def por_grupo(aval: pd.DataFrame, coluna: str = "PLANTA") -> list[dict]:
    """As oportunidades agrupadas -- por planta é como o campo trabalha, e
    por subsistema é como ele programa (130 das 203 TAGs das semanas 62 a 64
    saíram do mesmo SOP)."""
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


# --------------------------------------------------------- o que urge
JANELA_TESTE = 2        # semanas: teste nesse prazo e TAG sem semana = urgente


def semana_numero(rotulo: object) -> int:
    """"SEM87" ou "Semana 87" -> 87. Sem número, 0."""
    digitos = "".join(c for c in str(rotulo or "") if c.isdigit())
    return int(digitos) if digitos else 0


def teste_urgente(aval: pd.DataFrame, semana_atual: int,
                  janela: int = JANELA_TESTE) -> pd.DataFrame:
    """Aptas com teste de malha dentro da janela e ainda sem semana.

    O teste de malha é o prazo real: montar depois dele atrasa a malha.
    """
    if aval.empty:
        return aval
    limite = semana_atual + janela
    numeros = aval["TESTE_MALHA"].map(semana_numero)
    alvo = aval[(numeros > 0) & (numeros <= limite) & (aval["SITUACAO"] != "travada")]
    return alvo.assign(_n=numeros[alvo.index]).sort_values("_n").drop(columns="_n")


def sugerir(aval: pd.DataFrame, quantidade: int, fora: list | None = None) -> list:
    """As N mais urgentes ainda não escolhidas -- a ordem já é a da lista."""
    if quantidade <= 0 or aval.empty:
        return []
    fora = set(fora or [])
    livres = aval[(aval["SITUACAO"] != "travada") & (~aval["TAG"].isin(fora))]
    return livres["TAG"].head(quantidade).tolist()


def niveis_pulados(aval: pd.DataFrame, escolhidas: list) -> list:
    """Níveis mais urgentes que sobraram, abaixo do menor nível escolhido.

    Programar 6.01 com 3.13 ainda apto é fora de ordem -- o nível é a fila
    de prioridade da obra.
    """
    if not escolhidas or aval.empty:
        return []

    def numero(n):
        try:
            return float(n)
        except (TypeError, ValueError):
            return 99.0

    escolha = aval[aval["TAG"].isin(escolhidas)]
    niveis_escolhidos = [numero(n) for n in escolha["NIVEL"] if n]
    if not niveis_escolhidos:
        return []
    menor = min(niveis_escolhidos)
    sobrou = aval[(aval["SITUACAO"] != "travada") & (~aval["TAG"].isin(escolhidas))]
    pendentes = {}
    for nivel in sobrou["NIVEL"]:
        if nivel and numero(nivel) < menor:
            pendentes[nivel] = pendentes.get(nivel, 0) + 1
    return sorted(pendentes.items(), key=lambda x: numero(x[0]))[:3]


def historico(tags: pd.DataFrame, ultimas: int = 8) -> list:
    """Semana a semana: quanto foi programado e quanto foi montado."""
    semana = tags["SEMANA_PROGRAMADA"].astype(str).str.strip()
    mont = tags["STATUS_MONTAGEM"].astype(str).str.strip()
    com_semana = ~semana.isin(("Não Programado", "", "nan"))
    linhas = []
    for rotulo in sorted(set(semana[com_semana]), key=semana_numero):
        grupo = com_semana & semana.eq(rotulo)
        linhas.append({"semana": rotulo, "numero": semana_numero(rotulo),
                       "programadas": int(grupo.sum()),
                       "montadas": int((grupo & mont.eq("Montado")).sum())})
    return linhas[-ultimas:]


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
