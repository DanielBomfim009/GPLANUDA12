# -*- coding: utf-8 -*-
"""Extrai a marcação do PowerPoint para os assets da aba Planta.

O mapa que o Daniel mantém à mão -- "Mapa Visual - Avanço INFRA
Instrumentação.pptx" -- já traz o contorno de cada planta desenhado por cima do
arranjo. É dele que sai a geometria das zonas: nada é redesenhado no olho.

O app não lê o PPTX em tempo de execução; ele lê o que este script grava em
assets/mapa/. Rode-o de novo quando a marcação mudar de lugar:

    python gerar_mapa_assets.py "caminho/para/o.pptx"
"""
from __future__ import annotations

import base64
import io
import json
import pathlib
import re
import sys
import warnings
import xml.etree.ElementTree as ET
import zipfile

warnings.filterwarnings("ignore")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from PIL import Image

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

AQUI = pathlib.Path(__file__).parent
DESTINO = AQUI / "assets" / "mapa"
PADRAO = pathlib.Path.home() / "Downloads" / "Mapa Visual - Avanço INFRA Instrumentação.pptx"

# As três pranchas do slide, na ordem em que aparecem na aba. O nome do quadro
# é o que o PowerPoint deu à imagem; o resto é como ela se chama na tela.
PRANCHAS = {
    "Imagem 10": ("principal", "Arranjo geral da U-12"),
    "Imagem 35": ("piperack", "Pipe rack e estruturas"),
    "Imagem 8": ("se1200", "SE-1200 · subestação"),
}
# Largura em que a prancha é desenhada na aba: define até onde vale ampliar a
# imagem antes de virar peso morto no navegador.
LARGURA_ALVO = {"principal": 2000, "piperack": 2000, "se1200": 700}


# --------------------------------------------------------------- leitura do XML
def local(no, nome):
    """Primeiro descendente com esse nome local, seja qual for o namespace.

    O mesmo elemento troca de namespace conforme onde aparece: dentro do slide
    é p:spPr, dentro de um desenho é a:spPr.
    """
    for f in no.iter():
        if f.tag.split("}")[-1] == nome:
            return f
    return None


def xfrm_de(no):
    if no is None:
        return None
    x = local(no, "xfrm")
    if x is None:
        return None
    off, ext = x.find(f"{A}off"), x.find(f"{A}ext")
    if off is None or ext is None:
        return None
    d = {"x": int(off.get("x")), "y": int(off.get("y")),
         "cx": int(ext.get("cx")), "cy": int(ext.get("cy"))}
    ch_off, ch_ext = x.find(f"{A}chOff"), x.find(f"{A}chExt")
    if ch_off is not None and ch_ext is not None:
        d |= {"chx": int(ch_off.get("x")), "chy": int(ch_off.get("y")),
              "chcx": int(ch_ext.get("cx")), "chcy": int(ch_ext.get("cy"))}
    return d


def cor_de(no):
    pr = local(no, "spPr")
    if pr is None:
        return None
    for f in pr:
        n = f.tag.split("}")[-1]
        if n == "noFill":
            return "sem"
        if n == "solidFill":
            for tipo, prefixo in (("srgbClr", "#"), ("schemeClr", "tema:")):
                c = local(f, tipo)
                if c is not None:
                    return prefixo + c.get("val")
    return None


def texto_de(no):
    """O texto da forma, uma linha por parágrafo.

    "CHZ-326" e "100%" são dois parágrafos da mesma caixa. Emendados viram
    "CHZ-326100%", e o código do desenho deixa de ser reconhecível.
    """
    linhas = []
    for p in no.iter(f"{A}p"):
        t = "".join(r.text or "" for r in p.iter(f"{A}t")).strip()
        if t:
            linhas.append(t)
    return "\n".join(linhas)


def recorte_de(no, cx, cy):
    """A forma em "L" do PowerPoint virada em clip-path.

    A zona da CHZ-318 não é retângulo: foi desenhada como "corner" de propósito,
    para não engolir a CHZ-316/317, que fica no encaixe.
    """
    pr = local(no, "spPr")
    g = local(pr, "prstGeom") if pr is not None else None
    if g is None or g.get("prst") != "corner":
        return ""
    adj = {a.get("name"): int(a.get("fmla").split()[-1])
           for a in g.iter(f"{A}gd") if (a.get("fmla") or "").startswith("val")}
    ss = min(cx, cy)
    x1 = ss * adj.get("adj2", 50000) / 100000 / cx * 100
    y1 = 100 - ss * adj.get("adj1", 50000) / 100000 / cy * 100
    return (f"polygon(0 0, {x1:.2f}% 0, {x1:.2f}% {y1:.2f}%, "
            f"100% {y1:.2f}%, 100% 100%, 0 100%)")


def coletar(no, pai=None, saida=None):
    """Percorre a árvore do slide acumulando a transformação dos grupos."""
    saida = [] if saida is None else saida
    for f in no:
        tag = f.tag.split("}")[-1]
        if tag == "grpSp":
            g = xfrm_de(local(f, "grpSpPr"))
            novo = pai
            if g and "chcx" in g:
                sx, sy = g["cx"] / g["chcx"], g["cy"] / g["chcy"]
                ox, oy = g["x"] - g["chx"] * sx, g["y"] - g["chy"] * sy
                if pai:
                    sx, sy = sx * pai[0], sy * pai[1]
                    ox, oy = pai[2] + ox * pai[0], pai[3] + oy * pai[1]
                novo = (sx, sy, ox, oy)
            coletar(f, novo, saida)
        elif tag in ("sp", "pic"):
            x = xfrm_de(f)
            if not x:
                continue
            sx, sy, ox, oy = pai or (1, 1, 0, 0)
            nome = local(f, "cNvPr")
            cx, cy = x["cx"] * sx, x["cy"] * sy
            saida.append({
                "tipo": tag, "nome": nome.get("name") if nome is not None else "",
                "x": ox + x["x"] * sx, "y": oy + x["y"] * sy, "cx": cx, "cy": cy,
                "fill": cor_de(f), "txt": texto_de(f), "recorte": recorte_de(f, cx, cy),
                "rid": (local(f, "blip").get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships}embed") if local(f, "blip") is not None else None),
            })
    return saida


def centro(f):
    return f["x"] + f["cx"] / 2, f["y"] + f["cy"] / 2


# ------------------------------------------------------------------- execução
def main(origem: pathlib.Path) -> int:
    if not origem.exists():
        print(f"não achei o PPTX em {origem}")
        return 1

    with zipfile.ZipFile(origem) as z:
        slide = ET.fromstring(z.read("ppt/slides/slide1.xml"))
        rels = ET.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels"))
        arquivo_do_rid = {r.get("Id"): r.get("Target").split("/")[-1] for r in rels}
        midia = {n.split("/")[-1]: z.read(n) for n in z.namelist()
                 if n.startswith("ppt/media/")}

    formas = coletar(slide[0][0] if slide[0].tag.endswith("cSld") else slide)
    quadros = {f["nome"]: f for f in formas if f["nome"] in PRANCHAS}
    faltando = set(PRANCHAS) - set(quadros)
    if faltando:
        print(f"não achei as pranchas {sorted(faltando)} no slide")
        return 1

    zonas = [f for f in formas if f["fill"] in ("tema:accent6", "#FFFF99") and not f["txt"]]
    rotulos = [f for f in formas if re.search(r"\b(?:CHZ|JEI)-\d{3}\b", f["txt"] or "")]

    # Emparelha zona e rótulo pelo par mais próximo ainda livre: a caixa de
    # texto às vezes fica encostada por fora da forma, então conter o centro
    # não basta -- e um rótulo só serve a uma zona.
    import math
    pares = sorted((math.dist(centro(z), centro(t)), i, j)
                   for i, z in enumerate(zonas) for j, t in enumerate(rotulos))
    z_livres, t_livres, casado = set(range(len(zonas))), set(range(len(rotulos))), {}
    for _d, i, j in pares:
        if i in z_livres and j in t_livres:
            z_livres.discard(i)
            t_livres.discard(j)
            casado[i] = rotulos[j]

    saida = {"pranchas": []}
    DESTINO.mkdir(parents=True, exist_ok=True)
    for nome, (pid, rotulo) in PRANCHAS.items():
        q = quadros[nome]
        minhas = []
        for i, z in enumerate(zonas):
            cx, cy = centro(z)
            if not (q["x"] <= cx <= q["x"] + q["cx"] and q["y"] <= cy <= q["y"] + q["cy"]):
                continue
            t = casado.get(i)
            texto = t["txt"] if t else ""
            codigos = re.findall(r"(?:CHZ|JEI)-\d{3}\b", texto)
            # "JEI-001/003 à 007": a caixa cita a faixa em vez de listar uma a uma
            if "JEI" in texto and "007" in texto:
                codigos = [f"JEI-{n:03d}" for n in range(1, 8)]
            minhas.append({
                "desenhos": [f"800-{c}" for c in dict.fromkeys(codigos)],
                "x": (z["x"] - q["x"]) / q["cx"] * 100,
                "y": (z["y"] - q["y"]) / q["cy"] * 100,
                "l": z["cx"] / q["cx"] * 100,
                "a": z["cy"] / q["cy"] * 100,
                "recorte": z["recorte"],
            })
        if not minhas:
            print(f"prancha {pid}: nenhuma zona")
            return 1

        img = Image.open(io.BytesIO(midia[arquivo_do_rid[q["rid"]]])).convert("RGB")
        W, H = img.size
        # A folha do slide sobra: no pipe rack a faixa de interesse ocupa um
        # quinto de uma imagem quase vazia. O corte sai da própria marcação, com
        # margem, e as coordenadas são recalculadas no recorte novo.
        zx0 = min(z["x"] for z in minhas) / 100 * W
        zx1 = max(z["x"] + z["l"] for z in minhas) / 100 * W
        zy0 = min(z["y"] for z in minhas) / 100 * H
        zy1 = max(z["y"] + z["a"] for z in minhas) / 100 * H
        fx, fy = (zx1 - zx0) * 0.06 + 8, (zy1 - zy0) * 0.25 + 8
        x0, y0 = int(max(0, zx0 - fx)), int(max(0, zy0 - fy))
        x1, y1 = int(min(W, zx1 + fx)), int(min(H, zy1 + fy))
        corte = img.crop((x0, y0, x1, y1))
        alvo = LARGURA_ALVO[pid]
        if corte.width < alvo:
            corte = corte.resize((alvo, round(alvo * corte.height / corte.width)),
                                 Image.LANCZOS)
        buf = io.BytesIO()
        corte.save(buf, "PNG", optimize=True)
        (DESTINO / f"{pid}.png").write_bytes(buf.getvalue())

        for z in minhas:
            z["x"] = round((z["x"] / 100 * W - x0) / (x1 - x0) * 100, 3)
            z["y"] = round((z["y"] / 100 * H - y0) / (y1 - y0) * 100, 3)
            z["l"] = round(z["l"] / 100 * W / (x1 - x0) * 100, 3)
            z["a"] = round(z["a"] / 100 * H / (y1 - y0) * 100, 3)
        saida["pranchas"].append({
            "id": pid, "rotulo": rotulo, "arquivo": f"{pid}.png",
            "prop": round((y1 - y0) / (x1 - x0) * 100, 3),
            "zonas": sorted(minhas, key=lambda z: -z["l"] * z["a"]),
        })
        kb = len(buf.getvalue()) / 1024
        print(f"{pid:10} {corte.size!s:12} {len(minhas):2} zonas  {kb:6.0f} KB")

    (DESTINO / "zonas.json").write_text(
        json.dumps(saida, ensure_ascii=False, indent=1), encoding="utf-8")
    sem_codigo = [z for p in saida["pranchas"] for z in p["zonas"] if not z["desenhos"]]
    print(f"\nzonas.json: {sum(len(p['zonas']) for p in saida['pranchas'])} zonas"
          f"{f', {len(sem_codigo)} sem código' if sem_codigo else ''}")
    return 1 if sem_codigo else 0


if __name__ == "__main__":
    sys.exit(main(pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else PADRAO))
