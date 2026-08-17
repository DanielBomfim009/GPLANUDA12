# -*- coding: utf-8 -*-
"""Mede o contraste real de cada elemento visível, nos dois temas.

Ler a folha de estilo não prova nada: o que decide se um texto some é a cor
computada dele contra a cor computada do que está atrás -- que pode vir de um
ancestral, de um pseudo-elemento ou de uma regra de biblioteca. Aqui o
navegador é quem responde, elemento por elemento, em cada aba e em cada tema.

    python auditar_tema.py
    ALVO=https://gplan.onrender.com python auditar_tema.py

Sai com código 1 se algum elemento ficar abaixo do mínimo de contraste.
"""
from __future__ import annotations

import json
import os
import sys
import warnings

warnings.filterwarnings("ignore")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ALVO = os.environ.get("ALVO", "http://localhost:8501")
ESPERA = int(os.environ.get("ESPERA", "300")) * 1000

# WCGA AA pede 4,5:1 para texto pequeno e 3:1 para texto grande e para
# elemento gráfico. Aqui o corte é 3,0 para tudo: abaixo disso não é questão
# de gosto, é gente não conseguindo ler.
MINIMO = 3.0
# Rótulo em versalete e legenda de gráfico são texto pequeno de apoio; o corte
# maior existe porque é justamente onde a cor apagada costuma passar batido.
MINIMO_TEXTO = 4.0

ABAS = [("", "Dashboard", ".du-tela"),
        ("/progresso", "Progresso", "details.arv-n1"),
        ("/relatorios", "Relatórios", "table.gtbl"),
        ("/pesquisa", "Pesquisa tag", "table.gtbl"),
        ("/sigem", "Base SIGEM", "table.gtbl"),
        ("/gitec", "Gitec", ".fx-tiles"),
        ("/planta", "Planta", ".pl-tela"),
        ("/certificacao", "Certificação", ".ct-painel"),
        ("/atualizacao", "Última atualização", ".ua-mov")]

# O script roda no navegador: sobe a cor de fundo até achar quem de fato pinta,
# compõe as camadas semitransparentes e devolve o contraste de cada elemento.
SONDA = r"""
() => {
  const lum = ([r, g, b]) => {
    const f = v => { v /= 255; return v <= 0.03928 ? v/12.92 : Math.pow((v+0.055)/1.055, 2.4); };
    return 0.2126*f(r) + 0.7152*f(g) + 0.0722*f(b);
  };
  const razao = (a, b) => {
    const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p);
    return (x + 0.05) / (y + 0.05);
  };
  const rgba = txt => {
    txt = (txt || "").trim();
    // a paleta do :root vem em hex; o computed style dos elementos, em rgb()
    const h = txt.match(/^#([0-9a-f]{6})$/i);
    if (h) return [parseInt(h[1].slice(0,2),16), parseInt(h[1].slice(2,4),16),
                   parseInt(h[1].slice(4,6),16), 1];
    const m = txt.match(/[\d.]+/g);
    if (!m || m.length < 3) return null;
    return [+m[0], +m[1], +m[2], m.length > 3 ? +m[3] : 1];
  };
  const sobre = (frente, fundo) => frente.slice(0,3).map(
    (c, i) => Math.round(c * frente[3] + fundo[i] * (1 - frente[3])));

  // a cor que realmente aparece atras do elemento, empilhando os ancestrais
  const fundoDe = el => {
    const pilha = [];
    for (let n = el; n; n = n.parentElement) {
      const c = rgba(getComputedStyle(n).backgroundColor);
      if (c && c[3] > 0) { pilha.push(c); if (c[3] >= 0.999) break; }
    }
    let base = [255, 255, 255];
    for (let i = pilha.length - 1; i >= 0; i--) base = sobre(pilha[i], base);
    return base;
  };

  const visivel = el => {
    const s = getComputedStyle(el);
    if (s.display === "none" || s.visibility === "hidden" || +s.opacity < 0.15) return false;
    const r = el.getBoundingClientRect();
    return r.width > 2 && r.height > 2;
  };

  const achados = [];
  const marca = (el, tipo, cor, fundo, minimo, amostra) => {
    const r = razao(cor, fundo);
    if (r < minimo) achados.push({
      tipo, razao: +r.toFixed(2), minimo,
      classe: (el.className || "").toString().slice(0, 60),
      tag: el.tagName.toLowerCase(),
      cor: `rgb(${cor.join(",")})`, fundo: `rgb(${fundo.join(",")})`,
      texto: (amostra || "").trim().slice(0, 40)});
  };

  // ---- texto: so quem tem texto proprio, para nao medir o mesmo no pai ----
  for (const el of document.querySelectorAll("body *")) {
    if (!visivel(el)) continue;
    const proprio = [...el.childNodes]
      .filter(n => n.nodeType === 3 && n.textContent.trim())
      .map(n => n.textContent).join(" ").trim();
    if (!proprio) continue;
    const s = getComputedStyle(el);
    const cor = rgba(s.color);
    if (!cor || cor[3] < 0.5) continue;
    const fundo = fundoDe(el);
    const grande = parseFloat(s.fontSize) >= 19 || (+s.fontWeight >= 700 && parseFloat(s.fontSize) >= 15);
    marca(el, "texto", sobre(cor, fundo), fundo, grande ? MIN_G : MIN_T, proprio);
  }

  // ---- icone: mask-image (os do app) e svg com fill/stroke proprio -------
  for (const el of document.querySelectorAll("body *")) {
    if (!visivel(el)) continue;
    const s = getComputedStyle(el);
    const temMask = (s.maskImage && s.maskImage !== "none") ||
                    (s.webkitMaskImage && s.webkitMaskImage !== "none");
    if (!temMask) continue;
    const cor = rgba(s.backgroundColor);
    if (!cor || cor[3] < 0.2) continue;
    marca(el, "icone", sobre(cor, fundoDe(el.parentElement || el)),
          fundoDe(el.parentElement || el), MIN_G, "");
  }
  // <a>, <g>, <title> e <defs> nao pintam nada: herdam um fill preto de
  // fabrica que nao chega a aparecer. A marca do Gplan tambem fica de fora --
  // o anel de fundo dela e o vazio do medidor, discreto de proposito.
  const PINTA = new Set(["path", "circle", "rect", "polygon", "polyline", "line", "ellipse"]);
  for (const el of document.querySelectorAll("svg *")) {
    if (!PINTA.has(el.tagName.toLowerCase())) continue;
    // o anel de fundo da marca e da tela de carga e o vazio do medidor:
    // discreto de proposito, e nao carrega informacao nenhuma
    if (el.closest(".gplan-brand-mark, .gplan-brand, .gpl-mark")) continue;
    const s = getComputedStyle(el);
    for (const [prop, nome] of [["fill", "svg-fill"], ["stroke", "svg-stroke"]]) {
      const c = rgba(s[prop]);
      if (!c || c[3] < 0.5) continue;
      const svg = el.closest("svg");
      if (!svg || !visivel(svg)) continue;
      // traco de gauge/donut e area colorida grande, nao simbolo fino: o
      // limite util ali e o de elemento grafico
      marca(el, nome, sobre(c, fundoDe(svg)), fundoDe(svg), MIN_G, "");
    }
  }

  // ---- superficie que nao acompanhou o tema -------------------------------
  // Um campo escuro com letra clara tem contraste interno otimo -- a medicao
  // de texto passa direto por ele. O defeito e outro: a superficie ficou com a
  // cor do tema anterior e vira uma caixa escura no meio da tela clara. Foi
  // assim que o campo de busca da Pesquisa tag passou batido.
  {
    const fundoPagina = lum(fundoDe(document.querySelector(".stApp") || document.body));
    const claro = fundoPagina > 0.5;
    // Chip pintado com uma cor semantica e escuro de proposito no tema claro --
    // o ambar de leitura e #a16207. Comparar com a paleta viva do :root separa
    // "preenchimento da cor" de "superficie que ficou no tema anterior".
    const raiz = getComputedStyle(document.documentElement);
    const paleta = new Set(["--accent-blue", "--accent-purple", "--accent-green",
                            "--accent-red", "--accent-amber", "--accent-teal",
                            "--teal-2", "--neutro", "--neutro-2"]
      .map(n => (rgba(raiz.getPropertyValue(n).trim()) || []).slice(0, 3).join(",")));
    for (const el of document.querySelectorAll("body *")) {
      const s = getComputedStyle(el);
      const c = rgba(s.backgroundColor);
      if (!c || c[3] < 0.6) continue;
      const r = el.getBoundingClientRect();
      if (r.width < 40 || r.height < 12) continue;
      if (!visivel(el)) continue;
      if (paleta.has(c.slice(0, 3).join(","))) continue;
      const l = lum(c.slice(0, 3));
      // no claro, superficie quase preta; no escuro, superficie quase branca
      if ((claro && l < 0.25) || (!claro && l > 0.75)) {
        achados.push({tipo: "superficie", razao: +l.toFixed(2), minimo: claro ? 0.25 : 0.75,
          classe: (el.className || "").toString().slice(0, 60),
          tag: el.tagName.toLowerCase(), cor: s.backgroundColor,
          fundo: `luz da pagina ${fundoPagina.toFixed(2)}`,
          texto: (el.getAttribute("data-testid") || "")});
      }
    }
  }

  // ---- borda: separa o cartao do fundo? -----------------------------------
  for (const el of document.querySelectorAll(
      ".gplan-panel, .du-kpi, .pl-kpi, .fx-pn, .fx-tile, .fx-kpi, .gtbl, .pl-ch")) {
    if (!visivel(el)) continue;
    const s = getComputedStyle(el);
    if (parseFloat(s.borderTopWidth) < 0.5) continue;
    const c = rgba(s.borderTopColor);
    if (!c || c[3] < 0.02) continue;
    const fundo = fundoDe(el.parentElement || el);
    const r = razao(sobre(c, fundo), fundo);
    // borda e separacao, nao leitura: 1,15 ja da para enxergar o limite
    if (r < 1.12) achados.push({tipo: "borda", razao: +r.toFixed(2), minimo: 1.12,
      classe: (el.className || "").toString().slice(0, 60), tag: el.tagName.toLowerCase(),
      cor: s.borderTopColor, fundo: `rgb(${fundo.join(",")})`, texto: ""});
  }
  return achados;
}
"""


def main() -> int:
    from playwright.sync_api import sync_playwright

    problemas: list[str] = []
    with sync_playwright() as p:
        nav = p.chromium.launch()
        pg = nav.new_page(viewport={"width": 1600, "height": 1100})
        erros_js: list[str] = []
        pg.on("pageerror", lambda e: erros_js.append(str(e)))

        sonda = (SONDA.replace("MIN_T", str(MINIMO_TEXTO))
                 .replace("MIN_G", str(MINIMO)))

        for tema in ("escuro", "claro"):
            print(f"\nTEMA {tema.upper()}")
            for caminho, nome, marcador in ABAS:
                sep = "&" if "?" in caminho else "?"
                url = f"{ALVO}{caminho}{sep}tema={tema}"
                try:
                    pg.goto(url, wait_until="domcontentloaded", timeout=ESPERA)
                    pg.wait_for_selector(marcador, timeout=ESPERA)
                except Exception as e:
                    print(f"  [FALHA] {nome:14} não abriu ({type(e).__name__})")
                    problemas.append(f"{tema}/{nome}: não abriu")
                    continue
                pg.wait_for_timeout(3500)
                achados = pg.evaluate(sonda)
                # o mesmo defeito repetido em 300 linhas de tabela e um defeito
                vistos: dict[tuple, dict] = {}
                for a in achados:
                    vistos.setdefault((a["tipo"], a["classe"], a["cor"], a["fundo"]), a)
                unicos = list(vistos.values())
                marca = "ok  " if not unicos else "FALHA"
                print(f"  [{marca}] {nome:14} {len(achados):4} elementos abaixo do mínimo"
                      f"  ({len(unicos)} casos distintos)")
                for a in sorted(unicos, key=lambda x: x["razao"])[:6]:
                    print(f"          {a['razao']:5.2f} < {a['minimo']}  {a['tipo']:10}"
                          f" {a['tag']}.{a['classe'][:34]:34} {a['cor']} sobre {a['fundo']}"
                          f"  {a['texto']}")
                for a in unicos:
                    problemas.append(f"{tema}/{nome}: {a['tipo']} {a['classe'][:30]} "
                                     f"{a['razao']}:1")

        # ---- dentro das fichas -------------------------------------------
        # O modal nasce fechado, entao a varredura da aba nunca o alcanca --
        # e foi exatamente ali que o separador da trilha e a pilula "Montado"
        # passaram batido. Aqui uma ficha de cada tipo e aberta e medida.
        print("\nFICHAS (modal aberto)")
        FICHAS = [("/planta", ".pl-tela", "a.pl-zona"),
                  ("/progresso", "details.arv-n1", "a[href^='#n-FASE-']"),
                  ("/gitec", ".fx-tiles", "a.btn-detalhes")]
        for tema in ("escuro", "claro"):
            for caminho, marcador, gatilho in FICHAS:
                try:
                    pg.goto(f"{ALVO}{caminho}?tema={tema}", wait_until="domcontentloaded",
                            timeout=ESPERA)
                    pg.wait_for_selector(marcador, timeout=ESPERA)
                    pg.wait_for_timeout(4000)
                    alvo = pg.evaluate(
                        "sel => { const a = document.querySelector(sel);"
                        " return a ? a.getAttribute('href') : null; }", gatilho)
                    if not alvo or not alvo.startswith("#"):
                        print(f"  [ok  ] {tema:7} {caminho:11} sem ficha para abrir")
                        continue
                    pg.evaluate("h => location.hash = h", alvo)
                    pg.wait_for_timeout(1200)
                except Exception as e:
                    print(f"  [FALHA] {tema:7} {caminho:11} não abriu ({type(e).__name__})")
                    problemas.append(f"{tema}{caminho}: ficha não abriu")
                    continue
                achados = [a for a in pg.evaluate(sonda)
                           if pg.evaluate("() => !!document.querySelector('.fmodal:target')")]
                vistos = {(a["tipo"], a["classe"], a["cor"]): a for a in achados}
                marca = "ok  " if not vistos else "FALHA"
                print(f"  [{marca}] {tema:7} {caminho:11} {len(vistos)} casos na ficha aberta")
                for a in list(vistos.values())[:4]:
                    print(f"          {a['razao']:5.2f} {a['tipo']:10} .{a['classe'][:28]:28}"
                          f" {a['cor']} sobre {a['fundo']}  {a['texto']}")
                for a in vistos.values():
                    problemas.append(f"{tema} ficha{caminho}: {a['classe'][:26]} {a['razao']}:1")

        # ---- alternar varias vezes nao pode degradar -----------------------
        print("\nALTERNÂNCIA")
        pg.goto(f"{ALVO}?tema=escuro", wait_until="domcontentloaded", timeout=ESPERA)
        pg.wait_for_selector(".du-tela", timeout=ESPERA)
        pg.wait_for_timeout(3000)
        antes = pg.evaluate("() => getComputedStyle(document.body).backgroundColor")
        cores = []
        for i in range(4):
            # o segundo <button> e o alvo invisivel do tooltip, de tamanho zero
            pg.locator(".st-key-gplan_btn_tema button").first.click(force=True)
            pg.wait_for_timeout(2500)
            cores.append(pg.evaluate(
                "() => getComputedStyle(document.querySelector('.stApp')).backgroundColor"))
        alterna = len(set(cores)) == 2 and cores[0] != cores[1]
        print(f"  [{'ok  ' if alterna else 'FALHA'}] fundo alterna a cada clique  {cores}")
        if not alterna:
            problemas.append(f"alternância não troca o fundo: {cores}")
        depois = pg.evaluate("() => getComputedStyle(document.querySelector('.stApp')).backgroundColor")
        volta = depois == antes
        print(f"  [{'ok  ' if volta else 'FALHA'}] volta ao ponto de partida  {depois}")
        if not volta:
            problemas.append(f"depois de 4 trocas o fundo era {depois}, não {antes}")
        sobrando = pg.evaluate(
            "() => document.querySelectorAll('.st-key-gplan_btn_tema').length")
        print(f"  [{'ok  ' if sobrando == 1 else 'FALHA'}] um único botão de tema  {sobrando}")
        if sobrando != 1:
            problemas.append(f"{sobrando} botões de tema na tela")

        # ---- a preferência sobrevive a fechar e reabrir --------------------
        pg.goto(pg.url, wait_until="domcontentloaded", timeout=ESPERA)
        pg.wait_for_selector(".du-tela", timeout=ESPERA)
        pg.wait_for_timeout(2500)
        persistiu = pg.evaluate(
            "() => getComputedStyle(document.querySelector('.stApp')).backgroundColor") == depois
        print(f"  [{'ok  ' if persistiu else 'FALHA'}] preferência sobrevive ao recarregar")
        if not persistiu:
            problemas.append("a preferência de tema não sobreviveu ao recarregar")

        print(f"  [{'ok  ' if not erros_js else 'FALHA'}] erros de JavaScript "
              f"{erros_js[:1] or 'nenhum'}")
        problemas += erros_js[:1]
        nav.close()

    print()
    if problemas:
        print(f"{len(problemas)} problema(s):")
        for x in problemas[:40]:
            print(f"  - {x}")
        return 1
    print("Os dois temas passam: nada ilegível, nada invisível, alternância estável.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
