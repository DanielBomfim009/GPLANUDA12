# -*- coding: utf-8 -*-
"""Varredura de abrir/fechar/navegar ficha em toda aba.

O que quebra numa ficha nao e "abre ou nao abre" -- e o meio do caminho: abrir
pelo desenho, descer para um relatorio, voltar, fechar. Cada elo desse
caminho e um jeito diferente de abrir (classe .fmodal-on vs ancora :target) e
um jeito diferente de fechar (X, fundo, Esc), e o defeito mora na transicao
entre eles, nao num estado isolado.

Este script varre cada aba, acha toda ficha que da para abrir -- pelo desenho,
pela tabela, pela lista -- abre, confere que so uma esta visivel, se tiver
"Detalhes" ou degrau de nivel desce nela, volta, fecha pelo X e confere que
fechou de verdade (sem classe nem :target residual). Reporta cada falha com o
caminho exato que a causou, para dar para reproduzir na hora.

    python varrer_fichas.py
    ALVO=https://gplan.onrender.com python varrer_fichas.py
"""
from __future__ import annotations

import os
import sys
import warnings

warnings.filterwarnings("ignore")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright

ALVO = os.environ.get("ALVO", "http://localhost:8501")
ESPERA = int(os.environ.get("ESPERA", "180")) * 1000
AMOSTRA = int(os.environ.get("AMOSTRA", "3"))  # quantas fichas testar por origem, por aba

ABAS = [
    ("", "Dashboard", ".du-tela"),
    ("/progresso", "Progresso", "details.arv-n1"),
    ("/relatorios", "Relatórios", "table.gtbl"),
    ("/pesquisa", "Pesquisa tag", "table.gtbl"),
    ("/sigem", "Base SIGEM", "table.gtbl"),
    ("/gitec", "Gitec", ".fx-tiles"),
    ("/planta", "Planta", ".pl-tela"),
    ("/certificacao", "Certificação", ".ct-painel"),
    ("/atualizacao", "Última atualização", ".ua-mov"),
]

JS_ABERTOS = ("() => [...document.querySelectorAll('.fmodal')]"
              ".filter(e => getComputedStyle(e).display !== 'none').map(e => e.id)")
JS_RESIDUAL = ("() => [...document.querySelectorAll('.fmodal.fmodal-on')].map(e => e.id)")

achados: list[str] = []


def abertos(pg) -> list[str]:
    return pg.evaluate(JS_ABERTOS)


def fechar_tudo(pg) -> None:
    pg.evaluate("() => { document.querySelectorAll('.fmodal-on')"
                ".forEach(m => m.classList.remove('fmodal-on'));"
                " if (location.hash) history.replaceState(null, '', location.pathname); }")
    pg.wait_for_timeout(300)


def iframe_do_desenho(pg):
    """O iframe que tem cartoes clicaveis, se a aba tiver mais de um iframe."""
    n = pg.locator("iframe").count()
    for i in range(n):
        cand = pg.frame_locator("iframe").nth(i)
        try:
            if cand.locator("[data-ficha]").count():
                return cand
        except Exception:
            continue
    return None


def testar_origem(pg, nome_aba: str, origem: str, alvos, abrir_um,
                  pagina_original: str = "") -> None:
    """Testa ate AMOSTRA elementos desta origem: abrir, checar unicidade,
    descer num Detalhes/degrau se existir, voltar, fechar, checar residuo."""
    total_real = alvos.count() if hasattr(alvos, "count") else len(alvos)
    if not total_real:
        return
    # Links de tabela/lista se repetem aos milhares (cada ficha ja aberta cita
    # TAGs relacionadas dentro dela). Amostrar no universo inteiro pega indices
    # a milhares de linhas de distancia -- 30 s so para rolar ate la, e isso e
    # o tempo de teste, nao um defeito do produto. A amostra fica no comeco,
    # onde a pessoa de fato clica.
    total = min(total_real, 60)
    passo = max(1, total // AMOSTRA)
    testados = 0
    for i in range(0, total, passo):
        if testados >= AMOSTRA:
            break
        testados += 1
        rotulo = f"{nome_aba} · {origem} #{i}"
        try:
            fechar_tudo(pg)
            if not abrir_um(i):
                continue
            pg.wait_for_timeout(900)
            vistos = abertos(pg)
            if len(vistos) == 0:
                achados.append(f"{rotulo}: clique não abriu ficha nenhuma")
                continue
            if len(vistos) > 1:
                achados.append(f"{rotulo}: mais de uma ficha visível ao mesmo tempo: {vistos}")

            # descer: Detalhes de relatorio, ou degrau de nivel/trilha. O
            # degrau pode mandar para a Progresso (href absoluto) de proposito
            # -- fichas de nivel sao compartilhadas por varias TAGs, sem
            # "voltar para esta TAG" fixo -- e a Progresso demora ~90s pra
            # carregar, entao esse caso pede uma espera bem maior.
            desceu = pg.locator(".fmodal-on a.btn-detalhes, .fmodal:target a.btn-detalhes,"
                                " .fmodal-on .fx-trilha a, .fmodal:target .fx-trilha a").first
            if desceu.count():
                href_antes = pg.evaluate("() => location.hash")
                href_desceu = desceu.get_attribute("href") or ""
                sai_da_pagina = href_desceu.startswith("/progresso")
                desceu.click(timeout=15000)
                if sai_da_pagina:
                    try:
                        pg.wait_for_selector("details.arv-n1", timeout=ESPERA)
                    except Exception:
                        pass
                pg.wait_for_timeout(900 if not sai_da_pagina else 2000)
                v2 = abertos(pg)
                if sai_da_pagina:
                    if not v2:
                        achados.append(f"{rotulo}: degrau mandou para a Progresso "
                                       f"mas a ficha não abriu lá (hash "
                                       f"{pg.evaluate('() => location.hash')})")
                    # a saida da pagina foi de proposito -- nao faz sentido
                    # testar "voltar" nela. Volta para a pagina original antes
                    # da proxima amostra, senao os locators desta origem
                    # (presos ao iframe/tabela de antes) ficam invalidos.
                    if pagina_original:
                        pg.goto(pagina_original, wait_until="domcontentloaded", timeout=ESPERA)
                    continue
                if not v2:
                    achados.append(f"{rotulo}: descer para o detalhe/degrau fechou tudo "
                                   f"(era {vistos}, hash {href_antes} -> "
                                   f"{pg.evaluate('() => location.hash')})")
                else:
                    # voltar: X do que abriu
                    voltar = pg.locator(".fmodal:target .fmodal-x, .fmodal-on .fmodal-x").first
                    if voltar.count():
                        voltar.click(timeout=15000)
                        pg.wait_for_timeout(900)
                        v3 = abertos(pg)
                        if not v3:
                            achados.append(f"{rotulo}: voltar do detalhe fechou tudo, "
                                           f"esperava reabrir {vistos}")

            # fechar pelo X do que estiver aberto agora
            x = pg.locator(".fmodal:target .fmodal-x, .fmodal-on .fmodal-x").first
            if not x.count():
                achados.append(f"{rotulo}: sem botão de fechar visível com a ficha aberta")
                continue
            x.click(timeout=15000)
            pg.wait_for_timeout(900)
            v4 = abertos(pg)
            if v4:
                achados.append(f"{rotulo}: o X não fechou -- continua aberto: {v4}")
            residuo = pg.evaluate(JS_RESIDUAL)
            if residuo and not v4:
                achados.append(f"{rotulo}: fechou visualmente mas ficou classe "
                               f"fmodal-on residual em {residuo}")
        except Exception as e:
            achados.append(f"{rotulo}: exceção durante o teste -- {type(e).__name__}: "
                           f"{str(e)[:160]}")


def varrer_pagina(pg, caminho: str, nome: str, marcador: str) -> None:
    print(f"\n  {nome}")
    try:
        pg.goto(ALVO + caminho, wait_until="domcontentloaded", timeout=ESPERA)
        pg.wait_for_selector(marcador, timeout=ESPERA)
    except Exception as e:
        print(f"    [FALHA] não abriu ({type(e).__name__})")
        achados.append(f"{nome}: a página não abriu -- {type(e).__name__}")
        return
    pg.wait_for_timeout(4000)
    if "Traceback" in pg.inner_text("body"):
        print("    [FALHA] traceback na página")
        achados.append(f"{nome}: traceback visível na página")
        return

    antes = len(achados)

    pagina_original = ALVO + caminho

    def clicar_desenho(i):
        # busca o iframe de novo a cada clique: uma amostra anterior pode ter
        # navegado para fora (degrau -> Progresso) e voltado, e o iframe de
        # antes ja nao existe mais no DOM atual.
        fr_atual = iframe_do_desenho(pg)
        if fr_atual is None:
            return False
        fr_atual.locator("[data-ficha]").nth(i).click(force=True, timeout=15000)
        return True

    fr = iframe_do_desenho(pg)
    n = fr.locator("[data-ficha]").count() if fr is not None else 0
    if n:
        testar_origem(pg, nome, "desenho", fr.locator("[data-ficha]"),
                     clicar_desenho, pagina_original)

    def clicar_lista(i):
        atual = pg.locator("a.gtbl-tag[href^='#f-'], a.ua-tag[href^='#f-']")
        if i >= atual.count():
            return False
        atual.nth(i).scroll_into_view_if_needed(timeout=8000)
        atual.nth(i).click(timeout=15000, force=True)
        return True

    links_ficha = pg.locator("a.gtbl-tag[href^='#f-'], a.ua-tag[href^='#f-']")
    if links_ficha.count():
        testar_origem(pg, nome, "tabela/lista", links_ficha, clicar_lista, pagina_original)

    print(f"    {n if fr is not None else 0} no desenho, "
          f"{links_ficha.count()} em tabela/lista -- "
          f"{'ok' if len(achados) == antes else str(len(achados) - antes) + ' problema(s)'}")


def main() -> int:
    print(f"VARREDURA DE FICHAS  {ALVO}")
    with sync_playwright() as p:
        nav = p.chromium.launch()
        pg = nav.new_page(viewport={"width": 1600, "height": 1050})
        erros_js: list[str] = []
        pg.on("pageerror", lambda e: erros_js.append(str(e)))
        for caminho, nome, marcador in ABAS:
            varrer_pagina(pg, caminho, nome, marcador)
        nav.close()

    print("\n" + "=" * 70)
    if achados:
        print(f"{len(achados)} problema(s) encontrado(s):\n")
        for a in achados:
            print(f"  - {a}")
        if erros_js:
            print(f"\n  {len(erros_js)} erro(s) de JavaScript:")
            for e in erros_js[:5]:
                print(f"    {e[:200]}")
        return 1
    print("Nada travou: toda ficha testada abriu sozinha, navegou e fechou limpa.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
