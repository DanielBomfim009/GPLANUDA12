# -*- coding: utf-8 -*-
"""Quem entra no Gplan e o que cada um pode ver.

O problema que originou isto é concreto: apresentar o sistema para o cliente
sem mostrar valor em reais. Um botão de "modo apresentação" resolveria a tela,
mas não resolve o controle -- qualquer um clica de volta, inclusive na frente
do cliente. Então o que a tela mostra passa a ser consequência de quem entrou.

ONDE OS USUÁRIOS MORAM. No mesmo bucket do Supabase onde já mora a planilha,
num JSON. Não é banco, e para meia dúzia de logins não precisa ser: usa a
credencial que o app já tem, não pede migração de schema e não depende do
disco do Render, que é apagado a cada deploy. Em desenvolvimento
(GPLAN_LOCAL=1) o mesmo JSON fica num arquivo ao lado do app.

SENHA NUNCA É GRAVADA. O que fica guardado é o resultado do scrypt com sal
por usuário; a conferência é feita com compare_digest, que não vaza o
tamanho do acerto pelo tempo de resposta. Não existe caminho no código que
leia a senha de volta -- esquecer significa o administrador cadastrar outra.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import time

ARQUIVO = "gplan-acesso.json"
# O caminho do cofre local é sobrescritível por variável de ambiente, e isso
# não é conveniência: um teste que grava direto no cofre de verdade apaga os
# logins reais de quem estiver usando a máquina. Aconteceu em 20/08/2026 --
# o administrador recém-criado foi por cima. Teste aponta para o seu próprio
# arquivo; sem a variável, o caminho é o de sempre.
LOCAL = os.environ.get(
    "GPLAN_ACESSO_ARQUIVO",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), ".acesso.json"))
VALIDADE = 12 * 3600  # a sessão dura um dia de trabalho

# O catálogo é o contrato: a tela pergunta por estes nomes e a administração
# oferece exatamente estes. Permissão que não está aqui não existe.
PERMISSOES = {
    "ver_valores": "Ver valores em reais",
    "ver_gitec": "Ver a aba Gitec (medição de campo)",
    "ver_certificacao": "Ver a aba Certificação",
    "ver_planta": "Ver a aba Planta",
    "administrar": "Criar, editar e remover logins",
}
TODAS = list(PERMISSOES)

# O papel é o atalho: escolhe um e as permissões vêm prontas. Depois disso
# elas continuam editáveis uma a uma -- o papel é o ponto de partida e o
# rótulo que aparece na tela, não uma jaula.
PAPEIS = {
    "Administrador": TODAS,
    "Colaborador": ["ver_valores", "ver_gitec", "ver_certificacao", "ver_planta"],
    "Visualizador": ["ver_valores", "ver_certificacao", "ver_planta"],
    "Apresentador": ["ver_certificacao", "ver_planta"],
}
PAPEL_PADRAO = "Colaborador"

# a cor do selo: administrador se distingue de longe, apresentador avisa que
# aquela sessão está sem valores
COR_PAPEL = {"Administrador": "roxo", "Colaborador": "teal",
             "Visualizador": "azul", "Apresentador": "ambar"}


def papel_de(usuario: dict | None) -> str:
    """O papel gravado; na falta dele, o que as permissões dizem.

    Login criado antes de existir papel não fica sem rótulo: a permissão de
    administrar já denuncia o administrador, e ver_valores separa quem
    apresenta de quem trabalha.
    """
    if not usuario:
        return ""
    if usuario.get("papel") in PAPEIS:
        return usuario["papel"]
    perms = usuario.get("permissoes") or []
    if "administrar" in perms:
        return "Administrador"
    if "ver_valores" not in perms:
        return "Apresentador"
    return "Colaborador"


# ===================================================================== #
# Senha                                                                 #
# ===================================================================== #

ESPECIAIS = "!@#$%¨&*()-_=+[]{}^~/\\|;:,.<>?'\"`´"

REGRA_SENHA = ("Mínimo de 8 caracteres, com letra maiúscula, letra minúscula "
               "e caractere especial.")


def problema_na_senha(senha: str) -> str | None:
    """A senha serve? Devolve o que falta, ou None quando está boa.

    Uma reclamação de cada vez, na ordem em que se digita: listar as quatro
    de uma vez faz a pessoa relerem tudo para achar a que a pegou.
    """
    if len(senha) < 8:
        return "A senha precisa ter pelo menos 8 caracteres."
    if not any(c.isupper() for c in senha):
        return "A senha precisa ter pelo menos uma letra maiúscula."
    if not any(c.islower() for c in senha):
        return "A senha precisa ter pelo menos uma letra minúscula."
    if not any(c in ESPECIAIS for c in senha):
        return ("A senha precisa ter pelo menos um caractere especial "
                "(por exemplo ! @ # $ % & *).")
    return None


def cifrar(senha: str) -> str:
    """scrypt com sal novo a cada senha. n=2**14 é o custo recomendado para
    uso interativo -- alto o bastante para atrapalhar força bruta, baixo o
    bastante para o login não pesar."""
    sal = secrets.token_bytes(16)
    bruto = hashlib.scrypt(senha.encode("utf-8"), salt=sal, n=2 ** 14, r=8, p=1,
                           dklen=32)
    return f"scrypt${base64.b64encode(sal).decode()}${base64.b64encode(bruto).decode()}"


def confere(senha: str, guardado: str) -> bool:
    try:
        algoritmo, sal_b64, alvo_b64 = str(guardado).split("$")
        if algoritmo != "scrypt":
            return False
        bruto = hashlib.scrypt(senha.encode("utf-8"),
                               salt=base64.b64decode(sal_b64),
                               n=2 ** 14, r=8, p=1, dklen=32)
    except Exception:
        return False
    return hmac.compare_digest(bruto, base64.b64decode(alvo_b64))


# ===================================================================== #
# Onde ficam guardados                                                  #
# ===================================================================== #

def _sem_proxy_para(url: str) -> None:
    """Tira o Supabase do caminho do proxy corporativo.

    A rede da AG exporta HTTPS_PROXY, e o supabase-py (via httpx) obedece.
    O proxy faz interceptação TLS e apresenta um certificado próprio, que o
    Python recusa -- o erro que aparece é CERTIFICATE_VERIFY_FAILED, e ele
    parece problema de credencial quando é de rota. A conexão direta ao
    Supabase funciona e é verificada de verdade (TLSv1.3), então o certo é
    não mandar esse tráfego pelo proxy.

    Some sozinho fora da rede corporativa: sem proxy definido, não há nada a
    contornar e a variável simplesmente não atrapalha.
    """
    try:
        host = url.split("//", 1)[-1].split("/", 1)[0]
    except Exception:
        return
    for nome in ("NO_PROXY", "no_proxy"):
        atual = os.environ.get(nome, "")
        if host not in atual:
            os.environ[nome] = f"{atual},{host}".strip(",")


def _cliente():
    """O cliente do Supabase, ou None quando não há credencial.

    GPLAN_LOCAL=1 NÃO desliga o Supabase aqui, e essa é a diferença que
    importa: aquela variável quer dizer "leia a planilha do disco", e login é
    outro assunto. Desenvolver com a planilha local e os usuários de verdade é
    o caso normal -- sem isso, testar o acesso em localhost exigiria subir o
    sistema inteiro no Render, que é justamente o que trava.

    Sem credencial nenhuma, cai no arquivo local -- assim quem clonar o
    projeto ainda consegue rodar.
    """
    url = os.environ.get("SUPABASE_URL")
    chave = os.environ.get("SUPABASE_KEY")
    if not (url and chave):
        try:
            import streamlit as st
            url = url or st.secrets.get("SUPABASE_URL")
            chave = chave or st.secrets.get("SUPABASE_KEY")
        except Exception:
            pass
    if not (url and chave):
        return None
    _sem_proxy_para(url)
    from supabase import create_client
    return create_client(url, chave)


def ler() -> dict:
    """O cofre inteiro: usuários e o segredo que assina a sessão.

    A falha de leitura vai em "_erro" em vez de virar cofre vazio calado: sem
    isso, bucket fora do ar e primeiro uso se parecem na tela -- os dois
    mostram "crie o administrador" -- e criar por cima apaga quem já existia.
    """
    cli = _cliente()
    if cli is None:
        cofre = {"usuarios": {}, "segredo": ""}
        if os.path.exists(LOCAL):
            with open(LOCAL, encoding="utf-8") as f:
                cofre = json.load(f)
        cofre["_origem"] = "local"
        cofre["_erro"] = ""
        return cofre
    try:
        bruto = cli.storage.from_("gplan-data").download(ARQUIVO)
        cofre = json.loads(bruto.decode("utf-8"))
        cofre["_origem"] = "supabase"
        cofre["_erro"] = ""
        return cofre
    except Exception as erro:
        texto = f"{type(erro).__name__}: {erro}"
        # "não achei o arquivo" é o primeiro uso de verdade; qualquer outra
        # coisa é problema de acesso, e aí a tela não pode oferecer criar
        primeiro_uso = any(p in texto.lower()
                           for p in ("not_found", "not found", "404",
                                     "object not found"))
        return {"usuarios": {}, "segredo": "", "_origem": "supabase",
                "_erro": "" if primeiro_uso else texto}


def gravar(cofre: dict) -> None:
    """Grava e CONFERE relendo.

    O supabase-py nem sempre levanta exceção quando a escrita é recusada --
    em vários casos devolve um erro que, ignorado, faz o código concluir que
    salvou. Foi o que aconteceu em produção em 20/08/2026: a tela disse
    "administrador criado" e no arquivo não havia nada. Reler e comparar é o
    que transforma "mandei gravar" em "está gravado".
    """
    limpo = {k: v for k, v in cofre.items() if not k.startswith("_")}
    dados = json.dumps(limpo, ensure_ascii=False, indent=2).encode("utf-8")
    cli = _cliente()
    if cli is None:
        with open(LOCAL, "wb") as f:
            f.write(dados)
        return

    balde = cli.storage.from_("gplan-data")
    opcoes = {"content-type": "application/json", "upsert": "true"}
    erros = []
    for tentativa in (lambda: balde.update(ARQUIVO, dados, opcoes),
                      lambda: balde.upload(ARQUIVO, dados, opcoes)):
        try:
            tentativa()
        except Exception as erro:
            erros.append(f"{type(erro).__name__}: {erro}")
            continue
        # a prova: releia e veja se os logins bateram
        try:
            de_volta = json.loads(balde.download(ARQUIVO).decode("utf-8"))
        except Exception as erro:
            erros.append(f"releitura falhou: {type(erro).__name__}: {erro}")
            continue
        if set(de_volta.get("usuarios", {})) == set(limpo.get("usuarios", {})):
            return
        erros.append("o arquivo relido não tem os logins que acabaram de ser "
                     "gravados")
    raise RuntimeError("não consegui gravar o cofre no Supabase Storage. "
                       + " | ".join(erros))


def segredo(cofre: dict) -> str:
    """A chave que assina o cookie de sessão. Nasce junto com o primeiro
    login e fica no cofre: trocá-la derruba todas as sessões, que é o
    comportamento certo se alguém suspeitar de vazamento."""
    if not cofre.get("segredo"):
        cofre["segredo"] = secrets.token_hex(32)
        gravar(cofre)
    return cofre["segredo"]


# ===================================================================== #
# Usuários                                                              #
# ===================================================================== #

def novo_usuario(login: str, senha: str, nome: str, permissoes: list[str],
                 email: str = "", papel: str = PAPEL_PADRAO) -> dict:
    return {"nome": nome or login,
            "email": (email or "").strip(),
            "papel": papel if papel in PAPEIS else PAPEL_PADRAO,
            "foto": "",  # data URI, gravado pelo próprio dono no perfil
            "senha": cifrar(senha),
            "permissoes": [p for p in permissoes if p in PERMISSOES],
            "ativo": True}


def iniciais(nome: str, login: str = "") -> str:
    partes = [p for p in (nome or login or "?").split() if p]
    if not partes:
        return "?"
    if len(partes) == 1:
        return partes[0][:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


def autenticar(cofre: dict, login: str, senha: str) -> dict | None:
    u = cofre.get("usuarios", {}).get(str(login).strip().lower())
    if not u or not u.get("ativo", True):
        return None
    if not confere(senha, u.get("senha", "")):
        return None
    return u


def pode(usuario: dict | None, permissao: str) -> bool:
    if not usuario:
        return False
    return permissao in (usuario.get("permissoes") or [])


# ===================================================================== #
# Sessão                                                                #
# ===================================================================== #

def assinar(login: str, chave: str) -> str:
    """Cookie de sessão: login, validade e a assinatura dos dois. Sem a
    assinatura daria para trocar o login no navegador e virar outro."""
    ate = int(time.time()) + VALIDADE
    corpo = f"{login}|{ate}"
    marca = hmac.new(chave.encode(), corpo.encode(), hashlib.sha256).hexdigest()
    return f"{corpo}|{marca}"


def conferir_assinatura(token: str, chave: str) -> str | None:
    try:
        login, ate, marca = str(token).split("|")
    except ValueError:
        return None
    corpo = f"{login}|{ate}"
    esperado = hmac.new(chave.encode(), corpo.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(marca, esperado):
        return None
    if int(ate) < int(time.time()):
        return None
    return login
