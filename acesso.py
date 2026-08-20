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
LOCAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".acesso.json")
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


# ===================================================================== #
# Senha                                                                 #
# ===================================================================== #

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

def _cliente():
    """O mesmo cliente do resto do app; None em modo local."""
    if os.environ.get("GPLAN_LOCAL") == "1":
        return None
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
    from supabase import create_client
    return create_client(url, chave)


def ler() -> dict:
    """O cofre inteiro: usuários e o segredo que assina a sessão."""
    cli = _cliente()
    if cli is None:
        if os.path.exists(LOCAL):
            with open(LOCAL, encoding="utf-8") as f:
                return json.load(f)
        return {"usuarios": {}, "segredo": ""}
    try:
        bruto = cli.storage.from_("gplan-data").download(ARQUIVO)
        return json.loads(bruto.decode("utf-8"))
    except Exception:
        # arquivo ainda não existe: primeiro uso
        return {"usuarios": {}, "segredo": ""}


def gravar(cofre: dict) -> None:
    dados = json.dumps(cofre, ensure_ascii=False, indent=2).encode("utf-8")
    cli = _cliente()
    if cli is None:
        with open(LOCAL, "wb") as f:
            f.write(dados)
        return
    # upsert: o update falha quando o arquivo ainda não existe
    try:
        cli.storage.from_("gplan-data").update(
            ARQUIVO, dados, {"content-type": "application/json", "upsert": "true"})
    except Exception:
        cli.storage.from_("gplan-data").upload(
            ARQUIVO, dados, {"content-type": "application/json", "upsert": "true"})


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

def novo_usuario(login: str, senha: str, nome: str, permissoes: list[str]) -> dict:
    return {"nome": nome or login,
            "senha": cifrar(senha),
            "permissoes": [p for p in permissoes if p in PERMISSOES],
            "ativo": True}


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
