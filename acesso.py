# -*- coding: utf-8 -*-
"""Quem entra no Gplan e o que cada um pode ver.

Quem autentica é o Auth do Supabase. Não há senha guardada por este código,
nem hash, nem sal, nem token assinado à mão -- tudo isso é do Supabase, que
faz disso o ofício. O que sobra aqui é o que é do Gplan: papel e permissão,
numa tabela `perfis` ligada por id ao usuário do Auth.

A versão anterior guardava os logins num JSON no bucket e cifrava a senha
com scrypt. Funcionava, mas escrevia o arquivo inteiro a cada mudança -- dois
administradores editando ao mesmo tempo se sobrescreviam -- e a biblioteca
devolvia erro de escrita em vez de levantar, o que fez a tela anunciar um
administrador que nunca foi gravado. Numa tabela isso não acontece, e quem
esquece a senha passa a poder recuperá-la por e-mail em vez de depender de
alguém.

O que o Supabase precisa ter (feito em 20/08/2026):
  - tabela public.perfis, com RLS ligado e SEM política -- só a service_role
    enxerga, a API pública fica fechada;
  - gatilho ao_criar_usuario, que cria o perfil junto com a conta;
  - "Enable sign-ups" DESLIGADO: quem cria login é o administrador, pela tela.
"""
from __future__ import annotations

import os

TABELA = "perfis"

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
COR_PAPEL = {"Administrador": "roxo", "Colaborador": "teal",
             "Visualizador": "azul", "Apresentador": "ambar"}

ESPECIAIS = "!@#$%¨&*()-_=+[]{}^~/\\|;:,.<>?'\"`´"
REGRA_SENHA = ("Mínimo de 8 caracteres, com letra maiúscula, letra minúscula "
               "e caractere especial.")


class SemSupabase(RuntimeError):
    """Faltou credencial. É erro de configuração, não de senha."""


def problema_na_senha(senha: str) -> str | None:
    """A senha serve? Devolve o que falta, ou None quando está boa.

    Uma reclamação de cada vez, na ordem em que se digita: listar as quatro de
    uma vez faz a pessoa reler tudo para achar a que a pegou. O Supabase tem a
    própria exigência mínima; esta é a da casa, e a que dá a mensagem em
    português.
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


def papel_de(usuario: dict | None) -> str:
    """O papel gravado; na falta dele, o que as permissões dizem."""
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


def iniciais(nome: str, email: str = "") -> str:
    partes = [p for p in (nome or "").split() if p]
    if not partes:
        return (email or "?")[:2].upper()
    if len(partes) == 1:
        return partes[0][:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


def pode(usuario: dict | None, permissao: str) -> bool:
    if not usuario or not usuario.get("ativo", True):
        return False
    return permissao in (usuario.get("permissoes") or [])


# ===================================================================== #
# Conexão                                                               #
# ===================================================================== #

def _sem_proxy_para(url: str) -> None:
    """Tira o Supabase do caminho do proxy corporativo.

    A rede da AG exporta HTTPS_PROXY, e o supabase-py (via httpx) obedece. O
    proxy faz interceptação TLS e apresenta certificado próprio, que o Python
    recusa -- o erro é CERTIFICATE_VERIFY_FAILED, e parece problema de
    credencial quando é de rota. A conexão direta funciona e é verificada de
    verdade (TLSv1.3). Fora da rede corporativa não há o que contornar.
    """
    try:
        host = url.split("//", 1)[-1].split("/", 1)[0]
    except Exception:
        return
    for nome in ("NO_PROXY", "no_proxy"):
        atual = os.environ.get(nome, "")
        if host not in atual:
            os.environ[nome] = f"{atual},{host}".strip(",")


def _credenciais() -> tuple[str, str]:
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
        raise SemSupabase(
            "SUPABASE_URL e SUPABASE_KEY não estão configuradas. O login "
            "depende delas: sem as duas não há como autenticar ninguém.")
    return url, chave


def cliente():
    """Um cliente novo a cada chamada.

    De propósito: o cliente do supabase-py guarda a sessão de quem entrou
    dentro dele. Reaproveitar um único cliente entre pessoas faria a sessão de
    uma vazar para a próxima -- o mesmo tipo de contaminação que o cache já
    causou neste projeto.
    """
    url, chave = _credenciais()
    _sem_proxy_para(url)
    from supabase import create_client
    return create_client(url, chave)


# ===================================================================== #
# Perfil: o que é do Gplan, não do Auth                                 #
# ===================================================================== #

def _monta(uid: str, email: str, linha: dict | None) -> dict:
    linha = linha or {}
    perms = linha.get("permissoes") or []
    if isinstance(perms, str):
        import json
        try:
            perms = json.loads(perms)
        except Exception:
            perms = []
    return {"id": uid, "login": email, "email": email,
            "nome": linha.get("nome") or "",
            "papel": linha.get("papel") or PAPEL_PADRAO,
            "foto": linha.get("foto") or "",
            "permissoes": [p for p in perms if p in PERMISSOES],
            "ativo": bool(linha.get("ativo", True))}


def perfil(uid: str, email: str) -> dict:
    """O perfil da conta, lido SEMPRE com um cliente de serviço.

    Nunca com o cliente que acabou de autenticar: o supabase-py guarda a
    sessão dentro do cliente, e depois do sign_in ele deixa de falar como
    service_role e passa a falar como a pessoa. Como a tabela tem RLS ligado e
    nenhuma política, a linha some da consulta -- o código conclui que não
    existe, tenta inserir e leva "new row violates row-level security policy".
    A linha estava lá o tempo todo; quem mudou foi quem perguntou.

    Se o gatilho não tiver criado o perfil, cria agora: conta sem perfil
    entraria sem permissão nenhuma e pareceria defeito de permissão, quando é
    linha faltando.
    """
    cli = cliente()
    achado = cli.table(TABELA).select("*").eq("id", uid).execute().data
    if not achado:
        cli.table(TABELA).insert({"id": uid, "nome": ""}).execute()
        achado = cli.table(TABELA).select("*").eq("id", uid).execute().data
    return _monta(uid, email, achado[0] if achado else None)


# ===================================================================== #
# Entrar e sair                                                         #
# ===================================================================== #

def entrar(email: str, senha: str) -> tuple[dict, str] | None:
    """Autentica no Supabase. Devolve (usuário, refresh_token) ou None.

    O refresh_token é o que sobrevive no cookie: o access_token vence em uma
    hora, e guardar ele daria sessão que cai no meio do expediente.
    """
    cli = cliente()
    try:
        r = cli.auth.sign_in_with_password(
            {"email": str(email).strip().lower(), "password": senha})
    except Exception:
        return None
    if not (r and r.user and r.session):
        return None
    u = perfil(r.user.id, r.user.email or email)
    if not u["ativo"]:
        return None  # desativado entra como se a senha não conferisse
    return u, r.session.refresh_token


def retomar(refresh_token: str) -> tuple[dict, str] | None:
    """Recupera a sessão a partir do cookie, e devolve o token renovado."""
    if not refresh_token:
        return None
    cli = cliente()
    try:
        r = cli.auth.refresh_session(refresh_token)
    except Exception:
        return None
    if not (r and r.user and r.session):
        return None
    u = perfil(r.user.id, r.user.email or "")
    if not u["ativo"]:
        return None
    return u, r.session.refresh_token


def esquecer(refresh_token: str) -> None:
    """Encerra a sessão do lado do Supabase, e não só no navegador."""
    try:
        cli = cliente()
        cli.auth.refresh_session(refresh_token)
        cli.auth.sign_out()
    except Exception:
        pass


def recuperar_senha(email: str) -> None:
    """Manda o e-mail de redefinição. Exige SMTP configurado no projeto."""
    cliente().auth.reset_password_for_email(str(email).strip().lower())


# ===================================================================== #
# Administração                                                         #
# ===================================================================== #

def listar() -> list[dict]:
    """Todas as contas, já casadas com o perfil de cada uma."""
    cli = cliente()
    contas = cli.auth.admin.list_users()
    linhas = {r["id"]: r for r in cli.table(TABELA).select("*").execute().data}
    return sorted(
        (_monta(c.id, c.email or "", linhas.get(c.id)) for c in contas),
        key=lambda u: (u["nome"] or u["email"]).lower())


def criar(email: str, senha: str, nome: str, papel: str,
          permissoes: list[str] | None = None) -> dict:
    """Cria a conta já confirmada e grava o perfil.

    email_confirm=True porque quem cria é o administrador: exigir que a
    pessoa clique num link para existir só atrasaria, e o e-mail nem sempre
    chega (o mailer embutido do Supabase é limitado).
    """
    cli = cliente()
    r = cli.auth.admin.create_user({
        "email": str(email).strip().lower(),
        "password": senha,
        "email_confirm": True,
        "user_metadata": {"nome": nome or ""},
    })
    uid = r.user.id
    cli.table(TABELA).upsert({
        "id": uid, "nome": nome or "",
        "papel": papel if papel in PAPEIS else PAPEL_PADRAO,
        "permissoes": list(permissoes if permissoes is not None
                           else PAPEIS.get(papel, [])),
        "ativo": True,
    }).execute()
    return perfil(uid, r.user.email or email)


def salvar_perfil(uid: str, **campos) -> None:
    """Grava só o que veio, e nada além disso."""
    dados = {k: v for k, v in campos.items()
             if k in ("nome", "papel", "foto", "permissoes", "ativo")}
    if not dados:
        return
    cliente().table(TABELA).update(dados).eq("id", uid).execute()


def trocar_senha(uid: str, nova: str) -> None:
    cliente().auth.admin.update_user_by_id(uid, {"password": nova})


def trocar_minha_senha(email: str, atual: str, nova: str) -> bool:
    """Troca a própria senha, conferindo a atual.

    Confere entrando de novo com a senha atual: a sessão aberta prova quem é,
    não prova que a pessoa sabe a senha -- e máquina esquecida destrancada não
    pode virar troca de senha por quem passar.
    """
    if entrar(email, atual) is None:
        return False
    conta = next((c for c in cliente().auth.admin.list_users()
                  if (c.email or "").lower() == str(email).strip().lower()), None)
    if conta is None:
        return False
    trocar_senha(conta.id, nova)
    return True


def remover(uid: str) -> None:
    """Apaga a conta. O perfil vai junto, pelo on delete cascade."""
    cliente().auth.admin.delete_user(uid)
