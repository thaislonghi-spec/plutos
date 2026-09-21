# -*- coding: utf-8 -*-
"""
PLUTOS · Rebates · Multimóveis
Flask + arquivos JSON em DATA_DIR (multiempresa desde o início).
"""
from __future__ import annotations

import io
import gc
import json
from collections import OrderedDict
import os
import re
import secrets
import shutil
from datetime import date, datetime, timedelta, timezone
from functools import wraps
from typing import Any

from werkzeug.exceptions import HTTPException
from flask import (Flask, abort, flash, jsonify, redirect, render_template, request,
                   send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from motor import meli, erp, magalu, magalu_vendas, magalu_full, magalu_real, shopee, madeira, webcont, colombo, amazon
import planilhas

VERSAO = "2026-09-21h"
RAIZ = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR") or os.path.join(os.path.dirname(RAIZ), "dados")
os.makedirs(DATA_DIR, exist_ok=True)
BRT = timezone(timedelta(hours=-3))

app = Flask(__name__, template_folder=os.path.join(RAIZ, "templates"),
            static_folder=os.path.join(RAIZ, "static"))
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(16)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# --------------------------------------------------------------------------
# CANAIS — a ordem é a que a Thaís pediu (10/09/2026). Só o Meli tem motor;
# os outros aparecem como box "em construção" até cada um ser desenhado.
# --------------------------------------------------------------------------
CANAIS = [
    {"chave": "meli",     "nome": "Mercado Livre",   "ativo": True,
     "arquivo": "TABELA GERAL DE PEDIDOS", "arquivo_sub": "export do Power BI · filtro Pago · dia 01 a 31",
     "extra": None},
    {"chave": "magalu",   "nome": "Magazine Luiza",  "ativo": True,
     "arquivo": "FINANCEIRO POR PERÍODO", "arquivo_sub": "export do portal Magalu · 1 linha por pedido · dia 01 a 31",
     "extra": None},
    {"chave": "madeira",  "nome": "Madeira Madeira", "ativo": True,
     "arquivo": "RELATÓRIO DE PEDIDOS (portal Madeira)", "arquivo_sub": "MadeiraMadeira_<seller>-report_pedido_… · csv ; · 1 linha por item · dia 01 a 31",
     "extra": None},
    {"chave": "colombo",  "nome": "Colombo",         "ativo": True,
     "arquivo": "RELATÓRIO DE PEDIDOS (portal Colombo)", "arquivo_sub": "Colombo_Pedidos_DDMMateDDMMAA.csv · csv ; · 1 linha por item · dia 01 a 31",
     "extra": None},
    {"chave": "cbahia",   "nome": "Casas Bahia",     "ativo": False},
    {"chave": "amazon",   "nome": "Amazon",          "ativo": True,
     "arquivo": "TRANSAÇÕES (Seller Central → Pagamentos)", "arquivo_sub": "Amazon_Transações referentes ao período de DD_MM_AAAA a DD_MM_AAAA.csv · 1 linha por transação · acumula por ID do pedido",
     "extra": None},
    {"chave": "webcont",  "nome": "Webcontinental",  "ativo": True,
     "arquivo": "RELATÓRIO DE PEDIDOS (portal Webcontinental)", "arquivo_sub": "relatorio_pedidos_webcontinental_DDMMateDDMMAA.xlsx · aba Pedidos · 1 linha por pedido · dia 01 a 31",
     "extra": None},
    {"chave": "shopee",   "nome": "Shopee",          "ativo": True,
     "arquivo": "ORDER.ALL (Meus pedidos → Exportar)", "arquivo_sub": "export do portal Shopee · 1 linha por item · dia 01 a 31",
     "extra": None},
]
CANAIS_BASE = CANAIS


def canais() -> list[dict]:
    """Os 8 boxes da casa + os criados pela tela Arquivos (parametros.json → boxes_extra)."""
    extra = parametros().get("boxes_extra") or []
    out = [dict(c) for c in CANAIS_BASE]
    mapa = {c["chave"]: c for c in out}
    for e in extra:
        if e.get("chave") in mapa:
            mapa[e["chave"]]["nomes_erp"] = e.get("nomes_erp") or []
            continue
        out.append({"chave": e["chave"], "nome": e["nome"], "ativo": False, "extra": None,
                    "nomes_erp": e.get("nomes_erp") or [], "criado": True})
    return out


def canal_por_chave() -> dict:
    return {c["chave"]: c for c in canais()}


def mapa_erp_box() -> dict:
    """nome no ERP (sem 'MP - ') → chave do box. Base do motor + o que foi criado na tela."""
    m = dict(erp.CANAL_BOX)
    for c in canais():
        for n in c.get("nomes_erp") or []:
            m[n.strip().upper()] = c["chave"]
    return m


def reclassificar_erp():
    idx = erp_ler()
    m = mapa_erp_box()
    for l in idx["ocs"].values():
        l["box"] = m.get(l["canal_nome"], "outros")
    erp_gravar(idx)

PAPEIS = {"admin": "Admin", "gestor": "Gestor", "operador": "Operador", "equipe": "Equipe"}
PAPEL_DESC = {"admin": "faz tudo, inclusive criar, resetar, desativar e apagar usuário", "gestor": "tudo menos usuários",
              "operador": "sobe arquivos e preenche as tabelas manuais", "equipe": "só consulta"}
# quem pode o quê
PODE = {
    "arquivos":   {"admin", "gestor", "operador"},
    "faltante":   {"admin", "gestor", "operador"},
    "parametros": {"admin", "gestor"},
    "usuarios":   {"admin"},
    "exportar":   {"admin", "gestor", "operador"},
}

IRMAOS = [
    ("tropa",    "TROPA DE ELITE", os.environ.get("URL_TROPA", "https://tropadeelite.onrender.com/")),
    ("orion",    "ORION",          os.environ.get("URL_ORION", "https://orion-flash.onrender.com/")),
    ("hercules", "HÉRCULES",       os.environ.get("URL_HERCULES", "https://radar-hercules.onrender.com/")),
    ("atlas",    "ATLAS",          os.environ.get("URL_ATLAS", "https://atlas-pcp.onrender.com/")),
]


# --------------------------------------------------------------------------
# utilidades
# --------------------------------------------------------------------------
def agora():
    return datetime.now(BRT)


# Cache dos JSONs em memória, por (caminho, mtime, tamanho): cada tela lia o
# índice do ERP, as rodadas, a Planilha 2 e o cadastro várias vezes por
# requisição — no Render (512 MB) isso derrubava o worker (502). Quem altera
# um JSON sempre grava em seguida (_json_gravar), que invalida a entrada.
_CACHE: OrderedDict = OrderedDict()
CACHE_MAX_ARQ = 1_500_000     # bytes: JSON maior que isso NUNCA fica em memória
CACHE_MAX_ITENS = 30          # no máximo 30 arquivos em cache (LRU)
_CACHE_LOCK = __import__("threading").Lock()


def _json_ler(caminho, padrao):
    try:
        st = os.stat(caminho)
    except OSError:
        return padrao
    chave = (st.st_mtime_ns, st.st_size)
    with _CACHE_LOCK:
        hit = _CACHE.get(caminho)
        if hit and hit[0] == chave:
            _CACHE.move_to_end(caminho)
            return hit[1]
    try:
        with open(caminho, encoding="utf-8") as f:
            dado = json.load(f)
    except Exception:
        return padrao
    # o cache existe para as leituras pequenas e repetidas (parâmetros, índices,
    # resumos). Arquivo grande NUNCA fica em memória: no Render são 512 MB e foi
    # isso que derrubou o app em 13/09 (rodada gigante da Shopee).
    if st.st_size <= CACHE_MAX_ARQ:
        with _CACHE_LOCK:
            _CACHE[caminho] = (chave, dado)
            _CACHE.move_to_end(caminho)
            while len(_CACHE) > CACHE_MAX_ITENS:
                _CACHE.popitem(last=False)
    return dado


def _json_gravar(caminho, dado):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    tmp = caminho + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dado, f, ensure_ascii=False, indent=1, default=str)
    os.replace(tmp, caminho)
    with _CACHE_LOCK:
        _CACHE.pop(caminho, None)


def empresa_atual() -> str:
    return session.get("empresa") or "multimoveis"


def pasta(*p) -> str:
    c = os.path.join(DATA_DIR, empresa_atual(), *p)
    os.makedirs(os.path.dirname(c) if "." in os.path.basename(c) else c, exist_ok=True)
    return c


# COMISSÃO CADASTRADA POR CANAL — a tabela da casa (enviada em 10/09/2026).
# Comissão em %, taxa extra em R$ por item, tipo = base da comissão (Produto / GMV).
# No Meli a comissão é dupla: Premium 16,5% / Clássico 11,5%.
COMISSAO_PADRAO = [
    {"codigo": "9100", "canal": "CRC - MULTIMOVEIS", "gestor": "Marcão", "comissao": "0", "taxa": "0", "tipo": "-"},
    {"codigo": "9000", "canal": "E-COMMERCE MULTIMOVEIS", "gestor": "Geverton Hemsing", "comissao": "6", "taxa": "0", "tipo": "-"},
    {"codigo": "9997", "canal": "MARKETPLACE", "gestor": "Marketplace", "comissao": "0", "taxa": "0", "tipo": "-"},
    {"codigo": "762", "canal": "MP - ALI EXPRESS", "gestor": "Renan Stroeher", "comissao": "6", "taxa": "0", "tipo": "Produto"},
    {"codigo": "716", "canal": "MP - AMAZON.COM.BR", "gestor": "Renan Stroeher", "comissao": "9", "taxa": "0", "tipo": "GMV"},
    {"codigo": "788", "canal": "MP - BANCO DO BRASIL", "gestor": "Bruna Colares", "comissao": "17", "taxa": "0", "tipo": "GMV"},
    {"codigo": "754", "canal": "MP - BRADESCO - NEXT SHOP", "gestor": "Renan Stroeher", "comissao": "12", "taxa": "0", "tipo": "GMV"},
    {"codigo": "749", "canal": "MP - BANCO INTER", "gestor": "William Henzel", "comissao": "18,5", "taxa": "0", "tipo": "GMV"},
    {"codigo": "713", "canal": "MP - CARREFOUR", "gestor": "Bruna Colares", "comissao": "16", "taxa": "0", "tipo": "GMV"},
    {"codigo": "712", "canal": "MP - CASAS BAHIA", "gestor": "William Henzel", "comissao": "13", "taxa": "0", "tipo": "GMV"},
    {"codigo": "705", "canal": "MP - COLOMBO", "gestor": "Bruna Colares", "comissao": "7", "taxa": "0", "tipo": "GMV"},
    {"codigo": "776", "canal": "MP - IMPERIO", "gestor": "William Henzel", "comissao": "20", "taxa": "0", "tipo": "GMV"},
    {"codigo": "711", "canal": "MP - LEROY MERLIN", "gestor": "William Henzel", "comissao": "16", "taxa": "", "tipo": "GMV"},
    {"codigo": "795", "canal": "MP - LOJAS KOERICH", "gestor": "", "comissao": "5", "taxa": "", "tipo": "Produto"},
    {"codigo": "704", "canal": "MP - MADEIRA MADEIRA", "gestor": "Bruna Colares", "comissao": "17", "taxa": "", "tipo": "GMV"},
    {"codigo": "702", "canal": "MP - MAGAZINE LUIZA", "gestor": "William Henzel", "comissao": "11", "taxa": "5", "tipo": "GMV"},
    {"codigo": "700", "canal": "MP - MERCADO LIVRE", "gestor": "Eduardo Tomazi", "comissao": "16,5/11,5", "taxa": "", "tipo": "Produto"},
    {"codigo": "1002", "canal": "MP - MERCADO LIVRE - FULL", "gestor": "Full", "comissao": "16,5/11,5", "taxa": "", "tipo": "Produto"},
    {"codigo": "767", "canal": "MP - QUERO QUERO", "gestor": "Bruna Colares", "comissao": "18", "taxa": "0", "tipo": "GMV"},
    {"codigo": "792", "canal": "MP - SENFF SHOPPING", "gestor": "Bruna Colares", "comissao": "14,5", "taxa": "", "tipo": "Produto"},
    {"codigo": "752", "canal": "MP - SHOPEE", "gestor": "Renan Stroeher", "comissao": "12", "taxa": "12", "tipo": "Produto"},
    {"codigo": "782", "canal": "MP - SHOPEE XPRESS", "gestor": "Full", "comissao": "14", "taxa": "", "tipo": "Produto"},
    {"codigo": "770", "canal": "MP - SICREDI", "gestor": "Bruna Colares", "comissao": "15", "taxa": "", "tipo": "GMV"},
    {"codigo": "790", "canal": "MP - TIKTOK", "gestor": "Marketplace", "comissao": "12", "taxa": "", "tipo": "GMV"},
    {"codigo": "789", "canal": "MP - VALE BONUS", "gestor": "Bruna Colares", "comissao": "12", "taxa": "", "tipo": "GMV"},
    {"codigo": "728", "canal": "MP - WEBCONTINENTAL", "gestor": "Bruna Colares", "comissao": "19", "taxa": "", "tipo": "GMV"},
]
# A TABELA DE PARÂMETROS POR CANAL (21/09/2026) — tudo que muda de canal para
# canal mora aqui e vale para o sistema inteiro:
#   comissao     % negociado com o canal
#   tx_fin       % de taxa financeira (antecipação) — NÃO é comissão e fica FORA
#                do rebate; só separada no Linha a linha (Madeira 4%, Webcont 1%)
#   pct_extra    % de taxa cobrada pelo canal ALÉM da comissão (Amazon 1,5%);
#                entra na comissão do sistema, porque o canal cobra junto
#   taxa_pedido  R$ fixos por PEDIDO (Magalu R$ 5,00)
#   taxa_item    R$ fixos por ITEM (Shopee R$ 12,00 — já vem dentro da taxa de
#                serviço do relatório dela, por isso o motor não soma de novo)
#   tipo         base da comissão: Produto (produto + IPI) ou GMV (total da NF)
COMISSAO_CAMPOS = ["codigo", "canal", "gestor", "comissao", "tx_fin", "pct_extra",
                   "taxa_pedido", "taxa_item", "tipo"]
COMISSAO_LEGADO = "taxa"      # coluna antiga: virou taxa_pedido / taxa_item


def parametros() -> dict:
    padrao = {"empresa": "Multimóveis", "tolerancia_comissao": 0.50,
              "canais_ativos": [c["chave"] for c in CANAIS_BASE if c["ativo"]],
              "comissoes": COMISSAO_PADRAO,
              # canais em que o % do ERP embute uma taxa financeira (antecipação):
              # o Linha a linha separa "Comissão canal" e "Tx financeira canal"
              "tx_financeira": {"madeira": 4.0, "webcont": 1.0},
              # PRAZO DE ENVIO AO FULL, em dias, por CD (regra da Thaís, 17/09/2026):
              #   manuseio     = o nosso tempo de separar, embalar e faturar
              #   transferencia = o trânsito até o CD do canal
              # A sugestão de envio soma manuseio + transferência à cobertura desejada.
              "prazos_full": {"Guarulhos - SP": {"manuseio": 2, "transferencia": 4},
                              "Candeias - BA": {"manuseio": 2, "transferencia": 20}},
              "prazo_full_padrao": {"manuseio": 2, "transferencia": 10}}
    p = _json_ler(pasta("parametros.json"), {})
    p = _corrigir_comissoes(p)
    p = _migrar_comissoes(p)
    p = {**padrao, **p}
    # prazo antigo (um número só) → {manuseio, transferencia}, sem perder o valor
    def _pz(v, tra_padrao=0):
        if isinstance(v, dict):
            return {"manuseio": int(v.get("manuseio") or 0), "transferencia": int(v.get("transferencia") or 0)}
        try:
            return {"manuseio": 0, "transferencia": int(v or 0)}
        except (TypeError, ValueError):
            return {"manuseio": 0, "transferencia": tra_padrao}
    p["prazo_full_padrao"] = _pz(p.get("prazo_full_padrao"))
    p["prazos_full"] = {k: _pz(v) for k, v in (p.get("prazos_full") or {}).items()}
    return p


# Correções de cadastro que o app aplica UMA VEZ na tabela salva (e registra em
# "correcoes" para nunca repetir — se ela editar depois, o valor dela vale).
CORRECOES_COMISSAO = {
    # chave da correção: (pedaço do nome do canal, valor errado, valor certo, porquê)
    "colombo_7": ("COLOMBO", "9", "7", "o portal e o ERP trabalham com 7% (conferido em 100% dos pedidos)"),
    "amazon_105": ("AMAZON", "9", "10,5", "9% de comissão + 1,5% de taxa por pedido — a taxa da Amazon é em %, "
                                          "não em R$ (medido ao centavo em 575 pedidos)", "0"),
}


def _migrar_comissoes(p: dict) -> dict:
    """Leva a tabela antiga (comissao + taxa) para a tabela nova por canal, UMA
    vez. O que estava espalhado passa a morar numa coluna própria:
      Amazon  10,5%  → 9% de comissão + 1,5% de taxa % extra
      Magalu  taxa 5 → R$ 5,00 por PEDIDO
      Shopee  taxa 12 → R$ 12,00 por ITEM
      Madeira / Webcontinental → a taxa financeira que era parâmetro global
    Registrado em parametros["correcoes"]; se ela editar depois, o valor dela vale."""
    linhas = p.get("comissoes")
    if not linhas or "tabela_canal_v2" in (p.get("correcoes") or []):
        return p
    txg = p.get("tx_financeira") or {"madeira": 4.0, "webcont": 1.0}
    for lin in linhas:
        nome = unidecode_lower(lin.get("canal", ""))
        legado = str(lin.get(COMISSAO_LEGADO, "") or "").strip()
        for k in ("tx_fin", "pct_extra", "taxa_pedido", "taxa_item"):
            lin.setdefault(k, "")
        if "amazon" in nome:
            if str(lin.get("comissao", "")).replace(".", ",").startswith("10,5"):
                lin["comissao"], lin["pct_extra"] = "9", "1,5"
            elif not lin.get("pct_extra"):
                lin["pct_extra"] = "1,5"
        if "shopee" in nome and "xpress" not in nome:
            lin["taxa_item"] = lin.get("taxa_item") or legado or "12"
        elif legado and not lin.get("taxa_pedido"):
            lin["taxa_pedido"] = legado
        if "madeira" in nome and not lin.get("tx_fin"):
            lin["tx_fin"] = f_brl(txg.get("madeira", 4.0), 1)
        if ("webcontinental" in nome or "webcont" in nome) and not lin.get("tx_fin"):
            lin["tx_fin"] = f_brl(txg.get("webcont", 1.0), 1)
        lin.pop(COMISSAO_LEGADO, None)
    p["correcoes"] = list(p.get("correcoes") or []) + ["tabela_canal_v2"]
    _json_gravar(pasta("parametros.json"), p)
    return p


def _corrigir_comissoes(p: dict) -> dict:
    feitas = list(p.get("correcoes") or [])
    linhas = p.get("comissoes")
    if not linhas:
        return p
    mudou = False
    for chave, correcao in CORRECOES_COMISSAO.items():
        nome, errado, certo = correcao[0], correcao[1], correcao[2]
        taxa_certa = correcao[4] if len(correcao) > 4 else None
        if chave in feitas:
            continue
        for lin in linhas:
            if unidecode_lower(nome) in unidecode_lower(lin.get("canal", "")):
                atual = str(lin.get("comissao", "")).replace("%", "").replace(",", ".").strip()
                if atual in (errado, errado + ".0", errado + ".00"):
                    lin["comissao"] = certo
                    mudou = True
                if taxa_certa is not None and str(lin.get("taxa", "")).strip() not in (taxa_certa, ""):
                    # taxa que na verdade é percentual (Amazon 1,5%) não pode ficar
                    # no campo R$ — ela já está dentro do % cheio
                    lin["taxa"] = taxa_certa
                    mudou = True
        feitas.append(chave)
    if mudou or feitas != list(p.get("correcoes") or []):
        p["correcoes"] = feitas
        _json_gravar(pasta("parametros.json"), p)
    return p


def _num_br(v) -> float:
    s = str(v or "").strip().replace("R$", "").replace("%", "").replace(" ", "")
    if not s or s in ("-", "—"):
        return 0.0
    s = s.replace(".", "").replace(",", ".") if "," in s else s
    try:
        return float(s)
    except ValueError:
        return 0.0


# --------------------------------------------------------------------------
# usuários — hash werkzeug; contas iniciais criadas na 1ª subida
# --------------------------------------------------------------------------
def usuarios_caminho():
    return os.path.join(DATA_DIR, "usuarios.json")


def usuarios() -> dict:
    u = _json_ler(usuarios_caminho(), {})
    if not u:
        u = {
            "THAIS": {"papel": "admin", "senha": generate_password_hash(os.environ.get("PLUTOS_SENHA", "Plutos@2026")),
                      "trocar": True, "empresa": "multimoveis"},
            "GABI":  {"papel": "operador", "senha": generate_password_hash("Rebate@2026"),
                      "trocar": True, "empresa": "multimoveis"},
        }
        _json_gravar(usuarios_caminho(), u)
    return u


def gravar_usuarios(u):
    _json_gravar(usuarios_caminho(), u)


def logado(f):
    @wraps(f)
    def w(*a, **k):
        if not session.get("usuario"):
            return redirect(url_for("entrar", proximo=request.path))
        u = usuarios().get(session["usuario"])
        if not u:
            session.clear()
            return redirect(url_for("entrar"))
        if u.get("trocar") and request.endpoint not in ("trocar_senha", "sair", "static"):
            return redirect(url_for("trocar_senha"))
        return f(*a, **k)
    return w


def exige(perm):
    def deco(f):
        @wraps(f)
        def w(*a, **k):
            if session.get("papel") not in PODE[perm]:
                abort(403)
            return f(*a, **k)
        return w
    return deco


# --------------------------------------------------------------------------
# filtros de template
# --------------------------------------------------------------------------
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


@app.template_filter("brl")
def f_brl(v, casas=2):
    try:
        v = float(v or 0)
    except (TypeError, ValueError):
        return "—"
    s = f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s


@app.template_filter("pct")
def f_pct(v, casas=2):
    try:
        return f_brl(100 * float(v), casas) + "%"
    except (TypeError, ValueError):
        return "—"


@app.template_filter("dia")
def f_dia(s):
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return s or "—"


@app.template_filter("mesano")
def f_mesano(comp):
    try:
        a, m = comp.split("-")
        return f"{MESES[int(m) - 1]}/{a[2:]}"
    except Exception:
        return comp or "—"


@app.template_filter("quando")
def f_quando(s):
    try:
        return datetime.fromisoformat(str(s)).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return s or "—"


@app.context_processor
def contexto():
    papel = session.get("papel", "")
    return {
        "VERSAO": VERSAO, "CANAIS": canais(), "papel": papel, "PAPEIS": PAPEIS,
        "usuario": session.get("usuario"), "PODE": {k: (papel in v) for k, v in PODE.items()},
        "IRMAOS": [{"chave": c, "rotulo": r, "url": u, "arquivo": _icone_irmao(c)} for c, r, u in IRMAOS],
        "PARAM": parametros(), "COMPS": competencias(), "comp": comp_atual(), "HOJE_COMP": agora().strftime("%Y-%m"),
        "ULT": ultima_rodada_info(), "ATUALIZADO": ultima_atualizacao(), "OK_QUANDO": ok_quando_pagina(),
    }


def ultima_atualizacao():
    """A última vez que ALGUMA coisa entrou no PLUTOS (ERP, canais, Planilha 2, cadastros, comissões)."""
    if not session.get("usuario"):
        return None
    ts = []
    u = ultima_rodada_info()
    if u:
        ts.append(u["quando"])
    ts += [x["quando"] for x in erp_ler()["uploads"]]
    ts += [x["quando"] for x in mlbs_ler()["uploads"]]
    cad = descricoes_ler()
    ts += [v["quando"] for k, v in cad.items() if isinstance(v, dict) and k in ("anymarket", "custos") and v.get("quando")]
    p = parametros()
    if p.get("comissoes_quando"):
        ts.append(p["comissoes_quando"])
    return max(ts) if ts else None


def ok_quando_pagina():
    """O 'ok · data' de cada aba: quando o dado que a aba mostra entrou."""
    if not session.get("usuario") or not request:
        return None
    ep = request.endpoint or ""
    # devolve "" quando a aba TEM tabela mas ainda não recebeu nada (a pílula
    # aparece sempre: verde com a data, ou cinza "sem dados")
    if ep in ("mlbs", "custo_coletas", "coletas"):
        ups = mlbs_ler()["uploads"]
        return ups[-1]["quando"] if ups else ""
    if ep in ("canal", "linha", "faltante", "pedidos", "painel"):
        chave = request.view_args.get("chave", "meli") if request.view_args else "meli"
        r = rodada_res(chave, comp_atual())
        if r:
            return r["quando"]
        if ep in ("canal", "linha", "faltante"):
            return ""
        u = ultima_rodada_info()
        return u["quando"] if u else ""
    if ep == "parametros_tela":
        cad = descricoes_ler()
        ts = [v["quando"] for k, v in cad.items() if isinstance(v, dict) and k in ("anymarket", "custos") and v.get("quando")]
        if parametros().get("comissoes_quando"):
            ts.append(parametros()["comissoes_quando"])
        return max(ts) if ts else ""
    if ep == "arquivos":
        return ultima_atualizacao() or ""
    return None


def _icone_irmao(chave):
    for n in os.listdir(app.static_folder):
        if chave in n.lower() and n.lower().endswith((".png", ".svg", ".jpg", ".webp")) and "plutos" not in n.lower():
            return n
    return None


# --------------------------------------------------------------------------
# rodadas — uma por canal × competência
# --------------------------------------------------------------------------
def rodada_caminho(canal, comp):
    return pasta("rodadas", f"{canal}_{comp}.json")


def rodada_res_caminho(canal, comp):
    return pasta("rodadas", f"{canal}_{comp}.resumo.json")


def rodada(canal, comp) -> dict | None:
    """A rodada INTEIRA (com todas as linhas). Só use nas telas que precisam das
    linhas — Linha a linha, Por pedido, exportações."""
    return _json_ler(rodada_caminho(canal, comp), None)


def _gravar_resumo(canal, comp, r) -> dict:
    leve = {k: v for k, v in r.items() if k != "linhas"}
    leve["linhas_n"] = len(r.get("linhas") or [])
    _json_gravar(rodada_res_caminho(canal, comp), leve)
    return leve


def rodada_gravar(canal, comp, r):
    """Grava a rodada e, ao lado, o resumo leve (sem as linhas) que as telas usam."""
    _json_gravar(rodada_caminho(canal, comp), r)
    _gravar_resumo(canal, comp, r)


def _resumo_pronto(r) -> bool:
    """Rodada só conta quando o resumo está completo. Um recálculo interrompido
    (ex.: faltou memória) deixa 'resumo': {} — isso não pode derrubar as telas."""
    return bool(r) and bool((r.get("resumo") or {}).get("rebate_total") is not None)


def rodada_cab(canal, comp) -> dict | None:
    """O cabeçalho da rodada (arquivo, quem, quando, resumo) SEM as linhas.
    Serve para recalcular: aqui o resumo pode estar vazio (recálculo em curso)."""
    leve = _json_ler(rodada_res_caminho(canal, comp), None)
    if leve:
        return leve
    cheia = _json_ler(rodada_caminho(canal, comp), None)
    return _gravar_resumo(canal, comp, cheia) if cheia else None


def rodada_res(canal, comp) -> dict | None:
    """O que as telas usam: só vale com o resumo completo. Um recálculo
    interrompido (faltou memória) deixa o box como 'sem dados' em vez de quebrar."""
    cab = rodada_cab(canal, comp)
    return cab if _resumo_pronto(cab) else None


def competencias() -> list[str]:
    d = pasta("rodadas")
    comps = set(l["competencia"] for l in _json_ler(pasta("erp", "indice.json"), {"ocs": {}})["ocs"].values())
    for n in os.listdir(d):
        m = re.match(r"(\w+)_(\d{4}-\d{2})\.json$", n)
        if m:
            comps.add(m.group(2))
    comps |= {p["competencia"] for p in _json_ler(pasta("meli", "resumo_pedidos.json"), {"pedidos": {}})["pedidos"].values()}
    return sorted(comps)


def comp_atual() -> str:
    c = request.args.get("mes") if request else None
    comps = competencias()
    if c and re.match(r"^\d{4}-\d{2}$", c):
        return c
    return agora().strftime("%Y-%m")  # sempre o mês atual (Brasília)


def ultima_rodada_info():
    melhor = None
    for n in os.listdir(pasta("rodadas")):
        r = _json_ler(pasta("rodadas", n), None)
        if r and (melhor is None or r["quando"] > melhor["quando"]):
            melhor = {"quando": r["quando"], "canal": r["canal"], "comp": r["competencia"], "quem": r.get("quem")}
    return melhor


# --------------------------------------------------------------------------
# ERP · Pedidos Marketplace — um índice por OC (a chave de tudo), atualizado a
# cada upload; a OC nova substitui a antiga. Guarda também o histórico de uploads.
# --------------------------------------------------------------------------
def erp_ler() -> dict:
    return _json_ler(pasta("erp", "indice.json"), {"ocs": {}, "uploads": []})


def erp_gravar(d):
    _json_gravar(pasta("erp", "indice.json"), d)


def erp_comissao_sistema() -> dict:
    """{oc: {'pct': 0.165, 'rs': None}} — o que o motor do canal usa como comissão do sistema."""
    return {oc: {"pct": l["pct_comissao"], "rs": None} for oc, l in erp_ler()["ocs"].items()}


def linha_comissao(nomes: tuple[str, ...]) -> dict | None:
    """A linha da tabela de Parâmetros que descreve este canal."""
    for lin in parametros().get("comissoes") or COMISSAO_PADRAO:
        nome = unidecode_lower(lin.get("canal", ""))
        if any(unidecode_lower(n) in nome for n in nomes) and "full" not in nome and "xpress" not in nome:
            return lin
    return None


def parametros_canal(nomes: tuple[str, ...]) -> dict:
    """TODOS os parâmetros do canal, já em número:
       pct       = comissão + taxa % extra (é o que o canal cobra junto)
       comissao  = só o % negociado
       pct_extra = só a taxa % do canal (Amazon 1,5%)
       tx_fin    = taxa financeira em % (fica FORA do rebate)
       taxa_pedido / taxa_item = R$ fixos
       tipo      = Produto | GMV"""
    lin = linha_comissao(nomes) or {}
    com = _num_br(lin.get("comissao")) / 100 if lin.get("comissao") else 0.0
    ext = _num_br(lin.get("pct_extra")) / 100 if lin.get("pct_extra") else 0.0
    return {"comissao": com, "pct_extra": ext, "pct": round(com + ext, 6),
            "tx_fin": (_num_br(lin.get("tx_fin")) / 100 if lin.get("tx_fin") else 0.0),
            "taxa_pedido": _num_br(lin.get("taxa_pedido")),
            "taxa_item": _num_br(lin.get("taxa_item")),
            "tipo": (lin.get("tipo") or "GMV"), "linha": lin}


def comissao_cadastrada(nomes: tuple[str, ...], padrao: tuple[float, float]) -> tuple[float, float]:
    """(pct CHEIO, taxa R$ por pedido) da tabela de Parâmetros.
    pct cheio = comissão negociada + taxa % extra do canal — é o que o canal
    cobra junto e o que a comissão do sistema tem de prever."""
    lin = linha_comissao(nomes)
    if not lin:
        return padrao
    p = parametros_canal(nomes)
    if not p["pct"] and not p["taxa_pedido"]:
        return padrao
    return p["pct"], p["taxa_pedido"]


def comissao_tipo(nomes: tuple[str, ...], padrao: str = "GMV") -> str:
    """TIPO da tabela de comissões — regra oficial da casa (14/09/2026):
       Produto = comissão sobre valor produto + IPI
       GMV     = comissão sobre produto + IPI + frete (o valor total da NF)"""
    for lin in parametros().get("comissoes") or COMISSAO_PADRAO:
        nome = unidecode_lower(lin.get("canal", ""))
        if any(unidecode_lower(n) in nome for n in nomes) and "full" not in nome and "xpress" not in nome:
            t = (lin.get("tipo") or "").strip().upper()
            return "Produto" if t.startswith("PROD") else "GMV"
    return padrao


def base_do_tipo(tipo: str) -> str:
    return "produto + IPI" if tipo == "Produto" else "produto + IPI + frete = total da NF"


def erp_por_base(box: str) -> dict:
    """OC do ERP sem o sufixo -N → linha do ERP (o Magalu grava 'LU-…-1', 'LU-…-2').

    Um pedido do canal pode virar VÁRIAS OCs no ERP. Guardo a primeira (para os
    campos de texto) e SOMO produto, IPI e total de todas — é essa soma que vira
    a base da comissão do Promob (produto + IPI no Fulfillment, total na NF fora)."""
    out: dict[str, dict] = {}
    for oc, l in erp_ler()["ocs"].items():
        if l.get("box") != box:
            continue
        b = magalu.base_oc(oc)
        if b not in out:
            out[b] = dict(l, ocs_erp=0, prod_erp=0.0, ipi_erp=0.0, total_erp=0.0)
        a = out[b]
        a["ocs_erp"] += 1
        a["prod_erp"] += float(l.get("valor_prod") or 0.0)
        a["ipi_erp"] += float(l.get("ipi") or 0.0)
        a["total_erp"] += float(l.get("valor_total") or 0.0)
    return out


def erp_linhas(box: str | None = None, comp: str | None = None) -> list[dict]:
    out = list(erp_ler()["ocs"].values())
    if box:
        out = [l for l in out if l["box"] == box]
    if comp:
        out = [l for l in out if l["competencia"] == comp]
    out.sort(key=lambda l: (l["data"], l["oc"]))
    return out


def faltante_ler(canal="meli") -> dict:
    return _json_ler(pasta("manual", f"faltante_{canal}.json"), {})


def faltante_gravar(d, canal="meli"):
    _json_gravar(pasta("manual", f"faltante_{canal}.json"), d)


def mlbs_ler() -> dict:
    """Lista de MLB's do Mercado Livre (nasce da Planilha 2 · Resumo de Rebates)."""
    return _json_ler(pasta("meli", "mlbs.json"), {"mlbs": {}, "uploads": []})


def mlbs_gravar(d):
    _json_gravar(pasta("meli", "mlbs.json"), d)


def resumo_meli_ler() -> dict:
    """Pedidos da Planilha 2 (Resumo de Rebates), pedido a pedido. Daqui saem
    as telas de Coletas (frete coletas > 0 = pedido de coletas)."""
    return _json_ler(pasta("meli", "resumo_pedidos.json"), {"pedidos": {}})


def resumo_meli_gravar(d):
    _json_gravar(pasta("meli", "resumo_pedidos.json"), d)


RESUMO_CAMPOS = ("pedido_canal", "data", "competencia", "sku", "anuncio", "tipo", "qtd", "valor_prod", "frete",
                 "frete_coletas", "cupom_canal", "cupom_seller", "valor_meli", "pct_bruta", "com_bruta", "rebate_bi", "com_liq")


def _tabela_leve(f, nome: str):
    """Lê uma planilha grande sem pandas (o Anymarket tem 244 colunas × 10 mil
    linhas — em pandas estoura a memória do Render). Devolve (cabeçalho,
    iterador de linhas, função para fechar)."""
    import csv, io, tempfile
    if nome.lower().endswith((".csv", ".txt")):
        raw = f.read()
        txt = raw.decode("utf-8-sig") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("latin-1")
        sep = ";" if txt[:2000].count(";") > txt[:2000].count("\t") else "\t"
        if txt[:2000].count(",") > txt[:2000].count(sep):
            sep = ","
        rd = csv.reader(io.StringIO(txt), delimiter=sep)
        cab = next(rd)
        return cab, rd, (lambda: None)
    import openpyxl
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    f.save(tmp.name)
    wb = openpyxl.load_workbook(tmp.name, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    cab = [("" if c is None else str(c)) for c in next(it)]
    def fechar():
        wb.close()
        try:
            os.remove(tmp.name)
        except OSError:
            pass
    return cab, it, fechar


_IDX_SKU: dict = {}   # id(cadastro) → índice pela chave sem pontuação


def sku_chave(s) -> str:
    """SKU sem pontuação, em maiúsculas: o Magalu manda '0527156' e o cadastro
    tem '0527.156' — é o mesmo produto. Sem isso o item aparece 'sem peso'."""
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def cadastro_sku(sku, cad=None) -> dict:
    """A linha do cadastro daquele SKU, casando também sem a pontuação."""
    cad = cad if cad is not None else descricoes_ler()["skus"]
    t = str(sku or "").strip()
    e = cad.get(t) or cad.get(t.upper())
    if e:
        return e
    idx = _IDX_SKU.get(id(cad))
    if idx is None:
        idx = {}
        for k, v in cad.items():
            idx.setdefault(sku_chave(k), v)
        _IDX_SKU.clear()          # um cadastro por vez basta (17 mil SKUs)
        _IDX_SKU[id(cad)] = idx
    return idx.get(sku_chave(t)) or {}


def descricao_de(sku, cad=None) -> str:
    """Descrição do SKU pelo cadastro (Anymarket primeiro, CustoProduto se não houver)."""
    e = cadastro_sku(sku, cad)
    return e.get("descricao") or e.get("descricao_curta") or ""


def descricoes_ler() -> dict:
    """SKU → descrição (cadastro subido em Parâmetros). Vale para todos os canais."""
    return _json_ler(pasta("cadastro", "descricoes.json"), {"skus": {}, "quando": None, "quem": None, "arquivo": None})


# --------------------------------------------------------------------------
# BASE ACUMULADA por canal × competência: cada arquivo subido faz UPSERT por
# pedido (novo entra, igual é ignorado, diferente é atualizado). Nunca soma
# duas vezes, nunca sobrescreve o que outro arquivo trouxe. A rodada do canal
# é recalculada a partir da base, não do último arquivo.
# --------------------------------------------------------------------------
def base_ler(canal: str, comp: str) -> dict:
    return _json_ler(pasta("base", f"{canal}_{comp}.json"), {"pedidos": {}, "uploads": []})


def base_gravar(canal: str, comp: str, d: dict):
    _json_gravar(pasta("base", f"{canal}_{comp}.json"), d)


def _linha_serial(row: dict) -> dict:
    out = {}
    for k, v in row.items():
        if hasattr(v, "isoformat"):
            v = v.isoformat()
        elif hasattr(v, "item"):  # numpy
            v = v.item()
        if isinstance(v, float) and v != v:
            v = None
        out[k] = v
    return out


def base_upsert(canal: str, df, chave: str, nome: str, caminho: str, quem: str) -> dict:
    """Aplica o df (já normalizado, com coluna 'competencia') na base de cada
    competência. Devolve {comp: {novas, atualizadas, iguais, antigas, total}}.
    Um arquivo mais ANTIGO (período termina antes) nunca sobrescreve o que um
    arquivo mais novo já trouxe — só acrescenta pedidos que faltavam."""
    res = {}
    snap = str(df["data"].max()) if "data" in df and len(df) else ""
    for comp, sub in df.groupby("competencia"):
        b = base_ler(canal, comp)
        novas = atualizadas = iguais = antigas = 0
        for row in sub.to_dict("records"):
            k = str(row.get(chave) or "").strip()
            if not k:
                continue
            row = _linha_serial(row)
            row["_snap"] = snap
            atual = b["pedidos"].get(k)
            if atual is None:
                novas += 1
                b["pedidos"][k] = row
            elif {x: v for x, v in atual.items() if x != "_snap"} == {x: v for x, v in row.items() if x != "_snap"}:
                iguais += 1
            elif (atual.get("_snap") or "") > snap:
                antigas += 1  # a base já tem uma foto mais nova deste pedido
            else:
                atualizadas += 1
                b["pedidos"][k] = row
        b["uploads"].append({"nome": nome, "caminho": caminho, "quando": agora().isoformat(), "quem": quem,
                             "linhas": int(len(sub)), "novas": novas, "atualizadas": atualizadas, "iguais": iguais, "antigas": antigas})
        b["uploads"] = b["uploads"][-50:]
        base_gravar(canal, comp, b)
        res[comp] = {"novas": novas, "atualizadas": atualizadas, "iguais": iguais, "antigas": antigas, "total": len(b["pedidos"])}
    return res


def base_marcar_sumidos(canal: str, df, chave: str, nome: str) -> dict:
    """Pedido que ESTAVA na base e NÃO vem mais no relatório completo = cancelado
    pelo canal (o Mercado Livre simplesmente tira a linha do relatório).
    A varredura é fechada na janela do arquivo (menor e maior data) e nas CONTAS
    que aparecem nele — assim um export parcial, de uma conta só ou de meio mês
    nunca cancela pedido que ele nem deveria trazer. Se o pedido voltar a
    aparecer, a marca é desfeita. Devolve {comp: {sumiram, voltaram}}."""
    if df is None or not len(df) or "data" not in df:
        return {}
    de, ate = str(df["data"].min()), str(df["data"].max())
    contas = {str(c).strip() for c in df["conta"]} if "conta" in df else set()
    presentes = {str(k).strip() for k in df[chave]}
    quando = agora().isoformat()
    res = {}
    for comp in sorted({str(c) for c in df["competencia"]}):
        b = base_ler(canal, comp)
        sumiram = voltaram = 0
        for k, row in b["pedidos"].items():
            d = str(row.get("data") or "")
            dentro = bool(d) and de <= d <= ate and (not contas or str(row.get("conta") or "").strip() in contas)
            if k in presentes:
                if row.pop("sumiu", None):
                    row.pop("sumiu_em", None)
                    row.pop("sumiu_arquivo", None)
                    voltaram += 1
                continue
            if not dentro or row.get("sumiu"):
                continue
            if (row.get("_snap") or "") > (str(df["data"].max()) or ""):
                continue  # a base tem foto mais nova deste pedido: não é sumiço
            row["sumiu"] = True
            row["sumiu_em"] = quando
            row["sumiu_arquivo"] = nome
            sumiram += 1
        if sumiram or voltaram:
            base_gravar(canal, comp, b)
            res[comp] = {"sumiram": sumiram, "voltaram": voltaram}
    return res


def base_sumidos(canal: str, comp: str) -> dict:
    """{pedido: {quando, arquivo}} dos pedidos que sumiram do relatório."""
    return {k: {"quando": v.get("sumiu_em", ""), "arquivo": v.get("sumiu_arquivo", ""),
                "data": v.get("data", ""), "conta": v.get("conta", ""),
                "valor_prod": v.get("valor_prod", 0)}
            for k, v in base_ler(canal, comp)["pedidos"].items() if v.get("sumiu")}


def base_df(canal: str, comp: str):
    """DataFrame da base acumulada (data volta a ser date). Pedidos marcados como
    'sumiu' (cancelados pelo canal) ficam de fora — não geram rebate nem coleta."""
    import pandas as pd
    b = base_ler(canal, comp)
    vivos = [v for v in b["pedidos"].values() if not v.get("sumiu")]
    if not vivos:
        return None
    df = pd.DataFrame(vivos)
    for c in ("_snap", "sumiu", "sumiu_em", "sumiu_arquivo"):
        if c in df:
            df = df.drop(columns=[c])
    if "data" in df:
        df["data"] = pd.to_datetime(df["data"]).dt.date
    return df


def _txt_upsert(res: dict) -> str:
    return " · ".join(f"{f_mesano(c)}: {v['novas']} novas, {v['atualizadas']} atualizadas, {v['iguais']} iguais"
                      + (f", {v['antigas']} mantidas (arquivo mais antigo)" if v.get("antigas") else "") + f" → base {v['total']} pedidos"
                      for c, v in res.items())


def recalcular_meli(comp: str):
    """Recalcula a rodada do Meli daquela competência a partir dos arquivos
    guardados (usado quando a tabela manual muda ou o ADC002 chega)."""
    r = rodada_cab("meli", comp)   # só o cabeçalho: as linhas são refeitas abaixo
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    df = base_df("meli", comp)
    if df is None:
        df, diag = meli.ler_tabela_geral(r["arquivo"]["caminho"])
        df = df[df["competencia"] == comp]
    sis = erp_comissao_sistema()
    sis_diag = {"linhas": len(sis), "modo": "ERP · Pedidos Marketplace (coluna AB)"} if sis else None
    if not sis and r.get("sistema") and os.path.exists(r["sistema"]["caminho"]):
        sis, sis_diag = meli.ler_comissao_sistema(r["sistema"]["caminho"])
    linhas = meli.calcular(df, sis, faltante_ler("meli"), parametros()["tolerancia_comissao"])
    r["linhas"] = linhas
    r["sistema"] = ({"nome": "ERP · Pedidos Marketplace", "quando": erp_ler()["uploads"][-1]["quando"],
                     "diag": sis_diag} if sis and erp_ler()["uploads"] else r.get("sistema"))
    r["resumo"] = meli.resumo(linhas)
    r["sistema_diag"] = sis_diag
    r["recalculado"] = agora().isoformat()
    rodada_gravar("meli", comp, r)
    return r


def recalcular_magalu(comp: str):
    r = rodada_cab("magalu", comp)   # só o cabeçalho: as linhas são refeitas abaixo
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    df = base_df("magalu", comp)
    if df is None:
        df, diag = magalu.ler(r["arquivo"]["caminho"])
        df = df[df["competencia"] == comp]
    pct, taxa = comissao_cadastrada(("MAGAZINE", "MAGALU"), (0.11, 5.0))
    tipo = comissao_tipo(("MAGAZINE", "MAGALU"))
    linhas = magalu.calcular(df, pct, taxa, erp_por_base("magalu"), tipo)
    r["linhas"] = linhas
    r["resumo"] = magalu.resumo(linhas, int(df["cancelado"].sum()))
    sem_erp = sum(1 for l in linhas if not l.get("erp_ok"))
    r["sistema"] = {"nome": (f"% do pedido no ERP + R$ {taxa:.2f}/pedido · {tipo} ({base_do_tipo(tipo)}) · Fulfillment sobre produto + IPI"
                             + (f" · {sem_erp} sem par no ERP usam Parâmetros {pct * 100:.2f}%" if sem_erp else "")),
                    "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "comissão do ERP por pedido", "sem_erp": sem_erp}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    rodada_gravar("magalu", comp, r)
    return r


def vendas_magalu_ler() -> dict:
    """Planilha 2 do Magalu (Vendas no período): pacotes (frete Full) e itens
    (coparticipação de frete). Daqui saem as telas Custo Frete Full e
    Coparticipação de Frete."""
    return _json_ler(pasta("magalu", "vendas.json"), {"pacotes": {}, "itens": {}, "uploads": []})


def vendas_magalu_gravar(d):
    _json_gravar(pasta("magalu", "vendas.json"), d)


def processar_magalu_vendas(destino: str, nome: str, quem: str) -> str:
    """Planilha 2 · Vendas no período. O zip do portal traz dois csv; cada um
    entra por aqui e atualiza o seu lado (pacotes ou itens), sem apagar o outro."""
    qual = magalu_vendas.que_arquivo(destino)
    if not qual:
        raise ValueError("Não parece a Planilha 2 do Magalu. Suba o zip 'MagazineLuiza_Vendas_…' "
                         "(ele traz relatorio_vendas_pedidos e relatorio_vendas_pacotes).")
    d = vendas_magalu_ler()
    if qual == "pacotes":
        linhas, diag = magalu_vendas.ler_pacotes(destino)
        for p in linhas:
            d["pacotes"][p["pacote"] or p["pedido"]] = p
        txt = (f"Planilha 2 · PACOTES lida — {diag['linhas']} pacotes de {f_dia(diag['de'])} a {f_dia(diag['ate'])} · "
               f"{diag['full']} Full · {diag['cancelados']} cancelados fora · "
               f"custo do frete Full R$ {diag['frete_full']:,.2f}".replace(",", "@").replace(".", ",").replace("@", "."))
    else:
        linhas, diag = magalu_vendas.ler_pedidos(destino)
        vistos: dict[str, int] = {}
        for i in linhas:
            k = f"{i['pedido']}|{i['sku']}"
            vistos[k] = vistos.get(k, 0) + 1
            d["itens"][f"{k}|{vistos[k]}"] = i
        txt = (f"Planilha 2 · ITENS lida — {diag['linhas']} itens em {diag['pedidos']} pedidos de "
               f"{f_dia(diag['de'])} a {f_dia(diag['ate'])} · {diag['com_copart']} com coparticipação de frete · "
               f"R$ {diag['copart']:,.2f}".replace(",", "@").replace(".", ",").replace("@", "."))
    d["uploads"].append({"nome": nome, "caminho": destino, "quando": agora().isoformat(), "quem": quem,
                         "qual": qual, "diag": diag})
    d["uploads"] = d["uploads"][-30:]
    vendas_magalu_gravar(d)
    return txt


def full_magalu_ler() -> dict:
    """Cobranças do Fulfillment do Magalu (manuseio, armazenagem, tempo de
    estoque e coleta), uma entrada por cobrança."""
    return _json_ler(pasta("magalu", "full.json"), {"cobrancas": {}, "uploads": []})


def full_magalu_gravar(d):
    _json_gravar(pasta("magalu", "full.json"), d)


def processar_magalu_full(destino: str, nome: str, quem: str) -> str:
    """Cobranças do Fulfillment. Cada csv do zip entra por aqui e atualiza só o
    seu tipo — manuseio, armazenagem, tempo de estoque ou coleta."""
    qual, linhas, diag = magalu_full.ler(destino)
    d = full_magalu_ler()
    # REGRA DA THAÍS (16/09/2026): a cobrança do Full pertence ao MÊS EM QUE É
    # SUBIDA — este arquivo é de setembro, o próximo será de outubro, e assim
    # vai. O período da cobrança pode atravessar meses (a coleta é de agosto);
    # quem manda é a competência aberta na hora de subir.
    comp_ = comp_atual()
    # PARCIAL: o relatório do portal é ACUMULADO (do início do período até a
    # data da extração) e as linhas não têm id próprio. Então cada upload
    # SUBSTITUI o bloco daquele tipo na competência — subir o parcial do dia 10
    # e depois o fechado do dia 30 não duplica nada. Por isso: sempre exportar
    # desde o início do período, nunca só o pedaço novo.
    antes = sum(1 for c in d["cobrancas"].values()
                if c.get("comp") == comp_ and c.get("tipo") == qual)
    d["cobrancas"] = {k: c for k, c in d["cobrancas"].items()
                      if not (c.get("comp") == comp_ and c.get("tipo") == qual)}
    for c in linhas:
        c["comp"] = comp_
        d["cobrancas"][f"{comp_}|{c['chave']}"] = c
    d["uploads"].append({"nome": nome, "caminho": destino, "quando": agora().isoformat(), "quem": quem,
                         "qual": qual, "diag": diag})
    d["uploads"] = d["uploads"][-40:]
    full_magalu_gravar(d)
    per = f" de {f_dia(diag['de'])} a {f_dia(diag['ate'])}" if diag["de"] else ""
    troca = f" (substituiu as {antes} anteriores desta competência)" if antes else ""
    return (f"Fulfillment · {magalu_full.TIPOS[qual]} lido em {f_mesano(comp_)} — {diag['linhas']} cobranças{per} · "
            f"{diag['skus']} SKUs · R$ {diag['valor']:,.2f}{troca}".replace(",", "@").replace(".", ",").replace("@", "."))


def prazo_dias(v) -> int:
    """Prazo em dias a partir do cadastro: {manuseio, transferencia} ou um número."""
    if isinstance(v, dict):
        return int(v.get("manuseio") or 0) + int(v.get("transferencia") or 0)
    try:
        return int(v or 0)
    except (TypeError, ValueError):
        return 0


def cubagem_por_sku() -> dict:
    """SKU → cubagem (m³) do cadastro de SKUs. A coleta do Magalu é cobrada por
    m³, então é a cubagem que reparte esse custo entre os produtos."""
    cad = descricoes_ler()["skus"]
    out = {}
    for k, v in cad.items():
        c = (v or {}).get("cubagem")
        if c:
            out[k] = float(c)
            out[sku_chave(k)] = float(c)
    return out


def vendidos_full_sku(comp: str) -> dict:
    """Unidades vendidas no Fulfillment por SKU (Planilha 2 · Vendas), sem
    cancelados. Serve de divisor quando o SKU não teve manuseio no ciclo."""
    v = vendas_magalu_ler()
    canc, ff = set(), set()
    for p in v["pacotes"].values():
        if p.get("cancelado"):
            canc.add(p["pedido"])
        elif p.get("full"):
            ff.add(p["pedido"])
    out: dict[str, float] = {}
    for i in v["itens"].values():
        if i.get("competencia") != comp or i["pedido"] in canc or i["pedido"] not in ff:
            continue
        out[i.get("sku") or "—"] = out.get(i.get("sku") or "—", 0.0) + (i.get("qtd") or 0.0)
    return out


def copart_por_sku(comp: str) -> dict:
    """Coparticipação de frete por SKU da competência (Planilha 2 · Vendas)."""
    v = vendas_magalu_ler()
    canc = {p["pedido"] for p in v["pacotes"].values() if p.get("cancelado")}
    out: dict[str, float] = {}
    for i in v["itens"].values():
        if i.get("competencia") != comp or i["pedido"] in canc or not i.get("copart"):
            continue
        out[i.get("sku") or "—"] = round(out.get(i.get("sku") or "—", 0.0) + i["copart"], 2)
    return out


def processar_magalu(destino: str, nome: str, quem: str) -> str:
    df, diag = magalu.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    fora = 0
    for comp, n in list(comps.items()):
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            df = df[df["competencia"] != comp]
    res = base_upsert("magalu", df, "pedido", nome, destino, quem)
    feitos = []
    for comp in res:
        r = dict(rodada_cab("magalu", comp) or {"canal": "magalu", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("magalu", comp, r)
        r = recalcular_magalu(comp)
        if not _resumo_pronto(r):
            continue
        feitos.append((comp, r["resumo"]["pedidos"], r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    return f"Magazine Luiza lido — {_txt_upsert(res)}. {txt}. Cancelados fora: {diag['cancelados']}.{extra}"


def recalcular_shopee(comp: str):
    r = rodada_cab("shopee", comp)   # só o cabeçalho: as linhas são refeitas abaixo
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    df = base_df("shopee", comp)
    if df is None:
        df, diag = shopee.ler(r["arquivo"]["caminho"])
        df = df[df["competencia"] == comp]
    # as regras de exclusão valem também sobre o que já está na base (cancelado,
    # não pago, devolução aprovada) — assim "▶ Rodar" corrige sem subir o arquivo de novo
    df = shopee.marcar_fora(df)
    # na Shopee a taxa fixa é por ITEM (R$ 12,00) e já vem dentro da Taxa de
    # serviço bruta do relatório — o motor a recebe só para mostrar, não soma
    pct = comissao_cadastrada(("SHOPEE",), (0.12, 0.0))[0]
    taxa = parametros_canal(("SHOPEE",))["taxa_item"] or 12.0
    linhas = shopee.calcular(df, pct, taxa, erp_ler()["ocs"])
    r["linhas"] = linhas
    r["resumo"] = shopee.resumo(linhas, int(df["cancelado"].sum()))
    tipo = comissao_tipo(("SHOPEE",), "Produto")
    r["sistema"] = {"nome": f"Parâmetros · {pct * 100:.2f}% do {tipo} ({base_do_tipo(tipo)})",
                    "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    rodada_gravar("shopee", comp, r)
    return r


def processar_shopee(destino: str, nome: str, quem: str) -> str:
    df, diag = shopee.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    fora = 0
    for comp, n in list(comps.items()):
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            df = df[df["competencia"] != comp]
    res = base_upsert("shopee", df, "pedido", nome, destino, quem)
    feitos = []
    for comp in res:
        r = dict(rodada_cab("shopee", comp) or {"canal": "shopee", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("shopee", comp, r)
        r = recalcular_shopee(comp)
        if not _resumo_pronto(r):
            continue
        feitos.append((comp, r["resumo"]["pedidos"], r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    fora = f"Fora: {diag['cancelados']} (cancelados, {diag.get('nao_pagos', 0)} não pagos, {diag.get('devolvidos', 0)} com devolução aprovada)"
    return f"Shopee lida — {_txt_upsert(res)}. {txt}. {fora} · {diag['itens']} itens → {diag['linhas']} pedidos.{extra}"


def recalcular_madeira(comp: str):
    r = rodada_cab("madeira", comp)   # só o cabeçalho: as linhas são refeitas abaixo
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    df = base_df("madeira", comp)
    if df is None:
        df, diag = madeira.ler(r["arquivo"]["caminho"])
        df = df[df["competencia"] == comp]
    df = madeira.marcar_fora(df)   # as exclusões valem também sobre a base já gravada
    pct, _ = comissao_cadastrada(("MADEIRA",), (0.17, 0.0))
    linhas = madeira.calcular(df, pct, erp_ler()["ocs"])
    r["linhas"] = linhas
    r["resumo"] = madeira.resumo(linhas, int(df["cancelado"].sum()))
    tipo = comissao_tipo(("MADEIRA",))
    r["sistema"] = {"nome": f"Parâmetros · {pct * 100:.2f}% do {tipo} ({base_do_tipo(tipo)})", "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    rodada_gravar("madeira", comp, r)
    return r


def processar_madeira(destino: str, nome: str, quem: str) -> str:
    df, diag = madeira.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    fora = 0
    for comp, n in list(comps.items()):
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            df = df[df["competencia"] != comp]
    res = base_upsert("madeira", df, "pedido", nome, destino, quem)
    feitos = []
    for comp in res:
        r = dict(rodada_cab("madeira", comp) or {"canal": "madeira", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("madeira", comp, r)
        r = recalcular_madeira(comp)
        if not _resumo_pronto(r):
            continue
        feitos.append((comp, r["resumo"]["pedidos"], r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    return f"Madeira Madeira lido — {_txt_upsert(res)}. {txt}. Fora: {diag['cancelados']} ({diag['cancelados'] - diag.get('novos', 0)} cancelados + {diag.get('novos', 0)} novos) · {diag['itens']} itens → {diag['linhas']} pedidos.{extra}"


def recalcular_webcont(comp: str):
    r = rodada_cab("webcont", comp)   # só o cabeçalho: as linhas são refeitas abaixo
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    df = base_df("webcont", comp)
    if df is None:
        df, diag = webcont.ler(r["arquivo"]["caminho"])
        df = df[df["competencia"] == comp]
    pct, _ = comissao_cadastrada(("WEBCONTINENTAL", "WEBCONT"), (0.19, 0.0))
    linhas = webcont.calcular(df, pct, erp_ler()["ocs"])
    r["linhas"] = linhas
    r["resumo"] = webcont.resumo(linhas, int(df["cancelado"].sum()))
    tipo = comissao_tipo(("WEBCONTINENTAL", "WEBCONT"))
    r["sistema"] = {"nome": f"Parâmetros · {pct * 100:.2f}% do {tipo} ({base_do_tipo(tipo)})", "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    rodada_gravar("webcont", comp, r)
    return r


def processar_webcont(destino: str, nome: str, quem: str) -> str:
    df, diag = webcont.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    fora = 0
    for comp, n in list(comps.items()):
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            df = df[df["competencia"] != comp]
    res = base_upsert("webcont", df, "pedido", nome, destino, quem)
    feitos = []
    for comp in res:
        r = dict(rodada_cab("webcont", comp) or {"canal": "webcont", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("webcont", comp, r)
        r = recalcular_webcont(comp)
        if not _resumo_pronto(r):
            continue
        feitos.append((comp, r["resumo"]["pedidos"], r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    return f"Webcontinental lida — {_txt_upsert(res)}. {txt}. Cancelados fora: {diag['cancelados']}.{extra}"


def recalcular_colombo(comp: str):
    r = rodada_cab("colombo", comp)   # só o cabeçalho: as linhas são refeitas abaixo
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    df = base_df("colombo", comp)
    if df is None:
        df, diag = colombo.ler(r["arquivo"]["caminho"])
        df = df[df["competencia"] == comp]
    pct, _ = comissao_cadastrada(("COLOMBO",), (0.07, 0.0))
    linhas = colombo.calcular(df, pct, erp_ler()["ocs"])
    r["linhas"] = linhas
    r["resumo"] = colombo.resumo(linhas, int(df["cancelado"].sum()))
    tipo = comissao_tipo(("COLOMBO",))
    r["sistema"] = {"nome": f"Parâmetros · {pct * 100:.2f}% do {tipo} ({base_do_tipo(tipo)})", "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    rodada_gravar("colombo", comp, r)
    return r


def processar_colombo(destino: str, nome: str, quem: str) -> str:
    df, diag = colombo.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    fora = 0
    for comp, n in list(comps.items()):
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            df = df[df["competencia"] != comp]
    res = base_upsert("colombo", df, "pedido", nome, destino, quem)
    feitos = []
    for comp in res:
        r = dict(rodada_cab("colombo", comp) or {"canal": "colombo", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("colombo", comp, r)
        r = recalcular_colombo(comp)
        if not _resumo_pronto(r):
            continue
        feitos.append((comp, r["resumo"]["pedidos"], r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    dup = f" {diag['itens_duplicados']} linha(s) duplicada(s) do export ignorada(s)." if diag["itens_duplicados"] else ""
    pct, _ = comissao_cadastrada(("COLOMBO",), (0.07, 0.0))
    alerta = ("" if abs(pct - 0.07) < 0.0005 else
              f" ATENÇÃO: a tabela de Parâmetros está com {pct * 100:.2f}% para o Colombo, mas o portal e o ERP "
              f"trabalham com 7% — corrija em Parâmetros ou o rebate sai inflado.")
    return (f"Colombo lido — {_txt_upsert(res)}. {txt}. Fora (Cancelado/Incluído): {diag['cancelados']} · "
            f"{diag['itens']} itens → {diag['linhas']} pedidos ({diag['multi_item']} com mais de 1 item).{dup}{extra}{alerta}")


# --------------------------------------------------------------------------
# AMAZON — o relatório é de TRANSAÇÕES, não de pedidos. A base acumula
# transação a transação (chave própria) e o PEDIDO é montado a partir delas,
# para que o mesmo ID do pedido entre UMA VEZ SÓ, venha em quantos arquivos vier.
# --------------------------------------------------------------------------
def amazon_tx_todas():
    """Todas as transações da Amazon, de todas as competências, num df só."""
    import pandas as pd
    partes = []
    for cp in competencias():
        d = base_df("amazon", cp)
        if d is not None and len(d):
            partes.append(d)
    if not partes:
        return None
    df = pd.concat(partes, ignore_index=True)
    return df.drop_duplicates(subset=["chave"], keep="last")


def amazon_extra() -> dict:
    """Publicidade, serviços e reembolsos somados de TODAS as transações da base
    (não só do último arquivo) — por competência."""
    tx = amazon_tx_todas()
    out: dict[str, dict] = {}
    if tx is None:
        return out
    for cp, sub in tx.groupby("competencia"):
        serv = sub[sub["t"] == amazon.T_SERV]
        reem = sub[sub["t"].str.startswith(amazon.T_REEMB)]
        pub = serv[serv["produto"].astype(str).str.lower().str.contains("public")]
        out[str(cp)] = {
            "servicos": round(-float(serv["repasse"].sum()), 2),
            "publicidade": round(-float(pub["repasse"].sum()), 2),
            "servicos_dias": int(serv["data"].nunique()),
            "reembolsos": int(len(reem)),
            "reembolso_rs": round(-float(reem["repasse"].sum()), 2),
        }
    return out


def recalcular_amazon(comp: str):
    r = rodada_cab("amazon", comp)
    if not r:
        return None
    r = dict(r)
    r.pop("linhas_n", None)
    tx = amazon_tx_todas()
    if tx is None:
        tx, _d = amazon.ler(r["arquivo"]["caminho"])
    ped = amazon.pedidos(tx)
    if ped is None or not len(ped):
        return None
    ped = ped[ped["competencia"] == comp]
    if not len(ped):
        return None
    pct, taxa_rs = comissao_cadastrada(("AMAZON",), (0.105, 0.0))
    if taxa_rs:   # taxa em R$ na Amazon é erro de cadastro: o 1,5% já está no %
        pct = pct
    linhas = amazon.calcular(ped, pct, amazon.TAXA_PADRAO, erp_por_base("amazon"),
                             parametros()["tolerancia_comissao"])
    extra = dict(amazon_extra().get(comp) or {})
    extra["devolvidos"] = int(ped["devolvido"].sum())
    r["linhas"] = linhas
    r["resumo"] = amazon.resumo(linhas, extra)
    sem_faixa = r["resumo"]["sem_faixa"]
    r["sistema"] = {"nome": (f"Faixa da categoria medida no próprio relatório (negociada + {amazon.TAXA_PADRAO * 100:.1f}% de taxa)"
                             + (f" · {sem_faixa} pedido(s) fora de faixa usam o cadastro {pct * 100:.2f}%" if sem_faixa else "")),
                    "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "faixa por categoria × (produto − desconto + frete)",
                             "cadastro_pct": round(pct * 100, 2), "taxa_rs_cadastrada": taxa_rs,
                             "sem_faixa": sem_faixa}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    rodada_gravar("amazon", comp, r)
    return r


def processar_amazon(destino: str, nome: str, quem: str) -> str:
    tx, diag = amazon.ler(destino)
    res = base_upsert("amazon", tx, "chave", nome, destino, quem)
    # o pedido pode ter transações em meses diferentes (pagamento em um,
    # reembolso no outro): recalcula toda competência que a base conhece
    comps = sorted(set(res) | {c for c in competencias() if rodada_cab("amazon", c)})
    feitos = []
    for comp in comps:
        r = dict(rodada_cab("amazon", comp) or {"canal": "amazon", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("amazon", comp, r)
        r = recalcular_amazon(comp)
        if not _resumo_pronto(r):
            continue
        s = r["resumo"]
        feitos.append((comp, s["pedidos"], s["rebate_total"], s["desvio_cadastro"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, rebate R$ {f_brl(t)}" for c, n, t, _ in feitos)
    desv = sum(d for *_x, d in feitos)
    pct, taxa_rs = comissao_cadastrada(("AMAZON",), (0.105, 0.0))
    alerta = ""
    if taxa_rs:
        alerta = (f" ATENÇÃO: o cadastro está com R$ {f_brl(taxa_rs)} de taxa extra para a Amazon. Lá a taxa é de "
                  f"1,5% (percentual), e ela já está dentro do % cheio — zere a taxa em R$ em Parâmetros.")
    if abs(pct - 0.105) > 0.0005:
        alerta += (f" ATENÇÃO: Parâmetros está com {pct * 100:.2f}% para a Amazon; o negociado é 9% + 1,5% de taxa = 10,5%.")
    dev = f" {diag['reembolsos']} reembolso(s) (R$ {f_brl(diag['reembolso_rs'])}) — pedidos devolvidos saem do rebate." if diag["reembolsos"] else ""
    pub = f" Publicidade no período do arquivo: R$ {f_brl(diag['publicidade'])} em {diag['servicos_dias']} dia(s) — é custo de mídia, não entra no rebate." if diag["publicidade"] else ""
    dsv = f" Desvio de cadastro (não é rebate): R$ {f_brl(desv)}." if abs(desv) > 0.5 else ""
    return (f"Amazon lida — {diag['linhas']} transações ({diag['pagamentos']} pagamentos de "
            f"{diag['pedidos_no_arquivo']} pedidos) de {f_dia(diag['de'])} a {f_dia(diag['ate'])}. "
            f"{_txt_upsert(res)}. {txt}.{dsv}{dev}{pub}{alerta}")


def recalcular_canais(comp: str) -> list[str]:
    """Recalcula um canal de cada vez e SOLTA a memória entre eles (o Render tem
    512 MB: dois canais grandes juntos na memória derrubavam o app)."""
    feitos = []
    r = recalcular_meli(comp)
    if r:
        feitos.append(f"Mercado Livre {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    r = recalcular_magalu(comp)
    if r:
        feitos.append(f"Magazine Luiza {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    r = recalcular_shopee(comp)
    if r:
        feitos.append(f"Shopee {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    r = recalcular_madeira(comp)
    if r:
        feitos.append(f"Madeira Madeira {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    r = recalcular_webcont(comp)
    if r:
        feitos.append(f"Webcontinental {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    r = recalcular_colombo(comp)
    if r:
        feitos.append(f"Colombo {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    r = recalcular_amazon(comp)
    if r:
        feitos.append(f"Amazon {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = None
    gc.collect()
    return feitos


# --------------------------------------------------------------------------
# entrada / conta
# --------------------------------------------------------------------------
@app.route("/entrar", methods=["GET", "POST"])
def entrar():
    if request.method == "POST":
        login = (request.form.get("usuario") or "").strip().upper()
        senha = request.form.get("senha") or ""
        u = usuarios().get(login)
        if u and u.get("ativo") is False:
            flash("Este usuário está desativado. Fale com a direção.")
            return render_template("login.html")
        ok = u and check_password_hash(u["senha"], senha)
        if not ok and login == "THAIS" and senha == os.environ.get("PLUTOS_SENHA", "Plutos@2026"):
            ok = True  # porta de recuperação da direção
        if ok:
            session.clear()
            session["usuario"] = login
            session["papel"] = u["papel"] if u else "admin"
            session["empresa"] = (u or {}).get("empresa", "multimoveis")
            return redirect(url_for("painel"))  # sempre abre no GERAL, mês atual
        flash("Usuário ou senha não conferem.")
    return render_template("login.html")


@app.route("/sair")
def sair():
    session.clear()
    return redirect(url_for("entrar"))


@app.route("/trocar-senha", methods=["GET", "POST"])
def trocar_senha():
    if not session.get("usuario"):
        return redirect(url_for("entrar"))
    if request.method == "POST":
        s1, s2 = request.form.get("s1") or "", request.form.get("s2") or ""
        if len(s1) < 6 or s1 != s2:
            flash("A senha precisa ter 6+ caracteres e as duas iguais.")
        else:
            u = usuarios()
            u[session["usuario"]]["senha"] = generate_password_hash(s1)
            u[session["usuario"]]["trocar"] = False
            gravar_usuarios(u)
            flash("Senha trocada.")
            return redirect(url_for("painel"))
    return render_template("trocar_senha.html")


@app.route("/conta", methods=["GET", "POST"])
@logado
def conta():
    u = usuarios()
    eu = session["usuario"]
    if request.method == "POST":
        acao = request.form.get("acao")
        if acao == "minha_senha":
            atual = request.form.get("atual") or ""
            s1, s2 = request.form.get("s1") or "", request.form.get("s2") or ""
            if not check_password_hash(u[eu]["senha"], atual):
                flash("A senha atual não confere.")
            elif len(s1) < 6 or s1 != s2:
                flash("A nova senha precisa ter 6+ caracteres e as duas iguais.")
            else:
                u[eu]["senha"] = generate_password_hash(s1); u[eu]["trocar"] = False
                gravar_usuarios(u); flash("Senha trocada.")
            return redirect(url_for("conta"))
        if session.get("papel") != "admin":
            abort(403)
        alvo = (request.form.get("login") or "").strip().upper()
        senha_ini = None
        if acao == "novo" and alvo:
            if alvo in u:
                flash(f"{alvo} já existe.")
            else:
                senha_ini = secrets.token_urlsafe(6)
                u[alvo] = {"papel": request.form.get("papel", "equipe"), "senha": generate_password_hash(senha_ini),
                           "trocar": True, "empresa": empresa_atual(), "ativo": True}
                flash(f"{alvo} criado. Senha inicial: {senha_ini} — aparece só agora; quem recebe troca no primeiro acesso.")
        elif acao == "reset" and alvo in u:
            senha_ini = secrets.token_urlsafe(6)
            u[alvo]["senha"] = generate_password_hash(senha_ini); u[alvo]["trocar"] = True
            flash(f"Senha de {alvo} resetada. Senha inicial: {senha_ini} — aparece só agora.")
        elif acao == "papel" and alvo in u and alvo != eu:
            u[alvo]["papel"] = request.form.get("papel", u[alvo]["papel"])
        elif acao == "desativar" and alvo in u and alvo != eu:
            u[alvo]["ativo"] = not (u[alvo].get("ativo", True))
        elif acao == "excluir" and alvo in u and alvo != eu:
            del u[alvo]
        gravar_usuarios(u)
        return redirect(url_for("conta"))
    return render_template("conta.html", equipe=u, PAPEL_DESC=PAPEL_DESC)


# --------------------------------------------------------------------------
# telas
# --------------------------------------------------------------------------
@app.route("/")
@logado
def painel():
    comp = comp_atual()
    por_canal = []
    tot = {"rs": 0.0, "com": 0.0, "frete": 0.0, "total": 0.0, "venda": 0.0, "pedidos": 0, "pend": 0}
    for c in canais():
        r = rodada_res(c["chave"], comp) if c["ativo"] else None
        if _resumo_pronto(r):
            s = r["resumo"]
            por_canal.append({**c, "res": s, "quando": r["quando"]})
            v = lambda k: (s.get(k) or 0.0)  # noqa: E731 — resumo de versão antiga pode não ter tudo
            tot["rs"] += v("rebate_rs"); tot["com"] += v("rebate_comissao"); tot["frete"] += v("rebate_frete")
            tot["total"] += v("rebate_total"); tot["venda"] += v("venda"); tot["pedidos"] += v("pedidos")
            tot["pend"] += v("faltante_pendentes")
        else:
            por_canal.append({**c, "res": None})
    tot["pct"] = (100 * tot["total"] / tot["venda"]) if tot["venda"] else 0
    erp_res = erp.resumo_por_box([l for l in erp_ler()["ocs"].values() if l["competencia"] == comp])
    for c in por_canal:
        c["erp"] = erp_res.get(c["chave"])
    # aviso de cadastro: SKU do Full sem peso não tem R$/kg e não dá para
    # comparar Full × Coletas × transportadora própria
    try:
        sem_peso = [m["sku"] for m in _custo_full(comp) if not m.get("peso")]
    except Exception:  # noqa: BLE001
        sem_peso = []
    return render_template("painel.html", por_canal=por_canal, tot=tot, erp_outros=erp_res.get("outros"),
                           sem_peso=sem_peso)


@app.route("/canal/<chave>")
@logado
def canal(chave):
    c = canal_por_chave().get(chave) or abort(404)
    comp = comp_atual()
    r = rodada_res(chave, comp) if c["ativo"] else None
    if not _resumo_pronto(r):
        r = None
    if r and chave == "meli" and "com_sistema" not in (r.get("resumo") or {}):
        # rodada gravada por versão anterior: completa o resumo sem exigir rodar de novo
        cheia = rodada("meli", comp)
        if cheia:
            cheia["resumo"] = meli.resumo(cheia["linhas"])
            rodada_gravar("meli", comp, cheia)
            r = rodada_res("meli", comp)
            del cheia
            gc.collect()
    if chave == "meli":
        sm = base_sumidos("meli", comp)
        return render_template("canal.html", c=c, r=r, sumidos=len(sm),
                               sumidos_rs=round(sum(float(v.get("valor_prod") or 0) for v in sm.values()), 2))
    if chave == "magalu":
        return render_template("canal_magalu.html", c=c, r=r)
    if chave == "shopee":
        return render_template("canal_shopee.html", c=c, r=r)
    if chave == "madeira":
        return render_template("canal_madeira.html", c=c, r=r)
    if chave == "webcont":
        return render_template("canal_webcont.html", c=c, r=r)
    if chave == "colombo":
        return render_template("canal_colombo.html", c=c, r=r)
    if chave == "amazon":
        return render_template("canal_amazon.html", c=c, r=r)
    return render_template("canal.html", c=c, r=r)


# colunas da BASE LINHA A LINHA de cada canal: (rótulo, chave, tipo) — tipo: t texto, n número, p percentual, d data
LINHA_COLS = {
    "meli": [
        ("Data", "data", "d"), ("OC / Pedido mkt", "pedido_mkt", "t"), ("Pedido canal", "pedido_canal", "t"),
        ("ID mkt (Any)", "id_mkt", "t"), ("Pedido Any", "pedido_any", "t"), ("Conta", "conta", "t"),
        ("Status", "status", "t"), ("SKU", "sku", "t"), ("Descrição", "descricao", "t"), ("Anúncio", "anuncio", "t"), ("Tipo anúncio", "tipo", "t"),
        ("Valor produtos", "valor_prod", "n"), ("Tarifa venda", "tarifa", "n"), ("Frete pedido", "frete", "n"),
        ("Cupom seller", "cupom_seller", "n"), ("Cupom Meli", "cupom_meli", "n"),
        ("% comissão cobrada", "pct_comissao", "p"), ("% comissão sistema", "sis_pct", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("Diferença comissão", "diferenca", "n"),
        ("Tarifa zero?", "tarifa_zero", "b"), ("Faltante campanha", "faltante", "n"), ("Status faltante", "faltante_status", "t"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
    "magalu": [
        ("Data", "data", "d"), ("OC / Pedido", "pedido_mkt", "t"), ("Pedido Any", "pedido_any", "t"),
        ("Status", "status", "t"), ("Modalidade", "modalidade", "t"), ("CD", "conta", "t"), ("Forma pgto", "forma_pgto", "t"),
        ("Valor pago cliente", "valor_prod", "n"), ("Valor itens", "itens", "n"),
        ("% serviços Magalu", "pct_mkt", "p"), ("Serviços 1+2+3+4", "servicos", "n"), ("Intermediação", "intermediacao", "n"),
        ("Tecnologia", "tecnologia", "n"), ("MDR", "mdr", "n"), ("Tarifa fixa", "tarifa_fixa", "n"), ("Serviços pgto 2", "servicos_pgto2", "n"),
        ("Comissão real R$", "tarifa", "n"), ("% real", "pct_comissao", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("% sistema", "sis_pct", "p"), ("Diferença = rebate comissão", "diferenca", "n"),
        ("Desc. à vista Magalu", "desc_vista_magalu", "n"), ("Preço promo Magalu", "promo_magalu", "n"), ("Cupom Magalu", "cupom_magalu", "n"),
        ("Desc. à vista seller", "desc_vista_seller", "n"), ("Cupom seller", "cupom_seller", "n"),
        ("Copart. frete (logística)", "copart_frete", "n"), ("Custos logísticos", "custos_log", "n"), ("Repasse", "repasse", "n"), ("Líquido a receber", "liquido", "n"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
    "shopee": [
        ("Data", "data", "d"), ("OC / ID do pedido", "pedido_mkt", "t"), ("Pedido Any", "pedido_any", "t"), ("Rastreio", "id_mkt", "t"),
        ("Status", "status", "t"), ("Opção de envio", "envio", "t"), ("UF", "uf", "t"),
        ("SKU", "sku", "t"), ("Descrição", "descricao", "t"), ("SKUs do pedido", "anuncio", "t"), ("Itens", "itens", "n"), ("Qtd", "qtd", "n"),
        ("Subtotal produto", "valor_prod", "n"), ("Desc. vendedor", "desc_vendedor", "n"),
        ("Comissão bruta", "comissao_bruta", "n"), ("Serviço bruta", "servico_bruta", "n"), ("Ajuste ação comercial", "ajuste", "n"),
        ("Comissão real R$ = bruta + serviço − ajuste", "tarifa", "n"), ("% real", "pct_comissao", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("% sistema", "sis_pct", "p"),
        ("Diferença = rebate comissão", "diferenca", "n"),
        ("Incentivo Shopee (ação)", "incentivo", "n"), ("Incentivo de cupom", "cupom_shopee", "n"), ("Cupom vendedor", "cupom_seller", "n"), ("Moedas (qtd)", "moedas", "n"),
        ("Frete estimado", "frete", "n"), ("Frete pago comprador", "frete_comprador", "n"),
        ("Taxa transação", "taxa_transacao", "n"), ("Total global", "total_global", "n"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
    "madeira": [
        ("Data", "data", "d"), ("OC / Pedido", "pedido_mkt", "t"), ("Pedido Site MM", "pedido_canal", "t"), ("Pedido Any", "pedido_any", "t"),
        ("Status", "status", "t"), ("Pagamento", "pagamento", "t"), ("Parcelas", "parcelas", "n"), ("UF", "uf", "t"), ("Cidade", "cidade", "t"),
        ("SKU", "sku", "t"), ("Descrição", "descricao", "t"), ("SKUs do pedido", "anuncio", "t"), ("Itens", "itens", "n"), ("Qtd", "qtd", "n"),
        ("Valor pedido (GMV c/ frete)", "valor_prod", "n"), ("Valor itens", "valor_itens", "n"),
        ("Comissão cobrada R$", "tarifa", "n"), ("% real", "pct_comissao", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("% sistema", "sis_pct", "p"), ("Diferença = rebate comissão", "diferenca", "n"),
        ("NF", "nf", "t"), ("Data NF", "data_nf", "t"), ("Transportadora", "transportadora", "t"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
    "webcont": [
        ("Data", "data", "d"), ("OC / Pedido ERP", "pedido_mkt", "t"), ("Pedido Parceiro", "pedido_canal", "t"), ("Pedido Site", "id_mkt", "t"), ("Pedido Any", "pedido_any", "t"),
        ("Status", "status", "t"), ("Pagamento", "pagamento", "t"), ("UF", "uf", "t"), ("Cidade", "cidade", "t"),
        ("SKU", "sku", "t"), ("Descrição", "descricao", "t"), ("Qtd", "qtd", "n"),
        ("Total do pedido (GMV c/ frete)", "valor_prod", "n"), ("Valor produtos", "valor_itens", "n"), ("Valor frete", "frete", "n"), ("Desconto", "desconto", "n"), ("Valor repasse", "repasse", "n"),
        ("Comissão retida R$ (do relatório)", "comissao_retida", "n"),
        ("Comissão real R$ (negativa vira 0)", "tarifa", "n"), ("% real", "pct_comissao", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("% sistema", "sis_pct", "p"), ("Diferença = rebate comissão", "diferenca", "n"),
        ("NF", "nf", "t"), ("Transportadora", "transportadora", "t"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
        ],
    "amazon": [
        ("Data", "data", "d"), ("ID do pedido / OC", "pedido_mkt", "t"), ("Pedido Any", "pedido_any", "t"),
        ("Produto", "produto", "t"), ("Status", "status", "t"), ("Pagamento", "pagamento", "t"),
        ("Transações", "itens", "n"),
        ("Produto R$", "valor_prod", "n"), ("Desconto promocional", "desconto", "n"), ("Frete / outros", "frete", "n"),
        ("BASE da comissão", "sis_base", "n"), ("Repasse", "repasse", "n"),
        ("Comissão cobrada R$", "tarifa", "n"), ("% cobrado", "pct_comissao", "p"),
        ("Faixa da categoria", "tipo", "t"), ("% negociado (sem a taxa)", "negociada", "p"),
        ("Comissão pela faixa R$", "faixa_rs", "n"),
        ("Comissão cadastro R$", "sis_rs", "n"), ("% cadastro", "sis_pct", "p"), ("% no ERP", "erp_pct", "p"),
        ("Do rebate: categoria menor", "desvio_cadastro", "n"), ("Do rebate: fora da faixa", "erro_amazon", "n"),
        ("REBATE COMISSÃO (sistema − cobrado)", "diferenca", "n"), ("Reembolso", "reembolso", "n"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
    "colombo": [
        ("Data", "data", "d"), ("OC / Entrega", "pedido_mkt", "t"), ("Pedido Colombo", "pedido_canal", "t"), ("Pedido Any", "pedido_any", "t"),
        ("Status", "status", "t"), ("Pagamento", "pagamento", "t"), ("Parcelas", "parcelas", "n"), ("UF", "uf", "t"), ("Cidade", "cidade", "t"),
        ("SKU", "sku", "t"), ("Descrição", "descricao", "t"), ("SKUs do pedido", "anuncio", "t"), ("Itens", "itens", "n"), ("Qtd", "qtd", "n"),
        ("Total do pedido", "valor_prod", "n"), ("Valor mercadorias", "valor_itens", "n"), ("Valor frete", "frete", "n"), ("Desconto", "desconto", "n"),
        ("% do item", "tipo", "t"), ("Comissão cobrada R$", "tarifa", "n"), ("% real", "pct_comissao", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("% sistema", "sis_pct", "p"), ("Diferença = rebate comissão", "diferenca", "n"),
        ("Data entrega", "data_nf", "d"), ("Cliente", "cliente", "t"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
}



ERP_COLS = [
    ("ERP · Pedido", "erp_pedido_erp", "t"), ("ERP · Status", "erp_status", "t"), ("ERP · Natureza", "erp_natureza", "t"),
    ("ERP · Data emissão", "erp_data", "d"), ("ERP · Valor produtos", "erp_valor_prod", "n"),
    ("ERP · Valor frete", "erp_valor_frete", "n"), ("ERP · Valor total", "erp_valor_total", "n"),
    ("ERP · % comissão", "erp_pct_comissao", "p"), ("ERP · Comissão R$", "erp_comissao_erp_rs", "n"),
    ("ERP · IPI", "erp_ipi", "n"), ("ERP · NF", "erp_nf", "t"), ("ERP · Data NF", "erp_data_nf", "d"),
    ("ERP · UF", "erp_uf", "t"), ("ERP · Cidade", "erp_cidade", "t"), ("ERP · Transportadora", "erp_transportadora", "t"),
    ("ERP · Volumes", "erp_volumes", "n"), ("ERP · Obs 05", "erp_obs05", "t"),
]
ERP_SO_COLS = [("Data", "data", "d"), ("OC", "oc", "t"), ("Canal (ERP)", "canal_erp", "t")] + \
    [(rot.replace("ERP · ", ""), k[4:], t) for rot, k, t in ERP_COLS if k not in ("erp_data",)]


def _tem_tx_fin(chave: str) -> bool:
    """O canal tem taxa financeira cadastrada (tabela nova ou parâmetro antigo)?"""
    NOMES = {"madeira": ("MADEIRA",), "webcont": ("WEBCONTINENTAL", "WEBCONT"),
             "magalu": ("MAGAZINE", "MAGALU"), "meli": ("MERCADO LIVRE",), "shopee": ("SHOPEE",),
             "colombo": ("COLOMBO",), "amazon": ("AMAZON",), "cbahia": ("CASAS BAHIA",)}
    if parametros_canal(NOMES.get(chave, (chave.upper(),)))["tx_fin"]:
        return True
    return bool((parametros().get("tx_financeira") or {}).get(chave))


def separar_tx_financeira(chave: str, linhas: list[dict]) -> list[dict]:
    """Madeira Madeira: o ERP traz 21% = 17% de comissão + 4% de taxa financeira
    (antecipação). Cada linha ganha % e R$ separados; a base é a do ERP (valor produtos)."""
    NOMES_TX = {"madeira": ("MADEIRA",), "webcont": ("WEBCONTINENTAL", "WEBCONT"),
                "magalu": ("MAGAZINE", "MAGALU"), "meli": ("MERCADO LIVRE",), "shopee": ("SHOPEE",),
                "colombo": ("COLOMBO",), "amazon": ("AMAZON",), "cbahia": ("CASAS BAHIA",)}
    tx = parametros_canal(NOMES_TX.get(chave, (chave.upper(),)))["tx_fin"] * 100 or None
    if tx is None:   # tabela sem o campo: cai no parâmetro antigo
        tx = (parametros().get("tx_financeira") or {}).get(chave)
    if not tx:
        return linhas
    tx = float(tx) / 100.0
    for l in linhas:
        pct = l.get("pct_comissao") or 0.0
        base = l.get("valor_prod") or 0.0
        l["pct_canal"] = round(max(0.0, pct - tx), 6)
        l["pct_tx_fin"] = tx
        l["comissao_canal_rs"] = round(base * l["pct_canal"], 2)
        l["tx_fin_rs"] = round(base * tx, 2)
    return linhas


TX_FIN_COLS = [("Comissão canal %", "pct_canal", "p"), ("Comissão canal R$", "comissao_canal_rs", "n"),
               ("Tx financeira canal %", "pct_tx_fin", "p"), ("Tx financeira canal R$", "tx_fin_rs", "n")]


@app.route("/linha/<chave>")
@logado
def linha(chave):
    c = canal_por_chave().get(chave) or abort(404)
    comp = comp_atual()
    r = rodada(chave, comp) if c["ativo"] else None
    q = (request.args.get("q") or "").strip().lower()
    idx = erp_ler()["ocs"]
    if r:
        cols = LINHA_COLS.get(chave, []) + ERP_COLS
        cad = descricoes_ler()["skus"]
        base = erp_por_base(chave) if chave == "magalu" else {}
        linhas = []
        for l in r["linhas"]:
            e = idx.get(l["pedido_mkt"]) or idx.get(l["pedido_canal"]) or base.get(l["pedido_mkt"]) or {}
            m = dict(l)
            m["descricao"] = descricao_de(l.get("sku"), cad)
            for _, k, _ in ERP_COLS:
                m[k] = e.get(k[4:]) if e else None
            linhas.append(m)
        origem = "canal + ERP"
    else:
        cols = ERP_SO_COLS
        linhas = erp_linhas(chave, comp)
        if _tem_tx_fin(chave):
            linhas = separar_tx_financeira(chave, [dict(l) for l in linhas])
            i = next(i for i, (_, k, _) in enumerate(cols) if k == "comissao_erp_rs") + 1
            cols = cols[:i] + TX_FIN_COLS + cols[i:]
        origem = "só ERP"
    if q:
        linhas = [l for l in linhas if q in " ".join(str(l.get(k) or "") for _, k, _ in cols).lower()]
    ordem = request.args.get("ord") or "data"
    desc = request.args.get("desc") == "1"
    if any(k == ordem for _, k, _ in cols):
        linhas.sort(key=lambda l: (l.get(ordem) is None, l.get(ordem) if l.get(ordem) is not None else 0), reverse=desc)
    pag = max(1, int(request.args.get("p", 1) or 1))
    por = 200
    total = len(linhas)
    return render_template("linha.html", c=c, r=r, cols=cols, linhas=linhas[(pag - 1) * por: pag * por],
                           total=total, total_geral=len(linhas) if not q else None, pag=pag,
                           paginas=max(1, -(-total // por)), q=q, ordem=ordem, desc=desc, origem=origem)


@app.route("/pedidos")
@logado
def pedidos():
    comp = comp_atual()
    q = (request.args.get("q") or "").strip().lower()
    so = request.args.get("so") or ""
    canal_f = request.args.get("canal") or ""
    linhas = []
    for c in canais():
        if canal_f and c["chave"] != canal_f:
            continue
        r = rodada(c["chave"], comp)
        if r:
            linhas += r["linhas"]
    if q:
        linhas = [l for l in linhas if q in (l["pedido_mkt"] + " " + l["pedido_canal"] + " " + l["pedido_any"]
                                              + " " + l["sku"] + " " + l["anuncio"]).lower()]
    if so == "com":
        linhas = [l for l in linhas if l["rebate_total"]]
    elif so == "tz":
        linhas = [l for l in linhas if l["tarifa_zero"]]
    elif so == "dif":
        linhas = [l for l in linhas if l["rebate_comissao"]]
    linhas.sort(key=lambda l: (l["data"], l["pedido_mkt"]))
    pag = max(1, int(request.args.get("p", 1) or 1))
    por = 200
    total = len(linhas)
    fatia = linhas[(pag - 1) * por: pag * por]
    return render_template("pedidos.html", linhas=fatia, total=total, pag=pag,
                           paginas=max(1, -(-total // por)), q=q, so=so, canal_f=canal_f)


CONTA_APELIDO = {"loja_principal": "Loja Principal", "coletas_rs": "Loja RS"}


@app.template_filter("contanome")
def f_conta(v) -> str:
    """Nome da conta do Meli do jeito que a casa fala (LOJA_PRINCIPAL = Loja
    Principal · COLETAS_RS = Loja RS); conta nova aparece como veio no arquivo."""
    t = str(v or "").strip()
    return CONTA_APELIDO.get(unidecode_lower(t).replace(" ", "_"), t.replace("_", " ").title() or "—")


@app.route("/faltante", methods=["GET", "POST"])
@logado
def faltante():
    comp = comp_atual()
    if request.method == "POST":
        if session.get("papel") not in PODE["faltante"]:
            abort(403)
        tab = faltante_ler("meli")
        n = 0
        acao = request.form.get("acao")
        if acao == "aceitar_todas":
            for k, v in request.form.items():
                if k.startswith("sug_") and v not in ("", None):
                    ped = k[4:]
                    tab[ped] = {"valor": float(v), "quem": session["usuario"], "quando": agora().isoformat(),
                                "replicado": True}
                    n += 1
        else:
            for k, v in request.form.items():
                if not k.startswith("val_"):
                    continue
                ped = k[4:]
                v = (v or "").strip().replace(".", "").replace(",", ".") if "," in (v or "") else (v or "").strip()
                obs = (request.form.get("obs_" + ped) or "").strip()
                if v == "":
                    if ped in tab and request.form.get("limpar_" + ped):
                        del tab[ped]; n += 1
                    continue
                try:
                    val = float(v)
                except ValueError:
                    continue
                antigo = tab.get(ped) or {}
                if antigo.get("valor") != val or antigo.get("obs", "") != obs:
                    tab[ped] = {"valor": val, "obs": obs, "quem": session["usuario"],
                                "quando": agora().isoformat()}
                    n += 1
        faltante_gravar(tab, "meli")
        recalcular_meli(comp)
        flash(f"{n} pedido(s) atualizado(s) na tabela Faltante campanha. Rebates recalculados.")
        return redirect(url_for("faltante", mes=comp))
    r = rodada("meli", comp)
    itens = meli.sugestoes_faltante(r["linhas"], faltante_ler("meli")) if r else []
    pend = sum(1 for i in itens if i["faltante_status"] == "pendente")
    sug = sum(1 for i in itens if i["sugestao"] is not None and i["faltante_status"] == "pendente")
    # todas as contas do mês no Meli (não só as que têm tarifa zero) — assim a
    # barra de filtro mostra a divisão real, mesmo quando uma conta está zerada
    contas = sorted({(l.get("conta") or "") for l in (r["linhas"] if r else []) if l.get("conta")})
    return render_template("faltante.html", itens=itens, r=r, pend=pend, sug=sug, contas=contas)


@app.route("/faltante/um", methods=["POST"])
@logado
@exige("faltante")
def faltante_um():
    """OK individual: grava um pedido só e recalcula o Meli da competência."""
    d = request.get_json(silent=True) or {}
    ped = str(d.get("pedido") or "").strip()
    comp = d.get("comp") or comp_atual()
    if not ped:
        return jsonify({"ok": False, "erro": "pedido vazio"}), 400
    tab = faltante_ler("meli")
    v = str(d.get("valor") or "").strip()
    v = v.replace(".", "").replace(",", ".") if "," in v else v
    if v == "":
        tab.pop(ped, None); status = "pendente"
    else:
        try:
            val = float(v)
        except ValueError:
            return jsonify({"ok": False, "erro": "valor inválido"}), 400
        tab[ped] = {"valor": val, "obs": (d.get("obs") or "").strip(), "quem": session["usuario"],
                    "quando": agora().isoformat()}
        status = "preenchido"
    faltante_gravar(tab, "meli")
    r = recalcular_meli(comp)
    return jsonify({"ok": True, "status": status, "quem": session["usuario"], "quando": f_quando(agora().isoformat()),
                    "faltante_total": r["resumo"]["faltante"] if r else 0,
                    "pendentes": r["resumo"]["faltante_pendentes"] if r else 0})


@app.route("/faltante/varios", methods=["POST"])
@logado
@exige("faltante")
def faltante_varios():
    """OK nos selecionados: grava vários pedidos de uma vez e recalcula uma vez só."""
    d = request.get_json(silent=True) or {}
    comp = d.get("comp") or comp_atual()
    tab = faltante_ler("meli")
    feitos, status = 0, {}
    for it in d.get("itens") or []:
        ped = str(it.get("pedido") or "").strip()
        if not ped:
            continue
        v = str(it.get("valor") or "").strip()
        v = v.replace(".", "").replace(",", ".") if "," in v else v
        if v == "":
            tab.pop(ped, None); status[ped] = "pendente"; feitos += 1
            continue
        try:
            val = float(v)
        except ValueError:
            status[ped] = "erro"; continue
        tab[ped] = {"valor": val, "obs": (it.get("obs") or "").strip(), "quem": session["usuario"],
                    "quando": agora().isoformat()}
        status[ped] = "preenchido"; feitos += 1
    faltante_gravar(tab, "meli")
    r = recalcular_meli(comp)
    return jsonify({"ok": True, "feitos": feitos, "status": status, "quem": session["usuario"],
                    "quando": f_quando(agora().isoformat()),
                    "pendentes": r["resumo"]["faltante_pendentes"] if r else 0})


@app.route("/arquivos", methods=["GET"])
@logado
def arquivos():
    comp = comp_atual()
    hist = []
    for n in sorted(os.listdir(pasta("rodadas")), reverse=True):
        if not n.endswith(".resumo.json"):
            continue                      # o histórico lê só os resumos leves
        r = _json_ler(pasta("rodadas", n), None)
        if r:
            hist.append({"canal": r["canal"], "comp": r["competencia"], "quando": r["quando"],
                         "quem": r.get("quem"), "arquivo": r["arquivo"]["nome"],
                         "pedidos": (r.get("resumo") or {}).get("pedidos", 0),
                         "total": (r.get("resumo") or {}).get("rebate_total", 0),
                         "sistema": bool(r.get("sistema"))})
    hist.sort(key=lambda h: h["quando"], reverse=True)
    idx = erp_ler()
    for u in idx["uploads"]:
        if not u.get("bytes"):
            try:
                u["bytes"] = os.path.getsize(u["caminho"])
            except OSError:
                u["bytes"] = 0
    erp_res = erp.resumo_por_box([l for l in idx["ocs"].values() if l["competencia"] == comp])
    nomes_erp = sorted({l["canal_nome"] for l in idx["ocs"].values()})
    m = mapa_erp_box()
    sem_box = [n for n in nomes_erp if m.get(n, "outros") == "outros"]
    ml = mlbs_ler()
    pend = pend_ler()
    pend_por = {}
    for x in pend:
        pend_por.setdefault(x["chave"], []).append(x)
    rods = {c["chave"]: rodada_res(c["chave"], comp) for c in canais() if c["ativo"]}
    return render_template("arquivos.html", pend=pend, pend_por=pend_por, rods=rods, hist=hist, r_meli=rodada_res("meli", comp), erp_idx=idx, erp_res=erp_res,
                           nomes_erp=nomes_erp, sem_box=sem_box, mapa_erp=m,
                           mlbs_ult=(ml["uploads"][-1] if ml["uploads"] else None), mlbs_total=len(ml["mlbs"]),
                           mg_vendas=(vendas_magalu_ler()["uploads"] or [None])[-1],
                           mg_full=next((u for u in reversed(full_magalu_ler()["uploads"])
                                         if u.get("qual") != "coleta"), None),
                           mg_coleta=next((u for u in reversed(full_magalu_ler()["uploads"])
                                           if u.get("qual") == "coleta"), None),
                           mg_real=(full_real_ler()["uploads"] or [None])[-1])


def processar_meli(tipo: str, destino: str, nome: str, quem: str) -> str:
    """Lê um arquivo do box Mercado Livre já guardado no disco e roda o box.
    tipo = base (Planilha 1 · Completo) · rebates (Planilha 2 · Resumo) · sistema."""
    if tipo == "rebates":
        df, diag = meli.ler_resumo_rebates(destino)
        novos = meli.lista_mlbs(df)
        d = mlbs_ler()
        n_novos = n_mud = 0
        for k, v in novos.items():
            a = d["mlbs"].get(k)
            if a is None:
                n_novos += 1
                d["mlbs"][k] = v
            else:
                if a["tipo"] != v["tipo"] or a["sku"] != v["sku"]:
                    n_mud += 1
                v["primeira"] = min(a["primeira"], v["primeira"])
                v["pedidos"] = a["pedidos"] + v["pedidos"] if v["primeira"] > a["ultima"] else max(a["pedidos"], v["pedidos"])
                d["mlbs"][k] = v
        d["uploads"].append({"nome": nome, "caminho": destino, "quando": agora().isoformat(),
                             "quem": quem, "diag": diag})
        d["uploads"] = d["uploads"][-30:]
        mlbs_gravar(d)
        rp = resumo_meli_ler()
        for r_ in df.itertuples(index=False):
            rp["pedidos"][r_.pedido_mkt] = {k: (str(getattr(r_, k)) if k == "data" else getattr(r_, k)) for k in RESUMO_CAMPOS}
            for k in ("qtd", "valor_prod", "frete", "frete_coletas", "cupom_canal", "cupom_seller", "valor_meli", "pct_bruta", "com_bruta", "rebate_bi", "com_liq"):
                rp["pedidos"][r_.pedido_mkt][k] = float(rp["pedidos"][r_.pedido_mkt][k] or 0)
        resumo_meli_gravar(rp)
        return (f"Resumo de Rebates lido: {diag['linhas']} pedidos de {f_dia(diag['de'])} a {f_dia(diag['ate'])} · "
                f"{diag['mlbs']} MLB's no arquivo — {n_novos} novos · {n_mud} mudaram de tipo/SKU. "
                f"Lista de MLB's: {len(d['mlbs'])} anúncios.")
    if tipo == "sistema":
        comp = comp_atual()
        r = rodada("meli", comp)
        if not r:
            return f"Suba primeiro a Tabela Geral de {f_mesano(comp)}; a comissão do sistema entra em cima dela."
        sis, diag = meli.ler_comissao_sistema(destino)
        r["sistema"] = {"nome": nome, "caminho": destino, "quando": agora().isoformat(), "diag": diag}
        rodada_gravar("meli", comp, r)
        r = recalcular_meli(comp)
        return (f"Comissão do sistema lida ({diag['linhas']} pedidos, {diag['modo']}). "
                f"Rebate de comissão: R$ {f_brl(r['resumo']['rebate_comissao'])}.")
    df, diag = meli.ler_tabela_geral(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    fora = 0
    for comp, n in list(comps.items()):
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n  # linhas soltas de outro mês = sujeira do filtro do BI, não competência
            df = df[df["competencia"] != comp]
    canc = base_marcar_sumidos("meli", df, "pedido_mkt", nome)   # antes do upsert: compara com a base como ela está
    res = base_upsert("meli", df, "pedido_mkt", nome, destino, quem)
    feitos = []
    for comp in res:
        r = dict(rodada_cab("meli", comp) or {"canal": "meli", "competencia": comp}); r.pop("linhas_n", None)
        r.update({"quando": agora().isoformat(), "quem": quem,
                  "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                  "linhas": [], "resumo": {}})
        rodada_gravar("meli", comp, r)
        r = recalcular_meli(comp)
        if not _resumo_pronto(r):
            continue
        feitos.append((comp, r["resumo"]["pedidos"], r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora (filtro do BI)." if fora else ""
    s_ = sum(v["sumiram"] for v in canc.values())
    v_ = sum(v["voltaram"] for v in canc.values())
    cnc = ""
    if s_:
        cnc = (f" {s_} pedido(s) que estavam na base sumiram deste relatório (de {f_dia(diag['de'])} a {f_dia(diag['ate'])})"
               " — o Mercado Livre cancela tirando a linha, então saíram dos rebates e das coletas.")
    if v_:
        cnc += f" {v_} voltaram a aparecer e foram reativados."
    return f"Mercado Livre lido — {_txt_upsert(res)}. {txt}. Linhas rejeitadas: {diag['n_rejeitadas']}.{extra}{cnc}"


PEND_ORDEM = {"erp": 0, "base": 1, "rebates": 2, "sistema": 3}


def _guardar_upload(f, pasta_destino: str) -> list[tuple[str, str]]:
    """Salva o arquivo enviado. Se for .zip, extrai as planilhas de dentro
    (.xlsx/.xls/.csv/.txt) na ordem do nome — o portal quebra exports grandes
    em part_1_of_2, part_2_of_2… Devolve [(nome, caminho)]."""
    import zipfile
    nome = secure_filename(f.filename)
    carimbo = agora().strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(pasta_destino, f"{carimbo}_{nome}")
    os.makedirs(pasta_destino, exist_ok=True)
    f.save(destino)
    if not nome.lower().endswith(".zip"):
        return [(nome, destino)]
    out = []
    with zipfile.ZipFile(destino) as z:
        membros = sorted(m for m in z.namelist()
                         if m.lower().endswith((".xlsx", ".xls", ".csv", ".txt")) and not os.path.basename(m).startswith(("~", "."))
                         and not m.endswith("/"))
        for i, m in enumerate(membros, 1):
            nm = secure_filename(os.path.basename(m))
            cam = os.path.join(pasta_destino, f"{carimbo}_{i:02d}_{nm}")
            with z.open(m) as src, open(cam, "wb") as dst:
                shutil.copyfileobj(src, dst)
            out.append((nm, cam))
    if not out:
        raise ValueError("o .zip não tem nenhuma planilha (.xlsx/.csv) dentro")
    return out


def pend_ler() -> list[dict]:
    return _json_ler(pasta("pendentes.json"), [])


def pend_gravar(lista):
    _json_gravar(pasta("pendentes.json"), lista)


def pend_processar(chave: str | None = None) -> list[str]:
    """Roda os arquivos que estão aguardando (todos, ou só os de um box), na
    ordem certa: ERP primeiro (é a comissão do sistema), depois os canais."""
    lista = pend_ler()
    fila = [x for x in lista if chave is None or x["chave"] == chave]
    fila.sort(key=lambda x: (PEND_ORDEM.get(x["tipo"], 9), x["quando"]))
    msgs, feitos = [], []
    for x in fila:
        try:
            if x["chave"] == "erp":
                m = processar_erp(x["caminho"], x["nome"], x["quem"], recalcular=False)
            elif x["chave"] == "meli":
                m = processar_meli(x["tipo"], x["caminho"], x["nome"], x["quem"])
            elif x["chave"] == "magalu":
                m = (processar_magalu_vendas(x["caminho"], x["nome"], x["quem"]) if x.get("tipo") == "vendas"
                     else processar_magalu_full_real(x["caminho"], x["nome"], x["quem"]) if x.get("tipo") == "full_real"
                     else processar_magalu_full(x["caminho"], x["nome"], x["quem"]) if x.get("tipo") in ("full", "coleta")
                     else processar_magalu(x["caminho"], x["nome"], x["quem"]))
            elif x["chave"] == "shopee":
                m = processar_shopee(x["caminho"], x["nome"], x["quem"])
            elif x["chave"] == "madeira":
                m = processar_madeira(x["caminho"], x["nome"], x["quem"])
            elif x["chave"] == "webcont":
                m = processar_webcont(x["caminho"], x["nome"], x["quem"])
            elif x["chave"] == "colombo":
                m = processar_colombo(x["caminho"], x["nome"], x["quem"])
            elif x["chave"] == "amazon":
                m = processar_amazon(x["caminho"], x["nome"], x["quem"])
            else:
                m = f"{x['nome']}: o box {x['chave']} ainda não tem motor."
        except Exception as e:  # noqa: BLE001
            m = f"{x['nome']}: não consegui ler — {e}"
        msgs.append(m)
        feitos.append(x["quando"])
        gc.collect()   # solta a memória do arquivo anterior antes do próximo
    pend_gravar([x for x in pend_ler() if x["quando"] not in feitos])
    return msgs


@app.route("/arquivos/subir/<chave>", methods=["POST"])
@logado
@exige("arquivos")
def subir(chave):
    """Guarda o arquivo e deixa aguardando. Roda com '▶ Rodar <box>' ou com o
    '▶ Rodar o PLUTOS' (tudo de uma vez). Com rodar=1 no form, roda na hora."""
    c = canal_por_chave().get(chave) or abort(404)
    if not c["ativo"]:
        flash(f"O box {c['nome']} ainda está em construção.")
        return redirect(url_for("arquivos"))
    f = request.files.get("arquivo")
    if not f or not f.filename:
        flash("Escolha um arquivo.")
        return redirect(url_for("arquivos"))
    tipo = request.form.get("tipo", "base")
    try:
        arquivos_ = _guardar_upload(f, pasta("arquivos", chave))
    except Exception as e:  # noqa: BLE001
        flash(f"Não consegui guardar o arquivo: {e}")
        return redirect(url_for("arquivos"))
    lista = pend_ler()
    for i, (nome, destino) in enumerate(arquivos_):
        lista.append({"chave": chave, "tipo": tipo, "nome": nome, "caminho": destino,
                      "quando": (agora() + timedelta(milliseconds=i)).isoformat(), "quem": session["usuario"]})
    pend_gravar(lista)
    nome = " + ".join(n for n, _ in arquivos_) if len(arquivos_) > 1 else arquivos_[0][0]
    if request.form.get("rodar"):
        for m in pend_processar(chave):
            flash(m)
        # fica em Arquivos: quem sobe costuma subir vários seguidos. Só o
        # ▶ Rodar o PLUTOS leva para o GERAL, quando termina tudo.
        return redirect(url_for("arquivos", mes=comp_atual()))
    n = len([x for x in lista if x["chave"] == chave])
    flash(f"{nome} guardado no box {c['nome']} — {n} arquivo(s) aguardando. Clique ▶ Rodar {c['nome']} ou ▶ Rodar o PLUTOS.")
    return redirect(url_for("arquivos"))


@app.route("/arquivos/rodar/<chave>", methods=["POST"])
@logado
@exige("arquivos")
def rodar_box(chave):
    msgs = pend_processar(chave)
    if chave == "erp":
        reclassificar_erp()
        for comp in competencias():
            recalcular_canais(comp)
    if not msgs:
        flash("Nada aguardando neste box.")
    for m in msgs:
        flash(m)
    return redirect(url_for("arquivos"))


@app.route("/arquivos/erp/baixar/<quando>")
@logado
def erp_baixar(quando):
    u = next((u for u in erp_ler()["uploads"] if u["quando"] == quando), None)
    if not u or not os.path.exists(u.get("caminho") or ""):
        flash("Esse arquivo não está mais guardado no servidor.")
        return redirect(url_for("arquivos"))
    return send_file(u["caminho"], as_attachment=True, download_name=u["nome"])


@app.route("/arquivos/erp/remover/<quando>", methods=["POST"])
@logado
@exige("arquivos")
def erp_remover(quando):
    """Tira o arquivo da lista (e do disco). As OCs que ele trouxe continuam no
    índice — para zerar o índice use 'excluir todos'."""
    idx = erp_ler()
    u = next((u for u in idx["uploads"] if u["quando"] == quando), None)
    if u:
        idx["uploads"] = [x for x in idx["uploads"] if x["quando"] != quando]
        try:
            os.remove(u["caminho"])
        except OSError:
            pass
        erp_gravar(idx)
        flash(f"Arquivo {u['nome']} removido da lista. As OCs já lidas continuam no índice.")
    return redirect(url_for("arquivos"))


@app.route("/arquivos/erp/limpar", methods=["POST"])
@logado
@exige("arquivos")
def erp_limpar():
    """Excluir todos: apaga os arquivos do ERP E zera o índice de OCs. As rodadas
    dos canais são recalculadas sem comissão do sistema."""
    idx = erp_ler()
    for u in idx["uploads"]:
        try:
            os.remove(u["caminho"])
        except OSError:
            pass
    n = len(idx["ocs"])
    erp_gravar({"ocs": {}, "uploads": []})
    for comp in competencias():
        recalcular_canais(comp)
    flash(f"ERP zerado: {len(idx['uploads'])} arquivo(s) e {n} OCs excluídos. Suba o export de novo quando quiser.")
    return redirect(url_for("arquivos"))


@app.route("/arquivos/rodar", methods=["POST"])
@logado
@exige("arquivos")
def rodar_tudo():
    """RODAR: reclassifica o ERP pelos boxes atuais e recalcula todos os canais
    em todas as competências com os arquivos já guardados. Para depois de
    atualizar arquivos, tabelas manuais, comissões ou boxes."""
    t0 = agora()
    msgs = pend_processar()
    reclassificar_erp()
    feitos = []
    for comp in competencias():
        feitos += recalcular_canais(comp)
    seg = (agora() - t0).total_seconds()
    for m in msgs:
        flash(m)
    flash(f"PLUTOS rodado em {seg:.0f} s — {len(msgs)} arquivo(s) processado(s). ERP reclassificado ({len(erp_ler()['ocs'])} OCs)."
          + (" " + " · ".join(feitos) if feitos else " Nenhum canal com rodada ainda."))
    return redirect(url_for("painel", mes=comp_atual()))   # terminou tudo: abre o GERAL


def processar_erp(destino: str, nome: str, quem: str, recalcular: bool = True) -> str:
    """Lê o export do ERP já guardado e distribui pelos boxes (índice por OC)."""
    linhas, diag = erp.ler(destino, mapa_erp_box())
    # NUNCA soma duas vezes: a OC é a chave. OC que já existe e veio igual é
    # ignorada; OC que já existe e veio diferente (NF emitida, status, valor) é
    # atualizada — continua sendo UMA linha; OC nova entra.
    idx = erp_ler()
    novas = atualizadas = iguais = 0
    # dentro do próprio arquivo a OC pode vir repetida (o ERP repete a linha em
    # alguns casos): vale a ÚLTIMA ocorrência, uma vez só
    unicas = {}
    for l in linhas:
        unicas[l["oc"]] = l
    linhas = list(unicas.values())
    for l in linhas:
        atual = idx["ocs"].get(l["oc"])
        if atual is None:
            novas += 1
            idx["ocs"][l["oc"]] = l
        elif {k: v for k, v in atual.items() if k != "box"} == {k: v for k, v in l.items() if k != "box"}:
            iguais += 1
        else:
            atualizadas += 1
            idx["ocs"][l["oc"]] = l
    idx["uploads"].append({"nome": nome, "caminho": destino, "quando": agora().isoformat(), "quem": quem,
                           "bytes": os.path.getsize(destino), "linhas": diag["linhas"], "novas": novas, "atualizadas": atualizadas, "iguais": iguais,
                           "de": diag["de"], "ate": diag["ate"],
                           "canais": diag["canais"], "sem_box": diag["sem_box"], "rejeitadas": diag["rejeitadas"],
                           "ocs_duplicadas": diag["ocs_duplicadas"]})
    idx["uploads"] = idx["uploads"][-50:]
    erp_gravar(idx)
    # os canais ativos recalculam com a comissão do ERP
    recalc = []
    if recalcular:
        for comp in competencias():
            if recalcular_meli(comp):
                recalc.append(comp)
    canais_txt = " · ".join(f"{k} {v}" for k, v in sorted(diag["canais"].items(), key=lambda x: -x[1]))
    extra = (" Sem box ainda: " + ", ".join(f"{k} ({v})" for k, v in diag["sem_box"].items()) + ".") if diag["sem_box"] else ""
    return (f"ERP lido: {diag['linhas']} pedidos de {f_dia(diag['de'])} a {f_dia(diag['ate'])} — "
          f"{novas} OCs novas · {atualizadas} atualizadas · {iguais} já estavam iguais (ignoradas). "
          f"Índice: {len(idx['ocs'])} OCs, sem duplicar. {canais_txt}.{extra}"
          + (f" Mercado Livre recalculado ({', '.join(f_mesano(c) for c in recalc)})." if recalc else ""))


@app.route("/arquivos/subir-erp", methods=["POST"])
@logado
@exige("arquivos")
def subir_erp():
    f = request.files.get("arquivo")
    if not f or not f.filename:
        flash("Escolha o arquivo do ERP (.csv ou .xlsx).")
        return redirect(url_for("arquivos"))
    try:
        arquivos_ = _guardar_upload(f, pasta("arquivos", "erp"))
    except Exception as e:  # noqa: BLE001
        flash(f"Não consegui guardar o arquivo: {e}")
        return redirect(url_for("arquivos"))
    lista = pend_ler()
    for i, (nome, destino) in enumerate(arquivos_):
        lista.append({"chave": "erp", "tipo": "erp", "nome": nome, "caminho": destino,
                      "quando": (agora() + timedelta(milliseconds=i)).isoformat(), "quem": session["usuario"]})
    pend_gravar(lista)
    nome = " + ".join(n for n, _ in arquivos_) if len(arquivos_) > 1 else arquivos_[0][0]
    if request.form.get("rodar"):
        for m in pend_processar("erp"):
            flash(m)
        reclassificar_erp()
        for comp in competencias():
            recalcular_canais(comp)
        return redirect(url_for("arquivos"))
    flash(f"{nome} guardado como arquivo principal — aguardando. Clique ▶ Rodar ERP ou ▶ Rodar o PLUTOS.")
    return redirect(url_for("arquivos"))


@app.route("/arquivos/box", methods=["POST"])
@logado
@exige("parametros")
def box_criar():
    """Cria (ou ajusta) um box de canal pela tela Arquivos: nome + quais nomes do
    ERP caem nele. O box nasce sem motor (em construção), mas já com Linha a linha."""
    p = parametros()
    extra = p.get("boxes_extra") or []
    acao = request.form.get("acao") or "criar"
    if acao == "excluir":
        chave = request.form.get("chave")
        extra = [e for e in extra if e["chave"] != chave]
        flash("Box removido. Os pedidos dele voltam para 'sem box'.")
    else:
        nome = (request.form.get("nome") or "").strip()
        nomes_erp = [n.strip().upper() for n in request.form.getlist("nomes_erp") if n.strip()]
        livre = (request.form.get("nome_erp_livre") or "").strip().upper()
        if livre:
            nomes_erp.append(livre)
        chave = (request.form.get("chave") or "").strip() or re.sub(r"[^a-z0-9]", "", nome.lower())[:16]
        if not nome or not chave:
            flash("Dê um nome ao box.")
            return redirect(url_for("arquivos"))
        base = {c["chave"]: c for c in CANAIS_BASE}
        if chave in base:
            # box da casa: só ajusta os nomes do ERP que caem nele
            extra = [e for e in extra if e["chave"] != chave] + [{"chave": chave, "nome": base[chave]["nome"], "nomes_erp": nomes_erp}]
        else:
            extra = [e for e in extra if e["chave"] != chave] + [{"chave": chave, "nome": nome, "nomes_erp": nomes_erp}]
        flash(f"Box {nome} gravado" + (f" — recebe do ERP: {', '.join(nomes_erp)}." if nomes_erp else "."))
    p["boxes_extra"] = extra
    _json_gravar(pasta("parametros.json"), p)
    reclassificar_erp()
    return redirect(url_for("arquivos"))


@app.route("/parametros", methods=["GET", "POST"])
@logado
def parametros_tela():
    if request.method == "POST":
        if session.get("papel") not in PODE["parametros"]:
            abort(403)
        p = parametros()
        acao = request.form.get("acao") or "geral"
        if acao == "comissoes":
            linhas = []
            n = int(request.form.get("n") or 0)
            for i in range(n):
                lin = {k: (request.form.get(f"{k}_{i}") or "").strip() for k in COMISSAO_CAMPOS}
                if request.form.get(f"del_{i}") or not (lin["canal"] or lin["codigo"]):
                    continue
                linhas.append(lin)
            novo = {k: (request.form.get(f"{k}_novo") or "").strip() for k in COMISSAO_CAMPOS}
            if novo["canal"]:
                linhas.append(novo)
            p["comissoes"] = linhas
            p["comissoes_quando"] = agora().isoformat(); p["comissoes_quem"] = session["usuario"]
            _json_gravar(pasta("parametros.json"), p)
            flash(f"Tabela de comissões gravada: {len(linhas)} canais.")
            return redirect(url_for("parametros_tela"))
        if acao == "copart_vigente":
            v = (request.form.get("tabela") or "nova_40").strip()
            p["tabela_copart_vigente"] = v if v in magalu_real.TABELAS else "nova_40"
            _json_gravar(pasta("parametros.json"), p)
            flash(f"Tabela de coparticipação vigente: {magalu_real.NOMES[p['tabela_copart_vigente']]}.")
            return redirect(request.referrer or url_for("custo_full_real"))
        if acao == "prazos_full":
            # dois prazos por CD: o nosso manuseio + a transferência até o CD
            def _d(v):
                try:
                    return max(0, min(180, int(v or 0)))
                except ValueError:
                    return 0
            prazos = {}
            for k, v in request.form.items():
                if not k.startswith("man_"):
                    continue
                cd = k[4:].replace("__", " ")
                if not cd:
                    continue
                man = _d(v)
                tra = _d(request.form.get(f"tra_{k[4:]}"))
                if man or tra:
                    prazos[cd] = {"manuseio": man, "transferencia": tra}
            p["prazos_full"] = prazos
            p["prazo_full_padrao"] = {"manuseio": _d(request.form.get("man_padrao")),
                                      "transferencia": _d(request.form.get("tra_padrao"))}
            p["prazos_full_quando"] = agora().isoformat(); p["prazos_full_quem"] = session["usuario"]
            _json_gravar(pasta("parametros.json"), p)
            pad = p["prazo_full_padrao"]
            flash(f"Prazos de envio ao Full gravados: {len(prazos)} CD(s) · padrão "
                  f"{pad['manuseio']} + {pad['transferencia']} = {pad['manuseio'] + pad['transferencia']} dias.")
            return redirect(url_for("parametros_tela"))
        if acao in ("descricoes_arquivo", "custos_arquivo"):
            f = request.files.get("arquivo")
            if not f or not f.filename:
                flash("Escolha a planilha de cadastro (SKU · Descrição).")
                return redirect(url_for("parametros_tela"))
            try:
                nome = secure_filename(f.filename)
                cabec, linhas_iter, fechar = _tabela_leve(f, nome)
                cols = {}
                for i, c in enumerate(cabec):  # coluna repetida (ex.: dois "CUSTO TOTAL"): vale a primeira
                    cols.setdefault(unidecode_lower(c).replace(" ", ""), i)
                def col(*cands):  # prioridade = ordem dos candidatos, não a ordem do arquivo
                    return next((cols[k] for k in cands if k in cols), None)
                csku = col("codigofornecedor/skuinterno", "skuinterno", "sku", "modelo", "codigo", "cod", "produto", "codigoproduto", "referencia")
                cdes = col("titulosku", "descricao", "desc", "nomeproduto", "descricaoproduto", "titulo", "nome")
                cpeso = col("peso", "pesoproduto", "pesobruto", "pesokg", "peso(kg)", "pesoliquido")
                ccub = col("cubagem", "m3", "cubagemm3")
                ccus = col("custototal")  # "CUSTO" do Anymarket é preço, não custo
                ccat = col("categoria")
                if acao == "custos_arquivo" and ccus is None:
                    raise ValueError("essa não é a tabela de Custos MUL (falta a coluna CUSTO TOTAL) — suba-a no bloco Cadastro de SKUs")
                if acao == "descricoes_arquivo" and ccus is not None:
                    raise ValueError("essa é a tabela de Custos MUL — suba-a no bloco Tabela de Custos, ao lado")
                if csku is None or (cdes is None and cpeso is None):
                    raise ValueError("preciso de uma coluna SKU (ou Código) e uma Descrição (ou Nome) e/ou Peso")
                def cel(row, i):
                    v = row[i] if i is not None and i < len(row) else None
                    return "" if v is None else str(v).strip()
                skus = {}
                for row in linhas_iter:
                    k = cel(row, csku)
                    if not k:
                        continue
                    e = {}
                    if cdes is not None and cel(row, cdes):  # CustoProduto traz a descrição curta do ERP; a bonita vem do Anymarket
                        e["descricao_curta" if ccus is not None else "descricao"] = cel(row, cdes)
                    if ccat is not None and cel(row, ccat):
                        e["categoria"] = cel(row, ccat)
                    # PESO OFICIAL = o do CustoProduto (arquivo com CUSTO TOTAL). O peso do
                    # Anymarket fica guardado à parte (peso_any) e só vale se não houver o oficial.
                    for cc, kk in ((cpeso, "peso" if ccus is not None else "peso_any"), (ccub, "cubagem"), (ccus, "custo")):
                        if cc is not None and cel(row, cc):
                            try:
                                e[kk] = float(cel(row, cc).replace("kg", "").replace(",", ".").strip())
                            except ValueError:
                                pass
                    if e:
                        skus[k] = e
                fechar()
                cad = descricoes_ler()
                antes = len(cad["skus"])
                for k, e in skus.items():
                    cad["skus"].setdefault(k, {}).update(e)
                cad.update({"quando": agora().isoformat(), "quem": session["usuario"], "arquivo": nome})
                cad["custos" if ccus else "anymarket"] = {"quando": agora().isoformat(), "quem": session["usuario"], "arquivo": nome, "n": len(skus)}
                _json_gravar(pasta("cadastro", "descricoes.json"), cad)
                flash(f"Cadastro lido: {len(skus)} SKUs no arquivo · {len(cad['skus']) - antes} novos · cadastro com {len(cad['skus'])} SKUs.")
            except Exception as e:  # noqa: BLE001
                flash(f"Não consegui ler o cadastro: {e}")
            return redirect(url_for("parametros_tela"))
        if acao == "comissoes_arquivo":
            f = request.files.get("arquivo")
            if not f or not f.filename:
                flash("Escolha a planilha.")
                return redirect(url_for("parametros_tela"))
            import pandas as pd
            try:
                d = pd.read_excel(f, dtype=str).fillna("")
                cols = {re.sub(r"[^a-z]", "", str(c).lower().replace("ã", "a").replace("ç", "c").replace("ó", "o")): c for c in d.columns}
                def col(*nomes):
                    for nm in nomes:
                        if nm in cols:
                            return cols[nm]
                    return None
                cc, cn, cg, cm, ct, cti = col("codigo", "cod"), col("canal"), col("gestor"), col("comissao"), col("taxaextra", "taxa"), col("tipo")
                if not (cn and cm):
                    raise ValueError("preciso pelo menos das colunas Canal e Comissão")
                linhas = []
                for _, r in d.iterrows():
                    if not str(r[cn]).strip():
                        continue
                    linhas.append({"codigo": str(r[cc]).strip() if cc else "", "canal": str(r[cn]).strip(),
                                   "gestor": str(r[cg]).strip() if cg else "",
                                   "comissao": str(r[cm]).strip().replace("%", "").replace(".", ",") if "," not in str(r[cm]) else str(r[cm]).strip().replace("%", ""),
                                   "taxa": str(r[ct]).strip().replace("R$", "").strip() if ct else "",
                                   "tipo": str(r[cti]).strip() if cti else ""})
                p["comissoes"] = linhas
                p["comissoes_quando"] = agora().isoformat(); p["comissoes_quem"] = session["usuario"]
                _json_gravar(pasta("parametros.json"), p)
                flash(f"Tabela de comissões lida da planilha: {len(linhas)} canais.")
            except Exception as e:  # noqa: BLE001
                flash(f"Não consegui ler a planilha: {e}")
            return redirect(url_for("parametros_tela"))
        p["empresa"] = (request.form.get("empresa") or p["empresa"]).strip()
        try:
            p["tolerancia_comissao"] = float((request.form.get("tolerancia") or "0.5").replace(",", "."))
            p["tx_financeira"] = {**(p.get("tx_financeira") or {"madeira": 4.0, "webcont": 1.0}),
                                  "madeira": float((request.form.get("tx_madeira") or "4").replace("%", "").replace(",", ".")),
                                  "webcont": float((request.form.get("tx_webcont") or "1").replace("%", "").replace(",", "."))}
        except ValueError:
            pass
        _json_gravar(pasta("parametros.json"), p)
        for comp in competencias():
            recalcular_canais(comp)
        flash("Parâmetros gravados e rebates recalculados.")
        return redirect(url_for("parametros_tela"))
    # CDs que já apareceram nas cobranças do Full (para a tabela de prazos)
    cds_full = sorted({c["cd"] for c in full_magalu_ler()["cobrancas"].values() if c.get("cd")})
    return render_template("parametros.html", DESC=descricoes_ler(), CDS_FULL=cds_full)


# --------------------------------------------------------------------------
# saídas — planilha e JSON para o Tropa de Elite / ORION
# --------------------------------------------------------------------------
def _mlbs_filtrados(q: str, tipo: str) -> list[dict]:
    d = mlbs_ler()
    desc = descricoes_ler()["skus"]
    itens = []
    for m in d["mlbs"].values():
        m = dict(m)
        e = cadastro_sku(m["sku"], desc)
        m["descricao"] = e.get("descricao") or e.get("descricao_curta") or ""
        m["peso"] = e.get("peso") or e.get("peso_any")
        itens.append(m)
    if tipo and tipo != "todos":
        itens = [m for m in itens if unidecode_lower(m["tipo"]) == unidecode_lower(tipo)]
    if q:
        qs = [t for t in unidecode_lower(q).split() if t]
        itens = [m for m in itens if all(t in unidecode_lower(f"{m['mlb']} {m['sku']} {m['descricao']} {m['tipo']}") for t in qs)]
    itens.sort(key=lambda m: (m["sku"], m["mlb"]))
    return itens


def unidecode_lower(s) -> str:
    import unicodedata
    return "".join(ch for ch in unicodedata.normalize("NFKD", str(s or "")) if not unicodedata.combining(ch)).lower()


@app.route("/canal/meli/mlbs")
@logado
def mlbs():
    """Lista de MLB's — cada anúncio do Mercado Livre com o SKU, a descrição e o
    tipo de anúncio (Premium 16,5% · Clássico 11,5%). Nasce da Planilha 2."""
    q = (request.args.get("q") or "").strip()
    tipo = (request.args.get("tipo") or "todos").strip()
    d = mlbs_ler()
    todos = list(d["mlbs"].values())
    contagem = {"todos": len(todos)}
    for m in todos:
        contagem[m["tipo"]] = contagem.get(m["tipo"], 0) + 1
    itens = _mlbs_filtrados(q, tipo)
    tem_desc = bool(descricoes_ler()["skus"])
    return render_template("mlbs.html", c=canal_por_chave()["meli"], itens=itens, q=q, tipo=tipo,
                           contagem=contagem, ult=(d["uploads"][-1] if d["uploads"] else None), tem_desc=tem_desc,
                           desc_meta=descricoes_ler())


def _status_meli() -> dict:
    """pedido_mkt → (status, conta) vindos da Planilha 1 (rodadas do Mercado Livre)."""
    out = {}
    for cp in competencias():
        r = rodada("meli", cp)
        if r:
            for l in r["linhas"]:
                out[l["pedido_mkt"]] = (l.get("status") or "", l.get("conta") or "")
    return out


def _sumidos_meli() -> dict:
    """Pedidos que sumiram do relatório completo = cancelados pelo canal."""
    out = {}
    for cp in competencias():
        out.update(base_sumidos("meli", cp))
    return out


def _coletas_pedidos(so_validos: bool = True) -> list[dict]:
    """Pedidos de coletas = Frete Coletas > 0 na Planilha 2. 'Válido' = status
    Pago na Planilha 1; cancelado / devolvido fica fora. Sem Planilha 1 do mês,
    o status é desconhecido e o pedido entra (com aviso na tela).
    Pedido que sumiu do relatório completo é cancelamento do canal e fica fora
    sempre — inclusive da lista bruta."""
    st = _status_meli()
    sumidos = _sumidos_meli()
    out = []
    for k, p in resumo_meli_ler()["pedidos"].items():
        if (p.get("frete_coletas") or 0) <= 0 or k in sumidos:
            continue
        s_, conta = st.get(k, ("", ""))
        p = dict(p, pedido_mkt=k, status=s_ or "?", conta=conta)
        if so_validos and s_ and unidecode_lower(s_) != "pago":
            continue
        out.append(p)
    return out


def _soma(ps, k):
    return float(sum((p.get(k) or 0) for p in ps))


@app.route("/canal/magalu/fulfillment")
@logado
def fulfillment():
    """FULFILLMENT · Magalu — os pedidos com Modalidade de Entrega
    'Magalu entregas - Fulfillment'. Regra da Thaís (16/09/2026): nesses pedidos
    o Promob calcula a comissão SÓ sobre produto + IPI, não sobre o frete."""
    comp = comp_atual()
    r = rodada("magalu", comp)
    linhas = [l for l in (r["linhas"] if r else []) if l.get("fulfillment")]
    s = (r or {}).get("resumo") or {}
    ff = s.get("ff") or {}
    propria = s.get("propria") or {}
    por_dia: dict[str, dict] = {}
    for l in linhas:
        d = por_dia.setdefault(l["data"], {"dia": l["data"], "pedidos": 0, "venda": 0.0, "base": 0.0,
                                           "sistema": 0.0, "real": 0.0, "rebate": 0.0, "rebate_rs": 0.0})
        d["pedidos"] += 1
        d["venda"] += l.get("valor_prod") or 0.0
        d["base"] += l.get("sis_base") or 0.0
        sis = magalu.arred.sis_exato(l, True)
        d["sistema"] += sis
        d["real"] += l.get("tarifa") or 0.0
        d["rebate"] += sis - (l.get("tarifa") or 0.0)
        d["rebate_rs"] += l.get("rebate_rs") or 0.0
    dias = [dict(d, **{k: round(d[k], 2) for k in ("venda", "base", "sistema", "real", "rebate", "rebate_rs")})
            for d in sorted(por_dia.values(), key=lambda x: x["dia"])]
    piores = sorted(linhas, key=lambda l: (magalu.arred.sis_exato(l, True) - (l.get("tarifa") or 0.0)))[:40]
    # CUSTO TOTAL DO FULL PARA A MULTIMÓVEIS, por item
    fl = full_magalu_ler()
    cob = [c for c in fl["cobrancas"].values() if c.get("comp", comp) == comp]
    nomes = {}
    for i in vendas_magalu_ler()["itens"].values():
        if i.get("sku") and i.get("produto") and i["sku"] not in nomes:
            nomes[i["sku"]] = i["produto"]
    cub = cubagem_por_sku()
    cub_sku = {}
    for m in {c["sku"] for c in cob if c["sku"]} | set(copart_por_sku(comp)):
        v = cub.get(m) or cub.get(sku_chave(m))
        if v:
            cub_sku[m] = v
    custo = magalu_full.resumo(cob, copart_por_sku(comp), nomes, vendidos_full_sku(comp), cub_sku, comp)
    ult_full = fl["uploads"][-1] if fl["uploads"] else None
    return render_template("fulfillment.html", c=canal_por_chave()["magalu"], comp=comp, r=r,
                           ff=ff, propria=propria, dias=dias, piores=piores, n=len(linhas),
                           custo=custo, ult_full=ult_full)


def _custo_full(comp: str, q: str = "", cd: str = "todos") -> list[dict]:
    """Custo do Full por item: manuseio + armazenagem + tempo de estoque +
    coparticipação de frete + coleta rateada. Descrição e peso vêm do cadastro
    de SKUs (Parâmetros); o nome do anúncio, da Planilha 2 · Vendas."""
    fl = full_magalu_ler()
    cob = [c for c in fl["cobrancas"].values() if c.get("comp", comp) == comp]
    nomes = {}
    for i in vendas_magalu_ler()["itens"].values():
        if i.get("sku") and i.get("produto") and i["sku"] not in nomes:
            nomes[i["sku"]] = i["produto"]
    cub = cubagem_por_sku()
    cub_sku = {}   # aceita o SKU como vem do Magalu (sem ponto)
    for m in {c["sku"] for c in cob if c["sku"]} | set(copart_por_sku(comp)):
        v = cub.get(m) or cub.get(sku_chave(m))
        if v:
            cub_sku[m] = v
    s_ = magalu_full.resumo(cob, copart_por_sku(comp), nomes, vendidos_full_sku(comp), cub_sku, comp)
    desc = descricoes_ler()["skus"]
    itens = []
    for m in s_["por_sku"]:
        e = cadastro_sku(m["sku"], desc)
        m = dict(m)
        m["descricao"] = e.get("descricao") or e.get("descricao_curta") or m.get("produto") or ""
        m["peso"] = e.get("peso") or e.get("peso_any")
        # R$/kg é do CUSTO POR UNIDADE (o peso é de uma peça, não do lote)
        m["custo_kg"] = (m["por_unidade"] / m["peso"]) if (m.get("peso") and m.get("por_unidade")) else None
        m["custo_kg"] = round(m["custo_kg"], 2) if m["custo_kg"] else None
        m["cd"] = " · ".join(m.get("cds") or []) or "—"
        itens.append(m)
    if cd and cd != "todos":
        itens = [m for m in itens if cd in (m.get("cds") or [])]
    if q:
        qs = [t for t in unidecode_lower(q).split() if t]
        itens = [m for m in itens if all(t in unidecode_lower(f"{m['sku']} {m['descricao']} {m['cd']}") for t in qs)]
    itens.sort(key=lambda m: -m["custo_total"])
    for m in itens:
        m["_tarifa_m3"] = s_.get("tarifa_m3", 0.0)
        m["_m3_agendas"] = s_.get("m3_agendas", 0.0)
    return itens


def _amazon_faltantes(comp: str, q: str = "", filtro: str = "dentro") -> dict:
    """Pedidos do ERP (canal Amazon) do mês que NÃO apareceram em NENHUM
    relatório de transações subido — é a lista que a Gabi vai procurar.

    O pedido só entra no relatório de transações quando a Amazon PAGA. Por isso
    o faltante é separado em dois: o que está DENTRO da janela já coberta pelos
    relatórios (aí é buraco de verdade) e o que está FORA dela (só falta subir
    o arquivo daquele pedaço do mês)."""
    tx = amazon_tx_todas()
    vistos = set()
    de = ate = ""
    if tx is not None and len(tx):
        vistos = {str(p).strip() for p in tx["pedido"] if str(p).strip()}
        de, ate = str(tx["data"].min()), str(tx["data"].max())
    pct, _taxa = comissao_cadastrada(("AMAZON",), (0.105, 0.0))

    linhas, dentro, aguardando_l, fora, sem_id = [], [], [], [], 0
    for l in erp_linhas("amazon", comp):
        oc = (l.get("oc") or "").strip()
        if oc in vistos:
            continue
        tem_id = oc.startswith(("701-", "702-", "703-"))
        if not tem_id:
            sem_id += 1
        d = str(l.get("data") or "")
        # A transação de pagamento chega DEPOIS do pedido (medido: 2 a 17 dias).
        # Só é buraco de verdade quando a janela coberta vai além dessa folga.
        limite = ""
        if d:
            try:
                limite = str(date.fromisoformat(d) + timedelta(days=amazon.FOLGA_REPASSE))
            except ValueError:
                limite = d
        na_janela = bool(de and ate and de <= d <= ate and limite and limite <= ate)
        aguardando = bool(de and ate and de <= d <= ate and not na_janela)
        total = float(l.get("valor_total") or 0.0)
        m = {
            "oc": oc, "pedido_erp": l.get("pedido_erp", ""), "data": d,
            "nf": l.get("nf", ""), "data_nf": l.get("data_nf", ""), "status": l.get("status", ""),
            "cidade": l.get("cidade", ""), "uf": l.get("uf", ""), "cliente": l.get("cliente", ""),
            "valor_prod": float(l.get("valor_prod") or 0.0), "ipi": float(l.get("ipi") or 0.0),
            "frete": float(l.get("valor_frete") or 0.0), "total": total,
            "pct_erp": float(l.get("pct_comissao") or 0.0),
            "comissao_prevista": round(total * pct, 2),
            "tem_id": tem_id, "onde": "", "dias": ((agora().date() - date.fromisoformat(d)).days if d else 0),
        }
        m["onde"] = "dentro" if na_janela else ("aguardando" if aguardando else "fora")
        linhas.append(m)
        (dentro if na_janela else (aguardando_l if aguardando else fora)).append(m)

    sel = {"dentro": dentro, "aguardando": aguardando_l, "fora": fora, "todos": linhas}.get(filtro, dentro)
    if q:
        qs = [t for t in unidecode_lower(q).split() if t]
        sel = [m for m in sel
               if all(t in unidecode_lower(f"{m['oc']} {m['pedido_erp']} {m['nf']} {m['cidade']} {m['uf']} {m['cliente']}")
                      for t in qs)]
    sel.sort(key=lambda m: -m["total"])
    soma = lambda ls, k: round(sum(x[k] for x in ls), 2)  # noqa: E731
    return {
        "itens": sel, "de": de, "ate": ate, "pct": pct, "sem_id": sem_id,
        "folga": amazon.FOLGA_REPASSE,
        "contagem": {"dentro": len(dentro), "aguardando": len(aguardando_l), "fora": len(fora), "todos": len(linhas)},
        "erp_total": len(erp_linhas("amazon", comp)), "conferidos": len(erp_linhas("amazon", comp)) - len(linhas),
        "rs": {"dentro": soma(dentro, "total"), "aguardando": soma(aguardando_l, "total"),
               "fora": soma(fora, "total"), "todos": soma(linhas, "total")},
        "com": {"dentro": soma(dentro, "comissao_prevista"), "aguardando": soma(aguardando_l, "comissao_prevista"),
                "fora": soma(fora, "comissao_prevista"), "todos": soma(linhas, "comissao_prevista")},
    }


def _amazon_sem_erp(comp: str) -> list[dict]:
    """O contrário: pedido que veio no relatório da Amazon e não tem par no ERP."""
    r = rodada("amazon", comp)
    if not r:
        return []
    out = [dict(l) for l in r["linhas"] if not l.get("erp_ok")]
    out.sort(key=lambda l: -(l.get("sis_base") or 0))
    return out


@app.route("/canal/amazon/faltantes")
@logado
def amazon_faltantes():
    """PEDIDOS FALTANTES · Amazon — o que está no ERP e não apareceu em nenhum
    relatório de transações que a Gabi subiu."""
    comp = comp_atual()
    q = (request.args.get("q") or "").strip()
    filtro = (request.args.get("onde") or "dentro").strip()
    f = _amazon_faltantes(comp, q, filtro)
    sem_erp = _amazon_sem_erp(comp)
    return render_template("amazon_faltantes.html", c=canal_por_chave()["amazon"], comp=comp,
                           f=f, q=q, onde=filtro, sem_erp=sem_erp,
                           sem_erp_rs=round(sum((l.get("sis_base") or 0) for l in sem_erp), 2))


@app.route("/baixar/amazon-faltantes")
@logado
@exige("exportar")
def baixar_amazon_faltantes():
    q = (request.args.get("q") or "").strip()
    filtro = (request.args.get("onde") or "dentro").strip()
    comp = comp_atual()
    bio = planilhas.amazon_faltantes_xlsx(_amazon_faltantes(comp, q, filtro), _amazon_sem_erp(comp), comp)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_Amazon_Faltantes_{comp}_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def processar_magalu_full_real(destino: str, nome: str, quem: str) -> str:
    """CUSTO DO FULL POR ITEM · REAL — a tabela de coparticipação por SKU do
    canal. A tabela SUBSTITUI a anterior por inteiro (é uma tabela de preço,
    não um extrato que acumula)."""
    itens, diag = magalu_real.ler(destino)
    d = full_real_ler()
    d["itens"] = {i["sku"]: i for i in itens}
    d["uploads"].append({"nome": nome, "caminho": destino, "quando": agora().isoformat(), "quem": quem,
                         "diag": diag})
    d["uploads"] = d["uploads"][-30:]
    full_real_gravar(d)
    alerta = ""
    if diag["alta_nova"] > 20:
        alerta = (f" ATENÇÃO: a tabela NOVA é {diag['alta_nova']:.0f}% mais cara que a antiga "
                  f"(média R$ {f_brl(diag['media']['antiga'])} → R$ {f_brl(diag['media']['nova'])} por item); "
                  f"com o desconto de 40% a alta é de {diag['alta_40']:.0f}%. "
                  f"Escolha em Parâmetros qual tabela vale hoje.")
    return (f"Coparticipação do Full lida: {diag['skus']} SKUs. Média por item — antiga R$ "
            f"{f_brl(diag['media']['antiga'])} · nova −40% R$ {f_brl(diag['media']['nova_40'])} · "
            f"nova R$ {f_brl(diag['media']['nova'])}.{alerta}")


def full_real_ler() -> dict:
    """CUSTO DO FULL POR ITEM · REAL — o que o Magalu cobra de fato por SKU.
    Hoje a tela "Custo do Full por item · Estimado" chega ao custo por SKU
    RATEANDO o que vem por agenda/m³ (coleta) e por movimento (manuseio). Quando
    o canal manda o custo item a item, ele entra aqui e passa a ser o oficial —
    o Estimado vira conferência."""
    return _json_ler(pasta("magalu", "full_real.json"), {"itens": {}, "uploads": []})


def full_real_gravar(d):
    _json_gravar(pasta("magalu", "full_real.json"), d)


@app.route("/canal/magalu/custo-full-real")
@logado
def custo_full_real():
    """CUSTO DO FULL POR ITEM · REAL — oficial, quando o arquivo do canal traz
    o custo por SKU. Enquanto não houver arquivo, a tela explica e manda para o
    Estimado."""
    comp = comp_atual()
    d = full_real_ler()
    ult = d["uploads"][-1] if d["uploads"] else None
    vig = (request.args.get("tab") or parametros().get("tabela_copart_vigente") or "nova_40").strip()
    if vig not in magalu_real.TABELAS:
        vig = "nova_40"
    itens = list(d["itens"].values())
    s = {}
    est_copart = est_total = 0.0
    if itens:
        desc = descricoes_ler()["skus"]
        nomes = {k: (v.get("descricao") or v.get("descricao_curta") or "") for k, v in desc.items()}
        s = magalu_real.resumo(itens, vendidos_full_sku(comp), nomes, vig, sku_chave)
        est = _custo_full(comp, "", "todos")
        est_copart = round(sum(m["copart"] for m in est), 2)
        est_total = round(sum(m["custo_total"] for m in est), 2)
    return render_template("custo_full_real.html", c=canal_por_chave()["magalu"], comp=comp,
                           s=s, ult=ult, tem=bool(itens), vigente=vig,
                           TABELAS=magalu_real.TABELAS, NOMES=magalu_real.NOMES,
                           est_copart=est_copart, est_total=est_total,
                           pode_gravar=(session.get("papel") in PODE["parametros"]))


# ---------------------------------------------------------------------------
# ENTREGAS BY PLUTOS — o custo de ENTREGA por SKU e por canal, num lugar só.
# É ESTE o arquivo que sobe no ORION: uma linha por SKU × canal, com o custo
# por unidade que entra na formação do preço. Em vez de o ORION juntar dois
# ou três arquivos e ter de saber qual coluna usar de cada um, o PLUTOS
# entrega a conta pronta e diz de onde veio cada pedaço.
# ---------------------------------------------------------------------------
def entregas_por_sku(comp: str) -> dict:
    """Consolida o custo de entrega por SKU de TODOS os canais.

    MAGALU · FULL   custo do Full por unidade (manuseio + armazenagem + tempo de
                    estoque + coleta rateada, do que o canal COBROU no mês)
                    + coparticipação de frete da TABELA vigente (é o que vai ser
                    cobrado na próxima venda — precificação usa tabela, não o
                    realizado do mês passado).
    MELI · COLETAS  custo do Coletas por anúncio, do último pedido válido.
    Shopee Full e Amazon FBA entram aqui quando os arquivos chegarem."""
    desc = descricoes_ler()["skus"]
    nome_de = lambda k: ((desc.get(k) or {}).get("descricao")  # noqa: E731
                         or (desc.get(k) or {}).get("descricao_curta") or "")
    linhas: list[dict] = []

    # --- MAGALU · FULL -----------------------------------------------------
    tab = full_real_ler()["itens"]
    vig = (parametros().get("tabela_copart_vigente") or "nova_40")
    if vig not in magalu_real.TABELAS:
        vig = "nova_40"
    tab_idx = {sku_chave(k): v for k, v in tab.items()}
    for m in _custo_full(comp, "", "todos"):
        if not m.get("unidades"):
            continue
        t = tab_idx.get(sku_chave(m["sku"]))
        copart_tab = float(t[vig]) if t else None
        copart_un = copart_tab if copart_tab is not None else float(m.get("copart_un") or 0.0)
        full_un = float(m.get("full_un") or 0.0)
        linhas.append({
            "sku": m["sku"], "descricao": m.get("descricao") or nome_de(m["sku"]),
            "canal": "Magazine Luiza", "modalidade": "Full",
            "custo_un": round(full_un + copart_un, 2),
            "logistica_un": round(full_un, 2), "frete_un": round(copart_un, 2),
            "fonte_frete": ("tabela do canal · " + magalu_real.NOMES[vig]) if t else "medido no extrato",
            "unidades": m.get("unidades"), "peso": m.get("peso"), "cubagem": m.get("cubagem"),
            "manuseio_un": round(float(m.get("unit_manuseio") or 0.0), 2),
            "coleta_un": round(float(m.get("coleta_tabela_un") or 0.0), 2),
            "fonte": "Custo do Full cobrado + coparticipação",
        })

    # --- MERCADO LIVRE · COLETAS ------------------------------------------
    for m in _custo_coletas("", "todos"):
        if not m.get("sku"):
            continue
        linhas.append({
            "sku": m["sku"], "descricao": m.get("descricao") or nome_de(m["sku"]),
            "canal": "Mercado Livre", "modalidade": "Coletas",
            "custo_un": round(float(m.get("custo") or 0.0), 2),
            "logistica_un": 0.0, "frete_un": round(float(m.get("custo") or 0.0), 2),
            "fonte_frete": "último pedido válido", "unidades": m.get("pedidos"),
            "peso": m.get("peso"), "cubagem": None, "manuseio_un": 0.0, "coleta_un": 0.0,
            "mlb": m.get("mlb"), "tipo": m.get("tipo"),
            "fonte": "Custo coletas por anúncio",
        })

    # CONFIANÇA: custo por unidade calculado sobre 1 ou 2 unidades é ruído —
    # normalmente é item parado, que pagou armazenagem o mês todo e vendeu uma
    # peça. Para a precificação isso distorce; a linha vai marcada.
    for l in linhas:
        un = float(l.get("unidades") or 0)
        l["confianca"] = "ok" if un >= 5 else ("pouca base" if un >= 1 else "sem base")
        l["pouca_base"] = l["confianca"] != "ok"
    linhas.sort(key=lambda l: (l["sku"], l["canal"]))
    # pivô SKU × canal, do jeito que o ORION consome
    piv: dict[str, dict] = {}
    for l in linhas:
        p = piv.setdefault(l["sku"], {"sku": l["sku"], "descricao": l["descricao"],
                                      "canais": {}, "unidades": {}, "confianca": {}})
        ch = f"{l['canal']} · {l['modalidade']}"
        p["canais"][ch] = l["custo_un"]
        p["unidades"][ch] = l.get("unidades") or 0
        p["confianca"][ch] = l.get("confianca")
        if not p["descricao"]:
            p["descricao"] = l["descricao"]
    for p in piv.values():
        v = list(p["canais"].values())
        p["menor"] = round(min(v), 2) if v else 0.0
        p["maior"] = round(max(v), 2) if v else 0.0
        p["spread"] = round(p["maior"] - p["menor"], 2)
        p["spread_pct"] = (round(100 * p["spread"] / p["menor"], 1) if p["menor"] else 0.0)
        p["n_canais"] = len(v)
        p["mais_barato"] = min(p["canais"], key=p["canais"].get) if v else ""
        p["mais_caro"] = max(p["canais"], key=p["canais"].get) if v else ""
    colunas = sorted({f"{l['canal']} · {l['modalidade']}" for l in linhas})
    firmes = [l for l in linhas if not l["pouca_base"]]
    piv_l = sorted(piv.values(), key=lambda p: -p["spread"])
    comparativo = [p for p in piv_l if p["n_canais"] > 1]
    return {"linhas": linhas, "pivo": piv_l, "comparativo": comparativo,
            "comparativo_n": len(comparativo),
            "comparativo_rs": round(sum(p["spread"] for p in comparativo), 2),
            "colunas": colunas, "comp": comp, "vigente": vig,
            "skus": len({l["sku"] for l in linhas}),
            "canais": len(colunas),
            "pouca_base": sum(1 for l in linhas if l["pouca_base"]),
            "media": (round(sum(l["custo_un"] for l in firmes) / len(firmes), 2) if firmes else 0.0),
            "media_todas": (round(sum(l["custo_un"] for l in linhas) / len(linhas), 2) if linhas else 0.0)}


@app.route("/entregas")
@logado
def entregas():
    """ENTREGAS BY PLUTOS — o custo de entrega por SKU, canal a canal."""
    comp = comp_atual()
    e = entregas_por_sku(comp)
    q = (request.args.get("q") or "").strip()
    if q:
        qs = [t for t in unidecode_lower(q).split() if t]
        e["linhas"] = [l for l in e["linhas"]
                       if all(t in unidecode_lower(f"{l['sku']} {l['descricao']} {l['canal']} {l['modalidade']}") for t in qs)]
        e["pivo"] = [p for p in e["pivo"]
                     if all(t in unidecode_lower(f"{p['sku']} {p['descricao']}") for t in qs)]
    return render_template("entregas.html", c={"nome": "Entregas"}, comp=comp, e=e, q=q,
                           NOMES=magalu_real.NOMES)


@app.route("/baixar/entregas")
@logado
@exige("exportar")
def baixar_entregas():
    comp = comp_atual()
    bio = planilhas.entregas_xlsx(entregas_por_sku(comp), comp, magalu_real.NOMES)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_ENTREGAS_{comp}_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/entregas/<comp>.json")
def api_entregas(comp):
    """Mesmo conteúdo do Excel, para quem quiser puxar direto (ORION)."""
    if (request.args.get("token") or "") != (os.environ.get("PLUTOS_TOKEN") or "plutos"):
        abort(403)
    e = entregas_por_sku(comp)
    return jsonify({"competencia": comp, "gerado": agora().isoformat(), "versao": VERSAO,
                    "tabela_copart": e["vigente"], "itens": e["linhas"]})


@app.route("/baixar/copart-tabela")
@logado
@exige("exportar")
def baixar_copart_tabela():
    comp = comp_atual()
    d = full_real_ler()
    vig = (request.args.get("tab") or parametros().get("tabela_copart_vigente") or "nova_40").strip()
    if vig not in magalu_real.TABELAS:
        vig = "nova_40"
    desc = descricoes_ler()["skus"]
    nomes = {k: (v.get("descricao") or v.get("descricao_curta") or "") for k, v in desc.items()}
    s_ = magalu_real.resumo(list(d["itens"].values()), vendidos_full_sku(comp), nomes, vig, sku_chave)
    bio = planilhas.copart_tabela_xlsx(s_, comp, magalu_real.NOMES)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_Coparticipacao_Tabela_{comp}_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/canal/magalu/custo-full")
@logado
def custo_full():
    """CUSTO DO FULL POR ITEM — tudo que o Fulfillment cobra da Multimóveis,
    SKU a SKU, com busca e filtro por CD (mesmo padrão do Custo coletas)."""
    comp = comp_atual()
    q = (request.args.get("q") or "").strip()
    cd = (request.args.get("cd") or "todos").strip()
    todos = _custo_full(comp)
    contagem = {"todos": len(todos)}
    for m in todos:
        for x in (m.get("cds") or []):
            contagem[x] = contagem.get(x, 0) + 1
    itens = _custo_full(comp, q, cd)
    tot = {k: round(sum(m[k] for m in itens), 2) for k in
           ("manuseio", "armazenagem", "tempo_estoque", "copart", "coleta_rateio", "custo_total", "custo_full")}
    tot["qtd"] = round(sum(m["qtd"] for m in itens), 0)
    tot["unidades"] = round(sum(m["unidades"] for m in itens), 0)
    tot["full_un"] = round(tot["custo_full"] / tot["unidades"], 2) if tot["unidades"] else 0.0
    tot["copart_un"] = round(tot["copart"] / tot["unidades"], 2) if tot["unidades"] else 0.0
    tot["por_unidade"] = round(tot["custo_total"] / tot["unidades"], 2) if tot["unidades"] else 0.0
    sem_un = [m["sku"] for m in itens if not m["unidades"] and m["custo_total"]]
    sem_cub = [m["sku"] for m in itens if not m.get("cubagem")]
    tot["tarifa_m3"] = itens[0]["_tarifa_m3"] if itens else 0.0
    tot["volume"] = round(sum(m.get("volume") or 0 for m in itens), 2)
    cobs = [c for c in full_magalu_ler()["cobrancas"].values() if c.get("comp", comp) == comp]
    conc = {"por_pedido": {}, "por_cobranca": {}, "total": 0.0}
    for c in cobs:
        if c["tipo"] != "copart":
            continue
        cp = c.get("comp_pedido") or "?"
        mc = (c.get("data_cobranca") or "?")[:7]
        conc["por_pedido"][cp] = round(conc["por_pedido"].get(cp, 0.0) + c["valor"], 2)
        conc["por_cobranca"][mc] = round(conc["por_cobranca"].get(mc, 0.0) + c["valor"], 2)
        conc["total"] = round(conc["total"] + c["valor"], 2)
    conc["por_pedido"] = dict(sorted(conc["por_pedido"].items()))
    conc["por_cobranca"] = dict(sorted(conc["por_cobranca"].items()))
    dts = sorted({c["data"] for c in full_magalu_ler()["cobrancas"].values()
                  if c.get("comp", comp) == comp and c.get("data")})
    periodo = f"{f_dia(dts[0])} a {f_dia(dts[-1])}" if dts else ""
    return render_template("custo_full.html", c=canal_por_chave()["magalu"], comp=comp, itens=itens,
                           q=q, cd=cd, contagem=contagem, tot=tot, periodo=periodo, sem_un=sem_un, sem_cub=sem_cub, conc=conc,
                           sem_peso=sum(1 for m in todos if not m.get("peso")),
                           cds=sorted(k for k in contagem if k != "todos"))


@app.route("/baixar/custo-full")
@logado
@exige("exportar")
def baixar_custo_full():
    q = (request.args.get("q") or "").strip()
    cd = (request.args.get("cd") or "todos").strip()
    bio = planilhas.custo_full_xlsx(_custo_full(comp_atual(), q, cd), comp_atual(), q, cd)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_CustoFull_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/canal/magalu/estoque-full")
@logado
def estoque_full():
    """ESTOQUE FULL — o que está guardado no CD do Magalu hoje, quanto vale,
    quanto está custando de armazenagem e o que está parado pagando aniversário.
    Nasce da foto diária de estoque que vem nas Cobranças do Fulfillment."""
    comp = comp_atual()
    fl = full_magalu_ler()
    cob = [c for c in fl["cobrancas"].values() if c.get("comp", comp) == comp]
    vend = vendidos_full_sku(comp)
    dias_venda = len({i["data"] for i in vendas_magalu_ler()["itens"].values()
                      if i.get("competencia") == comp}) or 1
    cad = descricoes_ler()["skus"]
    cadastro = {}
    for k in {c["sku"] for c in cob if c["sku"]}:
        cadastro[k] = cadastro_sku(k, cad)
    # tarifa do m³ da coleta e a faixa de manuseio, para custear a sugestão
    m3 = round(sum(c.get("espaco") or 0.0 for c in cob if c["tipo"] == "coleta"), 4)
    coleta_rs = round(sum(c["valor"] for c in cob if c["tipo"] == "coleta"), 2)
    tarifa_m3 = round(coleta_rs / m3, 2) if m3 else 0.0
    unit_man = {}
    for c in cob:
        if c["tipo"] == "manuseio" and c["sku"] and c.get("unit"):
            unit_man[c["sku"]] = c["unit"]
    try:
        dias_alvo = max(7, min(120, int(request.args.get("dias") or 30)))
    except ValueError:
        dias_alvo = 30
    par = parametros()
    prazos_cad = par.get("prazos_full") or {}
    prazos = {k: prazo_dias(v) for k, v in prazos_cad.items()}
    prazo_padrao = prazo_dias(par.get("prazo_full_padrao"))
    e = magalu_full.estoque(cob, vend, dias_venda, cadastro, dias_alvo, tarifa_m3, unit_man,
                            prazos, prazo_padrao)
    q = (request.args.get("q") or "").strip()
    cd_sel = (request.args.get("cd") or "todos").strip()
    cds = sorted({x["cd"] for x in e["por_cd"] if x.get("cd")})
    por_cd = cd_sel != "todos"
    itens = [x for x in e["por_cd"] if x["cd"] == cd_sel] if por_cd else e["por_sku"]
    if q:
        qs = [t for t in unidecode_lower(q).split() if t]
        itens = [m for m in itens
                 if all(t in unidecode_lower(f"{m['sku']} {m.get('descricao', '')} {m.get('cd', '')}") for t in qs)]
    # o que a tela soma no rodapé muda quando o filtro é de um CD só
    tot = {"enviar_un": round(sum(m["sugestao"] for m in itens), 0),
           "estoque": round(sum(m["estoque"] for m in itens), 0),
           "m3_envio": round(sum((m.get("sug_m3") or 0) for m in itens), 2)}
    return render_template("estoque_full.html", c=canal_por_chave()["magalu"], comp=comp, e=e,
                           itens=itens, q=q, dias_venda=dias_venda, dias_alvo=dias_alvo,
                           cd_sel=cd_sel, cds=cds, por_cd=por_cd, tot=tot,
                           tarifa_m3=tarifa_m3, prazos=prazos, prazo_padrao=prazo_padrao,
                           prazos_cad=prazos_cad)


@app.route("/baixar/sugestao-full")
@logado
@exige("exportar")
def baixar_sugestao_full():
    """A lista de envio para o Full, do jeito que está na tela."""
    comp = comp_atual()
    fl = full_magalu_ler()
    cob = [c for c in fl["cobrancas"].values() if c.get("comp", comp) == comp]
    vend = vendidos_full_sku(comp)
    dias_venda = len({i["data"] for i in vendas_magalu_ler()["itens"].values()
                      if i.get("competencia") == comp}) or 1
    cad = descricoes_ler()["skus"]
    cadastro = {k: cadastro_sku(k, cad) for k in {c["sku"] for c in cob if c["sku"]}}
    m3 = round(sum(c.get("espaco") or 0.0 for c in cob if c["tipo"] == "coleta"), 4)
    coleta_rs = round(sum(c["valor"] for c in cob if c["tipo"] == "coleta"), 2)
    tarifa = round(coleta_rs / m3, 2) if m3 else 0.0
    unit_man = {c["sku"]: c["unit"] for c in cob if c["tipo"] == "manuseio" and c["sku"] and c.get("unit")}
    try:
        dias_alvo = max(7, min(120, int(request.args.get("dias") or 30)))
    except ValueError:
        dias_alvo = 30
    par = parametros()
    e = magalu_full.estoque(cob, vend, dias_venda, cadastro, dias_alvo, tarifa, unit_man,
                            {k: prazo_dias(v) for k, v in (par.get("prazos_full") or {}).items()},
                            prazo_dias(par.get("prazo_full_padrao")))
    bio = planilhas.sugestao_full_xlsx(e, comp, dias_alvo)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_EnvioFull_{dias_alvo}d_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/canal/magalu/frete-full")
@logado
def frete_full():
    """CUSTO FRETE FULL — o que o frete dos pedidos Fulfillment está custando.
    Nasce da Planilha 2 (Vendas no período), aba de PACOTES."""
    comp = comp_atual()
    v = vendas_magalu_ler()
    pacotes = list(v["pacotes"].values())
    s = magalu_vendas.resumo_frete_full(pacotes, comp)
    comps = sorted({p["competencia"] for p in pacotes})
    ult = v["uploads"][-1] if v["uploads"] else None
    return render_template("frete_full.html", c=canal_por_chave()["magalu"], comp=comp, s=s,
                           comps=comps, ult=ult, tem=bool(pacotes))


@app.route("/canal/magalu/coparticipacao")
@logado
def coparticipacao():
    """COPARTICIPAÇÃO DE FRETE — o pedaço do frete que NÓS pagamos no Magalu.
    Nasce da Planilha 2 (Vendas no período), aba de PEDIDOS/itens."""
    comp = comp_atual()
    v = vendas_magalu_ler()
    itens = list(v["itens"].values())
    pacotes = list(v["pacotes"].values())
    s = magalu_vendas.resumo_copart(itens, pacotes, comp)
    comps = sorted({i["competencia"] for i in itens})
    ult = v["uploads"][-1] if v["uploads"] else None
    return render_template("coparticipacao.html", c=canal_por_chave()["magalu"], comp=comp, s=s,
                           comps=comps, ult=ult, tem=bool(itens))


@app.route("/canal/meli/coletas")
@logado
def coletas():
    """Coletas · total — os pedidos do Mercado Livre entregues pelo Coletas
    (frete coletas > 0 na Planilha 2), por competência."""
    comp = comp_atual()
    todos = resumo_meli_ler()["pedidos"]
    col = _coletas_pedidos()
    brutos = _coletas_pedidos(so_validos=False)
    fora = {}
    for p in brutos:
        if p["status"] != "?" and unidecode_lower(p["status"]) != "pago":
            fora[p["competencia"]] = fora.get(p["competencia"], 0) + 1
    sem_status = sum(1 for p in brutos if p["status"] == "?" and p["competencia"] == comp)
    comps = sorted({p["competencia"] for p in col} | {p["competencia"] for p in todos.values()})
    por_comp = []
    st = _status_meli()
    for cp in comps:
        pc = [p for p in col if p["competencia"] == cp]
        tm = [p for k, p in todos.items() if p["competencia"] == cp and unidecode_lower(st.get(k, ("", ""))[0] or "pago") == "pago"]
        venda_meli = _soma(tm, "valor_prod")
        por_comp.append({"comp": cp, "pedidos": len(pc), "meli": len(tm), "share": (len(pc) / len(tm) if tm else 0),
                         "venda": _soma(pc, "valor_prod"), "venda_meli": venda_meli,
                         "share_venda": (_soma(pc, "valor_prod") / venda_meli if venda_meli else 0), "frete": _soma(pc, "frete"), "coletas": _soma(pc, "frete_coletas"),
                         "cupom": _soma(pc, "cupom_canal"), "rebate_com": _soma(pc, "rebate_bi"), "com_liq": _soma(pc, "com_liq")})
    atual = next((x for x in por_comp if x["comp"] == comp), None) or {"comp": comp, "pedidos": 0, "meli": 0, "share": 0, "venda": 0, "venda_meli": 0, "share_venda": 0, "frete": 0, "coletas": 0, "cupom": 0, "rebate_com": 0, "com_liq": 0}
    desc = descricoes_ler()["skus"]
    por_mlb: dict[str, dict] = {}
    for p in col:
        if p["competencia"] != comp:
            continue
        m = por_mlb.setdefault(p["anuncio"], {"mlb": p["anuncio"], "sku": p["sku"], "tipo": p["tipo"], "pedidos": 0, "venda": 0.0, "coletas": 0.0, "cupom": 0.0, "rebate_com": 0.0})
        m["pedidos"] += 1; m["venda"] += p["valor_prod"]; m["coletas"] += p["frete_coletas"]; m["cupom"] += p["cupom_canal"]; m["rebate_com"] += p["rebate_bi"]
    linhas = sorted(por_mlb.values(), key=lambda m: -m["coletas"])
    for m in linhas:
        e = desc.get(m["sku"]) or {}
        m["descricao"] = e.get("descricao") or e.get("descricao_curta") or ""
        m["custo_medio"] = m["coletas"] / m["pedidos"] if m["pedidos"] else 0
    sumidos = _sumidos_meli()
    canc = [p for k, p in sumidos.items() if (todos.get(k, {}).get("frete_coletas") or 0) > 0
            and todos.get(k, {}).get("competencia") == comp]
    return render_template("coletas.html", c=canal_por_chave()["meli"], comp=comp, atual=atual, por_comp=por_comp, linhas=linhas,
                           COMPS_COLETAS=comps, fora=fora.get(comp, 0), sem_status=sem_status,
                           cancelados=len(canc), cancelados_rs=round(sum(float(p.get("valor_prod") or 0) for p in canc), 2))


def _custo_coletas(q: str, tipo: str) -> list[dict]:
    """Custo coletas por MLB: o último pedido válido (frete coletas > 0) de cada
    anúncio dá o custo; descrição e peso vêm do cadastro de SKUs."""
    desc = descricoes_ler()["skus"]
    ult: dict[str, dict] = {}
    for p in sorted(_coletas_pedidos(), key=lambda p: p["data"]):
        m = ult.get(p["anuncio"])
        if m is None:
            m = ult[p["anuncio"]] = {"mlb": p["anuncio"], "pedidos": 0}
        m.update({"sku": p["sku"], "tipo": p["tipo"], "custo": p["frete_coletas"], "data": p["data"], "pedido": p["pedido_mkt"],
                  "valor_prod": p["valor_prod"]})
        m["pedidos"] += 1
    itens = []
    for m in ult.values():
        e = desc.get(m["sku"]) or desc.get(m["sku"].upper()) or {}
        m["descricao"] = e.get("descricao") or e.get("descricao_curta") or ""
        m["peso"] = e.get("peso") or e.get("peso_any")
        m["custo_kg"] = (m["custo"] / m["peso"]) if m.get("peso") else None
        itens.append(m)
    if tipo and tipo != "todos":
        itens = [m for m in itens if unidecode_lower(m["tipo"]) == unidecode_lower(tipo)]
    if q:
        qs = [t for t in unidecode_lower(q).split() if t]
        itens = [m for m in itens if all(t in unidecode_lower(f"{m['mlb']} {m['sku']} {m['descricao']} {m['tipo']}") for t in qs)]
    itens.sort(key=lambda m: (m["sku"], m["mlb"]))
    return itens


@app.route("/canal/meli/custo-coletas")
@logado
def custo_coletas():
    q = (request.args.get("q") or "").strip()
    tipo = (request.args.get("tipo") or "todos").strip()
    todos = _custo_coletas("", "todos")
    contagem = {"todos": len(todos)}
    for m in todos:
        contagem[m["tipo"]] = contagem.get(m["tipo"], 0) + 1
    itens = _custo_coletas(q, tipo)
    return render_template("custo_coletas.html", c=canal_por_chave()["meli"], itens=itens, q=q, tipo=tipo, contagem=contagem,
                           tem_desc=bool(descricoes_ler()["skus"]), sem_peso=sum(1 for m in todos if not m.get("peso")))


@app.route("/baixar/custo-coletas")
@logado
@exige("exportar")
def baixar_custo_coletas():
    q = (request.args.get("q") or "").strip()
    tipo = (request.args.get("tipo") or "todos").strip()
    bio = planilhas.custo_coletas_xlsx(_custo_coletas(q, tipo), q, tipo)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_CustoColetas_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/baixar/coletas")
@logado
@exige("exportar")
def baixar_coletas():
    comp = comp_atual()
    col = [p for p in _coletas_pedidos() if p["competencia"] == comp]
    desc = descricoes_ler()["skus"]
    bio = planilhas.coletas_xlsx(col, desc, comp)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_Coletas_{comp.replace('-', '')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/baixar/mlbs")
@logado
@exige("exportar")
def baixar_mlbs():
    q = (request.args.get("q") or "").strip()
    tipo = (request.args.get("tipo") or "todos").strip()
    bio = planilhas.mlbs_xlsx(_mlbs_filtrados(q, tipo), q, tipo)
    suf = "" if tipo == "todos" else "_" + secure_filename(tipo)
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_ListaMLBs{suf}_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/baixar/shopee-gabi")
@logado
@exige("exportar")
def baixar_shopee_gabi():
    """A planilha da Shopee no formato da Gabi (aba novo + Planilha1 com as
    fórmulas dela), já preenchida com o que o PLUTOS gerou no mês."""
    comp = comp_atual()
    r = rodada("shopee", comp)
    if not _resumo_pronto(r):
        flash("Ainda não há rodada da Shopee nesta competência.")
        return redirect(url_for("canal", chave="shopee", mes=comp))
    pct, _ = comissao_cadastrada(("SHOPEE",), (0.12, 12.0))
    bio = planilhas.shopee_gabi_xlsx(r["linhas"], comp, pct, agora())
    del r
    gc.collect()
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_SHOPEE_{comp.replace('-', '')}_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/baixar/<qual>")
@logado
@exige("exportar")
def baixar(qual):
    comp = comp_atual()
    rods = {c["chave"]: rodada(c["chave"], comp) for c in canais()}
    rods = {k: v for k, v in rods.items() if v}
    if not rods and not qual.startswith("linha_"):
        flash("Não há rodada nesta competência.")
        return redirect(url_for("arquivos"))
    if qual == "rebates":
        bio = planilhas.rebates_xlsx(rods, comp, canal_por_chave(), parametros())
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_Rebates_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if qual.startswith("linha_"):
        chave = qual[6:]
        r = rods.get(chave)
        idx = erp_ler()["ocs"]
        if r:
            linhas = []
            cad = descricoes_ler()["skus"]
            base = erp_por_base(chave) if chave == "magalu" else {}
            for l in r["linhas"]:
                e = idx.get(l["pedido_mkt"]) or idx.get(l["pedido_canal"]) or base.get(l["pedido_mkt"]) or {}
                m = dict(l); m["descricao"] = descricao_de(l.get("sku"), cad)
                m.update({k: (e.get(k[4:]) if e else None) for _, k, _ in ERP_COLS}); linhas.append(m)
            bio = planilhas.linha_xlsx({"linhas": linhas}, LINHA_COLS.get(chave, []) + ERP_COLS, canal_por_chave()[chave]["nome"], comp)
        else:
            cols_so, linhas_so = ERP_SO_COLS, erp_linhas(chave, comp)
            if _tem_tx_fin(chave):
                linhas_so = separar_tx_financeira(chave, [dict(l) for l in linhas_so])
                i = next(i for i, (_, k, _) in enumerate(cols_so) if k == "comissao_erp_rs") + 1
                cols_so = cols_so[:i] + TX_FIN_COLS + cols_so[i:]
            bio = planilhas.linha_xlsx({"linhas": linhas_so}, cols_so, canal_por_chave()[chave]["nome"], comp)
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_LinhaALinha_{chave.upper()}_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if qual == "faltante":
        bio = planilhas.faltante_xlsx(rods.get("meli"), faltante_ler("meli"), comp)
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_FaltanteCampanha_Meli_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    abort(404)


# "Frete cobrado pelo canal" = o frete que o CLIENTE pagou ao canal (receita de frete), definição da Thaís 12/09.
# Só entra quando o arquivo do canal traz esse valor: Meli = "Frete Pedido" (Tabela Geral).
# Magalu (Financeiro por período) NÃO traz o frete pago pelo cliente — só "Custos logísticos" (o que o Magalu
# cobra de nós) e a coparticipação → fica EM BRANCO. Shopee/Madeira/Webcont idem.
FRETE_CANAL = {"meli": "frete"}


def frete_cobrado_canal(chave: str, l: dict):
    """Frete pago pelo cliente ao canal, em R$, quando o canal informa; senão None → célula
    EM BRANCO no export (branco ≠ zero: o Tropa usa a NF/CT-e quando está em branco)."""
    campo = FRETE_CANAL.get(chave)
    if not campo:
        return None
    v = l.get(campo)
    return None if v is None else round(float(v), 2)


@app.route("/export-rebates")
@logado
@exige("exportar")
def export_rebates():
    """EXPORT REBATES — o apanhado de TUDO que o PLUTOS gerou, em todos os canais
    e competências, no formato que os outros apps importam:
    OC · Data · Canal · Rebate R$ · Rebate comissão · Rebate frete · Total."""
    linhas = []
    for comp in competencias():
        for c in canais():
            r = rodada(c["chave"], comp)
            if not r:
                continue
            linhas_r = r["linhas"]
            r = None
            for l in linhas_r:
                base = l.get("valor_prod") or 0.0
                linhas.append({"oc": l["pedido_mkt"], "data": l["data"], "canal": c["nome"],
                               "rebate_rs": l["rebate_rs"], "rebate_comissao": l["rebate_comissao"],
                               "rebate_frete": l["rebate_frete"], "rebate_total": l["rebate_total"],
                               "sis_rs": l.get("sis_rs"), "tarifa": l.get("tarifa"),
                               "frete_canal": frete_cobrado_canal(c["chave"], l),
                               "pct_real": ((l.get("tarifa") or 0.0) / base if base else None), "venda": base,
                               "competencia": comp, "sku": l.get("sku", ""), "pedido_canal": l.get("pedido_canal", ""),
                               "faltante_status": l.get("faltante_status", "")})
            linhas_r = None
            gc.collect()          # solta a rodada inteira antes de abrir a próxima
    if not linhas:
        flash("Nada gerado ainda para exportar.")
        return redirect(url_for("arquivos"))
    linhas.sort(key=lambda x: (x["data"], x["canal"], x["oc"]))
    bio = planilhas.export_rebates_xlsx(linhas, agora())
    return send_file(bio, as_attachment=True,
                     download_name=f"PLUTOS_EXPORT_REBATES_{agora().strftime('%d%m%Y_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/rebates/<comp>.json")
def api_rebates(comp):
    """O que o Tropa de Elite e o ORION consomem: rebate por pedido, já
    classificado nas três formas. Chave = pedido do marketplace (a OC do ERP)."""
    token = os.environ.get("PLUTOS_TOKEN")
    if token and request.args.get("token") != token and not session.get("usuario"):
        abort(403)
    saida = []
    for c in canais():
        r = rodada(c["chave"], comp)
        if not r:
            continue
        for l in r["linhas"]:
            saida.append({"canal": l["canal"], "oc": l["pedido_mkt"], "pedido_canal": l["pedido_canal"],
                          "pedido_any": l["pedido_any"], "data": l["data"], "sku": l["sku"],
                          "rebate_rs": l["rebate_rs"], "rebate_comissao": l["rebate_comissao"],
                          "rebate_frete": l["rebate_frete"], "rebate_total": l["rebate_total"],
                          "comissao_sistema": l.get("sis_rs"), "comissao_real": l.get("tarifa"), "base_venda": l.get("valor_prod"),
                          "frete_canal": frete_cobrado_canal(c["chave"], l),
                          "faltante_status": l["faltante_status"]})
    return jsonify({"competencia": comp, "gerado": agora().isoformat(), "versao": VERSAO,
                    "pedidos": len(saida), "linhas": saida})


@app.errorhandler(500)
@app.errorhandler(Exception)
def erro_500(e):
    """Em vez do 'Internal Server Error' cru: uma página que diz o que houve e
    libera a memória do processo (o cache de JSON) para o app voltar a responder."""
    if isinstance(e, HTTPException) and e.code != 500:
        return e
    with _CACHE_LOCK:
        _CACHE.clear()
    gc.collect()
    app.logger.exception("erro em %s", request.path if request else "?")
    memoria = isinstance(e, MemoryError)
    titulo = "Faltou memória para esta tela" if memoria else "Deu erro nesta tela"
    detalhe = ("O arquivo é grande demais para caber de uma vez. Já liberei a memória: recarregue a página. "
               "Se acontecer de novo, suba o arquivo em pedaços menores (por período) ou avise o Otto.") \
        if memoria else f"{type(e).__name__}: {e}"
    return (f"""<!doctype html><meta charset=utf-8><title>PLUTOS · erro</title>
<style>body{{font:15px/1.55 system-ui,Arial;margin:0;background:#0B0B0C;color:#F3F3F1}}
.cx{{max-width:640px;margin:12vh auto;padding:28px;background:#131315;border-top:3px solid #D4A017}}
h1{{font-size:20px;margin:0 0 10px}} p{{color:#BDBDB8}} a{{color:#F2C94C}}
code{{font:12px ui-monospace,monospace;color:#8A8A86;word-break:break-all}}</style>
<div class=cx><h1>{titulo}</h1><p>{detalhe}</p>
<p><a href="/">Voltar ao GERAL</a> &nbsp;·&nbsp; <a href="/arquivos">Arquivos</a></p>
<p><code>{request.path if request else ""} · PLUTOS {VERSAO}</code></p></div>""", 500)


def diag_disco() -> dict:
    """O DATA_DIR está num DISCO PERSISTENTE ou no sistema de arquivos do
    container? Essa é a diferença entre os dados sobreviverem ao deploy ou não.

    O teste é o número do dispositivo (st_dev): se o DATA_DIR está num
    dispositivo DIFERENTE da raiz do sistema, ele é um disco montado à parte —
    e o Render mantém esse disco entre deploys. Se for o MESMO dispositivo da
    raiz, os dados estão no container e somem no próximo deploy."""
    import shutil
    out: dict = {"data_dir": DATA_DIR}
    try:
        if not os.path.isdir(DATA_DIR):
            out["existe"] = False
            out["persistente"] = False
            out["aviso"] = "A pasta de dados não existe — será criada vazia no primeiro uso."
            return out
        out["existe"] = True
        dev_dados = os.stat(DATA_DIR).st_dev
        dev_raiz = os.stat("/").st_dev
        persistente = dev_dados != dev_raiz
        out["persistente"] = persistente
        out["aviso"] = ("Disco montado à parte: os dados sobrevivem aos deploys."
                        if persistente else
                        "ATENÇÃO: a pasta de dados está no sistema de arquivos do container, "
                        "NÃO num disco. Tudo que for apurado SOME no próximo deploy. "
                        "Crie um disco no Render e aponte o DATA_DIR para ele.")
        n = tam = 0
        antigo = novo_ = None
        for raiz, _ds, fs in os.walk(DATA_DIR):
            for f in fs:
                try:
                    st = os.stat(os.path.join(raiz, f))
                except OSError:
                    continue
                n += 1
                tam += st.st_size
                if antigo is None or st.st_mtime < antigo:
                    antigo = st.st_mtime
                if novo_ is None or st.st_mtime > novo_:
                    novo_ = st.st_mtime
        out["arquivos"] = n
        out["tamanho_mb"] = round(tam / 1048576, 1)
        if antigo:
            out["mais_antigo"] = datetime.fromtimestamp(antigo, BRT).strftime("%d/%m/%Y %H:%M")
            out["mais_recente"] = datetime.fromtimestamp(novo_, BRT).strftime("%d/%m/%Y %H:%M")
        u = shutil.disk_usage(DATA_DIR)
        out["disco_total_gb"] = round(u.total / 1073741824, 2)
        out["disco_livre_gb"] = round(u.free / 1073741824, 2)
        out["disco_usado_pct"] = round(100 * u.used / u.total, 1) if u.total else 0
    except Exception as e:  # noqa: BLE001
        out["erro"] = str(e)
    return out


@app.route("/saude")
def saude():
    cs = canais()
    return jsonify({"ok": True, "versao": VERSAO, "hora_brasilia": agora().strftime("%d/%m/%Y %H:%M:%S"),
                    "fuso": "America/Sao_Paulo (UTC-3), fixo no código — não depende do relógio do servidor",
                    "data_dir": DATA_DIR, "armazenamento": diag_disco(),
                    "boxes_ativos": [c["nome"] for c in cs if c["ativo"]],
                    "boxes_em_construcao": [c["nome"] for c in cs if not c["ativo"]],
                    "motores": {m: (m in globals() and globals()[m] is not None)
                                for m in ("meli", "magalu", "shopee", "madeira", "webcont", "colombo", "amazon")}})


@app.errorhandler(403)
def e403(_):
    return render_template("erro.html", msg="Seu nível de acesso não abre esta tela."), 403


@app.errorhandler(404)
def e404(_):
    return render_template("erro.html", msg="Essa tela não existe."), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5055)), debug=True)
