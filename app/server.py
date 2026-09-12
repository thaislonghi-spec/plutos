# -*- coding: utf-8 -*-
"""
PLUTOS · Rebates · Multimóveis
Flask + arquivos JSON em DATA_DIR (multiempresa desde o início).
"""
from __future__ import annotations

import io
import json
import os
import re
import secrets
import shutil
from datetime import datetime, timedelta, timezone
from functools import wraps
from typing import Any

from flask import (Flask, abort, flash, jsonify, redirect, render_template, request,
                   send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from motor import meli, erp, magalu, shopee
import planilhas

VERSAO = "2026-09-12e"
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
    {"chave": "madeira",  "nome": "Madeira Madeira", "ativo": False},
    {"chave": "colombo",  "nome": "Colombo",         "ativo": False},
    {"chave": "cbahia",   "nome": "Casas Bahia",     "ativo": False},
    {"chave": "amazon",   "nome": "Amazon",          "ativo": False},
    {"chave": "webcont",  "nome": "Webcontinental",  "ativo": False},
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
_CACHE: dict[str, tuple[tuple, Any]] = {}
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
            return hit[1]
    try:
        with open(caminho, encoding="utf-8") as f:
            dado = json.load(f)
    except Exception:
        return padrao
    with _CACHE_LOCK:
        _CACHE[caminho] = (chave, dado)
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
    {"codigo": "705", "canal": "MP - COLOMBO", "gestor": "Bruna Colares", "comissao": "9", "taxa": "0", "tipo": "GMV"},
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
COMISSAO_CAMPOS = ["codigo", "canal", "gestor", "comissao", "taxa", "tipo"]


def parametros() -> dict:
    padrao = {"empresa": "Multimóveis", "tolerancia_comissao": 0.50,
              "canais_ativos": [c["chave"] for c in CANAIS_BASE if c["ativo"]],
              "comissoes": COMISSAO_PADRAO}
    p = _json_ler(pasta("parametros.json"), {})
    return {**padrao, **p}


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
        r = rodada(chave, comp_atual())
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


def rodada(canal, comp) -> dict | None:
    return _json_ler(rodada_caminho(canal, comp), None)


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


def comissao_cadastrada(nomes: tuple[str, ...], padrao: tuple[float, float]) -> tuple[float, float]:
    """(pct, taxa R$ por pedido) da tabela de comissões de Parâmetros para o canal;
    nomes = pedaços que identificam a linha (ex.: 'MAGAZINE', 'MAGALU')."""
    for lin in parametros().get("comissoes") or COMISSAO_PADRAO:
        nome = unidecode_lower(lin.get("canal", ""))
        if any(unidecode_lower(n) in nome for n in nomes) and "full" not in nome and "xpress" not in nome:
            try:
                pct = float(str(lin.get("comissao", "0")).replace("%", "").replace(",", ".") or 0) / 100
                taxa = float(str(lin.get("taxa", "0")).replace("R$", "").replace(",", ".").strip() or 0)
                return pct, taxa
            except ValueError:
                break
    return padrao


def erp_por_base(box: str) -> dict:
    """OC do ERP sem o sufixo -N → linha do ERP (o Magalu grava 'LU-…-1', 'LU-…-2')."""
    out = {}
    for oc, l in erp_ler()["ocs"].items():
        if l.get("box") == box:
            out.setdefault(magalu.base_oc(oc), l)
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


def descricao_de(sku, cad=None) -> str:
    """Descrição do SKU pelo cadastro (Anymarket primeiro, CustoProduto se não houver)."""
    cad = cad if cad is not None else descricoes_ler()["skus"]
    e = cad.get(str(sku or "").strip()) or cad.get(str(sku or "").strip().upper()) or {}
    return e.get("descricao") or e.get("descricao_curta") or ""


def descricoes_ler() -> dict:
    """SKU → descrição (cadastro subido em Parâmetros). Vale para todos os canais."""
    return _json_ler(pasta("cadastro", "descricoes.json"), {"skus": {}, "quando": None, "quem": None, "arquivo": None})


def recalcular_meli(comp: str):
    """Recalcula a rodada do Meli daquela competência a partir dos arquivos
    guardados (usado quando a tabela manual muda ou o ADC002 chega)."""
    r = rodada("meli", comp)
    if not r:
        return None
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
    _json_gravar(rodada_caminho("meli", comp), r)
    return r


def recalcular_magalu(comp: str):
    r = rodada("magalu", comp)
    if not r:
        return None
    df, diag = magalu.ler(r["arquivo"]["caminho"])
    df = df[df["competencia"] == comp]
    pct, taxa = comissao_cadastrada(("MAGAZINE", "MAGALU"), (0.11, 5.0))
    linhas = magalu.calcular(df, pct, taxa, erp_por_base("magalu"))
    r["linhas"] = linhas
    r["resumo"] = magalu.resumo(linhas, int(df["cancelado"].sum()))
    r["sistema"] = {"nome": f"Parâmetros · {pct * 100:.2f}% + R$ {taxa:.2f}/pedido", "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    _json_gravar(rodada_caminho("magalu", comp), r)
    return r


def processar_magalu(destino: str, nome: str, quem: str) -> str:
    df, diag = magalu.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    pct, taxa = comissao_cadastrada(("MAGAZINE", "MAGALU"), (0.11, 5.0))
    feitos = []
    fora = 0
    for comp, n in comps.items():
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            continue
        sub = df[df["competencia"] == comp]
        linhas = magalu.calcular(sub, pct, taxa, erp_por_base("magalu"))
        r = {"canal": "magalu", "competencia": comp, "quando": agora().isoformat(), "quem": quem,
             "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
             "sistema": {"nome": f"Parâmetros · {pct * 100:.2f}% + R$ {taxa:.2f}/pedido", "quando": agora().isoformat(),
                         "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}},
             "linhas": linhas, "resumo": magalu.resumo(linhas, int(sub["cancelado"].sum()))}
        r["sistema_diag"] = r["sistema"]["diag"]
        _json_gravar(rodada_caminho("magalu", comp), r)
        feitos.append((comp, len(linhas), r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    return f"Magazine Luiza lido. {txt}. Cancelados fora: {diag['cancelados']}.{extra}"


def recalcular_shopee(comp: str):
    r = rodada("shopee", comp)
    if not r:
        return None
    df, diag = shopee.ler(r["arquivo"]["caminho"])
    df = df[df["competencia"] == comp]
    pct, taxa = comissao_cadastrada(("SHOPEE",), (0.12, 12.0))
    linhas = shopee.calcular(df, pct, taxa, erp_ler()["ocs"])
    r["linhas"] = linhas
    r["resumo"] = shopee.resumo(linhas, int(df["cancelado"].sum()))
    r["sistema"] = {"nome": f"Parâmetros · {pct * 100:.2f}% + R$ {taxa:.2f}/item", "quando": agora().isoformat(),
                    "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}}
    r["sistema_diag"] = r["sistema"]["diag"]
    r["recalculado"] = agora().isoformat()
    _json_gravar(rodada_caminho("shopee", comp), r)
    return r


def processar_shopee(destino: str, nome: str, quem: str) -> str:
    df, diag = shopee.ler(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    pct, taxa = comissao_cadastrada(("SHOPEE",), (0.12, 12.0))
    feitos = []
    fora = 0
    for comp, n in comps.items():
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n
            continue
        sub = df[df["competencia"] == comp]
        linhas = shopee.calcular(sub, pct, taxa, erp_ler()["ocs"])
        r = {"canal": "shopee", "competencia": comp, "quando": agora().isoformat(), "quem": quem,
             "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
             "sistema": {"nome": f"Parâmetros · {pct * 100:.2f}% + R$ {taxa:.2f}/item", "quando": agora().isoformat(),
                         "diag": {"linhas": len(linhas), "modo": "tabela de comissões (Parâmetros)"}},
             "linhas": linhas, "resumo": shopee.resumo(linhas, int(sub["cancelado"].sum()))}
        r["sistema_diag"] = r["sistema"]["diag"]
        _json_gravar(rodada_caminho("shopee", comp), r)
        feitos.append((comp, len(linhas), r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora." if fora else ""
    return f"Shopee lida. {txt}. Cancelados fora: {diag['cancelados']} · {diag['itens']} itens → {diag['linhas']} pedidos.{extra}"


def recalcular_canais(comp: str) -> list[str]:
    feitos = []
    r = recalcular_meli(comp)
    if r:
        feitos.append(f"Mercado Livre {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = recalcular_magalu(comp)
    if r:
        feitos.append(f"Magazine Luiza {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
    r = recalcular_shopee(comp)
    if r:
        feitos.append(f"Shopee {f_mesano(comp)}: R$ {f_brl(r['resumo']['rebate_total'])}")
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
        r = rodada(c["chave"], comp) if c["ativo"] else None
        if r:
            s = r["resumo"]
            por_canal.append({**c, "res": s, "quando": r["quando"]})
            tot["rs"] += s["rebate_rs"]; tot["com"] += s["rebate_comissao"]; tot["frete"] += 0.0
            tot["total"] += s["rebate_total"]; tot["venda"] += s["venda"]; tot["pedidos"] += s["pedidos"]
            tot["pend"] += s["faltante_pendentes"]
        else:
            por_canal.append({**c, "res": None})
    tot["pct"] = (100 * tot["total"] / tot["venda"]) if tot["venda"] else 0
    erp_res = erp.resumo_por_box([l for l in erp_ler()["ocs"].values() if l["competencia"] == comp])
    for c in por_canal:
        c["erp"] = erp_res.get(c["chave"])
    return render_template("painel.html", por_canal=por_canal, tot=tot, erp_outros=erp_res.get("outros"))


@app.route("/canal/<chave>")
@logado
def canal(chave):
    c = canal_por_chave().get(chave) or abort(404)
    comp = comp_atual()
    r = rodada(chave, comp) if c["ativo"] else None
    if r and chave == "meli" and "com_sistema" not in r["resumo"]:
        # rodada gravada por versão anterior: completa o resumo sem exigir rodar de novo
        r["resumo"] = meli.resumo(r["linhas"])
        _json_gravar(rodada_caminho("meli", comp), r)
    if chave == "magalu":
        return render_template("canal_magalu.html", c=c, r=r)
    if chave == "shopee":
        return render_template("canal_shopee.html", c=c, r=r)
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
        ("Comissão real R$", "tarifa", "n"), ("% real", "pct_comissao", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("% sistema", "sis_pct", "p"), ("R$/item sistema", "sis_taxa", "n"),
        ("Diferença = rebate comissão", "diferenca", "n"),
        ("Incentivo Shopee (ação)", "incentivo", "n"), ("Incentivo de cupom", "cupom_shopee", "n"), ("Cupom vendedor", "cupom_seller", "n"), ("Moedas (qtd)", "moedas", "n"),
        ("Frete estimado", "frete", "n"), ("Frete pago comprador", "frete_comprador", "n"),
        ("Taxa transação", "taxa_transacao", "n"), ("Total global", "total_global", "n"),
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
    return render_template("faltante.html", itens=itens, r=r, pend=pend, sug=sug)


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
        r = _json_ler(pasta("rodadas", n), None)
        if r:
            hist.append({"canal": r["canal"], "comp": r["competencia"], "quando": r["quando"],
                         "quem": r.get("quem"), "arquivo": r["arquivo"]["nome"],
                         "pedidos": r["resumo"]["pedidos"], "total": r["resumo"]["rebate_total"],
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
    rods = {c["chave"]: rodada(c["chave"], comp) for c in canais() if c["ativo"]}
    return render_template("arquivos.html", pend=pend, pend_por=pend_por, rods=rods, hist=hist, r_meli=rodada("meli", comp), erp_idx=idx, erp_res=erp_res,
                           nomes_erp=nomes_erp, sem_box=sem_box, mapa_erp=m,
                           mlbs_ult=(ml["uploads"][-1] if ml["uploads"] else None), mlbs_total=len(ml["mlbs"]))


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
        _json_gravar(rodada_caminho("meli", comp), r)
        r = recalcular_meli(comp)
        return (f"Comissão do sistema lida ({diag['linhas']} pedidos, {diag['modo']}). "
                f"Rebate de comissão: R$ {f_brl(r['resumo']['rebate_comissao'])}.")
    df, diag = meli.ler_tabela_geral(destino)
    comps = diag["competencias"]
    principal = max(comps, key=comps.get)
    feitos = []
    fora = 0
    for comp, n in comps.items():
        if comp != principal and n < 0.3 * comps[principal]:
            fora += n  # linhas soltas de outro mês = sujeira do filtro do BI, não competência
            continue
        sub = df[df["competencia"] == comp]
        sis = erp_comissao_sistema()
        sis_diag = {"linhas": len(sis), "modo": "ERP · Pedidos Marketplace (coluna AB)"} if sis else None
        linhas = meli.calcular(sub, sis, faltante_ler("meli"), parametros()["tolerancia_comissao"])
        r = {"canal": "meli", "competencia": comp, "quando": agora().isoformat(),
             "quem": quem,
             "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
             "sistema": ({"nome": "ERP · Pedidos Marketplace", "quando": erp_ler()["uploads"][-1]["quando"], "diag": sis_diag}
                         if sis and erp_ler()["uploads"] else None),
             "sistema_diag": sis_diag,
             "linhas": linhas, "resumo": meli.resumo(linhas)}
        _json_gravar(rodada_caminho("meli", comp), r)
        feitos.append((comp, len(linhas), r["resumo"]["rebate_total"]))
    txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
    extra = f" {fora} linha(s) de outro mês ficaram de fora (filtro do BI)." if fora else ""
    return f"Mercado Livre lido. {txt}. Linhas rejeitadas: {diag['n_rejeitadas']}.{extra}"


PEND_ORDEM = {"erp": 0, "base": 1, "rebates": 2, "sistema": 3}


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
                m = processar_magalu(x["caminho"], x["nome"], x["quem"])
            elif x["chave"] == "shopee":
                m = processar_shopee(x["caminho"], x["nome"], x["quem"])
            else:
                m = f"{x['nome']}: o box {x['chave']} ainda não tem motor."
        except Exception as e:  # noqa: BLE001
            m = f"{x['nome']}: não consegui ler — {e}"
        msgs.append(m)
        feitos.append(x["quando"])
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
    nome = secure_filename(f.filename)
    carimbo = agora().strftime("%Y%m%d_%H%M%S")
    destino = pasta("arquivos", chave, f"{carimbo}_{nome}")
    f.save(destino)
    tipo = request.form.get("tipo", "base")
    lista = pend_ler()
    lista.append({"chave": chave, "tipo": tipo, "nome": nome, "caminho": destino, "quando": agora().isoformat(), "quem": session["usuario"]})
    pend_gravar(lista)
    if request.form.get("rodar"):
        for m in pend_processar(chave):
            flash(m)
        return redirect(url_for("canal", chave=chave, mes=comp_atual()) if chave != "meli" or tipo == "base" else url_for("mlbs") if tipo == "rebates" else url_for("arquivos"))
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
    return redirect(url_for("arquivos"))


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
    nome = secure_filename(f.filename)
    destino = pasta("arquivos", "erp", f"{agora().strftime('%Y%m%d_%H%M%S')}_{nome}")
    f.save(destino)
    lista = pend_ler()
    lista.append({"chave": "erp", "tipo": "erp", "nome": nome, "caminho": destino, "quando": agora().isoformat(), "quem": session["usuario"]})
    pend_gravar(lista)
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
        except ValueError:
            pass
        _json_gravar(pasta("parametros.json"), p)
        for comp in competencias():
            recalcular_canais(comp)
        flash("Parâmetros gravados e rebates recalculados.")
        return redirect(url_for("parametros_tela"))
    return render_template("parametros.html", DESC=descricoes_ler())


# --------------------------------------------------------------------------
# saídas — planilha e JSON para o Tropa de Elite / ORION
# --------------------------------------------------------------------------
def _mlbs_filtrados(q: str, tipo: str) -> list[dict]:
    d = mlbs_ler()
    desc = descricoes_ler()["skus"]
    itens = []
    for m in d["mlbs"].values():
        m = dict(m)
        e = desc.get(m["sku"]) or desc.get(m["sku"].upper()) or {}
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


def _coletas_pedidos(so_validos: bool = True) -> list[dict]:
    """Pedidos de coletas = Frete Coletas > 0 na Planilha 2. 'Válido' = status
    Pago na Planilha 1; cancelado / devolvido fica fora. Sem Planilha 1 do mês,
    o status é desconhecido e o pedido entra (com aviso na tela)."""
    st = _status_meli()
    out = []
    for k, p in resumo_meli_ler()["pedidos"].items():
        if (p.get("frete_coletas") or 0) <= 0:
            continue
        s_, conta = st.get(k, ("", ""))
        p = dict(p, pedido_mkt=k, status=s_ or "?", conta=conta)
        if so_validos and s_ and unidecode_lower(s_) != "pago":
            continue
        out.append(p)
    return out


def _soma(ps, k):
    return float(sum((p.get(k) or 0) for p in ps))


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
    return render_template("coletas.html", c=canal_por_chave()["meli"], comp=comp, atual=atual, por_comp=por_comp, linhas=linhas,
                           COMPS_COLETAS=comps, fora=fora.get(comp, 0), sem_status=sem_status)


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
            bio = planilhas.linha_xlsx({"linhas": erp_linhas(chave, comp)}, ERP_SO_COLS, canal_por_chave()[chave]["nome"], comp)
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_LinhaALinha_{chave.upper()}_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if qual == "faltante":
        bio = planilhas.faltante_xlsx(rods.get("meli"), faltante_ler("meli"), comp)
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_FaltanteCampanha_Meli_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    abort(404)


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
            for l in r["linhas"]:
                linhas.append({"oc": l["pedido_mkt"], "data": l["data"], "canal": c["nome"],
                               "rebate_rs": l["rebate_rs"], "rebate_comissao": l["rebate_comissao"],
                               "rebate_frete": l["rebate_frete"], "rebate_total": l["rebate_total"],
                               "competencia": comp, "sku": l.get("sku", ""), "pedido_canal": l.get("pedido_canal", ""),
                               "faltante_status": l.get("faltante_status", "")})
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
                          "faltante_status": l["faltante_status"]})
    return jsonify({"competencia": comp, "gerado": agora().isoformat(), "versao": VERSAO,
                    "pedidos": len(saida), "linhas": saida})


@app.route("/saude")
def saude():
    return jsonify({"ok": True, "versao": VERSAO, "hora_brasilia": agora().strftime("%d/%m/%Y %H:%M:%S"),
                    "fuso": "America/Sao_Paulo (UTC-3), fixo no código — não depende do relógio do servidor",
                    "data_dir": DATA_DIR})


@app.errorhandler(403)
def e403(_):
    return render_template("erro.html", msg="Seu nível de acesso não abre esta tela."), 403


@app.errorhandler(404)
def e404(_):
    return render_template("erro.html", msg="Essa tela não existe."), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5055)), debug=True)
