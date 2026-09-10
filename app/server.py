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

from flask import (Flask, abort, flash, jsonify, redirect, render_template, request,
                   send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from motor import meli
import planilhas

VERSAO = "2026-09-10n"
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
     "extra": "Comissão do sistema (planilha \"rebate\" do ADC002)",
     "extra_sub": "a que a Gabi usa no PROCV — colunas Pedido / % / R$ do Promob"},
    {"chave": "magalu",   "nome": "Magazine Luiza",  "ativo": False},
    {"chave": "madeira",  "nome": "Madeira Madeira", "ativo": False},
    {"chave": "colombo",  "nome": "Colombo",         "ativo": False},
    {"chave": "cbahia",   "nome": "Casas Bahia",     "ativo": False},
    {"chave": "amazon",   "nome": "Amazon",          "ativo": False},
    {"chave": "webcont",  "nome": "Webcontinental",  "ativo": False},
    {"chave": "shopee",   "nome": "Shopee",          "ativo": False},
]
CANAL_POR_CHAVE = {c["chave"]: c for c in CANAIS}

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


def _json_ler(caminho, padrao):
    try:
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return padrao


def _json_gravar(caminho, dado):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    tmp = caminho + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dado, f, ensure_ascii=False, indent=1, default=str)
    os.replace(tmp, caminho)


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
              "canais_ativos": [c["chave"] for c in CANAIS if c["ativo"]],
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
        "VERSAO": VERSAO, "CANAIS": CANAIS, "papel": papel, "PAPEIS": PAPEIS,
        "usuario": session.get("usuario"), "PODE": {k: (papel in v) for k, v in PODE.items()},
        "IRMAOS": [{"chave": c, "rotulo": r, "url": u, "arquivo": _icone_irmao(c)} for c, r, u in IRMAOS],
        "PARAM": parametros(), "COMPS": competencias(), "comp": comp_atual(),
        "ULT": ultima_rodada_info(),
    }


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
    comps = set()
    for n in os.listdir(d):
        m = re.match(r"(\w+)_(\d{4}-\d{2})\.json$", n)
        if m:
            comps.add(m.group(2))
    return sorted(comps)


def comp_atual() -> str:
    c = request.args.get("mes") if request else None
    comps = competencias()
    if c and re.match(r"^\d{4}-\d{2}$", c):
        return c
    return comps[-1] if comps else agora().strftime("%Y-%m")


def ultima_rodada_info():
    melhor = None
    for n in os.listdir(pasta("rodadas")):
        r = _json_ler(pasta("rodadas", n), None)
        if r and (melhor is None or r["quando"] > melhor["quando"]):
            melhor = {"quando": r["quando"], "canal": r["canal"], "comp": r["competencia"], "quem": r.get("quem")}
    return melhor


def faltante_ler(canal="meli") -> dict:
    return _json_ler(pasta("manual", f"faltante_{canal}.json"), {})


def faltante_gravar(d, canal="meli"):
    _json_gravar(pasta("manual", f"faltante_{canal}.json"), d)


def recalcular_meli(comp: str):
    """Recalcula a rodada do Meli daquela competência a partir dos arquivos
    guardados (usado quando a tabela manual muda ou o ADC002 chega)."""
    r = rodada("meli", comp)
    if not r:
        return None
    df, diag = meli.ler_tabela_geral(r["arquivo"]["caminho"])
    df = df[df["competencia"] == comp]
    sis, sis_diag = None, None
    if r.get("sistema") and os.path.exists(r["sistema"]["caminho"]):
        sis, sis_diag = meli.ler_comissao_sistema(r["sistema"]["caminho"])
    linhas = meli.calcular(df, sis, faltante_ler("meli"), parametros()["tolerancia_comissao"])
    r["linhas"] = linhas
    r["resumo"] = meli.resumo(linhas)
    r["sistema_diag"] = sis_diag
    r["recalculado"] = agora().isoformat()
    _json_gravar(rodada_caminho("meli", comp), r)
    return r


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
            return redirect(request.args.get("proximo") or url_for("painel"))
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
    for c in CANAIS:
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
    return render_template("painel.html", por_canal=por_canal, tot=tot)


@app.route("/canal/<chave>")
@logado
def canal(chave):
    c = CANAL_POR_CHAVE.get(chave) or abort(404)
    comp = comp_atual()
    r = rodada(chave, comp) if c["ativo"] else None
    return render_template("canal.html", c=c, r=r)


# colunas da BASE LINHA A LINHA de cada canal: (rótulo, chave, tipo) — tipo: t texto, n número, p percentual, d data
LINHA_COLS = {
    "meli": [
        ("Data", "data", "d"), ("OC / Pedido mkt", "pedido_mkt", "t"), ("Pedido canal", "pedido_canal", "t"),
        ("ID mkt (Any)", "id_mkt", "t"), ("Pedido Any", "pedido_any", "t"), ("Conta", "conta", "t"),
        ("Status", "status", "t"), ("SKU", "sku", "t"), ("Anúncio", "anuncio", "t"), ("Tipo anúncio", "tipo", "t"),
        ("Valor produtos", "valor_prod", "n"), ("Tarifa venda", "tarifa", "n"), ("Frete pedido", "frete", "n"),
        ("Cupom seller", "cupom_seller", "n"), ("Cupom Meli", "cupom_meli", "n"),
        ("% comissão cobrada", "pct_comissao", "p"), ("% comissão sistema", "sis_pct", "p"),
        ("Comissão sistema R$", "sis_rs", "n"), ("Diferença comissão", "diferenca", "n"),
        ("Tarifa zero?", "tarifa_zero", "b"), ("Faltante campanha", "faltante", "n"), ("Status faltante", "faltante_status", "t"),
        ("Rebate R$", "rebate_rs", "n"), ("Rebate comissão", "rebate_comissao", "n"), ("Rebate frete", "rebate_frete", "n"),
        ("REBATE TOTAL", "rebate_total", "n"),
    ],
}


@app.route("/linha/<chave>")
@logado
def linha(chave):
    c = CANAL_POR_CHAVE.get(chave) or abort(404)
    comp = comp_atual()
    r = rodada(chave, comp) if c["ativo"] else None
    cols = LINHA_COLS.get(chave, [])
    q = (request.args.get("q") or "").strip().lower()
    linhas = list(r["linhas"]) if r else []
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
                           total=total, pag=pag, paginas=max(1, -(-total // por)), q=q, ordem=ordem, desc=desc)


@app.route("/pedidos")
@logado
def pedidos():
    comp = comp_atual()
    q = (request.args.get("q") or "").strip().lower()
    so = request.args.get("so") or ""
    canal_f = request.args.get("canal") or ""
    linhas = []
    for c in CANAIS:
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
    return render_template("arquivos.html", hist=hist, r_meli=rodada("meli", comp))


@app.route("/arquivos/subir/<chave>", methods=["POST"])
@logado
@exige("arquivos")
def subir(chave):
    c = CANAL_POR_CHAVE.get(chave) or abort(404)
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
    try:
        if chave == "meli":
            if tipo == "sistema":
                comp = request.form.get("comp") or comp_atual()
                r = rodada("meli", comp)
                if not r:
                    flash(f"Suba primeiro a Tabela Geral de {f_mesano(comp)}; a comissão do sistema entra em cima dela.")
                    return redirect(url_for("arquivos", mes=comp))
                sis, diag = meli.ler_comissao_sistema(destino)
                r["sistema"] = {"nome": nome, "caminho": destino, "quando": agora().isoformat(), "diag": diag}
                _json_gravar(rodada_caminho("meli", comp), r)
                r = recalcular_meli(comp)
                flash(f"Comissão do sistema lida ({diag['linhas']} pedidos, {diag['modo']}). "
                      f"Rebate de comissão: R$ {f_brl(r['resumo']['rebate_comissao'])}.")
                return redirect(url_for("canal", chave="meli", mes=comp))
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
                antiga = rodada("meli", comp) or {}
                sis, sis_diag = None, None
                if antiga.get("sistema") and os.path.exists(antiga["sistema"]["caminho"]):
                    sis, sis_diag = meli.ler_comissao_sistema(antiga["sistema"]["caminho"])
                linhas = meli.calcular(sub, sis, faltante_ler("meli"), parametros()["tolerancia_comissao"])
                r = {"canal": "meli", "competencia": comp, "quando": agora().isoformat(),
                     "quem": session["usuario"],
                     "arquivo": {"nome": nome, "caminho": destino, "diag": {k: v for k, v in diag.items() if k != "competencias"}},
                     "sistema": antiga.get("sistema"), "sistema_diag": sis_diag,
                     "linhas": linhas, "resumo": meli.resumo(linhas)}
                _json_gravar(rodada_caminho("meli", comp), r)
                feitos.append((comp, len(linhas), r["resumo"]["rebate_total"]))
            txt = " · ".join(f"{f_mesano(c)}: {n} pedidos, R$ {f_brl(t)}" for c, n, t in feitos)
            extra = f" {fora} linha(s) de outro mês ficaram de fora (filtro do BI)." if fora else ""
            flash(f"Mercado Livre lido. {txt}. Linhas rejeitadas: {diag['n_rejeitadas']}.{extra}")
            return redirect(url_for("canal", chave="meli", mes=principal))
    except Exception as e:  # noqa: BLE001
        flash(f"Não consegui ler o arquivo: {e}")
        return redirect(url_for("arquivos"))
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
            recalcular_meli(comp)
        flash("Parâmetros gravados e rebates recalculados.")
        return redirect(url_for("parametros_tela"))
    return render_template("parametros.html")


# --------------------------------------------------------------------------
# saídas — planilha e JSON para o Tropa de Elite / ORION
# --------------------------------------------------------------------------
@app.route("/baixar/<qual>")
@logado
@exige("exportar")
def baixar(qual):
    comp = comp_atual()
    rods = {c["chave"]: rodada(c["chave"], comp) for c in CANAIS}
    rods = {k: v for k, v in rods.items() if v}
    if not rods:
        flash("Não há rodada nesta competência.")
        return redirect(url_for("arquivos"))
    if qual == "rebates":
        bio = planilhas.rebates_xlsx(rods, comp, CANAL_POR_CHAVE, parametros())
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_Rebates_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if qual.startswith("linha_"):
        chave = qual[6:]
        r = rods.get(chave)
        if not r:
            abort(404)
        bio = planilhas.linha_xlsx(r, LINHA_COLS.get(chave, []), CANAL_POR_CHAVE[chave]["nome"], comp)
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_LinhaALinha_{chave.upper()}_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if qual == "faltante":
        bio = planilhas.faltante_xlsx(rods.get("meli"), faltante_ler("meli"), comp)
        return send_file(bio, as_attachment=True,
                         download_name=f"PLUTOS_FaltanteCampanha_Meli_{comp.replace('-', '')}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    abort(404)


@app.route("/api/rebates/<comp>.json")
def api_rebates(comp):
    """O que o Tropa de Elite e o ORION consomem: rebate por pedido, já
    classificado nas três formas. Chave = pedido do marketplace (a OC do ERP)."""
    token = os.environ.get("PLUTOS_TOKEN")
    if token and request.args.get("token") != token and not session.get("usuario"):
        abort(403)
    saida = []
    for c in CANAIS:
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
