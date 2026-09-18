"""Box MAGAZINE LUIZA · PLANILHA 2 — "Vendas no período" (regra da Thaís, 16/09/2026).

Entrada: o zip do portal Magalu "MagazineLuiza_Vendas_<loja>_<periodo>.zip", que
traz DOIS csv (UTF-8, vírgula):

  · relatorio_vendas_pedidos_<data>.csv   — 1 linha por ITEM (tem SKU, título,
    quantidade e a "Coparticipação de Fretes estimada" do pedido).
  · relatorio_vendas_pacotes_<data>.csv   — 1 linha por PACOTE (tem "Frete total
    do pacote", "Modalidade de entrega" e o endereço de entrega).

Daqui saem os DOIS custos de logística do Magalu:
  1. CUSTO FRETE FULL  = "Frete total do pacote" dos pacotes com Modalidade de
     entrega "Magalu entregas - Fulfillment".
  2. COPARTICIPAÇÃO DE FRETE = "Coparticipação de Fretes estimada" — o que NÓS
     pagamos do frete. Vem negativa no relatório; aqui vira positiva (custo).

Pedido CANCELADO fica fora dos dois, como em todos os boxes.

Atenção de operação: a coparticipação só é preenchida DEPOIS que o pedido é
despachado. Um relatório extraído cedo demais mostra quase zero — por isso a
tela mostra a data da extração.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from datetime import date, datetime
from typing import Any


def _norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v == v else 0.0
    t = str(v).strip()
    if not t or _norm(t) in ("nao se aplica", "em apuracao", "n/a", "-", "nan", "none", "sem comentarios"):
        return 0.0
    t = t.replace("R$", "").replace("%", "").replace(" ", "")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _data(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    t = str(v).strip()[:19]
    for f in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def _linhas(caminho: str):
    """csv do portal (UTF-8, vírgula). Devolve (cabeçalho normalizado, iterador)."""
    with open(caminho, "rb") as f:
        raw = f.read()
    txt = raw.decode("utf-8-sig") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("utf-8", "replace")
    amostra = txt[:4000]
    sep = ";" if amostra.count(";") > amostra.count(",") else ","
    r = csv.reader(io.StringIO(txt), delimiter=sep)
    cab = next(r)
    return [_norm(c) for c in cab], r


def que_arquivo(caminho: str) -> str:
    """'pedidos' (itens) · 'pacotes' · '' se não for a Planilha 2 do Magalu."""
    try:
        cab, _ = _linhas(caminho)
    except Exception:  # noqa: BLE001
        return ""
    if "numero do pacote" in cab and "frete total do pacote" in cab:
        return "pacotes"
    if "numero do pedido" in cab and "coparticipacao de fretes estimada" in cab and "codigo sku seller" in cab:
        return "pedidos"
    return ""


def _pega(cab, linha, nome, padrao=None):
    try:
        return linha[cab.index(nome)]
    except (ValueError, IndexError):
        return padrao


def ler_pacotes(caminho: str) -> tuple[list[dict], dict]:
    cab, it = _linhas(caminho)
    out, status, mods = [], {}, {}
    for l in it:
        ped = (_pega(cab, l, "numero do pedido") or "").strip()
        if not ped:
            continue
        d = _data(_pega(cab, l, "data do pacote"))
        if not d:
            continue
        st = (_pega(cab, l, "status pacote no momento que o relatorio foi solicitado") or "").strip()
        mod = (_pega(cab, l, "modalidade de entrega") or "").strip()
        status[st] = status.get(st, 0) + 1
        mods[mod] = mods.get(mod, 0) + 1
        out.append({
            "pedido": ped, "pacote": (_pega(cab, l, "numero do pacote") or "").strip(),
            "data": str(d), "competencia": f"{d.year}-{d.month:02d}",
            "status": st, "modalidade": mod, "full": "fulfillment" in _norm(mod),
            "cancelado": "cancelado" in _norm(st),
            "itens": int(_num(_pega(cab, l, "quantidade de itens do pacote"))),
            "produtos": _num(_pega(cab, l, "valor total dos produtos do pacote")),
            "desconto": _num(_pega(cab, l, "desconto totais do pacote")),
            "frete": _num(_pega(cab, l, "frete total do pacote")),
            "total": _num(_pega(cab, l, "valor total do pacote")),
            "uf": (_pega(cab, l, "estado") or "").strip(), "cidade": (_pega(cab, l, "cidade") or "").strip(),
            "entregue": (_pega(cab, l, "entregue em") or "").strip(),
        })
    comps: dict[str, int] = {}
    for p in out:
        comps[p["competencia"]] = comps.get(p["competencia"], 0) + 1
    diag = {"tipo": "pacotes", "linhas": len(out), "status": status, "modalidades": mods, "competencias": comps,
            "de": min((p["data"] for p in out), default=""), "ate": max((p["data"] for p in out), default=""),
            "cancelados": sum(1 for p in out if p["cancelado"]),
            "full": sum(1 for p in out if p["full"]),
            "frete_full": round(sum(p["frete"] for p in out if p["full"] and not p["cancelado"]), 2)}
    return out, diag


def ler_pedidos(caminho: str) -> tuple[list[dict], dict]:
    cab, it = _linhas(caminho)
    out = []
    for l in it:
        ped = (_pega(cab, l, "numero do pedido") or "").strip()
        if not ped:
            continue
        d = _data(_pega(cab, l, "data do pedido"))
        if not d:
            continue
        out.append({
            "pedido": ped, "data": str(d), "competencia": f"{d.year}-{d.month:02d}",
            "sku": (_pega(cab, l, "codigo sku seller") or "").strip(),
            "produto": (_pega(cab, l, "titulo do produto") or "").strip()[:90],
            "qtd": _num(_pega(cab, l, "quantidade de itens")),
            "valor_item": _num(_pega(cab, l, "valor total do item")),
            "bruto": _num(_pega(cab, l, "valor bruto do pedido")),
            # vem negativa (é desconto no repasse); guardo POSITIVA, como custo
            "copart": round(abs(_num(_pega(cab, l, "coparticipacao de fretes estimada"))), 2),
            "liquido": _num(_pega(cab, l, "valor liquido estimado a receber")),
        })
    comps: dict[str, int] = {}
    for p in out:
        comps[p["competencia"]] = comps.get(p["competencia"], 0) + 1
    diag = {"tipo": "pedidos", "linhas": len(out), "competencias": comps,
            "pedidos": len({p["pedido"] for p in out}),
            "de": min((p["data"] for p in out), default=""), "ate": max((p["data"] for p in out), default=""),
            "com_copart": sum(1 for p in out if p["copart"] > 0),
            "copart": round(sum(p["copart"] for p in out), 2)}
    return out, diag


def resumo_frete_full(pacotes: list[dict], comp: str) -> dict:
    """Custo do frete Full da competência: pacotes Full, sem cancelados."""
    ps = [p for p in pacotes if p["competencia"] == comp and p["full"] and not p["cancelado"]]
    fora = [p for p in pacotes if p["competencia"] == comp and p["full"] and p["cancelado"]]
    venda = round(sum(p["total"] for p in ps), 2)
    frete = round(sum(p["frete"] for p in ps), 2)
    por_dia: dict[str, dict] = {}
    for p in ps:
        d = por_dia.setdefault(p["data"], {"dia": p["data"], "pacotes": 0, "venda": 0.0, "frete": 0.0})
        d["pacotes"] += 1; d["venda"] += p["total"]; d["frete"] += p["frete"]
    por_uf: dict[str, dict] = {}
    for p in ps:
        u = por_uf.setdefault(p["uf"] or "—", {"uf": p["uf"] or "—", "pacotes": 0, "venda": 0.0, "frete": 0.0})
        u["pacotes"] += 1; u["venda"] += p["total"]; u["frete"] += p["frete"]
    for d in list(por_dia.values()) + list(por_uf.values()):
        d["venda"] = round(d["venda"], 2); d["frete"] = round(d["frete"], 2)
        d["medio"] = round(d["frete"] / d["pacotes"], 2) if d["pacotes"] else 0.0
        d["pct"] = round(100 * d["frete"] / d["venda"], 2) if d["venda"] else 0.0
    return {"pacotes": len(ps), "cancelados": len(fora), "venda": venda, "frete": frete,
            "sem_frete": sum(1 for p in ps if p["frete"] <= 0),
            "medio": round(frete / len(ps), 2) if ps else 0.0,
            "pct": round(100 * frete / venda, 2) if venda else 0.0,
            "por_dia": sorted(por_dia.values(), key=lambda x: x["dia"]),
            "por_uf": sorted(por_uf.values(), key=lambda x: -x["frete"])}


def resumo_copart(itens: list[dict], pacotes: list[dict], comp: str) -> dict:
    """Coparticipação de frete da competência (o que NÓS pagamos), por SKU."""
    mod = {}
    canc = set()
    for p in pacotes:
        mod.setdefault(p["pedido"], p["modalidade"])
        if p["cancelado"]:
            canc.add(p["pedido"])
    its = [i for i in itens if i["competencia"] == comp and i["pedido"] not in canc]
    fora = round(sum(i["copart"] for i in itens if i["competencia"] == comp and i["pedido"] in canc), 2)
    total = round(sum(i["copart"] for i in its), 2)
    venda = round(sum(i["valor_item"] for i in its), 2)
    por_dia: dict[str, dict] = {}
    for i in its:
        d = por_dia.setdefault(i["data"], {"dia": i["data"], "itens": 0, "venda": 0.0, "copart": 0.0, "com": 0})
        d["itens"] += 1; d["venda"] += i["valor_item"]; d["copart"] += i["copart"]
        d["com"] += 1 if i["copart"] > 0 else 0
    por_sku: dict[str, dict] = {}
    for i in its:
        s = por_sku.setdefault(i["sku"] or "—", {"sku": i["sku"] or "—", "produto": i["produto"], "itens": 0,
                                                 "qtd": 0.0, "venda": 0.0, "copart": 0.0, "com": 0})
        s["itens"] += 1; s["qtd"] += i["qtd"]; s["venda"] += i["valor_item"]; s["copart"] += i["copart"]
        s["com"] += 1 if i["copart"] > 0 else 0
    for d in list(por_dia.values()) + list(por_sku.values()):
        d["venda"] = round(d["venda"], 2); d["copart"] = round(d["copart"], 2)
        d["pct"] = round(100 * d["copart"] / d["venda"], 2) if d["venda"] else 0.0
        d["medio"] = round(d["copart"] / d["com"], 2) if d.get("com") else 0.0
    mods: dict[str, float] = {}
    for i in its:
        mods[mod.get(i["pedido"], "—") or "—"] = round(mods.get(mod.get(i["pedido"], "—") or "—", 0.0) + i["copart"], 2)
    return {"itens": len(its), "pedidos": len({i["pedido"] for i in its}), "com_copart": sum(1 for i in its if i["copart"] > 0),
            "copart": total, "venda": venda, "cancelados_fora": fora,
            "pct": round(100 * total / venda, 2) if venda else 0.0,
            "medio": round(total / max(1, sum(1 for i in its if i["copart"] > 0)), 2),
            "por_modalidade": mods,
            "por_dia": sorted(por_dia.values(), key=lambda x: x["dia"]),
            "por_sku": sorted(por_sku.values(), key=lambda x: -x["copart"])}
