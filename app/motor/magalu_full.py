"""Box MAGAZINE LUIZA · COBRANÇAS DO FULFILLMENT (regra da Thaís, 16/09/2026).

Entrada: os zips do portal Magalu
  · CSV-156171_CobrançasFullFilment_<periodo>.zip
  · CSV-160172_ColetaFullfilment_<periodo>.zip
que juntos trazem quatro csv (";" com número brasileiro):

  produtos_manuseados.csv   CD · sku · quantidade · valor unitário      → MANUSEIO
  produtos_armazenados.csv  CD · data · sku · quantidade · valor unit.  → ARMAZENAGEM
  tempo_estoque.csv         CD · data entrada · aniversário ·
                            "Descrição / SKU" · quantidade · valor unit. → TEMPO DE ESTOQUE
  produtos_coletados.csv    data · agenda · espaço (m³) · valor          → COLETA

Os três primeiros têm SKU e viram custo por item. A COLETA é cobrada por
agenda/m³ e NÃO tem SKU — ela aparece separada e, na tela, também rateada pela
participação de cada SKU no manuseio, sempre marcada como rateio.

Custo total do Full para a Multimóveis = manuseio + armazenagem + tempo de
estoque + coleta + coparticipação de frete (esta vem da Planilha 2 · Vendas).
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from datetime import datetime
from typing import Any

TIPOS = {"manuseio": "Manuseio", "armazenagem": "Armazenagem",
         "tempo_estoque": "Tempo de estoque", "coleta": "Coleta"}


def _norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v == v else 0.0
    t = str(v).strip().replace("R$", "").replace(" ", "")
    if not t or t in ("-", "nan", "None"):
        return 0.0
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _data(v):
    t = str(v or "").strip()[:10]
    for f in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def _ler(caminho: str):
    with open(caminho, "rb") as f:
        raw = f.read()
    txt = raw.decode("utf-8-sig") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("utf-8", "replace")
    amostra = txt[:3000]
    sep = ";" if amostra.count(";") >= amostra.count(",") else ","
    r = csv.reader(io.StringIO(txt), delimiter=sep)
    cab = [_norm(c) for c in next(r)]
    return cab, r


def que_arquivo(caminho: str) -> str:
    """manuseio · armazenagem · tempo_estoque · coleta · '' se não for do Full."""
    try:
        cab, _ = _ler(caminho)
    except Exception:  # noqa: BLE001
        return ""
    if "agenda" in cab and "espaco" in cab:
        return "coleta"
    if "data entrada" in cab and "aniversario" in cab:
        return "tempo_estoque"
    if "sku" in cab and "data" in cab and "valor unitario" in cab:
        return "armazenagem"
    if "sku" in cab and "valor unitario" in cab:
        return "manuseio"
    return ""


def _pega(cab, l, nome, padrao=""):
    try:
        return l[cab.index(nome)]
    except (ValueError, IndexError):
        return padrao


def ler(caminho: str) -> tuple[str, list[dict], dict]:
    """Devolve (tipo, linhas, diagnóstico). Uma linha por cobrança."""
    qual = que_arquivo(caminho)
    if not qual:
        raise ValueError("Não parece uma cobrança do Fulfillment do Magalu (manuseio, armazenagem, "
                         "tempo de estoque ou coleta).")
    cab, it = _ler(caminho)
    out = []
    for n, l in enumerate(it, 1):
        if not any((x or "").strip() for x in l):
            continue
        if qual == "coleta":
            d = _data(_pega(cab, l, "data"))
            v = _num(_pega(cab, l, "valor"))
            out.append({"tipo": qual, "cd": "", "sku": "", "produto": "", "data": str(d) if d else "",
                        "agenda": _pega(cab, l, "agenda").strip(), "espaco": _num(_pega(cab, l, "espaco")),
                        "qtd": 1.0, "unit": v, "valor": round(v, 2), "chave": f"coleta|{_pega(cab, l, 'agenda').strip()}|{d}"})
            continue
        if qual == "tempo_estoque":
            bruto = _pega(cab, l, "produto/sku")
            sku = bruto.rsplit("/", 1)[-1].strip() if "/" in bruto else bruto.strip()
            produto = bruto.rsplit("/", 1)[0].strip() if "/" in bruto else ""
            d = _data(_pega(cab, l, "data entrada"))
        else:
            sku = _pega(cab, l, "sku").strip()
            produto = ""
            d = _data(_pega(cab, l, "data")) if qual == "armazenagem" else None
        q = _num(_pega(cab, l, "quantidade")) or 1.0
        u = _num(_pega(cab, l, "valor unitario"))
        out.append({"tipo": qual, "cd": _pega(cab, l, "cd").strip(), "sku": sku, "produto": produto[:90],
                    "data": str(d) if d else "", "agenda": "", "espaco": 0.0,
                    "qtd": q, "unit": u, "valor": round(q * u, 2),
                    "aniversario": _pega(cab, l, "aniversario").strip() if qual == "tempo_estoque" else "",
                    "chave": f"{qual}|{_pega(cab, l, 'cd').strip()}|{sku}|{d}|{n}"})
    datas = [x["data"] for x in out if x["data"]]
    diag = {"tipo": qual, "linhas": len(out), "valor": round(sum(x["valor"] for x in out), 2),
            "skus": len({x["sku"] for x in out if x["sku"]}),
            "com_valor": sum(1 for x in out if x["valor"] > 0),
            "de": min(datas, default=""), "ate": max(datas, default="")}
    return qual, out, diag


def resumo(cobrancas: list[dict], copart_por_sku: dict | None = None, nomes: dict | None = None,
           de: str = "", ate: str = "") -> dict:
    """Custo total do Full por SKU. copart_por_sku = {sku: R$} da Planilha 2;
    nomes = {sku: descrição} para a tabela não ficar só com código."""
    copart_por_sku = copart_por_sku or {}
    nomes = nomes or {}
    cb = [c for c in cobrancas if (not de or not c["data"] or de <= c["data"] <= ate)]
    por_tipo = {k: round(sum(c["valor"] for c in cb if c["tipo"] == k), 2) for k in TIPOS}
    coleta = por_tipo["coleta"]
    com_sku = [c for c in cb if c["sku"]]
    por_sku: dict[str, dict] = {}
    for c in com_sku:
        s = por_sku.setdefault(c["sku"], {"sku": c["sku"], "produto": c["produto"], "qtd": 0.0,
                                          "manuseio": 0.0, "armazenagem": 0.0, "tempo_estoque": 0.0,
                                          "copart": 0.0, "coleta_rateio": 0.0, "cds": [], "ultimo": "",
                                          "cobrancas": 0, "estocado": 0.0})
        if c["produto"] and not s["produto"]:
            s["produto"] = c["produto"]
        if c["cd"] and c["cd"] not in s["cds"]:
            s["cds"].append(c["cd"])
        if c["data"] > s["ultimo"]:
            s["ultimo"] = c["data"]
        s["cobrancas"] += 1
        if c["tipo"] == "manuseio":
            s["qtd"] += c["qtd"]
        if c["tipo"] == "armazenagem":
            s["estocado"] += c["qtd"]
        s[c["tipo"]] = round(s[c["tipo"]] + c["valor"], 2)
    for sku, v in copart_por_sku.items():
        s = por_sku.setdefault(sku, {"sku": sku, "produto": "", "qtd": 0.0, "manuseio": 0.0,
                                     "armazenagem": 0.0, "tempo_estoque": 0.0, "copart": 0.0,
                                     "coleta_rateio": 0.0, "cds": [], "ultimo": "", "cobrancas": 0, "estocado": 0.0})
        s["copart"] = round(v, 2)
    # rateio da coleta pela participação no manuseio (não tem SKU na origem)
    base_rateio = sum(s["manuseio"] for s in por_sku.values())
    for s in por_sku.values():
        if not s["produto"]:
            s["produto"] = (nomes.get(s["sku"]) or "")[:90]
        s["coleta_rateio"] = round(coleta * s["manuseio"] / base_rateio, 2) if base_rateio else 0.0
        s["custo_sku"] = round(s["manuseio"] + s["armazenagem"] + s["tempo_estoque"] + s["copart"], 2)
        s["custo_total"] = round(s["custo_sku"] + s["coleta_rateio"], 2)
        s["por_unidade"] = round(s["custo_total"] / s["qtd"], 2) if s["qtd"] else 0.0
    copart = round(sum(copart_por_sku.values()), 2)
    total = round(sum(por_tipo.values()) + copart, 2)
    return {
        "linhas": len(cb), "skus": len(por_sku),
        "manuseio": por_tipo["manuseio"], "armazenagem": por_tipo["armazenagem"],
        "tempo_estoque": por_tipo["tempo_estoque"], "coleta": coleta, "copart": copart,
        "total": total,
        "cobrancas": round(sum(por_tipo.values()), 2),
        "por_sku": sorted(por_sku.values(), key=lambda s: -s["custo_total"]),
        "cds": sorted({c["cd"] for c in cb if c["cd"]}),
        "de": min((c["data"] for c in cb if c["data"]), default=""),
        "ate": max((c["data"] for c in cb if c["data"]), default=""),
    }
