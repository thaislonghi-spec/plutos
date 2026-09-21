"""CUSTO DO FULL POR ITEM · REAL (Magalu) — a TABELA DE COPARTICIPAÇÃO DE FRETE
por SKU, que o canal envia (Coparticipação Full_DDMMAAAA.xlsx).

Diferença para a tela "Estimado":
  · Estimado — o PLUTOS RATEIA o que o canal cobra por agenda/m³ e por movimento,
    porque na origem não vem SKU. É a melhor conta possível com o que ele entrega.
  · Real — aqui o valor já vem POR SKU, direto da tabela do canal. Não há rateio.

O ARQUIVO (1 linha por SKU):
  SKU MAGALU · SKU MULTI · SKU CANAL ·
  COPARTICIPAÇÃO TABELA ANTIGA · COPARTICIPAÇÃO TABELA NOVA -40% ·
  COPARTICIPAÇÃO TABELA NOVA · DIFERENCA DE CUSTO DE FRETE

TRÊS PREÇOS NA MESMA LINHA, e só um deles é o que vale hoje:
  antiga     — a tabela que vinha sendo praticada
  nova_40    — a tabela nova com 40% de desconto (é exatamente 60% da nova)
  nova       — a tabela nova cheia
Qual delas vigora é decisão da casa e fica em Parâmetros (tabela_copart_vigente);
o PLUTOS guarda as três para simular o impacto de uma virar a outra.

A chave é o SKU MULTI (o nosso), normalizado com e sem ponto.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any

import pandas as pd

TABELAS = ("antiga", "nova_40", "nova")
NOMES = {"antiga": "Tabela antiga", "nova_40": "Tabela nova −40%", "nova": "Tabela nova (cheia)"}


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
    ult_v, ult_p = t.rfind(","), t.rfind(".")
    if ult_v >= 0 and ult_p >= 0:
        t = t.replace(".", "").replace(",", ".") if ult_v > ult_p else t.replace(",", "")
    elif ult_v >= 0:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


COLS = {
    "sku magalu": "sku_magalu", "sku multi": "sku", "sku canal": "sku_canal",
    "coparticipacao tabela antiga": "antiga",
    "coparticipacao tabela nova -40%": "nova_40",
    "coparticipacao tabela nova": "nova",
    "diferenca de custo de frete": "diferenca",
}
OBRIGATORIAS = ["sku", "antiga", "nova"]


def ler(caminho: str) -> tuple[list[dict], dict]:
    d = (pd.read_excel(caminho, dtype=str) if caminho.lower().endswith((".xlsx", ".xls"))
         else pd.read_csv(caminho, dtype=str, sep=None, engine="python", encoding="utf-8-sig"))
    mapa = {}
    for c in d.columns:
        n = _norm(c)
        if n in COLS and COLS[n] not in mapa.values():
            mapa[c] = COLS[n]
        elif n.startswith("coparticipacao tabela nova") and "40" in n and "nova_40" not in mapa.values():
            mapa[c] = "nova_40"
    df = d.rename(columns=mapa).copy()
    if not all(k in df.columns for k in OBRIGATORIAS):
        raise ValueError("Não parece a tabela de COPARTICIPAÇÃO do Full (Coparticipação Full_DDMMAAAA.xlsx). "
                         "Preciso de SKU MULTI e das colunas de coparticipação (tabela antiga e tabela nova).")
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("antiga", "nova_40", "nova", "diferenca"):
        df[k] = df[k].map(_num)
    for k in ("sku", "sku_magalu", "sku_canal"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    df = df[df["sku"] != ""].copy()

    # a coluna "-40%" pode não vir: ela é exatamente 60% da nova
    falta40 = int((df["nova_40"] <= 0).sum())
    df.loc[df["nova_40"] <= 0, "nova_40"] = (df["nova"] * 0.6).round(2)

    itens = []
    for r in df.itertuples(index=False):
        itens.append({"sku": r.sku, "sku_magalu": r.sku_magalu, "sku_canal": r.sku_canal,
                      "antiga": round(r.antiga, 2), "nova_40": round(r.nova_40, 2),
                      "nova": round(r.nova, 2),
                      "diferenca": round(r.diferenca or (r.nova - r.antiga), 2)})
    s = lambda k: round(sum(i[k] for i in itens), 2)  # noqa: E731
    med = lambda k: round(s(k) / len(itens), 2) if itens else 0.0  # noqa: E731
    diag = {
        "linhas": len(itens), "skus": len({i["sku"] for i in itens}),
        "soma": {k: s(k) for k in TABELAS}, "media": {k: med(k) for k in TABELAS},
        "sem_40": falta40,
        "alta_nova": (round(100 * (s("nova") / s("antiga") - 1), 1) if s("antiga") else 0.0),
        "alta_40": (round(100 * (s("nova_40") / s("antiga") - 1), 1) if s("antiga") else 0.0),
    }
    return itens, diag


def resumo(itens: list[dict], vendidos: dict, nomes: dict | None = None,
           vigente: str = "nova_40", chave=lambda s: s) -> dict:
    """Junta a tabela com as UNIDADES VENDIDAS no Full do mês: é assim que o
    valor por SKU vira dinheiro do mês. `vendidos` = {sku: unidades}."""
    nomes = nomes or {}
    idx = {chave(i["sku"]): i for i in itens}
    linhas, sem_tabela = [], []
    tot = {k: 0.0 for k in TABELAS}
    un_com = un_sem = 0.0
    for sku, un in (vendidos or {}).items():
        it = idx.get(chave(sku))
        if not it:
            un_sem += un
            sem_tabela.append({"sku": sku, "unidades": un, "descricao": nomes.get(sku, "")})
            continue
        un_com += un
        l = {"sku": sku, "descricao": nomes.get(sku, "") or nomes.get(it["sku"], ""),
             "sku_canal": it["sku_canal"], "unidades": un}
        for k in TABELAS:
            l[k] = it[k]
            l[k + "_mes"] = round(it[k] * un, 2)
            tot[k] += it[k] * un
        l["vigente_un"] = it[vigente]
        l["vigente_mes"] = round(it[vigente] * un, 2)
        l["sobe_vs_antiga"] = round(it[vigente] - it["antiga"], 2)
        linhas.append(l)
    linhas.sort(key=lambda l: -l["vigente_mes"])
    tot = {k: round(v, 2) for k, v in tot.items()}
    return {
        "linhas": linhas, "sem_tabela": sorted(sem_tabela, key=lambda x: -x["unidades"]),
        "vigente": vigente, "vigente_nome": NOMES.get(vigente, vigente),
        "unidades": round(un_com, 0), "unidades_sem": round(un_sem, 0),
        "skus": len(linhas), "skus_tabela": len(itens),
        "total": tot.get(vigente, 0.0), "por_tabela": tot,
        "un_media": {k: (round(v / un_com, 2) if un_com else 0.0) for k, v in tot.items()},
        "por_unidade": (round(tot.get(vigente, 0.0) / un_com, 2) if un_com else 0.0),
        "delta_nova": round(tot["nova"] - tot["antiga"], 2),
        "delta_40": round(tot["nova_40"] - tot["antiga"], 2),
        "alta_nova": (round(100 * (tot["nova"] / tot["antiga"] - 1), 1) if tot["antiga"] else 0.0),
        "alta_40": (round(100 * (tot["nova_40"] / tot["antiga"] - 1), 1) if tot["antiga"] else 0.0),
    }
