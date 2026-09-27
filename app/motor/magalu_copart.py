"""COPARTICIPAÇÃO DE FRETE · REAL POR PEDIDO (Magazine Luiza) — o arquivo
SXC do portal (MagazineLuiza_CSV-<comp>_CoparticipacaoDeFretes_*.zip), que traz
UMA LINHA POR PEDIDO × SKU com o valor de frete que NÓS pagamos.

POR QUE ELE EXISTE (27/09/2026)
  Até aqui o PLUTOS tinha duas fontes e nenhuma das duas era o real:
    · o financeiro traz "Coparticipação de fretes ESTIMADA" — e vem quase
      sempre zerada (set/26: R$ 3.082 em 165 pedidos, de 838 do Fulfillment);
    · a tabela por SKU é PREÇO DE TABELA, serve para precificar, não para
      fechar o mês.
  Este arquivo é o realizado: pedido, SKU e o que foi efetivamente cobrado.

DUAS ARMADILHAS DO ARQUIVO, as duas tratadas aqui:

1) AS COLUNAS DE DATA VÊM TROCADAS.
   O portal grava "data do pedido" no lugar de "data cobrança" e vice-versa.
   Medido em set/26: em 2.862 de 2.862 linhas a "data cobrança" é ANTERIOR à
   "data do pedido" (9,7 dias em média) — cobrança antes da venda não existe.
   Cruzando 741 pedidos com a base do PLUTOS, a data REAL do pedido bateu
   729 vezes com a coluna rotulada "data cobrança" e ZERO vez com a rotulada
   "data do pedido". Então aqui: data_pedido = coluna "data cobrança".
   A COMPETÊNCIA é a data do pedido — a mesma régua dos outros canais.

2) NÃO EXISTE "A TABELA VIGENTE": existe tabela CHEIA e DESCONTO DE CAMPANHA.
   O canal cobra por serviço:
     · "Política Fulfillment"   → valor CHEIO do SKU
     · "Campanha ML Entregas"   → o mesmo SKU com desconto
   Medido em set/26, o desconto da campanha cai em degraus limpos sobre o
   cheio: 0,60× (−40%, o mais comum), 0,46× (−54%) e uma faixa de 0,10–0,17×
   (−83% a −90%, o frete quase grátis das campanhas pesadas).
   É por isso que precificar tudo pela "tabela nova −40%" erra: mais da
   metade do volume paga CHEIO.

O CHEIO DE CADA SKU sai do próprio arquivo: é o valor mais repetido nas
linhas de Política Fulfillment daquele SKU. Sem nenhuma linha cheia, cai
para o maior valor visto. Com o cheio na mão, toda linha ganha:
   desconto_pct = 1 − valor / cheio
e a linha acima do cheio (desconto negativo) vira PENDÊNCIA: o canal cobrou
mais do que a própria tabela dele — é dinheiro a reclamar, não é rebate.
"""
from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from collections import Counter
from datetime import date
from typing import Any

import pandas as pd

CANAL = "magalu"
S_CHEIO = "Política Fulfillment"
S_CAMPANHA = "Campanha ML Entregas"
# degraus de desconto que o canal pratica (para nomear a faixa na tela)
DEGRAUS = ((0.40, "campanha −40%"), (0.54, "campanha −54%"), (0.83, "campanha pesada"))
TOL_DEGRAU = 0.03
# acima disso a cobrança passou do cheio do próprio canal: pendência
TOL_ACIMA = 0.01


def _norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(v) -> float:
    """BR: 1.234,56 · US: 1,234.56 — decide pelo ÚLTIMO separador."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v == v else 0.0
    t = str(v).strip().replace("R$", "").replace(" ", "")
    if not t or t in ("-", "nan", "None"):
        return 0.0
    uv, up = t.rfind(","), t.rfind(".")
    if uv >= 0 and up >= 0:
        t = t.replace(".", "").replace(",", ".") if uv > up else t.replace(",", "")
    elif uv >= 0:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _data(v) -> str:
    t = str(v or "").strip()[:10]
    if not t:
        return ""
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return date(*reversed([int(x) for x in re.split(r"[/-]", t)])).isoformat() if f == "%d/%m/%Y" \
                else pd.to_datetime(t, format=f).date().isoformat()
        except Exception:
            continue
    d = pd.to_datetime(t, errors="coerce", dayfirst=True)
    return "" if d is not pd.NaT and d != d else (d.date().isoformat() if d is not pd.NaT else "")


COLS = {
    "pedido": "pedido", "sku": "sku",
    "grupo servico": "grupo", "servico": "servico", "beneficio(%)": "beneficio",
    # ATENÇÃO: os rótulos vêm TROCADOS no portal — ver o cabeçalho deste arquivo
    "data do pedido": "_rot_pedido", "data cobranca": "_rot_cobranca",
    "peso cubado(kg/m³)": "cubagem", "peso cubado(kg/m3)": "cubagem",
    "valor coparticipacao": "valor", "peso(kg)": "peso",
}
OBRIGATORIAS = ["pedido", "sku", "valor"]


def _abrir(caminho: str) -> pd.DataFrame:
    """O portal entrega .zip com um CSV dentro; aceita o CSV solto também."""
    if caminho.lower().endswith(".zip"):
        with zipfile.ZipFile(caminho) as z:
            nomes = [n for n in z.namelist() if n.lower().endswith((".csv", ".txt"))
                     and not n.startswith("__MACOSX")]
            if not nomes:
                raise ValueError("O .zip da coparticipação não tem nenhum CSV dentro.")
            bruto = z.read(nomes[0])
        return pd.read_csv(io.BytesIO(bruto), sep=";", dtype=str, encoding="utf-8-sig")
    if caminho.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(caminho, dtype=str)
    return pd.read_csv(caminho, sep=None, engine="python", dtype=str, encoding="utf-8-sig")


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    d = _abrir(caminho)
    mapa = {c: COLS[_norm(c)] for c in d.columns if _norm(c) in COLS}
    df = d.rename(columns=mapa).copy()
    if not all(k in df.columns for k in OBRIGATORIAS):
        raise ValueError(
            "Não parece o arquivo de COPARTICIPAÇÃO DE FRETES do portal Magalu "
            "(MagazineLuiza_CSV-AAAAMM_CoparticipacaoDeFretes_*.zip). "
            "Preciso das colunas pedido, sku e valor coparticipação.")
    for k in set(COLS.values()):
        if k not in df:
            df[k] = None
    for k in ("valor", "cubagem", "peso", "beneficio"):
        df[k] = df[k].map(_num)
    for k in ("pedido", "sku", "servico", "grupo"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    df = df[(df["pedido"] != "") & (df["valor"] != 0)].copy()

    # ---- AS DUAS DATAS, DESTROCADAS (ver o cabeçalho) --------------------
    a = df["_rot_pedido"].map(_data)     # rótulo "data do pedido"
    b = df["_rot_cobranca"].map(_data)   # rótulo "data cobrança"
    n_inv = int(((b != "") & (a != "") & (b < a)).sum())
    n_ok = int(((b != "") & (a != "") & (b > a)).sum())
    trocadas = n_inv > n_ok   # cobrança "antes" do pedido na maioria = rótulos trocados
    df["data"] = b if trocadas else a
    df["data_cobranca"] = a if trocadas else b
    df = df[df["data"] != ""].copy()
    df["competencia"] = df["data"].str[:7]

    # ---- CHEIO POR SKU, DESCONTO E FAIXA ---------------------------------
    cheio: dict[str, float] = {}
    for sku, g in df[df["servico"].map(_norm) == _norm(S_CHEIO)].groupby("sku"):
        c = Counter(round(v, 2) for v in g["valor"])
        cheio[sku] = c.most_common(1)[0][0]
    for sku, g in df.groupby("sku"):       # SKU que só apareceu em campanha
        cheio.setdefault(sku, round(float(g["valor"].max()), 2))
    df["cheio"] = df["sku"].map(cheio)
    df["desconto_pct"] = [round(1 - (v / c), 4) if c else 0.0 for v, c in zip(df["valor"], df["cheio"])]
    df["economia"] = [round(c - v, 2) for v, c in zip(df["valor"], df["cheio"])]
    df["faixa"] = [_faixa(p) for p in df["desconto_pct"]]
    df["acima_do_cheio"] = df["desconto_pct"] < -TOL_ACIMA
    # chave da base acumulada: o mesmo pedido pode ter 2 SKUs e 2 cobranças
    df["chave"] = (df["pedido"] + "|" + df["sku"] + "|" + df["servico"] + "|"
                   + df["data"] + "|" + df["valor"].map(lambda v: f"{v:.2f}"))
    df = df.drop(columns=["_rot_pedido", "_rot_cobranca"], errors="ignore")

    acima = df[df["acima_do_cheio"]]
    diag = {
        "linhas": len(df), "pedidos": int(df["pedido"].nunique()), "skus": int(df["sku"].nunique()),
        "total": round(float(df["valor"].sum()), 2),
        "de": (df["data"].min() if len(df) else ""), "ate": (df["data"].max() if len(df) else ""),
        "cob_de": (df["data_cobranca"].min() if len(df) else ""),
        "cob_ate": (df["data_cobranca"].max() if len(df) else ""),
        "datas_trocadas": bool(trocadas), "n_invertidas": n_inv,
        "competencias": sorted(set(df["competencia"])),
        "servicos": {k: int(v) for k, v in df["servico"].value_counts().items()},
        "acima_do_cheio": len(acima), "acima_rs": round(float((-acima["economia"]).sum()), 2),
    }
    return df, diag


def _faixa(pct: float) -> str:
    if pct < -TOL_ACIMA:
        return "acima do cheio"
    if abs(pct) <= TOL_DEGRAU:
        return "cheio"
    for alvo, nome in DEGRAUS:
        if abs(pct - alvo) <= TOL_DEGRAU:
            return nome
    if pct >= 0.70:
        return "campanha pesada"
    return f"campanha −{round(100 * pct)}%"


def resumo(df: pd.DataFrame, nomes: dict | None = None) -> dict:
    """O que a tela mostra. `nomes` = {sku: descrição} para dar nome ao SKU."""
    nomes = nomes or {}
    if df is None or not len(df):
        return {"linhas": 0, "total": 0.0, "pedidos": 0, "skus": 0, "por_sku": [],
                "por_servico": {}, "por_faixa": {}, "acima": [], "acima_rs": 0.0,
                "cheio_total": 0.0, "economia": 0.0, "economia_pct": 0.0,
                "por_dia": {}, "un_media": 0.0, "pct_cheio": 0.0}
    tot = round(float(df["valor"].sum()), 2)
    cheio_tot = round(float(df["cheio"].sum()), 2)

    por_servico = {}
    for s, g in df.groupby("servico"):
        por_servico[s] = {"linhas": len(g), "total": round(float(g["valor"].sum()), 2),
                          "media": round(float(g["valor"].mean()), 2),
                          "share": round(100 * float(g["valor"].sum()) / tot, 1) if tot else 0.0}
    por_faixa = {}
    for f, g in df.groupby("faixa"):
        por_faixa[f] = {"linhas": len(g), "total": round(float(g["valor"].sum()), 2),
                        "cheio": round(float(g["cheio"].sum()), 2),
                        "economia": round(float(g["economia"].sum()), 2),
                        "media": round(float(g["valor"].mean()), 2)}
    por_faixa = dict(sorted(por_faixa.items(), key=lambda kv: -kv[1]["total"]))

    por_sku = []
    for sku, g in df.groupby("sku"):
        v = float(g["valor"].sum())
        n_cheio = int((g["faixa"] == "cheio").sum())
        por_sku.append({
            "sku": sku, "descricao": nomes.get(sku, ""), "linhas": len(g),
            "total": round(v, 2), "un_media": round(v / len(g), 2),
            "cheio": round(float(g["cheio"].iloc[0]), 2),
            "min": round(float(g["valor"].min()), 2), "max": round(float(g["valor"].max()), 2),
            "cubagem": round(float(g["cubagem"].iloc[0]), 3),
            "peso": round(float(g["peso"].iloc[0]), 1),
            "n_cheio": n_cheio, "pct_cheio": round(100 * n_cheio / len(g), 0),
            "economia": round(float(g["economia"].sum()), 2),
            "acima": int(g["acima_do_cheio"].sum()),
        })
    por_sku.sort(key=lambda x: -x["total"])

    acima = [{"pedido": r.pedido, "sku": r.sku, "descricao": nomes.get(r.sku, ""), "data": r.data,
              "servico": r.servico, "valor": round(float(r.valor), 2),
              "cheio": round(float(r.cheio), 2), "excesso": round(float(-r.economia), 2)}
             for r in df[df["acima_do_cheio"]].itertuples()]
    acima.sort(key=lambda x: -x["excesso"])

    por_dia = {}
    for d, g in df.groupby("data"):
        por_dia[d] = {"linhas": len(g), "total": round(float(g["valor"].sum()), 2)}
    por_dia = dict(sorted(por_dia.items()))

    n_cheio = int((df["faixa"] == "cheio").sum())
    return {
        "linhas": len(df), "pedidos": int(df["pedido"].nunique()), "skus": int(df["sku"].nunique()),
        "total": tot, "un_media": round(tot / len(df), 2),
        "cheio_total": cheio_tot, "economia": round(cheio_tot - tot, 2),
        "economia_pct": round(100 * (cheio_tot - tot) / cheio_tot, 1) if cheio_tot else 0.0,
        "pct_cheio": round(100 * n_cheio / len(df), 1),
        "por_servico": por_servico, "por_faixa": por_faixa, "por_sku": por_sku,
        "acima": acima, "acima_rs": round(sum(a["excesso"] for a in acima), 2),
        "por_dia": por_dia,
        "de": df["data"].min(), "ate": df["data"].max(),
    }


def por_sku_medio(df: pd.DataFrame) -> dict:
    """{sku: coparticipação REAL média por unidade} — é o que o ORION soma no
    preço, no lugar do valor de tabela. Média ponderada do que foi cobrado."""
    if df is None or not len(df):
        return {}
    g = df.groupby("sku")["valor"]
    return {sku: round(float(v), 2) for sku, v in (g.sum() / g.size()).items()}
