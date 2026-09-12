"""Box WEBCONTINENTAL — regra oficial da Thaís (12/09/2026), mesma lógica do Madeira.

Fonte: export do portal "relatorio_pedidos_webcontinental_DDMMateDDMMAA.xlsx", aba "Pedidos",
40 colunas, 1 linha por pedido. Valores vêm como texto "527,99".

- Excluir Status Atual = "Cancelado", "Cancelado isento" e "Não Autorizado".
- Competência = Data Criação (dd/mm/aaaa).
- COMISSÃO SISTEMA = % da tabela de Parâmetros (Webcontinental 19% GMV) × Total do Pedido.
- COMISSÃO REAL = Valor de Comissão Retido.
- REBATE EM COMISSÃO = sistema − real. Rebate em R$ = 0 · Rebate em frete = 0.
- Chave: Pedido Parceiro (acumula por ele); Pedido ERP = OC do ERP.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from datetime import date, datetime
from typing import Any

import pandas as pd


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
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    t = str(v).strip()[:10]
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


COLS = {
    "pedido parceiro": "pedido", "pedido site": "pedido_site", "pedido erp": "oc",
    "data criacao": "data", "status atual": "status",
    "total do pedido": "total", "valor do frete": "frete", "valor dos produtos": "valor_prod_canal",
    "desconto": "desconto", "valor repasse": "repasse", "valor de comissao retido": "comissao",
    "cliente": "cliente", "uf": "uf", "cidade": "cidade", "forma de pagamento": "pagamento",
    "sku": "sku", "produto": "produto", "quantidade": "qtd", "nota fiscal": "nf", "transportadora": "transportadora",
}
OBRIGATORIAS = ["pedido", "data", "status", "total", "comissao"]


def _ler_bruto(caminho: str):
    with open(caminho, "rb") as f:
        cab = f.read(4)
    if cab[:2] == b"PK":
        xl = pd.ExcelFile(caminho)
        abas = [a for a in xl.sheet_names if _norm(a) == "pedidos"] or xl.sheet_names
        for aba in abas:
            d = xl.parse(aba, dtype=str)
            if any(_norm(c).startswith("pedido parceiro") for c in d.columns):
                return d
        return xl.parse(abas[0], dtype=str)
    with open(caminho, "rb") as f:
        raw = f.read()
    txt = raw.decode("utf-8-sig") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("latin-1")
    sep = ";" if txt[:3000].count(";") >= txt[:3000].count("\t") else "\t"
    return pd.read_csv(io.StringIO(txt), sep=sep, dtype=str, quoting=csv.QUOTE_MINIMAL)


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    d = _ler_bruto(caminho)
    mapa = {}
    for c in d.columns:
        n = _norm(c)
        for ini, k in COLS.items():
            if (n == ini or n.startswith(ini)) and k not in mapa.values():
                mapa[c] = k
                break
    if not all(k in mapa.values() for k in OBRIGATORIAS):
        raise ValueError("Não parece o relatório de pedidos da Webcontinental (preciso de Pedido Parceiro, Data Criação, "
                         "Status Atual, Total do Pedido, Valor de Comissão Retido).")
    df = d.rename(columns=mapa).copy()
    bruto = int(len(df))
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("pedido", "pedido_site", "oc", "status", "cliente", "uf", "cidade", "pagamento", "sku", "produto", "nf", "transportadora"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    for k in ("total", "frete", "valor_prod_canal", "desconto", "repasse", "comissao", "qtd"):
        df[k] = df[k].map(_num)
    df["data"] = df["data"].map(_data)
    df = df[(df["pedido"] != "") & df["data"].notna()].copy()
    df = df.drop_duplicates(subset=["pedido"], keep="last")
    st = df["status"].map(_norm)
    # fora: "Cancelado", "Cancelado isento" e "Não Autorizado" (decisão da Thaís 12/09: não soma)
    df["cancelado"] = st.str.startswith("cancelado") | st.eq("nao autorizado")
    df["competencia"] = df["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    diag: dict[str, Any] = {
        "aba": "Pedidos", "linhas_brutas": bruto, "linhas": int(len(df)),
        "cancelados": int(df["cancelado"].sum()),
        "status": df["status"].value_counts().to_dict(),
        "competencias": df["competencia"].value_counts().sort_index().to_dict(),
        "de": str(df["data"].min()), "ate": str(df["data"].max()),
        "com_oc": int((df["oc"] != "").sum()),
    }
    return df, diag


def calcular(df: pd.DataFrame, pct: float, erp_idx: dict | None = None) -> list[dict]:
    out = []
    erp_idx = erp_idx or {}
    for r in df.itertuples(index=False):
        if r.cancelado:
            continue
        sis_rs = round(r.total * pct, 2)
        real = round(r.comissao, 2)
        reb_com = round(sis_rs - real, 2)
        oc = r.oc or r.pedido
        e = erp_idx.get(oc) or erp_idx.get(r.pedido)
        out.append({
            "canal": "webcont",
            "pedido_mkt": oc, "pedido_canal": r.pedido, "pedido_any": (e or {}).get("obs05", "") or "",
            "id_mkt": r.pedido_site, "data": str(r.data), "competencia": r.competencia,
            "conta": "", "status": r.status, "pagamento": r.pagamento, "uf": r.uf, "cidade": r.cidade,
            "sku": r.sku, "anuncio": "", "tipo": r.pagamento, "produto": r.produto, "itens": 1, "qtd": float(r.qtd),
            "valor_prod": r.total, "valor_itens": r.valor_prod_canal, "tarifa": real, "frete": r.frete,
            "desconto": r.desconto, "repasse": r.repasse, "cupom_seller": 0.0, "cupom_meli": 0.0,
            "pct_comissao": (real / r.total if r.total else 0.0),
            "sis_pct": pct, "sis_rs": sis_rs, "diferenca": reb_com,
            "nf": r.nf, "transportadora": r.transportadora,
            "tarifa_zero": bool(real <= 0.005 and sis_rs > 0), "faltante": 0.0, "faltante_status": "", "sem_sistema": False,
            "rebate_rs": 0.0, "rebate_comissao": reb_com, "rebate_frete": 0.0, "rebate_total": reb_com,
            "erp_ok": bool(e),
        })
    return out


def resumo(linhas: list[dict], cancelados: int = 0) -> dict:
    n = len(linhas)
    s = lambda k: round(sum((l.get(k) or 0.0) for l in linhas), 2)  # noqa: E731
    venda = s("valor_prod")
    por_dia: dict[str, dict] = {}
    for l in linhas:
        d = por_dia.setdefault(l["data"], {"pedidos": 0, "venda": 0.0, "cupom": 0.0, "faltante": 0.0,
                                            "comissao": 0.0, "frete": 0.0, "total": 0.0, "tz": 0})
        d["pedidos"] += 1; d["venda"] += l["valor_prod"]; d["comissao"] += l["rebate_comissao"]; d["total"] += l["rebate_total"]
        d["tz"] += 1 if l["tarifa_zero"] else 0
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "frete", "total"):
            d[k] = round(d[k], 2)
    com_sis, com_real = s("sis_rs"), s("tarifa")
    sts: dict[str, int] = {}
    for l in linhas:
        sts[l["status"] or "—"] = sts.get(l["status"] or "—", 0) + 1
    return {
        "pedidos": n, "cancelados": cancelados, "venda": venda, "tarifa": com_real, "frete": s("frete"),
        "cupom_meli": 0.0, "cupom_seller": 0.0, "faltante": 0.0, "valor_itens": s("valor_itens"), "desconto": s("desconto"),
        "rebate_rs": 0.0, "rebate_comissao": s("rebate_comissao"), "rebate_frete": 0.0, "rebate_total": s("rebate_total"),
        "pct_sobre_venda": (round(100 * s("rebate_total") / venda, 2) if venda else 0.0),
        "tarifa_zero": sum(1 for l in linhas if l["tarifa_zero"]), "faltante_pendentes": 0, "faltante_preenchidos": 0,
        "dif_pos": sum(1 for l in linhas if l["diferenca"] > 0.5), "dif_pos_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] > 0.5), 2),
        "dif_neg": sum(1 for l in linhas if l["diferenca"] < -0.5), "dif_neg_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] < -0.5), 2),
        "sem_sistema": 0, "erp_ok": sum(1 for l in linhas if l.get("erp_ok")),
        "com_pedidos": n, "com_sistema": com_sis, "com_real": com_real, "com_venda": venda,
        "com_sistema_pct": (round(100 * com_sis / venda, 2) if venda else 0.0),
        "com_real_pct": (round(100 * com_real / venda, 2) if venda else 0.0),
        "com_dif": round(com_sis - com_real, 2),
        "status_validos": sts, "cheios": sum(1 for l in linhas if abs(l["pct_comissao"] - l["sis_pct"]) < 0.002),
        "contas": {}, "tipos": sts,
        "por_dia": dict(sorted(por_dia.items())),
    }
