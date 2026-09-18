"""Box MADEIRA MADEIRA — regra escrita pela Thaís (12/09/2026).

Fonte: relatório de pedidos do portal Madeira "MadeiraMadeira_<seller>-report_pedido_<hash>"
(csv separado por ';', latin-1, 41 colunas, 1 linha por ITEM; pode vir sem extensão).

- FORA: Status "Cancelado" (no Madeira a comissão NÃO zera no cancelado) e Status "Novo"
  (pedido que ainda não andou; entra sozinho quando o arquivo seguinte atualizar o status).
- 1 linha por pedido: "Valor Pedido" e "Comissão" repetem em cada item; "Valor" (item) soma.
- Competência = Data Pedido (dd/mm/aaaa hh:mm:ss → só a data).
- Chave = "Pedido" (= OC do ERP).
- COMISSÃO SISTEMA = % da tabela de Parâmetros (Madeira 17% GMV) × Valor Pedido (com frete).
- COMISSÃO REAL = coluna Comissão.
- REBATE EM COMISSÃO = sistema − real. Rebate em R$ = 0 · Rebate em frete = 0.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from datetime import date, datetime
from typing import Any

import pandas as pd

from . import arred


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
    for f in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


COLS = {
    "pedido": "pedido", "pedido site mm": "pedido_site", "data pedido": "data", "data aprovacao": "data_aprov",
    "valor pedido": "valor_pedido", "comissao": "comissao", "tipo pagamento": "pagamento", "status": "status",
    "cidade": "cidade", "uf": "uf", "sku": "sku", "produto": "produto", "quantidade": "qtd",
    "valor original": "valor_original", "% desconto": "pct_desc", "valor": "valor_item",
    "numero nf": "nf", "data emissao": "data_nf", "nome da transportadora": "transportadora", "parcelas": "parcelas",
}
OBRIGATORIAS = ["pedido", "data", "valor_pedido", "comissao", "status", "sku"]


def _ler_bruto(caminho: str) -> pd.DataFrame:
    """csv ';' latin-1 (com ou sem extensão) ou xlsx."""
    with open(caminho, "rb") as f:
        cab = f.read(4)
    if cab[:2] == b"PK":
        return pd.read_excel(caminho, dtype=str)
    with open(caminho, "rb") as f:
        raw = f.read()
    txt = raw.decode("utf-8-sig") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("latin-1")
    sep = ";" if txt[:3000].count(";") >= txt[:3000].count("\t") else "\t"
    return pd.read_csv(io.StringIO(txt), sep=sep, dtype=str, quoting=csv.QUOTE_MINIMAL)


def fora_da_conta(status) -> bool:
    """FORA DA CONTA (regra da Thaís, 16/09/2026): Status "Cancelado" (no Madeira a
    comissão não zera no cancelado, filtrar é obrigatório) e Status "Novo" (o pedido
    ainda nem começou — quando andar, o arquivo seguinte atualiza e ele entra sozinho).
    Vale na leitura E no recálculo, para alcançar o que já está na base."""
    st = _norm(status)
    return bool("cancelado" in st or st.startswith("novo"))


def marcar_fora(ped: pd.DataFrame) -> pd.DataFrame:
    st = ped["status"].map(_norm)
    ped["novo"] = st.str.startswith("novo")
    ped["cancelado"] = st.str.contains("cancelado") | ped["novo"]
    return ped


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    d = _ler_bruto(caminho)
    mapa = {}
    for c in d.columns:
        n = _norm(c).strip('"')
        if n in COLS and COLS[n] not in mapa.values():
            mapa[c] = COLS[n]
    if not all(k in mapa.values() for k in OBRIGATORIAS):
        raise ValueError("Não parece o relatório de pedidos do Madeira (preciso de Pedido, Data Pedido, Valor Pedido, Comissão, Status, SKU).")
    df = d.rename(columns=mapa).copy()
    itens = int(len(df))
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("pedido", "pedido_site", "status", "pagamento", "cidade", "uf", "sku", "produto", "nf", "transportadora"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    for k in ("valor_pedido", "comissao", "qtd", "valor_original", "pct_desc", "valor_item", "parcelas"):
        df[k] = df[k].map(_num)
    df["data"] = df["data"].map(_data)
    df["data_nf"] = df["data_nf"].map(lambda v: str(_data(v)) if _data(v) else "")
    df = df[(df["pedido"] != "") & df["data"].notna()].copy()
    primeira = df.groupby("pedido", sort=False).first()
    somas = df.groupby("pedido", sort=False).agg(valor_item=("valor_item", "sum"), qtd=("qtd", "sum"), itens=("sku", "size"),
                                                  skus=("sku", lambda s: ", ".join(dict.fromkeys(x for x in s if x))))
    ped = primeira.drop(columns=["valor_item", "qtd"]).join(somas).reset_index()
    ped = marcar_fora(ped)
    ped["competencia"] = ped["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    diag: dict[str, Any] = {
        "aba": "csv", "linhas_brutas": itens, "itens": itens, "linhas": int(len(ped)),
        "cancelados": int(ped["cancelado"].sum()), "multi_item": int((ped["itens"] > 1).sum()),
        "novos": int(ped["novo"].sum()),
        "status": ped["status"].value_counts().to_dict(),
        "competencias": ped["competencia"].value_counts().sort_index().to_dict(),
        "de": str(ped["data"].min()), "ate": str(ped["data"].max()),
    }
    return ped, diag


def calcular(df: pd.DataFrame, pct: float, erp_idx: dict | None = None) -> list[dict]:
    out = []
    erp_idx = erp_idx or {}
    for r in df.itertuples(index=False):
        if fora_da_conta(getattr(r, "status", "")):
            continue
        sis_rs = round(r.valor_pedido * pct, 2)
        real = round(r.comissao, 2)
        reb_com = round(sis_rs - real, 2)
        e = erp_idx.get(r.pedido)
        out.append({
            "canal": "madeira",
            "pedido_mkt": r.pedido, "pedido_canal": r.pedido_site, "pedido_any": (e or {}).get("obs05", "") or "",
            "id_mkt": "", "data": str(r.data), "competencia": r.competencia,
            "conta": "", "status": r.status, "pagamento": r.pagamento, "uf": r.uf, "cidade": r.cidade,
            "sku": r.sku, "anuncio": r.skus, "tipo": r.pagamento, "produto": r.produto, "itens": int(r.itens), "qtd": float(r.qtd),
            "valor_prod": r.valor_pedido, "valor_itens": r.valor_item, "tarifa": real, "frete": 0.0,
            "cupom_seller": 0.0, "cupom_meli": 0.0,
            "pct_comissao": (real / r.valor_pedido if r.valor_pedido else 0.0),
            "sis_pct": pct, "sis_rs": sis_rs, "diferenca": reb_com,
            "nf": r.nf, "data_nf": r.data_nf, "transportadora": r.transportadora, "parcelas": r.parcelas,
            "tarifa_zero": bool(real <= 0.005 and sis_rs > 0), "faltante": 0.0, "faltante_status": "", "sem_sistema": False,
            "rebate_rs": 0.0, "rebate_comissao": reb_com, "rebate_frete": 0.0,
            "rebate_total": reb_com,
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
        d["pedidos"] += 1; d["venda"] += l["valor_prod"]; d["comissao"] += arred.sis_exato(l, False) - (l["tarifa"] or 0.0); d["total"] += (l.get("rebate_rs") or 0.0) + (arred.sis_exato(l, False) - (l["tarifa"] or 0.0)) + (l.get("rebate_frete") or 0.0)
        d["tz"] += 1 if l["tarifa_zero"] else 0
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "frete", "total"):
            d[k] = round(d[k], 2)
    com_sis, com_real = arred.total(linhas, False), s("tarifa")  # arred. 1x no total
    reb_com = round(com_sis - com_real, 2)
    reb_tot = round(0.0 + reb_com + 0.0, 2)
    pag: dict[str, int] = {}
    for l in linhas:
        pag[l["pagamento"] or "—"] = pag.get(l["pagamento"] or "—", 0) + 1
    return {
        "pedidos": n, "cancelados": cancelados, "venda": venda, "tarifa": com_real, "frete": 0.0,
        "cupom_meli": 0.0, "cupom_seller": 0.0, "faltante": 0.0, "valor_itens": s("valor_itens"),
        "rebate_rs": 0.0, "rebate_comissao": reb_com, "rebate_frete": 0.0, "rebate_total": reb_tot,
        "pct_sobre_venda": (round(100 * reb_tot / venda, 2) if venda else 0.0),
        "tarifa_zero": sum(1 for l in linhas if l["tarifa_zero"]), "faltante_pendentes": 0, "faltante_preenchidos": 0,
        "dif_pos": sum(1 for l in linhas if l["diferenca"] > 0.5), "dif_pos_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] > 0.5), 2),
        "dif_neg": sum(1 for l in linhas if l["diferenca"] < -0.5), "dif_neg_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] < -0.5), 2),
        "sem_sistema": 0, "erp_ok": sum(1 for l in linhas if l.get("erp_ok")),
        "com_pedidos": n, "com_sistema": com_sis, "com_real": com_real, "com_venda": venda,
        "com_sistema_pct": (round(100 * com_sis / venda, 2) if venda else 0.0),
        "com_real_pct": (round(100 * com_real / venda, 2) if venda else 0.0),
        "com_dif": round(com_sis - com_real, 2),
        "pagamentos": pag, "multi_item": sum(1 for l in linhas if l["itens"] > 1), "contas": {}, "tipos": pag,
        "por_dia": dict(sorted(por_dia.items())),
    }
