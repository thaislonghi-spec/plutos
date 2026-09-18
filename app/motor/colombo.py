"""Box COLOMBO — regra da Gabi (14/09/2026), validada por Otto no arquivo
"Colombo_Pedidos_DDMMateDDMMAA.csv" (export do portal, csv `;` com aspas).

O arquivo vem com 1 LINHA POR ITEM e 40 colunas. O cuidado nº 1 é a DUPLA
CONTAGEM: em pedido com mais de um item, as colunas de CABEÇALHO (Valor
Mercadorias, Valor Frete, Valor Total) se REPETEM em cada linha; só as colunas
de ITEM (Valor Total Produto, Valor Comissão, Quantidade) é que somam.
No arquivo de 01–14/09 somar linha a linha inflava a venda em R$ 3.449,94.

- 1 linha por PEDIDO: cabeçalho = primeira linha · itens = soma.
- FORA: Status "Cancelado" e "Incluído" (decisão da Gabi).
- Competência = Data Pedido (dd/mm/aaaa).
- COMISSÃO REAL = soma de "Valor Comissão" (o que o Colombo cobrou de fato).
- COMISSÃO SISTEMA = % da tabela de Parâmetros (Colombo 7%) × Valor Total do
  pedido (mercadorias + frete − desconto = soma dos "Valor Total Produto").
- REBATE EM COMISSÃO = sistema − real. Rebate em R$ = 0 · Rebate em frete = 0.
  O % do item varia (5, 6 ou 7) conforme a negociação — é daí que vem o rebate.
- Chave = "Pedido" (acumula por ele). OC do ERP = "Entrega" (conferido:
  37 de 37 OCs do ERP casam por Entrega; por Pedido não casa nenhuma).
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
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


COLS = {
    "pedido": "pedido", "entrega": "oc", "data pedido": "data", "status": "status",
    "valor mercadorias": "merc", "valor frete": "frete", "valor total": "total",
    "data entrega": "data_entrega", "forma pagamento": "pagamento", "numero de parcelas": "parcelas",
    "nome do cliente": "cliente", "cidade": "cidade", "estado": "uf",
    "codigo produto colombo": "cod_colombo", "sku": "sku", "nome do produto": "produto",
    "quantidade": "qtd", "valor unitario": "valor_unit", "desconto": "desconto",
    "frete produto": "frete_item", "valor total produto": "total_item",
    "% comissao": "pct_item", "valor comissao": "comissao",
}
OBRIGATORIAS = ["pedido", "data", "status", "total", "total_item", "comissao", "sku"]


def _ler_bruto(caminho: str) -> pd.DataFrame:
    """csv ';' com aspas (utf-8 ou latin-1), ou xlsx."""
    with open(caminho, "rb") as f:
        cab = f.read(4)
    if cab[:2] == b"PK":
        return pd.read_excel(caminho, dtype=str)
    with open(caminho, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    sep = ";" if txt[:3000].count(";") >= txt[:3000].count("\t") else "\t"
    return pd.read_csv(io.StringIO(txt), sep=sep, dtype=str, quoting=csv.QUOTE_MINIMAL)


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    d = _ler_bruto(caminho)
    mapa = {}
    for c in d.columns:
        n = _norm(c).strip('"')
        if n in COLS and COLS[n] not in mapa.values():
            mapa[c] = COLS[n]
    if not all(k in mapa.values() for k in OBRIGATORIAS):
        raise ValueError("Não parece o relatório de pedidos do Colombo (preciso de Pedido, Data Pedido, Status, "
                         "Valor Total, Valor Total Produto, Valor Comissão e SKU).")
    df = d.rename(columns=mapa).copy()
    itens_brutos = int(len(df))
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("pedido", "oc", "status", "pagamento", "cliente", "cidade", "uf", "sku", "cod_colombo", "produto"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    for k in ("merc", "frete", "total", "qtd", "valor_unit", "desconto", "frete_item", "total_item",
              "pct_item", "comissao", "parcelas"):
        df[k] = df[k].map(_num)
    df["data"] = df["data"].map(_data)
    df["data_entrega"] = df["data_entrega"].map(lambda v: str(_data(v)) if _data(v) else "")
    df = df[(df["pedido"] != "") & df["data"].notna()].copy()

    # LINHAS IDÊNTICAS (mesmo pedido, mesmo item, mesmo valor) = duplicata do
    # export: entra uma vez só. Item repetido de verdade vem com Quantidade > 1.
    antes = len(df)
    df = df.drop_duplicates(subset=["pedido", "sku", "total_item", "comissao"], keep="first")
    dup_itens = antes - len(df)

    # 1 linha por pedido: cabeçalho da primeira linha + soma dos itens
    primeira = df.groupby("pedido", sort=False).first()
    somas = df.groupby("pedido", sort=False).agg(
        total_item=("total_item", "sum"), comissao=("comissao", "sum"), qtd=("qtd", "sum"),
        desconto=("desconto", "sum"), itens=("sku", "size"),
        skus=("sku", lambda s: ", ".join(dict.fromkeys(x for x in s if x))),
        pcts=("pct_item", lambda s: ", ".join(dict.fromkeys(f"{x:.0f}%" for x in s))))
    ped = primeira.drop(columns=["total_item", "comissao", "qtd", "desconto"]).join(somas).reset_index()

    st = ped["status"].map(_norm)
    # FORA: Cancelado e Incluído (regra da Gabi 14/09)
    ped["cancelado"] = st.str.startswith("cancelado") | st.str.startswith("incluido")
    ped["competencia"] = ped["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    diag: dict[str, Any] = {
        "aba": "csv", "linhas_brutas": itens_brutos, "itens": int(len(df)), "linhas": int(len(ped)),
        "itens_duplicados": int(dup_itens), "multi_item": int((ped["itens"] > 1).sum()),
        "cancelados": int(ped["cancelado"].sum()),
        "status": ped["status"].value_counts().to_dict(),
        "competencias": ped["competencia"].value_counts().sort_index().to_dict(),
        "de": str(ped["data"].min()), "ate": str(ped["data"].max()),
        "com_oc": int((ped["oc"] != "").sum()),
    }
    return ped, diag


def calcular(df: pd.DataFrame, pct: float, erp_idx: dict | None = None) -> list[dict]:
    out = []
    erp_idx = erp_idx or {}
    for r in df.itertuples(index=False):
        if r.cancelado:
            continue
        base = r.total_item or r.total        # soma dos itens = Valor Total do pedido
        sis_rs = round(base * pct, 2)
        real = round(r.comissao, 2)
        reb_com = round(sis_rs - real, 2)
        oc = r.oc or r.pedido                 # OC do ERP = "Entrega"
        e = erp_idx.get(oc) or erp_idx.get(r.pedido)
        out.append({
            "canal": "colombo",
            "pedido_mkt": oc, "pedido_canal": r.pedido, "pedido_any": (e or {}).get("obs05", "") or "",
            "id_mkt": "", "data": str(r.data), "competencia": r.competencia,
            "conta": "", "status": r.status, "pagamento": r.pagamento, "uf": r.uf, "cidade": r.cidade,
            "sku": r.sku, "anuncio": r.skus, "tipo": r.pcts, "produto": r.produto,
            "itens": int(r.itens), "qtd": float(r.qtd),
            "valor_prod": base, "valor_itens": r.merc, "tarifa": real, "frete": r.frete,
            "desconto": r.desconto, "cupom_seller": 0.0, "cupom_meli": 0.0,
            "pct_comissao": (real / base if base else 0.0),
            "sis_pct": pct, "sis_rs": sis_rs, "diferenca": reb_com,
            "data_nf": r.data_entrega, "parcelas": r.parcelas, "cliente": r.cliente,
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
        d["pedidos"] += 1; d["venda"] += l["valor_prod"]; d["comissao"] += arred.sis_exato(l, False) - (l["tarifa"] or 0.0); d["total"] += (l.get("rebate_rs") or 0.0) + (arred.sis_exato(l, False) - (l["tarifa"] or 0.0)) + (l.get("rebate_frete") or 0.0)
        d["tz"] += 1 if l["tarifa_zero"] else 0
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "frete", "total"):
            d[k] = round(d[k], 2)
    com_sis, com_real = arred.total(linhas, False), s("tarifa")  # arred. 1x no total
    reb_com = round(com_sis - com_real, 2)
    reb_tot = round(0.0 + reb_com + 0.0, 2)
    sts: dict[str, int] = {}
    for l in linhas:
        sts[l["status"] or "—"] = sts.get(l["status"] or "—", 0) + 1
    return {
        "pedidos": n, "cancelados": cancelados, "venda": venda, "tarifa": com_real, "frete": s("frete"),
        "cupom_meli": 0.0, "cupom_seller": 0.0, "faltante": 0.0, "valor_itens": s("valor_itens"), "desconto": s("desconto"),
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
        "status_validos": sts, "multi_item": sum(1 for l in linhas if l["itens"] > 1),
        "contas": {}, "tipos": sts,
        "por_dia": dict(sorted(por_dia.items())),
    }
