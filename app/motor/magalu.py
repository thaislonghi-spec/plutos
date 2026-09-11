"""Box MAGAZINE LUIZA — a regra da casa (método da Gabi, 11/09/2026).

Entrada: export do portal Magalu "FINANCEIRO POR PERÍODO" (.xlsx, 62 colunas,
1 linha por pedido "LU-…", com uma linha de TOTAL no fim que é ignorada).

- Pedido cancelado (status) fica FORA.
- COMISSÃO SISTEMA (o que deveria ser) = % cadastrado × valor pago pelo cliente
  + taxa fixa por pedido (Parâmetros · tabela de comissões: Magalu 11% + R$ 5,00).
- COMISSÃO REAL (cobrada) = "Serviços do marketplace (1+2+3+4)" + "Tarifa fixa"
  (vêm negativos; a 2ª forma de pagamento é ignorada — recebemos à vista).
- REBATE COMISSÃO = sistema − real.
- REBATE R$ = coparticipação de descontos pagos pelo Magalu (desconto à vista
  + preço promocional) + subsídio de cupom pago pelo Magalu.
- Coparticipação de fretes NÃO entra (vai para a conta logística).
"""
from __future__ import annotations

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
    t = str(v).strip()
    if not t or _norm(t) in ("nao se aplica", "em apuracao", "n/a", "-"):
        return 0.0
    t = t.replace("R$", "").replace(" ", "")
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
    for f in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


# nome normalizado (começo) → chave interna
COLS = {
    "numero do pedido": "pedido",
    "data/hora do pedido": "data",
    "status do pedido": "status",
    "canal de venda": "canal_venda",
    "modalidade de entrega": "modalidade",
    "cd de origem": "cd",
    "valor total do pedido pago pelo cliente": "pago",
    "valor total dos itens do pedido": "itens",
    "forma de pagamento 1": "forma_pgto",
    "servicos do marketplace (%) (forma de pagamento 1)": "pct_mkt",
    "servicos do marketplace (1+2+3+4) (forma de pagamento 1)": "servicos",
    "servicos de intermediacao (1) (forma de pagamento 1)": "intermediacao",
    "servicos de tecnologia (2) (forma de pagamento 1)": "tecnologia",
    "intermediacoes financeiras (mdr) (3) (forma de pagamento 1)": "mdr",
    "adm e gestao de recebiveis (4) (forma de pagamento 1)": "adm",
    "servicos do marketplace (1+2+3+4) (forma de pagamento 2)": "servicos_pgto2",
    "tarifa fixa": "tarifa_fixa",
    "coparticipacao de fretes estimada": "copart_frete",
    "custos logisticos": "custos_log",
    "descontos (*)": "descontos",
    "valor do repasse financeiro": "repasse",
    "coparticipacao de descontos (***) pago pelo magalu desconto": "desc_vista_magalu",
    "coparticipacao de descontos (***) pago por voce (seller) desconto": "desc_vista_seller",
    "coparticipacao de descontos (***) pago pelo magalu preco": "promo_magalu",
    "coparticipacao de descontos (***) pago por voce (seller) preco": "promo_seller",
    "valor subsidio cupom (r$) pago pelo magalu": "cupom_magalu",
    "valor subsidio cupom (r$) pago por voce (seller)": "cupom_seller",
    "valor liquido estimado a receber": "liquido",
}
OBRIGATORIAS = ["pedido", "data", "status", "pago", "servicos", "tarifa_fixa", "desc_vista_magalu", "cupom_magalu"]


def _mapear(colunas) -> dict:
    mapa = {}
    for c in colunas:
        n = _norm(c)
        for ini, k in COLS.items():
            if n.startswith(ini) and k not in mapa.values():
                mapa[c] = k
                break
    return mapa


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    """Lê o export FINANCEIRO POR PERÍODO. Devolve (df normalizado, diagnóstico)."""
    xl = pd.ExcelFile(caminho)
    escolhida, df = None, None
    for aba in xl.sheet_names:
        d = xl.parse(aba, dtype=str)
        mapa = _mapear(d.columns)
        if all(k in mapa.values() for k in OBRIGATORIAS):
            escolhida, df = aba, d.rename(columns=mapa)
            break
    if df is None:
        raise ValueError("Nenhuma aba tem as colunas do export 'FINANCEIRO POR PERÍODO' do Magalu "
                         "(Número do pedido, Data/Hora, Status, Valor pago, Serviços do marketplace, Tarifa fixa…).")
    df = df.copy()
    bruto = int(len(df))
    df["pedido"] = df["pedido"].fillna("").astype(str).str.strip()
    # a última linha é "Total dos pedidos capturados na pesquisa" — fora, senão dobra tudo
    df = df[df["pedido"].str.upper().str.startswith("LU-") | df["pedido"].str.match(r"^\d{6,}")]
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("status", "canal_venda", "modalidade", "cd", "forma_pgto"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    for k in ("pago", "itens", "pct_mkt", "servicos", "intermediacao", "tecnologia", "mdr", "adm", "servicos_pgto2",
              "tarifa_fixa", "copart_frete", "custos_log", "descontos", "repasse", "desc_vista_magalu",
              "desc_vista_seller", "promo_magalu", "promo_seller", "cupom_magalu", "cupom_seller", "liquido"):
        df[k] = df[k].map(_num)
    df["data"] = df["data"].map(_data)
    ruim = df["data"].isna()
    df = df[~ruim].copy()
    df["cancelado"] = df["status"].map(_norm).str.contains("cancelado")
    df["competencia"] = df["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    dup = df["pedido"].duplicated(keep="last")
    df = df[~dup].copy()
    diag: dict[str, Any] = {
        "aba": escolhida, "linhas_brutas": bruto, "linhas": int(len(df)), "sem_data": int(ruim.sum()),
        "duplicados": int(dup.sum()), "cancelados": int(df["cancelado"].sum()),
        "status": df["status"].value_counts().to_dict(),
        "modalidades": df["modalidade"].value_counts().to_dict(),
        "competencias": df["competencia"].value_counts().sort_index().to_dict(),
        "de": str(df["data"].min()), "ate": str(df["data"].max()),
        "com_pgto2": int((df["servicos_pgto2"] != 0).sum()),
        "pct_mkt": df["pct_mkt"].value_counts().head(6).to_dict(),
    }
    return df, diag


def base_oc(oc: str) -> str:
    """OC do ERP para o Magalu vem com sufixo -1/-2 (item/volume): 'LU-…-1' → 'LU-…'."""
    return re.sub(r"-\d+$", "", str(oc or "").strip())


def calcular(df: pd.DataFrame, pct: float, taxa: float, erp_por_base: dict | None = None) -> list[dict]:
    """Uma linha por pedido válido (não cancelado) com o rebate nas 3 formas.
    pct = comissão cadastrada (0.11), taxa = R$ por pedido (5.0)."""
    out = []
    erp_por_base = erp_por_base or {}
    for r in df.itertuples(index=False):
        if r.cancelado:
            continue
        sis_rs = round(r.pago * pct + taxa, 2)
        real = round(abs(r.servicos) + abs(r.tarifa_fixa), 2)
        reb_com = round(sis_rs - real, 2)
        reb_rs = round(r.desc_vista_magalu + r.promo_magalu + r.cupom_magalu, 2)
        e = erp_por_base.get(r.pedido)
        out.append({
            "canal": "magalu",
            "pedido_mkt": r.pedido, "pedido_canal": r.pedido, "pedido_any": (e or {}).get("obs05", "") or "",
            "id_mkt": "", "data": str(r.data), "competencia": r.competencia,
            "conta": r.cd, "status": r.status, "modalidade": r.modalidade, "forma_pgto": r.forma_pgto,
            "sku": "", "anuncio": "", "tipo": r.modalidade,
            "valor_prod": r.pago, "itens": r.itens, "tarifa": real, "frete": 0.0,
            "cupom_seller": r.cupom_seller, "cupom_meli": 0.0,
            "pct_comissao": (real / r.pago if r.pago else 0.0), "pct_mkt": r.pct_mkt / 100 if r.pct_mkt > 1 else r.pct_mkt,
            "servicos": r.servicos, "intermediacao": r.intermediacao, "tecnologia": r.tecnologia, "mdr": r.mdr,
            "tarifa_fixa": r.tarifa_fixa, "servicos_pgto2": r.servicos_pgto2,
            "sis_pct": pct, "sis_taxa": taxa, "sis_rs": sis_rs, "diferenca": reb_com,
            "desc_vista_magalu": r.desc_vista_magalu, "desc_vista_seller": r.desc_vista_seller,
            "promo_magalu": r.promo_magalu, "promo_seller": r.promo_seller,
            "cupom_magalu": r.cupom_magalu, "copart_frete": r.copart_frete, "custos_log": r.custos_log,
            "repasse": r.repasse, "liquido": r.liquido,
            "tarifa_zero": False, "faltante": 0.0, "faltante_status": "", "sem_sistema": False,
            "rebate_rs": reb_rs, "rebate_comissao": reb_com, "rebate_frete": 0.0,
            "rebate_total": round(reb_rs + reb_com, 2),
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
                                            "comissao": 0.0, "total": 0.0, "tz": 0})
        d["pedidos"] += 1; d["venda"] += l["valor_prod"]; d["cupom"] += l["rebate_rs"]
        d["comissao"] += l["rebate_comissao"]; d["total"] += l["rebate_total"]
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "total"):
            d[k] = round(d[k], 2)
    com_sis, com_real = s("sis_rs"), s("tarifa")
    mods: dict[str, int] = {}
    for l in linhas:
        mods[l["modalidade"] or "—"] = mods.get(l["modalidade"] or "—", 0) + 1
    return {
        "pedidos": n, "cancelados": cancelados, "venda": venda, "tarifa": com_real, "frete": 0.0,
        "cupom_meli": 0.0, "cupom_seller": s("cupom_seller"), "faltante": 0.0,
        "desc_vista_magalu": s("desc_vista_magalu"), "promo_magalu": s("promo_magalu"), "cupom_magalu": s("cupom_magalu"),
        "desc_vista_seller": s("desc_vista_seller"), "copart_frete": s("copart_frete"), "custos_log": s("custos_log"),
        "rebate_rs": s("rebate_rs"), "rebate_comissao": s("rebate_comissao"), "rebate_total": s("rebate_total"),
        "pct_sobre_venda": (round(100 * s("rebate_total") / venda, 2) if venda else 0.0),
        "tarifa_zero": 0, "faltante_pendentes": 0, "faltante_preenchidos": 0,
        "dif_pos": sum(1 for l in linhas if l["diferenca"] > 0.5), "dif_pos_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] > 0.5), 2),
        "dif_neg": sum(1 for l in linhas if l["diferenca"] < -0.5), "dif_neg_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] < -0.5), 2),
        "sem_sistema": 0, "erp_ok": sum(1 for l in linhas if l.get("erp_ok")),
        "com_pedidos": n, "com_sistema": com_sis, "com_real": com_real, "com_venda": venda,
        "com_sistema_pct": (round(100 * com_sis / venda, 2) if venda else 0.0),
        "com_real_pct": (round(100 * com_real / venda, 2) if venda else 0.0),
        "com_dif": round(com_sis - com_real, 2),
        "modalidades": mods, "contas": {}, "tipos": mods,
        "por_dia": dict(sorted(por_dia.items())),
    }
