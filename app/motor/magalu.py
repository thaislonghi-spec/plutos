"""Box MAGAZINE LUIZA — a regra da casa (método da Gabi, 11/09/2026).

Entrada: export do portal Magalu "FINANCEIRO POR PERÍODO" (.xlsx, 62 colunas,
1 linha por pedido "LU-…", com uma linha de TOTAL no fim que é ignorada).

- Pedido cancelado (status) fica FORA.
- COMISSÃO SISTEMA (o que deveria ser) = % do pedido no ERP (Promob) × base
  + R$ 5,00 de taxa por pedido (Parâmetros).
  Base: FULFILLMENT → produto + IPI (o Promob não cobra comissão sobre o frete
  nesses pedidos); demais → o TIPO cadastrado (GMV = valor pago pelo cliente).
  Sem par no ERP, cai para o % dos Parâmetros + a taxa fixa por pedido.
- COMISSÃO REAL (cobrada) = "Serviços do marketplace (1+2+3+4)" das DUAS formas
  de pagamento + "Tarifa fixa" (vêm negativos no relatório).
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


def fulfillment(modalidade) -> bool:
    """Modalidade de Entrega = 'Magalu entregas - Fulfillment' (regra da Thaís,
    16/09/2026: nesses pedidos o Promob cobra comissão SÓ sobre produto + IPI,
    não sobre o frete)."""
    return "fulfillment" in _norm(modalidade)


def calcular(df: pd.DataFrame, pct: float, taxa: float, erp_por_base: dict | None = None,
             tipo: str = "GMV") -> list[dict]:
    """Uma linha por pedido válido (não cancelado) com o rebate nas 3 formas.

    COMISSÃO DO SISTEMA (Promob), regra validada com a planilha da Gabi (16/09/2026):
      · % = o percentual DO PEDIDO no ERP (coluna de comissão da OC), não um % fixo;
        sem par no ERP, cai para o % dos Parâmetros.
      · + a taxa fixa por pedido dos Parâmetros (R$ 5,00): o % da OC é comissão
        pura, não embute a taxa (conferido pedido a pedido em set/26).
      · base = FULFILLMENT → produto + IPI (soma das OCs do ERP);
               demais → conforme o TIPO cadastrado (GMV = valor pago pelo cliente).
    A taxa fixa do canal (R$ ~5/pedido) NÃO entra na comissão do sistema — ela é
    custo cobrado pelo Magalu e já está na comissão real."""
    out = []
    erp_por_base = erp_por_base or {}
    for r in df.itertuples(index=False):
        if r.cancelado:
            continue
        e = erp_por_base.get(r.pedido)
        ff = fulfillment(r.modalidade)
        # A taxa fixa por pedido (Parâmetros, R$ 5,00) SEMPRE entra: conferido em
        # set/26 que o % gravado na OC é comissão pura — não embute a taxa
        # (0 de 1.301 pedidos batem com "serviços + taxa"; 1.012 batem sem ela).
        taxa_ped = taxa
        if e:
            pct_ped = float(e.get("pct_comissao") or 0.0) or pct
            base_prod = float(e.get("prod_erp") or 0.0) + float(e.get("ipi_erp") or 0.0)
        else:
            pct_ped = pct
            base_prod = r.itens or r.pago
        base = base_prod if (ff or tipo == "Produto") else r.pago
        base_nome = "produto + IPI" if (ff or tipo == "Produto") else "GMV (valor pago pelo cliente)"
        sis_rs = round(base * pct_ped + taxa_ped, 2)
        # Quando o cliente divide o pagamento em duas formas, o Magalu cobra
        # serviços nas DUAS (conferido com a planilha da Gabi, 16/09/2026: 9
        # pedidos de set/26). A 2ª forma entra na comissão real.
        real = round(abs(r.servicos) + abs(r.servicos_pgto2) + abs(r.tarifa_fixa), 2)
        reb_com = round(sis_rs - real, 2)
        reb_rs = round(r.desc_vista_magalu + r.promo_magalu + r.cupom_magalu, 2)
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
            "sis_pct": pct_ped, "sis_taxa": taxa_ped, "sis_rs": sis_rs, "diferenca": reb_com,
            "sis_base": round(base, 2), "sis_base_nome": base_nome, "fulfillment": ff,
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
        d["comissao"] += arred.sis_exato(l, True) - (l["tarifa"] or 0.0); d["total"] += (l.get("rebate_rs") or 0.0) + (arred.sis_exato(l, True) - (l["tarifa"] or 0.0)) + (l.get("rebate_frete") or 0.0)
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "total"):
            d[k] = round(d[k], 2)
    com_sis, com_real = arred.total(linhas, True), s("tarifa")  # arred. 1x no total
    reb_com = round(com_sis - com_real, 2)
    reb_tot = round(s("rebate_rs") + reb_com + 0.0, 2)
    mods: dict[str, int] = {}
    for l in linhas:
        mods[l["modalidade"] or "—"] = mods.get(l["modalidade"] or "—", 0) + 1
    # FULFILLMENT × entrega própria (a aba do Fulfillment vive desses números)
    def _bloco(sel):
        g = [l for l in linhas if bool(l.get("fulfillment")) is sel]
        sis = arred.total(g, True)
        real = round(sum((l.get("tarifa") or 0.0) for l in g), 2)
        rs = round(sum((l.get("rebate_rs") or 0.0) for l in g), 2)
        venda_g = round(sum((l.get("valor_prod") or 0.0) for l in g), 2)
        base_g = round(sum((l.get("sis_base") or 0.0) for l in g), 2)
        return {"pedidos": len(g), "venda": venda_g, "base": base_g, "com_sistema": sis, "com_real": real,
                "rebate_comissao": round(sis - real, 2), "rebate_rs": rs,
                "rebate_total": round(sis - real + rs, 2),
                "com_real_pct": (round(100 * real / venda_g, 2) if venda_g else 0.0),
                "com_sistema_pct": (round(100 * sis / venda_g, 2) if venda_g else 0.0)}
    ff, propria = _bloco(True), _bloco(False)
    return {
        "pedidos": n, "cancelados": cancelados, "venda": venda, "tarifa": com_real, "frete": 0.0,
        "cupom_meli": 0.0, "cupom_seller": s("cupom_seller"), "faltante": 0.0,
        "desc_vista_magalu": s("desc_vista_magalu"), "promo_magalu": s("promo_magalu"), "cupom_magalu": s("cupom_magalu"),
        "desc_vista_seller": s("desc_vista_seller"), "copart_frete": s("copart_frete"), "custos_log": s("custos_log"),
        "rebate_rs": s("rebate_rs"), "rebate_comissao": reb_com, "rebate_total": reb_tot,
        "pct_sobre_venda": (round(100 * reb_tot / venda, 2) if venda else 0.0),
        "tarifa_zero": 0, "faltante_pendentes": 0, "faltante_preenchidos": 0,
        "dif_pos": sum(1 for l in linhas if l["diferenca"] > 0.5), "dif_pos_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] > 0.5), 2),
        "dif_neg": sum(1 for l in linhas if l["diferenca"] < -0.5), "dif_neg_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] < -0.5), 2),
        "sem_sistema": 0, "erp_ok": sum(1 for l in linhas if l.get("erp_ok")),
        "com_pedidos": n, "com_sistema": com_sis, "com_real": com_real, "com_venda": venda,
        "com_sistema_pct": (round(100 * com_sis / venda, 2) if venda else 0.0),
        "com_real_pct": (round(100 * com_real / venda, 2) if venda else 0.0),
        "com_dif": round(com_sis - com_real, 2),
        "modalidades": mods, "contas": {}, "tipos": mods,
        "ff": ff, "propria": propria, "sem_erp": sum(1 for l in linhas if not l.get("erp_ok")),
        "por_dia": dict(sorted(por_dia.items())),
    }
