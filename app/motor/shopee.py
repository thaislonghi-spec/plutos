"""Box SHOPEE — método validado contra a apuração da Gabi (doc "PLUTOS · Box
Shopee — Método de cálculo", v1 12/09/2026).

Entrada: export do portal Shopee "Order.all.order_creation_date.AAAAMMDD_AAAAMMDD.xlsx"
(66 colunas, UMA LINHA POR ITEM do pedido).

- Status "Cancelado" fica fora (taxas zeradas, não gera rebate).
- FORA: Status "Cancelado", Status "Não pago" (entra quando for pago) e
  Status da Devolução / Reembolso = "Solicitação aprovada" (a venda voltou).
- Uma linha por PEDIDO: taxas, frete, incentivo e cupom se repetem em cada item
  → vale a primeira linha; Subtotal do produto e Quantidade são SOMADOS.
- COMISSÃO SISTEMA = % × subtotal + R$/item × quantidade (Parâmetros: Shopee 12% + R$ 12,00).
- COMISSÃO REAL = Taxa de comissão bruta + Taxa de serviço bruta − Ajuste por participação em ação comercial.
- REBATE COMISSÃO = sistema − real.
- COMISSÃO SISTEMA = % da tabela (12%) × Subtotal do produto (= valor produto + IPI).
- COMISSÃO REAL = comissão bruta + serviço bruta − ajuste (a taxa fixa de R$ 12,00/un já vem
  dentro da "Taxa de serviço bruta" e fica na conta: reduz o rebate).
- REBATE R$ = Incentivo Shopee para ação comercial + Incentivo de cupom (NÃO a coluna "Cupom").
- REBATE FRETE = Valor estimado do frete − Taxa de envio pagas pelo comprador (programa Frete Grátis, teto R$ 40).
- Chave com o ERP: ID do pedido = OC.
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
    for f in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


COLS = {
    "id do pedido": "pedido", "status do pedido": "status", "opcao de envio": "envio",
    "status da devolucao / reembolso": "devolucao", "status da devolucao/reembolso": "devolucao",
    "data de criacao do pedido": "data", "numero de rastreamento": "rastreio",
    "n de referencia do sku principal": "sku", "nome do produto": "produto", "quantidade": "qtd",
    "preco original": "preco_original", "preco acordado": "preco", "subtotal do produto": "subtotal",
    "desconto do vendedor": "desc_vendedor",
    "incentivo shopee para acao comercial": "incentivo", "ajuste por participacao em acao comercial": "ajuste",
    "cupom do vendedor": "cupom_vendedor", "cupom": "cupom_total", "incentivo de cupom": "cupom_shopee",
    "compensar moedas shopee": "moedas", "valor total": "valor_total",
    "taxa de envio pagas pelo comprador": "frete_comprador", "taxa de transacao": "taxa_transacao",
    "taxa de comissao bruta": "comissao_bruta", "taxa de comissao liquida": "comissao_liquida",
    "taxa de servico bruta": "servico_bruta", "taxa de servico liquida": "servico_liquida",
    "total global": "total_global", "valor estimado do frete": "frete_estimado", "uf": "uf",
    "peso total do pedido": "peso",
}
OBRIGATORIAS = ["pedido", "status", "data", "subtotal", "qtd", "comissao_bruta", "servico_bruta", "ajuste",
                "incentivo", "cupom_shopee", "frete_estimado", "frete_comprador"]


def _mapear(colunas) -> dict:
    mapa = {}
    for c in colunas:
        n = _norm(c).replace("º", "").replace("°", "")
        n = re.sub(r"\.\d+$", "", n)  # "Desconto do vendedor.1" (coluna repetida no pandas)
        if n in COLS and COLS[n] not in mapa.values():
            mapa[c] = COLS[n]
    return mapa


def fora_da_conta(status, devolucao="") -> bool:
    """FORA DA CONTA (regra oficial da Thaís, 14/09/2026), três casos:
       1. Status do pedido = Cancelado
       2. Status do pedido = Não pago — quando for pago, o arquivo seguinte atualiza
          o status e o pedido entra sozinho no rebate
       3. Status da Devolução / Reembolso = Solicitação aprovada (a venda voltou)
    Vale na leitura E no recálculo: assim a regra alcança o que já está na base,
    sem precisar subir o arquivo de novo."""
    st = _norm(status)
    return bool("cancelado" in st or st.startswith("nao pago")
                or "solicitacao aprovada" in _norm(devolucao))


def marcar_fora(ped: pd.DataFrame) -> pd.DataFrame:
    """Recalcula as marcas de exclusão a partir do status (e da devolução, quando o
    arquivo tem a coluna). Aplicado também na base já gravada."""
    if "devolucao" not in ped:
        ped["devolucao"] = ""
    st = ped["status"].map(_norm)
    dev = ped["devolucao"].fillna("").map(_norm)
    ped["devolvido"] = dev.str.contains("solicitacao aprovada")
    ped["nao_pago"] = st.str.startswith("nao pago")
    ped["cancelado"] = st.str.contains("cancelado") | ped["nao_pago"] | ped["devolvido"]
    return ped


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    """Lê o export Order.all (1 linha por item) e devolve 1 linha por PEDIDO."""
    xl = pd.ExcelFile(caminho)
    escolhida, df = None, None
    for aba in xl.sheet_names:
        d = xl.parse(aba, dtype=str)
        mapa = _mapear(d.columns)
        if all(k in mapa.values() for k in OBRIGATORIAS):
            escolhida, df = aba, d.rename(columns=mapa)
            break
    if df is None:
        raise ValueError("Nenhuma aba tem as colunas do export 'Order.all' da Shopee "
                         "(ID do pedido, Status, Data de criação, Subtotal do produto, Taxa de comissão bruta…).")
    df = df.copy()
    itens = int(len(df))
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("pedido", "status", "envio", "sku", "produto", "uf", "rastreio", "devolucao"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    for k in ("qtd", "preco_original", "preco", "subtotal", "desc_vendedor", "incentivo", "ajuste", "cupom_vendedor",
              "cupom_total", "cupom_shopee", "moedas", "valor_total", "frete_comprador", "taxa_transacao",
              "comissao_bruta", "comissao_liquida", "servico_bruta", "servico_liquida", "total_global",
              "frete_estimado", "peso"):
        df[k] = df[k].map(_num)
    df["data"] = df["data"].map(_data)
    df = df[(df["pedido"] != "") & df["data"].notna()].copy()
    # 1 linha por pedido: soma subtotal/qtd; o resto vale a primeira linha
    primeira = df.groupby("pedido", sort=False).first()
    somas = df.groupby("pedido", sort=False).agg(subtotal=("subtotal", "sum"), qtd=("qtd", "sum"), itens=("sku", "size"),
                                                  skus=("sku", lambda s: ", ".join(dict.fromkeys(x for x in s if x))))
    ped = primeira.drop(columns=["subtotal", "qtd"]).join(somas).reset_index()
    ped = marcar_fora(ped)
    ped["competencia"] = ped["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    diag: dict[str, Any] = {
        "aba": escolhida, "linhas_brutas": itens, "itens": itens, "linhas": int(len(ped)),
        "cancelados": int(ped["cancelado"].sum()), "multi_item": int((ped["itens"] > 1).sum()),
        "nao_pagos": int(ped["nao_pago"].sum()), "devolvidos": int(ped["devolvido"].sum()),
        "status": ped["status"].value_counts().to_dict(), "envio": ped["envio"].value_counts().to_dict(),
        "competencias": ped["competencia"].value_counts().sort_index().to_dict(),
        "de": str(ped["data"].min()), "ate": str(ped["data"].max()),
    }
    return ped, diag


def calcular(df: pd.DataFrame, pct: float, taxa_item: float, erp_idx: dict | None = None) -> list[dict]:
    out = []
    erp_idx = erp_idx or {}
    for r in df.itertuples(index=False):
        if fora_da_conta(getattr(r, "status", ""), getattr(r, "devolucao", "")):
            continue
        # REGRA OFICIAL DA THAÍS (14/09/2026):
        #   comissão REAL Shopee = Taxa de comissão bruta (AY) + Taxa de serviço bruta (BA)
        #                          − Ajuste por participação em ação comercial (AE)
        #   comissão SISTEMA (Promob) = % da tabela × Subtotal do produto (= produto + IPI)
        #   REBATE EM COMISSÃO = sistema − real
        # A taxa fixa de R$ 12,00/un JÁ ESTÁ dentro da Taxa de serviço bruta (conferido:
        # tirando 12 × quantidade, o serviço vira exatamente 2,00% do subtotal) e NÃO é
        # somada à comissão do sistema — ela reduz o rebate, como custo que o Promob não
        # previu. O campo taxa_fixa guarda o valor só para consulta.
        fixo = round(taxa_item * (r.qtd or 0), 2)
        sis_rs = round(r.subtotal * pct, 2)
        real = round(r.comissao_bruta + r.servico_bruta - r.ajuste, 2)
        real_cheia = real
        reb_com = round(sis_rs - real, 2)
        reb_rs = round(r.incentivo + r.cupom_shopee, 2)
        reb_frete = round(max(0.0, r.frete_estimado - r.frete_comprador), 2)
        e = erp_idx.get(r.pedido)
        out.append({
            "canal": "shopee",
            "pedido_mkt": r.pedido, "pedido_canal": r.pedido, "pedido_any": (e or {}).get("obs05", "") or "",
            "id_mkt": r.rastreio, "data": str(r.data), "competencia": r.competencia,
            "conta": "", "status": r.status, "envio": r.envio, "uf": r.uf,
            "sku": r.sku, "anuncio": r.skus, "tipo": r.envio, "produto": r.produto, "itens": int(r.itens), "qtd": float(r.qtd),
            "valor_prod": r.subtotal, "tarifa": real, "frete": r.frete_estimado, "frete_comprador": r.frete_comprador,
            "cupom_seller": r.cupom_vendedor, "cupom_meli": 0.0, "desc_vendedor": r.desc_vendedor,
            "pct_comissao": (real / r.subtotal if r.subtotal else 0.0),
            "pct_comissao_cheia": (real_cheia / r.subtotal if r.subtotal else 0.0),
            "comissao_bruta": r.comissao_bruta, "servico_bruta": r.servico_bruta, "ajuste": r.ajuste,
            "comissao_liquida": r.comissao_liquida, "servico_liquida": r.servico_liquida, "taxa_transacao": r.taxa_transacao,
            "incentivo": r.incentivo, "cupom_shopee": r.cupom_shopee, "moedas": r.moedas, "total_global": r.total_global,
            "sis_pct": pct, "sis_taxa": taxa_item, "sis_rs": sis_rs, "diferenca": reb_com,
            "taxa_fixa": fixo, "tarifa_cheia": real_cheia,
            "tarifa_zero": bool(r.ajuste >= r.comissao_bruta - 0.005 and r.comissao_bruta > 0), "faltante": 0.0, "faltante_status": "", "sem_sistema": False,
            "rebate_rs": reb_rs, "rebate_comissao": reb_com, "rebate_frete": reb_frete,
            "rebate_total": round(reb_rs + reb_com + reb_frete, 2),
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
        d["pedidos"] += 1; d["venda"] += l["valor_prod"]; d["cupom"] += l["rebate_rs"]
        d["comissao"] += l["rebate_comissao"]; d["frete"] += l["rebate_frete"]; d["total"] += l["rebate_total"]
        d["tz"] += 1 if l["tarifa_zero"] else 0
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "frete", "total"):
            d[k] = round(d[k], 2)
    com_sis, com_real = s("sis_rs"), s("tarifa")
    envios: dict[str, int] = {}
    for l in linhas:
        envios[l["envio"] or "—"] = envios.get(l["envio"] or "—", 0) + 1
    fretes = {"40": sum(1 for l in linhas if abs(l["rebate_frete"] - 40) < 0.02),
              "20": sum(1 for l in linhas if abs(l["rebate_frete"] - 20) < 0.02),
              "0": sum(1 for l in linhas if l["rebate_frete"] < 0.01)}
    return {
        "pedidos": n, "cancelados": cancelados, "venda": venda, "tarifa": com_real, "frete": s("frete"),
        "cupom_meli": 0.0, "cupom_seller": s("cupom_seller"), "faltante": 0.0,
        "incentivo": s("incentivo"), "cupom_shopee": s("cupom_shopee"), "frete_comprador": s("frete_comprador"),
        "comissao_bruta": s("comissao_bruta"), "servico_bruta": s("servico_bruta"), "ajuste": s("ajuste"),
        "taxa_fixa": s("taxa_fixa"), "tarifa_cheia": s("tarifa_cheia"),
        "rebate_rs": s("rebate_rs"), "rebate_comissao": s("rebate_comissao"), "rebate_frete": s("rebate_frete"),
        "rebate_total": s("rebate_total"),
        "pct_sobre_venda": (round(100 * s("rebate_total") / venda, 2) if venda else 0.0),
        "tarifa_zero": sum(1 for l in linhas if l["tarifa_zero"]), "faltante_pendentes": 0, "faltante_preenchidos": 0,
        "dif_pos": sum(1 for l in linhas if l["diferenca"] > 0.5), "dif_pos_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] > 0.5), 2),
        "dif_neg": sum(1 for l in linhas if l["diferenca"] < -0.5), "dif_neg_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] < -0.5), 2),
        "sem_sistema": 0, "erp_ok": sum(1 for l in linhas if l.get("erp_ok")),
        "com_pedidos": n, "com_sistema": com_sis, "com_real": com_real, "com_venda": venda,
        "com_sistema_pct": (round(100 * com_sis / venda, 2) if venda else 0.0),
        "com_real_pct": (round(100 * com_real / venda, 2) if venda else 0.0),
        "com_dif": round(com_sis - com_real, 2),
        "envios": envios, "fretes": fretes, "multi_item": sum(1 for l in linhas if l["itens"] > 1),
        "contas": {}, "tipos": envios,
        "por_dia": dict(sorted(por_dia.items())),
    }
