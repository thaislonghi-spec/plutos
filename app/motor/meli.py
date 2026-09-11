# -*- coding: utf-8 -*-
"""
PLUTOS · motor do box MERCADO LIVRE
====================================
Reproduz, sem mão humana, o que a Gabi faz na planilha "MERCADO LIVRE.xlsx"
a partir da TABELA GERAL DE PEDIDOS (export do Power BI, filtro Pago, 01–31):

  R  % COMISSÃO        = Tarifa Venda / Valor Produtos
  S  FALTANTE CAMPANHA = pedidos com tarifa ZERO → tabela manual (portal do Meli)
  T  % PROMOB          = comissão do SISTEMA (Promob) em %   ← planilha "rebate" do ADC002
  U  PROMOB R$         = comissão do SISTEMA (Promob) em R$  ← idem
  V  DIFERENÇA COM     = U − Tarifa Venda  (comissão sistema − comissão cobrada)

Rebates (definição da Thaís, 09/09/2026 — três formas):
  R$        = Cupom Meli  +  Faltante campanha (manual)
  COMISSÃO  = Diferença de comissão (só quando > tolerância)
  FRETE     = não existe no Meli (frete do Coletas é custo nosso, já no custo do produto)
"""
from __future__ import annotations

import re
from datetime import datetime, date
from typing import Any

import pandas as pd

CANAL = "MERCADO LIVRE"
CHAVE = "meli"

# colunas mínimas da TABELA GERAL DE PEDIDOS (nomes como saem do Power BI)
COLS = {
    "pedido_canal": "Pedido Canal",
    "pedido_mkt": "Pedido Marketplace (Any)",
    "id_mkt": "ID Marketplace (Any)",
    "pedido_any": "Pedido Any",
    "data": "Data Pedido",
    "conta": "Conta",
    "status": "Status Canal",
    "sku": "SKU",
    "anuncio": "Anúncio",
    "tipo": "Tipo Anúncio",
    "valor_prod": "Valor Produtos Meli (+)",
    "tarifa": "Tarifa Venda (-)",
    "frete": "Frete Pedido",
    "cupom_seller": "Cupom Seller",
    "cupom_meli": "Cupom Meli",
    "valor_meli": "Valor Pedido Meli",
    "valor_any": "Valor Pedido Any",
}
OBRIGATORIAS = ["pedido_mkt", "data", "sku", "valor_prod", "tarifa", "cupom_meli"]


def _norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = re.sub(r"[áàâã]", "a", s); s = re.sub(r"[éê]", "e", s); s = re.sub(r"[íî]", "i", s)
    s = re.sub(r"[óôõ]", "o", s); s = re.sub(r"[úû]", "u", s); s = s.replace("ç", "c")
    return re.sub(r"\s+", " ", s)


def _num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if not s or s in ("-", "—"):
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _pedido(v) -> str:
    """número de pedido como texto limpo (evita 2.0000153e15 do Excel)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def _data(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()[:10]
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


# --------------------------------------------------------------------------
# 1. TABELA GERAL DE PEDIDOS
# --------------------------------------------------------------------------
def ler_tabela_geral(caminho: str) -> tuple[pd.DataFrame, dict]:
    """Lê o export do Power BI. Aceita a aba 'Export' ou a primeira aba que
    tenha as colunas obrigatórias. Devolve (df normalizado, diagnóstico)."""
    xl = pd.ExcelFile(caminho)
    escolhida, df = None, None
    for aba in xl.sheet_names:
        d = xl.parse(aba, dtype={"Pedido Canal": str, "Pedido Marketplace (Any)": str,
                                 "ID Marketplace (Any)": str, "Pedido Any": str})
        mapa = _mapear(d.columns)
        if all(k in mapa.values() for k in OBRIGATORIAS):
            escolhida, df = aba, d.rename(columns=mapa)
            break
    if df is None:
        faltam = [COLS[k] for k in OBRIGATORIAS]
        raise ValueError("Nenhuma aba tem as colunas da TABELA GERAL DE PEDIDOS. "
                         f"Preciso pelo menos de: {', '.join(faltam)}.")

    diag: dict[str, Any] = {"aba": escolhida, "linhas_brutas": int(len(df)), "rejeitadas": []}

    df = df.copy()
    for k in ("pedido_canal", "pedido_mkt", "id_mkt", "pedido_any"):
        if k in df:
            df[k] = df[k].map(_pedido)
        else:
            df[k] = ""
    for k in ("valor_prod", "tarifa", "frete", "cupom_seller", "cupom_meli", "valor_meli", "valor_any"):
        df[k] = df[k].map(_num) if k in df else 0.0
    for k in ("conta", "status", "sku", "anuncio", "tipo"):
        df[k] = df[k].fillna("").astype(str).str.strip() if k in df else ""
    df["data"] = df["data"].map(_data)

    # linhas rejeitadas: sem pedido, sem data, ou rodapé de total
    ruim = df["pedido_mkt"].eq("") | df["data"].isna()
    for _, r in df[ruim].head(20).iterrows():
        diag["rejeitadas"].append({"pedido": r.get("pedido_canal") or r.get("pedido_mkt") or "—",
                                   "motivo": "sem pedido ou sem data"})
    diag["n_rejeitadas"] = int(ruim.sum())
    df = df[~ruim].copy()

    # status: o método da Gabi filtra PAGO no BI. Se vier outro, avisa e mantém
    # (o filtro vem do BI; aqui a decisão é dela, não do app).
    st = df["status"].map(_norm)
    diag["status"] = df["status"].value_counts().to_dict()
    diag["nao_pago"] = int((~st.isin(["pago", ""])).sum())

    # duplicidade de pedido (mesmo pedido em 2 linhas = 2 itens? no BI vem 1 linha por pedido)
    dup = df["pedido_mkt"].duplicated(keep=False)
    diag["duplicados"] = int(dup.sum())
    df["duplicado"] = dup

    df["competencia"] = df["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    diag["competencias"] = df["competencia"].value_counts().sort_index().to_dict()
    diag["de"] = str(df["data"].min())
    diag["ate"] = str(df["data"].max())
    diag["linhas"] = int(len(df))
    return df, diag


def _mapear(colunas) -> dict:
    inv = {_norm(v): k for k, v in COLS.items()}
    mapa = {}
    for c in colunas:
        n = _norm(c)
        if n in inv:
            mapa[c] = inv[n]
        else:
            # tolerância: "pedido marketplace" sem "(any)", "valor produtos" etc.
            for k, v in COLS.items():
                if _norm(v).replace(" (any)", "") == n.replace(" (any)", "") and k not in mapa.values():
                    mapa[c] = k
                    break
    return mapa


# --------------------------------------------------------------------------
# 2. COMISSÃO DO SISTEMA (planilha "rebate" do ADC002)
#    O PROCV da Gabi: chave na coluna D, % na coluna AB (25ª a partir de D),
#    R$ na coluna AK (34ª a partir de D). Primeiro tenta pelos NOMES das
#    colunas; se não achar, cai nas POSIÇÕES do PROCV.
# --------------------------------------------------------------------------
def ler_comissao_sistema(caminho: str) -> tuple[dict, dict]:
    xl = pd.ExcelFile(caminho)
    melhor, diag = {}, {"aba": None, "modo": None, "linhas": 0}
    for aba in xl.sheet_names:
        d = xl.parse(aba, header=None, dtype=str)
        if d.shape[1] < 5:
            continue
        # acha a linha de cabeçalho (a que tem "pedido" em alguma célula)
        hdr = None
        for i in range(min(15, len(d))):
            linha = [_norm(x) for x in d.iloc[i].tolist()]
            if any("pedido" in x for x in linha):
                hdr = i
                break
        if hdr is None:
            continue
        cab = [_norm(x) for x in d.iloc[hdr].tolist()]
        corpo = d.iloc[hdr + 1:]

        ic = ipct = irs = None
        for j, c in enumerate(cab):
            if ic is None and "pedido" in c and ("marketplace" in c or "any" in c or "canal" in c):
                ic = j
            if ipct is None and "comiss" in c and "%" in c:
                ipct = j
            if irs is None and "comiss" in c and "%" not in c and ("r$" in c or "valor" in c or "promob" in c):
                irs = j
        modo = "nomes"
        if ic is None or irs is None:
            # posições do PROCV: D = 3, AB = 27, AK = 36 (0-based)
            if d.shape[1] > 36:
                ic, ipct, irs, modo = 3, 27, 36, "posicoes do PROCV (D → AB / AK)"
            else:
                continue
        tab = {}
        for _, r in corpo.iterrows():
            k = _pedido(r.iloc[ic])
            if not k:
                continue
            tab[k] = {"pct": _num(r.iloc[ipct]) if ipct is not None else None,
                      "rs": _num(r.iloc[irs])}
        if len(tab) > len(melhor):
            melhor, diag = tab, {"aba": aba, "modo": modo, "linhas": len(tab)}
    if not melhor:
        raise ValueError("Não achei uma tabela com Pedido + Comissão do sistema nessa planilha.")
    return melhor, diag


# --------------------------------------------------------------------------
# 3. O CÁLCULO — uma linha por pedido
# --------------------------------------------------------------------------
def calcular(df: pd.DataFrame, comissao_sistema: dict | None, faltante: dict,
             tolerancia: float = 0.50) -> list[dict]:
    """faltante = {pedido_mkt: {"valor": float, ...}} — a tabela manual da Gabi."""
    linhas = []
    tem_sis = bool(comissao_sistema)
    for r in df.itertuples(index=False):
        vp, tarifa = r.valor_prod, r.tarifa
        pct = (tarifa / vp) if vp else 0.0
        sis = (comissao_sistema or {}).get(r.pedido_mkt) or (comissao_sistema or {}).get(r.pedido_canal)
        if sis:
            if sis.get("rs") is None and sis.get("pct") is not None:
                # ERP: só o % (coluna AB); o R$ é sobre o VALOR DE PRODUTOS do export do canal
                sis_rs = round(vp * sis["pct"], 2)
            else:
                sis_rs = sis["rs"]
            sis_pct = sis["pct"] if sis.get("pct") is not None else ((sis_rs / vp) if vp else None)
            dif = sis_rs - tarifa
        else:
            sis_rs = sis_pct = dif = None
        tarifa_zero = (tarifa == 0 and vp > 0)
        f = faltante.get(r.pedido_mkt) or {}
        falt_val = _num(f.get("valor")) if f else 0.0

        reb_rs = r.cupom_meli + (falt_val if tarifa_zero else 0.0)
        reb_com = dif if (dif is not None and abs(dif) > tolerancia) else 0.0
        linhas.append({
            "canal": CANAL,
            "pedido_canal": r.pedido_canal, "pedido_mkt": r.pedido_mkt, "id_mkt": r.id_mkt,
            "pedido_any": r.pedido_any,
            "data": r.data.isoformat(), "competencia": r.competencia,
            "conta": r.conta, "status": r.status, "sku": r.sku, "anuncio": r.anuncio, "tipo": r.tipo,
            "valor_prod": round(vp, 2), "tarifa": round(tarifa, 2), "frete": round(r.frete, 2),
            "cupom_seller": round(r.cupom_seller, 2), "cupom_meli": round(r.cupom_meli, 2),
            "pct_comissao": round(pct, 4),
            "sis_pct": (round(sis_pct, 4) if sis_pct is not None else None),
            "sis_rs": (round(sis_rs, 2) if sis_rs is not None else None),
            "diferenca": (round(dif, 2) if dif is not None else None),
            "tarifa_zero": tarifa_zero,
            "faltante": round(falt_val, 2) if tarifa_zero else 0.0,
            "faltante_status": ("preenchido" if (tarifa_zero and f) else ("pendente" if tarifa_zero else "")),
            "rebate_rs": round(reb_rs, 2),
            "rebate_comissao": round(reb_com, 2),
            "rebate_frete": 0.0,
            "rebate_total": round(reb_rs + reb_com, 2),
            "sem_sistema": (not sis) if tem_sis else True,
            "duplicado": bool(r.duplicado),
        })
    return linhas


def resumo(linhas: list[dict]) -> dict:
    n = len(linhas)
    s = lambda k: round(sum((l[k] or 0.0) for l in linhas), 2)  # noqa: E731
    tz = [l for l in linhas if l["tarifa_zero"]]
    pend = [l for l in tz if l["faltante_status"] == "pendente"]
    dif_pos = [l for l in linhas if (l["diferenca"] or 0) > 0.5]
    dif_neg = [l for l in linhas if (l["diferenca"] or 0) < -0.5]
    por_dia: dict[str, dict] = {}
    for l in linhas:
        d = por_dia.setdefault(l["data"], {"pedidos": 0, "venda": 0.0, "cupom": 0.0, "faltante": 0.0,
                                            "comissao": 0.0, "total": 0.0, "tz": 0})
        d["pedidos"] += 1; d["venda"] += l["valor_prod"]; d["cupom"] += l["cupom_meli"]
        d["faltante"] += l["faltante"]; d["comissao"] += l["rebate_comissao"]; d["total"] += l["rebate_total"]
        d["tz"] += 1 if l["tarifa_zero"] else 0
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "total"):
            d[k] = round(d[k], 2)
    venda = s("valor_prod")
    com = [l for l in linhas if l.get("sis_rs") is not None]  # pedidos com par no ERP
    com_sis = round(sum(l["sis_rs"] for l in com), 2)
    com_real = round(sum((l["tarifa"] or 0.0) for l in com), 2)
    venda_com = round(sum(l["valor_prod"] for l in com), 2)
    return {
        "com_pedidos": len(com), "com_sistema": com_sis, "com_real": com_real, "com_venda": venda_com,
        "com_sistema_pct": (round(100 * com_sis / venda_com, 2) if venda_com else 0.0),
        "com_real_pct": (round(100 * com_real / venda_com, 2) if venda_com else 0.0),
        "com_dif": round(com_sis - com_real, 2),
        "pedidos": n, "venda": venda, "tarifa": s("tarifa"), "frete": s("frete"),
        "cupom_meli": s("cupom_meli"), "cupom_seller": s("cupom_seller"),
        "faltante": s("faltante"), "rebate_rs": s("rebate_rs"), "rebate_comissao": s("rebate_comissao"),
        "rebate_total": s("rebate_total"),
        "pct_sobre_venda": (round(100 * s("rebate_total") / venda, 2) if venda else 0.0),
        "tarifa_zero": len(tz), "faltante_pendentes": len(pend),
        "faltante_preenchidos": len(tz) - len(pend),
        "dif_pos": len(dif_pos), "dif_pos_rs": round(sum(l["diferenca"] for l in dif_pos), 2),
        "dif_neg": len(dif_neg), "dif_neg_rs": round(sum(l["diferenca"] for l in dif_neg), 2),
        "sem_sistema": sum(1 for l in linhas if l["sem_sistema"]),
        "contas": _conta(linhas, "conta"), "tipos": _conta(linhas, "tipo"),
        "por_dia": dict(sorted(por_dia.items())),
    }


def _conta(linhas, k):
    out: dict[str, int] = {}
    for l in linhas:
        out[l[k] or "—"] = out.get(l[k] or "—", 0) + 1
    return out


def sugestoes_faltante(linhas: list[dict], faltante: dict) -> list[dict]:
    """A tabela FALTANTE CAMPANHA: todos os pedidos de tarifa zero, com a sugestão
    de valor quando um pedido IGUAL já foi preenchido — mesmo SKU, mesmo preço de
    produto, na mesma data ou na mais próxima (regra da Thaís, 10/09/2026).
    Segunda chance: mesmo anúncio (MLB). A sugestão nunca grava sozinha: só com o OK."""
    from datetime import date as _d
    feitos = []
    for l in linhas:
        if not l["tarifa_zero"]:
            continue
        f = faltante.get(l["pedido_mkt"])
        if f and f.get("valor") not in (None, ""):
            feitos.append((l["sku"], round(l["valor_prod"], 2), l["anuncio"], _d.fromisoformat(l["data"]),
                           _num(f["valor"]), l["pedido_mkt"]))
    out = []
    for l in linhas:
        if not l["tarifa_zero"]:
            continue
        f = faltante.get(l["pedido_mkt"]) or {}
        sug, origem = None, ""
        if not f and feitos:
            d0 = _d.fromisoformat(l["data"])
            cands = [x for x in feitos if x[0] == l["sku"] and abs(x[1] - round(l["valor_prod"], 2)) < 0.01]
            modo = "mesmo SKU e mesmo preço"
            if not cands:
                cands = [x for x in feitos if x[2] == l["anuncio"]]
                modo = "mesmo anúncio"
            if cands:
                cands.sort(key=lambda x: abs((x[3] - d0).days))
                melhor = cands[0]
                dias = abs((melhor[3] - d0).days)
                sug = melhor[4]
                origem = f"{modo} · {'mesma data' if dias == 0 else f'{dias} dia(s) de distância'} · pedido {melhor[5]}"
                vals = sorted(set(round(x[4], 2) for x in cands))
                if len(vals) > 1:
                    origem += f" · valores diferentes já usados: {', '.join(f'{v:.2f}' for v in vals)}"
        out.append({**l, "manual": f, "sugestao": sug, "sugestao_origem": origem})
    out.sort(key=lambda x: (x["faltante_status"] != "pendente", x["data"], x["sku"]))
    return out


# ---------------------------------------------------------------------------
# Planilha 2 · "Resumo de Rebates" (Relatorios_PedidosxRebates_BI_MercadoLivre)
# Uma linha por pedido, com o MLB, o SKU, o tipo de anúncio e a comissão bruta
# de tabela (Premium 16,5% · Clássico 11,5%). Daqui nasce a "Lista de MLB's".
# ---------------------------------------------------------------------------
COLS_REBATES = {
    "pedido_canal": "Pedido Canal", "pedido_mkt": "Pedido Marketplace (Any)", "data": "Data Pedido",
    "sku": "SKU", "anuncio": "Anúncio", "tipo": "Tipo Anúncio", "qtd": "Quantidade",
    "valor_prod": "Valor Produtos Meli (+)", "frete": "Frete Pedido", "frete_coletas": "Frete Coletas",
    "cupom_canal": "Cupom Canal", "cupom_seller": "Cupom Seller Meli", "valor_meli": "Valor Pedido Meli",
    "pct_bruta": "% Comissão (Bruta)", "com_bruta": "Comissão (Bruta)", "rebate_bi": "Rebate",
    "pct_liq": "% Comissão (Líquida)", "com_liq": "Comissão (Líquida)",
}
OBRIG_REBATES = ["pedido_mkt", "data", "sku", "anuncio", "tipo", "pct_bruta"]


def ler_resumo_rebates(caminho: str) -> tuple[pd.DataFrame, dict]:
    """Lê a Planilha 2 (Resumo de Rebates). Devolve (df normalizado, diagnóstico)."""
    inv = {_norm(v): k for k, v in COLS_REBATES.items()}
    xl = pd.ExcelFile(caminho)
    escolhida, df = None, None
    for aba in xl.sheet_names:
        d = xl.parse(aba, dtype=str)
        mapa = {c: inv[_norm(c)] for c in d.columns if _norm(c) in inv}
        if all(k in mapa.values() for k in OBRIG_REBATES):
            escolhida, df = aba, d.rename(columns=mapa)
            break
    if df is None:
        raise ValueError("Nenhuma aba tem as colunas do Resumo de Rebates. Preciso pelo menos de: "
                         + ", ".join(COLS_REBATES[k] for k in OBRIG_REBATES) + ".")
    df = df.copy()
    for k in ("pedido_canal", "pedido_mkt"):
        df[k] = df[k].map(_pedido) if k in df else ""
    for k in ("qtd", "valor_prod", "frete", "frete_coletas", "cupom_canal", "cupom_seller", "valor_meli",
              "pct_bruta", "com_bruta", "rebate_bi", "pct_liq", "com_liq"):
        df[k] = df[k].map(_num) if k in df else 0.0
    for k in ("sku", "anuncio", "tipo"):
        df[k] = df[k].fillna("").astype(str).str.strip() if k in df else ""
    df["data"] = df["data"].map(_data)
    ruim = df["pedido_mkt"].eq("") | df["data"].isna() | df["anuncio"].eq("")
    diag: dict[str, Any] = {"aba": escolhida, "linhas_brutas": int(len(df)), "n_rejeitadas": int(ruim.sum())}
    df = df[~ruim].copy()
    df["competencia"] = df["data"].map(lambda d: f"{d.year}-{d.month:02d}")
    diag["competencias"] = df["competencia"].value_counts().sort_index().to_dict()
    diag["de"], diag["ate"], diag["linhas"] = str(df["data"].min()), str(df["data"].max()), int(len(df))
    diag["tipos"] = df["tipo"].value_counts().to_dict()
    diag["mlbs"] = int(df["anuncio"].nunique())
    return df, diag


def lista_mlbs(df: pd.DataFrame) -> dict:
    """MLB → {sku, tipo, pct, primeira, ultima, pedidos}. Vale o que a última
    venda diz (um MLB pode mudar de Clássico para Premium); o histórico fica
    em 'pedidos'."""
    out: dict[str, dict] = {}
    d = df.sort_values("data")
    for r in d.itertuples(index=False):
        m = out.get(r.anuncio)
        if m is None:
            m = out[r.anuncio] = {"mlb": r.anuncio, "sku": r.sku, "tipo": r.tipo, "pct": float(r.pct_bruta),
                                  "primeira": str(r.data), "ultima": str(r.data), "pedidos": 0}
        m["sku"], m["tipo"], m["pct"], m["ultima"] = r.sku, r.tipo, float(r.pct_bruta), str(r.data)
        m["pedidos"] += 1
    return out
