"""Box AMAZON — nasce do relatório "Transações" do Seller Central
(Amazon_Transações referentes ao período de DD_MM_AAAA a DD_MM_AAAA_HHMM.csv).

O ARQUIVO (1 linha por TRANSAÇÃO, csv ',' com aspas, utf-8-sig):
  Data · Status da transação · Grupo de Pagamento · Tipo de transação ·
  ID do pedido · Detalhes do produto · Custo total do produto ·
  Total de descontos promocionais · Tarifas da Amazon · Outros · (total) (BRL)

TRÊS TIPOS DE TRANSAÇÃO, cada um com um destino diferente:
  · "Pagamento do pedido"  → o pedido (venda, comissão, frete, repasse)
  · "Reembolso"            → devolução: sai dos rebates (a Amazon estorna parte
                             da comissão junto) — mesma regra do Mercado Livre
  · "Tarifas de serviço"   → publicidade (Amazon Ads) e afins. NÃO é comissão e
                             NÃO entra no rebate: é custo de mídia, vai ao ORION

ACUMULATIVO (regra da casa, 18/09/2026): os relatórios chegam se sobrepondo.
A base guarda TRANSAÇÃO a TRANSAÇÃO, com chave própria, e o pedido é montado a
partir delas — assim o mesmo ID do pedido é usado UMA VEZ SÓ na apuração, mesmo
aparecendo em três arquivos, e um reembolso que chega depois encontra o
pagamento que veio antes.

A REGRA DA COMISSÃO (medida em 575 pedidos de 12–18/09/2026, ao centavo):
    Tarifas da Amazon = (% negociado da categoria + 1,5% de taxa)
                        × (produto − desconto promocional + frete)
  Sem taxa em R$. As faixas cobradas são 5,5 · 6,5 · 7,5 · 8,0 · 8,5 · 9,0 ·
  9,5 · 10,5% — que são as comissões negociadas (4 a 9%) mais os 1,5% de taxa.
  Em 302 pedidos da faixa cheia (10,5% = 9% + 1,5%) o total fechou com R$ 0,13
  de diferença.

  O cadastro da casa tem UM percentual para toda a Amazon (10,5%). Por isso a
  diferença é separada em duas — problemas diferentes, donos diferentes:
    · REBATE DE COMISSÃO = faixa da categoria × base − comissão cobrada.
      Erro de cobrança da Amazon; é o que a Gabi cobra de volta.
    · DESVIO DE CADASTRO = % do cadastro × base − faixa da categoria × base.
      Não é rebate: é o sistema prevendo comissão errada em categoria de
      comissão menor. Estraga margem e preço (ORION); corrige-se no cadastro.

Chave do pedido = "ID do pedido" (701-/702-…), que é a própria Ordem de compra
do ERP — o casamento é direto, sem de-para.
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

CANAL = "amazon"
TAXA_PADRAO = 0.015          # a taxa por pedido embutida no % cobrado
# faixas cobradas pela Amazon = % negociado da categoria + 1,5% de taxa
FAIXAS = (0.055, 0.060, 0.065, 0.070, 0.075, 0.080, 0.085, 0.090, 0.095, 0.100, 0.105,
          0.115, 0.120, 0.150, 0.165)
TOL_FAIXA = 0.0015           # 0,15 p.p. de folga (arredondamento de centavos)
# Dias entre o PEDIDO e a TRANSAÇÃO de pagamento: medido entre 2 e 17 dias
# (mediana 8) nos pedidos que casaram com o ERP. Abaixo dessa folga, o pedido
# que não apareceu no relatório ainda pode estar só esperando o repasse.
FOLGA_REPASSE = 20

T_PEDIDO, T_REEMB, T_SERV = "pagamento do pedido", "reembolso", "tarifas de servico"


def _norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v == v else 0.0
    t = str(v).strip().replace("R$", "").replace(" ", "").replace("\u00a0", "")
    if not t or t in ("-", "nan", "None", "--"):
        return 0.0
    # O relatório da Amazon vem em formato AMERICANO ("697.48", às vezes
    # "1,234.56"), mas o mesmo arquivo pode chegar convertido para o formato
    # brasileiro ("1.234,56"). Quem manda é o ÚLTIMO separador da string: ele é
    # o decimal, e o outro é separador de milhar.
    ult_v, ult_p = t.rfind(","), t.rfind(".")
    if ult_v >= 0 and ult_p >= 0:
        if ult_v > ult_p:                 # 1.234,56 → brasileiro
            t = t.replace(".", "").replace(",", ".")
        else:                             # 1,234.56 → americano
            t = t.replace(",", "")
    elif ult_v >= 0:
        t = t.replace(",", ".")           # só vírgula: é o decimal (99,99)
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
    "data": "data", "status da transacao": "status", "grupo de pagamento": "pagamento",
    "tipo de transacao": "tipo_tr", "id do pedido": "pedido", "detalhes do produto": "produto",
    "custo total do produto": "prod", "total de descontos promocionais": "desconto",
    "tarifas da amazon": "tarifa_raw", "outros": "outros", "(total) (brl)": "repasse",
}
OBRIGATORIAS = ["data", "tipo_tr", "prod", "tarifa_raw", "repasse"]


def faixa_de(pct: float) -> float | None:
    """O % cobrado cai em alguma faixa de categoria da Amazon?"""
    for f in FAIXAS:
        if abs(pct - f) < TOL_FAIXA:
            return f
    return None


def _ler_bruto(caminho: str) -> pd.DataFrame:
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
    # O Seller Central às vezes entrega o csv com a LINHA INTEIRA entre aspas e
    # as aspas internas dobradas ("01/09/2026,""Liberado"",…"). Sem desembrulhar,
    # o arquivo vira uma coluna só e nada é lido. Desfaz antes de parsear.
    linhas = txt.splitlines()
    envolvidas = sum(1 for l in linhas[:50] if len(l) > 2 and l[0] == '"' and l[-1] == '"' and '""' in l)
    if envolvidas >= max(2, len(linhas[:50]) // 2):
        txt = "\n".join((l[1:-1].replace('""', '"') if (len(l) > 2 and l[0] == '"' and l[-1] == '"') else l)
                        for l in linhas)
    amostra = txt[:4000]
    sep = max((",", ";", "\t"), key=lambda s: amostra.count(s))
    return pd.read_csv(io.StringIO(txt), sep=sep, dtype=str, quoting=csv.QUOTE_MINIMAL)


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    """Devolve as TRANSAÇÕES normalizadas (1 linha por transação, com chave
    própria) — é isso que a base acumula."""
    d = _ler_bruto(caminho)
    mapa = {}
    for c in d.columns:
        n = _norm(c).strip('"')
        if n in COLS and COLS[n] not in mapa.values():
            mapa[c] = COLS[n]
    df = d.rename(columns=mapa).copy()
    if not all(k in df.columns for k in OBRIGATORIAS):
        raise ValueError("Não parece o relatório de TRANSAÇÕES da Amazon (Seller Central → Pagamentos → "
                         "Todas as transações → Baixar). Preciso das colunas Data, Tipo de transação, "
                         "ID do pedido, Custo total do produto, Tarifas da Amazon e (total) (BRL).")
    for k in COLS.values():
        if k not in df:
            df[k] = None
    for k in ("pedido", "produto", "status", "pagamento", "tipo_tr"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    for k in ("prod", "desconto", "tarifa_raw", "outros", "repasse"):
        df[k] = df[k].map(_num)
    df["data"] = df["data"].map(_data)
    brutas = int(len(df))
    df = df[df["data"].notna()].copy()
    df["t"] = df["tipo_tr"].map(_norm)
    df["data"] = df["data"].map(str)

    # chave da TRANSAÇÃO: o relatório não tem id de linha, então a identidade é
    # (dia + tipo + pedido + valores) — e o ordinal separa duas iguais de verdade.
    def _ch(r, i):
        return f"{r['data']}|{r['t'][:12]}|{r['pedido']}|{r['prod']:.2f}|{r['tarifa_raw']:.2f}|{r['repasse']:.2f}|{i}"

    vistas: dict[str, int] = {}
    chaves = []
    for _, r in df.iterrows():
        b = _ch(r, 0)[:-2]
        i = vistas.get(b, 0)
        vistas[b] = i + 1
        chaves.append(f"{b}|{i}")
    df["chave"] = chaves
    df["competencia"] = df["data"].map(lambda s: s[:7])

    serv = df[df["t"] == T_SERV]
    reem = df[df["t"].str.startswith(T_REEMB)]      # inclui "Reembolso de estorno"
    pag = df[df["t"] == T_PEDIDO]
    diag: dict[str, Any] = {
        "aba": "transações", "linhas_brutas": brutas, "linhas": int(len(df)),
        "transacoes": {k: int(v) for k, v in df["tipo_tr"].value_counts().to_dict().items()},
        "pagamentos": int(len(pag)), "pedidos_no_arquivo": int(pag["pedido"].nunique()),
        "reembolsos": int(len(reem)), "reembolso_rs": round(-float(reem["repasse"].sum()), 2),
        "servicos": round(-float(serv["repasse"].sum()), 2), "servicos_dias": int(serv["data"].nunique()),
        "publicidade": round(-float(serv[serv["produto"].map(_norm).str.contains("public")]["repasse"].sum()), 2),
        "status": df["status"].value_counts().to_dict(),
        "competencias": df["competencia"].value_counts().sort_index().to_dict(),
        "de": str(df["data"].min()), "ate": str(df["data"].max()),
        "dias": int(df["data"].nunique()),
        "n_rejeitadas": brutas - int(len(df)),
    }
    if not len(pag):
        raise ValueError("O arquivo não tem nenhuma linha 'Pagamento do pedido' — é outro relatório da Amazon.")
    return df, diag


def pedidos(tx: pd.DataFrame) -> pd.DataFrame:
    """Monta 1 LINHA POR PEDIDO a partir das transações acumuladas — o ID do
    pedido entra uma vez só, somando o que veio de todos os arquivos."""
    tx = tx.copy()
    if "t" not in tx:
        tx["t"] = tx["tipo_tr"].map(_norm)
    pag = tx[tx["t"] == T_PEDIDO]
    pag = pag[pag["pedido"] != ""]
    if not len(pag):
        return pd.DataFrame()
    reem = tx[tx["t"].str.startswith(T_REEMB)]      # reembolso e reembolso de estorno
    reemb = reem.groupby("pedido")["repasse"].sum().to_dict() if len(reem) else {}
    reemb_n = reem.groupby("pedido")["repasse"].size().to_dict() if len(reem) else {}

    g = pag.groupby("pedido", sort=False)
    ped = g.agg(data=("data", "min"), status=("status", "first"), pagamento=("pagamento", "first"),
                produto=("produto", "first"), linhas=("prod", "size"),
                prod=("prod", "sum"), desconto=("desconto", "sum"), tarifa_raw=("tarifa_raw", "sum"),
                outros=("outros", "sum"), repasse=("repasse", "sum")).reset_index()

    ped["tarifa"] = ped["tarifa_raw"].map(lambda v: round(abs(v), 2))
    ped["desconto"] = ped["desconto"].map(lambda v: round(abs(v), 2))
    ped["prod"] = ped["prod"].round(2)
    ped["outros"] = ped["outros"].round(2)
    ped["repasse"] = ped["repasse"].round(2)
    ped["base"] = (ped["prod"] - ped["desconto"] + ped["outros"]).round(2)
    ped["pct_real"] = [(t / b if b else 0.0) for t, b in zip(ped["tarifa"], ped["base"])]
    ped["faixa"] = ped["pct_real"].map(faixa_de)
    ped["reembolso"] = ped["pedido"].map(lambda p: round(-float(reemb.get(p, 0.0)), 2))
    ped["reembolsos"] = ped["pedido"].map(lambda p: int(reemb_n.get(p, 0)))
    ped["devolvido"] = [bool(r > 0 and r >= 0.9 * rp) for r, rp in zip(ped["reembolso"], ped["repasse"])]
    ped["competencia"] = ped["data"].map(lambda s: str(s)[:7])
    return ped


def calcular(ped: pd.DataFrame, pct: float, taxa_pct: float = TAXA_PADRAO,
             erp_idx: dict | None = None, tolerancia: float = 0.50) -> list[dict]:
    """pct = % cheio dos Parâmetros (já com a taxa dentro, ex.: 10,5%).
    taxa_pct = a parte do % que é taxa por pedido (1,5%), só para mostrar.
    erp_idx = {OC: {...}} do ERP — a OC da Amazon é o próprio ID do pedido."""
    out = []
    erp_idx = erp_idx or {}
    for r in ped.itertuples(index=False):
        if r.devolvido:
            continue
        base = float(r.base)
        real = float(r.tarifa)
        e = erp_idx.get(r.pedido) or {}
        # COMISSÃO DO SISTEMA (regra da casa, 18/09/2026): % do cadastro (10,5%)
        # sobre o TOTAL DO PEDIDO do ERP (a NF: produto + IPI + frete). O pedido
        # que veio do ERP com % defasado (9,2%) é RECALCULADO aqui; o que já veio
        # com o % certo fica como está.
        pct_erp = float(e.get("pct_comissao") or 0.0) or None
        pct_cad = pct
        corrigido = bool(pct_erp and abs(pct_erp - pct) > 0.0005)
        base_erp = round(float(e.get("total_erp") or 0.0), 2)
        base_sis = base_erp if base_erp else base      # sem par no ERP: base do próprio relatório
        sis_rs = round(base_sis * pct_cad, 2)                   # o que o sistema deveria ter previsto
        faixa = float(r.faixa) if (r.faixa is not None and r.faixa == r.faixa) else None
        faixa_rs = round(base * faixa, 2) if faixa else None
        # Sem faixa identificada quase sempre é PEDIDO MISTO (itens de categorias
        # diferentes): o % efetivo cai entre duas faixas e a cobrança pode estar
        # certa. Nesses o alvo é a própria cobrança — não inventa rebate —, e a
        # linha fica marcada para conferência.
        alvo = faixa_rs if faixa_rs is not None else real
        # REBATE = comissão do sistema (ERP a 10,5% da NF) − comissão cobrada.
        # É a régua da casa, a mesma dos outros canais.
        dif = round(sis_rs - real, 2)
        reb_com = dif if abs(dif) > tolerancia else 0.0
        # DESVIO DE CADASTRO = o pedaço da diferença que a faixa da categoria
        # explica (categoria paga menos que o cadastro) — não é erro da Amazon.
        desvio = round(sis_rs - alvo, 2)
        erro_amazon = round(alvo - real, 2)
        out.append({
            "canal": CANAL,
            "pedido_mkt": r.pedido, "pedido_canal": r.pedido, "pedido_any": e.get("obs05", "") or "",
            "id_mkt": "", "data": str(r.data), "competencia": r.competencia,
            "conta": "", "status": r.status, "pagamento": r.pagamento,
            "sku": "", "anuncio": "", "produto": r.produto,
            "tipo": (f"{faixa * 100:.1f}%" if faixa else "fora de faixa"),
            "negociada": (round(faixa - taxa_pct, 4) if faixa else None),
            "itens": int(r.linhas), "qtd": 1.0,
            "valor_prod": float(r.prod), "frete": float(r.outros), "desconto": float(r.desconto),
            "repasse": float(r.repasse),
            "sis_base": base_sis, "sis_base_nome": ("total do pedido no ERP (NF)" if base_erp else "produto − desconto + frete (sem par no ERP)"),
            "base_amazon": base, "base_amazon_nome": "produto − desconto + frete",
            "tarifa": real, "pct_comissao": round(float(r.pct_real), 4),
            "faixa": faixa, "faixa_rs": faixa_rs, "sem_faixa": faixa is None,
            "sis_pct": round(pct_cad, 6), "sis_rs": sis_rs, "erp_ok": bool(e),
            "sis_base_erp": base_erp, "base_erp_nome": "total do pedido no ERP (NF)",
            "erp_pct": (round(pct_erp, 6) if pct_erp else None),
            "erp_difere": corrigido, "pct_corrigido": corrigido,
            "desvio_cadastro": desvio, "erro_amazon": erro_amazon,
            "diferenca": dif,
            "tarifa_zero": bool(real <= 0.005 and base > 0),
            "faltante": 0.0, "faltante_status": "", "sem_sistema": not bool(e),
            "cupom_seller": 0.0, "cupom_meli": 0.0,
            "reembolso": float(r.reembolso),
            "rebate_rs": 0.0, "rebate_comissao": reb_com, "rebate_frete": 0.0, "rebate_total": reb_com,
        })
    return out


def resumo(linhas: list[dict], extra: dict | None = None) -> dict:
    extra = extra or {}
    n = len(linhas)
    s = lambda k: round(sum((l.get(k) or 0.0) for l in linhas), 2)  # noqa: E731
    venda = s("valor_prod")
    base = s("base_amazon")          # base do relatório da Amazon
    base_sis = s("sis_base")         # base do sistema (NF do ERP)
    por_dia: dict[str, dict] = {}
    for l in linhas:
        d = por_dia.setdefault(l["data"], {"pedidos": 0, "venda": 0.0, "cupom": 0.0, "faltante": 0.0,
                                           "comissao": 0.0, "frete": 0.0, "total": 0.0, "tz": 0})
        d["pedidos"] += 1
        d["venda"] += l["valor_prod"]
        d["frete"] += l["frete"]
        d["comissao"] += l["rebate_comissao"]
        d["total"] += l["rebate_total"]
    for d in por_dia.values():
        for k in ("venda", "cupom", "faltante", "comissao", "frete", "total"):
            d[k] = round(d[k], 2)

    com_real = s("tarifa")
    com_faixa = round(sum((l["faixa_rs"] if l["faixa_rs"] is not None else l["sis_rs"]) for l in linhas), 2)
    com_sis = arred.total(linhas, False)          # cadastro × base, arredondado 1x no total
    reb_com = round(sum(l["rebate_comissao"] for l in linhas), 2)
    desvio = round(com_sis - com_faixa, 2)

    faixas: dict[str, dict] = {}
    for l in linhas:
        f = faixas.setdefault(l["tipo"], {"pedidos": 0, "base": 0.0, "real": 0.0, "faixa": 0.0, "sis": 0.0,
                                          "negociada": l.get("negociada")})
        f["pedidos"] += 1
        f["base"] += l["base_amazon"]
        f["real"] += l["tarifa"]
        f["faixa"] += (l["faixa_rs"] if l["faixa_rs"] is not None else l["sis_rs"])
        f["sis"] += l["sis_rs"]
    for f in faixas.values():
        for k in ("base", "real", "faixa", "sis"):
            f[k] = round(f[k], 2)
        f["desvio"] = round(f["sis"] - f["faixa"], 2)
        f["share"] = 0.0
    for f in faixas.values():
        f["share"] = round(100 * f["base"] / base, 1) if base else 0.0
    faixas = dict(sorted(faixas.items(), key=lambda kv: -kv[1]["base"]))

    return {
        "pedidos": n, "cancelados": int(extra.get("devolvidos") or 0),
        "venda": venda, "base": base, "tarifa": com_real, "frete": s("frete"),
        "desconto": s("desconto"), "repasse": s("repasse"),
        "cupom_meli": 0.0, "cupom_seller": 0.0, "faltante": 0.0,
        "rebate_rs": 0.0, "rebate_comissao": reb_com, "rebate_frete": 0.0, "rebate_total": reb_com,
        "pct_sobre_venda": (round(100 * reb_com / venda, 2) if venda else 0.0),
        "tarifa_zero": sum(1 for l in linhas if l["tarifa_zero"]),
        "faltante_pendentes": 0, "faltante_preenchidos": 0,
        "dif_pos": sum(1 for l in linhas if l["diferenca"] > 0.5),
        "dif_pos_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] > 0.5), 2),
        "dif_neg": sum(1 for l in linhas if l["diferenca"] < -0.5),
        "dif_neg_rs": round(sum(l["diferenca"] for l in linhas if l["diferenca"] < -0.5), 2),
        "sem_sistema": sum(1 for l in linhas if l["sem_sistema"]),
        "erp_ok": sum(1 for l in linhas if l.get("erp_ok")),
        "sem_faixa": sum(1 for l in linhas if l["sem_faixa"]),
        "erp_difere": sum(1 for l in linhas if l.get("erp_difere")),
        "corrigidos": sum(1 for l in linhas if l.get("pct_corrigido")),
        "base_erp": round(sum((l.get("sis_base_erp") or 0.0) for l in linhas), 2),
        "erro_amazon": round(sum((l.get("erro_amazon") or 0.0) for l in linhas), 2),
        "sem_erp_base": sum(1 for l in linhas if not l.get("sis_base_erp")),
        "sem_faixa_rs": round(sum(l["tarifa"] for l in linhas if l["sem_faixa"]), 2),
        "com_pedidos": n, "com_sistema": com_sis, "com_real": com_real, "com_venda": base_sis or base,
        "base_sistema": base_sis,
        "com_faixa": com_faixa, "desvio_cadastro": desvio,
        "com_sistema_pct": (round(100 * com_sis / base_sis, 2) if base_sis else 0.0),
        "com_real_pct": (round(100 * com_real / base, 2) if base else 0.0),
        "com_faixa_pct": (round(100 * com_faixa / base, 2) if base else 0.0),
        "com_dif": reb_com,
        "publicidade": float(extra.get("publicidade") or 0.0),
        "servicos": float(extra.get("servicos") or 0.0),
        "servicos_dias": int(extra.get("servicos_dias") or 0),
        "reembolso_rs": float(extra.get("reembolso_rs") or 0.0),
        "reembolsos": int(extra.get("reembolsos") or 0),
        "devolvidos": int(extra.get("devolvidos") or 0),
        "faixas": faixas, "contas": {}, "tipos": {k: v["pedidos"] for k, v in faixas.items()},
        "status_validos": {},
        "por_dia": dict(sorted(por_dia.items())),
    }
