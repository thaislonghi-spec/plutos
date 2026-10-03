"""CASAS BAHIA (Via) — REPASSE FINANCEIRO.

O QUE ESTE RELATÓRIO É, E POR QUE ELE NÃO É COMO OS OUTROS CANAIS
    Nos outros seis boxes o arquivo do canal é uma APURAÇÃO: um mês de vendas,
    com a comissão que o canal cobrou. Aqui não. O relatório da Via é um
    EXTRATO DE REPASSE: cada linha é dinheiro entrando (ou saindo) da conta da
    Multimóveis numa data. Um arquivo de setembro traz pedido de fevereiro que
    só foi liberado agora — o canal só libera o repasse quando o pedido está
    ENTREGUE (regra da Gabi, 02/10/2026).

    Daí a regra da casa para este canal, definida pela Gabi:
      · é REGIME DE CAIXA. O que veio no relatório é dinheiro recebido e conta
        no mês do REPASSE, mesmo que o pedido seja antigo;
      · o rebate de pedido de mês anterior NÃO volta para o mês dele (não
        mexe em mês já fechado): entra neste mês como SALDO DE MESES ANTERIORES,
        somado e separado;
      · pedido que nem está na nossa base entra igual — é dinheiro na conta.

DE ONDE SAEM OS ARQUIVOS
    Três relatórios por mês, mesmas 52 colunas, telas diferentes do portal:
      · BOLETO — ciclo do dia 06 ao 05, repasse no dia 20. Para fechar um mês
        vêm dois: 06/mês-1 a 05/mês e 06/mês a 05/mês+1.
      · CARTÃO — repasse semanal, datas soltas dentro do mês.
    O PLUTOS separa sozinho pela coluna "Origem Repasse": INTERNO = boleto,
    GETNET (ou outra adquirente) = cartão.

O REBATE — as duas parcelas da regra da Gabi
    1) linha de VENDA ....... soma "Desconto Ônus Via" (coluna O): o desconto
       que o CANAL bancou na venda;
    2) linha de AJUSTE com motivo "Credito de Campanha" ..... soma "Valor
       Líquido de Repasse" (coluna AE): campanha que não foi descontada da
       comissão na hora e o canal paga depois.
    3) REBATE DE COMISSÃO ... o que o canal cobrou ABAIXO dos 13% de contrato
       (regra da Thaís, 02/10/2026 — detalhe na seção da régua, abaixo).
    O mesmo pedido pode ter as três.

CANCELAMENTO
    Linha de CANCELAMENTO zera o rebate daquele pedido e ele fica SINALIZADO —
    pedido cancelado não gera rebate (regra da casa, vale em todos os canais).

A RÉGUA É O CONTRATO DA CASA: 13% (regra da Thaís, 02/10/2026)
    O relatório traz uma "Comissão Contratual %" que varia de 8% a 19% por
    categoria — mas essa é a tabela do canal, NÃO o que foi negociado. O
    contrato da Multimóveis com a Via é de 13%, e é essa a régua:

      · cobrou ACIMA de 13% ...... cobrança indevida. Vai para a subaba
        "Comissões a maior", para o canal ajustar a menos.
      · cobrou ABAIXO de 13% ..... é REBATE DE COMISSÃO. Entra na conta.

    Os dois NÃO se compensam: um é dinheiro a cobrar de volta, o outro é
    dinheiro que já ficou com a gente. Somar um no outro esconderia as duas
    coisas.

    O percentual vem dos Parâmetros (canal CASAS BAHIA), não está cravado no
    código: mudou o contrato, muda no cadastro e o mês recalcula — sem precisar
    subir arquivo de novo.

A BASE da comissão é o "Valor da Transação" (coluna U): testadas cinco bases,
essa fecha com a comissão cobrada em 841 de 868 linhas.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any

import pandas as pd

CANAL = "cbahia"
T_VENDA, T_AJUSTE, T_CANCEL = "VENDA", "AJUSTE", "CANCELAMENTO"
M_CREDITO = "credito de campanha"
M_ESTORNO = "estorno de campanha"
# MOTIVOS DE AJUSTE QUE NÃO ENTRAM NO REBATE (definido com a Gabi, 02/10/2026)
#   ads · outros · imposto de renda ... não são rebate, não entram. Ponto.
#   estorno de campanha ............. é a DEVOLUÇÃO do rebate quando o pedido
#     é cancelado. Conferido nos 3 casos dos arquivos de set/26: em 3 de 3 o
#     estorno vem junto com um CANCELAMENTO do mesmo pedido e o valor bate
#     exatamente com o "Desconto Ônus Via" dele, com o sinal trocado:
#         51411981001  ônus 494,00  estorno −494,00  cancelado
#         52052765601  ônus  42,50  estorno  −42,50  cancelado
#         50650468801  ônus  52,50  estorno  −52,50  cancelado
#     Como a regra da casa JÁ zera o rebate do pedido cancelado, somar o
#     estorno tiraria o mesmo dinheiro duas vezes. Por isso ele fica de fora
#     da conta — mas aparece na tela, ligado ao cancelamento que o explica.
M_FORA = ("outros", "ads", "imposto de renda", "estorno de campanha")
# MOTIVOS QUE JÁ CONHECEMOS. Qualquer motivo FORA desta lista é sinalizado alto
# na tela e na mensagem de leitura (pedido da Thaís, 02/10/2026): quando o canal
# percebe que cobrou comissão errada, ele costuma devolver como AJUSTE — e pode
# inventar uma nomenclatura nova para isso ("diferença de comissão", "acerto de
# comissão"…). Se esse motivo novo passar despercebido, o dinheiro entra na
# conta e não vira rebate para ninguém. Por isso: motivo desconhecido = alerta.
M_CONHECIDOS = ("credito de campanha", "estorno de campanha", "outros", "ads",
                "imposto de renda")
TOL_PCT = 0.01   # ponto percentual
TOL_RS = 0.05    # centavos: abaixo disso não é diferença, é arredondamento


def _norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(v) -> float:
    """BR com R$ e sinal: 'R$ 1.234,56' · '-R$ 849,76' · '1.234,56'."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v == v else 0.0
    t = str(v).strip()
    if not t or _norm(t) in ("nan", "none", "-", "--"):
        return 0.0
    neg = t.startswith("-") or t.startswith("(")
    t = t.replace("R$", "").replace("(", "").replace(")", "").replace("-", "").replace(" ", "").strip()
    uv, up = t.rfind(","), t.rfind(".")
    if uv >= 0 and up >= 0:
        t = t.replace(".", "").replace(",", ".") if uv > up else t.replace(",", "")
    elif uv >= 0:
        t = t.replace(",", ".")
    try:
        x = float(t)
    except ValueError:
        return 0.0
    return -x if neg else x


def _data(v) -> str:
    t = str(v or "").strip()[:19]
    if not t or _norm(t) in ("nan", "none"):
        return ""
    d = pd.to_datetime(t, errors="coerce", dayfirst=True)
    return "" if d is pd.NaT or d != d else d.date().isoformat()


COLS = {
    "numero pedido": "pedido", "id entrega": "entrega", "tipo da transacao": "tipo",
    "data do pedido incluido": "data_pedido", "data do pedido entregue": "data_entrega",
    "data liberacao": "data_liberacao", "data prevista do repasse": "data_prevista",
    "data do repasse": "data_repasse",
    "sku lojista": "sku", "descricao do produto": "produto",
    "departamento": "departamento", "categoria": "categoria",
    "valor do produto sem desconto": "produto_rs", "desconto onus via": "onus_via",
    "desconto onus lojista": "onus_lojista", "valor do frete": "frete",
    "frete promocional onus via": "frete_via", "frete promocional onus lojista": "frete_lojista",
    "valor da transacao": "transacao",
    "comissao contratual %": "pct_contratual", "comissao aplicada %": "pct_aplicada",
    "comissao aplicada r$": "comissao_rs",
    "parcela atual": "parcela", "numero de liquidacao": "liquidacao",
    "valor bruto de repasse": "repasse_bruto", "valor liquido de repasse": "repasse_liquido",
    "motivo do ajuste": "motivo", "observacoes": "obs",
    "codigo identificador operacao": "operacao", "origem repasse": "origem",
    "meio de pagamento": "meio_pgto", "tipo de campanha": "tipo_campanha",
    "nf cliente": "nf",
}
OBRIGATORIAS = ["pedido", "tipo", "onus_via", "repasse_liquido"]


def _ler_bruto(caminho: str) -> pd.DataFrame:
    if caminho.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(caminho, dtype=str)
    for sep in (";", ",", None):
        try:
            d = pd.read_csv(caminho, sep=sep, engine="python", dtype=str, encoding="utf-8-sig")
            if d.shape[1] > 5:
                return d
        except Exception:  # noqa: BLE001
            continue
    raise ValueError("não consegui abrir o arquivo da Casas Bahia")


def origem_de(v: str) -> str:
    """INTERNO = boleto · adquirente (GETNET…) = cartão."""
    n = _norm(v)
    return "boleto" if (not n or n == "interno") else "cartao"


def ler(caminho: str) -> tuple[pd.DataFrame, dict]:
    d = _ler_bruto(caminho)
    mapa = {c: COLS[_norm(c)] for c in d.columns if _norm(c) in COLS}
    df = d.rename(columns=mapa).copy()
    if not all(k in df.columns for k in OBRIGATORIAS):
        raise ValueError(
            "Não parece o relatório de REPASSE da Casas Bahia (Via). Preciso das colunas "
            "Número Pedido, Tipo da Transação, Desconto Ônus Via e Valor Líquido de Repasse — "
            "são as telas de repasse por BOLETO e por CARTÃO do portal.")
    for k in set(COLS.values()):
        if k not in df:
            df[k] = None
    for k in ("produto_rs", "onus_via", "onus_lojista", "frete", "frete_via", "frete_lojista",
              "transacao", "pct_contratual", "pct_aplicada", "comissao_rs",
              "repasse_bruto", "repasse_liquido"):
        df[k] = df[k].map(_num)
    for k in ("data_pedido", "data_entrega", "data_liberacao", "data_prevista", "data_repasse"):
        df[k] = df[k].map(_data)
    for k in ("pedido", "entrega", "tipo", "motivo", "sku", "produto", "categoria",
              "departamento", "origem", "meio_pgto", "parcela", "liquidacao", "operacao", "nf"):
        df[k] = df[k].fillna("").astype(str).str.strip()
    df = df[df["pedido"] != ""].copy()

    df["tipo"] = df["tipo"].str.upper()
    df["forma"] = df["origem"].map(origem_de)
    # A DATA DO DINHEIRO É A COLUNA G · "Data Liberação" (regra da Gabi,
    # 02/10/2026): não existe data fixa de repasse — é a liberação que manda.
    # As linhas de AJUSTE vêm SEM essa data (o ajuste não tem liberação
    # própria): nelas cai para a data do repasse e, por fim, para a prevista,
    # senão o crédito de campanha ficaria sem mês e sumiria da conta.
    df["data_caixa"] = [lb or rp or pv for lb, rp, pv in
                        zip(df["data_liberacao"], df["data_repasse"], df["data_prevista"])]
    df = df[df["data_caixa"] != ""].copy()
    df["competencia"] = df["data_caixa"].str[:7]          # mês do DINHEIRO
    df["comp_pedido"] = df["data_pedido"].str[:7]         # mês da VENDA
    df["motivo_n"] = df["motivo"].map(_norm)

    # ---- as duas parcelas do rebate ----------------------------------
    venda = df["tipo"] == T_VENDA
    cred = (df["tipo"] == T_AJUSTE) & (df["motivo_n"] == M_CREDITO)
    df["reb_onus"] = [round(v, 2) if t else 0.0 for v, t in zip(df["onus_via"], venda)]
    df["reb_campanha"] = [round(v, 2) if t else 0.0 for v, t in zip(df["repasse_liquido"], cred)]
    df["rebate"] = (df["reb_onus"] + df["reb_campanha"]).round(2)
    df["ajuste_fora"] = (df["tipo"] == T_AJUSTE) & (df["motivo_n"] != "") & ~df["motivo_n"].isin(("credito de campanha",))
    cancelados_ = set(df.loc[df["tipo"] == T_CANCEL, "pedido"])
    df["estorno"] = (df["tipo"] == T_AJUSTE) & (df["motivo_n"] == M_ESTORNO)
    df["estorno_com_cancel"] = df["estorno"] & df["pedido"].isin(cancelados_)
    # motivo de ajuste que o PLUTOS ainda não conhece — pode ser devolução de
    # comissão cobrada a maior com outro nome
    df["motivo_novo"] = ((df["tipo"] == T_AJUSTE) & (df["motivo_n"] != "")
                         & ~df["motivo_n"].isin(M_CONHECIDOS))

    # A comparação com o contrato (13%) NÃO é feita aqui: ela depende do % dos
    # Parâmetros e tem de reagir a uma mudança de cadastro sem re-subir
    # arquivo. Fica em comissao_vs_contrato(), chamada no recálculo.
    df["pct_dif_tabela"] = (df["pct_aplicada"] - df["pct_contratual"]).round(4)
    # ... mas a comparação APLICADA × CONTRATUAL do próprio relatório continua
    # valendo como SENTINELA: ela responde se o canal respeitou a tabela DELE.
    # Se um dia a aplicada sair da contratual, aí sim é erro do canal, e não
    # diferença de tabela por categoria.
    df["tabela_fora"] = venda & (df["pct_dif_tabela"].abs() > TOL_PCT)

    # ---- chave da base acumulada --------------------------------------
    # O mesmo pedido pode ter DUAS LINHAS IDÊNTICAS de verdade (2 unidades do
    # mesmo SKU, cada uma com venda e cancelamento). Por isso a chave leva a
    # ORDEM da repetição: subir o mesmo arquivo de novo casa linha a linha e
    # não duplica; um arquivo com uma ocorrência a mais acrescenta só ela.
    base = (df["pedido"] + "|" + df["entrega"] + "|" + df["tipo"] + "|" + df["motivo"] + "|"
            + df["sku"] + "|" + df["parcela"] + "|" + df["liquidacao"] + "|" + df["operacao"] + "|"
            + df["data_caixa"] + "|" + df["repasse_liquido"].map(lambda v: f"{v:.2f}"))
    df["chave"] = base + "|" + base.groupby(base).cumcount().add(1).astype(str)
    df["data"] = df["data_caixa"]        # a base do PLUTOS ordena por "data"

    canc = set(df.loc[df["tipo"] == T_CANCEL, "pedido"])
    fora = df[df["ajuste_fora"]]
    diag = {
        "linhas": len(df), "pedidos": int(df["pedido"].nunique()),
        "de": (df["data_caixa"].min() if len(df) else ""), "ate": (df["data_caixa"].max() if len(df) else ""),
        "forma": {k: int(v) for k, v in df["forma"].value_counts().items()},
        "tipos": {k: int(v) for k, v in df["tipo"].value_counts().items()},
        "rebate": round(float(df["rebate"].sum()), 2),
        "reb_onus": round(float(df["reb_onus"].sum()), 2),
        "reb_campanha": round(float(df["reb_campanha"].sum()), 2),
        "cancelados": len(canc),
        "venda_base": round(float(df.loc[venda, "transacao"].sum()), 2),
        "venda_comissao": round(float(df.loc[venda, "comissao_rs"].sum()), 2),
        "venda_linhas": int(venda.sum()),
        "tabela_fora": int(df["tabela_fora"].sum()),
        "tabela_fora_rs": round(float(df.loc[df["tabela_fora"], "comissao_rs"].sum()), 2),
        "tabela_pcts": {f"{k:.2f}": int(v) for k, v in
                        df.loc[venda, "pct_contratual"].round(2).value_counts().sort_index().items()},
        "ajustes_fora": len(fora),
        "ajustes_fora_rs": round(float(fora["repasse_liquido"].sum()), 2),
        "ajustes_fora_motivos": {k: int(v) for k, v in fora["motivo"].value_counts().items()},
        "estornos": int(df["estorno"].sum()),
        "estornos_com_cancel": int(df["estorno_com_cancel"].sum()),
        "estornos_rs": round(float(df.loc[df["estorno"], "repasse_liquido"].sum()), 2),
        "motivos_novos": {k: int(v) for k, v in df.loc[df["motivo_novo"], "motivo"].value_counts().items()},
        "motivos_novos_rs": round(float(df.loc[df["motivo_novo"], "repasse_liquido"].sum()), 2),
        "motivos_todos": {k: int(v) for k, v in df.loc[df["motivo"] != "", "motivo"].value_counts().items()},
        "competencias": {k: int(v) for k, v in df["competencia"].value_counts().items()},
        "pedido_de": (df.loc[df["data_pedido"] != "", "data_pedido"].min() if (df["data_pedido"] != "").any() else ""),
        "pedido_ate": (df.loc[df["data_pedido"] != "", "data_pedido"].max() if (df["data_pedido"] != "").any() else ""),
    }
    return df, diag


def comissao_vs_contrato(base: float, cobrada: float, pct_contrato: float) -> dict:
    """A régua do contrato (13%), aplicada a UM pedido.

    deveria = base × 13%. A partir daí, duas coisas que não se misturam:
      · cobrada ABAIXO do contrato → a diferença é REBATE DE COMISSÃO;
      · cobrada ACIMA do contrato  → a diferença é COBRANÇA A MAIOR, vai para
        a lista de ajuste com o canal (não abate do rebate).
    """
    deveria = round(base * pct_contrato, 2)
    dif = round(deveria - cobrada, 2)
    return {
        "com_contrato": deveria,
        "com_cobrada": round(cobrada, 2),
        "pct_cobrado": (round(100 * cobrada / base, 4) if base else 0.0),
        "rebate_comissao": (dif if dif > TOL_RS else 0.0),
        "cobrou_mais_rs": (round(-dif, 2) if dif < -TOL_RS else 0.0),
        "cobrou_mais": bool(dif < -TOL_RS),
    }


def por_pedido(df: pd.DataFrame, comp: str, erp_idx: dict | None = None,
               pct_cad: float = 0.13) -> list[dict]:
    """Uma linha por PEDIDO com repasse no mês `comp` (regime de caixa).

    `erp_idx` = {OC: linha do ERP} — serve para saber a data do pedido (e com
    ela separar o que é do mês do que é saldo de mês anterior) e para comparar
    o % do Promob com o que o canal aplicou."""
    erp_idx = erp_idx or {}
    if df is None or not len(df):
        return []
    d = df[df["competencia"] == comp]
    if not len(d):
        return []
    out = []
    for ped, g in d.groupby("pedido", sort=False):
        e = erp_idx.get(ped) or {}
        v = g[g["tipo"] == T_VENDA]
        cancelado = bool((g["tipo"] == T_CANCEL).any())
        rebate = round(float(g["rebate"].sum()), 2)
        comp_ped = (e.get("competencia") or
                    (g.loc[g["comp_pedido"] != "", "comp_pedido"].min() if (g["comp_pedido"] != "").any() else ""))
        base = round(float(v["transacao"].sum()), 2)
        com_canal = round(float(v["comissao_rs"].sum()), 2)
        pct_canal = round(float(v["pct_aplicada"].mean()), 4) if len(v) else 0.0
        pct_erp = float(e.get("pct_comissao") or 0.0) or None
        # A RÉGUA É O CONTRATO (13%, do cadastro de Parâmetros) — não o % que a
        # Via imprime como "contratual por categoria" no relatório dela.
        cc = comissao_vs_contrato(base, com_canal, pct_cad)
        # PEDIDO CANCELADO FICA FORA DAS DUAS PONTAS. Não gera rebate (regra da
        # casa) e também não gera cobrança a maior: no cancelamento o canal já
        # devolveu a comissão inteira, então pedir ajuste dela seria cobrar duas
        # vezes — e mandaria a Gabi reclamar de um pedido que nem existe mais.
        reb_com = (0.0 if cancelado else cc["rebate_comissao"])
        cob_mais_rs = (0.0 if cancelado else cc["cobrou_mais_rs"])
        cob_mais = (False if cancelado else cc["cobrou_mais"])
        com_promob = cc["com_contrato"]
        out.append({
            "canal": CANAL, "pedido_mkt": ped, "pedido_canal": ped,
            "pedido_any": e.get("obs05", "") or "", "id_mkt": g["entrega"].iloc[0],
            "data": g["data_caixa"].max(), "competencia": comp,
            "data_pedido": (e.get("data") or (g["data_pedido"].min() if (g["data_pedido"] != "").any() else "")),
            "comp_pedido": comp_ped,
            # SALDO = o dinheiro é deste mês, mas a venda é de mês anterior
            # (ou nem está na nossa base). Soma no mês, separado no box.
            "saldo_anterior": bool(comp_ped and comp_ped < comp) or not comp_ped,
            "erp_ok": bool(e), "conta": g["forma"].iloc[0],
            "forma": g["forma"].iloc[0], "meio_pgto": g["meio_pgto"].iloc[0],
            "status": ("Cancelado" if cancelado else "Repassado"),
            "sku": g["sku"].iloc[0], "anuncio": "", "produto": g["produto"].iloc[0],
            "categoria": g["categoria"].iloc[0], "tipo": g["categoria"].iloc[0],
            "itens": int(len(v)), "qtd": 1.0,
            "valor_prod": round(float(v["produto_rs"].sum()), 2),
            "frete": round(float(v["frete"].sum()), 2),
            "sis_base": base, "base_amazon": base,
            "tarifa": com_canal, "pct_comissao": round(pct_canal / 100, 6),
            "sis_pct": pct_cad, "sis_rs": com_promob,
            "pct_canal": pct_canal, "pct_promob": round(pct_cad * 100, 2),
            "pct_cobrado": cc["pct_cobrado"], "pct_contrato": round(pct_cad * 100, 2),
            "pct_erp": (round(pct_erp * 100, 2) if pct_erp else None),
            "com_contrato": cc["com_contrato"], "com_cobrada": cc["com_cobrada"],
            "cadastro_difere": bool(abs(cc["com_contrato"] - com_canal) > TOL_RS),
            "cadastro_dif_rs": round(com_promob - com_canal, 2),
            "cobrou_mais": cob_mais,
            "cobrou_mais_rs": cob_mais_rs,
            "repasse": round(float(g["repasse_liquido"].sum()), 2),
            "cancelado": cancelado,
            "reb_onus": round(float(g["reb_onus"].sum()), 2),
            "reb_campanha": round(float(g["reb_campanha"].sum()), 2),
            # cancelado não gera rebate — regra da casa
            "rebate_rs": (0.0 if cancelado else rebate),
            "rebate_bruto": rebate,
            "rebate_comissao": reb_com, "rebate_frete": 0.0,
            "rebate_total": (0.0 if cancelado else round(rebate + reb_com, 2)),
            "cupom_meli": 0.0, "cupom_seller": 0.0, "faltante": 0.0, "faltante_status": "",
            "tarifa_zero": False, "sem_sistema": not bool(e), "diferenca": 0.0,
        })
    out.sort(key=lambda x: -x["rebate_total"])
    return out


def resumo(linhas: list[dict], pendentes: list[dict] | None = None, extra: dict | None = None) -> dict:
    pendentes = pendentes or []
    extra = extra or {}
    s = lambda ls, k: round(sum(x.get(k) or 0.0 for x in ls), 2)  # noqa: E731
    do_mes = [l for l in linhas if not l["saldo_anterior"] and not l["cancelado"]]
    saldo = [l for l in linhas if l["saldo_anterior"] and not l["cancelado"]]
    canc = [l for l in linhas if l["cancelado"]]
    cad = [l for l in linhas if l["cadastro_difere"]]
    cob = [l for l in linhas if l["cobrou_mais"]]
    rcom = [l for l in linhas if (l.get("rebate_comissao") or 0.0) > 0]
    vivos = [l for l in linhas if not l["cancelado"]]
    venda = s(linhas, "valor_prod")
    return {
        "pedidos": len(linhas), "venda": venda,
        "rebate_mes": s(do_mes, "rebate_total"),
        "rebate_saldo": s(saldo, "rebate_total"),
        "rebate_total": s(linhas, "rebate_total"),
        "rebate_rs": s(linhas, "rebate_total"),
        "rebate_comissao": s(linhas, "rebate_comissao"), "rebate_frete": 0.0,
        "ped_rebate_comissao": len(rcom),
        "reb_repasse": round(s(linhas, "rebate_total") - s(linhas, "rebate_comissao"), 2),
        "ped_mes": len(do_mes), "ped_saldo": len(saldo),
        "reb_onus": s(linhas, "reb_onus"), "reb_campanha": s(linhas, "reb_campanha"),
        "cancelados": len(canc), "cancelados_rs": s(canc, "rebate_bruto"),
        "repasse": s(linhas, "repasse"),
        "com_real": s(linhas, "tarifa"), "com_sistema": s(linhas, "sis_rs"),
        # A régua só soma os pedidos que estão NELA: cancelado fica de fora das
        # duas pontas, senão a tela mostra "contrato − cobrada" sem bater com
        # "rebate − cobrança indevida" e a conta não fecha para quem confere.
        "com_contrato": s(vivos, "com_contrato"), "com_cobrada": s(vivos, "com_cobrada"),
        "com_base": s(vivos, "sis_base"),
        "canc_com_contrato": s(canc, "com_contrato"), "canc_com_cobrada": s(canc, "com_cobrada"),
        "cadastro_difere": len(cad), "cadastro_dif_rs": s(cad, "cadastro_dif_rs"),
        "cobrou_mais": len(cob), "cobrou_mais_rs": s(cob, "cobrou_mais_rs"),
        "pendentes": len(pendentes), "pendentes_rs": round(sum(p.get("a_receber") or 0.0 for p in pendentes), 2),
        "pct_sobre_venda": (round(100 * s(linhas, "rebate_total") / venda, 2) if venda else 0.0),
        "por_forma": {f: {"pedidos": len([l for l in linhas if l["forma"] == f]),
                          "rebate": s([l for l in linhas if l["forma"] == f], "rebate_total")}
                      for f in sorted({l["forma"] for l in linhas})},
        **extra,
    }
