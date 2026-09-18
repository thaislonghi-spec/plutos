"""Box MAGAZINE LUIZA · COBRANÇAS DO FULFILLMENT (regra da Thaís, 16/09/2026).

Entrada: os zips do portal Magalu
  · CSV-156171_CobrançasFullFilment_<periodo>.zip
  · CSV-160172_ColetaFullfilment_<periodo>.zip
que juntos trazem quatro csv (";" com número brasileiro):

  produtos_manuseados.csv   CD · sku · quantidade · valor unitário      → MANUSEIO
  produtos_armazenados.csv  CD · data · sku · quantidade · valor unit.  → ARMAZENAGEM
  tempo_estoque.csv         CD · data entrada · aniversário ·
                            "Descrição / SKU" · quantidade · valor unit. → TEMPO DE ESTOQUE
  produtos_coletados.csv    data · agenda · espaço (m³) · valor          → COLETA

Os três primeiros têm SKU e viram custo por item. A COLETA é cobrada por
agenda/m³ e NÃO tem SKU — ela aparece separada e, na tela, também rateada pela
participação de cada SKU no manuseio, sempre marcada como rateio.

Custo total do Full para a Multimóveis = manuseio + armazenagem + tempo de
estoque + coleta + coparticipação de frete (esta vem da Planilha 2 · Vendas).
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from datetime import datetime
from typing import Any

TIPOS = {"manuseio": "Manuseio", "armazenagem": "Armazenagem",
         "tempo_estoque": "Tempo de estoque", "coleta": "Coleta",
         "copart": "Coparticipação de frete"}
# a coparticipação medida (arquivo coparticipacao_a_pagar) entra no custo do Full
# no lugar da estimativa que vinha da Planilha 2 · Vendas
TIPOS_CUSTO = ("manuseio", "armazenagem", "tempo_estoque", "coleta")


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
    t = str(v or "").strip()[:10]
    for f in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def _ler(caminho: str):
    with open(caminho, "rb") as f:
        raw = f.read()
    txt = raw.decode("utf-8-sig") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("utf-8", "replace")
    amostra = txt[:3000]
    sep = ";" if amostra.count(";") >= amostra.count(",") else ","
    r = csv.reader(io.StringIO(txt), delimiter=sep)
    cab = [_norm(c) for c in next(r)]
    return cab, r


def que_arquivo(caminho: str) -> str:
    """manuseio · armazenagem · tempo_estoque · coleta · '' se não for do Full."""
    try:
        cab, _ = _ler(caminho)
    except Exception:  # noqa: BLE001
        return ""
    if "valor coparticipacao" in cab and "sku" in cab:
        return "copart"
    if "agenda" in cab and "espaco" in cab:
        return "coleta"
    if "data entrada" in cab and "aniversario" in cab:
        return "tempo_estoque"
    if "sku" in cab and "data" in cab and "valor unitario" in cab:
        return "armazenagem"
    if "sku" in cab and "valor unitario" in cab:
        return "manuseio"
    return ""


def _pega(cab, l, nome, padrao=""):
    try:
        return l[cab.index(nome)]
    except (ValueError, IndexError):
        return padrao


def ler(caminho: str) -> tuple[str, list[dict], dict]:
    """Devolve (tipo, linhas, diagnóstico). Uma linha por cobrança."""
    qual = que_arquivo(caminho)
    if not qual:
        raise ValueError("Não parece uma cobrança do Fulfillment do Magalu (manuseio, armazenagem, "
                         "tempo de estoque ou coleta).")
    cab, it = _ler(caminho)
    out = []
    for n, l in enumerate(it, 1):
        if not any((x or "").strip() for x in l):
            continue
        if qual == "coleta":
            d = _data(_pega(cab, l, "data"))
            v = _num(_pega(cab, l, "valor"))
            out.append({"tipo": qual, "cd": "", "sku": "", "produto": "", "data": str(d) if d else "",
                        "agenda": _pega(cab, l, "agenda").strip(), "espaco": _num(_pega(cab, l, "espaco")),
                        "qtd": 1.0, "unit": v, "valor": round(v, 2), "chave": f"coleta|{_pega(cab, l, 'agenda').strip()}|{d}"})
            continue
        if qual == "copart":
            # 1 linha por pedido × SKU, com o valor que NÓS pagamos do frete
            d = _data(_pega(cab, l, "data do pedido"))
            v = _num(_pega(cab, l, "valor coparticipacao"))
            sku = _pega(cab, l, "sku").strip()
            ped = _pega(cab, l, "pedido").strip()
            out.append({"tipo": qual, "cd": "", "sku": sku, "produto": "",
                        "data": str(d) if d else "", "agenda": "", "espaco": 0.0,
                        "pedido": ped, "servico": _pega(cab, l, "servico").strip(),
                        "peso_cubado": _num(_pega(cab, l, "peso cubado(kg/m3)")),
                        "peso": _num(_pega(cab, l, "peso(kg)")),
                        "qtd": 1.0, "unit": v, "valor": round(v, 2),
                        "chave": f"copart|{ped}|{sku}|{n}"})
            continue
        if qual == "tempo_estoque":
            bruto = _pega(cab, l, "produto/sku")
            sku = bruto.rsplit("/", 1)[-1].strip() if "/" in bruto else bruto.strip()
            produto = bruto.rsplit("/", 1)[0].strip() if "/" in bruto else ""
            d = _data(_pega(cab, l, "data entrada"))
        else:
            sku = _pega(cab, l, "sku").strip()
            produto = ""
            d = _data(_pega(cab, l, "data")) if qual == "armazenagem" else None
        q = _num(_pega(cab, l, "quantidade")) or 1.0
        u = _num(_pega(cab, l, "valor unitario"))
        out.append({"tipo": qual, "cd": _pega(cab, l, "cd").strip(), "sku": sku, "produto": produto[:90],
                    "data": str(d) if d else "", "agenda": "", "espaco": 0.0,
                    "qtd": q, "unit": u, "valor": round(q * u, 2),
                    "aniversario": _pega(cab, l, "aniversario").strip() if qual == "tempo_estoque" else "",
                    "chave": f"{qual}|{_pega(cab, l, 'cd').strip()}|{sku}|{d}|{n}"})
    datas = [x["data"] for x in out if x["data"]]
    diag = {"tipo": qual, "linhas": len(out), "valor": round(sum(x["valor"] for x in out), 2),
            "skus": len({x["sku"] for x in out if x["sku"]}),
            "com_valor": sum(1 for x in out if x["valor"] > 0),
            "de": min(datas, default=""), "ate": max(datas, default="")}
    return qual, out, diag


def resumo(cobrancas: list[dict], copart_por_sku: dict | None = None, nomes: dict | None = None,
           vendidos: dict | None = None, cubagens: dict | None = None, de: str = "", ate: str = "") -> dict:
    """Custo total do Full por SKU. copart_por_sku = {sku: R$} da Planilha 2;
    nomes = {sku: descrição}; vendidos = {sku: unidades vendidas no Full}.

    CUSTO FULL (cobranças) = manuseio + armazenagem + tempo de estoque + coleta rateada
    CUSTO TOTAL            = custo Full + coparticipação de frete
    POR UNIDADE            = custo total ÷ unidades (manuseadas; se o SKU não teve
                             manuseio no ciclo, cai para as unidades vendidas)."""
    copart_por_sku = copart_por_sku or {}
    nomes = nomes or {}
    vendidos = vendidos or {}
    cubagens = cubagens or {}
    cb = [c for c in cobrancas if (not de or not c["data"] or de <= c["data"] <= ate)]
    por_tipo = {k: round(sum(c["valor"] for c in cb if c["tipo"] == k), 2) for k in TIPOS}
    coleta = por_tipo["coleta"]
    # coparticipação MEDIDA (arquivo do canal) manda na estimada (Planilha 2)
    medida: dict[str, float] = {}
    for c in cb:
        if c["tipo"] == "copart" and c["sku"]:
            medida[c["sku"]] = round(medida.get(c["sku"], 0.0) + c["valor"], 2)
    if medida:
        copart_por_sku = medida
    com_sku = [c for c in cb if c["sku"] and c["tipo"] in TIPOS_CUSTO]
    por_sku: dict[str, dict] = {}
    for c in com_sku:
        s = por_sku.setdefault(c["sku"], {"sku": c["sku"], "produto": c["produto"], "qtd": 0.0,
                                          "manuseio": 0.0, "armazenagem": 0.0, "tempo_estoque": 0.0,
                                          "copart": 0.0, "coleta_rateio": 0.0, "cds": [], "ultimo": "",
                                          "cobrancas": 0, "estocado": 0.0, "unit_manuseio": 0.0})
        if c["produto"] and not s["produto"]:
            s["produto"] = c["produto"]
        if c["cd"] and c["cd"] not in s["cds"]:
            s["cds"].append(c["cd"])
        if c["data"] > s["ultimo"]:
            s["ultimo"] = c["data"]
        s["cobrancas"] += 1
        if c["tipo"] == "manuseio":
            s["qtd"] += c["qtd"]
            if c["unit"]:                    # tabela vigente: vale o último arquivo subido
                s["unit_manuseio"] = round(c["unit"], 2)
        if c["tipo"] == "armazenagem":
            s["estocado"] += c["qtd"]
        s[c["tipo"]] = round(s[c["tipo"]] + c["valor"], 2)
    for sku, v in copart_por_sku.items():
        s = por_sku.setdefault(sku, {"sku": sku, "produto": "", "qtd": 0.0, "manuseio": 0.0,
                                     "armazenagem": 0.0, "tempo_estoque": 0.0, "copart": 0.0,
                                     "coleta_rateio": 0.0, "cds": [], "ultimo": "", "cobrancas": 0,
                                     "estocado": 0.0, "unit_manuseio": 0.0})
        s["copart"] = round(v, 2)
    # A COLETA é cobrada por VOLUME (m³ da agenda) — conferido em set/26:
    # R$ 28.852,67 em 169,72 m³ = R$ 170,00/m³ exatos. Então o rateio dela é
    # pelo volume que cada SKU ocupou (cubagem × unidades), não pelo manuseio.
    # Sem cubagem no cadastro, cai para o manuseio (melhor que nada).
    m3_agendas = round(sum(c.get("espaco") or 0.0 for c in cb if c["tipo"] == "coleta"), 4)
    tarifa_m3 = round(coleta / m3_agendas, 2) if m3_agendas else 0.0
    for s in por_sku.values():
        s["cubagem"] = float(cubagens.get(s["sku"], 0.0) or 0.0)
        s["volume"] = round(s["cubagem"] * (s["qtd"] or 0.0), 4)
    base_vol = sum(s["volume"] for s in por_sku.values())
    base_rateio = sum(s["manuseio"] for s in por_sku.values())
    for s in por_sku.values():
        if not s["produto"]:
            s["produto"] = (nomes.get(s["sku"]) or "")[:90]
        if base_vol and s["volume"]:
            s["coleta_rateio"] = round(coleta * s["volume"] / base_vol, 2)
            s["coleta_base"] = "volume"
        else:
            s["coleta_rateio"] = round(coleta * s["manuseio"] / base_rateio, 2) if base_rateio else 0.0
            s["coleta_base"] = "manuseio" if s["manuseio"] else ""
        # custo de coleta pela TABELA: a tarifa do m³ vezes a cubagem da peça.
        # É o custo real de coletar UMA unidade — o rateio acima divide o que
        # foi cobrado no ciclo, este diz quanto custa a próxima peça.
        s["coleta_tabela_un"] = round(tarifa_m3 * s["cubagem"], 2)
        # CUSTO FULL = tudo que o Fulfillment cobra (sem o frete que nós pagamos)
        s["custo_full"] = round(s["manuseio"] + s["armazenagem"] + s["tempo_estoque"] + s["coleta_rateio"], 2)
        s["custo_sku"] = round(s["custo_full"] - s["coleta_rateio"] + s["copart"], 2)   # sem rateio, para negociação
        s["custo_total"] = round(s["custo_full"] + s["copart"], 2)
        s["vendidas"] = round(float(vendidos.get(s["sku"], 0.0)), 0)
        s["unidades"] = s["qtd"] or s["vendidas"]
        s["base_un"] = "manuseadas" if s["qtd"] else ("vendidas" if s["vendidas"] else "")
        u = s["unidades"]
        s["full_un"] = round(s["custo_full"] / u, 2) if u else 0.0
        s["copart_un"] = round(s["copart"] / u, 2) if u else 0.0
        s["por_unidade"] = round(s["custo_total"] / u, 2) if u else 0.0
    copart = round(sum(copart_por_sku.values()), 2)
    total = round(sum(por_tipo[k] for k in TIPOS_CUSTO) + copart, 2)
    sem_un = [s["sku"] for s in por_sku.values() if not s["unidades"] and s["custo_total"]]
    return {
        "linhas": len(cb), "skus": len(por_sku),
        "manuseio": por_tipo["manuseio"], "armazenagem": por_tipo["armazenagem"],
        "tempo_estoque": por_tipo["tempo_estoque"], "coleta": coleta, "copart": copart,
        "total": total,
        "cobrancas": round(sum(por_tipo.values()), 2),
        "custo_full": round(sum(por_tipo[k] for k in TIPOS_CUSTO), 2),
        "copart_medida": bool(medida),
        "m3_agendas": m3_agendas, "tarifa_m3": tarifa_m3,
        "volume": round(base_vol, 2),
        "sem_cubagem": [s["sku"] for s in por_sku.values() if not s["cubagem"]],
        "sem_unidade": sem_un,
        "sem_unidade_rs": round(sum(s["custo_total"] for s in por_sku.values() if not s["unidades"]), 2),
        "por_sku": sorted(por_sku.values(), key=lambda s: -s["custo_total"]),
        "cds": sorted({c["cd"] for c in cb if c["cd"]}),
        "de": min((c["data"] for c in cb if c["data"]), default=""),
        "ate": max((c["data"] for c in cb if c["data"]), default=""),
    }


def estoque(cobrancas: list[dict], vendidos: dict | None = None, dias_venda: int = 0,
            cadastro: dict | None = None, dias_alvo: int = 30, tarifa_m3: float = 0.0,
            unit_manuseio: dict | None = None, prazos: dict | None = None,
            prazo_padrao: int = 0) -> dict:
    """ESTOQUE FULL — o arquivo de armazenagem é uma FOTO DIÁRIA do estoque no CD
    do Magalu (CD · data · SKU · quantidade). Daqui sai:

      · a evolução do estoque no ciclo (quantas unidades por dia);
      · a última foto por SKU e por CD — o estoque de hoje;
      · a cobertura em dias (estoque ÷ venda média diária no Full);
      · o que já pagou ANIVERSÁRIO (tempo de estoque) — produto parado.

    vendidos = {sku: unidades vendidas no Full} · dias_venda = dias do período
    cadastro = {sku: {custo, peso, cubagem, descricao}} para valorizar o estoque.
    """
    vendidos = vendidos or {}
    cadastro = cadastro or {}
    unit_manuseio = unit_manuseio or {}
    prazos = prazos or {}
    arm = [c for c in cobrancas if c["tipo"] == "armazenagem" and c["data"]]
    if not arm:
        return {"tem": False, "por_dia": [], "por_sku": [], "foto": "", "unidades": 0}
    por_dia: dict[str, dict] = {}
    for c in arm:
        d = por_dia.setdefault(c["data"], {"dia": c["data"], "unidades": 0.0, "skus": set(), "custo": 0.0})
        d["unidades"] += c["qtd"]; d["skus"].add(c["sku"]); d["custo"] += c["valor"]
    dias = sorted(por_dia.values(), key=lambda x: x["dia"])
    for d in dias:
        d["skus"] = len(d["skus"]); d["unidades"] = round(d["unidades"], 0); d["custo"] = round(d["custo"], 2)
    foto = dias[-1]["dia"]
    aniver: dict[str, dict] = {}
    for c in cobrancas:
        if c["tipo"] != "tempo_estoque":
            continue
        a = aniver.setdefault(c["sku"], {"valor": 0.0, "qtd": 0.0, "entrada": c["data"], "aniversario": c.get("aniversario", "")})
        a["valor"] += c["valor"]; a["qtd"] += c["qtd"]
        if c["data"] and c["data"] < a["entrada"]:
            a["entrada"] = c["data"]
    por_sku: dict[str, dict] = {}
    for c in arm:
        s = por_sku.setdefault(c["sku"], {"sku": c["sku"], "estoque": 0.0, "cds": {}, "custo_arm": 0.0,
                                          "dias_no_ciclo": 0})
        s["custo_arm"] += c["valor"]
        s["dias_no_ciclo"] += 1
        if c["data"] == foto:
            s["estoque"] += c["qtd"]
            s["cds"][c["cd"]] = round(s["cds"].get(c["cd"], 0.0) + c["qtd"], 0)
    out = []
    for s in por_sku.values():
        e = cadastro.get(s["sku"]) or {}
        v = float(vendidos.get(s["sku"], 0.0))
        a = aniver.get(s["sku"]) or {}
        s["descricao"] = e.get("descricao") or e.get("descricao_curta") or ""
        s["custo_prod"] = float(e.get("custo") or 0.0)
        s["cubagem"] = float(e.get("cubagem") or 0.0)
        s["valor_estoque"] = round(s["estoque"] * s["custo_prod"], 2)
        s["volume"] = round(s["estoque"] * s["cubagem"], 3)
        s["vendidas"] = round(v, 0)
        s["media_dia"] = round(v / dias_venda, 2) if dias_venda else 0.0
        s["cobertura"] = round(s["estoque"] / s["media_dia"], 0) if s["media_dia"] else None
        s["custo_arm"] = round(s["custo_arm"], 2)
        s["aniversario_rs"] = round(a.get("valor", 0.0), 2)
        s["aniversario_qtd"] = round(a.get("qtd", 0.0), 0)
        s["parado"] = bool(s["estoque"] and not v)
        # SUGESTÃO DE ENVIO, POR CD: cada CD tem o seu prazo (manuseio +
        # transferência), então a conta é por CD, não por SKU inteiro.
        #
        # O relatório do Magalu NÃO diz de qual CD saiu cada venda (a coluna
        # "CD de Origem" vem vazia no Fulfillment). Até o canal mandar esse
        # dado, a venda do SKU é RATEADA pela participação de cada CD no
        # estoque dele — está marcado na tela como estimativa.
        cd_principal = max(s["cds"], key=lambda k: s["cds"][k]) if s["cds"] else ""
        s["cd_principal"] = cd_principal
        s["prazo"] = int(prazos.get(cd_principal, prazo_padrao) or prazo_padrao)
        s["dias_cobrir"] = dias_alvo + s["prazo"]
        base_cd = dict(s["cds"]) or ({cd_principal: 0.0} if cd_principal else {})
        tot_cd = sum(base_cd.values())
        linhas_cd = []
        for cdn, est_cd in sorted(base_cd.items(), key=lambda x: -x[1]):
            parte = (est_cd / tot_cd) if tot_cd else (1.0 / len(base_cd))
            pz = int(prazos.get(cdn, prazo_padrao) or prazo_padrao)
            md = round(s["media_dia"] * parte, 3)
            alvo_cd = round(md * (dias_alvo + pz), 0) if md else 0.0
            sug_cd = max(0.0, round(alvo_cd - est_cd, 0))
            cob_cd = round(est_cd / md, 0) if md else None
            linhas_cd.append({"cd": cdn, "estoque": est_cd, "parte": round(100 * parte, 1),
                              "prazo": pz, "dias_cobrir": dias_alvo + pz, "media_dia": md,
                              "vendidas": round(v * parte, 0), "alvo": alvo_cd, "sugestao": sug_cd,
                              "cobertura": cob_cd, "rompe": bool(md and cob_cd is not None and cob_cd < pz),
                              "sug_m3": round(sug_cd * s["cubagem"], 3)})
        s["por_cd"] = linhas_cd
        alvo = round(sum(x["alvo"] for x in linhas_cd), 0)
        s["alvo"] = alvo
        s["sugestao"] = round(sum(x["sugestao"] for x in linhas_cd), 0)
        s["sug_m3"] = round(s["sugestao"] * s["cubagem"], 3)
        s["sug_coleta"] = round(s["sug_m3"] * tarifa_m3, 2)
        s["sug_manuseio"] = round(s["sugestao"] * float(unit_manuseio.get(s["sku"], 0.0) or 0.0), 2)
        s["sug_custo"] = round(s["sug_coleta"] + s["sug_manuseio"], 2)
        s["sug_valor"] = round(s["sugestao"] * s["custo_prod"], 2)
        # rompe antes de a reposição chegar?
        s["rompe"] = any(x["rompe"] for x in linhas_cd)
        s["acao"] = ("RETIRAR" if s["parado"] else
                     ("URGENTE" if s["rompe"] and s["sugestao"] else
                      ("ENVIAR" if s["sugestao"] else
                       ("excesso" if s["cobertura"] and s["cobertura"] > 2 * dias_alvo else "ok"))))
        s["cd"] = " · ".join(f"{k} {int(q)}" for k, q in sorted(s["cds"].items(), key=lambda x: -x[1]))
        out.append(s)
    out.sort(key=lambda s: -s["valor_estoque"])
    unidades = round(sum(s["estoque"] for s in out), 0)
    return {
        "tem": True, "foto": foto, "de": dias[0]["dia"], "dias": len(dias),
        "unidades": unidades, "skus": sum(1 for s in out if s["estoque"]),
        "valor_estoque": round(sum(s["valor_estoque"] for s in out), 2),
        "volume": round(sum(s["volume"] for s in out), 2),
        "custo_arm": round(sum(s["custo_arm"] for s in out), 2),
        "aniversario_rs": round(sum(s["aniversario_rs"] for s in out), 2),
        "dias_alvo": dias_alvo, "prazo_padrao": prazo_padrao,
        "por_cd": sorted(
            [dict(x, sku=s["sku"], descricao=s.get("descricao", ""), cubagem=s["cubagem"],
                  custo_prod=s["custo_prod"], parado=s["parado"],
                  acao=("RETIRAR" if s["parado"] else ("URGENTE" if x["rompe"] and x["sugestao"]
                        else ("ENVIAR" if x["sugestao"] else
                              ("excesso" if x["cobertura"] and x["cobertura"] > 2 * dias_alvo else "ok")))))
             for s in out for x in s.get("por_cd", [])],
            key=lambda x: (x["acao"] != "URGENTE", -x["sugestao"])),
        "urgentes": [s["sku"] for s in out if s["acao"] == "URGENTE"],
        "urgentes_un": round(sum(s["sugestao"] for s in out if s["acao"] == "URGENTE"), 0),
        "enviar_skus": sum(1 for s in out if s["sugestao"]),
        "enviar_un": round(sum(s["sugestao"] for s in out), 0),
        "enviar_m3": round(sum(s["sug_m3"] for s in out), 2),
        "enviar_coleta": round(sum(s["sug_coleta"] for s in out), 2),
        "enviar_manuseio": round(sum(s["sug_manuseio"] for s in out), 2),
        "enviar_custo": round(sum(s["sug_custo"] for s in out), 2),
        "enviar_valor": round(sum(s["sug_valor"] for s in out), 2),
        "excesso": [s["sku"] for s in out if s["acao"] == "excesso"],
        "excesso_rs": round(sum(s["valor_estoque"] for s in out if s["acao"] == "excesso"), 2),
        "parados": [s["sku"] for s in out if s["parado"]],
        "parados_rs": round(sum(s["valor_estoque"] for s in out if s["parado"]), 2),
        "por_dia": dias, "por_sku": out,
        "cds": sorted({c["cd"] for c in arm if c["cd"]}),
    }
