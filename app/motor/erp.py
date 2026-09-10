# -*- coding: utf-8 -*-
"""
PLUTOS · ERP · PEDIDOS MARKETPLACE (Promob)
============================================
O export mensal do ERP com as vendas de TODOS os canais ("Pedidos_Marketplace_
DDMMateDDMMAAAA_HHMM"), em .csv (tab, latin-1) ou .xlsx.

Regras (Thaís + Gabi, 10/09/2026):
  • a OC (coluna D "Ordem de compra") é SEMPRE texto — é a chave com os canais
  • a coluna T "Representante - Fantasia" diz o canal: tira-se o "MP - " da
    frente para chegar no nome do canal e cair no box certo
  • a coluna AB "Percentual de comissão 01" é o % de comissão do ERP — sempre
    em %, e é o valor certo. Comissão do sistema (R$) = valor de produtos do
    canal × AB%  (validado: bate 100% com o PROMOB R$ da planilha da Gabi)
"""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd

# nome normalizado (só letras) → chave interna
COLS = {
    "pedidonumero": "pedido_erp", "naturezacodigo": "natureza", "situacaostatus": "status",
    "ordemdecompra": "oc", "registros": "registros", "nrovolumes": "volumes", "pesobruto": "peso",
    "datadeemissao": "data", "horadadigitacao": "hora", "valordosprodutos": "valor_prod",
    "valordofrete": "valor_frete", "valortotalpedido": "valor_total", "clientecidade": "cidade",
    "clienteestadosigla": "uf", "clientecodigo": "cliente", "observacao": "obs05",
    "representantefantasia": "canal_erp", "representanterazaosocialrepresentante": "canal_razao",
    "notafiscalnumero": "nf", "seriedanotafiscal": "nf_serie", "datadeemissaodanotafiscal": "data_nf",
    "reservadoflag": "reservado", "redespachofantasia": "redespacho",
    "percentualdecomissao": "pct_comissao", "pedimppedimpvlipi": "ipi",
    "transportadoraderedespachorazaosocialtranspostadora": "transportadora",
    "seriepedido": "serie", "enderecodeentregacep": "cep", "bloqueiofaturamento": "bloqueio",
}
OBRIGATORIAS = ["oc", "data", "canal_erp", "pct_comissao", "valor_prod"]

# "Representante - Fantasia" → box do PLUTOS (depois de tirar o "MP - ")
CANAL_BOX = {
    "MERCADO LIVRE": "meli", "MERCADO LIVRE - FULL": "meli",
    "MAGAZINE LUIZA": "magalu", "MADEIRA MADEIRA": "madeira", "COLOMBO": "colombo",
    "CASAS BAHIA": "cbahia", "AMAZON.COM.BR": "amazon", "AMAZON": "amazon",
    "WEBCONTINENTAL": "webcont", "SHOPEE": "shopee", "SHOPEE XPRESS": "shopee",
}


def _norm(s) -> str:
    s = "" if s is None else str(s)
    s = s.lower()
    # o export vem em latin-1 e às vezes com acentos quebrados (N·mero, C¾digo, SituaþÒo)
    s = re.sub(r"[áàâãäªº·]", "a", s); s = re.sub(r"[éèêë]", "e", s); s = re.sub(r"[íìîï]", "i", s)
    s = re.sub(r"[óòôõö¾]", "o", s); s = re.sub(r"[úùûü]", "u", s); s = s.replace("ç", "c").replace("þ", "c")
    return re.sub(r"[^a-z]", "", s)


def _num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _txt(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s


def _data(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    s = str(v).strip()[:10]
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def canal_de(fantasia: str, mapa: dict | None = None) -> tuple[str, str]:
    """('MP - MAGAZINE LUIZA') → ('MAGAZINE LUIZA', 'magalu'); sem box → chave 'outros'."""
    nome = re.sub(r"^\s*MP\s*-\s*", "", str(fantasia or "").strip(), flags=re.I).strip().upper()
    return nome, (mapa or CANAL_BOX).get(nome, "outros")


def ler(caminho: str, mapa: dict | None = None) -> tuple[list[dict], dict]:
    """Lê o export (csv tab latin-1 ou xlsx). Devolve (linhas normalizadas, diagnóstico)."""
    if caminho.lower().endswith((".csv", ".txt")):
        df = None
        for enc in ("latin-1", "cp1252", "utf-8-sig"):
            try:
                df = pd.read_csv(caminho, sep="\t", encoding=enc, dtype=str, keep_default_na=False)
                if df.shape[1] < 5:
                    df = pd.read_csv(caminho, sep=";", encoding=enc, dtype=str, keep_default_na=False)
                break
            except Exception:  # noqa: BLE001
                continue
        if df is None:
            raise ValueError("Não consegui ler o CSV (esperado: separado por TAB, latin-1).")
    else:
        xl = pd.ExcelFile(caminho)
        df = xl.parse(xl.sheet_names[0], dtype=str)
    colmap = {}
    for c in df.columns:
        n = _norm(c)
        for k, v in COLS.items():
            if n.startswith(k) and v not in colmap.values():
                colmap[c] = v
                break
    faltam = [k for k in OBRIGATORIAS if k not in colmap.values()]
    if faltam:
        raise ValueError(f"O arquivo não tem as colunas do ERP: faltam {', '.join(faltam)} "
                         f"(li {df.shape[1]} colunas). O .xlsx do ERP às vezes sai incompleto — use o .csv.")
    df = df.rename(columns=colmap)

    linhas, diag = [], {"linhas_brutas": int(len(df)), "rejeitadas": 0, "canais": {}, "sem_box": {},
                        "status": {}, "competencias": {}}
    for r in df.to_dict("records"):
        oc = _txt(r.get("oc"))
        d = _data(r.get("data"))
        if not oc or not d:
            diag["rejeitadas"] += 1
            continue
        nome, box = canal_de(r.get("canal_erp"), mapa)
        comp = f"{d.year}-{d.month:02d}"
        lin = {
            "oc": oc, "pedido_erp": _txt(r.get("pedido_erp")), "natureza": _txt(r.get("natureza")),
            "status": _txt(r.get("status")), "data": d.isoformat(), "competencia": comp,
            "hora": _txt(r.get("hora")), "volumes": _num(r.get("volumes")), "peso": _num(r.get("peso")),
            "valor_prod": round(_num(r.get("valor_prod")), 2), "valor_frete": round(_num(r.get("valor_frete")), 2),
            "valor_total": round(_num(r.get("valor_total")), 2),
            "cidade": _txt(r.get("cidade")), "uf": _txt(r.get("uf")), "cliente": _txt(r.get("cliente")),
            "obs05": _txt(r.get("obs05")), "canal_erp": _txt(r.get("canal_erp")), "canal_nome": nome, "box": box,
            "nf": _txt(r.get("nf")), "nf_serie": _txt(r.get("nf_serie")),
            "data_nf": (_data(r.get("data_nf")).isoformat() if _data(r.get("data_nf")) else ""),
            "pct_comissao": round(_num(r.get("pct_comissao")) / 100.0, 6),
            "ipi": round(_num(r.get("ipi")), 2), "transportadora": _txt(r.get("transportadora")),
            "redespacho": _txt(r.get("redespacho")), "cep": _txt(r.get("cep")), "bloqueio": _txt(r.get("bloqueio")),
        }
        lin["comissao_erp_rs"] = round(lin["valor_prod"] * lin["pct_comissao"], 2)
        linhas.append(lin)
        diag["canais"][nome] = diag["canais"].get(nome, 0) + 1
        if box == "outros":
            diag["sem_box"][nome] = diag["sem_box"].get(nome, 0) + 1
        diag["status"][lin["status"]] = diag["status"].get(lin["status"], 0) + 1
        diag["competencias"][comp] = diag["competencias"].get(comp, 0) + 1
    if not linhas:
        raise ValueError("Nenhuma linha válida (OC + data) no arquivo.")
    datas = sorted(l["data"] for l in linhas)
    diag["de"], diag["ate"], diag["linhas"] = datas[0], datas[-1], len(linhas)
    diag["ocs_duplicadas"] = len(linhas) - len({l["oc"] for l in linhas})
    return linhas, diag


def resumo_por_box(linhas: list[dict]) -> dict:
    out: dict[str, dict] = {}
    for l in linhas:
        b = out.setdefault(l["box"], {"pedidos": 0, "venda": 0.0, "frete": 0.0, "comissao_erp": 0.0, "nomes": set()})
        b["pedidos"] += 1; b["venda"] += l["valor_prod"]; b["frete"] += l["valor_frete"]
        b["comissao_erp"] += l["comissao_erp_rs"]; b["nomes"].add(l["canal_nome"])
    for b in out.values():
        for k in ("venda", "frete", "comissao_erp"):
            b[k] = round(b[k], 2)
        b["nomes"] = sorted(b["nomes"])
    return out
