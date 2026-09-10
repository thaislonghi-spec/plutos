# -*- coding: utf-8 -*-
"""Saídas em Excel do PLUTOS."""
from __future__ import annotations

import io
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PRETO, OURO, VERM = "0B0B0C", "D4A017", "C8102E"
CAB = PatternFill("solid", fgColor=PRETO)
CAB_F = Font(bold=True, color="FFFFFF", size=10)
OURO_F = PatternFill("solid", fgColor="FFF3D6")
BRL = '#,##0.00'


def _aba(wb, nome, cab, linhas, larguras=None, moeda=()):
    ws = wb.create_sheet(nome)
    ws.append(cab)
    for c in ws[1]:
        c.fill, c.font, c.alignment = CAB, CAB_F, Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for l in linhas:
        ws.append(l)
    for j, w in enumerate(larguras or [], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for j in moeda:
        for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
            for c in row:
                c.number_format = BRL
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def rebates_xlsx(rods: dict, comp: str, canais: dict, param: dict) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)

    # 1 · por pedido (todos os canais) — a chave é o pedido do marketplace = OC
    cab = ["Canal", "OC / Pedido marketplace", "Pedido canal", "Pedido Any", "Data", "SKU", "Anúncio",
           "Valor produto", "Tarifa cobrada", "Cupom Meli",
           "Rebate R$", "Rebate COMISSÃO", "Rebate FRETE", "REBATE TOTAL",
           "Tarifa zero?", "Faltante campanha", "Status faltante", "Comissão sistema R$", "Diferença"]
    linhas = []
    for chave, r in rods.items():
        nome = canais[chave]["nome"]
        for l in r["linhas"]:
            linhas.append([nome, l["pedido_mkt"], l["pedido_canal"], l["pedido_any"],
                           datetime.fromisoformat(l["data"]), l["sku"], l["anuncio"],
                           l["valor_prod"], l["tarifa"], l["cupom_meli"],
                           l["rebate_rs"], l["rebate_comissao"], l["rebate_frete"], l["rebate_total"],
                           "SIM" if l["tarifa_zero"] else "", l["faltante"], l["faltante_status"],
                           l["sis_rs"], l["diferenca"]])
    ws = _aba(wb, "1_Por_Pedido", cab, linhas,
              [16, 22, 20, 14, 11, 16, 16, 13, 13, 11, 12, 14, 12, 14, 10, 14, 12, 16, 12],
              moeda=(8, 9, 10, 11, 12, 13, 14, 16, 18, 19))
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for c in row:
            c.number_format = "DD/MM/YYYY"

    # 2 · por canal
    cab2 = ["Canal", "Pedidos", "Venda (produto)", "Rebate R$", "Rebate COMISSÃO", "Rebate FRETE",
            "REBATE TOTAL", "% sobre a venda", "Faltante pendente (pedidos)", "Rodada"]
    l2 = []
    for chave, r in rods.items():
        s = r["resumo"]
        l2.append([canais[chave]["nome"], s["pedidos"], s["venda"], s["rebate_rs"], s["rebate_comissao"],
                   0.0, s["rebate_total"], s["pct_sobre_venda"] / 100, s["faltante_pendentes"],
                   datetime.fromisoformat(r["quando"]).strftime("%d/%m/%Y %H:%M")])
    tot = [sum(x[i] for x in l2) for i in (1, 2, 3, 4, 5, 6)]
    l2.append(["TOTAL", *tot, (tot[5] / tot[1]) if tot[1] else 0, sum(x[8] for x in l2), ""])
    ws2 = _aba(wb, "2_Por_Canal", cab2, l2, [18, 10, 16, 14, 16, 14, 16, 14, 22, 18], moeda=(3, 4, 5, 6, 7))
    for row in ws2.iter_rows(min_row=2, min_col=8, max_col=8):
        for c in row:
            c.number_format = "0.00%"
    for c in ws2[ws2.max_row]:
        c.font = Font(bold=True); c.fill = OURO_F

    # 3 · por dia (Meli) — o mesmo resumo da "Planilha1" da Gabi
    if "meli" in rods:
        pd_ = rods["meli"]["resumo"]["por_dia"]
        cab3 = ["Dia", "Pedidos", "Venda", "CUPOM MELI (R$)", "FALTANTE CAMPANHA (R$)",
                "DIFERENÇA COMISSÃO", "SOMA", "Pedidos tarifa zero"]
        l3 = [[datetime.fromisoformat(d), v["pedidos"], v["venda"], v["cupom"], v["faltante"],
               v["comissao"], v["total"], v["tz"]] for d, v in pd_.items()]
        ws3 = _aba(wb, "3_Meli_Por_Dia", cab3, l3, [12, 10, 14, 16, 20, 18, 14, 16], moeda=(3, 4, 5, 6, 7))
        for row in ws3.iter_rows(min_row=2, min_col=1, max_col=1):
            for c in row:
                c.number_format = "DD/MM/YYYY"

    # 4 · como ler
    ws4 = wb.create_sheet("4_Como_ler")
    txt = [
        "PLUTOS · Rebates · competência " + comp,
        "",
        "As TRÊS formas de rebate (definição da direção, 09/09/2026):",
        "  1. REBATE R$        — volta em dinheiro na conta.",
        "  2. REBATE COMISSÃO  — desconto no % de comissão; não volta dinheiro, paga-se menos.",
        "  3. REBATE FRETE     — o canal ajuda a pagar o frete (devolve, não cobra, ou encontro de contas).",
        "",
        "MERCADO LIVRE (método da Gabi, automatizado):",
        "  R$        = Cupom Meli (coluna do export) + Faltante campanha (tabela manual, pedidos de tarifa zero)",
        "  COMISSÃO  = Comissão do sistema (Promob, planilha do ADC002) − Tarifa cobrada pelo Meli",
        f"              (diferenças até R$ {param.get('tolerancia_comissao', 0.5):.2f} são ignoradas — arredondamento)",
        "  FRETE     = não existe no Meli: o frete do Coletas é custo nosso, já dentro do custo do produto.",
        "",
        "TARIFA ZERO = o rebate da campanha foi maior que a comissão: zera a comissão E devolve um valor.",
        "  A comissão inteira do sistema entra como rebate de COMISSÃO; o valor devolvido (pesquisado no",
        "  portal pela Gabi) entra como rebate R$ pela tabela FALTANTE CAMPANHA. Enquanto não preenchido,",
        "  fica como PENDENTE e não soma.",
        "",
        "CHAVE: 'OC / Pedido marketplace' é o mesmo número que o Tropa de Elite usa como OC do ERP.",
        "COMPETÊNCIA: mês da DATA DO PEDIDO (o BI é exportado com filtro Pago, dia 01 a 31).",
    ]
    for t in txt:
        ws4.append([t])
    ws4.column_dimensions["A"].width = 110
    ws4["A1"].font = Font(bold=True, size=13, color=PRETO)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def faltante_xlsx(r, tab: dict, comp: str) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    cab = ["Pedido marketplace (OC)", "Pedido canal", "Data", "SKU", "Anúncio", "Valor produto",
           "Comissão sistema R$", "Valor devolvido (R$)", "Status", "Obs", "Quem", "Quando"]
    linhas = []
    for l in (r or {}).get("linhas", []):
        if not l["tarifa_zero"]:
            continue
        f = tab.get(l["pedido_mkt"]) or {}
        linhas.append([l["pedido_mkt"], l["pedido_canal"], datetime.fromisoformat(l["data"]), l["sku"], l["anuncio"],
                       l["valor_prod"], l["sis_rs"], f.get("valor"), l["faltante_status"], f.get("obs", ""),
                       f.get("quem", ""), f.get("quando", "")])
    ws = _aba(wb, "Faltante_Campanha_Meli", cab, linhas, [24, 20, 11, 16, 16, 13, 16, 18, 12, 30, 10, 20],
              moeda=(6, 7, 8))
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def linha_xlsx(r, cols, nome_canal: str, comp: str) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    cab = [rot for rot, _, _ in cols]
    linhas = []
    for l in r["linhas"]:
        lin = []
        for _, k, t in cols:
            v = l.get(k)
            if t == "d" and v:
                v = datetime.fromisoformat(v)
            elif t == "b":
                v = "SIM" if v else ""
            lin.append(v)
        linhas.append(lin)
    moeda = tuple(i + 1 for i, (_, _, t) in enumerate(cols) if t == "n")
    ws = _aba(wb, f"Linha_a_linha_{nome_canal[:20]}", cab, linhas, [14] * len(cols), moeda=moeda)
    for i, (_, _, t) in enumerate(cols, 1):
        if t == "d":
            for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                for c in row:
                    c.number_format = "DD/MM/YYYY"
        if t == "p":
            for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                for c in row:
                    c.number_format = "0.00%"
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio
