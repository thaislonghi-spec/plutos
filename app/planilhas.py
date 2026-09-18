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


def export_rebates_xlsx(linhas: list[dict], quando) -> io.BytesIO:
    """A saída única para os outros apps: 1 linha por OC, 7 colunas fixas + apoio.
    'Frete cobrado pelo canal' fica EM BRANCO (não zero) quando o canal não informa."""
    wb = Workbook()
    wb.remove(wb.active)
    cab = ["OC", "Data", "Canal", "Rebate R$", "Rebate comissão", "Rebate frete", "Rebate TOTAL",
           "Comissão SISTEMA R$", "Comissão REAL R$", "Frete cobrado pelo canal R$", "% real", "Base (venda)",
           "Competência", "SKU", "Pedido canal", "Faltante campanha"]
    ls = [[l["oc"], datetime.fromisoformat(l["data"]), l["canal"], l["rebate_rs"], l["rebate_comissao"],
           l["rebate_frete"], l["rebate_total"], l.get("sis_rs"), l.get("tarifa"), l.get("frete_canal"),
           l.get("pct_real"), l.get("venda"),
           l["competencia"], l["sku"], l["pedido_canal"], l["faltante_status"]]
          for l in linhas]
    ws = _aba(wb, "REBATES", cab, ls, [22, 12, 18, 13, 16, 13, 14, 18, 16, 20, 9, 14, 12, 16, 20, 16],
              moeda=(4, 5, 6, 7, 8, 9, 10, 12))
    for row in ws.iter_rows(min_row=2, min_col=11, max_col=11):
        for c in row:
            c.number_format = "0.00%"
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        for c in row:
            c.number_format = "@"
    # resumo por canal × competência
    res: dict[tuple, list] = {}
    for l in linhas:
        k = (l["competencia"], l["canal"])
        a = res.setdefault(k, [0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0])
        a[0] += 1; a[1] += l["rebate_rs"]; a[2] += l["rebate_comissao"]; a[3] += l["rebate_frete"]; a[4] += l["rebate_total"]
        a[5] += l.get("sis_rs") or 0.0; a[6] += l.get("tarifa") or 0.0; a[7] += l.get("venda") or 0.0
        if l.get("frete_canal") is not None:
            a[8] += l["frete_canal"]; a[9] += 1
    l2 = [[k[0], k[1], v[0], round(v[1], 2), round(v[2], 2), round(v[3], 2), round(v[4], 2), round(v[5], 2), round(v[6], 2),
           (round(v[8], 2) if v[9] else None), v[9],
           (round(v[6] / v[7], 4) if v[7] else None), round(v[7], 2)] for k, v in sorted(res.items())]
    ws2 = _aba(wb, "Resumo", ["Competência", "Canal", "Pedidos", "Rebate R$", "Rebate comissão", "Rebate frete", "Rebate TOTAL",
                              "Comissão SISTEMA R$", "Comissão REAL R$", "Frete cobrado pelo canal R$", "OCs com frete informado",
                              "% real", "Base (venda)"],
               l2, [12, 18, 10, 14, 16, 14, 14, 18, 16, 20, 14, 9, 14], moeda=(4, 5, 6, 7, 8, 9, 10, 13))
    for row in ws2.iter_rows(min_row=2, min_col=12, max_col=12):
        for c in row:
            c.number_format = "0.00%"
    ws3 = wb.create_sheet("Como_ler")
    for t in [f"PLUTOS · EXPORT REBATES · gerado em {quando.strftime('%d/%m/%Y %H:%M')} (Brasília)",
              "", "Uma linha por OC com o rebate nas três formas. É este arquivo que os outros apps",
              "(Tropa de Elite, ORION, Hércules, ATLAS) importam.",
              "OC = pedido do marketplace (a Ordem de compra do ERP), sempre texto.",
              "Rebate R$ = volta em dinheiro · Rebate comissão = paga-se menos · Rebate frete = o canal ajuda no frete.",
              "Comissão SISTEMA = a de tabela (Parâmetros / ERP) · Comissão REAL = a que o canal cobrou de fato · % real = REAL ÷ base.",
              "Comissão REAL NÃO é rebate: fica fora das três formas e fora do Rebate TOTAL (o Tropa lê para a margem).",
              "Rebate comissão = SISTEMA − REAL. O Tropa/ORION devem partir da comissão SISTEMA para não contar o rebate duas vezes.",
              "Frete cobrado pelo canal = o frete que o CLIENTE pagou ao canal (receita de frete), por pedido: Mercado Livre = Frete Pedido.",
              "EM BRANCO = o arquivo do canal não traz esse valor (Magalu, Shopee, Madeira, Webcontinental): o Tropa usa a NF/CT-e. Branco ≠ zero (zero = o cliente pagou zero).",
              "Faltante campanha 'pendente' = pedido de tarifa zero ainda sem o valor pesquisado no portal (não soma)."]:
        ws3.append([t])
    ws3.column_dimensions["A"].width = 110
    ws3["A1"].font = Font(bold=True, size=13)
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def mlbs_xlsx(itens: list[dict], q: str, tipo: str) -> io.BytesIO:
    """Lista de MLB's do Mercado Livre — MLB · SKU · Descrição · Tipo de anúncio · % comissão."""
    wb = Workbook()
    wb.remove(wb.active)
    cab = ["MLB", "SKU", "Descrição", "Tipo de anúncio", "% Comissão", "Pedidos", "Primeira venda", "Última venda"]
    ls = [[m["mlb"], m["sku"], m.get("descricao") or "", m["tipo"], m["pct"], m["pedidos"],
           datetime.fromisoformat(m["primeira"]), datetime.fromisoformat(m["ultima"])] for m in itens]
    ws = _aba(wb, "Lista_MLBs", cab, ls, [18, 16, 46, 16, 12, 10, 14, 14])
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for c in row:
            c.number_format = "0.0%"
    for row in ws.iter_rows(min_row=2, min_col=7, max_col=8):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    filtro = " · ".join(x for x in [f"busca: {q}" if q else "", f"tipo: {tipo}" if tipo and tipo != "todos" else ""] if x)
    ws2 = _aba(wb, "Como_ler", ["Item", "Explicação"], [
        ["Origem", "Planilha 2 · Resumo de Rebates (Relatorios_PedidosxRebates_BI_MercadoLivre)"],
        ["Tipo de anúncio", "Premium = comissão 16,5% · Clássico = comissão 11,5% (tabela do Mercado Livre)"],
        ["Regra", "Vale o tipo da última venda do MLB; um MLB que mudou de tipo aparece com o tipo atual"],
        ["Descrição", "Vem do cadastro de SKUs subido em Parâmetros; vazio = SKU sem cadastro"],
        ["Filtro aplicado", filtro or "nenhum (lista completa)"],
    ], [18, 90])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def custo_coletas_xlsx(itens: list[dict], q: str, tipo: str) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    cab = ["MLB", "SKU", "Descrição", "Tipo de anúncio", "Custo Coletas (R$)", "Peso produto (kg)", "R$/kg",
           "Pedidos coletas", "Último pedido", "Data do último", "Valor produto (último)"]
    ls = [[m["mlb"], m["sku"], m.get("descricao") or "", m["tipo"], m["custo"], m.get("peso"), m.get("custo_kg"),
           m["pedidos"], m["pedido"], datetime.fromisoformat(m["data"]), m["valor_prod"]] for m in itens]
    ws = _aba(wb, "Custo_Coletas", cab, ls, [18, 16, 46, 16, 16, 14, 10, 12, 22, 14, 16], moeda=(5, 7, 11))
    for row in ws.iter_rows(min_row=2, min_col=10, max_col=10):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    for row in ws.iter_rows(min_row=2, min_col=9, max_col=9):
        for c in row:
            c.number_format = "@"
    filtro = " · ".join(x for x in [f"busca: {q}" if q else "", f"tipo: {tipo}" if tipo and tipo != "todos" else ""] if x)
    _aba(wb, "Como_ler", ["Item", "Explicação"], [
        ["Origem", "Planilha 2 · Resumo de Rebates — pedidos com Frete Coletas > 0"],
        ["Custo Coletas", "O frete coletas do ÚLTIMO pedido válido daquele MLB (o valor mais recente que o Meli cobrou)"],
        ["Peso / Descrição", "Cadastro de SKUs (CustoProduto) subido em Parâmetros"],
        ["Filtro aplicado", filtro or "nenhum (lista completa)"],
    ], [18, 90])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def custo_full_xlsx(itens: list[dict], comp: str, q: str, cd: str) -> io.BytesIO:
    """Custo do Fulfillment do Magalu por SKU — aba ORION primeiro (as três
    colunas da precificação) e a abertura completa do mês na segunda aba."""
    wb = Workbook()
    wb.remove(wb.active)

    # 1) ORION — só o que entra no preço
    cab = ["SKU", "Descrição", "Custo Full por unidade (R$)", "Coparticipação de frete por unidade (R$)",
           "CUSTO TOTAL POR UNIDADE (R$)", "Unidades no ciclo", "Base das unidades",
           "Peso (kg)", "Cubagem (m³)", "R$/un manuseio (tabela)", "Coleta por unidade (tabela R$/m³)", "R$/kg"]
    ls = [[m["sku"], m.get("descricao") or "",
           (m.get("full_un") if m.get("unidades") else None),
           (m.get("copart_un") if m.get("unidades") else None),
           (m.get("por_unidade") if m.get("unidades") else None),
           m.get("unidades"), m.get("base_un") or "sem unidade no ciclo",
           m.get("peso"), m.get("cubagem"), m.get("unit_manuseio"), m.get("coleta_tabela_un"),
           m.get("custo_kg")] for m in itens]
    _aba(wb, "ORION_custo_por_unidade", cab, ls, [16, 46, 20, 24, 22, 14, 18, 11, 12, 16, 22, 10],
         moeda=(3, 4, 5, 10, 11, 12))

    # 2) abertura do mês, para conferência
    cab2 = ["SKU", "Descrição", "CD", "Unidades", "R$/un manuseio (tabela)", "Manuseio (R$)",
            "Armazenagem (R$)", "Tempo de estoque (R$)", "Coleta rateada (R$)", "CUSTO FULL do mês (R$)",
            "Coparticipação de frete do mês (R$)", "CUSTO TOTAL do mês (R$)", "Cobranças", "Última cobrança"]
    ls2 = [[m["sku"], m.get("descricao") or "", m.get("cd") or "", m.get("unidades"), m.get("unit_manuseio"),
            m["manuseio"], m["armazenagem"], m["tempo_estoque"], m["coleta_rateio"], m.get("custo_full"),
            m["copart"], m["custo_total"], m.get("cobrancas"),
            (datetime.fromisoformat(m["ultimo"]) if m.get("ultimo") else None)] for m in itens]
    ws2 = _aba(wb, "Abertura_do_mes", cab2, ls2, [16, 44, 26, 12, 16, 13, 14, 16, 15, 17, 20, 18, 11, 14],
               moeda=(5, 6, 7, 8, 9, 10, 11, 12))
    for row in ws2.iter_rows(min_row=2, min_col=14, max_col=14):
        for c in row:
            c.number_format = "DD/MM/YYYY"

    filtro = " · ".join(x for x in [f"busca: {q}" if q else "", f"CD: {cd}" if cd and cd != "todos" else ""] if x)
    _aba(wb, "Como_ler", ["Item", "Explicação"], [
        ["Para que serve", "levar para o ORION o custo do Fulfillment do Magalu por unidade, SKU a SKU"],
        ["CUSTO FULL por unidade", "manuseio + armazenagem + tempo de estoque + coleta rateada, dividido pelas unidades"],
        ["Coparticipação por unidade", "o pedaço do frete que a Multimóveis paga, dividido pelas unidades"],
        ["CUSTO TOTAL por unidade", "a soma dos dois — é esta coluna que entra na formação de preço"],
        ["Unidades no ciclo", "unidades manuseadas no Fulfillment; quando o SKU não teve manuseio, "
                              "usa as unidades vendidas no Full (a coluna 'Base das unidades' diz qual foi)"],
        ["Sem unidade no ciclo", "SKU que só pagou armazenagem ou tempo de estoque, sem manuseio nem venda. "
                                 "O custo entra no total do mês, mas NÃO vira custo por unidade: é penalidade "
                                 "de estoque parado, não custo do produto vendido"],
        ["Coleta rateada", "a coleta é cobrada por agenda e m³, SEM SKU. É rateada pelo VOLUME que cada SKU "
                           "ocupou (cubagem × unidades), que é como o Magalu cobra"],
        ["Coleta por unidade (tabela)", "a tarifa do m³ do ciclo × a cubagem da peça — quanto custa coletar UMA "
                                        "unidade. Serve para precificar produto novo, sem depender do rateio do mês"],
        ["R$/un manuseio (tabela)", "o manuseio é uma tabela por faixa de tamanho (em set/26: R$ 0,00 · 3,90 · "
                                    "19,90 · 32,90 · 37,90). É a faixa do SKU, do último arquivo subido"],
        ["Manuseio / Armazenagem / Tempo de estoque", "cobranças do Fulfillment (CSV-156171)"],
        ["Coparticipação de frete", f"Planilha 2 · Vendas no período, competência {comp}"],
        ["Peso / Descrição", "Cadastro de SKUs (Tabela de Custos MUL), subido em Parâmetros"],
        ["Filtro aplicado", filtro or "nenhum (lista completa)"],
    ], [30, 100])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def sugestao_full_xlsx(e: dict, comp: str, dias_alvo: int) -> io.BytesIO:
    """Lista de envio para o Fulfillment do Magalu: o que mandar, quanto, quanto
    ocupa e quanto custa colocar lá dentro."""
    wb = Workbook()
    wb.remove(wb.active)
    env = [m for m in e["por_sku"] if m.get("sugestao")]
    cab = ["SKU", "Descrição", "ENVIAR (un)", "Ação", "Estoque hoje", "Alvo (un)", "Venda média/dia",
           "Cobertura hoje (dias)", "Prazo do CD (dias)", "Dias a cobrir", "m³ do envio", "Custo da coleta (R$)", "Manuseio quando vender (R$)",
           "Coleta + manuseio (R$)", "Valor do envio a custo (R$)", "CD atual"]
    ls = [[m["sku"], m.get("descricao") or "", m["sugestao"], m.get("acao"), m["estoque"], m["alvo"],
           m["media_dia"], m["cobertura"], m.get("prazo"), m.get("dias_cobrir"), m["sug_m3"],
           m["sug_coleta"], m["sug_manuseio"], m["sug_custo"], m["sug_valor"], m.get("cd") or ""]
          for m in sorted(env, key=lambda x: (x.get("acao") != "URGENTE", -x["sugestao"]))]
    _aba(wb, f"Enviar_{dias_alvo}dias", cab, ls,
         [16, 44, 12, 11, 13, 11, 14, 16, 14, 13, 12, 15, 18, 20, 20, 26], moeda=(12, 13, 14, 15))

    # a mesma lista, aberta por CD (cada um com o seu prazo)
    cab_cd = ["CD", "SKU", "Descrição", "ENVIAR (un)", "Ação", "Estoque no CD", "% do estoque do SKU",
              "Venda rateada no ciclo", "Venda média/dia", "Cobertura (dias)", "Prazo do CD (dias)",
              "Dias a cobrir", "Alvo (un)", "m³ do envio"]
    ls_cd = [[x["cd"], x["sku"], x.get("descricao") or "", x["sugestao"], x.get("acao"), x["estoque"],
              x["parte"], x["vendidas"], x["media_dia"], x["cobertura"], x["prazo"], x["dias_cobrir"],
              x["alvo"], x["sug_m3"]] for x in e.get("por_cd", []) if x["sugestao"]]
    _aba(wb, "Enviar_por_CD", cab_cd, ls_cd, [22, 16, 42, 12, 11, 14, 18, 20, 15, 15, 16, 13, 11, 12])

    sair = [m for m in e["por_sku"] if m.get("acao") in ("RETIRAR", "excesso")]
    cab2 = ["SKU", "Descrição", "Ação", "Estoque hoje", "Vendidas no ciclo", "Cobertura (dias)",
            "Valor parado (R$)", "m³ ocupados", "Armazenagem paga (R$)", "Aniversário pago (R$)", "CD"]
    ls2 = [[m["sku"], m.get("descricao") or "", ("RETIRAR — sem venda no ciclo" if m["acao"] == "RETIRAR"
                                                 else "EXCESSO — cobertura acima do dobro do alvo"),
            m["estoque"], m["vendidas"], m["cobertura"], m["valor_estoque"], m["volume"],
            m["custo_arm"], m["aniversario_rs"], m.get("cd") or ""]
           for m in sorted(sair, key=lambda x: -x["valor_estoque"])]
    _aba(wb, "Tirar_do_CD", cab2, ls2, [16, 44, 34, 13, 16, 15, 16, 12, 18, 18, 26], moeda=(7, 9, 10))

    # a tela inteira, SKU a SKU
    cab3 = ["SKU", "Descrição", "Ação", "ENVIAR (un)", "Estoque hoje", "Valor do estoque (R$)",
            "Vendidas no ciclo", "Venda média/dia", "Cobertura (dias)", "Alvo (un)", "m³ ocupados",
            "Armazenagem no ciclo (R$)", "Aniversário pago (R$)", "Dias com estoque no ciclo", "CD"]
    ls3 = [[m["sku"], m.get("descricao") or "", m.get("acao"), m.get("sugestao"), m["estoque"],
            m["valor_estoque"], m["vendidas"], m["media_dia"], m["cobertura"], m.get("alvo"),
            m["volume"], m["custo_arm"], m["aniversario_rs"], m.get("dias_no_ciclo"), m.get("cd") or ""]
           for m in e["por_sku"]]
    _aba(wb, "Estoque_Full", cab3, ls3, [16, 44, 12, 12, 13, 18, 16, 15, 15, 11, 12, 20, 18, 20, 26],
         moeda=(6, 12, 13))

    # a foto dia a dia
    _aba(wb, "Estoque_dia_a_dia", ["Dia", "Unidades em estoque", "SKUs", "Armazenagem do dia (R$)"],
         [[datetime.fromisoformat(d["dia"]), d["unidades"], d["skus"], d["custo"]] for d in e["por_dia"]],
         [14, 18, 10, 20], moeda=(4,))

    _aba(wb, "Como_ler", ["Item", "Explicação"], [
        ["Base", f"foto de estoque de {e.get('foto', '')} e vendas do Fulfillment da competência {comp}"],
        ["ENVIAR", f"venda média por dia × ({dias_alvo} dias de cobertura + o prazo do CD) − estoque de hoje. "
                   "O prazo do CD é o tempo entre pedir a coleta e a peça ficar vendável lá (Parâmetros)"],
        ["URGENTE", "a cobertura de hoje é MENOR que o prazo de chegada — o SKU rompe antes da reposição entrar"],
        ["Venda média/dia", "unidades vendidas no Full no ciclo ÷ dias de venda do período"],
        ["Custo da coleta", "m³ do envio × a tarifa do m³ do ciclo — é o que o Magalu cobra para BUSCAR a mercadoria"],
        ["Manuseio quando vender", "unidades × a faixa de manuseio do SKU. NÃO é custo de colocar no Full: "
                                   "o manuseio é cobrado no movimento, então entra quando a peça sair"],
        ["RETIRAR", "estoque no CD sem UMA venda no ciclo: paga armazenagem, paga aniversário e ocupa m³"],
        ["EXCESSO", f"cobertura acima de {2 * dias_alvo} dias — mais que o dobro do alvo"],
        ["Atenção", "produto parado NÃO recebe sugestão de envio; ele tem que sair do CD"],
        ["Enviar_por_CD", "a mesma sugestão aberta por CD, cada um com o seu prazo"],
        ["Venda rateada", "o relatório do Magalu NÃO informa de qual CD saiu cada venda (a coluna 'CD de Origem' "
                          "vem vazia no Fulfillment). Até o canal mandar esse dado, a venda do SKU é rateada pela "
                          "participação de cada CD no estoque dele — é estimativa, não medição"],
    ], [26, 100])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def coletas_xlsx(pedidos: list[dict], desc: dict, comp: str) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    cab = ["Pedido", "Data", "MLB", "SKU", "Descrição", "Tipo de anúncio", "Qtd", "Valor produtos", "Frete pedido",
           "Frete Coletas", "Cupom canal", "Comissão bruta", "Comissão líquida", "Rebate comissão (BI)"]
    ls = [[p["pedido_mkt"], datetime.fromisoformat(p["data"]), p["anuncio"], p["sku"], ((desc.get(p["sku"]) or {}).get("descricao") or (desc.get(p["sku"]) or {}).get("descricao_curta") or ""),
           p["tipo"], p["qtd"], p["valor_prod"], p["frete"], p["frete_coletas"], p["cupom_canal"], p["com_bruta"], p["com_liq"], p["rebate_bi"]]
          for p in sorted(pedidos, key=lambda p: p["data"])]
    ws = _aba(wb, f"Coletas_{comp.replace('-', '')}", cab, ls, [22, 12, 18, 16, 40, 14, 6, 14, 13, 13, 12, 14, 14, 16],
              moeda=(8, 9, 10, 11, 12, 13, 14))
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        for c in row:
            c.number_format = "@"
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


# ---------------------------------------------------------------------------
# SHOPEE · planilha no formato da Gabi (aba "novo" + "Planilha1" com as fórmulas
# dela). Pedido a pedido do PLUTOS, para ela usar direto no fechamento.
# ---------------------------------------------------------------------------
SHOPEE_GABI_COLS = [
    ("COM PROMOB", None), ("PROMOB", None), ("ID do pedido", "pedido_mkt"), ("Status do pedido", "status"),
    ("Hot Listing", None), ("Cancelar Motivo", None), ("Status da Devolução / Reembolso", None),
    ("Número de rastreamento", "id_mkt"), ("Opção de envio", "envio"), ("Método de envio", None),
    ("Data", "data"), ("Hora do pagamento do pedido", None), ("Data prevista de envio", None),
    ("Tempo de Envio", None), ("Domestic Delivered Date", None), ("Hora completa do pedido", None),
    ("Data da Finalização do Cancelamento", None), ("Pedido FBS", None),
    ("Nº de referência do SKU principal", "sku"), ("Nome do Produto", "produto"),
    ("Número de referência SKU", "anuncio"), ("Nome da variação", None), ("Shopee Owned", None),
    ("Preço original", None), ("Preço acordado", None), ("Quantidade", "qtd"),
    ("Subtotal do produto", "valor_prod"), ("Desconto do vendedor", "desc_vendedor"), ("Desconto do vendedor ", None),
    ("Incentivo Shopee para ação comercial", "incentivo"), ("Ajuste por participação em ação comercial", "ajuste"),
    ("Peso total SKU", None), ("Número de produtos pedidos", "itens"), ("Peso total do pedido", None),
    ("Código do Cupom", None), ("Cupom do vendedor", "cupom_seller"),
    ("Coin Cashback Voucher Amount Sponsored by Seller", None), ("CUPOM GERAL", None),
    ("CUPOM SHOPEE", "cupom_shopee"), ("ps_csv_pix_discount_br", None),
    ("Indicador da Leve Mais por Menos", None), ("Desconto Shopee da Leve Mais por Menos", None),
    ("Desconto da Leve Mais por Menos do vendedor", None), ("Compensar", None),   # SEMPRE zero: a Shopee não paga mais moedas (decisão da Thaís 14/09)
    ("Total descontado Cartão de Crédito", None), ("Valor Total", None),
    ("Taxa de envio pagas pelo comprador", "frete_comprador"), ("", None), ("Taxa de Envio Reversa", None),
    ("Taxa de transação", "taxa_transacao"), ("Taxa de comissão bruta", "comissao_bruta"),
    ("Taxa de comissão líquida", "comissao_liquida"), ("Taxa de serviço bruta", "servico_bruta"),
    ("Taxa de serviço líquida", "servico_liquida"), ("Total global", "total_global"),
    ("Valor estimado do frete", "frete"), ("Nome de usuário (comprador)", None), ("Nome do destinatário", None),
    ("Telefone", None), ("Endereço de entrega", None), ("Cidade", None), ("Bairro", None), ("Cidade ", None),
    ("UF", "uf"), ("País", None), ("CEP", None), ("Observação do comprador", None), ("Nota", None),
    ("FRETE", "rebate_frete"),
]


def shopee_gabi_xlsx(linhas: list[dict], comp: str, pct: float, quando) -> io.BytesIO:
    """A planilha da Gabi (Shopee) já preenchida pelo PLUTOS: aba 'novo' com as
    mesmas 69 colunas (1 linha por PEDIDO, sem cancelado e sem duplicidade) e
    aba 'Planilha1' com os mesmos blocos e as mesmas fórmulas SUMIFS por dia."""
    wb = Workbook()
    wb.remove(wb.active)
    linhas = sorted(linhas, key=lambda l: (l["data"], l["pedido_mkt"]))

    ws = wb.create_sheet("novo")
    ws.append([c for c, _ in SHOPEE_GABI_COLS])
    for c in ws[1]:
        c.fill, c.font, c.alignment = CAB, CAB_F, Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for i, l in enumerate(linhas, start=2):
        linha = []
        for cab, campo in SHOPEE_GABI_COLS:
            if cab == "COM PROMOB":
                linha.append(pct)
            elif cab == "PROMOB":
                linha.append(f"=A{i}*AA{i}")
            elif campo == "data":
                linha.append(datetime.fromisoformat(l["data"]))
            elif cab == "Compensar":
                linha.append(0)          # sempre zero
            elif campo is None:
                linha.append(None)
            else:
                linha.append(l.get(campo))
        ws.append(linha)
    for col, larg in (("C", 20), ("D", 16), ("K", 12), ("S", 16), ("T", 40), ("U", 18)):
        ws.column_dimensions[col].width = larg
    for row in ws.iter_rows(min_row=2, min_col=11, max_col=11):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for c in row:
            c.number_format = "@"
    ws.freeze_panes = "D2"

    # ---- Planilha1: mesmos blocos e mesmas fórmulas da Gabi, um dia por coluna
    dias = sorted({l["data"] for l in linhas})
    p1 = wb.create_sheet("Planilha1")
    fim = get_column_letter(1 + len(dias))

    def bloco(linha_ini: int, titulo: str, itens: list[tuple[str, str]], ref: int):
        p1.cell(linha_ini, 1, titulo).font = Font(bold=True)
        for j, d in enumerate(dias, start=2):
            cel = p1.cell(linha_ini, j, datetime.fromisoformat(d))
            cel.number_format = "DD/MM/YYYY"
            cel.font = Font(bold=True)
        for k, (rot, colu) in enumerate(itens, start=1):
            p1.cell(linha_ini + k, 1, rot)
            for j, _ in enumerate(dias, start=2):
                L = get_column_letter(j)
                extra = "*0.01" if rot == "Compensar moedas" else ""
                p1.cell(linha_ini + k, j, f"=SUMIFS(novo!${colu}:${colu},novo!$K:$K,Planilha1!{L}${ref}){extra}").number_format = BRL

    bloco(1, "REBATE ATIVO", [("Incentivo Shopee para ação comercial", "AD"), ("CUPOM SHOPEE", "AM"),
                              ("Desconto de Frete Aproximado", "BQ"), ("Compensar moedas", "AR")], 1)
    p1.cell(7, 1, "Diferença a lançar").font = Font(bold=True)
    for j, _ in enumerate(dias, start=2):
        L = get_column_letter(j)
        p1.cell(7, j, f"=SUM({L}2:{L}5)").number_format = BRL

    bloco(11, "DIFERENÇA COMISSÃO", [("COMISSÃO PROMOB", "B"), ("Taxa de comissão bruta", "AY"),
                                     ("Taxa de serviço bruta", "BA"),
                                     ("Ajuste por participação em ação comercial", "AE"),
                                     ("Coin Cashback Voucher Amount Sponsored by Seller", "AK")], 11)
    p1.cell(18, 1, "Diferença a lançar").font = Font(bold=True)
    for j, _ in enumerate(dias, start=2):
        L = get_column_letter(j)
        p1.cell(18, j, f"=({L}12-(({L}13+{L}14)-{L}15-{L}16))").number_format = BRL

    p1.cell(20, 1, "REBATE ATIVO").font = Font(bold=True)
    p1.cell(20, 2, f"=SUM(B7:{fim}7)").number_format = BRL
    p1.cell(21, 1, "DIFERENÇA COMISSÃO").font = Font(bold=True)
    p1.cell(21, 2, f"=SUM(B18:{fim}18)").number_format = BRL
    p1.cell(22, 1, "TOTAL").font = Font(bold=True)
    p1.cell(22, 2, "=B20+B21").number_format = BRL

    bloco(24, "BI RODRIGO", [("Incentivo Shopee para ação comercial", "AD"),
                             ("Ajuste por participação em ação comercial", "AE"), ("CUPOM SHOPEE", "AM"),
                             ("Desconto de Frete Aproximado", "BQ"), ("Compensar moedas", "AR")], 24)
    p1.cell(31, 1, "Total BI").font = Font(bold=True)
    for j, _ in enumerate(dias, start=2):
        L = get_column_letter(j)
        p1.cell(31, j, f"=SUM({L}25:{L}29)").number_format = BRL
    p1.cell(33, 1, "Total do dia (rebate ativo + comissão)").font = Font(bold=True)
    for j, _ in enumerate(dias, start=2):
        L = get_column_letter(j)
        p1.cell(33, j, f"={L}7+{L}18").number_format = BRL
    p1.column_dimensions["A"].width = 42
    for j, _ in enumerate(dias, start=2):
        p1.column_dimensions[get_column_letter(j)].width = 13

    # ---- conferência com o número oficial do PLUTOS
    por_dia: dict[str, list] = {}
    for l in linhas:
        a = por_dia.setdefault(l["data"], [0, 0.0, 0.0, 0.0, 0.0, 0.0])
        a[0] += 1; a[1] += l["rebate_rs"]; a[2] += l["rebate_comissao"]; a[3] += l["rebate_frete"]
        a[4] += l["rebate_total"]; a[5] += l["sis_rs"]
    ls = [[datetime.fromisoformat(d), v[0], round(v[1], 2), round(v[2], 2), round(v[3], 2), round(v[4], 2), round(v[5], 2)]
          for d, v in sorted(por_dia.items())]
    ws3 = _aba(wb, "PLUTOS_por_dia", ["Dia", "Pedidos", "Rebate R$", "Rebate comissão", "Rebate frete",
                                      "REBATE TOTAL", "Comissão sistema (12% + R$/item)"],
               ls, [12, 10, 14, 16, 14, 15, 24], moeda=(3, 4, 5, 6, 7))
    for row in ws3.iter_rows(min_row=2, min_col=1, max_col=1):
        for c in row:
            c.number_format = "DD/MM/YYYY"

    mapa = [
        ["O que é", "Coluna do relatório da Shopee (Order.all)", "Conta do PLUTOS", "Entra em"],
        ["Pedido (chave)", "ID do pedido", "1 linha por pedido; o export vem 1 por item e taxas/incentivos repetem", "chave = OC do ERP"],
        ["Competência", "Data de criação do pedido", "mês da venda", "todos"],
        ["Fora da conta (1)", "Status do pedido = Cancelado", "não soma nada", "—"],
        ["Fora da conta (2)", "Status do pedido = Não pago", "não soma; quando o pedido for pago, o arquivo seguinte traz ele de volta", "—"],
        ["Fora da conta (3)", "Status da Devolução / Reembolso = Solicitação aprovada", "a venda voltou: não soma nada", "—"],
        ["Venda (base)", "Subtotal do produto", "soma dos itens do pedido", "base do % e do % sobre venda"],
        ["Quantidade", "Quantidade", "soma dos itens", "R$ 12,00 por ITEM"],
        ["COMISSÃO SISTEMA (Promob)", "— (tabela de Parâmetros)", "12% × Subtotal do produto (= produto + IPI)", "comissão sistema"],
        ["COMISSÃO REAL (Shopee)", "Taxa de comissão bruta (AY) + Taxa de serviço bruta (BA) − Ajuste por participação em ação comercial (AE)",
         "o que a Shopee cobrou de fato; o ajuste é desconto da campanha e entra subtraindo", "comissão real"],
        ["Taxa fixa R$ 12,00/un", "está DENTRO da Taxa de serviço bruta (BA)", "não é somada em lugar nenhum: já vem cobrada no relatório", "nada a fazer"],
        ["REBATE EM COMISSÃO", "as duas linhas acima", "sistema − real", "Rebate em comissão"],
        ["REBATE EM R$ (1ª parte)", "Incentivo Shopee para ação comercial", "soma", "Rebate em R$"],
        ["REBATE EM R$ (2ª parte)", "Incentivo de cupom", "soma — é ESTA coluna, NÃO a coluna 'Cupom'", "Rebate em R$"],
        ["REBATE EM FRETE", "Valor estimado do frete − Taxa de envio pagas pelo comprador",
         "o que sobra é o Frete Grátis que a Shopee banca (nunca negativo)", "Rebate em frete"],
        ["Compensar moedas", "Compensar Moedas Shopee", "SEMPRE ZERO — a Shopee não paga mais moedas", "não entra"],
        ["Cupom do vendedor", "Cupom do vendedor", "guardado para consulta — é desconto NOSSO", "não é rebate"],
        ["Coluna 'Cupom'", "Cupom", "NÃO usar: é o cupom total (nosso + Shopee)", "não entra"],
        ["Comissão/serviço líquida", "Taxa de comissão líquida · Taxa de serviço líquida", "guardadas para conferência", "não entram na conta"],
        ["Tarifa zero", "Ajuste ≥ Taxa de comissão bruta", "a campanha cobriu toda a comissão", "marcação"],
    ]
    _aba(wb, "De_onde_vem", mapa[0], mapa[1:], [26, 62, 66, 24])

    ws4 = wb.create_sheet("Como_ler")
    for t in [f"PLUTOS · SHOPEE no formato da Gabi · {comp} · gerado em {quando.strftime('%d/%m/%Y %H:%M')} (Brasília)",
              "",
              "aba 'novo'   = a base do PLUTOS nas mesmas 69 colunas do export da Shopee, já tratada:",
              "               1 LINHA POR PEDIDO (no export vem 1 por item, com taxas e incentivos repetidos),",
              "               cancelados fora e nada somado duas vezes. Coluna A = 12% · B = A × Subtotal.",
              "               Coluna BQ (FRETE) = valor estimado do frete − taxa de envio paga pelo comprador.",
              "aba 'Planilha1' = os mesmos blocos e as mesmas fórmulas SUMIFS do fechamento dela, um dia por coluna.",
              "aba 'PLUTOS_por_dia' = o número oficial do app para conferir.",
              "",
              "CONTA OFICIAL: comissão real Shopee = comissão bruta (AY) + serviço bruta (BA) − ajuste (AE);",
              "comissão Promob = 12% do Subtotal do produto (= produto + IPI); rebate = Promob − real.",
              "A taxa fixa de R$ 12,00/un JÁ VEM dentro da 'Taxa de serviço bruta' (conferido: tirando 12 ×",
              "quantidade, o serviço vira exatamente 2,00% do subtotal) — não se soma em lugar nenhum.",
              "",
              "Compensar moedas: SEMPRE ZERO (a Shopee não paga mais moedas para a Multimóveis) — a coluna AR",
              "sai zerada de propósito e o rebate do PLUTOS também não conta moedas.",
              ]:
        ws4.append([t])
    ws4.column_dimensions["A"].width = 105
    ws4["A1"].font = Font(bold=True, size=13)
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def amazon_faltantes_xlsx(f: dict, sem_erp: list[dict], comp: str) -> io.BytesIO:
    """Pedidos do ERP que não apareceram em nenhum relatório de transações da
    Amazon — a lista de procura da Gabi. Aba 1: os que estão dentro da janela já
    coberta (buraco de verdade). Aba 2: fora da janela (falta subir arquivo).
    Aba 3: o contrário — pedido da Amazon sem par no ERP."""
    wb = Workbook()
    wb.remove(wb.active)

    cab = ["ID do pedido (Amazon)", "Pedido ERP", "Data do pedido", "NF", "Data da NF", "Status ERP",
           "Cliente", "Cidade", "UF", "Valor produtos (R$)", "IPI (R$)", "Frete (R$)",
           "TOTAL do pedido / NF (R$)", "% no ERP", "Comissão prevista (R$)", "Tem ID da Amazon?"]
    larg = [26, 14, 14, 10, 13, 12, 14, 22, 6, 18, 12, 12, 22, 10, 20, 18]

    def _ls(itens):
        return [[m["oc"], m["pedido_erp"], m["data"], m["nf"], m["data_nf"], m["status"],
                 m["cliente"], m["cidade"], m["uf"], m["valor_prod"], m["ipi"], m["frete"],
                 m["total"], (m["pct_erp"] or 0) * 100, m["comissao_prevista"],
                 "sim" if m["tem_id"] else "NÃO — conferir a Ordem de compra no ERP"] for m in itens]

    dentro = [m for m in f["itens"] if m["onde"] == "dentro"] or \
             [m for m in f.get("itens", []) if f.get("onde") != "fora"]
    _aba(wb, "Procurar_dentro_da_janela", cab, _ls([m for m in f["itens"] if m["onde"] == "dentro"]),
         larg, moeda=(10, 11, 12, 13, 15))
    _aba(wb, "Aguardando_repasse", cab, _ls([m for m in f["itens"] if m["onde"] == "aguardando"]),
         larg, moeda=(10, 11, 12, 13, 15))
    _aba(wb, "Fora_da_janela", cab, _ls([m for m in f["itens"] if m["onde"] == "fora"]),
         larg, moeda=(10, 11, 12, 13, 15))

    cab3 = ["ID do pedido (Amazon)", "Data", "Produto", "Produto (R$)", "Desconto (R$)", "Frete (R$)",
            "Base da comissão (R$)", "Tarifas da Amazon (R$)", "% cobrado", "Faixa da categoria"]
    ls3 = [[l["pedido_mkt"], l["data"], l.get("produto") or "", l["valor_prod"], l["desconto"], l["frete"],
            l["sis_base"], l["tarifa"], (l["pct_comissao"] or 0) * 100, l["tipo"]] for l in sem_erp]
    _aba(wb, "Amazon_sem_par_no_ERP", cab3, ls3, [26, 13, 48, 15, 14, 12, 20, 22, 12, 18],
         moeda=(4, 5, 6, 7, 8))

    ws = wb.create_sheet("Como_ler", 0)
    for l in [
        [f"PLUTOS · Amazon · pedidos faltantes · {comp}"],
        [],
        ["Janela já coberta pelos relatórios subidos:", f"{f['de']} a {f['ate']}"],
        ["Pedidos da Amazon no ERP neste mês:", f["erp_total"]],
        ["Já conferidos (achados no relatório):", f["conferidos"]],
        ["FALTANDO dentro da janela (procurar):", f["contagem"]["dentro"], f["rs"]["dentro"]],
        ["Aguardando repasse (pedido recente):", f["contagem"]["aguardando"], f["rs"]["aguardando"]],
        ["Faltando fora da janela (falta subir o arquivo):", f["contagem"]["fora"], f["rs"]["fora"]],
        [],
        ["Dentro da janela", "o relatório de transações daquele período JÁ foi subido e mesmo assim o pedido não apareceu."],
        ["", "Causas prováveis: pedido ainda não pago pela Amazon (fica para o próximo repasse), cancelado no canal,"],
        ["", "ou Ordem de compra digitada diferente no ERP. É esta aba que precisa de procura."],
        ["Aguardando repasse", f"pedido dos ultimos {f['folga']} dias da janela: a Amazon paga de 2 a 17 dias depois do pedido, entao ainda pode entrar no proximo relatorio."],
        ["Fora da janela", "o pedido é de um período que ainda não foi subido. Não é erro: falta o arquivo daquele pedaço."],
        ["Comissão prevista", f"{f['pct'] * 100:.2f}% (cadastro) x TOTAL do pedido (NF). É o que deixa de ser conferido enquanto o pedido não aparece."],
        ["Amazon sem par no ERP", "o contrário: veio no relatório da Amazon e não achamos a Ordem de compra no ERP."],
    ]:
        ws.append(l)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 96
    ws.column_dimensions["C"].width = 18
    ws["A1"].font = Font(bold=True, size=13)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio
