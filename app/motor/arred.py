"""PLUTOS · regra oficial de arredondamento (Thaís, 16/09/2026).

A comissão do sistema é arredondada UMA VEZ NO TOTAL, não pedido a pedido.

Por quê: 17% × valor de cada pedido arredondado para 2 casas e depois somado
deixa um resíduo de centavos contra o cálculo feito sobre o total (no Madeira
set/26 deu R$ 0,38 em 258 pedidos). O rebate é cobrado do canal pelo valor
agregado, então o total manda. A linha a linha continua arredondada para
exibição — só os KPIs e o dia a dia usam o valor exato.

Uso no resumo de cada canal:
    ex = arred.mapa(linhas)                    # {id(linha): comissão sistema exata}
    com_sis = arred.total(linhas)              # arredondado 1x
    reb_com = round(com_sis - com_real, 2)
"""
from __future__ import annotations


def sis_exato(l: dict, taxa: bool = False) -> float:
    """Comissão do sistema do pedido, SEM arredondar.

    taxa=True soma a taxa fixa por pedido (Magalu: R$ 5,00/pedido). No Shopee o
    campo sis_taxa é só informativo (R$ 12/item já dentro da taxa de serviço do
    relatório) e por isso NÃO entra — por isso a taxa é opcional e explícita.
    """
    p = l.get("sis_pct")
    if p is None:
        return float(l.get("sis_rs") or 0.0)   # comissão cadastrada em R$ fixo
    v = (l.get("valor_prod") or 0.0) * p
    if taxa:
        v += (l.get("sis_taxa") or 0.0)
    return float(v)


def total(linhas: list[dict], taxa: bool = False) -> float:
    """Comissão do sistema do mês — soma exata, arredondada uma única vez."""
    return round(sum(sis_exato(l, taxa) for l in linhas), 2)
