# PLUTOS · Rebates

**Central de rebates dos marketplaces da Multimóveis.** Concentra, num só lugar e por OC, tudo o que os canais devolvem para recompor a margem — e entrega o resumo pronto para o Tropa de Elite e o ORION.

## O que é rebate (regra da casa)

Rebate é a forma como os canais ajudam a recompor a margem para rentabilizar a venda. Cada canal tem a sua particularidade, mas no fim são três formas:

| Forma | O que é | Prova |
|---|---|---|
| **1 · R$** | volta em dinheiro na conta | extrato / export do canal |
| **2 · Comissão** | desconto no % de comissão — não volta dinheiro, paga-se menos | comissão do sistema × comissão cobrada |
| **3 · Frete** | o canal ajuda a pagar o frete: devolve, não cobra, ou encontro de contas | extrato / ausência de cobrança / conciliação |

## Canais

Um box por canal, cada um com a sua regra, o arquivo que prova, a chave e o previsto. Ordem: Mercado Livre · Magazine Luiza · Madeira Madeira · Colombo · Casas Bahia · Amazon · Webcontinental · Shopee.

| Canal | Status |
|---|---|
| Mercado Livre | **ativo** — Cupom Meli (R$) · Diferença de comissão (Promob − tarifa) · Faltante campanha (tabela manual, tarifa zero) |
| demais | em construção — entram conforme cada box é desenhado |

## Telas

- **GERAL** — rebate total nas três formas, canal a canal, na competência
- **Por pedido** — cada OC com o rebate aberto (R$ / comissão / frete / total)
- **Canal** — o box: dia a dia, sinais da rodada, como o rebate nasce
- **Faltante campanha · Meli** — pedidos de tarifa zero; valor pesquisado no portal, com sugestão automática por anúncio/SKU
- **Arquivos** — upload por canal, histórico de rodadas, saídas
- **Parâmetros** — o que é editável por empresa (as regras ficam dentro)
- **Conta** — Admin · Gestor · Operador · Equipe; senha inicial trocada no 1º acesso

## Saídas

- `PLUTOS_Rebates_AAAAMM.xlsx` — por pedido, por canal, Meli por dia, como ler
- `/api/rebates/<AAAA-MM>.json` — rebate por OC nas três formas, para o Tropa de Elite e o ORION (token `PLUTOS_TOKEN`)

## Rodar

```
pip install -r requirements.txt
cd app && python server.py        # http://localhost:5055
```

Render: o `render.yaml` cria o serviço com disco em `/var/data`. Variáveis: `DATA_DIR`, `PLUTOS_SENHA`, `SECRET_KEY`, `PLUTOS_TOKEN`, `TZ=America/Sao_Paulo`.

## Estrutura

```
app/server.py        telas, contas, rodadas
app/motor/meli.py    box do Mercado Livre (o método da Gabi, automatizado)
app/planilhas.py     saídas em Excel
app/templates/       layout da família (Tropa de Elite) nas cores do PLUTOS
app/static/          estilo, marca, ícones
dados/<empresa>/     arquivos, rodadas, tabelas manuais, parâmetros (fora do git)
```

Padrões: hora de Brasília, datas dd/mm/aaaa, multiempresa desde o início, regras dentro do sistema e parâmetros editáveis por empresa.

---
Multimóveis · família EVEREST (Tropa de Elite · ORION · HÉRCULES · ATLAS · ARGOS · TALOS · PLUTOS)
