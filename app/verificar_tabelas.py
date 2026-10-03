"""VERIFICADOR DE TABELAS DO PLUTOS — roda o app num servidor local, abre
cada tela num navegador e confere, tabela por tabela:
  1) o cabecalho tem o MESMO numero de colunas que as linhas (contando colspan)
  2) nenhuma celula esta cortada com reticencias
  3) nenhum texto esta quebrando letra a letra

Uso:  DATA_DIR=<pasta dos dados> python verificar_tabelas.py
Precisa de playwright. Rodar SEMPRE antes de publicar uma mudanca de tela.
"""
import threading, time, os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import logging; logging.getLogger("werkzeug").setLevel(logging.ERROR)
import server as S
from flask.sessions import SecureCookieSessionInterface
from werkzeug.serving import make_server
srv=make_server("127.0.0.1",5091,S.app); threading.Thread(target=srv.serve_forever,daemon=True).start(); time.sleep(1)
ck=SecureCookieSessionInterface().get_signing_serializer(S.app).dumps({"usuario":"THAIS","empresa":"multimoveis","papel":"admin"})

URLS=["/","/canal/meli","/canal/magalu","/canal/shopee","/canal/madeira","/canal/webcont",
      "/canal/colombo","/canal/amazon","/canal/amazon/faltantes","/canal/cbahia",
      "/canal/cbahia/comissoes-a-maior","/canal/cbahia/pendentes",
      "/pedidos?canal=meli","/pedidos?canal=amazon","/pedidos?canal=cbahia",
      "/linha/meli","/linha/amazon","/linha/magalu","/linha/shopee","/linha/cbahia","/coletas","/custo-coletas",
      "/custo-full","/custo-full-real","/envio-full","/entregas","/faltante","/mlbs",
      "/arquivos","/parametros","/conta","/saude","/canais","/erp"]

JS = r"""
() => {
  const saida = [];
  document.querySelectorAll('table').forEach((t, idx) => {
    const cab = t.querySelector('thead tr:last-child');
    if (!cab) return;
    let nth = 0;
    cab.querySelectorAll('th,td').forEach(c => nth += (parseInt(c.colSpan)||1));
    const corpo = [];
    t.querySelectorAll('tbody tr').forEach((tr, i) => {
      let n = 0;
      tr.querySelectorAll('td,th').forEach(c => n += (parseInt(c.colSpan)||1));
      if (n !== nth) corpo.push({linha: i, td: n});
    });
    // celula cujo conteudo nao cabe (clip por overflow OU texto cortado)
    const apertadas = [];
    t.querySelectorAll('th,td').forEach(c => {
      if (c.colSpan > 1) return;
      const r = c.getBoundingClientRect();
      if (r.width < 1) return;
      if (c.scrollWidth > c.clientWidth + 2) apertadas.push((c.textContent||'').trim().slice(0,30));
      // texto quebrando letra a letra: mais de 1 linha para uma palavra so
      const cs = getComputedStyle(c), lh = parseFloat(cs.lineHeight) || 16;
      const linhas = Math.round((c.scrollHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom)) / lh);
      const palavras = (c.textContent||'').trim().split(/\s+/).filter(Boolean).length;
      if (linhas >= 3 && palavras <= 2 && (c.textContent||'').trim().length < 24)
        apertadas.push('QUEBRADA: ' + (c.textContent||'').trim().slice(0,22));
    });
    if (corpo.length || apertadas.length)
      saida.push({tabela: idx, classe: t.className || '(sem classe)', th: nth,
                  linhas_diferentes: corpo.slice(0,3), n_diferentes: corpo.length,
                  apertadas: [...new Set(apertadas)].slice(0,4)});
  });
  return saida;
}
"""
from playwright.sync_api import sync_playwright
achados = []
with sync_playwright() as p:
    b=p.chromium.launch(executable_path="/opt/pw-browsers/chromium")
    for w,h in ((1600,950),(1366,800)):
        ctx=b.new_context(viewport={"width":w,"height":h})
        ctx.add_cookies([{"name":"session","value":ck,"domain":"127.0.0.1","path":"/"}]); pg=ctx.new_page()
        for u in URLS:
            url="http://127.0.0.1:5091"+u+("&" if "?" in u else "?")+"mes=2026-09"
            try:
                pg.goto(url, timeout=25000); pg.wait_for_timeout(300)
                r = pg.evaluate(JS)
                for x in r: achados.append(dict(x, tela=u, tamanho=f"{w}x{h}"))
            except Exception as ex:
                achados.append({"tela":u,"tamanho":f"{w}x{h}","erro":str(ex)[:90]})
        ctx.close()
    b.close()
srv.shutdown()
print(json.dumps(achados, ensure_ascii=False, indent=1))
print("TOTAL DE ACHADOS:", len(achados))
