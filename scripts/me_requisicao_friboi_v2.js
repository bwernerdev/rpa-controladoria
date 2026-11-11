/* Bookmarklet ME: Relatorio de Requisicao - Friboi V2.
   Roda no navegador ja logado (sem login automatizado).
   Na lista de relatorios: abre o relatorio (clique de novo no favorito).
   No formulario: todas as filiais, Empresa FDS, inicio = hoje-60 dias, fim = hoje, Processar.
   Gerar o favorito de uma linha: .\scripts\gerar_favorito_me.ps1 */
(function () {
  var DIAS = 60, EMPRESA = 'FDS', NOME = 'Relatório de Requisição - Friboi V2';
  var dry = window.__ME_DRY === true;
  var ft = 'ctl00$conteudo$wzdExtracao$frmFiltros$';
  var ot = window[ft + 'Filial_optiontransfer'];
  if (!ot) {
    var link = Array.prototype.slice.call(document.querySelectorAll('a')).filter(function (a) {
      return a.textContent.trim() === NOME;
    })[0];
    if (!link) { alert('Abra https://www.me.com.br/ME/ExtracaoRelatorio.aspx (logado) e clique de novo.'); return; }
    link.click();
    return;
  }
  function dd(d) {
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear();
  }
  var fim = new Date(), ini = new Date();
  ini.setDate(ini.getDate() - DIAS);
  ot.transferAllRight();
  var emp = document.getElementById('ctl00_conteudo_wzdExtracao_frmFiltros_Empresa');
  var opt = Array.prototype.slice.call(emp.options).filter(function (o) { return o.text.trim() === EMPRESA; })[0];
  if (!opt) { alert('Empresa ' + EMPRESA + ' nao encontrada.'); return; }
  emp.value = opt.value;
  emp.dispatchEvent(new Event('change', { bubbles: true }));
  var id = 'ctl00_conteudo_wzdExtracao_frmFiltros_hd_ctl00_conteudo_wzdExtracao_frmFiltros_';
  [['DataInicio', ini], ['DataFim', fim]].forEach(function (c) {
    var el = document.getElementById(id + c[0]);
    el.value = dd(c[1]);
    el.dispatchEvent(new Event('change', { bubbles: true }));
  });
  var n = document.getElementById('Filial').options.length;
  if (!n) { alert('Nenhuma filial selecionada.'); return; }
  var btn = Array.prototype.slice.call(document.querySelectorAll('button')).filter(function (b) {
    return b.textContent.trim() === 'Processar';
  })[0];
  if (dry) { console.log('DRY', n, dd(ini), dd(fim), emp.value); return { filiais: n, ini: dd(ini), fim: dd(fim), emp: emp.options[emp.selectedIndex].text, btn: !!btn }; }
  btn.click();
})();
