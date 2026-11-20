
// TEMP v131 medição (remover depois)
function perfV131() {
  var out = {}, t = function (nome, f) { var t0 = Date.now(); var r; try { r = f(); } catch (e) { r = 'ERR ' + e.message; } out[nome] = Date.now() - t0; return r; };
  var q = 'Celio';
  t('conf', function () { return conferencias(q).pendentes.length; });
  t('rebaixas', function () { return rebaixas().vigentes.length; });
  t('apps', function () { return apps(q).length; });
  t('precosOff', function () { return precosOff(q).itens.length; });
  t('psvResumo', function () { return psvResumo_(q); });
  t('pontoTopo', function () { return pontoTopo_(q); });
  t('v9Resumo', function () { var r = {}; v9Resumo_(q, r); return r; });
  t('agc', function () { return agcCalc_(); });
  t('HOME_TOTAL', function () { return resumoHome_(q); });
  t('HOME_glaucia', function () { return resumoHome_('Gláucia'); });
  t('decGondola', function () { return decGondolaVencidas_(); });
  var m = t('decMapa', function () { return decMapa_(); });
  t('decReb', function () { return decRebCartoes_(m).length; });
  t('decNovo', function () { return decNovoCartoes_(m).length; });
  t('decCont', function () { return decContCartoes_().length; });
  t('decCod', function () { return decCodCartoes_(m).length; });
  t('decPode', function () { var pmc = {}, n = 0; Object.keys(m).forEach(function (k) { try { decPodeMudar_(k, m[k], pmc); n++; } catch (e) {} }); return n; });
  var sv = t('semVenda', function () { return produtoSemVenda('Gláucia'); });
  out.semVendaKB = Math.round(JSON.stringify(sv).length / 1024);
  try { out.semVendaN = (sv.itens || []).length; out.semVendaKeys = Object.keys(sv).join(','); out.itemKeys = sv.itens && sv.itens[0] ? Object.keys(sv.itens[0]).join(',') : ''; } catch (e) {}
  Logger.log(JSON.stringify(out));
  return out;
}
