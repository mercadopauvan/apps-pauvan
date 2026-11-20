// ===== CONSUMO INTERNO – Mercado Pauvan =====
// Página de bipagem (celular) + planilha de lançamentos.
// Abas: LANCAMENTOS (registros) e CADASTRO (cópia diária do cadastros ativos.xlsx).

var ID_CADASTRO_XLSX = '1swZ2ztQaM4FktQ2Slaihyh69JXzf86vq'; // ARQUIVOS CLAUDE / cadastros ativos.xlsx
var ABA_LANC = 'LANCAMENTOS';
var ABA_CAD = 'CADASTRO';
var ABA_ATALHOS = 'ATALHOS';
var ABA_ACESSOS = 'ACESSOS';
var ABA_AVARIAS = 'AVARIAS';
var NOME_PLANILHA = 'APPS PAUVAN';
var ABA_REB = 'REBAIXAS';
var CAB_REB = ['SKU', 'DESCRICAO', 'DEMANDA', 'DE', 'POR', 'VALIDADE', 'QTD', 'STATUS', 'ENCERRADA EM', 'MOTIVO', 'QUEM', 'SOBRA', 'ID'];
var ABA_SETORES = 'SETORES';
var ABA_CONT = 'CONTAGEM DIA D';
var CAB_CONT = ['DIA D', 'BLOCO', 'RESPONSAVEL', 'SKU', 'PRODUTO', 'AG', 'QTD GONDOLA', 'QUEM CONTOU', 'HORA', 'ID'];
var CAB_AVARIAS = ['DATA', 'HORA', 'QUEM', 'MOTIVO', 'VALIDADE', 'CODIGO BIPADO', 'SKU', 'DESCRICAO', 'QUANTIDADE', 'UNIDADE', 'FOTO 1', 'FOTO 2', 'FORNECEDOR', 'CUSTO', 'STATUS', 'ID'];
var CAB_ATALHOS = ['QUEM', 'TIPO', 'GRUPO', 'BOTAO', 'SKU'];
var ABA_PESSOAS = 'PESSOAS';
// Lista de pessoas: vem da aba PESSOAS (um nome por linha). Para incluir alguém, basta escrever o nome numa linha nova.
// v47: aba GRUPOS (EMPRESA -> GRUPO): notas da mesma empresa, ou de empresas que chegam juntas, aparecem agrupadas na conferência
function gruposForn_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('GRUPOS'), m = {};
  if (!sh) {
    sh = ss.insertSheet('GRUPOS');
    sh.getRange(1, 1, 5, 2).setValues([['EMPRESA', 'GRUPO'], ['AGIL', 'ÁGIL / DISNORTE'], ['DISNORTE', 'ÁGIL / DISNORTE'], ['DAC', 'DAC / ON'], ['ON', 'DAC / ON']]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff'); sh.setFrozenRows(1);
  }
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (l) { var e = semAcento_(String(l[0] || '')).toUpperCase().trim(); if (e && l[1]) m[e] = String(l[1]).trim(); });
  return m;
}
function criarAbaPessoas_(ss) {
  var sh = ss.getSheetByName(ABA_PESSOAS);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_PESSOAS);
  var seed = [['NOME'], ['Amanda'], ['Ana Paula'], ['Celio'], ['Cristiano'], ['Mateus'], ['Ney'], ['Regiane']];
  sh.getRange(1, 1, seed.length, 1).setValues(seed);
  sh.getRange(1, 1).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function pessoas_() {
  var sh = criarAbaPessoas_(SpreadsheetApp.getActive());
  var n = sh.getLastRow(), out = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (l) { var s = String(l[0]).trim(); if (s && out.indexOf(s) < 0) out.push(s); });
  out.sort(function (a, b) { return a.localeCompare(b, 'pt-BR'); });
  return out;
}
// Renomear uma pessoa levando junto PIN, acessos, atalhos e setores. Edite os dois nomes e execute no editor.
function renomearPessoa() { renomear_('Célio', 'Celio'); }
function renomear_(de, para) {
  var ss = SpreadsheetApp.getActive(), props = PropertiesService.getScriptProperties();
  var h = props.getProperty('PIN_' + de);
  if (h) {
    // o PIN é calculado com o nome; guardamos o hash antigo sob o nome novo junto com o nome de origem para validar
    props.setProperty('PIN_' + para, h); props.setProperty('PINDE_' + para, de); props.deleteProperty('PIN_' + de);
  }
  [[ABA_ACESSOS, 1], [ABA_ATALHOS, 1], [ABA_SETORES, 2], [ABA_PESSOAS, 1]].forEach(function (a) {
    var sh = ss.getSheetByName(a[0]); if (!sh || sh.getLastRow() < 2) return;
    var r = sh.getRange(2, a[1], sh.getLastRow() - 1, 1), v = r.getValues(), mud = false;
    v.forEach(function (l) { if (String(l[0]).trim() === de) { l[0] = para; mud = true; } });
    if (mud) r.setValues(v);
  });
}
var CAB_LANC = ['DATA', 'HORA', 'QUEM', 'TIPO', 'CODIGO BIPADO', 'SKU', 'DESCRICAO', 'QUANTIDADE', 'UNIDADE', 'STATUS', 'ID'];

// ---------- Configuração inicial (rodar uma vez) ----------
function setup() {
  var ss = SpreadsheetApp.getActive();
  var lanc = ss.getSheetByName(ABA_LANC) || ss.insertSheet(ABA_LANC);
  if (lanc.getLastRow() === 0) {
    lanc.getRange(1, 1, 1, CAB_LANC.length).setValues([CAB_LANC]).setFontWeight('bold');
    lanc.setFrozenRows(1);
  }
  var cad = ss.getSheetByName(ABA_CAD) || ss.insertSheet(ABA_CAD);
  criarAbaAtalhos_(ss);
  criarAbaAcessos_(ss);
  criarAbaAvarias_(ss);
  criarAbaRebaixas_(ss);
  criarAbaSetores_(ss);
  criarAbaPessoas_(ss);
  if (ss.getName() !== NOME_PLANILHA) ss.rename(NOME_PLANILHA);
  organizarPastas_(ss);
  var padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 2) ss.deleteSheet(padrao);
  // gatilho diário 07h para atualizar o CADASTRO
  var tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'atualizarCadastro'; });
  if (!tem) ScriptApp.newTrigger('atualizarCadastro').timeBased().atHour(7).everyDays(1).create();
  atualizarCadastro();
}

// ---------- Aba ATALHOS (grupos de botões por pessoa) ----------
function criarAbaAtalhos_(ss) {
  var sh = ss.getSheetByName(ABA_ATALHOS);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_ATALHOS);
  var seed = [
    CAB_ATALHOS,
    ['Cristiano', 'MERCADO', 'LUVAS', 'M (N8 c/100)', '40332'],
    ['Cristiano', 'MERCADO', 'LUVAS', 'M (N8 c/50)', '44993'],
    ['Cristiano', 'MERCADO', 'LUVAS', 'G (N9 c/100)', '44443'],
    ['Cristiano', 'MERCADO', 'LUVAS', 'G (N9 c/50)', '45392'],
    ['Cristiano', 'MERCADO', 'LUVAS', 'GG (N10 c/100)', '47415'],
    ['Cristiano', 'MERCADO', 'SACOLAS', 'P 30x45', '18193'],
    ['Cristiano', 'MERCADO', 'SACOLAS', 'M 40x50', '17665'],
    ['Cristiano', 'MERCADO', 'SACOLAS', 'G 50x60', '37353'],
    ['Cristiano', 'MERCADO', 'SACOLAS', 'Carvão 60x75', '4057'],
    ['Cristiano', 'MERCADO', 'SACOLAS', 'Carvão 60x80', '34240'],
    ['Cristiano', 'MERCADO', 'BANDEJAS', 'CRL01', '14348'],
    ['Cristiano', 'MERCADO', 'BANDEJAS', 'CRL02', '13184'],
    ['Cristiano', 'MERCADO', 'BANDEJAS', 'CRL03', '37424']
  ];
  sh.getRange(1, 1, seed.length, 5).setNumberFormat('@').setValues(seed);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// ---------- Aba ACESSOS (quais apps cada pessoa vê) ----------
// Pessoa sem linha na aba não vê nenhum app. Apps: MERCADO, PROPRIO, AVARIA, REBAIXAS.
function criarAbaAcessos_(ss) {
  var sh = ss.getSheetByName(ABA_ACESSOS);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_ACESSOS);
  sh.getRange(1, 1, 4, 2).setValues([['QUEM', 'APP'], ['Celio', 'MERCADO'], ['Celio', 'PROPRIO'], ['Celio', 'AVARIA']]);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function apps(quem) {
  var sh = criarAbaAcessos_(SpreadsheetApp.getActive());
  var n = sh.getLastRow(), out = [];
  var q = String(quem || '').trim().toLowerCase();
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) {
    if (String(l[0]).trim().toLowerCase() === q) { var a = String(l[1]).trim().toUpperCase().replace('Ó', 'O'); if (/^SEMVENDA/.test(semAcento_(a))) a = 'SEMVENDA'; if (a && out.indexOf(a) < 0) out.push(a); } // v87: SEMVENDA com abas (ver PRODUTO SEM VENDA)
  });
  return out; // sem linha na aba ACESSOS = sem acesso a nada
}

// ---------- Aba AVARIAS ----------
function criarAbaAvarias_(ss) {
  var sh = ss.getSheetByName(ABA_AVARIAS);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_AVARIAS);
  sh.getRange(1, 1, 1, CAB_AVARIAS.length).setValues([CAB_AVARIAS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
// Pasta do app no Drive: ARQUIVOS CLAUDE / APPS PAUVAN (a planilha fica dentro dela)
function pastaApp_(ss) {
  var arq = DriveApp.getFileById(ss.getId());
  var pais = arq.getParents();
  var pai = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  if (pai.getName() === NOME_PLANILHA) return pai;
  var it = pai.getFoldersByName(NOME_PLANILHA);
  var pasta = it.hasNext() ? it.next() : pai.createFolder(NOME_PLANILHA);
  arq.moveTo(pasta);
  return pasta;
}
function organizarPastas_(ss) {
  var pasta = pastaApp_(ss);
  ['AVARIAS', 'RELATORIOS FORNECEDOR'].forEach(function (n) { if (!pasta.getFoldersByName(n).hasNext()) pasta.createFolder(n); });
}
function pastaAvarias_(mesAno) {
  var base = pastaApp_(SpreadsheetApp.getActive());
  var it = base.getFoldersByName('AVARIAS');
  var pasta = it.hasNext() ? it.next() : base.createFolder('AVARIAS');
  var it2 = pasta.getFoldersByName(mesAno);
  return it2.hasNext() ? it2.next() : pasta.createFolder(mesAno);
}
function salvarAvaria(reg) {
  var dono = quemDoToken_(reg.token);
  if (!dono || dono !== String(reg.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var motivo = String(reg.motivo || '').toUpperCase();
  if (motivo !== 'VENCIDO' && motivo !== 'AVARIA') return { ok: false, erro: 'informe se foi vencido ou avaria' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = criarAbaAvarias_(ss);
    var agora = new Date();
    var tz = Session.getScriptTimeZone();
    var id = reg.id || Utilities.getUuid();
    var nCol = CAB_AVARIAS.length;
    if (sh.getLastRow() > 1) {
      var ids = sh.getRange(2, nCol, sh.getLastRow() - 1, 1).getValues().map(String);
      if (ids.indexOf(id) >= 0) return { ok: true, id: id, dup: true };
    }
    var pasta = null;
    function gravaFoto_(dataUrl, sufixo) {
      if (!dataUrl || !/^data:image\//.test(dataUrl)) return '';
      var m = dataUrl.match(/^data:(image\/\w+);base64,(.*)$/);
      if (!m) return '';
      var nome = Utilities.formatDate(agora, tz, 'yyyyMMdd_HHmm') + '_' + motivo + '_' + String(reg.sku || reg.codigo || '') + '_' + String(reg.desc || '').replace(/[^A-Za-z0-9]+/g, '-').slice(0, 40) + '_' + sufixo + '.jpg';
      if (!pasta) pasta = pastaAvarias_(Utilities.formatDate(agora, tz, 'yyyy-MM'));
      return pasta.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nome)).getUrl();
    }
    var urlFoto = gravaFoto_(reg.foto, motivo === 'VENCIDO' ? 'validade' : 'dano');
    var urlFoto2 = gravaFoto_(reg.foto2, 'produtos');
    var qtd = Number(String(reg.qtd).replace(',', '.'));
    sh.appendRow([
      Utilities.formatDate(agora, tz, 'dd/MM/yyyy'),
      Utilities.formatDate(agora, tz, 'HH:mm'),
      String(reg.quem || ''),
      motivo,
      String(reg.validade || ''),
      "'" + String(reg.codigo || ''),
      "'" + String(reg.sku || ''),
      String(reg.desc || ''),
      qtd,
      String(reg.un || ''),
      urlFoto, urlFoto2,
      '', '',
      'PENDENTE',
      id
    ]);
    marcarOrigem_(sh, sh.getLastRow(), 1, 16);
    if (reg.vsId) { try { vsMarcarRetirado_(String(reg.vsId), id, qtd, String(reg.quem || ''), String(reg.validade || '')); } catch (e) { } } // v79: veio da lista Vencidos da semana
    return { ok: true, id: id, foto: !!urlFoto, foto2: !!urlFoto2 };
  } finally {
    lock.releaseLock();
  }
}
function ultimosAvarias_(quem) {
  var sh = criarAbaAvarias_(SpreadsheetApp.getActive());
  var n = sh.getLastRow();
  if (n < 2) return [];
  var ini = Math.max(2, n - 60);
  var vals = sh.getRange(ini, 1, n - ini + 1, CAB_AVARIAS.length).getValues();
  var out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < 8; i--) {
    var l = vals[i];
    if (quem && String(l[2]) !== quem) continue;
    if (String(l[14]) !== 'PENDENTE') continue;
    out.push({ data: fmt_(l[0], 'dd/MM/yyyy'), hora: fmt_(l[1], 'HH:mm'), quem: String(l[2]), tipo: 'AVARIA', motivo: String(l[3]), validade: fmt_(l[4], 'dd/MM/yyyy'), desc: String(l[7]), qtd: Number(l[8]), un: String(l[9]), foto: String(l[10]), foto2: String(l[11]), id: String(l[15]) });
  }
  return out;
}
function excluirAvaria_(id) {
  var sh = criarAbaAvarias_(SpreadsheetApp.getActive());
  var n = sh.getLastRow();
  if (n < 2) return { ok: false };
  var nCol = CAB_AVARIAS.length;
  var ids = sh.getRange(2, nCol, n - 1, 1).getValues().map(String);
  var i = ids.indexOf(String(id));
  if (i < 0) return { ok: false };
  var lin = sh.getRange(i + 2, 1, 1, nCol).getValues()[0];
  if (String(lin[14]) !== 'PENDENTE') return { ok: false, motivo: 'já lançado' };
  [lin[10], lin[11]].forEach(function (url) {
    var m = String(url || '').match(/\/d\/([A-Za-z0-9_-]+)/);
    if (m) { try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) {} }
  });
  sh.deleteRow(i + 2);
  return { ok: true };
}

// ---------- Aba REBAIXAS (rebaixas vigentes para encerrar pelo app) ----------
function criarAbaRebaixas_(ss) {
  var sh = ss.getSheetByName(ABA_REB);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_REB);
  var seed = [CAB_REB,
    ['2879', 'MAIONESE HELLMANN´S PET 500G', 'MOLHOS E CONDIMENTOS', 15.49, 13.99, '30/10/2026', '120', 'VIGENTE', '', '', '', '', Utilities.getUuid()],
    ['17433', 'SHAMPOO HUGGIES CAMOMILA 200ML', 'HIGIENE E CUIDADOS PESSOAIS', 27.99, 19.99, '30/10/2026', '5', 'VIGENTE', '', '', '', '', Utilities.getUuid()],
    ['8637', 'INSETICIDA BAYGON AÇAO TOTAL SPRAY 360ML', 'INSETICIDAS', 26.99, 15.99, '31/10/2026', '5', 'VIGENTE', '', '', '', '', Utilities.getUuid()],
    ['46910', 'PETISCO PEDIGREE FILEZITOS CARNE 60G', 'LINHA PET', 8.49, 5.99, '29/10/2026', '6', 'VIGENTE', '', '', '', '', Utilities.getUuid()]
  ];
  sh.getRange(1, 1, seed.length, CAB_REB.length).setValues(seed);
  sh.getRange(2, 1, seed.length - 1, 1).setNumberFormat('@');
  sh.getRange(2, 6, seed.length - 1, 1).setNumberFormat('@');
  sh.getRange(1, 1, 1, CAB_REB.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
// ---------- Aba SETORES (demanda do Plenus -> pessoa responsável) ----------
function criarAbaSetores_(ss) {
  var sh = ss.getSheetByName(ABA_SETORES);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_SETORES);
  sh.getRange(1, 1, 3, 2).setValues([['DEMANDA', 'QUEM'], ['HIGIENE E CUIDADOS PESSOAIS', 'Celio'], ['INSETICIDAS', 'Celio']]);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function setores(quem) {
  var sh = criarAbaSetores_(SpreadsheetApp.getActive());
  var n = sh.getLastRow(), out = [];
  var q = String(quem || '').trim().toLowerCase();
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) {
    if (String(l[1]).trim().toLowerCase() === q) { var d = String(l[0]).trim().toUpperCase(); if (d && out.indexOf(d) < 0) out.push(d); }
  });
  return out;
}
// v57: produto -> seção / local / repositora (aba CADASTRO da planilha VALIDADE PRODUTOS) — usado em REBAIXAS e CONTAGEM DIA D
function mapaReposicao_() {
  var m = {};
  try {
    var sh = SpreadsheetApp.openById(ID_VALIDADE).getSheetByName('CADASTRO'), n = sh.getLastRow();
    if (n >= 2) sh.getRange(2, 1, n - 1, 5).getValues().forEach(function (l) { var k = normDesc_(l[0]), s = skuN_(l[4]); if (!l[3]) return; var o = { secao: String(l[1] || ''), local: String(l[2] || ''), rep: String(l[3] || '') }; if (k) m[k] = o; if (s) m['#' + s] = o; }); // v99: SKU
  } catch (e) {}
  return m;
}
function verTudo_(quem) { var q = semAcento_(String(quem || '')).toLowerCase().trim(); return q === 'celio' || q === 'ney'; }
function rebaixas() {
  var ult = Number(PropertiesService.getScriptProperties().getProperty('REB_SYNC') || 0);
  if (Date.now() - ult > 3 * 3600000) { try { sincronizarRebaixas(); } catch (e) {} }
  var sh = criarAbaRebaixas_(SpreadsheetApp.getActive());
  var n = sh.getLastRow();
  if (n < 2) return { vigentes: [], encerradas: [] };
  var vals = sh.getRange(2, 1, n - 1, 16).getValues(); // v-aviso: coluna P = aviso de promoção encerrada
  var vig = [], enc = [];
  vals.forEach(function (l) {
    if (!l[12]) return;
    var it = { sku: String(l[0]), desc: String(l[1]), demanda: String(l[2] || 'SEM DEMANDA').toUpperCase(), de: Number(l[3]) || 0, por: Number(l[4]) || 0, validade: fmt_(l[5], 'dd/MM/yyyy'), qtd: String(l[6]), status: String(l[7]), em: fmt_(l[8], 'dd/MM/yyyy HH:mm'), motivo: String(l[9]), quem: String(l[10]), sobra: String(l[11]), id: String(l[12]), aviso: String(l[15] || '') };
    if (it.status === 'VIGENTE') vig.push(it); else enc.push(it);
  });
  var mpR = mapaReposicao_(); vig.forEach(function (it) { var r = cadSku_(mpR, it.sku, it.desc); it.rep = r && r.rep ? r.rep : (quemDaDemanda_(it.demanda) || 'SEM CADASTRO'); /* v129: SETORES pela demanda original */ it.local = r ? r.local : ''; it.secao = r ? r.secao : ''; }); // v57
  vig.sort(function (a, b) { return a.validade.split('/').reverse().join('') < b.validade.split('/').reverse().join('') ? -1 : 1; });
  var tEm = function (x) { var m = String(x.em || '').match(/(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?/); return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime() : 0; };
  enc.sort(function (a, b) { return tEm(b) - tEm(a); }); // v102: encerradas recentemente em ordem cronológica (a mais recente primeiro)
  return { vigentes: vig, encerradas: enc.slice(0, 8) };
}
// ---------- Rebaixas automáticas (v26) ----------
// Fonte: promocoes vigentes.xlsx (Plenus, Drive/BD) cruzado com a planilha REBAIXAS (abas por mês).
// Regra: promoção INDIVIDUAL ATIVA (fora das campanhas da semana e da Scanntech) de produto que está na REBAIXAS
// com a validade igual ao fim da promoção (tolerância de 2 dias). Entra sozinha no app como VIGENTE.
// Ela NÃO sai porque sumiu do arquivo: só sai por ACABOU / VENCEU (funcionária) ou, na noite do fim,
// pelo fechamento automático "SEM REGISTRO" (ninguém registrou no app).
var ID_PROMO_XLSX = '1sNMaT1tBXrRiOQ5RtgLQnmTyqwqgQZLX';      // ARQUIVOS CLAUDE / BD / promocoes vigentes.xlsx
var ID_PLANILHA_REBAIXAS = '1-T_2qy_aUMgfjVgYLJ_wy7-XlqrbgPu0PsTaqq3b4aU'; // planilha REBAIXAS (abas JAN26, FEV26 ...)
var MOTIVO_SEM_REGISTRO = 'SEM REGISTRO';

function normDesc_(s) {
  return semAcento_(String(s || '')).toUpperCase().replace(/[¬*"'´`]/g, '').replace(/\s+/g, ' ').trim();
}
function dataBR_(v) { // aceita Date, 'dd/MM/yyyy[ hh:mm:ss]' ou número de série do Excel; devolve 'dd/MM/yyyy'
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return m[1] + '/' + m[2] + '/' + m[3];
  if (/^\d+(\.\d+)?$/.test(s)) {
    var d = new Date(Math.round((Number(s) - 25569) * 86400000));
    return Utilities.formatDate(d, 'UTC', 'dd/MM/yyyy');
  }
  return '';
}
function diaNum_(br) { var p = String(br).split('/'); return p.length === 3 ? Date.UTC(+p[2], +p[1] - 1, +p[0]) / 86400000 : NaN; }

function lerXlsx_(id) { // primeira planilha do xlsx -> array de linhas {LETRA: valor}
  var partes = Utilities.unzip(DriveApp.getFileById(id).getBlob().setContentType('application/zip'));
  var sheetXml = '', sstXml = '';
  partes.forEach(function (p) {
    var n = p.getName();
    if (n === 'xl/worksheets/sheet1.xml') sheetXml = p.getDataAsString('UTF-8');
    else if (n === 'xl/sharedStrings.xml') sstXml = p.getDataAsString('UTF-8');
  });
  var sst = [], m, reSi = /<si>([\s\S]*?)<\/si>/g;
  while ((m = reSi.exec(sstXml)) !== null) sst.push(desXml_(m[1].replace(/<[^>]+>/g, '')));
  var rows = [], reRow = /<row [^>]*>([\s\S]*?)<\/row>/g;
  var reCel = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  while ((m = reRow.exec(sheetXml)) !== null) {
    var cel = {}, c; reCel.lastIndex = 0;
    while ((c = reCel.exec(m[1])) !== null) {
      var attrs = c[2] || '', inner = c[3] || '', t = (attrs.match(/t="(\w+)"/) || [])[1];
      var v = (inner.match(/<v>([^<]*)<\/v>/) || [])[1];
      if (t === 'inlineStr') v = desXml_(inner.replace(/<[^>]+>/g, ''));
      else if (t === 's') v = sst[Number(v)] || '';
      else if (v != null) v = desXml_(v);
      cel[c[1]] = v == null ? '' : v;
    }
    rows.push(cel);
  }
  return rows;
}

// lista de rebaixas detectadas hoje: [{sku, desc, de, por, validade, qtd}]
function detectarRebaixas_() {
  var promo = lerXlsx_(ID_PROMO_XLSX).slice(1); // B=SKU C=DESC D=VENDA E=PROMO F=QMAX I=FIM J=TIPO N=STATUS O=DESCR P=SCANNTECH
  // planilha REBAIXAS: todas as abas de mês; B=descrição, J=validade, K=OBS (quantidade)
  var reb = {};
  SpreadsheetApp.openById(ID_PLANILHA_REBAIXAS).getSheets().forEach(function (sh) {
    if (!/^[A-Z]{3}\d{2}$/.test(sh.getName())) return;
    var n = sh.getLastRow(); if (n < 2) return;
    sh.getRange(2, 1, n - 1, 11).getDisplayValues().forEach(function (l) {
      var d = normDesc_(l[1]), val = dataBR_(l[9]); if (!d || !val) return;
      var q = (String(l[10]).match(/(\d+)\s*UN/gi) || []).pop();
      (reb[d] = reb[d] || []).push({ val: val, qtd: q ? String(parseInt(q, 10)) : '' });
    });
  });
  var out = [], vistos = {};
  promo.forEach(function (l) {
    if (String(l.N).toUpperCase() !== 'ATIVA' || String(l.J).toUpperCase() !== 'INDIVIDUAL') return;
    if (/^(TRUE|1)$/i.test(String(l.P))) return;                       // Scanntech
    if (/^PROMO (QUI|SEG|FIM|FDS|QUA|DIA)/i.test(String(l.O))) return; // campanhas da semana
    var sku = String(Number(String(l.B).replace(/\D/g, '')) || ''); if (!sku) return;
    var fim = dataBR_(l.I), lista = reb[normDesc_(l.C)]; if (!fim || !lista) return;
    var alvo = null;
    lista.forEach(function (r) { var dif = Math.abs(diaNum_(r.val) - diaNum_(fim)); if (dif <= 2 && (!alvo || dif < alvo.dif)) alvo = { dif: dif, qtd: r.qtd }; });
    if (!alvo || vistos[sku + '|' + fim]) return;
    vistos[sku + '|' + fim] = 1;
    out.push({ sku: sku, desc: String(l.C).replace(/[¬*]+$/g, '').trim(), de: Number(l.D) || 0, por: Number(l.E) || 0, validade: fim, qtd: alvo.qtd });
  });
  return out;
}

// inclui na aba REBAIXAS do app as rebaixas novas (não mexe nas que já existem)
function sincronizarRebaixas() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, erro: 'ocupado' };
  try {
    var det = detectarRebaixas_();
    var sh = criarAbaRebaixas_(SpreadsheetApp.getActive());
    var n = sh.getLastRow(), tem = {}, ex = n >= 2 ? sh.getRange(2, 1, n - 1, 8).getDisplayValues() : [];
    ex.forEach(function (l) { tem[String(Number(l[0]) || l[0]) + '|' + dataBR_(l[5])] = 1; });
    try { demHistDiario_(); } catch (e) {} // v129
    var cad = mapaCadastro_(), novas = [], ctx = demCtx_(ex);
    det.forEach(function (r) {
      if (tem[r.sku + '|' + r.validade]) return;
      var c = cad[r.sku] || {}, dm = String(c.dem || '').trim().toUpperCase();
      if (demTemp_(dm)) { var o = demOriginal_(r.sku, ctx); dm = o ? o.dem : (dm ? 'QUEIMA DE ESTOQUE (ORIGEM?)' : 'SEM DEMANDA'); } // v129: nunca gravar a demanda temporária (QUEIMA / ESPECIAL FIM DE ANO)
      novas.push([r.sku, r.desc, dm, r.de, r.por, r.validade, r.qtd, 'VIGENTE', '', '', '', '', Utilities.getUuid()]);
    });
    ex.forEach(function (l, i) { // v129: linha que ficou "(ORIGEM?)" é corrigida quando a demanda original aparecer
      if (!/\(ORIGEM\?\)/.test(String(l[2]))) return; var o = demOriginal_(String(Number(l[0]) || l[0]), ctx, i); if (o) sh.getRange(i + 2, 3).setValue(o.dem);
    });
    if (novas.length) {
      sh.getRange(n + 1, 1, novas.length, 1).setNumberFormat('@');
      sh.getRange(n + 1, 6, novas.length, 1).setNumberFormat('@');
      sh.getRange(n + 1, 1, novas.length, CAB_REB.length).setValues(novas);
      avRebNova_(novas); // v103: push de rebaixa nova
    }
    PropertiesService.getScriptProperties().setProperty('REB_SYNC', String(Date.now()));
    return { ok: true, detectadas: det.length, novas: novas.length };
  } finally { lock.releaseLock(); }
}

// cópia de segurança do código (CODIGO.GS.TXT / INDEX.HTML.TXT) na pasta APPS PAUVAN — enviada pelo editor via POST
function salvarBackup_(nome, conteudo) {
  nome = String(nome || '');
  if (['CODIGO.GS.TXT', 'INDEX.HTML.TXT'].indexOf(nome) < 0 || !conteudo) return { ok: false, erro: 'nome ou conteúdo inválido' };
  var pasta = pastaApp_(SpreadsheetApp.getActive()), it = pasta.getFilesByName(nome);
  if (it.hasNext()) it.next().setContent(String(conteudo)); else pasta.createFile(nome, String(conteudo), MimeType.PLAIN_TEXT);
  return { ok: true, nome: nome, tamanho: String(conteudo).length };
}

// à noite: rebaixa que chegou ao fim e ninguém encerrou no app -> ENCERRADA / SEM REGISTRO
function fecharRebaixasDoDia() {
  try { sincronizarRebaixas(); } catch (e) {}
  var sh = criarAbaRebaixas_(SpreadsheetApp.getActive());
  var n = sh.getLastRow(); if (n < 2) return { ok: true, fechadas: 0 };
  var tz = Session.getScriptTimeZone(), hoje = diaNum_(Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'));
  var agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  // v-aviso (30/10/2026): a rebaixa que chegou ao fim NÃO fecha mais sozinha; ganha o aviso na coluna P (AVISO)
  // e continua VIGENTE no app até a funcionária marcar ACABOU ou VENCEU
  if (String(sh.getRange(1, 16).getValue()) !== 'AVISO') sh.getRange(1, 16).setValue('AVISO').setFontWeight('bold');
  var v = sh.getRange(2, 1, n - 1, 16).getDisplayValues(), fechadas = 0;
  v.forEach(function (l, i) {
    if (l[7] !== 'VIGENTE' || l[15]) return;
    var d = diaNum_(dataBR_(l[5]));
    if (!isNaN(d) && d <= hoje) {
      sh.getRange(i + 2, 16).setValue('PROMOÇÃO ENCERRADA EM ' + agora.replace(' ', ' ÀS '));
      marcarPendRebaixa_(sh, i + 2, 'NAO PRECISA', 'PENDENTE');
      fechadas++;
    }
  });
  return { ok: true, fechadas: fechadas };
}

// colunas N (PLENUS) e O (ARTE) da aba REBAIXAS: o que a tarefa agendada ainda precisa fazer (encerrar no Plenus; tirar a arte da pasta e da TV)
function marcarPendRebaixa_(sh, r, plenus, arte) {
  if (String(sh.getRange(1, 14).getValue()) !== 'PLENUS') sh.getRange(1, 14, 1, 2).setValues([['PLENUS', 'ARTE']]).setFontWeight('bold');
  sh.getRange(r, 14, 1, 2).setValues([[plenus, arte]]);
}
function linhaRebaixa_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return -1;
  var ids = sh.getRange(2, CAB_REB.length, n - 1, 1).getValues().map(String);
  var i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}
function encerrarRebaixa(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var motivo = String(p.motivo || '').toUpperCase();
  if (motivo !== 'ACABOU' && motivo !== 'VENCEU') return { ok: false, erro: 'informe o motivo' };
  var sh = criarAbaRebaixas_(SpreadsheetApp.getActive());
  var r = linhaRebaixa_(sh, p.id);
  if (r < 0) return { ok: false, erro: 'rebaixa não encontrada' };
  if (String(sh.getRange(r, 8).getValue()) !== 'VIGENTE') return { ok: false, erro: 'já encerrada' };
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  sh.getRange(r, 8, 1, 5).setValues([['ENCERRADA', agora, motivo, dono, motivo === 'VENCEU' ? String(p.sobra || '') : '']]);
  var avisoFim = String(sh.getRange(r, 16).getValue() || ''), arteAt = String(sh.getRange(r, 15).getValue() || '');
  if (avisoFim) marcarPendRebaixa_(sh, r, 'NAO PRECISA', /^(FEITO|NAO ACHADA)/.test(arteAt) ? arteAt : 'PENDENTE'); // v-aviso: promoção já tinha acabado
  else marcarPendRebaixa_(sh, r, 'PENDENTE', 'PENDENTE');
  marcarOrigem_(sh, r, 1, 16);
  avisoRebaixa_(sh, r, motivo, dono, String(p.sobra || ''), p.id); // v63: notificação no celular
  try { var lrb = sh.getRange(r, 1, 1, 6).getDisplayValues()[0]; rebMarcarFim_(lrb[1], lrb[5], motivo, String(p.sobra || ''), agora.slice(0, 10)); } catch (e) {} // v99: planilha REBAIXAS
  return { ok: true };
}
function reabrirRebaixa(id, token) {
  if (!quemDoToken_(token)) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = criarAbaRebaixas_(SpreadsheetApp.getActive());
  var r = linhaRebaixa_(sh, id);
  if (r < 0) return { ok: false };
  try { var lrb = sh.getRange(r, 1, 1, 6).getDisplayValues()[0]; rebMarcarFim_(lrb[1], lrb[5], '', '', ''); } catch (e) {} // v99
  sh.getRange(r, 8, 1, 5).setValues([['VIGENTE', '', '', '', '']]);
  marcarPendRebaixa_(sh, r, '', '');
  return { ok: true };
}

// grupos/botões de uma pessoa para um tipo (MERCADO, PROPRIO ou AVARIA)
function atalhos(quem, tipo) {
  var q = String(quem || '').trim().toLowerCase(), t = String(tipo || '').toUpperCase();
  var cache = CacheService.getScriptCache(), ck = 'AT2_' + q + '_' + t;
  var c = cache.get(ck);
  if (c) return JSON.parse(c);
  var ss = SpreadsheetApp.getActive();
  var sh = criarAbaAtalhos_(ss);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, 6).getValues(); // v107: coluna F = BLOCO
  var grupos = [], idx = {}, skus = {};
  var linhas = vals.filter(function (l) {
    var lq = String(l[0]).trim().toLowerCase(), lt = String(l[1]).trim().toUpperCase().replace('Ó', 'O');
    if (!lq || !l[3]) return false;
    if (lq !== 'todos' && lq !== q) return false;
    if (lt && lt !== 'TODOS' && lt !== t) return false;
    return true;
  });
  linhas.forEach(function (l) { var s = String(l[4] || '').replace(/\D/g, ''); if (s) skus[s] = null; });
  // uma única leitura do cadastro para todos os SKUs
  if (Object.keys(skus).length) {
    var cad = ss.getSheetByName(ABA_CAD), nc = cad.getLastRow();
    if (nc >= 2) {
      var cv = cad.getRange(2, 2, nc - 1, 3).getValues(); // SKU, DESCRICAO, UNIDADE
      for (var i = 0; i < cv.length; i++) { var s = String(cv[i][0]); if (s in skus && !skus[s]) skus[s] = { desc: String(cv[i][1]), un: String(cv[i][2]) }; }
    }
  }
  linhas.forEach(function (l) {
    var g = String(l[2] || 'ATALHOS').trim();
    if (!(g in idx)) { idx[g] = grupos.length; grupos.push({ grupo: g, itens: [] }); }
    var sku = String(l[4] || '').replace(/\D/g, '');
    var p = sku ? skus[sku] : null;
    grupos[idx[g]].itens.push({ bloco: String(l[5] || '').trim(), botao: String(l[3]), sku: sku, desc: p ? p.desc : String(l[3]), un: p ? p.un : 'UNIDADE', ok: !!p });
  });
  try { cache.put(ck, JSON.stringify(grupos), 600); } catch (e) {}
  return grupos;
}

// v107: preenche a coluna F (BLOCO) da aba ATALHOS e encurta o nome do botão (uma vez; depois o Celio ajusta na planilha)
function atBlocosPreencher_() {
  var sh = criarAbaAtalhos_(SpreadsheetApp.getActive()), n = sh.getLastRow(); if (n < 2) return { ok: true, n: 0 };
  sh.getRange(1, 6).setValue('BLOCO').setFontWeight('bold');
  var v = sh.getRange(2, 1, n - 1, 6).getValues(), out = [];
  var m = function (g, b) {
    var G = semAcento_(String(g)).toUpperCase().trim(); b = String(b);
    if (G === 'LUVAS') { if (/^Preta /i.test(b)) return [b.replace(/^Preta /i, ''), 'PRETAS']; if (/^Branca /i.test(b)) return [b.replace(/^Branca /i, ''), 'BRANCAS']; }
    if (G === 'SACOLAS') { if (/^Carv/i.test(b)) return [b.replace(/^Carv\S* /i, '').replace('\u00ac', ''), 'CARV\u00c3O']; if (/^Branca /i.test(b)) return [b.replace(/^Branca /i, ''), 'BRANCAS']; return [b, 'PL\u00c1STICAS']; }
    if (G === 'BANDEJAS') { if (/^CRL/.test(b)) return [b, 'CRL']; if (/^CFL/.test(b)) return [b, 'CFL']; if (/CR02/.test(b)) return [b.replace(' (absorvente)', ''), 'ABSORVENTE']; return [b, 'OUTRAS']; }
    if (/BOBINAS/.test(G)) { if (/^Picotada /.test(b)) return [b.replace('Picotada ', ''), 'PICOTADAS']; if (/^P\u00e3o /.test(b)) return [b.replace('P\u00e3o ', ''), 'P\u00c3O']; if (/^Estrela /.test(b)) return [b.replace('Estrela ', ''), 'ESTRELA']; if (/^Etiqueta balan/.test(b)) return [b.replace(/^Etiqueta balan\S* /, ''), 'BALAN\u00c7A']; return [b, 'BOBINAS']; }
    if (G === 'PADARIA') { if (/^P\u00c3O \d/.test(b)) return [b, 'P\u00c3ES']; if (/QUEIJO|CHIPA/.test(b)) return [b, 'P\u00c3O DE QUEIJO E CHIPA']; if (/ESFIRRA|DOG/.test(b)) return [b, 'ASSADOS']; return [b, 'INSUMOS']; }
    if (G === 'ACOUGUE') { if (/^(Faca|Serra)/.test(b)) return [b, 'FACAS E SERRAS']; if (/^Disco picador /.test(b)) return [b.replace('Disco picador ', ''), 'DISCOS DO PICADOR']; return [b, 'TEMPEROS']; }
    return [b, ''];
  };
  v.forEach(function (l) { if (String(l[5]).trim()) { out.push([l[3], l[5]]); return; } out.push(m(l[2], l[3])); });
  sh.getRange(2, 4, out.length, 1).setValues(out.map(function (x) { return [x[0]]; }));
  sh.getRange(2, 6, out.length, 1).setValues(out.map(function (x) { return [x[1]]; }));
  var novos = [['DESMOLDANTE 5L', '46876'], ['GERGELIM 1KG', '46826'], ['PARMES\u00c3O RALADO 1KG', '46901']], tem = {};
  v.forEach(function (l) { tem[String(l[0]) + '|' + String(l[4])] = 1; });
  var add = [];
  ['Celio', 'Ney'].forEach(function (q) { novos.forEach(function (x) { if (!tem[q + '|' + x[1]]) add.push([q, 'MERCADO', 'PADARIA', x[0], x[1], 'INSUMOS']); }); });
  if (add.length) { sh.getRange(n + 1, 5, add.length, 1).setNumberFormat('@'); sh.getRange(n + 1, 1, add.length, 6).setValues(add); }
  CacheService.getScriptCache().removeAll(['AT2_celio_MERCADO', 'AT2_ney_MERCADO']);
  return { ok: true, linhas: out.length, novos: add.length };
}
function produtoPorSku_(sku) {
  var cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD);
  var n = cad.getLastRow();
  var cel = cad.getRange(1, 2, n, 1).createTextFinder(String(sku)).matchEntireCell(true).findNext();
  if (!cel) return null;
  var l = cad.getRange(cel.getRow(), 1, 1, 6).getValues()[0];
  return { sku: String(l[1]), desc: String(l[2]), un: String(l[3]), bal: String(l[4]), preco: Number(l[5]) || 0 };
}

// busca por nome (produto solto na balança, sem etiqueta)
function buscarNome(texto, kg) {
  var soKg = String(kg) === 'true' || kg === true;
  var palavras = semAcento_(String(texto || '')).toUpperCase().split(/\s+/).filter(function (w) { return w.length >= 2; });
  if (!palavras.length) return [];
  var cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD);
  var n = cad.getLastRow();
  if (n < 2) return [];
  var vals = cad.getRange(2, 2, n - 1, 6).getValues(); // SKU, DESCRICAO, UNIDADE, BALANCA, PRECO, NAO_VENDER
  var vistos = {}, out = [];
  for (var i = 0; i < vals.length; i++) {
    var sku = String(vals[i][0]);
    if (vistos[sku] || String(vals[i][5]) === 'NV') continue;
    var d = semAcento_(String(vals[i][1])).toUpperCase();
    var ok = palavras.every(function (w) { return d.indexOf(w) >= 0; }) || (palavras.length === 1 && /^\d{3,7}$/.test(palavras[0]) && sku === palavras[0]); // v123.3: busca também pelo SKU
    if (!ok) continue;
    var ehKg = /QUILO/.test(String(vals[i][2]));
    if (String(kg) !== 'todos' && soKg !== ehKg) continue; // v62: consulta preço procura em todos
    vistos[sku] = 1;
    out.push({ sku: sku, desc: String(vals[i][1]), un: String(vals[i][2]) });
  }
  // produtos por kg primeiro (é o caso típico de produto solto)
  var p0 = palavras[0];
  function rank(x) { var d = semAcento_(x.desc).toUpperCase(); return (/QUILO/.test(x.un) ? 0 : 2) + (d.indexOf(p0) === 0 ? 0 : 1); }
  out.sort(function (a, b) { return rank(a) - rank(b) || a.desc.length - b.desc.length || a.desc.localeCompare(b.desc); });
  return out.slice(0, 15);
}

function semAcento_(s) {
  return s.normalize ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : s;
}

// ---------- PIN por pessoa ----------
function pinHash_(nome, pin) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'pauvan:' + nome + ':' + pin, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function pinStatus(quem) {
  var h = PropertiesService.getScriptProperties().getProperty('PIN_' + quem);
  return { ok: true, temPin: !!h };
}
function pinDefinir(quem, pin) {
  pin = String(pin || '').replace(/\D/g, '');
  if (pessoas_().indexOf(quem) < 0) return { ok: false, erro: 'pessoa desconhecida' };
  if (pin.length !== 4) return { ok: false, erro: 'o PIN precisa ter 4 números' };
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('PIN_' + quem)) return { ok: false, erro: 'PIN já definido; peça ao Celio para zerar' };
  props.setProperty('PIN_' + quem, pinHash_(quem, pin)); props.deleteProperty('PINDE_' + quem);
  return pinValidar(quem, pin);
}
function pinValidar(quem, pin) {
  pin = String(pin || '').replace(/\D/g, '');
  var h = PropertiesService.getScriptProperties().getProperty('PIN_' + quem);
  if (!h) return { ok: false, erro: 'PIN não definido', semPin: true };
  var base = PropertiesService.getScriptProperties().getProperty('PINDE_' + quem) || quem;
  if (h !== pinHash_(base, pin)) return { ok: false, erro: 'PIN incorreto' };
  return { ok: true, token: gerarToken_(quem, h) };
}
// v126 (17/11/2026): consulta para o Pauvan Conecta — recebe só o PIN (4 números) e devolve o nome da pessoa.
// Nunca devolve PIN, hash nem token. Trava: 5 erros do mesmo aparelho em 15 min -> esse aparelho bloqueado 15 min;
// 20 erros no total em 15 min -> consulta bloqueada para todos por 15 min.
var CPIN_ERR_APAR = 5, CPIN_ERR_GERAL = 20, CPIN_JANELA = 900;
function conectaPin_(p) {
  var c = CacheService.getScriptCache(), pin = String(p.pin == null ? '' : p.pin).trim();
  var ap = String(p.aparelho || '').replace(/[^\w-]/g, '').slice(0, 60), kA = ap ? 'CPIN_A_' + ap : '', kG = 'CPIN_G';
  if (c.get('CPIN_BLOQ') || (kA && c.get(kA + '_B'))) return { ok: false, erro: 'muitas tentativas erradas — tente de novo em 15 minutos', bloqueado: true };
  var falhou = function (msg) {
    var g = Number(c.get(kG) || 0) + 1; c.put(kG, String(g), CPIN_JANELA); if (g >= CPIN_ERR_GERAL) c.put('CPIN_BLOQ', '1', CPIN_JANELA);
    var resta = null;
    if (kA) { var a = Number(c.get(kA) || 0) + 1; c.put(kA, String(a), CPIN_JANELA); if (a >= CPIN_ERR_APAR) c.put(kA + '_B', '1', CPIN_JANELA); resta = Math.max(0, CPIN_ERR_APAR - a); }
    var r = { ok: false, erro: msg }; if (resta != null) r.tentativasRestantes = resta; if (resta === 0 || g >= CPIN_ERR_GERAL) r.bloqueado = true; return r;
  };
  if (!/^\d{4}$/.test(pin)) return falhou('o PIN precisa ter 4 números');
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) {}
  try {
    var props = PropertiesService.getScriptProperties().getProperties(), achou = [];
    pessoas_().forEach(function (q) { var h = props['PIN_' + q]; if (h && h === pinHash_(props['PINDE_' + q] || q, pin)) achou.push(q); });
    if (achou.length !== 1) return falhou(achou.length ? 'PIN não identifica uma pessoa só — peça ao Celio para trocar' : 'PIN não encontrado');
    if (kA) c.remove(kA);
    return { ok: true, nome: achou[0] };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
// Token assinado, sem validade: o navegador fica "de confiança" até trocar de usuário.
// Se o PIN for zerado/recriado, o token antigo deixa de valer (leva os 8 primeiros caracteres do hash).
function segredo_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('SEGREDO');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SEGREDO', s); }
  return s;
}
function assinar_(txt) {
  var raw = Utilities.computeHmacSha256Signature(txt, segredo_());
  return raw.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function gerarToken_(quem, hash) {
  var corpo = Utilities.base64EncodeWebSafe(quem, Utilities.Charset.UTF_8) + '.' + Date.now() + '.' + hash.slice(0, 8);
  return corpo + '.' + assinar_(corpo);
}
function quemDoToken_(token) {
  if (!token) return null;
  var p = String(token).split('.');
  if (p.length !== 4) return CacheService.getScriptCache().get('TK_' + token) || null; // tokens antigos (6h)
  var corpo = p[0] + '.' + p[1] + '.' + p[2];
  if (assinar_(corpo) !== p[3]) return null;
  var quem;
  try { quem = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString('UTF-8'); } catch (e) { return null; }
  var h = PropertiesService.getScriptProperties().getProperty('PIN_' + quem);
  if (!h || h.slice(0, 8) !== p[2]) return null;
  return quem;
}
// Para zerar o PIN de alguém: edite o nome abaixo e execute esta função no editor.
function resetarPin() {
  var nome = 'Ney';
  PropertiesService.getScriptProperties().deleteProperty('PIN_' + nome);
}

// ---------- v62: Consulta preço (app hospedado) ----------
// preço (aba CADASTRO), promoção vigente (promocoes vigentes.xlsx + REBAIXAS do app), saldo (saldo atual.txt) e seção/local (VALIDADE PRODUTOS / CADASTRO)
var ID_SALDO_TXT = '1YO5Ow6xNJBMnELmqHHETZeS-YU643oPu'; // ARQUIVOS CLAUDE / BD / saldo atual.txt
function dataHoraPromo_(v) {
  var s = String(v == null ? '' : v).trim(), m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return Utilities.parseDate(m[3] + '-' + m[2] + '-' + m[1] + ' ' + (m[4] || '00') + ':' + (m[5] || '00') + ':' + (m[6] || '00'), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  if (/^\d+(\.\d+)?$/.test(s)) { var d = new Date(Math.round((Number(s) - 25569) * 86400000)); return Utilities.parseDate(Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd HH:mm:ss'), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
  return null;
}
function tipoPromo_(l) {
  if (/^(TRUE|1)$/i.test(String(l.P))) return 'SCANNTECH';
  var o = semAcento_(String(l.O || '')).toUpperCase();
  if (/ESQUENTA/.test(o)) return 'ESQUENTA DIA D';
  if (/DIA D/.test(o)) return 'DIA D';
  if (/QUARTA VERDE/.test(o)) return 'QUARTA VERDE';
  if (/QUI\w* E SEX/.test(o)) return 'QUINTA E SEXTA DA CARNE';
  if (/SEG\w* E TER/.test(o)) return 'SEGUNDA E TERCA';
  if (/FDS|FIM DE SEMANA/.test(o)) return 'FIM DE SEMANA';
  if (/REBAIXA/.test(o)) return 'REBAIXA';
  return 'PROMOCAO';
}
function promosVigentes_() {
  var cache = CacheService.getScriptCache(), c = cache.get('PROMO_MAP2');
  if (c) return JSON.parse(c);
  var agora = new Date(), tz = Session.getScriptTimeZone(), m = {};
  lerXlsx_(ID_PROMO_XLSX).slice(1).forEach(function (l) {
    var st = String(l.N || '').toUpperCase(); if (/VENCIDA|CANCEL|INATIV|ENCERR/.test(st)) return;
    var ini = dataHoraPromo_(l.H), fim = dataHoraPromo_(l.I); if (!ini || !fim || agora < ini || agora > fim) return;
    var sku = String(Number(String(l.B).replace(/\D/g, '')) || ''); if (!sku) return;
    var por = Number(l.E) || 0; if (!por) return;
    var it = { de: Number(l.D) || 0, por: por, fim: Utilities.formatDate(fim, tz, 'dd/MM/yyyy'), tipo: tipoPromo_(l), min: Number(l.K) || 0, agrup: String(l.J || '').toUpperCase() === 'AGRUPADO' };
    if (!m[sku] || por < m[sku].por) m[sku] = it;
  });
  var em = ''; try { em = Utilities.formatDate(DriveApp.getFileById(ID_PROMO_XLSX).getLastUpdated(), tz, 'dd/MM HH:mm'); } catch (e) {}
  var out = { itens: m, em: em };
  try { cache.put('PROMO_MAP2', JSON.stringify(out), 600); } catch (e) {}
  return out;
}
function rebaixaVigente_(sku) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_REB); if (!sh) return null;
  var n = sh.getLastRow(); if (n < 2) return null;
  var r = null;
  sh.getRange(2, 1, n - 1, 8).getDisplayValues().forEach(function (l) {
    if (String(Number(l[0]) || l[0]) === String(sku) && l[7] === 'VIGENTE') r = { de: Number(String(l[3]).replace(',', '.')) || 0, por: Number(String(l[4]).replace(',', '.')) || 0, validade: dataBR_(l[5]) };
  });
  return r;
}
// v110.4: o saldo atual.txt às vezes vem com separador duplo ";;" (10/11/2026) — normaliza para ";"
function saldoTxt_(f) { var t = (f || DriveApp.getFileById(ID_SALDO_TXT)).getBlob().getDataAsString('ISO-8859-1'), h = t.slice(0, Math.max(0, t.indexOf('\n'))); return /;;/.test(h) ? t.replace(/;;/g, ';') : t; }
function saldoSku_(sku) {
  var f = DriveApp.getFileById(ID_SALDO_TXT), t = saldoTxt_(f);
  var m = t.match(new RegExp('(?:^|\\n)0*' + sku + ';[^\\n]*')), q = null;
  if (m) { var c = m[0].replace(/^\n/, '').split(';'); q = Number(String(c[6] || '').replace(/\./g, '').replace(',', '.')); if (isNaN(q)) q = null; }
  return { qtd: q, em: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM HH:mm') };
}
function consultaPreco(codigo) {
  var p = buscarProduto(codigo);
  if (!p.ok) return p;
  var out = { ok: true, codigo: p.codigo, sku: p.sku, desc: p.desc, un: p.un, bal: p.bal, preco: p.preco };
  var pv = promosVigentes_(), promo = pv.itens[p.sku] || null, reb = rebaixaVigente_(p.sku);
  out.promoEm = pv.em;
  if (promo) { if (reb && (promo.tipo === 'PROMOCAO' || promo.tipo === 'REBAIXA')) { promo.tipo = 'REBAIXA'; promo.validade = reb.validade; } out.promo = promo; if (promo.de) out.preco = promo.de; }
  else if (reb && reb.por) out.promo = { de: reb.de || p.preco, por: reb.por, tipo: 'REBAIXA', validade: reb.validade, fim: reb.validade };
  try { var s = saldoSku_(p.sku); out.saldo = s.qtd; out.saldoEm = s.em; } catch (e) {}
  var r = cadSku_(mapaReposicao_(), p.sku, p.desc); if (r) { out.secao = r.secao; out.local = r.local; out.rep = r.rep; }
  return out;
}
function rhChave_(quem) { return 'RH131_' + semAcento_(String(quem || '')).toLowerCase().replace(/[^a-z0-9]/g, ''); } // v131
function resumoHome_(quem) {
  var rhc = CacheService.getScriptCache(), rhk = rhChave_(quem), rhv = null; // v131: contadores da tela inicial guardados 60 s por pessoa
  try { rhv = rhc.get(rhk); } catch (e) {} if (rhv) { try { var rr = JSON.parse(rhv); rr.cache = 1; return rr; } catch (e) {} }
  var r = resumoHomeCalc_(quem);
  try { rhc.put(rhk, JSON.stringify(r), 60); } catch (e) {}
  return r;
}
function resumoHomeCalc_(quem) {
  var r = { ok: true }, q = semAcento_(String(quem || '')).toLowerCase().trim();
  try { r.conf = conferencias(quem).pendentes.length; } catch (e) {}
  try {
    var vt = verTudo_(quem);
    r.reb = rebaixas().vigentes.filter(function (it) { if (vt) return true; var a = semAcento_(String(it.rep || '')).toLowerCase().trim(); return a === q || a.split(/\s*\/\s*/).indexOf(q) >= 0; }).length;
  } catch (e) {}
  try { // v76: preços off-line
    if (apps(quem).indexOf('PRECOS_OFF') >= 0) {
      var po = precosOff(quem).itens, vt2 = verTudo_(quem);
      r.poff = po.filter(function (i) { if (i.status || i.oculto) return false; if (vt2) return true; var a = semAcento_(String(i.rep || '')).toLowerCase().trim(); return a === q || a.split(/\s*\/\s*/).indexOf(q) >= 0; }).length;
      r.poffConf = po.filter(function (i) { return i.status === 'CONFIRMADO'; }).length;
    }
  } catch (e) {}
  try { if (apps(quem).indexOf('SEMVENDA') >= 0) r.psv = psvResumo_(quem); } catch (e) {} // v87
  try { if (apps(quem).indexOf('PONTO_TOPO') >= 0) r.ponto = pontoTopo_(quem); } catch (e) {} // v119.5: quadro do ponto no topo (só quem tem PONTO_TOPO na aba ACESSOS)
  try { v9Resumo_(quem, r); } catch (e) {} // v109
  try { if (ehCelio_(quem)) { var agc = agcCalc_(); r.agc = { p: agc.pend.length, a: agc.agenda.length }; } } catch (e) {} // v123
  return r;
}

// ---------- v63: Notificações no celular (Firebase Cloud Messaging) — só no app hospedado ----------
var FCM_PROJETO = 'apps-pauvan';
var ABA_PUSH = 'PUSH', ABA_AVISOS = 'AVISOS';
var AVISO_REBAIXA_PARA = ['Celio']; // quem recebe a notificação de rebaixa encerrada (acrescente nomes aqui)
var APP_URL = 'https://mercadopauvan.github.io/apps-pauvan/';
function abaPush_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ABA_PUSH);
  if (!sh) { sh = ss.insertSheet(ABA_PUSH); sh.getRange(1, 1, 1, 4).setValues([['QUEM', 'TOKEN', 'ATUALIZADO', 'APARELHO']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function registrarPush_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var tk = String(p.fcm || '').trim(); if (tk.length < 20) return { ok: false, erro: 'token inválido' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = abaPush_(), n = sh.getLastRow(), lin = [dono, tk, new Date(), String(p.aparelho || '').slice(0, 120)];
    if (n >= 2) { var v = sh.getRange(2, 2, n - 1, 1).getValues(); for (var i = 0; i < v.length; i++) if (String(v[i][0]) === tk) { sh.getRange(i + 2, 1, 1, 4).setValues([lin]); return { ok: true }; } }
    sh.appendRow(lin);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function fcmAcesso_() {
  var c = CacheService.getScriptCache(), t = c.get('FCM_AT'); if (t) return t;
  var sa = JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SA') || 'null');
  if (!sa || !sa.private_key) throw new Error('chave do Firebase não instalada');
  var agora = Math.floor(Date.now() / 1000);
  var b64 = function (o) { return Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/, ''); };
  var corpo = b64({ alg: 'RS256', typ: 'JWT' }) + '.' + b64({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging', aud: 'https://oauth2.googleapis.com/token', iat: agora, exp: agora + 3600 });
  var ass = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(corpo, sa.private_key)).replace(/=+$/, '');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: corpo + '.' + ass }, muteHttpExceptions: true });
  var j = JSON.parse(r.getContentText());
  if (!j.access_token) throw new Error('token FCM: ' + r.getContentText().slice(0, 200));
  c.put('FCM_AT', j.access_token, 3000);
  return j.access_token;
}
// para: 'TODOS', um nome ou uma lista de nomes. extra: {abrir:'rebaixas'|'aviso', tag}
function enviarPush_(para, titulo, texto, extra) {
  extra = extra || {};
  var sh = abaPush_(), n = sh.getLastRow(); if (n < 2) return { ok: true, enviados: 0, aparelhos: 0 };
  var lista = para === 'TODOS' ? null : [].concat(para).reduce(function (acc, x) { return acc.concat([x], membrosGrupo_(x)); }, []).map(function (x) { /* v119.9: grupo -> quem atende o grupo */ return semAcento_(String(x)).toLowerCase().trim(); });
  var linhas = sh.getRange(2, 1, n - 1, 2).getValues(), reqs = [], idx = [], at = fcmAcesso_();
  var link = APP_URL + (extra.abrir ? '?abrir=' + extra.abrir : '');
  linhas.forEach(function (l, i) {
    var q = semAcento_(String(l[0])).toLowerCase().trim();
    if (!l[1] || (lista && lista.indexOf(q) < 0)) return;
    var wn = { icon: APP_URL + 'icone-192.png' }; if (extra.tag) wn.tag = extra.tag;
    var msg = { message: { token: String(l[1]), notification: { title: titulo, body: texto }, data: { abrir: String(extra.abrir || '') }, webpush: { notification: wn, fcm_options: { link: link } } } };
    reqs.push({ url: 'https://fcm.googleapis.com/v1/projects/' + FCM_PROJETO + '/messages:send', method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + at }, payload: JSON.stringify(msg), muteHttpExceptions: true });
    idx.push(i);
  });
  if (!reqs.length) return { ok: true, enviados: 0, aparelhos: 0 };
  var env = 0, falhas = 0, apagar = [], erro = '';
  UrlFetchApp.fetchAll(reqs).forEach(function (r, k) {
    var c = r.getResponseCode(), t = r.getContentText();
    if (c === 200) env++;
    else { falhas++; erro = t.slice(0, 200); if (c === 404 || /UNREGISTERED/.test(t)) apagar.push(idx[k] + 2); }
  });
  apagar.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return { ok: true, enviados: env, falhas: falhas, aparelhos: reqs.length, erro: erro };
}
function enviarAviso_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (semAcento_(dono).toLowerCase().trim() !== 'celio') return { ok: false, erro: 'só o Celio pode enviar avisos' };
  var tit = String(p.titulo || '').trim() || 'Aviso do Pauvan', tx = String(p.texto || '').trim();
  if (!tx) return { ok: false, erro: 'escreva a mensagem' };
  var para = Array.isArray(p.para) ? p.para.map(String).filter(String) : String(p.para || 'TODOS'); // v91: várias pessoas
  if (Array.isArray(para) && !para.length) return { ok: false, erro: 'escolha para quem' };
  var r = enviarPush_(para, tit, tx, { abrir: 'aviso' });
  try {
    var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ABA_AVISOS);
    if (!sh) { sh = ss.insertSheet(ABA_AVISOS); sh.getRange(1, 1, 1, 6).setValues([['DATA', 'DE', 'PARA', 'TITULO', 'MENSAGEM', 'APARELHOS']]).setFontWeight('bold'); sh.setFrozenRows(1); }
    sh.appendRow([new Date(), dono, [].concat(para).join(', '), tit, tx, r.enviados + '/' + r.aparelhos]);
  } catch (e) {}
  return r;
}
// v91: quantos celulares cada pessoa tem com os avisos ativados (tela Enviar aviso)
function avisoPessoas_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = abaPush_(), n = sh.getLastRow(), at = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { if (!l[1]) return; var q = semAcento_(String(l[0])).toLowerCase().trim(); at[q] = (at[q] || 0) + 1; });
  return { ok: true, ativos: at };
}
function avisoRebaixa_(sh, r, motivo, dono, sobra, id) {
  try {
    var l = sh.getRange(r, 1, 1, 6).getDisplayValues()[0];
    if (!avLig_('REB_ENCERRADA')) return; // v103
    enviarPush_(avNomes_('REB_ENCERRADA', AVISO_REBAIXA_PARA), 'Rebaixa encerrada — ' + motivo, l[1] + ' · validade ' + dataBR_(l[5]) + (motivo === 'VENCEU' && sobra ? ' · sobrou ' + sobra : '') + ' · ' + dono, { abrir: 'rebaixas', tag: 'reb-' + id });
  } catch (e) {}
}

// v63: rode no editor para autorizar o envio de notificações (e testar no celular do Celio)
function testarNotificacao() { Logger.log(JSON.stringify(enviarPush_('Celio', 'Teste do Pauvan', 'Se chegou, as notificações estão funcionando ✅', { abrir: 'aviso' }))); }
// ---------- Atualiza a aba CADASTRO lendo o xlsx do Drive (sem precisar converter) ----------
var COL = { ativo: 'B', sku: 'E', desc: 'F', demanda: 'G', bal: 'I', naoVender: 'J', barras: 'K', preco: 'L', gtin: 'AH', un: 'AM', mestre: 'D', qtdEmb: 'AL', gtinTrib: 'AI', aux: 'BB' }; // v118.7: GTIN tributável e códigos auxiliares também entram como código; // v104: mestre e qtd por embalagem de entrada (fardo/caixa -> unidade)

function atualizarCadastro() {
  var blob = DriveApp.getFileById(ID_CADASTRO_XLSX).getBlob().setContentType('application/zip');
  var partes = Utilities.unzip(blob);
  var sheetXml = '', sstXml = '';
  partes.forEach(function (p) {
    var n = p.getName();
    if (n === 'xl/worksheets/sheet1.xml') sheetXml = p.getDataAsString('UTF-8');
    else if (n === 'xl/sharedStrings.xml') sstXml = p.getDataAsString('UTF-8');
  });
  if (!sheetXml) throw new Error('sheet1.xml não encontrado no xlsx');
  // strings compartilhadas
  var sst = [];
  var reSi = /<si>([\s\S]*?)<\/si>/g, m;
  while ((m = reSi.exec(sstXml)) !== null) {
    var t = m[1].replace(/<[^>]+>/g, '');
    sst.push(desXml_(t));
  }
  var out = [['CODIGO', 'SKU', 'DESCRICAO', 'UNIDADE', 'BALANCA', 'PRECO', 'NAO_VENDER', 'DEMANDA', 'MESTRE', 'QTD_EMB']]; // v104
  var reRow = /<row [^>]*>([\s\S]*?)<\/row>/g;
  var reCel = /<c r="([A-Z]+)\d+"(?:[^>]*?t="(\w+)")?[^>]*>(?:<v>([^<]*)<\/v>)?/g;
  var primeira = true;
  while ((m = reRow.exec(sheetXml)) !== null) {
    if (primeira) { primeira = false; continue; } // cabeçalho
    var cel = {}, c;
    reCel.lastIndex = 0;
    while ((c = reCel.exec(m[1])) !== null) {
      var v = c[3] == null ? '' : c[3];
      cel[c[1]] = (c[2] === 's') ? (sst[Number(v)] || '') : v;
    }
    if (String(cel[COL.ativo]) !== 'T') continue;
    var nv = String(cel[COL.naoVender]) === 'T' ? 'NV' : ''; // 'Não Vender': fica fora da busca/bipagem, mas pode ser usado em atalho
    var sku = numStr_(cel[COL.sku]);
    if (!sku) continue;
    var desc = String(cel[COL.desc] || ''), un = String(cel[COL.un] || 'UNIDADE');
    var bal = String(cel[COL.bal]) === 'T' ? 'S' : '';
    var preco = Number(cel[COL.preco]) || 0;
    var dem = String(cel[COL.demanda] || '');
    var mestre = numStr_(cel[COL.mestre]) || sku, qtdEmb = Number(cel[COL.qtdEmb]) || 1; // v104
    var cods = {};
    [cel[COL.barras], cel[COL.gtin], cel[COL.gtinTrib], sku].forEach(function (x) { x = numStr_(x); if (x && x !== '0') cods[x] = 1; });
    String(cel[COL.aux] || '').split(/\D+/).forEach(function (x) { if (x.length >= 8 && !/^0+$/.test(x)) cods[x] = 1; }); // v118.7: códigos auxiliares
    Object.keys(cods).forEach(function (x) { out.push([x, sku, desc, un, bal, preco, nv, dem, mestre, qtdEmb]); });
  }
  var ss = SpreadsheetApp.getActive();
  var cad = ss.getSheetByName(ABA_CAD) || ss.insertSheet(ABA_CAD);
  cad.clearContents();
  cad.getRange(1, 1, out.length, 5).setNumberFormat('@');
  cad.getRange(1, 1, out.length, 9).setNumberFormat('@'); // v104: MESTRE como texto
  cad.getRange(1, 1, out.length, out[0].length).setValues(out);
  cad.setFrozenRows(1);
  PropertiesService.getScriptProperties().setProperty('CAD_ATUALIZADO', new Date().toISOString());
  try { demHistGravar_(out.slice(1)); } catch (e) {} // v129: última demanda "verdadeira" de cada SKU
  return out.length - 1;
}

function numStr_(v) {
  if (v == null || v === '') return '';
  var s = String(v);
  if (/e\+?\d+$/i.test(s)) s = Number(s).toFixed(0);
  return s.replace(/\.0+$/, '').replace(/\D/g, '');
}

function desXml_(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&').replace(/_x000D_/g, ' ').replace(/\s+/g, ' ').trim();
}

// ---------- Consulta ----------
// v106: item ligado a um mestre vendido por kg (ex.: LING TOSCANA SADIA KG -> LING TOSCANA PERD/SAD KG) confere por KG, não por peça
function unPl_(cad, sku) { var c = cad[sku] || {}, me = c.mestre && c.mestre !== String(c.sku || sku) ? cad[c.mestre] : null; if (me && (c.qtdEmb || 1) === 1 && /QUILO/.test(String(me.un))) return String(me.un); return c.un || ''; }
function mapaCadastro_() {
  var cache = CacheService.getScriptCache();
  var ss = SpreadsheetApp.getActive();
  var cad = ss.getSheetByName(ABA_CAD);
  var n = cad.getLastRow();
  if (n < 2) return {};
  var vals = cad.getRange(2, 1, n - 1, 10).getValues();
  var m = {};
  vals.forEach(function (l) { m[String(l[0])] = { sku: String(l[1]), desc: String(l[2]), un: String(l[3]), bal: String(l[4]), dem: String(l[7] || ''), mestre: String(l[8] || ''), qtdEmb: Number(l[9]) || 1 }; });
  return m;
}
// v112.1 (regra do Celio, 10/11/2026): na aba PRODUTOS da VALIDADE a UN é sempre UN (ou KG) e CHEGOU é a quantidade em UNIDADE DE VENDA que entrou no estoque.
// Fator da embalagem da nota (quantas unidades por CX/FD/DP...): vínculo aprendido do fornecedor > fator da nota (qTrib/qCom ou já gravado) > descrição da nota > cadastro (fardo/caixa com mestre, ou descrição de cadastro de caixa). Sem certeza: CONFERIR.
var VAL_UNS_ = /^(UN|UNID|UND|UN1|PC|PCS|PECA|PÇ)$/;
function valSku_(x) { return String(x == null ? '' : x).replace(/^'/, '').trim(); }
function valFatorItem_(it, cad, aprend, cnpj) { // it = linha da CONFERENCIA ITENS
  var sku = valSku_(it[2]), kv = chaveVinc_(cnpj, valSku_(it[18]), it[15]), fa = Number(aprend['F|' + kv]) || 0;
  if (fa > 1) return { f: fa, fonte: 'vínculo do fornecedor' };
  var fn = Number(String(it[8]).replace(',', '.')) || 0; if (fn > 1) return { f: fn, fonte: 'nota' };
  var fd = Number(fatorDesc_(it[15])) || 0; if (fd > 1) return { f: fd, fonte: 'descrição da nota' };
  var ub = unidadeBase_(cad, sku); if (ub && ub.fator > 1) return { f: ub.fator, fonte: 'cadastro' };
  var c = cad[sku]; if (c && /^(CX|FD|FARDO|CAIXA|DP|DISPLAY|PCT|PACOTE)\b/i.test(String(c.desc))) { var fc = Number(fatorDesc_(c.desc)) || 0; if (fc > 1) return { f: fc, fonte: 'cadastro' }; }
  return { f: 0, fonte: '' };
}
function valKg_(it, cad, sku) { var c = cad[sku] || {}; return /QUILO/.test(String(c.un || '')) || (!!it && String(it[5]).toUpperCase() === 'KG' && !c.un); }
// um lote da CONFERENCIA LANC (l) -> { un: 'UN' | 'KG' | 'CONFERIR', qtd, conv, motivo }
function valLote_(it, l, cad, aprend, cnpj) {
  var sku = valSku_(it ? it[2] : l[2]), forma = String(l[4]), qd = Number(String(l[5]).replace(',', '.')) || 0;
  if (forma === 'KG' || forma === 'KGCX') { // v123.4 (regra do Celio, 16/11/2026): produto vendido por PEÇA que chega em kg -> validade em peças
    var c0 = cadPorSku_(cad, sku) || cad[sku] || {}, naLista = !!(KGP_ || (KGP_ = kgPecas_()))[skuN_(sku)], vendeUn = !!c0.un && !/QUILO/.test(String(c0.un));
    if (naLista || vendeUn) {
      var pc = Number(String(l[14] == null ? '' : l[14]).replace(',', '.')) || 0;
      if (pc > 0) return { un: 'UN', qtd: Math.round(pc), conv: String(qd).replace('.', ',') + ' kg = ' + Math.round(pc) + ' peças (contadas)' };
      var pz = pesoDesc_(c0.desc || l[3]);
      if (pz > 0 && qd > 0) { var e = qd / pz, er = Math.round(e); if (er > 0 && Math.abs(e - er) <= 0.15) return { un: 'UN', qtd: er, conv: String(qd).replace('.', ',') + ' kg ÷ ' + (pz < 1 ? Math.round(pz * 1000) + ' g' : String(pz).replace('.', ',') + ' kg') + ' = ' + er + ' peças', est: 'ESTIMADO (' + String(qd).replace('.', ',') + ' kg ÷ ' + (pz < 1 ? Math.round(pz * 1000) + ' g' : String(pz).replace('.', ',') + ' kg') + ')' }; }
      if (naLista) return { un: 'CONFERIR', qtd: qd, conv: '', motivo: 'vendido por peça: chegou em kg sem as peças contadas (sem gramatura na descrição ou o peso não fecha peças inteiras)' };
    }
    return { un: 'KG', qtd: qd, conv: '' };
  }
  if (forma === 'PECAS') return { un: 'CONFERIR', qtd: qd, conv: '', motivo: 'produto de kg contado em peças' };
  var kg = valKg_(it, cad, sku), unN = it ? String(it[5]).toUpperCase() : 'UN', pacote = !!it && !VAL_UNS_.test(unN) && unN !== 'KG';
  var fl = Number(String(l[6]).replace(',', '.')) || 0, fi = it ? valFatorItem_(it, cad, aprend, cnpj) : { f: 0 }, fNota = it ? Number(String(it[8]).replace(',', '.')) || 0 : 0, r;
  var nomeEmb = pacote ? unN : 'CX', uFim = kg ? 'KG' : 'UN';
  if (forma === 'CAIXA') { var f = fl > 0 ? fl : fi.f; r = f > 0 ? { q: qd * f, conv: '1 ' + nomeEmb + ' = ' + f + ' ' + uFim } : { erro: 'contado em caixa sem a quantidade por caixa' }; }
  else if (!pacote || fNota > 1) r = { q: qd, conv: '' }; // contou em unidades (a nota já estava em unidades, ou a conferência comparou em unidades)
  else if (fi.f > 0) r = { q: qd * fi.f, conv: '1 ' + unN + ' = ' + fi.f + ' ' + uFim };
  else r = { erro: 'nota em ' + unN + ' sem a quantidade por embalagem' };
  if (r.erro) return { un: 'CONFERIR', qtd: qd, conv: '', motivo: r.erro };
  return { un: uFim, qtd: Math.round(r.q * 1000) / 1000, conv: r.conv };
}
// texto da conversão para a tela de conferência ('1 CX = 10 UN'; '⚠ ...' quando não há como converter)
function valConvTxt_(it, cad, aprend, cnpj) {
  var sku = valSku_(it[2]), unN = String(it[5]).toUpperCase(); if (unN === 'KG') return '';
  var kg = valKg_(it, cad, sku), fi = valFatorItem_(it, cad, aprend, cnpj), fNota = Number(String(it[8]).replace(',', '.')) || 0;
  if (VAL_UNS_.test(unN)) return fNota > 1 ? '1 CX = ' + fNota + (kg ? ' KG' : ' UN') : '';
  return fi.f > 0 ? '1 ' + unN + ' = ' + fi.f + (kg ? ' KG' : ' UN') : '⚠ sem quantidade por ' + unN + ' — conte em UNIDADES';
}
// pendência para o Celio: item cuja quantidade não deu para converter em unidade (aba VALIDADE CONFERIR)
function valPendencias_(nNF, forn, itens) {
  if (!itens.length) return;
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('VALIDADE CONFERIR');
  if (!sh) { sh = ss.insertSheet('VALIDADE CONFERIR'); sh.getRange(1, 1, 1, 8).setValues([['DATA', 'NF', 'FORNECEDOR', 'SKU', 'PRODUTO', 'QTD (como veio)', 'MOTIVO', 'STATUS']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  var n = sh.getLastRow(), tem = {}; if (n >= 2) sh.getRange(2, 2, n - 1, 4).getValues().forEach(function (x) { tem[String(x[0]) + '|' + String(x[2]) + '|' + String(x[3])] = 1; });
  var add = itens.filter(function (x) { return !tem[nNF + '|' + x.sku + '|' + x.prod]; }).map(function (x) { return [agoraBr_(), "'" + nNF, forn, "'" + x.sku, x.prod, x.qtd, x.motivo, 'PENDENTE']; });
  if (!add.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, add.length, 8).setValues(add);
  try { v9Push_('Celio', '⚠ Validade para conferir', add.length + ' item(ns) da NF ' + nNF + ' ' + forn + ' sem conversão para unidade: ' + add.map(function (x) { return x[4]; }).slice(0, 2).join(', '), 'aviso'); } catch (e) {}
}
// v104 (regra do Celio, 07/11/2026): fardo/caixa de entrada (SKU filho com mestre diferente e quantidade por embalagem > 1)
// vira a UNIDADE (o cadastro mestre) nas contagens de validade: nome da unidade e quantidade x fator. Sem mestre no cadastro, fica como está.
function unidadeBase_(cad, sku) {
  var c = cad[String(sku || '')]; if (!c) return null;
  var f = Number(c.qtdEmb) || 1, m = String(c.mestre || '');
  if (f <= 1 || !m || m === String(c.sku)) return null;
  var u = cad[m]; if (!u || !u.desc) return null;
  return { sku: u.sku, desc: u.desc, fator: f };
}

function buscarProduto(codigo) {
  codigo = String(codigo || '').replace(/\D/g, '');
  if (!codigo) return { ok: false };
  var cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD);
  var n = cad.getLastRow();
  var etiq = null;
  var tentativas = [codigo, codigo.replace(/^0+/, '')];
  // etiqueta de balança Prix: 2 + SKU(5) + 0 + total(5, centavos) + dígito
  if (codigo.length === 13 && codigo.charAt(0) === '2') {
    etiq = { sku: String(parseInt(codigo.substr(1, 5), 10)), total: parseInt(codigo.substr(7, 5), 10) / 100 };
    tentativas = [etiq.sku].concat(tentativas);
  }
  var cel = null;
  for (var i = 0; i < tentativas.length && !cel; i++) {
    if (!tentativas[i]) continue;
    cel = (etiq && i === 0 ? cad.getRange(1, 2, n, 1) : cad.getRange(1, 1, n, 1)).createTextFinder(tentativas[i]).matchEntireCell(true).findNext();
  }
  if (!cel) return { ok: false, codigo: codigo };
  var l = cad.getRange(cel.getRow(), 1, 1, 7).getValues()[0];
  if (String(l[6]) === 'NV') return { ok: false, codigo: codigo, naoVender: true };
  var r = { ok: true, codigo: codigo, sku: String(l[1]), desc: String(l[2]), un: String(l[3]), bal: String(l[4]), preco: Number(l[5]) || 0 };
  if (etiq && r.preco > 0 && etiq.total > 0) { r.etiqTotal = etiq.total; r.pesoSug = Math.round(etiq.total / r.preco * 1000) / 1000; }
  return r;
}

// ---------- Gravação ----------
function salvarLancamento(reg) {
  var dono = quemDoToken_(reg.token);
  if (!dono || dono !== String(reg.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(ABA_LANC);
    var agora = new Date();
    var tz = Session.getScriptTimeZone();
    var id = reg.id || Utilities.getUuid();
    // evita duplicar se o celular reenviar
    if (sh.getLastRow() > 1) {
      var ids = sh.getRange(2, CAB_LANC.length, sh.getLastRow() - 1, 1).getValues().map(String);
      if (ids.indexOf(id) >= 0) return { ok: true, id: id, dup: true };
    }
    var qtd = Number(String(reg.qtd).replace(',', '.'));
    // v66: usuário Delivery (consumo próprio) informa quem pediu a baixa — coluna L PEDIDO POR
    var ped = String(reg.pedidoPor || '').trim();
    if (['delivery', 'ney'].indexOf(semAcento_(dono).toLowerCase().trim()) >= 0 && /PR[OÓ]PRIO/i.test(String(reg.tipo || '')) && PEDIDO_POR.indexOf(ped) < 0) return { ok: false, erro: 'escolha quem pediu' };
    sh.appendRow([
      Utilities.formatDate(agora, tz, 'dd/MM/yyyy'),
      Utilities.formatDate(agora, tz, 'HH:mm'),
      String(reg.quem || ''),
      String(reg.tipo || ''),
      "'" + String(reg.codigo || ''),
      "'" + String(reg.sku || ''),
      String(reg.desc || ''),
      qtd,
      String(reg.un || ''),
      'PENDENTE',
      id,
      ped
    ]);
    marcarOrigem_(sh, sh.getLastRow(), 1, 12);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

var PEDIDO_POR = ['Celio', 'Amanda', 'Ana Paula', 'Paulo', 'Vanda']; // v66: quem pode pedir baixa ao Delivery
function ultimos(quem, tipo) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_LANC);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var ini = Math.max(2, n - 60);
  var vals = sh.getRange(ini, 1, n - ini + 1, CAB_LANC.length + 1).getValues(); // v66: + PEDIDO POR
  var out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < 8; i--) {
    var l = vals[i];
    if (quem && String(l[2]) !== quem) continue;
    if (tipo && String(l[3]) !== tipo && !(tipo === 'CONSUMO DO MERCADO' && String(l[3]) === 'PERDA PADARIA')) continue; // v128: perdas da padaria aparecem na Baixa do mercado
    if (String(l[9]) !== 'PENDENTE') continue; // já lançado no Plenus sai da lista
    out.push({ data: fmt_(l[0], 'dd/MM/yyyy'), hora: fmt_(l[1], 'HH:mm'), quem: String(l[2]), tipo: String(l[3]), desc: String(l[6]), qtd: Number(l[7]), un: String(l[8]), id: String(l[10]), pedidoPor: String(l[11] || '') });
  }
  return out;
}

function fmt_(v, pat) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), pat);
  return String(v == null ? '' : v);
}

function excluirLancamento(id, token, tipo) {
  if (!quemDoToken_(token)) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (tipo === 'AVARIA') return excluirAvaria_(id);
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_LANC);
  var n = sh.getLastRow();
  if (n < 2) return { ok: false };
  var ids = sh.getRange(2, CAB_LANC.length, n - 1, 1).getValues().map(String);
  var i = ids.indexOf(String(id));
  if (i < 0) return { ok: false };
  if (String(sh.getRange(i + 2, 10).getValue()) !== 'PENDENTE') return { ok: false, motivo: 'já lançado' };
  sh.deleteRow(i + 2);
  return { ok: true };
}


// v122: módulo CONFERÊNCIA TATI (CONFIRMAÇÃO, v64) excluído a pedido do Celio em 15/11/2026; a aba CONFIRMACAO TATI ficou oculta como histórico.

// ========== FOTOS DO SITE (v66) ==========
// Aba FOTOS SITE: produtos ativos no site DeliveryVip sem foto. Claude preenche a lista;
// quem tem acesso FOTOS tira 2-3 fotos de cada produto pelo app; as fotos vão para
// APPS PAUVAN / FOTOS SITE / <SKU - PRODUTO> no Drive. Claude trata e coloca no site.
var ABA_FS = 'FOTOS SITE';
var CAB_FS = ['SKU', 'PRODUTO', 'SECAO', 'PRECO', 'STATUS', 'QUEM', 'ENVIADO EM', 'FOTOS', 'PASTA', 'OBS', 'NO SITE EM'];
function criarAbaFS_(ss) {
  var sh = ss.getSheetByName(ABA_FS);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_FS);
  sh.getRange(1, 1, 1, CAB_FS.length).setValues([CAB_FS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function pastaFotosSite_(nome) {
  var base = pastaApp_(SpreadsheetApp.getActive());
  var it = base.getFoldersByName('FOTOS SITE');
  var pasta = it.hasNext() ? it.next() : base.createFolder('FOTOS SITE');
  var it2 = pasta.getFoldersByName(nome);
  return it2.hasNext() ? it2.next() : pasta.createFolder(nome);
}
function fotosSite(quem) {
  if (apps(quem).indexOf('FOTOS') < 0) return { ok: false, erro: 'sem acesso' };
  var sh = criarAbaFS_(SpreadsheetApp.getActive()), n = sh.getLastRow(), out = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, CAB_FS.length).getValues().forEach(function (l) {
    var sku = String(l[0] || '').trim(); if (!sku) return;
    var st = String(l[4] || '').trim().toUpperCase() || 'PENDENTE';
    if (st === 'NO SITE' || st === 'CANCELADO') return; // já resolvido: sai da lista do app
    out.push({ sku: sku, desc: String(l[1] || ''), secao: String(l[2] || ''), preco: Number(l[3]) || 0, status: st, quem: String(l[5] || ''), em: l[6] instanceof Date ? fmt_(l[6], 'dd/MM HH:mm') : String(l[6] || ''), fotos: Number(l[7]) || 0, obs: String(l[9] || '') });
  });
  return { ok: true, itens: out };
}
function salvarFotosSite(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('FOTOS') < 0) return { ok: false, erro: 'sem acesso' };
  var fotos = (p.fotos || []).filter(function (d) { return d && /^data:image\//.test(d); });
  if (!fotos.length) return { ok: false, erro: 'tire pelo menos a foto da frente' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = criarAbaFS_(SpreadsheetApp.getActive()), n = sh.getLastRow(), r = -1, l = null;
    var vals = n >= 2 ? sh.getRange(2, 1, n - 1, CAB_FS.length).getValues() : [];
    for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === String(p.sku)) { r = i + 2; l = vals[i]; break; }
    if (r < 0) return { ok: false, erro: 'produto não está na lista' };
    var agora = new Date(), tz = Session.getScriptTimeZone();
    var nomePasta = String(p.sku) + ' - ' + String(l[1] || '').replace(/[\\/:*?"<>|]+/g, '-').slice(0, 60);
    var pasta = pastaFotosSite_(nomePasta);
    var rot = ['frente', 'frente-2', 'codigo-barras'];
    fotos.slice(0, 3).forEach(function (d, k) {
      var m = d.match(/^data:(image\/\w+);base64,(.*)$/); if (!m) return;
      var nome = String(p.sku) + '_' + Utilities.formatDate(agora, tz, 'yyyyMMdd_HHmm') + '_' + (k + 1) + '-' + rot[k] + '.jpg';
      pasta.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nome));
    });
    sh.getRange(r, 5, 1, 5).setValues([['ENVIADO', dono, agora, fotos.length, pasta.getUrl()]]);
    sh.getRange(r, 7).setNumberFormat('dd/MM/yyyy HH:mm');
    marcarOrigem_(sh, r, 1, CAB_FS.length);
    return { ok: true, fotos: fotos.length };
  } finally {
    lock.releaseLock();
  }
}

// ========== PREÇOS OFF-LINE (v76) ==========
// Relatório 'precos offline.xlsx' (pasta BD). A repositora coloca a etiqueta impressa pela Nath e confirma no app:
// a linha vai para a aba PRECOS OFF como CONFIRMADO e o Celio recebe aviso no celular. O Claude ativa o preço no Plenus e marca ATIVADO.
// Acesso: app 'PRECOS_OFF' na aba ACESSOS.
var ID_PASTA_BD = '14OzVSgxjAo_7I36F2jqfMTbEVQCx4nU0';
var ID_PRECOS_OFF_XLSX = '1ZQjGK1MB0SsRPuuRpOjufAqwChjiz-wh';
var ABA_POFF = 'PRECOS OFF';
var CAB_POFF = ['SKU', 'DESCRICAO', 'NOVO PRECO', 'PRECO ATUAL', 'ALTERACAO', 'STATUS', 'QUEM', 'CONFIRMADO EM', 'ATIVADO EM', 'OBS'];
var AVISO_POFF_PARA = ['Celio']; // quem recebe o aviso de etiqueta colocada (acrescente 'Nath' quando ela tiver o app)
function criarAbaPoff_(ss) {
  var sh = ss.getSheetByName(ABA_POFF);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_POFF);
  sh.getRange(1, 1, 1, CAB_POFF.length).setValues([CAB_POFF]).setFontWeight('bold').setBackground('#0369a1').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}
function arquivoPrecosOff_() {
  try {
    var it = DriveApp.getFolderById(ID_PASTA_BD).getFilesByName('precos offline.xlsx'), f = null;
    while (it.hasNext()) { var x = it.next(); if (!x.isTrashed() && (!f || x.getLastUpdated() > f.getLastUpdated())) f = x; }
    if (f) return f;
  } catch (e) {}
  return DriveApp.getFileById(ID_PRECOS_OFF_XLSX);
}
function chavePoff_(sku, novo) { return String(sku).trim() + '|' + (Number(novo) || 0).toFixed(2); }
// lê o xlsx (sem converter) — cache pela data de atualização do arquivo
function lerPrecosOff_() {
  var f = arquivoPrecosOff_(), ver = f.getId() + '_' + f.getLastUpdated().getTime(), c = CacheService.getScriptCache();
  var em = c.get('POFF_' + ver); if (em) return JSON.parse(em);
  var sheetXml = '', sstXml = '';
  Utilities.unzip(f.getBlob().setContentType('application/zip')).forEach(function (p) {
    var n = p.getName();
    if (n === 'xl/worksheets/sheet1.xml') sheetXml = p.getDataAsString('UTF-8');
    else if (n === 'xl/sharedStrings.xml') sstXml = p.getDataAsString('UTF-8');
  });
  var sst = [], m, reSi = /<si>([\s\S]*?)<\/si>/g;
  while ((m = reSi.exec(sstXml)) !== null) sst.push(desXml_(m[1].replace(/<[^>]+>/g, '')));
  var out = [], reRow = /<row [^>]*>([\s\S]*?)<\/row>/g, reCel = /<c r="([A-Z]+)\d+"(?:[^>]*?t="(\w+)")?[^>]*>(?:<v>([^<]*)<\/v>)?/g, prim = true;
  while ((m = reRow.exec(sheetXml)) !== null) {
    if (prim) { prim = false; continue; }
    var cel = {}, k; reCel.lastIndex = 0;
    while ((k = reCel.exec(m[1])) !== null) { var v = k[3] == null ? '' : k[3]; cel[k[1]] = k[2] === 's' ? (sst[Number(v)] || '') : v; }
    var sku = String(cel.B || '').trim(); if (!/^\d+$/.test(sku)) continue;
    out.push({ sku: sku, desc: String(cel.C || '').trim(), alt: String(cel.D || ''), novo: Number(cel.F) || 0, atual: Number(cel.G) || 0, secao: String(cel.J || ''), cat: String(cel.K || '') });
  }
  var r = { itens: out, arquivoEm: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM HH:mm') };
  try { c.put('POFF_' + ver, JSON.stringify(r), 21600); } catch (e) {}
  return r;
}
function statusPoff_() {
  var sh = criarAbaPoff_(SpreadsheetApp.getActive()), n = sh.getLastRow(), st = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, CAB_POFF.length).getValues().forEach(function (l, i) {
    var s = String(l[5] || '').toUpperCase().trim(); if (!l[0] || !s || s === 'DESFEITO') return;
    st[chavePoff_(l[0], l[2])] = { status: s, quem: String(l[6] || ''), em: l[7] instanceof Date ? fmt_(l[7], 'dd/MM HH:mm') : '', ativ: l[8] instanceof Date ? fmt_(l[8], 'dd/MM HH:mm') : '', lin: i + 2 };
  });
  return st;
}
// v77: fardo/caixa (FD/CX ... C/10UN) fica ligado à unidade. FD de arroz e feijão aparece logo abaixo da unidade (mesma repositora);
// os outros fardos/caixas não têm etiqueta: ficam ocultos e são ativados junto quando a unidade for confirmada.
function baseFardo_(d) {
  var s = normDesc_(d), re1 = /^(FD|FARDO|CX|CAIXA)\s+/, re2 = /\s+C\/\s*\d+\s*(UN|UND|UNID|UNIDADES)?\.?$/;
  if (!re1.test(s) && !re2.test(s)) return '';
  return s.replace(re1, '').replace(re2, '').trim();
}
function itensPoff_() {
  var rel = lerPrecosOff_(), st = statusPoff_(), mp = mapaReposicao_(), porDesc = {};
  rel.itens.forEach(function (it) { if (!baseFardo_(it.desc)) porDesc[normDesc_(it.desc)] = it; });
  var itens = rel.itens.map(function (it) {
    var s = st[chavePoff_(it.sku, it.novo)] || {}, base = baseFardo_(it.desc), pai = base ? porDesc[base] : null;
    var r = (pai && mp[normDesc_(pai.desc)]) || mp[normDesc_(it.desc)] || {};
    return { sku: it.sku, desc: it.desc, novo: it.novo, atual: it.atual, alt: it.alt, cat: it.cat, rep: r.rep || 'SEM RESPONSÁVEL', local: r.local || '', status: s.status === 'EM ATIVACAO' ? 'CONFIRMADO' : (s.status || ''), quem: s.quem || '', em: s.em || '', ativ: s.ativ || '',
      fardo: !!base, pai: pai ? pai.sku : '', ord: (base || normDesc_(it.desc)) + (base ? ' 2' : ' 1'), oculto: !!base && !/ARROZ|FEIJAO/.test(normDesc_(it.desc + ' ' + it.cat)) };
  });
  return { itens: itens, arquivoEm: rel.arquivoEm };
}
function precosOff(quem) {
  if (apps(quem).indexOf('PRECOS_OFF') < 0) return { ok: false, erro: 'sem acesso' };
  var r = itensPoff_();
  return { ok: true, itens: r.itens, arquivoEm: r.arquivoEm, verTudo: verTudo_(quem) };
}
function confirmarPrecosOff(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('PRECOS_OFF') < 0) return { ok: false, erro: 'sem acesso' };
  var pedidos = [].concat(p.skus || []).map(String); if (!pedidos.length) return { ok: false, erro: 'escolha os produtos' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var novos = [];
  try {
    var sh = criarAbaPoff_(SpreadsheetApp.getActive()), agora = new Date();
    itensPoff_().itens.forEach(function (it) { // v77: fardos/caixas ocultos vão junto com a unidade
      if (it.status || (pedidos.indexOf(it.sku) < 0 && !(it.oculto && it.pai && pedidos.indexOf(it.pai) >= 0))) return;
      novos.push([it.sku, it.desc, it.novo, it.atual, it.alt, 'CONFIRMADO', dono, agora, '', '']);
    });
    if (novos.length) {
      var n = sh.getLastRow();
      sh.getRange(n + 1, 1, novos.length, 1).setNumberFormat('@'); sh.getRange(n + 1, 5, novos.length, 1).setNumberFormat('@');
      sh.getRange(n + 1, 1, novos.length, CAB_POFF.length).setValues(novos);
      sh.getRange(n + 1, 8, novos.length, 2).setNumberFormat('dd/MM/yyyy HH:mm');
      marcarOrigem_(sh, n + 1, novos.length, CAB_POFF.length);
    }
  } finally { lock.releaseLock(); }
  if (novos.length) {
    try {
      var tx = novos.slice(0, 3).map(function (l) { return l[1] + ' R$ ' + Number(l[2]).toFixed(2).replace('.', ','); }).join(' · ') + (novos.length > 3 ? ' · +' + (novos.length - 3) : '');
      if (avLig_('POFF_PRONTO')) enviarPush_(avNomes_('POFF_PRONTO', AVISO_POFF_PARA), '🏷️ ' + dono + ': ' + novos.length + (novos.length === 1 ? ' preço pronto' : ' preços prontos') + ' para ativar', tx, { abrir: 'precosoff', tag: 'poff' });
      if (apps('Nath').indexOf('POFF_FILA') >= 0 && dono !== 'Nath') v9Push_('Nath', '💲 ' + novos.length + (novos.length === 1 ? ' preço para ativar' : ' preços para ativar'), dono + ': ' + tx, 'pffila'); // v110.3
    } catch (e) {}
  }
  return { ok: true, n: novos.length };
}
function desfazerPrecoOff(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('PRECOS_OFF') < 0) return { ok: false, erro: 'sem acesso' };
  var s = statusPoff_()[chavePoff_(p.sku, p.novo)];
  if (!s) return { ok: true };
  if (s.status !== 'CONFIRMADO') return { ok: false, erro: 'o preço já foi ativado' };
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_POFF), stt = statusPoff_(); sh.getRange(s.lin, 6).setValue('DESFEITO'); sh.getRange(s.lin, 10).setValue('desfeito por ' + dono);
  itensPoff_().itens.forEach(function (it) { if (it.oculto && it.pai === String(p.sku) && it.status === 'CONFIRMADO') { var x = stt[chavePoff_(it.sku, it.novo)]; if (x) { sh.getRange(x.lin, 6).setValue('DESFEITO'); sh.getRange(x.lin, 10).setValue('desfeito por ' + dono + ' (junto da unidade)'); } } }); // v77
  return { ok: true };
}

// ---------- Web app ----------
function doGet(e) {
  ORIGEM_ATUAL = 'NOVO'; // v78: chamada por link (app do GitHub)
  var p = (e && e.parameter) || {};
  if (p.acao) { var rg = tratar_(p); cache131Limpa_(p); return responder_(rg); } // v131
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Apps Pauvan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  ORIGEM_ATUAL = 'NOVO'; // v78: app hospedado no GitHub
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = (e && e.parameter) || {}; }
  return responder_(tratarMedido_(p));
}
// v118.2: registra a ação de cada chamada (aparece em Execuções) e grava na aba LOG LENTO as que passam de 20 s — para achar o que estoura os 6 min
// v131: depois de uma ação que grava, a tela inicial de quem gravou e a lista de Decisões são recalculadas
var LEITURA131 = /^(html|pessoas|buscar|ultimos|apps|rebaixas|atalhos|buscarNome|consultaPreco|resumoHome|fotosSite|semVenda|decLista|decResumo|ctaGestao|ctaMinhas|ctaInfo|pinStatus|conectaPin|etqClaude|ctaLancar|decExecutar|psvGerar)$/;
function cache131Limpa_(p) {
  try { var a = String(p.acao || ''); if (LEITURA131.test(a) && !p.op) return; var c = CacheService.getScriptCache(); c.remove('DECL131_n'); if (p.quem) c.remove(rhChave_(p.quem)); } catch (e) {}
}
function cacheGrande_(k, obj, ttl) { // valor grande no CacheService em pedaços (limite de 100 KB por chave)
  try { var s = JSON.stringify(obj), n = Math.ceil(s.length / 40000), m = {}; if (!n || n > 30) return; for (var i = 0; i < n; i++) m[k + '_' + i] = s.substr(i * 40000, 40000); m[k + '_n'] = String(n); CacheService.getScriptCache().putAll(m, ttl); } catch (e) {}
}
function cacheGrandeLer_(k) {
  try { var c = CacheService.getScriptCache(), n = Number(c.get(k + '_n')); if (!n) return null; var ks = []; for (var i = 0; i < n; i++) ks.push(k + '_' + i); var g = c.getAll(ks), s = ''; for (i = 0; i < n; i++) { if (g[ks[i]] == null) return null; s += g[ks[i]]; } return JSON.parse(s); } catch (e) { return null; }
}
function tratarMedido_(p) {
  var t0 = Date.now(), tag = 'acao=' + String(p.acao || '') + ' quem=' + String(p.quem || '') + (p.idNf ? ' nf=' + String(p.idNf).slice(-14) : '') + (p.idItem ? ' item=' + p.idItem : '');
  try { console.log('INICIO ' + tag); } catch (e) {}
  var r = tratar_(p), s = Math.round((Date.now() - t0) / 100) / 10;
  cache131Limpa_(p); // v131
  try { console.log('FIM ' + tag + ' ' + s + 's'); } catch (e) {}
  if (s > 20) { try { var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('LOG LENTO') || ss.insertSheet('LOG LENTO'); if (sh.getLastRow() < 1) sh.appendRow(['QUANDO', 'ACAO', 'QUEM', 'NF', 'ITEM', 'SEGUNDOS', 'ORIGEM']); sh.appendRow([agoraBr_(), String(p.acao || ''), String(p.quem || ''), String(p.idNf || ''), String(p.idItem || ''), s, ORIGEM_ATUAL]); } catch (e) {} }
  return r;
}

// chamada pela página servida pelo próprio Apps Script (google.script.run)
function api(p) { ORIGEM_ATUAL = 'ANTIGO'; var r = tratarMedido_(p || {}); try { return JSON.parse(JSON.stringify(r)); } catch (e) { return r; } } // v50: resposta sempre em JSON puro (datas etc. não derrubam o retorno)

// v78: ORIGEM APP — de qual app veio o registro (NOVO = GitHub, ANTIGO = link do Apps Script)
var ORIGEM_ATUAL = '';
function marcarOrigem_(sh, linha, qtd, piso) {
  try {
    if (!ORIGEM_ATUAL || !sh || !(linha > 1)) return;
    var lc = Math.max(sh.getLastColumn(), 1);
    var cab = sh.getRange(1, 1, 1, lc).getValues()[0].map(String);
    var c = cab.indexOf('ORIGEM APP') + 1;
    if (!c) { c = Math.max(lc, piso || 0) + 1; sh.getRange(1, c).setValue('ORIGEM APP').setFontWeight('bold'); }
    var v = []; for (var i = 0; i < (qtd || 1); i++) v.push([ORIGEM_ATUAL]);
    sh.getRange(linha, c, v.length, 1).setValues(v);
  } catch (e) { }
}

function tratar_(p) {
  try {
    switch (p.acao) {
      case 'html': return { ok: true, html: HtmlService.createHtmlOutputFromFile('index').getContent() }; // v60: tela do app para a página do GitHub (câmera ao vivo + ícone)
      case 'pessoas': return { ok: true, pessoas: pessoas_() };
      case 'buscar': return buscarProduto(p.codigo);
      case 'salvar': return salvarLancamento(p);
      case 'ultimos': return { ok: true, itens: (p.tipo === 'AVARIA' ? ultimosAvarias_(p.quem) : ultimos(p.quem, p.tipo)) };
      case 'excluir': return excluirLancamento(p.id, p.token, p.tipo);
      case 'salvarAvaria': return salvarAvaria(p);
      case 'apps': return { ok: true, apps: apps(p.quem) };
      case 'rebaixas': var rb = rebaixas(); return { ok: true, vigentes: rb.vigentes, encerradas: rb.encerradas, meus: setores(p.quem), verTudo: verTudo_(p.quem) };
      case 'encerrarRebaixa': return encerrarRebaixa(p);
      case 'reabrirRebaixa': return reabrirRebaixa(p.id, p.token);
      case 'atalhos': return { ok: true, grupos: atalhos(p.quem, p.tipo) };
      case 'buscarNome': return { ok: true, itens: buscarNome(p.texto, p.kg) };
      case 'consultaPreco': return consultaPreco(p.codigo); // v62
      case 'resumoHome': return resumoHome_(p.quem); // v62
      case 'fotosSite': return fotosSite(p.quem); // v66
      case 'precosOff': return precosOff(p.quem); // v76
      case 'confirmarPrecosOff': return confirmarPrecosOff(p); // v76
      case 'desfazerPrecoOff': return desfazerPrecoOff(p); // v76
      case 'vencidos': return vencidos(p.quem); // v79
      case 'vencidoNaoEncontrado': return vencidoNaoEncontrado(p); // v79
      case 'vencidoDesfazer': return vencidoDesfazer(p); // v79
      case 'validades': return validades(p.quem); // v81
      case 'semVenda': return produtoSemVenda(p.quem); // v87
      case 'psvResponder': return psvResponder(p); // v87
      case 'psvGerar': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; instalarGatilhoVencidos_(); return psvGerar(); // v87
      case 'psvAviso': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return psvAvisoDia(); // v87
      case 'trocaFeita': return trocaFeita(p); // v81
      case 'cvSalvar': return cvSalvar(p); // v82
      case 'cvRebSit': return cvRebSit(p); // v121
      case 'locaisOpcoes': return locaisOpcoes(p); // v99
      case 'localSugerir': return localSugerir(p); // v99
      case 'localDecidir': return localDecidir(p); // v99
      case 'cvValCheck': return cvValCheck(p); // v100
      case 'cvAvInfo': return cvAvInfo(p); // v104
      case 'cvAvulso': return cvAvulso(p); // v104
      case 'etqInfo': return etqInfo(p); // v109
      case 'etqPedir': return etqPedir(p);
      case 'etqMeus': return etqMeus(p);
      case 'etqCancelar': return etqCancelar(p);
      case 'etqFila': return etqFila(p);
      case 'etqAcao': return etqAcao(p);
      case 'pfFila': return pfFila(p); // v110.3
      case 'pfAcao': return pfAcao(p);
      case 'agcLista': return agcLista(p); // v123: Agenda Celio
      case 'agcAcao': return agcAcao(p);
      case 'decLista': return decLista(p); // v113 Decisões
      case 'decDecidir': return decDecidir(p);
      case 'decExecutar': return decExecutar(p);
      case 'agdLista': return agdLista(p); // v114 Tarefas agendadas
      case 'decDesfazer': return decDesfazer(p); // v115 mudar decisão
      case 'grmMinhas': return grmMinhas(p); // v118 conferir gramatura na loja
      case 'grmResponder': return grmResponder(p);
      case 'etqClaude': return etqClaude(p);
      case 'ctaInfo': return ctaInfo(p);
      case 'ctaPedir': return ctaPedir(p);
      case 'ctaMinhas': return ctaMinhas(p);
      case 'ctaContar': return ctaContar(p);
      case 'ctaGestao': return ctaGestao(p);
      case 'ctaDecidir': return ctaDecidir(p);
      case 'ctaCorrigir': return ctaCorrigir(p); // v127
      case 'ctaLancar': return ctaLancar(p);
      case 'vfLista': return vfLista(p);
      case 'vfResponder': return vfResponder(p);
      case 'vfAchei': return vfAchei(p);
      case 'vfResumo': return vfResumo(p);
      case 'vfSumido': return vfSumido(p);
      case 'acharLista': return acharLista(p); // v119.7
      case 'acharResp': return acharResp(p);
      case 'vfConfrontar': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return vfConfrontar();
      case 'tfAtraso': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return tfAtrasoRodar_(String(p.teste || '') === '1' ? 'Celio' : '');
      case 'tfAtrasoLigar': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; PropertiesService.getScriptProperties().setProperty('TF_ATRASO_EQUIPE', String(p.liga || '') === '1' ? 'SIM' : 'NAO'); return { ok: true, equipe: String(p.liga || '') === '1' };
      case 'tfAtrasoInstalar': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return tfAtrasoInstalar();
      case 'v129Queima': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return v129CorrigirQueima_(String(p.grava || '') === '1');
      case 'v124Legado': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return v124Legado_(String(p.grava || '') === '1');
      case 'hfMaisUsados': return hfMaisUsados(p);
      case 'hfAprender': return hfAprender(p);
      case 'v9Prod': return v9ProdInfo(p);
      case 'confItemRefazer': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return confItemRefazer_(p); // v101
      case 'kgPecas': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return { ok: true, skus: Object.keys(kgPecas_()) }; // v101
      case 'avisosInstalar': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return avisosInstalar(); // v103
      case 'avisosAgora': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return avisosHora(); // v103
      case 'atBlocos': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return atBlocosPreencher_(); // v107
      case 'cadSku': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return cadSkuPreencher(); // v99
      case 'mediasCV': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return cvAtualizarMedias();
      case 'montarCV': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; instalarGatilhoVencidos_(); return atualizarCV();
      case 'montarVencidos': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; instalarGatilhoVencidos_(); return atualizarVencidos();
      case 'salvarFotosSite': return salvarFotosSite(p); // v66
      case 'registrarPush': return registrarPush_(p); // v63
      case 'enviarAviso': return enviarAviso_(p); // v63
      case 'avisoPessoas': return avisoPessoas_(p); // v91
      case 'salvarChaveFcm': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; var sa = JSON.parse(p.conteudo); if (!sa.private_key || !sa.client_email) return { ok: false, erro: 'arquivo inválido' }; PropertiesService.getScriptProperties().setProperty('FCM_SA', JSON.stringify({ client_email: sa.client_email, private_key: sa.private_key })); CacheService.getScriptCache().remove('FCM_AT'); return { ok: true, email: sa.client_email };
      case 'testePush': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return enviarPush_(p.para || 'Celio', p.titulo || 'Teste do Pauvan', p.texto || 'Se chegou, as notificações estão funcionando ✅', { abrir: 'aviso' });
      case 'pinStatus': return pinStatus(p.quem);
      case 'pinDefinir': return pinDefinir(p.quem, p.pin);
      case 'pinValidar': return pinValidar(p.quem, p.pin);
      case 'conectaPin': return conectaPin_(p); // v126: Pauvan Conecta — PIN -> nome (nunca devolve PIN nem token)
      case 'contagem': var ct = contagem(p.quem); return { ok: true, dia: ct.dia, itens: ct.itens, verTudo: verTudo_(p.quem) };
      case 'salvarContagem': return salvarContagem(p);
      case 'carregarNF': return carregarNF(p);
      case 'conferencias': var cf = conferencias(p.quem); return { ok: true, pendentes: cf.pendentes, concluidas: cf.concluidas };
      case 'confItens': return confItens(p.idNf);
      case 'salvarConfItem': return salvarConfItem(p);
      case 'concluirConf': return concluirConf(p);
      case 'avisoGravar': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return avisoGravar_(p); // v93
      case 'avisoMarcar': return avisoMarcar_(p); // v93
      case 'recusarNota': return recusarNota_(p); // v93
      case 'salvarTroca': return salvarTroca(p);
      case 'extraAdd': return extraAdd_(p); // v96
      case 'extraDel': return extraDel_(p); // v96
      case 'excluirTroca': return excluirTroca(p);
      case 'varrerXml': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return varrerXml();
      case 'regravarValidades': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return regravarValidades_(p.idNf, p.diag);
      case 'instalarGatilhos': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return instalarGatilhos();
      case 'salvarBackup': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return salvarBackup_(p.nome, p.conteudo);
      case 'aprenderVinculos': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return aprenderVinculos_();
      case 'syncRebaixas': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return sincronizarRebaixas();
      case 'testeRebaixas': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return { ok: true, detectadas: detectarRebaixas_() };
      case 'recarregarCadastro': if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' }; return { ok: true, linhas: atualizarCadastro() };
      default: return { ok: false, erro: 'ação desconhecida' };
    }
  } catch (err) {
    return { ok: false, erro: String(err) };
  }
}

function responder_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ========== CONTAGEM DIA D ==========
// Aba CONTAGEM DIA D: itens que as meninas contam na gôndola antes do Dia D.
// Claude cruza depois com a CURVA A (venda de 15 dias) e o saldo atual para sugerir a separação.
function criarAbaContagem_(ss) {
  var sh = ss.getSheetByName(ABA_CONT);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_CONT);
  var dia = 'DIA D OUT/26';
  var seed = [
    ['1', 'TATI', 'ARROZ MASSON "T1" 5KG', ''],
    ['1', 'TATI', 'FEIJAO CARIOCA MASSON "T1" 1KG', ''],
    ['1', 'TATI', 'CAFE ROCHA SC500G', ''],
    ['1', 'TATI', 'MACARRAO GALO ESPAGUETE 8 500G (AZUL)', ''],
    ['1', 'TATI', 'MOLHO DE TOMATE OLE TRADICIONAL SC300G', ''],
    ['2', 'GLAUCIA', 'LEITE INTEGRAL PIRACANJUBA 1L', ''],
    ['2', 'GLAUCIA', 'LEITE CONDENSADO PIRACANJUBA 395G', ''],
    ['2', 'GLAUCIA', 'CREME DE LEITE PIRACANJUBA CX200G', ''],
    ['2', 'GLAUCIA', 'LEITE EM PO PIRACANJUBA SC400G', 'AG'],
    ['2', 'GLAUCIA', 'SALGADINHO CAMPITOS 40G', 'AG'],
    ['3', 'NEY / REGIANE', 'DETERGENTE YPE 500ML', 'AG'],
    ['3', 'NEY / REGIANE', 'AMACIANTE YPE 2L', 'AG'],
    ['3', 'NEY / REGIANE', 'SABAO EM PO OMO LAVAGEM PERFEITA 700G', ''],
    ['3', 'NEY / REGIANE', 'CX CERVEJA BRAHMA CHOPP LT269ML C/15UN MULTIPACK', ''],
    ['3', 'NEY / REGIANE', 'REFRIGERANTE COCA-COLA 2L ( COCA ZERO TAMBEM)', '']
  ];
  var linhas = [CAB_CONT];
  seed.forEach(function (s, i) {
    linhas.push([dia, s[0], s[1], skuPorDescricao_(s[2]), s[2], s[3], '', '', '', 'C' + (i + 1)]);
  });
  sh.getRange(1, 1, linhas.length, CAB_CONT.length).setValues(linhas);
  sh.getRange(1, 1, 1, CAB_CONT.length).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(5, 380);
  return sh;
}

// procura o SKU pela descrição exata na aba CADASTRO (vazio se não achar)
function skuPorDescricao_(desc) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CAD);
    if (!sh) return '';
    var n = sh.getLastRow(); if (n < 2) return '';
    var v = sh.getRange(2, 2, n - 1, 2).getValues();
    var d = String(desc).trim().toUpperCase();
    for (var i = 0; i < v.length; i++) if (String(v[i][1]).trim().toUpperCase() === d) return String(v[i][0]);
  } catch (e) {}
  return '';
}

// lista do Dia D mais recente (último valor da coluna DIA D), com todos os blocos
function contagem(quem) {
  var sh = criarAbaContagem_(SpreadsheetApp.getActive());
  var n = sh.getLastRow(), itens = [];
  if (n < 2) return { dia: '', itens: [] };
  var v = sh.getRange(2, 1, n - 1, CAB_CONT.length).getValues();
  var dia = String(v[v.length - 1][0]);
  v.forEach(function (l) {
    if (String(l[0]) !== dia || !String(l[4]).trim()) return; // v59: linha sem produto = Dia D arquivado (lista vazia)
    itens.push({ id: String(l[9]), bloco: String(l[1]), resp: String(l[2]), sku: String(l[3]), produto: String(l[4]), ag: String(l[5]) === 'AG',
      qtd: l[6] === '' ? null : Number(l[6]), quem: String(l[7]), hora: l[8] ? Utilities.formatDate(new Date(l[8]), 'America/Cuiaba', 'dd/MM HH:mm') : '' });
  });
  // v57: a divisão sai da repositora do cadastro (VALIDADE PRODUTOS); caixa, açougue, Ney e sem cadastro ficam com NEY / REGIANE
  var mpC = mapaReposicao_();
  itens.forEach(function (it) { var r = mpC[normDesc_(it.produto)], rep = r ? r.rep : ''; it.local = r ? r.local : ''; it.rep = (!rep || /CAIXA|ACOUGUE|^NEY$/.test(semAcento_(rep).toUpperCase().trim())) ? 'NEY / REGIANE' : rep; });
  return { dia: dia, itens: itens };
}

// grava a quantidade contada na gôndola (pode regravar para corrigir)
function salvarContagem(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = criarAbaContagem_(SpreadsheetApp.getActive());
  var n = sh.getLastRow();
  var ids = sh.getRange(2, 10, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(p.id)) {
      var q = p.qtd === '' || p.qtd === null || p.qtd === undefined ? '' : Number(p.qtd);
      sh.getRange(i + 2, 7, 1, 3).setValues([[q, q === '' ? '' : String(p.quem || ''), q === '' ? '' : new Date()]]);
      marcarOrigem_(sh, i + 2, 1, 10);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'item não encontrado' };
}

// ========== CONFERÊNCIA CEGA (recebimento de NF) ==========
// Abas: CONFERENCIA NF (cabeçalho), CONFERENCIA ITENS (itens da nota, com a quantidade que o app NÃO mostra),
// CONFERENCIA LANC (o que a conferente contou, uma linha por validade) e CONFERENCIA EXCECOES (fornecedores que não passam por conferência).
// Claude carrega a nota pelo XML (ação carregarNF) e compara depois (ação concluirConf).
var ABA_CNF = 'CONFERENCIA NF', ABA_CIT = 'CONFERENCIA ITENS', ABA_CLA = 'CONFERENCIA LANC', ABA_CEX = 'CONFERENCIA EXCECOES';
var CAB_CNF = ['ID_NF', 'FORNECEDOR', 'NF', 'EMISSAO', 'ENTRADA_PLENUS', 'STATUS', 'CONFERENTE', 'CONCLUIDA_EM', 'RESULTADO', 'ITENS', 'CONFERIDOS', 'RODADA', 'OBS', 'ARQUIVO'];
var CAB_CIT = ['ID_NF', 'ITEM', 'SKU', 'PRODUTO', 'CODIGO', 'UN', 'QTD_NF', 'QTD_AUX', 'FATOR', 'NOVO', 'STATUS', 'RODADA', 'DIVERG', 'QTD_CONFERIDA', 'ID_ITEM', 'DESC_NF'];
var CAB_CLA = ['ID_NF', 'ID_ITEM', 'SKU', 'PRODUTO', 'FORMA', 'QTD_DIGITADA', 'FATOR', 'QTD', 'VALIDADE', 'QUEM', 'DATA', 'HORA', 'RODADA', 'ID'];
var CHAVE_CARGA = 'pauvan-nf';

function criarAbaConf_(ss) {
  function aba(nome, cab, cor) {
    var sh = ss.getSheetByName(nome);
    if (sh) return sh;
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground(cor).setFontColor('#ffffff');
    sh.setFrozenRows(1);
    return sh;
  }
  var a = aba(ABA_CNF, CAB_CNF, '#0e7490');
  var b = aba(ABA_CIT, CAB_CIT, '#0e7490');
  var c = aba(ABA_CLA, CAB_CLA, '#0e7490');
  var d = ss.getSheetByName(ABA_CEX);
  if (!d) {
    d = ss.insertSheet(ABA_CEX);
    d.getRange(1, 1, 3, 2).setValues([['FORNECEDOR', 'OBS'], ['PADARIA 2000', 'entrega durante a semana e emite a NF depois'], ['', 'verduras da região: Celio vai indicar']]);
    d.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
    d.setFrozenRows(1);
  }
  return { nf: a, it: b, la: c, ex: d };
}

function nfeTexto_(el, nome, ns) { var x = el ? el.getChild(nome, ns) : null; return x ? x.getText() : ''; }

// Carrega uma NF-e a partir do XML salvo no Drive (pasta XML NF). p.arquivo = id ou nome do arquivo; p.entrada = 'dd/MM/aaaa' da entrada no Plenus ('' = ainda sem entrada); p.fornecedor = apelido (opcional).
function carregarNF(p) {
  if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' };
  var arq = null, ref = String(p.arquivo || '').trim();
  if (!ref) return { ok: false, erro: 'informe o arquivo' };
  try { arq = DriveApp.getFileById(ref); } catch (e) {
    var it = DriveApp.getFilesByName(/\.xml$/i.test(ref) ? ref : ref + '.xml');
    if (it.hasNext()) arq = it.next();
  }
  if (!arq) return { ok: false, erro: 'arquivo não encontrado: ' + ref };
  var xml = XmlService.parse(arq.getBlob().getDataAsString('UTF-8'));
  var ns = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');
  var root = xml.getRootElement();
  var nfe = root.getName() === 'NFe' ? root : root.getChild('NFe', ns);
  if (!nfe) return { ok: false, erro: 'XML não é uma NF-e' };
  var inf = nfe.getChild('infNFe', ns);
  var ide = inf.getChild('ide', ns), emit = inf.getChild('emit', ns);
  var chave = String(inf.getAttribute('Id') ? inf.getAttribute('Id').getValue() : '').replace(/^NFe/, '');
  var nNF = nfeTexto_(ide, 'nNF', ns), dh = nfeTexto_(ide, 'dhEmi', ns) || nfeTexto_(ide, 'dEmi', ns);
  var emissao = dh ? dh.slice(8, 10) + '/' + dh.slice(5, 7) + '/' + dh.slice(0, 4) : '';
  var apelido = String(p.fornecedor || '').trim();
  if (!apelido) { var m = arq.getName().match(/^(.*?)\s*\d{3,}/); apelido = m ? m[1].trim() : (nfeTexto_(emit, 'xFant', ns) || nfeTexto_(emit, 'xNome', ns)); }
  if (!apelido) apelido = nfeTexto_(emit, 'xNome', ns);
  var idNf = chave || ('nf_' + nfeTexto_(emit, 'CNPJ', ns) + '_' + nNF);
  var cadM = mapaCodMulti_(), cad = mapaCadastro_(), vinc = vinculosFixos_(), aprend = vinculosAprendidos_(), cnpjEmit = nfeTexto_(emit, 'CNPJ', ns) || nfeTexto_(emit, 'CPF', ns);
  var cob = cobrancaXml_(inf, ns);
  // bonificada: pelo nome do arquivo ("FORMULA 392005 BONIF.xml"), pela natureza da operação ou pelo parâmetro bonif=1
  var bonif = (/BONIF/i.test(arq.getName()) || /BONIF/i.test(nfeTexto_(ide, 'natOp', ns)) || String(p.bonif || '') === '1') ? 'S' : '';
  var itens = [];
  inf.getChildren('det', ns).forEach(function (det) {
    var pr = det.getChild('prod', ns);
    var cEAN = nfeTexto_(pr, 'cEAN', ns).replace(/\D/g, ''), cTrib = nfeTexto_(pr, 'cEANTrib', ns).replace(/\D/g, '');
    var uCom = nfeTexto_(pr, 'uCom', ns).toUpperCase(), qCom = Number(nfeTexto_(pr, 'qCom', ns)) || 0;
    var uTrib = nfeTexto_(pr, 'uTrib', ns).toUpperCase() || uCom, qTrib = Number(nfeTexto_(pr, 'qTrib', ns)) || qCom;
    var desc = nfeTexto_(pr, 'xProd', ns);
    var ad = nfeTexto_(det, 'infAdProd', ns), ma = ad.match(/Qtde_aux=(\d+(?:[.,]\d+)?)/i);
    var nu = normUnidade_(uCom, qCom, uTrib, qTrib, desc, String(ad).split('|')[0]); // v96: volume também do infAdProd ("FDO 1 X 6 900ML")
    var qAux = ma ? Number(ma[1].replace(',', '.')) : '';
    var cProd = nfeTexto_(pr, 'cProd', ns), kv = chaveVinc_(cnpjDoIdNf_(idNf) || cnpjEmit, cProd, desc), sv = aprend[kv];
    // v118.7 (pedido do Celio, 13/11): vínculo aprendido (fornecedor + código do produto) > vínculo fixo do código > código de barras da nota no CADASTRO (1 cadastro só; caixa/fardo -> unidade x qtd)
    var c = (sv && cadPorSku_(cad, sv)) || (cTrib && vinc[cTrib] && cadPorSku_(cad, vinc[cTrib])) || (cEAN && vinc[cEAN] && cadPorSku_(cad, vinc[cEAN])) || null, origem = c ? 'VINCULO' : '', nu0 = nu;
    if (!c) {
      var ea = eanAchar_(cadM, cad, [cTrib, cEAN]);
      if (ea.c) { c = ea.c; origem = 'EAN'; if (ea.caixa > 1) { origem = 'EAN CAIXA x' + ea.caixa; if (nu.un !== 'UN' && nu.un !== 'KG') nu = { un: 'UN', qtd: Math.round(nu.qtd * ea.caixa * 1000) / 1000, fator: '' }; } }
      else if (ea.cands.length > 1) origem = 'EAN AMBIGUO: ' + ea.cands.join(' / ');
    }
    if (/^(L|LT|LTS|LITRO|LITROS)$/.test(nu.un) && c && c.desc) nu = normUnidade_(uCom, qCom, uTrib, qTrib, desc, c.desc); // v96: ainda em litros -> volume pela descrição do cadastro (OLEO ... 900ML)
    if (nu.un === 'KG' && c && c.un && !/QUILO/.test(String(c.un)) && !(KGP_ || (KGP_ = kgPecas_()))[skuN_(c.sku)]) { // v101: lista KG + peças fica em kg // v34: nota em kg de produto vendido por unidade (Masson: pacote 5kg/1kg/500g) -> unidades
      var pk = pesoDesc_(c.desc) || pesoDesc_(desc);
      if (pk > 0) { var uq2 = nu.qtd / pk, ur2 = Math.round(uq2); if (ur2 > 0 && Math.abs(uq2 - ur2) <= Math.max(0.05, ur2 * 0.02)) nu = { un: 'UN', qtd: ur2, fator: '' }; }
    }
    itens.push({ cProd: cProd, item: String(det.getAttribute('nItem') ? det.getAttribute('nItem').getValue() : itens.length + 1), sku: c ? c.sku : '', produto: c ? c.desc : '** ' + desc, codigo: cTrib || cEAN, un: nu.un, qtd: nu.qtd, qAux: qAux, fator: nu.fator || (sv && nu.un !== 'UN' && nu.un !== 'KG' && aprend['F|' + kv]) || '', novo: c ? '' : 'S', descNf: desc, origem: origem, nu0: nu0 });
  });
  if (!itens.length) return { ok: false, erro: 'NF sem itens' };
  var ex = conferenciaExcecao_(apelido, nfeTexto_(emit, 'xNome', ns));
  var dsc = ehDesconhecida_(desconhecidasControle_(), nNF, apelido); // v38: nota marcada como desconhecida na CONTROLE NF não vai para a conferência
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss); criarAbasV24_(ss);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var linha = confLinhaNf_(abas.nf, idNf);
    // entrada no Plenus: informada no parâmetro ou, na varredura automática, lida da CONTROLE NF (coluna DT ENTRADA)
    var entrada = p.entrada != null ? String(p.entrada) : entradaControle_(nNF, apelido);
    if (linha > 0) {
      // já carregada: atualiza entrada e tenta de novo o cruzamento dos itens novos
      if (entrada) abas.nf.getRange(linha, 5).setValue(entrada);
      if (dsc && String(abas.nf.getRange(linha, 6).getValue()) === 'PENDENTE') { abas.nf.getRange(linha, 6).setValue('EXCECAO'); abas.nf.getRange(linha, 13).setValue('nota desconhecida (CONTROLE NF)'); }
      if (bonif) abas.nf.getRange(linha, CAB_CNF.length + 5).setValue(bonif);
      var n = abas.it.getLastRow();
      if (n >= 2) {
        var v = abas.it.getRange(2, 1, n - 1, CAB_CIT.length).getValues();
        v.forEach(function (l, i) {
          if (String(l[0]) !== idNf) return;
          var it = itens.filter(function (x) { return x.item === String(l[1]); })[0]; if (!it) return;
          var r = i + 2, novo = String(l[9]);
          // só as células do item (gravar a tabela inteira transformaria o ID da nota e os códigos em número)
          var skuL = String(l[2]).replace(/\D/g, ''), nuR = it; // v118.7
          if (novo === 'S' && !it.novo) { abas.it.getRange(r, 3, 1, 2).setValues([["'" + it.sku, it.produto]]); novo = ''; skuL = String(it.sku); abas.it.getRange(r, CAB_CIT.length + 8).setValue(it.origem); }
          else if (novo === 'S' && it.origem) abas.it.getRange(r, CAB_CIT.length + 8).setValue(it.origem); // ambíguo: candidatos para a decisão
          if (/^EAN CAIXA/.test(it.origem) && skuL !== String(it.sku)) nuR = it.nu0; // o Plenus vinculou outro SKU (ex.: a caixa): vale o Plenus, sem converter para unidade
          abas.it.getRange(r, 6, 1, 5).setValues([[nuR.un, nuR.qtd, it.qAux, nuR.fator, novo]]); // unidade/fator recalculados
        });
      }
      return { ok: true, idNf: idNf, atualizada: true, itens: itens.length, novos: itens.filter(function (i) { return i.novo; }).length };
    }
    abas.nf.appendRow([idNf, apelido, "'" + nNF, emissao, entrada, (ex || dsc) ? 'EXCECAO' : 'PENDENTE', '', '', '', itens.length, 0, 1, dsc ? 'nota desconhecida (CONTROLE NF)' : ex ? 'fornecedor na lista de exceções' : '', arq.getName(), '', '', cob.dup, cob.venc, bonif, '', arq.getId()]);
    if (!(ex || dsc)) avFilaNf_(apelido + ' ' + nNF); // v103: aviso de NF para conferir
    var rows = itens.map(function (i) { return [idNf, i.item, "'" + i.sku, i.produto, "'" + i.codigo, i.un, i.qtd, i.qAux, i.fator, i.novo, 'PENDENTE', '', '', '', idNf.slice(-8) + '-' + i.item, i.descNf, '', '', "'" + i.cProd, '', '', '', '', i.origem, '']; });
    abas.it.getRange(abas.it.getLastRow() + 1, 1, rows.length, CAB_CIT.length + EXTRA_CIT.length).setValues(rows);
    return { ok: true, idNf: idNf, fornecedor: apelido, nf: nNF, emissao: emissao, itens: itens.length, novos: itens.filter(function (i) { return i.novo; }).length, excecao: !!ex };
  } finally { lock.releaseLock(); }
}
// unidade da nota: UN quando um dos dois lados (comercial/tributável) é unidade; KG; senão a embalagem da nota (DP, CJ, FD, CX, BJ...)
// com o fator lido da descrição ("18UNX18GR", "12X26,6GR" → 18, 12)
// v34: quantidade por caixa/fardo lida da descrição da nota: CX12, PC24, FD20, C/15, SH C15, 20PCT, 500UN, 18X400G
function fatorDesc_(desc) {
  var s = String(desc || '').toUpperCase(), r = /(?:^|[\s\-\/(])(?:CX|PC|PCT|FD|FDO|FR|SH|DP|DISP|BJ|C)\s*[\/.]?\s*0*(\d{1,3})(?:\s*UN(?:ID)?)?(?=[\s\)]|$)/g, m, f = 0;
  while ((m = r.exec(s))) { if (/^\s*\d+(?:[.,]\d+)?\s*KG\b/.test(s.slice(r.lastIndex))) continue; if (Number(m[1]) > 1) f = Number(m[1]); } // v118.1: "CX 3 19KG" = caixa tipo 3 de 19 kg (não 3 por caixa)
  if (f) return f;
  m = s.match(/(?:^|\s)(\d{1,4})\s*(?:PCT|UNID|UN)(?=\s|$)/); if (m && Number(m[1]) > 1) return Number(m[1]);
  m = s.match(/(\d+)\s*(?:UN)?\s*X\s*[\d,.]+\s*(?:G|GR|KG|ML|L|UN)\b/); if (m && Number(m[1]) > 1) return Number(m[1]);
  return '';
}
// v34: peso do pacote na descrição do cadastro (5KG, 1KG, 500G) -> kg
function pesoDesc_(desc) {
  var m = String(desc || '').toUpperCase().match(/(\d+(?:[.,]\d+)?)\s*(KG|G|GR)\b/);
  if (!m) return 0;
  var v = Number(m[1].replace(',', '.')); return m[2] === 'KG' ? v : v / 1000;
}
function normUnidade_(uCom, qCom, uTrib, qTrib, desc, extra) { // v96: extra = infAdProd / descrição do cadastro (volume quando a nota vem em litros)
  var U = /^(UN|UNID|UND|UN1|PC|PCS|PECA|PÇ)$/;
  var r3 = function (x) { return Math.round(x * 1000) / 1000; };
  if (U.test(uTrib)) return { un: 'UN', qtd: qTrib, fator: (!U.test(uCom) && qCom > 0 && qTrib > qCom) ? r3(qTrib / qCom) : '' };
  if (U.test(uCom)) return { un: 'UN', qtd: qCom, fator: '' };
  if (/^(L|LT|LTS|LITRO|LITROS)$/.test(uTrib)) { // v33: nota em litros (Fera) -> unidades pelo volume da descrição (269ML, 2L...)
    var mv = String(desc || '').match(/(\d+(?:[.,]\d+)?)\s*(ML|L)\b/i) || String(extra || '').match(/(\d+(?:[.,]\d+)?)\s*(ML|L)\b/i);
    if (mv) {
      var vol = Number(mv[1].replace(',', '.')) / (/^ML$/i.test(mv[2]) ? 1000 : 1), uq = vol > 0 ? qTrib / vol : 0, ur = Math.round(uq);
      if (ur > 0 && Math.abs(uq - ur) <= Math.max(0.05, ur * 0.02)) return { un: 'UN', qtd: ur, fator: (qCom > 0 && ur > qCom) ? r3(ur / qCom) : '' };
    }
  }
  // v58: CX = KG com a mesma quantidade, mas a descrição traz o peso da caixa (ex.: PE DE FRANGO CX15 KG) -> caixas com fator em kg
  var mcx = String(desc || '').toUpperCase().match(/\bCX\s*(\d+(?:[.,]\d+)?)\s*KG\b/);
  if (uTrib === 'KG' && uCom && uCom !== 'KG' && Math.abs(qCom - qTrib) < 0.0001 && mcx && Number(mcx[1].replace(',', '.')) > 1) return { un: uCom, qtd: qCom, fator: Number(mcx[1].replace(',', '.')) };
  // v49: tributável em KG com a mesma quantidade da comercial (CX 1 = KG 1) e sem peso na descrição = erro do fornecedor -> vale a embalagem comercial
  if (uTrib === 'KG' && uCom && uCom !== 'KG' && Math.abs(qCom - qTrib) < 0.0001 && !/\d\s*(KG|G|GR)\b/i.test(String(desc || ''))) return { un: uCom, qtd: qCom, fator: fatorDesc_(desc) };
  if (uTrib === 'KG' || uCom === 'KG') return { un: 'KG', qtd: uTrib === 'KG' ? qTrib : qCom, fator: '' };
  return { un: uTrib || uCom, qtd: qTrib, fator: fatorDesc_(desc) };
}
function conferenciaExcecao_(apelido, razao) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CEX);
  if (!sh || sh.getLastRow() < 2) return false;
  var a = semAcento_(String(apelido || '')).toUpperCase(), r = semAcento_(String(razao || '')).toUpperCase();
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().some(function (l) {
    var f = semAcento_(String(l[0] || '')).trim().toUpperCase();
    return f && (a.indexOf(f) >= 0 || r.indexOf(f) >= 0);
  });
}
function confLinhaNf_(sh, idNf) {
  var n = sh.getLastRow(); if (n < 2) return -1;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues().map(String);
  var i = ids.indexOf(String(idNf)); return i < 0 ? -1 : i + 2;
}
// lista de notas para o app: pendentes/em conferência/divergentes primeiro, depois as últimas concluídas
function conferencias(quem) {
  var abas = criarAbaConf_(SpreadsheetApp.getActive());
  var n = abas.nf.getLastRow(); if (n < 2) return { pendentes: [], concluidas: [] };
  var v = abas.nf.getRange(2, 1, n - 1, CAB_CNF.length + EXTRA_CNF.length).getValues();
  var cA = avisoCol_(abas.nf); avisoMigrar_(abas.nf, cA); var va = abas.nf.getRange(2, cA, n - 1, AVISO_CNF.length).getValues(); // v93: avisos da nota
  var pend = [], conc = [], iB = CAB_CNF.length + 4, dv = {}, ni = abas.it.getLastRow(), grp = gruposForn_(); // v47: grupos de empresas
  if (ni >= 2) abas.it.getRange(2, 1, ni - 1, CAB_CIT.length + EXTRA_CIT.length).getValues().forEach(function (l) { if (String(l[19] || '')) dv[String(l[0])] = (dv[String(l[0])] || 0) + 1; }); // v36: itens a devolver por nota
  v.forEach(function (l, ix) {
    var st = String(l[5]); if (st === 'EXCECAO') return;
    var o = { idNf: String(l[0]), forn: String(l[1]), nf: String(l[2]), emissao: fmt_(l[3], 'dd/MM/yyyy'), entrada: fmt_(l[4], 'dd/MM/yyyy'), status: st, conferente: String(l[6]), em: fmt_(l[7], 'dd/MM/yyyy HH:mm'), resultado: String(l[8]), itens: Number(l[9]) || 0, feitos: Number(l[10]) || 0, bonif: String(l[iB] || '') === 'S', devolver: dv[String(l[0])] || 0, corrigida: /CORRIGIDA APÓS CONCLUSÃO/.test(String(l[12] || '')) };
    o.grupo = grp[semAcento_(o.forn).toUpperCase().trim()] || o.forn;
    o.aviso = avisoObj_(va[ix]); // v93
    if (st === 'CONCLUIDA' && o.resultado === 'OK' && !(o.aviso && o.aviso.tipo === 'DEVOLUCAO_EMITIDA')) { o.ts = tsConc_(l[7]); conc.push(o); } else pend.push(o);
  });
  pend.reverse();
  // v43: conferidas recentemente = fechadas hoje, ontem e anteontem, a última fechada em cima (mínimo 3 na lista)
  conc.sort(function (a, b) { return b.ts - a.ts; });
  var corte = new Date(); corte.setHours(0, 0, 0, 0); corte.setDate(corte.getDate() - 2);
  var rec = conc.filter(function (o, i) { return o.ts >= corte.getTime() || i < 3; });
  return { pendentes: pend, concluidas: rec.slice(0, 40) };
}
// v43: data/hora de conclusão (Date ou texto dd/MM/yyyy HH:mm) em milissegundos, para ordenar
function tsConc_(x) {
  if (x instanceof Date) return x.getTime();
  var m = String(x || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\D+(\d{1,2}):(\d{2}))?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime() : 0;
}
// itens de uma nota para o app — NUNCA manda QTD_NF nem QTD_AUX
function confItens(idNf) {
  var abas = criarAbaConf_(SpreadsheetApp.getActive()); criarAbasV24_(SpreadsheetApp.getActive());
  var cad = mapaCadastro_(), kgp = kgPecas_(); // v118.8
  var ln = confLinhaNf_(abas.nf, idNf); if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
  var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length + EXTRA_CNF.length).getValues()[0];
  var nf = { idNf: String(h[0]), forn: String(h[1]), nf: String(h[2]), emissao: fmt_(h[3], 'dd/MM/yyyy'), entrada: fmt_(h[4], 'dd/MM/yyyy'), status: String(h[5]), resultado: String(h[8]), rodada: Number(h[11]) || 1, bonif: String(h[CAB_CNF.length + 4] || '') === 'S' };
  nf.aviso = avisoObj_(abas.nf.getRange(ln, avisoCol_(abas.nf), 1, AVISO_CNF.length).getValues()[0]); // v93
  try { nf.extras = extrasDaNf_(idNf); } catch (e) { nf.extras = []; } // v96
  var itens = [], n = abas.it.getLastRow(), rowsNf = [];
  if (n >= 2) abas.it.getRange(2, 1, n - 1, CAB_CIT.length + EXTRA_CIT.length).getValues().forEach(function (l) {
    if (String(l[0]) !== idNf) return;
    rowsNf.push(l);
    itens.push({ kgUn: String(l[5]).toUpperCase() === 'KG' && !!kgp[skuN_(l[2])], avDev: avisoDoItem_(nf.aviso, l[2], l[3], l[15], l[6]), faltaQtd: Number(l[21]) || 0, faltaTipo: String(l[22] || ''), devolver: String(l[19] || ''), devMotivo: String(l[20] || ''), motivo: String(l[16] || ''), devolucao: String(l[17] || ''), idItem: String(l[14]), item: String(l[1]), sku: String(l[2]), produto: String(l[3]), codigo: String(l[4]), un: String(l[5]), fator: l[8] === '' ? null : Number(l[8]), novo: String(l[9]) === 'S', novoCad: /\*\*\s*$/.test(String(l[3])), status: String(l[10]), rodada: Number(l[11]) || 0, diverg: String(l[12]) === 'S', lotes: [], trocas: [], unPlenus: unPl_(cad, String(l[2])), horti: /HORTI|FRUT|VERD|LEGUM/.test(String((cad[String(l[2])] || {}).dem || '')) || (String(l[9]) === 'S' && confNovoHorti_(l[3], nf.forn)) }); // v111: produto novo só é hortifruti se o nome ou o fornecedor for de hortifruti
  });
  var byId = {}; itens.forEach(function (i) { byId[i.idItem] = i; });
  var m = abas.la.getLastRow(), nc = CAB_CLA.length + EXTRA_CLA.length;
  if (m >= 2) abas.la.getRange(2, 1, m - 1, nc).getValues().forEach(function (l) {
    if (String(l[0]) !== idNf) return;
    var it = byId[String(l[1])]; if (!it || Number(l[12]) !== it.rodada) return;
    it.lotes.push({ forma: String(l[4]), qtdDig: Number(l[5]), fator: l[6] === '' ? null : Number(l[6]), qtd: Number(l[7]), validade: fmt_(l[8], 'dd/MM/yyyy'), quem: String(l[9]), pecas: l[14] === '' ? null : Number(l[14]), emb: String(l[15] || ''), qtdEmb: l[16] === '' ? null : Number(l[16]), codLido: String(l[17] || ''), codStatus: String(l[18] || ''), autorizou: String(l[19] || '') });
  });
  trocasDaNf_(idNf).forEach(function (t) { var it = byId[t.idItem]; if (it) it.trocas.push(t); });
  try { var aprC = vinculosAprendidos_(), cnpjC = cnpjDoIdNf_(idNf); rowsNf.forEach(function (l) { var it = byId[String(l[14])]; if (it) it.vConv = valConvTxt_(l, cad, aprC, cnpjC); }); } catch (e) {} // v112.1: conversão para a validade (1 CX = 10 UN)
  try { var hist = confHistFator_(abas); itens.forEach(function (i) { var fh = hist.fat[i.sku + '|' + confFornK_(nf.forn)]; if (fh && !(i.fator > 0)) i.fatorHist = fh; if (hist.emb[i.sku]) i.embHist = true; }); } catch (e) {} // v111: fator pela última entrada do mesmo produto + fornecedor
  try { var cf = cfopsDaNf_(abas, idNf); itens.forEach(function (i) { i.cfop = cf[i.item] || ''; i.bonifItem = /^[56]910$/.test(i.cfop); }); } catch (e) {} // v94: CFOP 5910 = item bonificado
  try { // v95: compra + bonificado do mesmo produto viram uma linha só (contagem única)
    var grs = gruposBonif_(rowsNf, cf || {});
    Object.keys(grs).forEach(function (k) {
      var g = grs[k], pri = byId[k], mem = g.itens.map(function (l) { return byId[String(l[14])]; }).filter(Boolean); if (!pri) return;
      pri.grupo = mem.map(function (i) { return i.idItem; });
      mem.forEach(function (i) { if (i !== pri) i.grupoDe = pri.idItem; });
      pri.status = mem.some(function (i) { return i.status === 'PENDENTE'; }) ? 'PENDENTE' : (mem.every(function (i) { return i.status === pri.status; }) ? pri.status : 'CONFERIDO');
      pri.diverg = mem.some(function (i) { return i.diverg; });
      var mv = {}, ordV = [];
      mem.forEach(function (i) { i.lotes.forEach(function (l) { var key = l.validade + '|' + l.forma; if (!mv[key]) { mv[key] = { forma: l.forma === 'KG' ? 'KG' : 'UNIDADE', qtdDig: 0, fator: null, qtd: 0, validade: l.validade, quem: l.quem, pecas: null, emb: l.emb, qtdEmb: l.qtdEmb, codLido: l.codLido, codStatus: l.codStatus, autorizou: l.autorizou }; ordV.push(key); } mv[key].qtdDig += l.qtd; mv[key].qtd += l.qtd; }); });
      pri.lotes = ordV.map(function (key) { var x = mv[key]; x.qtdDig = Math.round(x.qtdDig * 1000) / 1000; x.qtd = x.qtdDig; return x; });
    });
  } catch (e) {}
  itens.forEach(function (i) { i.pares = itens.filter(function (o) { return o !== i && (i.sku ? o.sku === i.sku : o.produto === i.produto); }).map(function (o) { return { idItem: o.idItem, bonif: !!o.bonifItem }; }); }); // v94: mesmo produto em mais de um item
  return { ok: true, nf: nf, itens: itens };
}
// ---------- v111: conferência — produto novo hortifruti, fator e "Como veio?" aprendidos ----------
var HF_NOVO_RE = /\b(TOMATE|ALFACE|BANANA|MACA|MAMAO|MANGA|UVA|MORANGO|LARANJA|LIMAO|TANGERINA|MEXERICA|BATATA|CEBOLA|ALHO|CENOURA|BETERRABA|ABOBORA|ABOBRINHA|ABACAXI|ABACATE|MELANCIA|MELAO|PEPINO|PIMENTAO|REPOLHO|COUVE|BROCOLIS|CHUCHU|INHAME|MANDIOCA|COGUMELO|CHAMPIGNON|SHIMEJI|KIWI|PERA|AMEIXA|GOIABA|MARACUJA|COENTRO|CEBOLINHA|SALSA|RUCULA|AGRIAO|ESPINAFRE|GENGIBRE|QUIABO|JILO|BERINJELA|VAGEM|FRUTA|VERDURA|LEGUME|HORTALICA)\b/;
function confNovoHorti_(desc, forn) { return HF_NOVO_RE.test(normDesc_(desc)) || /SEMPRE BOM|VERDURAO|H\.?\s?B\.? HORTI|ADAO/.test(normDesc_(forn)); }
function confFornK_(f) { return normDesc_(f).replace(/[^A-Z0-9]/g, ''); }
function confHistFator_(abas) { // {fat: {sku|fornecedor: fator da última entrada em CAIXA}, emb: {sku: 1 se já veio em bandeja}} — cache 10 min
  var c = CacheService.getScriptCache(), k = 'CONF_HIST_V1111', em = c.get(k); if (em) return JSON.parse(em);
  var forn = {}, fat = {}, emb = {}, cnt = {}, n = abas.nf.getLastRow(), m = abas.la.getLastRow();
  if (n >= 2) abas.nf.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { forn[String(l[0])] = confFornK_(l[1]); });
  if (m >= 2) abas.la.getRange(2, 1, m - 1, 16).getValues().forEach(function (l) {
    var sku = String(l[2] || '').trim(); if (!sku) return;
    if (String(l[4]) === 'CAIXA' && Number(l[6]) > 0 && Number(l[6]) <= 120) { var kf = sku + '|' + (forn[String(l[0])] || ''), cf = cnt[kf] = cnt[kf] || {}; cf[Number(l[6])] = (cf[Number(l[6])] || 0) + 1; cf._u = Number(l[6]); } // v111.1: o fator mais usado (empate = o último); ignora digitação absurda (> 120)
    if (/BANDEJA/i.test(String(l[15] || ''))) emb[sku] = 1;
  });
  Object.keys(cnt).forEach(function (kf) { var cf = cnt[kf], best = cf._u, bn = cf[cf._u]; Object.keys(cf).forEach(function (f) { if (f !== '_u' && cf[f] > bn) { best = Number(f); bn = cf[f]; } }); fat[kf] = best; });
  var r = { fat: fat, emb: emb }; try { c.put(k, JSON.stringify(r), 600); } catch (e) {}
  return r;
}
// grava a contagem de um item: p.lotes = [{forma, qtdDig, fator, validade}] (validade 'dd/MM/yyyy' ou 'SEM VALIDADE'); p.naoVeio = true marca NÃO VEIO
function salvarConfItem(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var lotes = p.lotes || [];
  if (typeof lotes === 'string') { try { lotes = JSON.parse(lotes); } catch (e) { lotes = []; } }
  if (!p.naoVeio && !lotes.length) return { ok: false, erro: 'informe a quantidade' };
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ln = confLinhaNf_(abas.nf, p.idNf); if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length).getValues()[0];
    var rodada = Number(h[11]) || 1;
    var n = abas.it.getLastRow(), li = -1, item = null;
    var vi = n >= 2 ? abas.it.getRange(2, 1, n - 1, CAB_CIT.length).getValues() : [];
    for (var i = 0; i < vi.length; i++) if (String(vi[i][0]) === String(p.idNf) && String(vi[i][14]) === String(p.idItem)) { li = i; item = vi[i]; break; }
    if (!item) return { ok: false, erro: 'item não encontrado' };
    var posConc = String(h[5]) === 'CONCLUIDA'; // v44: correção depois de concluída -> nova rodada do item (o lançamento anterior fica guardado) e fica registrado
    if (posConc) rodada = Math.max(rodada, Number(item[11]) || 0) + 1;
    var grB = null; try { grB = gruposBonif_(vi.filter(function (r) { return String(r[0]) === String(p.idNf); }), cfopsDaNf_(abas, p.idNf))[String(p.idItem)] || null; } catch (e) { grB = null; } // v95
    if (grB) return salvarConfGrupo_(p, abas, ln, h, rodada, vi, grB, dono, posConc);
    // apaga lançamentos anteriores deste item na rodada atual
    var m = abas.la.getLastRow();
    if (m >= 2) {
      var vl = abas.la.getRange(2, 1, m - 1, CAB_CLA.length).getValues();
      for (var k = vl.length - 1; k >= 0; k--) if (String(vl[k][0]) === String(p.idNf) && String(vl[k][1]) === String(p.idItem) && Number(vl[k][12]) === rodada) abas.la.deleteRow(k + 2);
    }
    var agora = new Date(), tz = Session.getScriptTimeZone(), total = 0, rows = [];
    if (!p.naoVeio) {
      lotes.forEach(function (lt) {
        var qd = Number(String(lt.qtdDig).replace(',', '.')) || 0;
        var f = lt.fator === '' || lt.fator == null ? '' : Number(String(lt.fator).replace(',', '.'));
        var forma = String(lt.forma || 'UNIDADE').toUpperCase();
        var q = (forma === 'CAIXA' && f) ? qd * f : qd;
        total += q;
        rows.push([String(p.idNf), String(p.idItem), "'" + String(item[2]), String(item[3]), forma, qd, f, q, String(lt.validade || 'SEM VALIDADE'), dono, Utilities.formatDate(agora, tz, 'dd/MM/yyyy'), Utilities.formatDate(agora, tz, 'HH:mm'), rodada, Utilities.getUuid(), lt.pecas == null || lt.pecas === '' ? '' : Number(String(lt.pecas).replace(',', '.')), String(lt.emb || ''), lt.qtdEmb == null || lt.qtdEmb === '' ? '' : Number(lt.qtdEmb), "'" + String(lt.codLido || ''), String(lt.codStatus || ''), String(lt.autorizou || '')]);
      });
      if (rows.length) { var l0 = abas.la.getLastRow() + 1; abas.la.getRange(l0, 1, rows.length, CAB_CLA.length + EXTRA_CLA.length).setValues(rows); marcarOrigem_(abas.la, l0, rows.length, CAB_CLA.length + EXTRA_CLA.length); }
    }
    var nvSt = p.naoVeio ? (String(p.naoVeio).toUpperCase() === 'RECUSADO' ? 'RECUSADO' : 'NAO VEIO') : '';
    abas.it.getRange(li + 2, 11, 1, 4).setValues([[nvSt || 'CONFERIDO', rodada, '', nvSt ? 0 : total]]);
    var parc = null; if (!nvSt && p.parcial) { try { parc = typeof p.parcial === 'string' ? JSON.parse(p.parcial) : p.parcial; } catch (e) { parc = null; } if (parc && !(Number(parc.qtd) > 0)) parc = null; } // v46: recusa / falta parcial
    abas.it.getRange(li + 2, CAB_CIT.length + 1).setValue(nvSt ? String(p.motivo || '') : parc ? String(parc.motivo || '') : '');
    abas.it.getRange(li + 2, CAB_CIT.length + 6, 1, 2).setValues([[parc ? Number(parc.qtd) : '', parc ? (String(parc.st).toUpperCase() === 'RECUSADO' ? 'RECUSADO' : 'NAO VEIO') : '']]);
    var pcs = 0, temPc = false; if (!nvSt) lotes.forEach(function (lt) { var fo = String(lt.forma || '').toUpperCase(), pc = Number(String(lt.pecas == null ? '' : lt.pecas).replace(',', '.')); if (fo === 'KG' && pc > 0) { pcs += pc; temPc = true; } });
    abas.it.getRange(li + 2, CAB_CIT.length + EXTRA_CIT.indexOf('QTD_UNIDADES') + 1).setValue(temPc ? pcs : ''); // v118.8: unidades que chegaram (vem em KG, vende por UN) para o faturamento corrigir a nota
    // contador de conferidos no cabeçalho
    var feitos = 0; vi.forEach(function (l, j) { var st = j === li ? 'X' : String(l[10]); if (String(l[0]) === String(p.idNf) && st !== 'PENDENTE') feitos++; });
    abas.nf.getRange(ln, 6, 1, 1).setValue(String(h[5]) === 'PENDENTE' ? 'EM CONFERENCIA' : String(h[5]));
    abas.nf.getRange(ln, 11).setValue(feitos);
    var confere = true; // v40: avisa na hora se o item não bate (o app não mostra o número da nota)
    if (!nvSt) { try { confere = divergItem_(item, rows.map(function (r) { return r.map(function (c) { return typeof c === 'string' ? c.replace(/^'/, '') : c; }); }), trocasDaNf_(p.idNf), parc ? Number(parc.qtd) : 0).d !== 'S'; } catch (e) { confere = true; } }
    if (!posConc) abas.it.getRange(li + 2, 13).setValue(confere ? '' : 'S'); // v96: a divergência fica marcada já na contagem (o app continua mostrando NÃO CONFERE)
    if (posConc) {
      var obs = String(abas.nf.getRange(ln, 13).getValue() || '');
      abas.nf.getRange(ln, 13).setValue((obs ? obs + ' | ' : '') + 'CORRIGIDA APÓS CONCLUSÃO: ' + String(item[3]) + (nvSt ? ' (' + nvSt + ')' : ' (recontagem ' + total + ')') + ' por ' + dono + ' em ' + Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm'));
      abas.it.getRange(li + 2, 13).setValue(confere ? '' : 'S');
      if (!confere) abas.nf.getRange(ln, 9).setValue('DIVERGENTE');
      try { SpreadsheetApp.flush(); gravarValidades_(abas, ln, { idNf: p.idNf }, null); } catch (e) {} // v51: correção depois de concluída atualiza o CONTROLE DE VALIDADE
    }
    try { autoVincBipado_(p.idNf); } catch (e) {} // v118.5: item sem cadastro na nota, mas o código bipado está no cadastro -> vincula sozinho
    return { ok: true, feitos: feitos, total: total, confere: confere };
  } finally { lock.releaseLock(); }
}
// v40: compara um item com a nota (usado ao salvar o item — recontagem na hora, sem mostrar o número — e ao concluir a nota)
// lts = lançamentos do item na rodada: [4] forma, [5] qtd digitada, [6] fator, [7] qtd em unidades, [14] peças/caixas
function divergItem_(l, lts, trocas, falta) {
  var qNf = Number(l[6]) || 0, qAux = Number(l[7]) || 0, un = String(l[5]), fat = Number(l[8]) || 0;
  if (!fat && un !== 'KG') fat = Number(fatorDesc_(l[15])) || 0; // v34/v40: sem fator gravado -> lê da descrição da nota (display Tang 18 ou 15 na promo)
  var total = 0, formas = {}, cx = 0, pc = 0, fConf = 0;
  (lts || []).forEach(function (x) { total += Number(x[7]) || 0; formas[String(x[4])] = 1; if (String(x[4]) === 'CAIXA') { cx += Number(x[5]) || 0; if (Number(x[6]) > 0) fConf = Number(x[6]); } if (String(x[4]) === 'KGCX') { cx += Number(x[14]) || 0; if (Number(x[14]) > 0) fConf = Math.round(Number(x[7]) / Number(x[14]) * 10) / 10; } if (String(x[4]) === 'KG') pc += Number(x[14]) || 0; }); // KGCX = pesou tudo (kg) + nº de caixas
  var sub = 0; (trocas || []).forEach(function (t) { if (t.idItem === String(l[14])) sub += t.qtdSubst; });
  if (falta == null) falta = Number(l[21]) || 0; total += Number(falta) || 0; // v46: recusa/falta parcial conta para fechar com a nota
  var semComp = false;
  function cmp(total) {
    if (formas['PECAS']) { if (qAux > 0) return Math.abs(total - qAux) > 0.001 ? 'S' : ''; semComp = true; return ''; }
    if (un === 'UN' && formas['KG']) return Math.abs(pc - qNf) > 0.001 ? 'S' : ''; // v33: vem por unidade, pesado p/ o faturamento (brócolis): confere as unidades
    if (un === 'KG' || formas['KG']) return Math.abs(total - qNf) > Math.max(0.05, qNf * 0.03) ? 'S' : '';
    return (Math.abs(total - qNf) < 0.001 || (fat > 0 && Math.abs(total - qNf * fat) < 0.001) || ((un !== 'UN' || formas['CAIXA']) && Math.abs(cx - qNf) < 0.001)) ? '' : 'S'; // bate direto, por unidades (c/ fator) ou por nº de embalagens
  }
  // v42: com troca no recebimento vale tanto "contou só o que veio do item + trocados" quanto "contou junto o que veio no lugar"
  var d = cmp(total + sub); if (d === 'S' && sub > 0) d = cmp(total);
  return { d: d, semComp: semComp, fat: fat, fConf: fConf };
}
// fecha a conferência: compara com a nota e devolve só os nomes dos itens que não bateram
function concluirConf(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss); criarAbasV24_(ss);
  var trocas = trocasDaNf_(p.idNf);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ln = confLinhaNf_(abas.nf, p.idNf); if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length).getValues()[0];
    var rodada = Number(h[11]) || 1;
    var n = abas.it.getLastRow(), vi = n >= 2 ? abas.it.getRange(2, 1, n - 1, CAB_CIT.length + EXTRA_CIT.length).getValues() : [];
    var m = abas.la.getLastRow(), vl = m >= 2 ? abas.la.getRange(2, 1, m - 1, CAB_CLA.length + EXTRA_CLA.length).getValues() : [];
    var diverg = [], pend = 0, semComp = [], faltou = [], aprFator = false;
    for (var i = 0; i < vi.length; i++) {
      var l = vi[i]; if (String(l[0]) !== String(p.idNf)) continue;
      if (String(l[10]) === 'PENDENTE') { pend++; continue; }
      var rod = Number(l[11]) || 0, un = String(l[5]);
      var lts = vl.filter(function (x) { return String(x[0]) === String(p.idNf) && String(x[1]) === String(l[14]) && Number(x[12]) === rod; });
      var dv = divergItem_(l, lts, trocas), d = '';
      if (!dv.fat && dv.fConf > 0 && un !== 'UN' && un !== 'KG') { abas.it.getRange(i + 2, 9).setValue(dv.fConf); aprFator = true; } // v33: guarda kg/qtd por caixa -> vínculo
      if (String(l[10]) === 'NAO VEIO' || String(l[10]) === 'RECUSADO') { d = ''; faltou.push({ produto: String(l[3]), status: String(l[10]) }); } // não é divergência: vira devolução pendente
      else { d = dv.d; if (dv.semComp) semComp.push(String(l[3])); }
      if (Number(l[21]) > 0) faltou.push({ produto: String(l[3]), status: String(l[22]) === 'RECUSADO' ? 'RECUSADO' : 'NAO VEIO', qtd: Number(l[21]) }); // v46
      l[12] = d; if (d === 'S') diverg.push(String(l[3]));
    }
    if (pend > 0) return { ok: false, erro: 'ainda faltam ' + pend + ' item(ns) para conferir' };
    if (vi.length) abas.it.getRange(2, 13, vi.length, 1).setValues(vi.map(function (l) { return [l[12]]; }));
    var res = diverg.length ? 'DIVERGENTE' : 'OK';
    marcarOrigem_(abas.nf, ln, 1, CAB_CNF.length + EXTRA_CNF.length);
    abas.nf.getRange(ln, 6, 1, 4).setValues([['CONCLUIDA', dono, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), res]]);
    if (diverg.length) abas.nf.getRange(ln, 12).setValue(rodada + 1);
    if (aprFator) { try { aprenderVinculos_(); } catch (e) {} }
    if (p.dtRecebido) abas.nf.getRange(ln, CAB_CNF.length + 8).setValue(String(p.dtRecebido));
    var alerta = gravarCobranca_(abas, ln, p);
    SpreadsheetApp.flush(); // v50: o que importa já está gravado; o resto não pode derrubar a resposta
    posConcEnfileirar_(p, { dono: dono, res: res, diverg: diverg, semComp: semComp, alerta: alerta, rodada: rodada }); // v118.4: validades, e-mail e XML em segundo plano (o concluir responde na hora)
    return { ok: true, resultado: res, itens: diverg, naoVeio: faltou, semComparacao: semComp, alertaCobranca: alerta };
  } finally { lock.releaseLock(); }
}

// ===== v93: AVISOS DA NOTA (conferência) — colunas no fim da CONFERENCIA NF =====
// AVISO_TIPO: AVISO (texto livre) | DEVOLUCAO_PREVISTA (AVISO_ITENS começando com TOTAL = recusar a nota inteira) | DEVOLUCAO_EMITIDA (AVISO_TEXTO = nº da NF de devolução, AVISO_EM = data de emissão)
// AVISO_ITENS: "SKU|DESCRICAO|QTD|MOTIVO;..." (QTD vazia ou TUDO = o item inteiro) · AVISO_CIENTE: "quem dd/MM/yyyy HH:mm" ou "MOTORISTA LEVOU · quem dd/MM/yyyy HH:mm"
var AVISO_CNF = ['AVISO_TIPO', 'AVISO_TEXTO', 'AVISO_ITENS', 'AVISO_POR', 'AVISO_EM', 'AVISO_CIENTE'], AVISO_COL_ = 0;
function avisoCol_(sh) {
  if (AVISO_COL_) return AVISO_COL_;
  var lc = Math.max(sh.getLastColumn(), 1), cab = sh.getRange(1, 1, 1, lc).getValues()[0].map(String), c = cab.indexOf('AVISO_TIPO') + 1;
  if (!c) {
    c = lc + 1;
    var falta = c + AVISO_CNF.length - 1 - sh.getMaxColumns(); if (falta > 0) sh.insertColumnsAfter(sh.getMaxColumns(), falta);
    sh.getRange(1, c, 1, AVISO_CNF.length).setValues([AVISO_CNF]).setFontWeight('bold').setBackground('#b91c1c').setFontColor('#ffffff');
    sh.getRange(1, c).setNote('AVISO | DEVOLUCAO_PREVISTA | DEVOLUCAO_EMITIDA. Itens: SKU|DESCRICAO|QTD|MOTIVO separados por ; (TOTAL no início = nota inteira). Na EMITIDA, AVISO_TEXTO = nº da NF de devolução. Apagar o tipo tira o selo do app.');
  }
  AVISO_COL_ = c; return c;
}
function avisoItens_(s) {
  return String(s || '').split(';').map(function (x) { var p = x.split('|'); return { sku: String(p[0] || '').replace(/\D/g, ''), desc: String(p[1] || '').trim(), qtd: String(p[2] || '').trim(), motivo: String(p[3] || '').trim() }; }).filter(function (x) { return x.sku || x.desc; });
}
function avisoObj_(a) { // a = [tipo, texto, itens, por, em, ciente] -> objeto para o app (null = sem selo)
  if (!a) return null;
  var tipo = String(a[0] || '').toUpperCase().trim(); if (!tipo) return null;
  var ci = String(a[5] || '').trim(); if (tipo === 'DEVOLUCAO_EMITIDA' && /^MOTORISTA LEVOU/i.test(ci)) return null;
  var it = String(a[2] || '').trim(), total = /^TOTAL\b/i.test(it);
  return { tipo: tipo, texto: fmt_(a[1], 'dd/MM/yyyy'), itens: avisoItens_(total ? it.replace(/^TOTAL\s*;?/i, '') : it), total: total, por: String(a[3] || ''), em: fmt_(a[4], 'dd/MM/yyyy HH:mm'), ciente: ci };
}
// item da nota que está na lista de devolução prevista -> { qtd: 'TUDO' ou número, motivo }
function avisoDoItem_(av, sku, prod, descNf, qNf) {
  if (!av || av.tipo !== 'DEVOLUCAO_PREVISTA') return null;
  if (av.total) return { qtd: 'TUDO', motivo: av.texto };
  var P = semAcento_(String(prod || '')).toUpperCase().replace(/\*+/g, '').trim(), D = semAcento_(String(descNf || '')).toUpperCase().trim(), s = String(sku || '').replace(/\D/g, '');
  for (var k = 0; k < av.itens.length; k++) {
    var x = av.itens[k], X = semAcento_(x.desc).toUpperCase().replace(/\*+/g, '').trim();
    if ((x.sku && s && x.sku === s) || (X && (X === P || X === D || (X.length >= 8 && (P.indexOf(X) >= 0 || D.indexOf(X) >= 0))))) {
      var q = x.qtd, n = Number(String(q).replace(',', '.'));
      if (!q || /^(TUDO|TOTAL)$/i.test(q) || (n > 0 && Number(qNf) > 0 && n >= Number(qNf) - 0.0001)) q = 'TUDO';
      return { qtd: q, motivo: x.motivo || av.texto };
    }
  }
  return null;
}
// uma vez: OBS (coluna M) que começa com "AVISO" vira AVISO_TIPO=AVISO
function avisoMigrar_(sh, c) {
  var pr = PropertiesService.getScriptProperties(); if (pr.getProperty('AVISO_MIGRADO')) return 0;
  var n = sh.getLastRow(), k = 0;
  if (n >= 2) {
    var obs = sh.getRange(2, 13, n - 1, 1).getValues(), av = sh.getRange(2, c, n - 1, 1).getValues();
    obs.forEach(function (o, i) {
      var s = String(o[0] || '').trim(); if (!/^AVISO/i.test(s) || String(av[i][0] || '').trim()) return;
      var m = s.match(/^AVISO\s*(?:\(([^)]*)\))?\s*[:\-–—]?\s*/i), q = m && m[1] ? m[1].match(/^(.*?)\s*(\d{2}\/\d{2}\/\d{4})?\s*$/) : null;
      var txt = s.slice(m ? m[0].length : 0).split(' | ')[0].trim();
      sh.getRange(i + 2, c, 1, AVISO_CNF.length).setValues([['AVISO', txt, '', (q && q[1].trim()) || 'Celio', q && q[2] ? "'" + q[2] : '', '']]);
      k++;
    });
  }
  pr.setProperty('AVISO_MIGRADO', String(k)); return k;
}
function avisoLinha_(sh, p) {
  if (p.idNf) return confLinhaNf_(sh, p.idNf);
  var n = sh.getLastRow(); if (n < 2) return -1;
  var v = sh.getRange(2, 1, n - 1, 3).getValues(), nf = String(p.nf || '').replace(/\D/g, ''), f = semAcento_(String(p.forn || '')).toUpperCase().trim(), r = -1;
  if (!nf) return -1;
  v.forEach(function (l, i) { if (String(l[2]).replace(/\D/g, '') === nf && (!f || semAcento_(String(l[1])).toUpperCase().indexOf(f) >= 0)) r = i + 2; });
  return r;
}
// gravado pelo Claude/Celio (chave): p.idNf ou p.forn + p.nf; p.tipo (AVISO | DEVOLUCAO_PREVISTA | DEVOLUCAO_EMITIDA | LIMPAR); p.texto; p.itens (texto ou lista {sku, desc, qtd, motivo}); p.total; p.nfDev / p.emDev (emitida); p.por
function avisoGravar_(p) {
  var sh = criarAbaConf_(SpreadsheetApp.getActive()).nf, c = avisoCol_(sh), ln = avisoLinha_(sh, p);
  if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
  var tipo = String(p.tipo || '').toUpperCase().trim();
  if (!tipo || tipo === 'LIMPAR') { sh.getRange(ln, c, 1, AVISO_CNF.length).clearContent(); return { ok: true, linha: ln, limpo: true }; }
  if (['AVISO', 'DEVOLUCAO_PREVISTA', 'DEVOLUCAO_EMITIDA'].indexOf(tipo) < 0) return { ok: false, erro: 'tipo inválido' };
  var at = sh.getRange(ln, c, 1, AVISO_CNF.length).getValues()[0], it = p.itens;
  if (typeof it === 'string' && /^\s*\[/.test(it)) { try { it = JSON.parse(it); } catch (e) {} }
  if (Array.isArray(it)) it = (p.total ? ['TOTAL'] : []).concat(it.map(function (x) { return typeof x === 'string' ? x : [x.sku || '', String(x.desc || x.descricao || '').replace(/[|;]/g, ' '), x.qtd == null ? '' : x.qtd, String(x.motivo || '').replace(/[|;]/g, ' ')].join('|'); })).join(';');
  else if (it == null || it === '') it = p.total ? 'TOTAL' : (tipo === 'DEVOLUCAO_EMITIDA' ? String(at[2] || '') : '');
  else it = String(it);
  var texto = tipo === 'DEVOLUCAO_EMITIDA' ? String(p.nfDev || p.texto || '') : String(p.texto || '');
  var em = tipo === 'DEVOLUCAO_EMITIDA' && p.emDev ? String(p.emDev) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  var por = String(p.por || 'Celio');
  sh.getRange(ln, c, 1, AVISO_CNF.length).setValues([[tipo, "'" + texto, it, por, "'" + em, '']]);
  if (p.avisar) { try { var hh = sh.getRange(ln, 1, 1, 3).getValues()[0]; enviarPush_(String(p.avisar), 'Aviso na NF ' + hh[1] + ' ' + hh[2], String(p.pushTexto || 'Abra a conferência e leia o aviso da nota.').slice(0, 110), {}); } catch (e) {} } // v94: avisa a conferente no celular
  return { ok: true, linha: ln, aviso: avisoObj_([tipo, texto, it, por, em, '']) };
}
// conferente: p.o = CIENTE (deu ciência) | LEVOU (motorista levou a devolução)
function avisoMarcar_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = criarAbaConf_(SpreadsheetApp.getActive()).nf, c = avisoCol_(sh), ln = confLinhaNf_(sh, p.idNf);
  if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
  var a = sh.getRange(ln, c, 1, AVISO_CNF.length).getValues()[0], tipo = String(a[0] || '').toUpperCase().trim(), o = String(p.o || '').toUpperCase();
  if (!tipo) return { ok: false, erro: 'esta nota não tem aviso' };
  if (o === 'LEVOU' && tipo !== 'DEVOLUCAO_EMITIDA') return { ok: false, erro: 'a NF de devolução ainda não foi emitida' };
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  a[5] = (o === 'LEVOU' ? 'MOTORISTA LEVOU · ' : '') + dono + ' ' + agora;
  sh.getRange(ln, c + 5).setValue(a[5]);
  var h = sh.getRange(ln, 1, 1, 3).getValues()[0];
  try { enviarPush_('Celio', o === 'LEVOU' ? 'Devolução entregue' : 'Aviso lido — ' + h[1], o === 'LEVOU' ? h[1] + ' NF ' + h[2] + ': motorista levou (' + dono + ').' : dono + ' deu ciente no aviso da NF ' + h[2] + '.', {}); } catch (e) {}
  return { ok: true, aviso: avisoObj_(a) };
}
// devolução total prevista: marca todos os itens da nota como RECUSADO de uma vez
function recusarNota_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss), c = avisoCol_(abas.nf);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ln = confLinhaNf_(abas.nf, p.idNf); if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length).getValues()[0];
    if (String(h[5]) === 'CONCLUIDA') return { ok: false, erro: 'esta nota já foi concluída' };
    var av = avisoObj_(abas.nf.getRange(ln, c, 1, AVISO_CNF.length).getValues()[0]);
    if (!av || av.tipo !== 'DEVOLUCAO_PREVISTA' || !av.total) return { ok: false, erro: 'esta nota não tem devolução total prevista' };
    var rodada = Number(h[11]) || 1, mot = av.texto || 'devolução total prevista pelo Celio';
    var n = abas.it.getLastRow(); if (n < 2) return { ok: false, erro: 'nota sem itens' };
    var vi = abas.it.getRange(2, 1, n - 1, CAB_CIT.length + EXTRA_CIT.length).getValues(), ids = {}, cnt = 0;
    var st = vi.map(function (r) { return [r[10], r[11], r[12], r[13]]; }), mo = vi.map(function (r) { return [r[16]]; }), fa = vi.map(function (r) { return [r[21], r[22]]; });
    vi.forEach(function (r, i) { if (String(r[0]) !== String(p.idNf)) return; st[i] = ['RECUSADO', rodada, '', 0]; mo[i] = [mot]; fa[i] = ['', '']; ids[String(r[14])] = 1; cnt++; });
    if (!cnt) return { ok: false, erro: 'nota sem itens' };
    abas.it.getRange(2, 11, vi.length, 4).setValues(st);
    abas.it.getRange(2, CAB_CIT.length + 1, vi.length, 1).setValues(mo);
    abas.it.getRange(2, CAB_CIT.length + 6, vi.length, 2).setValues(fa);
    var m = abas.la.getLastRow();
    if (m >= 2) { var vl = abas.la.getRange(2, 1, m - 1, CAB_CLA.length).getValues(); for (var k = vl.length - 1; k >= 0; k--) if (String(vl[k][0]) === String(p.idNf) && ids[String(vl[k][1])] && Number(vl[k][12]) === rodada) abas.la.deleteRow(k + 2); }
    abas.nf.getRange(ln, 6).setValue('EM CONFERENCIA');
    abas.nf.getRange(ln, 11).setValue(cnt);
    return { ok: true, itens: cnt };
  } finally { lock.releaseLock(); }
}

// ===== v30: VINCULOS aprendidos (CNPJ do emitente + código do produto na nota → SKU) =====
// Item que não cruza pelo código de barras (sem EAN / "SEM GTIN" / EAN fora do cadastro) entra com "**".
// Quando alguém põe o SKU à mão na CONFERENCIA ITENS, a varredura de hora em hora grava o vínculo aqui,
// e as próximas notas do mesmo fornecedor já entram com o SKU certo.
var ABA_VINC = 'VINCULOS';
var CAB_VINC = ['FORNECEDOR', 'CÓDIGO DO PRODUTO NA NOTA', 'DESCRIÇÃO NA NOTA', 'SKU', 'DESCRIÇÃO DO SKU', 'DATA', 'FATOR']; // FATOR (v33) = kg ou unidades por caixa/saco
function criarAbaVinculos_(ss) {
  var sh = ss.getSheetByName(ABA_VINC);
  if (sh) { if (String(sh.getRange(1, 7).getValue()) !== 'FATOR') sh.getRange(1, 7).setValue('FATOR').setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff'); return sh; }
  sh = ss.insertSheet(ABA_VINC);
  sh.getRange(1, 1, 1, CAB_VINC.length).setValues([CAB_VINC]).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
  sh.getRange(2, 1, 999, 4).setNumberFormat('@');
  sh.setFrozenRows(1);
  return sh;
}
// v37: nota avulsa de produtor (mesmo CNPJ da SEFAZ, código 00000000 para tudo) -> a chave usa a descrição do item
function chaveVinc_(cnpj, cprod, desc) { var cp = String(cprod || '').trim().toUpperCase(); if (!cp || /^0+$/.test(cp)) cp = 'D:' + semAcento_(String(desc || '')).toUpperCase().replace(/\s+/g, ' ').trim(); return String(cnpj || '').replace(/\D/g, '') + '|' + cp; }
function vinculosAprendidos_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_VINC), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getDisplayValues().forEach(function (l) { var s = String(l[3]).replace(/\D/g, ''); if (l[0] && l[1] && s) { m[chaveVinc_(l[0], l[1], l[2])] = s; var f = Number(String(l[6] || '').replace(',', '.')); if (f > 0) m['F|' + chaveVinc_(l[0], l[1], l[2])] = f; } });
  return m;
}
function cnpjDoIdNf_(id) {
  id = String(id || '');
  if (/^\d{44}$/.test(id)) return id.substr(6, 14);
  var m = id.match(/^nf_(\d+)_/); return m ? m[1] : '';
}
// ===== v95: COMPRA + BONIFICADO DO MESMO PRODUTO = uma contagem só (o app separa: primeiro o bonificado, a sobra vai para a compra) =====
function baseItem_(l) { // unidade-base do item para somar: KG, ou unidades (UN, ou embalagem com fator)
  var un = String(l[5]); if (un === 'KG') return { base: 'KG', f: 1 }; if (un === 'UN') return { base: 'UN', f: 1 };
  var f = Number(l[8]) || Number(fatorDesc_(l[15])) || 0; return f > 0 ? { base: 'UN', f: f } : null;
}
// rows = linhas da CONFERENCIA ITENS desta nota; cf = CFOP por nº do item. Devolve { idItemPrincipal: { pri, itens, bon } }
function gruposBonif_(rows, cf) {
  var gk = {}, out = {};
  rows.forEach(function (l) { var k = String(l[2] || '').replace(/^'/, '') || String(l[3]); (gk[k] = gk[k] || []).push(l); });
  Object.keys(gk).forEach(function (k) {
    var g = gk[k]; if (g.length < 2) return;
    var bon = g.filter(function (l) { return /^[56]910$/.test(cf[String(l[1])] || ''); }); if (bon.length === g.length) bon = []; // v110: mesmo produto 2x na nota (preço normal + com desconto, sem bonificado) também vira uma contagem só
    var bs = g.map(baseItem_); if (bs.some(function (b) { return !b; }) || bs.some(function (b) { return b.base !== bs[0].base; })) return;
    var resto = g.filter(function (l) { return bon.indexOf(l) < 0; }).sort(function (a, b) { return (Number(a[6]) || 0) * baseItem_(a).f - (Number(b[6]) || 0) * baseItem_(b).f; }); // v110: a linha menor (desconto) enche primeiro; a maior fica com a sobra/falta
    var pri = bon.length ? g.filter(function (l) { return bon.indexOf(l) < 0; })[0] : resto[resto.length - 1];
    out[String(pri[14])] = { pri: pri, itens: bon.concat(resto), bon: bon };
  });
  return out;
}
function salvarConfGrupo_(p, abas, ln, h, rodada, vi, gr, dono, posConc) {
  var ids = {}; gr.itens.forEach(function (l) { ids[String(l[14])] = 1; });
  var m = abas.la.getLastRow();
  if (m >= 2) { var vl = abas.la.getRange(2, 1, m - 1, CAB_CLA.length).getValues(); for (var k = vl.length - 1; k >= 0; k--) if (String(vl[k][0]) === String(p.idNf) && ids[String(vl[k][1])] && Number(vl[k][12]) === rodada) abas.la.deleteRow(k + 2); }
  var lotes = p.lotes || []; if (typeof lotes === 'string') { try { lotes = JSON.parse(lotes); } catch (e) { lotes = []; } }
  var agora = new Date(), tz = Session.getScriptTimeZone(), bP = baseItem_(gr.pri);
  var nvSt = p.naoVeio ? (String(p.naoVeio).toUpperCase() === 'RECUSADO' ? 'RECUSADO' : 'NAO VEIO') : '';
  var pool = [], totalCont = 0;
  if (!nvSt) lotes.forEach(function (lt) {
    var qd = Number(String(lt.qtdDig).replace(',', '.')) || 0, f = lt.fator === '' || lt.fator == null ? 0 : Number(String(lt.fator).replace(',', '.')), forma = String(lt.forma || 'UNIDADE').toUpperCase();
    var q = forma === 'CAIXA' ? qd * (f || bP.f) : qd; totalCont += q; pool.push({ q: q, lt: lt });
  });
  var ordem = gr.bon.concat(gr.itens.filter(function (l) { return gr.bon.indexOf(l) < 0; })); // bonificado primeiro; a sobra (ou a falta) fica na compra
  var alvo = ordem.map(function (l) { return (Number(l[6]) || 0) * baseItem_(l).f; }), aloc = ordem.map(function () { return []; }), ix = 0;
  function soma(a) { return a.reduce(function (s, x) { return s + x.q; }, 0); }
  pool.forEach(function (pc) { var rest = pc.q; while (rest > 1e-9) { var last = ix === ordem.length - 1, cap = last ? Infinity : alvo[ix] - soma(aloc[ix]); if (cap <= 1e-9) { ix++; continue; } var t = Math.min(rest, cap); aloc[ix].push({ q: Math.round(t * 1000) / 1000, lt: pc.lt }); rest -= t; } });
  var rows = [], porItem = [];
  ordem.forEach(function (l, k) {
    var rr = aloc[k].map(function (a) {
      var lt = a.lt, kg = bP.base === 'KG';
      return [String(p.idNf), String(l[14]), "'" + String(l[2]).replace(/^'/, ''), String(l[3]), kg ? String(lt.forma || 'KG').toUpperCase() : 'UNIDADE', a.q, '', a.q, String(lt.validade || 'SEM VALIDADE'), dono, Utilities.formatDate(agora, tz, 'dd/MM/yyyy'), Utilities.formatDate(agora, tz, 'HH:mm'), rodada, Utilities.getUuid(), '', String(lt.emb || ''), lt.qtdEmb == null || lt.qtdEmb === '' ? '' : Number(lt.qtdEmb), "'" + String(lt.codLido || ''), String(lt.codStatus || ''), String(lt.autorizou || '')];
    });
    porItem.push(rr); rows = rows.concat(rr);
  });
  if (rows.length) { var l0 = abas.la.getLastRow() + 1; abas.la.getRange(l0, 1, rows.length, CAB_CLA.length + EXTRA_CLA.length).setValues(rows); marcarOrigem_(abas.la, l0, rows.length, CAB_CLA.length + EXTRA_CLA.length); }
  var parc = null; if (!nvSt && p.parcial) { try { parc = typeof p.parcial === 'string' ? JSON.parse(p.parcial) : p.parcial; } catch (e) { parc = null; } if (parc && !(Number(parc.qtd) > 0)) parc = null; }
  var confere = true, trocas = trocasDaNf_(p.idNf), linhaDe = {};
  vi.forEach(function (r, i) { if (String(r[0]) === String(p.idNf) && ids[String(r[14])]) linhaDe[String(r[14])] = i; });
  ordem.forEach(function (l, k) {
    var li = linhaDe[String(l[14])], ehPri = l === gr.pri, tot = soma(aloc[k]);
    abas.it.getRange(li + 2, 11, 1, 4).setValues([[nvSt || 'CONFERIDO', rodada, '', nvSt ? 0 : tot]]);
    abas.it.getRange(li + 2, CAB_CIT.length + 1).setValue(nvSt ? String(p.motivo || '') : (ehPri && parc) ? String(parc.motivo || '') : '');
    abas.it.getRange(li + 2, CAB_CIT.length + 6, 1, 2).setValues([[(ehPri && parc) ? Number(parc.qtd) : '', (ehPri && parc) ? (String(parc.st).toUpperCase() === 'RECUSADO' ? 'RECUSADO' : 'NAO VEIO') : '']]);
    var ok = true;
    if (!nvSt) { try { ok = divergItem_(l, porItem[k].map(function (r) { return r.map(function (c) { return typeof c === 'string' ? c.replace(/^'/, '') : c; }); }), trocas, (ehPri && parc) ? Number(parc.qtd) : 0).d !== 'S'; } catch (e) { ok = true; } }
    if (!ok) confere = false;
    abas.it.getRange(li + 2, 13).setValue(ok ? '' : 'S'); // v96
  });
  var feitos = 0; vi.forEach(function (l) { var st = ids[String(l[14])] && String(l[0]) === String(p.idNf) ? 'X' : String(l[10]); if (String(l[0]) === String(p.idNf) && st !== 'PENDENTE') feitos++; });
  abas.nf.getRange(ln, 6, 1, 1).setValue(String(h[5]) === 'PENDENTE' ? 'EM CONFERENCIA' : String(h[5]));
  abas.nf.getRange(ln, 11).setValue(feitos);
  if (posConc) {
    var obs = String(abas.nf.getRange(ln, 13).getValue() || '');
    abas.nf.getRange(ln, 13).setValue((obs ? obs + ' | ' : '') + 'CORRIGIDA APÓS CONCLUSÃO: ' + String(gr.pri[3]) + ' (compra + bonificado' + (nvSt ? ', ' + nvSt : ', recontagem ' + totalCont) + ') por ' + dono + ' em ' + Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm'));
    if (!confere) abas.nf.getRange(ln, 9).setValue('DIVERGENTE');
    try { SpreadsheetApp.flush(); gravarValidades_(abas, ln, { idNf: p.idNf }, null); } catch (e) {}
  }
  return { ok: true, feitos: feitos, total: totalCont, confere: confere, grupo: true };
}

// v94: CFOP de cada item (nItem) lido do XML da nota (guardado 6 h no cache)
function cfopsDaNf_(abas, idNf) {
  var ch = CacheService.getScriptCache(), k = 'cfop_' + String(idNf).slice(-30), c = ch.get(k);
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  var out = {};
  try {
    var ln = confLinhaNf_(abas.nf, idNf); if (ln < 0) return out;
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length + EXTRA_CNF.length).getValues()[0], f = null;
    var id = String(h[CAB_CNF.length + 6] || '');
    if (id) { try { f = DriveApp.getFileById(id); } catch (e) { f = null; } }
    if (!f) { var it = DriveApp.getFilesByName(String(h[13])); if (it.hasNext()) f = it.next(); }
    if (!f) return out;
    var ns = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');
    var root = XmlService.parse(f.getBlob().getDataAsString('UTF-8')).getRootElement();
    var nfe = root.getName() === 'NFe' ? root : root.getChild('NFe', ns);
    nfe.getChild('infNFe', ns).getChildren('det', ns).forEach(function (det, i) {
      var n = det.getAttribute('nItem') ? det.getAttribute('nItem').getValue() : String(i + 1);
      out[String(n)] = nfeTexto_(det.getChild('prod', ns), 'CFOP', ns);
    });
    if (Object.keys(out).length) ch.put(k, JSON.stringify(out), 21600);
  } catch (e) {}
  return out;
}
// cProd de cada item (nItem) lido do XML da nota
function cprodsDaNf_(abas, idNf) {
  var out = {};
  try {
    var ln = confLinhaNf_(abas.nf, idNf); if (ln < 0) return out;
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length + EXTRA_CNF.length).getValues()[0], f = null;
    var id = String(h[CAB_CNF.length + 6] || '');
    if (id) { try { f = DriveApp.getFileById(id); } catch (e) { f = null; } }
    if (!f) { var it = DriveApp.getFilesByName(String(h[13])); if (it.hasNext()) f = it.next(); }
    if (!f) return out;
    var ns = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');
    var root = XmlService.parse(f.getBlob().getDataAsString('UTF-8')).getRootElement();
    var nfe = root.getName() === 'NFe' ? root : root.getChild('NFe', ns);
    nfe.getChild('infNFe', ns).getChildren('det', ns).forEach(function (det, i) {
      var n = det.getAttribute('nItem') ? det.getAttribute('nItem').getValue() : String(i + 1);
      out[String(n)] = nfeTexto_(det.getChild('prod', ns), 'cProd', ns);
    });
  } catch (e) {}
  return out;
}
// aprende: item com NOVO = S e SKU preenchido à mão -> grava/atualiza o vínculo e tira o "novo" do item
function aprenderVinculos_() {
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss); criarAbasV24_(ss);
  var vs = criarAbaVinculos_(ss), n = abas.it.getLastRow();
  if (n < 2) return { ok: true, aprendidos: 0 };
  var nc = CAB_CIT.length + EXTRA_CIT.length;
  var vi = abas.it.getRange(2, 1, n - 1, nc).getValues(), cad = mapaCadastro_(), xmls = {}, novos = [], mud = false;
  var iCp = CAB_CIT.length + EXTRA_CIT.indexOf('CPROD');
  var m = vs.getLastRow(), atuais = m >= 2 ? vs.getRange(2, 1, m - 1, CAB_VINC.length).getDisplayValues() : [], pos = {};
  atuais.forEach(function (l, i) { pos[chaveVinc_(l[0], l[1], l[2])] = i; });
  var hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  vi.forEach(function (l, i) {
    var sku = String(l[2]).replace(/\D/g, '');
    if (String(l[9]) !== 'S' || !sku) { // já reconhecido: só completa o peso/qtd por caixa do vínculo (v33)
      if (sku && Number(l[8]) > 0 && String(l[5]) !== 'UN' && String(l[5]) !== 'KG') { var cp2 = String(l[iCp] || ''), k2 = chaveVinc_(cnpjDoIdNf_(l[0]), cp2, l[15]); if (cp2 && pos[k2] != null && !atuais[pos[k2]][6]) { atuais[pos[k2]][6] = String(l[8]); mud = true; } }
      return;
    }
    var idNf = String(l[0]), cnpj = cnpjDoIdNf_(idNf), cprod = String(l[iCp] || '');
    if (!cprod) { if (!xmls[idNf]) xmls[idNf] = cprodsDaNf_(abas, idNf); cprod = xmls[idNf][String(l[1])] || ''; }
    if (!cnpj || !cprod) return;
    var c = cad[sku] || {}, desc = c.desc || String(l[3]).replace(/^\*\*\s*/, '');
    var linha = [cnpj, cprod, String(l[15] || ''), sku, desc, hoje, Number(l[8]) > 0 && String(l[5]) !== 'UN' && String(l[5]) !== 'KG' ? String(l[8]) : ''], k = chaveVinc_(cnpj, cprod, l[15]);
    if (pos[k] != null && !linha[6]) linha[6] = atuais[pos[k]][6] || '';
    if (pos[k] != null) atuais[pos[k]] = linha; else { pos[k] = atuais.length; atuais.push(linha); }
    novos.push(k);
    // grava só as células que mudam (gravar a tabela inteira transformaria os códigos longos em número)
    var r = i + 2;
    abas.it.getRange(r, 10).setValue('');
    if (/^\*\*/.test(String(l[3])) && c.desc) abas.it.getRange(r, 4).setValue(c.desc);
    abas.it.getRange(r, iCp + 1).setValue("'" + cprod);
    mud = true;
  });
  if (mud) vs.getRange(2, 1, atuais.length, CAB_VINC.length).setNumberFormat('@').setValues(atuais);
  return { ok: true, aprendidos: novos.length };
}

// ===== v24: vínculos fixos, trocas, cobrança, peças, código da embalagem =====
var ABA_CVI = 'CONFERENCIA VINCULOS', ABA_CTR = 'CONFERENCIA TROCAS';
var CAB_CVI = ['CODIGO_NF', 'SKU', 'FORNECEDOR', 'OBS'];
var CAB_CTR = ['ID_NF', 'ID_ITEM', 'PRODUTO_NF', 'QTD_SUBST', 'SKU_NOVO', 'PRODUTO_NOVO', 'COD_NOVO', 'QTD_NOVO', 'VALIDADE', 'AUTORIZOU', 'QUEM', 'DATA', 'HORA', 'STATUS', 'ID'];
// colunas extras acrescentadas ao fim das abas já existentes (v23 → v24)
var EXTRA_CNF = ['COBRANCA', 'QTD_BOLETOS', 'DUP_XML', 'VENC_XML', 'BONIF', 'PLENUS', 'ID_ARQ', 'DT_RECEBIDO']; // colunas O..V da CONFERENCIA NF (PLENUS = PENDENTE quando há peças/não veio/troca para corrigir)
var EXTRA_CIT = ['MOTIVO', 'DEVOLUCAO', 'CPROD', 'DEVOLVER', 'DEV_MOTIVO', 'FALTA_QTD', 'FALTA_TIPO', 'ORIGEM_SKU', 'QTD_UNIDADES']; // v118.8: QTD_UNIDADES (Y) = unidades que chegaram de item que vem em KG na nota e é vendido por UN (aba CONFERENCIA KG PECAS); // v118.7: ORIGEM_SKU (X) = como o item foi reconhecido na importação (VINCULO, EAN, EAN CAIXA xN, EAN AMBIGUO: candidatos); // v46: FALTA_QTD (V) / FALTA_TIPO (W) = recusa ou falta PARCIAL do item; // v36: DEVOLVER (T) = qtd a devolver na unidade da nota ou TOTAL, DEV_MOTIVO (U) — devolução a pedido do Celio, gravada pelo Claude; // colunas Q..R da CONFERENCIA ITENS (motivo do não veio/recusado; nº da NF de devolução, preenchido pelo Claude)
var EXTRA_CLA = ['PECAS', 'EMBALAGEM', 'QTD_EMB', 'COD_LIDO', 'COD_STATUS', 'AUTORIZOU']; // colunas O..T da CONFERENCIA LANC (AUTORIZOU v33: validade < 30 dias)
function criarAbasV24_(ss) {
  function aba(nome, cab, seed) {
    var sh = ss.getSheetByName(nome);
    if (sh) return sh;
    sh = ss.insertSheet(nome);
    var rows = [cab].concat(seed || []);
    sh.getRange(1, 1, rows.length, cab.length).setNumberFormat('@').setValues(rows);
    sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    return sh;
  }
  var vi = aba(ABA_CVI, CAB_CVI, [
    ['7898552661144', '14749', 'NUTRIBRAS', 'carcaça suína → carne suína'],
    ['7898552662264', '36535', 'NUTRIBRAS', 'carré → bisteca'],
    ['7898552668068', '38141', 'NUTRIBRAS', 'costela inteira → costelinha'],
    ['7898552662042', '44394', 'NUTRIBRAS', 'mortadela bologna'],
    ['7898552669058', '43447', 'NUTRIBRAS', 'presunto cozido']
  ]);
  var tr = aba(ABA_CTR, CAB_CTR);
  // garante as colunas extras nas abas antigas
  var nf = ss.getSheetByName(ABA_CNF), la = ss.getSheetByName(ABA_CLA);
  if (nf) EXTRA_CNF.forEach(function (h, i) { var c = nf.getRange(1, CAB_CNF.length + 1 + i); if (String(c.getValue()) !== h) c.setValue(h).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff'); });
  var it = ss.getSheetByName(ABA_CIT);
  if (it) EXTRA_CIT.forEach(function (h, i) { var c = it.getRange(1, CAB_CIT.length + 1 + i); if (String(c.getValue()) !== h) c.setValue(h).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff'); });
  if (la && String(la.getRange(1, CAB_CLA.length + EXTRA_CLA.length).getValue()) !== EXTRA_CLA[EXTRA_CLA.length - 1]) la.getRange(1, CAB_CLA.length + 1, 1, EXTRA_CLA.length).setValues([EXTRA_CLA]).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
  return { vi: vi, tr: tr };
}
function vinculosFixos_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CVI), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (l) { var c = String(l[0]).replace(/\D/g, ''); if (c && l[1]) m[c] = String(l[1]).replace(/\D/g, ''); });
  return m;
}
// produto pelo SKU direto no mapa do cadastro (CODIGO = SKU também está no mapa)
// v118.7: código de barras -> todos os SKUs do CADASTRO com aquele código (sem zeros à esquerda)
function mapaCodMulti_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = sh ? sh.getLastRow() : 0, m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { var k = String(l[0]).replace(/\D/g, '').replace(/^0+/, ''), s = String(l[1]).replace(/\D/g, ''); if (k.length < 8 || !s) return; var a = m[k] || (m[k] = []); if (a.indexOf(s) < 0) a.push(s); });
  return m;
}
// códigos da nota (cEANTrib, cEAN) no CADASTRO: 1 cadastro -> usa; mais de 1 -> ambíguo (fica NOVO e os candidatos vão para a decisão)
function eanAchar_(cadM, cad, cods) {
  var amb = [], uni = function (s) { var c = cad[s] || {}; return Number(c.qtdEmb) > 1 && c.mestre && c.mestre !== String(c.sku) ? c.mestre : s; };
  for (var i = 0; i < cods.length; i++) {
    var k = String(cods[i] || '').replace(/\D/g, '').replace(/^0+/, ''); if (k.length < 8) continue;
    var ss = (cadM[k] || []).slice(); if (!ss.length) continue;
    if (ss.length > 1) { // o mesmo código na unidade e na caixa dela: vale a unidade
      var us = ss.filter(function (s) { return uni(s) === s; }), mm = {}; ss.forEach(function (s) { mm[uni(s)] = 1; });
      if (us.length === 1 && Object.keys(mm).length === 1) ss = us; else { if (!amb.length) amb = ss; continue; }
    }
    var c = cad[ss[0]]; if (!c) continue;
    var f = Number(c.qtdEmb) || 1;
    // (13/11, Celio: caixa/fardo continua com o SKU da caixa, como o Plenus vincula; o app já converte em unidades na contagem)
    return { c: c, caixa: 1, cands: [] };
  }
  return { c: null, caixa: 1, cands: amb };
}
// teste (rodar no editor, não grava nada): itens que hoje estão como NOVO e passariam a ser reconhecidos pelo código de barras
function testeEanConferencia() {
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss), cad = mapaCadastro_(), cadM = mapaCodMulti_(), nfs = {}, res = {};
  v9Linhas_(abas.nf, 21).forEach(function (l) { nfs[String(l[0])] = { forn: String(l[1]), nf: String(l[2]).replace(/^'/, ''), arq: String(l[20] || '') }; });
  var xmls = {}, ns = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');
  v9Linhas_(abas.it, CAB_CIT.length + EXTRA_CIT.length).forEach(function (l) {
    if (String(l[9]) !== 'S') return; var id = String(l[0]), nf = nfs[id] || {}, r = res[id] || (res[id] = { nota: (nf.forn || '') + ' ' + (nf.nf || ''), novos: 0, reconhecidos: 0, caixa: 0, ambiguos: 0, exemplos: [] });
    r.novos++;
    var cods = [String(l[4] || '').replace(/^'/, '')];
    if (nf.arq) { try { if (!xmls[id]) { var inf = XmlService.parse(DriveApp.getFileById(nf.arq).getBlob().getDataAsString('UTF-8')).getRootElement(); inf = inf.getName() === 'NFe' ? inf : inf.getChild('NFe', ns); inf = inf.getChild('infNFe', ns); xmls[id] = {}; inf.getChildren('det', ns).forEach(function (d) { var pr = d.getChild('prod', ns); xmls[id][d.getAttribute('nItem').getValue()] = [nfeTexto_(pr, 'cEANTrib', ns), nfeTexto_(pr, 'cEAN', ns)]; }); } cods = xmls[id][String(l[1])] || cods; } catch (e) {} }
    var ea = eanAchar_(cadM, cad, cods);
    if (ea.c) { r.reconhecidos++; if (ea.caixa > 1) r.caixa++; if (r.exemplos.length < 8) r.exemplos.push(String(l[15] || l[3]) + ' -> ' + ea.c.sku + ' ' + ea.c.desc + (ea.caixa > 1 ? ' (caixa x' + ea.caixa + ')' : '')); }
    else if (ea.cands.length > 1) { r.ambiguos++; if (r.exemplos.length < 8) r.exemplos.push(String(l[15] || l[3]) + ' -> AMBÍGUO: ' + ea.cands.join(' / ')); }
    else if (r.exemplos.length < 8) r.exemplos.push('SEM: ' + String(l[15] || l[3]).slice(0, 35) + ' | ' + cods.join(','));
  });
  Logger.log(JSON.stringify(res, null, 1));
  return res;
}
// teste 2 (só leitura): compara a regra antiga (último cadastro com o código) com a nova em todos os itens já importados
function testeEanConferencia2() {
  var abas = criarAbaConf_(SpreadsheetApp.getActive()), cad = mapaCadastro_(), cadM = mapaCodMulti_(), r = { codigosEmMaisDeUmCadastro: 0, itensComCodigo: 0, igual: 0, caixaViraUnidade: 0, ambiguoAntesEscolhia: 0, novoAcha: 0, nenhum: 0, exCaixa: [], exAmb: [] };
  Object.keys(cadM).forEach(function (k) { if (cadM[k].length > 1) r.codigosEmMaisDeUmCadastro++; });
  v9Linhas_(abas.it, 16).forEach(function (l) {
    var e = String(l[4] || '').replace(/^'/, '').replace(/\D/g, ''); if (e.length < 8) return; r.itensComCodigo++;
    var velho = cad[e] || null, ea = eanAchar_(cadM, cad, [e]);
    if (ea.c && velho && ea.c.sku === velho.sku) r.igual++;
    else if (ea.c && ea.caixa > 1) { r.caixaViraUnidade++; if (r.exCaixa.length < 6) r.exCaixa.push(String(l[15]).slice(0, 35) + ': ' + (velho ? velho.desc : '-') + ' -> ' + ea.c.desc + ' x' + ea.caixa + ' (hoje no item: ' + l[2] + ')'); }
    else if (ea.c && !velho) r.novoAcha++;
    else if (!ea.c && ea.cands.length > 1) { r.ambiguoAntesEscolhia++; if (r.exAmb.length < 6) r.exAmb.push(String(l[15]).slice(0, 35) + ': ' + ea.cands.join('/') + ' (antes: ' + (velho ? velho.sku : '-') + ', hoje no item: ' + l[2] + ')'); }
    else if (!ea.c) r.nenhum++;
  });
  Logger.log(JSON.stringify(r)); return r;
}
function cadPorSku_(cad, sku) { return sku && cad[String(sku)] ? cad[String(sku)] : null; }
// cobrança do XML: nº de duplicatas e vencimentos
function cobrancaXml_(inf, ns) {
  var cobr = inf.getChild('cobr', ns); if (!cobr) return { dup: 0, venc: '' };
  var dups = cobr.getChildren('dup', ns);
  return { dup: dups.length, venc: dups.map(function (d) { var v = nfeTexto_(d, 'dVenc', ns); return v ? v.slice(8, 10) + '/' + v.slice(5, 7) + '/' + v.slice(0, 4) : ''; }).filter(String).join(', ') };
}
// --- trocas no recebimento ---
function trocasDaNf_(idNf) {
  var sh = criarAbasV24_(SpreadsheetApp.getActive()).tr, n = sh.getLastRow(), out = [];
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, CAB_CTR.length).getValues().forEach(function (l) {
    if (String(l[0]) !== String(idNf)) return;
    out.push({ idItem: String(l[1]), produtoNf: String(l[2]), qtdSubst: Number(l[3]) || 0, skuNovo: String(l[4]), produtoNovo: String(l[5]), codNovo: String(l[6]), qtdNovo: Number(l[7]) || 0, validade: fmt_(l[8], 'dd/MM/yyyy'), autorizou: String(l[9]), quem: String(l[10]), status: String(l[13]), id: String(l[14]) });
  });
  return out;
}
// ===== v96: PRODUTO QUE VEIO A MAIS (fora da nota) — a conferente registra; o Claude dá entrada avulsa no Plenus depois =====
var ABA_CEXTRA = 'CONFERENCIA EXTRAS', CAB_CEXTRA = ['ID', 'ID_NF', 'FORNECEDOR', 'NF', 'SKU', 'PRODUTO', 'CODIGO', 'QTD', 'UNIDADE', 'PESO_KG', 'VALIDADE', 'OBS', 'QUEM', 'EM', 'STATUS', 'LANCAMENTO'];
function extrasAba_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ABA_CEXTRA);
  if (!sh) { sh = ss.insertSheet(ABA_CEXTRA); sh.getRange(1, 1, 1, CAB_CEXTRA.length).setValues([CAB_CEXTRA]).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff'); sh.setFrozenRows(1); }
  return sh;
}
function extrasDaNf_(idNf) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CEXTRA); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CAB_CEXTRA.length).getValues().filter(function (l) { return String(l[1]) === String(idNf) && String(l[14]) !== 'EXCLUIDO'; }).map(function (l) {
    return { id: String(l[0]), sku: String(l[4]), produto: String(l[5]), qtd: Number(l[7]) || 0, un: String(l[8]), peso: l[9] === '' ? '' : Number(l[9]), validade: fmt_(l[10], 'dd/MM/yyyy'), obs: String(l[11]), quem: String(l[12]), em: fmt_(l[13], 'dd/MM/yyyy HH:mm'), status: String(l[14]) };
  });
}
function extraAdd_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var q = Number(String(p.qtd).replace(',', '.')) || 0; if (!(q > 0)) return { ok: false, erro: 'informe a quantidade' };
  if (!String(p.produto || '').trim()) return { ok: false, erro: 'informe o produto' };
  var nfs = criarAbaConf_(SpreadsheetApp.getActive()).nf, ln = confLinhaNf_(nfs, p.idNf); if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
  var h = nfs.getRange(ln, 1, 1, 3).getValues()[0], sh = extrasAba_(), id = Utilities.getUuid(), tz = Session.getScriptTimeZone();
  var peso = String(p.peso || '').trim() === '' ? '' : Number(String(p.peso).replace(',', '.')) || '';
  var un = String(p.un || 'UN').toUpperCase();
  sh.appendRow([id, String(p.idNf), String(h[1]), "'" + String(h[2]), "'" + String(p.sku || ''), String(p.produto), "'" + String(p.codigo || ''), q, un, peso, String(p.validade || ''), String(p.obs || ''), dono, Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'), 'PENDENTE', '']);
  try { enviarPush_('Celio', 'Veio a mais — ' + h[1] + ' NF ' + h[2], (q + ' ' + un.toLowerCase() + ' ' + String(p.produto) + (peso ? ' (' + peso + ' kg)' : '') + ' · ' + dono).slice(0, 110), {}); } catch (e) {}
  return { ok: true, extras: extrasDaNf_(p.idNf) };
}
function extraDel_(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CEXTRA); if (!sh || sh.getLastRow() < 2) return { ok: false, erro: 'não encontrado' };
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (l) { return String(l[0]); }), i = ids.indexOf(String(p.id));
  if (i < 0) return { ok: false, erro: 'não encontrado' };
  if (String(sh.getRange(i + 2, 15).getValue()) !== 'PENDENTE') return { ok: false, erro: 'já lançado no Plenus' };
  sh.getRange(i + 2, 15).setValue('EXCLUIDO');
  return { ok: true, extras: extrasDaNf_(String(sh.getRange(i + 2, 2).getValue())) };
}

function salvarTroca(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var qs = Number(String(p.qtdSubst).replace(',', '.')) || 0, qn = Number(String(p.qtdNovo).replace(',', '.')) || 0;
  if (!(qs > 0) || !(qn > 0)) return { ok: false, erro: 'informe as duas quantidades' };
  var sh = criarAbasV24_(SpreadsheetApp.getActive()).tr;
  var agora = new Date(), tz = Session.getScriptTimeZone();
  var id = Utilities.getUuid();
  sh.appendRow([String(p.idNf), String(p.idItem), String(p.produtoNf || ''), qs, "'" + String(p.skuNovo || ''), String(p.produtoNovo || ''), "'" + String(p.codNovo || ''), qn, String(p.validade || 'SEM VALIDADE'), String(p.autorizou || ''), dono, Utilities.formatDate(agora, tz, 'dd/MM/yyyy'), Utilities.formatDate(agora, tz, 'HH:mm'), 'PENDENTE', id]);
  marcarOrigem_(sh, sh.getLastRow(), 1, 15);
  return { ok: true, id: id };
}
function excluirTroca(p) {
  if (!quemDoToken_(p.token)) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = criarAbasV24_(SpreadsheetApp.getActive()).tr, n = sh.getLastRow();
  if (n < 2) return { ok: false };
  var ids = sh.getRange(2, CAB_CTR.length, n - 1, 1).getValues().map(String), i = ids.indexOf(String(p.id));
  if (i < 0) return { ok: false };
  if (String(sh.getRange(i + 2, 14).getValue()) !== 'PENDENTE') return { ok: false, erro: 'já lançada no Plenus' };
  sh.deleteRow(i + 2);
  return { ok: true };
}
// e-mail para o Celio no momento em que a conferência é concluída (nunca derruba a conclusão se falhar)
var EMAIL_AVISO = 'mpauvan@gmail.com';
function emailConclusao_(abas, ln, p, dono, res, diverg, semComp, alerta, trocas, rodada) {
  try {
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length + EXTRA_CNF.length).getValues()[0];
    var forn = String(h[1]), nf = String(h[2]), bonif = String(h[CAB_CNF.length + 4] || '') === 'S';
    var tz = Session.getScriptTimeZone(), L = [];
    L.push('Conferente: ' + dono + ' · ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + ' · rodada ' + rodada);
    L.push('Nota: ' + forn + ' NF ' + nf + (bonif ? ' (BONIFICADA)' : '') + ' · emitida ' + fmt_(h[3], 'dd/MM/yyyy') + (h[4] ? ' · entrada no Plenus ' + fmt_(h[4], 'dd/MM/yyyy') : ' · AINDA SEM ENTRADA NO PLENUS') + (p.dtRecebido ? ' · recebida em ' + p.dtRecebido : ''));
    L.push('');
    if (res === 'OK') L.push('RESULTADO: TUDO CERTO'); else { L.push('RESULTADO: NÃO BATEU em ' + diverg.length + ' item(ns):'); diverg.forEach(function (d) { L.push('  - ' + d); }); }
    var n = abas.it.getLastRow(), vi = n >= 2 ? abas.it.getRange(2, 1, n - 1, CAB_CIT.length + EXTRA_CIT.length).getValues() : [];
    var nv = [], novos = [];
    vi.forEach(function (l) { if (String(l[0]) !== String(p.idNf)) return; if (String(l[10]) === 'NAO VEIO' || String(l[10]) === 'RECUSADO') nv.push(String(l[3]) + (String(l[10]) === 'RECUSADO' ? ' (RECUSADO)' : '') + (l[16] ? ' — ' + l[16] : '')); if (String(l[9]) === 'S') novos.push(String(l[3]).replace(/^\*\*\s*/, '') + ' (cód. ' + String(l[4]) + ')'); });
    vi.forEach(function (l) { if (String(l[0]) === String(p.idNf) && Number(l[21]) > 0) nv.push(String(l[3]) + ' (' + (String(l[22]) === 'RECUSADO' ? 'recusou ' : 'faltou ') + l[21] + (String(l[5]) === 'KG' ? ' kg' : ' un') + ')' + (l[16] ? ' — ' + l[16] : '')); }); // v46: parcial
    if (nv.length) L.push('NÃO VEIO / RECUSADO (devolução pendente): ' + nv.join('; '));
    var devs = []; vi.forEach(function (l) { if (String(l[0]) === String(p.idNf) && String(l[19] || '')) devs.push(String(l[3]) + ': ' + (String(l[19]).toUpperCase() === 'TOTAL' ? 'devolver tudo' : 'devolver ' + l[19] + ' ' + String(l[5])) + (l[20] ? ' — ' + l[20] : '')); });
    if (devs.length) L.push('DEVOLUÇÃO A PEDIDO DO CELIO: ' + devs.join('; '));
    if (semComp.length) L.push('Sem comparação (peças, nota sem Qtde_aux): ' + semComp.join('; '));
    var m = abas.la.getLastRow(), vl = m >= 2 ? abas.la.getRange(2, 1, m - 1, CAB_CLA.length + EXTRA_CLA.length).getValues() : [];
    var pecas = [], cods = [], pesos = [], curtas = [], cadE = mapaCadastro_();
    vl.forEach(function (l) {
      if (String(l[0]) !== String(p.idNf) || Number(l[12]) !== rodada) return;
      var unCad = (cadE[String(l[2]).replace(/\D/g, '')] || {}).un || '';
      if (String(l[4]) === 'KGCX') pesos.push(String(l[3]) + ': ' + l[14] + ' cx = ' + l[7] + ' kg (pesado)');
      else if (String(l[4]) === 'CAIXA' && /QUILO/.test(unCad)) pesos.push(String(l[3]) + ': ' + l[5] + ' cx × ' + l[6] + ' kg = ' + l[7] + ' kg');
      else if (l[14] !== '' && l[14] != null) pecas.push(String(l[3]) + ': ' + l[7] + ' kg / ' + l[14] + ' pç' + (l[8] && String(l[8]) !== 'SEM VALIDADE' ? ' (val. ' + fmt_(l[8], 'dd/MM/yyyy') + ')' : ''));
      if (l[19]) curtas.push(String(l[3]) + ': val. ' + fmt_(l[8], 'dd/MM/yyyy') + ' · autorizou: ' + (String(l[19]) === 'NINGUEM' ? 'NINGUÉM ⚠' : String(l[19])));
      if (String(l[18]) === 'NOVO') cods.push(String(l[3]) + ': cód. ' + String(l[17]) + (l[15] ? ' · ' + l[15] + (l[16] ? ' ×' + l[16] : '') : ''));
    });
    if (curtas.length) { L.push(''); L.push('VALIDADE CURTA (menos de 30 dias):'); curtas.forEach(function (x) { L.push('  - ' + x); }); }
    if (pesos.length) { L.push(''); L.push('PESO REAL (nota em caixa, vendido por kg no Plenus):'); pesos.forEach(function (x) { L.push('  - ' + x); }); }
    if (pecas.length) { L.push(''); L.push('PEÇAS (corrigir a quantidade no faturamento):'); pecas.forEach(function (x) { L.push('  - ' + x); }); }
    if (trocas && trocas.length) { L.push(''); L.push('TROCAS NO RECEBIMENTO (lançar Movimentação Avulsa quando autorizar):'); trocas.forEach(function (t) { L.push('  - ' + t.qtdSubst + ' × ' + t.produtoNf + ' → ' + t.qtdNovo + ' × ' + t.produtoNovo + (t.skuNovo ? ' (SKU ' + t.skuNovo + ')' : '') + ' · autorizou: ' + t.autorizou); }); }
    // pendência no Plenus: peças (quantidade real), não veio, trocas → Celio autoriza e o Claude corrige/fatura
    var nPend = pecas.length + nv.length + (trocas ? trocas.length : 0);
    abas.nf.getRange(ln, CAB_CNF.length + 6).setValue(nPend ? 'PENDENTE' : '');
    if (cods.length) { L.push(''); L.push('CÓDIGOS DE EMBALAGEM NÃO CADASTRADOS:'); cods.forEach(function (x) { L.push('  - ' + x); }); }
    if (novos.length) { L.push(''); L.push('PRODUTOS NOVOS (sem cadastro):'); novos.forEach(function (x) { L.push('  - ' + x); }); }
    var cob = String(p.cobranca || '').toUpperCase();
    if (cob) { L.push(''); L.push('Cobrança: ' + (cob === 'BOLETO' ? (Number(p.qtdBoletos) || 0) + ' boleto(s)' : cob)); }
    if (alerta) L.push('⚠ ' + alerta);
    L.push(''); L.push('Planilha: https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActive().getId());
    MailApp.sendEmail(EMAIL_AVISO, 'Conferência NF ' + nf + ' ' + forn + ' — ' + (res === 'OK' ? 'TUDO CERTO' : 'NÃO BATEU') + (bonif ? ' (bonificada)' : '') + (nPend ? ' · ⚠ ' + nPend + ' item(ns) para corrigir no Plenus' : ''), L.join(String.fromCharCode(10)));
  } catch (e) {}
}
// grava a cobrança informada na conclusão e compara com o XML
function gravarCobranca_(abas, ln, p) {
  var cob = String(p.cobranca || '').toUpperCase(), qb = Number(p.qtdBoletos) || 0;
  if (!cob) return null;
  var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length + EXTRA_CNF.length).getValues()[0];
  var vx = h[CAB_CNF.length + 3], dup = Number(h[CAB_CNF.length + 2]) || 0, venc = vx instanceof Date ? dataBR_(vx) : String(vx || ''); // v92: data sempre dd/mm/aaaa
  var bonif = String(h[CAB_CNF.length + 4] || '').toUpperCase().trim() === 'S';
  abas.nf.getRange(ln, CAB_CNF.length + 1, 1, 2).setValues([[cob, cob === 'BOLETO' ? qb : '']]);
  if (bonif && dup > 0) { // v92: bonificada com duplicata no XML — não cobra o conferente, avisa o Celio
    try { enviarPush_('Celio', 'NF bonificada com cobrança', h[1] + ' NF ' + h[2] + ': XML tem ' + dup + ' duplicata(s), venc. ' + venc + '. Confirme com o fornecedor.', {}); } catch (e) {}
    return 'Nota bonificada: o XML veio com cobrança (venc. ' + venc + '), mas não precisa de boleto. O Celio vai confirmar com o fornecedor.';
  }
  if (dup > 0 && cob !== 'BOLETO') return 'A nota tem ' + dup + ' boleto(s) no XML (venc. ' + venc + ') e foi informado ' + (cob === 'CHEQUE' ? 'CHEQUE' : 'NENHUM') + '. Confira com o motorista.';
  if (dup > 0 && cob === 'BOLETO' && qb !== dup) return 'A nota tem ' + dup + ' boleto(s) no XML e foram informados ' + qb + '. Confira com o motorista.';
  if (dup === 0 && cob === 'BOLETO') return 'O XML desta nota não tem boleto, mas foi informado boleto. Avise o Celio.';
  return null;
}

// ===== v26: varredura da pasta XML NF, pastas RECEBIDAS, validades na planilha VALIDADE PRODUTOS, gatilhos =====
var ID_PASTA_XML = '1ehLVmz1jgJNtPkXApm3IWldYwwxGU_EN';          // ARQUIVOS CLAUDE / XML NF
var ID_CONTROLE_NF = '1W2W7QTqeR_KCGXuk7c5ScaCFk-hI3Zk092DNaCJxFOM'; // planilha CONTROLE NF (abas OUT-26, NOV-26...)
var ID_VALIDADE = '1Xyf6noNgTVSmmPHyV-H7VIeJhss7olbD6WRKXNHKb5I';   // planilha VALIDADE PRODUTOS (aba PRODUTOS)
var PASTA_RECEBIDAS = 'RECEBIDAS';

// Varre a pasta XML NF (só a raiz) e carrega no app toda nota que ainda não está na CONFERENCIA NF.
// Fornecedor da lista de exceções: a nota é registrada como EXCECAO e o XML vai direto para RECEBIDAS.
function varrerXml() {
  var apr = null; try { apr = aprenderVinculos_(); } catch (e) { apr = { ok: false, erro: String(e) }; }
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss); criarAbasV24_(ss);
  var n = abas.nf.getLastRow(), vistos = {}, dsc = desconhecidasControle_(), tiradas = [], ent = null, entAtual = [];
  if (n >= 2) abas.nf.getRange(2, 1, n - 1, CAB_CNF.length + EXTRA_CNF.length).getValues().forEach(function (l, i) {
    vistos[String(l[13])] = 1; if (l[CAB_CNF.length + 6]) vistos[String(l[CAB_CNF.length + 6])] = 1;
    // v38: nota pendente que foi marcada como desconhecida na CONTROLE NF sai da conferência
    // v45: nota que ainda está sem data de entrada no app -> tenta de novo pela CONTROLE NF (coluna DT ENTRADA)
    if (!l[4] && String(l[5]) !== 'EXCECAO') { if (!ent) ent = entradasControle_(); var de = acharEntrada_(ent, l[2], l[1]); if (de) { abas.nf.getRange(i + 2, 5).setValue(de); entAtual.push(String(l[1]) + ' ' + String(l[2]) + ' = ' + de); } }
    if (String(l[5]) === 'PENDENTE' && ehDesconhecida_(dsc, l[2], l[1])) { abas.nf.getRange(i + 2, 6).setValue('EXCECAO'); abas.nf.getRange(i + 2, 13).setValue('nota desconhecida (CONTROLE NF)'); tiradas.push(String(l[1]) + ' ' + String(l[2])); }
  });
  var it = DriveApp.getFolderById(ID_PASTA_XML).getFiles(), out = [];
  while (it.hasNext()) {
    var f = it.next(), nome = f.getName();
    if (!/\.xml$/i.test(nome) || vistos[nome] || vistos[f.getId()]) continue;
    try {
      var r = carregarNF({ chave: CHAVE_CARGA, arquivo: f.getId() });
      if (r.ok && r.excecao) moverRecebidaArq_(f);
      out.push({ arquivo: nome, ok: r.ok, itens: r.itens, novos: r.novos, excecao: !!r.excecao, erro: r.erro });
    } catch (e) { out.push({ arquivo: nome, ok: false, erro: String(e) }); }
  }
  var devs = []; try { devs = vincularDevolucoes_(abas); } catch (e) { devs = ['erro: ' + e]; } // v48
  return { ok: true, carregadas: out, desconhecidas: tiradas, entradas: entAtual, devolucoes: devs, vinculos: apr };
}
// v48: devolução pendente (não veio / recusado, inteiro ou parcial, ou devolução a pedido do Celio) sem número de NF de devolução ->
// procura no Drive o PDF da NF de devolução que cita a chave da nota (Meu Drive > NF DEVOLUCAO > fornecedor, nome "FORNECEDOR 3073.pdf")
// e grava em DEVOLUCAO (R): "NF de devolução 3073 emitida em 24/10/2026"
function vincularDevolucoes_(abas) {
  var n = abas.it.getLastRow(); if (n < 2) return [];
  var v = abas.it.getRange(2, 1, n - 1, CAB_CIT.length + EXTRA_CIT.length).getValues(), porNf = {}, out = [], achados = {};
  v.forEach(function (l, i) {
    var id = String(l[0]); if (!/^\d{44}$/.test(id) || String(l[17] || '')) return;
    var st = String(l[10]);
    if (st === 'NAO VEIO' || st === 'RECUSADO' || Number(l[21]) > 0 || String(l[19] || '')) (porNf[id] = porNf[id] || []).push(i);
  });
  Object.keys(porNf).forEach(function (id) {
    if (achados[id] === undefined) {
      achados[id] = '';
      var it = DriveApp.searchFiles("fullText contains '" + id + "' and mimeType = 'application/pdf' and trashed = false"), best = null;
      var nfOrig = String(Number(id.substr(25, 9)));
      while (it.hasNext()) {
        var f = it.next(), m = f.getName().match(/(\d{2,})\s*\.pdf$/i); if (!m || String(Number(m[1])) === nfOrig) continue; // não confundir com o DANFE da própria nota
        var dev = false, ps = f.getParents(); while (ps.hasNext() && !dev) { var pa = ps.next(); if (/DEVOLU/i.test(pa.getName())) dev = true; else { var av = pa.getParents(); while (av.hasNext()) if (/DEVOLU/i.test(av.next().getName())) dev = true; } }
        if (dev && (!best || f.getDateCreated() < best.d)) best = { n: m[1].replace(/^0+/, ''), d: f.getDateCreated() };
      }
      if (best) achados[id] = 'NF de devolução ' + best.n + ' emitida em ' + Utilities.formatDate(best.d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    if (!achados[id]) return;
    porNf[id].forEach(function (i) { abas.it.getRange(i + 2, CAB_CIT.length + 2).setValue(achados[id]); out.push(String(v[i][3]) + ' -> ' + achados[id]); });
  });
  return out;
}
// v45: todas as datas de entrada da CONTROLE NF numa leitura só: nº da NF -> [{ap, data}]
function entradasControle_() {
  var m = {};
  try {
    SpreadsheetApp.openById(ID_CONTROLE_NF).getSheets().forEach(function (sh) {
      if (!/^[A-Z]{3}-\d{2}$/.test(sh.getName()) || sh.getLastRow() < 2 || sh.getLastColumn() < 14) return;
      sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues().forEach(function (l) {
        var nf = String(l[2]).replace(/\D/g, ''); if (!nf || !l[13]) return;
        (m[nf] = m[nf] || []).push({ ap: semAcento_(String(l[0])).toUpperCase().trim(), data: fmt_(l[13], 'dd/MM/yyyy') });
      });
    });
  } catch (e) {}
  return m;
}
function acharEntrada_(m, nNF, apelido) {
  var l = m[String(nNF).replace(/\D/g, '')]; if (!l) return '';
  var ap = semAcento_(String(apelido || '')).toUpperCase().trim();
  for (var i = 0; i < l.length; i++) if (!ap || !l[i].ap || l[i].ap.indexOf(ap) >= 0 || ap.indexOf(l[i].ap) >= 0) return l[i].data;
  return '';
}
// v38: números das notas com "desconhecid" na coluna OBS (M) de qualquer aba MÊS-AA da CONTROLE NF
function desconhecidasControle_() {
  var m = {};
  try {
    SpreadsheetApp.openById(ID_CONTROLE_NF).getSheets().forEach(function (sh) {
      if (!/^[A-Z]{3}-\d{2}$/.test(sh.getName()) || sh.getLastRow() < 2) return;
      sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues().forEach(function (l) {
        if (/DESCONHECID/.test(semAcento_(String(l[12] || '')).toUpperCase())) { var nf = String(l[2]).replace(/\D/g, ''); if (nf) (m[nf] = m[nf] || []).push(semAcento_(String(l[0])).toUpperCase().trim()); }
      });
    });
  } catch (e) {}
  return m;
}
function ehDesconhecida_(m, nNF, apelido) {
  var l = m[String(nNF).replace(/\D/g, '')]; if (!l) return false;
  var ap = semAcento_(String(apelido || '')).toUpperCase().trim();
  return l.some(function (a) { return !ap || !a || a.indexOf(ap) >= 0 || ap.indexOf(a) >= 0; });
}
// data de entrada no Plenus lida da CONTROLE NF: coluna C = NF, coluna N = DT ENTRADA (procura em todas as abas MÊS-AA)
function entradaControle_(nNF, apelido) {
  try {
    var ss = SpreadsheetApp.openById(ID_CONTROLE_NF), nf = String(nNF).replace(/\D/g, ''), ap = semAcento_(String(apelido || '')).toUpperCase();
    var abasMes = ss.getSheets().filter(function (s) { return /^[A-Z]{3}-\d{2}$/.test(s.getName()); }).reverse();
    for (var k = 0; k < abasMes.length; k++) {
      var sh = abasMes[k], m = sh.getLastRow(); if (m < 2) continue;
      var v = sh.getRange(2, 1, m - 1, 14).getValues();
      for (var i = v.length - 1; i >= 0; i--) {
        if (String(v[i][2]).replace(/\D/g, '') !== nf) continue;
        var a = semAcento_(String(v[i][0])).toUpperCase();
        if (ap && a && a.indexOf(ap) < 0 && ap.indexOf(a) < 0) continue;
        return v[i][13] ? fmt_(v[i][13], 'dd/MM/yyyy') : '';
      }
    }
  } catch (e) {}
  return '';
}
// move o XML da nota concluída para XML NF / RECEBIDAS / aaaa-MM
function moverRecebida_(abas, ln) {
  try {
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length + EXTRA_CNF.length).getValues()[0];
    var id = String(h[CAB_CNF.length + 6] || ''), f = null;
    if (id) { try { f = DriveApp.getFileById(id); } catch (e) { f = null; } }
    if (!f) { var it = DriveApp.getFolderById(ID_PASTA_XML).getFilesByName(String(h[13])); if (it.hasNext()) f = it.next(); }
    if (f) moverRecebidaArq_(f);
  } catch (e) {}
}
function moverRecebidaArq_(f) {
  var raiz = DriveApp.getFolderById(ID_PASTA_XML);
  var it = raiz.getFoldersByName(PASTA_RECEBIDAS), rec = it.hasNext() ? it.next() : raiz.createFolder(PASTA_RECEBIDAS);
  var mes = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var it2 = rec.getFoldersByName(mes), sub = it2.hasNext() ? it2.next() : rec.createFolder(mes);
  f.moveTo(sub);
}
// grava as validades contadas na aba PRODUTOS da planilha VALIDADE PRODUTOS (uma linha por validade; SEM VALIDADE não entra)
// colunas: N NF, DATA VAL, SESSAO, PRODUTOS, UN, CHEGOU, DT CHEGOU, FORNECEDOR, ACHOU
// v53: regrava na VALIDADE PRODUTOS as validades de uma nota já concluída (correção manual pela rotina de NF)
function regravarValidades_(idNf, diag) {
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss); criarAbasV24_(ss);
  var ln = confLinhaNf_(abas.nf, String(idNf || '')); if (ln < 0) return { ok: false, erro: 'nota não encontrada' };
  if (String(abas.nf.getRange(ln, 6).getValue()) !== 'CONCLUIDA') return { ok: false, erro: 'nota não concluída' };
  var r = gravarValidades_(abas, ln, { idNf: String(idNf) }, Number(abas.nf.getRange(ln, 12).getValue()) || 1, !!diag);
  return { ok: true, r: r || null };
}
function gravarValidades_(abas, ln, p, rodada, diag) {
  try {
    var h = abas.nf.getRange(ln, 1, 1, CAB_CNF.length).getValues()[0];
    var nNF = String(h[2]).replace(/\D/g, ''), forn = String(h[1]);
    var m = abas.la.getLastRow(), vl = m >= 2 ? abas.la.getRange(2, 1, m - 1, CAB_CLA.length + 1).getValues() : []; // v123.4: + PECAS (O)
    // v51: cada item vale pela rodada atual dele (item recontado numa reconferência ou corrigido depois de concluída)
    var rodItem = {}, ni = abas.it.getLastRow();
    var itRow = {}; // v112.1: linha do item (unidade da nota, fator, descrição, código) para converter em unidade de venda
    if (ni >= 2) abas.it.getRange(2, 1, ni - 1, CAB_CIT.length + EXTRA_CIT.length).getValues().forEach(function (l) { if (String(l[0]) === String(p.idNf)) { rodItem[String(l[14])] = Number(l[11]) || 0; itRow[String(l[14])] = l; } });
    var dtr = String(p.dtRecebido || ''); if (!dtr) { try { dtr = fmt_(abas.nf.getRange(ln, CAB_CNF.length + 8).getValue(), 'dd/MM/yyyy'); } catch (e) { dtr = ''; } }
    // v118.4: nota sem nenhuma validade lançada (hortifruti, sem validade) e 1ª rodada: não abre a VALIDADE PRODUTOS
    if (!diag && Number(rodada) === 1 && !vl.some(function (l) { return String(l[0]) === String(p.idNf) && /^\d{2}\/\d{2}\/\d{4}$/.test(fmt_(l[8], 'dd/MM/yyyy')); })) return { apagadas: 0, gravadas: 0, semValidade: true };
    // v53: as datas são montadas no fuso da planilha VALIDADE PRODUTOS (ela está em outro fuso; meia-noite daqui virava o dia anterior lá)
    var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone();
    var dV = function (d, m, a) { return Utilities.parseDate(d + '/' + m + '/' + a, tzV, 'dd/MM/yyyy'); };
    var cadV = mapaCadastro_(), aprV = vinculosAprendidos_(), cnpjV = cnpjDoIdNf_(p.idNf), convV = [], confV = [], estV = []; // v104 / v112.1
    var linhas = [], qUn = [], hoje = Utilities.parseDate(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'), tzV, 'dd/MM/yyyy'), md = String(dtr || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (md) hoje = dV(md[1], md[2], md[3]);
    vl.forEach(function (l) {
      if (String(l[0]) !== String(p.idNf)) return;
      var rIt = rodItem[String(l[1])]; if (Number(l[12]) !== (rIt != null ? rIt : rodada)) return;
      var val = fmt_(l[8], 'dd/MM/yyyy'), mv = val.match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if (!mv) return;
      qUn.push(0); // v112.1: a conversão para unidade agora é feita aqui (valLote_)
      var vc = valLote_(itRow[String(l[1])], l, cadV, aprV, cnpjV), ub = unidadeBase_(cadV, l[2]); // v104: nome da unidade (mestre) quando existe
      var nomeV = ub && vc.un !== 'KG' ? ub.desc : String(l[3]);
      if (vc.un === 'CONFERIR') confV.push({ sku: valSku_(l[2]), prod: nomeV, qtd: vc.qtd, motivo: vc.motivo });
      convV.push(vc.conv || ''); estV.push(vc.est || '');
      if (vc.est) confV.push({ sku: valSku_(l[2]), prod: nomeV, qtd: vc.qtd, motivo: vc.est + ' = ' + vc.qtd + ' peças — confirme a contagem' });
      linhas.push([nNF, dV(mv[1], mv[2], mv[3]), '', nomeV, vc.un, vc.qtd, hoje, forn, '']);
    });
    var sh = ssV.getSheetByName('PRODUTOS'), last = sh.getLastRow();
    var dados = last >= 2 ? sh.getRange(2, 1, last - 1, 9).getValues() : [];
    // seção pelo histórico do mesmo produto; linhas antigas desta mesma NF (reconferência) saem antes de regravar
    var sessao = {}, apagar = [];
    dados.forEach(function (l, i) {
      var k = String(l[3]).replace(/\*+\s*$/, '').trim().toUpperCase(); if (k && l[2]) sessao[k] = String(l[2]);
      if (String(l[0]).replace(/\D/g, '') === nNF && semAcento_(String(l[7])).toUpperCase() === semAcento_(forn).toUpperCase()) apagar.push(i + 2);
    });
    linhas.forEach(function (l) { l[2] = sessao[String(l[3]).replace(/\*+\s*$/, '').trim().toUpperCase()] || ''; });
    // v86: refrigerante/bebida que chegou em fardo ou caixa entra na validade por UNIDADE (quantidade x fator)
    // v104 (regra do Celio, 07/11/2026): vale para TODO produto contado em caixa/fardo: a validade é sempre em unidades
    linhas.forEach(function (l, i) { if (l[4] === 'CX' && qUn[i] > 0) { l[4] = 'UN'; l[5] = qUn[i]; } });
    if (!diag) { var nApag = {}; apagar.forEach(function (x) { nApag[x] = 1; }); tfGerar_(ssV, tzV, dados.filter(function (l, i) { return !nApag[i + 2]; }), linhas, nNF, forn); } // v81: Troca de frente (PVPS)
    if (diag) return { last: last, nNF: nNF, forn: forn, linhas: linhas.length, apagar: apagar.length, det: linhas.map(function (l, i) { return l[3] + ' | ' + l[4] + ' ' + l[5] + (convV[i] ? ' (' + convV[i] + ')' : ''); }), conferir: confV, amostra: dados.filter(function (l) { return String(l[0]).indexOf(nNF) >= 0; }).slice(0, 2).map(function (l) { return [typeof l[0], String(l[0]), String(l[7])]; }) };
    // v55: apaga em blocos contínuos (a aba PRODUTOS virou tabela; linha a linha é lento)
    for (var i = apagar.length - 1; i >= 0; ) { var fim = apagar[i], ini = fim; while (i > 0 && apagar[i - 1] === ini - 1) { i--; ini--; } sh.deleteRows(ini, fim - ini + 1); i--; }
    // v56: última linha com PRODUTO preenchido (a aba fica com linhas em branco no fim para lançar à mão; elas continuam lá embaixo)
    var colD = sh.getRange(1, 4, sh.getMaxRows(), 1).getValues(); last = 1;
    for (var k = colD.length - 1; k >= 1; k--) if (String(colD[k][0]).trim()) { last = k + 1; break; }
    if (!linhas.length) return;
    sh.insertRowsAfter(last, linhas.length);
    sh.getRange(last + 1, 1, linhas.length, 9).setValues(linhas); estV.forEach(function (e, i) { if (e) sh.getRange(last + 1 + i, 12).setValue(e); }); // v123.4: L = ESTIMADO
    // v55: sem setNumberFormat — as colunas de data da tabela já são do tipo DATA (dava erro)
    SpreadsheetApp.flush();
    cvCurta_(linhas, tzV); // v99: chegou com validade curta -> lista da repositora, no topo
    try { valPendencias_(nNF, forn, confV); } catch (e) {} // v112.1: sem conversão segura -> lista do Celio
    return { apagadas: apagar.length, gravadas: linhas.length };
  } catch (e) { return { erro: String(e) }; }
}
// gatilho de hora em hora para varrer a pasta XML NF (chamar uma vez via ?acao=instalarGatilhos&chave=...)
function instalarGatilhos() {
  var tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'varrerXml'; });
  if (!tem) ScriptApp.newTrigger('varrerXml').timeBased().everyHours(1).create();
  var gs = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (gs.indexOf('sincronizarRebaixas') < 0) ScriptApp.newTrigger('sincronizarRebaixas').timeBased().atHour(8).everyDays(1).create();
  if (gs.indexOf('atualizarVencidos') < 0) ScriptApp.newTrigger('atualizarVencidos').timeBased().atHour(5).everyDays(1).create(); // v79
  if (gs.indexOf('fecharRebaixasDoDia') < 0) ScriptApp.newTrigger('fecharRebaixasDoDia').timeBased().atHour(23).everyDays(1).create();
  return { ok: true, instalado: !tem, rebaixas: true };
}

// ========== VENCIDOS DA SEMANA (v79) ==========
// Lista semanal para a Thynaylly (módulo AVARIA / VENCIDO). Fonte: VALIDADE PRODUTOS, aba POR CHEGADA
// (produto + DATA VAL), com seção e local pela aba CADASTRO. Regras do Celio (02/11/2026):
// - seções LATICÍNIOS E FRIOS e FRIOS E EMBUTIDOS, e na MATINAL só as bebidas lácteas (pelo nome): retirar 3 dias antes do vencimento;
// - todas as demais seções: retirar no dia do vencimento;
// - se o dia de retirar cair no sábado ou domingo, retira na sexta;
// - semana = segunda a domingo; o que ficou PENDENTE da semana anterior vai para o topo da lista nova como ATRASADO;
// - produto sem seção no CADASTRO entra no bloco "Sem seção".
var ABA_VS = 'VENCIDOS SEMANA';
var CAB_VS = ['SEMANA', 'RETIRAR ATE', 'VALIDADE', 'PRODUTO', 'SECAO', 'LOCAL', 'REGRA', 'ATRASADO', 'STATUS', 'QTD', 'QUEM', 'EM', 'ID AVARIA', 'CHAVE', 'ID', 'SKU', 'UN', 'VALIDADE DIGITADA', 'DIVERGENCIA'];
var VS_SECOES_FRIAS = ['LATICINIOS E FRIOS', 'FRIOS E EMBUTIDOS'];
var VS_LACTEOS_MATINAL = /BEBIDA LACTEA|BEB\.? LACTEA|TODDYNHO|YOPRO|WHEY|PIRAKIDS/;
var VS_SECOES_FORA = ['ACOUGUE', 'FRACIONADOS', 'PRODUCAO']; // Celio 02/11/2026: açougue, fracionados e produção (padaria) não vão para a lista da Thynaylly
function vsFora_(secao) { return VS_SECOES_FORA.indexOf(semAcento_(String(secao || '')).toUpperCase().trim()) >= 0; }
function criarAbaVS_(ss) {
  var sh = ss.getSheetByName(ABA_VS);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_VS);
  sh.getRange(1, 1, 1, CAB_VS.length).setValues([CAB_VS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:C').setNumberFormat('@'); sh.getRange('L:L').setNumberFormat('@'); sh.getRange('P:P').setNumberFormat('@');
  return sh;
}
function vsDia_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function vsBr_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
function vsDeBr_(s) { var p = String(s || '').split('/'); return p.length === 3 ? new Date(+p[2], +p[1] - 1, +p[0]) : null; }
function vsSegunda_(d) { d = vsDia_(d); var w = d.getDay(); return new Date(d.getFullYear(), d.getMonth(), d.getDate() - ((w + 6) % 7)); }
function vsMais_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function vsRegra_(secao, prod) {
  var s = semAcento_(String(secao || '')).toUpperCase().trim(), p = normDesc_(prod);
  if (VS_SECOES_FRIAS.indexOf(s) >= 0) return 'FRIO';
  if (s === 'MATINAL' && VS_LACTEOS_MATINAL.test(p)) return 'FRIO';
  return 'DIA';
}
function vsRetirar_(val, regra) { // dia em que tem que sair da loja
  var r = vsMais_(val, regra === 'FRIO' ? -3 : 0), w = r.getDay();
  if (w === 6) r = vsMais_(r, -1); else if (w === 0) r = vsMais_(r, -2);
  return r;
}
function vsCadastroValidade_(ssV) {
  var sh = ssV.getSheetByName('CADASTRO'), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 3).getValues().forEach(function (l) { var k = normDesc_(l[0]); if (k) m[k] = { secao: String(l[1] || '').trim(), local: String(l[2] || '').trim() }; });
  return m;
}
function vsSkuPorNome_() { // cadastro do Plenus (aba CADASTRO desta planilha) por descrição
  var cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = cad ? cad.getLastRow() : 0, m = {};
  if (n >= 2) cad.getRange(2, 2, n - 1, 3).getValues().forEach(function (l) { var k = normDesc_(l[1]); if (k && !m[k]) m[k] = { sku: String(l[0]), un: String(l[2] || '') }; });
  return m;
}
// Monta / atualiza a lista da semana. Idempotente: pode rodar todo dia (gatilho) ou pelo Claude (acao=montarVencidos).
function atualizarVencidos() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActive(), sh = criarAbaVS_(ss), hoje = vsDia_(new Date());
    var seg = vsSegunda_(hoje), dom = vsMais_(seg, 6), semana = vsBr_(seg);
    var n = sh.getLastRow(), vals = n >= 2 ? sh.getRange(2, 1, n - 1, CAB_VS.length).getValues() : [];
    var chaves = {}, temSemana = false, novos = [], transf = 0, add = 0, atual = 0;
    vals.forEach(function (l) { chaves[String(l[13])] = 1; if (String(l[0]) === semana) temSemana = true; });
    // 1) virada da semana: PENDENTE de semanas anteriores -> TRANSFERIDO, e entra na semana nova como ATRASADO
    if (!temSemana) {
      vals.forEach(function (l, i) {
        if (String(l[8]) !== 'PENDENTE' || String(l[0]) === semana) return;
        sh.getRange(i + 2, 9).setValue('TRANSFERIDO'); transf++;
        novos.push([semana, String(l[1]), String(l[2]), String(l[3]), String(l[4]), String(l[5]), String(l[6]), 'S', 'PENDENTE', '', '', '', '', String(l[13]), Utilities.getUuid(), "'" + String(l[15] || ''), String(l[16] || ''), '', '']);
      });
    }
    // 2) produtos da VALIDADE PRODUTOS cujo dia de retirar cai nesta semana (ou na anterior e nunca entraram: ATRASADO)
    var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone(), cadV = vsCadastroValidade_(ssV), skus = vsSkuPorNome_();
    var pc = ssV.getSheetByName('POR CHEGADA'), np = pc.getLastRow();
    var linhas = np >= 2 ? pc.getRange(2, 1, np - 1, 10).getValues() : [], vistos = {};
    var ini = vsMais_(seg, -7), valMin = new Date(2026, 10, 1); // Celio 02/11/2026: a lista começa nos vencimentos a partir de 01/11/2026
    linhas.forEach(function (l) {
      var prod = String(l[3] || '').trim(), val = l[9];
      if (!prod || !(val instanceof Date)) return;
      var pv = Utilities.formatDate(val, tzV, 'yyyy-MM-dd').split('-'); val = new Date(+pv[0], +pv[1] - 1, +pv[2]); // data como está na planilha
      var k = normDesc_(prod) + '|' + Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyyMMdd');
      if (chaves[k] || vistos[k] || val < valMin) return;
      var c = cadV[normDesc_(prod)] || { secao: '', local: '' }, regra = vsRegra_(c.secao, prod), ret = vsRetirar_(val, regra);
      if (vsFora_(c.secao) || ret < ini || ret > dom) return;
      vistos[k] = 1;
      var sk = skus[normDesc_(prod)] || { sku: '', un: '' };
      novos.push([semana, vsBr_(ret), vsBr_(val), prod.replace(/[¬]/g, '').trim(), c.secao, c.local, regra, ret < seg ? 'S' : '', 'PENDENTE', '', '', '', '', k, Utilities.getUuid(), "'" + sk.sku, sk.un, '', '']);
      add++;
    });
    if (novos.length) {
      var l0 = sh.getLastRow() + 1;
      sh.getRange(l0, 1, novos.length, 3).setNumberFormat('@'); sh.getRange(l0, 12, novos.length, 1).setNumberFormat('@');
      sh.getRange(l0, 1, novos.length, CAB_VS.length).setValues(novos);
    }
    // 3) pendentes da semana: atualiza seção/local/regra/dia de retirar pelo CADASTRO (ex.: produto que ganhou seção)
    n = sh.getLastRow(); vals = n >= 2 ? sh.getRange(2, 1, n - 1, CAB_VS.length).getValues() : [];
    vals.forEach(function (l, i) {
      if (String(l[0]) !== semana || String(l[8]) !== 'PENDENTE') return;
      if (vsFora_(l[4])) { sh.getRange(i + 2, 9).setValue('FORA'); return; }
      var c = cadV[normDesc_(l[3])]; if (!c) return;
      var regra = vsRegra_(c.secao, l[3]), val = vsDeBr_(l[2]), ret = val ? vsBr_(vsRetirar_(val, regra)) : String(l[1]);
      if (c.secao !== String(l[4]) || c.local !== String(l[5]) || regra !== String(l[6])) {
        sh.getRange(i + 2, 2).setValue(String(l[7]) === 'S' ? String(l[1]) : ret);
        sh.getRange(i + 2, 5, 1, 3).setValues([[c.secao, c.local, regra]]); atual++;
      }
    });
    return { ok: true, semana: semana, novos: add, atrasados: transf, atualizados: atual };
  } finally { lock.releaseLock(); }
}
function vsItem_(l, i) {
  return { linha: i + 2, semana: String(l[0]), retirar: String(l[1]), validade: String(l[2]), desc: String(l[3]), secao: String(l[4]), local: String(l[5]), regra: String(l[6]), atrasado: String(l[7]) === 'S', status: String(l[8]), qtd: l[9] === '' ? '' : Number(l[9]), quem: String(l[10]), em: String(l[11]), id: String(l[14]), sku: String(l[15] || ''), un: String(l[16] || '') };
}
function vencidos(quem) {
  if (apps(quem).indexOf('VENCIDOS') < 0) return { ok: false, erro: 'sem acesso' };
  var sh = criarAbaVS_(SpreadsheetApp.getActive()), semana = vsBr_(vsSegunda_(new Date()));
  var n = sh.getLastRow(), out = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, CAB_VS.length).getValues().forEach(function (l, i) { if (String(l[0]) === semana && ['TRANSFERIDO', 'FORA'].indexOf(String(l[8])) < 0 && !(String(l[8]) === 'PENDENTE' && vfVend_(l[15], l[3], l[2]))) /* v109: lote vendido sai */ out.push(vsItem_(l, i)); });
  return { ok: true, semana: semana, hoje: vsBr_(new Date()), itens: out };
}
function vsLinha_(sh, id) {
  var n = sh.getLastRow(); if (n < 2) return -1;
  var ids = sh.getRange(2, 15, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function vencidoNaoEncontrado(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('VENCIDOS') < 0) return { ok: false, erro: 'sem acesso' };
  var sh = criarAbaVS_(SpreadsheetApp.getActive()), r = vsLinha_(sh, p.id);
  if (r < 0) return { ok: false, erro: 'item não encontrado' };
  if (String(sh.getRange(r, 9).getValue()) !== 'PENDENTE') return { ok: false, erro: 'esse item já foi resolvido' };
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  sh.getRange(r, 9, 1, 5).setValues([['NAO ENCONTRADO', '', dono, agora, '']]);
  marcarOrigem_(sh, r, 1, CAB_VS.length);
  return { ok: true };
}
// chamado pelo salvarAvaria quando a baixa veio do botão "Retirei" da lista (reg.vsId)
function vsMarcarRetirado_(vsId, idAv, qtd, dono, validadeDig) {
  var sh = criarAbaVS_(SpreadsheetApp.getActive()), r = vsLinha_(sh, vsId);
  if (r < 0) return;
  var l = sh.getRange(r, 1, 1, CAB_VS.length).getValues()[0];
  if (String(l[8]) !== 'PENDENTE') return;
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  sh.getRange(r, 9, 1, 5).setValues([['RETIRADO', qtd, dono, agora, idAv]]);
  var vp = String(l[2]).trim(), vd = String(validadeDig || '').trim();
  sh.getRange(r, 18, 1, 2).setNumberFormat('@').setValues([[vd, vd && vd !== vp ? 'S' : '']]);
  marcarOrigem_(sh, r, 1, CAB_VS.length);
}
function vencidoDesfazer(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var sh = criarAbaVS_(SpreadsheetApp.getActive()), r = vsLinha_(sh, p.id);
  if (r < 0) return { ok: false, erro: 'item não encontrado' };
  var l = sh.getRange(r, 1, 1, CAB_VS.length).getValues()[0];
  if (String(l[8]) === 'RETIRADO' && String(l[12])) { var ex = excluirAvaria_(String(l[12])); if (!ex.ok) return { ok: false, erro: ex.motivo === 'já lançado' ? 'a baixa já foi lançada no Plenus; peça ao Celio' : 'não achei a avaria' }; }
  sh.getRange(r, 9, 1, 5).setValues([['PENDENTE', '', '', '', '']]);
  sh.getRange(r, 18, 1, 2).setValues([['', '']]);
  return { ok: true };
}
function instalarGatilhoVencidos_() {
  var gs = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (gs.indexOf('atualizarVencidos') < 0) ScriptApp.newTrigger('atualizarVencidos').timeBased().atHour(5).everyDays(1).create();
  if (gs.indexOf('atualizarCV') < 0) ScriptApp.newTrigger('atualizarCV').timeBased().atHour(5).everyDays(1).create(); // v82
  if (gs.indexOf('psvGerar') < 0) ScriptApp.newTrigger('psvGerar').timeBased().atHour(5).nearMinute(30).everyDays(1).create(); // v87
  if (gs.indexOf('psvAvisoDia') < 0) ScriptApp.newTrigger('psvAvisoDia').timeBased().atHour(19).nearMinute(45).everyDays(1).create(); // v87
  return { ok: true };
}

// ========== PRODUTO SEM VENDA (v87) ==========
// Aba PRODUTO SEM VENDA: montada todo dia às 5h30 por psvGerar (abaixo), com o histórico que a tarefa agendada "Vendas diárias" (20h)
// grava em BD/vendas diarias.xlsx. Três tipos de item: ALERTA (vende todo dia e parou), SEM VENDA 30 DIAS e SEM VENDA 60 DIAS.
// O app mostra a lista e grava as respostas (ESTÁ NA GÔNDOLA / NÃO ESTÁ → REPUS AGORA, NÃO TEM NO ESTOQUE, OUTRO).
// Item respondido fica cinza no app e some das listas por 7 dias. R2 = dias de histórico diário (ALERTA só funciona com 7 ou mais).
// Acesso (aba ACESSOS, coluna APP): "SEMVENDA" = todas as abas; "SEMVENDA ALERTA" ou "SEMVENDA 30 60" = só essas abas;
// acrescente TUDO (ex.: "SEMVENDA 30 60 TUDO") para a pessoa ver os itens de todos. Celio e Ney sempre veem tudo.
var ABA_PSV = 'PRODUTO SEM VENDA';
var CAB_PSV = ['DATA', 'SKU', 'PRODUTO', 'FAMILIA', 'LOCAL', 'QUEM', 'TIPO', 'DIAS', 'MEDIA_DIA', 'SALDO', 'RESPOSTA', 'MOTIVO', 'QUEM_RESPONDEU', 'HORA', 'ID', 'AJUSTE', 'VALIDADE']; // v99: VALIDADE = 'dd/mm/aaaa|sobra|SUG ou REB R$ x'
var PSV_RESP = { GONDOLA: 'ESTÁ NA GÔNDOLA', REPUS: 'NÃO ESTÁ — REPUS AGORA', SEMESTOQUE: 'NÃO ESTÁ — NÃO TEM NO ESTOQUE', OUTRO: 'NÃO ESTÁ — OUTRO' };
function criarAbaPSV_(ss) {
  var sh = ss.getSheetByName(ABA_PSV);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_PSV);
  sh.getRange(1, 1, 1, CAB_PSV.length).setValues([CAB_PSV]).setFontWeight('bold');
  sh.getRange('R1:R2').setValues([['DIAS HISTORICO'], [0]]); sh.getRange('R1').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:B').setNumberFormat('@'); sh.getRange('N:N').setNumberFormat('@');
  return sh;
}
function psvAba_(tipo) { var t = String(tipo || '').toUpperCase(); return /ALERTA/.test(t) ? 'ALERTA' : (/60/.test(t) ? '60' : (/30/.test(t) ? '30' : '')); }
function psvAcesso_(quem) { // -> { abas: {ALERTA:1,'30':1,'60':1}, tudo: bool } ou null (sem acesso)
  var sh = criarAbaAcessos_(SpreadsheetApp.getActive()), n = sh.getLastRow(), q = semAcento_(String(quem || '')).toLowerCase().trim(), r = null;
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) {
    if (semAcento_(String(l[0])).toLowerCase().trim() !== q) return;
    var a = semAcento_(String(l[1])).toUpperCase().trim(); if (!/^SEMVENDA/.test(a)) return;
    r = r || { abas: {}, tudo: false };
    var tk = a.replace(/^SEMVENDA/, '').split(/[\s:,;\/_-]+/).filter(String);
    var ab = tk.filter(function (t) { return t === 'ALERTA' || t === '30' || t === '60'; });
    (ab.length ? ab : ['ALERTA', '30', '60']).forEach(function (t) { r.abas[t] = 1; });
    if (tk.indexOf('TUDO') >= 0) r.tudo = true;
  });
  if (r && verTudo_(quem)) r.tudo = true;
  return r;
}
function psvT_(s) { var m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}))?/); return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime() : 0; }
function psvDados_(quem, ac) {
  var sh = criarAbaPSV_(SpreadsheetApp.getActive()), n = sh.getLastRow();
  var hist = Number(String(sh.getRange('R2').getValue()).replace(',', '.')) || 0;
  if (n < 2) return { itens: [], hist: hist };
  var v = sh.getRange(2, 1, n - 1, CAB_PSV.length).getDisplayValues(), lim = Date.now() - 7 * 86400000, resp = {}, out = [], sg = lsugMapa_(); // v99
  v.forEach(function (l) { if (!l[10]) return; var t = psvT_(l[13]); if (t >= lim) resp[String(l[1])] = 1; });
  v.forEach(function (l) {
    if (!l[14]) return;
    var ab = psvAba_(l[6]); if (!ab || !ac.abas[ab]) return;
    var dono = String(l[5] || '').trim() || 'Ney'; // sem repositora no cadastro: fica com o Ney
    if (!ac.tudo && !tfMeu_(dono, quem)) return;
    if (l[10]) { if (psvT_(l[13]) < lim) return; }   // respondido há mais de 7 dias: sai do app
    else if (resp[String(l[1])]) return;              // respondido nos últimos 7 dias (em outra linha): some das próximas listas
    out.push({ id: l[14], sku: l[1], desc: l[2], fam: l[3], local: l[4], quem: dono, aba: ab, dias: l[7], media: l[8], saldo: l[9], resp: l[10], motivo: l[11], por: l[12], hora: l[13], val: l[16] || '', sug: sg['#' + skuN_(l[1])] || sg[normDesc_(l[2])] || null });
  });
  return { itens: out, hist: hist };
}
function produtoSemVenda(quem) {
  var ac = apps(quem).indexOf('SEMVENDA') >= 0 ? psvAcesso_(quem) : null;
  if (!ac) return { ok: false, erro: 'sem acesso' };
  var d = psvDados_(quem, ac);
  var itens = d.itens.map(function (i) { var o = {}; Object.keys(i).forEach(function (k) { if (i[k] !== '' && i[k] != null) o[k] = i[k]; }); return o; }); // v131: não manda campo vazio ao celular
  return { ok: true, abas: Object.keys(ac.abas), tudo: ac.tudo, hist: d.hist, itens: itens, locais: ehCelio_(quem) ? lsugPendentes_() : [] };
}
function psvResumo_(quem) { // subtítulo do botão na tela inicial
  var ac = psvAcesso_(quem); if (!ac) return '';
  var d = psvDados_(quem, ac), pend = d.itens.filter(function (i) { return !i.resp; });
  var nA = pend.filter(function (i) { return i.aba === 'ALERTA'; }).length, nP = pend.length - nA, p = [];
  if (ac.abas.ALERTA) p.push(d.hist < 7 ? 'alertas em construção' : (nA ? nA + (nA === 1 ? ' alerta' : ' alertas') : 'nenhum alerta'));
  if (ac.abas['30'] || ac.abas['60']) p.push(nP ? nP + ' sem venda' : 'nenhum sem venda');
  return p.join(' · ');
}
function psvResponder(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('SEMVENDA') < 0 || !psvAcesso_(dono)) return { ok: false, erro: 'sem acesso' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = criarAbaPSV_(SpreadsheetApp.getActive()), n = sh.getLastRow(), r = -1;
    if (n >= 2) { var ids = sh.getRange(2, 15, n - 1, 1).getValues(); for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(p.id)) { r = i + 2; break; } }
    if (r < 0) return { ok: false, erro: 'item não encontrado (a lista foi atualizada; abra de novo)' };
    var atual = sh.getRange(r, 11, 1, 6).getValues()[0];
    if (String(p.resposta) === 'DESFAZER') {
      if (String(atual[5])) return { ok: false, erro: 'o Celio já tratou essa resposta' };
      if (!verTudo_(dono) && semAcento_(String(atual[2])).toLowerCase() !== semAcento_(dono).toLowerCase()) return { ok: false, erro: 'só quem respondeu pode desfazer' };
      sh.getRange(r, 11, 1, 4).setValues([['', '', '', '']]);
      return { ok: true };
    }
    var txt = PSV_RESP[String(p.resposta)]; if (!txt) return { ok: false, erro: 'resposta inválida' };
    var mot = String(p.motivo || '').replace(/\s+/g, ' ').trim().slice(0, 120);
    if (p.resposta === 'OUTRO' && !mot) return { ok: false, erro: 'escreva o que aconteceu' };
    if (String(atual[0])) return { ok: false, erro: 'esse produto já foi respondido por ' + atual[2] };
    var hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    sh.getRange(r, 14).setNumberFormat('@');
    sh.getRange(r, 11, 1, 4).setValues([[txt, mot, dono, hora]]);
    return { ok: true, resposta: txt, hora: hora };
  } finally { lock.releaseLock(); }
}

// ---- Geração da lista (gatilho diário às 5h, ou ?acao=psvGerar&chave=...) ----
// Fontes: cadastros ativos.xlsx (Código Mestre = família), saldo atual.txt, vendas diarias.xlsx (histórico diário, até 90 dias)
// e, enquanto o histórico diário não cobre 60 dias, os relatórios "periodo AAAA-MM-DD a AAAA-MM-DD.xlsx" da pasta BD/VENDAS DIARIAS.
// Regra de família: se qualquer SKU da família (mestre, filhos agrupados, PROMO) vendeu, nenhum entra; saldo mostrado = saldo da família.
// ALERTA (só com 7+ dias de histórico): média/dia da família (últimos 30 dias) >= 0,5, saldo > 0 e dias seguidos sem vender >= limite
//   (>= 3/dia → 1 dia; 1 a 3/dia → 3 dias; 0,5 a 1/dia → 6 dias). Fora: seções USO INTERNO e VACA GORDA, famílias só com "Não Vender".
var ID_PASTA_VD = '1s3baSeZjT-La9nBgWxrzkrsDrMmqwBpl'; // ARQUIVOS CLAUDE / BD / VENDAS DIARIAS
function psvDia_(v) { // dia (contagem UTC) a partir de serial do Excel, aaaa-mm-dd ou dd/mm/aaaa
  var s = String(v == null ? '' : v).trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000;
  if ((m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/))) return Date.UTC(+m[3], +m[2] - 1, +m[1]) / 86400000;
  if (/^\d+(\.\d+)?$/.test(s)) return Math.floor(Number(s)) - 25569;
  return NaN;
}
function psvNum_(v) { var s = String(v == null ? '' : v).trim(); if (/,/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); return Number(s) || 0; }
function psvSku_(v) { return String(Number(String(v == null ? '' : v).replace(/\D/g, '')) || ''); }
// v88: produto sem linha no CADASTRO da VALIDADE PRODUTOS → local/repositora pela seção do Plenus (aba LOCAL POR SECAO da VALIDADE PRODUTOS).
// Colunas: A SECAO PLENUS, B PALAVRA (início da descrição, várias separadas por /; linha com palavra vale antes da linha sem), C SECAO, D LOCAL, E REPOSITORA.
function psvMapaSecao_(ssV) {
  var sh = ssV.getSheetByName('LOCAL POR SECAO'), m = {}; if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (l) {
    var s = semAcento_(String(l[0] || '')).toUpperCase().trim(); if (!s || !String(l[4] || '').trim()) return;
    var pal = String(l[1] || '').split('/').map(function (x) { return semAcento_(x).toUpperCase().trim(); }).filter(String);
    (m[s] = m[s] || []).push({ pal: pal, secao: String(l[2] || '').trim(), local: String(l[3] || '').trim(), rep: String(l[4] || '').trim() });
  });
  Object.keys(m).forEach(function (s) { m[s].sort(function (a, b) { return (b.pal.length ? 1 : 0) - (a.pal.length ? 1 : 0); }); });
  return m;
}
function psvPorSecao_(mapa, sec, desc) {
  var l = mapa[semAcento_(String(sec || '')).toUpperCase().trim()]; if (!l) return null;
  var d = semAcento_(String(desc || '')).toUpperCase().trim();
  for (var i = 0; i < l.length; i++) if (!l[i].pal.length || l[i].pal.some(function (p) { return d.indexOf(p) === 0; })) return l[i];
  return null;
}
function psvGerar() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var tz = Session.getScriptTimeZone(), hoje = psvDia_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
    // 1) cadastro: famílias pelo Código Mestre
    var fam = {}, mestreDe = {};
    lerXlsx_(ID_CADASTRO_XLSX).slice(1).forEach(function (r) {
      if (String(r.B) !== 'T') return;
      var sku = psvSku_(r.E); if (!sku) return;
      var m = psvSku_(r.D) || sku;
      (fam[m] = fam[m] || []).push({ sku: sku, desc: String(r.F || '').replace(/[¬*]+\s*$/, '').trim(), nv: String(r.J) === 'T', sec: String(r.Q || '') });
      mestreDe[sku] = m;
    });
    // 2) saldo atual
    var sal = {};
    saldoTxt_().split(/\r?\n/).forEach(function (l, i) {
      if (!i) return; var c = l.split(';'); if (c.length < 7) return; var s = psvSku_(c[0]); if (s) sal[s] = psvNum_(c[6]);
    });
    // 3) histórico diário
    var porDia = {}, dias = {};
    var itV = DriveApp.getFolderById(ID_PASTA_BD).getFilesByName('vendas diarias.xlsx');
    if (itV.hasNext()) lerXlsx_(itV.next().getId()).slice(1).forEach(function (r) {
      var d = psvDia_(r.A), s = psvSku_(r.B), q = psvNum_(r.D);
      if (isNaN(d) || !s) return; dias[d] = 1;
      if (q > 0) { var m = mestreDe[s] || s; porDia[m] = porDia[m] || {}; porDia[m][d] = (porDia[m][d] || 0) + q; }
    });
    var lista = Object.keys(dias).map(Number).sort(function (a, b) { return a - b; }), nHist = lista.length, ultDia = nHist ? lista[nHist - 1] : hoje - 1;
    var rec = {}, ant = {}; // família vendeu nos últimos 30 dias / nos 30 dias anteriores
    Object.keys(porDia).forEach(function (m) { Object.keys(porDia[m]).forEach(function (d) { var age = ultDia - Number(d); if (age < 30) rec[m] = (rec[m] || 0) + porDia[m][d]; else if (age < 60) ant[m] = (ant[m] || 0) + porDia[m][d]; }); });
    var fontes = ['vendas diarias.xlsx (' + nHist + ' dias)'];
    if (!(nHist && lista[0] <= ultDia - 59)) { // histórico diário ainda não cobre 60 dias: completa com os relatórios por período
      var it = DriveApp.getFolderById(ID_PASTA_VD).getFiles();
      while (it.hasNext()) {
        var f = it.next(), mm = f.getName().match(/^periodo (\d{4}-\d{2}-\d{2}) a (\d{4}-\d{2}-\d{2})/i); if (!mm) continue;
        var age = ultDia - psvDia_(mm[2]), alvo = age < 30 ? rec : (age < 60 ? ant : null); if (!alvo) continue;
        fontes.push(f.getName());
        lerXlsx_(f.getId()).slice(1).forEach(function (r) { var s = psvSku_(r.C), q = psvNum_(r.J); if (s && q > 0) { var m = mestreDe[s] || s; alvo[m] = (alvo[m] || 0) + q; } });
      }
    }
    // 4) listas
    var ssVal = SpreadsheetApp.openById(ID_VALIDADE), cadV = cvCad_(ssVal), mapaS = psvMapaSecao_(ssVal), dataBr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'), novos = [];
    var jan = lista.slice(-30), nJ = jan.length;
    Object.keys(fam).forEach(function (m) {
      var mem = fam[m];
      if (mem.some(function (c) { return /USO INTERNO/i.test(c.sec) || /VACA GORDA/i.test(c.desc); })) return;
      var vend = mem.filter(function (c) { return !c.nv; }); if (!vend.length) return;
      var saldo = 0; mem.forEach(function (c) { if (sal[c.sku] > 0) saldo += sal[c.sku]; }); if (!(saldo > 0)) return;
      var tipo = '', ndias = '', media = 0;
      if (!rec[m]) { tipo = ant[m] ? 'SEM VENDA 30 DIAS' : 'SEM VENDA 60 DIAS'; ndias = ant[m] ? 30 : 60; media = ant[m] ? Math.round(ant[m] / 30 * 100) / 100 : 0; }
      else if (nHist >= 7) {
        var tot = 0, ult = -1; jan.forEach(function (d, i) { var q = (porDia[m] && porDia[m][d]) || 0; tot += q; if (q > 0) ult = i; });
        media = Math.round(tot / nJ * 100) / 100;
        var semV = nJ - 1 - ult, lim = media >= 3 ? 1 : (media >= 1 ? 3 : (media >= 0.5 ? 6 : 0));
        if (lim && semV >= lim) { tipo = 'ALERTA'; ndias = semV; }
      }
      if (!tipo) return;
      var rep = vend.slice().sort(function (a, b) { return (sal[b.sku] || 0) - (sal[a.sku] || 0) || (a.sku === m ? -1 : (b.sku === m ? 1 : 0)); })[0];
      var c = cadSku_(cadV, rep.sku, rep.desc);
      if (!c || !c.rep) mem.some(function (x) { var y = cadSku_(cadV, x.sku, x.desc); if (y && y.rep) { c = y; return true; } return false; });
      if (!c || !c.rep) c = psvPorSecao_(mapaS, rep.sec, rep.desc) || c; // v88: sem cadastro → pela seção do Plenus
      novos.push([dataBr, rep.sku, rep.desc, mem.length > 1 ? mem.length + ' SKUs (mestre ' + m + ')' : '', c ? c.local : '', c ? c.rep : '', tipo, ndias, media, Math.round(saldo * 1000) / 1000, '', '', '', '', '', '']);
    });
    // v110.4: crítica de lote (aba PRODUTOS da VALIDADE PRODUTOS, saldo distribuído da chegada mais recente para a mais antiga)
    try { var crit = psvCriticaLotes_(ssVal, novos); if (crit.sug.length) acharCriar_(crit.sug, novos); } catch (e) {} // v119.7: antes de sugerir rebaixa, a repositora precisa ACHAR o produto na loja
    novos.forEach(function (r) { if (r.length < CAB_PSV.length) r[CAB_PSV.length - 1] = r[CAB_PSV.length - 1] || ''; });
    // 5) regrava a aba: mantém as respostas (60 dias, e as NÃO TEM NO ESTOQUE ainda sem ajuste), troca os pendentes pela lista nova
    var sh = criarAbaPSV_(SpreadsheetApp.getActive()), n = sh.getLastRow(), manter = [], idDe = {}, resp7 = {}, l7 = Date.now() - 7 * 86400000, l60 = Date.now() - 60 * 86400000;
    if (n >= 2) sh.getRange(2, 1, n - 1, CAB_PSV.length).getDisplayValues().forEach(function (l) {
      if (String(l[10])) { var t = psvT_(l[13]); if (t >= l60 || (/NÃO TEM NO ESTOQUE/.test(l[10]) && !l[15])) manter.push(l); if (t >= l7) resp7[String(l[1])] = 1; }
      else if (l[14]) idDe[String(l[1]) + '|' + String(l[6])] = String(l[14]);
    });
    var ordT = { 'ALERTA': 0, 'SEM VENDA 30 DIAS': 1, 'SEM VENDA 60 DIAS': 2 };
    novos = novos.filter(function (r) { return !resp7[r[1]]; });
    novos.forEach(function (r) { r[14] = idDe[r[1] + '|' + r[6]] || Utilities.getUuid(); });
    novos.sort(function (a, b) { return ordT[a[6]] - ordT[b[6]] || String(a[5]).localeCompare(String(b[5])) || String(a[4]).localeCompare(String(b[4])) || String(a[2]).localeCompare(String(b[2])); });
    var tudo = manter.concat(novos);
    if (sh.getMaxRows() < tudo.length + 1) sh.insertRowsAfter(sh.getMaxRows(), tudo.length + 1 - sh.getMaxRows());
    if (n >= 2) sh.getRange(2, 1, n - 1, CAB_PSV.length).clearContent();
    if (tudo.length) {
      sh.getRange(2, 1, tudo.length, 2).setNumberFormat('@'); sh.getRange(2, 14, tudo.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, tudo.length, CAB_PSV.length).setValues(tudo);
    }
    if (String(sh.getRange(1, 17).getValue()) !== 'VALIDADE') sh.getRange(1, 17).setValue('VALIDADE').setFontWeight('bold'); // v99
    sh.getRange('R1:R4').setValues([['DIAS HISTORICO'], [nHist], ['ATUALIZADO ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')], [fontes.join(' + ')]]);
    var cnt = { ALERTA: 0, 'SEM VENDA 30 DIAS': 0, 'SEM VENDA 60 DIAS': 0 }; novos.forEach(function (r) { cnt[r[6]]++; });
    return { ok: true, hist: nHist, alerta: cnt.ALERTA, d30: cnt['SEM VENDA 30 DIAS'], d60: cnt['SEM VENDA 60 DIAS'], respostasMantidas: manter.length, fontes: fontes };
  } finally { lock.releaseLock(); }
}
// ---------- v110.4: CRÍTICA DE LOTE (pedido do Celio em 10/11/2026) ----------
// O saldo do sistema (soma da família) é distribuído pelos lotes da aba PRODUTOS da VALIDADE PRODUTOS (ordem das linhas = ordem de chegada),
// da chegada mais recente para a mais antiga (cada lote recebe min(CHEGOU, saldo que resta)). Lote sem validade lançada fica com a parte
// dele, mas nunca gera sugestão. Cada validade com saldo que vence em até 60 dias e não vende até lá vira sugestão com a quantidade DESSA validade.
function psvChave_(s) { return normDesc_(s).replace(/[^A-Z0-9]+$/, ''); }
function psvDataL_(v, tz) {
  if (v instanceof Date && !isNaN(v)) { var p = Utilities.formatDate(v, tz, 'yyyy-MM-dd').split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  var m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (!m) return null; var y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1]);
}
function psvLotes_(ssVal) {
  var sh = ssVal.getSheetByName('PRODUTOS'), tz = ssVal.getSpreadsheetTimeZone(), m = {}; if (!sh) return m;
  var n = sh.getLastRow(); if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 7).getValues().forEach(function (l, i) {
    var k = psvChave_(l[3]); if (!k) return;
    var q = Number(String(l[5]).replace(',', '.')) || 0; if (!(q > 0)) return;
    (m[k] = m[k] || []).push({ lin: i + 2, val: psvDataL_(l[1], tz), ch: psvDataL_(l[6], tz), q: q, nf: String(l[0] || '').trim() });
  });
  return m;
}
function psvN_(x) { x = Math.round(Number(x) * 1000) / 1000; return String(x).replace('.', ','); }
// novos = linhas da PRODUTO SEM VENDA (r[1] SKU, r[2] produto, r[7] dias, r[8] média/dia, r[9] saldo); grava r[16] e devolve as sugestões
function psvCriticaLotes_(ssVal, novos, lotes) {
  lotes = lotes || psvLotes_(ssVal);
  var hojeD = vsDia_(new Date()), idxR = null, sug = [], det = [];
  novos.forEach(function (r) {
    r[16] = '';
    var ls = lotes[psvChave_(r[2])]; if (!ls || !ls.length) return;
    var resto = Number(r[9]) || 0, media = Number(r[8]) || 0, com = [];
    for (var i = ls.length - 1; i >= 0 && resto > 0.0005; i--) { var q = Math.min(ls[i].q, resto); resto -= q; com.push({ l: ls[i], q: q }); }
    resto = Math.round(resto * 1000) / 1000;
    var grp = {}, ordem = [];
    com.forEach(function (c) { var k = c.l.val ? String(c.l.val.getTime()) : 'SEM'; if (!grp[k]) { grp[k] = { val: c.l.val, q: 0, lins: [] }; ordem.push(k); } grp[k].q += c.q; grp[k].lins.push(c.l.lin); });
    var qtx = ordem.map(function (k) { var g = grp[k]; return (g.val ? Utilities.formatDate(g.val, Session.getScriptTimeZone(), 'dd/MM/yy') : 'SEM VALIDADE') + ':' + psvN_(g.q); });
    if (resto > 0) qtx.push('SEM LOTE:' + psvN_(resto));
    r[16] = qtx.join(' | ');
    ordem.forEach(function (k) {
      if (String(r[6]) === 'ALERTA') return; // v122: ALERTA (vende todo dia e parou) não gera sugestão de rebaixa (Celio 15/11)
      var g = grp[k]; if (!g.val) return;
      var dd = Math.round((g.val - hojeD) / 86400000); if (dd < 0 || dd > 60) return;
      var sobra = Math.ceil(g.q - media * dd); if (!(sobra > 0)) return;
      var vb = vsBr_(g.val), minLin = Math.min.apply(null, g.lins);
      var inv = com.filter(function (c) { return c.l.val && c.l.lin > minLin && c.l.val < g.val; })[0];
      var obs = obsQtd_(g.q) + ' DESTA VALIDADE (critica de lote: saldo sistema ' + psvN_(r[9]) + ') (sem venda ' + r[7] + ' dias) - deve sobrar ' + sobra;
      if (resto > 0) obs += ' - CONFERIR NA GONDOLA - saldo maior que os lotes registrados';
      if (inv) obs += ' - LOTE INVERTIDO - lote de ' + (inv.l.ch ? Utilities.formatDate(inv.l.ch, Session.getScriptTimeZone(), 'dd/MM') : 'NF ' + inv.l.nf) + ' vence antes, fazer troca de frente';
      idxR = idxR || rebIndice_();
      var st = rebSituacao_(idxR, r[1], r[2], vb), o = { sku: r[1], desc: r[2], validade: vb, obs: obs, sit: st ? st.tipo : '' };
      det.push(o); if (!st) sug.push(o);
    });
  });
  return { sug: sug, det: det };
}
// Teste pelo editor (não grava nada): compara a regra antiga (última chegada da POR CHEGADA) com a crítica de lote
function psvSimularCritica() {
  var ssVal = SpreadsheetApp.openById(ID_VALIDADE), sh = SpreadsheetApp.getActive().getSheetByName(ABA_PSV), n = sh.getLastRow();
  var novos = n >= 2 ? sh.getRange(2, 1, n - 1, CAB_PSV.length).getValues().filter(function (l) { return !String(l[10]); }).map(function (l) { return l.slice(); }) : [];
  // antiga
  var pcS = ssVal.getSheetByName('POR CHEGADA'), npc = pcS.getLastRow(), ultC = {}, tzV2 = ssVal.getSpreadsheetTimeZone(), hojeD = vsDia_(new Date()), idxR = rebIndice_(), antes = [];
  if (npc >= 2) pcS.getRange(2, 1, npc - 1, 10).getValues().forEach(function (l) { var k = normDesc_(l[3]), ch = l[0], vv = l[9]; if (!k || !(ch instanceof Date) || !(vv instanceof Date)) return; var o = ultC[k]; if (!o || ch > o.ch || (+ch === +o.ch && vv < o.val)) ultC[k] = { ch: ch, val: vv }; });
  novos.forEach(function (r) { var o = ultC[normDesc_(r[2])]; if (!o) return; var pv = Utilities.formatDate(o.val, tzV2, 'yyyy-MM-dd').split('-'), val = new Date(+pv[0], +pv[1] - 1, +pv[2]), dd = Math.round((val - hojeD) / 86400000); if (dd < 0 || dd > 60) return; var sobra = Math.ceil(Number(r[9]) - (Number(r[8]) || 0) * dd); if (!(sobra > 0)) return; var st = rebSituacao_(idxR, r[1], r[2], vsBr_(val)); antes.push({ desc: r[2], validade: vsBr_(val), qtd: r[9], sobra: sobra, sit: st ? st.tipo : '' }); });
  var depois = psvCriticaLotes_(ssVal, novos);
  var ex = /AMENDOIM TIPO JAPONES DORI 500G|SUCO TANG LARANJA 18G|SERENATA 90G|GELATINA LUAL CEREJA|WAFER STICKS PALU/;
  var res = { linhas: novos.length, antes: { total: antes.length, novas: antes.filter(function (o) { return !o.sit; }).length }, depois: { total: depois.det.length, novas: depois.sug.length },
    exemplosAntes: antes.filter(function (o) { return ex.test(normDesc_(o.desc)); }), exemplosDepois: depois.det.filter(function (o) { return ex.test(normDesc_(o.desc)); }),
    lotesExemplos: novos.filter(function (r) { return ex.test(normDesc_(r[2])); }).map(function (r) { return r[2] + ' saldo ' + r[9] + ' media ' + r[8] + ' -> ' + r[16]; }) };
  Logger.log(JSON.stringify(res, null, 1));
  return res;
}
// Push diário para o Celio (~19h45): alertas pendentes e respostas NÃO ESTÁ do dia
function psvAvisoDia() {
  try {
    var sh = criarAbaPSV_(SpreadsheetApp.getActive()), n = sh.getLastRow(), hist = Number(sh.getRange('R2').getValue()) || 0;
    var hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'), nA = 0, nR = 0, nN = 0, nSE = 0;
    if (n >= 2) sh.getRange(2, 1, n - 1, CAB_PSV.length).getDisplayValues().forEach(function (l) {
      if (!l[10]) { if (psvAba_(l[6]) === 'ALERTA') nA++; return; }
      if (String(l[13]).indexOf(hoje) !== 0) return;
      nR++; if (/^NÃO ESTÁ/.test(l[10])) nN++; if (/NÃO TEM NO ESTOQUE/.test(l[10])) nSE++;
    });
    var tx = (hist < 7 ? 'alertas em construção (' + hist + ' de 7 dias de histórico)' : nA + (nA === 1 ? ' alerta pendente' : ' alertas pendentes')) + ' · hoje: ' + nR + (nR === 1 ? ' resposta' : ' respostas') + ', ' + nN + ' NÃO ESTÁ' + (nSE ? ' (' + nSE + ' sem estoque → ajuste de saldo)' : '');
    return enviarPush_(['Celio'], 'Produto sem venda', tx, { abrir: 'semvenda', tag: 'psv-dia' });
  } catch (e) { return { ok: false, erro: String(e) }; }
}

// ========== VALIDADES — TROCA DE FRENTE / PVPS (v81) ==========
// Quando a conferência grava a validade de um produto na VALIDADE PRODUTOS e já existe o mesmo produto com
// validade mais longa, vira tarefa "Troca de frente" para a repositora da seção (CADASTRO da VALIDADE PRODUTOS) e para o Ney.
// Sem seção no CADASTRO: vai para o Ney, no bloco "Sem seção". Celio e Ney veem tudo.
var ABA_TF = 'TROCA FRENTE';
var CAB_TF = ['DATA', 'NF', 'FORNECEDOR', 'PRODUTO', 'SECAO', 'LOCAL', 'REPOSITORA', 'VALIDADE NOVA', 'VALIDADE EXPOSTA', 'STATUS', 'QUEM', 'FEITO EM', 'CHAVE', 'ID'];
function criarAbaTF_(ss) {
  var sh = ss.getSheetByName(ABA_TF);
  if (sh) return sh;
  sh = ss.insertSheet(ABA_TF);
  sh.getRange(1, 1, 1, CAB_TF.length).setValues([CAB_TF]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:B').setNumberFormat('@'); sh.getRange('H:I').setNumberFormat('@'); sh.getRange('L:L').setNumberFormat('@');
  return sh;
}
// chamada dentro do gravarValidades_ (antes de gravar as linhas novas). dados = aba PRODUTOS sem esta NF; linhas = validades novas.
function tfGerar_(ssV, tzV, dados, linhas, nNF, forn) {
  try {
    if (!linhas || !linhas.length) return 0;
    var cad = {}, shC = ssV.getSheetByName('CADASTRO'), nc = shC.getLastRow();
    if (nc >= 2) shC.getRange(2, 1, nc - 1, 4).getValues().forEach(function (l) { var k = normDesc_(l[0]); if (k) cad[k] = { secao: String(l[1] || '').trim(), local: String(l[2] || '').trim(), rep: String(l[3] || '').trim() }; });
    var dStr = function (d) { return d instanceof Date ? Utilities.formatDate(d, tzV, 'yyyyMMdd') : ''; };
    var dBr = function (d) { return Utilities.formatDate(d, tzV, 'dd/MM/yyyy'); };
    var maior = {}; // produto -> maior validade já existente
    dados.forEach(function (l) { var k = normDesc_(l[3]); if (!k || !(l[1] instanceof Date) || !(l[6] instanceof Date)) return; if (!maior[k] || dStr(l[1]) > dStr(maior[k])) maior[k] = l[1]; });
    var nova = {}, nomes = {}; // produto -> menor validade que chegou agora
    linhas.forEach(function (l) { var k = normDesc_(l[3]); if (!k || !(l[1] instanceof Date)) return; nomes[k] = String(l[3]).replace(/\*+\s*$/, '').trim(); if (!nova[k] || dStr(l[1]) < dStr(nova[k])) nova[k] = l[1]; });
    var sh = criarAbaTF_(SpreadsheetApp.getActive()), n = sh.getLastRow(), chaves = {};
    if (n >= 2) sh.getRange(2, 13, n - 1, 1).getValues().forEach(function (l) { chaves[String(l[0])] = 1; });
    var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), out = [];
    Object.keys(nova).forEach(function (k) {
      var vN = nova[k], vE = maior[k]; if (!vE || dStr(vE) <= dStr(vN)) return;
      var ch = nNF + '|' + k + '|' + dStr(vN); if (chaves[ch]) return;
      var c = cad[k] || { secao: '', local: '', rep: '' };
      out.push([agora, "'" + nNF, forn, nomes[k], c.secao, c.local, c.secao ? (c.rep || 'Ney') : 'Ney', dBr(vN), dBr(vE), 'PENDENTE', '', '', ch, Utilities.getUuid()]);
    });
    if (out.length) { var l0 = sh.getLastRow() + 1; sh.getRange(l0, 1, out.length, 3).setNumberFormat('@'); sh.getRange(l0, 8, out.length, 2).setNumberFormat('@'); sh.getRange(l0, 1, out.length, CAB_TF.length).setValues(out); avTfNova_(out); }
    return out.length;
  } catch (e) { return 0; }
}
function tfMeu_(rep, quem) { var a = semAcento_(String(rep || '')).toLowerCase().trim(), b = semAcento_(String(quem || '')).toLowerCase().trim(); if (!b) return false; var ps = a.split(/\s*\/\s*/); if (a === b || ps.indexOf(b) >= 0) return true; return grupos_(quem).some(function (g) { return a === g || ps.indexOf(g) >= 0; }); } // v119.8: também as tarefas do grupo da pessoa
var GRUPOS_MEMO = null, GRUPOS_NOME = {};
function membrosGrupo_(nome) { // v119.9: quem atende o grupo (ex.: Meninas do caixa -> Marttiele), para os avisos no celular
  grupos_(''); var k = semAcento_(String(nome || '')).toLowerCase().trim(), out = []; if (!k) return out;
  Object.keys(GRUPOS_MEMO).forEach(function (q) { if (GRUPOS_MEMO[q].indexOf(k) >= 0) out.push(GRUPOS_NOME[q] || q); }); return out;
}
function grupos_(quem) { // v119.8 (Celio 14/11/2026): ACESSOS coluna C = grupo que a pessoa atende (ex.: Marttiele -> Meninas do caixa)
  if (!GRUPOS_MEMO) { GRUPOS_MEMO = {}; try { var sh = SpreadsheetApp.getActive().getSheetByName(ABA_ACESSOS), n = sh.getLastRow(); if (n >= 2 && sh.getLastColumn() >= 3) sh.getRange(2, 1, n - 1, 3).getValues().forEach(function (l) { var q = semAcento_(String(l[0] || '')).toLowerCase().trim(), g = semAcento_(String(l[2] || '')).toLowerCase().trim(); if (q && g) { var x = GRUPOS_MEMO[q] = GRUPOS_MEMO[q] || []; if (x.indexOf(g) < 0) x.push(g); GRUPOS_NOME[q] = String(l[0]).trim(); } }); } catch (e) {} }
  return GRUPOS_MEMO[semAcento_(String(quem || '')).toLowerCase().trim()] || [];
}
function validades(quem) {
  if (apps(quem).indexOf('VALIDADES') < 0) return { ok: false, erro: 'sem acesso' };
  var tudo = verTudo_(quem), sh = criarAbaTF_(SpreadsheetApp.getActive()), n = sh.getLastRow(), pend = [], feitos = [];
  var lim = new Date(); lim.setDate(lim.getDate() - 3);
  if (n >= 2) sh.getRange(2, 1, n - 1, CAB_TF.length).getValues().forEach(function (l) {
    if (!tudo && !tfMeu_(l[6], quem)) return;
    var it = { id: String(l[13]), data: String(l[0]), nf: String(l[1]), forn: String(l[2]), desc: String(l[3]), secao: String(l[4]), local: String(l[5]), rep: String(l[6]), nova: String(l[7]), exposta: String(l[8]), status: String(l[9]), quem: String(l[10]), em: String(l[11]) };
    if (it.status === 'PENDENTE') pend.push(it);
    else if (it.status === 'FEITO') { var p = it.em.split(/[\/ ]/); var d = p.length >= 3 ? new Date(+p[2], +p[1] - 1, +p[0]) : null; if (!d || d >= lim) feitos.push(it); }
  });
  return { ok: true, verTudo: tudo, trocas: pend, feitos: feitos.reverse(), cv: cvLista_(quem), locais: ehCelio_(quem) ? lsugPendentes_() : [], corr: ehCelio_(quem) ? cvCorrLista_() : [], avs: cvAvMeus_(quem) }; // v104
}
function tfLinha_(sh, id) {
  var n = sh.getLastRow(); if (n < 2) return -1;
  var ids = sh.getRange(2, 14, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function trocaFeita(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('VALIDADES') < 0) return { ok: false, erro: 'sem acesso' };
  var sh = criarAbaTF_(SpreadsheetApp.getActive()), r = tfLinha_(sh, p.id);
  if (r < 0) return { ok: false, erro: 'troca não encontrada' };
  var desfazer = String(p.desfazer || '') === '1';
  if (desfazer) { sh.getRange(r, 10, 1, 3).setValues([['PENDENTE', '', '']]); return { ok: true }; }
  if (String(sh.getRange(r, 10).getValue()) !== 'PENDENTE') return { ok: false, erro: 'essa troca já foi marcada' };
  sh.getRange(r, 10, 1, 3).setValues([['FEITO', dono, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')]]);
  marcarOrigem_(sh, r, 1, CAB_TF.length);
  return { ok: true };
}

// ========== v99 (05/11/2026) ==========
// 1) CADASTRO de seções (VALIDADE PRODUTOS) com SKU na coluna E: procura pelo SKU primeiro, depois pela descrição.
// 2) "Não fica aqui": a repositora sugere local e repositora; o Celio aprova e o CADASTRO é corrigido (todas as linhas do produto).
// 3) Validade curta: produto que chega na conferência com validade antes da janela da lista entra na lista da repositora, no topo.
// 4) Rebaixas: sugestão do app vai para o bloco A DECIDIR da planilha REBAIXAS (aba do mês); produto que já está na rebaixa
//    não é sugerido de novo; ACABOU/VENCEU no app preenche COMO ACABOU / SOBRA / ACABOU EM.
// 5) Produto sem venda com validade próxima (última chegada) e saldo que não vende até lá: sugere rebaixa.
var ABA_LSUG = 'LOCAIS SUGERIDOS';
var CAB_LSUG = ['ID', 'DATA', 'QUEM', 'MODULO', 'SKU', 'PRODUTO', 'LOCAL ANTES', 'REP ANTES', 'SECAO NOVA', 'LOCAL NOVO', 'REP NOVO', 'OBS', 'STATUS', 'DECIDIDO POR', 'DECIDIDO EM'];
var MES_PT = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
function skuN_(v) { return String(v == null ? '' : v).replace(/^'/, '').replace(/\D/g, '').replace(/^0+(?=\d)/, ''); }
function cadSku_(cad, sku, desc) { var s = skuN_(sku); return (s && cad['#' + s]) || cad[normDesc_(desc)] || null; }
function ehCelio_(q) { return semAcento_(String(q || '')).toLowerCase().trim() === 'celio'; }
function agoraBr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); }

// ---- SKU na coluna E do CADASTRO (preenche o que estiver em branco, pela descrição do cadastro do Plenus) ----
function cadSkuPreencher() {
  var sh = SpreadsheetApp.openById(ID_VALIDADE).getSheetByName('CADASTRO'), n = sh.getLastRow(); if (n < 2) return { ok: true, preenchidos: 0 };
  if (String(sh.getRange(1, 5).getValue()) !== 'SKU') sh.getRange(1, 5).setValue('SKU').setFontWeight('bold');
  var skus = vsSkuPorNome_(), v = sh.getRange(2, 1, n - 1, 5).getValues(), col = [], mud = 0;
  v.forEach(function (l) { var s = skuN_(l[4]); if (!s && String(l[0]).trim()) { var x = skus[normDesc_(l[0])]; if (x && x.sku) { s = skuN_(x.sku); mud++; } } col.push([s]); });
  sh.getRange(2, 5, col.length, 1).setNumberFormat('@').setValues(col);
  return { ok: true, preenchidos: mud };
}
// grava seção/local/repositora do produto: corrige TODAS as linhas com o mesmo SKU ou a mesma descrição; sem linha, cria uma
function cadGravar_(sku, desc, secao, local, rep) {
  var sh = SpreadsheetApp.openById(ID_VALIDADE).getSheetByName('CADASTRO'), n = sh.getLastRow(), k = normDesc_(desc), s = skuN_(sku), achou = 0;
  if (String(sh.getRange(1, 5).getValue()) !== 'SKU') sh.getRange(1, 5).setValue('SKU').setFontWeight('bold');
  if (n >= 2) sh.getRange(2, 1, n - 1, 5).getValues().forEach(function (l, i) {
    if (!((s && skuN_(l[4]) === s) || (k && normDesc_(l[0]) === k))) return;
    sh.getRange(i + 2, 2, 1, 3).setValues([[secao || String(l[1] || ''), local, rep]]);
    if (s) sh.getRange(i + 2, 5).setNumberFormat('@').setValue(s);
    achou++;
  });
  if (!achou) {
    var colA = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues(), last = 1;
    for (var j = colA.length - 1; j >= 1; j--) if (String(colA[j][0]).trim()) { last = j + 1; break; }
    if (last + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    sh.getRange(last + 1, 1, 1, 4).setValues([[desc, secao, local, rep]]); sh.getRange(last + 1, 5).setNumberFormat('@').setValue(s);
  }
  try { CacheService.getScriptCache().remove('LOC_OPC'); } catch (e) {}
  return achou || 1;
}
// leva a correção para as listas que já estão no app (sem venda pendente, conferência de validade da semana, troca de frente)
function listasNovoLocal_(sku, desc, secao, local, rep) {
  var s = skuN_(sku), k = normDesc_(desc), ss = SpreadsheetApp.getActive();
  var psv = ss.getSheetByName(ABA_PSV);
  if (psv && psv.getLastRow() >= 2) psv.getRange(2, 1, psv.getLastRow() - 1, 11).getValues().forEach(function (l, i) {
    if (String(l[10])) return; if (!((s && skuN_(l[1]) === s) || normDesc_(l[2]) === k)) return;
    psv.getRange(i + 2, 5, 1, 2).setValues([[local, rep]]);
  });
  var cv = ss.getSheetByName(ABA_CVS), semana = vsBr_(vsSegunda_(new Date()));
  if (cv && cv.getLastRow() >= 2) cv.getRange(2, 1, cv.getLastRow() - 1, 14).getValues().forEach(function (l, i) {
    if (String(l[0]) !== semana || String(l[10]) !== 'PENDENTE' || String(l[2]) === '45') return;
    if (!((s && skuN_(l[5]) === s) || normDesc_(l[4]) === k)) return;
    cv.getRange(i + 2, 2).setValue(rep); cv.getRange(i + 2, 7, 1, 3).setValues([[secao || String(l[6]), local, rep]]);
    cv.getRange(i + 2, 12).setValue(rep + '|' + l[2] + '|' + normDesc_(l[4]) + '|' + l[3]); cv.getRange(i + 2, 14).setValue('');
  });
  var tf = ss.getSheetByName(ABA_TF);
  if (tf && tf.getLastRow() >= 2) tf.getRange(2, 1, tf.getLastRow() - 1, 10).getValues().forEach(function (l, i) {
    if (String(l[9]) !== 'PENDENTE' || normDesc_(l[3]) !== k) return;
    tf.getRange(i + 2, 5, 1, 3).setValues([[secao || String(l[4]), local, rep]]);
  });
}

// ---- "Não fica aqui" ----
function locaisOpcoes_() { // locais do CADASTRO (mais usados primeiro), com a seção e a repositora mais comuns de cada um
  var cache = CacheService.getScriptCache(), c = cache.get('LOC_OPC'); if (c) return JSON.parse(c);
  var sh = SpreadsheetApp.openById(ID_VALIDADE).getSheetByName('CADASTRO'), n = sh.getLastRow(), loc = {}, reps = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (l) {
    var lc = String(l[2] || '').trim(), sc = String(l[1] || '').trim(), rp = String(l[3] || '').trim();
    if (rp) reps[rp] = (reps[rp] || 0) + 1; if (!lc) return;
    var o = loc[lc] = loc[lc] || { n: 0, sec: {}, rep: {} }; o.n++;
    if (sc) o.sec[sc] = (o.sec[sc] || 0) + 1; if (rp) o.rep[rp] = (o.rep[rp] || 0) + 1;
  });
  function top(m) { return Object.keys(m).sort(function (a, b) { return m[b] - m[a]; })[0] || ''; }
  var out = {
    locais: Object.keys(loc).map(function (k) { return { local: k, n: loc[k].n, secao: top(loc[k].sec), rep: top(loc[k].rep) }; }).sort(function (a, b) { return b.n - a.n; }),
    reps: Object.keys(reps).filter(function (r) { return reps[r] >= 2 && !/\//.test(r); }).sort(function (a, b) { return reps[b] - reps[a]; })
  };
  if (out.reps.indexOf('Ney') < 0) out.reps.push('Ney');
  try { cache.put('LOC_OPC', JSON.stringify(out), 3600); } catch (e) {}
  return out;
}
function locaisOpcoes(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var o = locaisOpcoes_(); return { ok: true, locais: o.locais, reps: o.reps };
}
function lsugAba_() { return cvAba_(ABA_LSUG, CAB_LSUG); }
function lsugPendentes_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_LSUG), out = []; if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, CAB_LSUG.length).getValues().forEach(function (l) {
    if (String(l[12]) !== 'PENDENTE') return;
    out.push({ id: String(l[0]), data: String(l[1]), quem: String(l[2]), modulo: String(l[3]), sku: skuN_(l[4]), desc: String(l[5]), localAntes: String(l[6]), repAntes: String(l[7]), secao: String(l[8]), local: String(l[9]), rep: String(l[10]), obs: String(l[11]) });
  });
  return out;
}
function lsugMapa_() { var m = {}; lsugPendentes_().forEach(function (x) { var o = { local: x.local, rep: x.rep, quem: x.quem }; if (x.sku) m['#' + x.sku] = o; m[normDesc_(x.desc)] = o; }); return m; }
function localSugerir(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var ap = apps(dono); if (ap.indexOf('VALIDADES') < 0 && ap.indexOf('SEMVENDA') < 0) return { ok: false, erro: 'sem acesso' };
  var local = String(p.local || '').replace(/\s+/g, ' ').trim().slice(0, 80), rep = String(p.rep || '').trim().slice(0, 40), desc = String(p.desc || '').trim();
  if (!desc) return { ok: false, erro: 'produto não informado' };
  if (!local) return { ok: false, erro: 'escolha onde o produto fica' };
  if (!rep) return { ok: false, erro: 'escolha quem repõe' };
  var o = locaisOpcoes_().locais.filter(function (x) { return x.local === local; })[0], secao = String(p.secao || '').trim() || (o ? o.secao : '');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = lsugAba_(), n = sh.getLastRow(), s = skuN_(p.sku), k = normDesc_(desc), r = -1;
    if (n >= 2) sh.getRange(2, 1, n - 1, CAB_LSUG.length).getValues().forEach(function (l, i) { if (String(l[12]) === 'PENDENTE' && ((s && skuN_(l[4]) === s) || normDesc_(l[5]) === k)) r = i + 2; });
    var linha = [r > 0 ? String(sh.getRange(r, 1).getValue()) : Utilities.getUuid(), agoraBr_(), dono, String(p.modulo || ''), s, desc, String(p.localAntes || ''), String(p.repAntes || ''), secao, local, rep, String(p.obs || '').replace(/\s+/g, ' ').trim().slice(0, 120), 'PENDENTE', '', ''];
    sh.getRange(r > 0 ? r : n + 1, 1, 1, CAB_LSUG.length).setValues([linha]);
  } finally { lock.releaseLock(); }
  try { enviarPush_(['Celio'], 'Local a aprovar — ' + dono, (desc + ': ' + local + ' · ' + rep).slice(0, 110), { abrir: String(p.modulo) === 'CV' ? 'validades' : 'semvenda', tag: 'lsug' }); } catch (e) {}
  return { ok: true, sug: { local: local, rep: rep, quem: dono } };
}
function localDecidir(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (!ehCelio_(dono)) return { ok: false, erro: 'só o Celio aprova' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = lsugAba_(), n = sh.getLastRow(), r = -1;
    if (n >= 2) { var ids = sh.getRange(2, 1, n - 1, 1).getValues(); for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(p.id)) { r = i + 2; break; } }
    if (r < 0) return { ok: false, erro: 'sugestão não encontrada' };
    var l = sh.getRange(r, 1, 1, CAB_LSUG.length).getValues()[0];
    if (String(l[12]) !== 'PENDENTE') return { ok: true, pend: lsugPendentes_() };
    var aprova = String(p.aprovar || '') === '1';
    if (aprova) { cadGravar_(l[4], l[5], String(l[8]), String(l[9]), String(l[10])); listasNovoLocal_(l[4], l[5], String(l[8]), String(l[9]), String(l[10])); }
    sh.getRange(r, 13, 1, 3).setValues([[aprova ? 'APROVADO' : 'RECUSADO', dono, agoraBr_()]]);
    return { ok: true, pend: lsugPendentes_() };
  } finally { lock.releaseLock(); }
}

// ---- Validade curta (chamada logo depois da conferência gravar as validades) ----
function cvCurta_(linhas, tzV) {
  try {
    var hoje = vsDia_(new Date()), seg = vsSegunda_(hoje), lim = vsMais_(seg, 28), semana = vsBr_(seg);
    var ssV = SpreadsheetApp.openById(ID_VALIDADE), cad = cvCad_(ssV), skus = vsSkuPorNome_();
    var sh = cvAba_(ABA_CVS, CAB_CVS), n = sh.getLastRow(), chaves = {}, novos = [], avisar = {};
    if (n >= 2) sh.getRange(2, 1, n - 1, 12).getValues().forEach(function (l) { if (String(l[0]) === semana) chaves[String(l[11])] = 1; });
    linhas.forEach(function (l) {
      if (!(l[1] instanceof Date)) return;
      var pv = Utilities.formatDate(l[1], tzV, 'yyyy-MM-dd').split('-'), val = new Date(+pv[0], +pv[1] - 1, +pv[2]);
      if (val < hoje || val >= lim) return;
      var prod = String(l[3]).replace(/\*+\s*$/, '').trim(), k = normDesc_(prod), sk = skus[k] || { sku: '' };
      var c = cadSku_(cad, sk.sku, prod) || { secao: '', local: '', rep: '' }, sem = !c.secao, para = sem ? 'Ney' : (c.rep || 'Ney');
      var ch = para + '|CURTA|' + k + '|' + vsBr_(val); if (chaves[ch]) return; chaves[ch] = 1;
      novos.push([semana, para, 'CURTA', vsBr_(val), prod, sk.sku, c.secao, c.local, c.rep, '', 'PENDENTE', ch, Utilities.getUuid(), sem ? 'S' : '', '', '', '', '', '', '', '']);
      (avisar[para] = avisar[para] || []).push(prod + ' (' + vsBr_(val).slice(0, 5) + ')');
    });
    if (novos.length) {
      sh.getRange(sh.getLastRow() + 1, 1, novos.length, CAB_CVS.length).setValues(novos);
      avDisparar_('CV_CURTA', avisar, novos.map(function (x) { return x[4] + ' (' + String(x[3]).slice(0, 5) + ')'; }), { tit: 'Chegou com validade curta', tx: '{qtdp}: {lista}' }); // v103
    }
    return novos.length;
  } catch (e) { return 0; }
}

// ---- Rebaixas na planilha REBAIXAS ----
function rebCols_(sh) {
  var h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 16)).getDisplayValues()[0].map(function (x) { return semAcento_(String(x)).toUpperCase().trim(); });
  return { ORIGEM: h.indexOf('ORIGEM') + 1, COMO: h.indexOf('COMO ACABOU') + 1, SOBRA: h.indexOf('SOBRA') + 1, EM: h.indexOf('ACABOU EM') + 1 };
}
function rebAbasRecentes_(ss) { // aba do mês atual e a do mês anterior
  var d = new Date(), out = [];
  [0, -1].forEach(function (k) { var x = new Date(d.getFullYear(), d.getMonth() + k, 1), sh = ss.getSheetByName(MES_PT[x.getMonth()] + String(x.getFullYear()).slice(2)); if (sh) out.push(sh); });
  return out;
}
// índice da planilha: descrição -> [{val, por}] (por vazio = ainda A DECIDIR / sem preço)
function rebIndice_() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA_REBAIXAS), m = {};
  rebAbasRecentes_(ss).forEach(function (sh) {
    var n = sh.getLastRow(); if (n < 2) return;
    sh.getRange(2, 1, n - 1, 11).getDisplayValues().forEach(function (l) { var k = normDesc_(l[1]); if (!k) return; (m[k] = m[k] || []).push({ val: dataBR_(l[9]), por: String(l[4] || '').trim(), nao: /NAO FAZER/i.test(semAcento_(String(l[10] || ''))) }); });
  });
  return { ss: ss, m: m };
}
// situação do produto: {tipo:'REB', por} já rebaixado · {tipo:'SUG'} já está para decidir · null
function rebSituacao_(idx, sku, desc, validade) {
  var d0 = diaNum_(validade), r = rebaixaVigente_(skuN_(sku));
  if (r && (isNaN(diaNum_(r.validade)) || Math.abs(diaNum_(r.validade) - d0) <= 3)) return { tipo: 'REB', por: 'R$ ' + Number(r.por || 0).toFixed(2).replace('.', ',') };
  var lst = idx.m[normDesc_(desc)] || [], achou = null;
  lst.forEach(function (x) { if (Math.abs(diaNum_(x.val) - d0) <= 2) { if (x.nao) { if (!achou || achou.tipo === 'SUG') achou = { tipo: 'NAO' }; } else if (x.por && /\d/.test(x.por)) achou = { tipo: 'REB', por: x.por }; else if (!achou) achou = { tipo: 'SUG' }; } }); // v110.4: NAO FAZER no K não volta como sugestão
  return achou;
}
function precoSku_(sku) {
  var s = skuN_(sku), cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD); if (!s || !cad || cad.getLastRow() < 2) return 0;
  var v = cad.getRange(2, 2, cad.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === s) return Number(v[i][4]) || 0;
  return 0;
}
// lista = [{sku, desc, validade:'dd/mm/aaaa', obs}] -> linhas no bloco A DECIDIR da aba do mês, ORIGEM = APP (sem repetir)
function rebSugerirLote_(lista) {
  var res = { ok: 0, ja: 0 };
  try {
    if (!lista || !lista.length) return res;
    var ss = SpreadsheetApp.openById(ID_PLANILHA_REBAIXAS), d = new Date(), sh = ss.getSheetByName(MES_PT[d.getMonth()] + String(d.getFullYear()).slice(2)); if (!sh) return res;
    var n = sh.getLastRow(), v = sh.getRange(1, 1, n, 11).getDisplayValues(), iD = -1, c = rebCols_(sh), livres = [];
    for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim().toUpperCase() === 'A DECIDIR') { iD = i; break; }
    if (iD < 0) return res;
    for (i = iD + 1; i < v.length; i++) if (!String(v[i][1]).trim()) livres.push(i + 1);
    lista.forEach(function (o) {
      var k = normDesc_(o.desc), d0 = diaNum_(o.validade);
      if (v.some(function (l) { return normDesc_(l[1]) === k && Math.abs(diaNum_(dataBR_(l[9])) - d0) <= 2; })) { res.ja++; return; }
      var lin = livres.shift(); if (!lin) { sh.insertRowsAfter(sh.getMaxRows(), 1); lin = sh.getMaxRows(); }
      var preco = precoSku_(o.sku);
      sh.getRange(lin, 2).setValue(o.desc); if (preco) sh.getRange(lin, 3).setValue(preco);
      sh.getRange(lin, 10, 1, 2).setValues([[o.validade, o.obs]]);
      if (c.ORIGEM) sh.getRange(lin, c.ORIGEM).setValue('APP');
      var nova = []; for (var j = 0; j < 11; j++) nova.push(''); nova[1] = o.desc; nova[9] = o.validade; v.push(nova);
      res.ok++;
    });
  } catch (e) { res.erro = String(e); }
  return res;
}
// ACABOU / VENCEU no app -> COMO ACABOU / SOBRA / ACABOU EM na planilha (motivo vazio = limpa, quando a rebaixa é reaberta)
function rebMarcarFim_(desc, validade, motivo, sobra, quando) {
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILHA_REBAIXAS), k = normDesc_(desc), d0 = diaNum_(dataBR_(validade));
    var abas = ss.getSheets().filter(function (s) { return /^[A-Z]{3}\d{2}$/.test(s.getName()); });
    for (var a = abas.length - 1; a >= 0; a--) {
      var sh = abas[a], n = sh.getLastRow(); if (n < 2) continue;
      var c = rebCols_(sh); if (!c.COMO || !c.SOBRA || !c.EM) continue;
      var v = sh.getRange(1, 1, n, 10).getDisplayValues();
      for (var i = 1; i < v.length; i++) {
        if (normDesc_(v[i][1]) !== k || Math.abs(diaNum_(dataBR_(v[i][9])) - d0) > 2) continue;
        var como = motivo === 'ACABOU' ? 'VENDEU TUDO' : (motivo === 'VENCEU' ? 'NAO VENDEU TUDO' : '');
        sh.getRange(i + 1, c.COMO).setValue(como);
        sh.getRange(i + 1, c.SOBRA).setValue(motivo === 'ACABOU' ? 0 : (motivo ? String(sobra || '') : ''));
        sh.getRange(i + 1, c.EM).setValue(motivo ? quando : '');
        return true;
      }
    }
  } catch (e) {}
  return false;
}
function obsQtd_(n) { n = Math.round(Number(n) || 0); return (n < 10 ? '0' : '') + n + 'UN'; }

// ========== v101 (06/11/2026): produtos que vêm em KG na nota e são vendidos por UNIDADE com peso variável ==========
// Ficam na aba CONFERENCIA KG PECAS (SKU, PRODUTO, OBS). Para eles a conferência NÃO converte kg em unidades pelo peso
// da embalagem: a conferente informa os kg e as peças (KG + peças) e o Celio corrige a quantidade da NF antes de faturar.
var ABA_KGP = 'CONFERENCIA KG PECAS', KGP_ = null;
function kgPecas_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ABA_KGP), m = {};
  if (!sh) { sh = ss.insertSheet(ABA_KGP); sh.getRange(1, 1, 1, 3).setValues([['SKU', 'PRODUTO', 'OBS']]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange('A:A').setNumberFormat('@'); }
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (l) { var s = skuN_(l[0]); if (s) m[s] = 1; });
  return m;
}
// admin: refaz um item já carregado (unidade/quantidade da nota) e apaga a contagem dele para a conferente lançar de novo
function confItemRefazer_(p) {
  var ss = SpreadsheetApp.getActive(), it = ss.getSheetByName(ABA_CIT), la = ss.getSheetByName(ABA_CLA), n = it.getLastRow(), r = -1;
  var v = it.getRange(2, 1, n - 1, 15).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][14]) === String(p.idItem)) { r = i + 2; break; }
  if (r < 0) return { ok: false, erro: 'item não encontrado' };
  it.getRange(r, 6, 1, 2).setValues([[String(p.un), Number(String(p.qtd).replace(',', '.'))]]);
  it.getRange(r, 11).setValue('PENDENTE'); it.getRange(r, 13, 1, 2).setValues([['', '']]);
  var m = la.getLastRow(), apag = 0;
  if (m >= 2) { var lv = la.getRange(2, 2, m - 1, 1).getValues(); for (var j = lv.length - 1; j >= 0; j--) if (String(lv[j][0]) === String(p.idItem)) { la.deleteRow(j + 2); apag++; } }
  return { ok: true, linha: r, lancApagados: apag };
}

// ========== v100 (06/11/2026): VALIDADE DIFERENTE na conferência de validade ==========
// A repositora acha o produto com outra validade:
//  - até 15 dias de diferença: corrige o lançamento (aba PRODUTOS da VALIDADE PRODUTOS) e a conta usa a data certa;
//  - mais de 15 dias e a data já existe para o produto: não salva — esse lote aparece numa lista mais para frente;
//  - mais de 15 dias e a data não existe: entra como VALIDADE NOVA (lote novo) e o item da lista fecha como "não achou".
// Tudo fica registrado na aba VALIDADES CORRIGIDAS (o Celio vê no módulo, sem apontar culpado).
var ABA_VCOR = 'VALIDADES CORRIGIDAS';
var CAB_VCOR = ['DATA', 'QUEM', 'SKU', 'PRODUTO', 'VAL LISTA', 'VAL ACHADA', 'DIAS', 'TIPO', 'NF', 'FORNECEDOR', 'CONFERIDA POR', 'ID ITEM'];
function cvValCaso_(l, nova) {
  nova = String(nova || '').trim();
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(nova)) return { ok: false, erro: 'data inválida' };
  var dN = diaNum_(nova), dO = diaNum_(String(l[3])), hoje = diaNum_(vsBr_(vsDia_(new Date())));
  if (isNaN(dN)) return { ok: false, erro: 'data inválida' };
  var dif = Math.round(Math.abs(dN - dO));
  if (!dif) return { ok: false, erro: 'é a mesma validade da lista — use CONFERIDO' };
  if (dN < hoje) return { ok: false, erro: 'essa validade já passou — produto vencido vai para o módulo de avaria' };
  var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone(), sh = ssV.getSheetByName('PRODUTOS'), n = sh.getLastRow();
  var k = normDesc_(String(l[4]).replace(/\*+\s*$/, '')), v = n >= 2 ? sh.getRange(2, 1, n - 1, 8).getValues() : [], velhas = [], existe = null, ref = null, ult = null;
  v.forEach(function (x, i) {
    if (!(x[1] instanceof Date) || normDesc_(String(x[3]).replace(/\*+\s*$/, '')) !== k) return;
    var dd = diaNum_(Utilities.formatDate(x[1], tzV, 'dd/MM/yyyy'));
    if (Math.abs(dd - dO) < 1) { velhas.push(i + 2); if (!ref) ref = x; }
    if (Math.abs(dd - dN) < 1 && !existe) existe = { nf: String(x[0]), chegou: x[6] instanceof Date ? Utilities.formatDate(x[6], tzV, 'dd/MM/yyyy') : '', forn: String(x[7]) };
    if (!ult || (x[6] instanceof Date && (!(ult[6] instanceof Date) || x[6] >= ult[6]))) ult = x;
  });
  ref = ref || ult;
  var caso = dif <= 15 ? 'CORRIGIR' : (existe ? 'EXISTE' : 'NOVA');
  return { ok: true, caso: caso, dif: dif, nova: nova, lista: String(l[3]), velhas: velhas, existe: existe, ref: ref ? { nf: String(ref[0]).replace(/\D/g, ''), forn: String(ref[7] || ''), secao: String(ref[2] || '') } : { nf: '', forn: '', secao: '' } };
}
function cvConferente_(nf, forn) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CNF); if (!nf || !sh || sh.getLastRow() < 2) return '';
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues(), f = semAcento_(String(forn || '')).toUpperCase().slice(0, 6), achou = '';
    v.forEach(function (x) { if (String(x[2]).replace(/\D/g, '') === nf && (!f || semAcento_(String(x[1])).toUpperCase().indexOf(f) >= 0 || !achou)) achou = String(x[6] || ''); });
    return achou;
  } catch (e) { return ''; }
}
function cvCorrLog_(dono, l, vd, tipo, idItem) {
  try {
    var sh = cvAba_(ABA_VCOR, CAB_VCOR), n = sh.getLastRow();
    if (tipo === 'JA EXISTIA' && n >= 2) { var ja = sh.getRange(2, 1, n - 1, CAB_VCOR.length).getValues().some(function (x) { return String(x[11]) === String(idItem) && String(x[5]) === vd.nova && String(x[7]) === tipo; }); if (ja) return; }
    var nf = tipo === 'JA EXISTIA' && vd.existe ? String(vd.existe.nf).replace(/\D/g, '') : (tipo === 'VALIDADE NOVA' ? '' : vd.ref.nf), forn = tipo === 'JA EXISTIA' && vd.existe ? vd.existe.forn : (tipo === 'VALIDADE NOVA' ? '' : vd.ref.forn);
    sh.getRange(n + 1, 1, 1, CAB_VCOR.length).setValues([[agoraBr_(), dono, String(l[5]), String(l[4]), String(l[3]), vd.nova, String(vd.dif), tipo, nf, forn, tipo === 'VALIDADE NOVA' ? '' : cvConferente_(nf, forn), String(idItem)]]);
  } catch (e) {}
}
function cvValAplicar_(sh, r, l, vd, dono, total, idItem) {
  var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone(), pr = ssV.getSheetByName('PRODUTOS');
  var dNova = Utilities.parseDate(vd.nova, tzV, 'dd/MM/yyyy');
  if (vd.caso === 'CORRIGIR' && vd.velhas.length) vd.velhas.forEach(function (row) { pr.getRange(row, 2).setValue(dNova); });
  else { // VALIDADE NOVA (ou correção sem linha no levantamento): entra como lote encontrado na loja
    var colD = pr.getRange(1, 4, pr.getMaxRows(), 1).getValues(), last = 1;
    for (var k = colD.length - 1; k >= 1; k--) if (String(colD[k][0]).trim()) { last = k + 1; break; }
    var hojeV = Utilities.parseDate(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'), tzV, 'dd/MM/yyyy');
    pr.insertRowsAfter(last, 1);
    pr.getRange(last + 1, 1, 1, 8).setValues([['', dNova, vd.ref.secao, String(l[4]).replace(/\*+\s*$/, '').trim(), 'UN', total, hojeV, vd.ref.forn]]);
  }
  if (vd.caso === 'CORRIGIR') { sh.getRange(r, 4).setValue(vd.nova); sh.getRange(r, 12).setValue(String(l[1]) + '|' + l[2] + '|' + normDesc_(l[4]) + '|' + vd.nova); }
  cvCorrLog_(dono, l, vd, vd.caso === 'CORRIGIR' ? 'CORRIGIDA' : 'VALIDADE NOVA', idItem);
}
// o app pergunta antes de salvar: qual é o caso dessa data?
function cvValCheck(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('VALIDADES') < 0) return { ok: false, erro: 'sem acesso' };
  var sh = cvAba_(ABA_CVS, CAB_CVS), n = sh.getLastRow(), r = -1;
  if (n >= 2) { var ids = sh.getRange(2, 13, n - 1, 1).getValues(); for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(p.id)) { r = i + 2; break; } }
  if (r < 0) return { ok: false, erro: 'item não encontrado' };
  var l = sh.getRange(r, 1, 1, CAB_CVS.length).getValues()[0], vd = cvValCaso_(l, p.novaVal);
  if (!vd.ok) return vd;
  if (vd.caso === 'EXISTE') cvCorrLog_(dono, l, vd, 'JA EXISTIA', p.id);
  return { ok: true, caso: vd.caso, dif: vd.dif, nova: vd.nova, lista: vd.lista, existe: vd.existe };
}
function cvCorrLista_() { // últimos 7 dias, mais recente primeiro
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_VCOR), out = []; if (!sh || sh.getLastRow() < 2) return out;
  var lim = diaNum_(vsBr_(vsDia_(new Date()))) - 7;
  sh.getRange(2, 1, sh.getLastRow() - 1, 15).getValues().forEach(function (x) { // v104: + QTD / CRITICA / SOBRA
    if (diaNum_(String(x[0]).slice(0, 10)) < lim) return;
    out.push({ em: String(x[0]), quem: String(x[1]), desc: String(x[3]), lista: String(x[4]), achada: String(x[5]), dias: String(x[6]), tipo: String(x[7]), nf: String(x[8]), forn: String(x[9]), conf: String(x[10]), qtd: String(x[12] || ''), crit: String(x[13] || ''), sobra: String(x[14] || '') });
  });
  return out.reverse();
}

// ========== v103 (06/11/2026): AVISOS AUTOMÁTICOS (push) ==========
// Tudo configurável na aba AVISOS AUTOMATICOS da planilha APPS PAUVAN:
//  LIGADO  = SIM / NAO
//  PARA    = nomes separados por vírgula; REPOSITORAS = cada repositora recebe só o que é da seção dela
//            (pelo CADASTRO da VALIDADE PRODUTOS), e só quem tem o módulo liberado na aba ACESSOS
//  HORARIO = NA HORA (dispara quando acontece) ou a hora do dia (7 = 7h; o envio sai entre 7h e 8h)
//  DIAS    = avisos com horário: TODOS ou dias da semana (SEG,QUA,SEX); lembretes: de quantos em quantos dias
//  TITULO / TEXTO = mensagem; campos: {qtd} número, {qtdp} "3 produtos", {lista} nomes, {dias} dias parado
//  ULTIMO ENVIO = preenchido pelo app
// Para criar outro lembrete livre: copie a linha LEMBRETE, troque o código (LEMBRETE2, LEMBRETE3...) e ligue.
var ABA_AV = 'AVISOS AUTOMATICOS';
var CAB_AV = ['CODIGO', 'SITUACAO', 'LIGADO', 'PARA', 'HORARIO', 'DIAS', 'TITULO', 'TEXTO', 'ULTIMO ENVIO'];
var AV_PADRAO = [
  ['REB_NOVA', 'Rebaixa nova — quem leva as etiquetas', 'SIM', 'Nath', 'NA HORA', '', '🏷️ Rebaixa nova', '{qtdp} entraram na rebaixa. As etiquetas já estão impressas: pegue e leve para as repositoras.'],
  ['REB_NOVA_REP', 'Rebaixa nova — cada repositora (só a seção dela)', 'SIM', 'REPOSITORAS', 'NA HORA', '', '🏷️ Rebaixa nova na sua seção', '{qtdp}: {lista}'],
  ['REB_VENCE_HOJE', 'Rebaixa vence hoje', 'SIM', 'REPOSITORAS', '7', 'TODOS', '⏰ Rebaixa vence hoje', 'Vence hoje na rebaixa: {qtdp}. Confira na sua lista: {lista}'],
  ['REB_VENCE_AMANHA', 'Rebaixa vence amanhã', 'SIM', 'REPOSITORAS', '16', 'TODOS', 'Rebaixa vence amanhã', 'Vence amanhã na rebaixa: {qtdp}. {lista}'],
  ['REB_ENCERRADA', 'Rebaixa encerrada no app (ACABOU / VENCEU)', 'SIM', 'Celio', 'NA HORA', '', '', ''],
  ['CV_SEMANA', 'Lista de validades da semana', 'SIM', 'REPOSITORAS', '7', 'SEG', '📋 Validades da semana', 'Sua lista da semana tem {qtdp} para conferir.'],
  ['CV_CURTA', 'Chegou produto com validade curta', 'SIM', 'REPOSITORAS', 'NA HORA', '', 'Chegou com validade curta', '{qtdp}: {lista}'],
  ['POFF_PRONTO', 'Preço off-line pronto para ativar', 'SIM', 'Celio', 'NA HORA', '', '', ''],
  ['NF_CHEGOU', 'Nota fiscal chegou para conferir', 'SIM', 'Regiane', 'NA HORA', '', '📦 Nota para conferir', '{qtd} nota(s) para conferir: {lista}'],
  ['FOTOS_PEND', 'Fotos do site pendentes (lembrete)', 'SIM', 'Delivery', '9', '5', '📷 Fotos do site', 'Ainda tem {qtdp} esperando foto: {lista}'],
  ['TF_NOVA', 'Troca de frente nova', 'SIM', 'REPOSITORAS', 'NA HORA', '', '🔁 Troca de frente', '{qtdp} para trocar a frente: {lista}'],
  ['TF_LEMBRETE', 'Troca de frente parada (lembrete)', 'SIM', 'REPOSITORAS', '10', '5', '🔁 Troca de frente pendente', 'Faz {dias} dias sem baixa: ainda tem {qtdp} para trocar a frente.'],
  ['PSV_LEMBRETE', 'Produto sem venda parado (lembrete)', 'SIM', 'REPOSITORAS', '10', '3', 'Produto sem venda pendente', 'Faz {dias} dias sem resposta: ainda tem {qtdp} sem venda para conferir.'],
  ['LEMBRETE', 'Lembrete livre (exemplo — escreva e ligue)', 'NAO', 'Celio', '8', 'SEG', 'Lembrete', 'Escreva aqui a mensagem.']
];
var AV_ABRIR = { REB: 'rebaixas', CV: 'validades', TF: 'validades', POFF: 'precosoff', NF: 'conferencia', FOTOS: 'fotos', PSV: 'semvenda' };
var AV_APP = { REB: 'REBAIXAS', CV: 'VALIDADES', TF: 'VALIDADES', PSV: 'SEMVENDA', FOTOS: 'FOTOS', NF: 'CONFERENCIA', POFF: 'PRECOS_OFF' };
function avK_(s) { return semAcento_(String(s || '')).toLowerCase().trim(); }
function avAba_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ABA_AV);
  if (!sh) {
    sh = ss.insertSheet(ABA_AV);
    sh.getRange(1, 1, 1, CAB_AV.length).setValues([CAB_AV]).setFontWeight('bold');
    sh.getRange(2, 1, AV_PADRAO.length, 8).setNumberFormat('@').setValues(AV_PADRAO);
    sh.setFrozenRows(1); sh.setColumnWidth(2, 300); sh.setColumnWidth(8, 420);
  } else { // códigos novos que ainda não estão na aba entram no fim (não mexe no que o Celio mudou)
    var n = sh.getLastRow(), tem = {};
    if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (l) { tem[String(l[0]).trim().toUpperCase()] = 1; });
    var falta = AV_PADRAO.filter(function (l) { return !tem[l[0]]; });
    if (falta.length) sh.getRange(n + 1, 1, falta.length, 8).setNumberFormat('@').setValues(falta);
  }
  return sh;
}
function avTodos_() { // [{cod, lig, para, hora, dias, tit, tx, ult, row}]
  var c = CacheService.getScriptCache(), j = c.get('AV_CFG'); if (j) return JSON.parse(j);
  var sh = avAba_(), n = sh.getLastRow(), out = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, CAB_AV.length).getDisplayValues().forEach(function (l, i) {
    var cod = String(l[0]).trim().toUpperCase(); if (!cod) return;
    out.push({ cod: cod, lig: /^S/i.test(String(l[2]).trim()), para: String(l[3]), hora: String(l[4]).trim().toUpperCase(), dias: String(l[5]).trim().toUpperCase(), tit: String(l[6]), tx: String(l[7]), ult: String(l[8]), row: i + 2 });
  });
  c.put('AV_CFG', JSON.stringify(out), 300);
  return out;
}
function avCfg_(cod) { var a = avTodos_().filter(function (x) { return x.cod === cod; }); return a[0] || null; }
function avLig_(cod) { var c = avCfg_(cod); return !c || c.lig; } // sem linha na aba = ligado (comportamento antigo)
function avNomes_(cod, padrao) { var c = avCfg_(cod); if (!c || !String(c.para).trim()) return padrao; return String(c.para).split(/[,;]+/).map(function (s) { return s.trim(); }).filter(function (s) { return s && s.toUpperCase() !== 'REPOSITORAS'; }); }
function avMarcar_(cfg, txt) { try { if (cfg && cfg.row) { avAba_().getRange(cfg.row, 9).setValue(txt || agoraBr_()); CacheService.getScriptCache().remove('AV_CFG'); } } catch (e) {} }
function avAcessos_() { // pessoa (sem acento, minúscula) -> {APP:1}
  var sh = criarAbaAcessos_(SpreadsheetApp.getActive()), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { var q = avK_(l[0]); if (!q) return; var a = semAcento_(String(l[1]).trim().toUpperCase()); if (/^SEMVENDA/.test(a)) a = 'SEMVENDA'; (m[q] = m[q] || {})[a] = 1; });
  return m;
}
function avTexto_(modelo, itens, dias) {
  var q = itens.length, lista = itens.slice(0, 8).join(', ') + (q > 8 ? ' e mais ' + (q - 8) : '');
  return String(modelo || '').replace(/\{qtdp\}/g, q + (q === 1 ? ' produto' : ' produtos')).replace(/\{qtd\}/g, String(q)).replace(/\{lista\}/g, lista).replace(/\{dias\}/g, String(dias || '')).slice(0, 230);
}
function avGrupo_(cod) { return cod.split('_')[0]; }
// por = {pessoa: [itens]} (só vale para PARA = REPOSITORAS); todos = lista inteira (para os nomes fixos)
// op = {tit, tx (padrões quando a aba está em branco), lembrete: true}
function avDisparar_(cod, por, todos, op) {
  op = op || {};
  var cfg = avCfg_(cod); if (cfg && !cfg.lig) return 0;
  var para = cfg && String(cfg.para).trim() ? cfg.para : (op.para || 'REPOSITORAS');
  var g = avGrupo_(cod), app = AV_APP[g], ac = null, env = {}, alvos = [];
  String(para).split(/[,;]+/).map(function (s) { return s.trim(); }).filter(String).forEach(function (t) {
    if (t.toUpperCase() === 'REPOSITORAS') {
      ac = ac || avAcessos_();
      Object.keys(por || {}).forEach(function (p) {
        String(p).split(/\s*\/\s*/).reduce(function (acc, x) { return acc.concat([x], membrosGrupo_(x)); }, []).forEach(function (pp) { // "Tati / Gláucia" = as duas; v119.9: + quem atende o grupo
          var k = avK_(pp); if (!k || env[k] || !(por[p] || []).length) return;
          if (app && !(ac[k] && ac[k][app])) return; // sem o módulo: não recebe
          env[k] = 1; alvos.push({ nome: pp.trim(), itens: por[p] });
        });
      });
    } else { var k = avK_(t); if (!env[k] && (todos || []).length) { env[k] = 1; alvos.push({ nome: t, itens: todos }); } }
  });
  var tit = (cfg && cfg.tit) || op.tit || 'Aviso do Pauvan', modelo = (cfg && cfg.tx) || op.tx || '{qtdp}: {lista}', n = 0;
  var cada = cfg && !isNaN(Number(cfg.dias)) && Number(cfg.dias) > 0 ? Number(cfg.dias) : 0;
  alvos.forEach(function (a) {
    var dias = 0;
    if (op.lembrete) { dias = avParado_(cod, a.nome, a.itens.length, cada || 3); if (dias < 0) return; }
    try { var r = enviarPush_([a.nome], tit, avTexto_(modelo, a.itens, dias), { abrir: AV_ABRIR[g] || 'aviso', tag: 'av-' + cod.toLowerCase() }); if (r && r.aparelhos) n++; } catch (e) {}
  });
  if (cfg && alvos.length) avMarcar_(cfg, agoraBr_() + ' · ' + n + (n === 1 ? ' pessoa' : ' pessoas'));
  return n;
}
// lembrete: só avisa se a pessoa ficou "cada" dias sem baixar nenhum item (a quantidade não caiu)
// e se o último lembrete foi há "cada" dias ou mais. Devolve os dias parados, ou -1 (não avisa agora).
function avParado_(cod, nome, qtd, cada) {
  var pr = PropertiesService.getScriptProperties(), ch = 'AVL_' + cod + '_' + avK_(nome).replace(/\W+/g, '_');
  var hoje = diaNum_(vsBr_(vsDia_(new Date()))), st = JSON.parse(pr.getProperty(ch) || 'null');
  if (!st || qtd < st.c) st = { c: qtd, t: hoje, r: st ? st.r || 0 : 0 }; // primeira vez ou baixou algum: conta de novo a partir de hoje
  else st.c = qtd;
  var parado = hoje - st.t, ok = parado >= cada && hoje - (st.r || 0) >= cada;
  if (ok) st.r = hoje;
  pr.setProperty(ch, JSON.stringify(st));
  return ok ? parado : -1;
}
function avZerar_(cod, nomesComPendencia) { // quem não tem mais nada pendente: apaga o controle do lembrete
  var pr = PropertiesService.getScriptProperties(), pre = 'AVL_' + cod + '_', ok = {};
  nomesComPendencia.forEach(function (n) { ok[pre + avK_(n).replace(/\W+/g, '_')] = 1; });
  pr.getKeys().forEach(function (k) { if (k.indexOf(pre) === 0 && !ok[k]) pr.deleteProperty(k); });
}
function avPush_(m, pessoa, item) { String(pessoa || 'Ney').split(/\s*\/\s*/).forEach(function (p) { p = p.trim() || 'Ney'; (m[p] = m[p] || []).push(item); }); }

// ---- ganchos "na hora" ----
function avRebNova_(novas) { // novas = linhas gravadas na aba REBAIXAS por sincronizarRebaixas
  try {
    if (!novas || !novas.length) return;
    var mp = mapaReposicao_(), por = {}, todos = [];
    novas.forEach(function (l) {
      var nome = String(l[1]).replace(/\*+\s*$/, '').trim(); todos.push(nome);
      var r = cadSku_(mp, l[0], l[1]), rp = r && r.rep ? r.rep : quemDaDemanda_(l[2]); if (rp) avPush_(por, rp, nome); // v129: sem repositora no cadastro -> SETORES pela demanda original
    });
    avDisparar_('REB_NOVA', {}, todos, { para: 'Nath' });
    avDisparar_('REB_NOVA_REP', por, todos, {});
  } catch (e) {}
}
function avTfNova_(out) { // linhas novas da aba TROCA FRENTE
  try {
    if (!out || !out.length) return;
    var por = {}, todos = [];
    out.forEach(function (l) { todos.push(String(l[3])); avPush_(por, l[6], String(l[3])); });
    avDisparar_('TF_NOVA', por, todos, {});
  } catch (e) {}
}
function avFilaNf_(txt) { // NF nova: entra na fila e sai no próximo avisosHora (junta as notas da mesma varredura)
  try { var pr = PropertiesService.getScriptProperties(), f = JSON.parse(pr.getProperty('AV_NF_FILA') || '[]'); f.push(String(txt)); pr.setProperty('AV_NF_FILA', JSON.stringify(f.slice(-40))); } catch (e) {}
}

// ---- avisos com horário ----
function avRebVence_(dias) { // rebaixas VIGENTES que vencem hoje (0) ou amanhã (1), por repositora
  var sh = criarAbaRebaixas_(SpreadsheetApp.getActive()), n = sh.getLastRow(), por = {}, todos = []; if (n < 2) return { por: por, todos: todos };
  var alvo = diaNum_(vsBr_(vsDia_(new Date()))) + dias, mp = mapaReposicao_();
  sh.getRange(2, 1, n - 1, 8).getValues().forEach(function (l) {
    if (String(l[7]) !== 'VIGENTE' || diaNum_(dataBR_(l[5])) !== alvo) return;
    var nome = String(l[1]).replace(/\*+\s*$/, '').trim(), r = cadSku_(mp, l[0], l[1]);
    todos.push(nome); if (r && r.rep) avPush_(por, r.rep, nome);
  });
  return { por: por, todos: todos };
}
function avCvSemana_() { // lista de validades da semana pendente, por pessoa (coluna PARA)
  var sh = cvAba_(ABA_CVS, CAB_CVS), n = sh.getLastRow(), por = {}, todos = []; if (n < 2) return { por: por, todos: todos };
  var semana = vsBr_(vsSegunda_(vsDia_(new Date())));
  sh.getRange(2, 1, n - 1, 11).getDisplayValues().forEach(function (l) {
    if (String(l[0]) !== semana || String(l[10]) !== 'PENDENTE') return;
    todos.push(String(l[4])); avPush_(por, l[1], String(l[4]));
  });
  return { por: por, todos: todos };
}
function avTfPend_() {
  var sh = criarAbaTF_(SpreadsheetApp.getActive()), n = sh.getLastRow(), por = {}, todos = []; if (n < 2) return { por: por, todos: todos };
  sh.getRange(2, 1, n - 1, 10).getDisplayValues().forEach(function (l) { if (String(l[9]) !== 'PENDENTE') return; todos.push(String(l[3])); avPush_(por, l[6], String(l[3])); });
  return { por: por, todos: todos };
}
function avPsvPend_() { // produto sem venda sem resposta, por repositora (o mesmo que aparece no app de cada uma)
  var sh = criarAbaAcessos_(SpreadsheetApp.getActive()), n = sh.getLastRow(), nomes = {}, por = {}, todos = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { var q = String(l[0]).trim(); if (q && /^SEMVENDA/.test(semAcento_(String(l[1])).toUpperCase().trim()) && !verTudo_(q)) nomes[avK_(q)] = q; });
  Object.keys(nomes).forEach(function (k) {
    var nome = nomes[k], a = psvAcesso_(nome); if (!a) return;
    var it = psvDados_(nome, a).itens.filter(function (i) { return !i.resp; }).map(function (i) { return String(i.desc); });
    if (it.length) { por[nome] = it; todos = todos.concat(it); }
  });
  return { por: por, todos: todos };
}
function avFotosPend_() {
  var sh = criarAbaFS_(SpreadsheetApp.getActive()), n = sh.getLastRow(), todos = []; if (n < 2) return { por: {}, todos: todos };
  sh.getRange(2, 1, n - 1, 5).getValues().forEach(function (l) { if (!String(l[0]).trim()) return; var st = String(l[4] || '').trim().toUpperCase() || 'PENDENTE'; if (st === 'PENDENTE') todos.push(String(l[1])); });
  return { por: {}, todos: todos };
}
// v120.1 (Celio 14/11/2026): pendência parada há X dias (padrão 2) nas listas das repositoras / caixa / Thynaylly -> avisa o Ney para cobrar.
// Uma notificação por pessoa, com o que está parado em cada lista. Aba AVISOS AUTOMATICOS, código COBRAR_NEY (DIAS = quantos dias parado).
function avCobrar_(c) {
  var lim = Number(c && c.dias) > 0 ? Number(c.dias) : 2, hoje = diaNum_(v9Hoje_()), por = {}, ss = SpreadsheetApp.getActive();
  if (new Date().getDay() === 0) return 0; // domingo não cobra
  var dn = function (s) { return diaNum_(String(s || '').trim().slice(0, 10)); };
  var add = function (pessoa, lista, d) {
    if (isNaN(d) || hoje - d < lim) return;
    String(pessoa || '').split(/\s*\/\s*/).forEach(function (p) { p = p.trim(); if (!p || verTudo_(p)) return; var o = por[p] = por[p] || {}; o[lista] = (o[lista] || 0) + 1; });
  };
  var dv = function (nome, nc) { var sh = ss.getSheetByName(nome); if (!sh || sh.getLastRow() < 2) return []; return sh.getRange(2, 1, sh.getLastRow() - 1, nc).getDisplayValues(); };
  try { dv(ABA_CVS, 11).forEach(function (l) { if (String(l[10]) === 'PENDENTE' && !vfVend_(l[5], l[4], l[3])) add(l[1], 'validades da semana', dn(l[0])); }); } catch (e) {}
  try { dv(ABA_TF, 10).forEach(function (l) { if (String(l[9]) === 'PENDENTE') add(l[6], 'troca de frente', dn(l[0])); }); } catch (e) {}
  try { dv(ABA_ACH, 12).forEach(function (l) { if (String(l[11]) === 'PENDENTE') add(l[5], 'achar na loja', dn(l[1])); }); } catch (e) {}
  try { dv(ABA_CTA, 8).forEach(function (l) { if (String(l[7]) === 'AGUARDANDO CONTAGEM') add(l[5], 'contagem de estoque', dn(l[1])); }); } catch (e) {}
  try { // vencidos/avaria da semana: conta a partir do "retirar até" (ou do início da semana, o que for depois)
    var qv = v9ComAcesso_('VENCIDOS').filter(function (q) { return !verTudo_(q); }).join(' / ');
    if (qv) dv(ABA_VS, 9).forEach(function (l) { if (String(l[8]) === 'PENDENTE') add(qv, 'vencidos para retirar', Math.max(dn(l[0]) || 0, dn(l[1]) || 0)); });
  } catch (e) {}
  try { // produto sem venda: o mesmo que aparece no app de cada uma
    var dt = {}, vistos = {}; dv(ABA_PSV, 15).forEach(function (l) { if (l[14]) dt[String(l[14])] = dn(l[0]); });
    v9ComAcesso_('SEMVENDA').concat(v9ComAcesso_('SEMVENDA ALERTA')).forEach(function (q) {
      if (verTudo_(q) || vistos[q]) return; vistos[q] = 1; var a = psvAcesso_(q); if (!a) return;
      psvDados_(q, a).itens.forEach(function (i) { if (!i.resp) add(q, 'sem venda', dt[String(i.id)]); });
    });
  } catch (e) {}
  var para = String((c && c.para) || 'Ney').split(/[,;]+/).map(function (s) { return s.trim(); }).filter(String), n = 0, tit0 = (c && c.tit) || '📣 Cobrar: {pessoa}';
  Object.keys(por).forEach(function (p) {
    var o = por[p], partes = Object.keys(o).map(function (k) { return o[k] + ' ' + k; });
    var tx = p + ' está há ' + lim + '+ dias sem fazer: ' + partes.join(', ') + '.';
    try { var r = enviarPush_(para, tit0.replace(/\{pessoa\}/g, p), tx.slice(0, 230), { abrir: 'aviso', tag: 'av-cobrar-' + avK_(p).replace(/\W+/g, '_') }); if (r && r.aparelhos) n++; } catch (e) {}
  });
  return n;
}
var AV_SEMANA = ['DOM', 'SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB'];
function avDiaOk_(dias, d) { var s = semAcento_(String(dias || '')).toUpperCase(); if (!s || /TODOS|TODO DIA/.test(s) || !isNaN(Number(s))) return true; return s.indexOf(AV_SEMANA[d.getDay()]) >= 0; }
// gatilho de hora em hora (instalado por avisosInstalar)
function avisosHora() {
  var agora = new Date(), h = agora.getHours(), hoje = vsBr_(vsDia_(agora)), res = {};
  if (h < 6 || h > 21) return { ok: true, fora: true };
  CacheService.getScriptCache().remove('AV_CFG');
  if (h >= 7 && h <= 20) { try { sincronizarRebaixas(); } catch (e) {} }
  try { ctaAtualizar_(); } catch (e) {} // v130
  try { var shPf = pfAba_(); res.pfVolta = pfExpira_(shPf, pfLinhas_(shPf)); } catch (e) {} // v110.3: preço separado pela Nath há mais de 2 h volta para a fila do Claude
  if (h >= 8) { try { res.vf = vfDiario_(); } catch (e) {} }
  try { res.gond = decGondolaVencidas_(); } catch (e) {} // v121: conferir na gôndola com validade passada // v109: confrontação diária saldo × lotes de validade // rebaixa nova sai na hora em que entra no Plenus
  try { // notas que chegaram desde a última hora
    var pr = PropertiesService.getScriptProperties(), f = JSON.parse(pr.getProperty('AV_NF_FILA') || '[]');
    if (f.length) { pr.deleteProperty('AV_NF_FILA'); res.nf = avDisparar_('NF_CHEGOU', {}, f, { para: 'Regiane', tit: '📦 Nota para conferir', tx: '{qtd} nota(s) para conferir: {lista}' }); }
  } catch (e) {}
  avTodos_().forEach(function (c) {
    if (!c.lig || c.hora === 'NA HORA' || c.hora === '' || isNaN(Number(c.hora))) return;
    if (h < Number(c.hora) || String(c.ult).indexOf(hoje) === 0) return; // ainda não é a hora / já foi hoje
    if (!avDiaOk_(c.dias, agora)) return;
    try {
      var d = null, lemb = false;
      if (c.cod === 'REB_VENCE_HOJE') d = avRebVence_(0);
      else if (c.cod === 'REB_VENCE_AMANHA') d = avRebVence_(1);
      else if (c.cod === 'CV_SEMANA') d = avCvSemana_();
      else if (c.cod === 'TF_LEMBRETE') { d = avTfPend_(); lemb = true; }
      else if (c.cod === 'PSV_LEMBRETE') { d = avPsvPend_(); lemb = true; }
      else if (c.cod === 'FOTOS_PEND') { d = avFotosPend_(); lemb = true; }
      else if (c.cod === 'COBRAR_NEY') { var nc = avCobrar_(c); avMarcar_(c, agoraBr_() + ' · ' + nc + (nc === 1 ? ' aviso' : ' avisos')); res[c.cod] = nc; return; } // v120.1
      else if (/^LEMBRETE/.test(c.cod)) { // lembrete livre: DIAS = dias da semana ou de quantos em quantos dias
        var cada = Number(c.dias); if (cada > 0 && c.ult) { var u = diaNum_(String(c.ult).slice(0, 10)); if (!isNaN(u) && diaNum_(hoje) - u < cada) return; }
        var nn = 0; String(c.para).split(/[,;]+/).map(function (s) { return s.trim(); }).filter(String).forEach(function (p) { try { var r = enviarPush_([p], c.tit || 'Lembrete', String(c.tx).slice(0, 230), { abrir: 'aviso', tag: 'av-' + c.cod.toLowerCase() }); if (r && r.aparelhos) nn++; } catch (e) {} });
        avMarcar_(c, agoraBr_() + ' · ' + nn + (nn === 1 ? ' pessoa' : ' pessoas')); res[c.cod] = nn; return;
      } else return;
      if (lemb) avZerar_(c.cod, Object.keys(d.por).concat(d.todos.length ? String(c.para).split(/[,;]+/).map(function (s) { return s.trim(); }) : []));
      res[c.cod] = avDisparar_(c.cod, d.por, d.todos, { lembrete: lemb });
      avMarcar_(c, agoraBr_() + ' · ' + (res[c.cod] || 0) + ((res[c.cod] || 0) === 1 ? ' pessoa' : ' pessoas'));
    } catch (e) { res[c.cod] = 'erro: ' + e; }
  });
  return { ok: true, hora: h, res: res };
}
function avisosInstalar() {
  avAba_(); CacheService.getScriptCache().remove('AV_CFG');
  var tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'avisosHora'; });
  if (!tem) ScriptApp.newTrigger('avisosHora').timeBased().everyHours(1).create();
  return { ok: true, gatilho: tem ? 'já existia' : 'criado', avisos: avTodos_().map(function (c) { return c.cod + ':' + (c.lig ? 'SIM' : 'NAO'); }) };
}

// ========== v104 (07/11/2026): VALIDADE FORA DA LISTA ==========
// A repositora acha na reposição um produto com validade que não veio na lista dela (conferência deixou passar).
// Ela bipa (ou busca pelo nome), põe a validade e conta área + estoque. Ela só vê "Salvo!".
// Por trás: data nova -> entra como lote na VALIDADE PRODUTOS (aba PRODUTOS); data que já existe -> não duplica.
// Nos dois casos faz a conta de venda até o vencimento e, se for sobrar, manda a sugestão para o A DECIDIR da REBAIXAS.
// Fica registrado na aba VALIDADES CORRIGIDAS (tipo ACHADA NA LOJA), com QTD / CRITICA / SOBRA nas colunas M:O.
var CAB_VCOR_X = ['QTD', 'CRITICA', 'SOBRA'];
function cvAvProd_(p) { // produto pelo código de barras ou pelo SKU
  var r = null;
  if (String(p.codigo || '').replace(/\D/g, '')) r = buscarProduto(p.codigo);
  if ((!r || !r.ok) && skuN_(p.sku)) {
    var cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = cad.getLastRow();
    var cel = n >= 2 ? cad.getRange(2, 2, n - 1, 1).createTextFinder(skuN_(p.sku)).matchEntireCell(true).findNext() : null;
    if (cel) { var l = cad.getRange(cel.getRow(), 1, 1, 7).getValues()[0]; r = String(l[6]) === 'NV' ? { ok: false, naoVender: true } : { ok: true, sku: String(l[1]), desc: String(l[2]), un: String(l[3]) }; }
  }
  if (!r || !r.ok) return { ok: false, erro: r && r.naoVender ? 'produto marcado como NÃO VENDER' : 'produto não encontrado' };
  var c = cadSku_(cvCad_(SpreadsheetApp.openById(ID_VALIDADE)), r.sku, r.desc) || {}, med = cvMedias_()[skuN_(r.sku)];
  return { ok: true, sku: skuN_(r.sku), desc: String(r.desc).replace(/\*+\s*$/, '').trim(), secao: c.secao || '', local: c.local || '', media: med == null ? '' : Math.round(med * 10) / 10 };
}
function cvAvInfo(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('VALIDADES') < 0) return { ok: false, erro: 'sem acesso' };
  return cvAvProd_(p);
}
function cvAvulso(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('VALIDADES') < 0) return { ok: false, erro: 'sem acesso' };
  var val = String(p.validade || '').trim();
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(val) || isNaN(diaNum_(val))) return { ok: false, erro: 'escolha a validade' };
  if (diaNum_(val) < diaNum_(vsBr_(vsDia_(new Date())))) return { ok: false, erro: 'essa validade já passou — produto vencido vai para o módulo de avaria' };
  var num = function (x) { x = String(x == null ? '' : x).trim().replace(',', '.'); return x === '' ? null : Number(x); };
  var area = num(p.area), est = num(p.estoque);
  if (area == null || isNaN(area) || area < 0) return { ok: false, erro: 'informe a quantidade da área de venda' };
  if (est != null && (isNaN(est) || est < 0)) return { ok: false, erro: 'quantidade do estoque inválida' };
  var total = area + (est || 0); if (!total) return { ok: false, erro: 'quantidade zerada' };
  var pr = cvAvProd_({ sku: p.sku }); if (!pr.ok) return pr;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone(), sh = ssV.getSheetByName('PRODUTOS'), n = sh.getLastRow();
    var k = normDesc_(pr.desc), dN = diaNum_(val), existe = null, ult = null;
    (n >= 2 ? sh.getRange(2, 1, n - 1, 8).getValues() : []).forEach(function (x) {
      if (!(x[1] instanceof Date) || normDesc_(String(x[3]).replace(/\*+\s*$/, '')) !== k) return;
      if (Math.abs(diaNum_(Utilities.formatDate(x[1], tzV, 'dd/MM/yyyy')) - dN) < 1 && !existe) existe = x;
      if (!ult || (x[6] instanceof Date && (!(ult[6] instanceof Date) || x[6] >= ult[6]))) ult = x;
    });
    if (!existe) { // lote novo achado na loja
      var colD = sh.getRange(1, 4, sh.getMaxRows(), 1).getValues(), last = 1;
      for (var i = colD.length - 1; i >= 1; i--) if (String(colD[i][0]).trim()) { last = i + 1; break; }
      var hojeV = Utilities.parseDate(vsBr_(vsDia_(new Date())), tzV, 'dd/MM/yyyy');
      sh.insertRowsAfter(last, 1);
      sh.getRange(last + 1, 1, 1, 8).setValues([['', Utilities.parseDate(val, tzV, 'dd/MM/yyyy'), pr.secao || (ult ? String(ult[2] || '') : ''), pr.desc, 'UN', total, hojeV, 'ACHADO NA LOJA']]);
    }
    // conta de venda até o vencimento (mesma regra da conferência de validade)
    var med = cvMedias_()[pr.sku], dias = Math.max(0, Math.round((vsDeBr_(val) - vsDia_(new Date())) / 86400000)), sobra = '', crit = 'SEM MEDIA';
    if (med != null) { sobra = Math.max(0, Math.ceil(total - Math.round(med * dias * 10) / 10)); crit = sobra > 0 ? 'SUGERIR REBAIXA' : 'VAI VENDER'; if (!sobra) sobra = ''; }
    if (crit === 'SUGERIR REBAIXA') { try { var stR = rebSituacao_(rebIndice_(), pr.sku, pr.desc, val); if (stR && stR.tipo === 'REB') crit = 'JA NA REBAIXA ' + stR.por; else if (!stR) rebSugerirLote_([{ sku: pr.sku, desc: pr.desc, validade: val, obs: obsQtd_(total) + (est == null ? ' AREA (deposito nao conferido)' : '') + ' - deve sobrar ' + sobra + ' (achado na loja por ' + dono + ')' }]); } catch (e) {} }
    // registro para o Celio
    var ref = existe || ult, nf = ref ? String(ref[0]).replace(/\D/g, '') : '', forn = ref && String(ref[7]) !== 'ACHADO NA LOJA' ? String(ref[7] || '') : '';
    var lg = cvAba_(ABA_VCOR, CAB_VCOR), m = lg.getLastRow();
    if (String(lg.getRange(1, 13).getValue()) !== 'QTD') lg.getRange(1, 13, 1, 3).setValues([CAB_VCOR_X]).setFontWeight('bold');
    lg.getRange(m + 1, 1, 1, 15).setNumberFormat('@').setValues([[agoraBr_(), dono, pr.sku, pr.desc, '', val, '', existe ? 'ACHADA (JA EXISTIA)' : 'ACHADA NA LOJA', nf, forn, nf ? cvConferente_(nf, forn) : '', 'AVULSO', String(total) + (est == null ? ' (só área)' : ''), crit, String(sobra)]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function cvAvMeus_(quem) { // o que a pessoa lançou fora da lista nos últimos 7 dias (mais recente primeiro)
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_VCOR), out = []; if (!sh || sh.getLastRow() < 2) return out;
  var lim = diaNum_(vsBr_(vsDia_(new Date()))) - 7, q = semAcento_(String(quem || '')).toLowerCase().trim();
  sh.getRange(2, 1, sh.getLastRow() - 1, 15).getValues().forEach(function (x) {
    if (!/^ACHADA/.test(String(x[7])) || semAcento_(String(x[1])).toLowerCase().trim() !== q || diaNum_(String(x[0]).slice(0, 10)) < lim) return;
    out.push({ em: String(x[0]), desc: String(x[3]), val: String(x[5]), qtd: String(x[12]) });
  });
  return out.reverse();
}

// ========== v109 (08/11/2026): PEDIR ETIQUETA · CONTAGEM DE ESTOQUE · VALIDADE NA FRENTE · HORTIFRUTI (consumo próprio) ==========
function v9Dono_(p, cods) { // autenticação + acesso (o Celio entra em tudo)
  var d = quemDoToken_(p.token);
  if (!d || d !== String(p.quem || '')) return { erro: { ok: false, erro: 'sessão expirada', reauth: true } };
  if (cods && !ehCelio_(d)) { var ap = apps(d); if (![].concat(cods).some(function (c) { return ap.indexOf(c) >= 0; })) return { erro: { ok: false, erro: 'sem acesso' } }; }
  return { d: d };
}
function v9Aba_(nome, cab) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(nome);
  if (sh) return sh;
  sh = ss.insertSheet(nome);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), cab.length).setNumberFormat('@');
  return sh;
}
function v9Linhas_(sh, nc) { var n = sh.getLastRow(); return n >= 2 ? sh.getRange(2, 1, n - 1, nc).getValues() : []; }
function v9Linha_(sh, col, id) { // nº da linha pelo ID (coluna col)
  var n = sh.getLastRow(); if (n < 2 || !id) return -1;
  var v = sh.getRange(2, col, n - 1, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === String(id)) return i + 2;
  return -1;
}
function v9Hoje_() { return vsBr_(vsDia_(new Date())); }
function v9Num_(x) { x = String(x == null ? '' : x).trim().replace(',', '.'); return x === '' ? null : Number(x); }
function v9Br_(v) { // Date ou texto → dd/MM/yyyy
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (!m) return '';
  var a = m[3].length === 2 ? '20' + m[3] : m[3]; return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + a;
}
// v118.6: reposição e média de venda de um SKU sem ler as planilhas inteiras (TextFinder)
function v9RepRapido_(s, desc) {
  try {
    var sh = SpreadsheetApp.openById(ID_VALIDADE).getSheetByName('CADASTRO'), n = sh.getLastRow();
    if (s && n >= 2) { var cel = sh.getRange(2, 5, n - 1, 1).createTextFinder(String(s)).matchEntireCell(true).findNext(); if (cel) { var l = sh.getRange(cel.getRow(), 1, 1, 5).getValues()[0]; if (l[3]) return { secao: String(l[1] || ''), local: String(l[2] || ''), rep: String(l[3] || '') }; } }
  } catch (e) {}
  return cadSku_(mapaReposicao_(), s, desc);
}
function v9MediaRapida_(s) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_MED);
  if (!s || !sh || sh.getLastRow() < 2) return null;
  var cel = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder('^0*' + s + '$').useRegularExpression(true).findNext();
  if (!cel) return null;
  return Number(String(sh.getRange(cel.getRow(), 3).getValue()).replace(',', '.')) || 0;
}
function v9Prod_(p) { // produto pelo código de barras ou SKU: sku, desc, un, preço, seção, local, repositora, média de venda
  var r = null;
  if (String(p.codigo || '').replace(/\D/g, '')) r = buscarProduto(p.codigo);
  var s = skuN_(p.sku || (r && r.ok ? r.sku : ''));
  var cad = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = cad.getLastRow(), l = null;
  if (s && n >= 2) { var cel = cad.getRange(2, 2, n - 1, 1).createTextFinder(s).matchEntireCell(true).findNext(); if (cel) l = cad.getRange(cel.getRow(), 1, 1, 7).getValues()[0]; }
  if (!l) return { ok: false, erro: 'produto não encontrado' };
  if (String(l[6]) === 'NV') return { ok: false, erro: 'produto marcado como NÃO VENDER' };
  var desc = String(l[2]).replace(/\*+\s*$/, '').trim(), c = v9RepRapido_(s, desc) || {}, med = v9MediaRapida_(s); // v118.6: busca só a linha do produto (antes lia as planilhas inteiras)
  return { ok: true, sku: s, desc: desc, un: String(l[3]), preco: Number(l[5]) || 0, secao: c.secao || '', local: c.local || '', rep: c.rep || '', media: med == null ? '' : Math.round(med * 10) / 10 };
}
function v9Saldos_() { // saldo atual.txt → { itens: {sku: {q, custo}}, em: Date }
  var f = DriveApp.getFileById(ID_SALDO_TXT), m = {};
  saldoTxt_(f).split(/\r?\n/).forEach(function (l, i) {
    if (!i) return; var c = l.split(';'); if (c.length < 7) return; var s = skuN_(c[0]); if (!s) return;
    var n = function (x) { var v = Number(String(x || '').replace(/\./g, '').replace(',', '.')); return isNaN(v) ? null : v; };
    m[s] = { q: n(c[6]), custo: n(c[5]) };
  });
  return { itens: m, em: f.getLastUpdated() };
}
function v9SaldoSku_(sku) {
  var f = DriveApp.getFileById(ID_SALDO_TXT), t = saldoTxt_(f);
  var mm = t.match(new RegExp('(?:^|\\n)0*' + skuN_(sku) + ';[^\\n]*')), q = null, custo = null;
  if (mm) { var c = mm[0].replace(/^\n/, '').split(';'); q = Number(String(c[6] || '').replace(/\./g, '').replace(',', '.')); custo = Number(String(c[5] || '').replace(/\./g, '').replace(',', '.')); if (isNaN(q)) q = null; if (isNaN(custo)) custo = null; }
  return { qtd: q, custo: custo, em: f.getLastUpdated() };
}
function v9Push_(para, tit, tx, abrir) { try { return enviarPush_([].concat(para), tit, String(tx).slice(0, 230), { abrir: abrir || 'aviso', tag: 'v9-' + (abrir || 'aviso') }); } catch (e) { return null; } }
function v9ComAcesso_(cod) { // nomes com o módulo na aba ACESSOS
  var sh = criarAbaAcessos_(SpreadsheetApp.getActive()), out = [];
  v9Linhas_(sh, 2).forEach(function (l) { if (String(l[1]).trim().toUpperCase() === cod) { var q = String(l[0]).trim(); if (q && out.indexOf(q) < 0) out.push(q); } });
  return out;
}

// ---------- 🏷️ PEDIR ETIQUETA (repositoras: ETIQUETAS · fila de impressão da Nath: ETIQUETAS_FILA) ----------
var ABA_ETQ = 'ETIQUETAS';
var CAB_ETQ = ['ID', 'DATA', 'HORA', 'QUEM', 'SKU', 'PRODUTO', 'PRECO', 'LOCAL', 'TAMANHO', 'QTD', 'STATUS', 'RESERVADO POR', 'RESERVADO EM', 'IMPRESSO POR', 'IMPRESSO EM', 'OBS'];
function etqObj_(l) { return { id: String(l[0]), data: v9Br_(l[1]), hora: l[2] instanceof Date ? fmt_(l[2], 'HH:mm') : String(l[2]), // v111.3: a planilha vira data/hora em Date -> texto certo
 quem: String(l[3]), sku: String(l[4]), desc: String(l[5]), preco: String(l[6]), local: String(l[7]), tam: String(l[8]), qtd: Number(l[9]) || 1, status: String(l[10]), resPor: String(l[11]), resEm: pfBr_(l[12]), impPor: String(l[13]), impEm: pfBr_(l[14]), obs: String(l[15]) }; }
function etqRolo_(its) { var g = 0, p = 0; its.forEach(function (i) { if (i.tam === 'GRANDE') g += i.qtd; else p += i.qtd; }); return { grandes: g, pequenas: p, etiquetas: g + Math.ceil(p / 2) }; }
function etqInfo(p) {
  var a = v9Dono_(p, ['ETIQUETAS', 'ETIQUETAS_FILA']); if (a.erro) return a.erro;
  var r = v9Prod_(p); if (!r.ok) return r;
  return { ok: true, sku: r.sku, desc: r.desc, preco: r.preco, local: r.local, secao: r.secao };
}
function etqPedir(p) {
  var a = v9Dono_(p, ['ETIQUETAS', 'ETIQUETAS_FILA']); if (a.erro) return a.erro;
  var tam = String(p.tam || '').toUpperCase(), q = Math.round(Number(p.qtd) || 0);
  if (['PEQUENA', 'GRANDE'].indexOf(tam) < 0) return { ok: false, erro: 'escolha PEQUENA ou GRANDE' };
  if (q < 1 || q > 50) return { ok: false, erro: 'quantidade de 1 a 50' };
  var r = v9Prod_(p); if (!r.ok) return r;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_ETQ, CAB_ETQ), ag = new Date(), tz = Session.getScriptTimeZone(), id = String(p.id || Utilities.getUuid());
    if (v9Linha_(sh, 1, id) > 0) return { ok: true, id: id, dup: true };
    sh.appendRow([id, Utilities.formatDate(ag, tz, 'dd/MM/yyyy'), Utilities.formatDate(ag, tz, 'HH:mm'), a.d, r.sku, r.desc, String(r.preco.toFixed(2)).replace('.', ','), r.local, tam, String(q), 'NA FILA', '', '', '', '', '']);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
function etqMeus(p) {
  var a = v9Dono_(p, ['ETIQUETAS', 'ETIQUETAS_FILA']); if (a.erro) return a.erro;
  var hoje = v9Hoje_(), out = [];
  v9Linhas_(v9Aba_(ABA_ETQ, CAB_ETQ), CAB_ETQ.length).forEach(function (l) {
    var o = etqObj_(l); if (o.quem !== a.d || o.status === 'CANCELADA') return;
    if (o.data === hoje || o.status === 'NA FILA' || o.status === 'EM IMPRESSAO') out.push(o);
  });
  return { ok: true, itens: out.reverse() };
}
function etqCancelar(p) {
  var a = v9Dono_(p, ['ETIQUETAS', 'ETIQUETAS_FILA']); if (a.erro) return a.erro;
  var sh = v9Aba_(ABA_ETQ, CAB_ETQ), r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'pedido não encontrado' };
  var l = sh.getRange(r, 1, 1, CAB_ETQ.length).getValues()[0];
  if (String(l[3]) !== a.d && !ehCelio_(a.d)) return { ok: false, erro: 'só quem pediu pode cancelar' };
  if (String(l[10]) !== 'NA FILA') return { ok: false, erro: 'já está sendo impressa' };
  sh.getRange(r, 11).setValue('CANCELADA'); return { ok: true };
}
function etqFila(p) {
  var a = v9Dono_(p, 'ETIQUETAS_FILA'); if (a.erro) return a.erro;
  var hoje = v9Hoje_(), fila = [], emImp = [], impressas = [];
  v9Linhas_(v9Aba_(ABA_ETQ, CAB_ETQ), CAB_ETQ.length).forEach(function (l) {
    var o = etqObj_(l);
    if (o.status === 'NA FILA') fila.push(o); else if (o.status === 'EM IMPRESSAO') emImp.push(o);
    else if (o.status === 'IMPRESSA' && String(o.impEm).indexOf(hoje) === 0) impressas.push(o);
  });
  var h = new Date().getHours(), prox = h < 7 ? '7h' : (h < 12 ? '12h' : 'amanhã às 7h');
  return { ok: true, fila: fila, emImp: emImp, impressas: impressas.reverse(), rolo: etqRolo_(fila), roloEm: etqRolo_(emImp), roloHoje: etqRolo_(impressas), proxima: prox };
}
function etqMudar_(ids, de, para, quem) { // muda o status de vários pedidos; devolve os que mudaram
  var sh = v9Aba_(ABA_ETQ, CAB_ETQ), vals = v9Linhas_(sh, CAB_ETQ.length), agora = agoraBr_(), feitos = [];
  var set = {}; [].concat(ids || []).forEach(function (x) { set[String(x)] = 1; });
  vals.forEach(function (l, i) {
    if (!set[String(l[0])] || de.indexOf(String(l[10])) < 0) return;
    var r = i + 2;
    if (para === 'EM IMPRESSAO') sh.getRange(r, 11, 1, 3).setValues([['EM IMPRESSAO', quem, agora]]);
    else if (para === 'IMPRESSA') { sh.getRange(r, 11).setValue('IMPRESSA'); sh.getRange(r, 14, 1, 2).setValues([[quem, agora]]); }
    else if (para === 'NA FILA') { sh.getRange(r, 11, 1, 3).setValues([['NA FILA', '', '']]); if (quem) sh.getRange(r, 16).setValue(String(quem)); }
    feitos.push(etqObj_(l));
  });
  return feitos;
}
function etqAvisarProntas_(its, por) {
  var porQuem = {};
  its.forEach(function (o) { (porQuem[o.quem] = porQuem[o.quem] || []).push(o); });
  Object.keys(porQuem).forEach(function (q) {
    var l = porQuem[q], nomes = l.map(function (o) { return o.qtd + '× ' + o.desc; }).join(', ');
    v9Push_(q, '🏷️ Etiqueta pronta', (l.length === 1 ? 'Sua etiqueta ficou pronta: ' : l.length + ' etiquetas prontas: ') + nomes + (por ? ' (' + por + ')' : ''), 'etiquetas');
  });
}
function etqAcao(p) {
  var a = v9Dono_(p, 'ETIQUETAS_FILA'); if (a.erro) return a.erro;
  var ids = p.ids; if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) { ids = String(ids).split(','); } }
  if (!ids || !ids.length) return { ok: false, erro: 'marque pelo menos uma etiqueta' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var op = String(p.op || ''), f = [];
    if (op === 'RESERVAR') f = etqMudar_(ids, ['NA FILA'], 'EM IMPRESSAO', a.d);
    else if (op === 'CONFIRMAR') { f = etqMudar_(ids, ['EM IMPRESSAO', 'NA FILA'], 'IMPRESSA', a.d); etqAvisarProntas_(f, ''); }
    else if (op === 'DEVOLVER') f = etqMudar_(ids, ['EM IMPRESSAO'], 'NA FILA', '');
    else if (op === 'NAO_SAIU') f = etqMudar_(ids, ['EM IMPRESSAO', 'IMPRESSA'], 'NA FILA', 'NÃO SAIU ' + agoraBr_() + ' (' + a.d + ')');
    else return { ok: false, erro: 'ação inválida' };
    return { ok: true, n: f.length };
  } finally { lock.releaseLock(); }
}
function etqClaude(p) { // tarefa agendada do Claude (7h e 12h): LISTAR a fila / IMPRESSO (ids) depois de imprimir na Argox
  if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' };
  if (String(p.op || '') === 'IMPRESSO') {
    var ids = p.ids; if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) { ids = String(ids).split(','); } }
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { var f = etqMudar_(ids, ['NA FILA'], 'IMPRESSA', 'CLAUDE'); etqAvisarProntas_(f, 'impressa pelo Claude'); return { ok: true, n: f.length }; } finally { lock.releaseLock(); }
  }
  var fila = v9Linhas_(v9Aba_(ABA_ETQ, CAB_ETQ), CAB_ETQ.length).map(etqObj_).filter(function (o) { return o.status === 'NA FILA'; });
  return { ok: true, itens: fila, rolo: etqRolo_(fila) };
}

// ---------- v110.3: FILA DE PREÇOS ON-LINE (Nath e Celio: POFF_FILA) ----------
// PRECOS OFF ganha K RESERVADO POR, L RESERVADO EM, M ATIVADO POR. Status: CONFIRMADO (fila do Claude) -> EM ATIVACAO (Nath separou; o Claude não mexe)
// -> ATIVADO. Separado há mais de 2 h sem confirmar volta sozinho para CONFIRMADO.
function pfAba_() {
  var sh = criarAbaPoff_(SpreadsheetApp.getActive());
  if (String(sh.getRange(1, 11).getValue()) !== 'RESERVADO POR') sh.getRange(1, 11, 1, 3).setValues([['RESERVADO POR', 'RESERVADO EM', 'ATIVADO POR']]).setFontWeight('bold').setBackground('#0369a1').setFontColor('#ffffff');
  return sh;
}
function pfLinhas_(sh) { var n = sh.getLastRow(); return n >= 2 ? sh.getRange(2, 1, n - 1, 13).getValues() : []; }
function pfBr_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(v || ''); }
function pfMs_(s) { var m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/); return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]).getTime() : 0; }
function pfExpira_(sh, vals) { // separado há mais de 2 h sem confirmar -> volta para a fila do Claude
  var lim = Date.now() - 2 * 3600 * 1000, n = 0;
  vals.forEach(function (l, i) {
    if (String(l[5]).toUpperCase() !== 'EM ATIVACAO' || pfMs_(pfBr_(l[11])) > lim) return;
    sh.getRange(i + 2, 6).setValue('CONFIRMADO'); sh.getRange(i + 2, 10).setValue((l[9] ? l[9] + ' | ' : '') + 'voltou para a fila (' + l[10] + ' não confirmou em 2 h)');
    sh.getRange(i + 2, 11, 1, 2).setValues([['', '']]); l[5] = 'CONFIRMADO'; l[10] = ''; l[11] = ''; n++;
  });
  return n;
}
function pfObj_(l, i) {
  return { id: chavePoff_(l[0], l[2]), lin: i + 2, sku: String(l[0]), desc: String(l[1]), novo: Number(l[2]) || 0, atual: Number(l[3]) || 0, status: String(l[5]).toUpperCase(), quem: String(l[6]),
    confEm: pfBr_(l[7]), ativEm: pfBr_(l[8]), obs: String(l[9] || ''), resPor: String(l[10] || ''), resEm: pfBr_(l[11]), ativPor: String(l[12] || ''), fardo: !!baseFardo_(l[1]) };
}
function pfFila(p) {
  var a = v9Dono_(p, 'POFF_FILA'); if (a.erro) return a.erro;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = pfAba_(), vals = pfLinhas_(sh); pfExpira_(sh, vals);
    var hoje = v9Hoje_(), fila = [], em = [], ativ = [];
    vals.forEach(function (l, i) {
      var o = pfObj_(l, i); if (!o.sku) return;
      if (o.status === 'CONFIRMADO') fila.push(o); else if (o.status === 'EM ATIVACAO') em.push(o);
      else if (o.status === 'ATIVADO' && o.ativEm.indexOf(hoje) === 0) ativ.push(o);
    });
    var h = new Date(), hm = h.getHours() * 60 + h.getMinutes(), dw = h.getDay();
    var slots = [530, 590, 650, 830, 890, 950], prox = ''; for (var k = 0; k < slots.length; k++) if (dw !== 0 && slots[k] > hm) { prox = Math.floor(slots[k] / 60) + 'h' + ('0' + slots[k] % 60).slice(-2); break; }
    return { ok: true, fila: fila, emAt: em, ativados: ativ.reverse(), proxima: prox ? 'hoje às ' + prox : (dw === 6 ? 'segunda às 8h50' : 'amanhã às 8h50') };
  } finally { lock.releaseLock(); }
}
function pfAcao(p) {
  var a = v9Dono_(p, 'POFF_FILA'); if (a.erro) return a.erro;
  var ids = p.ids; if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) { ids = String(ids).split(','); } }
  if (!ids || !ids.length) return { ok: false, erro: 'marque pelo menos um produto' };
  var set = {}; ids.forEach(function (x) { set[String(x)] = 1; });
  var op = String(p.op || ''), de = { RESERVAR: ['CONFIRMADO'], CONFIRMAR: ['EM ATIVACAO', 'CONFIRMADO'], DEVOLVER: ['EM ATIVACAO'], NAO_ATIVOU: ['EM ATIVACAO'] }[op];
  if (!de) return { ok: false, erro: 'ação inválida' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = pfAba_(), vals = pfLinhas_(sh), agora = agoraBr_(), n = 0, avisar = {};
    vals.forEach(function (l, i) {
      var o = pfObj_(l, i); if (!set[o.id] || de.indexOf(o.status) < 0) return; var r = i + 2;
      if (op === 'RESERVAR') { sh.getRange(r, 6).setValue('EM ATIVACAO'); sh.getRange(r, 11, 1, 2).setValues([[a.d, agora]]); }
      else if (op === 'CONFIRMAR') { sh.getRange(r, 6).setValue('ATIVADO'); sh.getRange(r, 9).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm'); sh.getRange(r, 13).setValue(a.d); (avisar[o.quem] = avisar[o.quem] || []).push(o.desc); }
      else { sh.getRange(r, 6).setValue('CONFIRMADO'); sh.getRange(r, 11, 1, 2).setValues([['', '']]); if (op === 'NAO_ATIVOU') sh.getRange(r, 10).setValue((o.obs ? o.obs + ' | ' : '') + 'não ativou ' + agora + ' (' + a.d + ')'); }
      n++;
    });
    Object.keys(avisar).forEach(function (q) { if (!q || q === a.d) return; try { v9Push_(q, '✅ Preço ativado', avisar[q].length === 1 ? avisar[q][0] + ' já está com o preço novo no caixa' : avisar[q].length + ' preços ativados: ' + avisar[q].slice(0, 3).join(', '), 'precosoff'); } catch (e) {} });
    return { ok: true, n: n };
  } finally { lock.releaseLock(); }
}
function pfResumo_(quem, r) {
  var vals = pfLinhas_(pfAba_()), f = 0, e = 0;
  vals.forEach(function (l) { var s = String(l[5]).toUpperCase(); if (s === 'CONFIRMADO') f++; else if (s === 'EM ATIVACAO') e++; });
  r.pfFila = f; r.pfEm = e;
}

// ---------- 📦 CONTAGEM DE ESTOQUE (quem conta: CONTAGEM_ESTOQUE · pedir e decidir: só o Celio) ----------
var ABA_CTA = 'CONTAGEM ESTOQUE';
var CAB_CTA = ['ID', 'PEDIDO EM', 'PEDIDO POR', 'SKU', 'PRODUTO', 'PARA', 'MOTIVO', 'STATUS', 'AREA', 'ESTOQUE', 'TOTAL', 'CONTADO POR', 'CONTADO EM', 'SALDO ARQUIVO', 'ARQUIVO EM', 'VENDA ESTIMADA', 'SALDO SISTEMA', 'DIFERENCA', 'CUSTO UN', 'DIFERENCA R$', 'DECISAO', 'DECIDIDO EM', 'LANCAMENTO', 'LANCADO EM', 'ORIGEM', 'OBS'];
function ctaObj_(l) {
  var o = {}; l = l.map(function (x) { return x instanceof Date ? agdTxt_(x) : x; }); // v127.1: datas que a planilha converteu (PEDIDO EM etc.) — antes sumiam do histórico
  ['id', 'pedEm', 'pedPor', 'sku', 'desc', 'para', 'motivo', 'status', 'area', 'est', 'total', 'contPor', 'contEm', 'saldoArq', 'arqEm', 'vendaEst', 'saldoSis', 'dif', 'custo', 'difRs', 'decisao', 'decEm', 'lanc', 'lancEm', 'origem', 'obs'].forEach(function (k, i) { o[k] = String(l[i] == null ? '' : l[i]); });
  return o;
}
function ctaMinutosLoja_(ini, fim) { // minutos com a loja aberta (7h–19h30) entre duas datas
  var t = 0, d = vsDia_(ini);
  for (var k = 0; k < 4 && d <= fim; k++) {
    var a = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 7, 0), b = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 19, 30);
    var x = Math.max(a.getTime(), ini.getTime()), y = Math.min(b.getTime(), fim.getTime()); if (y > x) t += (y - x) / 60000;
    d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
  }
  return t;
}
function ctaInfo(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio pede contagem' };
  var r = v9Prod_(p); if (!r.ok) return r;
  var s = v9SaldoSku_(r.sku);
  var quem = v9ComAcesso_('CONTAGEM_ESTOQUE'); ['Ney', 'Regiane'].forEach(function (x) { if (quem.indexOf(x) < 0) quem.push(x); });
  if (r.rep) r.rep.split(/\s*\/\s*/).forEach(function (x) { x = x.trim(); var i = quem.indexOf(x); if (i >= 0) quem.splice(i, 1); if (x) quem.unshift(x); });
  r.saldo = s.qtd; r.saldoEm = Utilities.formatDate(s.em, Session.getScriptTimeZone(), 'dd/MM HH:mm'); r.pessoas = quem; r.ok = true;
  return r;
}
// ---------- v119.7 ACHAR NA LOJA (pedido do Celio, 14/11/2026) ----------
// Produto parado (produto sem venda) que o app ia sugerir rebaixar: antes, a repositora da seção precisa ACHAR o produto na loja.
// ACHOU -> validade + gôndola + depósito -> conta de venda até vencer -> se sobra, sugestão no A DECIDIR da REBAIXAS (OBS com achado x sistema).
// NÃO ACHOU -> não vira rebaixa: vai para a Contagem de estoque (contado 0, decidir ajuste).
var ABA_ACH = 'ACHAR NA LOJA';
var CAB_ACH = ['ID', 'CRIADO', 'SKU', 'PRODUTO', 'LOCAL', 'QUEM', 'TIPO', 'DIAS', 'MEDIA_DIA', 'SALDO', 'VALIDADES', 'STATUS', 'RESPONDEU', 'EM', 'VALIDADE', 'GONDOLA', 'DEPOSITO', 'TOTAL', 'SOBRA', 'DESTINO'];
function acharCriar_(sug, novos) {
  var sh = v9Aba_(ABA_ACH, CAB_ACH), lim = Date.now() - 15 * 86400000, ja = {}, novosL = [], lotes = null, porSku = {}, push = {};
  v9Linhas_(sh, CAB_ACH.length).forEach(function (l) { var s = skuN_(l[2]); if (String(l[11]) === 'PENDENTE' || psvT_(String(l[13] || l[1])) >= lim) ja[s] = 1; });
  novos.forEach(function (r) { porSku[skuN_(r[1])] = r; });
  sug.forEach(function (o) {
    var s = skuN_(o.sku); if (!s || ja[s]) return; ja[s] = 1;
    var r = porSku[s] || [], quem = String(r[5] || '').trim() || 'Ney';
    try { lotes = lotes || vfLotes_(); } catch (e) { lotes = {}; }
    var vs = ((lotes[s] || {}).lotes || []).map(function (x) { return v9Br_(x.v) || String(x.v); }).filter(String);
    if (o.validade && vs.indexOf(o.validade) < 0) vs.push(o.validade);
    vs.sort(function (a, b) { return diaNum_(a) - diaNum_(b); });
    novosL.push([Utilities.getUuid(), agoraBr_(), s, String(o.desc || r[2] || ''), String(r[4] || ''), quem, String(r[6] || ''), String(r[7] || ''), String(r[8] || ''), String(r[9] || ''), JSON.stringify(vs), 'PENDENTE', '', '', '', '', '', '', '', '']);
    push[quem] = (push[quem] || 0) + 1;
  });
  if (novosL.length) sh.getRange(sh.getLastRow() + 1, 1, novosL.length, CAB_ACH.length).setNumberFormat('@').setValues(novosL);
  Object.keys(push).forEach(function (q) { v9Push_(q, '🔎 Achar na loja', push[q] + (push[q] === 1 ? ' produto parado' : ' produtos parados') + ' para você achar na loja (Validades › Achar).', 'validades'); });
  return novosL.length;
}
function acharLista(p) {
  var a = v9Dono_(p, 'VALIDADES'); if (a.erro) return a.erro;
  var tudo = verTudo_(a.d), hoje = v9Hoje_(), out = [];
  v9Linhas_(v9Aba_(ABA_ACH, CAB_ACH), CAB_ACH.length).forEach(function (l) {
    var st = String(l[11]), dono = String(l[5] || '').trim() || 'Ney';
    if (!tudo && !tfMeu_(dono, a.d)) return;
    if (st !== 'PENDENTE' && String(l[13]).indexOf(hoje) !== 0) return; // respondidos só no dia
    var vs = []; try { vs = JSON.parse(String(l[10]) || '[]'); } catch (e) {}
    out.push({ id: String(l[0]), sku: String(l[2]), desc: String(l[3]), local: String(l[4]), quem: dono, dias: String(l[7]), validades: vs, status: st, validade: String(l[14]), total: String(l[17]) });
  });
  out.sort(function (x, y) { return (x.status === 'PENDENTE' ? 0 : 1) - (y.status === 'PENDENTE' ? 0 : 1); });
  return { ok: true, itens: out, tudo: tudo };
}
function acharResp(p) {
  var a = v9Dono_(p, 'VALIDADES'); if (a.erro) return a.erro;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_ACH, CAB_ACH), r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'item não encontrado' };
    var l = sh.getRange(r, 1, 1, CAB_ACH.length).getValues()[0]; if (String(l[11]) !== 'PENDENTE') return { ok: false, erro: 'esse produto já foi respondido' };
    var sku = skuN_(l[2]), desc = String(l[3]), sis = Number(String(l[9]).replace(',', '.')) || 0, media = Number(String(l[8]).replace(',', '.')) || 0, hora = agoraBr_(), dias = String(l[7] || '');
    if (String(p.achou) !== 'S') {
      acharContagemZero_(sku, desc, a.d, 'NÃO ACHADO NA LOJA por ' + a.d + (dias ? ' (sem venda ' + dias + ' dias)' : ''));
      sh.getRange(r, 12, 1, 9).setNumberFormat('@').setValues([['NAO ACHOU', a.d, hora, '', '0', '0', '0', '', 'CONTAGEM (Celio decide o ajuste)']]);
      return { ok: true, destino: 'CONTAGEM' };
    }
    var val = v9Br_(p.validade); if (!val) return { ok: false, erro: 'escolha a validade' };
    var g = v9Num_(p.gondola), d = v9Num_(p.deposito);
    if (g == null || isNaN(g) || g < 0) return { ok: false, erro: 'quantas tem na gôndola? (0 se não tem)' };
    if (d == null || isNaN(d) || d < 0) return { ok: false, erro: 'quantas tem no depósito? (0 se não tem)' };
    var tot = Math.round((g + d) * 1000) / 1000; if (!(tot > 0)) return { ok: false, erro: 'se não tem nenhum, toque em NÃO ACHEI' };
    var dd = Math.max(0, diaNum_(val) - diaNum_(v9Hoje_())), sobra = Math.ceil(tot - media * dd), dest = 'VENDE ATÉ VENCER';
    if (sobra > 0) {
      var obs = obsQtd_(tot) + ' ACHADAS NA LOJA (gondola ' + psvN_(g) + ' + deposito ' + psvN_(d) + ') por ' + a.d + ' em ' + v9Hoje_().slice(0, 5) + ' - sistema ' + psvN_(sis) + ' - SKU ' + sku + (dias ? ' (sem venda ' + dias + ' dias)' : '') + ' - deve sobrar ' + sobra;
      rebSugerirLote_([{ sku: sku, desc: desc, validade: val, obs: obs }]); dest = 'REBAIXA SUGERIDA';
    }
    sh.getRange(r, 12, 1, 9).setNumberFormat('@').setValues([['ACHOU', a.d, hora, val, psvN_(g), psvN_(d), psvN_(tot), String(sobra), dest]]);
    return { ok: true, destino: dest };
  } finally { lock.releaseLock(); }
}
function acharContagemZero_(sku, desc, contPor, motivo) { // registro na CONTAGEM ESTOQUE já contado (0 + 0), para o Celio decidir o ajuste
  var sh = v9Aba_(ABA_CTA, CAB_CTA), l = [], s = v9SaldoSku_(sku), tz = Session.getScriptTimeZone();
  for (var i = 0; i < CAB_CTA.length; i++) l.push('');
  var sis = s.qtd == null ? '' : s.qtd, dif = sis === '' ? '' : Math.round(-sis * 10) / 10, custo = s.custo == null ? '' : s.custo, difRs = dif === '' || custo === '' ? '' : Math.round(dif * custo * 100) / 100;
  l[0] = Utilities.getUuid(); l[1] = agoraBr_(); l[2] = 'Celio'; l[3] = sku; l[4] = desc; l[5] = contPor; l[6] = motivo; l[7] = dif !== '' && Math.abs(dif) < 0.5 ? 'BATEU - SEM AJUSTE' : 'CONTADA - DECIDIR';
  l[8] = '0'; l[9] = '0'; l[10] = '0'; l[11] = contPor; l[12] = agoraBr_(); l[13] = sis === '' ? '' : String(sis); l[14] = s.em ? Utilities.formatDate(s.em, tz, 'dd/MM/yyyy HH:mm') : '';
  l[15] = '0'; l[16] = String(sis); l[17] = String(dif); l[18] = String(custo); l[19] = String(difRs); l[24] = 'ACHAR NA LOJA';
  sh.appendRow(l);
  if (l[7] !== 'BATEU - SEM AJUSTE') v9Push_('Celio', '🔎 Não achado na loja', desc + ': ' + contPor + ' não achou' + (sis !== '' ? ' (sistema ' + String(sis).replace('.', ',') + ')' : ''), 'gcontagem');
}
function ctaPedir_(dono, sku, desc, para, motivo, origem) {
  var sh = v9Aba_(ABA_CTA, CAB_CTA), id = Utilities.getUuid(), l = [];
  for (var i = 0; i < CAB_CTA.length; i++) l.push('');
  l[0] = id; l[1] = agoraBr_(); l[2] = dono; l[3] = skuN_(sku); l[4] = desc; l[5] = para; l[6] = motivo || ''; l[7] = 'AGUARDANDO CONTAGEM'; l[24] = origem || 'APP';
  sh.appendRow(l);
  v9Push_(para, '📦 Contagem de estoque', 'O Celio pediu para contar: ' + desc + '. Conte área de venda + estoque.', 'contagem');
  return id;
}
function ctaPedir(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio pede contagem' };
  var para = String(p.para || '').trim(); if (!para) return { ok: false, erro: 'escolha quem vai contar' };
  var r = v9Prod_(p); if (!r.ok) return r;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ja = v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), CAB_CTA.length).some(function (l) { return skuN_(l[3]) === r.sku && String(l[7]) === 'AGUARDANDO CONTAGEM'; });
    if (ja) return { ok: false, erro: 'esse produto já está aguardando contagem' };
    return { ok: true, id: ctaPedir_(a.d, r.sku, r.desc, para, String(p.motivo || '').trim(), 'APP') };
  } finally { lock.releaseLock(); }
}
function ctaMinhas(p) { // quem conta: só os pedidos para ela, sem saldo do sistema
  var a = v9Dono_(p, 'CONTAGEM_ESTOQUE'); if (a.erro) return a.erro;
  var hoje = v9Hoje_(), pend = [], env = [], cad = mapaReposicao_();
  v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), CAB_CTA.length).forEach(function (l) {
    var o = ctaObj_(l); if (!tfMeu_(o.para, a.d)) return;
    var c = cadSku_(cad, o.sku, o.desc) || {};
    if (o.status === 'AGUARDANDO CONTAGEM') pend.push({ id: o.id, sku: o.sku, desc: o.desc, pedEm: o.pedEm, motivo: o.motivo, local: c.local || '', recont: /recontagem/i.test(o.motivo) });
    else if (o.contPor === a.d && o.contEm.indexOf(hoje) === 0) env.push({ id: o.id, desc: o.desc, total: o.total, em: o.contEm.slice(11) });
  });
  return { ok: true, pend: pend, enviados: env.reverse() };
}
function ctaContar(p) {
  var a = v9Dono_(p, 'CONTAGEM_ESTOQUE'); if (a.erro) return a.erro;
  var area = v9Num_(p.area), est = v9Num_(p.estoque);
  if (area == null || isNaN(area) || area < 0) return { ok: false, erro: 'informe a área de venda (0 se não tem)' };
  if (est == null || isNaN(est) || est < 0) return { ok: false, erro: 'informe o estoque (0 se não tem)' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_CTA, CAB_CTA), r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'pedido não encontrado' };
    var o = ctaObj_(sh.getRange(r, 1, 1, CAB_CTA.length).getValues()[0]);
    if (o.status !== 'AGUARDANDO CONTAGEM') return { ok: false, erro: 'essa contagem já foi enviada' };
    if (!tfMeu_(o.para, a.d) && !ehCelio_(a.d)) return { ok: false, erro: 'essa contagem é de outra pessoa' };
    var total = Math.round((area + est) * 1000) / 1000, agora = new Date(), s = v9SaldoSku_(o.sku), med = cvMedias_()[o.sku] || 0;
    var venda = s.em && s.em < agora ? Math.round(med * ctaMinutosLoja_(s.em, agora) / 750 * 10) / 10 : 0;
    var sis = s.qtd == null ? '' : Math.round((s.qtd - venda) * 10) / 10;
    var dif = sis === '' ? '' : Math.round((total - sis) * 10) / 10, custo = s.custo == null ? '' : s.custo;
    var difRs = dif === '' || custo === '' ? '' : Math.round(dif * custo * 100) / 100;
    var st = dif !== '' && Math.abs(dif) < 0.5 ? 'BATEU - SEM AJUSTE' : 'CONTADA - DECIDIR';
    var tz = Session.getScriptTimeZone();
    sh.getRange(r, 8, 1, 13).setValues([[st, String(area), String(est), String(total), a.d, agoraBr_(), s.qtd == null ? '' : String(s.qtd), s.em ? Utilities.formatDate(s.em, tz, 'dd/MM/yyyy HH:mm') : '', String(venda), String(sis), String(dif), String(custo), String(difRs)]]);
    if (st !== 'BATEU - SEM AJUSTE') v9Push_('Celio', '📦 Contagem feita — decidir', o.desc + ': contado ' + String(total).replace('.', ',') + (dif !== '' ? ' (diferença ' + (dif > 0 ? '+' : '') + String(dif).replace('.', ',') + ')' : ''), 'gcontagem');
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function ctaGestao(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var nfs = {}; try { nfs = ctaAtualizar_() || {}; } catch (e) {} // v130: saldo sempre do dia + aviso de NF
  var lim = diaNum_(v9Hoje_()) - 30, mes = v9Hoje_().slice(3), aguard = [], decidir = [], hist = [], totMes = 0;
  v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), CAB_CTA.length).forEach(function (l) {
    var o = ctaObj_(l);
    if (o.status === 'AGUARDANDO CONTAGEM') aguard.push(o);
    else if (o.status === 'CONTADA - DECIDIR') { o.nfs = nfs[o.id] || []; decidir.push(o); }
    if (diaNum_(o.pedEm.slice(0, 10)) >= lim) hist.push(o);
    if ((o.status === 'AJUSTADO' || /^AUTORIZADO/.test(o.status)) && o.decEm.slice(3, 10) === mes && o.difRs !== '') totMes += Number(o.difRs) || 0;
  });
  return { ok: true, aguard: aguard.reverse(), decidir: decidir.reverse(), hist: hist.reverse(), totMes: Math.round(totMes * 100) / 100, pessoas: v9ComAcesso_('CONTAGEM_ESTOQUE') };
}
function ctaDecidir(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio decide' };
  var dec = String(p.decisao || '').toUpperCase();
  if (['SAIDA MA', 'BAIXA NF', 'ENTRADA MA', 'NAO AJUSTAR', 'RECONTAGEM'].indexOf(dec) < 0) return { ok: false, erro: 'decisão inválida' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_CTA, CAB_CTA), r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'contagem não encontrada' };
    var o = ctaObj_(sh.getRange(r, 1, 1, CAB_CTA.length).getValues()[0]);
    if (['CONTADA - DECIDIR', 'BATEU - SEM AJUSTE'].indexOf(o.status) < 0) return { ok: false, erro: 'essa contagem já foi decidida' };
    var dif = Number(o.dif);
    if ((dec === 'SAIDA MA' || dec === 'BAIXA NF') && !(dif < 0)) return { ok: false, erro: 'não faltou produto — saída/baixa não se aplica' };
    if (dec === 'ENTRADA MA' && !(dif > 0)) return { ok: false, erro: 'entrada só quando sobra produto' };
    var st = dec === 'NAO AJUSTAR' ? 'NAO AJUSTADO' : (dec === 'RECONTAGEM' ? 'RECONTAGEM' : 'AUTORIZADO - A LANCAR');
    sh.getRange(r, 8).setValue(st); sh.getRange(r, 21, 1, 2).setValues([[dec, agoraBr_()]]);
    if (dec === 'RECONTAGEM') ctaPedir_(a.d, o.sku, o.desc, String(p.para || o.para), 'recontagem' + (o.total !== '' ? ' (1ª contagem: ' + o.total + ')' : ''), o.origem || 'APP');
    if (!p._viaDec) { try { decGrava_('CONTAGEM', 'CTA|' + o.id, 'DECIDIDA', '', dec, { id: o.id, prod: o.desc }, a.d, ''); } catch (e) {} } // v127: aparece também nas Decisões (Respondidas)
    return { ok: true, status: st };
  } finally { lock.releaseLock(); }
}
// v127 (17/11/2026): Corrigir resposta — a nova resposta substitui a anterior e fica registrado que foi corrigida
function decCorrPoe_(chave, o) { var pr = PropertiesService.getScriptProperties(), m = JSON.parse(pr.getProperty('DEC_CORR') || '{}'); o.em = agoraBr_(); m[chave] = o; pr.setProperty('DEC_CORR', JSON.stringify(m)); }
function decCorrPega_(chave) { var pr = PropertiesService.getScriptProperties(), m = JSON.parse(pr.getProperty('DEC_CORR') || '{}'), o = m[chave]; if (!o) return null; delete m[chave]; pr.setProperty('DEC_CORR', JSON.stringify(m)); return o; }
function ctaCorrigir(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var dec = String(p.decisao || '').toUpperCase(), sh = v9Aba_(ABA_CTA, CAB_CTA), o = null, r = -1, lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'contagem não encontrada' };
    o = ctaObj_(sh.getRange(r, 1, 1, CAB_CTA.length).getValues()[0]);
    if (o.status === 'AJUSTADO') return { ok: false, erro: 'já foi ajustada no Plenus — não dá para corrigir pelo app' };
    if (['AUTORIZADO - A LANCAR', 'NAO AJUSTADO'].indexOf(o.status) < 0) return { ok: false, erro: 'não dá para corrigir: ' + o.status };
    if (o.status === 'AUTORIZADO - A LANCAR' && decEmExec_('CTA_LANCAR_EM')) return { ok: false, erro: '🔄 sendo lançada no Plenus agora — não dá para corrigir' };
    if (dec === o.decisao) return { ok: false, erro: 'é a mesma resposta' };
    sh.getRange(r, 8).setValue('CONTADA - DECIDIR');
  } finally { lock.releaseLock(); }
  decCorrPoe_('CTA|' + o.id, { de: o.decisao, deEm: o.decEm });
  var res = ctaDecidir({ token: p.token, quem: p.quem, id: o.id, decisao: dec, para: p.para });
  if (!res.ok) { sh.getRange(r, 8).setValue(o.status); decCorrPega_('CTA|' + o.id); return res; }
  sh.getRange(r, 26).setValue((o.obs ? o.obs + ' | ' : '') + 'CORRIGIDA ' + agoraBr_() + ': ' + o.decisao + ' (' + o.decEm + ') -> ' + dec);
  return { ok: true, status: res.status, msg: 'Resposta corrigida' };
}
// v130 (Celio 19/11/2026): contagem a decidir sempre comparada com o saldo do DIA. O saldo da contagem era o do arquivo da manhã e ficava congelado;
// se a NF entrava depois do arquivo (ex.: Amstel 17/11: arquivo -14, NF de 60, contado 60 → dif +74,6 falsa), a decisão saía errada.
// Quando chega um saldo atual.txt mais novo que a contagem: SALDO SISTEMA (na hora da contagem) = saldo de hoje + vendido desde a contagem
// (vendas diarias.xlsx; dia da contagem proporcional às horas de loja que faltavam; dia sem venda gravada → média). Bateu → BATEU - SEM AJUSTE.
function ctaBrDt_(s) { var m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/); return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) : null; }
function ctaAtualizar_() {
  var sh = v9Aba_(ABA_CTA, CAB_CTA), pend = [];
  v9Linhas_(sh, CAB_CTA.length).forEach(function (l, i) { var o = ctaObj_(l); if (o.status === 'CONTADA - DECIDIR') { o._r = i + 2; pend.push(o); } });
  if (!pend.length) return {};
  var nfs = {}; try { nfs = ctaNfs_(pend); } catch (e) {}
  var f = DriveApp.getFileById(ID_SALDO_TXT), fEm = f.getLastUpdated(), tz = Session.getScriptTimeZone(), fTx = Utilities.formatDate(fEm, tz, 'dd/MM/yyyy HH:mm');
  var precisa = pend.filter(function (o) { var c = ctaBrDt_(o.contEm), ar = ctaBrDt_(o.arqEm); return c && fEm > c && (!ar || fEm.getTime() - ar.getTime() > 120000) && String(o.arqEm).slice(0, 16) !== fTx; });
  if (!precisa.length) return nfs;
  var sal = v9Saldos_().itens, vds = { s: {} }, med = cvMedias_(); try { vds = vfVendas_(); } catch (e) {}
  var fN = diaNum_(fTx.slice(0, 10)), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    precisa.forEach(function (o) {
      var sk = skuN_(o.sku), s = sal[sk]; if (!s || s.q == null) return;
      var c = ctaBrDt_(o.contEm), cN = diaNum_(String(o.contEm).slice(0, 10)), v = vds.s[sk] || {}, m = Number(med[sk]) || 0, venda = 0;
      for (var d = cN; d < fN; d++) {
        var q = v[d] != null ? v[d] : m;
        if (d === cN) q = q * ctaMinutosLoja_(c, new Date(c.getFullYear(), c.getMonth(), c.getDate(), 23, 59)) / 750;
        venda += q;
      }
      venda = Math.round(venda * 10) / 10;
      var sis = Math.round((s.q + venda) * 10) / 10, dif = Math.round((Number(o.total) - sis) * 10) / 10, custo = s.custo == null ? o.custo : s.custo;
      var difRs = custo === '' || custo == null ? '' : Math.round(dif * Number(custo) * 100) / 100;
      var obs = (String(o.obs || '').replace(/\s*\|?\s*v130:[^|]*/g, '').trim() + ' | v130: saldo de hoje ' + String(s.q).replace('.', ',') + ' (arquivo ' + fTx + ') + vendido desde a contagem ' + String(venda).replace('.', ',') + ' (antes: sistema ' + o.saldoSis + ', dif ' + o.dif + ')').replace(/^\s*\|\s*/, '');
      sh.getRange(o._r, 8).setValue(Math.abs(dif) < 0.5 ? 'BATEU - SEM AJUSTE' : 'CONTADA - DECIDIR');
      sh.getRange(o._r, 14, 1, 7).setValues([[String(s.q), fTx, String(venda), String(sis), String(dif), String(custo), String(difRs)]]);
      sh.getRange(o._r, 26).setValue(obs);
    });
  } finally { lock.releaseLock(); }
  return nfs;
}
// v130: NF do produto que chegou/entrou de 3 dias antes da contagem em diante (aviso no cartão)
function ctaNfs_(pend) {
  var out = {}, sk = {}, ck = 'CTA_NF_' + pend.map(function (o) { return o.id.slice(0, 8); }).join('').slice(0, 200), cache = CacheService.getScriptCache(), hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  pend.forEach(function (o) { sk[skuN_(o.sku)] = 1; });
  var ss = SpreadsheetApp.getActive(), ci = ss.getSheetByName(ABA_CIT), cn = ss.getSheetByName(ABA_CNF); if (!ci || !cn) return out;
  var nci = ci.getLastRow(), nn = cn.getLastRow(), por = {}; if (nci < 2 || nn < 2) return out;
  ci.getRange(2, 1, nci - 1, 7).getValues().forEach(function (l) { var s = skuN_(l[2]); if (!sk[s]) return; (por[String(l[0])] = por[String(l[0])] || []).push({ s: s, q: l[6] }); });
  var cab = {}; cn.getRange(2, 1, nn - 1, 5).getDisplayValues().forEach(function (l) { if (por[l[0]]) cab[l[0]] = l; });
  pend.forEach(function (o) {
    var s = skuN_(o.sku), cN = diaNum_(String(o.contEm).slice(0, 10)), L = [];
    Object.keys(cab).forEach(function (id) {
      var c = cab[id], ent = String(c[4]).trim(), eN = diaNum_(ent.slice(0, 10)), base = isNaN(eN) ? diaNum_(String(c[3]).slice(0, 10)) : eN; if (isNaN(base) || base < cN - 3) return;
      por[id].forEach(function (it) { if (it.s === s) L.push((ent ? '📥 Entrou NF ' : '🚚 Chegou NF ') + c[1] + ' ' + c[2] + ' (' + String(it.q).replace('.', ',') + ' na nota)' + (ent ? ' — entrada no Plenus ' + ent.slice(0, 5) : ' — ainda SEM entrada no Plenus')); });
    });
    if (L.length) out[o.id] = L;
  });
  try { cache.put(ck, JSON.stringify(out), 600); } catch (e) {}
  return out;
}
function ctaLancar(p) { // tarefa noturna do Claude: LISTAR os autorizados / FEITO (id, lancamento)
  if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' };
  var sh = v9Aba_(ABA_CTA, CAB_CTA);
  if (String(p.op || '') === 'FEITO') {
    var r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'não encontrada' };
    sh.getRange(r, 8).setValue('AJUSTADO'); sh.getRange(r, 23, 1, 2).setValues([[String(p.lancamento || ''), agoraBr_()]]);
    return { ok: true };
  }
  PropertiesService.getScriptProperties().setProperty('CTA_LANCAR_EM', String(Date.now())); // v115
  return { ok: true, itens: v9Linhas_(sh, CAB_CTA.length).map(ctaObj_).filter(function (o) { return o.status === 'AUTORIZADO - A LANCAR'; }) };
}

// ---------- 👀 VALIDADE NA FRENTE + confrontação diária saldo × lotes ----------
var ABA_LOTES = 'LOTES CONFRONTO', ABA_LVEND = 'LOTES VENDIDOS', ABA_VNF = 'VALIDADE NA FRENTE';
var CAB_LOTES = ['SKU', 'PRODUTO', 'VALIDADE', 'DT CHEGOU', 'NF', 'CHEGOU', 'RESTANTE', 'SITUACAO', 'SECAO', 'SALDO PRODUTO', 'ATUALIZADO', 'SALDO NA CHEGADA', 'ZEROU EM', 'OBS'];
var CAB_VNF = ['DIA', 'PARA', 'SKU', 'PRODUTO', 'LOCAL', 'VALIDADES', 'CERTA', 'QTD CERTA', 'STATUS', 'RESPOSTA', 'RESULTADO', 'ACHOU', 'QTD ACHADA', 'QUEM', 'EM', 'ID', 'LOTE SUMIDO', 'DECISAO SUMIDO'];
var VF_POR_DIA = 5;
function vfData_(v, tz) { if (v instanceof Date) { var p = Utilities.formatDate(v, tz, 'yyyy-MM-dd').split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); } var b = v9Br_(v); return b ? vsDeBr_(b) : null; }
function vfVendas_() { // v124: vendas diarias.xlsx -> { s: {sku: {dia: qtd}}, min: primeiro dia }
  var it = DriveApp.getFolderById(ID_PASTA_BD).getFilesByName('vendas diarias.xlsx'), m = {}, mn = null;
  if (!it.hasNext()) return { s: m, min: null };
  lerXlsx_(it.next().getId()).slice(1).forEach(function (r) {
    var s = skuN_(r.B), a = String(r.A || '').trim(), d = /^\d+(\.\d+)?$/.test(a) ? Math.floor(Number(a)) - 25569 : diaNum_(a.slice(0, 10)); if (!s || isNaN(d)) return;
    var q = Number(String(r.D || '').replace(',', '.')) || 0, g = m[s] = m[s] || {}; g[d] = (g[d] || 0) + q; if (mn == null || d < mn) mn = d;
  });
  return { s: m, min: mn };
}
// v124 (Celio 16/11): sobra por lote pela ordem de chegada (FIFO). Situação: COM SALDO = provável sobra; VENDIDO = provavelmente vendido (conferência rápida);
// ZEROU = o saldo chegou a zero por VENDA antes do lote seguinte chegar (lote acabou de verdade, sai da lista das meninas).
var VF_PK = /^(CX|CXA1|FD|FDO1|FARDO|DP|DIS|EXB|CT|PC)$/;
function vfConfrontar() {
  var tz = Session.getScriptTimeZone(), hoje = vsDia_(new Date()), lim = vsMais_(hoje, -30), hojeN = diaNum_(vsBr_(hoje));
  var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone(), pr = ssV.getSheetByName('PRODUTOS'), np = pr.getLastRow();
  var vals = np >= 2 ? pr.getRange(2, 1, np - 1, 8).getValues() : [], skus = vsSkuPorNome_(), sal = v9Saldos_().itens, vds = { s: {}, min: null };
  try { vds = vfVendas_(); } catch (e) {}
  var sumidos = {}; // lote que a repositora não achou (validade na frente) = vendido
  v9Linhas_(v9Aba_(ABA_VNF, CAB_VNF), CAB_VNF.length).forEach(function (l) { if (String(l[16])) sumidos[skuN_(l[2]) + '|' + String(l[16])] = 1; });
  var por = {};
  vals.forEach(function (l) {
    var prod = String(l[3] || '').replace(/[¬]/g, '').trim(), val = vfData_(l[1], tzV); if (!prod || !val || val < lim) return;
    var sk = skus[normDesc_(prod)]; if (!sk || /QUILO|KG|LITRO/i.test(sk.un)) return;
    var q = v9Num_(l[5]); if (!(q > 0)) return;
    var s = skuN_(sk.sku), dc = vfData_(l[6], tzV), g = por[s] = por[s] || { prod: prod, sec: String(l[2] || ''), rows: [] };
    g.rows.push({ val: val, vb: vsBr_(val), dc: dc, dn: dc ? diaNum_(vsBr_(dc)) : null, q: q, un: String(l[4] || '').trim().toUpperCase(), nf: String(l[0] || '') });
  });
  var out = [], vend = [], agora = agoraBr_(), lv0 = vsMais_(hoje, -10), lv1 = vsMais_(hoje, 70), nz = 0;
  Object.keys(por).forEach(function (s) {
    var g = por[s], it = sal[s]; if (!it || it.q == null) return;
    // quantidade lançada na embalagem (CX 1, FD 2...): vira unidades pela mediana dos outros lotes do produto
    var base = g.rows.filter(function (x) { return !VF_PK.test(x.un) && x.q >= 6; }).map(function (x) { return x.q; }).sort(function (a, b) { return a - b; });
    var med = base.length ? (base.length % 2 ? base[(base.length - 1) / 2] : (base[base.length / 2 - 1] + base[base.length / 2]) / 2) : 0;
    g.rows.forEach(function (x) { if (VF_PK.test(x.un) && x.q <= 5 && med) { x.q = x.q * med; x.est = 1; } });
    var lotes = {}, ent = {};
    g.rows.forEach(function (x) {
      var lt = lotes[x.vb] = lotes[x.vb] || { val: x.val, vb: x.vb, q: 0, dc: null, dn: null, nf: '', est: 0 };
      lt.q += x.q; if (x.dc && (!lt.dc || x.dc > lt.dc)) { lt.dc = x.dc; lt.dn = x.dn; } if (x.nf) lt.nf = x.nf; if (x.est) lt.est = 1;
      if (x.dn != null) ent[x.dn] = (ent[x.dn] || 0) + x.q;
    });
    var ls = Object.keys(lotes).map(function (k) { return lotes[k]; });
    ls.sort(function (a, b) { var x = (b.dc || b.val) - (a.dc || a.val); return x || (b.val - a.val); }); // entrada mais nova primeiro
    // saldo no fim de cada dia, de hoje para trás: saldo(t-1) = saldo(t) + vendas(t) - entradas(t)
    var vs = vds.s[s] || {}, S = {}, t0 = hojeN;
    ls.forEach(function (lt) { if (lt.dn != null && lt.dn < t0) t0 = lt.dn; });
    if (vds.min != null) { var cur = it.q; S[hojeN] = cur; for (var t = hojeN; t > Math.max(t0, vds.min) - 1; t--) { cur = cur + (vs[t] || 0) - (ent[t] || 0); S[t - 1] = cur; } }
    var resta = Math.max(0, it.q);
    ls.forEach(function (lt) {
      var r = sumidos[s + '|' + lt.vb] ? 0 : Math.min(lt.q, resta); resta -= r;
      var sit = r > 0 ? 'COM SALDO' : 'VENDIDO', zerou = '', naCheg = '';
      if (lt.dn != null && vds.min != null) {
        if (S[lt.dn - 1] != null) naCheg = String(Math.round(S[lt.dn - 1] * 1000) / 1000);
        var nd = null; ls.forEach(function (o) { if (o.dn != null && o.dn > lt.dn && (nd == null || o.dn < nd)) nd = o.dn; });
        if (nd != null) for (var d = Math.max(lt.dn, vds.min); d < nd && d <= hojeN; d++) if (S[d] != null && S[d] <= 0) { zerou = vsBr_(new Date(d * 86400000 + 43200000)); break; }
      }
      if (zerou) { sit = 'ZEROU'; nz++; }
      out.push(["'" + s, g.prod, lt.vb, lt.dc ? vsBr_(lt.dc) : '', lt.nf, String(Math.round(lt.q * 1000) / 1000), String(Math.round(r * 1000) / 1000), sit, g.sec, String(it.q), agora, naCheg, zerou, lt.est ? 'QTD ESTIMADA (embalagem x ' + med + ')' : '']);
      if (sit === 'ZEROU' && lt.val >= lv0 && lt.val <= lv1) vend.push(["'" + s, normDesc_(g.prod), lt.vb]);
    });
    if (resta > 0.5) out.push(["'" + s, g.prod, '', '', '', '', String(Math.round(resta * 1000) / 1000), 'SALDO SEM LOTE', g.sec, String(it.q), agora, '', '', '']);
  });
  var sh = v9Aba_(ABA_LOTES, CAB_LOTES); sh.getRange(1, 1, 1, CAB_LOTES.length).setValues([CAB_LOTES]).setFontWeight('bold'); sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), CAB_LOTES.length).clearContent();
  if (out.length) { if (sh.getMaxRows() < out.length + 1) sh.insertRowsAfter(sh.getMaxRows(), out.length + 1 - sh.getMaxRows()); sh.getRange(2, 1, out.length, CAB_LOTES.length).setNumberFormat('@').setValues(out); }
  var sv = v9Aba_(ABA_LVEND, ['SKU', 'CHAVE PRODUTO', 'VALIDADE']); sv.getRange(2, 1, Math.max(1, sv.getMaxRows() - 1), 3).clearContent();
  if (vend.length) { if (sv.getMaxRows() < vend.length + 1) sv.insertRowsAfter(sv.getMaxRows(), vend.length + 1 - sv.getMaxRows()); sv.getRange(2, 1, vend.length, 3).setNumberFormat('@').setValues(vend); }
  PropertiesService.getScriptProperties().setProperty('VF_DIA', v9Hoje_());
  CacheService.getScriptCache().remove('VF_VEND');
  return { ok: true, produtos: Object.keys(por).length, lotes: out.length, zerou: nz, zerouNaJanela: vend.length, vendasDesde: vds.min != null ? vsBr_(new Date(vds.min * 86400000 + 43200000)) : '' };
}
function vfDiario_() { if (PropertiesService.getScriptProperties().getProperty('VF_DIA') !== v9Hoje_()) return vfConfrontar(); return null; }
var VF_VEND_MEMO = null;
function vfVend_(sku, desc, val) { // lote já vendido pela confrontação?
  if (!VF_VEND_MEMO) {
    VF_VEND_MEMO = {};
    try { v9Linhas_(v9Aba_(ABA_LVEND, ['SKU', 'CHAVE PRODUTO', 'VALIDADE']), 3).forEach(function (l) { var s = skuN_(l[0]); if (s) VF_VEND_MEMO['#' + s + '|' + l[2]] = 1; if (l[1]) VF_VEND_MEMO[String(l[1]) + '|' + l[2]] = 1; }); } catch (e) {}
  }
  var vb = v9Br_(val); if (!vb) return false;
  var s = skuN_(sku); return !!((s && VF_VEND_MEMO['#' + s + '|' + vb]) || VF_VEND_MEMO[normDesc_(desc) + '|' + vb]);
}
var VF_SIT_MEMO = null;
function vfSit_(sku, val) { // v124: situação do lote na confrontação {sit, r}
  if (!VF_SIT_MEMO) { VF_SIT_MEMO = {}; try { v9Linhas_(v9Aba_(ABA_LOTES, CAB_LOTES), 8).forEach(function (l) { if (String(l[2])) VF_SIT_MEMO[skuN_(l[0]) + '|' + String(l[2])] = { sit: String(l[7]), r: Number(l[6]) || 0 }; }); } catch (e) {} }
  var vb = v9Br_(val); return vb ? (VF_SIT_MEMO[skuN_(sku) + '|' + vb] || null) : null;
}
function vfLotes_() { // lotes COM SALDO por SKU
  var m = {};
  v9Linhas_(v9Aba_(ABA_LOTES, CAB_LOTES), CAB_LOTES.length).forEach(function (l) {
    if (String(l[7]) !== 'COM SALDO') return; var s = skuN_(l[0]);
    var g = m[s] = m[s] || { sku: s, prod: String(l[1]), sec: String(l[8]), lotes: [] };
    g.lotes.push({ v: String(l[2]), q: Number(l[6]) || 0 });
  });
  Object.keys(m).forEach(function (s) { m[s].lotes.sort(function (a, b) { return diaNum_(a.v) - diaNum_(b.v); }); });
  return m;
}
function vfGerar_(quem) { // sorteia os produtos do dia da pessoa
  var hoje = v9Hoje_(), hn = diaNum_(hoje), sh = v9Aba_(ABA_VNF, CAB_VNF), vals = v9Linhas_(sh, CAB_VNF.length), recente = {};
  vals.forEach(function (l) { if (hn - diaNum_(String(l[0])) < 14) recente[skuN_(l[2])] = 1; });
  var lotes = vfLotes_(), cad = cvCad_(SpreadsheetApp.openById(ID_VALIDADE)), med = cvMedias_(), tudo = verTudo_(quem), cand = [];
  Object.keys(lotes).forEach(function (s) {
    var g = lotes[s]; if (recente[s] || !g.lotes.length) return;
    var c = cadSku_(cad, s, g.prod); if (!c) return;
    if (!tudo && !tfMeu_(c.rep, quem)) return;
    if (tudo && cvFora45_(c.secao)) return;
    var dias = diaNum_(g.lotes[0].v) - hn; if (dias < 0) return;
    var sc = (g.lotes.length >= 2 ? 100 : 0) + Math.max(0, 90 - dias) * 0.6 + Math.min(30, (med[s] || 0) * 3) + Math.random() * 25;
    cand.push({ s: s, g: g, c: c, sc: sc });
  });
  cand.sort(function (a, b) { return b.sc - a.sc; });
  var novos = cand.slice(0, VF_POR_DIA).map(function (x) {
    return [hoje, quem, x.s, x.g.prod, x.c.local || '', JSON.stringify(x.g.lotes.map(function (lt) { return lt.v; })), x.g.lotes[0].v, String(Math.round(x.g.lotes[0].q)), 'PENDENTE', '', '', '', '', '', '', Utilities.getUuid(), '', ''];
  });
  if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, CAB_VNF.length).setNumberFormat('@').setValues(novos);
  return novos.length;
}
function vfLista(p) {
  var a = v9Dono_(p, 'VALIDADES'); if (a.erro) return a.erro;
  try { vfDiario_(); } catch (e) {}
  var hoje = v9Hoje_(), sh = v9Aba_(ABA_VNF, CAB_VNF);
  var meus = function () { return v9Linhas_(sh, CAB_VNF.length).filter(function (l) { return String(l[0]) === hoje && String(l[1]) === a.d; }); };
  var ls = meus();
  if (!ls.length && PropertiesService.getScriptProperties().getProperty('VF_GER_' + a.d) !== hoje) {
    var lock = LockService.getScriptLock(); lock.waitLock(30000);
    try { ls = meus(); if (!ls.length) { vfGerar_(a.d); PropertiesService.getScriptProperties().setProperty('VF_GER_' + a.d, hoje); ls = meus(); } } finally { lock.releaseLock(); }
  }
  var itens = ls.map(function (l) {
    var vs = []; try { vs = JSON.parse(String(l[5]) || '[]'); } catch (e) {}
    var st = String(l[8]), o = { id: String(l[15]), sku: String(l[2]), desc: String(l[3]), local: String(l[4]), validades: vs, status: st, resposta: String(l[9]), resultado: String(l[10]) };
    if (st !== 'PENDENTE') { o.certa = String(l[6]); o.qtdCerta = String(l[7]); o.achou = String(l[11]); } // a certa só aparece depois de responder
    return o;
  });
  return { ok: true, itens: itens, feitos: itens.filter(function (i) { return i.status !== 'PENDENTE' && i.status !== 'AGUARDANDO ACHEI'; }).length, meta: VF_POR_DIA };
}
function vfResponder(p) {
  var a = v9Dono_(p, 'VALIDADES'); if (a.erro) return a.erro;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_VNF, CAB_VNF), r = v9Linha_(sh, 16, p.id); if (r < 0) return { ok: false, erro: 'item não encontrado' };
    var l = sh.getRange(r, 1, 1, CAB_VNF.length).getValues()[0];
    if (String(l[8]) !== 'PENDENTE') return { ok: false, erro: 'esse produto já foi respondido' };
    var resp = String(p.resposta || '').trim(), certa = String(l[6]), res, st = 'RESPONDIDO', vs = [];
    try { vs = JSON.parse(String(l[5]) || '[]'); } catch (e) {}
    if (resp === 'NAO TEM') res = 'NAO TEM NA GONDOLA';
    else if (/^OUTRA:/.test(resp)) { resp = v9Br_(resp.slice(6)); if (!resp) return { ok: false, erro: 'escolha a validade' }; res = vs.indexOf(resp) >= 0 ? (resp === certa ? 'CERTA' : 'MAIS NOVA') : 'OUTRA VALIDADE'; }
    else { resp = v9Br_(resp); if (vs.indexOf(resp) < 0) return { ok: false, erro: 'validade inválida' }; res = resp === certa ? 'CERTA' : 'MAIS NOVA'; }
    if (res === 'MAIS NOVA') st = 'AGUARDANDO ACHEI';
    sh.getRange(r, 9, 1, 3).setValues([[st, resp, res]]); sh.getRange(r, 14, 1, 2).setValues([[a.d, agoraBr_()]]);
    return { ok: true, resultado: res, certa: certa, qtdCerta: String(l[7]) };
  } finally { lock.releaseLock(); }
}
function vfAchei(p) {
  var a = v9Dono_(p, 'VALIDADES'); if (a.erro) return a.erro;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_VNF, CAB_VNF), r = v9Linha_(sh, 16, p.id); if (r < 0) return { ok: false, erro: 'item não encontrado' };
    var l = sh.getRange(r, 1, 1, CAB_VNF.length).getValues()[0];
    if (String(l[8]) !== 'AGUARDANDO ACHEI') return { ok: false, erro: 'esse produto já foi resolvido' };
    var achou = String(p.achou) === '1', q = v9Num_(p.qtd);
    if (achou && (q == null || isNaN(q) || q <= 0)) return { ok: false, erro: 'quantas unidades você achou?' };
    if (achou) {
      sh.getRange(r, 9).setValue('TROCADO'); sh.getRange(r, 12, 1, 2).setValues([['S', String(q)]]);
      try { // vira troca de frente feita
        var c = cadSku_(cvCad_(SpreadsheetApp.openById(ID_VALIDADE)), l[2], l[3]) || {}, tf = criarAbaTF_(SpreadsheetApp.getActive());
        tf.appendRow([v9Hoje_(), '', 'VALIDADE NA FRENTE', String(l[3]), c.secao || '', String(l[4] || c.local || ''), c.rep || '', String(l[6]), String(l[9]), 'FEITO', a.d, agoraBr_(), 'VNF|' + String(l[15]), Utilities.getUuid()]);
      } catch (e) {}
    } else {
      sh.getRange(r, 9).setValue('LOTE SUMIDO'); sh.getRange(r, 12).setValue('N'); sh.getRange(r, 17, 1, 2).setValues([[String(l[6]), 'AGUARDANDO']]);
      v9Push_('Celio', '👀 Lote sumido', String(l[3]) + ': lote de ' + String(l[6]).slice(0, 6) + String(l[6]).slice(8) + ' não foi achado (' + a.d + ')', 'gvalfrente');
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function vfResumo(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var seg = diaNum_(vsBr_(vsSegunda_(new Date()))), pess = {}, trocados = [], sumidos = [], naoTem = [], outras = [];
  v9Linhas_(v9Aba_(ABA_VNF, CAB_VNF), CAB_VNF.length).forEach(function (l) {
    var st = String(l[8]), res = String(l[10]), o = { id: String(l[15]), dia: String(l[0]), para: String(l[1]), sku: String(l[2]), desc: String(l[3]), certa: String(l[6]), qtdCerta: String(l[7]), resp: String(l[9]), quem: String(l[13]), em: String(l[14]), qtd: String(l[12]), dec: String(l[17]) };
    if (st === 'LOTE SUMIDO' && o.dec === 'AGUARDANDO') sumidos.push(o);
    if (diaNum_(o.dia) < seg) return;
    var pp = pess[o.para] = pess[o.para] || { nome: o.para, total: 0, resp: 0, certas: 0 };
    pp.total++; if (st !== 'PENDENTE') { pp.resp++; if (res === 'CERTA') pp.certas++; }
    if (st === 'TROCADO') trocados.push(o); else if (res === 'NAO TEM NA GONDOLA') naoTem.push(o); else if (res === 'OUTRA VALIDADE') outras.push(o);
  });
  var tirados = 0, semana = vsBr_(vsSegunda_(new Date()));
  try { v9Linhas_(cvAba_(ABA_CVS, CAB_CVS), CAB_CVS.length).forEach(function (l) { if (String(l[0]) === semana && String(l[10]) === 'PENDENTE' && vfVend_(l[5], l[4], l[3])) tirados++; }); } catch (e) {}
  try { v9Linhas_(criarAbaVS_(SpreadsheetApp.getActive()), CAB_VS.length).forEach(function (l) { if (String(l[0]) === semana && String(l[8]) === 'PENDENTE' && vfVend_(l[15], l[3], l[2])) tirados++; }); } catch (e) {}
  return { ok: true, pessoas: Object.keys(pess).map(function (k) { return pess[k]; }), trocados: trocados.reverse(), sumidos: sumidos.reverse(), naoTem: naoTem.reverse(), outras: outras.reverse(), tirados: tirados, confrontoEm: PropertiesService.getScriptProperties().getProperty('VF_DIA') || '' };
}
function vfSumido(p) { // Celio: PEDIR CONTAGEM ou IGNORAR um lote sumido
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = v9Aba_(ABA_VNF, CAB_VNF), r = v9Linha_(sh, 16, p.id); if (r < 0) return { ok: false, erro: 'item não encontrado' };
    var l = sh.getRange(r, 1, 1, CAB_VNF.length).getValues()[0];
    if (String(l[17]) !== 'AGUARDANDO') return { ok: false, erro: 'já decidido' };
    if (String(p.op) === 'CONTAGEM') {
      var para = String(p.para || l[1] || 'Ney');
      ctaPedir_(a.d, String(l[2]), String(l[3]), para, 'lote ' + String(l[16]).slice(0, 6) + String(l[16]).slice(8) + ' sumiu (validade na frente)', 'VALIDADE NA FRENTE');
      sh.getRange(r, 18).setValue('CONTAGEM PEDIDA');
    } else sh.getRange(r, 18).setValue('IGNORADO');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ---------- 🥬 CONSUMO PRÓPRIO — HORTIFRUTI (mais usados por pessoa + produtos aprendidos) ----------
var HF_SKUS = ['4922', '42098', '14697', '14696', '14695', '14700', '14755', '16329', '38966', '15996', '18742', '14694', '15196', '16876', '42173', '41747', '14699', '36546', '16449',
  '14701', '14693', '4924', '4926', '14751', '15536', '4925', '14692', '14691', '14705', '33795', '33793', '18723', '14698', '4923', '16584', '14702', '44492', '20723', '14752',
  '44846', '43825', '36788', '45590', '44710', '43827'];
var ABA_HFA = 'HORTI APRENDIDOS', CAB_HFA = ['QUEM', 'SKU', 'PRODUTO', 'EM'];
function hfMaisUsados(p) {
  var a = v9Dono_(p, 'PROPRIO'); if (a.erro) return a.erro;
  var ok = {}, novos = {}, q = a.d;
  HF_SKUS.forEach(function (s) { ok[s] = 1; });
  v9Linhas_(v9Aba_(ABA_HFA, CAB_HFA), 4).forEach(function (l) { if (String(l[0]) === q) { var s = skuN_(l[1]); ok[s] = 1; novos[s] = String(l[2]); } });
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_LANC), n = sh.getLastRow(), cont = {}, desc = {};
  if (n >= 2) { var ini = Math.max(2, n - 4000); sh.getRange(ini, 1, n - ini + 1, 9).getValues().forEach(function (l) {
    if (String(l[2]) !== q || !/PR[OÓ]PRIO/i.test(String(l[3]))) return;
    var s = skuN_(l[5]); if (!ok[s]) return; cont[s] = (cont[s] || 0) + 1; desc[s] = String(l[6]);
  }); }
  var out = Object.keys(cont).map(function (s) { return { sku: s, desc: desc[s], n: cont[s], novo: !!novos[s] && HF_SKUS.indexOf(s) < 0 }; });
  Object.keys(novos).forEach(function (s) { if (!cont[s] && HF_SKUS.indexOf(s) < 0) out.push({ sku: s, desc: novos[s], n: 0, novo: true }); });
  out.sort(function (x, y) { return y.n - x.n || x.desc.localeCompare(y.desc); });
  return { ok: true, itens: out.slice(0, 9) };
}
function hfAprender(p) {
  var a = v9Dono_(p, 'PROPRIO'); if (a.erro) return a.erro;
  var s = skuN_(p.sku); if (!s || HF_SKUS.indexOf(s) >= 0) return { ok: true };
  var sh = v9Aba_(ABA_HFA, CAB_HFA);
  if (v9Linhas_(sh, 2).some(function (l) { return String(l[0]) === a.d && skuN_(l[1]) === s; })) return { ok: true };
  sh.appendRow([a.d, "'" + s, String(p.desc || ''), agoraBr_()]);
  return { ok: true };
}

// ---------- selos da tela inicial ----------
function v9Resumo_(quem, r) {
  var ap = apps(quem), cel = ehCelio_(quem);
  try {
    if (ap.indexOf('CONTAGEM_ESTOQUE') >= 0 || cel) {
      var pc = 0, dec = 0, ag = 0;
      v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), 8).forEach(function (l) { var st = String(l[7]); if (st === 'AGUARDANDO CONTAGEM') { ag++; if (tfMeu_(l[5], quem)) pc++; } else if (st === 'CONTADA - DECIDIR') dec++; });
      r.cta = pc; if (cel) { r.ctaDec = dec; r.ctaAg = ag; }
    }
  } catch (e) {}
  try { if (cel) { var dz = decResumo_(); r.dec = dz.n; r.decUrg = dz.urg; } } catch (e) {} // v113
  try { var gq = grmPendQuem_(quem); if (gq) r.grm = gq; } catch (e) {} // v118
  try { if (cel) { var az = agdResumo_(); r.agdEsp = az.esp; r.agdErr = az.err; } } catch (e) {} // v114
  try {
    if (ap.indexOf('ETIQUETAS_FILA') >= 0 || ap.indexOf('ETIQUETAS') >= 0 || cel) {
      var fila = 0, prontas = 0, hoje = v9Hoje_();
      v9Linhas_(v9Aba_(ABA_ETQ, CAB_ETQ), 15).forEach(function (l) { if (String(l[10]) === 'NA FILA') fila++; if (String(l[3]) === quem && String(l[10]) === 'IMPRESSA' && String(l[14]).indexOf(hoje) === 0) prontas++; });
      r.etqFila = fila; r.etqProntas = prontas;
    }
  } catch (e) {}
  try { if (ap.indexOf('POFF_FILA') >= 0 || cel) pfResumo_(quem, r); } catch (e) {} // v110.3
  try {
    if (cel) { var sm = 0; v9Linhas_(v9Aba_(ABA_VNF, CAB_VNF), CAB_VNF.length).forEach(function (l) { if (String(l[8]) === 'LOTE SUMIDO' && String(l[17]) === 'AGUARDANDO') sm++; }); r.vfSum = sm; }
  } catch (e) {}
  return r;
}
function v9ProdInfo(p) { // produto pelo código/SKU para conferir na contagem (qualquer pessoa logada)
  var a = v9Dono_(p); if (a.erro) return a.erro;
  var r = v9Prod_(p); return r.ok ? { ok: true, sku: r.sku, desc: r.desc } : r;
}

// ========== VALIDADES — CONFERÊNCIA DE VALIDADE (v82) ==========
// Repositoras: produtos da seção delas (CADASTRO da VALIDADE PRODUTOS) com validade na semana que começa daqui a ~30 dias (seg+28..seg+34).
// Ney: todas as seções, semana daqui a ~45 dias (seg+42..seg+48). Uma linha por produto + validade. Sem seção → Ney, bloco "Sem seção".
// Pendente da semana anterior → ATRASADO no topo. Crítica: média diária (aba MEDIAS VENDA, mantida pelo Claude) × dias até vencer.
var ABA_CVS = 'CV SEMANA', ABA_CV = 'CONFERENCIA VALIDADE', ABA_MED = 'MEDIAS VENDA';
var CAB_CVS = ['SEMANA', 'PARA', 'JANELA', 'VALIDADE', 'PRODUTO', 'SKU', 'SECAO', 'LOCAL', 'REPOSITORA', 'ATRASADO', 'STATUS', 'CHAVE', 'ID', 'SEM SECAO', 'QTD AREA', 'QTD ESTOQUE', 'TOTAL', 'ESTOQUE PEND', 'CRITICA', 'QUEM', 'EM'];
var CAB_CV = ['DATA', 'HORA', 'QUEM', 'SKU', 'PRODUTO', 'VALIDADE', 'QTD AREA DE VENDA', 'QTD ESTOQUE', 'TOTAL', 'SITUACAO', 'ID ITEM', 'JANELA', 'MEDIA DIA', 'PREVISAO VENDA', 'SOBRA', 'CRITICA', 'CARTAO', 'DECISAO'];
function cvAba_(nome, cab) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(nome);
  if (sh) return sh;
  sh = ss.insertSheet(nome);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), cab.length).setNumberFormat('@');
  return sh;
}
function cvCad_(ssV) {
  var sh = ssV.getSheetByName('CADASTRO'), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 5).getValues().forEach(function (l) { var k = normDesc_(l[0]), o = { secao: String(l[1] || '').trim(), local: String(l[2] || '').trim(), rep: String(l[3] || '').trim() }, s = skuN_(l[4]); if (k) m[k] = o; if (s) m['#' + s] = o; }); // v99: também pelo SKU (coluna E)
  return m;
}
var CV_FORA_45 = ['ACOUGUE', 'FRACIONADOS']; // v85: seções fora da lista de 45 dias do Ney
function cvFora45_(sec) { return CV_FORA_45.indexOf(semAcento_(String(sec || '')).toUpperCase().trim()) >= 0; }
function atualizarCV() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    try { cadSkuPreencher(); } catch (e) {} // v99: SKU na coluna E do CADASTRO
    var sh = cvAba_(ABA_CVS, CAB_CVS), hoje = vsDia_(new Date()), seg = vsSegunda_(hoje), semana = vsBr_(seg);
    var n = sh.getLastRow(), vals = n >= 2 ? sh.getRange(2, 1, n - 1, CAB_CVS.length).getValues() : [];
    var chaves = {}, temSemana = false, novos = [], transf = 0, add = 0;
    var ssV = SpreadsheetApp.openById(ID_VALIDADE), tzV = ssV.getSpreadsheetTimeZone(), cad = cvCad_(ssV), skus = vsSkuPorNome_();
    // v83: item "sem seção" que ganhou seção no CADASTRO vai para a repositora certa (ou sai, se ela já tem o mesmo item)
    var ativos = {}; vals.forEach(function (l) { if (String(l[0]) === semana && String(l[10]) !== 'TRANSFERIDO' && String(l[10]) !== 'REMOVIDO') ativos[String(l[11])] = 1; });
    vals.forEach(function (l, i) {
      if (String(l[0]) !== semana || String(l[13]) !== 'S' || String(l[10]) !== 'PENDENTE') return;
      var k = normDesc_(l[4]), c = cadSku_(cad, l[5], l[4]); if (!c || !c.secao) return;
      var para = String(l[2]) === '45' ? 'Ney' : (c.rep || 'Ney'), ch = para + '|' + String(l[2]) + '|' + k + '|' + String(l[3]);
      if (ch !== String(l[11]) && ativos[ch]) { sh.getRange(i + 2, 11).setValue('REMOVIDO'); l[10] = 'REMOVIDO'; return; }
      delete ativos[String(l[11])]; ativos[ch] = 1;
      sh.getRange(i + 2, 2).setValue(para); sh.getRange(i + 2, 7, 1, 3).setValues([[c.secao, c.local, c.rep]]); sh.getRange(i + 2, 12).setValue(ch); sh.getRange(i + 2, 14).setValue('');
      l[1] = para; l[6] = c.secao; l[7] = c.local; l[8] = c.rep; l[11] = ch; l[13] = '';
    });
    // v85: seções que não entram na lista de 45 dias do Ney (pedido do Celio: AÇOUGUE)
    vals.forEach(function (l, i) { if (String(l[0]) === semana && String(l[2]) === '45' && String(l[10]) === 'PENDENTE' && cvFora45_(l[6])) { sh.getRange(i + 2, 11).setValue('REMOVIDO'); l[10] = 'REMOVIDO'; } });
    vals.forEach(function (l) { if (String(l[0]) === semana) { temSemana = true; chaves[String(l[11])] = 1; } });
    if (!temSemana) vals.forEach(function (l, i) {
      if (String(l[10]) !== 'PENDENTE') return;
      sh.getRange(i + 2, 11).setValue('TRANSFERIDO'); transf++;
      var r = l.slice(); r[0] = semana; r[9] = 'S'; r[12] = Utilities.getUuid(); chaves[String(l[11])] = 1; novos.push(r.map(String));
    });
    var pc = ssV.getSheetByName('POR CHEGADA'), np = pc.getLastRow(), linhas = np >= 2 ? pc.getRange(2, 1, np - 1, 10).getValues() : [];
    var jan = [{ dias: 30, ini: vsMais_(seg, 28), fim: vsMais_(seg, 34) }, { dias: 45, ini: vsMais_(seg, 42), fim: vsMais_(seg, 48) }];
    linhas.forEach(function (l) {
      var prod = String(l[3] || '').replace(/[¬]/g, '').trim(), val = l[9];
      if (!prod || !(val instanceof Date)) return;
      var pv = Utilities.formatDate(val, tzV, 'yyyy-MM-dd').split('-'); val = new Date(+pv[0], +pv[1] - 1, +pv[2]);
      var k = normDesc_(prod), sk = skus[k] || { sku: '' }, c = cadSku_(cad, sk.sku, prod) || { secao: '', local: '', rep: '' };
      jan.forEach(function (j) {
        if (val < j.ini || val > j.fim) return;
        if (j.dias === 45 && cvFora45_(c.secao)) return; // v85
        var sem = !c.secao, para = j.dias === 45 || sem ? 'Ney' : (c.rep || 'Ney');
        var ch = para + '|' + j.dias + '|' + k + '|' + vsBr_(val); if (chaves[ch]) return; chaves[ch] = 1;
        novos.push([semana, para, String(j.dias), vsBr_(val), prod, sk.sku, c.secao, c.local, c.rep, '', 'PENDENTE', ch, Utilities.getUuid(), sem ? 'S' : '', '', '', '', '', '', '', '']);
        add++;
      });
    });
    if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, CAB_CVS.length).setValues(novos);
    var med = null; if (!temSemana || !SpreadsheetApp.getActive().getSheetByName(ABA_MED)) { try { med = cvAtualizarMedias(); } catch (e) { med = { ok: false, erro: String(e) }; } } // médias: toda segunda
    return { ok: true, semana: semana, novos: add, atrasados: transf, medias: med };
  } finally { lock.releaseLock(); }
}
function cvMedias_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_MED), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (l) { var s = String(l[0]).replace(/^'/, '').trim().replace(/^0+(?=\d)/, ''); if (s) m[s] = Number(String(l[2]).replace(',', '.')) || 0; });
  return m;
}
function cvLista_(quem) {
  var tudo = verTudo_(quem), sh = cvAba_(ABA_CVS, CAB_CVS), semana = vsBr_(vsSegunda_(new Date())), n = sh.getLastRow(), out = [];
  if (n < 2) return { semana: semana, itens: out, tudo: tudo };
  var vals = sh.getRange(2, 1, n - 1, CAB_CVS.length).getValues(), ney = {}, sg = lsugMapa_(); // v99
  vals.forEach(function (l) { if (String(l[2]) === '45' && (String(l[10]) === 'CONFERIDO' || String(l[10]) === 'NAO TEM')) { var k = String(l[11]).split('|').slice(2).join('|'); ney[k] = { total: String(l[16]), area: String(l[14]), est: String(l[15]), st: String(l[10]), em: String(l[20]) }; } });
  vals.forEach(function (l) {
    if (String(l[0]) !== semana || String(l[10]) === 'TRANSFERIDO' || String(l[10]) === 'REMOVIDO') return;
    if (!tudo && !tfMeu_(l[1], quem)) return;
    if (String(l[10]) === 'PENDENTE' && vfVend_(l[5], l[4], l[3])) return; // v109: lote já vendido (confrontação) sai da lista
    var kk = String(l[11]).split('|').slice(2).join('|');
    out.push({ id: String(l[12]), para: String(l[1]), janela: String(l[2]), validade: String(l[3]), desc: String(l[4]), sku: String(l[5]), secao: String(l[6]), local: String(l[7]), rep: String(l[8]), atrasado: String(l[9]) === 'S', status: String(l[10]), semSecao: String(l[13]) === 'S', area: String(l[14]), estoque: String(l[15]), total: String(l[16]), estPend: String(l[17]) === 'S', critica: String(l[18]), quem: String(l[19]), em: String(l[20]), rapida: (function () { var x = String(l[10]) === 'PENDENTE' ? vfSit_(l[5], l[3]) : null; return !!(x && x.sit === 'VENDIDO'); })(), ney: String(l[2]) === '30' ? (ney[kk] || null) : null, sug: sg['#' + skuN_(l[5])] || sg[normDesc_(l[4])] || null });
  });
  return { semana: semana, itens: out, tudo: tudo };
}
function cvSalvar(p) {
  var dono = quemDoToken_(p.token);
  if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  if (apps(dono).indexOf('VALIDADES') < 0) return { ok: false, erro: 'sem acesso' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = cvAba_(ABA_CVS, CAB_CVS), n = sh.getLastRow(), r = -1;
    if (n >= 2) { var ids = sh.getRange(2, 13, n - 1, 1).getValues(); for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(p.id)) { r = i + 2; break; } }
    if (r < 0) return { ok: false, erro: 'item não encontrado' };
    var l = sh.getRange(r, 1, 1, CAB_CVS.length).getValues()[0];
    var vd = null; // v100: validade diferente
    if (String(p.novaVal || '').trim() && String(p.naoTem || '') !== '1') { vd = cvValCaso_(l, p.novaVal); if (!vd.ok) return vd; if (vd.caso === 'EXISTE') return { ok: false, existe: true, erro: 'essa validade já está no cadastro — ela aparece numa lista mais para frente' }; }
    var valItem = vd ? vd.nova : String(l[3]);
    var num = function (x) { x = String(x == null ? '' : x).trim().replace(',', '.'); return x === '' ? null : Number(x); };
    var naoTem = String(p.naoTem || '') === '1', area = naoTem ? 0 : num(p.area), est = naoTem ? 0 : num(p.estoque);
    if (!naoTem && (area == null || isNaN(area) || area < 0)) return { ok: false, erro: 'informe a quantidade da área de venda' };
    if (est != null && (isNaN(est) || est < 0)) return { ok: false, erro: 'quantidade do estoque inválida' };
    var total = (area || 0) + (est || 0), estPend = !naoTem && est == null;
    var situ = naoTem ? 'NAO TEM MAIS' : (estPend ? 'ESTOQUE NAO CONFERIDO' : 'CONFERIDO');
    if (vd && vd.caso === 'NOVA') situ = 'NAO ACHOU ' + l[3] + ' - ACHOU ' + vd.nova; else if (vd) situ += ' (VALIDADE CORRIGIDA ' + l[3] + ' -> ' + vd.nova + ')'; // v100
    // crítica: quanto vende até o vencimento
    var med = cvMedias_()[String(l[5]).replace(/^'/, '').trim().replace(/^0+(?=\d)/, '')], val = vsDeBr_(valItem), hoje = vsDia_(new Date());
    var dias = val ? Math.max(0, Math.round((val - hoje) / 86400000)) : 0, prev = '', sobra = '', crit = '';
    if (naoTem || total === 0) crit = '';
    else if (med == null || !String(l[5]).trim()) crit = 'SEM MEDIA';
    else { prev = Math.round(med * dias * 10) / 10; sobra = Math.max(0, Math.ceil(total - prev)); crit = sobra > 0 ? 'SUGERIR REBAIXA' : 'VAI VENDER'; if (!sobra) sobra = ''; }
    // v99: já está na rebaixa -> não sugere de novo; senão a sugestão vai para o bloco A DECIDIR da planilha REBAIXAS
    if (crit === 'SUGERIR REBAIXA') { try { var stR = rebSituacao_(rebIndice_(), l[5], l[4], valItem); if (stR && stR.tipo === 'REB') crit = 'JA NA REBAIXA ' + stR.por; else if (!stR) rebSugerirLote_([{ sku: String(l[5]), desc: String(l[4]), validade: valItem, obs: obsQtd_(total) + (estPend ? ' AREA (deposito nao conferido)' : '') + ' - deve sobrar ' + sobra }]); } catch (e) {} }
    var agora = new Date(), tz = Session.getScriptTimeZone(), em = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm');
    sh.getRange(r, 11).setValue(naoTem || (vd && vd.caso === 'NOVA') ? 'NAO TEM' : 'CONFERIDO');
    if (vd) cvValAplicar_(sh, r, l, vd, dono, total, p.id); // v100
    sh.getRange(r, 15, 1, 7).setValues([[naoTem ? '0' : String(area), est == null ? '' : String(est), String(total), estPend ? 'S' : '', crit, dono, em]]);
    // registro na aba CONFERENCIA VALIDADE (uma linha por item; correção atualiza a mesma linha)
    var cv = cvAba_(ABA_CV, CAB_CV), m = cv.getLastRow(), rr = -1;
    if (m >= 2) { var it = cv.getRange(2, 11, m - 1, 1).getValues(); for (var j = it.length - 1; j >= 0; j--) if (String(it[j][0]) === String(p.id)) { rr = j + 2; break; } }
    var linha = [Utilities.formatDate(agora, tz, 'dd/MM/yyyy'), Utilities.formatDate(agora, tz, 'HH:mm'), dono, String(l[5]), String(l[4]), valItem, naoTem ? '0' : String(area), est == null ? '' : String(est), String(total), situ, String(p.id), String(l[2]), med == null ? '' : String(med), String(prev), String(sobra), crit];
    if (rr < 0) cv.getRange(m + 1, 1, 1, linha.length).setValues([linha]);
    else cv.getRange(rr, 1, 1, linha.length).setValues([linha]);
    marcarOrigem_(cv, rr < 0 ? m + 1 : rr, 1, CAB_CV.length);
    var gk = null; try { gk = decGondolaRetorno_(l, { quem: dono, naoTem: naoTem, vd: vd, total: total, area: area, est: est, val: valItem, cvRow: rr < 0 ? m + 1 : rr }); } catch (e) {} // v121
    return { ok: true, critica: crit, sobra: sobra, prev: prev, total: total, vd: vd ? { caso: vd.caso, nova: vd.nova } : null, gondola: !!gk };
  } finally { lock.releaseLock(); }
}
// Médias de venda por dia (aba MEDIAS VENDA): vendas diarias.xlsx da pasta BD (média de 90 dias) ou, enquanto ele não existir, a Curva A mais recente.
var ID_PASTA_CURVA = '1xr7KI96FkNutO2zTgwBFpIdRz4lRnSai'; // RELATORIOS / CURVA A
function cvAtualizarMedias() {
  var out = [], fonte = '', hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var itV = DriveApp.getFolderById(ID_PASTA_BD).getFilesByName('vendas diarias.xlsx');
  if (itV.hasNext()) {
    var rows = lerXlsx_(itV.next().getId()), soma = {}, nome = {}, dias = {};
    rows.slice(1).forEach(function (r) { var s = String(r.B || '').replace(/^0+(?=\d)/, '').trim(); if (!s) return; soma[s] = (soma[s] || 0) + (Number(String(r.D || '').replace(',', '.')) || 0); nome[s] = String(r.C || ''); dias[String(r.A || '')] = 1; });
    var nd = Math.max(1, Math.min(90, Object.keys(dias).length));
    Object.keys(soma).forEach(function (s) { out.push(["'" + s, nome[s], String(Math.round(soma[s] / nd * 1000) / 1000), 'vendas diarias.xlsx (' + nd + ' dias)', hoje]); });
    fonte = 'vendas diarias (' + nd + ' dias)';
  } else {
    var it = DriveApp.getFolderById(ID_PASTA_CURVA).getFilesByType(MimeType.GOOGLE_SHEETS), f = null;
    while (it.hasNext()) { var x = it.next(); if (/^CURVA A /.test(x.getName()) && (!f || x.getDateCreated() > f.getDateCreated())) f = x; }
    if (!f) return { ok: false, erro: 'sem vendas diarias.xlsx e sem Curva A' };
    var v = SpreadsheetApp.openById(f.getId()).getSheets()[0].getDataRange().getValues();
    v.slice(1).forEach(function (r) { var s = String(r[0] || '').replace(/^0+(?=\d)/, '').trim(), q = typeof r[3] === 'number' ? r[3] : Number(String(r[3] || '').replace(/\./g, '').replace(',', '.')); if (!s || isNaN(q)) return; out.push(["'" + s, String(r[1] || ''), String(Math.round(q / 92 * 1000) / 1000), f.getName() + ' (vendido/92 dias)', hoje]); });
    fonte = f.getName();
  }
  var sh = cvAba_(ABA_MED, ['SKU', 'PRODUTO', 'MEDIA DIA', 'FONTE', 'ATUALIZADO']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  return { ok: true, fonte: fonte, produtos: out.length };
}

// v112.1: teste da conversão da validade (não grava nada) — NF 375935 Caapuã
function valTeste375935() { ['51261008768211000180550010003759351119353683', '51261026564534000215550010004071961669570166'].forEach(function (id) { var r = regravarValidades_(id, true); Logger.log(JSON.stringify({ nf: r.r && r.r.nNF, det: r.r && r.r.det, conferir: r.r && r.r.conferir })); }); }

// ========== v113 — 🗂️ DECISÕES (Gestão, só o Celio): Rebaixas, Produto novo (NF), Contagem, Código de barras — modelos aprovados em 10/11/2026 ==========
var ABA_DEC = 'DECISOES';
var CAB_DEC = ['ID', 'TIPO', 'CHAVE', 'CRIADO EM', 'STATUS', 'VOLTA EM', 'DECISAO', 'DETALHE', 'DECIDIDO POR', 'DECIDIDO EM', 'EXECUTADO EM', 'OBS'];
// STATUS: ESPERANDO (volta em VOLTA EM) | DECIDIDA (feita no app) | A EXECUTAR (o Claude faz no Plenus) | EXECUTADA
function decAba_() { return v9Aba_(ABA_DEC, CAB_DEC); }
// v131 (Fase 1 velocidade): leituras grandes (CONFERENCIA ITENS / LANC) feitas UMA vez por chamada da lista de Decisões
var DEC_MEMO_ = null;
function decMemoVals_(nome, sh, nc) {
  var n = sh.getLastRow(); if (n < 2) return [];
  if (!DEC_MEMO_) return sh.getRange(2, 1, n - 1, nc).getValues();
  var k = nome + '|' + nc; if (!DEC_MEMO_[k]) DEC_MEMO_[k] = sh.getRange(2, 1, n - 1, nc).getValues();
  return DEC_MEMO_[k];
}
function decMapa_() { // chave -> última linha (objeto)
  var m = {}; v9Linhas_(decAba_(), CAB_DEC.length).forEach(function (l, i) { m[String(l[2])] = { row: i + 2, tipo: String(l[1]), status: String(l[4]), volta: v9Br_(l[5]), decisao: String(l[6]), detalhe: String(l[7]), por: String(l[8]), em: agdTxt_(l[9]), exec: agdTxt_(l[10]), obs: String(l[11]) }; });
  return m;
}
function decGrava_(tipo, chave, status, volta, decisao, detalhe, quem, obs) {
  if (status !== 'EXECUTADA') { var corr = decCorrPega_(chave); if (corr) { detalhe = detalhe || {}; detalhe.corrigida = corr; obs = 'CORRIGIDA (antes: ' + corr.de + ' em ' + corr.deEm + ')' + (obs ? ' - ' + obs : ''); } } // v127
  try { CacheService.getScriptCache().remove('DECL131_n'); } catch (e) {} // v131
  var sh = decAba_(), m = decMapa_(), d = m[chave], l = [d ? '' : Utilities.getUuid(), tipo, chave, d ? '' : agoraBr_(), status, volta || '', decisao, JSON.stringify(detalhe || {}), quem, agoraBr_(), status === 'EXECUTADA' ? agoraBr_() : '', obs || ''];
  if (d) { var old = sh.getRange(d.row, 1, 1, CAB_DEC.length).getValues()[0]; l[0] = old[0]; l[3] = old[3]; sh.getRange(d.row, 1, 1, CAB_DEC.length).setValues([l]); }
  else sh.appendRow(l);
}
function decEsperando_(d) { return d && d.status === 'ESPERANDO' && diaNum_(d.volta) > diaNum_(v9Hoje_()); }
function decRs_(v) { var n = typeof v === 'number' ? v : Number(String(v || '').replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3})/g, '').replace(',', '.')); return isNaN(n) || !String(v).trim() ? null : n; }

// ---- REBAIXAS: linhas a decidir da aba do mês na planilha REBAIXAS ----
function decRebAba_() { var d = new Date(); return SpreadsheetApp.openById(ID_PLANILHA_REBAIXAS).getSheetByName(MES_PT[d.getMonth()] + String(d.getFullYear()).slice(2)); }
function decRebCartoes_(m) {
  var sh = decRebAba_(); if (!sh || sh.getLastRow() < 2) return [];
  var hoje = diaNum_(v9Hoje_()), out = [], nome = sh.getName(), oc = {}; // v116: oc = cache da origem
  sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues().forEach(function (l, i) {
    var prod = String(l[1] || '').trim(), val = v9Br_(l[9]), k = String(l[10] || ''); if (!prod || !val || /^A DECIDIR$/i.test(prod)) return;
    if (l[11] === true || String(l[11]).toUpperCase() === 'TRUE' || /^OK$/i.test(String(l[11]))) return; // já decidida (OK marcado)
    if (/NAO FAZER/i.test(k)) return; // desistiu, ou saldo do lote novo
    var dv = diaNum_(val); if (!(dv >= hoje)) return; // venceu: sai da lista (segue pela avaria)
    var chave = 'REB|' + nome + '|' + prod + '|' + val, d = m[chave];
    var q = (k.match(/(\d+)\s*UN DESTA VALIDADE/i) || k.match(/(\d+)\s*UN/i) || [])[1], sob = (k.match(/deve sobrar (\d+)/i) || [])[1], sv = (k.match(/sem venda (\d+) dias/i) || [])[1];
    var al = /CONFERIR NA GONDOLA/i.test(k) ? '🔎 Conferir na gôndola: saldo maior que os lotes registrados' : (/LOTE INVERTIDO/i.test(k) ? '🔁 Lote invertido: um lote posterior vence antes — trocar a frente' : '');
    out.push({ tipo: 'REBAIXA', chave: chave, linha: i + 2, aba: nome, prod: prod, validade: val, dias: dv - hoje, qtd: q || '', sobra: sob || '', semVenda: sv || '', conferido: /LOTE CONFERIDO/i.test(k), alerta: al, obs: k, pauvan: decRs_(l[2]), osso: decRs_(l[6]), pagou: decRs_(l[7]), origem: String(l[12] || ''), org: (function () { try { return decRebOrigem_(prod, val, k, oc); } catch (e) { return null; } })(), esperando: decEsperando_(d), volta: d && d.status === 'ESPERANDO' ? d.volta : '', gondola: decGondolaInfo_(d) });
  });
  return out.sort(function (a, b) { return a.dias - b.dias; });
}
function decRebLinha_(sh, c) { // acha de novo a linha (a planilha pode ter mudado)
  var n = sh.getLastRow(); if (n < 2) return -1; var v = sh.getRange(2, 1, n - 1, 10).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][1]).trim() === c.prod && v9Br_(v[i][9]) === c.validade) return i + 2;
  return -1;
}

// ---- PRODUTO NOVO: itens ** da conferência (nota ainda sem vínculo) ----
function decCadLista_() { // [{sku, desc, un}] únicos por SKU, com palavras
  var c = CacheService.getScriptCache(), k = 'DEC_CAD_V113', em = c.get(k); if (em) return JSON.parse(em);
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = sh.getLastRow(), vist = {}, out = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, 7).getValues().forEach(function (l) { var s = skuN_(l[1]); if (!s || vist[s]) return; vist[s] = 1; out.push([s, String(l[2]), String(l[3]), String(l[6]) === 'NV' ? 1 : 0]); });
  try { c.put(k, JSON.stringify(out), 900); } catch (e) {}
  return out;
}
function decPalavras_(t) { return semAcento_(String(t || '')).toUpperCase().replace(/\*+/g, ' ').replace(/[^A-Z0-9 ]/g, ' ').split(/\s+/).filter(function (w) { return w.length >= 3 && !/^(KG|UND|UNID|CX|FD|PCT|COM|PARA|SEM|TIPO)$/.test(w) && !/^\d+$/.test(w); }); }
var DEC_PAL_ = null; // v131: palavras dos cadastros (1x por chamada)
function decParecidos_(desc, lista, n) {
  var ws = decPalavras_(desc); if (!ws.length) return [];
  var r = [], pal = (DEC_PAL_ && DEC_PAL_.n === lista.length) ? DEC_PAL_.w : (DEC_PAL_ = { n: lista.length, w: lista.map(function (c) { return decPalavras_(c[1]); }) }).w; // v131
  lista.forEach(function (c, ic) { var cw = pal[ic], pt = 0; ws.forEach(function (w, i) { if (cw.indexOf(w) >= 0) pt += i === 0 ? 3 : 1; }); if (pt >= 2 && cw.indexOf(ws[0]) >= 0 || pt >= 3) r.push({ sku: c[0], desc: c[1], un: c[2], nv: !!c[3], pt: pt - cw.length * 0.05 }); });
  return r.sort(function (a, b) { return b.pt - a.pt; }).slice(0, n || 3);
}
function decNovoPar_(orig, dn, lista) { // v118.7: cadastros com o mesmo código de barras da nota vêm primeiro (sem pré-escolha)
  var ss = /^EAN AMBIGUO:/.test(String(orig || '')) ? String(orig).replace(/^EAN AMBIGUO:\s*/, '').split(/\s*\/\s*/).map(skuN_) : [], out = [];
  ss.forEach(function (s) { var c = lista.filter(function (x) { return x[0] === s; })[0]; out.push({ sku: s, desc: c ? c[1] : '(cadastro ' + s + ')', un: c ? c[2] : '', nv: !!(c && c[3]), ean: 1 }); });
  decParecidos_(dn, lista, 3).forEach(function (p) { if (ss.indexOf(String(p.sku)) < 0) out.push(p); });
  return out;
}
function decNovoCartoes_(m) {
  try { autoVincBipado_(); } catch (e) {} // v118.5
  var abas = criarAbaConf_(SpreadsheetApp.getActive()), nfs = {}, out = [], lista = null, med = null, apr = vinculosAprendidos_();
  v9Linhas_(abas.nf, 6).forEach(function (l) { nfs[String(l[0])] = { forn: String(l[1]), nf: String(l[2]), entrada: v9Br_(l[4]), status: String(l[5]) }; });
  decMemoVals_('CIT', abas.it, CAB_CIT.length + EXTRA_CIT.length).forEach(function (l) { // v131: memo
    if (String(l[9]) !== 'S') return; var nf = nfs[String(l[0])]; if (!nf || nf.status === 'EXCECAO') return;
    var chave = 'NOVO|' + String(l[14]), d = m[chave]; if (d && (d.status === 'A EXECUTAR' || d.status === 'EXECUTADA' || d.status === 'DECIDIDA')) return;
    if (!lista) lista = decCadLista_();
    var dn = String(l[15] || l[3]).replace(/^\*+\s*/, ''), fi = valFatorItem_(l, {}, apr, cnpjDoIdNf_(l[0])), un = String(l[5]).toUpperCase();
    out.push({ tipo: 'PRODUTO_NOVO', chave: chave, idNf: String(l[0]), idItem: String(l[14]), forn: nf.forn, nf: nf.nf, entrada: nf.entrada, statusNf: nf.status, prod: dn, un: un, qtd: String(l[6]), fator: fi.f || '', codNota: String(l[18] || ''), ean: String(l[4] || '').replace(/^'/, ''), parecidos: decNovoPar_(l[23], dn, lista), eanAmb: /^EAN AMBIGUO/.test(String(l[23] || '')),
      sugestao: { desc: decDescSug_(dn, un, fi.f), un: un === 'KG' || parseFloat(((dn.match(/\s(\d+(?:[,.]\d+)?)\s*KG\.?$/i) || [])[1] || '0').replace(',', '.')) >= 3 ? 'QUILO' : 'UNIDADE', entrada: fi.f > 1 ? '1 ' + (un === 'UN' ? 'CX' : un) + ' = ' + fi.f : '' },
      esperando: decEsperando_(d), volta: d && d.status === 'ESPERANDO' ? d.volta : '' });
  });
  return out;
}
function decDescSug_(d, un, f) { // descrição no padrão Pauvan: maiúscula, sem acento, sem a embalagem de entrada, terminada em **
  var s = semAcento_(String(d)).toUpperCase().replace(/\s+/g, ' ').trim();
  s = s.replace(/\s+(\d+\s*X\s*)?(CX|FD|FARDO|CAIXA|PCT|DP|DISPLAY)\s*(C\/)?\s*\d+\s*(UN|UND|PCT|DZ)?\.?$/, '').replace(/^(\d+)\s*X\s*/, '').replace(/\s+\d+\s*X\s*(?=\d)/, ' ');
  s = s.replace(/\s(\d+(?:[,.]\d+)?)\s*KG\.?$/, function (m0, n) { return parseFloat(n.replace(',', '.')) >= 3 ? ' KG' : m0; }); // hortifruti: caixa de 18 KG -> produto por KG
  return s.replace(/\*+$/, '').trim() + '**';
}

// ---- CÓDIGO DE BARRAS: bipado na conferência e não cadastrado (COD_STATUS = NOVO) ----
function decCodigosSku_() { // sku -> [códigos do cadastro]
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { var s = skuN_(l[1]), c = String(l[0]).replace(/^'/, '').trim(); if (s && c && c !== s) (m[s] = m[s] || []).push(c); });
  return m;
}
function decCodCartoes_(m, leve) {
  var abas = criarAbaConf_(SpreadsheetApp.getActive()), nfs = {}, vis = {}, out = [], cods = null, todos = null, med = null, cad = null, sal = {}, grm = null, rep = null; // v118: grm = conferir gramatura
  v9Linhas_(abas.nf, 3).forEach(function (l) { nfs[String(l[0])] = String(l[1]) + ' ' + String(l[2]); });
  var m2 = abas.la.getLastRow(); if (m2 < 2) return out;
  decMemoVals_('CLA', abas.la, CAB_CLA.length + EXTRA_CLA.length).forEach(function (l) { // v131: memo
    if (String(l[18]) !== 'NOVO') return; var cl = String(l[17] || '').replace(/^'/, '').trim(); if (!cl || vis[cl]) return; vis[cl] = 1;
    var chave = 'COD|' + cl, d = m[chave]; if (d && d.status !== 'ESPERANDO') return;
    if (/^\*\*/.test(String(l[3])) || !(diaNum_(v9Br_(l[10])) >= diaNum_(v9Hoje_()) - 30)) return; // item sem cadastro é do cartão Produto novo; só os últimos 30 dias
    if (!cods) { cods = decCodigosSku_(); todos = {}; Object.keys(cods).forEach(function (x) { cods[x].forEach(function (c) { todos[c] = 1; }); }); }
    if (todos[cl] || todos[cl.replace(/^0+/, '')]) return; // o código já foi cadastrado
    if (leve) { out.push({ esperando: decEsperando_(d) }); return; }
    if (!med) { med = cvMedias_(); cad = decCadLista_(); sal = v9Saldos_().itens; }
    var sku = skuN_(l[2]), atual = (cods[sku] || []).filter(function (c) { return c.length >= 8; }), m30 = Math.round((med[sku] || 0) * 30), s = { qtd: sal[sku] ? sal[sku].q : null };
    var soInterno = !atual.length || atual.every(function (c) { return c.length === 14; }), uva = /^UVA .*BDJ/i.test(String(l[3]));
    var sug = soInterno ? 'TROCAR' : (m30 > 0 ? 'PROMO' : 'TROCAR'), mot = soInterno ? 'o cadastro só tem código interno/de caixa' : (m30 > 0 ? 'o código atual vendeu ' + m30 + ' un em 30 dias → sugestão: cadastrar PROMO (o antigo continua valendo)' : 'o código atual está sem venda → sugestão: trocar');
    var irmao = null;
    if (uva) { var base = semAcento_(String(l[3])).toUpperCase().replace(/\*+/g, '').trim(); cad.forEach(function (c) { if (!irmao && semAcento_(c[1]).toUpperCase().replace(/\*+/g, '').trim() === base && c[0] !== sku && !(med[c[0]] > 0)) irmao = { sku: c[0], desc: c[1] }; }); }
    if (!grm) { grm = grmMapa_(); rep = mapaReposicao_(); } // v118
    var g = grm[chave] || null, gr = null, cr = cadSku_(rep, sku, String(l[3])) || {}, pp = ['Regiane'];
    [String(l[9]), cr.rep, 'Ney'].forEach(function (x) { String(x || '').split(/\s*\/\s*/).forEach(function (y) { y = y.trim(); if (y && pp.indexOf(y) < 0) pp.push(y); }); });
    if (g && (g.status === 'RESPONDIDO' || (g.status === 'PENDENTE' && d && d.decisao === 'CONFERIR GRAMATURA'))) gr = { status: g.status, resp: g.resp, peso: g.peso, foto: g.foto, por: g.por, em: g.em, para: g.para, cad: g.cad };
    if (gr && gr.resp === 'DIFERENTE') { sug = 'NOVO_PESO'; mot = 'peso diferente na embalagem: ' + gr.peso + ' (cadastro ' + (gr.cad || '?') + ') → é outro produto: cadastrar novo'; }
    else if (gr && gr.resp === 'MESMO') mot = 'mesmo peso, conferido por ' + gr.por + ' → só mudou o código; ' + mot;
    else if (gr && gr.resp === 'NAO ACHEI') mot = gr.por + ' não achou o produto na loja';
    out.push({ tipo: 'CODIGO', chave: chave, pesoCad: grmPeso_(String(l[3])), grm: gr, pessoas: pp, descNova: gr && gr.resp === 'DIFERENTE' ? grmDescNova_(String(l[3]), gr.peso.split(' ')[0], gr.peso.split(' ')[1]) : '', idNf: String(l[0]), idItem: String(l[1]), nf: nfs[String(l[0])] || '', sku: sku, prod: String(l[3]), bipado: cl, atual: atual.join(' / ') || '(só código interno)', quem: String(l[9]), data: v9Br_(l[10]), venda30: m30, saldo: s.qtd, sugestao: uva && irmao ? 'UVA' : sug, motivo: uva && irmao ? 'uva: código muda a cada compra — usar o cadastro ' + irmao.sku + ' (mesma uva, sem venda)' : mot, irmao: irmao, esperando: decEsperando_(d), volta: d && d.status === 'ESPERANDO' ? d.volta : '' });
  });
  return out;
}

// ---- CONTAGEM: contadas a decidir (aba CONTAGEM ESTOQUE) ----
function decContCartoes_() {
  return v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), CAB_CTA.length).map(ctaObj_).filter(function (o) { return o.status === 'CONTADA - DECIDIR'; }).map(function (o) { o.tipo = 'CONTAGEM'; o.chave = 'CTA|' + o.id; return o; });
}

// ---- API ----
function decLista(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  if (!p.fresco) { var hit = cacheGrandeLer_('DECL131'); if (hit) { hit.cache = 1; return hit; } } // v131: lista guardada 2 min (some ao gravar qualquer decisão)
  DEC_MEMO_ = {}; // v131
  var m = decMapa_(); try { if (decGondolaVencidas_(m)) m = decMapa_(); } catch (e) {} // v121 (v131: mapa lido uma vez)
  var r = { ok: true, aDecidir: [], esperando: [], decididas: [], erros: [] };
  [['REBAIXA', function () { return decRebCartoes_(m); }], ['PRODUTO_NOVO', function () { return decNovoCartoes_(m); }], ['CONTAGEM', decContCartoes_], ['CODIGO', function () { return decCodCartoes_(m); }]].forEach(function (t) {
    try { t[1]().forEach(function (c) { (c.esperando ? r.esperando : r.aDecidir).push(c); }); } catch (e) { r.erros.push(t[0] + ': ' + e.message); }
  });
  var lim = diaNum_(v9Hoje_()) - 15, pmc = {};
  Object.keys(m).forEach(function (k) { var d = m[k]; if (d.status !== 'ESPERANDO' && d.status !== 'PENDENTE' && diaNum_(String(d.em).slice(0, 10)) >= lim) { var det = {}; try { det = JSON.parse(d.detalhe || '{}'); } catch (e) {} var pm = { pode: false, motivo: '' }; try { pm = decPodeMudar_(k, d, pmc); } catch (e) {} r.decididas.push({ tipo: d.tipo, chave: k, status: d.status, decisao: d.decisao, det: det, por: d.por, em: d.em, exec: d.exec, obs: d.obs, pode: pm.pode, motivo: pm.motivo || '', fase: (d.status === 'EXECUTADA' || /^✓/.test(pm.motivo || '')) ? 'AJUSTADA' : 'RESPONDIDA' }); } });
  r.decididas.sort(function (x, y) { return String(y.em).slice(6, 10) + String(y.em).slice(3, 5) + String(y.em).slice(0, 2) + String(y.em).slice(11) > String(x.em).slice(6, 10) + String(x.em).slice(3, 5) + String(x.em).slice(0, 2) + String(x.em).slice(11) ? 1 : -1; });
  DEC_MEMO_ = null; // v131
  cacheGrande_('DECL131', r, 120); // v131
  return r;
}
function decDecidir(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio decide' };
  var tipo = String(p.tipo || ''), op = String(p.op || '').toUpperCase(), chave = String(p.chave || ''), obs = String(p.obs || '').trim();
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try { CacheService.getScriptCache().remove('DEC_RES_V113'); } catch (e) {}
  try {
    if (op === 'ESPERAR') {
      var dias = Math.max(1, Math.min(90, Number(p.dias) || 0)); if (!dias) return { ok: false, erro: 'escolha em quantos dias volta' };
      var volta = vsBr_(new Date(vsDia_(new Date()).getTime() + dias * 86400000));
      decGrava_(tipo, chave, 'ESPERANDO', volta, 'ESPERAR ' + dias + ' DIAS', { dias: dias }, a.d, obs);
      return { ok: true, msg: 'Volta para você em ' + volta };
    }
    if (tipo === 'REBAIXA') {
      var sh = decRebAba_(), c = { prod: String(p.prod || ''), validade: String(p.validade || '') }, r = decRebLinha_(sh, c); if (r < 0) return { ok: false, erro: 'não achei a linha na REBAIXAS (mudou?) — atualize a lista' };
      var k = String(sh.getRange(r, 11).getValue() || ''), antes = { e: sh.getRange(r, 5).getValue(), k: k, l: sh.getRange(r, 12).getValue() === true }; // v115: para poder desfazer
      if (op === 'FAZER') {
        var pr = decRs_(p.preco); if (!(pr > 0)) return { ok: false, erro: 'informe o preço da rebaixa' };
        sh.getRange(r, 5).setValue(pr); if (obs) sh.getRange(r, 11).setValue(k + ' - ' + obs);
        try { sh.getRange(r, 12).setValue(true); } catch (e) {}
        decGrava_(tipo, chave, 'DECIDIDA', '', 'FAZER REBAIXA', { preco: pr, prod: c.prod, validade: c.validade, antes: antes }, a.d, obs);
        return { ok: true, msg: 'Rebaixa decidida: R$ ' + String(pr.toFixed(2)).replace('.', ',') + ' — pronta para lançar' };
      }
      if (op === 'CONTAR') { // v119.7: diferença achado na loja x sistema -> contagem (Regiane)
        var sk = (k.match(/SKU (\d+)/) || [])[1]; if (!sk) return { ok: false, erro: 'SKU não encontrado na observação' };
        var jaC = v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), CAB_CTA.length).some(function (l) { return skuN_(l[3]) === sk && String(l[7]) === 'AGUARDANDO CONTAGEM'; });
        if (jaC) return { ok: true, msg: 'Já está aguardando contagem' };
        ctaPedir_(a.d, sk, c.prod, 'Regiane', 'diferença achado na loja x sistema (rebaixa)', 'ACHAR NA LOJA');
        return { ok: true, msg: 'Contagem pedida para a Regiane' };
      }
      if (op === 'GONDOLA') { // v121: conferir na gôndola primeiro
        var g = decGondolaCriar_(c.prod, c.validade, k), vt = vsBr_(new Date(vsDeBr_(c.validade).getTime() + 86400000));
        decGrava_(tipo, chave, 'ESPERANDO', vt, 'CONFERIR NA GONDOLA', { prod: c.prod, validade: c.validade, gondola: { chave: g.chave, pedido: v9Hoje_(), rep: g.rep, id: g.id, sku: g.sku }, antes: antes }, a.d, obs);
        return { ok: true, msg: 'Foi para a lista de ' + g.rep + ' conferir na gôndola' };
      }
      if (op === 'NAO FAZER') {
        sh.getRange(r, 11).setValue(k + ' - NAO FAZER (Celio ' + v9Hoje_().slice(0, 5) + (obs ? ': ' + obs : '') + ')');
        decGrava_(tipo, chave, 'DECIDIDA', '', 'NAO FAZER', { prod: c.prod, validade: c.validade, antes: antes }, a.d, obs);
        return { ok: true, msg: 'Marcado NÃO FAZER' };
      }
    }
    if (tipo === 'PRODUTO_NOVO') {
      var abas = criarAbaConf_(SpreadsheetApp.getActive()), idItem = chave.replace(/^NOVO\|/, ''), rr = v9Linha_(abas.it, 15, idItem); if (rr < 0) return { ok: false, erro: 'item não encontrado' };
      if (op === 'VINCULAR') {
        var s = skuN_(p.sku), cl = decCadLista_().filter(function (x) { return x[0] === s; })[0]; if (!cl) return { ok: false, erro: 'SKU ' + s + ' não está no cadastro' };
        var lv = abas.it.getRange(rr, 1, 1, 10).getValues()[0], antes = { c: lv[2], d: lv[3], i: lv[8], j: lv[9] }; // v115
        abas.it.getRange(rr, 3, 1, 2).setValues([["'" + s, cl[1]]]); abas.it.getRange(rr, 10).setValue('');
        if (Number(p.fator) > 1) abas.it.getRange(rr, 9).setValue(Number(p.fator));
        try { var lf = abas.it.getRange(rr, 1, 1, CAB_CIT.length + EXTRA_CIT.length).getValues()[0], cpv = String(lf[CAB_CIT.length + EXTRA_CIT.indexOf('CPROD')] || '').replace(/^'/, ''); if (cnpjDoIdNf_(lv[0]) && cpv) vincGravar_(cnpjDoIdNf_(lv[0]), cpv, String(lf[15] || ''), s, cl[1]); } catch (e) {} // v118.5: a decisão também ensina o vínculo para as próximas notas
        decGrava_(tipo, chave, 'EXECUTADA', '', 'VINCULAR', { sku: s, desc: cl[1], prod: String(p.prod || ''), idNf: String(lv[0]), antes: antes }, a.d, obs);
        return { ok: true, msg: 'Vinculado a ' + s + ' · ' + cl[1] };
      }
      if (op === 'CADASTRAR') {
        var desc = semAcento_(String(p.desc || '')).toUpperCase().trim(); if (!desc) return { ok: false, erro: 'informe a descrição' };
        decGrava_(tipo, chave, 'A EXECUTAR', '', 'CADASTRAR NOVO', { desc: /\*\*$/.test(desc) ? desc : desc + '**', un: String(p.un || 'UNIDADE'), entrada: String(p.entrada || ''), ean: String(p.ean || ''), prod: String(p.prod || ''), nf: String(p.nf || ''), forn: String(p.forn || ''), idNf: String(p.idNf || '') }, a.d, obs);
        return { ok: true, msg: 'Autorizado — o Claude cadastra na próxima rotina de NF' };
      }
    }
    if (tipo === 'CONTAGEM') {
      var mp = { 'SAIDA MA': 1, 'BAIXA NF': 1, 'ENTRADA MA': 1, 'NAO AJUSTAR': 1, 'RECONTAGEM': 1 }; if (!mp[op]) return { ok: false, erro: 'decisão inválida' };
      lock.releaseLock(); lock = null;
      var res = ctaDecidir({ token: p.token, quem: p.quem, id: chave.replace(/^CTA\|/, ''), decisao: op, para: p.para, _viaDec: 1 });
      if (res.ok) { try { var idc = chave.replace(/^CTA\|/, ''), shx = v9Aba_(ABA_CTA, CAB_CTA), rx = v9Linha_(shx, 1, idc); decGrava_('CONTAGEM', chave, 'DECIDIDA', '', op, { id: idc, prod: rx > 0 ? String(shx.getRange(rx, 5).getValue()) : '' }, a.d, obs); } catch (e) {} } // v115
      if (res.ok && obs) { var shc = v9Aba_(ABA_CTA, CAB_CTA), rc = v9Linha_(shc, 1, chave.replace(/^CTA\|/, '')); if (rc > 0) shc.getRange(rc, 26).setValue(obs); }
      return res.ok ? { ok: true, msg: op === 'RECONTAGEM' ? 'Recontagem pedida' : (op === 'NAO AJUSTAR' ? 'Sem ajuste' : 'Autorizado — o Claude lança na rotina da noite') } : res;
    }
    if (tipo === 'CODIGO') {
      var ops = { PROMO: 'CADASTRAR PROMO', TROCAR: 'TROCAR CODIGO', UVA: 'TROCAR CODIGO (UVA)', REBIPAR: 'PEDIR REBIPAGEM', IGNORAR: 'IGNORAR', GRAMATURA: 'CONFERIR GRAMATURA', NOVO_PESO: 'CADASTRAR NOVO (PESO)', CORRIGIR_PESO: 'CORRIGIR GRAMATURA' }; // v118 if (!ops[op]) return { ok: false, erro: 'decisão inválida' };
      var det = { sku: String(p.sku || ''), prod: String(p.prod || ''), bipado: chave.replace(/^COD\|/, ''), skuDestino: String(p.skuDestino || p.sku || ''), nf: String(p.nf || '') };
      if (op === 'GRAMATURA') { // v118: pede para alguém da loja conferir o peso na embalagem
        var paraG = String(p.para || '').trim(); if (!paraG) return { ok: false, erro: 'escolha quem vai conferir' };
        det.para = paraG; det.grmId = grmPedir_(chave, det, paraG, String(p.quemBipou || ''));
        decGrava_(tipo, chave, 'ESPERANDO', vsBr_(new Date(vsDia_(new Date()).getTime() + 7 * 86400000)), ops[op], det, a.d, obs);
        return { ok: true, msg: 'Enviado para ' + paraG + ' conferir — o cartão volta quando responder' };
      }
      if (op === 'NOVO_PESO' || op === 'CORRIGIR_PESO') {
        var dnv = semAcento_(String(p.desc || '')).toUpperCase().trim(); if (!dnv) return { ok: false, erro: 'informe a descrição' };
        det.desc = op === 'NOVO_PESO' && !/\*\*$/.test(dnv) ? dnv + '**' : dnv; det.pesoNovo = String(p.pesoNovo || ''); det.foto = String(p.foto || '');
      }
      if (op === 'REBIPAR') {
        decGrava_(tipo, chave, 'ESPERANDO', vsBr_(new Date(vsDia_(new Date()).getTime() + 2 * 86400000)), ops[op], det, a.d, obs);
        if (p.quemBipou) v9Push_(String(p.quemBipou), '📷 Bipar de novo', 'O Celio pediu para bipar de novo o código de: ' + det.prod + (det.nf ? ' (NF ' + det.nf + ')' : ''), 'aviso');
        return { ok: true, msg: 'Pedido enviado para ' + (p.quemBipou || 'a conferente') };
      }
      decGrava_(tipo, chave, op === 'IGNORAR' ? 'DECIDIDA' : 'A EXECUTAR', '', ops[op], det, a.d, obs);
      return { ok: true, msg: op === 'IGNORAR' ? 'Ignorado' : 'Autorizado — o Claude faz no Plenus na próxima rotina de NF' };
    }
    return { ok: false, erro: 'decisão inválida' };
  } finally { if (lock) lock.releaseLock(); }
}
function decExecutar(p) { // tarefa do Claude: LISTAR o que falta fazer no Plenus (cadastro novo, código) / FEITO (chave, obs)
  if (String(p.chave || '') !== CHAVE_CARGA) return { ok: false, erro: 'chave inválida' };
  var sh = decAba_(), m = decMapa_();
  if (String(p.op || '') === 'FEITO') {
    var d = m[String(p.id || '')]; if (!d) return { ok: false, erro: 'não encontrada' };
    sh.getRange(d.row, 5).setValue('EXECUTADA'); sh.getRange(d.row, 11, 1, 2).setValues([[agoraBr_(), String(p.obs || d.obs || '')]]);
    return { ok: true };
  }
  PropertiesService.getScriptProperties().setProperty('DEC_EXEC_EM', String(Date.now())); // v115: trava o "mudar decisão" por 3 h
  var out = []; Object.keys(m).forEach(function (k) { var d = m[k]; if (d.status === 'A EXECUTAR') { var det = {}; try { det = JSON.parse(d.detalhe); } catch (e) {} out.push({ chave: k, tipo: d.tipo, decisao: d.decisao, det: det, por: d.por, em: d.em, obs: d.obs }); } });
  return { ok: true, itens: out };
}
function decResumo_() { // selo da Gestão: total a decidir e quantas rebaixas vencem em até 3 dias (cache 5 min)
  var cc = CacheService.getScriptCache(), em = cc.get('DEC_RES_V113'); if (em) return JSON.parse(em);
  var m = decMapa_(), n = 0, urg = 0; DEC_MEMO_ = {}; // v131
  try { decRebCartoes_(m).forEach(function (c) { if (!c.esperando) { n++; if (c.dias <= 3) urg++; } }); } catch (e) {}
  try { n += decContCartoes_().length; } catch (e) {}
  try { n += decNovoCartoes_(m).filter(function (c) { return !c.esperando; }).length; } catch (e) {}
  try { n += decCodCartoes_(m, true).filter(function (c) { return !c.esperando; }).length; } catch (e) {}
  DEC_MEMO_ = null; // v131
  var r = { n: n, urg: urg }; try { cc.put('DEC_RES_V113', JSON.stringify(r), 300); } catch (e) {}
  return r;
}
// teste (não grava): Logger com o resumo de cada tipo
function decTeste() { var m = decMapa_(); var r = { reb: decRebCartoes_(m).slice(0, 5), novo: decNovoCartoes_(m).slice(0, 5), cod: decCodCartoes_(m).slice(0, 5), cont: decContCartoes_() }; Logger.log(JSON.stringify(r)); return r; }

// ========== v114 — ⏱️ TAREFAS AGENDADAS (Gestão, só o Celio) — modelo aprovado em 10/11/2026 ==========
// A aba AGENDAMENTOS é gravada pela tarefa "Vigia dos agendamentos" (o app não enxerga as tarefas direto).
var ABA_AGD = 'AGENDAMENTOS';
var CAB_AGD = ['ID', 'NOME', 'AUT', 'QUANDO', 'POR QUE', 'FAZ', 'PROXIMA', 'STATUS', 'RESUMO', 'ULTIMA', 'CHAT', 'HIST', 'ATUALIZADO', 'ATIVA', 'ORDEM', 'ETAPAS'];
// AUT: SEMPRE (🔴) | SE TIVER ACAO (🟡) | SEM (🟢) | UNICA (📅)
// STATUS: OK | PARCIAL | FALHOU | NAO RODOU | ESPERANDO | SEM RECADO | RODANDO
function agdTxt_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(v == null ? '' : v); }
function agdItens_() {
  var sh = v9Aba_(ABA_AGD, CAB_AGD);
  return v9Linhas_(sh, CAB_AGD.length).filter(function (l) { return String(l[0]).trim() && String(l[13]).toUpperCase() !== 'N'; }).map(function (l) {
    var h = []; try { h = JSON.parse(String(l[11] || '[]')); } catch (e) {}
    var et = []; try { et = JSON.parse(String(l[15] || '[]')); } catch (e) { et = String(l[15] || '').split(/\n/).filter(String); }
    return { id: String(l[0]), nome: String(l[1]), aut: String(l[2]).toUpperCase(), quando: String(l[3]), porque: String(l[4]), faz: String(l[5]), proxima: agdTxt_(l[6]), status: String(l[7]).toUpperCase(), resumo: String(l[8]), ultima: agdTxt_(l[9]), chat: String(l[10]), hist: h, atualizado: agdTxt_(l[12]), ordem: Number(l[14]) || 99, etapas: et, link: 'https://claude.ai/scheduled-task/' + String(l[0]).trim() };
  }).sort(function (a, b) { return a.ordem - b.ordem || (a.nome > b.nome ? 1 : -1); });
}
function agdLista(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var it = agdItens_(), at = '';
  var k = function (t) { t = String(t); return t.slice(6, 10) + t.slice(3, 5) + t.slice(0, 2) + t.slice(11); };
  it.forEach(function (o) { if (k(o.atualizado) > k(at)) at = o.atualizado; });
  return { ok: true, itens: it, atualizado: at, lista: 'https://claude.ai/code/routines' };
}
function agdResumo_() { // selo da Gestão: esperando você + erro/não rodou (cache 5 min)
  var cc = CacheService.getScriptCache(), em = cc.get('AGD_RES_V114'); if (em) return JSON.parse(em);
  var r = { esp: 0, err: 0, tot: 0 };
  try { agdItens_().forEach(function (o) { r.tot++; if (o.status === 'ESPERANDO') r.esp++; else if (o.status === 'FALHOU' || o.status === 'NAO RODOU') r.err++; }); } catch (e) {}
  try { cc.put('AGD_RES_V114', JSON.stringify(r), 300); } catch (e) {}
  return r;
}

function decTeste2() { var m = decMapa_(); Object.keys(m).forEach(function (k) { Logger.log(k + ' | ' + m[k].status + ' | volta ' + m[k].volta + ' | esperando ' + decEsperando_(m[k]) + ' | em ' + m[k].em + ' | mudar ' + JSON.stringify(decPodeMudar_(k, m[k], {}))); }); }
function agdTeste() { var r = agdItens_(); Logger.log(r.length + ' tarefas'); Logger.log(JSON.stringify(r[3])); Logger.log(JSON.stringify(agdResumo_())); }

// ========== v115 — ↩️ MUDAR DECISÃO (Decisões): desfaz enquanto não foi feito no Plenus — aprovado em 10/11/2026 ==========
var DEC_JANELA_EXEC = 3 * 3600000; // 3 h depois que a rotina pegou a lista, fica "em execução"
function decEmExec_(nome) { var t = Number(PropertiesService.getScriptProperties().getProperty(nome) || 0); return t && (Date.now() - t) < DEC_JANELA_EXEC; }
function decPodeMudar_(chave, d, cache) { // { pode: bool, motivo: texto }
  var det = {}; try { det = JSON.parse(d.detalhe || '{}'); } catch (e) {}
  if (d.status === 'EXECUTADA' && d.tipo !== 'PRODUTO_NOVO') return { pode: false, motivo: '✓ feito' + (d.exec ? ' em ' + String(d.exec).slice(0, 5) : '') };
  if (d.status === 'ESPERANDO') return { pode: true };
  if (d.tipo === 'REBAIXA') {
    if (!det.antes) return { pode: false, motivo: 'decidida antes do botão — desfazer com o Claude' };
    try { // v131: a aba REBAIXAS é lida uma vez só para todas as decisões
      if (!cache.rebIx) { var sh = decRebAba_(), nr = sh.getLastRow(); cache.rebIx = {}; if (nr >= 2) sh.getRange(2, 1, nr - 1, 10).getValues().forEach(function (v) { var kk = String(v[1]).trim() + '|' + v9Br_(v[9]); if (!(kk in cache.rebIx)) cache.rebIx[kk] = v[0] === true; }); }
      var kr = String(det.prod) + '|' + String(det.validade); if (!(kr in cache.rebIx)) return { pode: false, motivo: 'linha não encontrada na REBAIXAS' };
      if (cache.rebIx[kr]) return { pode: false, motivo: '✓ já lançada no Plenus' }; } catch (e) { return { pode: false, motivo: 'não consegui conferir a REBAIXAS' }; }
    return { pode: true };
  }
  if (d.tipo === 'PRODUTO_NOVO') {
    if (d.status === 'A EXECUTAR') return decEmExec_('DEC_EXEC_EM') ? { pode: false, motivo: '🔄 em execução na rotina de NF' } : { pode: true };
    if (d.status === 'EXECUTADA' && d.decisao !== 'VINCULAR') return { pode: false, motivo: '✓ feito' + (d.exec ? ' em ' + String(d.exec).slice(0, 5) : '') };
    if (!det.antes) return { pode: false, motivo: 'vinculado antes do botão — desfazer com o Claude' };
    if (!cache.nfEnt) { cache.nfEnt = {}; v9Linhas_(criarAbaConf_(SpreadsheetApp.getActive()).nf, 5).forEach(function (v) { cache.nfEnt[String(v[0])] = String(v[4]).trim(); }); } // v131
    if (cache.nfEnt[String(det.idNf || '')]) return { pode: false, motivo: '✓ a nota já entrou no Plenus' };
    return { pode: true };
  }
  if (d.tipo === 'CONTAGEM') {
    if (!cache.ctaSt) { cache.ctaSt = {}; v9Linhas_(v9Aba_(ABA_CTA, CAB_CTA), 8).forEach(function (v) { cache.ctaSt[String(v[0])] = String(v[7]); }); } // v131
    if (!(String(det.id || '') in cache.ctaSt)) return { pode: false, motivo: 'contagem não encontrada' };
    var st = cache.ctaSt[String(det.id || '')];
    if (st === 'AJUSTADO') return { pode: false, motivo: '✓ ajustado no Plenus' };
    if (st === 'RECONTAGEM') return { pode: false, motivo: 'recontagem já pedida' };
    if (st === 'AUTORIZADO - A LANCAR' && decEmExec_('CTA_LANCAR_EM')) return { pode: false, motivo: '🔄 em execução nas Vendas diárias' };
    return st === 'AUTORIZADO - A LANCAR' || st === 'NAO AJUSTADO' ? { pode: true } : { pode: false, motivo: 'situação: ' + st };
  }
  if (d.tipo === 'CODIGO') {
    if (d.status === 'A EXECUTAR' && decEmExec_('DEC_EXEC_EM')) return { pode: false, motivo: '🔄 em execução na rotina de NF' };
    return { pode: true };
  }
  return { pode: d.status !== 'EXECUTADA' };
}
// ---- v121 (Celio 14/11/2026): CONFERIR NA GÔNDOLA PRIMEIRO (cartão de rebaixa) ----
// Cria/usa a linha da CV SEMANA (mesma CHAVE da Central: rep|30|produto|validade); a decisão fica ESPERANDO até a repositora conferir.
function decGondolaCriar_(prod, val, k) {
  var ssV = SpreadsheetApp.openById(ID_VALIDADE), cad = cvCad_(ssV), pk = normDesc_(String(prod).replace(/\*+\s*$/, ''));
  var sku = (String(k || '').match(/SKU (\d+)/) || [])[1] || ''; if (!sku) { var sm = vsSkuPorNome_()[pk]; sku = sm ? String(sm.sku) : ''; }
  var c = cadSku_(cad, sku, prod) || { secao: '', local: '', rep: '' }, sem = !c.secao, rep = sem ? 'Ney' : (c.rep || 'Ney');
  var chave = rep + '|30|' + pk + '|' + val, semana = vsBr_(vsSegunda_(new Date())), sh = cvAba_(ABA_CVS, CAB_CVS), n = sh.getLastRow(), id = '';
  if (n >= 2) sh.getRange(2, 1, n - 1, 13).getValues().forEach(function (l) { if (!id && String(l[0]) === semana && String(l[11]) === chave && String(l[10]) === 'PENDENTE') id = String(l[12]); });
  if (!id) { id = Utilities.getUuid(); sh.getRange(sh.getLastRow() + 1, 1, 1, CAB_CVS.length).setValues([[semana, rep, '30', val, prod, sku, c.secao, c.local, c.rep, '', 'PENDENTE', chave, id, sem ? 'S' : '', '', '', '', '', '', '', '']].map(function (r) { return r.map(String); })); }
  try { v9Push_(rep, '🔎 Conferir na gôndola', String(prod) + ' · validade ' + String(val).slice(0, 5) + ' — conte área e estoque (Validades › Conferência).', 'validades'); } catch (e) {}
  return { chave: chave, rep: rep, id: id, sku: sku };
}
function decGondolaSuf_(ch) { return String(ch || '').split('|').slice(1).join('|'); }
// Chamado pelo cvSalvar: se a linha conferida é de um pedido "conferir na gôndola", devolve a decisão ou fecha como NAO FAZER.
function decGondolaRetorno_(l, x) {
  var m = decMapa_(), ch = String(l[11]), suf = decGondolaSuf_(ch), achou = null;
  Object.keys(m).forEach(function (key) {
    var d = m[key]; if (achou || d.tipo !== 'REBAIXA' || d.status !== 'ESPERANDO' || d.decisao !== 'CONFERIR NA GONDOLA') return;
    var det = {}; try { det = JSON.parse(d.detalhe || '{}'); } catch (e) {}
    var gc = det.gondola && det.gondola.chave; if (gc && (gc === ch || decGondolaSuf_(gc) === suf)) achou = { key: key, d: d, det: det };
  });
  if (!achou) return null;
  var det = achou.det, quem = x.quem, dm = v9Hoje_().slice(0, 5), em = agoraBr_();
  if (!x.naoTem && x.total > 0 && !(x.vd && x.vd.caso === 'NOVA')) {
    det.gondola.res = { quem: quem, em: em, total: x.total, area: x.area, est: x.est == null ? '' : x.est, val: x.val };
    decGrava_('REBAIXA', achou.key, 'PENDENTE', '', 'CONFERIDO NA GONDOLA', det, achou.d.por, '');
  } else {
    var mot = 'não tem essa validade na gôndola (' + quem + ' ' + dm + ')';
    det.gondola.res = { quem: quem, em: em, naoTem: true, achou: x.vd && x.vd.caso === 'NOVA' ? x.vd.nova : '' };
    try { var sh = decRebAba_(), r = decRebLinha_(sh, { prod: det.prod, validade: det.validade }); if (r > 0) { var k = String(sh.getRange(r, 11).getValue() || ''); sh.getRange(r, 11).setValue(k + ' - NAO FAZER REBAIXA - NAO TEM ESSA VALIDADE NA GONDOLA (' + quem + ' ' + dm + ')'); } } catch (e) {}
    decGrava_('REBAIXA', achou.key, 'DECIDIDA', '', 'NAO FAZER', det, achou.d.por, mot);
  }
  try { if (x.cvRow > 0) cvAba_(ABA_CV, CAB_CV).getRange(x.cvRow, 17).setValue('REB-GONDOLA'); } catch (e) {}
  try { CacheService.getScriptCache().remove('DEC_RES_V113'); } catch (e) {}
  return achou.key;
}
// Validade passou e ninguém conferiu: fecha a decisão.
function decGondolaVencidas_(m) {
  m = m || decMapa_(); var hoje = diaNum_(v9Hoje_()), n = 0;
  Object.keys(m).forEach(function (key) {
    var d = m[key]; if (d.tipo !== 'REBAIXA' || d.status !== 'ESPERANDO' || d.decisao !== 'CONFERIR NA GONDOLA') return;
    var det = {}; try { det = JSON.parse(d.detalhe || '{}'); } catch (e) {}
    if (!(diaNum_(det.validade) < hoje)) return;
    decGrava_('REBAIXA', key, 'DECIDIDA', '', 'VENCEU SEM CONFERENCIA', det, d.por, 'venceu sem conferência'); n++;
  });
  return n;
}
// v121 (fila): situação da rebaixa dos conferidos, calculada na hora (o app pede em segundo plano)
function cvRebSit(p) {
  var dono = quemDoToken_(p.token); if (!dono || dono !== String(p.quem || '')) return { ok: false, erro: 'sessão expirada', reauth: true };
  var its = cvLista_(dono).itens.filter(function (i) { return i.status === 'CONFERIDO' && (i.critica === 'SUGERIR REBAIXA' || /^JA NA REBAIXA/.test(i.critica)); }), out = {};
  if (!its.length) return { ok: true, sit: out };
  var idx = rebIndice_();
  its.forEach(function (i) {
    var s = null; try { s = rebSituacao_(idx, i.sku, i.desc, i.validade); } catch (e) {}
    out[i.id] = s ? { t: s.tipo, por: s.por || '' } : { t: '' };
  });
  return { ok: true, sit: out };
}
function decGondolaInfo_(d) { // v121: o que mostrar no cartão
  if (!d || !/GONDOLA/.test(d.decisao)) return null; var det = {}; try { det = JSON.parse(d.detalhe || '{}'); } catch (e) {}
  var g = det.gondola; if (!g) return null;
  return { rep: g.rep || '', pedido: g.pedido || '', esperando: d.status === 'ESPERANDO', res: d.status === 'PENDENTE' ? (g.res || null) : null };
}
function decDesfazer(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var chave = String(p.chave || ''), lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var m = decMapa_(), d = m[chave]; if (!d) return { ok: false, erro: 'decisão não encontrada (atualize a lista)' };
    var pm = decPodeMudar_(chave, d, {}); if (!pm.pode) return { ok: false, erro: 'não dá para mudar: ' + pm.motivo };
    var det = {}; try { det = JSON.parse(d.detalhe || '{}'); } catch (e) {}
    if (d.tipo === 'REBAIXA' && d.status !== 'ESPERANDO' && det.antes) {
      var sh = decRebAba_(), r = decRebLinha_(sh, { prod: det.prod, validade: det.validade });
      if (r > 0) { sh.getRange(r, 5).setValue(det.antes.e); sh.getRange(r, 11).setValue(det.antes.k); try { sh.getRange(r, 12).setValue(det.antes.l === true); } catch (e) {} }
    }
    if (d.tipo === 'PRODUTO_NOVO' && d.decisao === 'VINCULAR' && det.antes) {
      var abas = criarAbaConf_(SpreadsheetApp.getActive()), rr = v9Linha_(abas.it, 15, chave.replace(/^NOVO\|/, ''));
      if (rr > 0) { abas.it.getRange(rr, 3, 1, 2).setValues([[det.antes.c, det.antes.d]]); abas.it.getRange(rr, 9).setValue(det.antes.i); abas.it.getRange(rr, 10).setValue(det.antes.j); }
    }
    if (d.tipo === 'CONTAGEM' && d.status !== 'ESPERANDO') {
      var shc = v9Aba_(ABA_CTA, CAB_CTA), rc = v9Linha_(shc, 1, det.id || '');
      if (rc > 0) { shc.getRange(rc, 8).setValue('CONTADA - DECIDIR'); shc.getRange(rc, 21, 1, 2).setValues([['', '']]); }
    }
    if (d.tipo === 'CODIGO' && d.status === 'ESPERANDO' && d.decisao === 'CONFERIR GRAMATURA') grmCancelar_(chave); // v118
    if (String(p.corrigir || '') === '1') decCorrPoe_(chave, { de: d.decisao + (det.preco ? ' R$ ' + Number(det.preco).toFixed(2).replace('.', ',') : ''), deEm: d.em }); // v127: Corrigir resposta
    decAba_().deleteRow(d.row);
    try { CacheService.getScriptCache().remove('DEC_RES_V113'); } catch (e) {}
    return { ok: true, msg: 'Decisão desfeita — o cartão voltou para A DECIDIR' };
  } finally { lock.releaseLock(); }
}

// ========== v116 — ORIGEM no cartão de rebaixa (Decisões) — modelo aprovado em 11/11/2026 ==========
function decNorm_(s) { return semAcento_(String(s || '')).toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim(); }
function decDm_(v) { var b = v9Br_(v); return b ? b.slice(0, 5) : ''; }
function decOrgBase_(c) { // índices de CONFERENCIA VALIDADE e VALIDADE NA FRENTE (1x por lista)
  if (c.cv) return c;
  c.cv = {}; c.vf = {};
  try { var sh = SpreadsheetApp.getActive().getSheetByName(ABA_CV), n = sh ? sh.getLastRow() : 0;
    if (n >= 2) sh.getRange(2, 1, n - 1, 10).getValues().forEach(function (l) { var key = decNorm_(l[4]) + '|' + v9Br_(l[5]); c.cv[key] = { quem: String(l[2] || ''), data: decDm_(l[0]), area: l[6], dep: l[7], depNao: /NAO CONFERIDO/i.test(String(l[9])) }; }); } catch (e) {}
  try { var sv = SpreadsheetApp.getActive().getSheetByName(ABA_VNF), m = sv ? sv.getLastRow() : 0;
    if (m >= 2) sv.getRange(2, 1, m - 1, 15).getValues().forEach(function (l) { if (!String(l[13] || '').trim()) return; c.vf[decNorm_(l[3]) + '|' + decNorm_(l[13]).split(' ')[0]] = { qtd: l[12], data: decDm_(l[14] || l[0]) }; }); } catch (e) {}
  return c;
}
function decRebOrigem_(prod, val, k, c) {
  var n = function (re) { var x = k.match(re); return x ? Number(x[1]) : null; };
  var lj = k.match(/(\d+)\s*UN ACHADAS NA LOJA \(gondola ([\d.,]+) \+ deposito ([\d.,]+)\) por (.+?) em (\d{2}\/\d{2}) - sistema (-?[\d.,]+)/i); // v119.7
  if (lj) { var nb = function (x) { return Number(String(x).replace(',', '.')); }; return { t: 'LOJA', total: Number(lj[1]), gon: nb(lj[2]), dep: nb(lj[3]), quem: lj[4].trim(), data: lj[5], sis: nb(lj[6]) }; }
  if (/\(sobra\)/i.test(k)) { // contagem da conferência de validade (lista 30/45 dias)
    decOrgBase_(c); var r = c.cv[decNorm_(prod) + '|' + val];
    if (r) return { t: 'CONTAGEM', quem: r.quem, data: r.data, area: r.area === '' ? null : Number(r.area), dep: r.depNao || r.dep === '' ? null : Number(r.dep) };
    var est = /estoque 0/i.test(k) ? 0 : n(/(\d+)\s*UN ESTOQUE/i);
    return { t: 'CONTAGEM', quem: '', data: '', area: n(/(\d+)\s*UN AREA/i), dep: /nao conferido/i.test(k) ? null : est };
  }
  var ach = k.match(/achado na loja por ([^)]+)\)/i);
  if (ach) { decOrgBase_(c); var quem = ach[1].trim(), f = c.vf[decNorm_(prod) + '|' + decNorm_(quem).split(' ')[0]] || c.cv[decNorm_(prod) + '|' + val] || {}; if (f.qtd === undefined && f.area !== undefined) f = { data: f.data };
    return { t: 'ACHADO', quem: quem, qtd: f.qtd !== undefined && f.qtd !== '' ? Number(f.qtd) : n(/^(\d+)\s*UN/i), data: f.data || '' }; }
  if (/SALDO SISTEMA/i.test(k)) {
    var lote = /LOTE CONFERIDO/i.test(k) ? 'lote conferido' : (/CONFERIR NA GONDOLA/i.test(k) ? 'conferir na gôndola' : (n(/(\d+)\s*UN DESTA VALIDADE/i) != null ? n(/(\d+)\s*UN DESTA VALIDADE/i) + ' un desta validade' : ''));
    return { t: 'SISTEMA', saldo: n(/(\d+)\s*UN SALDO SISTEMA/i) != null ? n(/(\d+)\s*UN SALDO SISTEMA/i) : n(/saldo sistema (\d+)/i), lote: lote };
  }
  return { t: 'LANCADO' };
}

function decTeste3() { var o = decRebCartoes_(decMapa_()); var c = {}; o.forEach(function (x) { var t = x.org ? x.org.t : '-'; c[t] = (c[t] || 0) + 1; }); Logger.log(JSON.stringify(c)); ['CONTAGEM', 'ACHADO', 'SISTEMA', 'LANCADO'].forEach(function (t) { o.filter(function (y) { return y.org && y.org.t === t; }).slice(0, 2).forEach(function (x) { Logger.log(x.prod + ' => ' + JSON.stringify(x.org)); }); }); }

// ========== v118 — 📏 CONFERIR GRAMATURA NA LOJA (cartão Código das Decisões) — modelo aprovado em 11/11/2026 ==========
// O Celio pede, no cartão de código bipado diferente, para alguém da loja olhar o peso na embalagem.
// A pessoa responde no app (📏 Conferir na loja): mesmo peso / peso diferente (novo peso + foto) / não achei.
var ABA_GRM = 'CONFERIR LOJA';
var CAB_GRM = ['ID', 'CHAVE', 'SKU', 'PRODUTO', 'GRAMATURA CAD', 'BIPADO', 'NF', 'PARA', 'PEDIDO EM', 'STATUS', 'RESPOSTA', 'PESO NOVO', 'FOTO', 'RESPONDIDO POR', 'RESPONDIDO EM', 'QUEM BIPOU'];
// STATUS: PENDENTE | RESPONDIDO | CANCELADO   RESPOSTA: MESMO | DIFERENTE | NAO ACHEI
var GRM_RE = /(\d+(?:[.,]\d+)?)\s*(KG|GRS|GR|G|ML|LT|L)(?![A-Z])/g;
function grmPeso_(desc) { // último peso escrito no nome: '42 g', '1 kg', '500 ml', '2 L'
  var s = semAcento_(String(desc || '')).toUpperCase().replace(/\*+/g, ' '), m, u = null;
  GRM_RE.lastIndex = 0; while ((m = GRM_RE.exec(s))) u = m;
  if (!u) return '';
  var un = { KG: 'kg', GRS: 'g', GR: 'g', G: 'g', ML: 'ml', LT: 'L', L: 'L' }[u[2]];
  return u[1].replace('.', ',') + ' ' + un;
}
function grmDescNova_(desc, peso, un) { // troca o peso no nome pelo novo (padrão Pauvan: 40G, 1KG, 500ML, 2L)
  var s = String(desc || '').replace(/\*+\s*$/, '').trim(), nv = String(peso).replace('.', ',') + String(un || 'g').toUpperCase(), m, u = null;
  GRM_RE.lastIndex = 0; var up = semAcento_(s).toUpperCase(); while ((m = GRM_RE.exec(up))) u = m;
  s = u ? s.slice(0, u.index) + nv + s.slice(u.index + u[0].length) : s + ' ' + nv;
  return semAcento_(s).toUpperCase().replace(/\s+/g, ' ').trim() + '**';
}
function grmAba_() { return v9Aba_(ABA_GRM, CAB_GRM); }
function grmMapa_() { // chave -> último pedido
  var m = {}; v9Linhas_(grmAba_(), CAB_GRM.length).forEach(function (l, i) {
    m[String(l[1])] = { row: i + 2, id: String(l[0]), sku: String(l[2]), prod: String(l[3]), cad: String(l[4]), bipado: String(l[5]).replace(/^'/, ''), nf: String(l[6]), para: String(l[7]), pedEm: agdTxt_(l[8]), status: String(l[9]), resp: String(l[10]), peso: String(l[11]), foto: String(l[12]), por: String(l[13]), em: agdTxt_(l[14]), quemBipou: String(l[15]) };
  });
  return m;
}
function grmPedir_(chave, det, para, quemBipou) {
  var sh = grmAba_(), m = grmMapa_(), ant = m[chave];
  if (ant && ant.status === 'PENDENTE') sh.getRange(ant.row, 10).setValue('CANCELADO');
  var id = Utilities.getUuid();
  sh.appendRow([id, chave, "'" + det.sku, det.prod, grmPeso_(det.prod), "'" + det.bipado, det.nf || '', para, agoraBr_(), 'PENDENTE', '', '', '', '', '', quemBipou || '']);
  v9Push_(para, '📏 Conferir na loja', 'O Celio pediu para conferir o peso na embalagem de: ' + det.prod + (det.nf ? ' (NF ' + det.nf + ')' : ''), 'conferirloja');
  return id;
}
function grmCancelar_(chave) { try { var m = grmMapa_(), g = m[chave]; if (g && g.status === 'PENDENTE') grmAba_().getRange(g.row, 10).setValue('CANCELADO'); } catch (e) {} }
function grmPendQuem_(quem) { // selo da tela inicial
  var n = 0;
  v9Linhas_(grmAba_(), 10).forEach(function (l) { if (String(l[9]) === 'PENDENTE' && tfMeu_(l[7], quem)) n++; });
  return n;
}
function grmMinhas(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro;
  var cel = ehCelio_(a.d), hoje = v9Hoje_(), pend = [], env = [];
  v9Linhas_(grmAba_(), CAB_GRM.length).forEach(function (l) {
    var st = String(l[9]), para = String(l[7]); if (!tfMeu_(para, a.d) && !cel) return;
    var o = { id: String(l[0]), sku: String(l[2]).replace(/^'/, ''), prod: String(l[3]), peso: String(l[4]), un: (String(l[4]).split(' ')[1] || 'g'), bipado: String(l[5]).replace(/^'/, ''), nf: String(l[6]), para: para, pedEm: agdTxt_(l[8]).slice(0, 16) };
    if (st === 'PENDENTE') pend.push(o);
    else if (st === 'RESPONDIDO' && agdTxt_(l[14]).indexOf(hoje) === 0 && (tfMeu_(String(l[13]), a.d) || cel)) { o.resp = String(l[10]); o.pesoNovo = String(l[11]); env.push(o); }
  });
  return { ok: true, pend: pend, enviados: env.reverse() };
}
function grmResponder(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro;
  var resp = String(p.resp || '').toUpperCase(); if (['MESMO', 'DIFERENTE', 'NAO ACHEI'].indexOf(resp) < 0) return { ok: false, erro: 'escolha a resposta' };
  var peso = '', un = String(p.un || 'g');
  if (resp === 'DIFERENTE') { var pn = v9Num_(p.peso); if (!(pn > 0)) return { ok: false, erro: 'informe o peso que está na embalagem' }; if (!/^data:image\//.test(String(p.foto || ''))) return { ok: false, erro: 'tire a foto da embalagem' }; peso = String(pn).replace('.', ',') + ' ' + (['g', 'kg', 'ml', 'L'].indexOf(un) >= 0 ? un : 'g'); }
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var sh = grmAba_(), r = v9Linha_(sh, 1, p.id); if (r < 0) return { ok: false, erro: 'pedido não encontrado' };
    var l = sh.getRange(r, 1, 1, CAB_GRM.length).getValues()[0];
    if (String(l[9]) !== 'PENDENTE') return { ok: false, erro: String(l[9]) === 'CANCELADO' ? 'o Celio cancelou esse pedido' : 'esse pedido já foi respondido' };
    if (!tfMeu_(l[7], a.d) && !ehCelio_(a.d)) return { ok: false, erro: 'esse pedido é de outra pessoa' };
    var url = '';
    if (/^data:image\//.test(String(p.foto || ''))) {
      var mm = String(p.foto).match(/^data:(image\/\w+);base64,(.*)$/);
      if (mm) {
        var base = pastaApp_(SpreadsheetApp.getActive()), it = base.getFoldersByName('CONFERIR LOJA'), pasta = it.hasNext() ? it.next() : base.createFolder('CONFERIR LOJA');
        var nome = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '_' + String(l[2]).replace(/^'/, '') + '_' + String(l[3]).replace(/[^A-Za-z0-9]+/g, '-').slice(0, 40) + '.jpg';
        var f = pasta.createFile(Utilities.newBlob(Utilities.base64Decode(mm[2]), mm[1], nome));
        try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
        url = 'https://drive.google.com/thumbnail?sz=w600&id=' + f.getId();
      }
    }
    sh.getRange(r, 10, 1, 6).setValues([['RESPONDIDO', resp, peso, url, a.d, agoraBr_()]]);
    // o cartão volta para A DECIDIR com a resposta
    var chave = String(l[1]), dm = decMapa_()[chave];
    if (dm && dm.status === 'ESPERANDO') decAba_().getRange(dm.row, 6).setValue(v9Hoje_());
    try { CacheService.getScriptCache().remove('DEC_RES_V113'); } catch (e) {}
    var tx = resp === 'MESMO' ? 'mesmo peso (' + String(l[4] || 'do cadastro') + ')' : (resp === 'DIFERENTE' ? 'PESO DIFERENTE: ' + peso + ' (cadastro ' + String(l[4] || '?') + ')' : 'não achou na loja');
    v9Push_('Celio', '📏 Gramatura conferida — decidir', String(l[3]) + ': ' + tx + ' · ' + a.d, 'decisoes');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// v118: pedidos de conferência de gramatura feitos pelo Claude a pedido do Celio (11/11/2026) — rodar 1 vez pelo editor
function grmPedirClaude() {
  var volta = vsBr_(new Date(vsDia_(new Date()).getTime() + 7 * 86400000)), out = [];
  [['7896423439014', '39201', 'CHOCOLATE SNICKERS DARK 42G', 'NORTE-SUL 2780511', 'Regiane'],
   ['7896423439038', '45267', 'CHOCOLATE SNICKERS BRANCO 42G', 'NORTE-SUL 2780511', 'Regiane'],
   ['896994801958', '3351', 'CERA LIQ POLYLAR AMARELA 750ML', 'IRMAOS DOMINGOS 1310746', 'Regiane']].forEach(function (x) {
    var chave = 'COD|' + x[0], det = { sku: x[1], prod: x[2], bipado: x[0], skuDestino: x[1], nf: x[3], para: 'Regiane' };
    det.grmId = grmPedir_(chave, det, 'Regiane', x[4]);
    decGrava_('CODIGO', chave, 'ESPERANDO', volta, 'CONFERIR GRAMATURA', det, 'Celio', 'pedido pelo Claude a pedido do Celio');
    out.push(x[2] + ' -> ' + det.grmId);
  });
  try { CacheService.getScriptCache().remove('DEC_RES_V113'); } catch (e) {}
  Logger.log(out.join(' | ')); return out;
}

// ========== v118.4 — CONCLUIR CONFERÊNCIA RESPONDE NA HORA (12/11/2026) ==========
// Validades, e-mail e mudança do XML (o que deixava o concluir com ~1 min e às vezes estourava 6 min)
// rodam em segundo plano: a fila fica em POSCONC_FILA e um gatilho de uma vez processa ~20 s depois.
function posConcEnfileirar_(p, extra) {
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) {}
  try {
    var pr = PropertiesService.getScriptProperties(), f = []; try { f = JSON.parse(pr.getProperty('POSCONC_FILA') || '[]'); } catch (e) { f = []; }
    var q = {}; Object.keys(p || {}).forEach(function (k) { if (k !== 'token' && k !== 'lotes' && k !== 'foto') q[k] = p[k]; });
    f.push({ p: q, x: extra || {}, em: Date.now() }); pr.setProperty('POSCONC_FILA', JSON.stringify(f));
  } finally { try { lock.releaseLock(); } catch (e) {} }
  posConcAgendar_();
}
function posConcAgendar_() {
  try { var tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'posConclusaoFila'; }); if (!tem) ScriptApp.newTrigger('posConclusaoFila').timeBased().after(20 * 1000).create(); } catch (e) { console.log('posConc agendar erro ' + e); }
}
function posConclusaoFila() {
  try { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'posConclusaoFila') ScriptApp.deleteTrigger(t); }); } catch (e) {}
  var pr = PropertiesService.getScriptProperties(), t0 = Date.now();
  while (Date.now() - t0 < 4 * 60000) {
    var job = null, lock = LockService.getScriptLock();
    try { lock.waitLock(30000); var f = JSON.parse(pr.getProperty('POSCONC_FILA') || '[]'); job = f.shift() || null; pr.setProperty('POSCONC_FILA', JSON.stringify(f)); } catch (e) { job = null; } finally { try { lock.releaseLock(); } catch (e) {} }
    if (!job) return;
    try { posConcRodar_(job); } catch (e) { console.log('posConc erro ' + e); }
  }
  posConcAgendar_(); // sobrou fila: continua na próxima
}
function posConcRodar_(job) {
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss), p = job.p || {}, x = job.x || {};
  var ln = confLinhaNf_(abas.nf, p.idNf); if (ln < 0) return;
  var trocas = trocasDaNf_(p.idNf), t = Date.now();
  try { emailConclusao_(abas, ln, p, x.dono, x.res, x.diverg || [], x.semComp || [], x.alerta, trocas, x.rodada); } catch (e) {}
  var t1 = Date.now(); try { gravarValidades_(abas, ln, p, x.rodada); } catch (e) {}
  var t2 = Date.now(); try { moverRecebida_(abas, ln); } catch (e) {}
  console.log('posConc ' + String(p.idNf).slice(-14) + ' email ' + (t1 - t) + 'ms validades ' + (t2 - t1) + 'ms mover ' + (Date.now() - t2) + 'ms');
}


// ========== v118.5 — VÍNCULO AUTOMÁTICO PELO CÓDIGO BIPADO (pedido do Celio, 12/11/2026, NF Casterleite 208530) ==========
// Nota sem código de barras (ou com código fora do cadastro) entra com "**". Se a conferente bipou um código que está no
// cadastro, o item é vinculado sozinho (sem cartão de PRODUTO NOVO) e o vínculo fornecedor+código da nota fica gravado.
function vincGravar_(cnpj, cprod, descNf, sku, desc) {
  var vs = criarAbaVinculos_(SpreadsheetApp.getActive()), k = chaveVinc_(cnpj, cprod, descNf), n = vs.getLastRow();
  var hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  if (n >= 2) { var at = vs.getRange(2, 1, n - 1, 3).getDisplayValues(); for (var i = 0; i < at.length; i++) if (chaveVinc_(at[i][0], at[i][1], at[i][2]) === k) { vs.getRange(i + 2, 4, 1, 3).setNumberFormat('@').setValues([[String(sku), String(desc || ''), hoje]]); return; } }
  vs.getRange(n + 1, 1, 1, 6).setNumberFormat('@').setValues([[String(cnpj), String(cprod), String(descNf || ''), String(sku), String(desc || ''), hoje]]);
}
function autoVincBipado_(soNf) {
  var ss = SpreadsheetApp.getActive(), abas = criarAbaConf_(ss), n = abas.it.getLastRow(); if (n < 2) return 0;
  var nc = CAB_CIT.length + EXTRA_CIT.length, iCp = CAB_CIT.length + EXTRA_CIT.indexOf('CPROD'), vi = decMemoVals_('CIT', abas.it, nc), pend = {}, tem = false; // v131
  vi.forEach(function (l, i) { if (String(l[9]) === 'S' && !skuN_(l[2]) && (!soNf || String(l[0]) === String(soNf))) { pend[String(l[14])] = i; tem = true; } });
  if (!tem) return 0;
  var nl = abas.la.getLastRow(); if (nl < 2) return 0;
  var cod = {};
  decMemoVals_('CLA', abas.la, CAB_CLA.length + EXTRA_CLA.length).forEach(function (l) { var id = String(l[1]); if (pend[id] == null) return; /* v131 */ var c = String(l[17]).replace(/\D/g, ''); if (c) cod[id] = { c: c, st: String(l[18] || '') }; });
  var ids = Object.keys(cod); if (!ids.length) return 0;
  var cad = mapaCadastro_(), m = decMapa_(), feitos = 0;
  ids.forEach(function (id) {
    var x = cod[id], c = cad[x.c], mm = x.st.match(/^OK:(\d+)/), sku = c ? skuN_(c.sku) : (mm ? mm[1] : ''); if (!sku) return;
    var cc = cadPorSku_(cad, sku) || c; if (!cc) return;
    var chave = 'NOVO|' + id, d = m[chave]; if (d && /^(EXECUTADA|A EXECUTAR|DECIDIDA)$/.test(d.status)) return;
    var i = pend[id], l = vi[i], r = i + 2, antes = { c: l[2], d: l[3], i: l[8], j: l[9] };
    abas.it.getRange(r, 3, 1, 2).setValues([["'" + sku, cc.desc]]); abas.it.getRange(r, 10).setValue('');
    l[2] = "'" + sku; l[3] = cc.desc; l[9] = ''; // v131: memo igual à planilha
    var cnpj = cnpjDoIdNf_(l[0]), cprod = String(l[iCp] || '').replace(/^'/, '');
    if (!cprod) { try { cprod = (cprodsDaNf_(abas, String(l[0])) || {})[String(l[1])] || ''; if (cprod) abas.it.getRange(r, iCp + 1).setValue("'" + cprod); } catch (e) {} }
    if (cnpj && cprod) { try { vincGravar_(cnpj, cprod, String(l[15] || ''), sku, cc.desc); } catch (e) {} }
    decGrava_('PRODUTO_NOVO', chave, 'EXECUTADA', '', 'VINCULAR', { sku: sku, desc: cc.desc, prod: String(l[15] || ''), idNf: String(l[0]), antes: antes, auto: 'código bipado ' + x.c }, 'automático (código bipado)', '');
    feitos++;
  });
  if (feitos) { try { CacheService.getScriptCache().remove('DEC_RES_V113'); } catch (e) {} }
  return feitos;
}

// ---------- v123: Agenda Celio (planilha "PENDÊNCIAS DO DIA - Celio") ----------
var AGC_ID = '1_LzCBrNsoUW4eQO7HgPWDQDORM7m4_9vgS6kJmeiQVE';
function agcAba_() { var ss = SpreadsheetApp.openById(AGC_ID); return ss.getSheetByName('Untitled') || ss.getSheets()[0]; }
function agcHoje_() { var s = Utilities.formatDate(new Date(), 'America/Cuiaba', 'yyyy-MM-dd').split('-'); return new Date(+s[0], +s[1] - 1, +s[2]); }
function agcData_(v) { var m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (!m) return null; var a = +m[3]; if (a < 100) a += 2000; return new Date(a, +m[2] - 1, +m[1]); }
function agcFmt_(d) { return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear(); }
function agcTipo_(v) { return semAcento_(String(v || '')).toLowerCase().trim(); }
function agcLer_() {
  var sh = agcAba_(), n = sh.getLastRow(); if (n < 2) return { sh: sh, linhas: [] };
  var nc = Math.min(sh.getMaxColumns(), 10), v = sh.getRange(2, 1, n - 1, nc).getDisplayValues(); // v123.2: J = Link da conversa
  return { sh: sh, linhas: v.map(function (l, i) { return { row: i + 2, pri: String(l[0]).trim(), tarefa: String(l[1]).trim(), criada: l[2], prazo: agcData_(l[3]), status: String(l[4]).trim(), dias: Number(l[5]) || 0, obs: String(l[6]).trim(), tipo: agcTipo_(l[7]), hora: String(l[8]).trim(), link: String(l[9] || '').trim() }; }).filter(function (o) { return o.tarefa; }) };
}
function agcCalc_() {
  var h = agcHoje_().getTime(), L = agcLer_().linhas, PR = { alta: 0, media: 1, baixa: 2 };
  var feito = function (o) { return agcTipo_(o.status) === 'feito'; };
  var agenda = L.filter(function (o) { return o.tipo === 'agenda' && !feito(o) && o.prazo && o.prazo.getTime() === h; })
    .sort(function (a, b) { return (a.hora || '99').localeCompare(b.hora || '99'); });
  var pend = L.filter(function (o) { return o.tipo !== 'agenda' && !feito(o) && (!o.prazo || o.prazo.getTime() <= h); })
    .sort(function (a, b) { var pa = PR[agcTipo_(a.pri)], pb = PR[agcTipo_(b.pri)]; pa = pa == null ? 3 : pa; pb = pb == null ? 3 : pb; return pa - pb || b.dias - a.dias; });
  var lim = function (o) { return { row: o.row, tarefa: o.tarefa, pri: o.pri, dias: o.dias, obs: o.obs, hora: o.hora, link: o.link }; };
  return { agenda: agenda.map(lim), pend: pend.map(lim) };
}
function agcLista(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  try { agcGatilho_(); } catch (e) {}
  var c = agcCalc_(); c.ok = true; c.hoje = Utilities.formatDate(new Date(), 'America/Cuiaba', 'dd/MM/yyyy'); return c;
}
function agcAcao(p) {
  var a = v9Dono_(p); if (a.erro) return a.erro; if (!ehCelio_(a.d)) return { ok: false, erro: 'só o Celio' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = agcAba_(), row = Number(p.row);
    if (!(row >= 2) || row > sh.getLastRow()) return { ok: false, erro: 'linha não encontrada' };
    if (String(sh.getRange(row, 2).getDisplayValue()).trim() !== String(p.tarefa || '').trim()) return { ok: false, erro: 'a planilha mudou, atualize a tela' };
    if (p.op === 'FEITO') { sh.getRange(row, 5).setValue('Feito'); return { ok: true }; }
    if (p.op === 'ADIAR') {
      var d = agcHoje_();
      if (p.quando === 'DIA') d.setDate(d.getDate() + 1);
      else if (p.quando === 'SEMANA') d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7));
      else if (p.quando === 'MES') d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
      else return { ok: false, erro: 'opção inválida' };
      sh.getRange(row, 4).setValue(agcFmt_(d));
      return { ok: true, prazo: agcFmt_(d) };
    }
    return { ok: false, erro: 'ação inválida' };
  } finally { lock.releaseLock(); }
}
function agcRolar() { // gatilho diário 6h: pendência vencida volta para hoje e soma 1 em "Dias adiada"
  var r = agcLer_(), h = agcHoje_(), hs = agcFmt_(h), n = 0;
  r.linhas.forEach(function (o) {
    if (o.tipo === 'agenda' || agcTipo_(o.status) === 'feito' || !o.prazo || o.prazo.getTime() >= h.getTime()) return;
    r.sh.getRange(o.row, 4).setValue(hs); r.sh.getRange(o.row, 6).setValue(o.dias + 1); n++;
  });
  return { ok: true, roladas: n };
}
function agcGatilho_() {
  var c = CacheService.getScriptCache(); if (c.get('AGC_GAT')) return;
  var gs = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (gs.indexOf('agcRolar') < 0) ScriptApp.newTrigger('agcRolar').timeBased().atHour(6).nearMinute(0).everyDays(1).inTimezone('America/Cuiaba').create();
  c.put('AGC_GAT', '1', 21600);
}


// v124: sugestões antigas da REBAIXAS (A DECIDIR) que vieram só do saldo do sistema -> ACHAR NA LOJA (meninas conferem na gôndola) e saem da REBAIXAS
function v124Legado_(grava) {
  var ss = SpreadsheetApp.openById(ID_PLANILHA_REBAIXAS), d = new Date(), sh = ss.getSheetByName(MES_PT[d.getMonth()] + String(d.getFullYear()).slice(2));
  var n = sh.getLastRow(), v = sh.getRange(1, 1, n, 12).getDisplayValues(), iD = -1, skus = vsSkuPorNome_(), sug = [], lins = [], sem = [];
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim().toUpperCase() === 'A DECIDIR') { iD = i; break; }
  if (iD < 0) return { ok: false, erro: 'sem bloco A DECIDIR' };
  for (i = iD + 1; i < v.length; i++) {
    var l = v[i], obs = String(l[10]); if (!String(l[1]).trim() || !/SALDO SISTEMA|critica de lote/i.test(obs) || String(l[4]).trim() || String(l[11]).toUpperCase() === 'TRUE' || /NAO FAZER|NAO TEM ESSA VALIDADE|LOTE CONFERIDO|ACHAD|CONTAD|AREA|NAO TEM MAIS/i.test(obs)) continue; // só as que vieram do saldo do sistema e ainda não foram conferidas nem decididas
    var sk = skus[normDesc_(String(l[1]).replace(/[¬]/g, ''))];
    if (!sk) { sem.push(String(l[1])); continue; }
    sug.push({ sku: skuN_(sk.sku), desc: String(l[1]).replace(/[¬]/g, '').trim(), validade: v9Br_(l[9]) || String(l[9]), obs: obs }); lins.push(i + 1);
  }
  if (!grava) return { ok: true, teste: true, itens: sug.map(function (o, k) { return lins[k] + ' ' + o.sku + ' ' + o.desc + ' ' + o.validade; }), semSku: sem };
  var psv = {}; try { v9Linhas_(SpreadsheetApp.getActive().getSheetByName(ABA_PSV), CAB_PSV.length).forEach(function (r) { psv[skuN_(r[1])] = r; }); } catch (e) {}
  var criados = acharCriar_(sug, sug.map(function (o) { return psv[o.sku]; }).filter(Boolean));
  var c = rebCols_(sh); lins.forEach(function (r) { sh.getRange(r, 2, 1, 2).clearContent(); sh.getRange(r, 10, 1, 2).clearContent(); if (c.ORIGEM) sh.getRange(r, c.ORIGEM).clearContent(); });
  return { ok: true, achar: criados, tiradas: lins.length, semSku: sem };
}

// ========== v125 (16/11/2026): 🔁 TROCA DE FRENTE ATRASADA — push diário às 14h (Cuiabá), seg a sáb ==========
// Lê a aba TROCA FRENTE (só leitura). Repositora: trocas PENDENTE com 2+ dias (Açougue -> Cristiano). Ney: resumo das de 3+ dias, por repositora.
// Troca em que o próprio Ney é a repositora e tem 3+ dias entra só no resumo dele. Nada atrasado -> não manda nada.
var TF_ATR_REP = 2, TF_ATR_NEY = 3;
function tfAtrasoDiario() { // gatilho diário (tfAtrasoInstalar)
  if (new Date().getDay() === 0) return { ok: true, domingo: true };
  var equipe = PropertiesService.getScriptProperties().getProperty('TF_ATRASO_EQUIPE') === 'SIM'; // enquanto o Celio não liga, vai só para ele (teste)
  return tfAtrasoRodar_(equipe ? '' : 'Celio');
}
function tfAtrasoInstalar() {
  var tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'tfAtrasoDiario'; });
  if (!tem) ScriptApp.newTrigger('tfAtrasoDiario').timeBased().atHour(14).nearMinute(0).everyDays(1).inTimezone('America/Cuiaba').create();
  return { ok: true, funcao: 'tfAtrasoDiario', gatilho: tem ? 'já existia' : 'criado' };
}
function tfNomeCurto_(d) {
  var s = String(d || '').replace(/[¬*]/g, '').replace(/^(REFRIGERANTE|REFRIG|BEBIDA ENERG|BEBIDA|BISC RECHEADO|BISC|CERVEJA|CEREAL|NECTAR|AGUA MINERAL)\s+/i, '').trim();
  return s.toLowerCase().replace(/(^|[\s\/(-])([a-zà-ú])/g, function (m, a, b) { return a + b.toUpperCase(); }).replace(/\b(De|Da|Do|Dos|Das|E|C\/|Com)\b/g, function (m) { return m.toLowerCase(); }).replace(/(\d)(ml|l|g|kg|un)\b/gi, function (m, a, b) { b = b.toLowerCase(); return a + (b === 'l' ? 'L' : b); });
}
function tfLocalCurto_(l) { var s = String(l || '').split(/\s+[–-]\s+/)[0].trim(); return s.length > 18 ? s.slice(0, 18) : s; }
function tfAtrasoRodar_(soPara) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA_TF), n = sh ? sh.getLastRow() : 0, hj = diaNum_(v9Hoje_());
  if (n < 2) return { ok: true, nada: true };
  var ney = 'ney', porRep = {}, nomeRep = {}, neyL = [];
  sh.getRange(2, 1, n - 1, 10).getDisplayValues().forEach(function (l) {
    if (String(l[9]).trim().toUpperCase() !== 'PENDENTE') return;
    var d0 = diaNum_(String(l[0]).slice(0, 10)); if (isNaN(d0)) return;
    var dias = Math.floor(hj - d0); if (dias < TF_ATR_REP) return;
    var rep = String(l[6]).trim(); if (/^a[cç]ougue$/i.test(semAcento_(rep)) || /^acougue$/i.test(semAcento_(rep))) rep = 'Cristiano';
    if (!rep) rep = 'Ney';
    var k = semAcento_(rep).toLowerCase(), it = { prod: tfNomeCurto_(l[3]), loc: tfLocalCurto_(l[5]), desde: String(l[0]).slice(0, 5), dias: dias, rep: rep };
    if (dias >= TF_ATR_NEY) neyL.push(it);
    if (k === ney && dias >= TF_ATR_NEY) return; // já vai no resumo do Ney
    (porRep[k] = porRep[k] || []).push(it); nomeRep[k] = rep;
  });
  var ord = function (a, b) { return b.dias - a.dias; }, msgs = [];
  Object.keys(porRep).forEach(function (k) {
    var L = porRep[k].sort(ord), tx = L.length + (L.length === 1 ? ' troca: ' : ' trocas: ') + L.map(function (x) { return x.prod + (x.loc ? ' (' + x.loc + ')' : '') + ' desde ' + x.desde; }).join('; ');
    msgs.push({ para: nomeRep[k], tit: 'Troca de frente atrasada', tx: tx, tag: 'tf-atraso-' + k.replace(/\W+/g, '_') });
  });
  if (neyL.length) {
    neyL.sort(ord); var cnt = {}, ordem = [];
    neyL.forEach(function (x) { if (!cnt[x.rep]) { cnt[x.rep] = 0; ordem.push(x.rep); } cnt[x.rep]++; });
    ordem.sort(function (a, b) { return cnt[b] - cnt[a]; });
    var tx = ordem.map(function (r) { return r + ' ' + cnt[r]; }).join(', ') + ': ' + neyL.map(function (x) { return x.prod + ' desde ' + x.desde; }).join('; ');
    msgs.push({ para: 'Ney', tit: 'Trocas de frente atrasadas (3+ dias)', tx: tx, tag: 'tf-atraso-ney' });
  }
  if (!msgs.length) return { ok: true, nada: true };
  var res = msgs.map(function (m) {
    var tit = soPara ? '[TESTE p/ ' + m.para + '] ' + m.tit : m.tit, tx = m.tx.length > 230 ? m.tx.slice(0, 228) + '…' : m.tx;
    var r = null; try { r = enviarPush_([soPara || m.para], tit, tx, { abrir: 'trocafrente', tag: m.tag + (soPara ? '-teste' : '') }); } catch (e) { r = { erro: String(e) }; }
    return { para: m.para, enviadoPara: soPara || m.para, titulo: tit, texto: tx, aparelhos: r ? r.aparelhos : 0, enviados: r ? r.enviados : 0, erro: r && r.erro || '' };
  });
  try {
    var av = SpreadsheetApp.getActive().getSheetByName(ABA_AVISOS);
    if (av) res.forEach(function (x) { av.appendRow([new Date(), 'AUTOMATICO (troca de frente atrasada)', x.enviadoPara, x.titulo, x.texto, (x.enviados || 0) + '/' + (x.aparelhos || 0)]); });
  } catch (e) {}
  return { ok: true, msgs: res };
}

// ========== v129 (17/11/2026): DEMANDA ORIGINAL das rebaixas ==========
// A rotina de rebaixas move o produto para QUEIMA DE ESTOQUE; a aba REBAIXAS (coluna C) tem que guardar a demanda de antes.
// Ordem: a) aba DEMANDA ORIGINAL (linha mais recente do SKU com H vazio); b) linha anterior do SKU na REBAIXAS; c) DEMANDA HIST.
var ABA_DEMORIG = 'DEMANDA ORIGINAL', CAB_DEMORIG = ['SKU', 'DESCRICAO', 'DEMANDA ORIGINAL', 'MOVIDO PARA', 'MOVIDO EM', 'VALIDADE REBAIXA', 'QUEM', 'DEVOLVIDO EM'];
var ABA_DEMHIST = 'DEMANDA HIST', CAB_DEMHIST = ['SKU', 'DEMANDA', 'ATUALIZADO EM'];
function demTemp_(d) { d = String(d || '').trim().toUpperCase(); return !d || /^QUEIMA DE ESTOQUE|^ESPECIAL FIM DE ANO|^SEM DEMANDA$/.test(d); }
function demAba_(nome, cab) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange('A:A').setNumberFormat('@'); }
  return sh;
}
function demTs_(s) { var m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/); return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime() : 0; }
function demCtx_(rebVals) { // rebVals = linhas da aba REBAIXAS (display, a partir da linha 2)
  var ctx = { orig: {}, reb: {}, hist: {}, rebVals: rebVals || [] };
  try {
    var so = demAba_(ABA_DEMORIG, CAB_DEMORIG), no = so.getLastRow();
    if (no >= 2) so.getRange(2, 1, no - 1, 8).getDisplayValues().forEach(function (l) {
      var s = skuN_(l[0]), d = String(l[2]).trim().toUpperCase(); if (!s || demTemp_(d) || String(l[7]).trim()) return;
      var t = demTs_(l[4]), o = ctx.orig[s]; if (!o || t >= o.t) ctx.orig[s] = { dem: d, t: t };
    });
  } catch (e) {}
  try { var sh = demAba_(ABA_DEMHIST, CAB_DEMHIST), nh = sh.getLastRow(); if (nh >= 2) sh.getRange(2, 1, nh - 1, 2).getDisplayValues().forEach(function (l) { var s = skuN_(l[0]), d = String(l[1]).trim().toUpperCase(); if (s && !demTemp_(d)) ctx.hist[s] = d; }); } catch (e) {}
  return ctx;
}
function demOriginal_(sku, ctx, semLinha) { // { dem, fonte } ou null; semLinha = índice da própria linha na REBAIXAS (não usar ela mesma)
  var s = skuN_(sku); if (!s) return null;
  if (ctx.orig[s]) return { dem: ctx.orig[s].dem, fonte: 'DEMANDA ORIGINAL' };
  for (var i = ctx.rebVals.length - 1; i >= 0; i--) { if (i === semLinha) continue; var l = ctx.rebVals[i]; if (skuN_(l[0]) === s && !demTemp_(l[2]) && !/\(ORIGEM\?\)/.test(String(l[2]))) return { dem: String(l[2]).trim().toUpperCase(), fonte: 'REBAIXAS anterior' }; }
  if (ctx.hist[s]) return { dem: ctx.hist[s], fonte: 'DEMANDA HIST' };
  return null;
}
function quemDaDemanda_(dem) { // aba SETORES: DEMANDA -> QUEM
  var d = String(dem || '').trim().toUpperCase(); if (demTemp_(d)) return '';
  if (!DEM_SET_) { DEM_SET_ = {}; try { var sh = criarAbaSetores_(SpreadsheetApp.getActive()), n = sh.getLastRow(); if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (l) { var k = String(l[0]).trim().toUpperCase(); if (k && String(l[1]).trim() && !DEM_SET_[k]) DEM_SET_[k] = String(l[1]).trim(); }); } catch (e) {} }
  return DEM_SET_[d] || '';
}
var DEM_SET_ = null;
function demHistGravar_(rows) { // rows = linhas da aba CADASTRO [codigo, sku, desc, un, bal, preco, nv, dem, ...]
  var sh = demAba_(ABA_DEMHIST, CAB_DEMHIST), n = sh.getLastRow(), m = {}, ordem = [], agora = agoraBr_(), mud = 0;
  if (n >= 2) sh.getRange(2, 1, n - 1, 3).getDisplayValues().forEach(function (l) { var s = skuN_(l[0]); if (s && !m[s]) { m[s] = [s, String(l[1]), String(l[2])]; ordem.push(s); } });
  rows.forEach(function (r) {
    var s = skuN_(r[1]), d = String(r[7] || '').trim().toUpperCase(); if (!s || demTemp_(d)) return; // QUEIMA / ESPECIAL FIM DE ANO não entram
    if (!m[s]) { m[s] = [s, d, agora]; ordem.push(s); mud++; } else if (m[s][1] !== d) { m[s][1] = d; m[s][2] = agora; mud++; }
  });
  if (mud) { var out = ordem.map(function (s) { return m[s]; }); sh.getRange(2, 1, out.length, 3).setNumberFormat('@').setValues(out); }
  PropertiesService.getScriptProperties().setProperty('DEM_HIST_DIA', v9Hoje_());
  return { skus: ordem.length, mudou: mud };
}
function demHistDiario_() { // garante 1x por dia, mesmo se o atualizarCadastro não rodou
  if (PropertiesService.getScriptProperties().getProperty('DEM_HIST_DIA') === v9Hoje_()) return null;
  var c = SpreadsheetApp.getActive().getSheetByName(ABA_CAD), n = c ? c.getLastRow() : 0; if (n < 2) return null;
  return demHistGravar_(c.getRange(2, 1, n - 1, 8).getDisplayValues());
}
// uma vez: corrige as linhas da REBAIXAS com C = QUEIMA / ESPECIAL FIM DE ANO; as que não der, lista com sugestão (produtos parecidos do cadastro)
function v129CorrigirQueima_(grava) {
  var hist = demHistGravar_(SpreadsheetApp.getActive().getSheetByName(ABA_CAD).getRange(2, 1, SpreadsheetApp.getActive().getSheetByName(ABA_CAD).getLastRow() - 1, 8).getDisplayValues());
  var sh = criarAbaRebaixas_(SpreadsheetApp.getActive()), n = sh.getLastRow(), v = n >= 2 ? sh.getRange(2, 1, n - 1, 8).getDisplayValues() : [], ctx = demCtx_(v), ok = [], sem = [], cadv = null;
  v.forEach(function (l, i) {
    var c = String(l[2]).trim().toUpperCase(); if (!/^QUEIMA DE ESTOQUE|^ESPECIAL FIM DE ANO/.test(c)) return;
    var o = demOriginal_(l[0], ctx, i);
    if (o) { ok.push({ linha: i + 2, sku: l[0], prod: l[1], de: c, para: o.dem, fonte: o.fonte }); if (grava) sh.getRange(i + 2, 3).setValue(o.dem); return; }
    if (!cadv) cadv = SpreadsheetApp.getActive().getSheetByName(ABA_CAD).getRange(2, 2, SpreadsheetApp.getActive().getSheetByName(ABA_CAD).getLastRow() - 1, 7).getDisplayValues();
    var pal = semAcento_(String(l[1])).toUpperCase().split(/\s+/).filter(function (w) { return w.length > 2 && !/\d/.test(w); }).slice(0, 2), cnt = {};
    if (pal.length === 2) cadv.forEach(function (r) { var d = String(r[6]).trim().toUpperCase(); if (demTemp_(d) || skuN_(r[0]) === skuN_(l[0])) return; var t = semAcento_(String(r[1])).toUpperCase(); if (t.indexOf(pal[0]) >= 0 && t.indexOf(pal[1]) >= 0) cnt[d] = (cnt[d] || 0) + 1; });
    var sug = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; })[0] || '';
    sem.push({ linha: i + 2, sku: l[0], prod: l[1], status: l[7], sugestao: sug ? sug + ' (' + cnt[sug] + ' produtos parecidos)' : '' });
    if (grava && !/\(ORIGEM\?\)/.test(c)) sh.getRange(i + 2, 3).setValue('QUEIMA DE ESTOQUE (ORIGEM?)');
  });
  return { ok: true, teste: !grava, hist: hist, corrigidas: ok, semOrigem: sem };
}

// TEMP v131 medição 5 (remover depois)
function perfV131e() {
  var out = {}, t = function (nome, f) { var t0 = Date.now(); var r; try { r = f(); } catch (e) { out[nome + '_err'] = e.message; } out[nome] = Date.now() - t0; return r; };
  var cc = CacheService.getScriptCache(); cc.remove('DEC_RES_V113'); cc.remove('DECL131_n'); cc.remove(rhChave_('Celio')); DEC_PAL_ = null;
  var tk = 'perf'; // decLista sem login: simula o corpo
  var L = t('decListaFria', function () { return decListaCorpo131_(); });
  try { out.nA = L.aDecidir.length; out.nE = L.esperando.length; out.nD = L.decididas.length; out.erros = L.erros; out.kb = Math.round(JSON.stringify(L).length / 1024); } catch (e) { out.lerr = e.message; }
  t('decListaCache', function () { return cacheGrandeLer_('DECL131'); });
  t('decResumoFrio', function () { return decResumo_(); });
  t('homeCelioFrio', function () { return resumoHome_('Celio'); });
  t('homeCelioCache', function () { return resumoHome_('Celio'); });
  var sv = t('semVendaCelio', function () { return produtoSemVenda('Celio'); });
  try { out.svKB = Math.round(JSON.stringify(sv).length / 1024); out.svN = sv.itens.length; } catch (e) {}
  Logger.log(JSON.stringify(out));
  return out;
}
function decListaCorpo131_() { // cópia do corpo de decLista (sem login) para medir
  var hit = null;
  DEC_MEMO_ = {};
  var m = decMapa_(); try { if (decGondolaVencidas_(m)) m = decMapa_(); } catch (e) {}
  var r = { ok: true, aDecidir: [], esperando: [], decididas: [], erros: [] };
  [['REBAIXA', function () { return decRebCartoes_(m); }], ['PRODUTO_NOVO', function () { return decNovoCartoes_(m); }], ['CONTAGEM', decContCartoes_], ['CODIGO', function () { return decCodCartoes_(m); }]].forEach(function (t) {
    try { t[1]().forEach(function (c) { (c.esperando ? r.esperando : r.aDecidir).push(c); }); } catch (e) { r.erros.push(t[0] + ': ' + e.message); }
  });
  var lim = diaNum_(v9Hoje_()) - 15, pmc = {};
  Object.keys(m).forEach(function (k) { var d = m[k]; if (d.status !== 'ESPERANDO' && d.status !== 'PENDENTE' && diaNum_(String(d.em).slice(0, 10)) >= lim) { var pm = decPodeMudar_(k, d, pmc); r.decididas.push({ k: k, pode: pm.pode, motivo: pm.motivo || '' }); } });
  DEC_MEMO_ = null;
  cacheGrande_('DECL131', r, 120);
  return r;
}
