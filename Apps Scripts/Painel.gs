/**
 * Painel de Demandas em Aberto + Ações (menu) — Acompanhamento Comitê.
 *
 * Auto-adaptativo: detecta as abas de analista pela ESTRUTURA (cabeçalho com
 * Ticket + Status + Assunto), sem lista fixa de nomes. Conta como "em aberto"
 * tudo que NÃO for: resolvido/auxiliado/feito/cancelado/concluído/finalizado/fábrica.
 *
 * onOpen e excluirTicketComite ficam SÓ neste arquivo (não duplicar no Script.gs).
 */

var ABA_PAINEL      = 'Gráficos';
var LINHA_CABECALHO = 1;
var EXCLUIR_KEYWORDS = ['resolv', 'auxili', 'feito', 'cancel', 'conclu', 'finaliz', 'fabrica'];

// --- MENU ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Painel')
    .addItem('Atualizar gráficos', 'atualizarPainel')
    .addSeparator()
    .addItem('Criar atualização automática (30 min)', 'criarTriggerAuto')
    .addToUi();
  ui.createMenu('🚀 Ações')
    .addItem('🗑️ Excluir Ticket de Ambas as Planilhas', 'excluirTicketComite')
    .addToUi();
}

// --- HELPERS ---
function norm_(s) {
  return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}
function ehExcluido_(sNorm) {
  for (var i = 0; i < EXCLUIR_KEYWORDS.length; i++) if (sNorm.indexOf(EXCLUIR_KEYWORDS[i]) !== -1) return true;
  return false;
}
function ehN3_(sNorm) { return sNorm.indexOf('n3') !== -1; }

// Índice (0-based) das colunas pelo cabeçalho. -1 se não achar.
function pnColIdx_(sh) {
  var lc = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, lc).getValues()[0].map(norm_);
  function find(names) { for (var i = 0; i < head.length; i++) if (names.indexOf(head[i]) !== -1) return i; return -1; }
  return {
    ticket:  find(['ticket', 'chamado']),
    status:  find(['status']),
    assunto: find(['assunto', 'titulo', 'título', 'descricao', 'descrição'])
  };
}
// Abas de analista = têm Ticket+Status+Assunto e não são painel/apoios.
function pnAbasAnalistas_(ss) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var n = norm_(sh.getName());
    if (n === 'graficos' || n === 'painel' || n.indexOf('apoio') === 0) return;
    var ix = pnColIdx_(sh);
    if (ix.ticket >= 0 && ix.status >= 0 && ix.assunto >= 0) out.push(sh.getName());
  });
  return out;
}

// --- EXCLUSÃO (menu Ações) ---
function excluirTicketComite() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var ix = pnColIdx_(sheet);
  if (!(ix.ticket >= 0 && ix.status >= 0 && ix.assunto >= 0)) {
    ui.alert("Aviso", "Selecione uma célula numa aba de analista (com colunas Ticket/Status/Assunto).", ui.ButtonSet.OK);
    return;
  }
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { ui.alert("Aviso", "Você não pode excluir a linha de cabeçalho.", ui.ButtonSet.OK); return; }

  var ticketId = sheet.getRange(row, ix.ticket + 1).getValue();
  if (!ticketId) { ui.alert("Aviso", "Nenhum ticket encontrado na linha selecionada.", ui.ButtonSet.OK); return; }

  var numTicket = ticketId.toString().match(/(\d+)/); numTicket = numTicket ? numTicket[1] : ticketId;
  var resp = ui.alert('Confirmar Exclusão', 'Excluir permanentemente o ticket ' + numTicket + ' de AMBAS as planilhas?', ui.ButtonSet.YES_NO);
  if (resp != ui.Button.YES) return;

  try {
    var FABRICA_ID = "1SA3WD5srsNt-7nutVkKZwwc2rmkalK8N2MSCSX_sMp4";
    var ap = SpreadsheetApp.openById(FABRICA_ID).getSheetByName("Apoios");
    var at = pnColIdx_(ap).ticket;          // coluna Ticket da Apoios (por cabeçalho)
    var ult = ap.getLastRow();
    if (at >= 0 && ult >= 2) {
      var dados = ap.getRange(2, at + 1, ult - 1, 1).getValues();
      for (var i = 0; i < dados.length; i++) {
        var idOutra = dados[i][0].toString().match(/(\d+)/); idOutra = idOutra ? idOutra[1] : dados[i][0];
        if (idOutra == numTicket) { ap.deleteRow(i + 2); break; }
      }
    }
  } catch (err) { ui.alert("Erro ao excluir na Fábrica: " + err); return; }

  sheet.deleteRow(row);
  SpreadsheetApp.getActiveSpreadsheet().toast("Ticket " + numTicket + " excluído em ambas as planilhas.", "Sucesso", 5);
}

// --- COLETA (auto-detecta analistas e colunas) ---
function coletar_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resumo = [], matriz = {}, totaisStatus = {}, labelPorNorm = {};
  pnAbasAnalistas_(ss).forEach(function (nome) {
    var sh = ss.getSheetByName(nome); if (!sh) return;
    var ix = pnColIdx_(sh);
    var aberto = 0, n3 = 0; matriz[nome] = {};
    var dados = sh.getDataRange().getValues();
    for (var i = LINHA_CABECALHO; i < dados.length; i++) {
      var assunto = String(ix.assunto >= 0 ? dados[i][ix.assunto] : '').trim();
      if (!assunto) continue;
      var statusBruto = String(ix.status >= 0 ? dados[i][ix.status] : '').trim();
      var sNorm = norm_(statusBruto);
      if (ehExcluido_(sNorm)) continue;
      aberto++; if (ehN3_(sNorm)) n3++;
      var label = statusBruto || 'Sem Status', key = sNorm || 'sem status';
      labelPorNorm[key] = labelPorNorm[key] || label;
      matriz[nome][key] = (matriz[nome][key] || 0) + 1;
      totaisStatus[key] = (totaisStatus[key] || 0) + 1;
    }
    resumo.push({ analista: nome, aberto: aberto, n3: n3 });
  });
  var statusKeys = Object.keys(totaisStatus).sort(function (a, b) {
    if (a === 'sem status') return 1; if (b === 'sem status') return -1;
    return totaisStatus[b] - totaisStatus[a];
  });
  return { resumo: resumo, matriz: matriz, totaisStatus: totaisStatus, statusKeys: statusKeys, labelPorNorm: labelPorNorm };
}

// --- CONSTRÓI O PAINEL ---
function atualizarPainel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_PAINEL);
  if (!sh) sh = ss.insertSheet(ABA_PAINEL);
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clear();

  var d = coletar_();
  sh.getRange('A1').setValue('PAINEL DE DEMANDAS EM ABERTO').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Atualizado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')).setFontColor('#666');

  var r = 4, resumoIni = r;
  sh.getRange(r, 1, 1, 3).setValues([['Analista', 'Em Aberto', 'N3 em Aberto']]).setFontWeight('bold').setBackground('#1e4620').setFontColor('#fff'); r++;
  d.resumo.forEach(function (x) { sh.getRange(r, 1, 1, 3).setValues([[x.analista, x.aberto, x.n3]]); r++; });
  var resumoFim = r - 1;
  sh.getRange(r, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(r, 2).setFormula('=SUM(B' + (resumoIni + 1) + ':B' + resumoFim + ')').setFontWeight('bold');
  sh.getRange(r, 3).setFormula('=SUM(C' + (resumoIni + 1) + ':C' + resumoFim + ')').setFontWeight('bold');
  r += 2;

  var matrizIni = r;
  var header = ['Analista'].concat(d.statusKeys.map(function (k) { return d.labelPorNorm[k]; }));
  sh.getRange(r, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1e4620').setFontColor('#fff'); r++;
  d.resumo.forEach(function (x) {
    var linha = [x.analista].concat(d.statusKeys.map(function (k) { return (d.matriz[x.analista] && d.matriz[x.analista][k]) || 0; }));
    sh.getRange(r, 1, 1, linha.length).setValues([linha]); r++;
  });
  var matrizFim = r - 1, matrizCols = header.length; r += 2;

  var pieIni = r;
  sh.getRange(r, 1, 1, 2).setValues([['Status', 'Qtd']]).setFontWeight('bold').setBackground('#1e4620').setFontColor('#fff'); r++;
  d.statusKeys.forEach(function (k) { sh.getRange(r, 1, 1, 2).setValues([[d.labelPorNorm[k], d.totaisStatus[k]]]); r++; });
  var pieFim = r - 1;

  var ancCol = 9;
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(resumoIni, 1, resumoFim - resumoIni + 1, 1))
    .addRange(sh.getRange(resumoIni, 2, resumoFim - resumoIni + 1, 1))
    .setPosition(resumoIni, ancCol, 0, 0).setNumHeaders(1)
    .setOption('title', 'Demandas em aberto por analista').setOption('legend', { position: 'none' })
    .setOption('width', 480).setOption('height', 300).build());

  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(matrizIni, 1, matrizFim - matrizIni + 1, matrizCols))
    .setPosition(resumoIni + 17, ancCol, 0, 0).setNumHeaders(1).setOption('useFirstColumnAsDomain', true)
    .setOption('title', 'Backlog por status (empilhado)').setOption('isStacked', true).setOption('legend', { position: 'bottom' })
    .setOption('width', 560).setOption('height', 340).build());

  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(resumoIni, 1, resumoFim - resumoIni + 1, 1))
    .addRange(sh.getRange(resumoIni, 3, resumoFim - resumoIni + 1, 1))
    .setPosition(resumoIni + 34, ancCol, 0, 0).setNumHeaders(1)
    .setOption('title', 'N3 em aberto por analista').setOption('legend', { position: 'none' })
    .setOption('colors', ['#b71c1c']).setOption('width', 480).setOption('height', 300).build());

  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange(pieIni, 1, pieFim - pieIni + 1, 2))
    .setPosition(resumoIni + 51, ancCol, 0, 0).setNumHeaders(1)
    .setOption('title', 'Distribuição geral por status (aberto)').setOption('pieHole', 0.4)
    .setOption('width', 480).setOption('height', 300).build());

  sh.autoResizeColumns(1, matrizCols);
  SpreadsheetApp.flush();
  ss.toast('Painel atualizado!', '📊 Painel', 4);
}

// --- TRIGGER A CADA 30 MIN ---
function criarTriggerAuto() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'atualizarPainel') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('atualizarPainel').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert('Atualização automática criada: a cada 30 minutos.');
}
