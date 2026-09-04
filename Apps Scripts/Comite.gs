/**
 * ============================================================================
 * COMITÊ — ARQUIVO ÚNICO (Acompanhamento Comitê)
 * ============================================================================
 * Junta em um só arquivo o que antes eram três: KanbanAPI.gs + Painel.gs +
 * Scripts.gs. No Apps Script os .gs de um projeto já compartilham o MESMO
 * escopo global — separar em arquivos era só organização, nunca isolamento.
 * Portanto isto é concatenação pura: nenhum comportamento muda.
 *
 * CONTEÚDO
 *   [1] API DO KANBAN (Web App)  — doGet + kb*   -> responde o app Kanban.html
 *   [2] PAINEL E MENU           — onOpen + pn*   -> menus 📊 Painel e 🚀 Ações
 *   [3] SINCRONIZAÇÃO COMITÊ -> FÁBRICA — ce*    -> gatilho de edição
 *
 * NOMES QUE NÃO PODEM MUDAR (gatilhos e implantação apontam para eles):
 *   doGet ................. implantação como Aplicativo da Web (URL /exec)
 *   onOpen ................ gatilho simples, monta os menus
 *   processarEdicaoComite . gatilho INSTALÁVEL (ao editar) — Comitê -> Fábrica
 *   atualizarPainel ....... gatilho de tempo (30 min) + item de menu
 *   criarTriggerAuto ...... item de menu
 *   excluirTicketComite ... item de menu
 *
 * DEPOIS DE COLAR: salvar. Se mexeu em algo que o Kanban usa, publicar de novo
 * em Implantar > Gerenciar implantações > editar (lápis) > Versão: Nova versão.
 * A URL /exec continua a mesma; os gatilhos NÃO precisam ser reinstalados.
 *
 * Prefixos evitam colisão entre as três partes: kb* (API), pn* (painel),
 * ce* (sync). Ao adicionar função nova, siga o prefixo da seção.
 * Obs.: excluirTicketComite declara um var FABRICA_ID local — sombreia o global
 * de mesmo nome só dentro dela; os dois valem o mesmo ID da Fábrica.
 *
 * Fonte versionada: github.com/davi-ciss/Kanban -> Apps Scripts/Comite.gs
 * ============================================================================
 */

// ==========================================================================
// [1] API DO KANBAN (Web App)
// Antigo KanbanAPI.gs — responde ao Kanban.html via JSONP (doGet).
// ==========================================================================
/**
 * Kanban — endpoint Apps Script (Web App) da planilha de Acompanhamento.
 * Compatível com QUALQUER analista: a aba usada vem do parâmetro ?tab= (o nome
 * que a pessoa põe no Kanban). O Responsável gravado na Fábrica = nome dessa aba.
 * Mapeia colunas por NOME de cabeçalho (aba do analista e aba Apoios), então não
 * quebra quando inserem/movem/renomeiam colunas. Replica as regras do onEdit ao gravar.
 * Deploy: Implantar > Gerenciar implantações > editar > Nova versão. URL /exec.
 */

var FABRICA_ID = "1SA3WD5srsNt-7nutVkKZwwc2rmkalK8N2MSCSX_sMp4";

// Aba do analista (Comitê) — campo lógico -> nomes de cabeçalho aceitos (sem acento)
var KB_COLMAP = {
  ticket:  ['ticket','chamado'],
  status:  ['status'],
  assunto: ['assunto','titulo','título','descricao','descrição'],
  cliente: ['cliente'],
  obs:     ['observacao','observação','obs'],
  atrib:   ['atribuido','atribuído','atribuido?','atribuído?'],
  jira:    ['jira'],
  alt:     ['alteracao de status','alteração de status'],
  addedat: ['adicionado em','adicionado em:','adicionado']
};

// Aba "Apoios" (Fábrica) por NOME de cabeçalho — robusto a inserções de coluna.
var APOIOS_COLMAP = {
  equipe:    ['equipe fabrica','equipe fábrica','equipe'],
  ticket:    ['ticket','chamado'],
  descricao: ['descricao','descrição','assunto','titulo','título'],
  resp:      ['responsavel','responsável'],
  cliente:   ['cliente'],
  data:      ['data'],
  jira:      ['jira'],
  status:    ['status']
};
function kbApoiosIdx_(fab){
  var lastCol = Math.max(1, fab.getLastColumn());
  var head = fab.getRange(1,1,1,lastCol).getValues()[0].map(kbNorm_);
  var idx = {};
  for(var f in APOIOS_COLMAP){ idx[f] = -1; var ns = APOIOS_COLMAP[f];
    for(var i=0;i<head.length;i++){ if(ns.indexOf(head[i])!==-1){ idx[f]=i; break; } } }
  return idx;
}

function doGet(e){
  var q = (e && e.parameter) ? e.parameter : {};
  var cb = q.callback || 'callback';
  var out;
  try{
    var sh = kbGetSheet_(q.tab);
    switch((q.action||'').toLowerCase()){
      case 'ping':     out = {ok:true, tab:sh.getName(), count:Math.max(0, sh.getLastRow()-1)}; break;
      case 'statuses': out = {ok:true, statuses:kbStatuses_(sh)}; break;
      case 'equipes':  out = {ok:true, equipes:kbEquipes_()}; break;
      case 'list':     out = {ok:true, rows:kbListRows_(sh)}; break;
      case 'upsert':   out = kbUpsert_(sh, q); break;
      case 'delete':   out = kbDelete_(sh, q); break;
      default:         out = {ok:false, error:'Ação desconhecida: '+(q.action||'(vazia)')};
    }
  }catch(err){ out = {ok:false, error:String(err)}; }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function kbGetSheet_(tab){
  tab = String(tab||'').trim();
  if(!tab) throw new Error('Informe seu nome (aba) nas Configurações do Kanban');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(tab);
  if(!sh) throw new Error('Aba "'+tab+'" não encontrada nesta planilha');
  return sh;
}

function kbNorm_(s){
  return String(s==null?'':s).toLowerCase().trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,''); // remove acentos
}

function kbHeaderIndex_(sh){
  var lastCol = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1,1,1,lastCol).getValues()[0].map(kbNorm_);
  var idx = {};
  for(var field in KB_COLMAP){
    idx[field] = -1;
    var names = KB_COLMAP[field];
    for(var i=0;i<head.length;i++){ if(names.indexOf(head[i])!==-1){ idx[field]=i; break; } }
  }
  return idx;
}

function kbListRows_(sh){
  var lastRow = sh.getLastRow(), lastCol = Math.max(1, sh.getLastColumn());
  if(lastRow < 2) return [];
  var idx = kbHeaderIndex_(sh);
  var vals = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var rows = [];
  for(var r=0;r<vals.length;r++){
    var row = vals[r];
    var ticket = idx.ticket>=0 ? row[idx.ticket] : '';
    var assunto= idx.assunto>=0? row[idx.assunto]: '';
    if(String(ticket).trim()==='' && String(assunto).trim()==='') continue;
    rows.push({
      key: String(ticket||('R'+(r+2))),
      ticket: ticket, status: idx.status>=0?row[idx.status]:'',
      assunto: assunto, cliente: idx.cliente>=0?row[idx.cliente]:'',
      obs: idx.obs>=0?row[idx.obs]:'',
      jira: idx.jira>=0?row[idx.jira]:'',
      atrib: idx.atrib>=0?row[idx.atrib]:''
    });
  }
  return rows;
}

function kbUpsert_(sh, q){
  var idx = kbHeaderIndex_(sh);
  if(idx.ticket < 0 && idx.assunto < 0)
    return {ok:false, error:'Cabeçalhos Ticket/Assunto não encontrados na linha 1'};
  var lastRow = sh.getLastRow();
  var rawKey = String(q.key==null?'':q.key).trim();
  var ticketDigits = kbDigits_(q.ticket || '');
  var targetRow = -1;

  var mR = rawKey.match(/^(?:R|row)(\d+)$/i);
  if(mR){ var rn = parseInt(mR[1],10); if(rn>=2 && rn<=lastRow) targetRow = rn; }

  if(targetRow === -1 && ticketDigits && idx.ticket>=0 && lastRow>=2){
    var col = sh.getRange(2, idx.ticket+1, lastRow-1, 1).getValues();
    for(var i=0;i<col.length;i++){ if(kbDigits_(col[i][0])===ticketDigits){ targetRow = i+2; break; } }
  }

  if(targetRow === -1 && !ticketDigits && q.assunto && idx.assunto>=0 && lastRow>=2){
    var colA = sh.getRange(2, idx.assunto+1, lastRow-1, 1).getValues();
    var assTarget = String(q.assunto).trim().toLowerCase();
    for(var aI=0;aI<colA.length;aI++){ if(String(colA[aI][0]||'').trim().toLowerCase()===assTarget){ targetRow = aI+2; break; } }
  }

  var novo = false;
  if(targetRow === -1){ targetRow = lastRow + 1; novo = true; }

  var statusAtual = (idx.status>=0) ? String(sh.getRange(targetRow, idx.status+1).getValue()) : '';

  function set(field, val){ if(idx[field]>=0 && val!==undefined && val!=='') sh.getRange(targetRow, idx[field]+1).setValue(val); }

  // Ticket -> link Zendesk
  if(ticketDigits && idx.ticket>=0){
    var urlZ = 'https://cisssa.zendesk.com/agent/tickets/' + ticketDigits;
    var rtZ = SpreadsheetApp.newRichTextValue().setText(ticketDigits).setLinkUrl(urlZ).build();
    sh.getRange(targetRow, idx.ticket+1).setRichTextValue(rtZ);
  }

  set('assunto', q.assunto);
  set('cliente', q.cliente);
  set('obs',     q.obs);
  set('atrib',   q.atrib);

  var statusNovo = (q.status!==undefined) ? String(q.status) : statusAtual;
  if(idx.status>=0 && q.status!==undefined) sh.getRange(targetRow, idx.status+1).setValue(q.status);

  if(idx.alt>=0 && q.status!==undefined && statusNovo !== statusAtual){
    if(statusNovo === '') sh.getRange(targetRow, idx.alt+1).clearContent();
    else sh.getRange(targetRow, idx.alt+1).setValue(new Date());
  }

  // Jira -> link Atlassian, só reescreve se mudou
  if(idx.jira>=0 && q.jira!==undefined && String(q.jira)!==''){
    var curJ = String(sh.getRange(targetRow, idx.jira+1).getValue()).trim();
    var jiraKey = String(q.jira).trim();
    if(curJ !== jiraKey){
      var mj = jiraKey.match(/([A-Za-z0-9]+-\d+)/);
      if(mj){
        var urlJ = 'https://cisspoder.atlassian.net/browse/' + mj[1];
        var rtJ = SpreadsheetApp.newRichTextValue().setText(mj[1]).setLinkUrl(urlJ).build();
        sh.getRange(targetRow, idx.jira+1).setRichTextValue(rtJ);
      } else {
        sh.getRange(targetRow, idx.jira+1).setValue(jiraKey);
      }
    }
  }

  if(idx.addedat>=0){
    var addCell = sh.getRange(targetRow, idx.addedat+1);
    if(String(addCell.getValue()) === '') addCell.setValue(new Date());
  }

  if(ticketDigits) kbPropagarFabrica_(sh, idx, targetRow, q.equipe||'');

  var outKey = ticketDigits || ('R'+targetRow);
  return {ok:true, key:outKey, row:targetRow, novo:novo};
}

function kbPropagarFabrica_(sh, idx, row, equipe){
  try{
    var ticketId   = idx.ticket>=0  ? sh.getRange(row, idx.ticket+1).getValue()  : '';
    if(!ticketId) return;
    var statusLocal= idx.status>=0  ? sh.getRange(row, idx.status+1).getValue()  : '';
    var assunto    = idx.assunto>=0 ? sh.getRange(row, idx.assunto+1).getValue() : '';
    var cliente    = idx.cliente>=0 ? sh.getRange(row, idx.cliente+1).getValue() : '';
    var jira       = idx.jira>=0    ? sh.getRange(row, idx.jira+1).getValue()    : '';
    var dataAdic   = idx.addedat>=0 ? sh.getRange(row, idx.addedat+1).getValue() : '';

    var temEquipe   = !!(equipe && String(equipe).trim());
    var equipeFinal = temEquipe ? String(equipe).trim() : 'N3 Fábrica';
    var deveEscalar = (statusLocal === 'Pendente N3') || temEquipe;

    var statusParaFabrica = (statusLocal === 'Pendente N3') ? 'Pendente' : statusLocal;
    var dataApenasDia = (dataAdic instanceof Date)
      ? Utilities.formatDate(dataAdic, Session.getScriptTimeZone(), 'dd/MM/yyyy') : dataAdic;

    var jiraRT = kbJiraRT_(jira);

    var fab = SpreadsheetApp.openById(FABRICA_ID).getSheetByName("Apoios");
    var a = kbApoiosIdx_(fab);
    if(a.ticket < 0) return;

    function setA(field, r, val){ if(a[field]>=0 && val!==undefined && val!=='') fab.getRange(r, a[field]+1).setValue(val); }
    function setBlank(field, r, val){ if(a[field]>=0) fab.getRange(r, a[field]+1).setValue(val); }
    function setJira(r){ if(a.jira>=0){ if(jiraRT) fab.getRange(r, a.jira+1).setRichTextValue(jiraRT); else fab.getRange(r, a.jira+1).setValue(jira); } }

    var alvo = kbDigits_(ticketId);
    var ult = fab.getLastRow();
    var encontrado = false;

    if(ult >= 2){
      var dados = fab.getRange(2, a.ticket+1, ult-1, 1).getValues();
      for(var i=0;i<dados.length;i++){
        if(kbDigits_(dados[i][0]) === alvo){
          var rr = i+2;
          if(temEquipe) setA('equipe', rr, equipeFinal);
          setA('descricao', rr, assunto);
          setBlank('resp', rr, sh.getName());   // Responsável = nome da aba (analista)
          setBlank('cliente', rr, cliente);
          setA('data', rr, dataApenasDia);
          setJira(rr);
          setA('status', rr, statusParaFabrica);
          encontrado = true;
          break;
        }
      }
    }
    if(!encontrado && deveEscalar){
      var nova = fab.getLastRow() + 1;
      setA('equipe', nova, equipeFinal);
      setA('ticket', nova, ticketId);
      setA('descricao', nova, assunto);
      setBlank('resp', nova, sh.getName());   // Responsável = nome da aba (analista)
      setBlank('cliente', nova, cliente);
      setA('data', nova, dataApenasDia);
      setJira(nova);
      setA('status', nova, statusParaFabrica);
    }
  }catch(err){ /* Fábrica indisponível: ignora */ }
}

function kbDigits_(v){ var m = String(v==null?'':v).match(/(\d+)/); return m ? m[1] : String(v||'').trim(); }

function kbEquipes_(){
  try{
    var fab = SpreadsheetApp.openById(FABRICA_ID).getSheetByName("Apoios");
    var a = kbApoiosIdx_(fab);
    var ecol = (a.equipe>=0 ? a.equipe : 0) + 1;
    for(var row=2; row<=Math.min(fab.getLastRow(),30); row++){
      var rule = fab.getRange(row, ecol).getDataValidation();
      if(!rule) continue;
      var tipo=rule.getCriteriaType(), args=rule.getCriteriaValues();
      if(tipo===SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return (args[0]||[]).map(function(v){return String(v).trim();}).filter(String);
      if(tipo===SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return args[0].getValues().map(function(r){return String(r[0]).trim();}).filter(String);
    }
    var ult=fab.getLastRow(); if(ult<2) return [];
    var vals=fab.getRange(2, ecol, ult-1, 1).getValues(), seen={}, out=[];
    for(var i=0;i<vals.length;i++){ var v=String(vals[i][0]||'').trim(); if(v && !seen[v]){ seen[v]=1; out.push(v); } }
    return out;
  }catch(e){ return []; }
}

function kbJiraRT_(jiraLike){
  var m = String(jiraLike==null?'':jiraLike).match(/([A-Za-z0-9]+-\d+)/);
  if(!m) return null;
  var key = m[1].toUpperCase();
  return SpreadsheetApp.newRichTextValue()
    .setText(key).setLinkUrl('https://cisspoder.atlassian.net/browse/' + key).build();
}

function kbStatuses_(sh){
  var idx = kbHeaderIndex_(sh);
  var col = idx.status>=0 ? idx.status+1 : 2;
  var scanAte = Math.min(sh.getLastRow(), 60);
  for(var row=2; row<=scanAte; row++){
    var rule = sh.getRange(row, col).getDataValidation();
    if(!rule) continue;
    var tipo = rule.getCriteriaType(), args = rule.getCriteriaValues();
    if(tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return (args[0]||[]).map(function(v){return String(v).trim();}).filter(String);
    if(tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return args[0].getValues().map(function(r){return String(r[0]).trim();}).filter(String);
  }
  return [];
}

// Texto normalizado p/ comparar Assunto: minúsculo, sem acento, espaços/quebras colapsados.
function kbTxt_(v){ return kbNorm_(v).replace(/\s+/g,' '); }

/**
 * Exclui a linha da tarefa na aba do analista.
 *
 * NUNCA apaga por posição sem conferir: número de linha NÃO é identidade estável —
 * qualquer inclusão/exclusão acima (aqui, no menu Ações ou manual) desloca tudo, e a
 * key "R<n>" guardada no card fica velha. Ordem de identificação:
 *   1) Ticket (mais forte)  2) key "R<n>" CONFERIDA  3) Assunto normalizado
 *      (assunto atual do card e, se veio, o assunto da última sincronização — cobre card renomeado).
 * Sem prova de que a linha é a certa, não apaga nada e devolve local:false pro Kanban avisar.
 */
function kbDelete_(sh, q){
  var idx = kbHeaderIndex_(sh);
  if(idx.ticket < 0 && idx.assunto < 0) return {ok:false, error:'Coluna Ticket/Assunto não encontrada'};

  var lastRow = sh.getLastRow();
  if(lastRow < 2) return {ok:true, deleted:'', local:false, fabrica:false, motivo:'aba sem linhas'};

  var rawKey       = String(q.key==null?'':q.key).trim();
  var ticketDigits = kbDigits_(q.ticket || (rawKey.match(/^\d+$/) ? rawKey : ''));
  var assAlvos     = [];                                   // assuntos aceitos, em ordem de preferência
  [q.assunto, q.title, q.assuntoAnt].forEach(function(v){
    var t = kbTxt_(v || '');
    if(t && assAlvos.indexOf(t) === -1) assAlvos.push(t);
  });

  var colT = idx.ticket  >= 0 ? sh.getRange(2, idx.ticket+1,  lastRow-1, 1).getValues() : null;
  var colA = idx.assunto >= 0 ? sh.getRange(2, idx.assunto+1, lastRow-1, 1).getValues() : null;

  var alvo = -1, via = '';

  // 1) Ticket
  if(ticketDigits && colT){
    for(var i=0;i<colT.length;i++){
      if(kbDigits_(colT[i][0]) === ticketDigits){ alvo = i+2; via = 'ticket'; break; }
    }
  }

  // 2) key "R<n>" — só vale se a linha AINDA for a mesma tarefa
  if(alvo === -1){
    var mR = rawKey.match(/^(?:R|row)(\d+)$/i);
    if(mR){
      var rn = parseInt(mR[1],10);
      if(rn >= 2 && rn <= lastRow){
        var tLin = colT ? kbDigits_(colT[rn-2][0]) : '';
        var aLin = colA ? kbTxt_(colA[rn-2][0])    : '';
        var confere = ticketDigits ? (tLin === ticketDigits)
                                   : (assAlvos.length ? (assAlvos.indexOf(aLin) !== -1) : false);
        if(confere){ alvo = rn; via = 'linha'; }
      }
    }
  }

  // 3) Assunto normalizado — é por aqui que cai o card SEM ticket
  if(alvo === -1 && colA && assAlvos.length){
    for(var t=0; t<assAlvos.length && alvo===-1; t++){
      for(var k=0;k<colA.length;k++){
        if(kbTxt_(colA[k][0]) === assAlvos[t]){ alvo = k+2; via = 'assunto'; break; }
      }
    }
  }

  if(alvo === -1)
    return {ok:true, deleted:'', local:false, fabrica:false, motivo:'linha não localizada na aba '+sh.getName()};

  // Ticket REAL da linha: o card pode estar sem ticket e a linha ter (então a Fábrica também tem).
  var ticketLinha = colT ? kbDigits_(colT[alvo-2][0]) : '';
  sh.deleteRow(alvo);

  var deletedFabrica = false;
  if(ticketLinha){
    try{
      var fab = SpreadsheetApp.openById(FABRICA_ID).getSheetByName("Apoios");
      var a = kbApoiosIdx_(fab);
      var ult = fab.getLastRow();
      if(a.ticket >= 0 && ult >= 2){
        var dados = fab.getRange(2, a.ticket+1, ult-1, 1).getValues();
        for(var j=0;j<dados.length;j++){
          if(kbDigits_(dados[j][0]) === ticketLinha){ fab.deleteRow(j+2); deletedFabrica = true; break; }
        }
      }
    }catch(err){ /* Fábrica indisponível */ }
  }

  return {ok:true, deleted:ticketLinha || ('R'+alvo), local:true, fabrica:deletedFabrica, via:via};
}

// ==========================================================================
// [2] PAINEL DE DEMANDAS E MENUS
// Antigo Painel.gs — onOpen, gráficos e o menu 🚀 Ações.
// ==========================================================================
/**
 * Painel de Demandas em Aberto + Ações (menu) — Acompanhamento Comitê.
 *
 * Auto-adaptativo: detecta as abas de analista pela ESTRUTURA (cabeçalho com
 * Ticket + Status + Assunto), sem lista fixa de nomes. Conta como "em aberto"
 * tudo que NÃO for: resolvido/auxiliado/feito/cancelado/concluído/finalizado/fábrica.
 *
 * onOpen e excluirTicketComite existem UMA vez só — aqui, na seção [2].
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

  // Linha sem ticket (tarefa interna, criada direto no Kanban): existe só nesta aba,
  // não há nada correspondente na Fábrica — confirma e apaga só aqui.
  if (!ticketId) {
    var assuntoLinha = String(sheet.getRange(row, ix.assunto + 1).getValue() || '').trim();
    var r1 = ui.alert('Confirmar Exclusão',
      'Esta linha não tem ticket.\n\nExcluir permanentemente a linha ' + row +
      (assuntoLinha ? ' ("' + assuntoLinha + '")' : '') + ' desta aba?', ui.ButtonSet.YES_NO);
    if (r1 != ui.Button.YES) return;
    sheet.deleteRow(row);
    SpreadsheetApp.getActiveSpreadsheet().toast('Linha ' + row + ' excluída (sem ticket, nada a remover na Fábrica).', 'Sucesso', 5);
    return;
  }

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

// ==========================================================================
// [3] SINCRONIZAÇÃO COMITÊ -> FÁBRICA
// Antigo Scripts.gs — gatilho instalável processarEdicaoComite.
// ==========================================================================
// =========================================================================
// EDIÇÃO E SINCRONIZAÇÃO EM SEGUNDO PLANO (Comitê -> Fábrica)
//
// Auto-adaptativo: roda em QUALQUER aba de analista (detecta pela estrutura:
// tem cabeçalho Ticket + Status + Assunto). O Responsável gravado na Fábrica é
// o NOME da aba. Colunas mapeadas por cabeçalho — não depende de posição fixa
// nem de todas as colunas existirem.
//
// onOpen / excluirTicketComite ficam na seção [2] deste arquivo (não duplicar aqui).
// =========================================================================

var CE_FABRICA_ID = "1SA3WD5srsNt-7nutVkKZwwc2rmkalK8N2MSCSX_sMp4";

// Aba do analista (Comitê) — campo lógico -> cabeçalhos aceitos (sem acento)
var CE_COLMAP = {
  ticket:  ['ticket','chamado'],
  status:  ['status'],
  assunto: ['assunto','titulo','título','descricao','descrição'],
  cliente: ['cliente'],
  jira:    ['jira'],
  alt:     ['alteracao de status','alteração de status'],
  addedat: ['adicionado em','adicionado em:','adicionado']
};
// Aba "Apoios" (Fábrica)
var CE_APOIOS_COLMAP = {
  equipe:    ['equipe fabrica','equipe fábrica','equipe'],
  ticket:    ['ticket','chamado'],
  descricao: ['descricao','descrição','assunto','titulo','título'],
  resp:      ['responsavel','responsável'],
  cliente:   ['cliente'],
  data:      ['data'],
  jira:      ['jira'],
  status:    ['status']
};

function ceNorm_(x){ return String(x==null?'':x).toLowerCase().trim().normalize('NFD').replace(/[̀-ͯ]/g,''); }
function ceDigits_(v){ var m=String(v==null?'':v).match(/(\d+)/); return m?m[1]:String(v||'').trim(); }
function ceIdx_(sh, colmap){
  var lc=Math.max(1, sh.getLastColumn());
  var head=sh.getRange(1,1,1,lc).getValues()[0].map(ceNorm_);
  var idx={};
  for(var f in colmap){ idx[f]=-1; var ns=colmap[f];
    for(var i=0;i<head.length;i++){ if(ns.indexOf(head[i])!==-1){ idx[f]=i; break; } } }
  return idx;
}
// É aba de analista? tem Ticket + Status + Assunto e não é a aba de painel.
function ceIsAnalyst_(sh, idx){
  var n = ceNorm_(sh.getName());
  if(n==='graficos' || n==='painel' || n.indexOf('apoio')===0) return false;
  return idx.ticket>=0 && idx.status>=0 && idx.assunto>=0;
}

function processarEdicaoComite(e){
  if(!e || !e.source) return;
  var s = e.source.getActiveSheet();
  if(e.range.getRow() <= 1) return;
  var idx = ceIdx_(s, CE_COLMAP);
  if(!ceIsAnalyst_(s, idx)) return;          // não é aba de analista -> ignora

  var r = e.range, linha = r.getRow(), col = r.getColumn()-1;   // col 0-based
  var resp = s.getName();                    // Responsável = nome da aba

  // Links: Ticket (Zendesk) e Jira
  if(col===idx.ticket || col===idx.jira){
    var value = r.getValue();
    if(typeof value === 'string'){
      if(col===idx.ticket){
        var mZ = value.match(/cisssa\.zendesk\.com\/agent\/tickets\/(\d+)/);
        if(mZ){ r.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(mZ[1]).setLinkUrl(value.trim()).build()); r.setFontColor(null); }
      }
      if(col===idx.jira){
        var mJ = value.match(/([A-Za-z0-9]+-\d+)/);
        if(mJ) r.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(mJ[1]).setLinkUrl('https://cisspoder.atlassian.net/browse/'+mJ[1]).build());
      }
    }
  }

  // Adicionado em: preenche na 1ª vez que a linha tem conteúdo
  if(idx.addedat>=0 && col!==idx.addedat){
    var cAdd = s.getRange(linha, idx.addedat+1);
    if(cAdd.getValue()==="" && r.getValue()!=="") cAdd.setValue(new Date());
  }

  // Propaga ao mudar Status, Assunto, Cliente ou Jira
  var gatilho = (col===idx.status)||(col===idx.assunto)||(col===idx.cliente)||(col===idx.jira);
  if(!gatilho) return;

  // Alteração de Status
  if(col===idx.status && idx.alt>=0){
    var cAlt = s.getRange(linha, idx.alt+1);
    if(e.value===undefined || r.getValue()==="") cAlt.clearContent(); else cAlt.setValue(new Date());
  }

  try{
    var ticketId = idx.ticket>=0 ? s.getRange(linha, idx.ticket+1).getValue() : '';
    if(!ticketId) return;
    var statusLocal = idx.status>=0 ? s.getRange(linha, idx.status+1).getValue() : '';
    var assunto     = idx.assunto>=0? s.getRange(linha, idx.assunto+1).getValue(): '';
    var cliente     = idx.cliente>=0? s.getRange(linha, idx.cliente+1).getValue(): '';
    var jira        = idx.jira>=0   ? s.getRange(linha, idx.jira+1).getValue()   : '';
    var dataAdic    = idx.addedat>=0? s.getRange(linha, idx.addedat+1).getValue(): '';
    cePropagarFabrica_(ticketId, statusLocal, assunto, cliente, jira, dataAdic, resp);
  }catch(err){ console.error(err); }
}

// Espelha a linha na Fábrica "Apoios" (header-based). Cria linha nova só em "Pendente N3".
function cePropagarFabrica_(ticketId, statusLocal, assunto, cliente, jira, dataAdic, resp){
  var statusFab = (statusLocal==="Pendente N3") ? "Pendente" : statusLocal;
  var dataDia = (dataAdic instanceof Date) ? Utilities.formatDate(dataAdic, Session.getScriptTimeZone(), "dd/MM/yyyy") : dataAdic;
  var ap = SpreadsheetApp.openById(CE_FABRICA_ID).getSheetByName("Apoios");
  if(!ap) return;
  var a = ceIdx_(ap, CE_APOIOS_COLMAP);
  if(a.ticket < 0) return;

  function setA(field, row, val){ if(a[field]>=0 && val!=='' && val!==undefined) ap.getRange(row, a[field]+1).setValue(val); }
  function setB(field, row, val){ if(a[field]>=0) ap.getRange(row, a[field]+1).setValue(val); }

  var alvo = ceDigits_(ticketId);
  var ult = ap.getLastRow(), achou=false;
  if(ult>=2){
    var ids = ap.getRange(2, a.ticket+1, ult-1, 1).getValues();
    for(var i=0;i<ids.length;i++){
      if(ceDigits_(ids[i][0])===alvo){
        var rr=i+2;
        setA('descricao', rr, assunto);
        setB('resp', rr, resp);
        setB('cliente', rr, cliente);
        setA('data', rr, dataDia);
        setA('jira', rr, jira);
        setA('status', rr, statusFab);
        achou=true; break;
      }
    }
  }
  if(!achou && statusLocal==="Pendente N3"){
    var nl = ap.getLastRow()+1;
    setB('equipe', nl, "N3 Fábrica");
    setB('ticket', nl, ticketId);
    setB('descricao', nl, assunto);
    setB('resp', nl, resp);
    setB('cliente', nl, cliente);
    setB('data', nl, dataDia);
    setB('jira', nl, jira);
    setB('status', nl, statusFab);
  }
}
