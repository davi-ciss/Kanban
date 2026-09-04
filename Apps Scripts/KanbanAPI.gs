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