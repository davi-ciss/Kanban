// =========================================================================
// EDIÇÃO E SINCRONIZAÇÃO EM SEGUNDO PLANO (Comitê -> Fábrica)
//
// Auto-adaptativo: roda em QUALQUER aba de analista (detecta pela estrutura:
// tem cabeçalho Ticket + Status + Assunto). O Responsável gravado na Fábrica é
// o NOME da aba. Colunas mapeadas por cabeçalho — não depende de posição fixa
// nem de todas as colunas existirem.
//
// onOpen / excluirTicketComite ficam no arquivo Painel.gs (não duplicar aqui).
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
