// ============================================================
// ALMOXARIFADO - Google Apps Script (Backend)
// Cole este código no Apps Script da sua planilha
// ============================================================

const SHEET_ITENS       = 'Itens';
const SHEET_HISTORICO   = 'Historico';
const SHEET_FORNECEDORES= 'Fornecedores';
const SHEET_LOG         = 'Log_Auditoria';

// Cabeçalhos de cada aba
const HDR_ITENS        = ['Codigo','Descricao','Unidade','Quantidade','Estoque_Minimo','Localizacao'];
const HDR_HISTORICO    = ['ID','Data','Hora','Tipo','Codigo','Descricao','Quantidade','Unidade','Fornecedor','Referencia','Observacao','Usuario'];
const HDR_FORNECEDORES = ['Nome','CNPJ','Contato','Observacao'];
const HDR_LOG          = ['Data','Hora','Usuario','Acao','Detalhes'];

// ---- Inicializa abas se não existirem ----
function inicializar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[SHEET_ITENS, HDR_ITENS],[SHEET_HISTORICO, HDR_HISTORICO],[SHEET_FORNECEDORES, HDR_FORNECEDORES],[SHEET_LOG, HDR_LOG]].forEach(([nome, hdr]) => {
    let sh = ss.getSheetByName(nome);
    if (!sh) {
      sh = ss.insertSheet(nome);
      sh.appendRow(hdr);
      sh.getRange(1,1,1,hdr.length).setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });
}

// ---- Ponto de entrada HTTP ----
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    inicializar();
    const params = e.parameter || {};
    const body   = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    const action = params.action || body.action;
    const usuario= params.usuario || body.usuario || 'Desconhecido';

    let result;
    switch(action) {
      case 'getAll':          result = getAll(); break;
      case 'addItem':         result = addItem(body); break;
      case 'editItem':        result = editItem(body, usuario); break;
      case 'deleteItem':      result = deleteItem(body, usuario); break;
      case 'registrarMov':    result = registrarMov(body, usuario); break;
      case 'addFornecedor':   result = addFornecedor(body); break;
      case 'editFornecedor':  result = editFornecedor(body, usuario); break;
      case 'deleteFornecedor':result = deleteFornecedor(body, usuario); break;
      case 'getLog':          result = getLog(); break;
      default: result = {erro: 'Ação desconhecida: ' + action};
    }
    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({erro: err.message});
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Helpers de planilha ----
function getSheet(nome) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
}

function sheetToArray(nome) {
  const sh = getSheet(nome);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function logAuditoria(usuario, acao, detalhes) {
  const sh = getSheet(SHEET_LOG);
  const agora = new Date();
  sh.appendRow([
    Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy'),
    Utilities.formatDate(agora, 'America/Sao_Paulo', 'HH:mm:ss'),
    usuario, acao, detalhes
  ]);
}

// ---- Ações ----
function getAll() {
  return {
    itens:        sheetToArray(SHEET_ITENS),
    historico:    sheetToArray(SHEET_HISTORICO),
    fornecedores: sheetToArray(SHEET_FORNECEDORES)
  };
}

function addItem(body) {
  const sh = getSheet(SHEET_ITENS);
  const itens = sheetToArray(SHEET_ITENS);
  if (itens.find(i => i.Codigo === body.cod)) return {erro: 'Código já cadastrado.'};
  sh.appendRow([body.cod, body.desc, body.unid, body.qtd, body.min, body.local||'']);
  return {ok: true};
}

function editItem(body, usuario) {
  const sh = getSheet(SHEET_ITENS);
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.codOrig) {
      sh.getRange(r+1, 1, 1, 6).setValues([[body.cod, body.desc, body.unid, body.qtd, body.min, body.local||'']]);
      logAuditoria(usuario, 'EDITAR_ITEM', `${body.codOrig} → ${body.cod} | ${body.desc}`);
      return {ok: true};
    }
  }
  return {erro: 'Item não encontrado.'};
}

function deleteItem(body, usuario) {
  const sh = getSheet(SHEET_ITENS);
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.cod) {
      logAuditoria(usuario, 'EXCLUIR_ITEM', `Código: ${body.cod} | ${data[r][1]}`);
      sh.deleteRow(r+1);
      return {ok: true};
    }
  }
  return {erro: 'Item não encontrado.'};
}

function registrarMov(body, usuario) {
  // Atualiza saldo
  const shI = getSheet(SHEET_ITENS);
  const dataI = shI.getDataRange().getValues();
  let found = false;
  for (let r = 1; r < dataI.length; r++) {
    if (dataI[r][0] === body.cod) {
      const saldoAtual = parseFloat(dataI[r][3]) || 0;
      const qtd = parseFloat(body.qtd) || 0;
      if (body.tipo === 'saida' && saldoAtual < qtd) return {erro: 'Saldo insuficiente. Disponível: ' + saldoAtual + ' ' + dataI[r][2]};
      const novoSaldo = body.tipo === 'entrada' ? saldoAtual + qtd : saldoAtual - qtd;
      shI.getRange(r+1, 4).setValue(novoSaldo);
      found = true;
      break;
    }
  }
  if (!found) return {erro: 'Item não encontrado.'};

  // Grava no histórico
  const shH = getSheet(SHEET_HISTORICO);
  const agora = new Date();
  const id = 'MOV' + agora.getTime();
  shH.appendRow([
    id,
    Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy'),
    Utilities.formatDate(agora, 'America/Sao_Paulo', 'HH:mm'),
    body.tipo, body.cod, body.desc, body.qtd, body.unid,
    body.forn||'', body.ref||'', body.obs||'', usuario
  ]);
  return {ok: true};
}

function addFornecedor(body) {
  const sh = getSheet(SHEET_FORNECEDORES);
  const lista = sheetToArray(SHEET_FORNECEDORES);
  if (lista.find(f => f.Nome === body.nome)) return {erro: 'Fornecedor já cadastrado.'};
  sh.appendRow([body.nome, body.cnpj||'', body.contato||'', body.obs||'']);
  return {ok: true};
}

function editFornecedor(body, usuario) {
  const sh = getSheet(SHEET_FORNECEDORES);
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.nomeOrig) {
      sh.getRange(r+1,1,1,4).setValues([[body.nome, body.cnpj||'', body.contato||'', body.obs||'']]);
      logAuditoria(usuario, 'EDITAR_FORNECEDOR', `${body.nomeOrig} → ${body.nome}`);
      return {ok: true};
    }
  }
  return {erro: 'Fornecedor não encontrado.'};
}

function deleteFornecedor(body, usuario) {
  const sh = getSheet(SHEET_FORNECEDORES);
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.nome) {
      logAuditoria(usuario, 'EXCLUIR_FORNECEDOR', `Nome: ${body.nome}`);
      sh.deleteRow(r+1);
      return {ok: true};
    }
  }
  return {erro: 'Fornecedor não encontrado.'};
}

function getLog() {
  return { log: sheetToArray(SHEET_LOG) };
}
