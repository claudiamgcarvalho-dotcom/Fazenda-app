// API de gravação dos relatórios diários (Fazenda App - Fase 3).
// Cole este código em Extensões > Apps Script da planilha
// "Fazenda App - Relatórios Diários" e implante como "App da Web"
// (executar como "Eu", acesso "Qualquer pessoa com o link").
//
// Cada fazenda tem sua própria aba (CB, PB, SH), criada automaticamente na
// primeira gravação. Se o envio chegar sem o código da fazenda (ex.: ela não
// foi selecionada no formulário), a linha cai na aba "RelatoriosDiarios" em
// vez de se perder, para facilitar identificar o problema.

var FARM_CODES = ['CB', 'PB', 'SH'];
var FALLBACK_SHEET_NAME = 'RelatoriosDiarios';
var HEADERS = ['Timestamp', 'ID', 'Data', 'Fazenda', 'Responsavel', 'ResumoTexto', 'DadosJSON'];

function doPost(e) {
  var raw = (e && e.postData && e.postData.contents) || '';
  var dados = {};
  try {
    dados = JSON.parse(raw);
  } catch (err) {
    dados = {};
  }

  var fazenda = dados.fazenda || '';
  var sheet = getOrCreateSheet(fazenda);
  var id = [dados.data || '', fazenda || 'SEM_FAZENDA', new Date().getTime()].join('_');

  sheet.appendRow([
    new Date(),
    id,
    dados.data || '',
    fazenda,
    dados.responsavel || '',
    dados.resumoTexto || '',
    raw
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(fazenda) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nomeAba = FARM_CODES.indexOf(fazenda) >= 0 ? fazenda : FALLBACK_SHEET_NAME;
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}
