// API de gravação dos relatórios diários (Fazenda App - Fase 3).
// Cole este código em Extensões > Apps Script da planilha
// "Fazenda App - Relatórios Diários" e implante como "App da Web"
// (executar como "Eu", acesso "Qualquer pessoa com o link").

var SHEET_NAME = 'RelatoriosDiarios';
var HEADERS = ['Timestamp', 'ID', 'Data', 'Fazenda', 'Responsavel', 'ResumoTexto', 'DadosJSON'];

function doPost(e) {
  var sheet = getOrCreateSheet();
  var dados = JSON.parse(e.postData.contents);

  var id = [dados.data, dados.fazenda, new Date().getTime()].join('_');
  sheet.appendRow([
    new Date(),
    id,
    dados.data || '',
    dados.fazenda || '',
    dados.responsavel || '',
    dados.resumoTexto || '',
    JSON.stringify(dados)
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}
