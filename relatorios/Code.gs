// API de gravação dos relatórios diários (Fazenda App - Fase 3).
// Cole este código em Extensões > Apps Script da planilha
// "Fazenda App - Relatórios Diários" e implante como "App da Web"
// (executar como "Eu", acesso "Qualquer pessoa com o link").
//
// Cada fazenda ganha sua própria aba, criada automaticamente com o código
// que vier no envio (CB, PB, SH, ou qualquer fazenda nova cadastrada depois
// — não precisa editar este script quando uma fazenda nova for cadastrada).
// Se o envio chegar sem o código da fazenda (não deveria acontecer, já que o
// campo é obrigatório no formulário), a linha cai na aba "RelatoriosDiarios"
// em vez de se perder, para facilitar identificar o problema.

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

  var fazenda = (dados.fazenda || '').trim();
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
  var nomeAba = fazenda || FALLBACK_SHEET_NAME;
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    // ResumoTexto (F) e DadosJSON (G): corta o texto em vez de expandir a
    // altura da linha. Clicar na célula mostra o texto completo na barra de
    // fórmulas.
    sheet.getRange('F:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  return sheet;
}
