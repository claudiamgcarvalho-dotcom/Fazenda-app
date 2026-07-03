// API de leitura dos cadastros (Fazenda App - Fase 2).
// Cole este código em Extensões > Apps Script da planilha de cadastros e
// implante como "App da Web" (acesso: Qualquer pessoa com o link).

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fazendas = lerFazendas(ss);
  // Códigos de fazenda vêm da própria aba "Fazendas" (em vez de uma lista
  // fixa no código) — uma fazenda nova cadastrada lá já aparece aqui sem
  // precisar editar este script.
  var farmCodes = Object.keys(fazendas);

  var result = {
    fazendas: fazendas,
    funcionarios: lerPorFazenda(ss, 'Funcionarios', farmCodes),
    pastos: lerPorFazenda(ss, 'Pastos', farmCodes),
    produtosNutricao: lerCatalogo(ss, 'ProdutosNutricao', farmCodes, function (r) {
      return { nome: r.Nome, unidade: r.Unidade, tamanhoSaco: Number(r.TamanhoSaco) || 0 };
    }),
    categoriasAnimal: lerCatalogo(ss, 'CategoriasAnimal', farmCodes, function (r) {
      return r.Nome;
    }),
    vacinas: lerCatalogo(ss, 'Vacinas', farmCodes, function (r) {
      return r.Nome;
    }),
    tiposProduto: lerCatalogo(ss, 'TiposProduto', farmCodes, function (r) {
      return { nome: r.Nome, descExemplo: r.DescricaoExemplo, qtdExemplo: r.QuantidadeExemplo };
    })
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function isAtivo(valor) {
  return !(valor === false || valor === 'FALSE' || valor === 'false');
}

// Lê uma aba e devolve um array de objetos {NomeDaColuna: valor}, usando a
// linha 1 como cabeçalho. Ignora linhas totalmente vazias.
function lerLinhas(ss, nomeAba) {
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) throw new Error('Aba não encontrada: ' + nomeAba);
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var linhas = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var vazia = row.every(function (c) { return c === ''; });
    if (vazia) continue;
    var obj = {};
    header.forEach(function (h, idx) { obj[h] = row[idx]; });
    linhas.push(obj);
  }
  return linhas;
}

// Aba "Fazendas": Codigo, Nome, Ativo -> { CB: "Nome", ... }
function lerFazendas(ss) {
  var linhas = lerLinhas(ss, 'Fazendas');
  var out = {};
  linhas.forEach(function (r) {
    if (!isAtivo(r.Ativo)) return;
    out[r.Codigo] = r.Nome;
  });
  return out;
}

// Abas "por fazenda" (Funcionarios, Pastos): Fazenda, Nome, Ativo
// -> { CB: ["Nome", ...], PB: [...], SH: [...] }
function lerPorFazenda(ss, nomeAba, farmCodes) {
  var linhas = lerLinhas(ss, nomeAba);
  var out = {};
  farmCodes.forEach(function (c) { out[c] = []; });
  linhas.forEach(function (r) {
    if (!isAtivo(r.Ativo)) return;
    if (out[r.Fazenda]) out[r.Fazenda].push(r.Nome);
  });
  return out;
}

// Abas "catálogo compartilhado" (ProdutosNutricao, CategoriasAnimal,
// Vacinas, TiposProduto): cada linha tem FazendasAtivas = "CB,PB,SH" e é
// distribuída para os buckets dos códigos listados.
function lerCatalogo(ss, nomeAba, farmCodes, mapLinha) {
  var linhas = lerLinhas(ss, nomeAba);
  var out = {};
  farmCodes.forEach(function (c) { out[c] = []; });
  linhas.forEach(function (r) {
    var ativas = String(r.FazendasAtivas || '').split(',')
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s; });
    ativas.forEach(function (codigo) {
      if (out[codigo]) out[codigo].push(mapLinha(r));
    });
  });
  return out;
}
