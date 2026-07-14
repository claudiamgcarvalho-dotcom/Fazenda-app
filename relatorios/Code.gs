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

// Planilha separada "Solicitações" (uma aba por fazenda), usada pra guardar o
// link do cartão do Trello de cada solicitação depois que a administração cria
// o cartão. Cole aqui o ID da planilha (está na URL dela) depois de criá-la —
// até lá, fica desativado sem quebrar o envio do relatório diário.
var SOLICITACOES_SPREADSHEET_ID = '1cYVvpPC9Xfi8E8yen2UVzjPyctlIYKKqqjAjOCuV81M';
var SOLICITACOES_HEADERS = ['Timestamp', 'ID', 'Data', 'Fazenda', 'Descricao', 'Quantidade', 'Urgencia', 'LinkTrello'];

// Planilha separada "Controle de Horas" (uma aba por fazenda, mais uma aba
// "<FAZENDA>_ResumoMensal" com o total por mês e funcionário). Cole aqui o ID
// dela depois de criada — até lá, fica desativado sem quebrar o envio do
// relatório diário (mesmo padrão de proteção da SOLICITACOES_SPREADSHEET_ID).
var CONTROLE_HORAS_SPREADSHEET_ID = '12YGHFdE8hPEyr4HQmE23n1l2jN-uxOjirBCuLUu1TZU';
var HORAS_HEADERS = ['Data', 'Fazenda', 'Funcionario', 'DiaSemana', 'Presenca', 'Horas', 'Dias', 'ObservacoesGerais'];
var RESUMO_MENSAL_HEADERS = ['Ano', 'Mes', 'Funcionario', 'TotalHoras', 'TotalDias'];
var DIAS_SEMANA = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];

// Planilha separada "Estoque Nutrição" (uma aba por fazenda). Cole aqui o ID
// da planilha depois de criá-la — até lá, fica desativado sem quebrar o envio
// do relatório diário (mesmo padrão de proteção das outras planilhas auxiliares).
var ESTOQUE_SPREADSHEET_ID = '1uwVG9sbCZNUFFzU84BwUu7MHXlWK-7yBjqJEOWk-JJw';
var ESTOQUE_HEADERS = ['Timestamp', 'Data', 'Fazenda', 'Produto', 'Tipo', 'QtdSacos', 'QtdKg', 'Observacoes'];

var SAUDE_ANIMAL_SPREADSHEET_ID = '1u58XdIQaR9ht87A8ngMDXEXAVbC8yXnEUZ_PHLb1iaU';
var VACINAS_HEADERS     = ['Timestamp', 'Data', 'Fazenda', 'Categoria Animal', 'Qtd Animais', 'Nome', 'Tipo', 'Dose', 'Observações'];
var MEDICAMENTOS_HEADERS = ['Timestamp', 'Data', 'Fazenda', 'Categoria Animal', 'Qtd Animais', 'ID Animal', 'Sintomas', 'Medicamento', 'Observações'];

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
  var id = dados._id || [dados.data || '', fazenda || 'SEM_FAZENDA', new Date().getTime()].join('_');

  // Deduplicação: se já existe linha com este id, retorna ok sem inserir
  var existingData = sheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (String(existingData[i][1]) === String(id)) {
      return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id, duplicate: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  sheet.appendRow([
    new Date(),
    id,
    dados.data || '',
    fazenda,
    dados.responsavel || '',
    dados.resumoTexto || '',
    raw
  ]);

  registrarSolicitacoes(dados, fazenda);
  registrarHoras(dados, fazenda);
  registrarConsumo(dados, fazenda);
  registrarVacinas(dados, fazenda);
  registrarMedicamentos(dados, fazenda);

  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Copia cada solicitação de produto do envio do dia pra planilha "Solicitações"
// (uma linha por item, na aba da fazenda), pra administração acompanhar e
// colar o link do Trello depois. Envolto em try/catch: se a planilha ainda não
// foi criada/configurada (SOLICITACOES_SPREADSHEET_ID com o valor placeholder),
// o relatório diário continua sendo salvo normalmente, só sem essa cópia.
function registrarSolicitacoes(dados, fazenda) {
  var itens = (dados.solicitacoesProdutos || {}).itens || [];
  if (!itens.length) return;
  try {
    var ss = SpreadsheetApp.openById(SOLICITACOES_SPREADSHEET_ID);
    var aba = getOrCreateAbaSolicitacoes(ss, fazenda);
    itens.forEach(function (item) {
      aba.appendRow([
        new Date(),
        item.id || '',
        dados.data || '',
        fazenda,
        item.descricao || '',
        item.quantidade || '',
        item.urgencia || '',
        ''
      ]);
    });
  } catch (err) {
    // Planilha de Solicitações ainda não configurada — ignora silenciosamente.
  }
}

function getOrCreateAbaSolicitacoes(ss, fazenda) {
  var nomeAba = fazenda || FALLBACK_SHEET_NAME;
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  if (aba.getLastRow() === 0) aba.appendRow(SOLICITACOES_HEADERS);
  return aba;
}

// Devolve o nome do dia da semana (em português) de uma data "yyyy-MM-dd",
// construindo a data a partir dos números diretamente (evita ambiguidade de
// fuso horário que o new Date(string) padrão do JS poderia causar).
function diaDaSemana(dataISO) {
  var p = String(dataISO).split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  return DIAS_SEMANA[d.getDay()];
}

// Regra de horas/dias trabalhados (definida com a usuária):
// Integral: 8h/1 dia de segunda a domingo, exceto sábado (4h/1 dia).
// Meio período: 4h/0,5 dia de segunda a domingo, exceto sábado (4h/1 dia).
// Falta e Folga combinada: sempre 0h/0 dia.
function calcularHorasDias(presenca, dataISO) {
  var sabado = diaDaSemana(dataISO) === 'Sábado';
  if (presenca === 'Integral') return { horas: sabado ? 4 : 8, dias: 1 };
  if (presenca === 'Meio período') return { horas: 4, dias: sabado ? 1 : 0.5 };
  return { horas: 0, dias: 0 };
}

// Copia a equipe do dia pra planilha "Controle de Horas" (uma linha por
// funcionário) e atualiza o resumo mensal. Se já houver mais de um envio no
// mesmo dia/fazenda (já visto acontecer na prática), as linhas desse dia são
// substituídas a cada novo envio — só o último envio do dia vale, igual o
// painel diário já prioriza. Envolto em try/catch: enquanto a planilha não
// estiver configurada, o relatório diário continua sendo salvo normalmente.
function registrarHoras(dados, fazenda) {
  var equipe = (dados.fechamento || {}).equipe || [];
  if (!equipe.length || !dados.data) return;
  try {
    var ss = SpreadsheetApp.openById(CONTROLE_HORAS_SPREADSHEET_ID);
    var aba = getOrCreateAbaHoras(ss, fazenda);

    removerLinhasDoDia(aba, dados.data);

    var observacoesGerais = (dados.fechamento || {}).observacoesGerais || '';
    equipe.forEach(function (p) {
      var calc = calcularHorasDias(p.presenca, dados.data);
      aba.appendRow([dados.data, fazenda, p.nome || '', diaDaSemana(dados.data), p.presenca || '', calc.horas, calc.dias, observacoesGerais]);
    });

    atualizarResumoMensal(ss, fazenda, dados.data);
  } catch (err) {
    // Planilha de Controle de Horas ainda não configurada — ignora silenciosamente.
  }
}

function getOrCreateAbaHoras(ss, fazenda) {
  var nomeAba = fazenda || FALLBACK_SHEET_NAME;
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  if (aba.getLastRow() === 0) aba.appendRow(HORAS_HEADERS);
  return aba;
}

// Remove as linhas já existentes daquela data antes de regravar — garante que
// múltiplos envios no mesmo dia não dupliquem horas (só o último prevalece).
function removerLinhasDoDia(aba, dataISO) {
  var linhas = aba.getDataRange().getValues();
  for (var i = linhas.length - 1; i >= 1; i--) {
    if (formatarData(linhas[i][0]) === dataISO) {
      aba.deleteRow(i + 1);
    }
  }
}

// Recalcula o total de horas/dias do mês (ano+mês da data informada) de cada
// funcionário, varrendo a aba "Controle de Horas" daquela fazenda, e regrava
// a aba "<FAZENDA>_ResumoMensal" com os totais atualizados.
function atualizarResumoMensal(ss, fazenda, dataISO) {
  var p = String(dataISO).split('-');
  var ano = Number(p[0]);
  var mes = Number(p[1]);

  var abaHoras = ss.getSheetByName(fazenda);
  var linhas = abaHoras.getDataRange().getValues();
  var totais = {};
  for (var i = 1; i < linhas.length; i++) {
    var dataLinha = formatarData(linhas[i][0]);
    if (!dataLinha) continue;
    var partesLinha = dataLinha.split('-');
    if (Number(partesLinha[0]) !== ano || Number(partesLinha[1]) !== mes) continue;
    var nome = linhas[i][2];
    if (!totais[nome]) totais[nome] = { horas: 0, dias: 0 };
    totais[nome].horas += Number(linhas[i][5]) || 0;
    totais[nome].dias += Number(linhas[i][6]) || 0;
  }

  var nomeAbaResumo = fazenda + '_ResumoMensal';
  var abaResumo = ss.getSheetByName(nomeAbaResumo);
  if (!abaResumo) abaResumo = ss.insertSheet(nomeAbaResumo);
  if (abaResumo.getLastRow() === 0) abaResumo.appendRow(RESUMO_MENSAL_HEADERS);

  var linhasResumo = abaResumo.getDataRange().getValues();
  for (var j = linhasResumo.length - 1; j >= 1; j--) {
    if (Number(linhasResumo[j][0]) === ano && Number(linhasResumo[j][1]) === mes) {
      abaResumo.deleteRow(j + 1);
    }
  }
  Object.keys(totais).forEach(function (nome) {
    abaResumo.appendRow([ano, mes, nome, totais[nome].horas, totais[nome].dias]);
  });
}

// Copia o consumo de nutrição do dia pra planilha "Estoque Nutrição" (uma
// linha por produto consumido na mistura). Soma todas as misturas do dia em
// cada produto antes de gravar — se o dia tiver múltiplos envios, apaga o
// consumo anterior e regrava só o último, igual ao padrão do Controle de Horas.
function registrarConsumo(dados, fazenda) {
  if (!(dados.nutricaoMisturas || {}).houveMistura) return;
  var misturas = (dados.nutricaoMisturas || {}).misturas || [];
  if (!misturas.length || !dados.data) return;

  var consumoPorProduto = {};
  misturas.forEach(function (mistura) {
    Object.keys(mistura).forEach(function (produto) {
      var kg = Number(mistura[produto]) || 0;
      if (kg > 0) consumoPorProduto[produto] = (consumoPorProduto[produto] || 0) + kg;
    });
  });
  if (!Object.keys(consumoPorProduto).length) return;

  try {
    var ss = SpreadsheetApp.openById(ESTOQUE_SPREADSHEET_ID);
    var aba = getOrCreateAbaEstoque(ss, fazenda);

    removerConsumoDoDia(aba, dados.data);

    // Grava a data como objeto Date pra que as fórmulas da aba Dashboard
    // consigam filtrar por mês usando TEXT(Data,"yyyy-MM") no Sheets.
    var p = String(dados.data).split('-');
    var dataDate = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));

    Object.keys(consumoPorProduto).forEach(function (produto) {
      aba.appendRow([new Date(), dataDate, fazenda, produto, 'Consumo', '', consumoPorProduto[produto], '']);
    });
  } catch (err) {
    // Planilha de Estoque ainda não configurada — ignora silenciosamente.
  }
}

function getOrCreateAbaEstoque(ss, fazenda) {
  var nomeAba = fazenda || FALLBACK_SHEET_NAME;
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  if (aba.getLastRow() === 0) aba.appendRow(ESTOQUE_HEADERS);
  return aba;
}

// Remove linhas de Consumo do dia antes de regravar (evita duplicidade em
// múltiplos envios no mesmo dia). Usa a coluna Data (índice 1) como chave,
// não o Timestamp — pra não afetar entradas manuais de outras datas.
function removerConsumoDoDia(aba, dataISO) {
  var linhas = aba.getDataRange().getValues();
  for (var i = linhas.length - 1; i >= 1; i--) {
    if (formatarData(linhas[i][1]) === dataISO && linhas[i][4] === 'Consumo') {
      aba.deleteRow(i + 1);
    }
  }
}

// API de leitura para os painéis (Fase 4). Devolve os registros diários
// (já com o DadosJSON parseado) de uma fazenda dentro de um intervalo de
// datas. A agregação por dia/semana/mês é feita no front-end (relatorios.html)
// para manter este script simples e fácil de evoluir sem precisar reimplantar
// toda vez que um novo indicador for adicionado ao painel.
// Exemplo: ?fazenda=CB&inicio=2026-06-01&fim=2026-06-30
function doGet(e) {
  var params = (e && e.parameter) || {};

  // Rota separada para o painel de Estoque (não precisa de fazenda/inicio/fim).
  if (params.action === 'estoque') return doGetEstoque();

  var fazenda = (params.fazenda || '').trim();
  var inicio = params.inicio || '';
  var fim = params.fim || '';

  if (!fazenda || !inicio || !fim) {
    return jsonOutput({ ok: false, error: 'Parâmetros fazenda, inicio e fim são obrigatórios.', registros: [] });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(fazenda);
  if (!sheet) {
    return jsonOutput({ ok: true, registros: [], aviso: 'Aba "' + fazenda + '" não encontrada.' });
  }

  var linhas = sheet.getDataRange().getValues();

  // Modo de diagnóstico temporário: ?debug=1 devolve o que foi lido em cada
  // linha (sem filtrar por data), pra investigar incompatibilidade de formato.
  if (params.debug) {
    var diagnostico = [];
    for (var d = 1; d < linhas.length; d++) {
      diagnostico.push({
        linha: d + 1,
        valorBruto: String(linhas[d][2]),
        tipo: Object.prototype.toString.call(linhas[d][2]),
        dataFormatada: formatarData(linhas[d][2]),
        fazendaCelula: linhas[d][3]
      });
    }
    return jsonOutput({ ok: true, diagnostico: diagnostico });
  }

  var registros = [];
  for (var i = 1; i < linhas.length; i++) {
    var row = linhas[i];
    var dataRegistro = formatarData(row[2]);
    if (!dataRegistro || dataRegistro < inicio || dataRegistro > fim) continue;
    var dadosJson;
    try {
      dadosJson = JSON.parse(row[6] || '{}');
    } catch (err) {
      continue;
    }
    registros.push({ data: dataRegistro, responsavel: row[4] || '', dados: dadosJson });
  }
  registros.sort(function (a, b) { return a.data < b.data ? -1 : (a.data > b.data ? 1 : 0); });

  enriquecerComLinksTrello(registros, fazenda);

  return jsonOutput({ ok: true, registros: registros });
}

// Adiciona o campo "linkTrello" em cada solicitação (casando pelo ID) com o
// que estiver salvo na planilha "Solicitações". Envolto em try/catch pelo
// mesmo motivo do registrarSolicitacoes: não quebra o painel se essa planilha
// ainda não estiver configurada.
function enriquecerComLinksTrello(registros, fazenda) {
  var temSolicitacao = registros.some(function (r) {
    return ((r.dados.solicitacoesProdutos || {}).itens || []).length > 0;
  });
  if (!temSolicitacao) return;

  try {
    var ss = SpreadsheetApp.openById(SOLICITACOES_SPREADSHEET_ID);
    var aba = ss.getSheetByName(fazenda);
    if (!aba) return;

    var linhas = aba.getDataRange().getValues();
    var linkPorId = {};
    for (var i = 1; i < linhas.length; i++) {
      var idLinha = linhas[i][1];
      var link = linhas[i][7];
      if (idLinha && link) linkPorId[idLinha] = link;
    }

    registros.forEach(function (r) {
      var itens = (r.dados.solicitacoesProdutos || {}).itens || [];
      itens.forEach(function (item) {
        if (item.id && linkPorId[item.id]) item.linkTrello = linkPorId[item.id];
      });
    });
  } catch (err) {
    // Planilha de Solicitações ainda não configurada — segue sem o link.
  }
}

// Lê a planilha "Fazenda App - Estoque" e devolve, por fazenda e por produto,
// o estoque atual (kg e sacos) e o consumo médio dos últimos 3 meses completos.
function doGetEstoque() {
  var FAZENDAS = ['CB', 'PB', 'SH'];
  try {
    var ss = SpreadsheetApp.openById(ESTOQUE_SPREADSHEET_ID);

    // Lê TamanhoSaco da aba Referencia (col A = produto, col B = kg/saco).
    var tamanhoSaco = {};
    var refSheet = ss.getSheetByName('Referencia');
    if (refSheet) {
      var refData = refSheet.getDataRange().getValues();
      for (var r = 1; r < refData.length; r++) {
        var nomeProd = String(refData[r][0] || '').trim();
        if (nomeProd) tamanhoSaco[nomeProd] = Number(refData[r][1]) || 0;
      }
    }

    // Os 3 meses completos antes do mês atual (ex.: jul/26 → jun, mai, abr).
    var hoje = new Date();
    var ultimos3 = [];
    for (var m = 1; m <= 3; m++) {
      var d = new Date(hoje.getFullYear(), hoje.getMonth() - m, 1);
      var mn = d.getMonth() + 1;
      ultimos3.push(d.getFullYear() + '-' + (mn < 10 ? '0' : '') + mn);
    }

    var resultado = {};
    FAZENDAS.forEach(function (fazenda) {
      var aba = ss.getSheetByName(fazenda);
      resultado[fazenda] = {};
      if (!aba || aba.getLastRow() <= 1) return;

      var linhas = aba.getDataRange().getValues();
      var porProduto = {};

      for (var i = 1; i < linhas.length; i++) {
        var produto = String(linhas[i][3] || '').trim(); // col D
        var tipo    = String(linhas[i][4] || '').trim(); // col E
        var qtdKg   = Number(linhas[i][6]) || 0;         // col G
        var dataLinha = linhas[i][1];                     // col B

        // Se QtdKg estiver vazio, calcula a partir de QtdSacos × TamanhoSaco da Referencia.
        if (qtdKg === 0) {
          var qtdSacos = Number(linhas[i][5]) || 0;
          qtdKg = qtdSacos * (tamanhoSaco[produto] || 0);
        }
        if (!produto || !tipo || qtdKg === 0) continue;
        if (!porProduto[produto]) porProduto[produto] = { estoqueKg: 0, consumoPorMes: {} };

        if (tipo === 'Estoque Inicial' || tipo === 'Compra') {
          porProduto[produto].estoqueKg += qtdKg;
        } else if (tipo === 'Consumo' || tipo === 'Saída' || tipo === 'Saida') {
          porProduto[produto].estoqueKg -= qtdKg;
          if (tipo === 'Consumo') {
            var mesAno = formatarData(dataLinha).substring(0, 7); // "yyyy-MM"
            if (mesAno.length === 7) {
              porProduto[produto].consumoPorMes[mesAno] = (porProduto[produto].consumoPorMes[mesAno] || 0) + qtdKg;
            }
          }
        }
      }

      Object.keys(porProduto).forEach(function (produto) {
        var dp = porProduto[produto];
        // Conta só a sequência consecutiva de meses com dados a partir do mais
        // recente. Se um mês no meio estiver zerado, para — recomeça do mês
        // seguinte com consumo (ex: jul✓ jun✗ mai✓ → usa só julho).
        var totalConsec = 0;
        var mesesConsec = 0;
        for (var m = 0; m < ultimos3.length; m++) {
          var v = dp.consumoPorMes[ultimos3[m]] || 0;
          if (v === 0) break;
          totalConsec += v;
          mesesConsec++;
        }
        var consumoMedio = mesesConsec > 0 ? totalConsec / mesesConsec : 0;
        var saco = tamanhoSaco[produto] || 0;
        var estoqueArred = Math.round(dp.estoqueKg * 10) / 10;

        resultado[fazenda][produto] = {
          estoqueKg: estoqueArred,
          estoqueSacos: saco > 0 ? Math.round(estoqueArred / saco * 10) / 10 : null,
          consumoMedioMensal: Math.round(consumoMedio * 10) / 10,
          diasRestantes: consumoMedio > 0 ? Math.round(estoqueArred / (consumoMedio / 30)) : null,
          tamanhoSaco: saco
        };
      });
    });

    var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    return jsonOutput({ ok: true, estoque: resultado, atualizadoEm: agora });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message, estoque: {} });
  }
}

// A coluna "Data" pode conter texto ("2026-06-24") ou, se o Sheets converter
// automaticamente o valor ao salvar, um objeto Date — normaliza pros dois casos.
// Usa Object.prototype.toString em vez de "instanceof Date": o runtime do Apps
// Script devolve os valores de data de getValues() de um contexto diferente
// do construtor Date global, então "instanceof" falha mesmo sendo uma data.
function formatarData(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(valor || '');
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Grava vacinas/vermifugos/carrapaticidas do dia na aba DadosVacinas da
// planilha Saude Animal. Apaga e regrava para evitar duplicidade em
// múltiplos envios no mesmo dia (só o último prevalece).
function registrarVacinas(dados, fazenda) {
  var itens = (dados.vacinasVermifugos || {}).itens || [];
  if (!dados.data) return;
  try {
    var ss  = SpreadsheetApp.openById(SAUDE_ANIMAL_SPREADSHEET_ID);
    var aba = getOrCreateAbaSaudeAnimal(ss, 'DadosVacinas', VACINAS_HEADERS);
    removerSaudeAnimalDoDia(aba, dados.data, fazenda);
    if (!itens.length) return;
    var p = String(dados.data).split('-');
    var dataDate = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    itens.forEach(function (item) {
      var dose = item.dose ? (item.dose + (item.doseUnidade ? ' ' + item.doseUnidade : '')) : '';
      aba.appendRow([new Date(), dataDate, fazenda, item.categoria || '', item.qtdAnimais || '', item.produto || '', item.tipo || '', dose, '']);
    });
  } catch (err) {}
}

// Grava tratamentos com medicamentos do dia na aba DadosMedicamentos.
// Mesmo padrão: apaga e regrava para evitar duplicidade.
function registrarMedicamentos(dados, fazenda) {
  var tratamentos = (dados.sanidadeAnimal || {}).tratamentos || [];
  if (!dados.data) return;
  try {
    var ss  = SpreadsheetApp.openById(SAUDE_ANIMAL_SPREADSHEET_ID);
    var aba = getOrCreateAbaSaudeAnimal(ss, 'DadosMedicamentos', MEDICAMENTOS_HEADERS);
    removerSaudeAnimalDoDia(aba, dados.data, fazenda);
    if (!tratamentos.length) return;
    var p = String(dados.data).split('-');
    var dataDate = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    tratamentos.forEach(function (t) {
      var categorias = (t.categorias || []).join(', ');
      var sintomas   = (t.sintomas   || []).join(', ');
      var obs = t.responsavel ? 'Resp: ' + t.responsavel : '';
      aba.appendRow([new Date(), dataDate, fazenda, categorias, t.qtdAnimais || '', t.idAnimal || '', sintomas, t.tratamentoAplicado || '', obs]);
    });
  } catch (err) {}
}

function getOrCreateAbaSaudeAnimal(ss, nomeAba, headers) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  if (aba.getLastRow() === 0) aba.appendRow(headers);
  return aba;
}

// Remove linhas de uma data+fazenda específica (cols B e C, índices 1 e 2).
function removerSaudeAnimalDoDia(aba, dataISO, fazenda) {
  var linhas = aba.getDataRange().getValues();
  for (var i = linhas.length - 1; i >= 1; i--) {
    if (formatarData(linhas[i][1]) === dataISO && String(linhas[i][2]) === fazenda) {
      aba.deleteRow(i + 1);
    }
  }
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
