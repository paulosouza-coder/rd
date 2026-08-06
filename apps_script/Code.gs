/**
 * Integrador RD Station -> Google Sheets (Google Apps Script)
 *
 * Substitui o fluxo local (api_rdstation.py + save.py) por um script que roda
 * dentro da própria planilha, agendado por um gatilho de tempo. Sem custo,
 * sem depender de máquina local.
 *
 * SETUP:
 * 1. Na planilha do Google Sheets, abra "Extensões > Apps Script".
 * 2. Cole o conteúdo deste arquivo em "Code.gs" (substituindo o padrão).
 * 3. Rode a função `configurarToken` uma vez (ela vai pedir o token via prompt
 *    e salvá-lo em Script Properties, sem deixá-lo no código-fonte).
 * 4. Rode `sincronizarRDStation` uma vez manualmente para autorizar o script
 *    (Google vai pedir permissão de acesso à planilha e à internet).
 * 5. Rode `criarGatilhoHorario` uma vez para agendar a sincronização automática.
 * 6. Recarregue a planilha: vai aparecer um menu "RD Station" com as mesmas
 *    ações, caso queira rodar manualmente depois.
 */

// ============================================================
// CONFIGURAÇÃO
// ============================================================
// Cole seu token do RD Station CRM na linha abaixo, entre as aspas, SE
// preferir deixá-lo fixo aqui no código (mais simples de configurar).
// Ex: var RD_API_TOKEN = 'a1b2c3d4e5';
//
// Atenção: quem tiver acesso a este projeto do Apps Script (ou a uma cópia
// deste arquivo) consegue ver o token em texto puro. Não cole o token em
// nenhum lugar fora daqui (chat, repositório Git, etc).
//
// Se preferir não deixar o token fixo no código, deixe esta linha vazia
// (''): o script vai pedir o token por uma caixa de diálogo na primeira
// vez que rodar (menu "RD Station > Configurar token da API"), e vai
// guardá-lo separado do código, no Script Properties do Google.
var RD_API_TOKEN = '';
// ============================================================

var BASE_URL = 'https://crm.rdstation.com/api/v1/';
var TOKEN_PROPERTY = 'RD_API_TOKEN';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RD Station')
    .addItem('Sincronizar agora', 'sincronizarRDStation')
    .addItem('Configurar token da API', 'configurarToken')
    .addItem('Agendar sincronização automática (1x por hora)', 'criarGatilhoHorario')
    .addItem('Depurar estrutura da API (ver Logs)', 'depurarEstrutura_')
    .addItem('Depurar Funis/Etapas (ver Logs)', 'depurarFunis_')
    .addToUi();
}

/**
 * Diagnóstico específico do problema "só aparece 1 funil": lista, para todos
 * os deals coletados, as combinações distintas de etapa/funil encontradas
 * (e quantas vezes cada uma aparece), e mostra a lista completa retornada
 * por /deal_stages — para descobrir se esse endpoint só devolve as etapas
 * de um único funil (o que explicaria o problema).
 */
function depurarFunis_() {
  var deals = fetchAllPages_('deals', 'deals');
  var stages = fetchAllPages_('deal_stages', 'deal_stages');

  Logger.log('Total de deals coletados: ' + deals.length);

  var combos = {};
  deals.forEach(function (deal) {
    var chave = JSON.stringify({
      deal_stage_nome: get_(deal, 'deal_stage.name'),
      deal_stage_id: get_(deal, 'deal_stage.id'),
      deal_pipeline_direto: get_(deal, 'deal_pipeline.name'),
      deal_pipeline_aninhado: get_(deal, 'deal_stage.deal_pipeline.name')
    });
    combos[chave] = (combos[chave] || 0) + 1;
  });

  Logger.log('===== COMBINAÇÕES DISTINTAS DE ETAPA/FUNIL NOS DEALS =====');
  Object.keys(combos).forEach(function (chave) {
    Logger.log(combos[chave] + 'x -> ' + chave);
  });

  Logger.log('===== TODAS AS ETAPAS RETORNADAS POR /deal_stages (total: ' + stages.length + ') =====');
  Logger.log(JSON.stringify(stages, null, 2));

  Object.keys(OUTROS_FUNIS_POR_ID).forEach(function (pipelineId) {
    var pipeline = buscarFunilPorId_(pipelineId);
    Logger.log('===== GET /deal_pipelines/' + pipelineId + ' (' + OUTROS_FUNIS_POR_ID[pipelineId] + ') =====');
    Logger.log(JSON.stringify(pipeline, null, 2));
  });

  SpreadsheetApp.getUi().alert('Diagnóstico concluído! Veja o resultado em "Execuções" no editor do Apps Script e me envie o conteúdo.');
}

/**
 * Busca a primeira página de deals, tasks, deal_stages e organizations e
 * despeja a estrutura bruta no Log de execução, para conferir os nomes
 * reais dos campos (ex: onde fica o funil/pipeline, campos personalizados,
 * classificação de saúde da conta) direto na sua conta do RD Station.
 * Rode e veja o resultado em: editor do Apps Script > "Execuções" (ícone
 * de lista à esquerda) ou Ver > Registros de execução.
 */
function depurarEstrutura_() {
  var token = getToken_();

  function primeiraPagina_(endpoint, key) {
    try {
      var url = BASE_URL + endpoint + '?page=1&token=' + encodeURIComponent(token);
      var response = UrlFetchApp.fetch(url, { method: 'get', contentType: 'application/json', muteHttpExceptions: true });
      var body = JSON.parse(response.getContentText());
      return key ? body[key] : body;
    } catch (e) {
      return { erro: 'Não foi possível buscar "' + endpoint + '": ' + e.message };
    }
  }

  var deals = primeiraPagina_('deals', 'deals');
  var tasks = primeiraPagina_('tasks', 'tasks');
  var stages = primeiraPagina_('deal_stages', 'deal_stages');
  var orgs = primeiraPagina_('organizations', 'organizations');

  Logger.log('===== EXEMPLO DE DEAL =====\n' + JSON.stringify((deals && deals[0]) || deals, null, 2));
  Logger.log('===== EXEMPLO DE TASK =====\n' + JSON.stringify((tasks && tasks[0]) || tasks, null, 2));
  Logger.log('===== ETAPAS DE FUNIL (deal_stages) =====\n' + JSON.stringify(stages, null, 2));
  Logger.log('===== EXEMPLO DE ORGANIZAÇÃO =====\n' + JSON.stringify((orgs && orgs[0]) || orgs, null, 2));

  SpreadsheetApp.getUi().alert('Depuração concluída! No editor do Apps Script, abra "Execuções" (ícone de lista à esquerda) para ver os dados coletados.');
}

function configurarToken() {
  var ui = SpreadsheetApp.getUi();
  var resposta = ui.prompt('Token da API do RD Station CRM', 'Cole aqui o token (Configurações > Integrações no RD Station):', ui.ButtonSet.OK_CANCEL);
  if (resposta.getSelectedButton() !== ui.Button.OK) return;
  var token = resposta.getResponseText().trim();
  if (!token) {
    ui.alert('Token vazio, nada foi salvo.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(TOKEN_PROPERTY, token);
  ui.alert('Token salvo com sucesso.');
}

function getToken_() {
  if (RD_API_TOKEN) return RD_API_TOKEN;

  var token = PropertiesService.getScriptProperties().getProperty(TOKEN_PROPERTY);
  if (!token) {
    throw new Error('Token da API não configurado. Preencha RD_API_TOKEN no topo do código, ou rode "RD Station > Configurar token da API" primeiro.');
  }
  return token;
}

/** Busca todas as páginas de um endpoint, igual ao get_all_pages do api_rdstation.py */
function fetchAllPages_(endpoint, key) {
  var token = getToken_();
  var allData = [];
  var page = 1;

  while (true) {
    var url = BASE_URL + endpoint + '?page=' + page + '&token=' + encodeURIComponent(token);
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      contentType: 'application/json',
      muteHttpExceptions: true
    });

    var body;
    try {
      body = JSON.parse(response.getContentText());
    } catch (e) {
      throw new Error('Erro ao converter resposta JSON de ' + endpoint + ': ' + response.getContentText());
    }

    if (body && body.message) {
      throw new Error('Erro ao buscar ' + endpoint + ': ' + body.message);
    }

    var data = (key && body[key]) ? body[key] : body;
    allData = allData.concat(data);

    if (!body.has_more) break;
    page += 1;
  }

  return allData;
}

/** Acesso seguro a campos aninhados, tipo get(obj, "a.b.c", "") em Python */
function get_(obj, path, fallback) {
  fallback = fallback === undefined ? '' : fallback;
  var parts = path.split('.');
  var current = obj;
  for (var i = 0; i < parts.length; i++) {
    if (current === null || current === undefined) return fallback;
    current = current[parts[i]];
  }
  return current === null || current === undefined ? fallback : current;
}

var FORMATO_DATA_HORA = 'dd/mm/yyyy hh:mm:ss';

/** Converte uma string de data da API (ISO 8601) em um Date real do Sheets */
function parseData_(valor) {
  if (!valor) return '';
  var data = new Date(valor);
  if (isNaN(data.getTime())) return valor; // não era uma data válida, mantém como texto
  return data;
}

function sincronizarRDStation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tasks = fetchAllPages_('tasks', 'tasks');
  var deals = fetchAllPages_('deals', 'deals');

  escreverTasks_(ss, tasks);
  escreverDeals_(ss, deals);
}

function escreverTasks_(ss, tasks) {
  var headers = ['ID', 'deal_id', 'Assunto', 'Tipo', 'Hora', 'Status', 'Data', 'Criado em',
    'Feito', 'Data Conclusão', 'Negócio', 'Rating', 'Usuário', 'Email Usuário'];

  var rows = tasks.map(function (task) {
    var usuario = (task.users && task.users.length > 0) ? task.users[0] : {};
    return [
      get_(task, 'id'),
      get_(task, 'deal_id'),
      get_(task, 'subject'),
      get_(task, 'type'),
      get_(task, 'hour'),
      get_(task, 'status'),
      parseData_(get_(task, 'date')),
      parseData_(get_(task, 'created_at')),
      get_(task, 'done'),
      parseData_(get_(task, 'done_date')),
      get_(task, 'deal.name'),
      get_(task, 'deal.rating'),
      get_(usuario, 'name'),
      get_(usuario, 'email')
    ];
  });

  // Colunas de data/hora (1-based): 7=Data, 8=Criado em, 10=Data Conclusão
  escreverAba_(ss, 'Tasks', headers, rows, [7, 8, 10]);
}

function escreverDeals_(ss, deals) {
  var baseHeaders = ['ID', 'Nome', 'Valor Total', 'Data de Criação', 'Última Atualização',
    'Organização', 'Endereço', 'Usuário Responsável', 'Email Usuário', 'Funil', 'Estágio',
    'Ordem Etapa', 'Status Negociação', 'Fonte', 'Campanha', 'Próxima Tarefa', 'Data Próxima Tarefa'];

  var mapaEtapas = buscarMapaEtapas_();

  // Descobre dinamicamente todos os labels de campos personalizados usados nos deals
  var customLabels = [];
  var seen = {};
  deals.forEach(function (deal) {
    (deal.deal_custom_fields || []).forEach(function (cf) {
      var label = get_(cf, 'custom_field.label');
      if (label && !seen[label]) {
        seen[label] = true;
        customLabels.push(label);
      }
    });
  });

  var headers = baseHeaders.concat(customLabels);

  var rows = deals.map(function (deal) {
    var row = [
      get_(deal, 'id'),
      get_(deal, 'name'),
      get_(deal, 'amount_total'),
      parseData_(get_(deal, 'created_at')),
      parseData_(get_(deal, 'updated_at')),
      get_(deal, 'organization.name'),
      get_(deal, 'organization.address'),
      get_(deal, 'user.name'),
      get_(deal, 'user.email'),
      nomeFunil_(deal, mapaEtapas),
      get_(deal, 'deal_stage.name'),
      ordemEtapa_(deal, mapaEtapas),
      statusNegociacao_(deal),
      get_(deal, 'deal_source.name'),
      get_(deal, 'campaign.name'),
      get_(deal, 'next_task.subject'),
      parseData_(get_(deal, 'next_task.date'))
    ];

    var customValues = {};
    (deal.deal_custom_fields || []).forEach(function (cf) {
      var label = get_(cf, 'custom_field.label');
      if (label) customValues[label] = get_(cf, 'value');
    });
    customLabels.forEach(function (label) {
      row.push(customValues.hasOwnProperty(label) ? customValues[label] : '');
    });

    return row;
  });

  // Colunas de data/hora (1-based): 4=Data de Criação, 5=Última Atualização, 17=Data Próxima Tarefa
  escreverAba_(ss, 'Deals', headers, rows, [4, 5, 17]);
}

/** Deriva o status da negociação a partir dos campos "win" e "deal_lost_reason" da API */
function statusNegociacao_(deal) {
  if (deal.win === true) return 'Vendida';
  if (deal.win === false || get_(deal, 'deal_lost_reason.name')) return 'Perdida';
  return 'Em andamento';
}

// /deal_stages sem filtro só retorna as etapas do funil "padrão" da conta (confirmado por
// depuração). Os demais funis precisam ser buscados explicitamente por ID via
// /deal_pipelines/{id}. IDs específicos desta conta RD Station:
var OUTROS_FUNIS_POR_ID = {
  '67cb2f7d04bf6d00167e4fc4': 'Sucesso do Cliente',
  '67df427fb6a6ee0028d86cae': 'Gestão de Contratos'
};

/** Mapa id_da_etapa -> {nome, ordem}, cobrindo o funil padrão (via /deal_stages) e os
 *  demais funis da conta (via /deal_pipelines/{id}), como fallback para quando o funil
 *  não vem aninhado diretamente no deal. A ordem é usada para desenhar o funil no
 *  dashboard na sequência certa das etapas. */
function buscarMapaEtapas_() {
  var mapa = {};

  var stagesPadrao = fetchAllPages_('deal_stages', 'deal_stages');
  (stagesPadrao || []).forEach(function (stage) {
    var nomeFunil = get_(stage, 'deal_pipeline.name') || get_(stage, 'deal_pipeline_name');
    if (stage && stage.id && nomeFunil) {
      mapa[stage.id] = { nome: nomeFunil, ordem: get_(stage, 'order', 0) };
    }
  });

  Object.keys(OUTROS_FUNIS_POR_ID).forEach(function (pipelineId) {
    var nomeFunil = OUTROS_FUNIS_POR_ID[pipelineId];
    var pipeline = buscarFunilPorId_(pipelineId);
    var etapas = (pipeline && (pipeline.deal_stages || pipeline.stages)) || [];
    etapas.forEach(function (etapa) {
      if (etapa && etapa.id) mapa[etapa.id] = { nome: nomeFunil, ordem: get_(etapa, 'order', 0) };
    });
  });

  return mapa;
}

/** Busca um funil específico por ID (GET /deal_pipelines/{id}) */
function buscarFunilPorId_(pipelineId) {
  var token = getToken_();
  var url = BASE_URL + 'deal_pipelines/' + encodeURIComponent(pipelineId) + '?token=' + encodeURIComponent(token);
  var response = UrlFetchApp.fetch(url, { method: 'get', contentType: 'application/json', muteHttpExceptions: true });
  try {
    return JSON.parse(response.getContentText());
  } catch (e) {
    return null;
  }
}

/** Tenta descobrir o nome do funil do deal em diferentes formatos possíveis da API,
 *  usando o mapa de etapas como último recurso. */
function nomeFunil_(deal, mapaEtapas) {
  var direto = get_(deal, 'deal_pipeline.name');
  if (direto) return direto;

  var aninhado = get_(deal, 'deal_stage.deal_pipeline.name');
  if (aninhado) return aninhado;

  var stageId = get_(deal, 'deal_stage.id');
  if (stageId && mapaEtapas[stageId]) return mapaEtapas[stageId].nome;

  return '';
}

/** Posição da etapa dentro do seu funil (para ordenar o funil no dashboard) */
function ordemEtapa_(deal, mapaEtapas) {
  var stageId = get_(deal, 'deal_stage.id');
  if (stageId && mapaEtapas[stageId]) return mapaEtapas[stageId].ordem;
  return '';
}

function escreverAba_(ss, nomeAba, headers, rows, colunasData) {
  var sheet = ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    (colunasData || []).forEach(function (coluna) {
      sheet.getRange(2, coluna, rows.length, 1).setNumberFormat(FORMATO_DATA_HORA);
    });
  }
}

// ============================================================
// DASHBOARD "GESTÃO À VISTA" (Web App)
// ============================================================
// Painel público (sem valores em R$) publicado como Web App do Apps Script.
// Para publicar: no editor, "Implantar" > "Nova implantação" > tipo "App da Web".
// Executar como: "Eu". Quem tem acesso: escolha conforme a sensibilidade
// (ex: "Qualquer pessoa da [seu domínio]"). Isso gera uma URL fixa que pode
// ser aberta numa TV/monitor.

var FUNIS_DISPONIVEIS = ['Todos', 'Vendas Consultiva', 'Sucesso do Cliente', 'Gestão de Contratos'];

// Metas do formulário de KPIs (Win Rate e Taxa de Conversão), usadas só para
// colorir os cards do dashboard (sem expor valor financeiro nenhum).
var METAS_KPI = {
  winRate: { meta: 0.35, minimoAceitavel: 0.25 },
  taxaConversao: { meta: 0.35, minimoAceitavel: 0.30 }
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Gestão à Vista — Comercial')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lê uma aba inteira e devolve como lista de objetos {cabeçalho: valor} */
function lerAbaComoObjetos_(nomeAba) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var valores = sheet.getDataRange().getValues();
  var cabecalhos = valores[0];
  return valores.slice(1).map(function (linha) {
    var obj = {};
    cabecalhos.forEach(function (cabecalho, i) { obj[cabecalho] = linha[i]; });
    return obj;
  });
}

function limitesPeriodo_(periodo) {
  var agora = new Date();
  if (periodo === 'trimestre') {
    var inicioTrimestre = new Date(agora.getFullYear(), Math.floor(agora.getMonth() / 3) * 3, 1);
    return { inicio: inicioTrimestre, fim: agora };
  }
  if (periodo === 'tudo') {
    return { inicio: null, fim: agora };
  }
  var inicioMes = new Date(agora.getFullYear(), agora.getMonth(), 1);
  return { inicio: inicioMes, fim: agora };
}

function dentroDoPeriodo_(data, limites) {
  if (!(data instanceof Date)) return false;
  if (limites.inicio && data < limites.inicio) return false;
  if (data > limites.fim) return false;
  return true;
}

/** Função chamada pelo dashboard (google.script.run) para buscar os dados já calculados */
function getDadosDashboard(funilSelecionado, periodo) {
  var deals = lerAbaComoObjetos_('Deals');
  var tasks = lerAbaComoObjetos_('Tasks');
  var limites = limitesPeriodo_(periodo);

  if (funilSelecionado && funilSelecionado !== 'Todos') {
    deals = deals.filter(function (d) { return d['Funil'] === funilSelecionado; });
  }

  var abertos = deals.filter(function (d) { return d['Status Negociação'] === 'Em andamento'; });
  var vendidosPeriodo = deals.filter(function (d) {
    return d['Status Negociação'] === 'Vendida' && dentroDoPeriodo_(d['Última Atualização'], limites);
  });
  var perdidosPeriodo = deals.filter(function (d) {
    return d['Status Negociação'] === 'Perdida' && dentroDoPeriodo_(d['Última Atualização'], limites);
  });
  var criadosPeriodo = deals.filter(function (d) { return dentroDoPeriodo_(d['Data de Criação'], limites); });
  var vendidosCriadosPeriodo = criadosPeriodo.filter(function (d) { return d['Status Negociação'] === 'Vendida'; });

  var fechadosPeriodo = vendidosPeriodo.length + perdidosPeriodo.length;
  var winRate = fechadosPeriodo > 0 ? vendidosPeriodo.length / fechadosPeriodo : 0;
  var taxaConversao = criadosPeriodo.length > 0 ? vendidosCriadosPeriodo.length / criadosPeriodo.length : 0;

  var porEtapa = {};
  abertos.forEach(function (d) {
    var etapa = d['Estágio'] || 'Sem etapa';
    var ordem = Number(d['Ordem Etapa']) || 0;
    if (!porEtapa[etapa]) porEtapa[etapa] = { etapa: etapa, ordem: ordem, quantidade: 0 };
    porEtapa[etapa].quantidade++;
  });
  var funil = Object.keys(porEtapa)
    .map(function (k) { return porEtapa[k]; })
    .sort(function (a, b) { return a.ordem - b.ordem; });

  var tarefasPendentes = tasks.filter(function (t) { return t['Feito'] === false; });
  var hoje = new Date();
  var tarefasAtrasadas = tarefasPendentes.filter(function (t) {
    return t['Data'] instanceof Date && t['Data'] < hoje;
  });

  return {
    atualizadoEm: Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
    funisDisponiveis: FUNIS_DISPONIVEIS,
    metas: METAS_KPI,
    negociosAbertos: abertos.length,
    negociosVendidos: vendidosPeriodo.length,
    negociosPerdidos: perdidosPeriodo.length,
    winRate: winRate,
    taxaConversao: taxaConversao,
    tarefasPendentes: tarefasPendentes.length,
    tarefasAtrasadas: tarefasAtrasadas.length,
    funil: funil
  };
}

/** Cria um gatilho para rodar sincronizarRDStation automaticamente a cada hora */
function criarGatilhoHorario() {
  removerGatilhos_();
  ScriptApp.newTrigger('sincronizarRDStation')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('Sincronização automática agendada para rodar a cada hora.');
}

function removerGatilhos_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'sincronizarRDStation') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
