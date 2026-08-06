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
    .addItem('Criar/Atualizar aba METAS', 'garantirAbaMetas_')
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
    'Organização', 'Endereço', 'Usuário Responsável', 'ID Usuário Responsável', 'Email Usuário',
    'Funil', 'ID Funil', 'Estágio', 'Ordem Etapa', 'Status Negociação', 'Fonte', 'Campanha',
    'Próxima Tarefa', 'Data Próxima Tarefa'];

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
      get_(deal, 'user.id'),
      get_(deal, 'user.email'),
      nomeFunil_(deal, mapaEtapas),
      idFunil_(deal, mapaEtapas),
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

  // Colunas de data/hora (1-based): 4=Data de Criação, 5=Última Atualização, 19=Data Próxima Tarefa
  escreverAba_(ss, 'Deals', headers, rows, [4, 5, 19]);
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

// Mapa completo nome -> ID do funil (para gravar "ID Funil" nos deals e casar
// com a coluna ID_Funil da aba METAS).
var FUNIL_NOME_PARA_ID = {
  'Vendas Consultiva': '67cd09d3a4c0590017632c38',
  'Sucesso do Cliente': '67cb2f7d04bf6d00167e4fc4',
  'Gestão de Contratos': '67df427fb6a6ee0028d86cae'
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

/** ID do funil do deal (para casar com a coluna ID_Funil da aba METAS) */
function idFunil_(deal, mapaEtapas) {
  var nome = nomeFunil_(deal, mapaEtapas);
  return FUNIL_NOME_PARA_ID[nome] || '';
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
// Painel público publicado como Web App do Apps Script.
//
// RESTRIÇÃO DE EXPOSIÇÃO DE VALORES: este dashboard e a função
// getDadosDashboard() abaixo NUNCA devem ler ou devolver campos monetários
// (amount_total / "Valor Total" ou qualquer outro valor em R$), mesmo que
// existam na aba Deals. Só quantidades, percentuais e prazos em dias.
//
// Para publicar: no editor, "Implantar" > "Nova implantação" > tipo "App da
// Web". Executar como: "Eu". Quem tem acesso: escolha conforme a
// sensibilidade (ex: "Qualquer pessoa da [seu domínio]").

var FUNIS_DISPONIVEIS = ['Todos', 'Vendas Consultiva', 'Sucesso do Cliente', 'Gestão de Contratos'];

// Nomes canônicos de indicador aceitos na coluna Nome_Indicador da aba METAS
// (lista de "Indicadores iniciais para comparação" definida com o usuário).
var INDICADORES_CONHECIDOS = [
  'Quantidade de oportunidades criadas',
  'Quantidade de empresas prospectadas',
  'Quantidade de tarefas concluídas',
  'Quantidade de reuniões realizadas',
  'Quantidade de propostas enviadas',
  'Quantidade de negócios ganhos',
  'Taxa de conversão',
  'Taxa de cumprimento de tarefas',
  'Tempo médio de permanência no funil',
  'Tempo médio de resposta',
  'Quantidade de tarefas atrasadas',
  'Percentual de oportunidades sem atividade'
];

var CABECALHOS_METAS = [
  'ID_Meta', 'Nome_Indicador', 'Tipo_Indicador', 'Ano', 'Mês', 'Trimestre',
  'ID_Funil', 'ID_Responsável', 'ID_Empresa', 'Meta_Quantidade', 'Meta_Percentual',
  'Meta_Prazo_Dias', 'Sentido_Indicador', 'Faixa_Atencao', 'Faixa_Critica',
  'Data_Inicio_Vigencia', 'Data_Fim_Vigencia', 'Ativo', 'Observação',
  'Atualizado_Por', 'Data_Atualização'
];

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

/** Cria a aba METAS (se ainda não existir) com a estrutura definida, validações e um
 *  exemplo de preenchimento. Rode pelo menu "RD Station > Criar/Atualizar aba METAS". */
function garantirAbaMetas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('METAS');
  var jaExistia = !!sheet;
  if (!sheet) sheet = ss.insertSheet('METAS');

  sheet.getRange(1, 1, 1, CABECALHOS_METAS.length).setValues([CABECALHOS_METAS]);
  sheet.setFrozenRows(1);

  if (!jaExistia) {
    // Linha de exemplo (meta geral de Taxa de conversão para o mês atual, sem
    // funil/responsável específico = meta da equipe toda).
    var agora = new Date();
    sheet.getRange(2, 1, 1, CABECALHOS_METAS.length).setValues([[
      'EXEMPLO-001', 'Taxa de conversão', 'Percentual', agora.getFullYear(), agora.getMonth() + 1, '',
      '', '', '', '', 0.35,
      '', 'MAIOR_MELHOR', 0.8, 0.6,
      '', '', true, 'Linha de exemplo — pode apagar esta linha',
      Session.getActiveUser().getEmail(), agora
    ]]);
  }

  // Validações (dropdown) nas colunas de tipo fixo
  var ultimaLinha = Math.max(sheet.getMaxRows(), 200);
  sheet.getRange(2, 2, ultimaLinha - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(INDICADORES_CONHECIDOS, true).setAllowInvalid(true).build()
  );
  sheet.getRange(2, 3, ultimaLinha - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(
      ['Quantidade', 'Percentual', 'Prazo', 'Conversão', 'Produtividade', 'Atividade', 'Qualidade'], true
    ).build()
  );
  sheet.getRange(2, 13, ultimaLinha - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['MAIOR_MELHOR', 'MENOR_MELHOR'], true).build()
  );
  sheet.getRange(2, 18, ultimaLinha - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );

  sheet.getRange(1, 7).setNote(
    'IDs dos funis conhecidos:\n' +
    Object.keys(FUNIL_NOME_PARA_ID).map(function (nome) { return FUNIL_NOME_PARA_ID[nome] + ' = ' + nome; }).join('\n')
  );
  sheet.getRange(1, 8).setNote('ID do usuário responsável no RD Station (coluna "ID Usuário Responsável" na aba Deals). Deixe em branco para meta geral da equipe.');
  sheet.getRange(1, 4).setNote('Preencha Mês OU Trimestre (não os dois) para uma meta específica do período. Deixe ambos em branco para uma meta ANUAL (usada só quando não há meta mais específica pro período).');

  sheet.autoResizeColumns(1, CABECALHOS_METAS.length);

  SpreadsheetApp.getUi().alert(
    jaExistia
      ? 'Aba METAS já existia — cabeçalhos e validações atualizados.'
      : 'Aba METAS criada com uma linha de exemplo. Preencha suas metas reais e apague a linha de exemplo.'
  );
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

/** true se a data de hoje está dentro da janela de vigência da meta (campos podem ficar vazios = sem limite) */
function metaVigente_(meta) {
  var hoje = new Date();
  if (meta.Data_Inicio_Vigencia instanceof Date && hoje < meta.Data_Inicio_Vigencia) return false;
  if (meta.Data_Fim_Vigencia instanceof Date && hoje > meta.Data_Fim_Vigencia) return false;
  return true;
}

function metaCobrePeriodoEspecifico_(meta, contexto) {
  if (Number(meta.Ano) !== contexto.ano) return false;
  if (contexto.periodo === 'mes') return Number(meta.Mês) === contexto.mes;
  if (contexto.periodo === 'trimestre') return Number(meta.Trimestre) === contexto.trimestre;
  return false;
}

function metaAnual_(meta, contexto) {
  return Number(meta.Ano) === contexto.ano && !meta.Mês && !meta.Trimestre;
}

/**
 * Busca a meta aplicável para um indicador, seguindo a ordem de prioridade definida:
 * 1) responsável + funil + período · 2) responsável + período · 3) funil + período ·
 * 4) geral + período · 5) meta anual do indicador · 6) sem meta cadastrada (null).
 */
function buscarMeta_(nomeIndicador, contexto, todasMetas) {
  var ativas = todasMetas.filter(function (m) {
    return m.Ativo === true && m.Nome_Indicador === nomeIndicador && metaVigente_(m);
  });

  if (contexto.periodo !== 'tudo') {
    var doPeriodo = ativas.filter(function (m) { return metaCobrePeriodoEspecifico_(m, contexto); });

    var comFunil = contexto.funilId ? doPeriodo.filter(function (m) { return m.ID_Funil === contexto.funilId; }) : [];
    var semFunil = doPeriodo.filter(function (m) { return !m.ID_Funil; });
    var comResp = contexto.responsavelId ? doPeriodo.filter(function (m) { return String(m['ID_Responsável']) === String(contexto.responsavelId); }) : [];

    var nivel1 = comResp.filter(function (m) { return contexto.funilId && m.ID_Funil === contexto.funilId; })[0];
    if (nivel1) return nivel1;

    var nivel2 = comResp.filter(function (m) { return !m.ID_Funil; })[0];
    if (nivel2) return nivel2;

    var nivel3 = comFunil.filter(function (m) { return !m['ID_Responsável']; })[0];
    if (nivel3) return nivel3;

    var nivel4 = semFunil.filter(function (m) { return !m['ID_Responsável']; })[0];
    if (nivel4) return nivel4;
  }

  var anual = ativas.filter(function (m) {
    return metaAnual_(m, contexto) && !m.ID_Funil && !m['ID_Responsável'];
  })[0];
  if (anual) return anual;

  return null;
}

/** Extrai o valor-alvo numérico da meta, conforme o Tipo_Indicador */
function valorAlvoDaMeta_(meta) {
  if (!meta) return null;
  if (meta.Tipo_Indicador === 'Percentual' || meta.Tipo_Indicador === 'Conversão') return Number(meta.Meta_Percentual) || null;
  if (meta.Tipo_Indicador === 'Prazo') return Number(meta.Meta_Prazo_Dias) || null;
  return Number(meta.Meta_Quantidade) || null;
}

/** % de atingimento, respeitando o sentido (maior-melhor / menor-melhor) */
function calcularAtingimento_(resultado, meta) {
  var alvo = valorAlvoDaMeta_(meta);
  if (alvo === null || alvo <= 0) return null;
  if (meta.Sentido_Indicador === 'MENOR_MELHOR') {
    return resultado > 0 ? alvo / resultado : 1;
  }
  return resultado / alvo;
}

function statusPorFaixas_(atingimento, meta) {
  if (atingimento === null || !meta) return 'neutro';
  var faixaCritica = Number(meta.Faixa_Critica) || 0;
  var faixaAtencao = Number(meta.Faixa_Atencao) || 0;
  if (faixaCritica && atingimento < faixaCritica) return 'critico';
  if (faixaAtencao && atingimento < faixaAtencao) return 'atencao';
  return 'bom';
}

/** Monta o objeto {resultado, meta, atingimento, status} devolvido ao dashboard para um indicador */
function montarIndicadorComMeta_(nomeIndicador, resultado, contexto, todasMetas) {
  var meta = buscarMeta_(nomeIndicador, contexto, todasMetas);
  var atingimento = meta ? calcularAtingimento_(resultado, meta) : null;
  return {
    resultado: resultado,
    metaEncontrada: !!meta,
    valorAlvo: valorAlvoDaMeta_(meta),
    tipoIndicador: meta ? meta.Tipo_Indicador : null,
    atingimento: atingimento,
    status: statusPorFaixas_(atingimento, meta)
  };
}

/** Função chamada pelo dashboard (google.script.run) para buscar os dados já calculados.
 *  IMPORTANTE: nunca ler/expor "Valor Total" (amount_total) aqui — só quantidades, % e dias. */
function getDadosDashboard(funilSelecionado, periodo) {
  var deals = lerAbaComoObjetos_('Deals');
  var tasks = lerAbaComoObjetos_('Tasks');
  var todasMetas = lerAbaComoObjetos_('METAS');
  var limites = limitesPeriodo_(periodo);

  var idFunilSelecionado = (funilSelecionado && funilSelecionado !== 'Todos')
    ? FUNIL_NOME_PARA_ID[funilSelecionado] : '';

  if (funilSelecionado && funilSelecionado !== 'Todos') {
    deals = deals.filter(function (d) { return d['Funil'] === funilSelecionado; });
  }

  var agora = new Date();
  var contexto = {
    ano: agora.getFullYear(),
    mes: agora.getMonth() + 1,
    trimestre: Math.floor(agora.getMonth() / 3) + 1,
    periodo: periodo,
    funilId: idFunilSelecionado,
    responsavelId: null
  };

  var abertos = deals.filter(function (d) { return d['Status Negociação'] === 'Em andamento'; });
  var vendidosPeriodo = deals.filter(function (d) {
    return d['Status Negociação'] === 'Vendida' && dentroDoPeriodo_(d['Última Atualização'], limites);
  });
  var criadosPeriodo = deals.filter(function (d) { return dentroDoPeriodo_(d['Data de Criação'], limites); });
  var vendidosCriadosPeriodo = criadosPeriodo.filter(function (d) { return d['Status Negociação'] === 'Vendida'; });

  var taxaConversao = criadosPeriodo.length > 0 ? vendidosCriadosPeriodo.length / criadosPeriodo.length : 0;

  // Empresas em acompanhamento: organizações distintas com negócio em aberto
  var empresas = {};
  abertos.forEach(function (d) {
    var nomeOrg = d['Organização'];
    if (nomeOrg) empresas[nomeOrg] = true;
  });

  // Tempo médio no funil: idade média (em dias) dos negócios em aberto — não temos
  // histórico de troca de etapa via API, então é a permanência desde a criação.
  var idadesDias = abertos
    .filter(function (d) { return d['Data de Criação'] instanceof Date; })
    .map(function (d) { return (agora - d['Data de Criação']) / (1000 * 60 * 60 * 24); });
  var tempoMedioFunilDias = idadesDias.length > 0
    ? idadesDias.reduce(function (soma, dias) { return soma + dias; }, 0) / idadesDias.length
    : 0;

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
  var tarefasAtrasadas = tarefasPendentes.filter(function (t) {
    return t['Data'] instanceof Date && t['Data'] < agora;
  });

  return {
    atualizadoEm: Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
    funisDisponiveis: FUNIS_DISPONIVEIS,

    oportunidadesAbertas: montarIndicadorComMeta_('Quantidade de oportunidades criadas', abertos.length, contexto, todasMetas),
    empresasAcompanhamento: montarIndicadorComMeta_('Quantidade de empresas prospectadas', Object.keys(empresas).length, contexto, todasMetas),
    taxaConversao: montarIndicadorComMeta_('Taxa de conversão', taxaConversao, contexto, todasMetas),
    negociosGanhos: montarIndicadorComMeta_('Quantidade de negócios ganhos', vendidosPeriodo.length, contexto, todasMetas),
    tarefasPendentes: { resultado: tarefasPendentes.length },
    tarefasAtrasadas: montarIndicadorComMeta_('Quantidade de tarefas atrasadas', tarefasAtrasadas.length, contexto, todasMetas),
    tempoMedioFunil: montarIndicadorComMeta_('Tempo médio de permanência no funil', tempoMedioFunilDias, contexto, todasMetas),

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
