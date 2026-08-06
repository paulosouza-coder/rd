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
    .addToUi();
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
      get_(task, 'date'),
      get_(task, 'created_at'),
      get_(task, 'done'),
      get_(task, 'done_date'),
      get_(task, 'deal.name'),
      get_(task, 'deal.rating'),
      get_(usuario, 'name'),
      get_(usuario, 'email')
    ];
  });

  escreverAba_(ss, 'Tasks', headers, rows);
}

function escreverDeals_(ss, deals) {
  var baseHeaders = ['ID', 'Nome', 'Valor Total', 'Data de Criação', 'Última Atualização',
    'Organização', 'Endereço', 'Usuário Responsável', 'Email Usuário', 'Estágio', 'Fonte',
    'Campanha', 'Próxima Tarefa', 'Data Próxima Tarefa'];

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
      get_(deal, 'created_at'),
      get_(deal, 'updated_at'),
      get_(deal, 'organization.name'),
      get_(deal, 'organization.address'),
      get_(deal, 'user.name'),
      get_(deal, 'user.email'),
      get_(deal, 'deal_stage.name'),
      get_(deal, 'deal_source.name'),
      get_(deal, 'campaign.name'),
      get_(deal, 'next_task.subject'),
      get_(deal, 'next_task.date')
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

  escreverAba_(ss, 'Deals', headers, rows);
}

function escreverAba_(ss, nomeAba, headers, rows) {
  var sheet = ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
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
