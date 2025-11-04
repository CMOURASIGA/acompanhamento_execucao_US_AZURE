/***********************************************************************************
 * CONTROLE DE USER STORIES - AZURE DEVOPS → GOOGLE SHEETS
 *
 * Objetivo:
 *   - Ler User Stories dos projetos do Azure DevOps
 *   - Calcular responsáveis e datas de cada etapa (Validado, Aceito, Resolvido, Concluído)
 *   - Gravar tudo em uma planilha única no Google Drive (uma aba por projeto)
 *
 * Pré-requisitos:
 *   - Script Property "AZURE_PAT" configurada com o Personal Access Token do Azure DevOps
 *   - Pasta no Google Drive para armazenar a planilha (FOLDER_ID)
 ***********************************************************************************/

/**
 * Configurações gerais
 */
const ORG = 'CNC-TI';
const AZURE_API_VERSION = '7.1';

const FOLDER_ID = '14T0UZyoqoBwdBk_fb8B8tA3-CSDCD9Gh'; // Pasta "Azure" no Drive
const SPREADSHEET_NAME = 'Controle_US_Azure';
const PAT_PROPERTY_KEY = 'AZURE_PAT';                   // Nome da Script Property com o PAT

// Mapeamento de estados para cada “etapa” da US
const VALIDATED_STATES = ['Ready'];
const ACCEPTED_STATES  = ['Active'];
const RESOLVED_STATES  = ['Resolved'];
const CLOSED_STATES    = ['Closed', 'Done'];


/***********************************************************************************
 * HELPERS GENÉRICOS
 ***********************************************************************************/

/**
 * Extrai um nome legível de um objeto IdentityRef do Azure DevOps
 * @param {Object|string} identity
 * @return {string}
 */
function extractIdentityName_(identity) {
  if (!identity) return '';

  // Se já veio como string, devolve direto
  if (typeof identity === 'string') return identity;

  // Padrão IdentityRef do Azure DevOps
  if (identity.displayName) return identity.displayName;
  if (identity.uniqueName) return identity.uniqueName;

  return '';
}

/**
 * Converte datas para o formato DD/MM/YYYY
 * Trata datas "9999-01-01" como vazias
 * @param {string|Date} value
 * @return {string}
 */
function formatDate_(value) {
  if (!value) return '';

  var d = new Date(value);
  if (isNaN(d)) return '';

  var year = d.getUTCFullYear();
  if (year >= 9000) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}

/**
 * Chamada genérica à API do Azure DevOps
 * @param {string} path   Caminho da API a partir de https://dev.azure.com/{ORG}
 * @param {string} method GET ou POST
 * @param {Object} body   Corpo JSON (opcional)
 * @return {Object}       JSON já parseado
 */
function callAzureDevOps_(path, method, body) {
  var token = PropertiesService.getScriptProperties().getProperty(PAT_PROPERTY_KEY);
  if (!token) {
    throw new Error('PAT não encontrado nas Script Properties (' + PAT_PROPERTY_KEY + ').');
  }

  var url = 'https://dev.azure.com/' + ORG + path;

  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(':' + token),
    'Content-Type': 'application/json'
  };

  var options = {
    method: method || 'get',
    headers: headers,
    muteHttpExceptions: true
  };

  if (body) {
    options.payload = JSON.stringify(body);
  }

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code >= 300) {
    throw new Error('Erro Azure DevOps ' + code + ': ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}


/***********************************************************************************
 * AZURE DEVOPS – CAMADA DE DADOS
 ***********************************************************************************/

/**
 * Lista os nomes dos projetos da organização
 * @return {string[]} Array de nomes de projetos
 */
function listProjects_() {
  var path = '/_apis/projects?api-version=' + AZURE_API_VERSION;
  var result = callAzureDevOps_(path, 'get');
  return result.value.map(function(p) { return p.name; });
}

/**
 * Busca os IDs das User Stories de um projeto via WIQL
 * @param {string} projectName
 * @return {number[]} IDs das US
 */
function fetchUserStoryIds_(projectName) {
  var wiql = {
    query: (
      "SELECT [System.Id] " +
      "FROM WorkItems " +
      "WHERE [System.TeamProject] = '" + projectName + "' " +
      "AND [System.WorkItemType] = 'User Story' " +
      "ORDER BY [System.ChangedDate] DESC"
    )
  };

  var path = '/' + encodeURIComponent(projectName) +
             '/_apis/wit/wiql?api-version=' + AZURE_API_VERSION;

  var result = callAzureDevOps_(path, 'post', wiql);
  return result.workItems.map(function(w) { return w.id; });
}

/**
 * Busca detalhes das US em lotes, com os campos necessários
 * @param {string} projectName
 * @param {number[]} ids
 * @return {Object[]} Work items com campos
 */
function fetchUserStoriesDetails_(projectName, ids) {
  if (!ids || ids.length === 0) return [];

  var allItems = [];
  var batchSize = 180; // Azure limita quantidade por chamada

  for (var i = 0; i < ids.length; i += batchSize) {
    var chunk = ids.slice(i, i + batchSize);

    var fields = [
      'System.Id',
      'System.Title',
      'System.State',
      'System.CreatedBy',
      'System.CreatedDate',
      'Microsoft.VSTS.Common.ResolvedBy',
      'Microsoft.VSTS.Common.ResolvedDate',
      'Microsoft.VSTS.Common.ClosedBy',
      'Microsoft.VSTS.Common.ClosedDate'
    ];

    var path = '/_apis/wit/workitems?ids=' + chunk.join(',') +
               '&fields=' + fields.join(',') +
               '&api-version=' + AZURE_API_VERSION;

    var result = callAzureDevOps_(path, 'get');
    allItems = allItems.concat(result.value);
  }

  return allItems;
}

/**
 * Retorna quem e quando colocou a US em um dos estados alvo
 * (lendo o histórico de updates da US)
 *
 * @param {string} projectName
 * @param {number} workItemId
 * @param {string[]} targetStates Lista de estados que representam a etapa (Ready, Active, etc.)
 * @return {{by: string, date: string}}
 */
function getFirstTransitionInfo_(projectName, workItemId, targetStates) {
  var path = '/' + encodeURIComponent(projectName) +
             '/_apis/wit/workitems/' + workItemId +
             '/updates?api-version=' + AZURE_API_VERSION;

  var result = callAzureDevOps_(path, 'get');
  var updates = result.value || [];

  // Garante ordem crescente por data
  updates.sort(function(a, b) {
    return new Date(a.revisedDate) - new Date(b.revisedDate);
  });

  for (var i = 0; i < updates.length; i++) {
    var upd = updates[i];
    var fields = upd.fields || {};
    var stateChange = fields['System.State'];

    if (stateChange && stateChange.newValue) {
      var newState = stateChange.newValue;
      if (targetStates.indexOf(newState) !== -1) {
        var by = '';
        if (upd.revisedBy) {
          by = extractIdentityName_(upd.revisedBy);
        }
        var date = upd.revisedDate || '';
        return { by: by, date: date };
      }
    }
  }

  // Nunca entrou nesse estado
  return { by: '', date: '' };
}


/***********************************************************************************
 * GOOGLE SHEETS – CRIAÇÃO E ESTRUTURA
 ***********************************************************************************/

/**
 * Cria (se necessário) ou abre a planilha Controle_US_Azure dentro da pasta Azure
 * @return {Spreadsheet}
 */
function getOrCreateSpreadsheet_() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var files = folder.getFilesByName(SPREADSHEET_NAME);
  
  if (files.hasNext()) {
    var file = files.next();
    return SpreadsheetApp.open(file);
  }

  // Cria novo arquivo na pasta
  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // tira da raiz "Meu Drive"

  return ss;
}

/**
 * Garante que a aba do projeto existe e escreve o cabeçalho
 * @param {Spreadsheet} ss
 * @param {string} projectName
 * @return {Sheet}
 */
function ensureProjectSheet_(ss, projectName) {
  var sheet = ss.getSheetByName(projectName);
  if (!sheet) {
    sheet = ss.insertSheet(projectName);
  }
  
  var headers = [
    'ID',
    'Link',
    'Status',
    'Criado por',
    'Data Criação',
    'Validado por',
    'Data Validação',
    'Aceito por',
    'Data Aceite',
    'Resolvido por',
    'Data Resolvido',
    'Concluído por',
    'Data Concluído'
  ];
  
  sheet.clear(); // zera conteúdo da aba
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  return sheet;
}


/***********************************************************************************
 * ORQUESTRAÇÃO – SINCRONIZAÇÃO
 ***********************************************************************************/

/**
 * Sincroniza um projeto específico para a aba correspondente
 * @param {string} projectName
 */
function syncProjectToSheet_(projectName) {
  Logger.log('Iniciando sincronização do projeto: ' + projectName);

  var ss = getOrCreateSpreadsheet_();
  Logger.log('Planilha aberta/criada: ' + SPREADSHEET_NAME);

  var sheet = ensureProjectSheet_(ss, projectName);
  Logger.log('Aba do projeto garantida: ' + projectName);

  Logger.log('Buscando IDs das User Stories...');
  var ids = fetchUserStoryIds_(projectName);
  Logger.log('Total de US encontradas: ' + ids.length);

  if (!ids || ids.length === 0) {
    Logger.log('Nenhuma US encontrada para o projeto: ' + projectName);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    return;
  }

  Logger.log('Buscando detalhes das US...');
  var items = fetchUserStoriesDetails_(projectName, ids);
  Logger.log('Detalhes obtidos para ' + items.length + ' US.');

  var rows = [];
  var counter = 0;

  items.forEach(function(item) {
    counter++;
    if (counter % 10 === 0) {
      Logger.log('Processando US ' + counter + ' de ' + items.length);
    }

    var f = item.fields || {};
    var id = f['System.Id'];
    var link = 'https://dev.azure.com/' + ORG + '/' +
               encodeURIComponent(projectName) + '/_workitems/edit/' + id;

    var status = f['System.State'] || '';

    // Criador e data de criação
    var criadoPor   = extractIdentityName_(f['System.CreatedBy']);
    var dataCriacao = formatDate_(f['System.CreatedDate']);

    // Histórico de transições
    var infoValidado   = getFirstTransitionInfo_(projectName, id, VALIDATED_STATES);
    var infoAceito     = getFirstTransitionInfo_(projectName, id, ACCEPTED_STATES);
    var infoResolvido  = getFirstTransitionInfo_(projectName, id, RESOLVED_STATES);
    var infoConcluido  = getFirstTransitionInfo_(projectName, id, CLOSED_STATES);

    var validadoPor    = infoValidado.by;
    var dataValidacao  = formatDate_(infoValidado.date);

    var aceitoPor      = infoAceito.by;
    var dataAceite     = formatDate_(infoAceito.date);

    // Resolvido por / data resolvido
    var resolvidoPor = infoResolvido.by ||
                      extractIdentityName_(f['Microsoft.VSTS.Common.ResolvedBy']);

    var dataResolvido = formatDate_(infoResolvido.date);
    if (!dataResolvido) {
      dataResolvido = formatDate_(f['Microsoft.VSTS.Common.ResolvedDate']);
    }

    // Concluído por / data concluído
    var concluidoPor = infoConcluido.by ||
                      extractIdentityName_(f['Microsoft.VSTS.Common.ClosedBy']);

    var dataConcluido = formatDate_(infoConcluido.date);
    if (!dataConcluido) {
      dataConcluido = formatDate_(f['Microsoft.VSTS.Common.ClosedDate']);
    }

    rows.push([
      id,
      link,
      status,
      criadoPor,
      dataCriacao,
      validadoPor,
      dataValidacao,
      aceitoPor,
      dataAceite,
      resolvidoPor,
      dataResolvido,
      concluidoPor,
      dataConcluido
    ]);
  });

  Logger.log('Escrevendo ' + rows.length + ' linhas na planilha...');
  sheet.getRange(2, 1, rows.length, 13).setValues(rows);

  Logger.log('Sincronização concluída para o projeto: ' + projectName);
}

/**
 * Sincroniza todos os projetos da organização
 */
function syncAllProjects() {
  var projectNames = listProjects_(); // Ex: SNCC, Arrecadação, SEI, etc.
  projectNames.forEach(function(name) {
    syncProjectToSheet_(name);
  });
}

/**
 * Sincroniza apenas um projeto específico – útil para testes
 */
function syncProjetoSNCC() {
  syncProjectToSheet_('Representações'); // troque o nome do projeto se quiser testar outro
}


/***********************************************************************************
 * FUNÇÕES DE TESTE / DIAGNÓSTICO
 ***********************************************************************************/

/**
 * Testa apenas se o PAT e a chamada à API de projetos estão corretos
 */
function testAzureToken() {
  var token = PropertiesService.getScriptProperties().getProperty(PAT_PROPERTY_KEY);
  var url = 'https://dev.azure.com/' + ORG + '/_apis/projects?api-version=' + AZURE_API_VERSION;

  var options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(':' + token)
    },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());
}

/**
 * Teste isolado de WIQL para um projeto específico
 * (ajuste o nome do projeto se quiser testar outro)
 */
function testListUserStories() {
  var projectName = 'SNCC';

  var wiqlQuery = {
    query: (
      "SELECT [System.Id], [System.Title], [System.State], [System.CreatedBy], [System.CreatedDate] " +
      "FROM WorkItems " +
      "WHERE [System.TeamProject] = '" + projectName + "' " +
      "AND [System.WorkItemType] = 'User Story' " +
      "ORDER BY [System.ChangedDate] DESC"
    )
  };

  var path = '/' + encodeURIComponent(projectName) +
             '/_apis/wit/wiql?api-version=' + AZURE_API_VERSION;

  var result = callAzureDevOps_(path, 'post', wiqlQuery);
  Logger.log(JSON.stringify(result, null, 2));
}

function debugConclusao_5523() {
  var projectName = 'Representações'; // projeto dessa US
  var workItemId = 5523;              // ID da US

  // 1) Ver campos principais da US
  var pathItem = '/' + encodeURIComponent(projectName) +
                 '/_apis/wit/workitems/' + workItemId +
                 '?fields=System.State,Microsoft.VSTS.Common.ClosedBy,Microsoft.VSTS.Common.ClosedDate&api-version=' + AZURE_API_VERSION;

  var item = callAzureDevOps_(pathItem, 'get');
  Logger.log('WORK ITEM:');
  Logger.log(JSON.stringify(item, null, 2));

  // 2) Ver histórico de updates (para entender a transição de estado)
  var pathUpdates = '/' + encodeURIComponent(projectName) +
                    '/_apis/wit/workitems/' + workItemId +
                    '/updates?api-version=' + AZURE_API_VERSION;

  var updates = callAzureDevOps_(pathUpdates, 'get');
  Logger.log('UPDATES (primeiros 5):');

  updates.value.slice(0, 5).forEach(function(u, idx) {
    Logger.log('Update ' + idx + ': ' + JSON.stringify(u.fields && u.fields['System.State']));
  });

  // 3) Ver o que nossa função está calculando hoje
  var infoConcluido = getFirstTransitionInfo_(projectName, workItemId, CLOSED_STATES);
  Logger.log('getFirstTransitionInfo_ (CLOSED): ' + JSON.stringify(infoConcluido));
}

