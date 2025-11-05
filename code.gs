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
// Link direto da planilha de controle
// Substitua pelo link real da sua planilha
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1G0slu7kbVqXucRrZTnEpcgsfGEjr4S1b6RHS7IGs9VM/edit';


/**
 * Configurações de e-mail
 * Observação: o Apps Script sempre envia a partir da conta que está executando o script.
 * Estas constantes controlam nome, destinatário e se o envio está habilitado.
 */
const EMAIL_NOTIFICATIONS_ENABLED = false; // coloque false se quiser desligar

const EMAIL_FROM_NAME = 'Christian Moura dos Santos';          // nome que aparece no "De"
const EMAIL_REPLY_TO  = 'christian_7c@cnc.org.br';             // para onde vão as respostas CNC - Sistemas de TI <sistemas@cnc.org.br>
const EMAIL_FROM_ADDRESS = 'contactconsultservices@gmail.com'; // conta técnica que executa o script (informativo)

const EMAIL_TO = 'christian_7c@cnc.org.br';                        // destinatário padrão

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
    return [];
  }

  Logger.log('Buscando detalhes das US...');
  var items = fetchUserStoriesDetails_(projectName, ids);
  Logger.log('Detalhes obtidos para ' + items.length + ' US.');

  var rows = [];
  var snapshot = []; // usado para comparação e envio de e-mail
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

    // snapshot com informações extras para o e-mail
    snapshot.push({
      projectName: projectName,
      id: id,
      status: status,
      link: link,
      dataCriacao: dataCriacao,
      dataValidacao: dataValidacao,
      dataAceite: dataAceite,
      dataResolvido: dataResolvido,
      dataConcluido: dataConcluido
    });
  }); // <- fecha o forEach aqui ✅

  Logger.log('Escrevendo ' + rows.length + ' linhas na planilha...');
  sheet.getRange(2, 1, rows.length, 13).setValues(rows);

  Logger.log('Sincronização concluída para o projeto: ' + projectName);

  // devolve snapshot para quem chamou (syncAllProjects)
  return snapshot;
}



/**
 * Sincroniza todos os projetos da organização
 */
function syncAllProjects() {
  var projectNames = listProjects_(); // Ex: SNCC, Arrecadação, SEI, etc.
  var globalSnapshot = [];

  projectNames.forEach(function(name) {
    var projectSnapshot = syncProjectToSheet_(name);
    globalSnapshot = globalSnapshot.concat(projectSnapshot);
  });

  if (EMAIL_NOTIFICATIONS_ENABLED) {
    notifyUserStoryChanges_(globalSnapshot);
  } else {
    Logger.log('Envio de e-mails desabilitado (EMAIL_NOTIFICATIONS_ENABLED = false).');
  }
}


/**
 * Sincroniza apenas um projeto específico – útil para testes
 */
function syncProjetoSNCC() {
  syncProjectToSheet_('SEI'); // troque o nome do projeto se quiser testar outro
}

/***********************************************************************************
 * NOTIFICAÇÃO POR E-MAIL – MUDANÇA DE STATUS DE USER STORIES
 ***********************************************************************************/

/**
 * Monta o HTML do e-mail com a lista de mudanças
 * @param {Object[]} changedItems
 * @return {string} html
 */
function buildEmailBody_(changedItems) {
  var dataExecucao = formatDate_(new Date());

  // Cores da identidade visual CNC
  var cncBlue = '#004d73';      // Azul CNC (principal)
  var cncGold = '#c9a055';      // Dourado CNC (destaque)
  var lightBlue = '#e8f4f8';    // Azul claro para alternância
  var borderColor = '#d0d0d0';
  var textColor = '#333333';

  var html = '';

  // Container principal com borda sutil
  html += '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: ' + textColor + '; max-width: 900px; margin: 0 auto; background-color: #ffffff;">';

  // Header com título (sem logo)
  html += '<div style="background-color: ' + cncBlue + '; padding: 25px 20px; text-align: center; border-radius: 8px 8px 0 0;">';
  html += '<h1 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 600; letter-spacing: 0.5px;">CNC</h1>';
  html += '<h2 style="color: ' + cncGold + '; margin: 8px 0 0 0; font-size: 18px; font-weight: 500;">Atualização de User Stories - Azure DevOps</h2>';
  html += '</div>';

  // Conteúdo principal
  html += '<div style="padding: 30px 20px; background-color: #ffffff;">';

  html += '<p style="line-height: 1.6; margin-bottom: 15px;">Prezada equipe de Sistemas de TI,</p>';

  html += '<p style="line-height: 1.6; margin-bottom: 15px;">Seguem abaixo as User Stories do Azure DevOps que tiveram alteração de status após o último processamento.</p>';

  // Badge com data de execução
  html += '<div style="background-color: ' + lightBlue + '; border-left: 4px solid ' + cncGold + '; padding: 12px 15px; margin: 20px 0; border-radius: 4px;">';
  html += '<p style="margin: 0; font-weight: 600;"><span style="color: ' + cncBlue + ';">📅 Data da execução:</span> ' + dataExecucao + '</p>';
  html += '</div>';

  // Link para planilha em destaque
  html += '<div style="background-color: #f9f9f9; padding: 15px; border-radius: 6px; margin: 20px 0; border: 1px solid ' + borderColor + ';">';
  html += '<p style="margin: 0 0 8px 0; font-weight: 600; color: ' + cncBlue + ';">📊 Planilha de Controle</p>';
  html += '<p style="margin: 0;">Consulte o detalhamento completo em: ';
  html += '<a href="' + SPREADSHEET_URL + '" target="_blank" style="color: ' + cncGold + '; text-decoration: none; font-weight: 600;">Controle_US_Azure →</a></p>';
  html += '</div>';

  // Título da tabela
  html += '<h3 style="color: ' + cncBlue + '; font-size: 16px; margin: 25px 0 15px 0; border-bottom: 2px solid ' + cncGold + '; padding-bottom: 8px;">Alterações Identificadas</h3>';

  // Tabela responsiva com melhor compatibilidade
  html += '<div style="overflow-x: auto;">';
  html += '<table border="0" cellpadding="12" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 13px; width: 100%; border: 1px solid ' + borderColor + ';">';

  // Cabeçalho da tabela (sem gradiente para melhor compatibilidade)
  html += '<thead>';
  html += '<tr style="background-color: ' + cncBlue + '; color: #ffffff;">';
  html += '<th align="left" style="padding: 14px 12px; border-bottom: 3px solid ' + cncGold + '; font-weight: 600; white-space: nowrap;">Projeto</th>';
  html += '<th align="center" style="padding: 14px 12px; border-bottom: 3px solid ' + cncGold + '; font-weight: 600; white-space: nowrap;">ID</th>';
  html += '<th align="left" style="padding: 14px 12px; border-bottom: 3px solid ' + cncGold + '; font-weight: 600; white-space: nowrap;">Status Anterior</th>';
  html += '<th align="left" style="padding: 14px 12px; border-bottom: 3px solid ' + cncGold + '; font-weight: 600; white-space: nowrap;">Status Atual</th>';
  html += '<th align="center" style="padding: 14px 12px; border-bottom: 3px solid ' + cncGold + '; font-weight: 600; white-space: nowrap;">Data da Ação</th>';
  html += '<th align="center" style="padding: 14px 12px; border-bottom: 3px solid ' + cncGold + '; font-weight: 600; white-space: nowrap;">Ação</th>';
  html += '</tr>';
  html += '</thead>';

  html += '<tbody>';
  changedItems.forEach(function(item, index) {
    var rowBg = (index % 2 === 0) ? '#ffffff' : lightBlue;

    html += '<tr style="background-color: ' + rowBg + ';">';
    html += '<td style="padding: 12px; border-bottom: 1px solid ' + borderColor + ';">' + item.projectName + '</td>';
    html += '<td align="center" style="padding: 12px; border-bottom: 1px solid ' + borderColor + '; font-family: monospace; font-weight: 600; color: ' + cncBlue + ';">' + item.id + '</td>';
    html += '<td style="padding: 12px; border-bottom: 1px solid ' + borderColor + '; color: #666;">' + (item.oldStatus || '—') + '</td>';
    html += '<td style="padding: 12px; border-bottom: 1px solid ' + borderColor + '; font-weight: 600; color: ' + cncBlue + ';">' + item.newStatus + '</td>';
    html += '<td align="center" style="padding: 12px; border-bottom: 1px solid ' + borderColor + '; font-size: 12px;">' + (item.actionDate || '—') + '</td>';
    html += '<td align="center" style="padding: 12px; border-bottom: 1px solid ' + borderColor + ';">';
    html += '<a href="' + item.link + '" target="_blank" style="background-color: ' + cncGold + '; color: #ffffff; padding: 6px 16px; text-decoration: none; border-radius: 4px; font-weight: 600; font-size: 12px; display: inline-block;">Abrir</a>';
    html += '</td>';
    html += '</tr>';
  });
  html += '</tbody>';

  html += '</table>';
  html += '</div>';

  // Nota informativa
  html += '<div style="margin-top: 25px; padding: 15px; background-color: #f0f0f0; border-radius: 6px; font-size: 12px; color: #666;">';
  html += '<p style="margin: 0;"><strong>ℹ️ Informação:</strong> Este e-mail foi enviado automaticamente pela rotina de integração entre Azure DevOps e Google Sheets.</p>';
  html += '</div>';

  html += '</div>'; // Fim do conteúdo principal

  // Assinatura (alinhada à direita)
  html += '<div style="margin-top: 30px; text-align: left; padding-right: 20px;">';
  html += '<p style="color: #333; margin: 0 0 5px 0; font-size: 14px;">Atenciosamente,</p>';
  html += '<p style="color: ' + cncBlue + '; margin: 0; font-weight: 600; font-size: 15px;">' + EMAIL_FROM_NAME + '</p>';
  html += '</div>';

  html += '</div>'; // Fim do conteúdo principal

  // Footer institucional
  html += '<div style="background-color: ' + cncBlue + '; padding: 15px 20px; text-align: center; border-radius: 0 0 8px 8px;">';
  html += '<p style="color: rgba(255,255,255,0.8); margin: 0; font-size: 11px;">Confederação Nacional do Comércio de Bens, Serviços e Turismo</p>';
  html += '</div>';

  html += '</div>'; // Fim do container principal

  return html;
}


/**
 * Decide qual data usar como "data da ação" de acordo com o status atual
 * Closed  -> Data Concluído
 * Resolved -> Data Resolvido
 * Active  -> Data Aceite
 * Ready   -> Data Validação
 * New / outros -> Data Criação
 */
function getActionDateForStatus_(item) {
  var s = item.status || '';

  if (CLOSED_STATES.indexOf(s) !== -1) {
    return item.dataConcluido || item.dataResolvido || item.dataAceite || item.dataValidacao || item.dataCriacao || '';
  }

  if (RESOLVED_STATES.indexOf(s) !== -1) {
    return item.dataResolvido || item.dataConcluido || item.dataAceite || item.dataValidacao || item.dataCriacao || '';
  }

  if (ACCEPTED_STATES.indexOf(s) !== -1) {
    return item.dataAceite || item.dataValidacao || item.dataCriacao || '';
  }

  if (VALIDATED_STATES.indexOf(s) !== -1) {
    return item.dataValidacao || item.dataCriacao || '';
  }

  // fallback para itens novos ou estados não mapeados
  return item.dataCriacao || '';
}



/**
 * Compara o snapshot atual com o LOG_EXECUCAO e, se houver mudanças, envia e-mail
 * @param {Object[]} currentItems [{projectName, id, status, link}, ...]
 */
function notifyUserStoryChanges_(currentItems) {
  var ss = getOrCreateSpreadsheet_();
  var logSheetName = 'LOG_EXECUCAO';
  var logSheet = ss.getSheetByName(logSheetName);
  var firstRun = false;
  var previousMap = {}; // chave: "Projeto|ID" -> status anterior

  // Carrega LOG_EXECUCAO anterior
  if (logSheet) {
    var data = logSheet.getDataRange().getValues();
    if (data.length > 1) {
      data.slice(1).forEach(function(row) {
        var project = row[0];
        var id = row[1];
        var status = row[2];
        var key = project + '|' + id;
        previousMap[key] = status;
      });
    } else {
      firstRun = true;
    }
  } else {
    logSheet = ss.insertSheet(logSheetName);
    firstRun = true;
  }

  // Detecta mudanças
  var changedItems = [];

  currentItems.forEach(function(item) {
    var key = item.projectName + '|' + item.id;
    var prevStatus = previousMap[key];

    var actionDate = getActionDateForStatus_(item);

    if (!prevStatus) {
      // US nova no snapshot (não existia antes)
      changedItems.push({
        projectName: item.projectName,
        id: item.id,
        link: item.link,
        oldStatus: 'N/A',
        newStatus: item.status,
        actionDate: actionDate
      });
    } else if (prevStatus !== item.status) {
      // US com mudança de status
      changedItems.push({
        projectName: item.projectName,
        id: item.id,
        link: item.link,
        oldStatus: prevStatus,
        newStatus: item.status,
        actionDate: actionDate
      });
    }

  });

  // Atualiza LOG_EXECUCAO com o snapshot atual (sempre)
  logSheet.clear();
  var header = ['Projeto', 'ID', 'Status', 'Link'];
  logSheet.getRange(1, 1, 1, header.length).setValues([header]);

  var logRows = currentItems.map(function(item) {
    return [item.projectName, item.id, item.status, item.link];
  });

  if (logRows.length > 0) {
    logSheet.getRange(2, 1, logRows.length, header.length).setValues(logRows);
  }

  // Primeira execução: só monta o LOG, mas não manda e-mail
  if (firstRun) {
    Logger.log('Primeira execução: LOG_EXECUCAO inicializado, sem envio de e-mail.');
    return;
  }

  // Se não houve mudanças, não manda e-mail
  if (changedItems.length === 0) {
    Logger.log('Nenhuma alteração de status encontrada, não enviando e-mail.');
    return;
  }

  // Monta corpo do e-mail
  var htmlBody = buildEmailBody_(changedItems);

  // Envia e-mail
  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: 'Atualização de User Stories – Azure DevOps',
    name: EMAIL_FROM_NAME,
    replyTo: EMAIL_REPLY_TO,
    htmlBody: htmlBody
  });

  Logger.log('E-mail enviado para ' + EMAIL_TO + ' com ' + changedItems.length + ' alterações.');
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

function testNotifySingleProject() {
  // roda só o projeto SEI (pode trocar o nome se quiser)
  var snapshot = syncProjectToSheet_('SEI'); 

  // usa o snapshot real para comparar com o LOG e, se tiver mudança, mandar e-mail
  notifyUserStoryChanges_(snapshot);
}

