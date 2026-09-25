/**
 * Configuracoes que variam entre DEV e futura PROD do Meu Giro.
 *
 * Importante: este arquivo nao calcula nenhum valor na inicializacao global.
 * As propriedades sao lidas apenas quando uma funcao realmente precisa delas.
 */
function dgmbMeuGiroConfigGet_(chave, fallback) {
  var nome = String(chave || '').trim();
  if (!nome) throw new Error('Nome da configuracao do Meu Giro ausente.');

  var valor = '';
  try {
    valor = String(PropertiesService.getScriptProperties().getProperty(nome) || '').trim();
  } catch (erro) {
    Logger.log('[MEU_GIRO][ENV] Falha ao ler Script Properties para ' + nome + ': ' +
      (erro && erro.message ? erro.message : erro));
  }

  if (valor) return valor;

  var padrao = String(fallback == null ? '' : fallback).trim();
  if (padrao) return padrao;

  throw new Error('Configuracao do Meu Giro nao encontrada: ' + nome);
}

function dgmbMeuGiroSpreadsheetId_() {
  return dgmbMeuGiroConfigGet_(
    'DGMB_SPREADSHEET_ID',
    '1sFxmiWmIrPlwXgSb56M8maHF5niRYVLWRMFSUQYXgHM'
  );
}

function dgmbMeuGiroPortalWebappUrl_() {
  return dgmbMeuGiroConfigGet_(
    'DGMB_PORTAL_WEBAPP_URL',
    'https://script.google.com/macros/s/AKfycbxqA6LmqyTca8i9af5EWKOzuaibTDQKFa6Mtsht4jm7tR29iVZeohNZYLdc3WjNFFJA5Q/exec'
  );
}

function dgmbMeuGiroCertificadosFolderId_() {
  return dgmbMeuGiroConfigGet_(
    'DGMB_CERTIFICADOS_FOLDER_ID',
    '1GncBumQM3RAS6WIT0jHQPaIMKBlT7OHi'
  );
}

function dgmbMeuGiroCertificadoTemplateId_() {
  return dgmbMeuGiroConfigGet_(
    'DGMB_CERTIFICADO_TEMPLATE_ID',
    '13BP2rHBiqymQyOk1bJsNFsSPxJAhQEPHj5FRIuAssfo'
  );
}
