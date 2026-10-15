/**
 * Configuracoes que variam entre DEV e futura PROD do Meu Giro.
 *
 * Importante: este arquivo nao calcula nenhum valor na inicializacao global.
 * As propriedades sao lidas apenas quando uma funcao realmente precisa delas.
 */
function dgmbMeuGiroConfigGet_(chave, fallback) {
  var nome = String(chave || '').trim();
  if (!nome) throw new Error('Nome da configuracao do Meu Giro ausente.');

  // STAGING: aceitar apenas projeto novo e propriedades explicitas da base copiada.
  if (ScriptApp.getScriptId() !== '1Qkn_ko5Wv6fNwvVfC9lIQOMWZdncArqdWntHtspBEVyp8X1CC9gSsGs4') {
    throw new Error('[STAGING_MEU_GIRO_PROJETO_INCORRETO] Publicacao bloqueada neste projeto.');
  }
  var propsStaging = PropertiesService.getScriptProperties();
  if (String(propsStaging.getProperty('DGMB_ENVIRONMENT') || '').trim() !== 'STAGING') {
    throw new Error('[STAGING_MEU_GIRO_AMBIENTE_NAO_CONFIGURADO]');
  }
  var idStaging = String(propsStaging.getProperty('DGMB_SPREADSHEET_ID') || '').trim();
  if (!idStaging) throw new Error('[STAGING_MEU_GIRO_PLANILHA_NAO_CONFIGURADA]');
  if (SpreadsheetApp.openById(idStaging).getName() !== 'STAGING ISOLADO — Portal Giro + Meu Giro — 2026-10-15 — NAO PROD') {
    throw new Error('[STAGING_MEU_GIRO_PLANILHA_DIVERGENTE]');
  }
  if (nome === 'DGMB_ENVIRONMENT') return 'STAGING';
  // Somente neste projeto STAGING: URL do Portal implantado e restrito.
  // Recusar propriedades divergentes para nunca encaminhar o atleta ao DEV/PROD.
  if (nome === 'DGMB_PORTAL_WEBAPP_URL') {
    var destinoStaging = 'https://script.google.com/macros/s/AKfycbzkI5HpnfLpMjfokvZGj7lAj5sQGA_CYAz7UQhOp4UPcxlgX-tuAFUxBTIu1QF5dfjs/exec';
    var existente = String(propsStaging.getProperty(nome) || '').trim();
    if (existente && existente !== destinoStaging) {
      throw new Error('[STAGING_MEU_GIRO_URL_PORTAL_DIVERGENTE]');
    }
    return destinoStaging;
  }

  var configuradoStaging = String(propsStaging.getProperty(nome) || '').trim();
  if (!configuradoStaging) throw new Error('[STAGING_MEU_GIRO_CONFIG_OBRIGATORIA_AUSENTE] ' + nome);
  return configuradoStaging;


  var valor = '';
  try {
    valor = String(PropertiesService.getScriptProperties().getProperty(nome) || '').trim();
  } catch (erro) {
    Logger.log('[MEU_GIRO][ENV] Falha ao ler Script Properties para ' + nome + ': ' +
      (erro && erro.message ? erro.message : erro));
  }

  if (valor) return valor;

  var ambiente = '';
  try {
    ambiente = String(
      PropertiesService.getScriptProperties().getProperty('DGMB_ENVIRONMENT') || ''
    ).trim().toUpperCase();
  } catch (erroAmbiente) {}

  // Em PROD, nunca usar silenciosamente os valores padrão do DEV.
  if (ambiente === 'PROD' && nome !== 'DGMB_ENVIRONMENT') {
    throw new Error('Configuracao obrigatoria ausente no Meu Giro PROD: ' + nome);
  }

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
