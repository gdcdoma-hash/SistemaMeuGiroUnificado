const PERFORMANCE_DEBUG = false;

/**
 * Configuracoes que mudam entre DEV e a futura producao do Meu Giro.
 * Script Properties tem prioridade. Os valores abaixo preservam o DEV atual
 * enquanto a separacao de ambientes esta sendo preparada.
 */
const DGMB_MEU_GIRO_ENV_DEFAULTS_ = Object.freeze({
  DGMB_ENVIRONMENT: 'DEV',
  DGMB_SPREADSHEET_ID: '1sFxmiWmIrPlwXgSb56M8maHF5niRYVLWRMFSUQYXgHM',
  DGMB_PORTAL_WEBAPP_URL: 'https://script.google.com/macros/s/AKfycbxqA6LmqyTca8i9af5EWKOzuaibTDQKFa6Mtsht4jm7tR29iVZeohNZYLdc3WjNFFJA5Q/exec',
  DGMB_CERTIFICADOS_FOLDER_ID: '1GncBumQM3RAS6WIT0jHQPaIMKBlT7OHi',
  DGMB_CERTIFICADO_TEMPLATE_ID: '13BP2rHBiqymQyOk1bJsNFsSPxJAhQEPHj5FRIuAssfo'
});

function dgmbMeuGiroConfigGet_(chave) {
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

  if (Object.prototype.hasOwnProperty.call(DGMB_MEU_GIRO_ENV_DEFAULTS_, nome)) {
    return String(DGMB_MEU_GIRO_ENV_DEFAULTS_[nome] || '').trim();
  }

  throw new Error('Configuracao do Meu Giro nao encontrada: ' + nome);
}

function dgmbMeuGiroPortalWebappUrl_() {
  return dgmbMeuGiroConfigGet_('DGMB_PORTAL_WEBAPP_URL');
}

function dgmbMeuGiroCertificadosFolderId_() {
  return dgmbMeuGiroConfigGet_('DGMB_CERTIFICADOS_FOLDER_ID');
}

function dgmbMeuGiroCertificadoTemplateId_() {
  return dgmbMeuGiroConfigGet_('DGMB_CERTIFICADO_TEMPLATE_ID');
}

// Fonte oficial compartilhada com o Portal Giro.
const SPREADSHEET_ID = dgmbMeuGiroConfigGet_('DGMB_SPREADSHEET_ID');

const SHEETS = {
  PESSOAS: 'DadosPessoais',
  DESAFIO: 'dgmbDesafios',
  REGISTRO_KM: 'REGISTRO_KM',
  MEU_GIRO_RESUMO: 'MEU_GIRO_RESUMO',
  FRASES: 'FRASES',
  LISTA_DESAFIOS: 'ListaDesafios',
  CONFIG_CERTIFICADO_TEMPLATE: 'CONFIG_CERTIFICADO_TEMPLATE'
};
