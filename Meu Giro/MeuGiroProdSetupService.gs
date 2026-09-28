/**
 * Configuração inicial do Meu Giro PROD.
 * Travada pelo Script ID do projeto PROD para impedir execução no DEV.
 */
var DGMB_MEU_GIRO_PROD_SCRIPT_ID_ = '1gCtXj8xlNLhQi6XSJhQF7c8bxgvVvMuq3-ZmZ4_BIGoubMUs09D8VKY2';
var DGMB_MEU_GIRO_PROD_SPREADSHEET_ID_ = '1hQlopo_pbsp_KGEtXsDUZQ8ilPLjCwHhAyKivYlXRug';
var DGMB_MEU_GIRO_PROD_FILES_ROOT_ID_ = '1x6kIh-PzC7eGjJkS5gy0ggBw9KqRh4Y-';
var DGMB_MEU_GIRO_PROD_CERT_FOLDER_NAME_ = 'CERTIFICADOS_DGMB';
var DGMB_MEU_GIRO_PROD_CERT_TEMPLATE_ID_ = '13BP2rHBiqymQyOk1bJsNFsSPxJAhQEPHj5FRIuAssfo';
var DGMB_PORTAL_PROD_WEBAPP_URL_ = 'https://script.google.com/macros/s/AKfycbx3I_pc35_LWv1SpeEJbJ_3pq9aLE1lIc6tE4BccQGdD18nyKzmw5quivht3zOEdyCP/exec';
var DGMB_MEU_GIRO_PROD_WEBAPP_URL_ = 'https://script.google.com/macros/s/AKfycbzgCdo-kgL9Hf-f43wPIrt_hFYlamgZD5AaIEM9l3f2q8oM7cgOHUnMEDctvztqTC6s3w/exec';

function SETUP_MEU_GIRO_PROD_INICIAL() {
  var scriptAtual = ScriptApp.getScriptId();
  if (scriptAtual !== DGMB_MEU_GIRO_PROD_SCRIPT_ID_) {
    throw new Error(
      'BLOQUEADO: esta configuração só pode ser executada no projeto Meu Giro PROD.'
    );
  }

  var ss = SpreadsheetApp.openById(DGMB_MEU_GIRO_PROD_SPREADSHEET_ID_);
  if (!ss) throw new Error('Não foi possível abrir a planilha PROD.');

  var root = DriveApp.getFolderById(DGMB_MEU_GIRO_PROD_FILES_ROOT_ID_);
  var folders = root.getFoldersByName(DGMB_MEU_GIRO_PROD_CERT_FOLDER_NAME_);
  var certFolder = folders.hasNext()
    ? folders.next()
    : root.createFolder(DGMB_MEU_GIRO_PROD_CERT_FOLDER_NAME_);

  // Valida acesso ao template padrão sem modificá-lo.
  var template = DriveApp.getFileById(DGMB_MEU_GIRO_PROD_CERT_TEMPLATE_ID_);
  if (!template) throw new Error('Template padrão do certificado não acessível.');

  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    DGMB_ENVIRONMENT: 'PROD',
    DGMB_SPREADSHEET_ID: DGMB_MEU_GIRO_PROD_SPREADSHEET_ID_,
    DGMB_CERTIFICADOS_FOLDER_ID: certFolder.getId(),
    DGMB_CERTIFICADO_TEMPLATE_ID: DGMB_MEU_GIRO_PROD_CERT_TEMPLATE_ID_
  }, false);

  // Portal PROD ainda não possui URL /exec. Remove eventual valor anterior
  // para impedir retorno acidental ao Portal DEV.
  props.deleteProperty('DGMB_PORTAL_WEBAPP_URL');

  var retorno = {
    status: 'OK',
    ambiente: String(props.getProperty('DGMB_ENVIRONMENT') || ''),
    scriptId: scriptAtual,
    spreadsheetId: String(props.getProperty('DGMB_SPREADSHEET_ID') || ''),
    spreadsheetName: ss.getName(),
    certificadosFolderId: String(props.getProperty('DGMB_CERTIFICADOS_FOLDER_ID') || ''),
    certificadoTemplateId: String(props.getProperty('DGMB_CERTIFICADO_TEMPLATE_ID') || ''),
    portalConfigurado: !!String(props.getProperty('DGMB_PORTAL_WEBAPP_URL') || '').trim()
  };

  Logger.log('[MEU_GIRO][PROD_SETUP] ' + JSON.stringify(retorno));
  return retorno;
}

function VERIFICAR_MEU_GIRO_PROD_CONFIG() {
  var scriptAtual = ScriptApp.getScriptId();
  var props = PropertiesService.getScriptProperties();

  var spreadsheetAcessivel = false;
  var spreadsheetName = '';
  var certificadosFolderAcessivel = false;
  var templateAcessivel = false;

  try {
    var ss = SpreadsheetApp.openById(String(props.getProperty('DGMB_SPREADSHEET_ID') || ''));
    spreadsheetAcessivel = !!ss;
    spreadsheetName = ss ? ss.getName() : '';
  } catch (e) {}

  try {
    certificadosFolderAcessivel = !!DriveApp.getFolderById(
      String(props.getProperty('DGMB_CERTIFICADOS_FOLDER_ID') || '')
    );
  } catch (e2) {}

  try {
    templateAcessivel = !!DriveApp.getFileById(
      String(props.getProperty('DGMB_CERTIFICADO_TEMPLATE_ID') || '')
    );
  } catch (e3) {}

  var portalUrl = String(props.getProperty('DGMB_PORTAL_WEBAPP_URL') || '');

  var retorno = {
    scriptIdCorreto: scriptAtual === DGMB_MEU_GIRO_PROD_SCRIPT_ID_,
    scriptId: scriptAtual,
    ambiente: String(props.getProperty('DGMB_ENVIRONMENT') || ''),
    spreadsheetId: String(props.getProperty('DGMB_SPREADSHEET_ID') || ''),
    spreadsheetAcessivel: spreadsheetAcessivel,
    spreadsheetName: spreadsheetName,
    certificadosFolderId: String(props.getProperty('DGMB_CERTIFICADOS_FOLDER_ID') || ''),
    certificadosFolderAcessivel: certificadosFolderAcessivel,
    certificadoTemplateId: String(props.getProperty('DGMB_CERTIFICADO_TEMPLATE_ID') || ''),
    templateAcessivel: templateAcessivel,
    portalWebappUrl: portalUrl,
    portalUrlCorreta: portalUrl === DGMB_PORTAL_PROD_WEBAPP_URL_,
    prontoSemPortalUrl: (
      scriptAtual === DGMB_MEU_GIRO_PROD_SCRIPT_ID_ &&
      String(props.getProperty('DGMB_ENVIRONMENT') || '') === 'PROD' &&
      String(props.getProperty('DGMB_SPREADSHEET_ID') || '') === DGMB_MEU_GIRO_PROD_SPREADSHEET_ID_ &&
      spreadsheetAcessivel &&
      certificadosFolderAcessivel &&
      templateAcessivel
    ),
    prontoCompleto: (
      scriptAtual === DGMB_MEU_GIRO_PROD_SCRIPT_ID_ &&
      String(props.getProperty('DGMB_ENVIRONMENT') || '') === 'PROD' &&
      String(props.getProperty('DGMB_SPREADSHEET_ID') || '') === DGMB_MEU_GIRO_PROD_SPREADSHEET_ID_ &&
      spreadsheetAcessivel &&
      certificadosFolderAcessivel &&
      templateAcessivel &&
      portalUrl === DGMB_PORTAL_PROD_WEBAPP_URL_
    )
  };

  Logger.log('[MEU_GIRO][PROD_CONFIG] ' + JSON.stringify(retorno));
  return retorno;
}


function CONFIGURAR_URL_PORTAL_PROD() {
  var scriptAtual = ScriptApp.getScriptId();
  if (scriptAtual !== DGMB_MEU_GIRO_PROD_SCRIPT_ID_) {
    throw new Error(
      'BLOQUEADO: esta função só pode ser executada no projeto Meu Giro PROD.'
    );
  }

  var props = PropertiesService.getScriptProperties();
  if (String(props.getProperty('DGMB_ENVIRONMENT') || '').trim().toUpperCase() !== 'PROD') {
    throw new Error('Ambiente PROD ainda não configurado neste projeto.');
  }

  props.setProperty('DGMB_PORTAL_WEBAPP_URL', DGMB_PORTAL_PROD_WEBAPP_URL_);

  var retorno = {
    status: 'OK',
    meuGiroProdUrl: DGMB_MEU_GIRO_PROD_WEBAPP_URL_,
    portalProdUrl: String(props.getProperty('DGMB_PORTAL_WEBAPP_URL') || ''),
    configurado: String(props.getProperty('DGMB_PORTAL_WEBAPP_URL') || '') === DGMB_PORTAL_PROD_WEBAPP_URL_
  };

  Logger.log('[MEU_GIRO][PROD_URL_PORTAL] ' + JSON.stringify(retorno));
  return retorno;
}
