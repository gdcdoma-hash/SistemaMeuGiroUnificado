/**
 * PREPARACAO STAGING 15/10/2026 — executar manualmente APENAS no
 * projeto NOVO Meu Giro STAGING Integrado.
 * Localiza a copia restrita por nome; NAO publica IDs da planilha no GitHub.
 * Não grava km, atividades, pagamentos, inscrições ou certificados.
 */
function STAGING_CONFIGURAR_MEU_GIRO_20261015() {
  var scriptEsperado = '1Qkn_ko5Wv6fNwvVfC9lIQOMWZdncArqdWntHtspBEVyp8X1CC9gSsGs4';
  var planilhaNome = 'STAGING ISOLADO — Portal Giro + Meu Giro — 2026-10-15 — NAO PROD';
  var pastaNome = 'DGMB_STAGING_INTEGRADO_ARQUIVOS_2026-10-15';
  if (ScriptApp.getScriptId() !== scriptEsperado) {
    throw new Error('PROIBIDO: execute somente no NOVO projeto Meu Giro STAGING.');
  }
  var itens = DriveApp.getFilesByName(planilhaNome);
  if (!itens.hasNext()) throw new Error('Planilha STAGING nao acessivel para esta conta.');
  var arquivo = itens.next();
  if (itens.hasNext() || arquivo.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    throw new Error('Planilha STAGING ausente, duplicada ou do tipo incorreto.');
  }
  var id = arquivo.getId();
  if (SpreadsheetApp.openById(id).getName() !== planilhaNome) {
    throw new Error('A planilha localizada nao passou na validacao.');
  }
  var pastas = DriveApp.getFoldersByName(pastaNome);
  if (!pastas.hasNext()) throw new Error('Pasta STAGING nao acessivel para esta conta.');
  var pasta = pastas.next();
  if (pastas.hasNext()) throw new Error('Pasta STAGING duplicada. Interrompido.');
  var props = PropertiesService.getScriptProperties();
  var antigo = String(props.getProperty('DGMB_ENVIRONMENT') || '').trim();
  if (antigo && antigo !== 'STAGING') {
    throw new Error('Ambiente preexistente diferente de STAGING: nao modificar.');
  }
  props.setProperties({
    DGMB_ENVIRONMENT:'STAGING',
    DGMB_SPREADSHEET_ID:id,
    DGMB_CERTIFICADOS_FOLDER_ID:pasta.getId()
  }, false);
  var final = props.getProperties();
  if (final.DGMB_ENVIRONMENT !== 'STAGING' || final.DGMB_SPREADSHEET_ID !== id) {
    throw new Error('Falha na verificacao das propriedades.');
  }
  var retorno = {ok:true, ambiente:'STAGING', planilha:planilhaNome,
    pasta:pastaNome, deploymentCriado:false,
    urlPortalConfigurada:!!String(final.DGMB_PORTAL_WEBAPP_URL||'').trim()};
  Logger.log('[STAGING_MEU_GIRO] ' + JSON.stringify(retorno));
  return retorno;
}
