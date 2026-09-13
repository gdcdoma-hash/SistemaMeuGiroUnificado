function doGet(e) {
  var page = e && e.parameter ? String(e.parameter.page || '').trim().toLowerCase() : '';
  var arquivo = page === 'admin' ? 'AdminMeuGiro' : 'Index';
  var titulo = page === 'admin' ? 'MEU GIRO — ADMIN' : 'MEU GIRO';

  return HtmlService
    .createTemplateFromFile(arquivo)
    .evaluate()
    .setTitle(titulo)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Permite incluir arquivos HTML dentro do Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
