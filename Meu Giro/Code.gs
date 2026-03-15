function doGet(e) {
  const view = e && e.parameter ? e.parameter.view : '';

  if (view === 'inscricao') {
    return renderInscricaoDGMB();
  }

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('MEU GIRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Permite incluir arquivos HTML dentro do Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
