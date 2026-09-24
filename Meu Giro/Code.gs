function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var parametros = e && e.parameter ? e.parameter : {};
  template.meuGiroBootJson = JSON.stringify({
    embedded: String(parametros.embedded || '').trim() === '1',
    page: String(parametros.page || '').trim().toLowerCase(),
    handoff: String(parametros.handoff || '').trim()
  });

  return template
    .evaluate()
    .setTitle('MEU GIRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Permite incluir arquivos HTML dentro do Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
