function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var parametros = e && e.parameter ? e.parameter : {};
  template.meuGiroBootJson = JSON.stringify({
    embedded: String(parametros.embedded || '').trim() === '1'
  });

  return template
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
