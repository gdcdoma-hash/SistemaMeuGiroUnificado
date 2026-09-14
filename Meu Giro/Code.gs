function doGet(e) {
  var parametros = e && e.parameter ? e.parameter : {};

  if (String(parametros.api || '').trim().toLowerCase() === 'admin') {
    return portalAdminApiDoGet_(parametros);
  }

  var page = String(parametros.page || '').trim().toLowerCase();

  if (page === 'admin') {
    var token = String(parametros.handoff || '').trim();
    var acesso = portalHandoffValidarAdmin_(token);
    if (!acesso || acesso.ok !== true) {
      return HtmlService
        .createHtmlOutput('<!doctype html><html><body style="margin:0;background:#0f172a;color:#e5e7eb;font-family:Arial,sans-serif;padding:28px"><h2>Acesso administrativo indisponível</h2><p>Abra o Meu Giro — Admin a partir do Portal Administrativo.</p></body></html>')
        .setTitle('MEU GIRO — ADMIN')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var templateAdmin = HtmlService.createTemplateFromFile('AdminMeuGiro');
    templateAdmin.handoffToken = token;
    return templateAdmin
      .evaluate()
      .setTitle('MEU GIRO — ADMIN')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var template = HtmlService.createTemplateFromFile('Index');
  var atletaHandoffToken = page === 'atleta' ? String(parametros.handoff || '').trim() : '';
  template.atletaHandoffToken = atletaHandoffToken;
  template.portalGiroUrl = 'https://script.google.com/macros/s/AKfycbxqA6LmqyTca8i9af5EWKOzuaibTDQKFa6Mtsht4jm7tR29iVZeohNZYLdc3WjNFFJA5Q/exec';

  var avaliado = template.evaluate();
  var html = avaliado.getContent();
  var tokenJson = JSON.stringify(atletaHandoffToken || '');
  var integrado = page === 'atleta' && !!atletaHandoffToken;

  var guard = '<script>(function(){' +
    'function limparLegado(){try{["meuGiro.loginSession","meuGiro.painelState","meuGiro.desafioEmFocoKey"].forEach(function(k){localStorage.removeItem(k);});}catch(e){}}' +
    'if(typeof clearUserSession==="function"){var _clearUserSession=clearUserSession;clearUserSession=function(){try{_clearUserSession();}finally{limparLegado();}};}' +
    'if(typeof logoutUser==="function"){var _logoutUser=logoutUser;logoutUser=function(){limparLegado();return _logoutUser.apply(this,arguments);};}' +
    (integrado
      ? 'window.__MEU_GIRO_HANDOFF_INTEGRADO__=' + tokenJson + ';limparLegado();try{localStorage.removeItem("MEU_GIRO_CURRENT_USER");localStorage.removeItem("MEU_GIRO_SERVER_SESSION");}catch(e){}' +
        'if(typeof tentarRestaurarSessaoPersistida==="function"){tentarRestaurarSessaoPersistida=function(){return false;};}' +
        'if(typeof iniciarHandoffAtleta_==="function"){iniciarHandoffAtleta_(window.__MEU_GIRO_HANDOFF_INTEGRADO__);}'
      : '') +
    '})();<\/script>';

  if (/<\/body>/i.test(html)) {
    html = html.replace(/<\/body>/i, guard + '\n</body>');
  } else {
    html += guard;
  }

  return HtmlService
    .createHtmlOutput(html)
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

function doPost(e) {
  return portalAdminApiDoPost_(e);
}
