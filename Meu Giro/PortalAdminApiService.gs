function portalAdminApiJson_(dados) {
  return ContentService
    .createTextOutput(JSON.stringify(dados || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function portalAdminApiDoGet_(parametros) {
  var p = parametros || {};
  var acao = String(p.acao || '').trim().toLowerCase();
  var token = String(p.handoff || '').trim();

  try {
    if (acao === 'buscar_atletas') {
      return portalAdminApiJson_(
        portalAdminMeuGiroBuscarAtletas(token, String(p.termo || ''))
      );
    }

    if (acao === 'consultar_atleta') {
      return portalAdminApiJson_(
        portalAdminMeuGiroConsultarAtleta(token, String(p.id_dgmb || ''))
      );
    }

    return portalAdminApiJson_({
      ok: false,
      code: 'ACAO_INVALIDA',
      msg: 'Ação administrativa inválida.'
    });
  } catch (erro) {
    return portalAdminApiJson_({
      ok: false,
      code: 'ERRO_INTERNO',
      msg: erro && erro.message ? erro.message : 'Erro interno no Meu Giro.'
    });
  }
}
