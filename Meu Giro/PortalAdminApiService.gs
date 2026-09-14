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


function portalAdminApiDoPost_(e) {
  var corpo = {};
  try {
    corpo = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (erro) {
    return portalAdminApiJson_({ ok: false, code: 'JSON_INVALIDO', msg: 'Dados da requisição inválidos.' });
  }

  var acao = String(corpo.acao || '').trim().toLowerCase();
  var token = String(corpo.handoff || '').trim();

  try {
    if (acao === 'editar_atividade') {
      return portalAdminApiJson_(
        portalAdminMeuGiroEditarAtividade(token, corpo.payload || {})
      );
    }
    return portalAdminApiJson_({ ok: false, code: 'ACAO_INVALIDA', msg: 'Ação administrativa inválida.' });
  } catch (erro) {
    return portalAdminApiJson_({
      ok: false,
      code: 'ERRO_INTERNO',
      msg: erro && erro.message ? erro.message : 'Erro interno no Meu Giro.'
    });
  }
}
