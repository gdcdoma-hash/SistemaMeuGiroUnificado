var PORTAL_ATLETA_RETURN_TTL_MS_ = 5 * 60 * 1000;
var PORTAL_ATLETA_RETURN_URL_ = 'https://script.google.com/macros/s/AKfycbxqA6LmqyTca8i9af5EWKOzuaibTDQKFa6Mtsht4jm7tR29iVZeohNZYLdc3WjNFFJA5Q/exec';

function portalAtletaReturnDestino_(valor) {
  var destino = String(valor || '').trim().toLowerCase();
  return destino === 'inscricao' ? 'inscricao' : 'minhas';
}

function portalAtletaCriarHandoff(sessionToken, destinoSolicitado) {
  var sessao = atletaValidarSessao(sessionToken);
  if (!sessao || sessao.ok !== true || !sessao.id_dgmb) {
    return {
      ok: false,
      code: sessao && sessao.code ? sessao.code : 'SESSAO_INVALIDA',
      msg: sessao && sessao.msg ? sessao.msg : 'Sessão do atleta inválida.'
    };
  }

  var agora = Date.now();
  var expiraEm = agora + PORTAL_ATLETA_RETURN_TTL_MS_;
  var token = Utilities.getUuid() + '-' + Utilities.getUuid();
  var destino = portalAtletaReturnDestino_(destinoSolicitado);
  var sh = portalHandoffObterSheet_();

  sh.appendRow([
    token,
    sessao.id_dgmb,
    'PORTAL_ATLETA',
    new Date(agora),
    new Date(expiraEm),
    'ATIVO',
    'MEU_GIRO_ATLETA'
  ]);

  return {
    ok: true,
    expira_em: expiraEm,
    destino: destino,
    url: PORTAL_ATLETA_RETURN_URL_ + '?page=portal&handoff=' + encodeURIComponent(token) + '&destino=' + encodeURIComponent(destino)
  };
}
