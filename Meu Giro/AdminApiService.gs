var MEU_GIRO_ADMIN_AUDITORIA_SHEET_ = '_MEU_GIRO_ADMIN_AUDITORIA';
var MEU_GIRO_ADMIN_AUDITORIA_HEADERS_ = [
  'ID_AUDITORIA',
  'DATA_HORA',
  'ID_DGMB',
  'ACTIVITY_ID',
  'CHAVE_EDICAO',
  'DATA_ANTERIOR',
  'KM_ANTERIOR',
  'DATA_NOVA',
  'KM_NOVO',
  'MOTIVO',
  'ADMIN_ID_DGMB'
];

function meuGiroAdminApiJson_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function meuGiroAdminApiFalha_(code, msg) {
  return {
    ok: false,
    code: String(code || 'ADMIN_API_ERROR'),
    msg: String(msg || 'Erro na API administrativa do Meu Giro.')
  };
}

function meuGiroAdminApiValidarAcesso_(handoffToken) {
  var acesso = portalHandoffValidarAdmin_(handoffToken);
  if (!acesso || acesso.ok !== true) {
    return {
      ok: false,
      resposta: meuGiroAdminApiFalha_(
        acesso && acesso.code ? acesso.code : 'ADMIN_HANDOFF_INVALIDO',
        acesso && acesso.msg ? acesso.msg : 'Acesso administrativo inválido.'
      )
    };
  }
  return {
    ok: true,
    admin_id_dgmb: normalizeText_(acesso.id_dgmb)
  };
}

function meuGiroAdminApiBuscarAtletas_(termo) {
  var alvo = normalizeText_(termo).toLowerCase();
  if (!alvo) {
    return { ok: true, data: [] };
  }

  var somenteDigitos = onlyDigits_(alvo);
  var cpfBase64 = somenteDigitos.length === 11 ? Utilities.base64Encode(somenteDigitos) : '';
  var sh = getSheetByName_(SHEETS.PESSOAS);
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { ok: true, data: [] };
  }

  var map = buildHeaderMap_(values[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], SHEETS.PESSOAS);
  var idxNome = getRequiredColumnIndex_(map, ['nome'], SHEETS.PESSOAS);
  var idxCidade = getOptionalColumnIndex_(map, ['cidade-uf', 'cidade_uf', 'cidade uf']);
  var idxCpf = getOptionalColumnIndex_(map, ['cpf']);
  var idxCod = getOptionalColumnIndex_(map, ['cod', 'cód', 'codigo', 'código']);

  var encontrados = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    var id = normalizeText_(row[idxId]);
    if (!id) continue;

    var nome = normalizeText_(row[idxNome]);
    var cidade = idxCidade > -1 ? normalizeText_(row[idxCidade]) : '';
    var cod = idxCod > -1 ? normalizeText_(row[idxCod]) : '';
    var cpfSalvo = idxCpf > -1 ? normalizeText_(row[idxCpf]) : '';

    var corresponde =
      id.toLowerCase().indexOf(alvo) !== -1 ||
      nome.toLowerCase().indexOf(alvo) !== -1 ||
      (cod && cod.toLowerCase().indexOf(alvo) !== -1) ||
      (cpfBase64 && cpfSalvo === cpfBase64);

    if (!corresponde) continue;

    encontrados.push({
      id_dgmb: id,
      nome: nome,
      cidade_uf: cidade,
      cod: cod
    });

    if (encontrados.length >= 30) break;
  }

  return { ok: true, data: encontrados };
}

function meuGiroAdminAuditoriaSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(MEU_GIRO_ADMIN_AUDITORIA_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(MEU_GIRO_ADMIN_AUDITORIA_SHEET_);
    sh.getRange(1, 1, 1, MEU_GIRO_ADMIN_AUDITORIA_HEADERS_.length)
      .setValues([MEU_GIRO_ADMIN_AUDITORIA_HEADERS_]);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function meuGiroAdminApiFormatarDataHora_(valor) {
  if (!(valor instanceof Date) || isNaN(valor.getTime())) return normalizeText_(valor);
  var tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  return Utilities.formatDate(valor, tz, 'dd/MM/yyyy HH:mm:ss');
}

function meuGiroAdminApiListarAuditorias_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var sh = meuGiroAdminAuditoriaSheet_();
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  var out = [];
  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i] || [];
    if (normalizeText_(row[2]) !== id) continue;

    out.push({
      id_auditoria: normalizeText_(row[0]),
      data_hora: meuGiroAdminApiFormatarDataHora_(row[1]),
      id_dgmb: normalizeText_(row[2]),
      activity_id: normalizeText_(row[3]),
      chave_edicao: normalizeText_(row[4]),
      data_anterior: normalizarDataISO_(row[5]) || normalizeText_(row[5]),
      km_anterior: parseLocalizedNumber_(row[6]),
      data_nova: normalizarDataISO_(row[7]) || normalizeText_(row[7]),
      km_novo: parseLocalizedNumber_(row[8]),
      motivo: normalizeText_(row[9]),
      admin_id_dgmb: normalizeText_(row[10])
    });

    if (out.length >= 50) break;
  }

  return out;
}

function meuGiroAdminApiEncontrarAtividade_(idDgmb, activityId, chaveEdicao) {
  var atividades = buscarAtividadesUsuario_(idDgmb) || [];
  var activity = normalizeText_(activityId);
  var chave = normalizeText_(chaveEdicao);

  for (var i = 0; i < atividades.length; i++) {
    var item = atividades[i] || [];
    if (activity && normalizeText_(item.activity_id) === activity) return item;
    if (!activity && chave && normalizeText_(item.chave_edicao) === chave) return item;
  }
  return null;
}

function meuGiroAdminApiRegistrarAuditoria_(dados) {
  dados = dados || {};
  var idAuditoria = 'MGA-' + new Date().getTime() + '-' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();

  var sh = meuGiroAdminAuditoriaSheet_();
  sh.appendRow([
    idAuditoria,
    new Date(),
    normalizeText_(dados.id_dgmb),
    normalizeText_(dados.activity_id),
    normalizeText_(dados.chave_edicao),
    normalizeText_(dados.data_anterior),
    Number(dados.km_anterior || 0),
    normalizeText_(dados.data_nova),
    Number(dados.km_novo || 0),
    normalizeText_(dados.motivo),
    normalizeText_(dados.admin_id_dgmb)
  ]);

  return idAuditoria;
}

function meuGiroAdminApiMontarPainel_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return meuGiroAdminApiFalha_('ID_OBRIGATORIO', 'ID do atleta não informado.');

  var painel = getPainelUsuario(id, { somenteLeitura: true });
  if (!painel || painel.ok !== true || !painel.data) {
    return meuGiroAdminApiFalha_(
      painel && painel.code ? painel.code : 'PAINEL_INDISPONIVEL',
      painel && painel.msg ? painel.msg : 'Não foi possível carregar os dados do atleta.'
    );
  }

  var data = painel.data || {};
  return {
    ok: true,
    id_dgmb: id,
    atleta: {
      id_dgmb: id,
      nome: normalizeText_(data.nome),
      cidade_uf: normalizeText_(data.cidade_uf)
    },
    total_pedalado: Number(data.total_pedalado != null ? data.total_pedalado : (data.totalPedalado || 0)),
    desafios_ativos: Array.isArray(data.desafios_ativos) ? data.desafios_ativos : [],
    desafios_historico: Array.isArray(data.desafios_historico) ? data.desafios_historico : [],
    atividades: Array.isArray(data.atividades) ? data.atividades : [],
    auditorias: meuGiroAdminApiListarAuditorias_(id)
  };
}

function meuGiroAdminApiDoGet_(parametros) {
  try {
    parametros = parametros || {};
    var acesso = meuGiroAdminApiValidarAcesso_(parametros.handoff);
    if (!acesso.ok) return meuGiroAdminApiJson_(acesso.resposta);

    var acao = normalizeText_(parametros.acao).toLowerCase();
    var resposta;

    if (acao === 'buscar_atletas') {
      resposta = meuGiroAdminApiBuscarAtletas_(parametros.termo);
    } else if (acao === 'consultar_atleta') {
      resposta = meuGiroAdminApiMontarPainel_(parametros.id_dgmb);
    } else {
      resposta = meuGiroAdminApiFalha_('ACAO_INVALIDA', 'Ação administrativa não reconhecida.');
    }

    return meuGiroAdminApiJson_(resposta);
  } catch (err) {
    Logger.log('[MEU_GIRO_ADMIN_API][GET] ' + (err && err.stack ? err.stack : err));
    return meuGiroAdminApiJson_(meuGiroAdminApiFalha_(
      'ADMIN_API_GET_EXCEPTION',
      err && err.message ? err.message : 'Erro interno na consulta administrativa.'
    ));
  }
}

function meuGiroAdminApiDoPost_(e) {
  try {
    var texto = e && e.postData ? String(e.postData.contents || '') : '';
    if (!texto) {
      return meuGiroAdminApiJson_(meuGiroAdminApiFalha_('PAYLOAD_AUSENTE', 'Dados da operação não informados.'));
    }

    var body = JSON.parse(texto);
    var acesso = meuGiroAdminApiValidarAcesso_(body.handoff);
    if (!acesso.ok) return meuGiroAdminApiJson_(acesso.resposta);

    var acao = normalizeText_(body.acao).toLowerCase();
    if (acao !== 'editar_atividade') {
      return meuGiroAdminApiJson_(meuGiroAdminApiFalha_('ACAO_INVALIDA', 'Ação administrativa não reconhecida.'));
    }

    var payload = body.payload || {};
    var idDgmb = normalizeText_(payload.id_dgmb);
    var motivo = normalizeText_(payload.motivo);
    if (motivo.length < 5) {
      return meuGiroAdminApiJson_(meuGiroAdminApiFalha_(
        'MOTIVO_OBRIGATORIO',
        'Informe o motivo da correção com pelo menos 5 caracteres.'
      ));
    }

    var atividadeAnterior = meuGiroAdminApiEncontrarAtividade_(
      idDgmb,
      payload.activity_id,
      payload.chave_edicao
    );
    if (!atividadeAnterior) {
      return meuGiroAdminApiJson_(meuGiroAdminApiFalha_(
        'ATIVIDADE_NAO_ENCONTRADA',
        'Atividade não encontrada antes da correção.'
      ));
    }

    var resultadoEdicao = editarAtividade(payload);
    if (!resultadoEdicao || resultadoEdicao.ok !== true) {
      return meuGiroAdminApiJson_(resultadoEdicao || meuGiroAdminApiFalha_(
        'EDICAO_FALHOU',
        'Não foi possível corrigir a atividade.'
      ));
    }

    var auditoriaId = meuGiroAdminApiRegistrarAuditoria_({
      id_dgmb: idDgmb,
      activity_id: atividadeAnterior.activity_id || payload.activity_id,
      chave_edicao: atividadeAnterior.chave_edicao || payload.chave_edicao,
      data_anterior: atividadeAnterior.data,
      km_anterior: atividadeAnterior.km,
      data_nova: payload.data_atividade,
      km_novo: parseLocalizedNumber_(payload.km),
      motivo: motivo,
      admin_id_dgmb: acesso.admin_id_dgmb
    });

    var painelAtualizado = meuGiroAdminApiMontarPainel_(idDgmb);
    painelAtualizado.auditoria_id = auditoriaId;
    painelAtualizado.msg = 'Atividade corrigida e auditoria registrada.';
    return meuGiroAdminApiJson_(painelAtualizado);
  } catch (err) {
    Logger.log('[MEU_GIRO_ADMIN_API][POST] ' + (err && err.stack ? err.stack : err));
    return meuGiroAdminApiJson_(meuGiroAdminApiFalha_(
      'ADMIN_API_POST_EXCEPTION',
      err && err.message ? err.message : 'Erro interno na operação administrativa.'
    ));
  }
}
