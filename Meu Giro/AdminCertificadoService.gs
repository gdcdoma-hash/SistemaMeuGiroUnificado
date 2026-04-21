var ADMIN_VALIDACAO_CERTIFICADO_ID = '1133';

function isAdminValidacaoCertificado_(idDgmb) {
  return normalizeText_(idDgmb) === ADMIN_VALIDACAO_CERTIFICADO_ID;
}

function validarAcessoAdminCertificado_(adminIdDgmb) {
  if (isAdminValidacaoCertificado_(adminIdDgmb)) {
    return { ok: true };
  }

  return {
    ok: false,
    code: 'ACESSO_NEGADO_ADMIN_CERTIFICADO',
    msg: 'Acesso negado. Esta área é restrita ao administrador de validação.'
  };
}

function listarPendenciasValidacaoCertificado(adminIdDgmb) {
  try {
    var acesso = validarAcessoAdminCertificado_(adminIdDgmb);
    if (!acesso.ok) return acesso;

    var ss = getSpreadsheet_();
    var sheetName = SHEETS.DESAFIO || 'dgmbDesafios';
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      return { ok: false, code: 'ABA_DESAFIO_NAO_ENCONTRADA', msg: 'Aba dgmbDesafios não encontrada.' };
    }

    var values = sh.getDataRange().getValues();
    if (!values || values.length < 2) {
      return { ok: true, data: [] };
    }

    var map = buildHeaderMap_(values[0]);
    var idxIdDgmb = getRequiredColumnIndex_(map, ['id_dgmb'], sheetName);
    var idxIdDesafio = getOptionalColumnIndex_(map, ['id_desafio']);
    var idxIdItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
    var idxObsRegistro = getOptionalColumnIndex_(map, ['observacao', 'observação']);
    var idxMeta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km', 'meta_km', 'meta km']);
    var idxRealizado = getOptionalColumnIndex_(map, ['distancia_realizada', 'distancia realizada']);
    var idxStatusApuracao = getOptionalColumnIndex_(map, ['status_apuracao', 'status apuracao', 'status apuração', 'status_desafio', 'status desafio']);
    var idxStatusValidacao = getRequiredColumnIndex_(map, ['status_validacao_certificado'], sheetName);
    var idxPrint = getOptionalColumnIndex_(map, ['print_strava_certificado']);
    var idxLinkPrint = getRequiredColumnIndex_(map, ['link_print_strava'], sheetName);
    var idxDataEnvio = getOptionalColumnIndex_(map, ['data_envio_print_strava']);
    var idxDataAprovacao = getOptionalColumnIndex_(map, ['data_aprovacao_certificado']);
    var idxObsValidacao = getOptionalColumnIndex_(map, ['obs_validacao_certificado']);

    var nomesPorId = adminCertificadoBuildMapaNomesPessoas_();
    var nomesDesafiosPorId = adminCertificadoBuildMapaNomesDesafios_();
    var resumoPorChave = adminCertificadoBuildResumoPorChave_();

    var out = [];
    var statusPendentes = {
      PENDENTE: true,
      EM_ANALISE: true
    };

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var statusValidacao = normalizeText_(row[idxStatusValidacao]).toUpperCase();
      if (!statusPendentes[statusValidacao]) continue;

      var idDgmb = normalizeText_(row[idxIdDgmb]);
      if (!idDgmb) continue;

      var idDesafio = idxIdDesafio > -1 ? normalizeText_(row[idxIdDesafio]) : '';
      if (!idDesafio && idxObsRegistro > -1) {
        idDesafio = extrairIdDesafioObservacao_(row[idxObsRegistro]);
      }

      var idItem = idxIdItem > -1 ? normalizeText_(row[idxIdItem]) : '';
      var chaveExata = [idDgmb, idDesafio, idItem].join('|');
      var chaveSemItem = [idDgmb, idDesafio, ''].join('|');
      var resumo = resumoPorChave[chaveExata] || resumoPorChave[chaveSemItem] || null;

      var metaKm = resumo ? Number(resumo.meta_km || 0) : (idxMeta > -1 ? parseLocalizedNumber_(row[idxMeta]) : 0);
      var kmRealizado = resumo ? Number(resumo.distancia_realizada || 0) : (idxRealizado > -1 ? parseLocalizedNumber_(row[idxRealizado]) : 0);
      var metaValida = isFinite(metaKm) && metaKm > 0;
      var metaAtingida = metaValida && isFinite(kmRealizado) && kmRealizado >= metaKm;
      if (!metaAtingida) continue;

      out.push({
        row_number: i + 1,
        id_dgmb: idDgmb,
        nome_participante: nomesPorId[idDgmb] || '',
        id_desafio: idDesafio,
        id_item_estoque: idItem,
        nome_desafio: nomesDesafiosPorId[idDesafio] || '',
        meta_km: metaKm,
        km_realizado: kmRealizado,
        status_apuracao: resumo ? normalizeText_(resumo.status_apuracao).toUpperCase() : (idxStatusApuracao > -1 ? normalizeText_(row[idxStatusApuracao]).toUpperCase() : ''),
        status_validacao_certificado: statusValidacao,
        print_strava_certificado: idxPrint > -1 ? normalizeText_(row[idxPrint]) : '',
        link_print_strava: normalizeText_(row[idxLinkPrint]),
        data_envio_print_strava: idxDataEnvio > -1 ? row[idxDataEnvio] || '' : '',
        data_aprovacao_certificado: idxDataAprovacao > -1 ? row[idxDataAprovacao] || '' : '',
        obs_validacao_certificado: idxObsValidacao > -1 ? normalizeText_(row[idxObsValidacao]) : ''
      });
    }

    out.sort(function(a, b) {
      return String(a.status_validacao_certificado || '').localeCompare(String(b.status_validacao_certificado || '')) ||
        String(a.nome_participante || '').localeCompare(String(b.nome_participante || ''));
    });

    return { ok: true, data: out };
  } catch (err) {
    return {
      ok: false,
      code: 'ADMIN_CERTIFICADO_LISTAR_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao listar pendências de certificado.'
    };
  }
}

function atualizarStatusValidacaoCertificadoAdmin(payload) {
  try {
    var data = payload || {};
    var acesso = validarAcessoAdminCertificado_(data.admin_id_dgmb);
    if (!acesso.ok) return acesso;

    var idDgmb = normalizeText_(data.id_dgmb);
    var idDesafio = normalizeText_(data.id_desafio);
    var idItem = normalizeText_(data.id_item_estoque);
    var novoStatus = normalizeText_(data.novo_status).toUpperCase();
    var observacao = normalizeText_(data.observacao);

    if (!idDgmb || !idDesafio) {
      return { ok: false, code: 'PARAMETROS_INVALIDOS', msg: 'ID_DGMB e ID_DESAFIO são obrigatórios.' };
    }

    var statusPermitidos = {
      PENDENTE: true,
      EM_ANALISE: true,
      APROVADO: true,
      REPROVADO: true
    };

    if (!statusPermitidos[novoStatus]) {
      return { ok: false, code: 'STATUS_INVALIDO', msg: 'Status de validação inválido.' };
    }

    if (novoStatus === 'REPROVADO' && !observacao) {
      return { ok: false, code: 'OBS_OBRIGATORIA_REPROVACAO', msg: 'Informe uma observação ao reprovar o certificado.' };
    }

    var ss = getSpreadsheet_();
    var sheetName = SHEETS.DESAFIO || 'dgmbDesafios';
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      return { ok: false, code: 'ABA_DESAFIO_NAO_ENCONTRADA', msg: 'Aba dgmbDesafios não encontrada.' };
    }

    var values = sh.getDataRange().getValues();
    if (!values || values.length < 2) {
      return { ok: false, code: 'DESAFIOS_VAZIO', msg: 'Não há dados para atualização.' };
    }

    var map = buildHeaderMap_(values[0]);
    var idxIdDgmb = getRequiredColumnIndex_(map, ['id_dgmb'], sheetName);
    var idxIdDesafio = getOptionalColumnIndex_(map, ['id_desafio']);
    var idxIdItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
    var idxObsRegistro = getOptionalColumnIndex_(map, ['observacao', 'observação']);
    var idxStatusValidacao = getRequiredColumnIndex_(map, ['status_validacao_certificado'], sheetName);
    var idxDataAprovacao = getRequiredColumnIndex_(map, ['data_aprovacao_certificado'], sheetName);
    var idxObsValidacao = getRequiredColumnIndex_(map, ['obs_validacao_certificado'], sheetName);

    var linhaAtualizacao = -1;

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (normalizeText_(row[idxIdDgmb]) !== idDgmb) continue;

      var rowDesafio = idxIdDesafio > -1 ? normalizeText_(row[idxIdDesafio]) : '';
      if (!rowDesafio && idxObsRegistro > -1) {
        rowDesafio = extrairIdDesafioObservacao_(row[idxObsRegistro]);
      }
      if (rowDesafio !== idDesafio) continue;

      var rowItem = idxIdItem > -1 ? normalizeText_(row[idxIdItem]) : '';
      if (idItem && rowItem !== idItem) continue;

      linhaAtualizacao = i + 1;
      break;
    }

    if (linhaAtualizacao === -1) {
      return { ok: false, code: 'LINHA_NAO_ENCONTRADA', msg: 'Registro do desafio não encontrado para atualização.' };
    }

    sh.getRange(linhaAtualizacao, idxStatusValidacao + 1).setValue(novoStatus);

    if (novoStatus === 'APROVADO') {
      sh.getRange(linhaAtualizacao, idxDataAprovacao + 1).setValue(new Date());
    }

    if (idxObsValidacao > -1) {
      if (novoStatus === 'REPROVADO' || observacao) {
        sh.getRange(linhaAtualizacao, idxObsValidacao + 1).setValue(observacao);
      }
    }

    return {
      ok: true,
      data: {
        id_dgmb: idDgmb,
        id_desafio: idDesafio,
        id_item_estoque: idItem,
        status_validacao_certificado: novoStatus,
        observacao: observacao,
        row_number: linhaAtualizacao
      }
    };
  } catch (err) {
    return {
      ok: false,
      code: 'ADMIN_CERTIFICADO_ATUALIZAR_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao atualizar validação de certificado.'
    };
  }
}

function adminCertificadoBuildMapaNomesPessoas_() {
  var out = {};
  var pessoas = [];

  try {
    pessoas = getAllObjects_(SHEETS.PESSOAS) || [];
  } catch (e) {
    return out;
  }

  for (var i = 0; i < pessoas.length; i++) {
    var row = pessoas[i] || {};
    var id = normalizeText_(firstFilledValue_(row, ['ID_DGMB', 'id_dgmb']));
    if (!id) continue;
    out[id] = normalizeText_(firstFilledValue_(row, ['nome', 'Nome']));
  }

  return out;
}

function adminCertificadoBuildMapaNomesDesafios_() {
  var out = {};
  var itens = [];

  try {
    itens = getAllObjects_(SHEETS.LISTA_DESAFIOS) || [];
  } catch (e) {
    return out;
  }

  for (var i = 0; i < itens.length; i++) {
    var row = itens[i] || {};
    var id = normalizeText_(firstFilledValue_(row, ['ID_DESAFIO', 'id_desafio', 'id desafio']));
    if (!id) continue;
    out[id] = normalizeText_(firstFilledValue_(row, ['Nome_Desafio', 'nome_desafio', 'nome desafio', 'desafio', 'nome']));
  }

  return out;
}

function adminCertificadoBuildResumoPorChave_() {
  var out = {};
  var itens = [];

  try {
    itens = getAllObjects_(SHEETS.MEU_GIRO_RESUMO) || [];
  } catch (e) {
    return out;
  }

  for (var i = 0; i < itens.length; i++) {
    var row = itens[i] || {};
    var id = normalizeText_(firstFilledValue_(row, ['ID_DGMB', 'id_dgmb']));
    var idDesafio = normalizeText_(firstFilledValue_(row, ['ID_DESAFIO', 'id_desafio']));
    var idItem = normalizeText_(firstFilledValue_(row, ['id_item_estoque', 'ID_ITEM_ESTOQUE', 'id item estoque']));

    if (!id || !idDesafio) continue;

    out[[id, idDesafio, idItem].join('|')] = {
      meta_km: parseLocalizedNumber_(firstFilledValue_(row, ['Meta_KM', 'meta_km', 'meta km'])),
      distancia_realizada: parseLocalizedNumber_(firstFilledValue_(row, ['Distancia_Realizada', 'distancia_realizada', 'distancia realizada'])),
      status_apuracao: normalizeText_(firstFilledValue_(row, ['Status_Apuracao', 'status_apuracao', 'status apuração']))
    };
  }

  return out;
}
