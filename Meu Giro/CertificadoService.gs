function gerarOuObterCertificadoDesafio(payload) {
  try {
    var contexto = certificadoBuscarContextoDesafio_(payload || {});

    if (!contexto || !contexto.ok) {
      return contexto || {
        ok: false,
        code: 'CERTIFICADO_CONTEXT_INVALIDO',
        msg: 'Não foi possível carregar o contexto do certificado.'
      };
    }

    if (!contexto.desafio_elegivel) {
      return {
        ok: false,
        code: 'DESAFIO_NAO_ELEGIVEL_CERTIFICADO',
        msg: 'O certificado só fica disponível para desafios concluídos, finalizados ou encerrados.'
      };
    }

    var statusValidacao = String(contexto.status_validacao_certificado || '').trim().toUpperCase();
    if (statusValidacao !== 'APROVADO') {
      return {
        ok: false,
        code: 'CERTIFICADO_NAO_LIBERADO',
        msg: 'Seu certificado ainda não foi liberado pela administração.'
      };
    }

    if (contexto.link_certificado_existente) {
      return {
        ok: true,
        url: contexto.link_certificado_existente,
        reused: true
      };
    }

    var gerado = gerarCertificadoDesafio_(contexto);
    if (!gerado || gerado.ok === false) {
      return gerado || {
        ok: false,
        code: 'CERTIFICADO_NAO_GERADO',
        msg: 'Não foi possível gerar o certificado.'
      };
    }

    return {
      ok: true,
      url: String(gerado.url || ''),
      reused: false
    };
  } catch (err) {
    return {
      ok: false,
      code: 'CERTIFICADO_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao gerar certificado.'
    };
  }
}

function gerarCertificadoDesafio_(contexto) {
  var ctx = contexto || {};
  var nomeArquivo = [
    'Certificado',
    ctx.id_dgmb || 'sem-id',
    ctx.id_desafio || 'desafio'
  ].join('_') + '.pdf';
  var conteudo = [
    'CERTIFICADO DE CONCLUSÃO',
    '',
    'ID DGMB: ' + String(ctx.id_dgmb || ''),
    'Desafio: ' + String(ctx.id_desafio || ''),
    'Item: ' + String(ctx.id_item_estoque || ''),
    'Status de validação: APROVADO',
    'Data: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  ];
  var html = HtmlService.createHtmlOutput(
    '<html><body style=\"font-family:Arial,sans-serif;padding:24px;\">' +
      conteudo.map(function(linha) {
        return '<p style=\"margin:0 0 10px 0;\">' + String(linha || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</p>';
      }).join('') +
    '</body></html>'
  );
  var blobPdf = html.getBlob().getAs(MimeType.PDF).setName(nomeArquivo);
  var arquivo = DriveApp.createFile(blobPdf);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = String(arquivo.getUrl() || '').trim();
  if (!url) {
    return { ok: false, code: 'CERTIFICADO_URL_INVALIDA', msg: 'Não foi possível gerar a URL do certificado.' };
  }

  if (ctx.sheet_name && ctx.rowNumber && ctx.idx_link_certificado > -1) {
    var sh = getSpreadsheet_().getSheetByName(ctx.sheet_name);
    if (sh) {
      sh.getRange(ctx.rowNumber, ctx.idx_link_certificado + 1).setValue(url);
    }
  }

  return { ok: true, url: url };
}

function certificadoBuscarContextoDesafio_(payload) {
  var params = payload || {};
  var idDgmb = normalizeText_(params.id_dgmb || params.idDgmb);
  var idDesafioFiltro = normalizeText_(params.id_desafio || params.idDesafio);
  var idItemFiltro = normalizeText_(params.id_item_estoque || params.idItemEstoque);

  if (!idDgmb) {
    return { ok: false, code: 'ID_DGMB_OBRIGATORIO', msg: 'ID do usuário é obrigatório para buscar o certificado.' };
  }

  var ss = getSpreadsheet_();
  var sheetName = SHEETS.DESAFIO || 'dgmbDesafios';
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    return { ok: false, code: 'ABA_DESAFIO_NAO_ENCONTRADA', msg: 'Aba dgmbDesafios não encontrada.' };
  }

  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { ok: false, code: 'DESAFIO_SEM_DADOS', msg: 'Não há registros de desafios para este usuário.' };
  }

  var map = buildHeaderMap_(values[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], sheetName);
  var idxIdDesafio = getOptionalColumnIndex_(map, ['id_desafio']);
  var idxIdItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxStatusApuracao = getOptionalColumnIndex_(map, ['status_apuracao', 'status apuracao', 'status apuração', 'status_desafio', 'status desafio']);
  var idxStatusValidacao = getRequiredColumnIndex_(map, ['status_validacao_certificado'], sheetName);

  var idxPrintCert = getRequiredColumnIndex_(map, ['print_strava_certificado'], sheetName);
  var idxLinkPrint = getRequiredColumnIndex_(map, ['link_print_strava'], sheetName);
  var idxDataEnvio = getRequiredColumnIndex_(map, ['data_envio_print_strava'], sheetName);
  var idxDataAprov = getRequiredColumnIndex_(map, ['data_aprovacao_certificado'], sheetName);
  var idxObs = getRequiredColumnIndex_(map, ['obs_validacao_certificado'], sheetName);

  var idxLinkCert = getOptionalColumnIndex_(map, ['link_certificado', 'url_certificado', 'certificado_url']);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = normalizeText_(row[idxId]);
    if (rowId !== idDgmb) continue;

    var rowDesafio = idxIdDesafio > -1 ? normalizeText_(row[idxIdDesafio]) : '';
    var rowItem = idxIdItem > -1 ? normalizeText_(row[idxIdItem]) : '';

    if (idDesafioFiltro && rowDesafio !== idDesafioFiltro) continue;
    if (idItemFiltro && rowItem !== idItemFiltro) continue;

    var statusApuracao = idxStatusApuracao > -1
      ? normalizeText_(row[idxStatusApuracao]).toUpperCase()
      : '';

    var desafioElegivel = {
      CONCLUIDO: true,
      FINALIZADO: true,
      ENCERRADO: true
    }[statusApuracao] === true;

    return {
      ok: true,
      rowNumber: i + 1,
      id_dgmb: rowId,
      id_desafio: rowDesafio,
      id_item_estoque: rowItem,
      status_apuracao: statusApuracao,
      desafio_elegivel: desafioElegivel,
      status_validacao_certificado: normalizeText_(row[idxStatusValidacao]).toUpperCase(),
      print_strava_certificado: normalizeText_(row[idxPrintCert]),
      link_print_strava: normalizeText_(row[idxLinkPrint]),
      data_envio_print_strava: row[idxDataEnvio] || '',
      data_aprovacao_certificado: row[idxDataAprov] || '',
      obs_validacao_certificado: normalizeText_(row[idxObs]),
      link_certificado_existente: idxLinkCert > -1 ? normalizeText_(row[idxLinkCert]) : '',
      sheet_name: sheetName,
      idx_link_certificado: idxLinkCert
    };
  }

  return {
    ok: false,
    code: 'DESAFIO_NAO_ENCONTRADO',
    msg: 'Desafio não encontrado para este usuário.'
  };
}
