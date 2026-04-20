var CERTIFICADO_BACKGROUND_FILE_ID_ = '1vPHVb07i5fc5oKIM0g6LVdIE3JBze-ZR';
var CERTIFICADO_PASTA_BASE_ID_ = '1GncBumQM3RAS6WIT0jHQPaIMKBlT7OHi';

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
    'certificado',
    ctx.id_dgmb || 'sem-id',
    ctx.id_desafio || 'desafio'
  ].join('_') + '.pdf';
  var dadosVisuais = certificadoBuscarDadosVisuais_(ctx);
  var pastaDestino = certificadoGetOuCriarPastaDesafio_(ctx.id_desafio);
  var arquivoExistente = certificadoBuscarArquivoExistente_(pastaDestino, nomeArquivo);
  if (arquivoExistente) {
    var urlExistente = String(arquivoExistente.getUrl() || '').trim();
    if (urlExistente) {
      certificadoSalvarLinkPlanilha_(ctx, urlExistente);
      return { ok: true, url: urlExistente, reused: true };
    }
  }

  var html = HtmlService.createHtmlOutput(gerarHtmlCertificadoDesafio_(ctx, dadosVisuais));
  var blobPdf = html.getBlob().getAs(MimeType.PDF).setName(nomeArquivo);
  var arquivo = pastaDestino.createFile(blobPdf);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = String(arquivo.getUrl() || '').trim();
  if (!url) {
    return { ok: false, code: 'CERTIFICADO_URL_INVALIDA', msg: 'Não foi possível gerar a URL do certificado.' };
  }

  certificadoSalvarLinkPlanilha_(ctx, url);

  return { ok: true, url: url };
}

function gerarHtmlCertificadoDesafio_(ctx, dados) {
  var payload = dados || {};
  var frase = 'Você não apenas concluiu o desafio. Você provou que é capaz de ir além.';
  var dataGeracao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var backgroundDataUri = certGetBackgroundDataUri_();

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
      '<meta charset="UTF-8">',
      '<style>',
        '@page { size: A4 landscape; margin: 0; }',
        'html, body { margin: 0; padding: 0; width: 100%; height: 100%; font-family: Arial, sans-serif; }',
        '.page { width: 1123px; height: 794px; position: relative; overflow: hidden; color: #fff; }',
        '.bg { position: absolute; inset: 0; background-image: url("' + certEscapeHtml_(backgroundDataUri) + '"); background-size: cover; background-position: center; }',
        '.overlay { position: absolute; inset: 0; background: linear-gradient(90deg, rgba(0,0,0,0.66) 0%, rgba(0,0,0,0.50) 45%, rgba(0,0,0,0.10) 70%, rgba(0,0,0,0.00) 100%); }',
        '.content { position: relative; z-index: 2; width: 58%; padding: 122px 0 0 84px; }',
        '.title { font-size: 42px; font-weight: 800; letter-spacing: 1px; margin: 0 0 14px 0; text-transform: uppercase; }',
        '.name { font-size: 36px; font-weight: 700; margin: 0 0 10px 0; color: #ffd44d; line-height: 1.15; }',
        '.subtitle { font-size: 20px; margin: 0 0 22px 0; color: #f4f4f4; }',
        '.phrase { font-size: 18px; line-height: 1.5; margin: 0 0 26px 0; max-width: 95%; }',
        '.grid { width: 92%; border-collapse: collapse; border-spacing: 0; }',
        '.grid td { padding: 7px 0; vertical-align: top; border-bottom: 1px solid rgba(255,255,255,0.25); font-size: 16px; }',
        '.grid td.label { width: 165px; color: #ffe8a2; font-weight: 700; text-transform: uppercase; font-size: 13px; letter-spacing: .6px; }',
        '.footer { margin-top: 18px; font-size: 12px; color: #f2f2f2; opacity: .9; }',
      '</style>',
    '</head>',
    '<body>',
      '<div class="page">',
        '<div class="bg"></div>',
        '<div class="overlay"></div>',
        '<div class="content">',
          '<p class="title">Certificado</p>',
          '<p class="name">' + certEscapeHtml_(payload.nome_participante || ('Participante ' + (ctx.id_dgmb || ''))) + '</p>',
          '<p class="subtitle">Concluiu com sucesso o desafio <strong>' + certEscapeHtml_(payload.nome_desafio || ('ID ' + (ctx.id_desafio || ''))) + '</strong>.</p>',
          '<p class="phrase">' + certEscapeHtml_(frase) + '</p>',
          '<table class="grid">',
            '<tr><td class="label">Meta</td><td>' + certEscapeHtml_(payload.meta_km) + '</td></tr>',
            '<tr><td class="label">KM realizado</td><td>' + certEscapeHtml_(payload.km_realizado) + '</td></tr>',
            '<tr><td class="label">Status</td><td>' + certEscapeHtml_(payload.status_desafio || 'CONCLUÍDO') + '</td></tr>',
            '<tr><td class="label">Período</td><td>' + certEscapeHtml_(payload.periodo) + '</td></tr>',
          '</table>',
          '<p class="footer">ID DGMB: ' + certEscapeHtml_(ctx.id_dgmb || '') + ' • Desafio: ' + certEscapeHtml_(ctx.id_desafio || '') + ' • Emitido em: ' + certEscapeHtml_(dataGeracao) + '</p>',
        '</div>',
      '</div>',
    '</body>',
    '</html>'
  ].join('');
}

function certificadoSalvarLinkPlanilha_(ctx, url) {
  if (!(ctx && ctx.sheet_name && ctx.rowNumber && ctx.idx_link_certificado > -1)) return;
  var sh = getSpreadsheet_().getSheetByName(ctx.sheet_name);
  if (!sh) return;
  var range = sh.getRange(ctx.rowNumber, ctx.idx_link_certificado + 1);
  var atual = String(range.getValue() || '').trim();
  if (atual === String(url || '').trim()) return;
  range.setValue(url);
}

function certificadoBuscarArquivoExistente_(pasta, nomeArquivo) {
  if (!pasta || !nomeArquivo) return null;
  var arquivos = pasta.getFilesByName(nomeArquivo);
  if (arquivos.hasNext()) return arquivos.next();
  return null;
}

function certificadoBuscarDadosVisuais_(ctx) {
  var resumo = certificadoBuscarResumoDesafio_(ctx.id_dgmb, ctx.id_desafio, ctx.id_item_estoque);
  var nome = certificadoBuscarNomeParticipante_(ctx.id_dgmb);
  var status = normalizeText_(resumo.status_apuracao || ctx.status_apuracao).toUpperCase();

  return {
    nome_participante: nome || '',
    nome_desafio: resumo.nome_desafio || ('Desafio ' + String(ctx.id_desafio || '')),
    meta_km: certFormatKm_(resumo.meta_km),
    km_realizado: certFormatKm_(resumo.distancia_realizada),
    status_desafio: status || 'CONCLUÍDO',
    periodo: certFormatPeriodo_(resumo.periodo_inicio, resumo.periodo_fim)
  };
}

function certificadoBuscarResumoDesafio_(idDgmb, idDesafio, idItemEstoque) {
  var id = normalizeText_(idDgmb);
  var desafio = normalizeText_(idDesafio);
  var item = normalizeText_(idItemEstoque);
  if (!id || !desafio) return {};

  var resumo = [];
  try {
    resumo = atualizarMeuGiroResumo_(id) || [];
  } catch (e) {
    resumo = [];
  }

  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    if (normalizeText_(row.id_desafio) !== desafio) continue;
    if (item && normalizeText_(row.id_item_estoque) !== item) continue;
    return row;
  }

  return {};
}

function certificadoBuscarNomeParticipante_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return '';

  var sh = getSheetByName_(SHEETS.PESSOAS);
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return '';

  var map = buildHeaderMap_(values[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], SHEETS.PESSOAS);
  var idxNome = getRequiredColumnIndex_(map, ['nome'], SHEETS.PESSOAS);

  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    if (normalizeText_(row[idxId]) !== id) continue;
    return normalizeText_(row[idxNome]);
  }

  return '';
}

function certificadoGetOuCriarPastaDesafio_(idDesafio) {
  var pastaBase = DriveApp.getFolderById(CERTIFICADO_PASTA_BASE_ID_);
  var nomeSubpasta = 'ID_DESAFIO_' + String(idDesafio || 'sem-id').trim();
  var subpastas = pastaBase.getFoldersByName(nomeSubpasta);
  if (subpastas.hasNext()) return subpastas.next();
  return pastaBase.createFolder(nomeSubpasta);
}

function certFormatKm_(valor) {
  var n = parseLocalizedNumber_(valor);
  if (!isFinite(n) || n <= 0) return '-';
  return n.toFixed(1).replace('.', ',') + ' km';
}

function certFormatPeriodo_(inicio, fim) {
  var inicioFmt = certFormatDataPt_(inicio);
  var fimFmt = certFormatDataPt_(fim);
  if (inicioFmt && fimFmt) return inicioFmt + ' a ' + fimFmt;
  return inicioFmt || fimFmt || '-';
}

function certFormatDataPt_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  var txt = String(valor || '').trim();
  if (!txt) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) {
    return txt.split('-').reverse().join('/');
  }
  return txt;
}

function certEscapeHtml_(valor) {
  return String(valor || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function certGetBackgroundDataUri_() {
  var file = DriveApp.getFileById(CERTIFICADO_BACKGROUND_FILE_ID_);
  var blob = file.getBlob();
  var mime = String(blob.getContentType() || 'image/png').trim();
  var b64 = Utilities.base64Encode(blob.getBytes());
  return 'data:' + mime + ';base64,' + b64;
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
  var idxObservacao = getOptionalColumnIndex_(map, ['observacao', 'observação']);
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
    if (!rowDesafio && idxObservacao > -1) {
      var observacao = String(row[idxObservacao] || '');
      var matchDesafio = observacao.match(/\[\s*ID_DESAFIO\s*:\s*([0-9]+)\s*\]/i);
      rowDesafio = matchDesafio && matchDesafio[1] ? normalizeText_(matchDesafio[1]) : '';
    }
    var rowItem = idxIdItem > -1 ? normalizeText_(row[idxIdItem]) : '';

    if (idDesafioFiltro && rowDesafio !== idDesafioFiltro) continue;
    if (idItemFiltro && rowItem !== idItemFiltro) continue;

    var statusApuracao = certificadoBuscarStatusApuracaoResumo_(rowId, rowDesafio, rowItem);
    if (!statusApuracao && idxStatusApuracao > -1) {
      statusApuracao = normalizeText_(row[idxStatusApuracao]).toUpperCase();
    }

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

function certificadoBuscarStatusApuracaoResumo_(idDgmb, idDesafio, idItemEstoque) {
  var id = normalizeText_(idDgmb);
  var desafio = normalizeText_(idDesafio);
  var item = normalizeText_(idItemEstoque);
  if (!id || !desafio) return '';

  var resumo = [];
  try {
    resumo = atualizarMeuGiroResumo_(id) || [];
  } catch (e) {
    resumo = [];
  }

  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    var rowDesafio = normalizeText_(row.id_desafio);
    var rowItem = normalizeText_(row.id_item_estoque);
    if (rowDesafio !== desafio) continue;
    if (item && rowItem !== item) continue;
    if (!item && rowItem) continue;
    return normalizeText_(row.status_apuracao).toUpperCase();
  }

  for (var j = 0; j < resumo.length; j++) {
    var rowFallback = resumo[j] || {};
    if (normalizeText_(rowFallback.id_desafio) !== desafio) continue;
    return normalizeText_(rowFallback.status_apuracao).toUpperCase();
  }

  return '';
}
