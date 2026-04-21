var CERTIFICADO_PASTA_BASE_ID_ = '1GncBumQM3RAS6WIT0jHQPaIMKBlT7OHi';
var TEMPLATE_CERTIFICADO_SLIDES_ID_ = '';

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

    var linkPlanilha = certificadoLerLinkPlanilha_(contexto);
    if (certLinkValido_(linkPlanilha)) {
      return {
        ok: true,
        url: linkPlanilha,
        reused: true
      };
    }

    if (certLinkValido_(contexto.link_certificado_existente)) {
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
      var saveExistente = certificadoSalvarLinkPlanilha_(ctx, urlExistente);
      if (!saveExistente.ok) return saveExistente;
      return { ok: true, url: urlExistente, reused: true };
    }
  }

  var templateId = String(TEMPLATE_CERTIFICADO_SLIDES_ID_ || '').trim();
  if (!templateId) {
    return {
      ok: false,
      code: 'CERTIFICADO_TEMPLATE_SLIDES_NAO_CONFIGURADO',
      msg: 'Template do certificado em Google Slides não configurado.'
    };
  }

  var arquivoTemporario = null;
  var arquivo = null;
  try {
    var templateFile = DriveApp.getFileById(templateId);
    arquivoTemporario = templateFile.makeCopy('tmp_' + nomeArquivo.replace(/\.pdf$/i, '') + '_' + new Date().getTime());
    var apresentacao = SlidesApp.openById(arquivoTemporario.getId());
    var slides = apresentacao.getSlides();
    if (!slides || !slides.length) {
      return {
        ok: false,
        code: 'CERTIFICADO_TEMPLATE_SEM_SLIDE',
        msg: 'Template de certificado sem slide válido.'
      };
    }

    var slide = slides[0];
    var frase = 'Você não apenas concluiu o desafio. Você provou que é capaz de ir além.';
    var dataGeracao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var placeholders = {
      '{{NOME}}': dadosVisuais.nome_participante || ('Participante ' + (ctx.id_dgmb || '')),
      '{{DESAFIO}}': dadosVisuais.nome_desafio || ('ID ' + (ctx.id_desafio || '')),
      '{{META}}': dadosVisuais.meta_km || '-',
      '{{KM_REALIZADO}}': dadosVisuais.km_realizado || '-',
      '{{STATUS}}': dadosVisuais.status_desafio || 'CONCLUÍDO',
      '{{PERIODO}}': dadosVisuais.periodo || '-',
      '{{FRASE}}': frase,
      '{{DATA_EMISSAO}}': dataGeracao,
      '{{ID_DGMB}}': String(ctx.id_dgmb || ''),
      '{{ID_DESAFIO}}': String(ctx.id_desafio || '')
    };

    Object.keys(placeholders).forEach(function(chave) {
      slide.replaceAllText(chave, String(placeholders[chave] || ''));
    });

    apresentacao.saveAndClose();
    var blobPdf = DriveApp.getFileById(arquivoTemporario.getId()).getBlob().getAs(MimeType.PDF).setName(nomeArquivo);
    arquivo = pastaDestino.createFile(blobPdf);
  } catch (e) {
    return {
      ok: false,
      code: 'CERTIFICADO_GERACAO_SLIDES_ERROR',
      msg: e && e.message ? e.message : 'Erro ao gerar certificado via Google Slides.'
    };
  } finally {
    if (arquivoTemporario) {
      try {
        arquivoTemporario.setTrashed(true);
      } catch (trashErr) {}
    }
  }

  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = String(arquivo.getUrl() || '').trim();
  if (!url) {
    return { ok: false, code: 'CERTIFICADO_URL_INVALIDA', msg: 'Não foi possível gerar a URL do certificado.' };
  }

  var saveNovo = certificadoSalvarLinkPlanilha_(ctx, url);
  if (!saveNovo.ok) return saveNovo;

  return { ok: true, url: url };
}

function gerarHtmlCertificadoDesafio_(ctx, dados) {
  var payload = dados || {};
  var frase = 'Você não apenas concluiu o desafio. Você provou que é capaz de ir além.';
  var dataGeracao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
      '<meta charset="UTF-8">',
      '<style>',
        '@page { size: A4 landscape; margin: 0; }',
        'html, body { margin: 0; padding: 0; width: 100%; height: 100%; font-family: Arial, sans-serif; background: #efefef; }',
        '.page { width: 1123px; height: 794px; position: relative; overflow: hidden; color: #111; background: #f5f5f5; }',
        '.top-band { height: 22px; background: #f1c40f; }',
        '.frame { position: absolute; inset: 22px 28px 26px 28px; border: 2px solid #1e1e1e; background: #fff; }',
        '.header { background: linear-gradient(90deg, #111 0%, #222 58%, #2f2f2f 100%); color: #fff; padding: 24px 42px 20px 42px; border-bottom: 6px solid #f1c40f; }',
        '.title { margin: 0; font-size: 42px; font-weight: 800; letter-spacing: 1.6px; text-transform: uppercase; }',
        '.subtitle { margin: 8px 0 0 0; font-size: 17px; color: #d7d7d7; letter-spacing: .3px; }',
        '.content { padding: 30px 42px 0 42px; }',
        '.name-label { margin: 0; font-size: 13px; text-transform: uppercase; font-weight: 700; color: #666; letter-spacing: 1px; }',
        '.name { margin: 8px 0 8px 0; font-size: 50px; font-weight: 800; color: #101010; line-height: 1.1; }',
        '.challenge { margin: 0; font-size: 24px; line-height: 1.32; color: #222; }',
        '.challenge strong { color: #000; background: #f7d85c; padding: 2px 8px; }',
        '.divider { margin: 20px 0 16px 0; height: 4px; background: linear-gradient(90deg, #f1c40f 0%, #f1c40f 38%, #3d3d3d 38%, #3d3d3d 100%); }',
        '.cards { width: 100%; border-collapse: separate; border-spacing: 12px 12px; }',
        '.card { width: 50%; border: 1px solid #d9d9d9; background: #f8f8f8; padding: 12px 16px; }',
        '.card-label { margin: 0 0 7px 0; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .9px; color: #575757; }',
        '.card-value { margin: 0; font-size: 23px; font-weight: 700; color: #111; }',
        '.impact { margin: 10px 0 0 0; padding: 16px 18px; border-left: 8px solid #f1c40f; background: #191919; color: #fff; font-size: 21px; line-height: 1.36; font-weight: 700; }',
        '.footer { position: absolute; left: 42px; right: 42px; bottom: 26px; font-size: 12px; color: #6d6d6d; border-top: 1px solid #d8d8d8; padding-top: 10px; }',
      '</style>',
    '</head>',
    '<body>',
      '<div class="page">',
        '<div class="top-band"></div>',
        '<div class="frame">',
          '<div class="header">',
            '<p class="title">Certificado</p>',
            '<p class="subtitle">Meu Giro / DGMB • Reconhecimento Oficial de Conclusão</p>',
          '</div>',
          '<div class="content">',
            '<p class="name-label">Participante</p>',
            '<p class="name">' + certEscapeHtml_(payload.nome_participante || ('Participante ' + (ctx.id_dgmb || ''))) + '</p>',
            '<p class="challenge">Concluiu com sucesso o desafio <strong>' + certEscapeHtml_(payload.nome_desafio || ('ID ' + (ctx.id_desafio || ''))) + '</strong>.</p>',
            '<div class="divider"></div>',
            '<table class="cards">',
              '<tr>',
                '<td class="card"><p class="card-label">Meta</p><p class="card-value">' + certEscapeHtml_(payload.meta_km) + '</p></td>',
                '<td class="card"><p class="card-label">KM realizado</p><p class="card-value">' + certEscapeHtml_(payload.km_realizado) + '</p></td>',
              '</tr>',
              '<tr>',
                '<td class="card"><p class="card-label">Status</p><p class="card-value">' + certEscapeHtml_(payload.status_desafio || 'CONCLUÍDO') + '</p></td>',
                '<td class="card"><p class="card-label">Período</p><p class="card-value">' + certEscapeHtml_(payload.periodo) + '</p></td>',
              '</tr>',
            '</table>',
            '<p class="impact">' + certEscapeHtml_(frase) + '</p>',
          '</div>',
          '<p class="footer">Emitido em: ' + certEscapeHtml_(dataGeracao) + ' • ID DGMB: ' + certEscapeHtml_(ctx.id_dgmb || '') + ' • Referência do desafio: ' + certEscapeHtml_(ctx.id_desafio || '') + '</p>',
        '</div>',
      '</div>',
    '</body>',
    '</html>'
  ].join('');
}

function certificadoSalvarLinkPlanilha_(ctx, url) {
  if (!(ctx && ctx.sheet_name && ctx.rowNumber && ctx.idx_link_certificado > -1)) {
    return {
      ok: false,
      code: 'CERTIFICADO_LINK_COLUNA_INVALIDA',
      msg: 'Não foi possível salvar o LINK_CERTIFICADO: coluna não encontrada na aba de desafios.'
    };
  }
  var sh = getSpreadsheet_().getSheetByName(ctx.sheet_name);
  if (!sh) {
    return {
      ok: false,
      code: 'CERTIFICADO_LINK_ABA_INVALIDA',
      msg: 'Não foi possível salvar o LINK_CERTIFICADO: aba de desafios não encontrada.'
    };
  }
  var range = sh.getRange(ctx.rowNumber, ctx.idx_link_certificado + 1);
  var atual = String(range.getValue() || '').trim();
  if (atual === String(url || '').trim()) return { ok: true };
  range.setValue(url);
  SpreadsheetApp.flush();
  var salvo = String(range.getValue() || '').trim();
  if (salvo !== String(url || '').trim()) {
    return {
      ok: false,
      code: 'CERTIFICADO_LINK_NAO_PERSISTIDO',
      msg: 'Não foi possível confirmar a gravação do LINK_CERTIFICADO na planilha.'
    };
  }
  return { ok: true };
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

function certificadoLerLinkPlanilha_(ctx) {
  if (!(ctx && ctx.sheet_name && ctx.rowNumber && ctx.idx_link_certificado > -1)) return '';
  var sh = getSpreadsheet_().getSheetByName(ctx.sheet_name);
  if (!sh) return '';
  return String(sh.getRange(ctx.rowNumber, ctx.idx_link_certificado + 1).getValue() || '').trim();
}

function certLinkValido_(url) {
  var u = String(url || '').trim();
  return /^https?:\/\/\S+/i.test(u);
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
