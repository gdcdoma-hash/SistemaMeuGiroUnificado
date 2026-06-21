var CERTIFICADO_PASTA_BASE_ID_ = '1GncBumQM3RAS6WIT0jHQPaIMKBlT7OHi';
var TEMPLATE_CERTIFICADO_SLIDES_ID_ = '13BP2rHBiqymQyOk1bJsNFsSPxJAhQEPHj5FRIuAssfo';

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

    if (contexto.status_usuario_desafio === 'CANCELADO') {
      return {
        ok: false,
        code: 'INSCRICAO_CANCELADA',
        msg: 'Inscrição cancelada.'
      };
    }

    if (contexto.status_apuracao !== 'CONCLUIDO') {
      return {
        ok: false,
        code: 'DESAFIO_NAO_ELEGIVEL_CERTIFICADO',
        msg: contexto.status_apuracao === 'EXPIRADO' && contexto.status_usuario_desafio === 'NAO_CONCLUIDO'
          ? 'Desafio encerrado sem conclusão da meta.'
          : 'Conclua sua meta para liberar seu certificado.'
      };
    }

    if (contexto.status_usuario_desafio !== 'CONCLUIDO') {
      return {
        ok: false,
        code: 'CONCLUSAO_ADMINISTRATIVA_PENDENTE',
        msg: 'Meta atingida. Aguardando validação da organização.'
      };
    }

    var linkPlanilha = certificadoLerLinkPlanilha_(contexto);
    if (certLinkValido_(linkPlanilha)) {
      var extrasPlanilha = certificadoMontarExtrasImagem_(contexto);
      return {
        ok: true,
        url: linkPlanilha,
        pdfDownloadUrl: montarUrlDownloadPdfCertificado_(linkPlanilha),
        reused: true,
        imageUrl: extrasPlanilha.imageUrl || '',
        downloadImageUrl: extrasPlanilha.downloadImageUrl || '',
        whatsAppUrl: extrasPlanilha.whatsAppUrl || ''
      };
    }

    if (certLinkValido_(contexto.link_certificado_existente)) {
      var extrasExistente = certificadoMontarExtrasImagem_(contexto);
      return {
        ok: true,
        url: contexto.link_certificado_existente,
        pdfDownloadUrl: montarUrlDownloadPdfCertificado_(contexto.link_certificado_existente),
        reused: true,
        imageUrl: extrasExistente.imageUrl || '',
        downloadImageUrl: extrasExistente.downloadImageUrl || '',
        whatsAppUrl: extrasExistente.whatsAppUrl || ''
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

    var extrasNovo = certificadoMontarExtrasImagem_(contexto);
    return {
      ok: true,
      url: String(gerado.url || ''),
      pdfDownloadUrl: String(gerado.pdfDownloadUrl || ''),
      reused: false,
      imageUrl: extrasNovo.imageUrl || '',
      downloadImageUrl: extrasNovo.downloadImageUrl || '',
      whatsAppUrl: extrasNovo.whatsAppUrl || ''
    };
  } catch (err) {
    return {
      ok: false,
      code: 'CERTIFICADO_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao gerar certificado.'
    };
  }
}

function certificadoMontarExtrasImagem_(contexto) {
  var ctx = contexto || {};
  if (ctx._certificado_imagem_processado_) {
    return ctx._certificado_imagem_processado_;
  }
  try {
    var imagem = gerarCertificadoImagem_(ctx);
    if (!imagem || imagem.ok !== true) {
      ctx._certificado_imagem_processado_ = {};
      return ctx._certificado_imagem_processado_;
    }
    ctx._certificado_imagem_processado_ = {
      imageUrl: String(imagem.imageUrl || '').trim(),
      downloadImageUrl: String(imagem.downloadImageUrl || '').trim(),
      whatsAppUrl: String(imagem.whatsAppUrl || '').trim()
    };
    return ctx._certificado_imagem_processado_;
  } catch (e) {
    ctx._certificado_imagem_processado_ = {};
    return ctx._certificado_imagem_processado_;
  }
}

function gerarCertificadoDesafio_(contexto) {
  var ctx = contexto || {};
  var nomeArquivoPartes = [
    'certificado',
    ctx.id_dgmb || 'sem-id',
    ctx.id_desafio || 'desafio'
  ];
  if (ctx.id_inscricao) nomeArquivoPartes.push(ctx.id_inscricao);
  var nomeArquivo = nomeArquivoPartes.join('_') + '.pdf';
  var dadosVisuais = certificadoBuscarDadosVisuais_(ctx);
  var pastaDestino = certificadoGetOuCriarPastaDesafio_(ctx.id_desafio);
  var arquivoExistente = certificadoBuscarArquivoExistente_(pastaDestino, nomeArquivo);
  if (arquivoExistente) {
    var urlExistente = String(arquivoExistente.getUrl() || '').trim();
    if (urlExistente) {
      var saveExistente = certificadoSalvarLinkPlanilha_(ctx, urlExistente);
      if (!saveExistente.ok) return saveExistente;
      return {
        ok: true,
        url: urlExistente,
        pdfDownloadUrl: montarUrlDownloadPdfCertificado_(arquivoExistente.getId() || urlExistente),
        reused: true
      };
    }
  }

  var templatePadraoId = String(TEMPLATE_CERTIFICADO_SLIDES_ID_ || '').trim();
  if (!templatePadraoId) {
    return {
      ok: false,
      code: 'CERTIFICADO_TEMPLATE_SLIDES_NAO_CONFIGURADO',
      msg: 'Template do certificado em Google Slides não configurado.'
    };
  }

  var resolucao = certificadoResolverTemplateSlides_(ctx.id_desafio, ctx.id_item_estoque);
  var templateResolvidoId = resolucao.templateId || templatePadraoId;
  var geracao = certificadoGerarPdfComTemplate_(
    templateResolvidoId,
    nomeArquivo,
    pastaDestino,
    ctx,
    dadosVisuais
  );
  if (!geracao.ok && resolucao.source !== 'PADRAO') {
    templateResolvidoId = templatePadraoId;
    geracao = certificadoGerarPdfComTemplate_(templatePadraoId, nomeArquivo, pastaDestino, ctx, dadosVisuais);
  }
  if (!geracao.ok) return geracao;
  ctx._certificado_template_slides_id_ = templateResolvidoId;

  var arquivo = geracao.arquivo;
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = String(arquivo.getUrl() || '').trim();
  if (!url) {
    return { ok: false, code: 'CERTIFICADO_URL_INVALIDA', msg: 'Não foi possível gerar a URL do certificado.' };
  }

  var saveNovo = certificadoSalvarLinkPlanilha_(ctx, url);
  if (!saveNovo.ok) return saveNovo;

  return {
    ok: true,
    url: url,
    pdfDownloadUrl: montarUrlDownloadPdfCertificado_(arquivo.getId() || url)
  };
}

function certificadoResolverTemplateSlides_(idDesafio, idItemEstoque) {
  var templatePadrao = String(TEMPLATE_CERTIFICADO_SLIDES_ID_ || '').trim();
  var padrao = function(reason) {
    return { templateId: templatePadrao, source: 'PADRAO', fallback: true, reason: reason };
  };
  var desafio = String(idDesafio == null ? '' : idDesafio).trim();
  var item = String(idItemEstoque == null ? '' : idItemEstoque).trim();
  if (!desafio) return padrao('ID_DESAFIO_VAZIO');

  try {
    var nomeAba = (typeof SHEETS !== 'undefined' && SHEETS.CONFIG_CERTIFICADO_TEMPLATE)
      ? SHEETS.CONFIG_CERTIFICADO_TEMPLATE
      : 'CONFIG_CERTIFICADO_TEMPLATE';
    var sh = getSpreadsheet_().getSheetByName(nomeAba);
    if (!sh) return padrao('ABA_AUSENTE');
    var values = sh.getDataRange().getValues();
    if (!values || !values.length) return padrao('ABA_VAZIA');

    var layout = certificadoTemplateObterLayout_(values[0] || []);
    if (!layout) return padrao('CABECALHO_INCOMPLETO');

    var especificas = [];
    var gerais = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i] || [];
      if (!certificadoTemplateAtivo_(row[layout.ativo])) continue;
      if (String(row[layout.idDesafio] == null ? '' : row[layout.idDesafio]).trim() !== desafio) continue;
      var templateId = String(row[layout.templateId] == null ? '' : row[layout.templateId]).trim();
      if (!templateId) continue;
      var rowItem = String(row[layout.idItemEstoque] == null ? '' : row[layout.idItemEstoque]).trim();
      if (item && rowItem === item) especificas.push(templateId);
      if (!rowItem) gerais.push(templateId);
    }

    if (item && especificas.length > 1) return padrao('DUPLICIDADE_ATIVA_ITEM');
    if (item && especificas.length === 1) {
      return { templateId: especificas[0], source: 'ITEM', fallback: false, reason: 'CONFIG_ITEM' };
    }
    if (gerais.length > 1) return padrao('DUPLICIDADE_ATIVA_DESAFIO');
    if (gerais.length === 1) {
      return { templateId: gerais[0], source: 'DESAFIO', fallback: false, reason: 'CONFIG_DESAFIO' };
    }
    return padrao('CONFIGURACAO_NAO_ENCONTRADA');
  } catch (e) {
    return padrao('ERRO_LEITURA_CONFIGURACAO');
  }
}

function certificadoTemplateAtivo_(valor) {
  if (valor === true || valor === 1) return true;
  var texto = String(valor == null ? '' : valor).trim().toUpperCase();
  return texto === 'TRUE' || texto === 'VERDADEIRO' || texto === 'SIM' || texto === '1';
}

function certificadoTemplateObterLayout_(headers) {
  var indices = {};
  for (var i = 0; i < headers.length; i++) {
    indices[String(headers[i] == null ? '' : headers[i]).trim().toUpperCase()] = i;
  }
  var obrigatorios = ['ID_DESAFIO', 'ID_ITEM_ESTOQUE', 'ID_SLIDE_TEMPLATE', 'ATIVO'];
  for (var j = 0; j < obrigatorios.length; j++) {
    if (!Object.prototype.hasOwnProperty.call(indices, obrigatorios[j])) return null;
  }
  return {
    idDesafio: indices.ID_DESAFIO,
    idItemEstoque: indices.ID_ITEM_ESTOQUE,
    templateId: indices.ID_SLIDE_TEMPLATE,
    ativo: indices.ATIVO
  };
}

function certificadoGerarPdfComTemplate_(templateId, nomeArquivo, pastaDestino, ctx, dadosVisuais) {
  var arquivoTemporario = null;
  try {
    var templateFile = DriveApp.getFileById(templateId);
    arquivoTemporario = templateFile.makeCopy('tmp_' + nomeArquivo.replace(/\.pdf$/i, '') + '_' + new Date().getTime(), pastaDestino);
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
    return { ok: true, arquivo: pastaDestino.createFile(blobPdf) };
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

  var identidadeAtual = certificadoRevalidarLinhaInscricao_(sh, ctx);
  if (!identidadeAtual.ok) return identidadeAtual;

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

function certificadoRevalidarLinhaInscricao_(sh, ctx) {
  var lastColumn = Math.max(Number(sh.getLastColumn() || 0), Number(ctx.idx_link_certificado || 0) + 1);
  var row = sh.getRange(ctx.rowNumber, 1, 1, lastColumn).getValues()[0] || [];
  var rowIdDgmb = normalizeText_(row[ctx.idx_id_dgmb]);
  var rowIdInscricao = ctx.idx_id_inscricao > -1 ? normalizeText_(row[ctx.idx_id_inscricao]) : '';

  if (rowIdDgmb !== normalizeText_(ctx.id_dgmb) ||
      (ctx.id_inscricao && rowIdInscricao !== normalizeText_(ctx.id_inscricao))) {
    return {
      ok: false,
      code: 'CERTIFICADO_INSCRICAO_ALTERADA',
      msg: 'A inscrição do certificado foi alterada antes da gravação. Recarregue os dados e tente novamente.'
    };
  }

  if (!ctx.id_inscricao) {
    var rowDesafio = obterIdDesafioRegistro_(row, ctx.idx_id_desafio, ctx.idx_observacao);
    var rowItem = ctx.idx_id_item > -1 ? normalizeText_(row[ctx.idx_id_item]) : '';
    if (rowDesafio !== normalizeText_(ctx.id_desafio) || rowItem !== normalizeText_(ctx.id_item_estoque)) {
      return {
        ok: false,
        code: 'CERTIFICADO_INSCRICAO_ALTERADA',
        msg: 'O registro legado do certificado foi alterado antes da gravação. Recarregue os dados e tente novamente.'
      };
    }
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
  var resumo = certificadoBuscarResumoDesafio_(ctx);
  var nome = certificadoBuscarNomeParticipante_(ctx.id_dgmb);
  var status = normalizeText_(resumo.status_apuracao || ctx.status_apuracao).toUpperCase();

  return {
    nome_participante: nome || '',
    nome_desafio: resumo.nome_desafio || ('Desafio ' + String(ctx.id_desafio || '')),
    meta_km: certFormatKm_(resumo.meta_km),
    km_realizado: certFormatKm_(resumo.distancia_realizada),
    percentual_concluido: parseLocalizedNumber_(resumo.percentual_concluido),
    status_desafio: status || 'CONCLUÍDO',
    periodo: certFormatPeriodo_(resumo.periodo_inicio, resumo.periodo_fim)
  };
}

function certificadoBuscarResumoDesafio_(contextoOuIdDgmb, idDesafio, idItemEstoque) {
  var ctx = typeof contextoOuIdDgmb === 'object'
    ? (contextoOuIdDgmb || {})
    : { id_dgmb: contextoOuIdDgmb, id_desafio: idDesafio, id_item_estoque: idItemEstoque };
  var id = normalizeText_(ctx.id_dgmb);
  var inscricao = normalizeText_(ctx.id_inscricao);
  var desafio = normalizeText_(ctx.id_desafio);
  var item = normalizeText_(ctx.id_item_estoque);
  if (!id || (!inscricao && !desafio)) return {};

  var resumo = [];
  try {
    resumo = atualizarMeuGiroResumo_(id) || [];
  } catch (e) {
    resumo = [];
  }

  if (inscricao) {
    for (var i = 0; i < resumo.length; i++) {
      var rowInscricao = resumo[i] || {};
      if (normalizeText_(rowInscricao.id_inscricao) !== inscricao) continue;
      if (normalizeText_(rowInscricao.id_dgmb || id) !== id) continue;
      return rowInscricao;
    }
    return {};
  }

  for (var j = 0; j < resumo.length; j++) {
    var row = resumo[j] || {};
    if (normalizeText_(row.id_desafio) !== desafio) continue;
    if (normalizeText_(row.id_item_estoque) !== item) continue;
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

function montarUrlDownloadPdfCertificado_(urlOuId) {
  var fileId = extrairDriveFileIdCertificado_(urlOuId);
  if (!fileId) return '';
  return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId);
}

function extrairDriveFileIdCertificado_(urlOuId) {
  var txt = String(urlOuId || '').trim();
  if (!txt) return '';
  var byIdMatch = txt.match(/[-\w]{25,}/);
  if (!/^https?:\/\//i.test(txt)) {
    return byIdMatch ? byIdMatch[0] : '';
  }
  var patterns = [
    /\/d\/([-\w]{25,})/i,
    /[?&]id=([-\w]{25,})/i,
    /\/file\/d\/([-\w]{25,})/i
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = txt.match(patterns[i]);
    if (match && match[1]) return match[1];
  }
  return byIdMatch ? byIdMatch[0] : '';
}

function certificadoBuscarContextoDesafio_(payload) {
  var params = payload || {};
  var idInscricao = normalizeText_(params.id_inscricao || params.idInscricao);
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
  var idxIdInscricao = getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxIdDesafio = getIdDesafioColumnIndex_(map);
  var idxIdItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxObservacao = getOptionalColumnIndex_(map, ['observacao', 'observação']);
  var idxStatusApuracao = getOptionalColumnIndex_(map, ['status_apuracao', 'status apuracao', 'status apuração', 'status_desafio', 'status desafio']);
  var idxStatusUsuarioDesafio = getRequiredColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio'], sheetName);
  var idxStatusValidacao = getRequiredColumnIndex_(map, ['status_validacao_certificado'], sheetName);
  var idxPrintCert = getRequiredColumnIndex_(map, ['print_strava_certificado'], sheetName);
  var idxLinkPrint = getRequiredColumnIndex_(map, ['link_print_strava'], sheetName);
  var idxDataEnvio = getRequiredColumnIndex_(map, ['data_envio_print_strava'], sheetName);
  var idxDataAprov = getRequiredColumnIndex_(map, ['data_aprovacao_certificado'], sheetName);
  var idxObs = getRequiredColumnIndex_(map, ['obs_validacao_certificado'], sheetName);
  var idxLinkCert = getOptionalColumnIndex_(map, ['link_certificado', 'url_certificado', 'certificado_url']);

  var linhaSelecionada = -1;
  if (idInscricao && idxIdInscricao > -1) {
    var candidatasInscricao = [];
    for (var i = 1; i < values.length; i++) {
      if (normalizeText_(values[i][idxId]) !== idDgmb) continue;
      if (normalizeText_(values[i][idxIdInscricao]) !== idInscricao) continue;
      candidatasInscricao.push(i);
    }
    if (candidatasInscricao.length > 1) {
      return {
        ok: false,
        code: 'CERTIFICADO_INSCRICAO_AMBIGUA',
        msg: 'Mais de uma linha possui o mesmo ID_INSCRICAO para este participante.'
      };
    }
    if (candidatasInscricao.length === 1) linhaSelecionada = candidatasInscricao[0];
  }

  if (linhaSelecionada === -1) {
    var candidatasLegadas = [];
    for (var j = 1; j < values.length; j++) {
      var rowLegada = values[j] || [];
      if (normalizeText_(rowLegada[idxId]) !== idDgmb) continue;
      if (idInscricao && idxIdInscricao > -1 && normalizeText_(rowLegada[idxIdInscricao])) continue;

      var desafioLegado = obterIdDesafioRegistro_(rowLegada, idxIdDesafio, idxObservacao);
      var itemLegado = idxIdItem > -1 ? normalizeText_(rowLegada[idxIdItem]) : '';
      if (idDesafioFiltro && desafioLegado !== idDesafioFiltro) continue;
      if (idItemFiltro && itemLegado !== idItemFiltro) continue;
      candidatasLegadas.push(j);
    }

    if (candidatasLegadas.length > 1) {
      return {
        ok: false,
        code: 'CERTIFICADO_INSCRICAO_AMBIGUA',
        msg: 'Mais de uma inscrição corresponde aos identificadores legados informados.'
      };
    }
    if (candidatasLegadas.length === 1) linhaSelecionada = candidatasLegadas[0];
  }

  if (linhaSelecionada === -1) {
    return { ok: false, code: 'DESAFIO_NAO_ENCONTRADO', msg: 'Desafio não encontrado para este usuário.' };
  }

  var row = values[linhaSelecionada] || [];
  var rowId = normalizeText_(row[idxId]);
  var rowInscricao = idxIdInscricao > -1 ? normalizeText_(row[idxIdInscricao]) : '';
  var rowDesafio = obterIdDesafioRegistro_(row, idxIdDesafio, idxObservacao);
  var rowItem = idxIdItem > -1 ? normalizeText_(row[idxIdItem]) : '';
  var statusApuracao = certificadoBuscarStatusApuracaoResumo_(rowInscricao, rowId, rowDesafio, rowItem);
  if (!statusApuracao && !rowInscricao && idxStatusApuracao > -1) {
    statusApuracao = normalizeText_(row[idxStatusApuracao]).toUpperCase();
  }

  var statusUsuarioDesafio = normalizeText_(row[idxStatusUsuarioDesafio]).toUpperCase();
  var statusValidacaoCertificado = normalizeText_(row[idxStatusValidacao]).toUpperCase();
  var desafioElegivel = statusApuracao === 'CONCLUIDO' &&
    statusUsuarioDesafio === 'CONCLUIDO';

  return {
    ok: true,
    rowNumber: linhaSelecionada + 1,
    id_inscricao: rowInscricao,
    id_dgmb: rowId,
    id_desafio: rowDesafio,
    id_item_estoque: rowItem,
    status_apuracao: statusApuracao,
    status_usuario_desafio: statusUsuarioDesafio,
    desafio_elegivel: desafioElegivel,
    status_validacao_certificado: statusValidacaoCertificado,
    print_strava_certificado: normalizeText_(row[idxPrintCert]),
    link_print_strava: normalizeText_(row[idxLinkPrint]),
    data_envio_print_strava: row[idxDataEnvio] || '',
    data_aprovacao_certificado: row[idxDataAprov] || '',
    obs_validacao_certificado: normalizeText_(row[idxObs]),
    link_certificado_existente: idxLinkCert > -1 ? normalizeText_(row[idxLinkCert]) : '',
    sheet_name: sheetName,
    idx_id_dgmb: idxId,
    idx_id_inscricao: idxIdInscricao,
    idx_id_desafio: idxIdDesafio,
    idx_id_item: idxIdItem,
    idx_observacao: idxObservacao,
    idx_link_certificado: idxLinkCert
  };
}

function certificadoBuscarStatusApuracaoResumo_(idInscricao, idDgmb, idDesafio, idItemEstoque) {
  var inscricao = normalizeText_(idInscricao);
  var id = normalizeText_(idDgmb);
  var desafio = normalizeText_(idDesafio);
  var item = normalizeText_(idItemEstoque);
  if (!id || (!inscricao && !desafio)) return '';

  var resumo = [];
  try {
    resumo = atualizarMeuGiroResumo_(id) || [];
  } catch (e) {
    resumo = [];
  }

  if (inscricao) {
    for (var i = 0; i < resumo.length; i++) {
      var rowInscricao = resumo[i] || {};
      if (normalizeText_(rowInscricao.id_inscricao) !== inscricao) continue;
      if (normalizeText_(rowInscricao.id_dgmb || id) !== id) continue;
      return normalizeText_(rowInscricao.status_apuracao).toUpperCase();
    }
    return '';
  }

  for (var j = 0; j < resumo.length; j++) {
    var row = resumo[j] || {};
    if (normalizeText_(row.id_desafio) !== desafio) continue;
    if (normalizeText_(row.id_item_estoque) !== item) continue;
    return normalizeText_(row.status_apuracao).toUpperCase();
  }

  return '';
}
