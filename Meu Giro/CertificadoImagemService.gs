var TEMPLATE_CERTIFICADO_IMAGEM_SLIDES_ID_ = '';

function gerarCertificadoImagem_(contexto) {
  var ctx = contexto || {};
  var dadosVisuais = certificadoBuscarDadosVisuais_(ctx);
  var pastaDestino = certificadoGetOuCriarPastaDesafio_(ctx.id_desafio);
  var nomeArquivo = [
    'certificado_imagem',
    ctx.id_dgmb || 'sem-id',
    ctx.id_desafio || 'desafio'
  ].join('_') + '.png';

  var arquivoExistente = certificadoBuscarArquivoExistente_(pastaDestino, nomeArquivo);
  if (arquivoExistente) {
    arquivoExistente.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var urlPublicaExistente = montarUrlPublicaImagemCertificado_(arquivoExistente.getId());
    return {
      ok: true,
      reused: true,
      imageUrl: urlPublicaExistente,
      downloadImageUrl: montarUrlDownloadImagemCertificado_(arquivoExistente.getId()),
      whatsAppUrl: montarLinkWhatsAppCertificado_(dadosVisuais.nome_participante, urlPublicaExistente)
    };
  }

  if (!TEMPLATE_CERTIFICADO_IMAGEM_SLIDES_ID_) {
    return {
      ok: false,
      code: 'CERTIFICADO_IMAGEM_TEMPLATE_NAO_CONFIGURADO',
      msg: 'Template do certificado em imagem não configurado.'
    };
  }
  var templateId = String(TEMPLATE_CERTIFICADO_IMAGEM_SLIDES_ID_).trim();

  var arquivoTemporario = null;
  var arquivoImagem = null;
  try {
    var templateFile = DriveApp.getFileById(templateId);
    arquivoTemporario = templateFile.makeCopy('tmp_img_' + nomeArquivo.replace(/\.png$/i, '') + '_' + new Date().getTime(), pastaDestino);
    var apresentacao = SlidesApp.openById(arquivoTemporario.getId());
    var slides = apresentacao.getSlides();
    if (!slides || !slides.length) {
      return {
        ok: false,
        code: 'CERTIFICADO_IMAGEM_TEMPLATE_SEM_SLIDE',
        msg: 'Template de certificado em imagem sem slide válido.'
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

    var blob = exportarSlideComoImagem_(arquivoTemporario.getId(), slide.getObjectId());
    arquivoImagem = salvarImagemCertificadoNoDrive_(blob, nomeArquivo, pastaDestino);
  } catch (e) {
    return {
      ok: false,
      code: 'CERTIFICADO_IMAGEM_GERACAO_ERROR',
      msg: e && e.message ? e.message : 'Erro ao gerar imagem do certificado.'
    };
  } finally {
    if (arquivoTemporario) {
      try {
        arquivoTemporario.setTrashed(true);
      } catch (trashErr) {}
    }
  }

  if (!arquivoImagem) {
    return {
      ok: false,
      code: 'CERTIFICADO_IMAGEM_ARQUIVO_INVALIDO',
      msg: 'Não foi possível salvar a imagem do certificado.'
    };
  }

  arquivoImagem.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var imageUrl = montarUrlPublicaImagemCertificado_(arquivoImagem.getId());
  return {
    ok: true,
    reused: false,
    imageUrl: imageUrl,
    downloadImageUrl: montarUrlDownloadImagemCertificado_(arquivoImagem.getId()),
    whatsAppUrl: montarLinkWhatsAppCertificado_(dadosVisuais.nome_participante, imageUrl)
  };
}

function exportarSlideComoImagem_(presentationId, slideId) {
  var presId = String(presentationId || '').trim();
  var slId = String(slideId || '').trim();
  if (!presId || !slId) {
    throw new Error('PresentationId e slideId são obrigatórios para exportar imagem do certificado.');
  }
  var url = 'https://docs.google.com/presentation/d/' + encodeURIComponent(presId) + '/export/png?pageid=' + encodeURIComponent(slId);
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  var status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Falha ao exportar slide em imagem. HTTP ' + status);
  }
  return response.getBlob().setContentType(MimeType.PNG);
}

function salvarImagemCertificadoNoDrive_(blob, nomeArquivo, pasta) {
  if (!blob) throw new Error('Blob da imagem do certificado é obrigatório.');
  if (!pasta) throw new Error('Pasta de destino da imagem do certificado é obrigatória.');
  var nome = String(nomeArquivo || 'certificado.png').trim();
  return pasta.createFile(blob.setName(nome));
}

function montarLinkWhatsAppCertificado_(nomePessoa, url) {
  var link = String(url || '').trim();
  if (!link) return '';
  var nome = String(nomePessoa || '').trim();
  var saudacao = nome ? ('Olá! Eu sou ' + nome + '.') : 'Olá!';
  var mensagem = '🚴‍♂️ Conquistei meu certificado de conclusão!\nConfira aqui:\n' + link + '\n\n' + saudacao;
  return 'https://wa.me/?text=' + encodeURIComponent(mensagem);
}

function montarUrlPublicaImagemCertificado_(fileId) {
  return 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(String(fileId || '').trim());
}

function montarUrlDownloadImagemCertificado_(fileId) {
  return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(String(fileId || '').trim());
}
