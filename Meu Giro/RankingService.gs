function getRanking(idDgmb, idDesafio, idItemEstoque, idInscricao) {
  var perfRankingInicio = typeof painelMG_perfNow_ === 'function' ? painelMG_perfNow_() : new Date().getTime();
  var rankingPerfLogado = false;
  if (typeof painelMG_incrementarAuditoriaCarregamentoInicial_ === 'function') {
    painelMG_incrementarAuditoriaCarregamentoInicial_('getRanking_chamadas');
  }
  try {
    var idUsuario = rankingMG_norm_(idDgmb);
    var desafioSolicitado = rankingMG_norm_(idDesafio);
    var itemSolicitado = rankingMG_norm_(idItemEstoque);
    var inscricaoSolicitada = rankingMG_norm_(idInscricao);
    var diagnosticos = rankingMG_criarDiagnosticos_(inscricaoSolicitada);

    if (!idUsuario) {
      return { ok: false, data: [], total: 0, msg: 'ID do usuário não informado.', diagnosticos: diagnosticos };
    }

    var pessoas = getAllObjects_(SHEETS.PESSOAS) || [];
    var resumo = getAllObjects_(SHEETS.MEU_GIRO_RESUMO) || [];

    if (!resumo.length) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Nenhum atleta encontrado no ranking.',
        diagnosticos: diagnosticos
      };
    }

    // Sem ID_INSCRICAO, o fallback legado não consegue confirmar um foco histórico.
    var statusLocalizacaoLegada = { ATIVO: true, CONCLUIDO: true };
    var referencia = null;

    if (inscricaoSolicitada) {
      referencia = rankingMG_localizarReferenciaPorInscricao_(resumo, idUsuario, inscricaoSolicitada);
      diagnosticos.referencia_por = 'ID_INSCRICAO';

      if (!referencia) {
        diagnosticos.id_inscricao_nao_encontrado = true;
        rankingMG_logDiagnostico_('ID_INSCRICAO recebido, mas não encontrado no MEU_GIRO_RESUMO.', {
          id_dgmb: idUsuario,
          id_inscricao: inscricaoSolicitada,
          id_desafio: desafioSolicitado,
          id_item_estoque: itemSolicitado
        });
        return {
          ok: true,
          data: [],
          total: 0,
          posicao_usuario: 0,
          msg: 'A inscrição em foco não foi encontrada para o atleta informado.',
          diagnosticos: diagnosticos
        };
      }
    } else {
      diagnosticos.referencia_por = 'LEGADO';
      referencia = rankingMG_localizarReferenciaLegada_(
        resumo,
        idUsuario,
        desafioSolicitado,
        itemSolicitado,
        statusLocalizacaoLegada
      );
    }

    if (!referencia) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Usuário sem desafio elegível para o ranking.',
        diagnosticos: diagnosticos
      };
    }

    var statusReferencia = rankingMG_obterStatus_(referencia);
    var statusValidos = rankingMG_obterStatusElegiveis_(statusReferencia);
    var idInscricaoReferencia = rankingMG_obterIdInscricao_(referencia);
    var desafioPrincipal = rankingMG_norm_(rankingMG_firstFilled_(referencia, ['ID_DESAFIO', 'id_desafio']));
    var itemPrincipal = rankingMG_norm_(rankingMG_firstFilled_(referencia, ['id_item_estoque', 'id item estoque']));
    diagnosticos.id_inscricao_referencia = idInscricaoReferencia;
    diagnosticos.status_referencia = statusReferencia;
    diagnosticos.status_elegiveis = rankingMG_chavesMapa_(statusValidos);
    diagnosticos.id_desafio_referencia = desafioPrincipal;

    if (!desafioPrincipal) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Desafio-base não identificado para o ranking.',
        diagnosticos: diagnosticos
      };
    }

    var pessoasMap = rankingMG_buildPessoasMap_(pessoas);
    var selecao;

    if (inscricaoSolicitada) {
      var indiceCompetitivo = rankingMG_criarIndiceCompetitivo_();
      var atributosReferencia = rankingMG_resolverAtributosCompetitivos_(referencia, indiceCompetitivo);

      if (!atributosReferencia) {
        diagnosticos.atributos_competitivos_nao_resolvidos = true;
        rankingMG_logDiagnostico_('Atributos competitivos da inscrição em foco não foram resolvidos.', {
          id_dgmb: idUsuario,
          id_inscricao: inscricaoSolicitada
        });
        return {
          ok: true,
          data: [],
          total: 0,
          posicao_usuario: 0,
          msg: 'Não foi possível resolver os atributos competitivos da inscrição em foco.',
          diagnosticos: diagnosticos
        };
      }

      var validacaoReferencia = rankingMG_validarAtributosCompetitivos_(atributosReferencia);
      if (!validacaoReferencia.valido) {
        diagnosticos.atributos_competitivos_invalidos_referencia = validacaoReferencia.campos_invalidos;
        rankingMG_logDiagnostico_('Inscrição em foco excluída por atributos competitivos inválidos.', {
          id_dgmb: idUsuario,
          id_inscricao: inscricaoSolicitada,
          campos_invalidos: validacaoReferencia.campos_invalidos
        });
        return {
          ok: true,
          data: [],
          total: 0,
          posicao_usuario: 0,
          msg: 'A inscrição em foco não possui todos os atributos competitivos válidos.',
          diagnosticos: diagnosticos
        };
      }

      diagnosticos.chave_competitiva = rankingMG_montarChaveCompetitiva_(atributosReferencia);
      diagnosticos.atributos_competitivos_referencia = atributosReferencia;
      selecao = rankingMG_selecionarGrupoCompetitivo_(
        resumo,
        atributosReferencia,
        indiceCompetitivo,
        statusValidos,
        diagnosticos
      );
    } else {
      diagnosticos.grupo_base_referencia = rankingMG_extrairGrupoBaseDesafio_(itemPrincipal);
      selecao = {
        rows: rankingMG_selecionarGrupoLegado_(resumo, desafioPrincipal, itemPrincipal, statusValidos),
        deduplicado: false
      };
      diagnosticos.quantidade_antes_deduplicacao = selecao.rows.length;
      diagnosticos.quantidade_depois_deduplicacao = selecao.rows.length;
      rankingMG_atualizarDiagnosticosLegados_(diagnosticos, selecao.rows);
    }

    var ranking = [];
    for (var i = 0; i < selecao.rows.length; i++) {
      ranking.push(rankingMG_montarParticipante_(selecao.rows[i], pessoasMap));
    }

    ranking.sort(rankingMG_compararRanking_);

    var posicaoUsuario = 0;
    for (var p = 0; p < ranking.length; p++) {
      ranking[p].posicao = p + 1;
      ranking[p].posicao_ranking = p + 1;
      if (!posicaoUsuario && ranking[p].id_dgmb === idUsuario) posicaoUsuario = p + 1;
    }

    if (typeof painelMG_perfLog_ === 'function') {
      rankingPerfLogado = true;
      painelMG_perfLog_('painel-inicial', 'getRanking', perfRankingInicio, {
        id_desafio: desafioSolicitado || desafioPrincipal || '',
        id_item_estoque: itemSolicitado || itemPrincipal || '',
        id_inscricao: inscricaoSolicitada || idInscricaoReferencia || '',
        total_participantes: ranking.length,
        posicao_usuario: posicaoUsuario
      });
    }

    return {
      ok: true,
      data: ranking,
      total: ranking.length,
      posicao_usuario: posicaoUsuario,
      id_inscricao_referencia: idInscricaoReferencia,
      diagnosticos: diagnosticos
    };
  } catch (err) {
    if (typeof painelMG_perfLog_ === 'function') {
      rankingPerfLogado = true;
      painelMG_perfLog_('painel-inicial', 'getRanking', perfRankingInicio, {
        ok: false,
        erro: err && err.message ? err.message : 'Erro ao carregar ranking.'
      });
    }
    return {
      ok: false,
      data: [],
      total: 0,
      msg: err && err.message ? err.message : 'Erro ao carregar ranking.'
    };
  } finally {
    if (!rankingPerfLogado && typeof painelMG_perfLog_ === 'function') {
      painelMG_perfLog_('painel-inicial', 'getRanking', perfRankingInicio, {
        ok: false,
        retorno_antecipado: true
      });
    }
  }
}

function rankingMG_localizarReferenciaPorInscricao_(resumo, idDgmb, idInscricao) {
  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    if (rankingMG_obterIdInscricao_(row) !== idInscricao) continue;
    if (rankingMG_obterIdDgmb_(row) === idDgmb) return row;
  }
  return null;
}

function rankingMG_localizarReferenciaLegada_(resumo, idDgmb, idDesafio, idItemEstoque, statusValidos) {
  var referencia = null;
  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    if (rankingMG_obterIdDgmb_(row) !== idDgmb) continue;

    var status = rankingMG_obterStatus_(row);
    var desafio = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DESAFIO', 'id_desafio']));
    var item = rankingMG_norm_(rankingMG_firstFilled_(row, ['id_item_estoque', 'id item estoque']));
    if (!statusValidos[status]) continue;
    if (idDesafio && desafio !== idDesafio) continue;
    if (idDesafio && idItemEstoque && item !== idItemEstoque) continue;

    if (status === 'ATIVO') return row;
    if (!referencia) referencia = row;
  }
  return referencia;
}

// ID_INSCRICAO identifica a linha, mas não participa da chave que reúne atletas equivalentes.
function rankingMG_montarChaveCompetitiva_(atributos) {
  var src = atributos || {};
  if (!rankingMG_validarAtributosCompetitivos_(src).valido) return '';
  return [
    'COMPETICAO_V1',
    rankingMG_norm_(src.id_desafio),
    rankingMG_normalizarTipo_(src.tipo_desafio),
    rankingMG_normalizarDataCompetitiva_(src.periodo_inicio),
    rankingMG_normalizarDataCompetitiva_(src.periodo_fim),
    rankingMG_formatarMetaCompetitiva_(src.meta_km),
    rankingMG_obterItemCompetitivo_(src)
  ].join('|');
}

function rankingMG_criarIndiceCompetitivo_() {
  if (typeof painelMG_incrementarAuditoriaCarregamentoInicial_ === 'function') {
    painelMG_incrementarAuditoriaCarregamentoInicial_('leituras_dgmbDesafios');
  }
  var rows = getAllObjects_(SHEETS.DESAFIO) || [];
  var periodos = buildPeriodoOficialPorAbaEId_(getSpreadsheet_());
  var porInscricao = {};

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var idInscricao = rankingMG_obterIdInscricao_(row);
    var idDgmb = rankingMG_obterIdDgmb_(row);
    if (!idInscricao || !idDgmb) continue;

    var idDesafio = rankingMG_norm_(rankingMG_firstFilled_(row, [
      'ID_DESAFIO', 'id_desafio', 'ID Desafio', 'id desafio'
    ]));
    if (!idDesafio) {
      idDesafio = rankingMG_norm_(extrairIdDesafioObservacao_(rankingMG_firstFilled_(row, [
        'Observacao', 'observacao', 'Observação', 'observação'
      ])));
    }
    var tipo = rankingMG_norm_(rankingMG_firstFilled_(row, [
      'tipo_do_desafio', 'Tipo_Do_Desafio', 'tipo_desafio', 'Tipo_Desafio', 'tipo desafio'
    ]));
    var tipoNormalizado = rankingMG_normalizarTipo_(tipo).toLowerCase();
    var ehNormal = tipoNormalizado === 'normal';
    var abaDesafio = SHEETS.DESAFIO || 'dgmbDesafios';
    var periodoLista = (idDesafio && periodos.byId[idDesafio]) ||
      (!ehNormal && periodos.byAba[abaDesafio]) ||
      { inicio: '', fim: '', nome_desafio: '' };
    var periodo = rankingMG_resolverPeriodoCompetitivo_(row, periodoLista);

    porInscricao[rankingMG_chaveInscricao_(idDgmb, idInscricao)] = {
      id_inscricao: idInscricao,
      id_dgmb: idDgmb,
      id_desafio: idDesafio,
      tipo_desafio: tipo,
      periodo_inicio: periodo.inicio,
      periodo_fim: periodo.fim,
      meta_km: rankingMG_toNumber_(rankingMG_firstFilled_(row, [
        'Distancia_KM', 'distancia_km', 'Distancia KM', 'Meta_KM', 'meta_km', 'meta'
      ])),
      id_item_estoque: rankingMG_norm_(rankingMG_firstFilled_(row, [
        'id_item_estoque', 'ID_ITEM_ESTOQUE', 'id item estoque'
      ]))
    };
  }

  return { por_inscricao: porInscricao };
}

function rankingMG_resolverPeriodoCompetitivo_(row, periodoLista) {
  var periodoDatas = {
    inicio: normalizarDataISO_(rankingMG_firstFilled_(row, [
      'data_inicio_desafio', 'Data_Inicio_Desafio', 'data inicio desafio', 'data início desafio'
    ])),
    fim: normalizarDataISO_(rankingMG_firstFilled_(row, [
      'data_fim_desafio', 'Data_Fim_Desafio', 'data fim desafio'
    ]))
  };
  if (periodoCompletoValido_(periodoDatas)) return periodoDatas;

  var periodoTexto = rankingMG_firstFilled_(row, [
    'periodo_desafio', 'Periodo_Desafio', 'período_desafio', 'periodo desafio', 'período desafio'
  ]);
  var periodoHistorico = extrairPeriodoDesafioTexto_(periodoTexto);
  if (periodoCompletoValido_(periodoHistorico)) return periodoHistorico;
  if (periodoCompletoValido_(periodoLista)) {
    return { inicio: periodoLista.inicio, fim: periodoLista.fim };
  }
  return { inicio: '', fim: '' };
}

function rankingMG_resolverAtributosCompetitivos_(rowResumo, indice) {
  var idInscricao = rankingMG_obterIdInscricao_(rowResumo);
  var idDgmb = rankingMG_obterIdDgmb_(rowResumo);
  if (!idInscricao || !idDgmb) return null;

  var vinculo = indice.por_inscricao[rankingMG_chaveInscricao_(idDgmb, idInscricao)];
  if (!vinculo) return null;

  return {
    id_inscricao: idInscricao,
    id_dgmb: idDgmb,
    id_desafio: rankingMG_norm_(rankingMG_firstFilled_(rowResumo, ['ID_DESAFIO', 'id_desafio'])) || vinculo.id_desafio,
    tipo_desafio: vinculo.tipo_desafio,
    periodo_inicio: vinculo.periodo_inicio,
    periodo_fim: vinculo.periodo_fim,
    meta_km: rankingMG_obterMeta_(rowResumo),
    id_item_estoque: rankingMG_norm_(rankingMG_firstFilled_(rowResumo, [
      'id_item_estoque', 'ID_ITEM_ESTOQUE', 'id item estoque'
    ])) || vinculo.id_item_estoque
  };
}

function rankingMG_selecionarGrupoCompetitivo_(resumo, referencia, indice, statusValidos, diagnosticos) {
  var chaveReferencia = rankingMG_montarChaveCompetitiva_(referencia);
  var candidatos = [];
  var diferencas = {
    metas: {},
    tipos: {},
    periodos: {},
    itens: {}
  };

  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    if (!statusValidos[rankingMG_obterStatus_(row)]) continue;

    var atributos = rankingMG_resolverAtributosCompetitivos_(row, indice);
    if (!atributos) {
      if (rankingMG_obterIdInscricao_(row)) diagnosticos.candidatos_sem_atributos_competitivos++;
      continue;
    }

    var validacaoAtributos = rankingMG_validarAtributosCompetitivos_(atributos);
    if (!validacaoAtributos.valido) {
      diagnosticos.inscricoes_excluidas_atributos_invalidos.push({
        id_inscricao: rankingMG_obterIdInscricao_(row),
        id_dgmb: rankingMG_obterIdDgmb_(row),
        campos_invalidos: validacaoAtributos.campos_invalidos
      });
      continue;
    }
    if (atributos.id_desafio !== referencia.id_desafio) continue;

    rankingMG_registrarDiferencasCompetitivas_(diferencas, referencia, atributos);
    if (rankingMG_montarChaveCompetitiva_(atributos) !== chaveReferencia) continue;
    candidatos.push(row);
  }

  diagnosticos.quantidade_antes_deduplicacao = candidatos.length;
  diagnosticos.quantidade_inscricoes_excluidas_atributos_invalidos =
    diagnosticos.inscricoes_excluidas_atributos_invalidos.length;
  diagnosticos.metas_diferentes_detectadas = rankingMG_valoresMapa_(diferencas.metas, true);
  diagnosticos.tipos_diferentes_detectados = rankingMG_valoresMapa_(diferencas.tipos, false);
  diagnosticos.periodos_diferentes_detectados = rankingMG_valoresMapa_(diferencas.periodos, false);
  diagnosticos.itens_diferentes_detectados = rankingMG_valoresMapa_(diferencas.itens, false);

  var porAtleta = {};
  for (var c = 0; c < candidatos.length; c++) {
    var candidato = candidatos[c];
    var idDgmb = rankingMG_obterIdDgmb_(candidato);
    if (!idDgmb) continue;

    if (!porAtleta[idDgmb]) {
      porAtleta[idDgmb] = { canonica: candidato, removidas: [] };
      continue;
    }

    var atual = porAtleta[idDgmb].canonica;
    if (rankingMG_compararInscricaoCanonica_(candidato, atual) < 0) {
      porAtleta[idDgmb].removidas.push(atual);
      porAtleta[idDgmb].canonica = candidato;
    } else {
      porAtleta[idDgmb].removidas.push(candidato);
    }
  }

  var deduplicados = [];
  for (var id in porAtleta) {
    if (!Object.prototype.hasOwnProperty.call(porAtleta, id)) continue;
    var grupoAtleta = porAtleta[id];
    deduplicados.push(grupoAtleta.canonica);
    if (grupoAtleta.removidas.length) {
      diagnosticos.atletas_duplicados_removidos.push({
        id_dgmb: id,
        quantidade_removida: grupoAtleta.removidas.length,
        id_inscricao_canonica: rankingMG_obterIdInscricao_(grupoAtleta.canonica),
        ids_inscricao_removidos: grupoAtleta.removidas.map(function(item) {
          return rankingMG_obterIdInscricao_(item);
        })
      });
    }
  }

  diagnosticos.atletas_duplicados_removidos.sort(function(a, b) {
    return String(a.id_dgmb).localeCompare(String(b.id_dgmb));
  });
  diagnosticos.quantidade_depois_deduplicacao = deduplicados.length;
  diagnosticos.possui_dgmb_duplicado = diagnosticos.atletas_duplicados_removidos.length > 0;
  diagnosticos.dgmb_duplicados = diagnosticos.atletas_duplicados_removidos;

  if (diagnosticos.possui_dgmb_duplicado) {
    rankingMG_logDiagnostico_('Inscrições duplicadas do mesmo atleta foram removidas do grupo competitivo.', {
      chave_competitiva: chaveReferencia,
      atletas: diagnosticos.atletas_duplicados_removidos
    });
  }

  return { rows: deduplicados, deduplicado: true };
}


function rankingMG_validarAtributosCompetitivos_(atributos) {
  var src = atributos || {};
  var camposInvalidos = [];
  var inicio = rankingMG_normalizarDataCompetitiva_(src.periodo_inicio);
  var fim = rankingMG_normalizarDataCompetitiva_(src.periodo_fim);

  if (!rankingMG_norm_(src.id_desafio)) camposInvalidos.push('ID_DESAFIO');
  if (!rankingMG_normalizarTipo_(src.tipo_desafio)) camposInvalidos.push('TIPO_DESAFIO');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(inicio)) camposInvalidos.push('PERIODO_INICIO');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fim) || (inicio && fim && inicio > fim)) {
    camposInvalidos.push('PERIODO_FIM');
  }
  if (!(rankingMG_toNumber_(src.meta_km) > 0)) camposInvalidos.push('META_KM');
  if (!rankingMG_tipoNormal_(src.tipo_desafio) && !rankingMG_norm_(src.id_item_estoque)) {
    camposInvalidos.push('ID_ITEM_ESTOQUE');
  }

  return {
    valido: camposInvalidos.length === 0,
    campos_invalidos: camposInvalidos
  };
}

function rankingMG_registrarDiferencasCompetitivas_(mapas, referencia, candidato) {
  var metaReferencia = rankingMG_formatarMetaCompetitiva_(referencia.meta_km);
  var metaCandidato = rankingMG_formatarMetaCompetitiva_(candidato.meta_km);
  var tipoReferencia = rankingMG_normalizarTipo_(referencia.tipo_desafio);
  var tipoCandidato = rankingMG_normalizarTipo_(candidato.tipo_desafio);
  var periodoReferencia = rankingMG_normalizarDataCompetitiva_(referencia.periodo_inicio) + '|' +
    rankingMG_normalizarDataCompetitiva_(referencia.periodo_fim);
  var periodoCandidato = rankingMG_normalizarDataCompetitiva_(candidato.periodo_inicio) + '|' +
    rankingMG_normalizarDataCompetitiva_(candidato.periodo_fim);
  var itemReferencia = rankingMG_obterItemCompetitivo_(referencia);
  var itemCandidato = rankingMG_obterItemCompetitivo_(candidato);

  if (metaCandidato !== metaReferencia) mapas.metas[metaCandidato] = rankingMG_toNumber_(candidato.meta_km);
  if (tipoCandidato !== tipoReferencia) mapas.tipos[tipoCandidato] = tipoCandidato;
  if (periodoCandidato !== periodoReferencia) mapas.periodos[periodoCandidato] = {
    periodo_inicio: rankingMG_normalizarDataCompetitiva_(candidato.periodo_inicio),
    periodo_fim: rankingMG_normalizarDataCompetitiva_(candidato.periodo_fim)
  };
  if (itemCandidato !== itemReferencia) mapas.itens[itemCandidato] = itemCandidato;
}

function rankingMG_selecionarGrupoLegado_(resumo, idDesafio, idItemEstoque, statusValidos) {
  var grupoBase = rankingMG_extrairGrupoBaseDesafio_(idItemEstoque);
  var rows = [];
  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    var desafio = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DESAFIO', 'id_desafio']));
    var item = rankingMG_norm_(rankingMG_firstFilled_(row, ['id_item_estoque', 'id item estoque']));
    if (desafio !== idDesafio) continue;
    if (grupoBase && rankingMG_extrairGrupoBaseDesafio_(item) !== grupoBase) continue;
    if (!statusValidos[rankingMG_obterStatus_(row)]) continue;
    rows.push(row);
  }
  return rows;
}

function rankingMG_atualizarDiagnosticosLegados_(diagnosticos, rows) {
  var ocorrencias = {};
  var metas = {};
  for (var i = 0; i < rows.length; i++) {
    var idDgmb = rankingMG_obterIdDgmb_(rows[i]);
    if (idDgmb) ocorrencias[idDgmb] = (ocorrencias[idDgmb] || 0) + 1;
    var meta = rankingMG_obterMeta_(rows[i]);
    metas[String(meta)] = meta;
  }

  var duplicados = [];
  for (var id in ocorrencias) {
    if (Object.prototype.hasOwnProperty.call(ocorrencias, id) && ocorrencias[id] > 1) {
      duplicados.push({ id_dgmb: id, quantidade: ocorrencias[id] });
    }
  }
  duplicados.sort(function(a, b) { return String(a.id_dgmb).localeCompare(String(b.id_dgmb)); });

  diagnosticos.dgmb_duplicados = duplicados;
  diagnosticos.possui_dgmb_duplicado = duplicados.length > 0;
  diagnosticos.metas_distintas = rankingMG_valoresMapa_(metas, true);
  diagnosticos.possui_multiplas_metas = diagnosticos.metas_distintas.length > 1;
  diagnosticos.metas_diferentes_no_ranking = diagnosticos.possui_multiplas_metas;
  diagnosticos.grupo_competitivo_multiplas_metas = diagnosticos.possui_multiplas_metas;
}

// Ordenação exclusiva para escolher uma inscrição canônica por atleta dentro do grupo.
function rankingMG_compararInscricaoCanonica_(a, b) {
  var prioridadeStatus = { CONCLUIDO: 2, ATIVO: 1 };
  var statusA = prioridadeStatus[rankingMG_obterStatus_(a)] || 0;
  var statusB = prioridadeStatus[rankingMG_obterStatus_(b)] || 0;
  if (statusA !== statusB) return statusB - statusA;

  var distanciaA = rankingMG_obterDistancia_(a);
  var distanciaB = rankingMG_obterDistancia_(b);
  if (distanciaA !== distanciaB) return distanciaB - distanciaA;

  var percentualA = rankingMG_obterPercentual_(a);
  var percentualB = rankingMG_obterPercentual_(b);
  if (percentualA !== percentualB) return percentualB - percentualA;

  return rankingMG_obterIdInscricao_(a).localeCompare(rankingMG_obterIdInscricao_(b));
}

function rankingMG_montarParticipante_(row, pessoasMap) {
  var idDgmb = rankingMG_obterIdDgmb_(row);
  var meta = rankingMG_obterMeta_(row);
  var realizado = rankingMG_obterDistancia_(row);
  var percentual = rankingMG_obterPercentual_(row);
  if (percentual <= 0 && meta > 0) percentual = rankingMG_round1_((realizado / meta) * 100);
  var pessoa = pessoasMap[idDgmb] || {};

  return {
    id_inscricao: rankingMG_obterIdInscricao_(row),
    id_dgmb: idDgmb,
    nome: pessoa.nome || rankingMG_norm_(rankingMG_firstFilled_(row, ['Nome_Avatar', 'nome_avatar'])) || 'Participante',
    cidade_uf: pessoa.cidade_uf || '',
    distancia_realizada: realizado,
    percentual_concluido: percentual,
    meta: meta,
    realizado: realizado,
    restante: rankingMG_round1_(Math.max(meta - realizado, 0)),
    percentual: percentual
  };
}

function rankingMG_compararRanking_(a, b) {
  if (b.distancia_realizada !== a.distancia_realizada) return b.distancia_realizada - a.distancia_realizada;
  if (b.percentual_concluido !== a.percentual_concluido) return b.percentual_concluido - a.percentual_concluido;
  return String(a.id_dgmb || '').localeCompare(String(b.id_dgmb || ''));
}

function rankingMG_obterIdInscricao_(row) {
  return rankingMG_norm_(rankingMG_firstFilled_(row || {}, [
    'ID_INSCRICAO', 'id_inscricao', 'ID Inscricao', 'ID Inscrição', 'id inscricao', 'id inscrição'
  ]));
}

function rankingMG_obterIdDgmb_(row) {
  return rankingMG_norm_(rankingMG_firstFilled_(row || {}, ['ID_DGMB', 'id_dgmb']));
}

function rankingMG_obterStatus_(row) {
  return rankingMG_norm_(rankingMG_firstFilled_(row || {}, ['Status_Apuracao', 'status_apuracao'])).toUpperCase();
}

function rankingMG_obterMeta_(row) {
  return rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row || {}, [
    'Meta_KM', 'meta_km', 'Distancia_KM', 'distancia_km', 'Distancia KM', 'meta'
  ])));
}

function rankingMG_obterDistancia_(row) {
  return rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row || {}, [
    'Distancia_Realizada', 'distancia_realizada', 'Distancia Realizada'
  ])));
}

function rankingMG_obterPercentual_(row) {
  return rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row || {}, [
    'Percentual_Concluido', 'percentual_concluido'
  ])));
}

function rankingMG_chaveInscricao_(idDgmb, idInscricao) {
  return rankingMG_norm_(idDgmb) + '|' + rankingMG_norm_(idInscricao);
}

function rankingMG_normalizarTipo_(value) {
  var tipo = rankingMG_norm_(value).toUpperCase();
  return tipo.normalize ? tipo.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : tipo;
}

function rankingMG_tipoNormal_(value) {
  return rankingMG_normalizarTipo_(value) === 'NORMAL';
}

function rankingMG_obterItemCompetitivo_(atributos) {
  var src = atributos || {};
  return rankingMG_tipoNormal_(src.tipo_desafio) ? '' : rankingMG_norm_(src.id_item_estoque);
}

function rankingMG_obterStatusElegiveis_(statusReferencia) {
  var status = rankingMG_norm_(statusReferencia).toUpperCase();
  if (status === 'EXPIRADO') return { EXPIRADO: true, CONCLUIDO: true };
  return { ATIVO: true, CONCLUIDO: true };
}

function rankingMG_normalizarDataCompetitiva_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    var ano = value.getFullYear();
    var mes = String(value.getMonth() + 1).padStart(2, '0');
    var dia = String(value.getDate()).padStart(2, '0');
    return ano + '-' + mes + '-' + dia;
  }

  var texto = rankingMG_norm_(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
  if (/^\d{2}\/\d{2}\/\d{4}/.test(texto)) {
    return texto.slice(6, 10) + '-' + texto.slice(3, 5) + '-' + texto.slice(0, 2);
  }
  return texto;
}

function rankingMG_formatarMetaCompetitiva_(value) {
  return rankingMG_round1_(rankingMG_toNumber_(value)).toFixed(1);
}

function rankingMG_chavesMapa_(map) {
  var out = [];
  for (var key in map) {
    if (Object.prototype.hasOwnProperty.call(map, key)) out.push(key);
  }
  return out.sort();
}

function rankingMG_valoresMapa_(map, numerico) {
  var out = [];
  for (var key in map) {
    if (Object.prototype.hasOwnProperty.call(map, key)) out.push(map[key]);
  }
  out.sort(function(a, b) {
    return numerico ? Number(a) - Number(b) : JSON.stringify(a).localeCompare(JSON.stringify(b));
  });
  return out;
}

function rankingMG_criarDiagnosticos_(idInscricaoSolicitada) {
  return {
    id_inscricao_recebido: rankingMG_norm_(idInscricaoSolicitada),
    id_inscricao_nao_encontrado: false,
    atributos_competitivos_nao_resolvidos: false,
    atributos_competitivos_invalidos_referencia: [],
    referencia_por: '',
    id_inscricao_referencia: '',
    id_desafio_referencia: '',
    status_referencia: '',
    status_elegiveis: [],
    grupo_base_referencia: '',
    chave_competitiva: '',
    atributos_competitivos_referencia: null,
    quantidade_antes_deduplicacao: 0,
    quantidade_depois_deduplicacao: 0,
    atletas_duplicados_removidos: [],
    dgmb_duplicados: [],
    possui_dgmb_duplicado: false,
    candidatos_sem_atributos_competitivos: 0,
    quantidade_inscricoes_excluidas_atributos_invalidos: 0,
    inscricoes_excluidas_atributos_invalidos: [],
    metas_diferentes_detectadas: [],
    tipos_diferentes_detectados: [],
    periodos_diferentes_detectados: [],
    itens_diferentes_detectados: [],
    metas_distintas: [],
    possui_multiplas_metas: false,
    metas_diferentes_no_ranking: false,
    grupo_competitivo_multiplas_metas: false
  };
}

function rankingMG_logDiagnostico_(mensagem, dados) {
  try {
    Logger.log('[Meu Giro][Ranking] ' + mensagem + ' ' + JSON.stringify(dados || {}));
  } catch (e) {}
}

function rankingMG_buildPessoasMap_(pessoas) {
  var map = {};
  for (var i = 0; i < pessoas.length; i++) {
    var row = pessoas[i] || {};
    var idDgmb = rankingMG_obterIdDgmb_(row);
    if (!idDgmb) continue;
    map[idDgmb] = {
      id_dgmb: idDgmb,
      nome: rankingMG_norm_(rankingMG_firstFilled_(row, ['nome', 'Nome'])),
      cidade_uf: rankingMG_norm_(rankingMG_firstFilled_(row, ['Cidade-UF', 'Cidade_UF', 'cidade_uf', 'cidade-uf']))
    };
  }
  return map;
}

function rankingMG_firstFilled_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, keys[i])) {
      var value = obj[keys[i]];
      if (value !== '' && value !== null && value !== undefined) return value;
    }
  }
  return '';
}

function rankingMG_norm_(value) {
  return String(value || '').trim();
}

function rankingMG_toNumber_(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  var s = String(value == null ? '' : value).trim();
  if (!s) return 0;
  s = s.replace(/\s+/g, '');
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  }
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function rankingMG_round1_(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 10) / 10;
}

function rankingMG_extrairGrupoBaseDesafio_(idItemEstoque) {
  var item = rankingMG_norm_(idItemEstoque);
  if (!item) return '';
  var semKm = item.replace(/_[0-9]+(?:[.,][0-9]+)?$/g, '');
  return rankingMG_norm_(semKm || item);
}
