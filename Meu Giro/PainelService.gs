
function painelMG_logBug03ListaDesafios_(etapa, lista, origem) {
  if (typeof bug03PeriodoDesafioLogBackend_ !== 'function') return;

  (Array.isArray(lista) ? lista : []).forEach(function(item) {
    bug03PeriodoDesafioLogBackend_(etapa, {
      id_dgmb: item && item.id_dgmb,
      id_desafio: item && item.id_desafio,
      id_inscricao: item && item.id_inscricao,
      id_item_estoque: item && item.id_item_estoque,
      nome_desafio: item && item.nome_desafio,
      periodo_desafio: item && item.periodo_desafio,
      periodo_inicio: item && item.periodo_inicio,
      periodo_fim: item && item.periodo_fim,
      origem: origem
    });
  });
}

function painelMG_logBug03Desafio_(etapa, item, origem) {
  if (typeof bug03PeriodoDesafioLogBackend_ !== 'function') return;
  if (!item) {
    bug03PeriodoDesafioLogBackend_(etapa, { origem: origem });
    return;
  }

  painelMG_logBug03ListaDesafios_(etapa, [item], origem);
}

function getPainelUsuario(idDgmb) {
  var opcoes = arguments.length > 1 ? arguments[1] : null;
  var somenteLeitura = !!(opcoes && opcoes.somenteLeitura);
  var perfTotalInicio = painelMG_perfNow_();
  var perfEtapaInicio = perfTotalInicio;
  var auditoria = painelMG_criarAuditoriaCarregamentoInicial_();
  painelMG_definirAuditoriaCarregamentoInicial_(auditoria);

  try {
    var id = String(idDgmb || '').trim();
    if (!id) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do usuário não informado.' };
    }

    perfEtapaInicio = painelMG_perfNow_();
    var pessoa = buscarPessoaPainelMG_(id);
    painelMG_perfLog_('painel-inicial', 'buscarPessoaPainelMG_', perfEtapaInicio, { encontrado: !!pessoa });

    perfEtapaInicio = painelMG_perfNow_();
    var resumoDesafios = obterMeuGiroResumoAtualizadoLeve_(id, { reconciliar: !somenteLeitura }) || [];
    painelMG_perfLog_('painel-inicial', 'lerMeuGiroResumoAtualizadoLogin_', perfEtapaInicio, {
      total_desafios_resumo: resumoDesafios.length,
      fallback: false
    });

    if (!resumoDesafios.length && !somenteLeitura) {
      perfEtapaInicio = painelMG_perfNow_();
      resumoDesafios = atualizarMeuGiroResumo_(id) || [];
      painelMG_perfLog_('painel-inicial', 'atualizarMeuGiroResumo_fallback_login_', perfEtapaInicio, {
        total_desafios_resumo: resumoDesafios.length,
        fallback: true
      });
    }

    perfEtapaInicio = painelMG_perfNow_();
    var desafio = buscarInscricaoPainelMGLeve_(id, resumoDesafios);
    painelMG_perfLog_('painel-inicial', desafio && desafio.usou_fallback ? 'buscarInscricaoPainelMG_' : 'buscarInscricaoPainelMGLeve_', perfEtapaInicio, {
      ok: !!(desafio && desafio.ok),
      total_desafios: desafio && desafio.desafios ? desafio.desafios.length : 0,
      fallback: !!(desafio && desafio.usou_fallback),
      motivo_fallback: desafio && desafio.motivo_fallback ? desafio.motivo_fallback : ''
    });

    if (!pessoa) {
      return { ok: false, code: 'USUARIO_NAO_ENCONTRADO', msg: 'Usuário não encontrado na base de pessoas para carregar o painel.' };
    }

    if (!desafio.ok) {
      return {
        ok: false,
        code: desafio.code,
        motivo_inscricao: desafio.motivo,
        msg: desafio.msg
      };
    }

    if (!desafio.data) {
      return {
        ok: false,
        code: 'CONTRATO_INSCRICAO_INVALIDO',
        msg: 'Dados de inscrição inválidos para o painel.'
      };
    }

    var desafioData = desafio.data;
    var desafiosConsolidados = (desafio.desafios || []).map(function(item) {
      var desafioPainel = {};
      Object.keys(item || {}).forEach(function(chave) {
        desafioPainel[chave] = item[chave];
      });

      var operacional = painelMG_montarMensagemOperacional_(
        desafioPainel.status_apuracao,
        desafioPainel.status_usuario_desafio
      );
      desafioPainel.status_operacional = operacional.codigo_operacional;
      desafioPainel.mensagem_operacional = operacional.mensagem_operacional;
      return desafioPainel;
    });

    var desafiosAtivosPainel = desafiosConsolidados.filter(function(d) {
      return painelMG_isDesafioAtivoParaFoco_(d);
    }).sort(painelMG_compareFocoDesafiosAtivos_);
    var desafiosHistoricoPainel = desafiosConsolidados.filter(function(d) {
      return painelMG_isStatusHistorico_(d && d.status_apuracao);
    }).sort(painelMG_compareHistoricoDesafios_);

    var desafioPrincipalPainel = painelMG_selecionarDesafioPrincipal_(desafiosConsolidados);
    var meta = painelMG_toNumber_(desafioPrincipalPainel ? desafioPrincipalPainel.meta_km : desafioData.meta);
    var realizado = painelMG_toNumber_(desafioPrincipalPainel ? desafioPrincipalPainel.distancia_realizada : desafioData.realizado);
    var realizadoPainel = painelMG_round1_(realizado);

    var statusOperacionalPainel = desafioPrincipalPainel
      ? {
          codigo_operacional: desafioPrincipalPainel.status_operacional,
          mensagem_operacional: desafioPrincipalPainel.mensagem_operacional
        }
      : painelMG_montarMensagemOperacional_('', '');
    var progresso = painelMG_calcularProgresso_(meta, realizadoPainel);
    var ritmo = painelMG_calcularRitmo_(meta, realizadoPainel, desafioData.periodo_inicio, desafioData.periodo_fim);
    perfEtapaInicio = painelMG_perfNow_();
    var atividades = buscarAtividadesUsuario_(id);
    painelMG_perfLog_('painel-inicial', 'buscarAtividadesUsuario_', perfEtapaInicio, {
      total_atividades: atividades && atividades.length ? atividades.length : 0
    });

    // P23: ranking sob demanda. O carregamento inicial do painel não deve
    // acionar getRanking(), nem os agregadores que chamam getRanking() para
    // cada desafio. A tela Ranking mantém o fluxo próprio via carregarRanking().
    var rankingInfo = { posicao: 0, total: 0 };
    var rankingPorDesafio = {};

    var frase = '';
    var contextoFrase = '';

    try {
      var painelBaseFrase = painelMG_montarBaseFrase_({
        id_dgmb: pessoa.id_dgmb || id,
        nome: pessoa.nome || '',
        meta: meta,
        realizado: realizadoPainel,
        restante: progresso.restante,
        percentual: progresso.percentual,
        ritmo_status: ritmo.status,
        ritmo_mensagem: ritmo.mensagem
      });

      var payloadFrase = montarPayloadFraseMotivacional_(painelBaseFrase);
      frase = payloadFrase && payloadFrase.frase_motivacional ? payloadFrase.frase_motivacional : '';
      contextoFrase = payloadFrase && payloadFrase.contexto_frase ? payloadFrase.contexto_frase : '';
    } catch (e) {
      frase = '';
      contextoFrase = '';
    }

    painelMG_logResumoAuditoriaCarregamentoInicial_(perfTotalInicio, auditoria, {
      ok: true,
      total_desafios: desafiosConsolidados.length,
      total_atividades: atividades && atividades.length ? atividades.length : 0,
      total_rankings_por_desafio: rankingPorDesafio ? Object.keys(rankingPorDesafio).length : 0
    });

    painelMG_logBug03Desafio_('getPainelUsuario/desafio_em_foco', desafioPrincipalPainel, 'payload.data.desafio_em_foco');
    painelMG_logBug03ListaDesafios_('getPainelUsuario/desafios_ativos', desafiosAtivosPainel, 'payload.data.desafios_ativos');
    painelMG_logBug03ListaDesafios_('getPainelUsuario/desafios_historico', desafiosHistoricoPainel, 'payload.data.desafios_historico');
    return {
      ok: true,
      data: {
        // Compatibilidade temporária de contrato entre payloads do painel.
        // Manter aliases duplicados até convergência explícita dos consumidores.
        nome: pessoa.nome || '',
        cidade_uf: pessoa.cidade_uf || '',
        id_dgmb: pessoa.id_dgmb || '',
        status_inscricao: desafioData.status_inscricao || 'inscrito',
        criterio_validacao_inscricao: desafioData.criterio_validacao || 'presenca_id_dgmb',
        desafio_usuario: desafioData.aba_desafio || '',

        meta: painelMG_round1_(meta),
        realizado: realizadoPainel,
        restante: progresso.restante,
        percentual: progresso.percentual,

        diaAtual: ritmo.diaAtual,
        diasRestantes: ritmo.diasRestantes,
        kmIdealAtual: ritmo.kmIdealHoje,
        kmPorDiaRestante: ritmo.kmPorDiaRestante,

        ritmo_status: ritmo.status,
        ritmo_mensagem: ritmo.mensagem,
        status_operacional: statusOperacionalPainel.codigo_operacional,
        mensagem_operacional: statusOperacionalPainel.mensagem_operacional,

        posicao_ranking: rankingInfo.posicao,
        total_participantes: rankingInfo.total,
        posicaoRanking: rankingInfo.posicao,
        totalParticipantes: rankingInfo.total,
        ranking_por_desafio: rankingPorDesafio,

        frase: frase || desafioData.frase_incentivo || 'Cada quilômetro conta. Continue no seu ritmo.',
        frase_motivacional: frase || desafioData.frase_incentivo || 'Cada quilômetro conta. Continue no seu ritmo.',
        contexto_frase: contextoFrase || '',

        atividades: atividades,
        desafios: desafiosConsolidados,
        desafio_em_foco: desafioPrincipalPainel,
        desafios_ativos: desafiosAtivosPainel,
        desafios_historico: desafiosHistoricoPainel,
        totalPedalado: realizadoPainel,
        total_pedalado: realizadoPainel
      }
    };
  } catch (err) {
    painelMG_logResumoAuditoriaCarregamentoInicial_(perfTotalInicio, auditoria, {
      ok: false,
      erro: err && err.message ? err.message : 'Erro interno ao carregar o painel do usuário.'
    });
    return {
      ok: false,
      code: 'PAINEL_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao carregar o painel do usuário.'
    };
  } finally {
    painelMG_limparAuditoriaCarregamentoInicial_();
  }
}

var PAINEL_MG_AUDITORIA_CARREGAMENTO_INICIAL_ = null;

function painelMG_criarAuditoriaCarregamentoInicial_() {
  return {
    getRanking_chamadas: 0,
    obterVinculosDesafioUsuario_chamadas: 0,
    leituras_ListaDesafios: 0,
    leituras_dgmbDesafios: 0
  };
}

function painelMG_definirAuditoriaCarregamentoInicial_(auditoria) {
  PAINEL_MG_AUDITORIA_CARREGAMENTO_INICIAL_ = auditoria || null;
}

function painelMG_obterAuditoriaCarregamentoInicial_() {
  return PAINEL_MG_AUDITORIA_CARREGAMENTO_INICIAL_;
}

function painelMG_limparAuditoriaCarregamentoInicial_() {
  PAINEL_MG_AUDITORIA_CARREGAMENTO_INICIAL_ = null;
}

function painelMG_incrementarAuditoriaCarregamentoInicial_(campo) {
  var auditoria = painelMG_obterAuditoriaCarregamentoInicial_();
  if (!auditoria || !campo) return;
  auditoria[campo] = Number(auditoria[campo] || 0) + 1;
}

function painelMG_logResumoAuditoriaCarregamentoInicial_(inicio, auditoria, extras) {
  var payload = {
    quantidade_chamadas_getRanking: auditoria ? auditoria.getRanking_chamadas : 0,
    quantidade_chamadas_obterVinculosDesafioUsuario: auditoria ? auditoria.obterVinculosDesafioUsuario_chamadas : 0,
    quantidade_leituras_ListaDesafios: auditoria ? auditoria.leituras_ListaDesafios : 0,
    quantidade_leituras_dgmbDesafios: auditoria ? auditoria.leituras_dgmbDesafios : 0
  };
  Object.keys(extras || {}).forEach(function(chave) {
    payload[chave] = extras[chave];
  });
  painelMG_perfLog_('painel-inicial', 'getPainelUsuario_total', inicio, payload);
}


function painelMG_perfDebugAtivo_() {
  try {
    return typeof PERFORMANCE_DEBUG !== 'undefined' && !!PERFORMANCE_DEBUG;
  } catch (e) {
    return false;
  }
}

function painelMG_perfNow_() {
  try {
    return Date.now();
  } catch (e) {
    return new Date().getTime();
  }
}

function painelMG_perfLog_(escopo, etapa, inicio, extras) {
  if (!painelMG_perfDebugAtivo_()) return;

  try {
    var fim = painelMG_perfNow_();
    var payload = {
      etapa: etapa,
      duracao_ms: fim - inicio
    };

    Object.keys(extras || {}).forEach(function(chave) {
      payload[chave] = extras[chave];
    });

    Logger.log('[Meu Giro][performance][' + escopo + '] ' + JSON.stringify(payload));
  } catch (e) {}
}

function getPainelUsuarioPosSalvarLeve(idDgmb) {
  var perfTotalInicio = painelMG_perfNow_();
  var perfEtapaInicio = perfTotalInicio;

  try {
    var id = String(idDgmb || '').trim();
    if (!id) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do usuário não informado.' };
    }

    perfEtapaInicio = painelMG_perfNow_();
    var pessoa = buscarPessoaPainelMG_(id);
    painelMG_perfLog_('painel-leve-pos-salvar', 'buscarPessoaPainelMG_', perfEtapaInicio, {
      encontrado: !!pessoa
    });

    perfEtapaInicio = painelMG_perfNow_();
    var resumoDesafios = obterMeuGiroResumoAtualizadoLeve_(id) || [];
    painelMG_perfLog_('painel-leve-pos-salvar', 'lerMeuGiroResumoAtualizadoLeve_', perfEtapaInicio, {
      total_desafios_resumo: resumoDesafios.length,
      fallback: false
    });

    if (!resumoDesafios.length) {
      perfEtapaInicio = painelMG_perfNow_();
      resumoDesafios = atualizarMeuGiroResumo_(id) || [];
      painelMG_perfLog_('painel-leve-pos-salvar', 'atualizarMeuGiroResumo_fallback_', perfEtapaInicio, {
        total_desafios_resumo: resumoDesafios.length,
        fallback: true
      });
    }

    perfEtapaInicio = painelMG_perfNow_();
    var desafio = buscarInscricaoPainelMGLeve_(id, resumoDesafios);
    painelMG_perfLog_('painel-leve-pos-salvar', desafio && desafio.usou_fallback ? 'buscarInscricaoPainelMG_' : 'buscarInscricaoPainelMGLeve_', perfEtapaInicio, {
      ok: !!(desafio && desafio.ok),
      total_desafios: desafio && desafio.desafios ? desafio.desafios.length : 0,
      fallback: !!(desafio && desafio.usou_fallback),
      motivo_fallback: desafio && desafio.motivo_fallback ? desafio.motivo_fallback : ''
    });

    if (!pessoa) {
      return { ok: false, code: 'USUARIO_NAO_ENCONTRADO', msg: 'Usuário não encontrado na base de pessoas para carregar o painel.' };
    }

    if (!desafio.ok) {
      return {
        ok: false,
        code: desafio.code,
        motivo_inscricao: desafio.motivo,
        msg: desafio.msg
      };
    }

    if (!desafio.data) {
      return {
        ok: false,
        code: 'CONTRATO_INSCRICAO_INVALIDO',
        msg: 'Dados de inscrição inválidos para o painel.'
      };
    }

    var desafioData = desafio.data;
    perfEtapaInicio = painelMG_perfNow_();
    var desafiosConsolidados = (desafio.desafios || []).map(function(item) {
      var desafioPainel = {};
      Object.keys(item || {}).forEach(function(chave) {
        desafioPainel[chave] = item[chave];
      });

      var operacional = painelMG_montarMensagemOperacional_(
        desafioPainel.status_apuracao,
        desafioPainel.status_usuario_desafio
      );
      desafioPainel.status_operacional = operacional.codigo_operacional;
      desafioPainel.mensagem_operacional = operacional.mensagem_operacional;
      return desafioPainel;
    });
    painelMG_perfLog_('painel-leve-pos-salvar', 'montagem_desafios_consolidados', perfEtapaInicio, {
      total_desafios: desafiosConsolidados.length
    });

    var desafiosAtivosPainel = desafiosConsolidados.filter(function(d) {
      return painelMG_isDesafioAtivoParaFoco_(d);
    }).sort(painelMG_compareFocoDesafiosAtivos_);
    var desafiosHistoricoPainel = desafiosConsolidados.filter(function(d) {
      return painelMG_isStatusHistorico_(d && d.status_apuracao);
    }).sort(painelMG_compareHistoricoDesafios_);

    var desafioPrincipalPainel = painelMG_selecionarDesafioPrincipal_(desafiosConsolidados);
    var meta = painelMG_toNumber_(desafioPrincipalPainel ? desafioPrincipalPainel.meta_km : desafioData.meta);
    var realizado = painelMG_toNumber_(desafioPrincipalPainel ? desafioPrincipalPainel.distancia_realizada : desafioData.realizado);
    var realizadoPainel = painelMG_round1_(realizado);

    var statusOperacionalPainel = desafioPrincipalPainel
      ? {
          codigo_operacional: desafioPrincipalPainel.status_operacional,
          mensagem_operacional: desafioPrincipalPainel.mensagem_operacional
        }
      : painelMG_montarMensagemOperacional_('', '');
    var progresso = painelMG_calcularProgresso_(meta, realizadoPainel);
    var ritmo = painelMG_calcularRitmo_(meta, realizadoPainel, desafioData.periodo_inicio, desafioData.periodo_fim);
    perfEtapaInicio = painelMG_perfNow_();
    var atividades = buscarAtividadesUsuario_(id);
    painelMG_perfLog_('painel-leve-pos-salvar', 'buscarAtividadesUsuario_', perfEtapaInicio, {
      total_atividades: atividades && atividades.length ? atividades.length : 0
    });

    painelMG_perfLog_('painel-leve-pos-salvar', 'getPainelUsuarioPosSalvarLeve_total', perfTotalInicio, {
      total_desafios: desafiosConsolidados.length,
      total_atividades: atividades && atividades.length ? atividades.length : 0
    });


    painelMG_logBug03Desafio_('getPainelUsuarioPosSalvarLeve/desafio_em_foco', desafioPrincipalPainel, 'payload leve data.desafio_em_foco');
    painelMG_logBug03ListaDesafios_('getPainelUsuarioPosSalvarLeve/desafios_ativos', desafiosAtivosPainel, 'payload leve data.desafios_ativos');
    painelMG_logBug03ListaDesafios_('getPainelUsuarioPosSalvarLeve/desafios_historico', desafiosHistoricoPainel, 'payload leve data.desafios_historico');

    return {
      ok: true,
      data: {
        nome: pessoa.nome || '',
        cidade_uf: pessoa.cidade_uf || '',
        id_dgmb: pessoa.id_dgmb || '',
        status_inscricao: desafioData.status_inscricao || 'inscrito',
        criterio_validacao_inscricao: desafioData.criterio_validacao || 'presenca_id_dgmb',
        desafio_usuario: desafioData.aba_desafio || '',

        meta: painelMG_round1_(meta),
        realizado: realizadoPainel,
        restante: progresso.restante,
        percentual: progresso.percentual,

        diaAtual: ritmo.diaAtual,
        diasRestantes: ritmo.diasRestantes,
        kmIdealAtual: ritmo.kmIdealHoje,
        kmPorDiaRestante: ritmo.kmPorDiaRestante,

        ritmo_status: ritmo.status,
        ritmo_mensagem: ritmo.mensagem,
        status_operacional: statusOperacionalPainel.codigo_operacional,
        mensagem_operacional: statusOperacionalPainel.mensagem_operacional,

        atividades: atividades,
        desafios: desafiosConsolidados,
        desafio_em_foco: desafioPrincipalPainel,
        desafios_ativos: desafiosAtivosPainel,
        desafios_historico: desafiosHistoricoPainel,
        totalPedalado: realizadoPainel,
        total_pedalado: realizadoPainel
      }
    };
  } catch (err) {
    return {
      ok: false,
      code: 'PAINEL_LEVE_POS_SALVAR_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao atualizar atividades do painel.'
    };
  }
}

function painelMG_montarMensagemOperacional_(statusApuracao, statusUsuarioDesafio) {
  var apuracao = painelMG_norm_(statusApuracao).toUpperCase();
  var usuario = painelMG_norm_(statusUsuarioDesafio).toUpperCase();

  if (usuario === 'CANCELADO') {
    return {
      codigo_operacional: 'INSCRICAO_CANCELADA',
      mensagem_operacional: 'Inscrição cancelada.'
    };
  }

  if (apuracao === 'INAPTO') {
    return {
      codigo_operacional: 'APURACAO_INAPTA',
      mensagem_operacional: 'Não foi possível apurar este desafio. Entre em contato com a organização.'
    };
  }

  if (usuario === 'EM_ANDAMENTO' && apuracao === 'ATIVO') {
    return {
      codigo_operacional: 'DESAFIO_EM_ANDAMENTO',
      mensagem_operacional: 'Desafio em andamento. Continue registrando suas atividades.'
    };
  }

  if (usuario === 'EM_ANDAMENTO' && apuracao === 'CONCLUIDO') {
    return {
      codigo_operacional: 'META_ATINGIDA_AGUARDANDO_VALIDACAO',
      mensagem_operacional: 'Meta atingida. Aguardando validação da organização.'
    };
  }

  if (usuario === 'CONCLUIDO' && apuracao === 'CONCLUIDO') {
    return {
      codigo_operacional: 'DESAFIO_CONCLUIDO_OFICIALMENTE',
      mensagem_operacional: 'Desafio concluído oficialmente.'
    };
  }

  if (usuario === 'NAO_CONCLUIDO' && apuracao === 'EXPIRADO') {
    return {
      codigo_operacional: 'DESAFIO_ENCERRADO_SEM_CONCLUSAO',
      mensagem_operacional: 'Desafio encerrado sem conclusão da meta.'
    };
  }

  return {
    codigo_operacional: 'STATUS_EM_ANALISE',
    mensagem_operacional: 'Status do desafio em análise.'
  };
}

function painelMG_montarBaseFrase_(dados) {
  var src = dados || {};

  return {
    id_dgmb: String(src.id_dgmb || '').trim(),
    nome: String(src.nome || '').trim(),
    meta: painelMG_round1_(painelMG_toNumber_(src.meta)),
    realizado: painelMG_round1_(painelMG_toNumber_(src.realizado)),
    restante: painelMG_round1_(painelMG_toNumber_(src.restante)),
    percentual: painelMG_round1_(painelMG_toNumber_(src.percentual)),
    ritmo_status: String(src.ritmo_status || '').trim(),
    ritmo_mensagem: String(src.ritmo_mensagem || '').trim()
  };
}

function buscarAtividadesUsuario_(idDgmb) {
  var items = [];
  try {
    items = getAllObjects_(SHEETS.REGISTRO_KM) || [];
  } catch (e) {
    return [];
  }
  var out = [];
  var activityIdsIncluidos = {};

  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var rowId = painelMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']);

    if (painelMG_norm_(rowId) === painelMG_norm_(idDgmb)) {
      var dataOriginal = painelMG_firstFilled_(row, ['Data_Atividade', 'Data', 'data_atividade', 'data']);
      var dataNormalizada = painelMG_normalizarDataAtividade_(dataOriginal);

      var timestampOriginal = painelMG_firstFilled_(row, ['Timestamp', 'timestamp']);
      var chaveEdicao = normalizarTimestampEdicao_(timestampOriginal);

      var activityId = obterActivityIdRegistroKm_(row);
      if (activityId) {
        if (activityIdsIncluidos[activityId]) continue;
        activityIdsIncluidos[activityId] = true;
      }

      out.push({
        activity_id: activityId,
        chave_edicao: String(chaveEdicao || '').trim(),
        data: dataNormalizada || painelMG_norm_(dataOriginal),
        km: painelMG_round1_(painelMG_toNumber_(painelMG_firstFilled_(row, ['KM', 'km'])))
      });
    }
  }

  out.sort(function(a, b) {
    return String(b.data || '').localeCompare(String(a.data || ''));
  });

  return out;
}

function buscarPessoaPainelMG_(idDgmb) {
  var items = getAllObjects_(SHEETS.PESSOAS);

  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var rowId = painelMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']);

    if (painelMG_norm_(rowId) === painelMG_norm_(idDgmb)) {
      return {
        id_dgmb: painelMG_norm_(rowId),
        nome: painelMG_norm_(painelMG_firstFilled_(row, ['nome', 'Nome'])),
        cidade_uf: painelMG_norm_(painelMG_firstFilled_(row, ['Cidade-UF', 'Cidade_UF', 'cidade_uf', 'cidade-uf']))
      };
    }
  }

  return null;
}

function buscarInscricaoPainelMGLeve_(idDgmb, resumoAtualizado) {
  var resumo = resumoAtualizado && resumoAtualizado.length ? resumoAtualizado : [];
  var fallback = function(motivo) {
    var atual = buscarInscricaoPainelMG_(idDgmb, resumo, true);
    atual.usou_fallback = true;
    atual.motivo_fallback = motivo || '';
    return atual;
  };

  if (!resumo.length) return fallback('resumo_ausente');

  var desafioPrincipal = painelMG_selecionarDesafioPrincipal_(resumo);
  if (!desafioPrincipal) return fallback('desafio_principal_ausente');

  var inscricao = painelMG_obterInscricaoLevePorDesafio_(idDgmb, desafioPrincipal);
  if (!inscricao || inscricao.inscricao_valida === false) {
    if (inscricao) {
      var erro = montarErroInscricaoInvalida_(inscricao);
      return { ok: false, code: erro.code, motivo: erro.motivo, msg: erro.msg };
    }
    return fallback('inscricao_nao_localizada');
  }

  if (!inscricao.periodo_inicio || !inscricao.periodo_fim) {
    return fallback('periodo_ausente');
  }

  return {
    ok: true,
    data: {
      id_dgmb: painelMG_norm_(inscricao.id_dgmb),
      meta: desafioPrincipal.meta_km,
      realizado: desafioPrincipal.distancia_realizada,
      status_inscricao: painelMG_norm_(inscricao.status_inscricao),
      criterio_validacao: painelMG_norm_(inscricao.criterio_validacao),
      aba_desafio: painelMG_norm_(inscricao.aba_desafio),
      frase_incentivo: painelMG_norm_(inscricao.frase_incentivo),
      periodo_inicio: inscricao.periodo_inicio,
      periodo_fim: inscricao.periodo_fim
    },
    desafios: resumo,
    usou_fallback: false
  };
}

// Validação manual do Desafio em foco (GAS):
// A) inscrição em julho/2026 sem REGISTRO_KM => julho/2026 aparece com 0 km.
// B) abril/2026 incompleto + julho/2026 ativo => julho/2026 assume o foco.
// C) dois ativos em julho/2026 => maior Meta_KM assume o foco.
// D) sem desafio no mês corrente + ativo anterior => ativo mais recente assume o foco.
// E) concluído antigo + ativo atual => ativo atual assume o foco.
// F) julho/2026 prorrogado e ainda ativo em agosto => não perde foco só pelo mês do relógio.
// G) mesma meta e mesmo período => desempate previsível por inscrição, desafio, item e nome.
// H) 0 km sem linhas em REGISTRO_KM => painel carrega usando inscrição/resumo com distância 0.
function painelMG_selecionarDesafioPrincipal_(resumo) {
  var lista = Array.isArray(resumo) ? resumo.slice() : [];
  if (!lista.length) return null;

  var ativos = lista.filter(function(item) {
    return painelMG_isDesafioAtivoParaFoco_(item);
  }).sort(painelMG_compareFocoDesafiosAtivos_);
  if (ativos.length) return ativos[0];

  var historicos = lista.filter(function(d) {
    return painelMG_isStatusHistorico_(d && d.status_apuracao);
  }).sort(painelMG_compareHistoricoDesafios_);
  return historicos.length ? historicos[0] : lista.sort(painelMG_compareHistoricoDesafios_)[0];
}

function painelMG_obterInscricaoLevePorDesafio_(idDgmb, desafioPrincipal) {
  var perfMontagemInicio = painelMG_perfNow_();
  var id = painelMG_norm_(idDgmb);
  if (!id) {
    painelMG_perfLog_('painel-leve-pos-salvar', 'montagem_inscricao_leve', perfMontagemInicio, {
      encontrada: false,
      motivo: 'id_ausente'
    });
    return null;
  }

  var perfLeituraInicio = painelMG_perfNow_();
  var cacheDesafios = obterDgmbDesafiosCacheExecucao_('painelMG_obterInscricaoLevePorDesafio_');
  var abaDesafio = cacheDesafios.aba;
  var values = cacheDesafios.values;
  painelMG_perfLog_('painel-leve-pos-salvar', 'leitura_dgmbDesafios_inscricao_leve', perfLeituraInicio, {
    quantidade_linhas_dgmbDesafios: values && values.length ? values.length - 1 : 0,
    usou_cache_dgmbDesafios: cacheDesafios.usouCache
  });
  if (!values || values.length < 2) {
    painelMG_perfLog_('painel-leve-pos-salvar', 'montagem_inscricao_leve', perfMontagemInicio, {
      encontrada: false,
      motivo: 'dgmbDesafios_sem_dados'
    });
    return null;
  }

  var map = buildHeaderMap_(values[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], abaDesafio);
  var idxMeta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km']);
  var idxRealizado = getOptionalColumnIndex_(map, ['distancia_realizada', 'distancia realizada']);
  var idxFrase = getOptionalColumnIndex_(map, ['frase_incentivo']);
  var idxStatus = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição', 'status', 'situacao', 'situação']);
  var idxStatusUsuarioDesafio = getOptionalColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  var idxPagamento = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagto_status', 'pagamento', 'pix_status']);
  var idxInscricao = getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxObs = getOptionalColumnIndex_(map, ['observacao', 'observação']);
  var idxIdDesafio = getIdDesafioColumnIndex_(map);
  var idxPeriodo = getOptionalColumnIndex_(map, ['periodo_desafio', 'periodo desafio', 'período_desafio', 'período desafio']);
  var idxInicio = getOptionalColumnIndex_(map, ['data_inicio_desafio', 'data inicio desafio', 'data início desafio']);
  var idxFim = getOptionalColumnIndex_(map, ['data_fim_desafio', 'data fim desafio']);
  var periodosLista = buildListaDesafiosContexto_(getSpreadsheet_()).periodos;

  var alvoInscricao = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_inscricao);
  var alvoDesafio = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_desafio);
  var alvoItem = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_item_estoque);
  var primeiraInvalida = null;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (painelMG_norm_(row[idxId]) !== id) continue;

    var idInscricao = idxInscricao > -1 ? painelMG_norm_(row[idxInscricao]) : '';
    var idDesafio = obterIdDesafioRegistro_(row, idxIdDesafio, idxObs);
    var idItem = idxItem > -1 ? painelMG_norm_(row[idxItem]) : '';
    if (alvoInscricao && idInscricao !== alvoInscricao) continue;
    if (!alvoInscricao && alvoDesafio && idDesafio !== alvoDesafio) continue;
    if (alvoItem && idItem !== alvoItem) continue;

    var statusInscricao = idxStatus > -1 ? painelMG_norm_(row[idxStatus]) : '';
    if (!statusInscricao && idxStatusUsuarioDesafio > -1) statusInscricao = painelMG_norm_(row[idxStatusUsuarioDesafio]);
    var statusConfirmacao = idxConfirmacao > -1 ? painelMG_norm_(row[idxConfirmacao]) : '';
    var statusPagamento = idxPagamento > -1 ? painelMG_norm_(row[idxPagamento]) : '';
    var validacao = validarInscricaoMinima_({
      status_inscricao: statusInscricao,
      status_confirmacao: statusConfirmacao,
      status_pagamento: statusPagamento
    });
    var periodoDatas = {
      inicio: normalizarDataISO_(idxInicio > -1 ? row[idxInicio] : ''),
      fim: normalizarDataISO_(idxFim > -1 ? row[idxFim] : '')
    };
    var periodoTexto = idxPeriodo > -1 ? extrairPeriodoDesafioTexto_(row[idxPeriodo]) : { inicio: '', fim: '' };
    var periodoLista = (idDesafio && periodosLista.byId[idDesafio]) || { inicio: '', fim: '' };
    var periodoSelecionado = periodoCompletoValido_(periodoTexto)
      ? periodoTexto
      : periodoCompletoValido_(periodoLista)
        ? periodoLista
        : periodoDatas;
    var inicio = periodoCompletoValido_(periodoSelecionado) ? periodoSelecionado.inicio : '';
    var fim = periodoCompletoValido_(periodoSelecionado) ? periodoSelecionado.fim : '';

    var inscricao = {
      id_dgmb: id,
      aba_desafio: abaDesafio,
      status_inscricao: statusInscricao || 'inscrito',
      status_confirmacao: statusConfirmacao,
      status_pagamento: statusPagamento,
      inscricao_valida: validacao.valida,
      criterio_validacao: validacao.criterio,
      meta: idxMeta > -1 ? row[idxMeta] : '',
      distancia_realizada: idxRealizado > -1 ? row[idxRealizado] : '',
      frase_incentivo: idxFrase > -1 ? painelMG_norm_(row[idxFrase]) : '',
      periodo_inicio: inicio || '',
      periodo_fim: fim || ''
    };

    if (validacao.valida) {
      painelMG_perfLog_('painel-leve-pos-salvar', 'montagem_inscricao_leve', perfMontagemInicio, {
        encontrada: true,
        valida: true,
        linha: i + 1,
        possui_periodo: !!(inscricao.periodo_inicio && inscricao.periodo_fim)
      });
      return inscricao;
    }
    if (!primeiraInvalida) primeiraInvalida = inscricao;
  }

  painelMG_perfLog_('painel-leve-pos-salvar', 'montagem_inscricao_leve', perfMontagemInicio, {
    encontrada: !!primeiraInvalida,
    valida: false,
    motivo: primeiraInvalida ? 'inscricao_invalida' : 'inscricao_nao_localizada'
  });
  return primeiraInvalida;
}

function buscarInscricaoPainelMG_(idDgmb, resumoAtualizado, evitarRecalculoResumo) {
  var inscricao = obterDadosInscricaoUsuario_(idDgmb);

  if (!inscricao || inscricao.inscricao_valida === false) {
    var erro = montarErroInscricaoInvalida_(inscricao);
    return {
      ok: false,
      code: erro.code,
      motivo: erro.motivo,
      msg: erro.msg
    };
  }

  var resumo = resumoAtualizado && resumoAtualizado.length ? resumoAtualizado : [];
  if (!resumo.length && !evitarRecalculoResumo) {
    try {
      resumo = atualizarMeuGiroResumo_(idDgmb) || [];
    } catch (e) {
      resumo = [];
    }
  }

  var desafioPrincipal = null;
  for (var i = 0; i < resumo.length; i++) {
    if (resumo[i].status_apuracao === 'ATIVO') {
      desafioPrincipal = resumo[i];
      break;
    }
  }
  if (!desafioPrincipal && resumo.length) {
    var historicos = resumo.filter(function(d) {
      return painelMG_isStatusHistorico_(d && d.status_apuracao);
    }).sort(painelMG_compareHistoricoDesafios_);
    desafioPrincipal = historicos.length ? historicos[0] : resumo[0];
  }

  var vinculoPrincipal = painelMG_buscarVinculoPrincipal_(idDgmb, desafioPrincipal, evitarRecalculoResumo);
  var periodoPorAba = (vinculoPrincipal.periodo_inicio && vinculoPrincipal.periodo_fim)
    ? { periodo_inicio: '', periodo_fim: '' }
    : painelMG_obterPeriodoOficialPorAba_(inscricao.aba_desafio);

  return {
    ok: true,
    data: {
      id_dgmb: painelMG_norm_(inscricao.id_dgmb),
      meta: desafioPrincipal ? desafioPrincipal.meta_km : inscricao.meta,
      realizado: desafioPrincipal ? desafioPrincipal.distancia_realizada : inscricao.distancia_realizada,
      status_inscricao: painelMG_norm_(inscricao.status_inscricao),
      criterio_validacao: painelMG_norm_(inscricao.criterio_validacao),
      aba_desafio: painelMG_norm_(inscricao.aba_desafio),
      frase_incentivo: painelMG_norm_(inscricao.frase_incentivo),
      periodo_inicio: vinculoPrincipal.periodo_inicio || periodoPorAba.periodo_inicio,
      periodo_fim: vinculoPrincipal.periodo_fim || periodoPorAba.periodo_fim
    },
    desafios: resumo
  };
}

function painelMG_calcularProgresso_(meta, realizado) {
  var restante = Math.max(painelMG_toNumber_(meta) - painelMG_toNumber_(realizado), 0);
  var percentual = painelMG_toNumber_(meta) > 0 ? (painelMG_toNumber_(realizado) / painelMG_toNumber_(meta)) * 100 : 0;

  return {
    restante: painelMG_round1_(restante),
    percentual: painelMG_round1_(painelMG_clamp_(percentual, 0, 100))
  };
}


function percentualMetaConcluida_(meta, realizado, restante) {
  var metaNumero = painelMG_toNumber_(meta);
  var realizadoNumero = painelMG_toNumber_(realizado);
  var restanteNumero = isFinite(restante) ? painelMG_toNumber_(restante) : Math.max(metaNumero - realizadoNumero, 0);
  var percentual = metaNumero > 0 ? (realizadoNumero / metaNumero) * 100 : 0;

  return percentual >= 100 || restanteNumero <= 0;
}

function painelMG_calcularRitmo_(meta, realizado, periodoInicio, periodoFim) {
  var now = new Date();
  var inicio = painelMG_parseDataISO_(periodoInicio) || new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var fim = painelMG_parseDataISO_(periodoFim) || inicio;

  if (fim.getTime() < inicio.getTime()) {
    var swap = inicio;
    inicio = fim;
    fim = swap;
  }

  var msDia = 24 * 60 * 60 * 1000;
  var diasTotal = Math.max(Math.floor((fim.getTime() - inicio.getTime()) / msDia) + 1, 1);
  var diaAtual = Math.floor((now.getTime() - inicio.getTime()) / msDia) + 1;
  if (diaAtual < 1) diaAtual = 1;
  if (diaAtual > diasTotal) diaAtual = diasTotal;

  var diasRestantes = Math.max(diasTotal - diaAtual, 0);
  var kmIdealPorDia = painelMG_toNumber_(meta) > 0 ? painelMG_toNumber_(meta) / diasTotal : 0;
  var kmIdealHoje = kmIdealPorDia * diaAtual;
  var restante = Math.max(painelMG_toNumber_(meta) - painelMG_toNumber_(realizado), 0);
  var kmPorDiaRestante = diasRestantes > 0 ? restante / diasRestantes : restante;

  var diferenca = painelMG_toNumber_(realizado) - kmIdealHoje;
  var tolerancia = Math.max(2, painelMG_toNumber_(meta) * 0.01);

  var status = 'Você está no ritmo.';
  var mensagem = 'Continue pedalando para manter sua evolução.';

  if (percentualMetaConcluida_(meta, realizado, restante)) {
    status = 'Desafio concluído!';
    mensagem = 'Você alcançou sua meta. Parabéns pela conquista.';
  } else if (diferenca > tolerancia) {
    status = 'Você está à frente da meta.';
    mensagem = 'Seu progresso está acima do ritmo esperado para este momento do desafio.';
  } else if (diferenca < -tolerancia) {
    status = 'Você está um pouco abaixo do ritmo.';
    mensagem = 'Alguns pedais extras podem ajudar a recuperar o ritmo do desafio.';
  }

  return {
    diaAtual: diaAtual,
    diasRestantes: diasRestantes,
    kmIdealHoje: painelMG_round1_(kmIdealHoje),
    kmPorDiaRestante: painelMG_round1_(kmPorDiaRestante),
    status: status,
    mensagem: mensagem
  };
}


function painelMG_buscarVinculoPrincipal_(idDgmb, desafioPrincipal, usarResumoComoFonte) {
  var vazio = { periodo_inicio: '', periodo_fim: '' };
  try {
    if (usarResumoComoFonte) {
      if (desafioPrincipal) {
        var periodoResumo = {
          periodo_inicio: painelMG_norm_(desafioPrincipal.periodo_inicio),
          periodo_fim: painelMG_norm_(desafioPrincipal.periodo_fim)
        };
        if (periodoResumo.periodo_inicio && periodoResumo.periodo_fim) {
          return periodoResumo;
        }
      }

      return vazio;
    }

    var vinculos = obterVinculosDesafioUsuario_(idDgmb) || [];
    var idInscricaoPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_inscricao);
    var idDesafioPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_desafio);
    var idItemPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_item_estoque);

    for (var i = 0; i < vinculos.length; i++) {
      var v = vinculos[i] || {};
      var corresponde = idInscricaoPrincipal
        ? painelMG_norm_(v.id_inscricao) === idInscricaoPrincipal
        : painelMG_norm_(v.id_desafio) === idDesafioPrincipal &&
          painelMG_norm_(v.id_item_estoque) === idItemPrincipal;
      if (corresponde) {
        return {
          periodo_inicio: painelMG_norm_(v.periodo_inicio),
          periodo_fim: painelMG_norm_(v.periodo_fim)
        };
      }
    }
  } catch (e) {}

  return vazio;
}

function painelMG_parseDataISO_(valor) {
  var texto = painelMG_norm_(valor);
  if (!texto) return null;

  var match = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}


function painelMG_obterPeriodoOficialPorAba_(abaDesafio) {
  var vazio = { periodo_inicio: '', periodo_fim: '' };
  var aba = painelMG_norm_(abaDesafio);
  if (!aba) return vazio;

  try {
    var periodos = buildPeriodoOficialPorAbaEId_(getSpreadsheet_());
    var periodo = (periodos && periodos.byAba && periodos.byAba[aba]) || {};
    return {
      periodo_inicio: painelMG_norm_(periodo.inicio),
      periodo_fim: painelMG_norm_(periodo.fim)
    };
  } catch (e) {
    return vazio;
  }
}

function painelMG_calcularPosicaoRanking_(idDgmb, idDesafio, idItemEstoque, idInscricao) {
  var resposta = getRanking(idDgmb, idDesafio, idItemEstoque, idInscricao);
  if (!resposta || !resposta.ok) return { posicao: 0, total: 0 };
  return {
    posicao: Number(resposta.posicao_usuario || 0),
    total: Number(resposta.total || 0),
    diagnosticos: resposta.diagnosticos || {}
  };
}

function painelMG_montarRankingPorDesafio_(idDgmb, desafios) {
  var out = {};
  var lista = Array.isArray(desafios) ? desafios : [];

  for (var i = 0; i < lista.length; i++) {
    var item = lista[i] || {};
    var idDesafio = painelMG_norm_(item.id_desafio);
    var idItem = painelMG_norm_(item.id_item_estoque);
    if (!idDesafio) continue;

    var chave = painelMG_chaveDesafioPainel_(item);
    if (out[chave]) continue;

    out[chave] = painelMG_calcularPosicaoRanking_(idDgmb, idDesafio, idItem, item.id_inscricao);
  }

  return out;
}

function painelMG_chaveDesafioPainel_(desafio) {
  var item = desafio || {};
  var idInscricao = painelMG_norm_(item.id_inscricao);
  var idDesafio = painelMG_norm_(item.id_desafio);
  var idItem = painelMG_norm_(item.id_item_estoque);
  var inicio = painelMG_norm_(item.periodo_inicio);
  var fim = painelMG_norm_(item.periodo_fim);
  var meta = painelMG_norm_(item.meta_km);

  if (idInscricao) return ['INSCRICAO', idInscricao].join('|');
  if (!idDesafio) return '';
  if (idItem) return [idDesafio, idItem].join('|');
  if (inicio || fim || meta) return [idDesafio, inicio, fim, meta].join('|');
  return idDesafio;
}


function painelMG_firstFilled_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, keys[i])) {
      var value = obj[keys[i]];
      if (value !== '' && value !== null && value !== undefined) {
        return value;
      }
    }
  }
  return '';
}

function painelMG_normalizarDataAtividade_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var s = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    var p = s.split('/');
    return p[2] + '-' + p[1] + '-' + p[0];
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return '';
}

function painelMG_norm_(value) {
  return String(value || '').trim();
}

function painelMG_normalizarStatus_(status) {
  return painelMG_norm_(status).toUpperCase();
}

function painelMG_isStatusAtivo_(status) {
  return painelMG_normalizarStatus_(status) === 'ATIVO';
}

function painelMG_isDesafioAtivoParaFoco_(desafio) {
  var item = desafio || {};
  var statusUsuario = painelMG_normalizarStatus_(item.status_usuario_desafio);
  var statusApuracao = painelMG_normalizarStatus_(item.status_apuracao);
  var statusDesafio = painelMG_normalizarStatus_(item.status_desafio);
  var statusLista = painelMG_normalizarStatus_(item.status_lista_desafios);
  var possuiPeriodo = !!(painelMG_normalizarDataISO_(item.periodo_inicio) || painelMG_normalizarDataISO_(item.periodo_fim));

  // STATUS_EM_ANALISE é mantido como visível somente para desafios com período,
  // pois MEU_GIRO_RESUMO usa esse status para inscrição acompanhável ainda não concluída.
  // Status finais em qualquer fonte continuam bloqueando o foco ativo.
  if (painelMG_isStatusFinalFoco_(statusUsuario) || painelMG_isStatusFinalFoco_(statusApuracao) ||
      painelMG_isStatusFinalFoco_(statusDesafio) || painelMG_isStatusFinalFoco_(statusLista)) return false;

  if (statusUsuario === 'EM_ANDAMENTO') return true;
  if (statusApuracao === 'ATIVO') return true;
  if (statusDesafio === 'ATIVO') return true;
  if (statusLista === 'ATIVO') return true;

  return possuiPeriodo && (statusUsuario === 'STATUS_EM_ANALISE' || statusApuracao === 'STATUS_EM_ANALISE');
}

function painelMG_isStatusFinalFoco_(status) {
  var finais = {
    CONCLUIDO: true,
    CANCELADO: true,
    DESISTENTE: true,
    EXPIRADO: true,
    ENCERRADO: true,
    INAPTO: true,
    NAO_CONCLUIDO: true
  };
  return !!finais[painelMG_normalizarStatus_(status)];
}

function painelMG_compareFocoDesafiosAtivos_(a, b) {
  var ativoAgoraA = painelMG_desafioAtivoNaDataReferencia_(a) ? 0 : 1;
  var ativoAgoraB = painelMG_desafioAtivoNaDataReferencia_(b) ? 0 : 1;
  if (ativoAgoraA !== ativoAgoraB) return ativoAgoraA - ativoAgoraB;

  var periodoA = painelMG_periodoOrdenacaoFoco_(a);
  var periodoB = painelMG_periodoOrdenacaoFoco_(b);
  var mesmoPeriodo = periodoA.chave && periodoA.chave === periodoB.chave;

  if (!mesmoPeriodo) {
    if (periodoA.fim !== periodoB.fim) return periodoB.fim.localeCompare(periodoA.fim);
    if (periodoA.inicio !== periodoB.inicio) return periodoB.inicio.localeCompare(periodoA.inicio);
  }

  var metaA = painelMG_toNumber_(a && a.meta_km);
  var metaB = painelMG_toNumber_(b && b.meta_km);
  if (metaA !== metaB) return metaB - metaA;

  if (mesmoPeriodo) {
    if (periodoA.fim !== periodoB.fim) return periodoB.fim.localeCompare(periodoA.fim);
    if (periodoA.inicio !== periodoB.inicio) return periodoB.inicio.localeCompare(periodoA.inicio);
  }

  var inscricaoA = painelMG_norm_(a && a.id_inscricao);
  var inscricaoB = painelMG_norm_(b && b.id_inscricao);
  if (inscricaoA !== inscricaoB) return inscricaoA.localeCompare(inscricaoB);

  var desafioA = painelMG_norm_(a && a.id_desafio);
  var desafioB = painelMG_norm_(b && b.id_desafio);
  if (desafioA !== desafioB) return desafioA.localeCompare(desafioB);

  var itemA = painelMG_norm_(a && a.id_item_estoque);
  var itemB = painelMG_norm_(b && b.id_item_estoque);
  if (itemA !== itemB) return itemA.localeCompare(itemB);

  return String((a && a.nome_desafio) || '').localeCompare(String((b && b.nome_desafio) || ''));
}

function painelMG_desafioAtivoNaDataReferencia_(desafio) {
  var hoje = new Date();
  var inicio = painelMG_parseDataISO_(desafio && desafio.periodo_inicio);
  var fim = painelMG_parseDataISO_(desafio && desafio.periodo_fim);

  if (inicio && fim) return inicio.getTime() <= hoje.getTime() && fim.getTime() >= hoje.getTime();
  if (inicio) return inicio.getTime() <= hoje.getTime();
  if (fim) return fim.getTime() >= hoje.getTime();
  return false;
}

function painelMG_periodoOrdenacaoFoco_(desafio) {
  var inicio = painelMG_normalizarDataISO_(desafio && desafio.periodo_inicio);
  var fim = painelMG_normalizarDataISO_(desafio && desafio.periodo_fim);
  return {
    inicio: inicio || '',
    fim: fim || inicio || '',
    chave: [inicio || '', fim || ''].join('|')
  };
}

function painelMG_isStatusHistorico_(status) {
  var normalizado = painelMG_normalizarStatus_(status);
  var statusFinais = {
    CONCLUIDO: true,
    EXPIRADO: true,
    CANCELADO: true,
    DESISTENTE: true,
    ENCERRADO: true
  };
  return !!statusFinais[normalizado];
}

function painelMG_normalizarDataISO_(value) {
  return normalizarDataISO_(value);
}

function painelMG_compareHistoricoDesafios_(a, b) {
  var fimA = painelMG_normalizarDataISO_(a && a.periodo_fim);
  var fimB = painelMG_normalizarDataISO_(b && b.periodo_fim);
  if (fimA !== fimB) return String(fimB || '').localeCompare(String(fimA || ''));

  var inicioA = painelMG_normalizarDataISO_(a && a.periodo_inicio);
  var inicioB = painelMG_normalizarDataISO_(b && b.periodo_inicio);
  if (inicioA !== inicioB) return String(inicioB || '').localeCompare(String(inicioA || ''));

  return String((a && a.nome_desafio) || '').localeCompare(String((b && b.nome_desafio) || ''));
}

function painelMG_toNumber_(value) {
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }

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

function painelMG_clamp_(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function painelMG_round1_(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 10) / 10;
}

function painelMG_obterFraseSeguro_(frasePadrao) {
  var frase = String(frasePadrao || '').trim();
  if (frase) return frase;
  return 'Cada quilômetro conta. Continue no seu ritmo.';
}
