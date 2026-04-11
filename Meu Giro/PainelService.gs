function getPainelUsuario(idDgmb) {
  try {
    var id = String(idDgmb || '').trim();
    if (!id) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do usuário não informado.' };
    }

    var pessoa = buscarPessoaPainelMG_(id);
    var resumoDesafios = atualizarMeuGiroResumo_(id) || [];
    var desafio = buscarInscricaoPainelMG_(id, resumoDesafios);

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
    var desafiosConsolidados = desafio.desafios || [];

    var meta = painelMG_toNumber_(desafioData.meta);
    var realizado = painelMG_toNumber_(desafioData.realizado);
    var realizadoPainel = painelMG_round1_(realizado);

    var desafioPrincipalPainel = null;
    for (var idxResumo = 0; idxResumo < desafiosConsolidados.length; idxResumo++) {
      if (desafiosConsolidados[idxResumo].status_apuracao === 'ATIVO') {
        desafioPrincipalPainel = desafiosConsolidados[idxResumo];
        break;
      }
    }
    if (!desafioPrincipalPainel && desafiosConsolidados.length) {
      desafioPrincipalPainel = desafiosConsolidados[0];
    }

    var progresso = painelMG_calcularProgresso_(meta, realizadoPainel);
    var ritmo = painelMG_calcularRitmo_(meta, realizadoPainel, desafioData.periodo_inicio, desafioData.periodo_fim);
    var atividades = buscarAtividadesUsuario_(id);
    var rankingInfo = painelMG_calcularPosicaoRanking_(
      id,
      desafioPrincipalPainel && desafioPrincipalPainel.id_desafio,
      desafioPrincipalPainel && desafioPrincipalPainel.id_item_estoque
    );
    var rankingPorDesafio = painelMG_montarRankingPorDesafio_(id, desafiosConsolidados);

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

    return {
      ok: true,
      data: {
        // Compatibilidade temporária de contrato (Meu Giro + legado InscriçãoDesafio).
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
        desafios_ativos: desafiosConsolidados.filter(function(d) {
          return painelMG_norm_(d && d.status_apuracao).toUpperCase() === 'ATIVO';
        }),
        desafios_historico: desafiosConsolidados.filter(function(d) {
          return painelMG_norm_(d && d.status_apuracao).toUpperCase() !== 'ATIVO';
        }),
        totalPedalado: realizadoPainel,
        total_pedalado: realizadoPainel
      }
    };
  } catch (err) {
    return {
      ok: false,
      code: 'PAINEL_ERROR',
      msg: err && err.message ? err.message : 'Erro interno ao carregar o painel do usuário.'
    };
  }
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

  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var rowId = painelMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']);

    if (painelMG_norm_(rowId) === painelMG_norm_(idDgmb)) {
      var dataOriginal = painelMG_firstFilled_(row, ['Data_Atividade', 'Data', 'data_atividade', 'data']);
      var dataNormalizada = painelMG_normalizarDataAtividade_(dataOriginal);

      var timestampOriginal = painelMG_firstFilled_(row, ['Timestamp', 'timestamp']);
      var chaveEdicao = normalizarTimestampEdicao_(timestampOriginal);

      var activityIdOriginal = painelMG_firstFilled_(row, ['activity_id', 'Activity_ID', 'activity id', 'id_atividade', 'ID_Atividade']);
      out.push({
        activity_id: painelMG_norm_(activityIdOriginal),
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

function buscarInscricaoPainelMG_(idDgmb, resumoAtualizado) {
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
  if (!resumo.length) {
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
  if (!desafioPrincipal && resumo.length) desafioPrincipal = resumo[0];

  var vinculoPrincipal = painelMG_buscarVinculoPrincipal_(idDgmb, desafioPrincipal);
  var periodoPorAba = painelMG_obterPeriodoOficialPorAba_(inscricao.aba_desafio);

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

  var status = 'Você está à frente da meta.'
  var mensagem = 'Seu progresso está acima do ritmo esperado para este momento do desafio.'

  if (diferenca > tolerancia) {
    status = 'Você está à frente da meta.';
      mensagem = 'Seu progresso está acima do ritmo esperado para este momento do desafio.';

  } else if (diferenca < -tolerancia) {
    status = 'Você está um pouco abaixo do ritmo.'
    mensagem =  'Alguns pedais extras podem ajudar a recuperar o ritmo do desafio.'
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


function painelMG_buscarVinculoPrincipal_(idDgmb, desafioPrincipal) {
  var vazio = { periodo_inicio: '', periodo_fim: '' };
  try {
    var vinculos = obterVinculosDesafioUsuario_(idDgmb) || [];
    var idDesafioPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_desafio);
    var idItemPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_item_estoque);

    for (var i = 0; i < vinculos.length; i++) {
      var v = vinculos[i] || {};
      if (
        painelMG_norm_(v.id_desafio) === idDesafioPrincipal &&
        painelMG_norm_(v.id_item_estoque) === idItemPrincipal
      ) {
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

function painelMG_calcularPosicaoRanking_(idDgmb, idDesafio, idItemEstoque) {
  var idUsuario = painelMG_norm_(idDgmb);
  var desafioPrincipal = painelMG_norm_(idDesafio);
  var itemPrincipal = painelMG_norm_(idItemEstoque);
  var grupoBasePrincipal = painelMG_extrairGrupoBaseDesafio_(itemPrincipal);

  if (!idUsuario || !desafioPrincipal) {
    return { posicao: 0, total: 0 };
  }

  var resumo = [];
  try {
    resumo = getAllObjects_(SHEETS.MEU_GIRO_RESUMO) || [];
  } catch (e) {
    resumo = [];
  }

  if (!resumo.length) {
    return { posicao: 0, total: 0 };
  }

  var statusValidos = { ATIVO: true, CONCLUIDO: true };
  var lista = [];

  for (var i = 0; i < resumo.length; i++) {
    var row = resumo[i] || {};
    var did = painelMG_norm_(painelMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']));
    var rowDesafio = painelMG_norm_(painelMG_firstFilled_(row, ['ID_DESAFIO', 'id_desafio']));
    var rowItem = painelMG_norm_(painelMG_firstFilled_(row, ['id_item_estoque', 'id item estoque']));
    var rowGrupoBase = painelMG_extrairGrupoBaseDesafio_(rowItem);
    var rowStatus = painelMG_norm_(painelMG_firstFilled_(row, ['Status_Apuracao', 'status_apuracao'])).toUpperCase();

    if (!did) continue;
    if (rowDesafio !== desafioPrincipal) continue;
    if (grupoBasePrincipal && rowGrupoBase !== grupoBasePrincipal) continue;
    if (!statusValidos[rowStatus]) continue;

    lista.push({
      id_dgmb: did,
      distancia_realizada: painelMG_round1_(painelMG_toNumber_(painelMG_firstFilled_(row, ['Distancia_Realizada', 'distancia_realizada']))),
      percentual_concluido: painelMG_round1_(painelMG_toNumber_(painelMG_firstFilled_(row, ['Percentual_Concluido', 'percentual_concluido'])))
    });
  }

  if (!lista.length) {
    return { posicao: 0, total: 0 };
  }

  lista.sort(function(a, b) {
    if (b.distancia_realizada !== a.distancia_realizada) return b.distancia_realizada - a.distancia_realizada;
    if (b.percentual_concluido !== a.percentual_concluido) return b.percentual_concluido - a.percentual_concluido;
    return String(a.id_dgmb || '').localeCompare(String(b.id_dgmb || ''));
  });

  for (var j = 0; j < lista.length; j++) {
    if (lista[j].id_dgmb === idUsuario) {
      return { posicao: j + 1, total: lista.length };
    }
  }

  return { posicao: 0, total: lista.length };
}

function painelMG_extrairGrupoBaseDesafio_(idItemEstoque) {
  var item = painelMG_norm_(idItemEstoque);
  if (!item) return '';

  var semKm = item.replace(/_[0-9]+(?:[.,][0-9]+)?$/g, '');
  return painelMG_norm_(semKm || item);
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

    out[chave] = painelMG_calcularPosicaoRanking_(idDgmb, idDesafio, idItem);
  }

  return out;
}

function painelMG_chaveDesafioPainel_(desafio) {
  var item = desafio || {};
  var idDesafio = painelMG_norm_(item.id_desafio);
  var idItem = painelMG_norm_(item.id_item_estoque);
  var inicio = painelMG_norm_(item.periodo_inicio);
  var fim = painelMG_norm_(item.periodo_fim);

  if (!idDesafio) return '';
  if (idItem) return [idDesafio, idItem].join('|');
  if (inicio || fim) return [idDesafio, inicio, fim].join('|');
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
