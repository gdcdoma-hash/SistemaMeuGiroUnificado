/**
 * Simula, somente em memória, a reconstrução completa da aba MEU_GIRO_RESUMO.
 *
 * Classificações do comparativo:
 * - corretas: primeira ocorrência da chave atual igual ao resumo esperado;
 * - a criar: chave esperada inexistente no resumo atual;
 * - a atualizar: primeira ocorrência existente com valores divergentes;
 * - órfãs: chave atual sem vínculo correspondente no resumo esperado;
 * - legadas: linhas atuais que ainda dependem da chave sem ID_INSCRICAO;
 * - duplicadas: ocorrências adicionais de uma mesma chave no resumo atual.
 *
 * Esta função é deliberadamente somente leitura e não altera nenhuma aba.
 */
function simularReconstrucaoMeuGiroResumo() {
  var ss = getSpreadsheet_();
  var nomes = {
    desafios: SHEETS.DESAFIO || 'dgmbDesafios',
    registros: SHEETS.REGISTRO_KM || 'REGISTRO_KM',
    lista: SHEETS.LISTA_DESAFIOS || 'ListaDesafios',
    resumo: SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO'
  };

  var dadosDesafios = simularResumoLerAba_(ss, nomes.desafios);
  var dadosRegistros = simularResumoLerAba_(ss, nomes.registros);
  var dadosLista = simularResumoLerAba_(ss, nomes.lista);
  var dadosResumo = simularResumoLerAba_(ss, nomes.resumo);
  var layoutResumo = meuGiroResumoObterLayout_(dadosResumo[0] || [], nomes.resumo);
  var esperado = simularResumoCalcularEsperado_(
    dadosDesafios,
    dadosRegistros,
    dadosLista,
    layoutResumo.possuiIdInscricao
  );
  var comparativo = simularResumoComparar_(esperado, dadosResumo, layoutResumo);

  Logger.log('[Meu Giro][simulação reconstrução] total esperado: ' + comparativo.total_esperado);
  Logger.log('[Meu Giro][simulação reconstrução] total atual: ' + comparativo.total_atual);
  Logger.log('[Meu Giro][simulação reconstrução] linhas corretas: ' + comparativo.linhas_corretas);
  Logger.log('[Meu Giro][simulação reconstrução] linhas a criar: ' + comparativo.linhas_a_criar);
  Logger.log('[Meu Giro][simulação reconstrução] linhas a atualizar: ' + comparativo.linhas_a_atualizar);
  Logger.log('[Meu Giro][simulação reconstrução] órfãs: ' + comparativo.orfas);
  Logger.log('[Meu Giro][simulação reconstrução] legadas: ' + comparativo.legadas);
  Logger.log('[Meu Giro][simulação reconstrução] duplicadas: ' + comparativo.duplicadas);

  return comparativo;
}

/**
 * Reconstrói o resumo em uma aba temporária para validação, sem tocar na aba
 * oficial MEU_GIRO_RESUMO ou nas abas de origem.
 *
 * As únicas operações de escrita desta rotina são clearContents() e
 * setValues() na aba MEU_GIRO_RESUMO_REBUILD_TESTE. A aba é criada somente
 * quando ainda não existe.
 */
function reconstruirMeuGiroResumoEmAbaTeste() {
  var nomeAbaTeste = 'MEU_GIRO_RESUMO_REBUILD_TESTE';
  var cabecalho = [
    'Timestamp_Atualizacao',
    'ID_INSCRICAO',
    'ID_DGMB',
    'ID_DESAFIO',
    'id_item_estoque',
    'Meta_KM',
    'Distancia_Realizada',
    'Percentual_Concluido',
    'Status_Apuracao'
  ];
  var ss = getSpreadsheet_();
  var dadosDesafios = simularResumoLerAba_(ss, SHEETS.DESAFIO || 'dgmbDesafios');
  var dadosRegistros = simularResumoLerAba_(ss, SHEETS.REGISTRO_KM || 'REGISTRO_KM');
  var dadosLista = simularResumoLerAba_(ss, SHEETS.LISTA_DESAFIOS || 'ListaDesafios');

  // O layout da aba de teste é sempre canônico e inclui ID_INSCRICAO.
  var esperado = simularResumoCalcularEsperado_(
    dadosDesafios,
    dadosRegistros,
    dadosLista,
    true
  );
  var itens = [];

  for (var i = 0; i < esperado.chaves.length; i++) {
    itens.push(esperado.porChave[esperado.chaves[i]]);
  }

  itens.sort(reconstruirResumoCompararItens_);

  var timestampAtualizacao = new Date();
  var valores = [cabecalho];
  var relatorio = {
    total_linhas_gravadas: itens.length,
    total_com_id_inscricao: 0,
    total_sem_id_inscricao: 0,
    total_concluido: 0,
    total_status_em_analise: 0
  };

  for (var j = 0; j < itens.length; j++) {
    var item = itens[j];
    var idInscricao = normalizeText_(item.id_inscricao);
    var status = normalizeText_(item.status_apuracao).toUpperCase();

    if (idInscricao) {
      relatorio.total_com_id_inscricao++;
    } else {
      relatorio.total_sem_id_inscricao++;
    }

    if (status === 'CONCLUIDO') relatorio.total_concluido++;
    if (status === 'STATUS_EM_ANALISE') relatorio.total_status_em_analise++;

    valores.push([
      timestampAtualizacao,
      idInscricao,
      item.id_dgmb,
      item.id_desafio,
      item.id_item_estoque,
      item.meta_km,
      item.distancia_realizada,
      item.percentual_concluido,
      status
    ]);
  }

  var abaTeste = ss.getSheetByName(nomeAbaTeste);
  if (!abaTeste) abaTeste = ss.insertSheet(nomeAbaTeste);

  abaTeste.clearContents();
  abaTeste.getRange(1, 1, valores.length, cabecalho.length).setValues(valores);

  Logger.log('[Meu Giro][rebuild teste] total de linhas gravadas: ' + relatorio.total_linhas_gravadas);
  Logger.log('[Meu Giro][rebuild teste] total com ID_INSCRICAO: ' + relatorio.total_com_id_inscricao);
  Logger.log('[Meu Giro][rebuild teste] total sem ID_INSCRICAO: ' + relatorio.total_sem_id_inscricao);
  Logger.log('[Meu Giro][rebuild teste] total CONCLUIDO: ' + relatorio.total_concluido);
  Logger.log('[Meu Giro][rebuild teste] total STATUS_EM_ANALISE: ' + relatorio.total_status_em_analise);

  return relatorio;
}

function reconstruirResumoCompararItens_(a, b) {
  var camposTexto = ['id_dgmb', 'id_desafio', 'id_item_estoque'];

  for (var i = 0; i < camposTexto.length; i++) {
    var comparacao = normalizeText_(a[camposTexto[i]]).localeCompare(
      normalizeText_(b[camposTexto[i]])
    );
    if (comparacao !== 0) return comparacao;
  }

  var diferencaMeta = Number(a.meta_km || 0) - Number(b.meta_km || 0);
  if (diferencaMeta !== 0) return diferencaMeta;

  return normalizeText_(a.id_inscricao).localeCompare(normalizeText_(b.id_inscricao));
}

function simularResumoLerAba_(ss, nomeAba) {
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) throw new Error('Aba não encontrada: ' + nomeAba);
  return sheet.getDataRange().getValues();
}

function simularResumoCalcularEsperado_(dadosDesafios, dadosRegistros, dadosLista, possuiIdInscricao) {
  var periodos = simularResumoBuildPeriodos_(dadosLista);
  var statusLista = simularResumoBuildStatusLista_(dadosLista);
  var registrosPorId = simularResumoBuildRegistrosPorId_(dadosRegistros);
  var vinculos = simularResumoBuildVinculos_(dadosDesafios, periodos, statusLista);
  var esperadoPorChave = {};
  var ordemChaves = [];

  for (var i = 0; i < vinculos.length; i++) {
    var vinculo = vinculos[i];
    var meta = Number(vinculo.meta_km || 0);
    var metaArredondada = simularResumoArredondar_(meta);
    var inicio = normalizarDataISO_(vinculo.periodo_inicio);
    var fim = normalizarDataISO_(vinculo.periodo_fim);
    var apto = !!vinculo.apto && !!inicio && !!fim && !!vinculo.id_desafio;
    var distancia = 0;
    var registros = registrosPorId[vinculo.id_dgmb] || [];

    if (apto) {
      for (var r = 0; r < registros.length; r++) {
        if (atividadeDentroPeriodoOficial_(registros[r].data_atividade, inicio, fim)) {
          distancia += Number(registros[r].km || 0);
        }
      }
    }

    var percentual = meta > 0 ? (distancia / meta) * 100 : 0;
    var percentualArredondado = simularResumoArredondar_(percentual);
    var status = calcularStatusMeuGiroPorPercentual_(percentualArredondado);

    var item = {
      id_inscricao: normalizeText_(vinculo.id_inscricao),
      id_dgmb: normalizeText_(vinculo.id_dgmb),
      id_desafio: normalizeText_(vinculo.id_desafio),
      id_item_estoque: normalizeText_(vinculo.id_item_estoque),
      meta_km: metaArredondada,
      distancia_realizada: simularResumoArredondar_(distancia),
      percentual_concluido: percentualArredondado,
      status_apuracao: status
    };
    var chave = meuGiroResumoBuildChave_(
      item.id_dgmb,
      item.id_desafio,
      item.id_item_estoque,
      item.meta_km,
      possuiIdInscricao ? item.id_inscricao : ''
    );

    if (!Object.prototype.hasOwnProperty.call(esperadoPorChave, chave)) {
      ordemChaves.push(chave);
    }
    esperadoPorChave[chave] = item;
  }

  return {
    porChave: esperadoPorChave,
    chaves: ordemChaves
  };
}

function simularResumoBuildPeriodos_(dadosLista) {
  var out = { byAba: {}, byId: {} };
  if (!dadosLista || dadosLista.length < 2) return out;

  var map = buildHeaderMap_(dadosLista[0]);
  var idxAba = getOptionalColumnIndex_(map, ['aba', 'aba desafio', 'abadesafio']);
  var idxId = getOptionalColumnIndex_(map, [
    'id', 'id_desafio', 'id desafio', 'id_desafio_lista', 'id desafio lista',
    'id_desafio_base', 'id desafio base'
  ]);
  var idxPeriodo = getOptionalColumnIndex_(map, ['periodo', 'período']);
  var idxNome = getOptionalColumnIndex_(map, [
    'nome_desafio', 'nome desafio', 'nome_desafio_lista', 'nome desafio lista',
    'desafio', 'nome'
  ]);
  if (idxAba === -1) idxAba = 1;

  for (var i = 1; i < dadosLista.length; i++) {
    var row = dadosLista[i];
    var aba = normalizeText_(row[idxAba]);
    if (!aba) continue;

    var periodoMensal = idxPeriodo > -1
      ? normalizarPeriodoMensal_(row[idxPeriodo])
      : { inicio: '', fim: '' };
    var periodo = {
      inicio: periodoMensal.inicio,
      fim: periodoMensal.fim,
      periodo_desafio: idxPeriodo > -1 ? normalizeText_(row[idxPeriodo]) : '',
      nome_desafio: (idxNome > -1 ? normalizeText_(row[idxNome]) : '') || aba
    };
    out.byAba[aba] = periodo;

    if (idxId > -1) {
      var idDesafio = normalizeText_(row[idxId]);
      if (idDesafio) out.byId[idDesafio] = periodo;
    }
  }

  return out;
}

function simularResumoBuildStatusLista_(dadosLista) {
  var out = {};
  if (!dadosLista || dadosLista.length < 2) return out;

  var map = buildHeaderMap_(dadosLista[0]);
  var idxId = getOptionalColumnIndex_(map, [
    'id_desafio_lista', 'id desafio lista', 'id_desafio', 'id desafio', 'id'
  ]);
  var idxStatus = getOptionalColumnIndex_(map, [
    'status', 'status_desafio', 'status desafio', 'status_lista', 'situacao', 'situação'
  ]);
  if (idxId === -1 || idxStatus === -1) return out;

  for (var i = 1; i < dadosLista.length; i++) {
    var idDesafio = normalizeText_(dadosLista[i][idxId]);
    if (idDesafio) out[idDesafio] = normalizeText_(dadosLista[i][idxStatus]).toLowerCase();
  }

  return out;
}

function simularResumoBuildRegistrosPorId_(dadosRegistros) {
  var out = {};
  var activityIdsPorId = {};
  if (!dadosRegistros || dadosRegistros.length < 2) return out;

  var map = buildHeaderMap_(dadosRegistros[0]);
  var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
  var idxData = getOptionalColumnIndex_(map, ['data_atividade', 'data atividade', 'data']);
  var idxKm = getOptionalColumnIndex_(map, ['km']);
  var idxActivity = getOptionalColumnIndex_(map, [
    'activity_id', 'activity id', 'id_atividade', 'id atividade'
  ]);
  if (idxId === -1) return out;

  for (var i = 1; i < dadosRegistros.length; i++) {
    var row = dadosRegistros[i];
    var id = normalizeText_(row[idxId]);
    if (!id) continue;

    var activityId = idxActivity > -1 ? normalizeText_(row[idxActivity]) : '';
    if (activityId) {
      if (!activityIdsPorId[id]) activityIdsPorId[id] = {};
      if (activityIdsPorId[id][activityId]) continue;
      activityIdsPorId[id][activityId] = true;
    }

    if (!out[id]) out[id] = [];
    out[id].push({
      data_atividade: idxData > -1 ? normalizarDataISO_(row[idxData]) : '',
      km: idxKm > -1 ? parseLocalizedNumber_(row[idxKm]) : 0
    });
  }

  return out;
}

function simularResumoBuildVinculos_(dadosDesafios, periodos, statusLista) {
  var vinculos = [];
  var chaves = {};
  if (!dadosDesafios || dadosDesafios.length < 2) return vinculos;

  var map = buildHeaderMap_(dadosDesafios[0]);
  var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
  if (idxId === -1) return vinculos;

  var idxMeta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km']);
  var idxInscricao = getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxIdDesafio = getIdDesafioColumnIndex_(map);
  var idxObs = getOptionalColumnIndex_(map, ['observacao', 'observação']);
  var idxItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxTipo = getOptionalColumnIndex_(map, ['tipo_do_desafio', 'tipo do desafio', 'tipo_desafio', 'tipo desafio']);
  var idxStatusUsuario = getOptionalColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio']);
  var idxStatusPagamento = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagamento', 'pix_status']);
  var idxStatusInscricao = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  var idxPeriodo = getOptionalColumnIndex_(map, ['periodo_desafio', 'periodo desafio', 'período_desafio', 'período desafio']);
  var idxInicio = getOptionalColumnIndex_(map, ['data_inicio_desafio', 'data inicio desafio', 'data início desafio']);
  var idxFim = getOptionalColumnIndex_(map, ['data_fim_desafio', 'data fim desafio']);
  var abaDesafio = SHEETS.DESAFIO || 'dgmbDesafios';

  for (var i = 1; i < dadosDesafios.length; i++) {
    var row = dadosDesafios[i];
    var id = normalizeText_(row[idxId]);
    if (!id) continue;

    var idDesafio = obterIdDesafioRegistro_(row, idxIdDesafio, idxObs);
    var idInscricao = idxInscricao > -1 ? normalizeText_(row[idxInscricao]) : '';
    var idItem = idxItem > -1 ? normalizeText_(row[idxItem]) : '';
    var metaKm = idxMeta > -1 ? parseLocalizedNumber_(row[idxMeta]) : 0;
    var tipo = idxTipo > -1 ? normalizeText_(row[idxTipo]) : '';
    var tipoSemAcento = tipo.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    var ehNormal = tipoSemAcento === 'normal';
    var statusUsuario = idxStatusUsuario > -1 ? normalizeText_(row[idxStatusUsuario]) : '';
    var validacao = validarInscricaoMinima_({
      status_inscricao: (idxStatusInscricao > -1 ? normalizeText_(row[idxStatusInscricao]) : '') || statusUsuario,
      status_confirmacao: idxConfirmacao > -1 ? normalizeText_(row[idxConfirmacao]) : '',
      status_pagamento: idxStatusPagamento > -1 ? normalizeText_(row[idxStatusPagamento]) : ''
    });
    var aptoBase = validacao.valida && !inscricaoTemBloqueioMinimo_(statusUsuario);
    var apto = ehNormal ? aptoBase && !!idDesafio && metaKm > 0 : aptoBase;
    var periodoLista = (idDesafio && periodos.byId[idDesafio]) ||
      (!ehNormal && periodos.byAba[abaDesafio]) ||
      { inicio: '', fim: '', periodo_desafio: '', nome_desafio: '' };
    var periodo = simularResumoMontarPeriodo_(row, {
      periodo: idxPeriodo,
      inicio: idxInicio,
      fim: idxFim
    }, periodoLista);
    var chaveVinculo = [
      id,
      idInscricao,
      idDesafio,
      idItem || ('META_' + simularResumoArredondar_(metaKm))
    ].join('|');
    if (chaves[chaveVinculo]) continue;
    chaves[chaveVinculo] = true;

    vinculos.push({
      id_dgmb: id,
      id_inscricao: idInscricao,
      id_desafio: idDesafio,
      id_item_estoque: idItem,
      meta_km: metaKm,
      apto: apto,
      periodo_inicio: periodo.inicio || '',
      periodo_fim: periodo.fim || '',
      status_lista_desafios: idDesafio ? (statusLista[idDesafio] || '') : ''
    });
  }

  return vinculos;
}

function simularResumoMontarPeriodo_(row, indices, periodoLista) {
  var periodoTexto = indices.periodo > -1 ? normalizeText_(row[indices.periodo]) : '';
  var periodoTextoEspecifico = extrairPeriodoDesafioTexto_(periodoTexto);
  var periodoDatasEspecificas = {
    inicio: indices.inicio > -1 ? normalizarDataISO_(row[indices.inicio]) : '',
    fim: indices.fim > -1 ? normalizarDataISO_(row[indices.fim]) : ''
  };
  var periodo = { inicio: '', fim: '' };

  if (periodoCompletoValido_(periodoDatasEspecificas)) {
    periodo = periodoDatasEspecificas;
  } else if (periodoCompletoValido_(periodoTextoEspecifico)) {
    periodo = periodoTextoEspecifico;
  } else if (periodoCompletoValido_(periodoLista)) {
    periodo = periodoLista;
  }

  return {
    inicio: periodo.inicio || '',
    fim: periodo.fim || ''
  };
}

function simularResumoComparar_(esperado, dadosResumo, layoutResumo) {
  var map = layoutResumo.map;
  var idxInscricao = getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
  var idxDesafio = getOptionalColumnIndex_(map, ['id_desafio']);
  var idxItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxMeta = getOptionalColumnIndex_(map, ['meta_km', 'meta km']);
  var idxDistancia = getOptionalColumnIndex_(map, ['distancia_realizada', 'distancia realizada']);
  var idxPercentual = getOptionalColumnIndex_(map, ['percentual_concluido', 'percentual concluido', 'percentual concluído']);
  var idxStatus = getOptionalColumnIndex_(map, ['status_apuracao', 'status apuracao', 'status apuração']);
  var atuaisPorChave = {};
  var totalAtual = 0;
  var legadas = 0;

  for (var i = 1; i < dadosResumo.length; i++) {
    var row = dadosResumo[i];
    if (!simularResumoLinhaPreenchida_(row)) continue;

    totalAtual++;
    var idInscricao = idxInscricao > -1 ? normalizeText_(row[idxInscricao]) : '';
    if (!idInscricao) legadas++;

    var chave = meuGiroResumoBuildChave_(
      row[idxId],
      row[idxDesafio],
      row[idxItem],
      row[idxMeta],
      idInscricao
    );
    if (!atuaisPorChave[chave]) atuaisPorChave[chave] = [];
    atuaisPorChave[chave].push(row);
  }

  var corretas = 0;
  var criar = 0;
  var atualizar = 0;
  var duplicadas = 0;
  var orfas = 0;

  for (var c = 0; c < esperado.chaves.length; c++) {
    var chaveEsperada = esperado.chaves[c];
    var linhasAtuais = atuaisPorChave[chaveEsperada] || [];
    if (!linhasAtuais.length) {
      criar++;
      continue;
    }

    if (simularResumoLinhaIgual_(linhasAtuais[0], esperado.porChave[chaveEsperada], {
      inscricao: idxInscricao,
      id: idxId,
      desafio: idxDesafio,
      item: idxItem,
      meta: idxMeta,
      distancia: idxDistancia,
      percentual: idxPercentual,
      status: idxStatus
    })) {
      corretas++;
    } else {
      atualizar++;
    }
    if (linhasAtuais.length > 1) duplicadas += linhasAtuais.length - 1;
  }

  for (var chaveAtual in atuaisPorChave) {
    if (!Object.prototype.hasOwnProperty.call(atuaisPorChave, chaveAtual)) continue;
    if (Object.prototype.hasOwnProperty.call(esperado.porChave, chaveAtual)) continue;
    orfas += atuaisPorChave[chaveAtual].length;
    if (atuaisPorChave[chaveAtual].length > 1) {
      duplicadas += atuaisPorChave[chaveAtual].length - 1;
    }
  }

  return {
    total_esperado: esperado.chaves.length,
    total_atual: totalAtual,
    linhas_corretas: corretas,
    linhas_a_criar: criar,
    linhas_a_atualizar: atualizar,
    orfas: orfas,
    legadas: legadas,
    duplicadas: duplicadas
  };
}

function simularResumoLinhaIgual_(row, esperado, indices) {
  return (indices.inscricao === -1 || normalizeText_(row[indices.inscricao]) === esperado.id_inscricao) &&
    normalizeText_(row[indices.id]) === esperado.id_dgmb &&
    normalizeText_(row[indices.desafio]) === esperado.id_desafio &&
    normalizeText_(row[indices.item]) === esperado.id_item_estoque &&
    parseLocalizedNumber_(row[indices.meta]) === esperado.meta_km &&
    parseLocalizedNumber_(row[indices.distancia]) === esperado.distancia_realizada &&
    parseLocalizedNumber_(row[indices.percentual]) === esperado.percentual_concluido &&
    normalizeText_(row[indices.status]) === esperado.status_apuracao;
}

function simularResumoLinhaPreenchida_(row) {
  for (var i = 0; i < row.length; i++) {
    if (normalizeText_(row[i])) return true;
  }
  return false;
}

function simularResumoArredondar_(valor) {
  return Math.round((Number(valor || 0) + Number.EPSILON) * 10) / 10;
}
