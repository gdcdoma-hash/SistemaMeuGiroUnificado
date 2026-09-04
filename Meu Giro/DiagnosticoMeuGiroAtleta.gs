/**
 * Diagnóstico temporário e somente-leitura do fluxo Meu Giro para um atleta.
 *
 * Execute no editor do Apps Script:
 *   diagnosticarMeuGiroAtleta('ID_DGMB_DO_ATLETA')
 *
 * A função não escreve em abas, não reconstrói resumo e não chama
 * atualizarMeuGiroResumo_ como correção. O getPainelUsuario real só é chamado
 * quando já existem linhas em MEU_GIRO_RESUMO para evitar o fallback que grava.
 */
function diagnosticarMeuGiroAtleta(idDgmb) {
  var id = normalizeText_(idDgmb);
  var relatorio = {
    id_dgmb: id,
    gerado_em: new Date().toISOString(),
    somente_leitura: true,
    dgmbDesafios: { total_vinculos: 0, ativos: 0, concluidos: 0, linhas: [] },
    MEU_GIRO_RESUMO: { total_linhas: 0, ativos: 0, concluidos: 0, linhas: [] },
    REGISTRO_KM: { total_atividades: 0, atividades: [], total_por_periodo_resumo: [] },
    buscarInscricaoPainelMGLeve_: null,
    getPainelUsuario: null,
    comparacao: {},
    diagnostico_textual: []
  };

  if (!id) {
    relatorio.diagnostico_textual.push('ID_DGMB não informado.');
    Logger.log(JSON.stringify(relatorio, null, 2));
    return relatorio;
  }

  relatorio.dgmbDesafios = diagnosticoMeuGiroLerDgmbDesafios_(id);
  relatorio.MEU_GIRO_RESUMO = diagnosticoMeuGiroLerResumo_(id);
  relatorio.REGISTRO_KM = diagnosticoMeuGiroLerRegistroKm_(id, relatorio.MEU_GIRO_RESUMO.linhas);

  var resumoLeve = obterMeuGiroResumoAtualizadoLeve_(id, { reconciliar: false }) || [];
  relatorio.MEU_GIRO_RESUMO.total_obterMeuGiroResumoAtualizadoLeve = resumoLeve.length;
  relatorio.MEU_GIRO_RESUMO.linhas_obterMeuGiroResumoAtualizadoLeve = diagnosticoMeuGiroProjetarDesafios_(resumoLeve);

  relatorio.buscarInscricaoPainelMGLeve_ = diagnosticoMeuGiroExecutarBuscaLeve_(id, resumoLeve);
  relatorio.getPainelUsuario = diagnosticoMeuGiroExecutarGetPainelUsuarioSeguro_(id, relatorio.MEU_GIRO_RESUMO.total_linhas);
  relatorio.comparacao = diagnosticoMeuGiroCompararFontes_(relatorio);
  relatorio.diagnostico_textual = diagnosticoMeuGiroGerarTexto_(relatorio);

  Logger.log(JSON.stringify(relatorio, null, 2));
  return relatorio;
}

function diagnosticoMeuGiroLerDgmbDesafios_(id) {
  var nomeAba = SHEETS.DESAFIO || 'dgmbDesafios';
  var saida = { total_vinculos: 0, ativos: 0, concluidos: 0, linhas: [] };
  var sh = getSpreadsheet_().getSheetByName(nomeAba);
  if (!sh) {
    saida.erro = 'Aba não encontrada: ' + nomeAba;
    return saida;
  }
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return saida;
  var map = buildHeaderMap_(values[0]);
  var idx = diagnosticoMeuGiroIndicesComuns_(map);
  idx.id = getOptionalColumnIndex_(map, ['id_dgmb']);
  idx.meta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km', 'meta_km', 'meta km']);
  idx.statusConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  idx.statusLista = getOptionalColumnIndex_(map, ['status_lista_desafios']);
  idx.periodo = getOptionalColumnIndex_(map, MEU_GIRO_PERIODO_DESAFIO_ALIASES_);
  idx.inicio = getOptionalColumnIndex_(map, ['data_inicio_desafio', 'data inicio desafio', 'data início desafio', 'periodo_inicio']);
  idx.fim = getOptionalColumnIndex_(map, ['data_fim_desafio', 'data fim desafio', 'periodo_fim']);
  if (idx.id === -1) return saida;

  var periodosLista = buildListaDesafiosContexto_(getSpreadsheet_()).periodos;

  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    if (normalizeText_(row[idx.id]) !== id) continue;
    var periodoTexto = idx.periodo > -1 ? extrairPeriodoDesafioTexto_(row[idx.periodo]) : { inicio: '', fim: '' };
    var periodoDatas = {
      inicio: normalizarDataISO_(idx.inicio > -1 ? row[idx.inicio] : ''),
      fim: normalizarDataISO_(idx.fim > -1 ? row[idx.fim] : '')
    };
    var statusUsuario = idx.statusUsuario > -1 ? normalizeText_(row[idx.statusUsuario]) : '';
    var statusPagamento = idx.statusPagamento > -1 ? normalizeText_(row[idx.statusPagamento]) : '';
    var statusConfirmacao = idx.statusConfirmacao > -1 ? normalizeText_(row[idx.statusConfirmacao]) : '';
    var validacao = validarInscricaoMinima_({ status_inscricao: statusUsuario, status_confirmacao: statusConfirmacao, status_pagamento: statusPagamento });
    var item = diagnosticoMeuGiroProjetarLinha_(row, idx, id);
    item.linha_planilha = i + 1;
    item.meta_km = idx.meta > -1 ? parseLocalizedNumber_(row[idx.meta]) : 0;
    item.status_lista_desafios = idx.statusLista > -1 ? normalizeText_(row[idx.statusLista]) : '';
    var periodoLista = (item.id_desafio && periodosLista.byId[item.id_desafio]) || { inicio: '', fim: '' };
    var periodoSelecionado = periodoCompletoValido_(periodoTexto)
      ? periodoTexto
      : periodoCompletoValido_(periodoLista)
        ? periodoLista
        : periodoDatas;
    item.periodo_inicio = periodoCompletoValido_(periodoSelecionado) ? periodoSelecionado.inicio : '';
    item.periodo_fim = periodoCompletoValido_(periodoSelecionado) ? periodoSelecionado.fim : '';
    item.apto_elegivel = validacao.valida;
    item.criterio_elegibilidade = validacao.criterio;
    saida.linhas.push(item);
  }
  saida.total_vinculos = saida.linhas.length;
  saida.ativos = saida.linhas.filter(function(x) { return diagnosticoMeuGiroEhAtivo_(x); }).length;
  saida.concluidos = saida.linhas.filter(function(x) { return diagnosticoMeuGiroEhConcluido_(x); }).length;
  return saida;
}

function diagnosticoMeuGiroLerResumo_(id) {
  var nomeAba = SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO';
  var saida = { total_linhas: 0, ativos: 0, concluidos: 0, linhas: [] };
  var sh = getSpreadsheet_().getSheetByName(nomeAba);
  if (!sh) return saida;
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return saida;
  var map = buildHeaderMap_(values[0]);
  var idx = diagnosticoMeuGiroIndicesComuns_(map);
  idx.id = getOptionalColumnIndex_(map, ['id_dgmb']);
  idx.meta = getOptionalColumnIndex_(map, ['meta_km', 'meta km']);
  idx.distancia = getOptionalColumnIndex_(map, ['distancia_realizada', 'distancia realizada']);
  idx.percentual = getOptionalColumnIndex_(map, ['percentual_concluido', 'percentual concluido', 'percentual concluído']);
  idx.inicio = getOptionalColumnIndex_(map, ['periodo_inicio', 'data_inicio_desafio', 'data inicio desafio']);
  idx.fim = getOptionalColumnIndex_(map, ['periodo_fim', 'data_fim_desafio', 'data fim desafio']);
  if (idx.id === -1) return saida;
  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    if (normalizeText_(row[idx.id]) !== id) continue;
    var item = diagnosticoMeuGiroProjetarLinha_(row, idx, id);
    item.linha_planilha = i + 1;
    item.meta_km = idx.meta > -1 ? parseLocalizedNumber_(row[idx.meta]) : 0;
    item.distancia_realizada = idx.distancia > -1 ? parseLocalizedNumber_(row[idx.distancia]) : 0;
    item.percentual_concluido = idx.percentual > -1 ? parseLocalizedNumber_(row[idx.percentual]) : 0;
    item.periodo_inicio = normalizarDataISO_(idx.inicio > -1 ? row[idx.inicio] : '');
    item.periodo_fim = normalizarDataISO_(idx.fim > -1 ? row[idx.fim] : '');
    saida.linhas.push(item);
  }
  saida.total_linhas = saida.linhas.length;
  saida.ativos = saida.linhas.filter(function(x) { return diagnosticoMeuGiroEhAtivo_(x); }).length;
  saida.concluidos = saida.linhas.filter(function(x) { return diagnosticoMeuGiroEhConcluido_(x); }).length;
  return saida;
}

function diagnosticoMeuGiroLerRegistroKm_(id, periodosResumo) {
  var saida = { total_atividades: 0, atividades: [], total_por_periodo_resumo: [] };
  var registros = obterRegistrosKmUsuario_(id) || [];
  saida.atividades = registros.map(function(r) {
    return {
      data: normalizarDataISO_(r.data || r.Data || r.data_atividade || r.DATA || ''),
      distancia: parseLocalizedNumber_(r.distancia_km || r.distancia || r.Distancia || r.DISTANCIA || 0),
      origem: r.origem || ''
    };
  });
  saida.total_atividades = saida.atividades.length;
  saida.total_por_periodo_resumo = (periodosResumo || []).map(function(d) {
    var total = 0;
    saida.atividades.forEach(function(a) {
      if (a.data && d.periodo_inicio && d.periodo_fim && a.data >= d.periodo_inicio && a.data <= d.periodo_fim) total += a.distancia;
    });
    return { id_inscricao: d.id_inscricao, id_desafio: d.id_desafio, periodo_inicio: d.periodo_inicio, periodo_fim: d.periodo_fim, distancia_no_periodo: Math.round((total + Number.EPSILON) * 10) / 10 };
  });
  return saida;
}

function diagnosticoMeuGiroExecutarBuscaLeve_(id, resumoLeve) {
  var r = buscarInscricaoPainelMGLeve_(id, resumoLeve || []);
  return {
    ok: !!(r && r.ok),
    code: r && r.code,
    motivo: r && (r.motivo || r.motivo_fallback),
    usou_fallback: !!(r && r.usou_fallback),
    desafio_em_foco: r && r.desafios ? painelMG_selecionarDesafioPrincipal_(r.desafios) : null,
    quantidade_desafios_retornada: r && r.desafios ? r.desafios.length : 0,
    desafios_retornados: diagnosticoMeuGiroProjetarDesafios_(r && r.desafios ? r.desafios : [])
  };
}

function diagnosticoMeuGiroExecutarGetPainelUsuarioSeguro_(id, totalResumo) {
  if (!totalResumo) {
    return { executado: false, motivo: 'Não executado para preservar somente-leitura: getPainelUsuario chama atualizarMeuGiroResumo_ quando MEU_GIRO_RESUMO está vazio.' };
  }
  var payload = getPainelUsuario(id, { somenteLeitura: true });
  var data = payload && payload.data ? payload.data : {};
  return {
    executado: true,
    ok: !!(payload && payload.ok),
    code: payload && payload.code,
    desafio_em_foco: data.desafio_em_foco || null,
    data_desafio_em_foco: data.desafio_em_foco || null,
    desafios_ativos: diagnosticoMeuGiroProjetarDesafios_(data.desafios_ativos || []),
    desafios_historico: diagnosticoMeuGiroProjetarDesafios_(data.desafios_historico || []),
    desafios_payload: diagnosticoMeuGiroProjetarDesafios_(data.desafios || []),
    quantidade_final_desafios_payload: data.desafios ? data.desafios.length : 0
  };
}

function diagnosticoMeuGiroCompararFontes_(relatorio) {
  var dgmbKeys = diagnosticoMeuGiroKeys_(relatorio.dgmbDesafios.linhas);
  var resumoKeys = diagnosticoMeuGiroKeys_(relatorio.MEU_GIRO_RESUMO.linhas);
  var payloadKeys = diagnosticoMeuGiroKeys_(relatorio.getPainelUsuario && relatorio.getPainelUsuario.desafios_payload || []);
  return {
    dgmbDesafios_vs_MEU_GIRO_RESUMO: diagnosticoMeuGiroDiff_(dgmbKeys, resumoKeys),
    MEU_GIRO_RESUMO_vs_payload_final: diagnosticoMeuGiroDiff_(resumoKeys, payloadKeys),
    REGISTRO_KM_usado_para_decidir_aparicao: 'Não no caminho leve: buscarInscricaoPainelMGLeve_ recebe MEU_GIRO_RESUMO e não consulta REGISTRO_KM; REGISTRO_KM só é lido depois por getPainelUsuario para atividades.'
  };
}

function diagnosticoMeuGiroGerarTexto_(r) {
  var textos = [];
  if (!r.dgmbDesafios.total_vinculos) textos.push('os desafios atuais não existem em dgmbDesafios ou não há vínculos para este ID_DGMB.');
  if (r.dgmbDesafios.total_vinculos && !r.MEU_GIRO_RESUMO.total_linhas) textos.push('existem em dgmbDesafios, mas não foram para MEU_GIRO_RESUMO.');
  if (r.MEU_GIRO_RESUMO.total_linhas && r.buscarInscricaoPainelMGLeve_ && r.buscarInscricaoPainelMGLeve_.quantidade_desafios_retornada < r.MEU_GIRO_RESUMO.total_linhas) textos.push('existem em MEU_GIRO_RESUMO, mas foram removidos em buscarInscricaoPainelMGLeve_ ou no fallback acionado por ela.');
  if (r.getPainelUsuario && r.getPainelUsuario.executado && r.getPainelUsuario.quantidade_final_desafios_payload === r.MEU_GIRO_RESUMO.total_linhas) textos.push('chegam no payload; se o front não mostra, o próximo ponto a auditar é o Script.html.');
  if (!textos.length) textos.push('Nenhum desaparecimento óbvio foi identificado; compare as chaves e status no bloco comparacao.');
  return textos;
}

function diagnosticoMeuGiroIndicesComuns_(map) {
  return {
    inscricao: getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']),
    desafio: getIdDesafioColumnIndex_(map),
    item: getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']),
    nome: getOptionalColumnIndex_(map, ['nome_desafio', 'nome desafio', 'desafio']),
    statusDesafio: getOptionalColumnIndex_(map, ['status_desafio', 'status desafio']),
    statusUsuario: getOptionalColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio']),
    statusPagamento: getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagamento', 'pix_status']),
    statusApuracao: getOptionalColumnIndex_(map, ['status_apuracao', 'status apuracao', 'status apuração']),
    obs: getOptionalColumnIndex_(map, ['observacao', 'observação'])
  };
}

function diagnosticoMeuGiroProjetarLinha_(row, idx, id) {
  return {
    id_dgmb: id,
    id_inscricao: idx.inscricao > -1 ? normalizeText_(row[idx.inscricao]) : '',
    id_desafio: obterIdDesafioRegistro_(row, idx.desafio, idx.obs),
    id_item_estoque: idx.item > -1 ? normalizeText_(row[idx.item]) : '',
    nome_desafio: idx.nome > -1 ? normalizeText_(row[idx.nome]) : '',
    status_desafio: idx.statusDesafio > -1 ? normalizeText_(row[idx.statusDesafio]) : '',
    status_usuario_desafio: idx.statusUsuario > -1 ? normalizeText_(row[idx.statusUsuario]) : '',
    status_pagamento: idx.statusPagamento > -1 ? normalizeText_(row[idx.statusPagamento]) : '',
    status_apuracao: idx.statusApuracao > -1 ? normalizeText_(row[idx.statusApuracao]) : ''
  };
}

function diagnosticoMeuGiroProjetarDesafios_(lista) {
  return (lista || []).map(function(d) {
    return {
      id_inscricao: d.id_inscricao || '',
      id_desafio: d.id_desafio || '',
      id_item_estoque: d.id_item_estoque || '',
      nome_desafio: d.nome_desafio || '',
      meta_km: d.meta_km,
      distancia_realizada: d.distancia_realizada,
      percentual_concluido: d.percentual_concluido,
      status_apuracao: d.status_apuracao || '',
      status_usuario_desafio: d.status_usuario_desafio || '',
      periodo_inicio: d.periodo_inicio || '',
      periodo_fim: d.periodo_fim || ''
    };
  });
}

function diagnosticoMeuGiroEhAtivo_(d) {
  return normalizeText_(d.status_apuracao).toUpperCase() === 'ATIVO' || painelMG_isDesafioAtivoParaFoco_(d);
}

function diagnosticoMeuGiroEhConcluido_(d) {
  return normalizeText_(d.status_apuracao).toUpperCase() === 'CONCLUIDO' || normalizeText_(d.status_usuario_desafio).toUpperCase() === 'CONCLUIDO';
}

function diagnosticoMeuGiroKeys_(lista) {
  var out = {};
  (lista || []).forEach(function(d) {
    var key = [d.id_inscricao || '', d.id_desafio || '', d.id_item_estoque || ''].join('|');
    out[key] = d;
  });
  return out;
}

function diagnosticoMeuGiroDiff_(origem, destino) {
  var ausentes = [];
  Object.keys(origem || {}).forEach(function(k) {
    if (!destino[k]) ausentes.push(k);
  });
  return { origem_total: Object.keys(origem || {}).length, destino_total: Object.keys(destino || {}).length, ausentes_no_destino: ausentes };
}
