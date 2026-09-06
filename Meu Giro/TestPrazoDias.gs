function assertPrazoDias_(condicao, mensagem) {
  if (!condicao) throw new Error('[PRAZO_DIAS][TESTE] ' + mensagem);
}

function testarPrazoDiasAceitacao_() {
  var periodoLista = {
    inicio: '2026-09-01',
    fim: '2026-09-30',
    periodo_desafio: 'Setembro/2026',
    nome_desafio: 'DESAFIO 1000 KM',
    tipo_meta: 'PRAZO_DIAS'
  };
  var row = ['', 'Setembro/2026', new Date(2026, 8, 6), new Date(2026, 9, 25)];
  var periodo = montarPeriodoHistoricoVinculo_(
    row,
    { periodo: 1, inicio: 2, fim: 3 },
    periodoLista,
    { id_dgmb: 'TESTE', id_desafio: '1000K', id_inscricao: 'INSCRICAO_TESTE' },
    'PRAZO_DIAS'
  );

  assertPrazoDias_(periodo.inicio === '2026-09-06', 'PRAZO_DIAS deve priorizar data_inicio_desafio.');
  assertPrazoDias_(periodo.fim === '2026-10-25', 'PRAZO_DIAS deve priorizar data_fim_desafio.');

  var atividades = [
    { data_atividade: '2026-09-01', km: 34 },
    { data_atividade: '2026-09-06', km: 20 },
    { data_atividade: '2026-09-07', km: 30 },
    { data_atividade: '2026-10-26', km: 40 }
  ];
  var realizado = atividades.reduce(function(total, atividade) {
    return total + (atividadeDentroPeriodoOficial_(atividade.data_atividade, periodo.inicio, periodo.fim) ? atividade.km : 0);
  }, 0);

  assertPrazoDias_(realizado === 50, 'Cenário principal deve totalizar 50 km.');
  assertPrazoDias_(!atividadeDentroPeriodoOficial_('2026-09-01', periodo.inicio, periodo.fim), 'Atividade anterior não pode contar.');
  assertPrazoDias_(atividadeDentroPeriodoOficial_('2026-09-06', periodo.inicio, periodo.fim), 'Primeiro dia deve contar.');
  assertPrazoDias_(atividadeDentroPeriodoOficial_('2026-10-25', periodo.inicio, periodo.fim), 'Último dia deve contar.');
  assertPrazoDias_(!atividadeDentroPeriodoOficial_('2026-10-26', periodo.inicio, periodo.fim), 'Dia posterior ao fim não pode contar.');

  var naoConsolidado = montarPeriodoHistoricoVinculo_(
    ['', 'Setembro/2026', '', ''],
    { periodo: 1, inicio: 2, fim: 3 },
    periodoLista,
    { id_dgmb: 'TESTE', id_desafio: '1000K', id_inscricao: 'NAO_CONSOLIDADO' },
    'PRAZO_DIAS'
  );
  assertPrazoDias_(!naoConsolidado.inicio && !naoConsolidado.fim, 'PRAZO_DIAS não consolidado não deve herdar a janela mensal.');

  var tradicional = montarPeriodoHistoricoVinculo_(
    ['', 'Setembro/2026', new Date(2026, 8, 6), new Date(2026, 9, 25)],
    { periodo: 1, inicio: 2, fim: 3 },
    { inicio: '2026-09-01', fim: '2026-09-30', periodo_desafio: 'Setembro/2026', nome_desafio: 'NORMAL', tipo_meta: 'KM' },
    { id_dgmb: 'TESTE', id_desafio: 'NORMAL' },
    'KM'
  );
  assertPrazoDias_(tradicional.inicio === '2026-09-01' && tradicional.fim === '2026-09-30', 'Desafio tradicional deve preservar período mensal.');

  assertPrazoDias_(ehTipoMetaPrazoDias_('PRAZO_DIAS'), 'Marcador PRAZO_DIAS deve ser reconhecido.');
  assertPrazoDias_(!ehTipoMetaPrazoDias_('KM'), 'Tipo KM não pode ser tratado como PRAZO_DIAS.');

  return {
    ok: true,
    realizado: realizado,
    faltam: 950,
    percentual: 5,
    periodo_inicio: periodo.inicio,
    periodo_fim: periodo.fim
  };
}
