function portalAdminMeuGiroConsultarAtleta(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) {
    return { ok: false, code: 'ID_DGMB_OBRIGATORIO', msg: 'ID_DGMB obrigatorio.' };
  }

  var painel = getPainelUsuario(id, { somenteLeitura: true });
  if (!painel || painel.ok === false) {
    return painel || { ok: false, code: 'PAINEL_NAO_DISPONIVEL', msg: 'Painel nao disponivel.' };
  }

  return {
    ok: true,
    id_dgmb: id,
    atleta: painel.atleta || {},
    desafio_em_foco: painel.desafio_em_foco || null,
    desafios: painel.desafios || [],
    desafios_ativos: painel.desafios_ativos || [],
    desafios_historico: painel.desafios_historico || [],
    atividades: painel.atividades || [],
    total_pedalado: painel.total_pedalado || 0,
    somente_leitura: true
  };
}
