/**
 * LAB — classificação esportiva do Meu Giro.
 * Uma atividade pertence à modalidade real, não a uma inscrição específica.
 * Registros e inscrições anteriores à migração (sem coluna) são Ciclismo.
 * Caminhada e corrida compõem a mesma modalidade esportiva.
 */
function meuGiroModalidadeCanonica_(valor) {
  var texto = String(valor == null ? '' : valor).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s\-\/]+/g,'_');
  if (!texto || texto === 'CICLISMO' || texto === 'PEDAL' ||
      texto === 'BIKE' || texto === 'RIDE' || texto === 'VIRTUAL_RIDE' ||
      texto === 'CYCLING') return 'CICLISMO';
  if (texto === 'CAMINHADA_CORRIDA' || texto === 'CAMINHADA' ||
      texto === 'CORRIDA' || texto === 'WALK' || texto === 'RUN' ||
      texto === 'VIRTUAL_RUN' || texto === 'HIKE') return 'CAMINHADA_CORRIDA';
  throw new Error('Modalidade esportiva não reconhecida: ' + texto);
}
function meuGiroModalidadeRegistro_(registro) {
  var r = registro || {};
  var valor = r.Modalidade_Esportiva || r.modalidade_esportiva ||
    r.modalidadeEsportiva || r.Modalidade || r.modalidade || '';
  return meuGiroModalidadeCanonica_(valor);
}
function meuGiroAtividadeNoVinculo_(registro, vinculo) {
  var atividade = registro || {}, desafio = vinculo || {};
  // Registros futuros legados/testes não entram no progresso antes do dia.
  var dataAtividade = normalizarDataISO_(atividade.data_atividade || atividade.data);
  return dataAtividade <= meuGiroHojeBrasilISO_() &&
    meuGiroModalidadeRegistro_(atividade) ===
    meuGiroModalidadeCanonica_(desafio.modalidade_esportiva || desafio.modalidade || '') &&
    atividadeDentroPeriodoOficial_(
      atividade.data_atividade || atividade.data,
      desafio.periodo_inicio,
      desafio.periodo_fim
    );
}
function meuGiroSomarKmPorVinculo_(atividades, vinculo) {
  var total = 0;
  (atividades || []).forEach(function(reg) {
    if (meuGiroAtividadeNoVinculo_(reg, vinculo)) total += Number(reg.km || 0);
  });
  return Math.round((total + Number.EPSILON) * 100) / 100;
}

/**
 * Política de histórico: registro manual pode ser feito em qualquer data
 * passada ou em hoje, mesmo fora do prazo de uma inscrição.
 * A contagem para metas é independente: só entra quando modalidade e janela
 * da inscrição correspondem. Proíbe datas futuras no servidor.
 */
function meuGiroHojeBrasilISO_() {
  return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
}
function meuGiroValidarDataModalidadeInscricao_(dataAtividade, modalidade, vinculos) {
  var data = normalizarDataISO_(dataAtividade);
  if (!isDataIsoValida_(data)) {
    return {ok:false,code:'DATA_INVALIDA',msg:'Informe uma data válida para a atividade.'};
  }
  var esporte = meuGiroModalidadeCanonica_(modalidade);
  var hoje = meuGiroHojeBrasilISO_();
  if (data > hoje) {
    return {ok:false,code:'DATA_FUTURA',
      msg:'Não é permitido registrar atividades com data futura. Informe hoje ou uma data anterior.'};
  }
  // Não exigir desafio aberto ou dentro do prazo: histórico pessoal pertence
  // ao atleta e pode ser preenchido depois do encerramento do desafio.
  return {ok:true,modalidade_esportiva:esporte,data_atividade:data};
}
