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
  return meuGiroModalidadeRegistro_(atividade) ===
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
 * Validação autoritativa de data na gravação/edição.
 * O atleta pode registrar uma atividade se E SOMENTE SE houver pelo menos
 * uma inscrição elegível da mesma modalidade cuja janela contenha a data.
 * Janelas podem atravessar meses, sem agrupamento por competência mensal.
 */
function meuGiroValidarDataModalidadeInscricao_(dataAtividade, modalidade, vinculos) {
  var data = normalizarDataISO_(dataAtividade);
  if (!isDataIsoValida_(data)) {
    return {ok:false,code:'DATA_INVALIDA',msg:'Informe uma data válida.'};
  }
  var esporte = meuGiroModalidadeCanonica_(modalidade);
  var elegiveis = (vinculos || []).filter(function(v) {
    if (!v || !v.apto) return false;
    if (meuGiroModalidadeCanonica_(v.modalidade_esportiva || v.modalidade || '') !== esporte)
      return false;
    var inicio = normalizarDataISO_(v.periodo_inicio);
    var fim = normalizarDataISO_(v.periodo_fim);
    return isDataIsoValida_(inicio) && isDataIsoValida_(fim) && inicio <= fim;
  });
  if (!elegiveis.length) {
    return {ok:false,code:'SEM_DESAFIO_MODALIDADE',
      msg:'Não há desafio disponível nesta modalidade com período válido para registrar atividades.'};
  }
  var aceito = elegiveis.some(function(v) {
    return atividadeDentroPeriodoOficial_(
      data, normalizarDataISO_(v.periodo_inicio), normalizarDataISO_(v.periodo_fim));
  });
  if (!aceito) {
    return {ok:false,code:'DATA_FORA_DO_PRAZO',
      msg:'Esta data está fora do período dos seus desafios desta modalidade. Confira o início e a data limite.'};
  }
  return {ok:true,modalidade_esportiva:esporte,data_atividade:data};
}
