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
