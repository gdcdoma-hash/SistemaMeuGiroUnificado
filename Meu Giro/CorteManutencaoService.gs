/**
 * PROTECAO TEMPORARIA DE CORTE — Desafio Giro, 16/10/2026.
 *
 * ABA COMPARTILHADA na mesma planilha oficial, sem cache:
 *   'DGMB_CORTE_MANUTENCAO'!A1 = DGMB_CORTE_MODALIDADES_20261016
 *   'DGMB_CORTE_MANUTENCAO'!B1 = PAUSADO  (bloqueia)
 *   'DGMB_CORTE_MANUTENCAO'!B1 = LIBERADO (retoma)
 *
 * Aba ausente = operacao normal. Aba existente mas ilegivel/invalida = bloqueio
 * preventivo, inclusive em falha temporaria de acesso à planilha.
 * So protege as cinco APIs de inscricao/reativacao e KM/edicao/exclusao
 * explicitamente instrumentadas; operacoes Admin devem ficar suspensas.
 */
function dgmbCorteGravacaoPausada_() {
  try {
    var ss = SpreadsheetApp.openById(dgmbMeuGiroSpreadsheetId_());
    var sheet = ss.getSheetByName('DGMB_CORTE_MANUTENCAO');
    if (!sheet) return false;
    var vals = sheet.getRange('A1:B1').getDisplayValues()[0];
    var chave = String(vals[0] || '').trim();
    var status = String(vals[1] || '').trim().toUpperCase();
    if (chave !== 'DGMB_CORTE_MODALIDADES_20261016') {
      Logger.log('[CORTE] aba de controle invalida, bloquear novas gravacoes');
      return true;
    }
    if (status === 'LIBERADO') return false;
    if (status === 'PAUSADO') return true;
    Logger.log('[CORTE] status desconhecido, bloquear novas gravacoes');
    return true;
  } catch (erro) {
    Logger.log('[CORTE] nao foi possivel consultar a trava de escrita: ' + (erro && erro.message || erro));
    return true;
  }
}
