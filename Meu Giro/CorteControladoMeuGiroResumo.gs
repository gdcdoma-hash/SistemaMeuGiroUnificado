/**
 * Substitui o conteúdo da aba oficial MEU_GIRO_RESUMO pelo conteúdo já
 * validado da aba MEU_GIRO_RESUMO_REBUILD_TESTE.
 *
 * Esta é uma operação de corte controlado: não recalcula, não reconstrói e
 * não altera a aba de origem nem qualquer outra aba operacional. Antes da
 * substituição, cria uma cópia integral da aba oficial para auditoria e
 * eventual recuperação manual.
 *
 * @return {Object} Resumo da operação executada.
 */
function substituirMeuGiroResumoPorRebuildTeste() {
  var nomeOrigem = 'MEU_GIRO_RESUMO_REBUILD_TESTE';
  var nomeDestino = SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO';
  var ss = getSpreadsheet_();
  var origem = ss.getSheetByName(nomeOrigem);
  var destino = ss.getSheetByName(nomeDestino);

  // Todas as pré-condições são verificadas antes da primeira escrita.
  if (!origem) {
    throw new Error('Corte abortado: aba não encontrada: ' + nomeOrigem);
  }

  if (!destino) {
    throw new Error('Corte abortado: aba não encontrada: ' + nomeDestino);
  }

  var totalLinhasOrigem = origem.getLastRow();
  if (totalLinhasOrigem <= 300) {
    throw new Error(
      'Corte abortado: ' + nomeOrigem +
      ' possui somente ' + totalLinhasOrigem +
      ' linhas; eram esperadas mais de 300.'
    );
  }

  var totalColunasOrigem = origem.getLastColumn();
  if (totalColunasOrigem < 1) {
    throw new Error('Corte abortado: a aba de origem não possui colunas para copiar.');
  }

  var executadoEm = new Date();
  var timezone = Session.getScriptTimeZone();
  var sufixoBackup = Utilities.formatDate(executadoEm, timezone, 'yyyyMMdd_HHmmss');
  var nomeBackup = 'MEU_GIRO_RESUMO_BACKUP_' + sufixoBackup;

  if (ss.getSheetByName(nomeBackup)) {
    throw new Error('Corte abortado: já existe uma aba com o nome ' + nomeBackup + '.');
  }

  var valoresOrigem = origem
    .getRange(1, 1, totalLinhasOrigem, totalColunasOrigem)
    .getValues();

  // Primeira escrita: o backup integral da aba oficial é obrigatório.
  var backup = destino.copyTo(ss);
  backup.setName(nomeBackup);

  // A expansão ocorre somente na aba oficial e somente após o backup.
  var linhasAdicionadas = Math.max(0, totalLinhasOrigem - destino.getMaxRows());
  var colunasAdicionadas = Math.max(0, totalColunasOrigem - destino.getMaxColumns());

  if (linhasAdicionadas > 0) {
    destino.insertRowsAfter(destino.getMaxRows(), linhasAdicionadas);
  }

  if (colunasAdicionadas > 0) {
    destino.insertColumnsAfter(destino.getMaxColumns(), colunasAdicionadas);
  }

  // Somente o conteúdo da aba oficial é limpo; formatação e estrutura ficam.
  destino.clearContents();
  destino
    .getRange(1, 1, totalLinhasOrigem, totalColunasOrigem)
    .setValues(valoresOrigem);

  SpreadsheetApp.flush();

  var totalLinhasDestino = destino.getLastRow();
  var totalColunasDestino = destino.getLastColumn();
  if (
    totalLinhasDestino !== totalLinhasOrigem ||
    totalColunasDestino !== totalColunasOrigem
  ) {
    throw new Error(
      'Corte abortado após a cópia: a origem possui ' + totalLinhasOrigem +
      ' linhas e ' + totalColunasOrigem + ' colunas, mas a aba oficial possui ' +
      totalLinhasDestino + ' linhas e ' + totalColunasDestino +
      ' colunas. Backup disponível em ' + nomeBackup + '.'
    );
  }

  var dataHora = Utilities.formatDate(executadoEm, timezone, 'yyyy-MM-dd HH:mm:ss');
  var resultado = {
    linhas_copiadas: totalLinhasOrigem,
    linhas_de_dados_copiadas: totalLinhasOrigem - 1,
    colunas_copiadas: totalColunasOrigem,
    linhas_adicionadas: linhasAdicionadas,
    colunas_adicionadas: colunasAdicionadas,
    backup_criado: nomeBackup,
    executado_em: dataHora,
    sucesso: true
  };

  Logger.log('[Meu Giro][corte controlado] linhas copiadas: ' + resultado.linhas_copiadas);
  Logger.log('[Meu Giro][corte controlado] colunas copiadas: ' + resultado.colunas_copiadas);
  Logger.log('[Meu Giro][corte controlado] linhas adicionadas: ' + resultado.linhas_adicionadas);
  Logger.log('[Meu Giro][corte controlado] colunas adicionadas: ' + resultado.colunas_adicionadas);
  Logger.log('[Meu Giro][corte controlado] backup criado: ' + resultado.backup_criado);
  Logger.log('[Meu Giro][corte controlado] data/hora: ' + resultado.executado_em);
  Logger.log('[Meu Giro][corte controlado] sucesso: true');

  return resultado;
}
