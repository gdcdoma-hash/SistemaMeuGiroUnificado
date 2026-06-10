function registrarAtividade(idDgmb, dataAtividade, km, force) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    idDgmb = String(idDgmb || '').trim();
    dataAtividade = normalizarDataISO_(dataAtividade);
    km = parseKmInputSeguro_(km);

    if (!idDgmb) {
      return { ok:false, code:'ID_OBRIGATORIO', msg:'ID do atleta é obrigatório.' };
    }

    if (!dataAtividade) {
      return { ok:false, code:'DATA_OBRIGATORIA', msg:'Informe o dia da atividade.' };
    }

    if (!km || km <= 0) {
      return { ok:false, code:'KM_INVALIDO', msg:'Informe um valor de KM maior que zero.' };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(SHEETS.REGISTRO_KM);

    var dados = sheet.getDataRange().getValues();
    var cols = getRegistroKmColumnIndexes_(dados);

    cols = ensureRegistroKmActivityIdColumn_(sheet, dados, cols);
    var activityId = gerarActivityId_();

    for (var i = 1; i < dados.length; i++) {
      var rowId = String(dados[i][cols.idxId] || '').trim();
      var rowData = normalizarDataISO_(dados[i][cols.idxData]);
      var rowKm = normalizarKmEdicao_(dados[i][cols.idxKm]);

      if (rowId === idDgmb && rowData === dataAtividade && kmsIguaisEdicao_(rowKm, km)) {
        if (!force) {
          return {
            ok:false,
            code:'DUPLICIDADE',
            msg:'Já existe atividade com mesmo ID, data e KM informado.'
          };
        }
      }
    }

    var vinculos = obterVinculosDesafioUsuario_(idDgmb).filter(function(vinculo) {
      return !!vinculo.apto &&
        !!normalizeText_(vinculo.id_desafio) &&
        atividadeDentroPeriodoOficial_(
          dataAtividade,
          normalizarDataISO_(vinculo.periodo_inicio),
          normalizarDataISO_(vinculo.periodo_fim)
        );
    });

    if (!vinculos.length) {
      return {
        ok:false,
        code:'SEM_INSCRICAO_VALIDA',
        msg:'Nenhuma inscrição válida foi encontrada para a data da atividade.'
      };
    }

    var rowLength = Math.max(sheet.getLastColumn(), maiorIndiceRegistroKm_(cols) + 1);
    var timestamp = new Date();
    var rows = vinculos.map(function(vinculo) {
      var row = preencherLinhaRegistroKm_(rowLength, cols, {
        timestamp: timestamp,
        id_dgmb: idDgmb,
        id_inscricao: vinculo.id_inscricao,
        id_desafio: vinculo.id_desafio,
        id_item_estoque: vinculo.id_item_estoque,
        periodo_desafio: vinculo.periodo_desafio,
        data_atividade: dataAtividade,
        km: km,
        origem_registro: 'MANUAL',
        observacao: 'Lançamento manual Meu Giro',
        status_validacao: 'PENDENTE',
        activity_id: activityId
      });
      return row;
    });

    var primeiraLinhaInserida = sheet.getLastRow() + 1;
    sheet.getRange(primeiraLinhaInserida, 1, rows.length, rowLength).setValues(rows);

    try {
      atualizarDistanciaRealizada_(idDgmb);
      atualizarMeuGiroResumo_(idDgmb);
    } catch (syncErr) {
      sheet.deleteRows(primeiraLinhaInserida, rows.length);
      throw syncErr;
    }

    return {
      ok:true,
      activity_id: activityId,
      registros_criados: rows.length,
      msg:'Atividade registrada com sucesso.'
    };

  } catch(err) {
    Logger.log('registrarAtividade erro: ' + (err && err.stack ? err.stack : err));

    return {
      ok:false,
      code:'REGISTRAR_ATIVIDADE_EXCEPTION',
      msg:'Erro interno ao registrar atividade na aba REGISTRO_KM.'
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function gerarActivityId_() {
  return Utilities.getUuid();
}

function ensureRegistroKmActivityIdColumn_(sheet, dados, cols) {
  if (cols && cols.idxActivityId > -1) {
    return cols;
  }

  var headerLength = (dados && dados[0] && dados[0].length) ? dados[0].length : 0;
  var newIndex = headerLength;

  sheet.getRange(1, newIndex + 1).setValue('activity_id');

  cols.idxActivityId = newIndex;
  return cols;
}

function maiorIndiceRegistroKm_(cols) {
  var maior = -1;
  Object.keys(cols || {}).forEach(function(chave) {
    if (typeof cols[chave] === 'number' && cols[chave] > maior) {
      maior = cols[chave];
    }
  });
  return maior;
}

function preencherLinhaRegistroKm_(rowLength, cols, registro) {
  var row = [];
  for (var i = 0; i < rowLength; i++) row[i] = '';

  row[cols.idxTimestamp] = registro.timestamp;
  row[cols.idxId] = registro.id_dgmb;
  if (cols.idxInscricao > -1) row[cols.idxInscricao] = registro.id_inscricao || '';
  if (cols.idxDesafio > -1) row[cols.idxDesafio] = registro.id_desafio || '';
  if (cols.idxItemEstoque > -1) row[cols.idxItemEstoque] = registro.id_item_estoque || '';
  if (cols.idxPeriodoDesafio > -1) row[cols.idxPeriodoDesafio] = registro.periodo_desafio || '';
  row[cols.idxData] = registro.data_atividade;
  row[cols.idxKm] = registro.km;
  if (cols.idxOrigemRegistro > -1) row[cols.idxOrigemRegistro] = registro.origem_registro || '';
  if (cols.idxObservacao > -1) row[cols.idxObservacao] = registro.observacao || '';
  if (cols.idxStatusValidacao > -1) row[cols.idxStatusValidacao] = registro.status_validacao || '';
  row[cols.idxActivityId] = registro.activity_id;

  return row;
}

function atualizarDistanciaRealizada_(idDgmb){

  var registros = getAllObjects_(SHEETS.REGISTRO_KM);
  var total = 0;
  var activityIdsSomados = {};

  registros.forEach(function(r){
    if(String(r.ID_DGMB).trim() !== String(idDgmb).trim()) return;

    var activityId = obterActivityIdRegistroKm_(r);

    if (activityId) {
      if (activityIdsSomados[activityId]) return;
      activityIdsSomados[activityId] = true;
    }

    total += Number(r.KM || 0);
  });

  var inscricao = obterDadosInscricaoUsuario_(idDgmb);
  if (!inscricao || !inscricao.aba_desafio) return;

  var abaDesafio = inscricao.aba_desafio;
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(abaDesafio);

  var dados = sheet.getDataRange().getValues();
  if (!dados || dados.length < 2) return;

  var map = buildHeaderMap_(dados[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], abaDesafio);
  var idxRealizado = getRequiredColumnIndex_(map, ['distancia_realizada', 'distancia realizada'], abaDesafio);

  for(var i=1;i<dados.length;i++){
    if(String(dados[i][idxId]).trim() === String(idDgmb).trim()){
      sheet.getRange(i + 1, idxRealizado + 1).setValue(total);
      break;
    }
  }
}

function editarAtividade(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var activityId = String(payload.activity_id || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();
    var novaDataAtividade = normalizarDataISO_(payload.data_atividade);
    var novoKm = parseKmInputSeguro_(payload.km);

    if (!idDgmb) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do atleta é obrigatório.' };
    }
    if (!activityId && !chaveEdicao) {
      return { ok: false, code: 'IDENTIFICADOR_ATIVIDADE_OBRIGATORIO', msg: 'activity_id ou chave_edicao é obrigatório para edição.' };
    }
    if (!novaDataAtividade) {
      return { ok: false, code: 'DATA_OBRIGATORIA', msg: 'Informe o dia da atividade.' };
    }
    if (!novoKm || novoKm <= 0) {
      return { ok: false, code: 'KM_INVALIDO', msg: 'Informe um valor de KM maior que zero.' };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.REGISTRO_KM);
    var dados = sheet.getDataRange().getValues();
    var cols = getRegistroKmColumnIndexes_(dados);
    var linhasEncontradas = localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);

    if (!linhasEncontradas.length) {
      return { ok: false, code: 'ATIVIDADE_NAO_ENCONTRADA', msg: 'Atividade não encontrada para edição com a chave e ID informados.' };
    }

    var linhasDoLancamento = {};
    linhasEncontradas.forEach(function(linha) { linhasDoLancamento[linha] = true; });

    for (var j = 1; j < dados.length; j++) {
      var checkId = String(dados[j][cols.idxId] || '').trim();
      var checkData = normalizarDataEdicao_(dados[j][cols.idxData]);
      var checkKm = normalizarKmEdicao_(dados[j][cols.idxKm]);
      var linhaAtual = j + 1;

      if (!linhasDoLancamento[linhaAtual] && checkId === idDgmb &&
          checkData === novaDataAtividade && kmsIguaisEdicao_(checkKm, novoKm)) {
        return { ok: false, code: 'DUPLICIDADE_EDICAO', msg: 'Já existe uma atividade com esta mesma data e KM.' };
      }
    }

    var valoresOriginais = linhasEncontradas.map(function(linha) {
      return [dados[linha - 1][cols.idxData], dados[linha - 1][cols.idxKm]];
    });

    linhasEncontradas.forEach(function(linha) {
      sheet.getRange(linha, cols.idxData + 1).setValue(novaDataAtividade);
      sheet.getRange(linha, cols.idxKm + 1).setValue(novoKm);
    });

    try {
      atualizarDistanciaRealizada_(idDgmb);
      atualizarMeuGiroResumo_(idDgmb);
    } catch (syncErr) {
      linhasEncontradas.forEach(function(linha, index) {
        sheet.getRange(linha, cols.idxData + 1).setValue(valoresOriginais[index][0]);
        sheet.getRange(linha, cols.idxKm + 1).setValue(valoresOriginais[index][1]);
      });
      throw syncErr;
    }

    return { ok: true, registros_atualizados: linhasEncontradas.length, msg: 'Atividade atualizada com sucesso.' };
  } catch (err) {
    Logger.log('editarAtividade erro: ' + (err && err.stack ? err.stack : err));
    return { ok: false, code: 'EDITAR_ATIVIDADE_EXCEPTION', msg: 'Erro interno ao editar atividade na aba REGISTRO_KM.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function parseKmInputSeguro_(value) {
  var text = String(value === null || value === undefined ? '' : value).trim();
  if (!text) return NaN;

  text = text.replace(/\s+/g, '');
  if (!/^\d+(?:[.,]\d+)?$/.test(text)) return NaN;

  var parsed = parseLocalizedNumber_(text);
  return isFinite(parsed) ? parsed : NaN;
}

function normalizarDataEdicao_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var s = String(valor).trim();
  if (!s) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    return s.slice(6, 10) + '-' + s.slice(3, 5) + '-' + s.slice(0, 2);
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return s;
}

function normalizarKmEdicao_(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;

  var s = String(valor).trim().replace(/\s/g, '').replace(',', '.');
  var n = Number(s);

  if (isNaN(n)) return 0;

  return Math.round(n * 1000) / 1000;
}

function kmsIguaisEdicao_(a, b) {
  return Math.abs(Number(a || 0) - Number(b || 0)) < 0.0001;
}

function excluirAtividade(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var activityId = String(payload.activity_id || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();

    if (!idDgmb) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do atleta é obrigatório.' };
    }
    if (!activityId && !chaveEdicao) {
      return { ok: false, code: 'IDENTIFICADOR_ATIVIDADE_OBRIGATORIO', msg: 'activity_id ou chave_edicao é obrigatório para exclusão.' };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.REGISTRO_KM);
    var dados = sheet.getDataRange().getValues();
    var cols = getRegistroKmColumnIndexes_(dados);
    var linhasEncontradas = localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);

    if (!linhasEncontradas.length) {
      return { ok: false, code: 'ATIVIDADE_NAO_ENCONTRADA', msg: 'Atividade não encontrada para exclusão com a chave e ID informados.' };
    }

    var linhasOriginais = linhasEncontradas.map(function(linha) {
      return { numero: linha, valores: dados[linha - 1].slice() };
    });

    linhasEncontradas.slice().sort(function(a, b) { return b - a; }).forEach(function(linha) {
      sheet.deleteRow(linha);
    });

    try {
      atualizarDistanciaRealizada_(idDgmb);
      atualizarMeuGiroResumo_(idDgmb);
    } catch (syncErr) {
      linhasOriginais.forEach(function(item) {
        sheet.insertRowBefore(item.numero);
        sheet.getRange(item.numero, 1, 1, item.valores.length).setValues([item.valores]);
      });
      throw syncErr;
    }

    return { ok: true, registros_excluidos: linhasEncontradas.length, msg: 'Atividade excluída com sucesso.' };
  } catch (err) {
    Logger.log('excluirAtividade erro: ' + (err && err.stack ? err.stack : err));
    return { ok: false, code: 'EXCLUSAO_ATIVIDADE_ERROR', msg: 'Erro interno ao excluir atividade na aba REGISTRO_KM.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getRegistroKmColumnIndexes_(dados) {
  var fallback = {
    idxTimestamp: 0,
    idxId: 1,
    idxData: 2,
    idxKm: 3,
    idxInscricao: -1,
    idxDesafio: -1,
    idxItemEstoque: -1,
    idxPeriodoDesafio: -1,
    idxOrigemRegistro: -1,
    idxObservacao: -1,
    idxStatusValidacao: -1,
    idxActivityId: -1
  };

  if (!dados || !dados.length || !dados[0] || !dados[0].length) return fallback;

  var map = buildHeaderMap_(dados[0]);
  function indice(aliases, fallbackIndex) {
    var idx = getOptionalColumnIndex_(map, aliases);
    return idx > -1 ? idx : fallbackIndex;
  }

  return {
    idxTimestamp: indice(['timestamp', 'data_hora', 'data hora', 'criado_em', 'criado em'], fallback.idxTimestamp),
    idxId: indice(['id_dgmb'], fallback.idxId),
    idxInscricao: indice(['id_inscricao', 'id inscrição', 'id inscricao'], -1),
    idxDesafio: indice(['id_desafio', 'id desafio'], -1),
    idxItemEstoque: indice(['id_item_estoque', 'id item estoque'], -1),
    idxPeriodoDesafio: indice(['periodo_desafio', 'periodo desafio', 'período_desafio', 'período desafio'], -1),
    idxData: indice(['data_atividade', 'data atividade', 'data'], fallback.idxData),
    idxKm: indice(['km', 'distancia_km', 'distancia km'], fallback.idxKm),
    idxOrigemRegistro: indice(['origem_registro', 'origem registro'], -1),
    idxObservacao: indice(['observacao', 'observação'], -1),
    idxStatusValidacao: indice(['status_validacao', 'status validação', 'status validacao'], -1),
    idxActivityId: indice(['activity_id', 'activity id', 'id_atividade', 'id atividade'], -1)
  };
}

function localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao) {
  var idNormalizado = String(idDgmb || '').trim();
  var activityIdNormalizado = String(activityId || '').trim();
  var chaveNormalizada = String(chaveEdicao || '').trim();
  var linhas = [];

  if (!dados || dados.length < 2) return linhas;

  if (activityIdNormalizado && cols.idxActivityId > -1) {
    for (var i = 1; i < dados.length; i++) {
      var rowIdByActivity = String(dados[i][cols.idxId] || '').trim();
      var rowActivityId = String(dados[i][cols.idxActivityId] || '').trim();
      if (rowIdByActivity === idNormalizado && rowActivityId === activityIdNormalizado) {
        linhas.push(i + 1);
      }
    }
    if (linhas.length) return linhas;
  }

  if (!chaveNormalizada) return linhas;

  for (var j = 1; j < dados.length; j++) {
    var rowTimestamp = normalizarTimestampEdicao_(dados[j][cols.idxTimestamp]);
    var rowId = String(dados[j][cols.idxId] || '').trim();
    if (rowTimestamp === chaveNormalizada && rowId === idNormalizado) {
      linhas.push(j + 1);
    }
  }

  return linhas;
}

function localizarLinhaAtividade_(dados, cols, idDgmb, activityId, chaveEdicao) {
  var linhas = localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);
  return linhas.length ? linhas[0] : -1;
}
