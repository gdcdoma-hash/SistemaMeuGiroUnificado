function registrarAtividade(idDgmb, dataAtividade, km, force) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    idDgmb = String(idDgmb || '').trim();
    dataAtividade = String(dataAtividade || '').trim();
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
      var rowData = String(dados[i][cols.idxData] || '').trim();
      var rowKm = Number(dados[i][cols.idxKm] || 0);

      if (rowId === idDgmb && rowData === dataAtividade && rowKm === km) {

        if (!force) {
          return {
            ok:false,
            code:'DUPLICIDADE',
            msg:'Já existe atividade com mesmo ID, data e KM informado.'
          };
        }

      }
    }

    var row = [];
    var rowLength = Math.max(cols.idxTimestamp, cols.idxId, cols.idxData, cols.idxKm, cols.idxActivityId) + 1;

    for (var idx = 0; idx < rowLength; idx++) {
      row[idx] = '';
    }

    row[cols.idxTimestamp] = new Date();
    row[cols.idxId] = idDgmb;
    row[cols.idxData] = dataAtividade;
    row[cols.idxKm] = km;
    row[cols.idxActivityId] = activityId;

    sheet.appendRow(row);

    try {
      atualizarDistanciaRealizada_(idDgmb);
      atualizarMeuGiroResumo_(idDgmb);
    } catch (syncErr) {
      var linhaInserida = localizarLinhaAtividade_(sheet.getDataRange().getValues(), cols, idDgmb, activityId, '');
      if (linhaInserida > -1) {
        sheet.deleteRow(linhaInserida);
      }
      throw syncErr;
    }

    return {
      ok:true,
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

  return {
    idxTimestamp: cols.idxTimestamp,
    idxId: cols.idxId,
    idxData: cols.idxData,
    idxKm: cols.idxKm,
    idxActivityId: newIndex
  };
}

function atualizarDistanciaRealizada_(idDgmb){

  var registros = getAllObjects_(SHEETS.REGISTRO_KM);
  var total = 0;

  registros.forEach(function(r){

    if(String(r.ID_DGMB).trim() === String(idDgmb).trim()){

      total += Number(r.KM || 0);

    }

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
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var activityId = String(payload.activity_id || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();
    var novaDataAtividade = String(payload.data_atividade || '').trim();
    var novoKm = parseKmInputSeguro_(payload.km);

    if (!idDgmb) {
      return {
        ok: false,
        code: 'ID_OBRIGATORIO',
        msg: 'ID do atleta é obrigatório.'
      };
    }

    if (!activityId && !chaveEdicao) {
      return {
        ok: false,
        code: 'IDENTIFICADOR_ATIVIDADE_OBRIGATORIO',
        msg: 'activity_id ou chave_edicao é obrigatório para edição.'
      };
    }

    if (!novaDataAtividade) {
      return {
        ok: false,
        code: 'DATA_OBRIGATORIA',
        msg: 'Informe o dia da atividade.'
      };
    }

    if (!novoKm || novoKm <= 0) {
      return {
        ok: false,
        code: 'KM_INVALIDO',
        msg: 'Informe um valor de KM maior que zero.'
      };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(SHEETS.REGISTRO_KM);

    var dados = sheet.getDataRange().getValues();
    var cols = getRegistroKmColumnIndexes_(dados);
    var linhaEncontrada = localizarLinhaAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);

    if (linhaEncontrada === -1) {
      return {
        ok: false,
        code: 'ATIVIDADE_NAO_ENCONTRADA',
        msg: 'Atividade não encontrada para edição com a chave e ID informados.'
      };
    }

    for (var j = 1; j < dados.length; j++) {
      var checkId = String(dados[j][cols.idxId] || '').trim();
      var checkData = normalizarDataEdicao_(dados[j][cols.idxData]);
      var checkKm = normalizarKmEdicao_(dados[j][cols.idxKm]);
      var linhaAtual = j + 1;

      if (
        linhaAtual !== linhaEncontrada &&
        checkId === idDgmb &&
        checkData === novaDataAtividade &&
        kmsIguaisEdicao_(checkKm, novoKm)
      ) {
        return {
          ok: false,
          code: 'DUPLICIDADE_EDICAO',
          msg: 'Já existe uma atividade com esta mesma data e KM.'
        };
      }
    }

    var valorDataOriginal = dados[linhaEncontrada - 1][cols.idxData];
    var valorKmOriginal = dados[linhaEncontrada - 1][cols.idxKm];

    sheet.getRange(linhaEncontrada, cols.idxData + 1).setValue(novaDataAtividade);
    sheet.getRange(linhaEncontrada, cols.idxKm + 1).setValue(novoKm);

    try {
      atualizarDistanciaRealizada_(idDgmb);
      atualizarMeuGiroResumo_(idDgmb);
    } catch (syncErr) {
      sheet.getRange(linhaEncontrada, cols.idxData + 1).setValue(valorDataOriginal);
      sheet.getRange(linhaEncontrada, cols.idxKm + 1).setValue(valorKmOriginal);
      throw syncErr;
    }

    return {
      ok: true,
      msg: 'Atividade atualizada com sucesso.'
    };

  } catch (err) {
    Logger.log('editarAtividade erro: ' + (err && err.stack ? err.stack : err));
    return {
      ok: false,
      code: 'EDITAR_ATIVIDADE_EXCEPTION',
      msg: 'Erro interno ao editar atividade na aba REGISTRO_KM.'
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
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
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var activityId = String(payload.activity_id || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();

    if (!idDgmb) {
      return {
        ok: false,
        code: 'ID_OBRIGATORIO',
        msg: 'ID do atleta é obrigatório.'
      };
    }

    if (!activityId && !chaveEdicao) {
      return {
        ok: false,
        code: 'IDENTIFICADOR_ATIVIDADE_OBRIGATORIO',
        msg: 'activity_id ou chave_edicao é obrigatório para exclusão.'
      };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(SHEETS.REGISTRO_KM);

    var dados = sheet.getDataRange().getValues();
    var cols = getRegistroKmColumnIndexes_(dados);
    var linhaEncontrada = localizarLinhaAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);

    if (linhaEncontrada === -1) {
      return {
        ok: false,
        code: 'ATIVIDADE_NAO_ENCONTRADA',
        msg: 'Atividade não encontrada para exclusão com a chave e ID informados.'
      };
    }

    var linhaOriginal = dados[linhaEncontrada - 1];
    sheet.deleteRow(linhaEncontrada);

    try {
      atualizarDistanciaRealizada_(idDgmb);
      atualizarMeuGiroResumo_(idDgmb);
    } catch (syncErr) {
      sheet.insertRowBefore(linhaEncontrada);
      sheet.getRange(linhaEncontrada, 1, 1, linhaOriginal.length).setValues([linhaOriginal]);
      throw syncErr;
    }

    return {
      ok: true,
      msg: 'Atividade excluída com sucesso.'
    };

  } catch (err) {
    Logger.log('excluirAtividade erro: ' + (err && err.stack ? err.stack : err));
    return {
      ok: false,
      code: 'EXCLUSAO_ATIVIDADE_ERROR',
      msg: 'Erro interno ao excluir atividade na aba REGISTRO_KM.'
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function getRegistroKmColumnIndexes_(dados) {
  var fallback = {
    idxTimestamp: 0,
    idxId: 1,
    idxData: 2,
    idxKm: 3
  };

  var aliases = {
    timestamp: ['timestamp', 'data_hora', 'data hora', 'criado_em', 'criado em'],
    id: ['id_dgmb'],
    data: ['data_atividade', 'data atividade', 'data'],
    km: ['km', 'distancia_km', 'distancia km'],
    activityId: ['activity_id', 'activity id', 'id_atividade', 'id atividade']
  };

  if (!dados || !dados.length || !dados[0] || !dados[0].length) {
    return fallback;
  }

  var map = buildHeaderMap_(dados[0]);
  var idxTimestamp = getOptionalColumnIndex_(map, aliases.timestamp);
  var idxId = getOptionalColumnIndex_(map, aliases.id);
  var idxData = getOptionalColumnIndex_(map, aliases.data);
  var idxKm = getOptionalColumnIndex_(map, aliases.km);
  var idxActivityId = getOptionalColumnIndex_(map, aliases.activityId);

  return {
    idxTimestamp: idxTimestamp > -1 ? idxTimestamp : fallback.idxTimestamp,
    idxId: idxId > -1 ? idxId : fallback.idxId,
    idxData: idxData > -1 ? idxData : fallback.idxData,
    idxKm: idxKm > -1 ? idxKm : fallback.idxKm,
    idxActivityId: idxActivityId
  };
}

function localizarLinhaAtividade_(dados, cols, idDgmb, activityId, chaveEdicao) {
  var idNormalizado = String(idDgmb || '').trim();
  var activityIdNormalizado = String(activityId || '').trim();
  var chaveNormalizada = String(chaveEdicao || '').trim();

  if (!dados || dados.length < 2) {
    return -1;
  }

  if (activityIdNormalizado && cols.idxActivityId > -1) {
    for (var i = 1; i < dados.length; i++) {
      var rowIdByActivity = String(dados[i][cols.idxId] || '').trim();
      var rowActivityId = String(dados[i][cols.idxActivityId] || '').trim();

      if (rowIdByActivity === idNormalizado && rowActivityId === activityIdNormalizado) {
        return i + 1;
      }
    }
  }

  if (!chaveNormalizada) {
    return -1;
  }

  for (var j = 1; j < dados.length; j++) {
    var rowTimestamp = normalizarTimestampEdicao_(dados[j][cols.idxTimestamp]);
    var rowId = String(dados[j][cols.idxId] || '').trim();

    if (rowTimestamp === chaveNormalizada && rowId === idNormalizado) {
      return j + 1;
    }
  }

  return -1;
}
