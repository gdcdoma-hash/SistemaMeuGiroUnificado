function registrarAtividade(idDgmb, dataAtividade, km, force) {
  try {

    idDgmb = String(idDgmb || '').trim();
    dataAtividade = String(dataAtividade || '').trim();
    km = Number(String(km || '').replace(',', '.'));

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

    for (var i = 1; i < dados.length; i++) {

      var rowId = String(dados[i][1] || '').trim();
      var rowData = String(dados[i][2] || '').trim();
      var rowKm = Number(dados[i][3] || 0);

      if (rowId === idDgmb && rowData === dataAtividade && rowKm === km) {

        if (!force) {
          return {
            ok:false,
            code:'DUPLICIDADE',
            msg:'Atividade já registrada.'
          };
        }

      }
    }

    sheet.appendRow([
      new Date(),
      idDgmb,
      dataAtividade,
      km
    ]);

    atualizarDistanciaRealizada_(idDgmb);

    return {
      ok:true,
      msg:'Atividade registrada com sucesso.'
    };

  } catch(err) {

    return {
      ok:false,
      msg:err.message
    };

  }
}

function atualizarDistanciaRealizada_(idDgmb){

  var registros = getAllObjects_(SHEETS.REGISTRO_KM);
  var total = 0;

  registros.forEach(function(r){

    if(String(r.ID_DGMB).trim() === String(idDgmb).trim()){

      total += Number(r.KM || 0);

    }

  });

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(SHEETS.DESAFIO);

  var dados = sheet.getDataRange().getValues();

  for(var i=1;i<dados.length;i++){

    if(String(dados[i][1]).trim() === String(idDgmb).trim()){

      sheet.getRange(i+1,11).setValue(total);
      break;

    }

  }

}

function editarAtividade(payload) {
  try {
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();
    var novaDataAtividade = String(payload.data_atividade || '').trim();
    var novoKm = Number(String(payload.km || '').replace(',', '.'));

    if (!idDgmb) {
      return {
        ok: false,
        code: 'ID_OBRIGATORIO',
        msg: 'ID do atleta é obrigatório.'
      };
    }

    if (!chaveEdicao) {
      return {
        ok: false,
        code: 'CHAVE_EDICAO_OBRIGATORIA',
        msg: 'Chave da atividade é obrigatória para edição.'
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
    var linhaEncontrada = -1;

    for (var i = 1; i < dados.length; i++) {
      var rowTimestamp = normalizarTimestampEdicao_(dados[i][0]);
      var rowId = String(dados[i][1] || '').trim();

      if (rowTimestamp === chaveEdicao && rowId === idDgmb) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return {
        ok: false,
        code: 'ATIVIDADE_NAO_ENCONTRADA',
        msg: 'Atividade não encontrada para edição.'
      };
    }

    for (var j = 1; j < dados.length; j++) {
      var checkId = String(dados[j][1] || '').trim();
      var checkData = normalizarDataEdicao_(dados[j][2]);
      var checkKm = normalizarKmEdicao_(dados[j][3]);
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

    sheet.getRange(linhaEncontrada, 3).setValue(novaDataAtividade);
    sheet.getRange(linhaEncontrada, 4).setValue(novoKm);

    atualizarDistanciaRealizada_(idDgmb);

    return {
      ok: true,
      msg: 'Atividade atualizada com sucesso.'
    };

  } catch (err) {
    return {
      ok: false,
      msg: err.message
    };
  }
}

function normalizarTimestampEdicao_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  }

  return String(valor).trim();
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

function datasIguaisEdicao_(a, b) {
  return String(a || '').trim() === String(b || '').trim();
}

function kmsIguaisEdicao_(a, b) {
  return Math.abs(Number(a || 0) - Number(b || 0)) < 0.0001;
}

function normalizarTimestampEdicao_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  }

  return String(valor).trim();
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
  try {
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();

    if (!idDgmb) {
      return {
        ok: false,
        code: 'ID_OBRIGATORIO',
        msg: 'ID do atleta é obrigatório.'
      };
    }

    if (!chaveEdicao) {
      return {
        ok: false,
        code: 'CHAVE_EDICAO_OBRIGATORIA',
        msg: 'Chave da atividade é obrigatória para exclusão.'
      };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(SHEETS.REGISTRO_KM);

    var dados = sheet.getDataRange().getValues();
    var linhaEncontrada = -1;

    for (var i = 1; i < dados.length; i++) {
      var rowTimestamp = normalizarTimestampEdicao_(dados[i][0]);
      var rowId = String(dados[i][1] || '').trim();

      if (rowTimestamp === chaveEdicao && rowId === idDgmb) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return {
        ok: false,
        code: 'ATIVIDADE_NAO_ENCONTRADA',
        msg: 'Atividade não encontrada para exclusão.'
      };
    }

    sheet.deleteRow(linhaEncontrada);
    atualizarDistanciaRealizada_(idDgmb);

    return {
      ok: true,
      msg: 'Atividade excluída com sucesso.'
    };

  } catch (err) {
    return {
      ok: false,
      msg: err && err.message ? err.message : 'Erro ao excluir atividade.'
    };
  }
}