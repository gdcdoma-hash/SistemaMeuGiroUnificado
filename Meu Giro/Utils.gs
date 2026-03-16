function onlyDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeCell_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function toNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  var n = Number(String(value).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function getSpreadsheet_() {
  if (!SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID não informado no Config.gs');
  }

  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Não foi possível abrir a planilha pelo SPREADSHEET_ID');
  }
}

function getSheetByName_(name) {
  var sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) {
    throw new Error('Aba não encontrada: ' + name);
  }
  return sh;
}

function getSheetByNameOrThrow_(name) {
  return getSheetByName_(name);
}

function localizarAbaDesafioUsuario_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) {
    return {
      abaDesafio: SHEETS.DESAFIO,
      fallback: true
    };
  }

  var ss = getSpreadsheet_();
  var lista = ss.getSheetByName('ListaDesafios');

  if (!lista) {
    return {
      abaDesafio: SHEETS.DESAFIO,
      fallback: true
    };
  }

  var desafios = lista.getDataRange().getDisplayValues();
  if (!desafios || desafios.length < 2) {
    return {
      abaDesafio: SHEETS.DESAFIO,
      fallback: true
    };
  }

  var mapLista = buildHeaderMap_(desafios[0]);
  var idxAba = getOptionalColumnIndex_(mapLista, ['aba', 'aba desafio', 'abadesafio']);
  var idxStatus = getOptionalColumnIndex_(mapLista, ['status', 'situacao', 'situação']);

  if (idxAba === -1) idxAba = 1;
  if (idxStatus === -1) idxStatus = 3;

  for (var i = 1; i < desafios.length; i++) {
    var row = desafios[i];
    var abaOriginal = String(row[idxAba] || '').trim();
    var status = normalizeText_(row[idxStatus]).toLowerCase();

    if (!abaOriginal || status !== 'ativo') {
      continue;
    }

    var sheetDesafio = ss.getSheetByName(abaOriginal);
    if (!sheetDesafio) {
      continue;
    }

    var values = sheetDesafio.getDataRange().getValues();
    if (!values || values.length < 2) {
      continue;
    }

    var map = buildHeaderMap_(values[0]);
    var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
    if (idxId === -1) {
      continue;
    }

    for (var j = 1; j < values.length; j++) {
      var rowId = normalizeText_(values[j][idxId]);
      if (rowId === id) {
        return {
          abaDesafio: abaOriginal,
          fallback: false
        };
      }
    }
  }

  return {
    abaDesafio: SHEETS.DESAFIO,
    fallback: true
  };
}

function buildHeaderMap_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = normalizeHeaderKey_(headerRow[i]);
    if (key) {
      map[key] = i;
    }
  }
  return map;
}

function normalizeHeaderKey_(value) {
  return normalizeCell_(value)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function getRequiredColumnIndex_(map, candidates, sheetName) {
  var idx = getOptionalColumnIndex_(map, candidates);
  if (idx === -1) {
    throw new Error(
      'Coluna obrigatória não encontrada na aba ' +
      sheetName +
      ': ' +
      candidates.join(' / ')
    );
  }
  return idx;
}

function getOptionalColumnIndex_(map, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_(candidates[i]);
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      return map[key];
    }
  }
  return -1;
}

function formatDateToYMD_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

/**
 * Lê uma aba inteira e devolve um array de objetos
 * usando a primeira linha como cabeçalho.
 */
function getAllObjects_(sheetName) {
  var sh = getSheetByName_(sheetName);
  var values = sh.getDataRange().getValues();

  if (!values || values.length < 2) {
    return [];
  }

  var headers = values[0].map(function(h) {
    return normalizeCell_(h);
  });

  var items = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};

    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }

    items.push(obj);
  }

  return items;
}

function normalizeText_(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/\s+/g, ' ')
    .trim();
}

function obterDadosInscricaoUsuario_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return null;

  var localizacao = localizarAbaDesafioUsuario_(id);
  var abaDesafio = localizacao.abaDesafio;
  var sh = getSheetByName_(abaDesafio);
  var values = sh.getDataRange().getValues();

  if (!values || values.length < 2) {
    return null;
  }

  var header = values[0];
  var map = buildHeaderMap_(header);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], abaDesafio);
  var idxMeta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km']);
  var idxRealizado = getOptionalColumnIndex_(map, ['distancia_realizada', 'distancia realizada']);
  var idxFrase = getOptionalColumnIndex_(map, ['frase_incentivo']);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = normalizeText_(row[idxId]);

    if (rowId === id) {
      return {
        id_dgmb: rowId,
        aba_desafio: abaDesafio,
        status_inscricao: 'inscrito',
        meta: idxMeta > -1 ? row[idxMeta] : '',
        distancia_realizada: idxRealizado > -1 ? row[idxRealizado] : '',
        frase_incentivo: idxFrase > -1 ? normalizeText_(row[idxFrase]) : ''
      };
    }
  }

  return null;
}

function parseNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;

  var text = String(value).trim();

  text = text.replace(/\s/g, '');

  if (text.indexOf(',') > -1 && text.indexOf('.') > -1) {
    text = text.replace(/\./g, '').replace(',', '.');
  } else if (text.indexOf(',') > -1) {
    text = text.replace(',', '.');
  }

  var n = Number(text);
  return isNaN(n) ? 0 : n;
}
function toNumber_(value) {
  return parseNumber_(value);
}

function firstFilledValue_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, keys[i])) {
      var value = obj[keys[i]];
      if (value !== '' && value !== null && value !== undefined) {
        return value;
      }
    }
  }
  return '';
}
