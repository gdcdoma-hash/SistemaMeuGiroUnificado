function onlyDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeCell_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
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
  var lista = ss.getSheetByName(SHEETS.LISTA_DESAFIOS || 'ListaDesafios');

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

function normalizarTimestampEdicao_(valor) {
  if (!valor) return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var s = String(valor).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
    return s.slice(6, 10) + '-' + s.slice(3, 5) + '-' + s.slice(0, 2);
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return '';
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
  var idxStatus = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição', 'status', 'situacao', 'situação']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  var idxPagamento = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagto_status', 'pagamento', 'pix_status']);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = normalizeText_(row[idxId]);

    if (rowId === id) {
      var statusInscricao = idxStatus > -1 ? normalizeText_(row[idxStatus]) : '';
      var statusConfirmacao = idxConfirmacao > -1 ? normalizeText_(row[idxConfirmacao]) : '';
      var statusPagamento = idxPagamento > -1 ? normalizeText_(row[idxPagamento]) : '';
      var validacao = validarInscricaoMinima_({
        status_inscricao: statusInscricao,
        status_confirmacao: statusConfirmacao,
        status_pagamento: statusPagamento
      });

      return {
        id_dgmb: rowId,
        aba_desafio: abaDesafio,
        status_inscricao: statusInscricao || 'inscrito',
        status_confirmacao: statusConfirmacao,
        status_pagamento: statusPagamento,
        inscricao_valida: validacao.valida,
        criterio_validacao: validacao.criterio,
        meta: idxMeta > -1 ? row[idxMeta] : '',
        distancia_realizada: idxRealizado > -1 ? row[idxRealizado] : '',
        frase_incentivo: idxFrase > -1 ? normalizeText_(row[idxFrase]) : ''
      };
    }
  }

  return null;
}

function montarErroInscricaoInvalida_(inscricao) {
  if (!inscricao) {
    return {
      code: 'NAO_INSCRITO',
      motivo: 'inscricao_nao_localizada',
      msg: 'Seu cadastro foi localizado, mas não há inscrição válida na aba do desafio atual.'
    };
  }

  return {
    code: 'INSCRICAO_INVALIDA',
    motivo: normalizeText_(inscricao.criterio_validacao) || 'criterio_minimo_inscricao',
    msg: 'Seu cadastro foi localizado, mas a inscrição está bloqueada pelos critérios mínimos do desafio.'
  };
}

function validarInscricaoMinima_(dadosInscricao) {
  var dados = dadosInscricao || {};
  var marcadores = [
    { campo: 'status_inscricao', valor: normalizeText_(dados.status_inscricao) },
    { campo: 'status_confirmacao', valor: normalizeText_(dados.status_confirmacao) },
    { campo: 'status_pagamento', valor: normalizeText_(dados.status_pagamento) }
  ];

  var possuiMarcador = false;

  for (var i = 0; i < marcadores.length; i++) {
    var marcador = marcadores[i];
    if (marcador.valor) {
      possuiMarcador = true;
    }

    if (inscricaoTemBloqueioMinimo_(marcador.valor)) {
      return {
        valida: false,
        criterio: 'bloqueio_em_' + marcador.campo
      };
    }
  }

  return {
    valida: true,
    criterio: possuiMarcador ? 'marcadores_sem_bloqueio' : 'presenca_id_dgmb'
  };
}

function inscricaoTemBloqueioMinimo_(valor) {
  var texto = normalizeText_(valor).toLowerCase();
  if (!texto) return false;

  var textoSemAcento = texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  var bloqueios = [
    'cancelad',
    'desist',
    'inativ',
    'exclu',
    'remov',
    'indefer',
    'recus',
    'nao confirmado',
    'nao conf',
    'nao pago',
    'estorn'
  ];

  for (var i = 0; i < bloqueios.length; i++) {
    if (textoSemAcento.indexOf(bloqueios[i]) !== -1) {
      return true;
    }
  }

  return false;
}

function parseLocalizedNumber_(value) {
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
  return parseLocalizedNumber_(value);
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

function extrairIdDesafioObservacao_(observacao) {
  var texto = String(observacao || '');
  var match = texto.match(/\[\s*ID_DESAFIO\s*:\s*([0-9]+)\s*\]/i);
  return match ? String(match[1]).trim() : '';
}

function normalizarDataISO_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    return s.slice(6, 10) + '-' + s.slice(3, 5) + '-' + s.slice(0, 2);
  }
  if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}(?::\d{2})?$/.test(s)) {
    return s.slice(6, 10) + '-' + s.slice(3, 5) + '-' + s.slice(0, 2);
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}

function isDataIsoValida_(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || '').trim());
}

function atividadeDentroPeriodoOficial_(dataAtividadeIso, periodoInicioIso, periodoFimIso) {
  var dataAtividade = String(dataAtividadeIso || '').trim();
  var inicio = String(periodoInicioIso || '').trim();
  var fim = String(periodoFimIso || '').trim();

  if (!isDataIsoValida_(dataAtividade) || !isDataIsoValida_(inicio) || !isDataIsoValida_(fim)) {
    return false;
  }

  return dataAtividade >= inicio && dataAtividade <= fim;
}

function buildPeriodoOficialPorAbaEId_(ss) {
  var out = { byAba: {}, byId: {} };
  var lista = ss.getSheetByName(SHEETS.LISTA_DESAFIOS || 'ListaDesafios');
  if (!lista) return out;

  var rows = lista.getDataRange().getValues();
  if (!rows || rows.length < 2) return out;

  var map = buildHeaderMap_(rows[0]);
  var idxAba = getOptionalColumnIndex_(map, ['aba', 'aba desafio', 'abadesafio']);
  var idxId = getOptionalColumnIndex_(map, [
    'id',
    'id_desafio',
    'id desafio',
    'id_desafio_lista',
    'id desafio lista',
    'id_desafio_base',
    'id desafio base'
  ]);
  var idxInicio = getOptionalColumnIndex_(map, ['data_inicio', 'data início', 'inicio', 'início', 'dt_inicio']);
  var idxFim = getOptionalColumnIndex_(map, ['data_fim', 'data fim', 'fim', 'dt_fim']);
  var idxNome = getOptionalColumnIndex_(map, [
    'nome_desafio',
    'nome desafio',
    'nome_desafio_lista',
    'nome desafio lista',
    'desafio',
    'nome'
  ]);

  if (idxAba === -1) idxAba = 1;

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var aba = normalizeText_(row[idxAba]);
    if (!aba) continue;

    var nomeDesafio = idxNome > -1 ? normalizeText_(row[idxNome]) : '';
    var periodo = {
      inicio: idxInicio > -1 ? normalizarDataISO_(row[idxInicio]) : '',
      fim: idxFim > -1 ? normalizarDataISO_(row[idxFim]) : '',
      nome_desafio: nomeDesafio || aba
    };

    out.byAba[aba] = periodo;

    if (idxId > -1) {
      var idDesafio = normalizeText_(row[idxId]);
      if (idDesafio) out.byId[idDesafio] = periodo;
    }
  }

  return out;
}

function buildMapaStatusDesafioListaPorId_(ss) {
  var out = { byId: {}, possuiColunaId: false };
  var lista = ss.getSheetByName(SHEETS.LISTA_DESAFIOS || 'ListaDesafios');
  if (!lista) return out;

  var rows = lista.getDataRange().getValues();
  if (!rows || rows.length < 2) return out;

  var map = buildHeaderMap_(rows[0]);
  var idxId = getOptionalColumnIndex_(map, [
    'id_desafio_lista',
    'id desafio lista',
    'id_desafio',
    'id desafio',
    'id'
  ]);
  var idxStatus = getOptionalColumnIndex_(map, [
    'status',
    'status_desafio',
    'status desafio',
    'status_lista',
    'situacao',
    'situação'
  ]);
  if (idxId === -1 || idxStatus === -1) return out;

  out.possuiColunaId = true;

  for (var i = 1; i < rows.length; i++) {
    var idDesafio = normalizeText_(rows[i][idxId]);
    if (!idDesafio) continue;
    out.byId[idDesafio] = normalizeText_(rows[i][idxStatus]).toLowerCase();
  }

  return out;
}

function obterVinculosDesafioUsuario_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var ss = getSpreadsheet_();
  var periodos = buildPeriodoOficialPorAbaEId_(ss);
  var statusListaDesafios = buildMapaStatusDesafioListaPorId_(ss);
  var abaDesafio = SHEETS.DESAFIO || 'dgmbDesafios';
  var sh = ss.getSheetByName(abaDesafio);
  if (!sh) return [];

  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  var map = buildHeaderMap_(values[0]);
  var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
  if (idxId === -1) return [];

  var idxMeta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km']);
  var idxObs = getOptionalColumnIndex_(map, ['observacao', 'observação']);
  var idxItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxTipoDesafio = getOptionalColumnIndex_(map, ['tipo_do_desafio', 'tipo do desafio', 'tipo_desafio', 'tipo desafio']);
  var idxStatusDesafio = getOptionalColumnIndex_(map, ['status_desafio', 'status desafio']);
  var idxStatusValidacaoCertificado = getOptionalColumnIndex_(map, ['status_validacao_certificado']);
  var idxStatusPag = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagamento', 'pix_status']);
  var idxStatusInscricao = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição', 'status', 'situacao', 'situação']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);

  var vinculos = [];
  var chaves = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = normalizeText_(row[idxId]);
    if (rowId !== id) continue;

    var observacao = idxObs > -1 ? row[idxObs] : '';
    var idDesafio = extrairIdDesafioObservacao_(observacao);
    var idItem = idxItem > -1 ? normalizeText_(row[idxItem]) : '';
    var tipoDesafio = idxTipoDesafio > -1 ? normalizeText_(row[idxTipoDesafio]) : '';
    var tipoSemAcento = tipoDesafio.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    var ehNormal = tipoSemAcento === 'normal';
    var metaKm = idxMeta > -1 ? parseLocalizedNumber_(row[idxMeta]) : 0;

    var statusInscricao = idxStatusInscricao > -1 ? normalizeText_(row[idxStatusInscricao]) : '';
    var statusConfirmacao = idxConfirmacao > -1 ? normalizeText_(row[idxConfirmacao]) : '';
    var statusPagamento = idxStatusPag > -1 ? normalizeText_(row[idxStatusPag]) : '';
    var statusDesafio = idxStatusDesafio > -1 ? normalizeText_(row[idxStatusDesafio]) : '';
    var statusValidacaoCertificado = idxStatusValidacaoCertificado > -1 ? normalizeText_(row[idxStatusValidacaoCertificado]) : '';

    var validacao = validarInscricaoMinima_({
      status_inscricao: statusInscricao,
      status_confirmacao: statusConfirmacao,
      status_pagamento: statusPagamento
    });
    var aptoBase = validacao.valida && !inscricaoTemBloqueioMinimo_(statusDesafio);
    var desafioAtivoNaLista = !statusListaDesafios.possuiColunaId
      ? !!idDesafio
      : statusListaDesafios.byId[idDesafio] === 'ativo';
    // Desafio normal não depende de item de estoque; repescagem mantém fluxo atual.
    var apto = ehNormal
      ? aptoBase && !!idDesafio && metaKm > 0 && desafioAtivoNaLista
      : aptoBase;

    var periodo = ehNormal
      ? (idDesafio && periodos.byId[idDesafio]) || { inicio: '', fim: '', nome_desafio: '' }
      : (idDesafio && periodos.byId[idDesafio]) || periodos.byAba[abaDesafio] || { inicio: '', fim: '', nome_desafio: '' };
    var chave = [id, idDesafio, idItem].join('|');
    if (chaves[chave]) continue;
    chaves[chave] = true;

    vinculos.push({
      id_dgmb: id,
      id_desafio: idDesafio,
      id_item_estoque: idItem,
      meta_km: metaKm,
      status_desafio: statusDesafio,
      status_validacao_certificado: statusValidacaoCertificado,
      apto: apto,
      periodo_inicio: periodo.inicio || '',
      periodo_fim: periodo.fim || '',
      nome_desafio: periodo.nome_desafio || abaDesafio || '',
      aba_desafio: abaDesafio
    });
  }

  return vinculos;
}

function obterRegistrosKmUsuario_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var registros = getAllObjects_(SHEETS.REGISTRO_KM);
  var out = [];

  for (var i = 0; i < registros.length; i++) {
    var r = registros[i];
    var rowId = normalizeText_(firstFilledValue_(r, ['ID_DGMB', 'id_dgmb']));
    if (rowId !== id) continue;

    out.push({
      data_atividade: normalizarDataISO_(firstFilledValue_(r, ['Data_Atividade', 'Data', 'data_atividade', 'data'])),
      km: parseLocalizedNumber_(firstFilledValue_(r, ['KM', 'km']))
    });
  }

  return out;
}

function ensureMeuGiroResumoSheet_() {
  var ss = getSpreadsheet_();
  var sheetName = SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO';
  var sh = ss.getSheetByName(sheetName);
  var headers = [
    'Timestamp_Atualizacao',
    'ID_DGMB',
    'ID_DESAFIO',
    'id_item_estoque',
    'Meta_KM',
    'Distancia_Realizada',
    'Percentual_Concluido',
    'Status_Apuracao'
  ];

  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  var atual = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var ok = true;
  for (var i = 0; i < headers.length; i++) {
    if (normalizeText_(atual[i]) !== headers[i]) {
      ok = false;
      break;
    }
  }
  if (!ok) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sh;
}

function atualizarMeuGiroResumo_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var vinculos = obterVinculosDesafioUsuario_(id);
  var registros = obterRegistrosKmUsuario_(id);
  var shResumo = ensureMeuGiroResumoSheet_();
  var valoresResumo = shResumo.getDataRange().getValues();
  var mapResumo = buildHeaderMap_(valoresResumo[0] || []);
  var idxId = getOptionalColumnIndex_(mapResumo, ['id_dgmb']);
  var idxDesafio = getOptionalColumnIndex_(mapResumo, ['id_desafio']);
  var idxItem = getOptionalColumnIndex_(mapResumo, ['id_item_estoque', 'id item estoque']);
  var idxMetaResumo = getOptionalColumnIndex_(mapResumo, ['meta_km', 'meta km']);
  var idxDistanciaResumo = getOptionalColumnIndex_(mapResumo, ['distancia_realizada', 'distancia realizada']);
  var idxPercentualResumo = getOptionalColumnIndex_(mapResumo, ['percentual_concluido', 'percentual concluido', 'percentual concluído']);
  var idxStatusResumo = getOptionalColumnIndex_(mapResumo, ['status_apuracao', 'status apuracao', 'status apuração']);
  var linhasPorChave = {};

  for (var i = 1; i < valoresResumo.length; i++) {
    var row = valoresResumo[i];
    if (normalizeText_(row[idxId]) !== id) continue;
    var chaveExistente = [normalizeText_(row[idxId]), normalizeText_(row[idxDesafio]), normalizeText_(row[idxItem])].join('|');
    linhasPorChave[chaveExistente] = i + 1;
  }

  var hoje = normalizarDataISO_(new Date());
  var saida = [];

  for (var v = 0; v < vinculos.length; v++) {
    var vinculo = vinculos[v];
    var chave = [id, vinculo.id_desafio, vinculo.id_item_estoque].join('|');
    var inicio = vinculo.periodo_inicio;
    var fim = vinculo.periodo_fim;
    var apto = !!vinculo.apto && !!inicio && !!fim && !!vinculo.id_desafio;
    var distancia = 0;

    if (apto) {
      for (var r = 0; r < registros.length; r++) {
        var reg = registros[r];
        if (atividadeDentroPeriodoOficial_(reg.data_atividade, inicio, fim)) {
          distancia += Number(reg.km || 0);
        }
      }
    }

    var meta = Number(vinculo.meta_km || 0);
    var percentual = meta > 0 ? (distancia / meta) * 100 : 0;
    var status = 'INAPTO';

    var dentroDoPeriodo = hoje >= inicio && hoje <= fim;
    if (apto) {
      if (distancia >= meta && meta > 0) {
        status = 'CONCLUIDO';
      } else if (dentroDoPeriodo) {
        status = 'ATIVO';
      } else {
        status = 'EXPIRADO';
      }
    }

    var metaArredondada = Math.round((meta + Number.EPSILON) * 10) / 10;
    var distanciaArredondada = Math.round((distancia + Number.EPSILON) * 10) / 10;
    var percentualArredondado = Math.round((percentual + Number.EPSILON) * 10) / 10;

    var linha = [
      new Date(),
      id,
      vinculo.id_desafio,
      vinculo.id_item_estoque,
      metaArredondada,
      distanciaArredondada,
      percentualArredondado,
      status
    ];

    if (linhasPorChave[chave]) {
      var numeroLinha = linhasPorChave[chave];
      var rowAtual = valoresResumo[numeroLinha - 1] || [];
      var houveMudanca =
        (idxMetaResumo > -1 ? parseLocalizedNumber_(rowAtual[idxMetaResumo]) : 0) !== metaArredondada ||
        (idxDistanciaResumo > -1 ? parseLocalizedNumber_(rowAtual[idxDistanciaResumo]) : 0) !== distanciaArredondada ||
        (idxPercentualResumo > -1 ? parseLocalizedNumber_(rowAtual[idxPercentualResumo]) : 0) !== percentualArredondado ||
        normalizeText_(idxStatusResumo > -1 ? rowAtual[idxStatusResumo] : '') !== status;

      if (houveMudanca) {
        shResumo.getRange(numeroLinha, 1, 1, linha.length).setValues([linha]);
      }
    } else {
      shResumo.appendRow(linha);
    }

    saida.push({
      id_dgmb: id,
      id_desafio: vinculo.id_desafio,
      id_item_estoque: vinculo.id_item_estoque,
      nome_desafio: vinculo.nome_desafio || '',
      meta_km: linha[4],
      distancia_realizada: linha[5],
      percentual_concluido: linha[6],
      status_apuracao: status,
      status_validacao_certificado: normalizeText_(vinculo.status_validacao_certificado).toUpperCase(),
      periodo_inicio: inicio || '',
      periodo_fim: fim || ''
    });
  }

  return saida;
}

function atualizarMeuGiroResumoEmLote_() {
  var ss = getSpreadsheet_();
  var abaDesafio = SHEETS.DESAFIO || 'dgmbDesafios';
  var sh = ss.getSheetByName(abaDesafio);

  if (!sh) {
    return { total_ids: 0, atualizados: 0, ids: [] };
  }

  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { total_ids: 0, atualizados: 0, ids: [] };
  }

  var map = buildHeaderMap_(values[0]);
  var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
  if (idxId === -1) {
    return { total_ids: 0, atualizados: 0, ids: [] };
  }

  var idxStatusDesafio = getOptionalColumnIndex_(map, ['status_desafio', 'status desafio']);
  var idxStatusPag = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagamento', 'pix_status']);
  var idxStatusInscricao = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição', 'status', 'situacao', 'situação']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  var ids = [];
  var idsMap = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var id = normalizeText_(row[idxId]);
    if (!id || idsMap[id]) continue;

    var statusInscricao = idxStatusInscricao > -1 ? normalizeText_(row[idxStatusInscricao]) : '';
    var statusConfirmacao = idxConfirmacao > -1 ? normalizeText_(row[idxConfirmacao]) : '';
    var statusPagamento = idxStatusPag > -1 ? normalizeText_(row[idxStatusPag]) : '';
    var statusDesafio = idxStatusDesafio > -1 ? normalizeText_(row[idxStatusDesafio]) : '';
    var validacao = validarInscricaoMinima_({
      status_inscricao: statusInscricao,
      status_confirmacao: statusConfirmacao,
      status_pagamento: statusPagamento
    });
    var apto = validacao.valida && !inscricaoTemBloqueioMinimo_(statusDesafio);
    if (!apto) continue;

    idsMap[id] = true;
    ids.push(id);
  }

  for (var j = 0; j < ids.length; j++) {
    atualizarMeuGiroResumo_(ids[j]);
  }

  return {
    total_ids: ids.length,
    atualizados: ids.length,
    ids: ids
  };
}

function atualizarMeuGiroResumoEmLote() {
  return atualizarMeuGiroResumoEmLote_();
}
