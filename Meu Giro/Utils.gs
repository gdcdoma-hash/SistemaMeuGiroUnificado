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

function converterValoresRegistroKmEmObjetos_(values) {
  if (!values || values.length < 2 || !values[0] || !values[0].length) return [];

  var headers = values[0].map(function(h) {
    return normalizeCell_(h);
  });

  return values.slice(1).map(function(row) {
    var obj = {};
    for (var i = 0; i < headers.length; i++) {
      obj[headers[i]] = row[i];
    }
    return obj;
  });
}

function obterRegistrosKmObjetosReaproveitados_(idDgmb, opcoes) {
  var id = normalizeText_(idDgmb);
  var contextoId = normalizeText_(opcoes && opcoes.idDgmb);
  var objetos = opcoes && opcoes.registrosKmObjetos;
  var valores = opcoes && opcoes.registrosKmValores;
  var layout = opcoes && opcoes.layoutRegistroKm;
  var contextoValido = !!id && contextoId === id &&
    Array.isArray(objetos) && Array.isArray(valores) &&
    valores.length === objetos.length + 1 &&
    valores.length > 0 && Array.isArray(valores[0]) &&
    layout && typeof layout.idxId === 'number' &&
    layout.idxId > -1 && layout.idxId < valores[0].length;

  if (contextoValido) {
    return { valores: objetos, reaproveitados: true };
  }

  return { valores: getAllObjects_(SHEETS.REGISTRO_KM), reaproveitados: false };
}

function normalizeText_(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/\s+/g, ' ')
    .trim();
}

function obterDadosInscricaoUsuario_(idDgmb, contextoDesafios) {
  var id = normalizeText_(idDgmb);
  if (!id) return null;

  // A inscrição do atleta é identificada diretamente na base consolidada.
  // ListaDesafios é catálogo operacional e não deve autorizar ou bloquear o login.
  var abaDesafio = contextoDesafios && contextoDesafios.abaDesafio
    ? contextoDesafios.abaDesafio
    : SHEETS.DESAFIO || 'dgmbDesafios';
  var values = contextoDesafios && Array.isArray(contextoDesafios.values)
    ? contextoDesafios.values
    : getSheetByName_(abaDesafio).getDataRange().getValues();

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
  var idxStatusUsuarioDesafio = getOptionalColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  var idxPagamento = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagto_status', 'pagamento', 'pix_status']);
  var primeiraInscricaoInvalida = null;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = normalizeText_(row[idxId]);

    if (rowId === id) {
      var statusInscricao = idxStatus > -1 ? normalizeText_(row[idxStatus]) : '';
      if (!statusInscricao && idxStatusUsuarioDesafio > -1) {
        statusInscricao = normalizeText_(row[idxStatusUsuarioDesafio]);
      }
      var statusConfirmacao = idxConfirmacao > -1 ? normalizeText_(row[idxConfirmacao]) : '';
      var statusPagamento = idxPagamento > -1 ? normalizeText_(row[idxPagamento]) : '';
      var validacao = validarInscricaoMinima_({
        status_inscricao: statusInscricao,
        status_confirmacao: statusConfirmacao,
        status_pagamento: statusPagamento
      });
      var inscricao = {
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

      if (validacao.valida) {
        return inscricao;
      }

      if (!primeiraInscricaoInvalida) {
        primeiraInscricaoInvalida = inscricao;
      }
    }
  }

  return primeiraInscricaoInvalida;
}

function montarErroInscricaoInvalida_(inscricao) {
  if (!inscricao) {
    return {
      code: 'NAO_INSCRITO',
      motivo: 'inscricao_nao_localizada',
      msg: 'Seu cadastro foi localizado, mas não há inscrição válida registrada para acesso ao Meu Giro.'
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

function getIdDesafioColumnIndex_(map) {
  return getOptionalColumnIndex_(map, ['id_desafio', 'id desafio']);
}

function obterIdDesafioRegistro_(row, idxIdDesafio, idxObservacao) {
  if (idxIdDesafio > -1) {
    return normalizeText_(row[idxIdDesafio]);
  }

  return idxObservacao > -1
    ? extrairIdDesafioObservacao_(row[idxObservacao])
    : '';
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

function normalizarPeriodoMensal_(value) {
  if (!value) return { inicio: '', fim: '' };

  var ano = 0;
  var mes = 0;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    ano = Number(Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy'));
    mes = Number(Utilities.formatDate(value, Session.getScriptTimeZone(), 'MM'));
  } else {
    var texto = normalizeText_(value).toLowerCase();
    var numerico = texto.match(/^(\d{2})\/(\d{4})$/);
    var iso = texto.match(/^(\d{4})-(\d{2})$/);

    if (numerico) {
      mes = Number(numerico[1]);
      ano = Number(numerico[2]);
    } else if (iso) {
      ano = Number(iso[1]);
      mes = Number(iso[2]);
    } else {
      var porExtenso = texto.match(/^([a-zçãáàâéêíóôõú]+)\s*\/\s*(\d{4})$/i);
      if (porExtenso) {
        var nomeMes = porExtenso[1]
          .replace(/[áàâã]/g, 'a')
          .replace(/[éê]/g, 'e')
          .replace(/í/g, 'i')
          .replace(/[óôõ]/g, 'o')
          .replace(/ú/g, 'u')
          .replace(/ç/g, 'c');
        var meses = {
          janeiro: 1,
          fevereiro: 2,
          marco: 3,
          abril: 4,
          maio: 5,
          junho: 6,
          julho: 7,
          agosto: 8,
          setembro: 9,
          outubro: 10,
          novembro: 11,
          dezembro: 12
        };
        mes = meses[nomeMes] || 0;
        ano = Number(porExtenso[2]);
      }
    }
  }

  if (!ano || mes < 1 || mes > 12) return { inicio: '', fim: '' };

  var bissexto = ano % 4 === 0 && (ano % 100 !== 0 || ano % 400 === 0);
  var diasPorMes = [31, bissexto ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  var mesTexto = String(mes).padStart(2, '0');

  return {
    inicio: String(ano) + '-' + mesTexto + '-01',
    fim: String(ano) + '-' + mesTexto + '-' + String(diasPorMes[mes - 1]).padStart(2, '0')
  };
}

function meuGiroPerfDebugAtivo_() {
  try {
    return typeof PERFORMANCE_DEBUG !== 'undefined' && !!PERFORMANCE_DEBUG;
  } catch (e) {
    return false;
  }
}

function meuGiroPerfNow_() {
  try {
    return Date.now();
  } catch (e) {
    return new Date().getTime();
  }
}

var MEU_GIRO_PERF_OPERACAO_ATUAL_ = '';

function meuGiroPerfLog_(escopo, etapa, inicio, extras) {
  if (!meuGiroPerfDebugAtivo_()) return;

  try {
    var payload = {
      etapa: etapa,
      duracao_ms: meuGiroPerfNow_() - inicio
    };
    if (MEU_GIRO_PERF_OPERACAO_ATUAL_) {
      payload.operacao = MEU_GIRO_PERF_OPERACAO_ATUAL_;
    }
    Object.keys(extras || {}).forEach(function(chave) {
      payload[chave] = extras[chave];
    });
    Logger.log('[Meu Giro][performance][' + escopo + '] ' + JSON.stringify(payload));
  } catch (e) {}
}

var MEU_GIRO_DIAGNOSTICO_LOGS_EXECUCAO_ = 0;
var LISTA_DESAFIOS_CACHE_EXECUCAO_ = null;

function buildListaDesafiosContexto_(ss) {
  var perfTotalInicio = meuGiroPerfNow_();
  if (LISTA_DESAFIOS_CACHE_EXECUCAO_ !== null) {
    meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'cache_hit_lista_desafios', perfTotalInicio);
    meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'buildListaDesafiosContexto_total', perfTotalInicio, {
      cache_lista_desafios: 'hit'
    });
    return LISTA_DESAFIOS_CACHE_EXECUCAO_;
  }

  meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'cache_miss_lista_desafios', perfTotalInicio);
  var contexto = {
    periodos: { byAba: {}, byId: {} },
    status: { byId: {}, possuiColunaId: false }
  };
  var lista = ss.getSheetByName(SHEETS.LISTA_DESAFIOS || 'ListaDesafios');
  if (!lista) {
    LISTA_DESAFIOS_CACHE_EXECUCAO_ = contexto;
    meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'buildListaDesafiosContexto_total', perfTotalInicio, {
      cache_lista_desafios: 'miss',
      quantidade_linhas_lista_desafios: 0
    });
    return contexto;
  }

  var perfLeituraInicio = meuGiroPerfNow_();
  if (typeof painelMG_incrementarAuditoriaCarregamentoInicial_ === 'function') {
    painelMG_incrementarAuditoriaCarregamentoInicial_('leituras_ListaDesafios');
  }
  var rows = lista.getDataRange().getValues();
  meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'leitura_ListaDesafios_contexto', perfLeituraInicio, {
    quantidade_linhas_lista_desafios: rows && rows.length ? rows.length - 1 : 0
  });
  if (!rows || rows.length < 2) {
    LISTA_DESAFIOS_CACHE_EXECUCAO_ = contexto;
    meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'buildListaDesafiosContexto_total', perfTotalInicio, {
      cache_lista_desafios: 'miss',
      quantidade_linhas_lista_desafios: rows && rows.length ? rows.length - 1 : 0
    });
    return contexto;
  }

  var map = buildHeaderMap_(rows[0]);
  var idxAba = getOptionalColumnIndex_(map, ['aba', 'aba desafio', 'abadesafio']);
  var idxIdPeriodo = getOptionalColumnIndex_(map, [
    'id',
    'id_desafio',
    'id desafio',
    'id_desafio_lista',
    'id desafio lista',
    'id_desafio_base',
    'id desafio base'
  ]);
  var idxPeriodo = getOptionalColumnIndex_(map, ['periodo', 'período']);
  var idxNome = getOptionalColumnIndex_(map, [
    'nome_desafio',
    'nome desafio',
    'nome_desafio_lista',
    'nome desafio lista',
    'desafio',
    'nome'
  ]);
  var idxIdStatus = getOptionalColumnIndex_(map, [
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

  if (idxAba === -1) idxAba = 1;
  if (idxIdStatus > -1 && idxStatus > -1) contexto.status.possuiColunaId = true;

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var aba = normalizeText_(row[idxAba]);

    if (aba) {
      var nomeDesafio = idxNome > -1 ? normalizeText_(row[idxNome]) : '';
      var periodoTexto = idxPeriodo > -1 ? normalizeText_(row[idxPeriodo]) : '';
      var periodoMensal = idxPeriodo > -1
        ? normalizarPeriodoMensal_(row[idxPeriodo])
        : { inicio: '', fim: '' };
      var periodo = {
        inicio: periodoMensal.inicio,
        fim: periodoMensal.fim,
        periodo_desafio: periodoTexto,
        nome_desafio: nomeDesafio || aba
      };

      contexto.periodos.byAba[aba] = periodo;

      if (idxIdPeriodo > -1) {
        var idDesafioPeriodo = normalizeText_(row[idxIdPeriodo]);
        if (idDesafioPeriodo) contexto.periodos.byId[idDesafioPeriodo] = periodo;
      }
    }

    if (contexto.status.possuiColunaId) {
      var idDesafioStatus = normalizeText_(row[idxIdStatus]);
      if (idDesafioStatus) contexto.status.byId[idDesafioStatus] = normalizeText_(row[idxStatus]).toLowerCase();
    }
  }

  LISTA_DESAFIOS_CACHE_EXECUCAO_ = contexto;
  meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'buildListaDesafiosContexto_total', perfTotalInicio, {
    cache_lista_desafios: 'miss',
    quantidade_linhas_lista_desafios: rows.length - 1
  });
  return contexto;
}

function buildPeriodoOficialPorAbaEId_(ss) {
  return buildListaDesafiosContexto_(ss).periodos;
}

function buildMapaStatusDesafioListaPorId_(ss) {
  return buildListaDesafiosContexto_(ss).status;
}

function logMeuGiroDiagnostico_(mensagem, dados) {
  MEU_GIRO_DIAGNOSTICO_LOGS_EXECUCAO_++;
  if (!meuGiroPerfDebugAtivo_()) return;

  try {
    Logger.log('[Meu Giro][diagnostico] ' + mensagem + (dados ? ' ' + JSON.stringify(dados) : ''));
  } catch (e) {}
}

function extrairPeriodoDesafioTexto_(texto) {
  var bruto = normalizeText_(texto);
  if (!bruto) return { inicio: '', fim: '' };

  var iso = bruto.match(/(\d{4}-\d{2}-\d{2}).*?(\d{4}-\d{2}-\d{2})/);
  if (iso) return { inicio: normalizarDataISO_(iso[1]), fim: normalizarDataISO_(iso[2]) };

  var br = bruto.match(/(\d{2}\/\d{2}\/\d{4}).*?(\d{2}\/\d{2}\/\d{4})/);
  if (br) return { inicio: normalizarDataISO_(br[1]), fim: normalizarDataISO_(br[2]) };

  return normalizarPeriodoMensal_(bruto);
}

function periodoCompletoValido_(periodo) {
  return !!periodo &&
    isDataIsoValida_(periodo.inicio) &&
    isDataIsoValida_(periodo.fim) &&
    periodo.inicio <= periodo.fim;
}

function montarPeriodoHistoricoVinculo_(row, indices, periodoLista, contextoLog) {
  var periodoTexto = indices.periodo > -1 ? normalizeText_(row[indices.periodo]) : '';
  var periodoDatasEspecificas = {
    inicio: indices.inicio > -1 ? normalizarDataISO_(row[indices.inicio]) : '',
    fim: indices.fim > -1 ? normalizarDataISO_(row[indices.fim]) : ''
  };
  var periodoTextoEspecifico = extrairPeriodoDesafioTexto_(periodoTexto);
  var periodo = { inicio: '', fim: '' };
  var usouFallbackLista = false;
  var usouFallbackDatasEspecificas = false;

  if (periodoCompletoValido_(periodoTextoEspecifico)) {
    periodo = periodoTextoEspecifico;
  } else if (periodoCompletoValido_(periodoLista)) {
    periodo = periodoLista;
    usouFallbackLista = true;
  } else if (periodoCompletoValido_(periodoDatasEspecificas)) {
    periodo = periodoDatasEspecificas;
    usouFallbackDatasEspecificas = true;
  }

  if (usouFallbackLista) {
    logMeuGiroDiagnostico_('Fallback de período via ListaDesafios.Periodo usado.', contextoLog);
  }

  if (usouFallbackDatasEspecificas) {
    logMeuGiroDiagnostico_('Fallback legado de período via data_inicio_desafio/data_fim_desafio usado.', contextoLog);
  }

  if (!periodoCompletoValido_(periodo)) {
    logMeuGiroDiagnostico_('Desafio sem período de apuração válido para filtro histórico.', contextoLog);
  }

  return {
    inicio: periodo.inicio || '',
    fim: periodo.fim || '',
    periodo_desafio: periodoTexto || (periodoLista && periodoLista.periodo_desafio) || '',
    nome_desafio: (periodoLista && periodoLista.nome_desafio) || ''
  };
}

function obterVinculosDesafioUsuario_(idDgmb) {
  if (typeof painelMG_incrementarAuditoriaCarregamentoInicial_ === 'function') {
    painelMG_incrementarAuditoriaCarregamentoInicial_('obterVinculosDesafioUsuario_chamadas');
  }
  var perfTotalInicio = meuGiroPerfNow_();
  var perfEtapaInicio = perfTotalInicio;
  var logsDiagnosticoInicio = MEU_GIRO_DIAGNOSTICO_LOGS_EXECUCAO_ || 0;
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var ss = getSpreadsheet_();
  var contextoLista = buildListaDesafiosContexto_(ss);
  var periodos = contextoLista.periodos;
  var statusListaDesafios = contextoLista.status;
  var abaDesafio = SHEETS.DESAFIO || 'dgmbDesafios';
  var sh = ss.getSheetByName(abaDesafio);
  if (!sh) return [];

  perfEtapaInicio = meuGiroPerfNow_();
  if (typeof painelMG_incrementarAuditoriaCarregamentoInicial_ === 'function') {
    painelMG_incrementarAuditoriaCarregamentoInicial_('leituras_dgmbDesafios');
  }
  var values = sh.getDataRange().getValues();
  meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'leitura_dgmbDesafios', perfEtapaInicio, {
    quantidade_linhas_dgmbDesafios: values && values.length ? values.length - 1 : 0
  });
  if (!values || values.length < 2) return [];

  var map = buildHeaderMap_(values[0]);
  var idxId = getOptionalColumnIndex_(map, ['id_dgmb']);
  if (idxId === -1) return [];

  var idxMeta = getOptionalColumnIndex_(map, ['distancia_km', 'distancia km']);
  var idxInscricao = getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxIdDesafio = getIdDesafioColumnIndex_(map);
  var idxObs = getOptionalColumnIndex_(map, ['observacao', 'observação']);
  var idxItem = getOptionalColumnIndex_(map, ['id_item_estoque', 'id item estoque']);
  var idxTipoDesafio = getOptionalColumnIndex_(map, ['tipo_do_desafio', 'tipo do desafio', 'tipo_desafio', 'tipo desafio']);
  var idxStatusDesafio = getOptionalColumnIndex_(map, ['status_desafio', 'status desafio']);
  var idxStatusUsuarioDesafio = getOptionalColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio']);
  var idxStatusValidacaoCertificado = getOptionalColumnIndex_(map, ['status_validacao_certificado']);
  var idxStatusPag = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagamento', 'pix_status']);
  var idxStatusInscricao = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição']);
  var idxConfirmacao = getOptionalColumnIndex_(map, ['confirmacao', 'confirmação', 'confirmado', 'inscricao_confirmada']);
  var idxPeriodoHistorico = getOptionalColumnIndex_(map, ['periodo_desafio', 'periodo desafio', 'período_desafio', 'período desafio']);
  var idxInicioHistorico = getOptionalColumnIndex_(map, ['data_inicio_desafio', 'data inicio desafio', 'data início desafio']);
  var idxFimHistorico = getOptionalColumnIndex_(map, ['data_fim_desafio', 'data fim desafio']);

  var vinculos = [];
  var chaves = {};

  perfEtapaInicio = meuGiroPerfNow_();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = normalizeText_(row[idxId]);
    if (rowId !== id) continue;

    var idDesafio = obterIdDesafioRegistro_(row, idxIdDesafio, idxObs);
    var idInscricao = idxInscricao > -1 ? normalizeText_(row[idxInscricao]) : '';
    var idItem = idxItem > -1 ? normalizeText_(row[idxItem]) : '';
    var tipoDesafio = idxTipoDesafio > -1 ? normalizeText_(row[idxTipoDesafio]) : '';
    var tipoSemAcento = tipoDesafio.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    var ehNormal = tipoSemAcento === 'normal';
    var metaKm = idxMeta > -1 ? parseLocalizedNumber_(row[idxMeta]) : 0;

    var statusInscricao = idxStatusInscricao > -1 ? normalizeText_(row[idxStatusInscricao]) : '';
    var statusConfirmacao = idxConfirmacao > -1 ? normalizeText_(row[idxConfirmacao]) : '';
    var statusPagamento = idxStatusPag > -1 ? normalizeText_(row[idxStatusPag]) : '';
    var statusDesafio = idxStatusDesafio > -1 ? normalizeText_(row[idxStatusDesafio]) : '';
    var statusUsuarioDesafio = idxStatusUsuarioDesafio > -1 ? normalizeText_(row[idxStatusUsuarioDesafio]) : '';
    var statusValidacaoCertificado = idxStatusValidacaoCertificado > -1 ? normalizeText_(row[idxStatusValidacaoCertificado]) : '';

    if (!idDesafio) {
      logMeuGiroDiagnostico_('Vínculo sem ID_DESAFIO identificável.', {
        id_dgmb: id,
        linha: i + 1,
        id_item_estoque: idItem || ''
      });
    }

    var validacao = validarInscricaoMinima_({
      status_inscricao: statusInscricao || statusUsuarioDesafio,
      status_confirmacao: statusConfirmacao,
      status_pagamento: statusPagamento
    });
    var aptoBase = validacao.valida && !inscricaoTemBloqueioMinimo_(statusUsuarioDesafio);
    var statusLista = idDesafio ? (statusListaDesafios.byId[idDesafio] || '') : '';

    if (ehNormal && idDesafio && statusLista && statusLista !== 'ativo') {
      logMeuGiroDiagnostico_('Desafio histórico preservado mesmo com ListaDesafios.Status inativo.', {
        id_dgmb: id,
        id_desafio: idDesafio,
        status_lista: statusLista
      });
    }

    // Para o Meu Giro, vínculo normal em dgmbDesafios não é bloqueado por ListaDesafios.Status.
    // ListaDesafios segue como fallback de período/status visual e permanece preservada para inscrições.
    var apto = ehNormal
      ? aptoBase && !!idDesafio && metaKm > 0
      : aptoBase;

    var periodoLista = (idDesafio && periodos.byId[idDesafio]) || (!ehNormal && periodos.byAba[abaDesafio]) || { inicio: '', fim: '', nome_desafio: '' };
    var periodo = montarPeriodoHistoricoVinculo_(row, {
      periodo: idxPeriodoHistorico,
      inicio: idxInicioHistorico,
      fim: idxFimHistorico
    }, periodoLista, {
      id_dgmb: id,
      id_desafio: idDesafio || '',
      linha: i + 1
    });

    var chave = [id, idInscricao, idDesafio, idItem || ('META_' + Math.round((metaKm + Number.EPSILON) * 10) / 10)].join('|');
    if (chaves[chave]) continue;
    chaves[chave] = true;

    vinculos.push({
      id_dgmb: id,
      id_inscricao: idInscricao,
      id_desafio: idDesafio,
      id_item_estoque: idItem,
      meta_km: metaKm,
      status_desafio: statusDesafio,
      status_usuario_desafio: statusUsuarioDesafio,
      status_pagamento: statusPagamento,
      status_lista_desafios: statusLista,
      status_validacao_certificado: statusValidacaoCertificado,
      apto: apto,
      periodo_inicio: periodo.inicio || '',
      periodo_fim: periodo.fim || '',
      periodo_desafio: periodo.periodo_desafio || '',
      nome_desafio: periodo.nome_desafio || abaDesafio || '',
      aba_desafio: abaDesafio
    });
  }

  meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'montagem_vinculos', perfEtapaInicio, {
    quantidade_linhas_dgmbDesafios: values.length - 1,
    quantidade_vinculos_do_usuario: vinculos.length,
    quantidade_logs_diagnostico: (MEU_GIRO_DIAGNOSTICO_LOGS_EXECUCAO_ || 0) - logsDiagnosticoInicio
  });
  meuGiroPerfLog_('obter-vinculos-desafio-usuario', 'obterVinculosDesafioUsuario_total', perfTotalInicio, {
    quantidade_linhas_dgmbDesafios: values.length - 1,
    quantidade_vinculos_do_usuario: vinculos.length,
    quantidade_logs_diagnostico: (MEU_GIRO_DIAGNOSTICO_LOGS_EXECUCAO_ || 0) - logsDiagnosticoInicio
  });

  return vinculos;
}

function obterActivityIdRegistroKm_(registro) {
  return normalizeText_(firstFilledValue_(registro || {}, [
    'activity_id', 'Activity_ID', 'activity id', 'id_atividade', 'ID_Atividade'
  ]));
}

function obterRegistrosKmUsuario_(idDgmb, opcoes) {
  var perfTotalInicio = meuGiroPerfNow_();
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var perfEtapaInicio = meuGiroPerfNow_();
  var registrosContexto = obterRegistrosKmObjetosReaproveitados_(id, opcoes);
  var registros = registrosContexto.valores;
  meuGiroPerfLog_('obter-registros-km-usuario', registrosContexto.reaproveitados ? 'leitura_REGISTRO_KM_reaproveitada' : 'leitura_REGISTRO_KM', perfEtapaInicio, {
    quantidade_linhas_registro_km: registros.length
  });
  var out = [];
  var activityIdsIncluidos = {};

  for (var i = 0; i < registros.length; i++) {
    var r = registros[i];
    var rowId = normalizeText_(firstFilledValue_(r, ['ID_DGMB', 'id_dgmb']));
    if (rowId !== id) continue;

    var activityId = obterActivityIdRegistroKm_(r);
    if (activityId) {
      if (activityIdsIncluidos[activityId]) continue;
      activityIdsIncluidos[activityId] = true;
    }

    out.push({
      data_atividade: normalizarDataISO_(firstFilledValue_(r, ['Data_Atividade', 'Data', 'data_atividade', 'data'])),
      km: parseLocalizedNumber_(firstFilledValue_(r, ['KM', 'km']))
    });
  }

  meuGiroPerfLog_('obter-registros-km-usuario', 'obterRegistrosKmUsuario_total', perfTotalInicio, {
    quantidade_registros_usuario: out.length
  });
  return out;
}

function meuGiroResumoHeaders_() {
  return [
    'Timestamp_Atualizacao',
    'ID_INSCRICAO',
    'ID_DGMB',
    'ID_DESAFIO',
    'id_item_estoque',
    'Meta_KM',
    'Distancia_Realizada',
    'Percentual_Concluido',
    'Status_Apuracao'
  ];
}

function meuGiroResumoObterLayout_(headerRow, sheetName) {
  var map = buildHeaderMap_(headerRow || []);
  var camposObrigatorios = [
    ['timestamp_atualizacao'],
    ['id_dgmb'],
    ['id_desafio'],
    ['id_item_estoque', 'id item estoque'],
    ['meta_km', 'meta km'],
    ['distancia_realizada', 'distancia realizada'],
    ['percentual_concluido', 'percentual concluido', 'percentual concluído'],
    ['status_apuracao', 'status apuracao', 'status apuração']
  ];

  for (var i = 0; i < camposObrigatorios.length; i++) {
    if (getOptionalColumnIndex_(map, camposObrigatorios[i]) === -1) {
      throw new Error(
        'Estrutura inválida na aba ' + sheetName +
        '. Cabeçalho obrigatório não encontrado: ' + camposObrigatorios[i][0]
      );
    }
  }

  return {
    map: map,
    possuiIdInscricao: getOptionalColumnIndex_(map, ['id_inscricao', 'id inscrição', 'id inscricao']) > -1
  };
}

function ensureMeuGiroResumoSheet_() {
  var ss = getSpreadsheet_();
  var sheetName = SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO';
  var sh = ss.getSheetByName(sheetName);
  var headers = meuGiroResumoHeaders_();

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  var ultimaColuna = Math.max(sh.getLastColumn(), 1);
  var atual = sh.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  var possuiCabecalho = false;
  for (var i = 0; i < atual.length; i++) {
    if (normalizeText_(atual[i])) {
      possuiCabecalho = true;
      break;
    }
  }

  if (!possuiCabecalho) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  // Aceita tanto a estrutura legada quanto a futura. A validação por nome evita
  // sobrescrever ou deslocar dados durante a inserção manual de ID_INSCRICAO.
  meuGiroResumoObterLayout_(atual, sheetName);
  return sh;
}

function meuGiroResumoBuildChave_(idDgmb, idDesafio, idItemEstoque, metaKm, idInscricao) {
  var inscricao = normalizeText_(idInscricao);
  if (inscricao) return 'INSCRICAO|' + inscricao;

  var id = normalizeText_(idDgmb);
  var desafio = normalizeText_(idDesafio);
  var item = normalizeText_(idItemEstoque);
  var meta = Math.round((parseLocalizedNumber_(metaKm) + Number.EPSILON) * 10) / 10;

  // Fallback legado para abas sem a coluna ID_INSCRICAO e para linhas ou
  // vínculos em que esse identificador ainda não esteja preenchido.
  return [id, desafio, item || ('META_' + meta)].join('|');
}


function obterMeuGiroResumoAtualizado_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var ss = getSpreadsheet_();
  var sheetName = SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO';
  var shResumo = ss.getSheetByName(sheetName);
  if (!shResumo) return [];

  var valoresResumo = shResumo.getDataRange().getValues();
  if (!valoresResumo || valoresResumo.length < 2) return [];

  var layoutResumo = meuGiroResumoObterLayout_(valoresResumo[0] || [], sheetName);
  var mapResumo = layoutResumo.map;
  var idxInscricaoResumo = getOptionalColumnIndex_(mapResumo, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxId = getOptionalColumnIndex_(mapResumo, ['id_dgmb']);
  var idxDesafio = getOptionalColumnIndex_(mapResumo, ['id_desafio']);
  var idxItem = getOptionalColumnIndex_(mapResumo, ['id_item_estoque', 'id item estoque']);
  var idxMetaResumo = getOptionalColumnIndex_(mapResumo, ['meta_km', 'meta km']);
  var idxDistanciaResumo = getOptionalColumnIndex_(mapResumo, ['distancia_realizada', 'distancia realizada']);
  var idxPercentualResumo = getOptionalColumnIndex_(mapResumo, ['percentual_concluido', 'percentual concluido', 'percentual concluído']);
  var idxStatusResumo = getOptionalColumnIndex_(mapResumo, ['status_apuracao', 'status apuracao', 'status apuração']);

  var vinculos = obterVinculosDesafioUsuario_(id) || [];
  var vinculosPorChave = {};
  for (var v = 0; v < vinculos.length; v++) {
    var vinculo = vinculos[v] || {};
    var chaveVinculo = meuGiroResumoBuildChave_(
      id,
      vinculo.id_desafio,
      vinculo.id_item_estoque,
      vinculo.meta_km,
      idxInscricaoResumo > -1 ? vinculo.id_inscricao : ''
    );
    vinculosPorChave[chaveVinculo] = vinculo;
  }

  var saida = [];
  for (var i = 1; i < valoresResumo.length; i++) {
    var row = valoresResumo[i] || [];
    if (normalizeText_(row[idxId]) !== id) continue;

    var idInscricao = idxInscricaoResumo > -1 ? normalizeText_(row[idxInscricaoResumo]) : '';
    var meta = parseLocalizedNumber_(row[idxMetaResumo]);
    var chave = meuGiroResumoBuildChave_(
      row[idxId],
      row[idxDesafio],
      row[idxItem],
      meta,
      idxInscricaoResumo > -1 ? idInscricao : ''
    );
    var vinculoAtual = vinculosPorChave[chave] || {};

    saida.push({
      id_inscricao: idInscricao,
      id_dgmb: id,
      id_desafio: normalizeText_(row[idxDesafio]),
      id_item_estoque: normalizeText_(row[idxItem]),
      nome_desafio: normalizeText_(vinculoAtual.nome_desafio),
      meta_km: Math.round((meta + Number.EPSILON) * 10) / 10,
      distancia_realizada: Math.round((parseLocalizedNumber_(row[idxDistanciaResumo]) + Number.EPSILON) * 10) / 10,
      percentual_concluido: Math.round((parseLocalizedNumber_(row[idxPercentualResumo]) + Number.EPSILON) * 10) / 10,
      status_apuracao: normalizeText_(row[idxStatusResumo]),
      status_validacao_certificado: normalizeText_(vinculoAtual.status_validacao_certificado).toUpperCase(),
      status_desafio: normalizeText_(vinculoAtual.status_desafio),
      status_usuario_desafio: normalizeText_(vinculoAtual.status_usuario_desafio),
      status_pagamento: normalizeText_(vinculoAtual.status_pagamento),
      status_lista_desafios: normalizeText_(vinculoAtual.status_lista_desafios),
      periodo_inicio: normalizarDataISO_(vinculoAtual.periodo_inicio) || '',
      periodo_fim: normalizarDataISO_(vinculoAtual.periodo_fim) || '',
      periodo_desafio: normalizeText_(vinculoAtual.periodo_desafio)
    });
  }

  return saida;
}


function obterMeuGiroResumoAtualizadoLeve_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var ss = getSpreadsheet_();
  var sheetName = SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO';
  var shResumo = ss.getSheetByName(sheetName);
  if (!shResumo) return [];

  var valoresResumo = shResumo.getDataRange().getValues();
  if (!valoresResumo || valoresResumo.length < 2) return [];

  var layoutResumo = meuGiroResumoObterLayout_(valoresResumo[0] || [], sheetName);
  var mapResumo = layoutResumo.map;
  var idxInscricaoResumo = getOptionalColumnIndex_(mapResumo, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxId = getOptionalColumnIndex_(mapResumo, ['id_dgmb']);
  var idxDesafio = getOptionalColumnIndex_(mapResumo, ['id_desafio']);
  var idxItem = getOptionalColumnIndex_(mapResumo, ['id_item_estoque', 'id item estoque']);
  var idxMetaResumo = getOptionalColumnIndex_(mapResumo, ['meta_km', 'meta km']);
  var idxDistanciaResumo = getOptionalColumnIndex_(mapResumo, ['distancia_realizada', 'distancia realizada']);
  var idxPercentualResumo = getOptionalColumnIndex_(mapResumo, ['percentual_concluido', 'percentual concluido', 'percentual concluído']);
  var idxStatusResumo = getOptionalColumnIndex_(mapResumo, ['status_apuracao', 'status apuracao', 'status apuração']);

  if (idxId === -1 || idxDesafio === -1 || idxItem === -1 || idxMetaResumo === -1 || idxDistanciaResumo === -1 || idxPercentualResumo === -1 || idxStatusResumo === -1) {
    return obterMeuGiroResumoAtualizado_(id);
  }

  var saida = [];
  for (var i = 1; i < valoresResumo.length; i++) {
    var row = valoresResumo[i] || [];
    if (normalizeText_(row[idxId]) !== id) continue;

    var meta = parseLocalizedNumber_(row[idxMetaResumo]);
    saida.push({
      id_inscricao: idxInscricaoResumo > -1 ? normalizeText_(row[idxInscricaoResumo]) : '',
      id_dgmb: id,
      id_desafio: normalizeText_(row[idxDesafio]),
      id_item_estoque: normalizeText_(row[idxItem]),
      nome_desafio: '',
      meta_km: Math.round((meta + Number.EPSILON) * 10) / 10,
      distancia_realizada: Math.round((parseLocalizedNumber_(row[idxDistanciaResumo]) + Number.EPSILON) * 10) / 10,
      percentual_concluido: Math.round((parseLocalizedNumber_(row[idxPercentualResumo]) + Number.EPSILON) * 10) / 10,
      status_apuracao: normalizeText_(row[idxStatusResumo]),
      status_validacao_certificado: '',
      status_desafio: '',
      status_usuario_desafio: '',
      status_pagamento: '',
      status_lista_desafios: '',
      periodo_inicio: '',
      periodo_fim: '',
      periodo_desafio: ''
    });
  }

  return saida;
}

function meuGiroResumoAgruparLinhasContiguas_(linhas) {
  var blocos = [];
  for (var i = 0; i < linhas.length; i++) {
    var numeroLinha = linhas[i];
    var blocoAtual = blocos.length ? blocos[blocos.length - 1] : null;
    if (blocoAtual && numeroLinha === blocoAtual.linhaFinal + 1) {
      blocoAtual.linhaFinal = numeroLinha;
      blocoAtual.quantidadeLinhas++;
    } else {
      blocos.push({
        linhaInicial: numeroLinha,
        linhaFinal: numeroLinha,
        quantidadeLinhas: 1
      });
    }
  }
  return blocos;
}

function meuGiroResumoLeituraIntegralFallback_(shResumo) {
  return shResumo.getDataRange().getValues();
}

function meuGiroResumoLerLinhasAlvo_(
  shResumo,
  cabecalho,
  totalColunasResumo,
  id,
  vinculos,
  idxId,
  idxInscricaoResumo
) {
  var perfIndicesInicio = meuGiroPerfNow_();
  var ultimaLinha = shResumo.getLastRow();
  var valoresResumo = [];
  valoresResumo[0] = cabecalho;
  var linhasCandidatas = [];
  var quantidadeLinhasConsultadas = Math.max(ultimaLinha - 1, 0);
  var usouIdInscricao = idxInscricaoResumo > -1;

  if (quantidadeLinhasConsultadas > 0) {
    var idsDgmb = shResumo.getRange(2, idxId + 1, quantidadeLinhasConsultadas, 1).getValues();
    var idsInscricao = usouIdInscricao
      ? shResumo.getRange(2, idxInscricaoResumo + 1, quantidadeLinhasConsultadas, 1).getValues()
      : [];
    var inscricoesDoAtleta = {};

    if (usouIdInscricao) {
      for (var v = 0; v < vinculos.length; v++) {
        var idInscricaoVinculo = normalizeText_((vinculos[v] || {}).id_inscricao);
        if (idInscricaoVinculo) inscricoesDoAtleta[idInscricaoVinculo] = true;
      }
    }

    for (var i = 0; i < quantidadeLinhasConsultadas; i++) {
      var pertenceAoId = normalizeText_((idsDgmb[i] || [])[0]) === id;
      var idInscricaoLinha = usouIdInscricao
        ? normalizeText_((idsInscricao[i] || [])[0])
        : '';
      if (pertenceAoId || (idInscricaoLinha && inscricoesDoAtleta[idInscricaoLinha])) {
        linhasCandidatas.push(i + 2);
      }
    }
  }

  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'leitura_MEU_GIRO_RESUMO_indices', perfIndicesInicio, {
    quantidade_linhas_consultadas: quantidadeLinhasConsultadas,
    usou_id_inscricao: usouIdInscricao
  });

  var perfLinhasAlvoInicio = meuGiroPerfNow_();
  linhasCandidatas.sort(function(a, b) { return a - b; });
  var blocos = meuGiroResumoAgruparLinhasContiguas_(linhasCandidatas);
  for (var b = 0; b < blocos.length; b++) {
    var bloco = blocos[b];
    var valoresBloco = shResumo.getRange(
      bloco.linhaInicial,
      1,
      bloco.quantidadeLinhas,
      totalColunasResumo
    ).getValues();
    for (var linhaBloco = 0; linhaBloco < valoresBloco.length; linhaBloco++) {
      valoresResumo[bloco.linhaInicial + linhaBloco - 1] = valoresBloco[linhaBloco];
    }
  }

  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'leitura_MEU_GIRO_RESUMO_linhas_alvo', perfLinhasAlvoInicio, {
    quantidade_linhas_completas_lidas: linhasCandidatas.length,
    quantidade_blocos_lidos: blocos.length,
    quantidade_celulas_estimadas: (quantidadeLinhasConsultadas * (usouIdInscricao ? 2 : 1)) +
      (linhasCandidatas.length * totalColunasResumo)
  });

  return {
    valores: valoresResumo,
    linhasCandidatas: linhasCandidatas,
    fallbackIntegral: false
  };
}

function atualizarMeuGiroResumo_(idDgmb, opcoes) {
  var perfTotalInicio = meuGiroPerfNow_();
  var id = normalizeText_(idDgmb);
  if (!id) return [];

  var perfEtapaInicio = meuGiroPerfNow_();
  var vinculos = obterVinculosDesafioUsuario_(id);
  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'obterVinculosDesafioUsuario_', perfEtapaInicio, {
    quantidade_vinculos: vinculos.length
  });
  perfEtapaInicio = meuGiroPerfNow_();
  var registros = obterRegistrosKmUsuario_(id, opcoes);
  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'obterRegistrosKmUsuario_', perfEtapaInicio, {
    quantidade_registros_usuario: registros.length
  });
  perfEtapaInicio = meuGiroPerfNow_();
  var shResumo = ensureMeuGiroResumoSheet_();
  var totalColunasResumo = Math.max(shResumo.getLastColumn(), 1);
  var cabecalhoResumo = shResumo.getRange(1, 1, 1, totalColunasResumo).getValues()[0] || [];
  var layoutResumo = meuGiroResumoObterLayout_(
    cabecalhoResumo,
    SHEETS.MEU_GIRO_RESUMO || 'MEU_GIRO_RESUMO'
  );
  var mapResumo = layoutResumo.map;
  var idxTimestamp = getOptionalColumnIndex_(mapResumo, ['timestamp_atualizacao']);
  var idxInscricaoResumo = getOptionalColumnIndex_(mapResumo, ['id_inscricao', 'id inscrição', 'id inscricao']);
  var idxId = getOptionalColumnIndex_(mapResumo, ['id_dgmb']);
  var idxDesafio = getOptionalColumnIndex_(mapResumo, ['id_desafio']);
  var idxItem = getOptionalColumnIndex_(mapResumo, ['id_item_estoque', 'id item estoque']);
  var idxMetaResumo = getOptionalColumnIndex_(mapResumo, ['meta_km', 'meta km']);
  var idxDistanciaResumo = getOptionalColumnIndex_(mapResumo, ['distancia_realizada', 'distancia realizada']);
  var idxPercentualResumo = getOptionalColumnIndex_(mapResumo, ['percentual_concluido', 'percentual concluido', 'percentual concluído']);
  var idxStatusResumo = getOptionalColumnIndex_(mapResumo, ['status_apuracao', 'status apuracao', 'status apuração']);
  var leituraResumo;
  try {
    leituraResumo = meuGiroResumoLerLinhasAlvo_(
      shResumo,
      cabecalhoResumo,
      totalColunasResumo,
      id,
      vinculos,
      idxId,
      idxInscricaoResumo
    );
  } catch (erroLeituraCirurgica) {
    leituraResumo = {
      valores: meuGiroResumoLeituraIntegralFallback_(shResumo),
      linhasCandidatas: null,
      fallbackIntegral: true
    };
    meuGiroPerfLog_('atualizar-meu-giro-resumo', 'leitura_MEU_GIRO_RESUMO_fallback_integral', perfEtapaInicio, {
      erro: String(erroLeituraCirurgica)
    });
  }
  var valoresResumo = leituraResumo.valores;
  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'leitura_MEU_GIRO_RESUMO', perfEtapaInicio, {
    quantidade_linhas_meu_giro_resumo: Math.max(shResumo.getLastRow() - 1, 0),
    leitura_integral_fallback: leituraResumo.fallbackIntegral
  });
  var linhasPorChave = {};
  var linhasParaIndexar = leituraResumo.linhasCandidatas;
  if (!linhasParaIndexar) {
    linhasParaIndexar = [];
    for (var linhaIntegral = 2; linhaIntegral <= valoresResumo.length; linhaIntegral++) {
      linhasParaIndexar.push(linhaIntegral);
    }
  }

  for (var i = 0; i < linhasParaIndexar.length; i++) {
    var numeroLinhaExistente = linhasParaIndexar[i];
    var row = valoresResumo[numeroLinhaExistente - 1] || [];
    var idInscricaoExistente = idxInscricaoResumo > -1
      ? normalizeText_(row[idxInscricaoResumo])
      : '';
    if (!idInscricaoExistente && normalizeText_(row[idxId]) !== id) continue;

    var chaveExistente = meuGiroResumoBuildChave_(
      row[idxId],
      row[idxDesafio],
      row[idxItem],
      row[idxMetaResumo],
      idInscricaoExistente
    );
    linhasPorChave[chaveExistente] = numeroLinhaExistente;
  }

  var hoje = normalizarDataISO_(new Date());
  var saida = [];
  var quantidadeEscritasResumo = 0;
  var duracaoEscritasResumoMs = 0;

  for (var v = 0; v < vinculos.length; v++) {
    var vinculo = vinculos[v];
    var idInscricao = normalizeText_(vinculo.id_inscricao);
    var meta = Number(vinculo.meta_km || 0);
    var metaArredondada = Math.round((meta + Number.EPSILON) * 10) / 10;
    var chave = meuGiroResumoBuildChave_(
      id,
      vinculo.id_desafio,
      vinculo.id_item_estoque,
      metaArredondada,
      idxInscricaoResumo > -1 ? idInscricao : ''
    );
    var inicio = normalizarDataISO_(vinculo.periodo_inicio);
    var fim = normalizarDataISO_(vinculo.periodo_fim);
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

    var distanciaArredondada = Math.round((distancia + Number.EPSILON) * 10) / 10;
    var percentualArredondado = Math.round((percentual + Number.EPSILON) * 10) / 10;
    var numeroLinha = linhasPorChave[chave] || 0;
    var rowAtual = numeroLinha ? (valoresResumo[numeroLinha - 1] || []) : [];
    var houveMudanca = !numeroLinha ||
      (idxInscricaoResumo > -1 && normalizeText_(rowAtual[idxInscricaoResumo]) !== idInscricao) ||
      parseLocalizedNumber_(rowAtual[idxMetaResumo]) !== metaArredondada ||
      parseLocalizedNumber_(rowAtual[idxDistanciaResumo]) !== distanciaArredondada ||
      parseLocalizedNumber_(rowAtual[idxPercentualResumo]) !== percentualArredondado ||
      normalizeText_(rowAtual[idxStatusResumo]) !== status;

    if (houveMudanca) {
      var linha = [];
      for (var c = 0; c < totalColunasResumo; c++) {
        linha[c] = numeroLinha ? rowAtual[c] : '';
      }
      linha[idxTimestamp] = new Date();
      if (idxInscricaoResumo > -1) linha[idxInscricaoResumo] = idInscricao;
      linha[idxId] = id;
      linha[idxDesafio] = vinculo.id_desafio;
      linha[idxItem] = vinculo.id_item_estoque;
      linha[idxMetaResumo] = metaArredondada;
      linha[idxDistanciaResumo] = distanciaArredondada;
      linha[idxPercentualResumo] = percentualArredondado;
      linha[idxStatusResumo] = status;

      var perfEscritaResumoInicio = meuGiroPerfNow_();
      if (numeroLinha) {
        shResumo.getRange(numeroLinha, 1, 1, totalColunasResumo).setValues([linha]);
      } else {
        shResumo.appendRow(linha);
      }
      duracaoEscritasResumoMs += meuGiroPerfNow_() - perfEscritaResumoInicio;
      quantidadeEscritasResumo++;
    }

    saida.push({
      id_inscricao: idInscricao,
      id_dgmb: id,
      id_desafio: vinculo.id_desafio,
      id_item_estoque: vinculo.id_item_estoque,
      nome_desafio: vinculo.nome_desafio || '',
      meta_km: metaArredondada,
      distancia_realizada: distanciaArredondada,
      percentual_concluido: percentualArredondado,
      status_apuracao: status,
      status_validacao_certificado: normalizeText_(vinculo.status_validacao_certificado).toUpperCase(),
      status_desafio: normalizeText_(vinculo.status_desafio),
      status_usuario_desafio: normalizeText_(vinculo.status_usuario_desafio),
      status_pagamento: normalizeText_(vinculo.status_pagamento),
      status_lista_desafios: normalizeText_(vinculo.status_lista_desafios),
      periodo_inicio: inicio || '',
      periodo_fim: fim || '',
      periodo_desafio: normalizeText_(vinculo.periodo_desafio)
    });
  }

  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'escrita_MEU_GIRO_RESUMO', meuGiroPerfNow_() - duracaoEscritasResumoMs, {
    quantidade_linhas_escritas: quantidadeEscritasResumo,
    quantidade_vinculos_processados: vinculos.length
  });
  meuGiroPerfLog_('atualizar-meu-giro-resumo', 'atualizarMeuGiroResumo_total', perfTotalInicio, {
    quantidade_vinculos: vinculos.length,
    quantidade_registros_usuario: registros.length,
    quantidade_linhas_escritas: quantidadeEscritasResumo
  });
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

  var idxStatusUsuarioDesafio = getOptionalColumnIndex_(map, ['status_usuario_desafio', 'status usuário desafio', 'status usuario desafio']);
  var idxStatusPag = getOptionalColumnIndex_(map, ['status_pagamento', 'pagamento_status', 'pagamento', 'pix_status']);
  var idxStatusInscricao = getOptionalColumnIndex_(map, ['status_inscricao', 'status inscrição']);
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
    var statusUsuarioDesafio = idxStatusUsuarioDesafio > -1 ? normalizeText_(row[idxStatusUsuarioDesafio]) : '';
    var validacao = validarInscricaoMinima_({
      status_inscricao: statusInscricao || statusUsuarioDesafio,
      status_confirmacao: statusConfirmacao,
      status_pagamento: statusPagamento
    });
    var apto = validacao.valida && !inscricaoTemBloqueioMinimo_(statusUsuarioDesafio);
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
