/**
 * FrasesService.gs
 * Camada isolada e segura para frases motivacionais.
 * NÃO altera lógica existente.
 * NÃO depende de função antiga.
 * Sempre retorna fallback seguro.
 */

/** =========================
 *  CONFIG / FALLBACK
 *  ========================= */
var FRASE_CONTEXTOS_ = {
  COMECANDO: 'COMECANDO_O_DESAFIO',
  RITMO_IDEAL: 'RITMO_IDEAL',
  ATRASADO: 'UM_POUCO_ATRASADO',
  ADIANTADO: 'MUITO_ADIANTADO',
  FASE_FINAL: 'FASE_FINAL',
  CONCLUIDO: 'DESAFIO_CONCLUIDO'
};

var FRASES_FALLBACK_ = {
  COMECANDO_O_DESAFIO: [
    'Todo grande desafio começa com os primeiros quilômetros. Continue.',
    'Você já começou. Agora é seguir em frente, um pedal por vez.',
    'O importante não é largar perfeito, é continuar em movimento.'
  ],
  RITMO_IDEAL: [
    'Você está no ritmo certo. Continue constante.',
    'Seu progresso está dentro do esperado. Mantenha a regularidade.',
    'Pedalando assim, você segue firme rumo à meta.'
  ],
  UM_POUCO_ATRASADO: [
    'Ainda dá tempo de recuperar. Um passo de cada vez.',
    'Seu desafio continua totalmente possível. Retome o ritmo.',
    'Não desanime. Pequenos avanços agora fazem diferença no final.'
  ],
  MUITO_ADIANTADO: [
    'Excelente ritmo. Você está muito à frente do esperado.',
    'Parabéns! Seu desempenho está acima do planejado.',
    'Você construiu uma ótima vantagem. Continue focado.'
  ],
  FASE_FINAL: [
    'Falta pouco. Seu objetivo já está logo ali.',
    'Você entrou na fase final. Continue firme até completar.',
    'Agora é manter o foco: o final do desafio está próximo.'
  ],
  DESAFIO_CONCLUIDO: [
    'Parabéns! Você concluiu seu desafio com sucesso.',
    'Meta atingida! Seu esforço valeu a pena.',
    'Desafio concluído. Resultado construído com consistência.'
  ],
  GENERICA: [
    'Siga firme no seu desafio.',
    'Cada quilômetro conta. Continue.',
    'Seu progresso é construído pedal por pedal.'
  ]
};

/** =========================
 *  API PRINCIPAL
 *  ========================= */

/**
 * Retorna uma frase motivacional segura para o atleta/painel.
 * Nunca lança erro para fora.
 */
function getFraseMotivacionalSegura_(painel) {
  try {
    var contexto = identificarContextoFrase_(painel);
    var frasesPorContexto = carregarFrasesMotivacionaisDaPlanilha_();
    var lista = frasesPorContexto[contexto];

    if (!lista || !lista.length) {
      lista = FRASES_FALLBACK_[contexto] || FRASES_FALLBACK_.GENERICA;
    }

    return escolherFraseDeterministica_(lista, painel, contexto);
  } catch (err) {
    return FRASES_FALLBACK_.GENERICA[0];
  }
}

/**
 * Opcional: monta payload com aliases para o frontend não quebrar,
 * independente do nome que ele já esteja lendo.
 */
function montarPayloadFraseMotivacional_(painel) {
  var frase = getFraseMotivacionalSegura_(painel);
  var contexto = identificarContextoFrase_(painel);

  return {
    frase_motivacional: frase,
    fraseMotivacional: frase,
    contexto_frase: contexto
  };
}

/** =========================
 *  IDENTIFICAÇÃO DE CONTEXTO
 *  ========================= */

function identificarContextoFrase_(painel) {
  var meta = extrairNumeroPainel_(painel, ['meta', 'meta_km', 'km_meta', 'objetivo', 'objetivo_km']);
  var realizado = extrairNumeroPainel_(painel, ['realizado', 'realizado_km', 'km_realizado', 'total_km']);
  var restante = extrairNumeroPainel_(painel, ['restante', 'restante_km', 'km_restante']);
  var percentual = extrairNumeroPainel_(painel, ['percentual', 'percentual_concluido', 'progresso_percentual']);
  var ritmoStatus = extrairTextoPainel_(painel, ['ritmo_status', 'status_ritmo', 'alerta_ritmo', 'ritmo']);
  var hoje = new Date();

  if (!isFinite(meta) || meta <= 0) meta = 0;
  if (!isFinite(realizado) || realizado < 0) realizado = 0;

  if (!isFinite(restante)) {
    restante = meta > 0 ? Math.max(meta - realizado, 0) : 0;
  }

  if (!isFinite(percentual)) {
    percentual = meta > 0 ? (realizado / meta) * 100 : 0;
  }

  if (percentual >= 100 || restante <= 0) {
    return FRASE_CONTEXTOS_.CONCLUIDO;
  }

  if (meta > 0) {
    var limiteFinalKm = Math.max(5, meta * 0.10); // últimos 10% ou no mínimo 5 km
    if (restante > 0 && (restante <= limiteFinalKm || percentual >= 90)) {
      return FRASE_CONTEXTOS_.FASE_FINAL;
    }
  }

  if (meta > 0) {
    var limiteInicioKm = Math.max(3, meta * 0.05); // até 5% da meta ou mínimo 3 km
    if (realizado <= limiteInicioKm) {
      return FRASE_CONTEXTOS_.COMECANDO;
    }
  } else if (realizado <= 3) {
    return FRASE_CONTEXTOS_.COMECANDO;
  }

  var statusNorm = normalizarTextoFrase_(ritmoStatus);

  if (statusNorm) {
    if (statusNorm.indexOf('atras') !== -1) {
      return FRASE_CONTEXTOS_.ATRASADO;
    }
    if (statusNorm.indexOf('adiant') !== -1) {
      return FRASE_CONTEXTOS_.ADIANTADO;
    }
    if (statusNorm.indexOf('em dia') !== -1) {
      return FRASE_CONTEXTOS_.RITMO_IDEAL;
    }
  }

  // Fallback por cálculo simples do ritmo ideal mensal,
  // sem depender de nenhuma função antiga do sistema.
  var deltaPercentual = calcularDeltaPercentualRitmo_(meta, realizado, hoje);

  if (deltaPercentual <= -8) {
    return FRASE_CONTEXTOS_.ATRASADO;
  }

  if (deltaPercentual >= 15) {
    return FRASE_CONTEXTOS_.ADIANTADO;
  }

  return FRASE_CONTEXTOS_.RITMO_IDEAL;
}

function calcularDeltaPercentualRitmo_(meta, realizado, dataRef) {
  if (!isFinite(meta) || meta <= 0) return 0;
  if (!isFinite(realizado) || realizado < 0) realizado = 0;

  var d = dataRef instanceof Date ? dataRef : new Date();
  var diaAtual = d.getDate();
  var ultimoDiaMes = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();

  if (!ultimoDiaMes || ultimoDiaMes <= 0) return 0;

  var percentualIdeal = (diaAtual / ultimoDiaMes) * 100;
  var percentualReal = (realizado / meta) * 100;

  return percentualReal - percentualIdeal;
}

/** =========================
 *  LEITURA SEGURA DA PLANILHA
 *  ========================= */

function carregarFrasesMotivacionaisDaPlanilha_() {
  var base = copiarMapaFrases_(FRASES_FALLBACK_);

  try {
    var sh = getFrasesSheetSafe_();
    if (!sh) return base;

    var values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return base;

    var header = values[0] || [];
    var map = buildHeaderMapFrases_(header);

    var idxContexto = firstHeaderIndex_(map, ['contexto', 'categoria', 'tipo', 'grupo', 'status']);
    var idxFrase = firstHeaderIndex_(map, ['frase', 'mensagem', 'texto', 'conteudo', 'descricao']);

    if (idxFrase < 0) {
      return base;
    }

    // Sem coluna de contexto, não arrisca comportamento inesperado.
    if (idxContexto < 0) {
      return base;
    }

    for (var i = 1; i < values.length; i++) {
      var row = values[i] || [];
      var contextoRaw = row[idxContexto];
      var fraseRaw = row[idxFrase];

      var contexto = mapearContextoPlanilha_(contextoRaw);
      var frase = limparFrase_(fraseRaw);

      if (!contexto || !frase) continue;

      if (!base[contexto]) base[contexto] = [];
      if (base[contexto].indexOf(frase) === -1) {
        base[contexto].push(frase);
      }
    }

    return base;
  } catch (err) {
    return base;
  }
}

function getFrasesSheetSafe_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    return ss.getSheetByName('FRASES');
  } catch (err) {
    return null;
  }
}

function buildHeaderMapFrases_(header) {
  var map = {};
  for (var i = 0; i < header.length; i++) {
    var key = normalizarTextoFrase_(header[i]);
    if (key) map[key] = i;
  }
  return map;
}

function firstHeaderIndex_(map, nomes) {
  for (var i = 0; i < nomes.length; i++) {
    var k = normalizarTextoFrase_(nomes[i]);
    if (map.hasOwnProperty(k)) return map[k];
  }
  return -1;
}

function mapearContextoPlanilha_(valor) {
  var v = normalizarTextoFrase_(valor);

  if (!v) return '';

  if (v === 'comecando' || v === 'comecando o desafio' || v === 'inicio' || v === 'iniciando') {
    return FRASE_CONTEXTOS_.COMECANDO;
  }
  if (v === 'ritmo ideal' || v === 'em dia' || v === 'ritmo') {
    return FRASE_CONTEXTOS_.RITMO_IDEAL;
  }
  if (v === 'um pouco atrasado' || v === 'atrasado' || v === 'abaixo do ritmo') {
    return FRASE_CONTEXTOS_.ATRASADO;
  }
  if (v === 'muito adiantado' || v === 'adiantado' || v === 'acima do ritmo') {
    return FRASE_CONTEXTOS_.ADIANTADO;
  }
  if (v === 'fase final' || v === 'final' || v === 'reta final') {
    return FRASE_CONTEXTOS_.FASE_FINAL;
  }
  if (v === 'desafio concluido' || v === 'concluido' || v === 'meta atingida' || v === 'finalizado') {
    return FRASE_CONTEXTOS_.CONCLUIDO;
  }

  return '';
}

/** =========================
 *  ESCOLHA DETERMINÍSTICA
 *  ========================= */

function escolherFraseDeterministica_(lista, painel, contexto) {
  if (!lista || !lista.length) return FRASES_FALLBACK_.GENERICA[0];
  if (lista.length === 1) return lista[0];

  var seedBase = '';
  seedBase += String(contexto || '');
  seedBase += '|';
  seedBase += String(extrairTextoPainel_(painel, ['id_dgmb', 'id', 'cpf', 'nome']) || '');
  seedBase += '|';
  seedBase += String(Math.floor(extrairNumeroPainel_(painel, ['realizado', 'realizado_km', 'total_km']) || 0));
  seedBase += '|';
  seedBase += String(new Date().getDate());

  var hash = simpleHashFrase_(seedBase);
  var idx = Math.abs(hash) % lista.length;
  return lista[idx];
}

function simpleHashFrase_(str) {
  var h = 0;
  str = String(str || '');
  for (var i = 0; i < str.length; i++) {
    h = ((h << 5) - h) + str.charCodeAt(i);
    h |= 0;
  }
  return h;
}

/** =========================
 *  HELPERS
 *  ========================= */

function extrairNumeroPainel_(obj, chaves) {
  if (!obj || !chaves || !chaves.length) return NaN;

  for (var i = 0; i < chaves.length; i++) {
    var k = chaves[i];
    if (obj.hasOwnProperty(k)) {
      var n = parseNumeroFrase_(obj[k]);
      if (isFinite(n)) return n;
    }
  }

  return NaN;
}

function extrairTextoPainel_(obj, chaves) {
  if (!obj || !chaves || !chaves.length) return '';
  for (var i = 0; i < chaves.length; i++) {
    var k = chaves[i];
    if (obj.hasOwnProperty(k) && obj[k] != null && obj[k] !== '') {
      return String(obj[k]);
    }
  }
  return '';
}

function parseNumeroFrase_(valor) {
  if (typeof valor === 'number') return valor;
  if (valor == null || valor === '') return NaN;

  var s = String(valor).trim();
  if (!s) return NaN;

  s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isFinite(n) ? n : NaN;
}

function normalizarTextoFrase_(valor) {
  return String(valor == null ? '' : valor)
    .toLowerCase()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function limparFrase_(valor) {
  var s = String(valor == null ? '' : valor).trim();
  return s;
}

function copiarMapaFrases_(origem) {
  var out = {};
  for (var k in origem) {
    if (!origem.hasOwnProperty(k)) continue;
    out[k] = (origem[k] || []).slice();
  }
  return out;
}


function listarFrasesDoContextoAtual_(painel) {
  try {
    var contexto = identificarContextoFrase_(painel);
    var frasesPorContexto = carregarFrasesMotivacionaisDaPlanilha_();
    var lista = frasesPorContexto[contexto];

    if (!lista || !lista.length) {
      lista = FRASES_FALLBACK_[contexto] || FRASES_FALLBACK_.GENERICA;
    }

    return {
      ok: true,
      contexto: contexto,
      frases: (lista || []).slice()
    };
  } catch (err) {
    return {
      ok: true,
      contexto: 'RITMO_IDEAL',
      frases: FRASES_FALLBACK_.GENERICA.slice()
    };
  }
}

function getOutraFraseMotivacional(idDgmb, fraseAtual) {
  try {
    var painelResp = getPainelUsuario(idDgmb);
    if (!painelResp || !painelResp.ok || !painelResp.data) {
      return {
        ok: false,
        code: 'PAINEL_INDISPONIVEL',
        msg: 'Não foi possível carregar os dados do atleta.'
      };
    }

    var d = painelResp.data || {};

    var painelBaseFrase = {
      id_dgmb: d.id_dgmb || '',
      nome: d.nome || '',
      meta: d.meta,
      realizado: d.realizado,
      restante: d.restante,
      percentual: d.percentual,
      ritmo_status: d.ritmo_status || '',
      ritmo_mensagem: d.ritmo_mensagem || ''
    };

    var pack = listarFrasesDoContextoAtual_(painelBaseFrase);
    var lista = (pack && pack.frases) ? pack.frases : [];

    if (!lista.length) {
      return {
        ok: true,
        data: {
          contexto_frase: pack.contexto || '',
          frase: 'Cada quilômetro conta. Continue no seu ritmo.'
        }
      };
    }

    var atual = String(fraseAtual || '').trim();
    var idx = lista.indexOf(atual);

    if (idx < 0) idx = 0;
    else idx = (idx + 1) % lista.length;

    return {
      ok: true,
      data: {
        contexto_frase: pack.contexto || '',
        frase: lista[idx]
      }
    };
  } catch (err) {
    return {
      ok: true,
      data: {
        contexto_frase: '',
        frase: 'Cada quilômetro conta. Continue no seu ritmo.'
      }
    };
  }
}


/**
 * Lista todas as frases disponíveis do contexto atual.
 * Nunca quebra.
 */
function listarFrasesDoContextoAtual_(painel) {
  try {
    var contexto = identificarContextoFrase_(painel);
    var frasesPorContexto = carregarFrasesMotivacionaisDaPlanilha_();
    var lista = frasesPorContexto[contexto];

    if (!lista || !lista.length) {
      lista = FRASES_FALLBACK_[contexto] || FRASES_FALLBACK_.GENERICA;
    }

    return {
      ok: true,
      contexto: contexto,
      frases: (lista || []).slice()
    };
  } catch (err) {
    return {
      ok: true,
      contexto: 'RITMO_IDEAL',
      frases: FRASES_FALLBACK_.GENERICA.slice()
    };
  }
}

/**
 * Retorna a próxima frase do mesmo contexto atual do atleta.
 * Seguro para uso no botão "Outra mensagem".
 */
function getOutraFraseMotivacional(idDgmb, fraseAtual) {
  try {
    var id = String(idDgmb || '').trim();
    if (!id) {
      return {
        ok: false,
        code: 'ID_OBRIGATORIO',
        msg: 'ID do atleta não informado.'
      };
    }

    var pessoa = buscarPessoaPainelMG_(id);
    var desafio = buscarInscricaoPainelMG_(id);

    if (!pessoa) {
      return {
        ok: false,
        code: 'USUARIO_NAO_ENCONTRADO',
        msg: 'Atleta não encontrado no desafio.'
      };
    }

    if (!desafio.ok) {
      return {
        ok: false,
        code: desafio.code,
        motivo_inscricao: desafio.motivo,
        msg: desafio.msg
      };
    }

    if (!desafio.data) {
      return {
        ok: false,
        code: 'CONTRATO_INSCRICAO_INVALIDO',
        msg: 'Dados de inscrição inválidos para frases motivacionais.'
      };
    }

    var desafioData = desafio.data;
    var meta = painelMG_toNumber_(desafioData.meta);
    var realizado = painelMG_toNumber_(desafioData.realizado);

    var progresso = painelMG_calcularProgresso_(meta, realizado);
    var ritmo = painelMG_calcularRitmo_(meta, realizado);

    var painelBaseFrase = {
      id_dgmb: pessoa.id_dgmb || id,
      nome: pessoa.nome || '',
      meta: painelMG_round1_(meta),
      realizado: painelMG_round1_(realizado),
      restante: progresso.restante,
      percentual: progresso.percentual,
      ritmo_status: ritmo.status,
      ritmo_mensagem: ritmo.mensagem
    };

    var pack = listarFrasesDoContextoAtual_(painelBaseFrase);
    var lista = (pack && pack.frases) ? pack.frases : [];

    if (!lista.length) {
      return {
        ok: true,
        data: {
          contexto_frase: pack && pack.contexto ? pack.contexto : '',
          frase: 'Cada quilômetro conta. Continue no seu ritmo.'
        }
      };
    }

    var atual = String(fraseAtual || '').trim();
    var idx = lista.indexOf(atual);

    if (idx < 0) {
      idx = 0;
    } else {
      idx = (idx + 1) % lista.length;
    }

    return {
      ok: true,
      data: {
        contexto_frase: pack && pack.contexto ? pack.contexto : '',
        frase: lista[idx]
      }
    };
  } catch (err) {
    return {
      ok: true,
      data: {
        contexto_frase: '',
        frase: 'Cada quilômetro conta. Continue no seu ritmo.'
      }
    };
  }
}