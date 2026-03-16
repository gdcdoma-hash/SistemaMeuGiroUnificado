function getPainelUsuario(idDgmb) {
  try {
    var id = String(idDgmb || '').trim();
    if (!id) {
      return { ok: false, msg: 'ID do usuário não informado.' };
    }

    var pessoa = buscarPessoaPainelMG_(id);
    var desafio = buscarInscricaoPainelMG_(id);

    if (!pessoa || !desafio) {
      return { ok: false, msg: 'Usuário não encontrado no desafio.' };
    }

    var meta = painelMG_toNumber_(desafio.meta);
    var realizado = painelMG_toNumber_(desafio.realizado);

    var progresso = painelMG_calcularProgresso_(meta, realizado);
    var ritmo = painelMG_calcularRitmo_(meta, realizado);
    var atividades = buscarAtividadesUsuario_(id);
    var rankingInfo = painelMG_calcularPosicaoRanking_(id);

    var frase = '';
    var contextoFrase = '';

    try {
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

      var payloadFrase = montarPayloadFraseMotivacional_(painelBaseFrase);
      frase = payloadFrase && payloadFrase.frase_motivacional ? payloadFrase.frase_motivacional : '';
      contextoFrase = payloadFrase && payloadFrase.contexto_frase ? payloadFrase.contexto_frase : '';
    } catch (e) {
      frase = '';
      contextoFrase = '';
    }

    return {
      ok: true,
      data: {
        nome: pessoa.nome || '',
        cidade_uf: pessoa.cidade_uf || '',
        id_dgmb: pessoa.id_dgmb || '',
        status_inscricao: desafio.status_inscricao || 'inscrito',
        criterio_validacao_inscricao: desafio.criterio_validacao || 'presenca_id_dgmb',
        desafio_usuario: desafio.aba_desafio || '',

        meta: painelMG_round1_(meta),
        realizado: painelMG_round1_(realizado),
        restante: progresso.restante,
        percentual: progresso.percentual,

        diaAtual: ritmo.diaAtual,
        diasRestantes: ritmo.diasRestantes,
        kmIdealAtual: ritmo.kmIdealHoje,
        kmPorDiaRestante: ritmo.kmPorDiaRestante,

        ritmo_status: ritmo.status,
        ritmo_mensagem: ritmo.mensagem,

        posicao_ranking: rankingInfo.posicao,
        total_participantes: rankingInfo.total,

        frase: frase || desafio.frase_incentivo || 'Cada quilômetro conta. Continue no seu ritmo.',
        frase_motivacional: frase || desafio.frase_incentivo || 'Cada quilômetro conta. Continue no seu ritmo.',
        contexto_frase: contextoFrase || '',

        atividades: atividades,
        totalPedalado: painelMG_round1_(realizado)
      }
    };
  } catch (err) {
    return {
      ok: false,
      msg: err && err.message ? err.message : 'Erro ao carregar painel.'
    };
  }
}
function buscarAtividadesUsuario_(idDgmb) {
  var items = getAllObjects_(SHEETS.REGISTRO_KM);
  var out = [];

  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var rowId = painelMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']);

    if (painelMG_norm_(rowId) === painelMG_norm_(idDgmb)) {
      var dataOriginal = painelMG_firstFilled_(row, ['Data_Atividade', 'Data', 'data_atividade', 'data']);
      var dataNormalizada = painelMG_normalizarDataAtividade_(dataOriginal);

      var timestampOriginal = painelMG_firstFilled_(row, ['Timestamp', 'timestamp']);
      var chaveEdicao = normalizarTimestampEdicao_(timestampOriginal);

      out.push({
        chave_edicao: String(chaveEdicao || '').trim(),
        data: dataNormalizada,
        km: painelMG_round1_(painelMG_toNumber_(painelMG_firstFilled_(row, ['KM', 'km'])))
      });
    }
  }

  out.sort(function(a, b) {
    return String(b.data || '').localeCompare(String(a.data || ''));
  });

  return out;
}

function buscarPessoaPainelMG_(idDgmb) {
  var items = getAllObjects_(SHEETS.PESSOAS);

  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var rowId = painelMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']);

    if (painelMG_norm_(rowId) === painelMG_norm_(idDgmb)) {
      return {
        id_dgmb: painelMG_norm_(rowId),
        nome: painelMG_norm_(painelMG_firstFilled_(row, ['nome', 'Nome'])),
        cidade_uf: painelMG_norm_(painelMG_firstFilled_(row, ['Cidade-UF', 'Cidade_UF', 'cidade_uf', 'cidade-uf']))
      };
    }
  }

  return null;
}

function buscarInscricaoPainelMG_(idDgmb) {
  var inscricao = obterDadosInscricaoUsuario_(idDgmb);

  if (!inscricao || inscricao.inscricao_valida === false) {
    return null;
  }

  return {
    id_dgmb: painelMG_norm_(inscricao.id_dgmb),
    meta: inscricao.meta,
    realizado: inscricao.distancia_realizada,
    status_inscricao: painelMG_norm_(inscricao.status_inscricao),
    criterio_validacao: painelMG_norm_(inscricao.criterio_validacao),
    aba_desafio: painelMG_norm_(inscricao.aba_desafio),
    frase_incentivo: painelMG_norm_(inscricao.frase_incentivo)
  };
}

function painelMG_calcularProgresso_(meta, realizado) {
  var restante = Math.max(painelMG_toNumber_(meta) - painelMG_toNumber_(realizado), 0);
  var percentual = painelMG_toNumber_(meta) > 0 ? (painelMG_toNumber_(realizado) / painelMG_toNumber_(meta)) * 100 : 0;

  return {
    restante: painelMG_round1_(restante),
    percentual: painelMG_round1_(painelMG_clamp_(percentual, 0, 100))
  };
}

function painelMG_calcularRitmo_(meta, realizado) {
  var now = new Date();
  var diaAtual = now.getDate();
  var diasTotal = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();

  if (diaAtual < 1) diaAtual = 1;
  if (diaAtual > diasTotal) diaAtual = diasTotal;

  var diasRestantes = Math.max(diasTotal - diaAtual, 0);
  var kmIdealPorDia = painelMG_toNumber_(meta) > 0 ? painelMG_toNumber_(meta) / diasTotal : 0;
  var kmIdealHoje = kmIdealPorDia * diaAtual;
  var restante = Math.max(painelMG_toNumber_(meta) - painelMG_toNumber_(realizado), 0);
  var kmPorDiaRestante = diasRestantes > 0 ? restante / diasRestantes : restante;

  var diferenca = painelMG_toNumber_(realizado) - kmIdealHoje;
  var tolerancia = Math.max(2, painelMG_toNumber_(meta) * 0.01);

  var status = 'Você está à frente da meta.'
  var mensagem = 'Seu progresso está acima do ritmo esperado para este momento do desafio.'

  if (diferenca > tolerancia) {
    status = 'Você está à frente da meta.';
      mensagem = 'Seu progresso está acima do ritmo esperado para este momento do desafio.';

  } else if (diferenca < -tolerancia) {
    status = 'Você está um pouco abaixo do ritmo.'
    mensagem =  'Alguns pedais extras podem ajudar a recuperar o ritmo do desafio.'
  }

  return {
    diaAtual: diaAtual,
    diasRestantes: diasRestantes,
    kmIdealHoje: painelMG_round1_(kmIdealHoje),
    kmPorDiaRestante: painelMG_round1_(kmPorDiaRestante),
    status: status,
    mensagem: mensagem
  };
}

function painelMG_calcularPosicaoRanking_(idDgmb) {
  var pessoas = getAllObjects_(SHEETS.PESSOAS);
  var abaDesafio = localizarAbaDesafioUsuario_(idDgmb).abaDesafio;
  var desafio = getAllObjects_(abaDesafio);
  var pessoasMap = {};

  for (var i = 0; i < pessoas.length; i++) {
    var p = pessoas[i];
    var pid = painelMG_norm_(painelMG_firstFilled_(p, ['ID_DGMB', 'id_dgmb']));
    if (!pid) continue;

    pessoasMap[pid] = {
      nome: painelMG_norm_(painelMG_firstFilled_(p, ['nome', 'Nome'])),
      cidade_uf: painelMG_norm_(painelMG_firstFilled_(p, ['Cidade-UF', 'Cidade_UF', 'cidade_uf', 'cidade-uf']))
    };
  }

  var lista = [];

  for (var j = 0; j < desafio.length; j++) {
    var d = desafio[j];
    var did = painelMG_norm_(painelMG_firstFilled_(d, ['ID_DGMB', 'id_dgmb']));
    if (!did) continue;

    var meta = painelMG_toNumber_(painelMG_firstFilled_(d, ['Distancia_KM', 'distancia_km', 'Distancia KM']));
    var realizado = painelMG_toNumber_(painelMG_firstFilled_(d, ['Distancia_Realizada', 'distancia_realizada', 'Distancia Realizada']));
    var restante = Math.max(meta - realizado, 0);
    var percentual = meta > 0 ? (realizado / meta) * 100 : 0;

    lista.push({
      id_dgmb: did,
      nome: (pessoasMap[did] && pessoasMap[did].nome) || painelMG_norm_(painelMG_firstFilled_(d, ['Nome_Avatar', 'nome_avatar'])) || 'Participante',
      percentual: painelMG_round1_(percentual),
      restante: painelMG_round1_(restante),
      realizado: painelMG_round1_(realizado)
    });
  }

  lista.sort(function(a, b) {
    if (b.percentual !== a.percentual) return b.percentual - a.percentual;
    if (a.restante !== b.restante) return a.restante - b.restante;
    if (b.realizado !== a.realizado) return b.realizado - a.realizado;
    return String(a.nome || '').localeCompare(String(b.nome || ''));
  });

  for (var k = 0; k < lista.length; k++) {
    if (lista[k].id_dgmb === painelMG_norm_(idDgmb)) {
      return {
        posicao: k + 1,
        total: lista.length
      };
    }
  }

  return {
    posicao: 0,
    total: lista.length
  };
}

function painelMG_firstFilled_(obj, keys) {
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

function painelMG_normalizarDataAtividade_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var s = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    var p = s.split('/');
    return p[2] + '-' + p[1] + '-' + p[0];
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return '';
}

function painelMG_norm_(value) {
  return String(value || '').trim();
}

function painelMG_toNumber_(value) {
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }

  var s = String(value == null ? '' : value).trim();
  if (!s) return 0;

  s = s.replace(/\s+/g, '');

  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  }

  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function painelMG_clamp_(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function painelMG_round1_(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 10) / 10;
}

function painelMG_obterFraseSeguro_(frasePadrao) {
  var frase = String(frasePadrao || '').trim();
  if (frase) return frase;
  return 'Cada quilômetro conta. Continue no seu ritmo.';
}


