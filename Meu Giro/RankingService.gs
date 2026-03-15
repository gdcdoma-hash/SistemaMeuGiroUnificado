function getRanking() {
  try {
    var pessoas = getAllObjects_(SHEETS.PESSOAS);
    var desafio = getAllObjects_(SHEETS.DESAFIO);

    if (!desafio || !desafio.length) {
      return {
        ok: true,
        data: [],
        msg: 'Nenhum atleta encontrado no desafio.'
      };
    }

    var pessoasMap = rankingMG_buildPessoasMap_(pessoas);
    var ranking = [];

    for (var i = 0; i < desafio.length; i++) {
      var row = desafio[i];

      var idDgmb = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']));
      if (!idDgmb) continue;

      var meta = rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row, [
        'Distancia_KM',
        'distancia_km',
        'Distancia KM'
      ])));

      var realizado = rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row, [
        'Distancia_Realizada',
        'distancia_realizada',
        'Distancia Realizada'
      ])));

      var restante = rankingMG_round1_(Math.max(meta - realizado, 0));
      var percentual = meta > 0 ? rankingMG_round1_((realizado / meta) * 100) : 0;

      var pessoa = pessoasMap[idDgmb] || {};

      if (realizado <= 0 || percentual <= 0) continue;

      ranking.push({
        id_dgmb: idDgmb,
        nome: pessoa.nome || rankingMG_norm_(rankingMG_firstFilled_(row, ['Nome_Avatar', 'nome_avatar'])) || 'Participante',
        cidade_uf: pessoa.cidade_uf || '',
        meta: meta,
        realizado: realizado,
        restante: restante,
        percentual: percentual
      });
    }

    ranking.sort(function(a, b) {
      if (b.percentual !== a.percentual) return b.percentual - a.percentual;
      if (a.restante !== b.restante) return a.restante - b.restante;
      if (b.realizado !== a.realizado) return b.realizado - a.realizado;
      return String(a.nome || '').localeCompare(String(b.nome || ''));
    });

    return {
      ok: true,
      data: ranking
    };
  } catch (err) {
    return {
      ok: false,
      msg: err && err.message ? err.message : 'Erro ao carregar ranking.'
    };
  }
}

function rankingMG_buildPessoasMap_(pessoas) {
  var map = {};

  for (var i = 0; i < pessoas.length; i++) {
    var row = pessoas[i];

    var idDgmb = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']));
    if (!idDgmb) continue;

    map[idDgmb] = {
      id_dgmb: idDgmb,
      nome: rankingMG_norm_(rankingMG_firstFilled_(row, ['nome', 'Nome'])),
      cidade_uf: rankingMG_norm_(rankingMG_firstFilled_(row, ['Cidade-UF', 'Cidade_UF', 'cidade_uf', 'cidade-uf']))
    };
  }

  return map;
}

function rankingMG_firstFilled_(obj, keys) {
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

function rankingMG_norm_(value) {
  return String(value || '').trim();
}

function rankingMG_toNumber_(value) {
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

function rankingMG_round1_(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 10) / 10;
}