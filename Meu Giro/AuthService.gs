function loginCPF(cpf) {
  try {
    var cpfLimpo = onlyDigits_(cpf);

    if (!cpfLimpo) {
      return { ok: false, code: 'CPF_VAZIO', msg: 'Informe o CPF.' };
    }

    if (!isValidCPF_(cpfLimpo)) {
      return { ok: false, code: 'CPF_INVALIDO', msg: 'CPF inválido.' };
    }

    var usuario = buscarUsuarioPorCPF_(cpfLimpo);
    if (!usuario) {
      return { ok: false, code: 'CPF_NAO_ENCONTRADO', msg: 'CPF não encontrado.' };
    }

    var inscricao = verificarInscricaoDesafio_(usuario.id_dgmb);
    if (!inscricao) {
      return {
        ok: false,
        code: 'NAO_INSCRITO',
        msg: 'Seu CPF foi localizado, mas você não está inscrito no desafio atual.'
      };
    }

    return {
      ok: true,
      data: {
        id_dgmb: normalizeText_(usuario.id_dgmb),
        nome: normalizeText_(usuario.nome),
        cidade_uf: normalizeText_(usuario.cidade_uf),
        cpf: cpfLimpo
      }
    };
  } catch (err) {
    return {
      ok: false,
      code: 'LOGIN_ERROR',
      msg: err && err.message ? err.message : 'Erro ao realizar login.'
    };
  }
}

function buscarUsuarioPorCPF_(cpfLimpo) {
  var sh = getSheetByName_(SHEETS.PESSOAS);
  var values = sh.getDataRange().getValues();

  if (!values || values.length < 2) {
    return null;
  }

  var header = values[0];
  var map = buildHeaderMap_(header);

  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], SHEETS.PESSOAS);
  var idxCpf = getRequiredColumnIndex_(map, ['cpf'], SHEETS.PESSOAS);
  var idxNome = getRequiredColumnIndex_(map, ['nome'], SHEETS.PESSOAS);
  var idxCidade = getOptionalColumnIndex_(map, ['cidade-uf', 'cidade_uf', 'cidade uf']);

  var cpfCriptografado = Utilities.base64Encode(cpfLimpo);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var cpfSalvo = normalizeText_(row[idxCpf]);

    if (cpfSalvo === cpfCriptografado) {
      return {
        id_dgmb: normalizeText_(row[idxId]),
        nome: normalizeText_(row[idxNome]),
        cidade_uf: idxCidade > -1 ? normalizeText_(row[idxCidade]) : ''
      };
    }
  }

  return null;
}

function verificarInscricaoDesafio_(idDgmb) {
  var sh = getSheetByName_(SHEETS.DESAFIO);
  var values = sh.getDataRange().getValues();

  if (!values || values.length < 2) {
    return null;
  }

  var header = values[0];
  var map = buildHeaderMap_(header);

  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], SHEETS.DESAFIO);

  for (var i = 1; i < values.length; i++) {
    var rowId = normalizeText_(values[i][idxId]);
    if (rowId === normalizeText_(idDgmb)) {
      return { id_dgmb: rowId };
    }
  }

  return null;
}

function isValidCPF_(cpf) {
  var c = onlyDigits_(cpf);
  if (c.length !== 11 || /(\d)\1{10}/.test(c)) return false;

  function calcDigit(base, factor) {
    var total = 0;
    for (var i = 0; i < base.length; i++) {
      total += parseInt(base.charAt(i), 10) * (factor - i);
    }
    var rest = (total * 10) % 11;
    return rest === 10 ? 0 : rest;
  }

  var d1 = calcDigit(c.substring(0, 9), 10);
  var d2 = calcDigit(c.substring(0, 10), 11);

  return d1 === parseInt(c.charAt(9), 10) && d2 === parseInt(c.charAt(10), 10);
}