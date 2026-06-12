/**
 * Auditoria temporária, estritamente somente leitura, dos valores existentes em
 * dgmbDesafios.Status_Usuario_Desafio.
 *
 * @return {Object} Relatório com contagens, exemplos, variações e recomendações.
 */
function auditarStatusUsuarioDesafio() {
  var LIMITE_EXEMPLOS = 5;
  var nomeAba = (typeof SHEETS !== 'undefined' && SHEETS.DESAFIO)
    ? SHEETS.DESAFIO
    : 'dgmbDesafios';
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(nomeAba);

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + nomeAba);
  }

  var dados = sheet.getDataRange().getDisplayValues();
  if (!dados || dados.length === 0) {
    throw new Error('A aba ' + nomeAba + ' não possui cabeçalho.');
  }

  function semAcentos_(valor) {
    return String(valor).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function normalizarCabecalho_(valor) {
    return semAcentos_(valor)
      .toLowerCase()
      .replace(/_/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizarStatus_(valor) {
    var texto = String(valor)
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
    return texto ? semAcentos_(texto) : 'VAZIO';
  }

  function localizarColuna_(cabecalhos, candidatos, obrigatoria) {
    var candidatosNormalizados = {};
    for (var c = 0; c < candidatos.length; c++) {
      candidatosNormalizados[normalizarCabecalho_(candidatos[c])] = true;
    }

    for (var h = 0; h < cabecalhos.length; h++) {
      if (candidatosNormalizados[normalizarCabecalho_(cabecalhos[h])]) {
        return h;
      }
    }

    if (obrigatoria) {
      throw new Error(
        'Coluna obrigatória não encontrada na aba ' + nomeAba + ': ' + candidatos.join(' / ')
      );
    }
    return -1;
  }

  function valorDaColuna_(linha, indice) {
    return indice > -1 && indice < linha.length ? String(linha[indice]) : '';
  }

  function adicionarExemplo_(lista, valor, incluirVazio) {
    var repetido = lista.indexOf(valor) !== -1;
    if ((!incluirVazio && valor === '') || lista.length >= LIMITE_EXEMPLOS || repetido) return;
    lista.push(valor);
  }

  function rotuloOriginal_(valor) {
    return valor === '' ? '(vazio)' : JSON.stringify(valor);
  }

  var cabecalhos = dados[0];
  var idxStatus = localizarColuna_(cabecalhos, [
    'Status_Usuario_Desafio',
    'status_usuario_desafio',
    'status usuário desafio',
    'status usuario desafio'
  ], true);
  var idxInscricao = localizarColuna_(cabecalhos, [
    'ID_INSCRICAO',
    'id inscrição',
    'id inscricao'
  ], false);
  var idxDgmb = localizarColuna_(cabecalhos, ['ID_DGMB', 'id dgmb'], false);
  var idxDesafio = localizarColuna_(cabecalhos, ['ID_DESAFIO', 'id desafio'], false);

  var porOriginal = {};
  var ordemOriginais = [];
  var porNormalizado = {};
  var ordemNormalizados = [];
  var totalVazias = 0;

  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    var numeroLinha = i + 1;
    var original = valorDaColuna_(linha, idxStatus);
    var normalizado = normalizarStatus_(original);
    var chaveOriginal = '$' + original;
    var chaveNormalizada = '$' + normalizado;

    if (original === '') totalVazias++;

    if (!Object.prototype.hasOwnProperty.call(porOriginal, chaveOriginal)) {
      porOriginal[chaveOriginal] = {
        valor_original: original,
        valor_normalizado: normalizado,
        quantidade: 0,
        exemplos_linhas: [],
        exemplos_id_inscricao: [],
        exemplos_id_dgmb: [],
        exemplos_id_desafio: []
      };
      ordemOriginais.push(chaveOriginal);
    }

    var itemOriginal = porOriginal[chaveOriginal];
    itemOriginal.quantidade++;
    adicionarExemplo_(itemOriginal.exemplos_linhas, numeroLinha, true);
    adicionarExemplo_(itemOriginal.exemplos_id_inscricao, valorDaColuna_(linha, idxInscricao));
    adicionarExemplo_(itemOriginal.exemplos_id_dgmb, valorDaColuna_(linha, idxDgmb));
    adicionarExemplo_(itemOriginal.exemplos_id_desafio, valorDaColuna_(linha, idxDesafio));

    if (!Object.prototype.hasOwnProperty.call(porNormalizado, chaveNormalizada)) {
      porNormalizado[chaveNormalizada] = {
        valor_normalizado: normalizado,
        quantidade: 0,
        valores_originais: []
      };
      ordemNormalizados.push(chaveNormalizada);
    }

    var itemNormalizado = porNormalizado[chaveNormalizada];
    itemNormalizado.quantidade++;
    if (itemNormalizado.valores_originais.indexOf(original) === -1) {
      itemNormalizado.valores_originais.push(original);
    }
  }

  var valoresOriginais = ordemOriginais.map(function(chave) {
    return porOriginal[chave];
  });
  var valoresNormalizados = ordemNormalizados.map(function(chave) {
    return porNormalizado[chave];
  });
  var inconsistencias = valoresNormalizados.filter(function(item) {
    return item.valores_originais.length > 1;
  }).map(function(item) {
    return {
      valor_normalizado: item.valor_normalizado,
      variacoes: item.valores_originais.slice()
    };
  });

  valoresOriginais.sort(function(a, b) {
    return b.quantidade - a.quantidade || a.valor_original.localeCompare(b.valor_original);
  });
  valoresNormalizados.sort(function(a, b) {
    return b.quantidade - a.quantidade || a.valor_normalizado.localeCompare(b.valor_normalizado);
  });
  inconsistencias.sort(function(a, b) {
    return a.valor_normalizado.localeCompare(b.valor_normalizado);
  });

  var recomendacoes = [
    'Definir o enum oficial somente após validação administrativa dos valores normalizados encontrados.',
    'Criar uma tabela explícita de equivalência entre cada grafia original e o futuro valor oficial.',
    'Tratar vazios e valores compostos apenas por espaços como casos distintos durante o saneamento.',
    'Executar qualquer correção de dados em uma missão separada, com backup e validação prévia.'
  ];

  var relatorio = {
    aba: nomeAba,
    cabecalho_status_encontrado: cabecalhos[idxStatus],
    total_linhas_analisadas: Math.max(dados.length - 1, 0),
    total_celulas_vazias: totalVazias,
    colunas_de_exemplo_encontradas: {
      ID_INSCRICAO: idxInscricao > -1,
      ID_DGMB: idxDgmb > -1,
      ID_DESAFIO: idxDesafio > -1
    },
    valores_originais: valoresOriginais,
    valores_normalizados: valoresNormalizados,
    possiveis_inconsistencias_grafia: inconsistencias,
    recomendacoes_saneamento: recomendacoes
  };

  Logger.log('=== AUDITORIA Status_Usuario_Desafio ===');
  Logger.log('Aba: ' + nomeAba);
  Logger.log('Cabeçalho encontrado: ' + cabecalhos[idxStatus]);
  Logger.log('Total de linhas analisadas: ' + relatorio.total_linhas_analisadas);
  Logger.log('Total de células vazias: ' + totalVazias);
  Logger.log('--- Valores originais ---');
  valoresOriginais.forEach(function(item) {
    Logger.log(rotuloOriginal_(item.valor_original) + ' = ' + item.quantidade);
  });
  Logger.log('--- Valores normalizados ---');
  valoresNormalizados.forEach(function(item) {
    Logger.log(item.valor_normalizado + ' = ' + item.quantidade);
  });
  Logger.log('--- Exemplos por valor original ---');
  valoresOriginais.forEach(function(item) {
    Logger.log(
      rotuloOriginal_(item.valor_original) +
      ' | normalizado=' + item.valor_normalizado +
      ' | linhas=' + JSON.stringify(item.exemplos_linhas) +
      ' | ID_INSCRICAO=' + JSON.stringify(item.exemplos_id_inscricao) +
      ' | ID_DGMB=' + JSON.stringify(item.exemplos_id_dgmb) +
      ' | ID_DESAFIO=' + JSON.stringify(item.exemplos_id_desafio)
    );
  });
  Logger.log('--- Possíveis inconsistências de grafia ---');
  if (inconsistencias.length === 0) {
    Logger.log('Nenhuma variação de grafia agrupada pela normalização foi encontrada.');
  } else {
    inconsistencias.forEach(function(item) {
      Logger.log(
        'Valor normalizado ' + item.valor_normalizado +
        ' possui variações: ' + item.variacoes.map(rotuloOriginal_).join(', ')
      );
    });
  }
  Logger.log('--- Recomendações de saneamento (nenhuma alteração realizada) ---');
  recomendacoes.forEach(function(recomendacao) {
    Logger.log('- ' + recomendacao);
  });
  Logger.log('=== FIM DA AUDITORIA ===');

  return relatorio;
}
