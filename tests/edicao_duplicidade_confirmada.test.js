'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '..', 'Meu Giro', 'RegistroService.gs'), 'utf8');
const helperStart = source.indexOf('function meuGiroPodeConfirmarDuplicidadeEdicao_(');
const end = source.indexOf('\nfunction parseKmInputSeguro_(', helperStart);
assert.ok(helperStart >= 0 && end > helperStart, 'trecho do backend de edição encontrado');
const code = source.slice(helperStart, end);

function simularEdicao(options = {}) {
  // Duas atividades reais e independentes no LAB 1603 em 01/10: 4 km + 3 km.
  const linhas = [
    ['Timestamp','ID_DGMB','Data_Atividade','KM','activity_id','Modalidade_Esportiva'],
    ['T-A','1603','2026-10-01',4,'id-distinto-A','CICLISMO'],
    ['T-B','1603','2026-10-01',3,'id-distinto-B','CICLISMO']
  ];
  let escritas = 0, recalculos = 0;
  const planilha = {
    getDataRange: () => ({getValues: () => linhas.map(row => row.slice())}),
    getRange: (row, col) => ({setValue: (value) => {
      linhas[row-1][col-1] = value;
      escritas++;
    }})
  };
  const colunas = {idxTimestamp:0,idxId:1,idxData:2,idxKm:3,idxActivityId:4,idxModalidade:5};
  const localizar = (dados, cols, id, activityId, chave) => {
    const porId = activityId ? dados
      .map((r,i) => r[cols.idxId] === id && r[cols.idxActivityId] === activityId ? i+1 : -1)
      .filter(row => row > 1) : [];
    if (porId.length) return porId;
    return chave ? dados
      .map((r,i) => r[cols.idxId] === id && r[cols.idxTimestamp] === chave ? i+1 : -1)
      .filter(row => row > 1) : [];
  };
  const deps = {
    MEU_GIRO_PERF_OPERACAO_ATUAL_:'',
    meuGiroPerfNow_: () => 0,
    meuGiroPerfLog_: () => {},
    // A aplicação PROD consulta a trava de manutenção antes de editar os dados.
    // Aqui simulamos explicitamente o estado LIBERADO; nenhum GAS real é chamado.
    dgmbCorteGravacaoPausada_: () => false,
    LockService:{getScriptLock: () => ({waitLock(){},releaseLock(){}})},
    Logger:{log(){}},
    normalizarDataISO_: x => String(x || ''),
    normalizarDataEdicao_: x => String(x || ''),
    parseKmInputSeguro_: x => Number(x),
    normalizarKmEdicao_: x => Number(x),
    kmsIguaisEdicao_: (a,b) => Number(a) === Number(b),
    meuGiroModalidadeCanonica_: x => String(x || 'CICLISMO'),
    meuGiroValidarDataModalidadeInscricao_: () => ({ok:true}),
    getSpreadsheet_: () => ({getSheetByName: () => planilha}),
    SHEETS:{REGISTRO_KM:'REGISTRO_KM'},
    getRegistroKmColumnIndexes_: () => colunas,
    ensureRegistroKmModalidadeColumn_: (s,d,c) => c,
    localizarLinhasAtividade_: localizar,
    criarOpcoesRegistroKmReaproveitado_: () => ({}),
    atualizarDistanciaRealizada_: () => {},
    atualizarMeuGiroResumoComLockAdquirido_: () => {recalculos++;}
  };
  const editar = new Function(...Object.keys(deps), code + '\nreturn editarAtividade;')(...Object.values(deps));
  const payload = {
    id_dgmb:'1603', activity_id:'id-distinto-A', chave_edicao:'T-A',
    data_atividade:'2026-10-01', km:3, modalidade_esportiva:'CICLISMO',
    force:false, ...options
  };
  return {res:editar(payload), linhas, escritas, recalculos};
}

test('sem confirmar, a edição 4→3 é bloqueada sem alterar a planilha', () => {
  const sim = simularEdicao();
  assert.equal(sim.res.code, 'DUPLICIDADE_EDICAO');
  assert.equal(sim.linhas[1][3],4);
  assert.equal(sim.linhas[2][3],3);
  assert.equal(sim.escritas,0);
  assert.equal(sim.recalculos,0);
});

test('com confirmação explícita, edição 4→3 altera somente a atividade identificada', () => {
  const sim = simularEdicao({force:true});
  assert.equal(sim.res.ok,true);
  assert.equal(sim.res.registros_atualizados,1);
  assert.equal(sim.linhas.length,3); // cabeçalho + duas atividades distintas
  assert.equal(sim.linhas[1][3],3);
  assert.equal(sim.linhas[2][3],3);
  assert.equal(sim.linhas[1][4],'id-distinto-A');
  assert.equal(sim.linhas[2][4],'id-distinto-B');
  assert.equal(sim.recalculos,1);
});

test('não aceitar confirmação forçada quando o activity_id está incorreto', () => {
  const sim = simularEdicao({activity_id:'id-incorreto',force:true});
  assert.equal(sim.res.code,'DUPLICIDADE_EDICAO');
  assert.equal(sim.escritas,0);
});

test('registros legados sem activity_id não recebem liberação indiscriminada', () => {
  const sim = simularEdicao({activity_id:'',force:true});
  assert.equal(sim.res.code,'DUPLICIDADE_EDICAO');
  assert.equal(sim.escritas,0);
});

test('force em texto não substitui a confirmação booleana explícita', () => {
  const sim = simularEdicao({force:'true'});
  assert.equal(sim.res.code,'DUPLICIDADE_EDICAO');
  assert.equal(sim.escritas,0);
});
