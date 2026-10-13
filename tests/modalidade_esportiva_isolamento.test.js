'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const raiz = path.resolve(__dirname, '..');
const src = nome => fs.readFileSync(path.join(raiz, 'Meu Giro', nome), 'utf8');
const periodo = (dia, inicio, fim) => dia >= inicio && dia <= fim;
const dataValida = v => /^\d{4}-\d{2}-\d{2}$/.test(v) &&
  !Number.isNaN(new Date(v+'T12:00:00Z').getTime()) &&
  new Date(v+'T12:00:00Z').toISOString().slice(0,10)===v;
const meuGiroSimulado = hoje => new Function(
  'atividadeDentroPeriodoOficial_', 'normalizarDataISO_', 'isDataIsoValida_', 'Utilities',
  src('ModalidadeEsportivaService.gs') +
  '\nreturn {canon:meuGiroModalidadeCanonica_, sum:meuGiroSomarKmPorVinculo_,' +
  'registro:meuGiroModalidadeRegistro_, validar:meuGiroValidarDataModalidadeInscricao_};'
)(periodo, v=>String(v||''),dataValida,{formatDate:()=>hoje});
const helpers = meuGiroSimulado('2026-11-15');

test('atividade legada sem modalidade permanece Ciclismo', () => {
  assert.equal(helpers.registro({}), 'CICLISMO');
  assert.equal(helpers.canon('Ride'), 'CICLISMO');
  assert.equal(helpers.canon('Run'), 'CAMINHADA_CORRIDA');
  assert.equal(helpers.canon('Walk'), 'CAMINHADA_CORRIDA');
  assert.throws(() => helpers.canon('Surf'), /Modalidade esportiva não reconhecida/);
});

test('KM não atravessam modalidade ou período; atividade vale para desafios compatíveis', () => {
  const registros = [
    { data_atividade: '2026-10-04', km: 20, modalidade_esportiva: 'CICLISMO' },
    { data_atividade: '2026-10-04', km: 5, modalidade_esportiva: 'CAMINHADA_CORRIDA' },
    { data_atividade: '2026-10-07', km: 12 }, // LEGADO = ciclismo
    { data_atividade: '2026-09-30', km: 3, modalidade_esportiva: 'CAMINHADA_CORRIDA' },
    { data_atividade: '2026-11-01', km: 13, modalidade_esportiva: 'CICLISMO' }
  ];
  const periodoOutubro = { periodo_inicio: '2026-10-01', periodo_fim: '2026-10-31' };
  assert.equal(helpers.sum(registros, {...periodoOutubro, modalidade_esportiva:'CICLISMO'}), 32);
  assert.equal(helpers.sum(registros, {...periodoOutubro, modalidade_esportiva:'CAMINHADA_CORRIDA'}), 5);
  const outroDesafioCiclismoMesmoPrazo = {...periodoOutubro, modalidade_esportiva:'CICLISMO'};
  assert.equal(helpers.sum(registros, outroDesafioCiclismoMesmoPrazo), 32);
  assert.equal(helpers.sum(registros, {...periodoOutubro, periodo_inicio:'2026-11-01',
    periodo_fim:'2026-11-30', modalidade_esportiva:'CICLISMO'}), 13);
});

test('regra de pontuação, gravação e histórico estão conectados às modalidades', () => {
  const utils = src('Utils.gs');
  const registro = src('RegistroService.gs');
  const painel = src('PainelService.gs');
  const ui = src('Script.html');
  const formulario = src('Index.html');
  assert.match(utils, /meuGiroAtividadeNoVinculo_\(reg, vinculo\)/);
  assert.match(utils, /modalidade_esportiva: modalidadeVinculo/);
  assert.match(utils, /modalidade_esportiva: meuGiroModalidadeRegistro_\(r\)/);
  // Tanto o caminho de leitura rápida quanto a reconstrução completa do
  // painel devem recuperar a modalidade da inscrição.
  assert.match(utils, /statusDgmbResumo\.modalidade_esportiva/);
  assert.match(utils, /modalidade_esportiva: meuGiroModalidadeCanonica_\(vinculo\.modalidade_esportiva\)/);
  assert.match(utils, /idxModalidadeInscricao > -1 && normalizeText_\(row\[idxModalidadeInscricao\]\)/);

  assert.match(registro, /ensureRegistroKmModalidadeColumn_/);
  assert.match(registro, /checkModalidade === modalidadeEdicao/);
  assert.match(registro, /modalidadeLinha === modalidadeRegistro/);
  assert.match(registro, /meuGiroSomarKmPorVinculo_\(registros, vinculo\)/);
  assert.match(registro, /MODALIDADES_MISTURADAS/);
  assert.match(painel, /modalidade_esportiva: meuGiroModalidadeRegistro_\(row\)/);
  assert.match(ui, /registrarAtividade\(idDgmb, data, km, !!forcarDuplicidade, meuGiroModalidadeFormularioAtividade_\(\)\)/);
  assert.match(ui, /modos\.CAMINHADA_CORRIDA/);
  assert.match(ui, /var chaveFoco = String/);
  assert.match(formulario, /id="modalidade-atividade"/);
});

test('histórico retroativo livre e proibição de data futura em ambos os esportes', () => {
  const fixed = meuGiroSimulado('2026-10-13');
  const vinculos=[
    {apto:true,modalidade_esportiva:'CICLISMO',periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31'},
    {apto:true,modalidade_esportiva:'CAMINHADA_CORRIDA',periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31'}
  ];
  assert.equal(fixed.validar('2026-09-30','CICLISMO',vinculos).ok,true);
  assert.equal(fixed.validar('2026-09-30','CAMINHADA_CORRIDA',[]).ok,true);
  assert.equal(fixed.validar('2026-10-13','CICLISMO',vinculos).ok,true);
  assert.equal(fixed.validar('2026-10-13','CAMINHADA_CORRIDA',vinculos).ok,true);
  assert.equal(fixed.validar('2026-10-14','CICLISMO',vinculos).code,'DATA_FUTURA');
  assert.equal(fixed.validar('2026-10-14','CAMINHADA_CORRIDA',vinculos).code,'DATA_FUTURA');
  assert.equal(fixed.validar('2026-11-01','CICLISMO',vinculos).code,'DATA_FUTURA');
  assert.equal(fixed.validar('2026-02-30','CICLISMO',vinculos).code,'DATA_INVALIDA');
  assert.equal(fixed.sum([
    {data_atividade:'2026-09-30',km:4,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-13',km:5,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-14',km:1,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-13',km:3,modalidade_esportiva:'CAMINHADA_CORRIDA'}
  ],vinculos[0]),5);
  assert.equal(fixed.sum([
    {data_atividade:'2026-10-13',km:3,modalidade_esportiva:'CAMINHADA_CORRIDA'}
  ],vinculos[1]),3);
});

test('histórico de desafio separa atividades por modalidade além das datas', () => {
  const script=src('Script.html');
  const a=script.indexOf('function filtrarAtividadesPorPeriodoDesafio(');
  const b=script.indexOf('\nfunction getDesafioPainelPorChave',a);
  assert.ok(a>0&&b>a);
  const filtrar=new Function(script.slice(a,b)+'\nreturn filtrarAtividadesPorPeriodoDesafio;')();
  const atividades=[
    {data:'2026-09-30',km:3,modalidade_esportiva:'CAMINHADA_CORRIDA'},
    {data:'2026-10-12',km:20},
    {data:'2026-10-12',km:5.5,modalidade_esportiva:'CAMINHADA_CORRIDA'},
    {data:'2026-11-01',km:7,modalidade_esportiva:'CAMINHADA_CORRIDA'}
  ];
  assert.equal(filtrar(atividades,{periodo_inicio:'2026-09-20',periodo_fim:'2026-11-03',
    modalidade_esportiva:'CAMINHADA_CORRIDA'}).items.length,3);
  assert.equal(filtrar(atividades,{periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31',
    modalidade_esportiva:'CICLISMO'}).items.length,1);
  const backend=src('RegistroService.gs');
  assert.ok(backend.includes('var dataPermitida = meuGiroValidarDataModalidadeInscricao_'));
  assert.ok(backend.includes('var dataPermitidaEdicao = meuGiroValidarDataModalidadeInscricao_'));
  assert.ok(!backend.includes('modalidadeRegistro, obterVinculosDesafioUsuario_(idDgmb)'));
  assert.ok(script.includes('meuGiroAvisoDataFutura_'));
  assert.ok(script.includes('await mostrarErroAcaoAtividade(res)'));
});
