'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const raiz = path.resolve(__dirname, '..');
const src = nome => fs.readFileSync(path.join(raiz, 'Meu Giro', nome), 'utf8');
const periodo = (dia, inicio, fim) => dia >= inicio && dia <= fim;
const helpers = new Function('atividadeDentroPeriodoOficial_', src('ModalidadeEsportivaService.gs') +
  '\nreturn {canon:meuGiroModalidadeCanonica_, sum:meuGiroSomarKmPorVinculo_, registro:meuGiroModalidadeRegistro_}')(periodo);

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
  assert.match(registro, /ensureRegistroKmModalidadeColumn_/);
  assert.match(registro, /checkModalidade === modalidadeEdicao/);
  assert.match(registro, /modalidadeLinha === modalidadeRegistro/);
  assert.match(registro, /meuGiroSomarKmPorVinculo_\(registros, vinculo\)/);
  assert.match(painel, /modalidade_esportiva: meuGiroModalidadeRegistro_\(row\)/);
  assert.match(ui, /registrarAtividade\(idDgmb, data, km, !!forcarDuplicidade, meuGiroModalidadeFormularioAtividade_\(\)\)/);
  assert.match(ui, /modos\.CAMINHADA_CORRIDA/);
  assert.match(formulario, /id="modalidade-atividade"/);
});
