const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');
const script = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Script.html'), 'utf8');

function getFunctionSlice(name, nextName) {
  const start = script.indexOf(`function ${name}`);
  const end = script.indexOf(`function ${nextName}`, start + 1);
  assert.notEqual(start, -1, `Função ${name} não encontrada`);
  assert.notEqual(end, -1, `Limite ${nextName} não encontrado`);
  return script.slice(start, end);
}

const helperCode =
  getFunctionSlice('getDesafioKey', 'salvarDesafioEmFocoNaSessao') + '\n' +
  getFunctionSlice('normalizarListaDesafiosPainel_', 'isDesafioAtivo_');

const ctx = {};
vm.createContext(ctx);
vm.runInContext(helperCode, ctx);

test('deduplicação preserva desafios atuais diferentes sem id_item_estoque usando meta alternativa', () => {
  const desafios = [
    {
      id_desafio: 'mensal-2026-07',
      periodo_inicio: '2026-07-01',
      periodo_fim: '2026-07-31',
      meta: 200,
      nome_desafio: 'Desafio Julho 200 km'
    },
    {
      id_desafio: 'mensal-2026-07',
      periodo_inicio: '2026-07-01',
      periodo_fim: '2026-07-31',
      meta: 400,
      nome_desafio: 'Desafio Julho 400 km'
    }
  ];

  assert.notEqual(ctx.getDesafioBaseKey_(desafios[0]), ctx.getDesafioBaseKey_(desafios[1]));
  const normalizados = ctx.normalizarListaDesafiosPainel_(desafios);
  assert.equal(normalizados.length, 2);
  assert.equal(normalizados[0].nome_desafio, 'Desafio Julho 200 km');
  assert.equal(normalizados[1].nome_desafio, 'Desafio Julho 400 km');
});

test('deduplicação usa nome como desempate quando IDs, período e meta vêm iguais ou incompletos', () => {
  const desafios = [
    {
      id_desafio: 'campanha-atual',
      periodo_inicio: '2026-07-01',
      periodo_fim: '2026-07-31',
      meta_km: 200,
      nome_desafio: 'Pedal da Cidade'
    },
    {
      id_desafio: 'campanha-atual',
      periodo_inicio: '2026-07-01',
      periodo_fim: '2026-07-31',
      meta_km: 200,
      nome_desafio: 'Pedal da Serra'
    }
  ];

  assert.notEqual(ctx.getDesafioBaseKey_(desafios[0]), ctx.getDesafioBaseKey_(desafios[1]));
  assert.equal(ctx.normalizarListaDesafiosPainel_(desafios).length, 2);
});

test('id_inscricao continua sendo a chave prioritária para representar a mesma inscrição', () => {
  const desafios = [
    { id_inscricao: 'abc-123', id_desafio: '1', meta_km: 200, nome_desafio: 'Original' },
    { id_inscricao: 'abc-123', id_desafio: '1', meta_km: 400, nome_desafio: 'Duplicado enriquecido' }
  ];

  assert.equal(ctx.getDesafioBaseKey_(desafios[0]), 'INSCRICAO|abc-123');
  assert.equal(ctx.normalizarListaDesafiosPainel_(desafios).length, 1);
});
