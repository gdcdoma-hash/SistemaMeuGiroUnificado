const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');

const index = fs.readFileSync('Meu Giro/Index.html', 'utf8');
const styles = fs.readFileSync('Meu Giro/Styles.html', 'utf8');
const script = fs.readFileSync('Meu Giro/Script.html', 'utf8');

test('card principal do desafio em foco usa donut e remove barra horizontal', () => {
  const painelInicio = index.slice(index.indexOf('<section id="screen-painel"'), index.indexOf('<div class="painel-acoes-rapidas"'));

  assert.match(painelInicio, /class="painel-meta-destaque"[\s\S]*id="meta-km"/);
  assert.match(painelInicio, /id="progress-donut"[\s\S]*id="progress-percent"[\s\S]*concluído/);
  assert.match(painelInicio, /class="painel-progresso-donut-grid"[\s\S]*id="realizado-km"[\s\S]*id="restante-km"/);
  assert.doesNotMatch(painelInicio, /painel-progress-track|id="progress-bar"|painel-progress-fill/);
});

test('donut usa conic-gradient e percentual normalizado do contexto', () => {
  assert.match(styles, /\.painel-progress-donut[\s\S]*conic-gradient\(#22c55e 0deg, #22c55e calc\(var\(--percent\) \* 3\.6deg\), #fde7c8 0deg\)/);
  assert.match(styles, /\.painel-progress-donut::after[\s\S]*background: #ffffff/);
  assert.match(script, /function atualizarDonutProgressoPainel_\(percentual\)[\s\S]*Math\.max\(0, Math\.min\(100, toFiniteNumber\(percentual, 0\)\)\)[\s\S]*style\.setProperty\('--percent', percentualNormalizado\)/);
  assert.match(script, /setTextById\('progress-percent', formatNumber\(contexto\.percentual\) \+ '%'\);\s*atualizarDonutProgressoPainel_\(contexto\.percentual\)/);
});

test('layout do donut é responsivo sem sobrepor números', () => {
  assert.match(styles, /\.painel-progresso-donut-grid[\s\S]*grid-template-columns: minmax\(148px, 0\.95fr\) minmax\(148px, 1fr\)/);
  assert.match(styles, /@media \(max-width: 420px\)[\s\S]*\.painel-progresso-donut-grid[\s\S]*grid-template-columns: 1fr/);
  assert.match(styles, /\.painel-progress-donut[\s\S]*width: min\(100%, 190px\)[\s\S]*aspect-ratio: 1/);
});
