const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');

const index = fs.readFileSync('Meu Giro/Index.html', 'utf8');
const styles = fs.readFileSync('Meu Giro/Styles.html', 'utf8');
const script = fs.readFileSync('Meu Giro/Script.html', 'utf8');

test('card principal do desafio em foco usa donut e remove barra horizontal', () => {
  const painelInicio = index.slice(index.indexOf('<section id="screen-painel"'), index.indexOf('<div class="painel-acoes-rapidas"'));

  assert.doesNotMatch(painelInicio, /painel-meta-destaque/);
  assert.match(painelInicio, /class="painel-progresso-donut-grid"/);
  assert.match(painelInicio, /id="progress-donut"[\s\S]*id="progress-percent"[\s\S]*concluído/);
  assert.match(painelInicio, /class="painel-metricas-emocionais"[\s\S]*>Meta<[\s\S]*id="meta-km"[\s\S]*<small>km<\/small>[\s\S]*>Realizado<[\s\S]*id="realizado-km"[\s\S]*<small>km<\/small>[\s\S]*>Faltam<[\s\S]*id="restante-km"[\s\S]*<small>km<\/small>/);
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
  assert.match(styles, /\.painel-metricas-emocionais strong[\s\S]*white-space: nowrap/);
  assert.match(styles, /\.painel-metricas-emocionais small[\s\S]*font-size: 0\.58em/);
});

test('card de desafio reutilizavel nao foi alterado para o donut do painel principal', () => {
  const buildDesafioCard = script.slice(script.indexOf('function buildDesafioCard(item, options)'), script.indexOf('function buildCertificadoDesafioHtml_'));

  assert.match(buildDesafioCard, /desafio-progress-track/);
  assert.match(buildDesafioCard, /id="progress-bar"/);
  assert.doesNotMatch(buildDesafioCard, /painel-progresso-donut-grid|progress-donut|painel-metricas-emocionais/);
});
