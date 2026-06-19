const assert = require('assert');
const fs = require('fs');
const path = require('path');

const script = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'Script.html'),
  'utf8'
);

const helperInicio = script.indexOf('function obterUltimaAtividadePainel_(atividades)');
const renderInicio = script.indexOf('function renderUltimaAtividadePainel(lista)');
const syncInicio = script.indexOf('function sincronizarPainelComDesafioEmFoco()');
const renderPainelInicio = script.indexOf('function renderPainel(p)');

assert.ok(helperInicio >= 0, 'deve existir helper obterUltimaAtividadePainel_');
assert.ok(renderInicio >= 0, 'renderUltimaAtividadePainel deve existir');
assert.ok(syncInicio >= 0, 'sincronizarPainelComDesafioEmFoco deve existir');
assert.ok(renderPainelInicio >= 0, 'renderPainel deve existir');

const helper = script.slice(helperInicio, renderInicio);
const render = script.slice(renderInicio, script.indexOf('\nfunction atualizarContextoVisualRanking', renderInicio));
const sync = script.slice(syncInicio, script.indexOf('\nfunction setTextById', syncInicio));
const renderPainel = script.slice(renderPainelInicio, script.indexOf('\nfunction renderDesafiosV2', renderPainelInicio));

assert.match(
  render,
  /var ultima = obterUltimaAtividadePainel_\(lista\);/,
  'renderUltimaAtividadePainel deve usar a lista real recebida pelo helper'
);
assert.doesNotMatch(
  render,
  /atividades\s*\[\s*0\s*\]/,
  'renderUltimaAtividadePainel não deve depender exclusivamente de atividades[0]'
);
assert.match(
  helper,
  /normalizada\.data\.timestamp > ultima\.data\.timestamp/,
  'helper deve escolher pela maior data normalizada'
);
assert.match(
  helper,
  /normalizada\.data\.timestamp === ultima\.data\.timestamp[\s\S]*?normalizada\.index < ultima\.index/,
  'helper deve desempatar pela ordem original da lista'
);
assert.match(
  script,
  /'data_atividade',[\s\S]*?'data',[\s\S]*?'Data',[\s\S]*?'DATA'/,
  'helper deve aceitar data_atividade, data, Data e DATA'
);
assert.match(
  script,
  /'km',[\s\S]*?'KM',[\s\S]*?'distancia',[\s\S]*?'distancia_km'/,
  'helper deve aceitar km, KM, distancia e distancia_km'
);
assert.match(
  render,
  /diaEl\.innerText = 'Dia ' \+ ultima\.data\.br;/,
  'render deve exibir Dia DD/MM/AAAA'
);
assert.match(
  render,
  /kmEl\.innerText = formatNumber\(ultima\.km\) \+ ' km';/,
  'render deve exibir o km normalizado'
);
assert.match(
  sync,
  /renderUltimaAtividadePainel\(contexto\.atividades \|\| \(currentPainel && currentPainel\.atividades\) \|\| \[\]\);/,
  'sincronizarPainelComDesafioEmFoco deve enviar a lista atualizada para a última atividade'
);
assert.match(
  renderPainel,
  /sincronizarPainelComDesafioEmFoco\(\);/,
  'renderPainel deve sincronizar o painel após receber dados atualizados'
);

console.log('OK: última atividade do painel usa atividades reais, campos compatíveis e maior data.');
