'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const dir=path.join(__dirname,'..','Meu Giro');
const html=fs.readFileSync(path.join(dir,'Index.html'),'utf8');
const css=fs.readFileSync(path.join(dir,'Styles.html'),'utf8');
const script=fs.readFileSync(path.join(dir,'Script.html'),'utf8');

const registrar=html.slice(
  html.indexOf('<section id="screen-registrar"'),
  html.indexOf('<section id="screen-conquistas"')
);
const form=registrar.slice(
  registrar.indexOf('<div class="registro-campos-grid">'),
  registrar.indexOf('<p class="field-helper registro-km-helper">')
);

test('modalidade, data e KM estão na mesma grade e nessa ordem',()=>{
  const p1=form.indexOf('id="registro-modalidade-box"');
  const p2=form.indexOf('id="data-atividade"');
  const p3=form.indexOf('id="km-atividade"');
  assert.ok(p1>0 && p1<p2 && p2<p3,'ordem Modalidade | Data | KM');
  assert.match(form,/^<div class="registro-campos-grid">/);
  assert.equal((registrar.match(/class="registro-campos-grid"/g)||[]).length,1);
  assert.equal((registrar.match(/id="registro-modalidade-box"/g)||[]).length,1);
  assert.match(form,/id="registro-modalidade-box"[^>]*hidden/);
  assert.match(form,/<label for="modalidade-atividade">Modalidade<\/label>/);
  assert.match(form,/<label for="data-atividade">Data da atividade<\/label>/);
  assert.match(form,/<label for="km-atividade">KM realizados<\/label>/);
  assert.ok(registrar.indexOf('id="btn-salvar"')>registrar.indexOf('id="km-atividade"'));
  assert.ok(registrar.indexOf('Atividades registradas')>registrar.indexOf('id="btn-salvar"'));
});

test('permanece dropdown obrigatório sem modalidade selecionada',()=>{
  assert.match(form,/<select id="modalidade-atividade"[^>]*required>/);
  assert.match(form,/<option value="" disabled selected>Escolha a modalidade<\/option>/);
  assert.match(form,/<option value="CICLISMO">Ciclismo<\/option>/);
  assert.match(form,/<option value="CAMINHADA_CORRIDA">Caminhada\/Corrida<\/option>/);
  assert.doesNotMatch(form,/type="radio"|id="botao-modalidade/);
  assert.match(script,/function meuGiroValidarModalidadeRegistro_\(\)/);
  assert.match(script,/meuGiroAvisoValidacaoRegistro_\(/);
});

test('a grade tem 3 colunas com modalidades múltiplas e 2 com uma só',()=>{
  assert.match(css,/\.registro-campos-grid\s*\{[\s\S]*?grid-template-columns:\s*minmax\(0,\s*1fr\)\s+minmax\(0,\s*1fr\)/);
  assert.match(css,/\.registro-campos-grid:has\(>\s*\.registro-modalidade-box:not\(\[hidden\]\)\)\s*\{[\s\S]*?grid-template-columns:\s*minmax\(0,\s*1\.35fr\)\s+minmax\(0,\s*1\.13fr\)\s+minmax\(0,\s*\.86fr\)/);
  assert.match(css,/\.registro-modalidade-box\[hidden\]\s*\{\s*display:\s*none!important/);
  assert.match(css,/@media\s*\(max-width:\s*420px\)\s*\{[\s\S]*?\.registro-campos-grid:has\(>\s*\.registro-modalidade-box:not\(\[hidden\]\)\)/);
  assert.match(css,/#screen-registrar \.registro-campos-grid \.registro-modalidade-box select\s*\{[\s\S]*?font-size:\s*12px/);
});

test('mensagens úteis continuam abaixo da grade, sem bloco longo acima da data',()=>{
  assert.doesNotMatch(form,/Você pode registrar atividades anteriores/);
  assert.match(registrar,/<p class="field-helper registro-km-helper">KM: vírgula ou ponto\./);
  assert.match(registrar,/Datas futuras não são permitidas/);
  assert.match(registrar,/id="btn-salvar"/);
});
