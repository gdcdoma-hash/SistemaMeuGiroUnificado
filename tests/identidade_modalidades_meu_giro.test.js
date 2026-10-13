'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const test=require('node:test');
const root=path.resolve(__dirname,'..');
const script=fs.readFileSync(path.join(root,'Meu Giro','Script.html'),'utf8');
const html=fs.readFileSync(path.join(root,'Meu Giro','Index.html'),'utf8');
const css=fs.readFileSync(path.join(root,'Meu Giro','Styles.html'),'utf8');
function extrair(nome) {
  const inicio=script.indexOf('function '+nome+'(');
  const fim=script.indexOf('\nfunction ',inicio+4);
  assert.ok(inicio>=0,'Função ausente: '+nome);
  return script.slice(inicio,fim<0?undefined:fim);
}
const identidade=new Function(extrair('meuGiroIdentidadeModalidade_')+
  '\nreturn meuGiroIdentidadeModalidade_;')();
test('modalidade vem da inscrição e ciclismo legado é preservado',()=>{
  assert.equal(identidade({modalidade_esportiva:'CAMINHADA_CORRIDA'}).nome,'Caminhada/Corrida');
  assert.equal(identidade({modalidade_esportiva:'CICLISMO'}).nome,'Ciclismo');
  assert.equal(identidade({}).classe,'mg-modalidade-ciclismo');
});
test('ícones e lemas são independentes por esporte',()=>{
  const c=identidade({modalidade_esportiva:'CICLISMO'});
  const p=identidade({modalidade_esportiva:'CAMINHADA_CORRIDA'});
  assert.equal(c.icone,'🚴');
  assert.equal(p.icone,'🚶🏃');
  assert.notEqual(c.classe,p.classe);
  assert.equal(c.lema,'Paixão por pedalar');
  assert.equal(p.lema,'Cada Passo, Uma Conquista.');
  assert.doesNotMatch(p.incentivo,/pedal/i);
});
test('troca de foco remove tema anterior sem alterar dados ou Ranking',()=>{
  const names=new Set();
  const el={classList:{
    remove(...arr){arr.forEach(v=>names.delete(v))},
    add(v){names.add(v)}
  }};
  const aplicar=new Function(extrair('aplicarIdentidadeModalidadeElemento_')+
    '\nreturn aplicarIdentidadeModalidadeElemento_;')();
  aplicar(el,identidade({modalidade_esportiva:'CAMINHADA_CORRIDA'}));
  assert.deepEqual([...names],['mg-modalidade-caminhada']);
  aplicar(el,identidade({modalidade_esportiva:'CICLISMO'}));
  assert.deepEqual([...names],['mg-modalidade-ciclismo']);
});
test('Início, Desafios e Detalhes exibem identificação visível e não só cores',()=>{
  for(const id of ['painel-modalidade','painel-card-lema','desafio-detalhe-modalidade'])
    assert.ok(html.includes('id="'+id+'"'), 'Ausente: '+id);
  assert.match(script, /class="mg-modalidade-tag"/);
  assert.match(script, /setTextById\('painel-card-lema', identidadeVisual\.lema\)/);
  assert.match(script, /modalidadeVisual\.nome/);
  assert.match(script, /badge\.setAttribute\('aria-label', 'Modalidade: '/);
});
test('cores distintas, contraste legível e largura móvel',()=>{
  assert.match(css, /--mg-modalidade-cor: #c2410c/);
  assert.match(css, /--mg-modalidade-cor: #0f766e/);
  assert.match(css, /\.desafio-v2-card\.mg-modalidade-caminhada/);
  assert.match(css, /\.painel-inicio-card\.mg-modalidade-caminhada/);
  assert.match(css, /@media \(max-width: 420px\) \{\s*\.mg-modalidade-tag/);
});
