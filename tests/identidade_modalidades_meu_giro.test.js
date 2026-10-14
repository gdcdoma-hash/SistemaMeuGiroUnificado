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
const identidade=new Function(extrair('meuGiroIconeModalidade_')+
  '\n'+extrair('meuGiroIdentidadeModalidade_')+
  '\nreturn meuGiroIdentidadeModalidade_;')();
test('modalidade vem da inscrição e ciclismo legado é preservado',()=>{
  assert.equal(identidade({modalidade_esportiva:'CAMINHADA_CORRIDA'}).nome,'Caminhada/Corrida');
  assert.equal(identidade({modalidade_esportiva:'CICLISMO'}).nome,'Ciclismo');
  assert.equal(identidade({}).classe,'mg-modalidade-ciclismo');
});
test('ícones e lemas são independentes por esporte',()=>{
  const c=identidade({modalidade_esportiva:'CICLISMO'});
  const p=identidade({modalidade_esportiva:'CAMINHADA_CORRIDA'});
  assert.match(c.icone,/^<svg/);
  assert.match(p.icone,/^<svg/);
  assert.match(c.icone,/stroke="currentColor"/);
  assert.match(p.icone,/stroke="currentColor"/);
  assert.match(p.icone,/mg-modalidade-icon-duo/);
  assert.equal((p.icone.match(/<circle/g)||[]).length,2);
  assert.doesNotMatch(c.icone + p.icone,/🚴|🚶|🏃/);
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
  assert.match(script, /badge\.innerHTML = '<span aria-hidden="true">'/);
  assert.match(script, /badgeModalidade\.innerHTML = '<span aria-hidden="true">'/);
  assert.match(html, /id="painel-modalidade"[\s\S]*?mg-modalidade-icon/);
  assert.match(html, /<option value="CICLISMO">Ciclismo<\/option>/);
  assert.match(html, /<option value="CAMINHADA_CORRIDA">Caminhada\/Corrida<\/option>/);
});
test('cores distintas, contraste legível e largura móvel',()=>{
  assert.match(css, /--mg-modalidade-cor: #c2410c/);
  assert.match(css, /--mg-modalidade-cor: #0f766e/);
  assert.match(css, /\.mg-modalidade-icon \{/);
  assert.match(css, /color: inherit/);
  assert.match(css, /\.desafio-v2-card\.mg-modalidade-caminhada/);
  assert.match(css, /\.painel-inicio-card\.mg-modalidade-caminhada/);
  assert.match(css, /@media \(max-width: 420px\) \{\s*\.mg-modalidade-tag/);
});

test('aviso motivacional separa o texto do SVG em ambas as modalidades',()=>{
  const montarStatus=new Function(
    'meuGiroIdentidadeModalidade_','normalizeRitmoText','painelMetaConcluida_','getRitmoAgoraVariant',
    extrair('obterStatusMotivacionalFeed_')+'\nreturn obterStatusMotivacionalFeed_;'
  )(identidade,x=>x,()=>false,()=> 'alerta-ritmo-atencao');
  for (const modalidade of ['CICLISMO','CAMINHADA_CORRIDA']) {
    const status=montarStatus({ritmo_status:'abaixo'}, {modalidade_esportiva:modalidade});
    assert.equal(status.titulo,'Hora de reagir.');
    assert.match(status.icone,/^<svg/);
    assert.doesNotMatch(status.titulo,/<svg/);
  }
  // A arte vai para um elemento próprio; somente o título textual é escapado.
  assert.match(script,/status\.icone \? '<span class="mg-status-modalidade-icon"/);
  assert.match(script,/escapeHtml\(status\.titulo\)/);
  assert.doesNotMatch(script,/titulo:\s*identidadeVisual\.icone\s*\+/);
});

test('aviso do painel insere o ícone como elemento e não mostra a marcação SVG',()=>{
  const elems={
    'alerta-ritmo-titulo':{textContent:'',inserido:'',insertAdjacentHTML(pos,html){this.inserido+=html;}},
    'alerta-ritmo-texto':{textContent:''},
    'alerta-ritmo':{classList:{remove(){},add(){}}}
  };
  const mockDoc={getElementById(id){return elems[id]||null;}};
  const setTexto=(id,txt)=>{elems[id].textContent=txt;};
  const render=new Function(
    'document','normalizeRitmoText','painelMetaConcluida_','getRitmoAgoraVariant',
    'meuGiroIdentidadeModalidade_','getDesafioEmFocoAtual','setTextById',
    extrair('renderRitmoAgora')+'\nreturn renderRitmoAgora;'
  )(mockDoc,x=>x,()=>false,()=> 'alerta-ritmo-atencao',
    identidade,()=>({modalidade_esportiva:'CAMINHADA_CORRIDA'}),setTexto);
  render({ritmo_status:'abaixo'});
  assert.equal(elems['alerta-ritmo-titulo'].textContent,'Hora de reagir.');
  assert.match(elems['alerta-ritmo-titulo'].inserido,/^<span class="mg-status-modalidade-icon"/);
  assert.match(elems['alerta-ritmo-titulo'].inserido,/<svg/);
  assert.doesNotMatch(elems['alerta-ritmo-titulo'].textContent,/<svg/);
  assert.equal(elems['alerta-ritmo-texto'].textContent,'Sua meta está ficando distante.');
});
