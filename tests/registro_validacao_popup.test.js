'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const test=require('node:test');
const source=fs.readFileSync(path.join(__dirname,'..','Meu Giro','Script.html'),'utf8');

function extrair(nome) {
  const start=source.indexOf('function '+nome+'(');
  assert.ok(start>=0, 'Função '+nome+' presente');
  const end=source.indexOf('\nfunction ',start+8);
  return source.slice(start,end<0?undefined:end);
}
function ambiente(modalidade) {
  const chamadas=[];
  let aceitouModal;
  const modalPromise=new Promise(resolve=>{aceitouModal=resolve;});
  const elementos={
    'modalidade-atividade':{value:modalidade,focus(){this.focused=true;}}
  };
  const funcoes=new Function(
    'document','setRegistroMsg','mostrarModalMeuGiro',
    'releaseRegistroSubmitGuard','meuGiroModalidadeFormularioAtividade_',
    extrair('meuGiroAvisoValidacaoRegistro_')+'\n'+
    extrair('meuGiroValidarModalidadeRegistro_')+
    '\nreturn {validar:meuGiroValidarModalidadeRegistro_,aviso:meuGiroAvisoValidacaoRegistro_};'
  )(
    {getElementById:id=>elementos[id]},
    (msg,variant)=>chamadas.push({tipo:'inline',msg,variant}),
    opcoes=>{chamadas.push({tipo:'popup',opcoes});return modalPromise;},
    ()=>chamadas.push({tipo:'release'}),
    ()=>modalidade
  );
  return {funcoes,chamadas,elementos,aceitouModal};
}
test('modalidade vazia mostra popup no padrão do Meu Giro, sem erro inline',async()=>{
  const app=ambiente('');
  assert.equal(app.funcoes.validar(),false);
  const modal=app.chamadas.find(x=>x.tipo==='popup');
  assert.ok(modal);
  assert.equal(modal.opcoes.titulo,'Escolha a modalidade');
  assert.equal(modal.opcoes.tipo,'alerta');
  assert.equal(modal.opcoes.origem,'decisao');
  assert.equal(modal.opcoes.processando,false);
  assert.equal(modal.opcoes.botoes[0].texto,'Entendi');
  assert.ok(app.chamadas.some(x=>x.tipo==='release'));
  assert.ok(!app.chamadas.some(x=>x.tipo==='inline'&&x.variant==='error'));
  app.aceitouModal(true);
  await Promise.resolve();
  assert.equal(app.elementos['modalidade-atividade'].focused,true);
});
test('modalidade preenchida não apresenta popup de erro',()=>{
  for(const modalidade of ['CICLISMO','CAMINHADA_CORRIDA']){
    const app=ambiente(modalidade);
    assert.equal(app.funcoes.validar(),true);
    assert.equal(app.chamadas.length,0);
  }
});
test('validação local de data e km usa popup no registro e na edição',()=>{
  for(const nome of ['salvarAtividade','salvarEdicaoAtividade']){
    const inicio=source.indexOf('async function '+nome+'(');
    assert.ok(inicio>0);
    const fim=source.indexOf(nome==='salvarAtividade'?'async function salvarEdicaoAtividade(':'function iniciarEdicaoAtividadeUI(',inicio+5);
    const corpo=source.slice(inicio,fim);
    assert.match(corpo,/meuGiroAvisoValidacaoRegistro_\('Informe os dados da atividade'/);
    assert.match(corpo,/meuGiroAvisoValidacaoRegistro_\('Confira a data'/);
    assert.match(corpo,/meuGiroAvisoValidacaoRegistro_\('Confira a distância'/);
    assert.match(corpo,/meuGiroAvisoValidacaoRegistro_\('Data futura não permitida'/);
    assert.doesNotMatch(corpo,/setRegistroMsg\('Informe uma data válida\.', 'error'\)/);
    assert.doesNotMatch(corpo,/setRegistroMsg\('Informe um KM válido maior que zero\.', 'error'\)/);
  }
});
