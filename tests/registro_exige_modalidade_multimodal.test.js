'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const test=require('node:test');
const root=path.join(__dirname,'..','Meu Giro');
const script=fs.readFileSync(path.join(root,'Script.html'),'utf8');
const html=fs.readFileSync(path.join(root,'Index.html'),'utf8');
const server=fs.readFileSync(path.join(root,'RegistroService.gs'),'utf8');

function controlador(modalidades,focusModalidade='CICLISMO'){
  const elementos={
    'modalidade-atividade':{value:'CICLISMO',required:false,focus(){this.focused=true;}},
    'registro-modalidade-box':{hidden:false}
  };
  let mensagem='',liberou=0;
  const documento={getElementById:id=>elementos[id]||null};
  const a=script.indexOf('var meuGiroContextoModalidadeFormulario_ = ');
  const b=script.indexOf('\nfunction renderAtividades(',a);
  assert.ok(a>=0&&b>a);
  const funcoes=new Function('document','meuGiroConfigurarLimiteDataRegistro_',
    'getDesafioEmFocoAtual','currentPainel','activeEditKey','releaseRegistroSubmitGuard',
    'setRegistroMsg','mostrarModalMeuGiro',script.slice(a,b)+
    '\nreturn {sync:sincronizarModalidadeFormularioAtividade_,'+
    'get:meuGiroModalidadeFormularioAtividade_,'+
    'validate:meuGiroValidarModalidadeRegistro_};');
  const state=funcoes(documento,()=>{},()=>({modalidade_esportiva:focusModalidade}),{
    id_dgmb:'TESTE',desafios_ativos:modalidades.map(modalidade_esportiva=>({modalidade_esportiva}))
  },'',()=>liberou++,t=>mensagem=t,opcoes=>{mensagem=opcoes.mensagem;return Promise.resolve(true);});
  return {state,elementos,getMensagem:()=>mensagem,getLiberou:()=>liberou};
}

test('com Ciclismo e Caminhada/Corrida, não há modalidade pré-selecionada',async()=>{
  const c=controlador(['CICLISMO','CAMINHADA_CORRIDA']);
  assert.equal(c.state.sync(),'');
  assert.equal(c.state.get(),'');
  assert.equal(c.elementos['registro-modalidade-box'].hidden,false);
  assert.equal(c.elementos['modalidade-atividade'].required,true);
  assert.equal(c.state.validate(),false);
  assert.match(c.getMensagem(),/Selecione Ciclismo ou Caminhada\/Corrida/);
  await Promise.resolve(); // foco vem apos a confirmacao do popup
  assert.equal(c.elementos['modalidade-atividade'].focused,true);
  assert.equal(c.getLiberou(),1);
});

test('escolha manual é mantida durante atualização mas zerada em novo registro',()=>{
  const c=controlador(['CICLISMO','CAMINHADA_CORRIDA']);
  c.state.sync();
  c.elementos['modalidade-atividade'].value='CAMINHADA_CORRIDA';
  assert.equal(c.state.sync(),'CAMINHADA_CORRIDA');
  assert.equal(c.state.validate(),true);
  assert.equal(c.state.sync('',true),'');
  assert.equal(c.state.validate(),false);
});

test('edição explicita a modalidade original, mas não troca silenciosamente',()=>{
  const c=controlador(['CICLISMO','CAMINHADA_CORRIDA']);
  c.state.sync('CAMINHADA_CORRIDA');
  assert.equal(c.state.get(),'CAMINHADA_CORRIDA');
});

test('atleta com uma modalidade mantém preenchimento automático',()=>{
  for(const modalidade of ['CICLISMO','CAMINHADA_CORRIDA']){
    const c=controlador([modalidade],modalidade);
    assert.equal(c.state.sync('',true),modalidade);
    assert.equal(c.state.get(),modalidade);
    assert.equal(c.elementos['registro-modalidade-box'].hidden,true);
    assert.equal(c.state.validate(),true);
  }
});

test('não há fallback silencioso a Ciclismo para múltiplas modalidades',()=>{
  assert.match(html,/<option value="" disabled selected>Escolha a modalidade<\/option>/);
  assert.match(script,/if \(destinoNormalizado === 'register' && !activeEditKey\)/);
  assert.match(script,/sincronizarModalidadeFormularioAtividade_\('', true\)/);
  assert.match(script,/if \(!meuGiroValidarModalidadeRegistro_\(\)\) return;/);
  assert.match(server,/MODALIDADE_OBRIGATORIA/);
  assert.match(server,/obterVinculosDesafioUsuario_\(idDgmb\)/);
  assert.match(server,/var modalidadeRegistro = meuGiroModalidadeCanonica_\(modalidadeEsportiva\)/);
});
