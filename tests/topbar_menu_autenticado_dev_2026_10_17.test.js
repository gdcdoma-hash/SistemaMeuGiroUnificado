'use strict';
const fs=require('node:fs'),cp=require('node:child_process'),assert=require('node:assert/strict');
const before=cp.execFileSync('git',['show','85fa45af3bc7a2c0ca072215de9d52c011a0eb77:Meu Giro/Styles.html'],{encoding:'utf8'});
const current=fs.readFileSync('Meu Giro/Styles.html','utf8');
const old=`  .menu-authenticated {
    grid-template-columns: minmax(0, 120px);
    justify-content: end;
    margin-top: 10px;
  }`;
const revised=`  .menu-authenticated {
    grid-template-columns: repeat(4, minmax(0, 112px));
    justify-content: end;
    margin-top: 10px;
  }

  /* A barra autenticada usa fundo escuro pela navegação integrada. */
  .topbar .brand,
  .topbar .topbar-user-name {
    color: #ffffff;
  }

  .topbar .topbar-user-name::before,
  .topbar .topbar-user-name::after {
    color: #94a3b8;
  }

  @media (max-width: 600px) {
    .menu-authenticated {
      grid-template-columns: repeat(4, minmax(0, 1fr));
    }
  }`;
assert.equal(before.split(old).length,2);
assert.equal(current,before.replace(old,revised),'Somente CSS do menu autenticado pode mudar');
const html=fs.readFileSync('Meu Giro/Index.html','utf8');
const auth=fs.readFileSync('Meu Giro/AuthSession.html','utf8');
assert.match(html,/id="menu-autenticado-top" class="menu menu-top-icons menu-authenticated"/);
assert.match(auth,/auth\.innerHTML='<button class="dgmb-unified-nav-btn"/);
assert.match(auth,/\.topbar\{background:#0f172a!important/);
assert.match(current,/\.menu-authenticated\s*\{\s*grid-template-columns:\s*repeat\(4,\s*minmax\(0,\s*112px\)\)/);
assert.match(current,/@media\s*\(max-width:\s*600px\)\s*\{\s*\.menu-authenticated\s*\{\s*grid-template-columns:\s*repeat\(4,\s*minmax\(0,\s*1fr\)\)/);
assert.match(current,/\.topbar \.brand,\s*\.topbar \.topbar-user-name\s*\{\s*color:\s*#ffffff/);
for(const p of ['Meu Giro/Script.html','Meu Giro/PainelService.gs','Meu Giro/RegistroService.gs','Meu Giro/CertificadoService.gs']){
 const actual=fs.readFileSync(p,'utf8');
 const expected=cp.execFileSync('git',['show','85fa45af3bc7a2c0ca072215de9d52c011a0eb77:'+p],{encoding:'utf8'});
 assert.equal(actual,expected,'Fluxo e dados não podem mudar: '+p);
}
console.log('TOPBAR_DEV_CSS=PASS: navegação em 4 colunas responsiva, marca visível, nenhum backend alterado.');
