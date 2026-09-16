# Continuidade técnica — Meu Giro

Atualizado em: 2026-09-16

## Finalidade
Ponto de retomada do Meu Giro. Antes de alterar código em novo PC/chat, ler este arquivo e a documentação principal do Portal Giro no repositório `gdcdoma-hash/dgmb-inscricoes-repescagem`.

## Estado atual
- Repositório: `gdcdoma-hash/SistemaMeuGiroUnificado`
- Branch: `feat-meu-giro-admin-2026-09-13`
- Ambiente: desenvolvimento/homologação
- Pasta Apps Script conhecida: `D:\Projetos\SistemaMeuGiro-DEV-LIMPO-2026-09-15\Meu Giro`
- Script ID GAS DEV: `1N-10MQMYaq_91O75W_TFF837vPR6KfSDpJwVQJNS7776IwGo56NuppzG`
- Web App DEV usado pelo Portal no último estado certificado: `https://script.google.com/macros/s/AKfycbxCV-6fDri2y2ppdPC1JPLqCTxEOwXLuSameujgVkoYokm-sfQgPtDY4oQQ0Z_uVCmRKg/exec`

## Estado estável certificado
- Portal -> Meu Giro identifica corretamente o atleta.
- Handoff integrado é aplicado antes que estados antigos do navegador possam assumir a sessão.
- Navegação embutida: `Portal Giro | Início | Desafios | Registrar | Conquistas`.
- Ranking permanece retirado até reconstrução.
- Retorno ao Portal funciona sem novo login.
- Popup de retorno ao Portal aparece.
- Nome do atleta permanece visível no fluxo autenticado.
- Retorno interno para `Início` ficou rápido após reaproveitamento do painel carregado.
- Rolagem vertical anormal na primeira abertura desktop foi corrigida no conjunto Portal/Meu Giro.

## Commits relevantes conhecidos
- `c4b172c0a144a97a780240ecff043376d7d0907b` — aplica modo embutido antes da renderização / identidade correta.
- `cde2ce51c4c1d69e55efadecc6f6deb9e724657a` — reaproveita painel carregado ao voltar ao Início.
- `4339414e4da0455a10622002187ba4e2dab07075` — popup imediato ao retornar ao Portal.
- `4c7ba3a01ddd5bab08f5a95e86f8a13413f3e451` — restaura abertura estável após tentativa problemática de cache integral do iframe.

## Estratégias revertidas — não repetir sem nova análise
1. ResizeObserver/altura dinâmica/READY entre iframe e Portal: causou corte/área escura no desktop, lentidão e mobile preso no loading.
2. Manter o iframe completo do Meu Giro carregado para retorno instantâneo: causou `Reconectando sua inscrição...`.

## Revisão administrativa pendente
O Meu Giro Admin faz parte da lista-mestra de 15 fases documentada no Portal. Nenhuma fase foi implementada ainda.

Itens específicos do Meu Giro Admin já aprovados no planejamento:
- topo com `BUSCAR ATLETA | DADOS DO ATLETA`;
- preservar busca por Código/Nome/CPF/ID_DGMB e `Trocar atleta`;
- desafios ativos em bloco navegável anterior/próximo;
- desafios encerrados em bloco separado navegável anterior/próximo;
- atividades com filtros `Ano | Mês` e paginação quando necessário;
- histórico administrativo com filtros `Ano | Mês | Tipo` e paginação;
- remover navegação/nome do atleta da área administrativa e usar padrão `Painel Admin`.

## Ponto exato onde paramos
Planejamento administrativo concluído; nenhuma alteração dessa revisão foi implementada.

## Próxima missão de desenvolvimento
A próxima missão global é **Fase 1 — Padrão único da área administrativa**, começando pelo Portal. Não iniciar isoladamente alterações do Meu Giro Admin antes de validar o padrão global.

## Regra para novo chat
Não refazer auditorias, backups ou certificações já aprovadas sem mudança relacionada. Consultar primeiro a documentação do Portal e esta documentação. Produção não deve ser modificada automaticamente.