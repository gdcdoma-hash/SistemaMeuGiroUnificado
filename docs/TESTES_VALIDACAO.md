# Testes e validações — Meu Giro

Atualizado em: 2026-09-16

## Certificados
- Branch DEV correta clonada no PC atual.
- `.clasp.json` com Script ID correto.
- `clasp status` funciona.
- Portal -> Meu Giro abre o atleta correto.
- Sessão antiga/atleta errado corrigidos.
- Navegação embutida funciona.
- Retorno `Portal Giro` funciona.
- Popup de retorno aparece.
- Nome do atleta aparece corretamente.
- Retorno para Início dentro do Meu Giro ficou rápido; usuário confirmou.
- Regressão de rolagem vertical da primeira abertura desktop corrigida no conjunto integrado.

## Rejeitados/revertidos
- ResizeObserver/altura dinâmica/READY: regressões desktop/mobile; não repetir sem nova análise.
- Cache integral do iframe: travamento em `Reconectando sua inscrição...`; revertido.

## Regra
Não repetir estes testes apenas por troca de chat ou PC. Retestar se código relacionado for alterado.

## Regressão administrativa futura
Após a revisão do Meu Giro Admin, testar busca, troca de atleta, desafios ativos/encerrados, KM, atividades, correções, auditoria e histórico.