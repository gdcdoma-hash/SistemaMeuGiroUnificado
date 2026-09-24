# Estado atual do Meu Giro — 24/09/2026

Este documento é a referência principal para retomada do desenvolvimento do Meu Giro.

## 1. Repositório, ambiente e deploy

- Repositório: `gdcdoma-hash/SistemaMeuGiroUnificado`
- Branch operacional: `main`
- Código Apps Script ativo: pasta `Meu Giro/`
- Script ID DEV: `1N-10MQMYaq_91O75W_TFF837vPR6KfSDpJwVQJNS7776IwGo56NuppzG`
- Deployment DEV: `AKfycbxCV-6fDri2y2ppdPC1JPLqCTxEOwXLuSameujgVkoYokm-sfQgPtDY4oQQ0Z_uVCmRKg`
- Versão DEV confirmada em 24/09/2026: `@82`
- Workflow: `.github/workflows/sync-gas.yml`
- O workflow executa `clasp push --force` e atualiza automaticamente este deployment DEV.
- Produção não deve ser atualizada automaticamente.

## 2. Planilha oficial

O Meu Giro usa a mesma planilha oficial do Portal Giro:

`1sFxmiWmIrPlwXgSb56M8maHF5niRYVLWRMFSUQYXgHM`

Abas relevantes:
- `DadosPessoais`
- `dgmbDesafios`
- `REGISTRO_KM`
- `MEU_GIRO_RESUMO`
- `ListaDesafios`
- `FRASES`
- `CONFIG_CERTIFICADO_TEMPLATE`
- `_PORTAL_HANDOFF`

Não voltar a IDs históricos de planilha sem migração explícita.

## 3. Relação com o Portal Giro

Portal Giro e Meu Giro são Web Apps separados.

Portal Giro:
- autentica o atleta;
- gerencia inscrições;
- abre o Meu Giro por handoff seguro.

Meu Giro:
- recebe o handoff;
- cria sua sessão;
- exibe painel/desafios/atividades/conquistas;
- pode retornar ao Portal sem exigir novo login.

### Handoff
Portal -> Meu Giro:
1. Portal valida o atleta autenticado.
2. Cria token temporário em `_PORTAL_HANDOFF`.
3. Abre o Meu Giro com `page=atleta&handoff=TOKEN`.
4. Quando embutido no Portal, acrescenta `embedded=1`.
5. Meu Giro troca o handoff por sessão própria.

Meu Giro -> Portal:
- no modo embutido usa `postMessage`/ACK;
- fora do modo embutido pode usar handoff de retorno.

## 4. Regra crítica de iframe

O `doGet()` deve manter:

```javascript
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
```

Sem isso, o Meu Giro não carrega dentro do iframe do Portal e o navegador exibe:

`script.google.com recusou estabelecer ligação`

Esta configuração foi restaurada e homologada em 24/09/2026.

## 5. Navegação

### Quando aberto dentro do Portal
`Portal Giro | Início | Desafios | Registrar | Conquistas`

### Quando aberto diretamente
O Meu Giro preserva acesso/login próprio e não depende obrigatoriamente do Portal para existir.

## 6. Tela Início / desafio em foco

O card principal atual mostra:
- Desafio em foco
- Nome do desafio
- Nome do atleta
- Início
- Término
- Prazo
- Tempo restante
- Meta
- Realizado
- Faltam
- Percentual concluído
- Mensagem motivacional
- assinatura visual do Desafio Giro

O campo `Tempo restante` usa a data final e apresenta:
- N dias;
- `Último dia`;
- `Encerrado`.

## 7. PRAZO_DIAS

Regra alinhada ao Portal Giro:
- `Prazo_Dias > 0` identifica desafio de janela individual;
- `data_consolidacao` define que a janela individual foi consolidada;
- após consolidado, usar:
  - `data_inicio_desafio`
  - `data_fim_desafio`
- nunca substituir `data_fim_desafio` pelo fim mensal da `ListaDesafios`.

O resumo leve e o painel devem respeitar essa mesma regra.

Exemplo homologado:
- DESAFIO 1000 KM
- Início: 01/09/2026
- Término: 15/10/2026
- Prazo: 45 dias

## 8. Compartilhamento do progresso

O card da tela Início pode ser transformado em PNG e compartilhado.

Implementação atual:
- donut de progresso em SVG;
- html2canvas carregado sob demanda;
- Web Share API quando o navegador aceita arquivo;
- download PNG como fallback;
- botão de compartilhar compacto com ícone.

O SVG foi adotado porque `conic-gradient` não era reproduzido corretamente na imagem gerada.

### Próxima missão já registrada
- identidade/texto do botão como **Compartilha Meu Giro**;
- permitir compartilhamento individual dos cards da tela Desafios, mantendo navegação horizontal.

## 9. Cálculo e dados de progresso

O Meu Giro usa:
- `dgmbDesafios` para vínculo, meta, status e janela;
- `REGISTRO_KM` para atividades;
- `MEU_GIRO_RESUMO` para resumo persistente de cálculo;
- `ListaDesafios` para metadados do desafio.

O resumo persistente não pode reintroduzir fallback mensal em PRAZO_DIAS.

## 10. Status de conclusão

Regra operacional:
- 100% de KM não significa, sozinho, conclusão administrativa oficial.
- `Status_Usuario_Desafio` continua sendo relevante.
- O painel pode indicar meta atingida aguardando validação antes da conclusão oficial.

## 11. Sessão e identidade

Preservar:
- `AuthSession.html`
- `AtletaSessionService.gs`
- `PortalAtletaReturnService.gs`
- `PortalHandoffService.gs`

Não simplificar sessão/handoff durante alteração visual.

A sessão integrada deve sempre corresponder ao mesmo atleta autenticado no Portal.

## 12. Deploy e testes

Fluxo normal:
1. commit na `main`;
2. aguardar workflow;
3. confirmar `Sincronizar com Meu Giro GAS: success`;
4. confirmar `Atualizar deployment DEV: success`;
5. só então testar o `/exec`.

Um `clasp push` sem atualização do deployment pode deixar o usuário testando uma versão antiga. Por isso o workflow atualiza os dois passos.

## 13. Invariantes

Não quebrar:
1. GitHub é canônico.
2. Código ativo fica em `Meu Giro/`.
3. Branch atual é `main`.
4. Planilha oficial compartilhada permanece a mesma do Portal.
5. `ALLOWALL` permanece no `doGet()`.
6. Handoff temporário permanece como ponte de identidade.
7. PRAZO_DIAS usa datas individuais após consolidação.
8. Deployment DEV usado pelo Portal deve ser o mesmo atualizado pelo workflow.
9. Produção não é automática.

## 14. Pendências conhecidas

- Compartilha Meu Giro: refinamento do botão/texto.
- Compartilhamento dos cards individuais da tela Desafios.
- Melhorias futuras no Ranking, quando priorizadas.
- Integração Strava permanece futura.
- Melhorias de performance devem evitar reintroduzir cache integral de iframe ou outras estratégias já revertidas por regressão.
