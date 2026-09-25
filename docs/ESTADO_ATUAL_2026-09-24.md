# Estado atual do Meu Giro — 24/09/2026

Este documento é a referência principal para retomada do desenvolvimento do Meu Giro.

## 1. Repositório, ambiente e deploy

- Repositório: `gdcdoma-hash/SistemaMeuGiroUnificado`
- Branch operacional: `main`
- Código Apps Script ativo: pasta `Meu Giro/`
- Script ID DEV: `1N-10MQMYaq_91O75W_TFF837vPR6KfSDpJwVQJNS7776IwGo56NuppzG`
- Deployment DEV: `AKfycbxCV-6fDri2y2ppdPC1JPLqCTxEOwXLuSameujgVkoYokm-sfQgPtDY4oQQ0Z_uVCmRKg`
- Versão DEV confirmada em 24/09/2026: `@84`
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

## 5.1. Bootstrap da navegação integrada

Em 24/09/2026 foi corrigido um estado neutro em que o Meu Giro abria pelo Portal sem a navegação própria e sem o botão de retorno ao Portal.

Regra atual:
- o servidor injeta no boot `embedded`, `page` e `handoff`;
- quando existe `page=atleta` + `handoff`, o handoff integrado tem prioridade sobre restauração de sessão local;
- o modo embedded é aplicado antes de a UI autenticada ser sincronizada;
- no modo embedded deve aparecer `Portal Giro | Início | Desafios | Registrar | Conquistas`;
- no modo independente autenticado deve aparecer a navegação autenticada do Meu Giro, e não apenas o botão Sair.

A função `syncTopNavigation` não deve voltar a ocultar incondicionalmente `menu-autenticado-top`.

## 5.2. Estado estabilizado da integração

**Homologado pelo usuário em 24/09/2026.**

Estado estável:
- Portal Giro abre o Meu Giro corretamente em iframe.
- Meu Giro embutido exibe `Portal Giro | Início | Desafios | Registrar | Conquistas`.
- O retorno por `Portal Giro` funciona sem novo login.
- O modo independente do Meu Giro preserva sua navegação autenticada.
- Não há mais estado neutro sem navegação.
- Não há mais bloqueio `script.google.com recusou estabelecer ligação` no fluxo homologado.

Ponto técnico certificado:
- deployment Meu Giro DEV: `@84`;
- `ALLOWALL` é obrigatório no `doGet()`;
- o boot integrado contém `embedded`, `page` e `handoff`;
- o handoff integrado tem prioridade sobre restauração local;
- `syncTopNavigation` não deve ocultar incondicionalmente o menu autenticado.

**Regra de continuidade:** não alterar esta integração por refatoração ampla. Só mexer novamente em iframe, handoff, bootstrap ou navegação se houver regressão reproduzível ou uma nova missão explícita.

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


## Correcao Meu Giro Admin - 25/09/2026
- Sintoma identificado no Portal Admin: ao buscar atleta no modulo Meu Giro Admin, aparecia "Resposta invalida recebida do Meu Giro".
- Causa: o Portal chamava a rota `?api=admin`, mas o `doGet` atual do Meu Giro sempre devolvia a pagina HTML do atleta.
- Correcao implementada no commit `25c073633384a19041df5f3e3006765077669f44`.
- Criado `Meu Giro/AdminApiService.gs` com as operacoes administrativas protegidas pelo handoff ADMIN:
  - buscar atleta;
  - consultar painel do atleta;
  - corrigir atividade;
  - registrar historico administrativo da correcao.
- `Meu Giro/Code.gs` agora direciona chamadas `api=admin` para resposta JSON e aceita POST administrativo.
- GitHub Action #38 concluida com SUCCESS, incluindo sincronizacao do GAS e atualizacao do deployment DEV.
- Fluxo do atleta nao foi alterado.
- Homologacao pendente: repetir busca no Meu Giro Admin pelo Portal.


## Separacao DEV/PROD do Meu Giro - etapa inicial - 25/09/2026
- Busca e abertura de atleta no Meu Giro Admin foram homologadas pelo responsavel apos a correcao da API administrativa.
- Commit funcional da configuracao por ambiente: `f92f3b5a338e4703a607d9c58a4baced134f2a68`.
- `Meu Giro/Config.gs` passou a centralizar:
  - `DGMB_ENVIRONMENT`;
  - `DGMB_SPREADSHEET_ID`;
  - `DGMB_PORTAL_WEBAPP_URL`;
  - `DGMB_CERTIFICADOS_FOLDER_ID`;
  - `DGMB_CERTIFICADO_TEMPLATE_ID`.
- Script Properties tem prioridade; os valores DEV atuais permanecem como fallback para nao alterar o funcionamento durante a migracao.
- `PortalAtletaReturnService.gs` passou a obter a URL do Portal pela configuracao central.
- `CertificadoService.gs` passou a obter pasta e template padrao pela configuracao central.
- A planilha oficial do DEV permanece a mesma.
- Nenhum GAS PROD foi criado e nenhum ambiente publico antigo foi alterado.
- GitHub Action #39: SUCCESS, incluindo sincronizacao GAS DEV e atualizacao do deployment DEV.
- Homologacao pendente: abrir Meu Giro pelo Portal, confirmar carregamento dos dados e retornar ao Portal.


## Configuracao segura por ambiente do Meu Giro - 25/09/2026
- O teste apos o rollback confirmou ida Portal -> Meu Giro e retorno Meu Giro -> Portal funcionando normalmente.
- A tentativa anterior com configuracao calculada na inicializacao global foi descartada.
- Nova abordagem implementada no commit `c9fd947c5bce1ce568cdd77279f9d68022d2ce62`.
- Criado `Meu Giro/EnvironmentConfig.gs` somente com funcoes de leitura em tempo de execucao; nenhum valor e calculado na abertura global do Apps Script.
- Chaves preparadas: `DGMB_SPREADSHEET_ID`, `DGMB_PORTAL_WEBAPP_URL`, `DGMB_CERTIFICADOS_FOLDER_ID` e `DGMB_CERTIFICADO_TEMPLATE_ID`.
- Os valores DEV atuais permanecem como fallback, portanto nenhum ID, URL, pasta ou template foi trocado.
- `getSpreadsheet_()`, retorno ao Portal e servicos de certificado passaram a usar a leitura segura nos pontos migrados.
- `SPREADSHEET_ID` antigo permanece temporariamente para compatibilidade com arquivos ainda nao migrados.
- GitHub Action #41 concluiu com SUCCESS, incluindo sincronizacao do GAS e atualizacao do deployment DEV.
- Homologacao pendente: repetir Portal -> Meu Giro -> Portal.
