# Meu Giro — checkpoint técnico e funcional de 14/10/2026

**Estado:** fotografia do código da branch `main` após PRs #279–#291, complementada pelas homologações expressas do organizador. Este documento **não** é uma declaração de que cada fluxo já foi testado no PROD.

**Leitura obrigatória para alterações futuras:** este checkpoint + [estado operacional de 24/09](ESTADO_ATUAL_2026-09-24.md) + [README do projeto](../README.md). O documento de 24/09 conserva histórico técnico e identificadores de ambiente; havendo divergência, verificar o código e as configurações em execução antes de decidir.

## 1. Fronteiras de projeto e ambientes

- Repositório canônico do Meu Giro: `gdcdoma-hash/SistemaMeuGiroUnificado`, código Apps Script em `Meu Giro/`, branch operacional DEV: `main`.
- Portal Giro e Meu Giro são Web Apps independentes. Portal gerencia autenticação/inscrições; Meu Giro gerencia painel, desafios, registros e conquistas. Handoff temporário, sessão própria e, quando embutido, iframe + `postMessage` devem continuar funcionando.
- `doGet()` precisa preservar `HtmlService.XFrameOptionsMode.ALLOWALL` para integração embutida. Não alterar handoff, sessão nem navegação por refatoração visual.
- O workflow `.github/workflows/sync-gas.yml` sincroniza `Meu Giro/` e atualiza o deployment **DEV** ao mudar arquivos do aplicativo em `main`. Conferir sucesso do workflow antes de afirmar que o deployment foi atualizado. Arquivos apenas em `docs/` e `tests/` não acionam esse deploy pela configuração vigente.
- Há Meu Giro PROD separado e integração com Portal PROD já homologada em 28/09 (navegação/leitura). **Nenhuma** mudança DEV/LAB é promoção automática ao PROD; deploy e testes PROD são tarefas separadas.
- Manter segregação dos dados e não criar inscrições ou atividades em PROD durante homologações DEV. Não imprimir CPF, tokens de handoff ou dados de atletas em documentos de teste.

## 2. Fontes dos dados e vínculos

- `dgmbDesafios`: inscrições/vínculos, meta, modalidade, status e período individual.
- `REGISTRO_KM`: atividades e suas modalidades, data e distância; uma atividade pertence ao atleta/esporte, não é duplicada por desafio.
- `MEU_GIRO_RESUMO`: progresso e resumo persistido; apuração deve conciliar com registros e janelas das inscrições.
- `ListaDesafios`: catálogo/metadados; nunca substituir pelo mês do catálogo o limite individual de `PRAZO_DIAS`.
- `DadosPessoais`, `_PORTAL_HANDOFF` e tabelas de certificados possuem funções específicas; preservar contratos existentes.

## 3. Modalidades e registros — estado de 14/10

- Modalidades canônicas: `CICLISMO` e `CAMINHADA_CORRIDA`; caminhada e corrida compartilham modalidade. Atividade legada sem modalidade continua sendo Ciclismo por compatibilidade.
- Com inscrições nas duas modalidades, novo registro inicia em `Escolha a modalidade` (valor vazio). Não herdar esporte do card em foco nem selecionar Ciclismo automaticamente. Escolha explícita é obrigatória; validação também existe no backend. Quando o atleta tem apenas uma modalidade, ela pode ser aplicada sem mostrar o seletor.
- Em edição, preservar a modalidade originalmente registrada. Editar/excluir atividades está funcionando no fluxo de homologação; não trocar modalidade ou inscrição silenciosamente.
- Registrar e editar: datas reais **até o dia corrente** (fuso `America/Sao_Paulo`) podem integrar o histórico pessoal mesmo fora da janela da meta; datas futuras são bloqueadas no cliente e no servidor. Atividade fora de prazo não avança desafio incompatível.
- Avisos de validação (modalidade, data e KM) devem aparecer em **popup padrão**, não escondidos como mensagens inline no formulário. Estados de carregamento/sucesso podem usar feedback contextual.
- Campos usam, na tela **Registrar**, a **mesma linha e nesta ordem exata**: `Modalidade | Data da atividade | KM realizados`. **Manter dropdown**, não substituir por botões. Rótulos e controles alinhados verticalmente; botão `Registrar atividade` imediatamente abaixo, seguido de `Atividades registradas`, ampliando área útil do histórico.
- Correção da grade: PR #290; alinhamento vertical corrigido: PR #291; questão #289 **homologada e encerrada** pelo organizador em 14/10 com a instrução `ficou perfeito, mantenha assim`. Evitar reintroduzir seletor grande, deslocamento vertical ou texto de ajuda extenso acima dos campos.

## 4. Progresso, períodos e simultaneidade

- Desafios com `PRAZO_DIAS`: depois da consolidação, respeitar `data_inicio_desafio` e `data_fim_desafio` da **inscrição individual**, mesmo atravessando meses. Ex.: 45 dias, 01/09/2026–15/10/2026, sem corte em 30/09.
- Mensais por KM seguem a própria janela oficial.
- Uma atividade pode contar simultaneamente para dois desafios da **mesma modalidade** cujas janelas incluam a data, sem duplicar a linha em `REGISTRO_KM`.
- Jamais compartilhar KM entre Ciclismo e Caminhada/Corrida. Fora da janela, conservar no histórico pessoal mas **não** somar à meta daquele desafio.
- A data da atividade também precisa ser igual ou anterior a hoje para apuração; registros futuros preexistentes/testes não antecipam progresso.
- Percentual de meta atingido não equivale automaticamente à conclusão administrativa: respeitar `Status_Usuario_Desafio` e fluxos de validação.
- PR #286 corrigiu o fallback de modalidade legada na construção de vínculos (`contextoLista`, não variável inexistente) e incluiu testes de 45/90 dias, meses e modalidades.

## 5. Telas e identidade visual

- Navegação autenticada embutida: `Portal Giro | Início | Desafios | Registrar | Conquistas`. Retorno ao Portal sem novo login permanece essencial.
- Na tela **Início**, cada desafio simultâneo deve aparecer como **um card inteiro por vez**, com controles Anterior/Próximo e indicador `1 de N`; não mostrar pedaços de outro card como falsa prévia (PR #282).
- Identificar esportes por **ícones monocromáticos e texto**, não apenas por cor (PRs #283–#285). Ciclismo: `Paixão por pedalar`; Caminhada/Corrida: `Cada Passo, Uma Conquista.`.
- Metas, realizado, faltante, percentuais, data inicial/final, prazo e progresso devem refletir o desafio em foco. PRAZO_DIAS exibe a faixa individual compacta.
- Compartilhamento PNG do cartão principal existe; refinamentos de nomenclatura/compartilhamento individual de outros cards seguem pendentes conforme documentação anterior.

## 6. Testes e evidências

**Já confirmado:** cadastro do atleta de teste 1603 com as duas modalidades; registro independente de Ciclismo e Caminhada/Corrida; operações básicas de edição e exclusão; histórico retroativo; bloqueio de data futura; formulário e popup; navegação por desafio e identidade visual; alinhamento final aprovado em 14/10.

**Regra no código validada em memória em 14/10** com `meuGiroSomarKmPorVinculo_` e data simulada 14/10:
- 30/09: Ciclismo 10 km; 12/10: Ciclismo 6 km; 12/10: Caminhada/Corrida 5,5 km.
- Desafio Ciclismo 1000 KM, 45 dias, 01/09–15/10: **16 km**.
- Outubro Rosa (Ciclismo), 01/10–31/10: **6 km**.
- Desafio de Caminhada/Corrida, 01/10–31/10: **5,5 km**.
- Registro de 16/10 é futuro na data simulada e não contribui, embora esteja no fixture.
- Isto é **fixture fictício**, não prova dos totais reais da planilha.

**Pendente de homologação na interface DEV/LAB:** no atleta de teste 1603, conferir presença (e evitar duplicação) de registro de **Ciclismo em 30/09/2026, 10 km**. Se ainda não existir, cadastrá-lo conscientemente; conferir em `Início` que somente os desafios de Ciclismo com janela que inclua setembro aumentaram em 10 km, Outubro Rosa não aumentou e Caminhada/Corrida permanece intacta. Conferir `Setembro/2026` no histórico e se o card de 45 dias mantém 01/09–15/10. Comparar variação antes/depois em vez de assumir totais absolutos, pois há registros antigos de teste.

## 7. Próxima sequência e restrições

1. **Prioridade atual:** concluir teste funcional multi-mês do atleta 1603 acima, incluindo recálculo do resumo e conferência visual dos cards.
2. Se divergir: ler `REGISTRO_KM`, vínculo correspondente de `dgmbDesafios` e `MEU_GIRO_RESUMO`, verificar modalidade, janela individual, data local e atualização do resumo. Não alterar dados reais nem limpar o resumo indiscriminadamente.
3. Testar separadamente troca de foco entre desafios de mesma modalidade, inclusive sobreposição de janelas e edição/exclusão após o registro de homologação, com cuidado para não deixar dados duplicados.
4. Somente após concluir a lógica de registros, retomar o que foi priorizado para Portal Giro (seleção consciente da modalidade na inscrição pública, UX de vagas e demais pendências em seu próprio repositório); não inferir que implementações DEV estão no PROD.
5. **Ranking fora de escopo por decisão expressa do organizador.** Não implementar nem modificar comparações entre atletas nesta fase.
6. STRAVA, entregas de medalhas, finanças, avatares, operações Admin e Landing Page pertencem a missões paralelas: preservar contratos, mas não acoplar mudanças dessas áreas à auditoria de KM atual.

## 8. Guardrails para futuras implementações

- Antes de qualquer ajuste: identificar DEV/LAB ou PROD, branch, versão do deployment e base de dados usada.
- Reexecutar testes de modalidade, datas, janelas multi-mês, registros históricos, responsividade e iframe após mudar serviços comuns.
- Não substituir o dropdown aprovado por botões, nem ocultar seleção obrigatória para atleta multimodal.
- Não reinserir corte do mês no prazo individual.
- Não duplicar atividades por inscrição; progredir cada meta usando data + modalidade + janela.
- Não avançar alterações visuais já homologadas enquanto pendências lógicas estiverem em teste.
- Não declarar homologação funcional por apenas um teste automatizado; distinguir teste unitário, deploy e confirmação na interface.


## 9. Homologação visual do atleta 1603 — 14/10/2026 (após checkpoint)

**Evidência:** quatro capturas do Meu Giro DEV/LAB enviadas pelo organizador no chat; não foram utilizadas leituras nem gravações diretas em planilhas nesta conferência.

- Histórico **Setembro/2026**: Ciclismo **12 km** (dia 30: **10 km** + **2 km**); Caminhada/Corrida **44 km** (dia 29).
- Card **DESAFIO 1000 KM**, modalidade Ciclismo, início 01/09/2026, término 15/10/2026, prazo 45 dias: **18 km realizados**, 982 km faltantes e 1,8% concluído. A soma corresponde a 12 km em setembro + 6 km em outubro.
- Card **MEDALHA OUTUBRO ROSA**, Ciclismo, 01/10/2026 a 31/10/2026, meta 50 km: **6 km realizados**, 44 km faltantes e 12% concluído. Os 12 km do Ciclismo de setembro não entraram nessa meta.
- Card **LAB TESTE CAMINHADA/CORRIDA (SEM MEDALHA)**, outubro, meta 50 km: **5,5 km realizados**, 44,5 km faltantes e 11% concluído. Os 44 km registrados em setembro ficaram no histórico e não entraram na meta mensal de outubro. Tampouco houve transferência dos 6 km de Ciclismo de outubro para a meta de Caminhada/Corrida.
- Formulário com Modalidade | Data da atividade | KM realizados permanece visualmente alinhado e compacto conforme homologação anterior.

**Conclusão:** aprovado na interface LAB o compartilhamento por modalidade e por período, com PRAZO_DIAS atravessando setembro/outubro e separação entre atividades dos esportes. Isto não constitui teste de regressão PROD.

**Próxima missão lógica:** verificar atualização em ambos os desafios de Ciclismo quando UMA atividade de outubro é editada, com restauração posterior. Base antes da edição: 1000 KM **18 km**, Outubro Rosa **6 km**, Caminhada/Corrida de outubro **5,5 km**. Preferir um registro de Ciclismo de **3 km em outubro**, alterar **somente KM para 4 km**, e esperar **19 / 7 / 5,5 km** nos três cards. Restaurar **4→3 km** no mesmo registro e esperar **18 / 6 / 5,5 km** novamente. Não tocar nas atividades de setembro ou excluir registros nesta fase. Se a base inicial mudar por outros registros, comparar a diferença +1/-1, sem fixar totais absolutos.


## 10. Teste de edição compartilhada — etapa 1 (14/10/2026)

Quatro prints do Meu Giro DEV/LAB enviados pelo organizador **após editar 3→4 km uma atividade de Ciclismo de outubro**:

| Desafio no Início | Antes | Depois | Verificação |
|---|---:|---:|---|
| DESAFIO 1000 KM, Ciclismo, 45 dias | 18 km | 19 km | +1 km correto; 981 km faltantes; 1,9% |
| MEDALHA OUTUBRO ROSA, Ciclismo | 6 km | 7 km | +1 km correto; 43 km faltantes; 14% |
| LAB TESTE CAMINHADA/CORRIDA (SEM MEDALHA), 50 km | 5,5 km | 5,5 km | não afetado |
| LAB TESTE CAMINHADA/CORRIDA 2, 75 km | 5,5 km | 5,5 km | não afetado; mesma atividade de Caminhada/Corrida elegível compartilha progresso em duas metas |

**Homologado:** edição de uma atividade de Ciclismo atualizou os dois desafios compatíveis, mantendo inalteradas as duas metas de Caminhada/Corrida. **Não homologado ainda:** reversão de 4→3 km e confirmação de integridade do histórico (sem duplicatas). Issue [#294](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/issues/294) permanece aberta. Solicitar restauração do mesmo registro a 3 km e capturas ou confirmação dos quatro totais originais 18 / 6 / 5,5 / 5,5. Nenhum deploy, banco ou PROD alterado pela documentação.

## 11. Regressão de restauração 4→3 km — diagnóstico LAB (14/10/2026)

Na missão #294, após editar uma atividade de Ciclismo de 3→4 km, o usuário tentou restaurá-la para 3 km. O popup exibiu "Já existe uma atividade com esta mesma data e KM", sem aplicar a restauração.

**Causa raiz confirmada por leitura exclusiva da base LAB e do código:**
- A execução DEV do Meu Giro resolve sua planilha pela função `dgmbMeuGiroSpreadsheetId_` em `Meu Giro/EnvironmentConfig.gs`. Para o Script ID DEV em uso, a base atual é `1ivfrue_toKUiopYq1Xlx4UR_yFrKHFpyIgVn19xqEUA` (Portal Giro HOMOLOGACAO - REFRESH 2026-10-08 1159). A planilha `1sFxmi...`, presente no README histórico, é um fallback legado, **não a base atualmente selecionada pelo DEV**.
- Para o atleta sintético 1603, a aba `REGISTRO_KM` continha **duas atividades independentes de Ciclismo em 01/10/2026**, com `activity_id` diferentes e KM 4 e 3. A reversão legítima geraria 3 + 3 km, mantendo duas atividades.
- A interface `salvarEdicaoAtividade` já tratava `DUPLICIDADE_EDICAO`: pedia confirmação em popup padrão e reenviava `force: true`. O backend `editarAtividade` **ignorava** esse parâmetro e retornava novamente `DUPLICIDADE_EDICAO`, gerando o erro final após a confirmação.

**Correção DEV/LAB:** no `RegistroService.gs`, respeitar `payload.force === true` somente se a edição for identificada por **um `activity_id` existente e único**. Confirmar não cria linhas, altera somente a atividade identificada e atualiza o progresso normalmente. Sem confirmação, com IDs inconsistentes, com múltiplas linhas ou em registros legados sem `activity_id`, a proteção de duplicidade permanece. A interface/popup homologados não foram redesenhados.

**Teste automatizado acrescentado:** `tests/edicao_duplicidade_confirmada.test.js` executa o `editarAtividade` verdadeiro com planilha simulada, verificando cinco cenários: sem confirmação bloqueia, com confirmação edita somente a linha correta, ID incorreto bloqueia, legado sem ID bloqueia, confirmação não booleana bloqueia. Todos PASS em simulação sem escrita real.

**Homologação real ainda pendente:** após sincronizar e publicar o DEV, o organizador deve repetir 4→3 km no mesmo registro de Ciclismo de 01/10. O popup de atividade semelhante é esperado; confirmar consciente e conferir 1000 KM 18 km, Outubro Rosa 6 km, os dois desafios Caminhada/Corrida 5,5 km cada, e histórico de outubro com **duas atividades de Ciclismo de 3 km**, totalizando 6 km. Não apagar nem criar atividades adicionais. Ranking/PROD fora do escopo.

## 12. Homologação da restauração 4→3 km e encerramento #294

**Evidência:** captura enviada pelo organizador após edição com confirmação no DEV/LAB; mostra em Outubro/2026 **Dia 1 — 3 km — Ciclismo** duas vezes, mais **Dia 12 — 5,5 km — Caminhada/Corrida**.

**Base efetiva LAB, conferência somente leitura:**
- `REGISTRO_KM`: duas linhas 01/10 Ciclismo 3 km com `activity_id` independentes e preservados; 7 linhas do atleta 1603, sem registro adicional.
- `MEU_GIRO_RESUMO`: desafio 1000 KM **18 km**; Outubro Rosa **6 km**; Caminhada/Corrida meta 50 **5,5 km**; Caminhada/Corrida meta 75 **5,5 km**.
- A edição 3→4→3 foi concluída e calculada sem mesclar nem duplicar atividades. Correção da PR #296 homologada. [Issue #294](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/issues/294) **encerrada com sucesso**.

## 13. Auditoria de exclusão individual — próxima missão #297

Após a restauração, a inspeção de `excluirAtividade` identificou um risco anterior à homologação destrutiva: `localizarLinhasAtividade_` tenta `chave_edicao`/timestamp quando um `activity_id` fornecido não é localizado. Em exclusões isso pode apontar para registro errado ou múltiplas linhas.

**Proteção DEV/LAB:** exigir que o `activity_id` presente identifique exatamente uma linha. Para registros legados sem ID, permitir timestamp somente quando corresponder a uma única linha. Se houver ID inconsistente ou ambiguidade, bloquear com mensagem de erro. Não mudar a rotina de edição nem o layout. Não descartar a confirmação de exclusão já presente no frontend.

**Testes em memória:** `tests/exclusao_id_atividade_seguro.test.js` usa `excluirAtividade` e localizador reais, com planilha em memória. Sete casos PASS: exclusão única, dois registros iguais com IDs distintos, ID inexistente sem fallback perigoso, legado inequívoco, legado ambíguo, modalidades diferentes no mesmo timestamp e rollback da linha se falhar a atualização do resumo. Nenhum dado LAB/PROD foi alterado nesses testes.

**Plano de homologação #297 (somente após sucesso do workflow DEV):** não excluir os dois registros originais de 3 km. Na conta sintética 1603, criar atividade temporária de **Ciclismo de 2 km em 02/10/2026**, desde que não haja registro nessa data. Confirmação de +2 km nos dois desafios Ciclismo: 1000 KM 18→20; Outubro Rosa 6→8. Ambos os desafios Caminhada/Corrida devem continuar 5,5. Excluir somente a atividade temporária e verificar retorno 18/6/5,5/5,5 e preservação das duas atividades de 3 km. Registrar evidência e encerrar issue #297 apenas após validação visual e do resumo, com PROD intocado.


## 14. Homologação de exclusão segura — 15/10/2026 (#297)

**Missão concluída e encerrada:** [GitHub issue #297](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/issues/297), implementação DEV pela [PR #298](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/pull/298).

**Evidência do participante de homologação 1603** — duas capturas de `Registrar > Outubro/2026`:
- Antes: Ciclismo **8 km**, Caminhada/Corrida **5,5 km**, atividade temporária **02/10 — 2 km — Ciclismo** presente; duas atividades **01/10 — 3 km — Ciclismo** presentes.
- Depois: Ciclismo **6 km**, Caminhada/Corrida **5,5 km**; atividade temporária de 02/10 não aparece mais; **as duas atividades de 3 km continuam**.
- Confirmação independente de **somente leitura na planilha LAB em uso**: `REGISTRO_KM` tem **7 linhas** para 1603, não contém a linha temporária, e mantém as duas atividades de 01/10 com `activity_id` distintos. `MEU_GIRO_RESUMO` mantém **1000 KM 18 km**, **Outubro Rosa 6 km**, **Caminhada/Corrida 50 km 5,5 km** e **75 km 5,5 km**.
- Guardas implementadas e cobertas por sete testes em memória: ID fornecido deve identificar exatamente uma linha; ID incorreto não faz fallback destrutivo por timestamp; legado ambíguo é bloqueado; rollback previsto se falhar recálculo. A confirmação de exclusão em popup permanece.
- Não houve gravação administrativa na planilha de teste nem publicação de código PROD pelo processo desta auditoria.

**Próxima fronteira lógica sugerida:** verificar as regras de encerramento da janela individual `PRAZO_DIAS`, em particular **último dia incluído (15/10/2026)**, **dia seguinte excluído do progresso (16/10/2026)**, permanência no histórico pessoal, modalidades independentes e eventual atualização do status. A data futura nunca pode ser registrada antes de ocorrer no Brasil; usar casos simulados para o dia seguinte enquanto ainda for futuro. Ranking permanece fora de escopo.


## 15. Missão #300 — auditoria técnica autônoma do último dia (15/10/2026)

**Referência:** [Issue #300](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/issues/300), ainda **aberta** até homologação funcional. Data/hora local aferida em 15/10/2026, 07:45 (UTC-03).

**Base de dados DEV/LAB examinada em leitura exclusiva, sem gravação:**
- Atleta sintético 1603, `REGISTRO_KM`: 7 atividades; nenhuma em 15/10 nem 16/10 no instante da consulta. Baseline Ciclismo = 12 km de setembro + 6 km outubro; Caminhada/Corrida = 44 km setembro e 5,5 km outubro.
- `MEU_GIRO_RESUMO`: desafio 1000 KM **18 km**, Outubro Rosa Ciclismo **6 km**, metas de Caminhada/Corrida 50 km e 75 km **5,5 km** cada.
- Em `dgmbDesafios`, inscrição de 1000 KM com `Prazo_Dias=45`, `data_inicio_desafio=2026-09-01`, `data_fim_desafio=2026-10-15`; `Status_Usuario_Desafio=EM_ANDAMENTO`. Manter status administrativo original até conferência/validação oficial.

**Regressões automatizadas:** arquivo `tests/prazo_ultimo_dia_historico_1603.test.js` executado em memória com **7 PASS**, utilizando as funções reais de `ModalidadeEsportivaService.gs`, `atividadeDentroPeriodoOficial_` de `Utils.gs` e `obterTempoRestanteDesafioPainel_` de `Script.html` com relógio simulado. Confirma:
1. 14/10: 15 e 16 são datas futuras e são bloqueadas; relógio UI mostra `1 dia`.
2. 15/10: +1 km Ciclismo nesse dia vai de 18→19 no 1000 KM, de 6→7 no Outubro Rosa, e não toca nas metas de Caminhada/Corrida; UI indica `Último dia`.
3. 16/10: registros desse dia são aceitos no histórico, entram em Outubro Rosa mas não no 1000 KM; UI indica `Encerrado`.
4. Novembro: atividades históricas válidas permanecem no total de 45 dias; atividade posterior não entra em nenhum dos desafios encerrados em outubro.
5. Modalidade diferente nunca contamina os km do Ciclismo, e exclusão simulada do km temporário restaura a base; datas inválidas são recusadas.

**Atenção: classificação do desafio após o prazo.** `obterTempoRestanteDesafioPainel_` informa `Encerrado` após 15/10, mas `isDesafioAtivo_` e `classificarDesafioV2_` usam principalmente os status de origem, e não a data final. Enquanto a inscrição 1603 continuar `EM_ANDAMENTO` e o resumo `STATUS_EM_ANALISE`, o card poderá ficar em "Seus desafios em andamento" com "Encerrado". **Não transformar automaticamente a situação administrativa em CONCLUIDO, EXPIRADO ou NAO_CONCLUIDO**; o tratamento visual e o fluxo de apuração de encerramento precisam ser definidos e homologados separadamente.

**Pendência operacional da #300:** simulação em memória não equivale a chamada autenticada ao Web App. Não houve registro/exclusão direta na planilha, para não contornar o código real ou alterar a base de homologação de modo artificial. A verificação end-to-end de registrar/excluir 1 km em 15/10 (ou test harness integrado usando a execução real do GAS DEV) permanece pendente. Depois da virada para 16/10 (Brasil), conferir novamente somente leitura o comportamento do limite e o status visual; eventual registro de 16/10 precisa respeitar a data efetiva, sem retroagir o relógio em produção.


## 16. Teste de integração do registro no corte de 15/16 outubro — 15/10/2026

**Avanço da missão [#300](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/issues/300).** Criado `tests/registro_servico_real_prazo_1603.test.js`: executa a implementação real de `registrarAtividade`, `atualizarDistanciaRealizada_` e a classificação do vínculo por modalidade/data do código GAS. As chamadas a `LockService`, `Utilities`, `getSpreadsheet_`, `MEU_GIRO_RESUMO` e às funções de acesso a planilhas são simuladas **em memória**, sem sessões autenticadas nem gravações reais. O resumo simulado aplica as mesmas regras de data/modalidade, mas **não** executa a persistência completa do serviço de resumo em planilha.

**Oito cenários passaram em teste isolado**:
1. Hoje 15/10: registrar **1 km Ciclismo em 15/10** ⇒ 1000 KM **18→19**; Outubro Rosa **6→7**; Caminhada/Corrida **5,5** nas duas metas.
2. Hoje 15/10: tentar **16/10** ⇒ `DATA_FUTURA`, nenhum registro nem recálculo.
3. Hoje 16/10: registrar **1 km Ciclismo em 16/10** ⇒ 1000 KM mantém **18**, Outubro Rosa vai **6→7**; histórico recebe o registro.
4. Hoje 16/10: registrar retroativamente em **15/10** ⇒ 1000 KM **19**, Outubro Rosa **7**.
5. Atleta multimodal sem seleção explícita ⇒ `MODALIDADE_OBRIGATORIA`, sem gravar.
6. Mesmo atleta/data/modalidade/KM já existentes ⇒ `DUPLICIDADE`, sem gravar.
7. Registrar 1 km Caminhada/Corrida em 15/10 ⇒ ambas metas dessa modalidade **5,5→6,5**, Ciclismo inalterado.
8. Atividade de Ciclismo anterior à janela do 1000 KM ⇒ permanece no histórico, sem somar às metas.

**Resultado:** camada de registro e recálculo exercitada com serviço real + mocks, sem divergências. Não declarar homologação da interface e persistência GAS/Sheets somente por esse teste. A base LAB permanece intacta. Auditoria automática em modo somente leitura já programada para 16/10.

**Risco separado:** após a data final, `obterTempoRestanteDesafioPainel_` sinaliza "Encerrado", porém `isDesafioAtivo_` e `classificarDesafioV2_` seguem status administrativos, podendo manter o card em "Em andamento". Aberta a [issue #302](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/issues/302) para tratamento visual seguro **sem mutação de status na planilha**. Não alterar o layout aprovado, inscrições, Ranking ou PROD inadvertidamente.

**Outros limites não cobertos:** falha intermediária no recálculo pode deixar `dgmbDesafios` parcialmente sincronizada antes de um rollback da linha de `REGISTRO_KM`; isso exige uma auditoria transacional independente, não faz parte da regressão de datas.
