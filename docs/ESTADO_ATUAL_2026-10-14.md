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
