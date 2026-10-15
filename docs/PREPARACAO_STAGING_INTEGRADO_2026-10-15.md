# Preparo isolado — Meu Giro — 15/10/2026

**Status: projeto Apps Script STAGING NOVO criado, com codigo enviado, mas SEM DEPLOYMENT WEB.** As linhas de preparacao sao independentes dos fluxos normais de DEV, LAB e PROD. As propriedades ainda precisam ser ativadas.

## Destino obrigatorio
- Criar projeto Google Apps Script **Meu Giro — STAGING INTEGRADO**, diferente do Meu Giro DEV e PROD.
- Ambos os projetos de staging devem usar a **mesma copia restrita da planilha LAB de 08/10/2026**, indicada no controle interno do Portal (GitHub privado, issue #398). **Nao publicar ID da planilha ou dados de atleta neste repositorio publico.**
- No projeto novo configurar explicitamente `DGMB_ENVIRONMENT=STAGING`, `DGMB_SPREADSHEET_ID` (copia de staging), `DGMB_PORTAL_WEBAPP_URL` (deployment Portal staging). Se teste envolver certificados, `DGMB_CERTIFICADOS_FOLDER_ID` e `DGMB_CERTIFICADO_TEMPLATE_ID` devem ser apenas de preparacao; nunca reutilizar documentos ou pastas PROD.
- `EnvironmentConfig.gs` possui valores de fallback para ambientes de desenvolvimento: **nao iniciar o staging antes de validar todas as Script Properties efetivas**. Nao executar operacoes se valores estiverem faltando ou apontando para DEV/PROD.
- Atualizar o Script ID placeholder neste ramo **somente depois** de provisionar projeto GAS independente. Nenhum workflow existente de `main` ou de PROD deve ser acionado com este ramo; podem sobrescrever destinos.
- Sem abertura publica; base de staging contem copia de dados operacionais LAB. Restringir acesso.

## Homologacao minima
- Confirmar rota `page=atleta&handoff=...`, consumo unico do token e sessao associada ao ID autenticado.
- Testar retorno ao Portal de staging sem reiniciar login.
- Registro Ciclismo/Caminhada-Corrida: escolha obrigatoria, km dentro do prazo, registro entre meses, nenhuma mistura entre modalidades.
- Certificado apenas se necessario para regressao: regra oficial **KM 100% + status CONCLUIDO atualizado manualmente pelo Admin**.
- Ranking e melhorias adicionais fora do escopo.
- Proibido deployment PROD ate homologacao integrada aprovada e autorizacao especifica.

**Controle interno:** issue #398 do repositorio Portal Giro.

## Atualizacao de infraestrutura (15/10/2026) — prevalece sobre placeholders acima

- Projeto Apps Script STAGING exclusivamente novo criado: `https://script.google.com/d/1Qkn_ko5Wv6fNwvVfC9lIQOMWZdncArqdWntHtspBEVyp8X1CC9gSsGs4/edit`.
- `.clasp.json` alterado somente neste ramo para o ID do projeto novo; o placeholder propositalmente invalido foi substituido com seguranca.
- Codigo enviado para esse projeto via [workflow restrito de preparacao](https://github.com/gdcdoma-hash/SistemaMeuGiroUnificado/actions/runs/38790299722), **sem deployment**.
- Trava adicionada a `EnvironmentConfig.gs`: rejeita Script ID incorreto e impede fallback de propriedades antigas.
- Funcao disponivel no novo projeto: `STAGING_CONFIGURAR_MEU_GIRO_20261015()`, criada no arquivo `Meu Giro/StagingBootstrapService.gs`. Executar **uma vez no editor Apps Script novo**, autorizando Google Drive/Sheets. Ela valida projeto, planilha e pasta e grava somente Script Properties do staging, sem alterar dados das planilhas.
- `DGMB_MEU_GIRO_WEBAPP_URL` e `DGMB_PORTAL_WEBAPP_URL` **nao podem ser apontadas ate existir o par de deployments WEB stagings**; ainda bloqueado.
- Sem acionar workflows antigos, sem promover PROD, sem transacao bancaria.
