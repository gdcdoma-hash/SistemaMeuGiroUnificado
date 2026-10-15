# Preparo isolado — Meu Giro — 15/10/2026

**Status: NAO PUBLICAR.** Branch derivada de `release/staging-meugiro-modalidades-2026-10-15`. Arquivo `.clasp.json` usa Script ID invalido propositalmente para evitar sincronizacao com GAS DEV atual. O codigo de atividades/certificado nao foi modificado nesta preparacao.

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
