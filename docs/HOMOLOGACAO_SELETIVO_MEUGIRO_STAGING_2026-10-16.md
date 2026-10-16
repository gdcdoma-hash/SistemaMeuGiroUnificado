# Homologação seletiva do Meu Giro — somente STAGING privado

Base: `prep/corte-seletivo-meugiro-modalidades-2026-10-16` commit `649fe3077a9eadaa4bd49f326aa4dc79ffe1fecc`.
Arquivos operacionais selecionados: oito, iguais por hash à release STAGING já testada.
Configurações privadas sobrepostas **somente nesta branch de teste**: `.clasp.json`, `.claspignore`, `Meu Giro/EnvironmentConfig.gs`, `Meu Giro/StagingBootstrapService.gs`, `Meu Giro/appsscript.json`, todas provenientes da branch STAGING original.

O serviço `Meu Giro/CertificadoService.gs` permanece **idêntico ao último corte PROD @8**, diferentemente do STAGING atual, a fim de testar não regressão. Ranking permanece idêntico ao corte e ao STAGING.

Destino único: Apps Script STAGING privado `1Qkn_ko5Wv6fNwvVfC9lIQOMWZdncArqdWntHtspBEVyp8X1CC9gSsGs4`, Web App STAGING privado existente, acesso `MYSELF`. Planilha STAGING isolada; o código de EnvironmentConfig bloqueia IDs, configurações e planilhas divergentes. NUNCA enviar esta branch ao PROD ou DEV.

Workflow criado com guardas de branch, token explícito e ID fixo de STAGING. O push e redeploy STAGING somente acontecerão quando houver commit de gatilho autorizado. O workflow não cria novo deployment nem altera PROD.
