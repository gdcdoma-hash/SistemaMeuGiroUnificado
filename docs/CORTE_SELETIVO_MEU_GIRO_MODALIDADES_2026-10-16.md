# Pacote seletivo Meu Giro — modalidades e handoff, revisão 16/10/2026

> **Somente preparação.** NÃO sincronizar, mesclar na `main`, publicar, editar Script Properties ou migrar banco antes da autorização expressa. Controle do corte [Portal #398](https://github.com/gdcdoma-hash/dgmb-inscricoes-repescagem/issues/398).

**Base protegida:** commit `1df3d6431838f9f7399dc6c9401fa4ca18f0b9d3`, o commit usado no último corte Meu Giro PROD @8 de 04/10. A branch é separada da `main` que auto-sincroniza DEV.

**Oito arquivos operacionais copiados da release homologada:**
- `Meu Giro/AuthSession.html`
- `Meu Giro/Index.html`
- `Meu Giro/ModalidadeEsportivaService.gs`
- `Meu Giro/PainelService.gs`
- `Meu Giro/RegistroService.gs`
- `Meu Giro/Script.html`
- `Meu Giro/Styles.html`
- `Meu Giro/Utils.gs`

**Excluídos deliberadamente:**
- `Meu Giro/RankingService.gs`: Ranking não faz parte deste corte;
- `Meu Giro/CertificadoService.gs`: mudanças de certificado e CONCLUIDO manual não fazem parte da missão das modalidades;
- `Meu Giro/EnvironmentConfig.gs` e `Meu Giro/appsscript.json`: preservados do último corte PROD. A versão `main` possui um tratamento específico de DEV na configuração, que não deve ser transportado sem revisão;
- `tests/`, documentos e outras mudanças da `main` posteriores a 04/10, que **não foram aprovados neste pacote**.

**Segurança de empacotamento:** `.clasp.json` preservado da base apontando para a raiz `Meu Giro` do repositório e **não para o PROD**. O futuro corte exigirá apontar explicitamente ao Script ID PROD, depois de autorizar e verificar destino; `clasp show-file-status` é bloqueio obrigatório pré-push. O workflow atual `sync-meu-giro-prod.yml` envia HEAD mas não atualiza deployment público. É necessário realizar versão e atualização separadas do deployment já existente, somente após autorização formal.

**Notas de compatibilidade:** o pacote não inclui as correções independentes de Ranking nem de Certificados. Antes do corte, rodar testes de modalidades/handoff/prazo e smoke test conservador de certificados/Ranking com o código existente. Se houver dependência comprovada, documentar e obter aprovação do acréscimo — não incluir automaticamente.

**Modelo de publicação:** duas fases, com manutenção breve e backup da planilha oficial: atualizar Meu Giro PROD de forma seletiva e seu deployment existente; depois atualizar Portal PROD com o pacote seletivo correspondente. Não reabrir inscrições nem registros antes dos testes de ambos. Ver [plano de checkpoint/rollback](https://github.com/gdcdoma-hash/dgmb-inscricoes-repescagem/blob/prep/corte-seletivo-modalidades-2026-10-16/docs/CHECKPOINT_E_REVERSAO_MODALIDADES_2026-10-16.md).

**Verificação local antes de construir branch:** os quatro arquivos GS candidatos passaram em parse JavaScript; isto não comprova execução do GAS nem regressão completa.
