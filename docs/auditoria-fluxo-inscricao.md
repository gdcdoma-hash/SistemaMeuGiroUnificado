# Auditoria do fluxo de inscrição e sincronização do MEU_GIRO_RESUMO

## Conclusão

A auditoria deste repositório não encontrou o fluxo atual que cria novas inscrições em `dgmbDesafios`.

O código operacional presente aqui está em `Meu Giro/`. Ele lê `dgmbDesafios`, atualiza campos derivados como `distancia_realizada` e gera/atualiza `MEU_GIRO_RESUMO`, mas não contém função ativa de criação de inscrição, repescagem ou reativação.

Portanto, a sincronização automática após nova inscrição não deve ser inventada neste repositório. A correção funcional deve ser aplicada no repositório que contém o fluxo atual de inscrição (por exemplo, o Portal Giro), imediatamente após a gravação principal em `dgmbDesafios` concluir com sucesso.

## Localização do código atual

- Código operacional deste repositório: `Meu Giro/`.
- Fluxo atual de criação de inscrição em `dgmbDesafios`: não localizado neste repositório após busca por `appendRow`, `setValues`, `dgmbDesafios`, `ID_INSCRICAO`, `Status_Usuario_Desafio`, `gravarInscricao`, `nova inscrição`, `repescagem`, `reativar` e `adminAtualizarStatus`.
- Pasta legada removida: `InscriçãoDesafio/`. O histórico permanece no Git.

## Pontos encontrados

| Termo / ponto | Arquivo | Função | Dado gravado | ID_DGMB disponível | Altera inscrição? | Deve chamar `atualizarMeuGiroResumo_`? |
| --- | --- | --- | --- | --- | --- | --- |
| `appendRow` | `Meu Giro/Utils.gs` | `atualizarMeuGiroResumo_` | Nova linha na `MEU_GIRO_RESUMO` | Sim, parâmetro `idDgmb` normalizado em `id` | Não; grava resumo derivado | Não; é a própria função de resumo |
| `setValues` | `Meu Giro/RegistroService.gs` | `registrarAtividade` | Nova linha em `REGISTRO_KM` | Sim, `idDgmb` | Não; registra atividade | Já chama após escrita e atualização de distância |
| `setValues` | `Meu Giro/RegistroService.gs` | `excluirAtividade` | Rollback em `REGISTRO_KM` dentro do `catch` | Sim, `idDgmb` | Não; desfaz exclusão de atividade | Não no rollback; a sincronização ocorre antes do `catch` |
| `setValue` em `dgmbDesafios` | `Meu Giro/RegistroService.gs` | `atualizarDistanciaRealizada_` | `distancia_realizada` em `dgmbDesafios` | Sim, `idDgmb` | Não cria inscrição; atualiza total derivado | É chamada antes de `atualizarMeuGiroResumo_` nos fluxos de atividade |
| `dgmbDesafios` | `Meu Giro/Utils.gs` | `obterDgmbDesafiosCacheExecucao_` | Leitura/cache da aba | Não grava; depende do chamador | Não | Não |
| `dgmbDesafios` | `Meu Giro/Utils.gs` | `obterVinculosDesafioUsuario_` | Leitura dos vínculos por usuário | Sim, parâmetro `idDgmb` | Não | Não; alimenta `atualizarMeuGiroResumo_` |
| `ID_INSCRICAO` | `Meu Giro/Utils.gs` | `meuGiroResumoHeaders_` | Cabeçalho da `MEU_GIRO_RESUMO` | Não se aplica | Não | Não |
| `setValues` | `Meu Giro/Utils.gs` | `ensureMeuGiroResumoSheet_` | Cabeçalho da `MEU_GIRO_RESUMO` | Não se aplica | Não | Não |
| `setValues` / `appendRow` | `Meu Giro/Utils.gs` | `atualizarMeuGiroResumo_` | Atualização/criação de linhas da `MEU_GIRO_RESUMO` | Sim, parâmetro `idDgmb` | Não; materializa resumo | Não; é a própria sincronização |
| `dgmbDesafios` | `Meu Giro/AdminCertificadoService.gs` | `listarPendenciasValidacaoCertificado` | Leitura de pendências | Sim, `adminIdDgmb` para acesso admin; lê atletas | Não | Não |
| `dgmbDesafios` / status admin | `Meu Giro/AdminCertificadoService.gs` | `atualizarStatusValidacaoCertificadoAdmin` | `status_validacao_certificado`, data/observação de validação | Sim, `idDgmb` | Não cria inscrição; altera validação de certificado | Não, sem evidência de impacto no resumo |
| `dgmbDesafios` | `Meu Giro/CertificadoService.gs` | `certificadoBuscarContextoDesafio_` | Leitura de contexto do certificado | Sim, `idDgmb` | Não | Não |
| `Status_Usuario_Desafio` | `Meu Giro/AuditoriaStatusUsuarioDesafio.gs` | `auditarStatusUsuarioDesafio` | Relatório/log de auditoria | Não cria inscrição | Não | Não |
| `gravarInscricao` | — | — | Nenhuma ocorrência no código ativo | — | — | Correção pertence ao repositório externo que possuir o fluxo |
| `nova inscrição` | — | — | Nenhuma ocorrência no código ativo | — | — | Correção pertence ao repositório externo que possuir o fluxo |
| `repescagem` | — | — | Nenhuma ocorrência no código ativo | — | — | Correção pertence ao repositório externo que possuir o fluxo |
| `reativar` | — | — | Nenhuma ocorrência no código ativo | — | — | Correção pertence ao repositório externo que possuir o fluxo |
| `adminAtualizarStatus` | — | — | Nenhuma ocorrência no código ativo | — | — | Não se aplica |

## Ação necessária no repositório do fluxo de inscrição

No repositório que contém a função real de nova inscrição/repescagem/reativação, aplicar a sincronização somente depois da gravação principal em `dgmbDesafios` ter sucesso:

```javascript
try {
  atualizarMeuGiroResumo_(idDgmb);
} catch (erroResumo) {
  Logger.log(
    '[MEU_GIRO_RESUMO][ERRO_SINCRONIZACAO] origem=NOME_DA_FUNCAO id_dgmb=' +
    idDgmb +
    ' erro=' +
    (erroResumo && erroResumo.message ? erroResumo.message : String(erroResumo))
  );
}
```

Este repositório não deve receber serviço paralelo nem função fictícia de inscrição apenas para cobrir essa ausência.
