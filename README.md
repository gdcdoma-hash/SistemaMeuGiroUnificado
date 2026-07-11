# Sistema Meu Giro Unificado

O código operacional ativo deste repositório está na pasta `Meu Giro/`.

A auditoria do fluxo de inscrição está documentada em `docs/auditoria-fluxo-inscricao.md`. Neste repositório não foi localizado o fluxo atual que cria nova inscrição, repescagem ou reativação em `dgmbDesafios`; a sincronização automática após nova inscrição deve ser aplicada no repositório que contém esse fluxo (por exemplo, o Portal Giro), caso ele seja externo.

A pasta histórica `InscriçãoDesafio/` foi removida desta branch para evitar que arquivos legados de cadastro/login participem do deploy ou recebam correções por engano. O histórico permanece preservado no Git.
