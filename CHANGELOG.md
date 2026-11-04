# Changelog

Uma entrada por alteração, da mais recente para a mais antiga. Cada entrada traz o que mudou, o impacto e a ação manual necessária (se houver).

## 2026-11-04 — Sinalização de desatualizado em 15 dias
- **Mudou:** linha vermelha e aviso "DESATUALIZADO" passam a disparar com mais de 15 dias (antes 20). Nova constante `DIAS_SINALIZACAO_DESATUALIZADO` em `WebAppFunctions.gs`.
- **Não mudou:** bloqueio de saída segue em 20 dias (`DIAS_VALIDADE_ATUALIZACAO`).
- **Arquivos:** `WebAppFunctions.gs`, `Código.gs`, `WebApp.html`.
- **Ação manual:** publicar no Apps Script.

## 2026-11-04 — Não bloqueia item com última entrada em até 20 dias
- **Mudou:** coluna H de `ÍNDICE_ITENS` passa a guardar a data da última entrada (`NAO` = só teve saídas). Saída de item com entrada nos últimos 20 dias não é bloqueada. ACERTO com entrada conta como entrada.
- **Arquivos:** `WebAppFunctions.gs`.
- **Ação manual:** rodar `reconstruirIndiceCompleto()` de novo; até lá, coluna H vazia ou `SIM` continua bloqueando.

## 2026-11-04 — Libera saída de item que nunca teve entrada
- **Mudou:** item só com saídas (sem nenhuma entrada) não é bloqueado nem avisado. Coluna H "Teve Entrada" (`SIM`/`NAO`/vazio), depois substituída pela data da última entrada (entrada acima).
- **Arquivos:** `WebAppFunctions.gs`.
- **Ação manual:** `reconstruirIndiceCompleto()`.

## 2026-11-03 — Bloqueio restrito ao grupo MATÉRIA PRIMA
- **Mudou:** só itens com Grupo `MATÉRIA PRIMA` (coluna D de `ÍNDICE_ITENS`) têm a saída bloqueada por falta de ATUALIZAÇÃO. Demais grupos (químico, peça, embalagem, consumo) ficam apenas com a sinalização. Constante `GRUPOS_BLOQUEIA`.
- **Pendência:** materiais que não são fio dentro de MATÉRIA PRIMA (ex.: `ACETATO TRANSP 40MM`) continuam bloqueados.
- **Arquivos:** `WebAppFunctions.gs`.
- **Ação manual:** publicar no Apps Script.

## 2026-10-28 — Bloqueio por ATUALIZAÇÃO (histórico)
- PR #24 e anteriores: bloqueio de saída de item sem ATUALIZAÇÃO nos últimos 20 dias, restrito a saídas, com aviso na seleção do item e por linha no lote.
