# Controle de Estoque (Google Apps Script)

## Arquitetura
- `Código.gs`: fluxo antigo da planilha (menus, diálogos, relatórios). `Dialog*.html`, `SidebarCores.html`.
- `WebAppFunctions.gs`: backend do WebApp (lançamentos, índice, bloqueio). `WebApp.html`: frontend.
- `Fornecedores.gs`: fornecedores. `OTIMIZAÇÕES.md`: histórico de otimizações.
- Arquivos `.gs` compartilham escopo global (constantes de um arquivo valem nos outros).

## Abas da planilha
- `ESTOQUE`: A Grupo, B Item, C Unidade, D Data, E NF, F Obs, G Saldo Anterior, H Entrada, I Saída, J Saldo, K Valor Unit., L Alterado Em, M Alterado Por. Mais de 40 mil linhas, cronológicas.
- `ÍNDICE_ITENS` (um registro por item, evita ler o `ESTOQUE`): A Item, B Saldo, C Última Data, D Grupo, E Linha ESTOQUE, F carimbo, G Últ. ATUALIZAÇÃO, H Últ. Entrada (data, ou `NAO` se só teve saídas; vazio ou `SIM` = desconhecido).
- Grupos vistos: MATÉRIA PRIMA (fios e alguns materiais), QUIMICO, PEÇA PARA MÁQUINA, EMBALAGEM, MANUTENÇÃO.

## Regras de negócio (valores em constantes no topo de `WebAppFunctions.gs`)
- **Bloqueio de saída** (`_bloqueioItem`): só SAÍDA, só grupo em `GRUPOS_BLOQUEIA` (MATÉRIA PRIMA), sem ATUALIZAÇÃO há mais de `DIAS_VALIDADE_ATUALIZACAO` (20).
  - Libera: obs `ATUALIZAÇÃO`, `ACERTO` ou `ATUAL. SAIDA`; item novo; item com coluna H = `NAO`; última entrada em até 20 dias (ACERTO com entrada conta).
  - Coluna H vazia ou `SIM` = desconhecido, continua bloqueando. Erro na regra libera o lançamento e grava no log.
  - Ligada pela propriedade `BLOQUEIO_ATUALIZACAO_ATIVO` (criada por `reconstruirIndiceCompleto()`); apagar a propriedade desliga.
- **Sinalização** (linha vermelha + aviso "DESATUALIZADO"): `DIAS_SINALIZACAO_DESATUALIZADO` (15). Entrada pinta de amarelo. Independe do bloqueio.
- Relatório "Cores Desatualizadas": 15 dias. "Consulta de atualização" do menu antigo (`Código.gs`): 20 dias.

## Cuidados técnicos
- Cache do índice (`CacheService`) tem limite de 100 KB: não acrescentar campos ao objeto do índice em cache.
- Códigos como `6300/1` viram data/número no Sheets: coluna Item em texto e leitura com `getDisplayValues()`.
- Evitar ler o `ESTOQUE` inteiro por item (tempo de execução do Apps Script, ~6 min). Usar o índice e memo por execução.
- Mudou colunas do `ÍNDICE_ITENS`: rodar `reconstruirIndiceCompleto()` uma vez após publicar.

## Fluxo de trabalho
- Responder em português, direto. Para alteração: entendimento, plano, riscos e casos extremos, perguntar; depois código só dos blocos alterados.
- Validar: copiar o `.gs` para `.js` no scratchpad e rodar `node --check`; testar a regra com simulação (mock) quando possível.
- Commit e push só na branch designada da sessão. Não abrir PR sem pedido.

## Pendências conhecidas
- Aba `Credenciais` guarda senhas em texto puro; migrar para hash ou `PropertiesService`.
- Materiais que não são fio dentro de MATÉRIA PRIMA (ex.: `ACETATO TRANSP 40MM`) continuam bloqueados; solução prevista: aba de exceções por item.
