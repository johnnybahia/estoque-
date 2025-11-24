// ==============================
// Code.gs
// ==============================

/**
 * doGet: Função principal do Web App
 * Serve a interface web quando alguém acessa a URL do Web App
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('WebApp');

  // Passa parâmetros da URL para o template, se houver
  template.params = e.parameter;

  return template.evaluate()
    .setTitle('Sistema de Gestão de Estoque')
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/drive_2020q4_48dp.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * include: Função helper para incluir arquivos HTML parciais
 * Permite modularizar o código HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Atualiza o menu principal e adiciona um menu separado para processar cores.
 */
function updateMenus() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GESTÃO DO ESTOQUE")
    .addItem("Inserir Estoque", "showEstoqueSidebar")
    .addItem("Inserir Grupo", "showGrupoDialog")
    .addSeparator()
    .addItem("Localizar Produto", "localizarProduto")
    .addItem("Mostrar Todos", "mostrarTodos")
    .addSeparator()
    .addItem("Gerar Relatório", "abrirDialogRelatorioEstoque")
    .addItem("Relatório por Grupo", "abrirDialogRelatorioPorGrupo")
    .addItem("Listagem de Estoque", "showListagemEstoqueSidebar")
    .addItem("Atualizar Compra de Fio e Histórico", "atualizarCompraDeFioEHistorico")
    .addSeparator()
    .addItem("Atualizar Total Embarcado", "atualizarTotalEmbarcado")
    .addItem("Alternar Restauração", "toggleRestore")
    .addItem("Apagar Última Linha", "apagarUltimaLinha")
    .addSeparator()
    .addItem("ÚLTIMA LINHA", "select10RowsBelow")
    .addSeparator()
    .addItem("Estoque por Período", "abrirDialogEstoquePorPeriodo")
    .addItem("Limpar Filtro", "limparFiltroEstoque")
    .addSeparator()
    .addItem("Estoque 3 Meses", "showEstoque3MesesSidebar")
    .addSeparator()
    .addItem("Cores Desatualizadas", "showCoresDesatualizadasDialog")
    .addToUi();
}

/**
 * onOpen: Executada quando a planilha é aberta.
 */
function onOpen() {
  PropertiesService.getUserProperties().deleteProperty("loggedUser");
  Logger.log("onOpen: Propriedade 'loggedUser' apagada.");
  
  backupEstoqueData();
  removeFilterOnOpen();
  showLoginDialog();
  // O updateMenus() só é chamado após login bem-sucedido
}

/* ... (demais funções já existentes no seu script, como backupEstoqueData, showEstoqueSidebar, etc.) ... */


/* ================================
   NOVAS FUNÇÕES: Processamento de Cores Desatualizadas via Sidebar
   ================================ */

/**
 * showCoresSidebar: Abre a sidebar que lista os valores (cores) da coluna E da aba "CORES DESATUALIZADAS".
 */
function showCoresSidebar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) {
    SpreadsheetApp.getUi().alert("A aba 'CORES DESATUALIZADAS' não foi encontrada.");
    return;
  }
  
  // Considera que os dados da coluna E começam na linha 2 (com cabeçalho na linha 1)
  var lastRow = sheetCores.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Não há dados na coluna E da aba CORES DESATUALIZADAS.");
    return;
  }
  var coresData = sheetCores.getRange(2, 5, lastRow - 1, 1).getValues();
  var cores = [];
  for (var i = 0; i < coresData.length; i++) {
    var val = coresData[i][0];
    if (val && cores.indexOf(val) === -1) {
      cores.push(val);
    }
  }
  
  var template = HtmlService.createTemplateFromFile("SidebarCores");
  template.cores = JSON.stringify(cores);
  var html = template.evaluate().setTitle("Cores Desatualizadas");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * processCoresFromSidebar: Para cada item (cor) selecionado na sidebar,
 * procura na aba "ESPELHO DO ESTOQUE" os registros cujo valor da coluna A seja igual (case-insensitive)
 * e pega os 5 últimos registros. Em seguida, escreve na aba "CORES DESATUALIZADAS" nas colunas:
 *   - A: Item (coluna A do ESPELHO)
 *   - B: Valor da coluna B do ESPELHO
 *   - C: Data (coluna C do ESPELHO)
 *   - D: Valor da coluna E do ESPELHO
 */
function processCoresFromSidebar(selectedCores) {
  if (!selectedCores || selectedCores.length === 0) {
    return "Nenhuma cor foi selecionada.";
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEspelho = ss.getSheetByName("ESPELHO DO ESTOQUE");
  if (!sheetEspelho) {
    throw new Error("A aba 'ESPELHO DO ESTOQUE' não foi encontrada.");
  }
  
  // Lê todos os registros da aba ESPELHO DO ESTOQUE (supondo cabeçalho na linha 1)
  var lastRowEspelho = sheetEspelho.getLastRow();
  if (lastRowEspelho < 2) {
    throw new Error("Não há dados na aba 'ESPELHO DO ESTOQUE'.");
  }
  var espelhoData = sheetEspelho.getRange(2, 1, lastRowEspelho - 1, sheetEspelho.getLastColumn()).getValues();
  
  var resultados = [];
  
  // Para cada item selecionado, filtra os registros cujo valor da coluna A (índice 0) seja igual (case-insensitive)
  selectedCores.forEach(function(item) {
    var filtrados = espelhoData.filter(function(row) {
      return row[0] && row[0].toString().toLowerCase() === item.toString().toLowerCase();
    });
    // Pega os 5 últimos registros
    var ultimos5 = filtrados.slice(-5);
    ultimos5.forEach(function(row) {
      // Mapeia: Coluna A: item (índice 0), Coluna B: valor (índice 1), Coluna C: data (índice 2), Coluna D: valor da coluna E (índice 4)
      resultados.push([row[0], row[1], row[2], row[4]]);
    });
  });
  
  // Escreve os resultados na aba "CORES DESATUALIZADAS" sobrescrevendo as colunas A:D
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) {
    sheetCores = ss.insertSheet("CORES DESATUALIZADAS");
  }
  // Limpa as colunas A a D
  sheetCores.getRange("A:D").clearContent();
  // Cabeçalho opcional
  sheetCores.getRange("A1:D1").setValues([["Item", "Valor B", "Data", "Valor E"]]);
  if (resultados.length > 0) {
    sheetCores.getRange(2, 1, resultados.length, 4).setValues(resultados);
  }
  
  return "Processamento concluído. Foram encontrados " + resultados.length + " registros.";
}

/* ================================
   Fim das Novas Funções
   ================================ */


/**
 * onOpen: Executada quando a planilha é aberta.
 * Apaga a propriedade "loggedUser", remove filtros na aba "ESTOQUE" e faz backup dos dados.
 * Exibe o diálogo de login (o menu só é criado após um login bem-sucedido).
 */
function onOpen() {
  PropertiesService.getUserProperties().deleteProperty("loggedUser");
  Logger.log("onOpen: Propriedade 'loggedUser' apagada.");
  
  backupEstoqueData();
  removeFilterOnOpen();
  showLoginDialog();
  // updateMenus() não é chamado aqui para restringir acesso sem login.
}

/**
 * removeFilterOnOpen: Remove o filtro ativo na aba "ESTOQUE", se existir.
 */
function removeFilterOnOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (sheetEstoque && sheetEstoque.getFilter()) {
    sheetEstoque.getFilter().remove();
    Logger.log("removeFilterOnOpen: Filtro removido na aba ESTOQUE.");
  }
}

/**
 * backupEstoqueData: Copia as últimas 500 linhas da aba "ESTOQUE" para a aba "BACKUP_ESTOQUE".
 */
function backupEstoqueData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) return;
  
  var lastRow = sheetEstoque.getLastRow();
  var startRow = Math.max(1, lastRow - 500 + 1);
  var numRows = lastRow - startRow + 1;
  var lastColumn = sheetEstoque.getLastColumn();
  var values = sheetEstoque.getRange(startRow, 1, numRows, lastColumn).getValues();
  
  var sheetBackup = ss.getSheetByName("BACKUP_ESTOQUE");
  if (!sheetBackup) {
    sheetBackup = ss.insertSheet("BACKUP_ESTOQUE");
  }
  if (sheetBackup.getMaxRows() < lastRow) {
    sheetBackup.insertRowsAfter(sheetBackup.getMaxRows(), lastRow - sheetBackup.getMaxRows());
  }
  sheetBackup.getRange(startRow, 1, numRows, lastColumn).clearContent();
  sheetBackup.getRange(startRow, 1, numRows, lastColumn).setValues(values);
  sheetBackup.hideSheet();
  Logger.log("backupEstoqueData: Backup das linhas de " + startRow + " até " + lastRow + " realizado.");
}

/**
 * onEdit: Se a edição ocorrer na aba EMBARQUES (colunas A, B ou E), chama atualizarTotalEmbarcado;
 * se ocorrer na aba ESTOQUE, impede edições manuais.
 */
function onEdit(e) {
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  if (sheetName === "EMBARQUES") {
    var col = e.range.getColumn();
    if (col === 1 || col === 2 || col === 5) {
      atualizarTotalEmbarcado();
    }
    return;
  }
  
  if (sheetName !== "ESTOQUE") return;
  
  var restoreEnabled = PropertiesService.getScriptProperties().getProperty("restoreEnabled");
  if (restoreEnabled === "false") {
    Logger.log("onEdit: Restauração desativada, nenhuma ação realizada.");
    return;
  }
  
  if (PropertiesService.getScriptProperties().getProperty("editingViaScript") === "true") {
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetBackup = ss.getSheetByName("BACKUP_ESTOQUE");
  if (!sheetBackup) {
    Logger.log("onEdit: Aba BACKUP_ESTOQUE não encontrada.");
    return;
  }
  
  var editedRange = e.range;
  var numRows = editedRange.getNumRows();
  var numCols = editedRange.getNumColumns();
  var startRow = editedRange.getRow();
  var startCol = editedRange.getColumn();
  
  var backupValues = sheetBackup.getRange(startRow, startCol, numRows, numCols).getValues();
  var newValues = [];
  for (var r = 0; r < numRows; r++) {
    var row = [];
    for (var c = 0; c < numCols; c++) {
      row.push(backupValues[r][c] !== "" ? backupValues[r][c] : "");
    }
    newValues.push(row);
  }
  
  PropertiesService.getScriptProperties().setProperty("editingViaScript", "true");
  editedRange.setValues(newValues);
  PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
  
  SpreadsheetApp.getUi().alert("Edição manual não é permitida. Utilize o sidebar para inserir dados.");
  Logger.log("onEdit: Edição manual detectada e revertida na faixa " + editedRange.getA1Notation());
}

/**
 * toggleRestore: Alterna a restauração de dados para permitir edições manuais temporariamente.
 */
function toggleRestore() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Digite a senha para alternar a restauração dos dados:");
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var senha = response.getResponseText();
  if (senha !== "919633") {
    ui.alert("Senha incorreta!");
    return;
  }
  var restoreEnabled = PropertiesService.getScriptProperties().getProperty("restoreEnabled");
  if (restoreEnabled === null || restoreEnabled === "true") {
    PropertiesService.getScriptProperties().setProperty("restoreEnabled", "false");
    ui.alert("Restauração desativada. Agora você poderá editar manualmente.");
  } else {
    PropertiesService.getScriptProperties().setProperty("restoreEnabled", "true");
    ui.alert("Restauração ativada. As edições manuais serão revertidas automaticamente.");
  }
  updateMenus();
}

/**
 * apagarUltimaLinha: Apaga a última linha preenchida da aba ESTOQUE.
 */
function apagarUltimaLinha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) {
    SpreadsheetApp.getUi().alert("A aba ESTOQUE não foi encontrada.");
    return;
  }
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Não há dados para apagar.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty("editingViaScript", "true");
  sheetEstoque.deleteRow(lastRow);
  PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
  backupEstoqueData();
  SpreadsheetApp.getUi().alert("Última linha apagada com sucesso.");
}

/**
 * showGrupoDialog: Abre o diálogo para inserir um novo grupo na aba DADOS.
 */
function showGrupoDialog() {
  var template = HtmlService.createTemplateFromFile("DialogInserirGrupo");
  template.groupList = JSON.stringify(getGroupList());
  var htmlOutput = template.evaluate().setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "INSERIR GRUPO");
}

/**
 * inserirGrupo: Insere o grupo na aba DADOS.
 */
function inserirGrupo(formData) {
  var group = formData.group;
  if (!group || group.trim() === "") {
    throw new Error("⚠️ Informe um grupo.");
  }
  group = group.trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) throw new Error("A aba DADOS não foi encontrada.");
  var existingGroups = getGroupList();
  if (existingGroups.indexOf(group) !== -1) {
    SpreadsheetApp.getUi().alert("Grupo já cadastrado.");
    return "Grupo já cadastrado.";
  }
  var lastRow = sheetDados.getLastRow();
  var newRow = lastRow < 2 ? 2 : lastRow + 1;
  sheetDados.getRange(newRow, 4).setValue(group);
  SpreadsheetApp.getUi().alert("Grupo inserido com sucesso.");
  return "Grupo inserido com sucesso!";
}

/**
 * atualizarTotalEmbarcado: Atualiza a aba TOTAL EMBARCADO com os cadastros exclusivos e seus totais.
 * Os cadastros são gravados como texto para evitar formatação como data.
 * Se na coluna E de EMBARQUES houver "CHEGOU", subtrai o valor (sem deixar negativo).
 * Cria filtro na faixa A:B. (Mensagem de alerta removida)
 */
function atualizarTotalEmbarcado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEmbarques = ss.getSheetByName("EMBARQUES");
  if (!sheetEmbarques) throw new Error("A aba EMBARQUES não foi encontrada.");
  
  var lastRow = sheetEmbarques.getLastRow();
  if (lastRow < 2) {
    return "Sem dados na aba EMBARQUES.";
  }
  
  var dataRange = sheetEmbarques.getRange(2, 1, lastRow - 1, sheetEmbarques.getLastColumn());
  var dataValues = dataRange.getValues();
  
  var totais = {};
  dataValues.forEach(function(row) {
    var cadastro = row[0] ? row[0].toString().trim() : "";
    if (cadastro === "") return;
    var valor = parseFloat(row[1]) || 0;
    var status = row[4] ? row[4].toString().trim().toLowerCase() : "";
    if (!totais.hasOwnProperty(cadastro)) {
      totais[cadastro] = 0;
    }
    if (status === "chegou") {
      totais[cadastro] = Math.max(totais[cadastro] - valor, 0);
    } else {
      totais[cadastro] += valor;
    }
  });
  
  var sheetTotal = ss.getSheetByName("TOTAL EMBARCADO");
  if (!sheetTotal) {
    sheetTotal = ss.insertSheet("TOTAL EMBARCADO");
  }
  sheetTotal.clearContents();
  sheetTotal.getRange(1, 1, 1, 2).setValues([["CADASTRO", "TOTAL"]]);
  
  var output = [];
  for (var cadastro in totais) {
    if (totais.hasOwnProperty(cadastro)) {
      output.push(["'" + cadastro, totais[cadastro]]);
    }
  }
  
  if (output.length > 0) {
    sheetTotal.getRange(2, 1, output.length, 2).setValues(output);
    sheetTotal.getRange(2, 1, output.length, 1).setNumberFormat("@");
  }
  
  if (sheetTotal.getFilter()) {
    sheetTotal.getFilter().remove();
  }
  sheetTotal.getRange(1, 1, sheetTotal.getLastRow(), 2).createFilter();
  
  return "Total embarcado atualizado com sucesso!";
}

/**
 * atualizarCompraDeFio: Atualiza a aba COMPRA DE FIO com os valores das abas RELATORIO e TOTAL EMBARCADO.
 * Compara o Total Compra com o threshold definido em J1 para definir "URGENTE" ou "ESTOQUE".
 */
function atualizarCompraDeFio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetCompra = ss.getSheetByName("COMPRA DE FIO");
  if (!sheetCompra) {
    throw new Error("A aba COMPRA DE FIO não foi encontrada.");
  }
  var compraData = sheetCompra.getDataRange().getValues();
  if (compraData.length < 2) {
    SpreadsheetApp.getUi().alert("Não há cadastros na aba COMPRA DE FIO para atualizar.");
    return;
  }
  var cadastrosCompra = compraData.slice(1).map(function(row) {
    return row[0] ? row[0].toString().replace(/^'/, "").trim() : "";
  });
  
  var sheetRelatorio = ss.getSheetByName("RELATORIO");
  if (!sheetRelatorio) {
    throw new Error("A aba RELATORIO não foi encontrada.");
  }
  var relData = sheetRelatorio.getDataRange().getValues();
  relData.shift();
  var relMap = {};
  relData.forEach(function(row) {
    var cad = row[0] ? row[0].toString().trim() : "";
    var valor = parseFloat(row[1]) || 0;
    if (cad) {
      relMap[cad] = valor;
    }
  });
  
  var sheetTotal = ss.getSheetByName("TOTAL EMBARCADO");
  if (!sheetTotal) {
    throw new Error("A aba TOTAL EMBARCADO não foi encontrada.");
  }
  var totalData = sheetTotal.getDataRange().getValues();
  totalData.shift();
  var totalMap = {};
  totalData.forEach(function(row) {
    var cad = row[0] ? row[0].toString().replace(/^'/, "").trim() : "";
    var valor = parseFloat(row[1]) || 0;
    if (cad) {
      totalMap[cad] = valor;
    }
  });
  
  var notFound = [];
  var totalCompra = [];
  var breakdownRel = [];
  var breakdownTot = [];
  
  cadastrosCompra.forEach(function(cad) {
    if (!cad) return;
    var valorRel = relMap.hasOwnProperty(cad) ? relMap[cad] : 0;
    var valorTotal = totalMap.hasOwnProperty(cad) ? totalMap[cad] : 0;
    var soma = valorRel + valorTotal;
    if (!relMap.hasOwnProperty(cad)) {
      notFound.push(cad);
    }
    totalCompra.push([soma]);
    breakdownRel.push([valorRel]);
    breakdownTot.push([valorTotal]);
  });
  
  var lastRowCompra = sheetCompra.getLastRow();
  if (lastRowCompra >= 2) {
    sheetCompra.getRange(2, 2, lastRowCompra - 1, 1).clearContent();
    sheetCompra.getRange(2, 5, lastRowCompra - 1, 1).clearContent();
    sheetCompra.getRange(2, 6, lastRowCompra - 1, 2).clearContent();
  }
  
  var threshold = parseFloat(sheetCompra.getRange("J1").getValue());
  if (isNaN(threshold)) {
    threshold = 0;
  }
  
  for (var i = 0; i < totalCompra.length; i++) {
    var totalValue = totalCompra[i][0];
    sheetCompra.getRange(i + 2, 2).setValue(totalValue);
    var label = parseFloat(totalValue) < threshold ? "URGENTE" : "ESTOQUE";
    sheetCompra.getRange(i + 2, 5).setValue(label);
    sheetCompra.getRange(i + 2, 6).setValue(breakdownRel[i][0]);
    sheetCompra.getRange(i + 2, 7).setValue(breakdownTot[i][0]);
  }
  
  var existingFilter = sheetCompra.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  sheetCompra.getRange(1, 1, sheetCompra.getLastRow(), 7).createFilter();
  
  if (notFound.length > 0) {
    SpreadsheetApp.getUi().alert("Os seguintes cadastros não foram encontrados no RELATORIO: " + notFound.join(", "));
  } else {
    SpreadsheetApp.getUi().alert("Compra de fio atualizada com sucesso!");
  }
  
  return "Compra de fio atualizada com sucesso!";
}

/**
 * copyCompraToHistorico: Copia os dados da aba COMPRA DE FIO para a aba HISTORICO, adicionando a data/hora atual.
 */
function copyCompraToHistorico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCompra = ss.getSheetByName("COMPRA DE FIO");
  var historicoSheet = ss.getSheetByName("HISTORICO");
  if (!historicoSheet) {
    historicoSheet = ss.insertSheet("HISTORICO");
  }
  
  var numRowsToCopy = sheetCompra.getLastRow() - 1;
  Logger.log("Número de linhas para copiar: " + numRowsToCopy);
  if (numRowsToCopy > 0) {
    var compData = sheetCompra.getRange(2, 1, numRowsToCopy, 7).getValues();
    var now = new Date();
    var historicoData = compData.map(function(row) {
      return row.concat([now]);
    });
    var lastRowHistorico = historicoSheet.getLastRow();
    var startRowHistorico = lastRowHistorico < 1 ? 1 : lastRowHistorico + 1;
    historicoSheet.getRange(startRowHistorico, 1, historicoData.length, historicoData[0].length).setValues(historicoData);
    Logger.log("Dados copiados para HISTORICO a partir da linha " + startRowHistorico);
  } else {
    Logger.log("Não há linhas para copiar na aba COMPRA DE FIO.");
  }
}

/**
 * atualizarCompraDeFioEHistorico: Executa atualizarCompraDeFio() e, em seguida, copyCompraToHistorico().
 */
function atualizarCompraDeFioEHistorico() {
  atualizarCompraDeFio();
  copyCompraToHistorico();
}

/**
 * showLoginDialog: Exibe o diálogo de login.
 */
function showLoginDialog() {
  var html = HtmlService.createTemplateFromFile("DialogLogin")
    .evaluate()
    .setWidth(350)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "LOGIN");
}

/**
 * processLogin: Valida as credenciais na aba DADOS e, se bem-sucedido, define "loggedUser" e cria o menu.
 */
function processLogin(formData) {
  Logger.log("processLogin: Dados recebidos: " + JSON.stringify(formData));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) {
    throw new Error("A aba DADOS não foi encontrada.");
  }
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 1) {
    throw new Error("Não há usuários cadastrados.");
  }
  var data = sheetDados.getRange(1, 2, lastRow, 2).getValues();
  var valid = false;
  for (var i = 0; i < data.length; i++) {
    var username = data[i][0];
    var password = data[i][1];
    if (username && password) {
      if (username.toString().trim() === formData.username.toString().trim() &&
          password.toString().trim() === formData.password.toString().trim()) {
        valid = true;
        break;
      }
    }
  }
  if (!valid) {
    throw new Error("Credenciais inválidas.");
  }
  PropertiesService.getUserProperties().setProperty("loggedUser", formData.username.toString().trim());
  Logger.log("processLogin: Login efetuado para " + formData.username);
  updateMenus();
  return "Login efetuado com sucesso!";
}

/**
 * getLoggedUser: Retorna o usuário logado.
 */
function getLoggedUser() {
  return PropertiesService.getUserProperties().getProperty("loggedUser");
}

/**
 * showEstoqueSidebar: Abre o formulário de cadastro de estoque na sidebar.
 */
function showEstoqueSidebar() {
  var nextRow = updateUnprotectedRange();
  Logger.log("showEstoqueSidebar: Próxima linha para cadastro: " + nextRow);

  // OTIMIZADO: 1 busca em vez de 4
  var autocompleteData = getAllAutocompleteData();

  var template = HtmlService.createTemplateFromFile("DialogEstoque");
  template.itemList = JSON.stringify(autocompleteData.items);
  template.groupList = JSON.stringify(autocompleteData.groups);
  template.nfList = JSON.stringify(autocompleteData.nfs);
  template.obsList = JSON.stringify(autocompleteData.obs);
  template.currentRow = nextRow;

  var htmlOutput = template.evaluate().setTitle("CADASTRO DE ESTOQUE");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * updateUnprotectedRange: Retorna a próxima linha livre na aba ESTOQUE.
 */
function updateUnprotectedRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  var nextRow = sheet.getLastRow() + 1;
  return nextRow;
}

/**
 * setActiveNextEmptyCell: Seleciona a célula da coluna A que está 15 linhas abaixo da última preenchida.
 */
function setActiveNextEmptyCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet) {
    var nextRow = sheet.getLastRow() + 15;
    sheet.activate();
    sheet.setActiveSelection("A" + nextRow);
    Logger.log("setActiveNextEmptyCell: Célula A" + nextRow + " selecionada.");
  }
}

/**
 * select4RowsBelow: Seleciona a célula da coluna A que está 4 linhas abaixo da última linha preenchida.
 */
function select4RowsBelow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet) {
    var nextRow = sheet.getLastRow() + 4;
    sheet.activate();
    sheet.setActiveSelection("A" + nextRow);
    Logger.log("select4RowsBelow: Célula A" + nextRow + " selecionada.");
  }
}

/**
 * select10RowsBelow: Seleciona a célula da coluna A que está 10 linhas abaixo da última linha preenchida.
 */
function select10RowsBelow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet) {
    var nextRow = sheet.getLastRow() + 10;
    sheet.activate();
    sheet.setActiveSelection("A" + nextRow);
    Logger.log("select10RowsBelow: Célula A" + nextRow + " selecionada.");
  }
}

/**
 * getItemList: Retorna a lista única de itens da aba DADOS (Coluna A).
 */
function getItemList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) return [];
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 1) return [];
  var values = sheetDados.getRange(1, 1, lastRow, 1).getValues().flat();
  var items = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i] && values[i].toString().trim() !== "") {
      items.push(values[i].toString().trim());
    }
  }
  return Array.from(new Set(items));
}

/**
 * getGroupList: Retorna a lista única de grupos da aba DADOS (Coluna D).
 */
function getGroupList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) return [];
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 1) return [];
  var values = sheetDados.getRange(1, 4, lastRow, 1).getValues().flat();
  var groups = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i] && values[i].toString().trim() !== "") {
      groups.push(values[i].toString().trim());
    }
  }
  return Array.from(new Set(groups));
}

/**
 * getNfList: Retorna a lista única de valores da coluna D da aba ESTOQUE (Nota Fiscal/Pedido).
 */
function getNfList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();
  var nfList = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  return Array.from(new Set(nfList));
}

/**
 * getObsList: Retorna a lista única de valores da coluna E da aba ESTOQUE (Cliente/Observações).
 */
function getObsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat();
  var obsList = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  return Array.from(new Set(obsList));
}

/**
 * getMedidasList: Retorna a lista de unidades de medida.
 * Combina opções da aba DADOS (coluna MEDIDAS) com opções já usadas na aba ESTOQUE (coluna C).
 */
function getMedidasList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var medidasSet = new Set();

  // 1. Busca na aba DADOS (coluna MEDIDAS)
  var sheetDados = ss.getSheetByName("DADOS");
  if (sheetDados) {
    var headers = sheetDados.getRange(1, 1, 1, sheetDados.getLastColumn()).getValues()[0];
    var medidasCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var headerUpper = headers[i].toString().toUpperCase().trim();
      if (headerUpper === "MEDIDAS" || headerUpper === "MEDIDA" || headerUpper === "UNIDADE" || headerUpper === "UNIDADES") {
        medidasCol = i + 1;
        break;
      }
    }

    if (medidasCol !== -1) {
      var lastRow = sheetDados.getLastRow();
      if (lastRow >= 2) {
        var values = sheetDados.getRange(2, medidasCol, lastRow - 1, 1).getValues().flat();
        values.forEach(function(v) {
          var val = v.toString().trim();
          if (val !== "") medidasSet.add(val);
        });
      }
    }
  }

  // 2. Busca na aba ESTOQUE (coluna C - Unidade) para incluir opções já usadas
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (sheetEstoque) {
    var lastRowEstoque = sheetEstoque.getLastRow();
    if (lastRowEstoque >= 2) {
      var valuesEstoque = sheetEstoque.getRange(2, 3, lastRowEstoque - 1, 1).getValues().flat();
      valuesEstoque.forEach(function(v) {
        var val = v.toString().trim();
        if (val !== "") medidasSet.add(val);
      });
    }
  }

  // Converte Set para Array e ordena
  var medidasList = Array.from(medidasSet).sort();
  return medidasList;
}

/**
 * getObservacoesList: Retorna a lista de observações.
 * Busca da aba DADOS, coluna F (OBSERVAÇÃO).
 */
function getObservacoesList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var obsSet = new Set();

  // Busca na aba DADOS, coluna F (índice 6)
  var sheetDados = ss.getSheetByName("DADOS");
  if (sheetDados) {
    var lastRow = sheetDados.getLastRow();
    if (lastRow >= 2) {
      // Coluna F = índice 6
      var values = sheetDados.getRange(2, 6, lastRow - 1, 1).getDisplayValues().flat();
      values.forEach(function(v) {
        var val = v.toString().trim();
        if (val !== "") obsSet.add(val);
      });
    }
  }

  // Converte Set para Array e ordena
  var obsList = Array.from(obsSet).sort();
  return obsList;
}

/**
 * normalize: Função auxiliar para normalizar texto.
 */
function normalize(text) {
  if (!text) return "";
  return text.toString().trim().toLowerCase().replace(/\s+/g, " ");
}

/* ================================
   FUNÇÕES DE CACHE E AUTOCOMPLETE
   ================================ */

/**
 * getCachedData: Busca dados no cache ou executa função e armazena no cache.
 */
function getCachedData(key, fetchFunction, ttl) {
  ttl = ttl || 120; // 2 minutos padrão
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log("Cache parse error: " + key);
    }
  }

  var data = fetchFunction();
  try {
    var jsonData = JSON.stringify(data);
    if (jsonData.length < 100000) {
      cache.put(key, jsonData, ttl);
    }
  } catch (e) {
    Logger.log("Cache save error: " + e.message);
  }

  return data;
}

/**
 * invalidateCache: Invalida caches.
 */
function invalidateCache(keys) {
  var cache = CacheService.getScriptCache();
  var keysToInvalidate = typeof keys === 'string' ? [keys] : (keys || []);
  keysToInvalidate.forEach(function(key) { cache.remove(key); });
  cache.remove("autocompleteData");
}

/**
 * invalidateAllAutocompleteCache: Invalida todos os caches de autocomplete.
 */
function invalidateAllAutocompleteCache() {
  invalidateCache(["itemList", "groupList", "nfList", "obsList", "autocompleteData"]);
}

/**
 * getAllAutocompleteData: Busca todos os dados de autocomplete em uma única operação.
 * OTIMIZADO: Usa cache de 10 minutos
 */
function getAllAutocompleteData() {
  return getCachedData("autocompleteData", function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1ª Leitura: DADOS (apenas grupos)
    var sheetDados = ss.getSheetByName("DADOS");
    var groups = [];
    if (sheetDados) {
      var lastRowDados = sheetDados.getLastRow();
      if (lastRowDados >= 1) {
        var dadosData = sheetDados.getRange(1, 4, lastRowDados, 1).getDisplayValues();
        for (var i = 0; i < dadosData.length; i++) {
          if (dadosData[i][0] && dadosData[i][0].toString().trim() !== "") {
            groups.push(dadosData[i][0].toString().trim());
          }
        }
      }
    }

    // 2ª Leitura: ESTOQUE (itens da coluna B e NFs da coluna E)
    // Estrutura: A=Grupo, B=Item, C=Unidade, D=Data, E=NF, F=Obs
    var sheetEstoque = ss.getSheetByName("ESTOQUE");
    var items = [], nfs = [];
    if (sheetEstoque) {
      var lastRowEstoque = sheetEstoque.getLastRow();
      if (lastRowEstoque >= 2) {
        // Lê colunas B até E (4 colunas: B, C, D, E)
        var estoqueData = sheetEstoque.getRange(2, 2, lastRowEstoque - 1, 4).getDisplayValues();
        for (var j = 0; j < estoqueData.length; j++) {
          // Coluna B (índice 0) = Item
          if (estoqueData[j][0] && estoqueData[j][0].toString().trim() !== "") {
            items.push(estoqueData[j][0].toString().trim());
          }
          // Coluna E (índice 3) = NF (já em formato texto com getDisplayValues)
          if (estoqueData[j][3] && estoqueData[j][3].toString().trim() !== "") {
            nfs.push(estoqueData[j][3].toString().trim());
          }
        }
      }
    }

    return {
      items: Array.from(new Set(items)),
      groups: Array.from(new Set(groups)),
      nfs: Array.from(new Set(nfs)),
      medidas: getMedidasList(),
      observacoes: getObservacoesList()
    };
  }, 120); // 2 minutos
}

/**
 * getLastRegistration: Retorna o último registro de um item (data, estoque e grupo).
 * OTIMIZADO: Lê apenas as últimas 2000 linhas da planilha ESTOQUE
 */
function getLastRegistration(item, currentRow) {
  Logger.log("=== getLastRegistration INICIADO ===");
  Logger.log("Item buscado: '" + item + "'");
  Logger.log("CurrentRow: " + currentRow);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) {
    Logger.log("ERRO: Aba ESTOQUE não encontrada!");
    return { lastDate: null, lastStock: 0, lastGroup: null };
  }

  var lastRow = sheetEstoque.getLastRow();
  Logger.log("Última linha da planilha: " + lastRow);
  if (lastRow < 2) {
    Logger.log("Planilha vazia - sem dados");
    return { lastDate: null, lastStock: 0, lastGroup: null };
  }

  // Lê TODA a planilha como TEXTO (getDisplayValues) para evitar problemas de formato
  var startRow = 2;
  var numRows = lastRow - startRow + 1;
  Logger.log("Lendo TODA a planilha - linhas de " + startRow + " até " + lastRow + " (" + numRows + " linhas)");

  // USA getDisplayValues() para forçar conversão para texto
  // Lê 10 colunas (A-J) para incluir o Saldo que está na coluna J
  var data = sheetEstoque.getRange(startRow, 1, numRows, 10).getDisplayValues();
  Logger.log("Usando getDisplayValues() para forçar formato de TEXTO");

  var result = { lastDate: null, lastStock: 0, lastGroup: null };
  var itemNormalized = normalize(item);
  Logger.log("Item normalizado: '" + itemNormalized + "'");

  var encontrados = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    var rowNum = startRow + i;
    if (rowNum >= currentRow) continue;

    var currentItem = data[i][1]; // Coluna B (Item) - agora em formato TEXTO
    if (currentItem && currentItem.toString().trim() !== "") {
      var currentItemNormalized = normalize(currentItem);

      if (currentItemNormalized === itemNormalized) {
        encontrados++;
        result.lastGroup = data[i][0];  // Coluna A (Grupo)
        result.lastDate = data[i][3];   // Coluna D (Data) - como texto
        result.lastStock = data[i][9];  // Coluna J (Saldo) - como texto
        Logger.log("✓ ENCONTRADO na linha " + rowNum);
        Logger.log("  Grupo: '" + result.lastGroup + "'");
        Logger.log("  Data: " + result.lastDate);
        Logger.log("  Estoque: " + result.lastStock);
        Logger.log("  Item raw: '" + currentItem + "'");
        break;
      }
    }
  }

  if (encontrados === 0) {
    Logger.log("✗ NENHUM REGISTRO ENCONTRADO para o item '" + item + "'");
  }

  Logger.log("=== getLastRegistration FINALIZADO ===");
  return result;
}

/**
 * getLastInfoFromDados: Retorna a última informação não vazia da coluna C da aba DADOS para um produto.
 */
function getLastInfoFromDados(product) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) return "";
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 2) return "";
  var data = sheetDados.getRange(2, 1, lastRow - 1, sheetDados.getLastColumn()).getValues();
  var lastInfo = "";
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === product && data[i][2].toString().trim() !== "") {
      lastInfo = data[i][2];
    }
  }
  return lastInfo;
}

/**
 * showCustomDialog: Exibe um diálogo HTML customizado com uma mensagem.
 */
function showCustomDialog(message) {
  var template = HtmlService.createTemplateFromFile("CustomDialog");
  template.message = message;
  var html = template.evaluate().setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, "AVISO");
}

/**
 * localizarProduto: Abre o diálogo para localizar um produto.
 */
function localizarProduto() {
  var template = HtmlService.createTemplateFromFile("DialogLocalizarProduto");
  template.produtos = JSON.stringify(getProdutosEstoque());
  var htmlOutput = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "LOCALIZAR PRODUTO");
}

/**
 * getProdutosEstoque: Retorna a lista única de produtos da aba ESTOQUE (Coluna B).
 */
function getProdutosEstoque() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("produtosEstoque");
  if (cached) {
    return JSON.parse(cached);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var range = sheet.getRange("B2:B" + lastRow);
  var values = range.getDisplayValues().flat();
  var produtos = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  var unique = Array.from(new Set(produtos));
  cache.put("produtosEstoque", JSON.stringify(unique), 300);
  return unique;
}

/**
 * filtrarProduto: Aplica um filtro na aba ESTOQUE para exibir apenas as linhas cujo valor da coluna B seja igual ao produto.
 */
function filtrarProduto(produto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  var range = sheet.getDataRange();
  var filter = range.createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(produto).build();
  filter.setColumnFilterCriteria(2, criteria);
}

/**
 * mostrarTodos: Remove o filtro, ordena a aba ESTOQUE pela data (Coluna C) e seleciona uma célula.
 */
function mostrarTodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).sort({ column: 3, ascending: true });
  }
  setActiveNextEmptyCell();
}

/**
 * abrirDialogRelatorioEstoque: Abre o diálogo para definir a faixa de data do relatório.
 */
function abrirDialogRelatorioEstoque() {
  var html = HtmlService.createTemplateFromFile("DialogRelatorioEstoque")
      .evaluate()
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "RELATÓRIO DE ESTOQUE");
}

/**
 * gerarRelatorioEstoque: Gera o relatório geral para o período definido.
 */
function gerarRelatorioEstoque(dataInicio, dataFim) {
  Logger.log("gerarRelatorioEstoque: Início " + dataInicio + " - Fim " + dataFim);
  var partsInicio = dataInicio.split("/");
  var partsFim = dataFim.split("/");
  var startDate = new Date(partsInicio[2], partsInicio[1] - 1, partsInicio[0], 0, 0, 0);
  var endDate = new Date(partsFim[2], partsFim[1] - 1, partsFim[0], 23, 59, 59);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  var lastColumn = sheetEstoque.getLastColumn();
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, lastColumn);
  var dataValues = dataRange.getValues();
  
  var filtered = dataValues.filter(function(row) {
    var dt = new Date(row[2]);
    return dt >= startDate && dt <= endDate;
  });
  
  var grupos = {};
  filtered.forEach(function(row) {
    var prod = row[1];
    if (!grupos[prod]) {
      grupos[prod] = row;
    } else {
      var currentDate = new Date(row[2]);
      var storedDate = new Date(grupos[prod][2]);
      if (currentDate > storedDate) {
        grupos[prod] = row;
      }
    }
  });
  
  var reportData = [];
  for (var prod in grupos) {
    var row = grupos[prod];
    reportData.push([prod, row[8], row[4], row[2]]);
  }
  
  reportData.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });
  
  var sheetRelatorio = ss.getSheetByName("RELATORIO");
  if (!sheetRelatorio) {
    sheetRelatorio = ss.insertSheet("RELATORIO");
    sheetRelatorio.getRange("J1").setValue(0);
  }
  var threshold = parseFloat(sheetRelatorio.getRange("J1").getValue());
  if (isNaN(threshold)) {
    threshold = 0;
  }
  
  for (var i = 0; i < reportData.length; i++) {
    var novoSaldo = parseFloat(reportData[i][1]);
    if (novoSaldo < threshold) {
      reportData[i].push("URGENTE");
    } else {
      reportData[i].push("ESTOQUE");
    }
  }
  
  sheetRelatorio.clearContents();
  sheetRelatorio.getRange("J1").setValue(threshold);
  sheetRelatorio.getRange(1, 1, 1, 5).setValues([["PRODUTO", "NOVO SALDO", "OBS", "DATA/HORA", "STATUS"]]);
  if (reportData.length > 0) {
    sheetRelatorio.getRange(2, 1, reportData.length, 5).setValues(reportData);
  }
  
  var relFilter = sheetRelatorio.getFilter();
  if (relFilter) {
    relFilter.remove();
  }
  sheetRelatorio.getRange(1, 1, sheetRelatorio.getLastRow(), 5).createFilter();
  
  Logger.log("gerarRelatorioEstoque: Relatório gerado com " + reportData.length + " registros.");
  return "Relatório gerado com sucesso!";
}

/**
 * abrirDialogRelatorioPorGrupo: Abre o diálogo para definir o grupo do relatório.
 */
function abrirDialogRelatorioPorGrupo() {
  var template = HtmlService.createTemplateFromFile("DialogRelatorioPorGrupo");
  template.grupos = JSON.stringify(getGruposEstoque());
  var htmlOutput = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "RELATÓRIO POR GRUPO");
}

/**
 * getGruposEstoque: Retorna os grupos únicos da aba ESTOQUE (Coluna A).
 */
function getGruposEstoque() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  var grupos = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  return Array.from(new Set(grupos));
}

/**
 * gerarRelatorioPorGrupo: Gera o relatório para um grupo específico.
 */
function gerarRelatorioPorGrupo(grupoSelecionado) {
  Logger.log("gerarRelatorioPorGrupo: Grupo selecionado: " + grupoSelecionado);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  var lastColumn = sheetEstoque.getLastColumn();
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, lastColumn);
  var dataValues = dataRange.getValues();
  
  var filtered = dataValues.filter(function(row) {
    return row[0].toString().trim() === grupoSelecionado;
  });
  
  var gruposItens = {};
  filtered.forEach(function(row) {
    var item = row[1];
    if (!gruposItens[item]) {
      gruposItens[item] = row;
    } else {
      var currentDate = new Date(row[2]);
      var storedDate = new Date(gruposItens[item][2]);
      if (currentDate > storedDate) {
        gruposItens[item] = row;
      }
    }
  });
  
  var reportData = [];
  for (var item in gruposItens) {
    var row = gruposItens[item];
    reportData.push([row[0], row[1], row[8], row[2]]);
  }
  
  reportData.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });
  
  var sheetRelatorio = ss.getSheetByName("RELATORIO POR GRUPO DE ITEM");
  if (!sheetRelatorio) {
    sheetRelatorio = ss.insertSheet("RELATORIO POR GRUPO DE ITEM");
  }
  sheetRelatorio.clearContents();
  sheetRelatorio.getRange(1, 1, 1, 4).setValues([["GRUPO", "ITEM", "NOVO SALDO", "DATA/HORA"]]);
  if (reportData.length > 0) {
    sheetRelatorio.getRange(2, 1, reportData.length, 4).setValues(reportData);
  }
  
  Logger.log("gerarRelatorioPorGrupo: Relatório gerado com " + reportData.length + " registros.");
  return "Relatório por grupo gerado com sucesso!";
}

/**
 * showListagemEstoqueSidebar: Abre a sidebar para a listagem de estoque.
 */
function showListagemEstoqueSidebar() {
  var template = HtmlService.createTemplateFromFile('DialogListagemEstoque');
  template.produtos = JSON.stringify(getProdutosEstoque());
  var html = template.evaluate()
    .setTitle('Listagem de Estoque')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * gerarListagemEstoque: Processa os itens da sidebar e gera/atualiza a aba "LISTAGEM DE ESTOQUE".
 */
function gerarListagemEstoque(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) coleta até 20 itens
  var items = [];
  for (var i = 1; i <= 20; i++) {
    var v = formData['item' + i];
    if (v) items.push(v.toString().trim().toLowerCase());
  }
  if (items.length === 0) throw new Error("⚠️ Informe pelo menos um item.");

  // 2) lê dados da aba ESTOQUE - OTIMIZADO: lê apenas colunas necessárias
  var sheetEst = ss.getSheetByName('ESTOQUE');
  var lastRow = sheetEst.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  // Lê apenas colunas B (item), C (data), I (novo saldo)
  var dados = sheetEst.getRange(2, 1, lastRow - 1, 9).getValues();

  // 3) OTIMIZADO: Cria índice de últimos registros (1 passada em vez de N*M)
  var itemIndex = {};
  for (var j = dados.length - 1; j >= 0; j--) {
    var prod = dados[j][1] ? dados[j][1].toString().trim().toLowerCase() : '';
    if (prod && !itemIndex[prod]) {
      itemIndex[prod] = {
        saldo: dados[j][8],
        data: dados[j][2]
      };
    }
  }

  // 4) monta listagem (item, último saldo e data)
  var listagemData = items.map(function(item) {
    var key = item.toLowerCase();
    var registro = itemIndex[key];
    if (registro) {
      return [item, registro.saldo, registro.data];
    }
    return [item, '', ''];
  });

  // 4) grava na aba LISTAGEM DE ESTOQUE
  var sheetL = ss.getSheetByName('LISTAGEM DE ESTOQUE')
             || ss.insertSheet('LISTAGEM DE ESTOQUE');

  // remove filtro antigo, se houver
  if (sheetL.getFilter()) sheetL.getFilter().remove();
  sheetL.clearContents();

  // escreve cabeçalho e dados
  sheetL.getRange(1, 1, 1, 3)
        .setValues([['ITEM','ÚLTIMO SALDO','DATA/HORA']]);
  if (listagemData.length) {
    sheetL.getRange(2, 1, listagemData.length, 3)
          .setValues(listagemData);
  }
  sheetL.getRange(1, 1, sheetL.getLastRow(), 3).createFilter();

  // 5) monta e retorna o HTML para a sidebar, com data formatada
  var tz = Session.getScriptTimeZone();
  var html = '<table style="width:100%;border-collapse:collapse;">'
           + '<tr><th>ITEM</th><th>ÚLTIMO SALDO</th><th>DATA</th></tr>';
  listagemData.forEach(function(r) {
    // formata só a data como dd/MM/yyyy
    var dataStr = '';
    if (r[2] instanceof Date) {
      dataStr = Utilities.formatDate(r[2], tz, 'dd/MM/yyyy');
    } else {
      dataStr = r[2] || '';
    }
    html += '<tr>'
         +   '<td>'+r[0]+'</td>'
         +   '<td>'+r[1]+'</td>'
         +   '<td>'+dataStr+'</td>'
         + '</tr>';
  });
  html += '</table>';
  return html;
}

/**
 * testarCadastro: Função de teste para simular um cadastro.
 */
function testarCadastro() {
  processEstoque({
    group: "Grupo Teste",
    item: "Produto Teste",
    nf: "NF123",
    obs: "Observação Teste",
    entrada: 10,
    saida: 3
  });
  Logger.log("testarCadastro: Cadastro de teste executado.");
}

/**
 * processEstoque: Processa os dados do cadastro inserido via formulário.
 */
/**
 * processEstoque: Processa os dados do cadastro inserido via formulário.
 * Atualizado para não pintar de vermelho se a coluna E conter 'ACERTO' ou 'ATUALIZAÇÃO' (variações).
 */
function processEstoque(formData) {
  Logger.log("processEstoque: Dados inseridos: " + JSON.stringify(formData));
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  var now = new Date();
  var nextRow = sheetEstoque.getLastRow() + 1;
  
  PropertiesService.getScriptProperties().setProperty("editingViaScript", "true");
  
  // Recupera último registro para cálculo de saldo e data
  var lastReg = getLastRegistration(formData.item, nextRow);
  var previousSaldo = parseFloat(lastReg.lastStock) || 0;
  var newSaldo = previousSaldo + parseFloat(formData.entrada) - parseFloat(formData.saida);
  var rowData = [
    formData.group,
    formData.item,
    now,
    formData.nf,
    formData.obs,
    previousSaldo,
    formData.entrada,
    formData.saida,
    newSaldo,
    now,
    getLoggedUser()
  ];
  
  try {
    sheetEstoque.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    Logger.log("processEstoque: Dados inseridos na linha " + nextRow);
  } catch (err) {
    Logger.log("processEstoque: Erro ao inserir dados: " + err);
    showCustomDialog("Erro ao inserir os dados. Por favor, contate o administrador.");
    PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
    return;
  }
  
  PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
  backupEstoqueData();

  // Verifica se passou mais de 20 dias desde a última data de registro
  if (lastReg.lastDate) {
    var lastDate = new Date(lastReg.lastDate);
    var diffDays = (now.getTime() - lastDate.getTime()) / (1000 * 3600 * 24);
    if (diffDays > 20) {
      // Verifica coluna E (obs) por palavras-chave
      var textoObs = formData.obs ? formData.obs.toString().toLowerCase() : "";
      var temKeyword = /acerto|atualiza[cç][ãa]o/.test(textoObs);
      // Se não conter 'acerto' ou 'atualização', pinta de vermelho
      if (!temKeyword) {
        var lastColumn = sheetEstoque.getLastColumn();
        sheetEstoque.getRange(nextRow, 1, 1, lastColumn).setBackground("red");
        showCustomDialog("⚠️ PRODUTO DESATUALIZADO (ÚLTIMA ATUALIZAÇÃO HÁ MAIS DE 20 DIAS). POR FAVOR, ATUALIZAR URGENTE.");
        return;
      }
    }
  }

  // Verifica se houve ENTRADA de estoque - aviso para atualização
  if (parseFloat(formData.entrada) > 0) {
    var lastColumn = sheetEstoque.getLastColumn();
    sheetEstoque.getRange(nextRow, 1, 1, lastColumn).setBackground("yellow");
    showCustomDialog("⚠️ ENTRADA DE ESTOQUE REGISTRADA!\n\nÉ NECESSÁRIO ATUALIZAR O ESTOQUE DESTE ITEM PARA EVITAR FUROS DE ESTOQUE.\n\nRealize uma contagem física e registre uma atualização completa do saldo.");
    return;
  }

  // Se não entrou no critério, não exibe diálogo de sucesso para agilizar o cadastro.
}

/* ================================
   NOVAS FUNÇÕES: Estoque por Período e Limpar Filtro
   ================================ */

/**
 * abrirDialogEstoquePorPeriodo: Abre um diálogo para que o usuário informe as datas de início e fim.
 */
function abrirDialogEstoquePorPeriodo() {
  var html = HtmlService.createTemplateFromFile("DialogEstoquePorPeriodo")
    .evaluate()
    .setWidth(350)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, "Filtrar Estoque por Período");
}

/**
 * filtrarEstoquePorPeriodo: Copia as linhas da aba ESTOQUE, cuja data na coluna C
 * esteja entre dataInicio e dataFim (formato dd/mm/yyyy), e as cola na aba "FILTRO POR PERIODO".
 * Antes de colar, apaga o conteúdo anterior da aba.
 */
function filtrarEstoquePorPeriodo(dataInicio, dataFim) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  // Divide as datas informadas (formato dd/mm/yyyy) em partes e cria objetos Date.
  var partsInicio = dataInicio.split("/");
  var partsFim = dataFim.split("/");
  if (partsInicio.length !== 3 || partsFim.length !== 3) {
    throw new Error("Formato de data inválido. Use dd/mm/yyyy");
  }
  
  var startDate = new Date(
    parseInt(partsInicio[2], 10), 
    parseInt(partsInicio[1], 10) - 1, 
    parseInt(partsInicio[0], 10)
  );
  
  var endDate = new Date(
    parseInt(partsFim[2], 10), 
    parseInt(partsFim[1], 10) - 1, 
    parseInt(partsFim[0], 10),
    23, 59, 59, 999
  );
  
  // Obtém os dados da aba ESTOQUE (assumindo que a primeira linha é o cabeçalho)
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, sheetEstoque.getLastColumn());
  var dataValues = dataRange.getValues();
  
  // Prepara a aba de destino "FILTRO POR PERIODO"
  var sheetFiltro = ss.getSheetByName("FILTRO POR PERIODO");
  if (!sheetFiltro) {
    sheetFiltro = ss.insertSheet("FILTRO POR PERIODO");
  } else {
    sheetFiltro.clear();
  }
  
  // Copia o cabeçalho da aba ESTOQUE para a aba "FILTRO POR PERIODO"
  var header = sheetEstoque.getRange(1, 1, 1, sheetEstoque.getLastColumn()).getValues();
  sheetFiltro.getRange(1, 1, 1, header[0].length).setValues(header);
  
  var targetData = [];
  
  // Percorre cada linha e copia as que tiverem data na coluna C (índice 2) dentro do período
  for (var i = 0; i < dataValues.length; i++) {
    var row = dataValues[i];
    var dateValue = row[2];
    if (!(dateValue instanceof Date)) continue;
    if (dateValue >= startDate && dateValue <= endDate) {
      targetData.push(row);
    }
  }
  
  if (targetData.length > 0) {
    sheetFiltro.getRange(2, 1, targetData.length, targetData[0].length).setValues(targetData);
  }
  
  var targetLastRow = sheetFiltro.getLastRow();
  if (targetLastRow > 1) {
    sheetFiltro.getRange(2, 1, targetLastRow - 1, sheetFiltro.getLastColumn())
              .sort({ column: 3, ascending: true });
  }
  
  return "Dados do período de " + dataInicio + " a " + dataFim + " foram copiados para a aba 'FILTRO POR PERIODO'.";
}

/**
 * limparFiltroEstoque: Remove o filtro da aba ESTOQUE, ordena pela coluna C (datas) de forma ascendente
 * e seleciona a célula 4 linhas abaixo da última linha preenchida.
 */
function limparFiltroEstoque() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 3, ascending: true });
  }
  
  select4RowsBelow();
  
  return "Filtro removido e planilha ordenada por data.";
}

/**
 * convertDateFormat: Converte uma data do formato dd/mm/yyyy para mm/dd/yyyy.
 */
function convertDateFormat(dateStr) {
  var parts = dateStr.split("/");
  if (parts.length !== 3) throw new Error("Data inválida: " + dateStr);
  return parts[1] + "/" + parts[0] + "/" + parts[2];
}

/* ================================
   NOVAS FUNÇÕES: Estoque 3 Meses
   ================================ */

function showEstoque3MesesSidebar() {
  var template = HtmlService.createTemplateFromFile("DialogEstoque3Meses");
  template.itemList = JSON.stringify(getItemList());
  template.evaluate()
          .setTitle("Estoque 3 Meses")
          .setWidth(350)
          .setHeight(400);
  SpreadsheetApp.getUi().showSidebar(template);
}


/* ================================
   NOVAS FUNÇÕES: Cores Desatualizadas
   ================================ */
// ==============================
// Code.gs
// ==============================

/**
 * updateMenus: Cria o menu customizado na interface do Google Sheets.
 * O menu inclui agora as opções "Estoque por Período", "Limpar Filtro", "Estoque 3 Meses" e "Cores Desatualizadas".
 */
function updateMenus() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GESTÃO DO ESTOQUE")
    .addItem("Inserir Estoque", "showEstoqueSidebar")
    .addItem("Inserir Grupo", "showGrupoDialog")
    .addSeparator()
    .addItem("Localizar Produto", "localizarProduto")
    .addItem("Mostrar Todos", "mostrarTodos")
    .addSeparator()
    .addItem("Gerar Relatório", "abrirDialogRelatorioEstoque")
    .addItem("Relatório por Grupo", "abrirDialogRelatorioPorGrupo")
    .addItem("Listagem de Estoque", "showListagemEstoqueSidebar")
    .addItem("Atualizar Compra de Fio e Histórico", "atualizarCompraDeFioEHistorico")
    .addSeparator()
    .addItem("Atualizar Total Embarcado", "atualizarTotalEmbarcado")
    .addItem("Alternar Restauração", "toggleRestore")
    .addItem("Apagar Última Linha", "apagarUltimaLinha")
    .addSeparator()
    .addItem("ÚLTIMA LINHA", "select10RowsBelow")
    .addSeparator()
    .addItem("Estoque por Período", "abrirDialogEstoquePorPeriodo")
    .addItem("Limpar Filtro", "limparFiltroEstoque")
    .addSeparator()
    .addItem("Estoque 3 Meses", "showEstoque3MesesSidebar")
    .addSeparator()
    .addItem("Cores Desatualizadas", "showCoresDesatualizadasDialog")
    .addToUi();
}

/**
 * onOpen: Executada quando a planilha é aberta.
 * Apaga a propriedade "loggedUser", remove filtros na aba "ESTOQUE" e faz backup dos dados.
 * Exibe o diálogo de login (o menu só é criado após um login bem-sucedido).
 */
function onOpen() {
  PropertiesService.getUserProperties().deleteProperty("loggedUser");
  Logger.log("onOpen: Propriedade 'loggedUser' apagada.");
  
  backupEstoqueData();
  removeFilterOnOpen();
  showLoginDialog();
  // updateMenus() não é chamado aqui para restringir acesso sem login.
}

/**
 * removeFilterOnOpen: Remove o filtro ativo na aba "ESTOQUE", se existir.
 */
function removeFilterOnOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (sheetEstoque && sheetEstoque.getFilter()) {
    sheetEstoque.getFilter().remove();
    Logger.log("removeFilterOnOpen: Filtro removido na aba ESTOQUE.");
  }
}

/**
 * backupEstoqueData: Copia as últimas 500 linhas da aba "ESTOQUE" para a aba "BACKUP_ESTOQUE".
 */
function backupEstoqueData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) return;
  
  var lastRow = sheetEstoque.getLastRow();
  var startRow = Math.max(1, lastRow - 500 + 1);
  var numRows = lastRow - startRow + 1;
  var lastColumn = sheetEstoque.getLastColumn();
  var values = sheetEstoque.getRange(startRow, 1, numRows, lastColumn).getValues();
  
  var sheetBackup = ss.getSheetByName("BACKUP_ESTOQUE");
  if (!sheetBackup) {
    sheetBackup = ss.insertSheet("BACKUP_ESTOQUE");
  }
  if (sheetBackup.getMaxRows() < lastRow) {
    sheetBackup.insertRowsAfter(sheetBackup.getMaxRows(), lastRow - sheetBackup.getMaxRows());
  }
  sheetBackup.getRange(startRow, 1, numRows, lastColumn).clearContent();
  sheetBackup.getRange(startRow, 1, numRows, lastColumn).setValues(values);
  sheetBackup.hideSheet();
  Logger.log("backupEstoqueData: Backup das linhas de " + startRow + " até " + lastRow + " realizado.");
}

/**
 * onEdit: Se a edição ocorrer na aba EMBARQUES (colunas A, B ou E), chama atualizarTotalEmbarcado;
 * se ocorrer na aba ESTOQUE, impede edições manuais.
 */
function onEdit(e) {
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  if (sheetName === "EMBARQUES") {
    var col = e.range.getColumn();
    if (col === 1 || col === 2 || col === 5) {
      atualizarTotalEmbarcado();
    }
    return;
  }
  
  if (sheetName !== "ESTOQUE") return;
  
  var restoreEnabled = PropertiesService.getScriptProperties().getProperty("restoreEnabled");
  if (restoreEnabled === "false") {
    Logger.log("onEdit: Restauração desativada, nenhuma ação realizada.");
    return;
  }
  
  if (PropertiesService.getScriptProperties().getProperty("editingViaScript") === "true") {
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetBackup = ss.getSheetByName("BACKUP_ESTOQUE");
  if (!sheetBackup) {
    Logger.log("onEdit: Aba BACKUP_ESTOQUE não encontrada.");
    return;
  }
  
  var editedRange = e.range;
  var numRows = editedRange.getNumRows();
  var numCols = editedRange.getNumColumns();
  var startRow = editedRange.getRow();
  var startCol = editedRange.getColumn();
  
  var backupValues = sheetBackup.getRange(startRow, startCol, numRows, numCols).getValues();
  var newValues = [];
  for (var r = 0; r < numRows; r++) {
    var row = [];
    for (var c = 0; c < numCols; c++) {
      row.push(backupValues[r][c] !== "" ? backupValues[r][c] : "");
    }
    newValues.push(row);
  }
  
  PropertiesService.getScriptProperties().setProperty("editingViaScript", "true");
  editedRange.setValues(newValues);
  PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
  
  SpreadsheetApp.getUi().alert("Edição manual não é permitida. Utilize o sidebar para inserir dados.");
  Logger.log("onEdit: Edição manual detectada e revertida na faixa " + editedRange.getA1Notation());
}

/**
 * toggleRestore: Alterna a restauração de dados para permitir edições manuais temporariamente.
 */
function toggleRestore() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Digite a senha para alternar a restauração dos dados:");
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var senha = response.getResponseText();
  if (senha !== "919633") {
    ui.alert("Senha incorreta!");
    return;
  }
  var restoreEnabled = PropertiesService.getScriptProperties().getProperty("restoreEnabled");
  if (restoreEnabled === null || restoreEnabled === "true") {
    PropertiesService.getScriptProperties().setProperty("restoreEnabled", "false");
    ui.alert("Restauração desativada. Agora você poderá editar manualmente.");
  } else {
    PropertiesService.getScriptProperties().setProperty("restoreEnabled", "true");
    ui.alert("Restauração ativada. As edições manuais serão revertidas automaticamente.");
  }
  updateMenus();
}

/**
 * apagarUltimaLinha: Apaga a última linha preenchida da aba ESTOQUE.
 */
function apagarUltimaLinha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) {
    SpreadsheetApp.getUi().alert("A aba ESTOQUE não foi encontrada.");
    return;
  }
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Não há dados para apagar.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty("editingViaScript", "true");
  sheetEstoque.deleteRow(lastRow);
  PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
  backupEstoqueData();
  SpreadsheetApp.getUi().alert("Última linha apagada com sucesso.");
}

/**
 * showGrupoDialog: Abre o diálogo para inserir um novo grupo na aba DADOS.
 */
function showGrupoDialog() {
  var template = HtmlService.createTemplateFromFile("DialogInserirGrupo");
  template.groupList = JSON.stringify(getGroupList());
  var htmlOutput = template.evaluate().setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "INSERIR GRUPO");
}

/**
 * inserirGrupo: Insere o grupo na aba DADOS.
 */
function inserirGrupo(formData) {
  var group = formData.group;
  if (!group || group.trim() === "") {
    throw new Error("⚠️ Informe um grupo.");
  }
  group = group.trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) throw new Error("A aba DADOS não foi encontrada.");
  var existingGroups = getGroupList();
  if (existingGroups.indexOf(group) !== -1) {
    SpreadsheetApp.getUi().alert("Grupo já cadastrado.");
    return "Grupo já cadastrado.";
  }
  var lastRow = sheetDados.getLastRow();
  var newRow = lastRow < 2 ? 2 : lastRow + 1;
  sheetDados.getRange(newRow, 4).setValue(group);
  SpreadsheetApp.getUi().alert("Grupo inserido com sucesso.");
  return "Grupo inserido com sucesso!";
}

/**
 * atualizarTotalEmbarcado: Atualiza a aba TOTAL EMBARCADO com os cadastros exclusivos e seus totais.
 * Os cadastros são gravados como texto para evitar formatação como data.
 * Se na coluna E de EMBARQUES houver "CHEGOU", subtrai o valor (sem deixar negativo).
 * Cria filtro na faixa A:B. (Mensagem de alerta removida)
 */
function atualizarTotalEmbarcado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEmbarques = ss.getSheetByName("EMBARQUES");
  if (!sheetEmbarques) throw new Error("A aba EMBARQUES não foi encontrada.");
  
  var lastRow = sheetEmbarques.getLastRow();
  if (lastRow < 2) {
    return "Sem dados na aba EMBARQUES.";
  }
  
  var dataRange = sheetEmbarques.getRange(2, 1, lastRow - 1, sheetEmbarques.getLastColumn());
  var dataValues = dataRange.getValues();
  
  var totais = {};
  dataValues.forEach(function(row) {
    var cadastro = row[0] ? row[0].toString().trim() : "";
    if (cadastro === "") return;
    var valor = parseFloat(row[1]) || 0;
    var status = row[4] ? row[4].toString().trim().toLowerCase() : "";
    if (!totais.hasOwnProperty(cadastro)) {
      totais[cadastro] = 0;
    }
    if (status === "chegou") {
      totais[cadastro] = Math.max(totais[cadastro] - valor, 0);
    } else {
      totais[cadastro] += valor;
    }
  });
  
  var sheetTotal = ss.getSheetByName("TOTAL EMBARCADO");
  if (!sheetTotal) {
    sheetTotal = ss.insertSheet("TOTAL EMBARCADO");
  }
  sheetTotal.clearContents();
  sheetTotal.getRange(1, 1, 1, 2).setValues([["CADASTRO", "TOTAL"]]);
  
  var output = [];
  for (var cadastro in totais) {
    if (totais.hasOwnProperty(cadastro)) {
      output.push(["'" + cadastro, totais[cadastro]]);
    }
  }
  
  if (output.length > 0) {
    sheetTotal.getRange(2, 1, output.length, 2).setValues(output);
    sheetTotal.getRange(2, 1, output.length, 1).setNumberFormat("@");
  }
  
  if (sheetTotal.getFilter()) {
    sheetTotal.getFilter().remove();
  }
  sheetTotal.getRange(1, 1, sheetTotal.getLastRow(), 2).createFilter();
  
  return "Total embarcado atualizado com sucesso!";
}

/**
 * atualizarCompraDeFio: Atualiza a aba COMPRA DE FIO com os valores das abas RELATORIO e TOTAL EMBARCADO.
 * Compara o Total Compra com o threshold definido em J1 para definir "URGENTE" ou "ESTOQUE".
 */
function atualizarCompraDeFio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetCompra = ss.getSheetByName("COMPRA DE FIO");
  if (!sheetCompra) {
    throw new Error("A aba COMPRA DE FIO não foi encontrada.");
  }
  var compraData = sheetCompra.getDataRange().getValues();
  if (compraData.length < 2) {
    SpreadsheetApp.getUi().alert("Não há cadastros na aba COMPRA DE FIO para atualizar.");
    return;
  }
  var cadastrosCompra = compraData.slice(1).map(function(row) {
    return row[0] ? row[0].toString().replace(/^'/, "").trim() : "";
  });
  
  var sheetRelatorio = ss.getSheetByName("RELATORIO");
  if (!sheetRelatorio) {
    throw new Error("A aba RELATORIO não foi encontrada.");
  }
  var relData = sheetRelatorio.getDataRange().getValues();
  relData.shift();
  var relMap = {};
  relData.forEach(function(row) {
    var cad = row[0] ? row[0].toString().trim() : "";
    var valor = parseFloat(row[1]) || 0;
    if (cad) {
      relMap[cad] = valor;
    }
  });
  
  var sheetTotal = ss.getSheetByName("TOTAL EMBARCADO");
  if (!sheetTotal) {
    throw new Error("A aba TOTAL EMBARCADO não foi encontrada.");
  }
  var totalData = sheetTotal.getDataRange().getValues();
  totalData.shift();
  var totalMap = {};
  totalData.forEach(function(row) {
    var cad = row[0] ? row[0].toString().replace(/^'/, "").trim() : "";
    var valor = parseFloat(row[1]) || 0;
    if (cad) {
      totalMap[cad] = valor;
    }
  });
  
  var notFound = [];
  var totalCompra = [];
  var breakdownRel = [];
  var breakdownTot = [];
  
  cadastrosCompra.forEach(function(cad) {
    if (!cad) return;
    var valorRel = relMap.hasOwnProperty(cad) ? relMap[cad] : 0;
    var valorTotal = totalMap.hasOwnProperty(cad) ? totalMap[cad] : 0;
    var soma = valorRel + valorTotal;
    if (!relMap.hasOwnProperty(cad)) {
      notFound.push(cad);
    }
    totalCompra.push([soma]);
    breakdownRel.push([valorRel]);
    breakdownTot.push([valorTotal]);
  });
  
  var lastRowCompra = sheetCompra.getLastRow();
  if (lastRowCompra >= 2) {
    sheetCompra.getRange(2, 2, lastRowCompra - 1, 1).clearContent();
    sheetCompra.getRange(2, 5, lastRowCompra - 1, 1).clearContent();
    sheetCompra.getRange(2, 6, lastRowCompra - 1, 2).clearContent();
  }
  
  var threshold = parseFloat(sheetCompra.getRange("J1").getValue());
  if (isNaN(threshold)) {
    threshold = 0;
  }
  
  for (var i = 0; i < totalCompra.length; i++) {
    var totalValue = totalCompra[i][0];
    sheetCompra.getRange(i + 2, 2).setValue(totalValue);
    var label = parseFloat(totalValue) < threshold ? "URGENTE" : "ESTOQUE";
    sheetCompra.getRange(i + 2, 5).setValue(label);
    sheetCompra.getRange(i + 2, 6).setValue(breakdownRel[i][0]);
    sheetCompra.getRange(i + 2, 7).setValue(breakdownTot[i][0]);
  }
  
  var existingFilter = sheetCompra.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  sheetCompra.getRange(1, 1, sheetCompra.getLastRow(), 7).createFilter();
  
  if (notFound.length > 0) {
    SpreadsheetApp.getUi().alert("Os seguintes cadastros não foram encontrados no RELATORIO: " + notFound.join(", "));
  } else {
    SpreadsheetApp.getUi().alert("Compra de fio atualizada com sucesso!");
  }
  
  return "Compra de fio atualizada com sucesso!";
}

/**
 * copyCompraToHistorico: Copia os dados da aba COMPRA DE FIO para a aba HISTORICO, adicionando a data/hora atual.
 */
function copyCompraToHistorico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCompra = ss.getSheetByName("COMPRA DE FIO");
  var historicoSheet = ss.getSheetByName("HISTORICO");
  if (!historicoSheet) {
    historicoSheet = ss.insertSheet("HISTORICO");
  }
  
  var numRowsToCopy = sheetCompra.getLastRow() - 1;
  Logger.log("Número de linhas para copiar: " + numRowsToCopy);
  if (numRowsToCopy > 0) {
    var compData = sheetCompra.getRange(2, 1, numRowsToCopy, 7).getValues();
    var now = new Date();
    var historicoData = compData.map(function(row) {
      return row.concat([now]);
    });
    var lastRowHistorico = historicoSheet.getLastRow();
    var startRowHistorico = lastRowHistorico < 1 ? 1 : lastRowHistorico + 1;
    historicoSheet.getRange(startRowHistorico, 1, historicoData.length, historicoData[0].length).setValues(historicoData);
    Logger.log("Dados copiados para HISTORICO a partir da linha " + startRowHistorico);
  } else {
    Logger.log("Não há linhas para copiar na aba COMPRA DE FIO.");
  }
}

/**
 * atualizarCompraDeFioEHistorico: Executa atualizarCompraDeFio() e, em seguida, copyCompraToHistorico().
 */
function atualizarCompraDeFioEHistorico() {
  atualizarCompraDeFio();
  copyCompraToHistorico();
}

/**
 * showLoginDialog: Exibe o diálogo de login.
 */
function showLoginDialog() {
  var html = HtmlService.createTemplateFromFile("DialogLogin")
    .evaluate()
    .setWidth(350)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "LOGIN");
}

/**
 * processLogin: Valida as credenciais na aba DADOS e, se bem-sucedido, define "loggedUser" e cria o menu.
 */
function processLogin(formData) {
  Logger.log("processLogin: Dados recebidos: " + JSON.stringify(formData));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) {
    throw new Error("A aba DADOS não foi encontrada.");
  }
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 1) {
    throw new Error("Não há usuários cadastrados.");
  }
  var data = sheetDados.getRange(1, 2, lastRow, 2).getValues();
  var valid = false;
  for (var i = 0; i < data.length; i++) {
    var username = data[i][0];
    var password = data[i][1];
    if (username && password) {
      if (username.toString().trim() === formData.username.toString().trim() &&
          password.toString().trim() === formData.password.toString().trim()) {
        valid = true;
        break;
      }
    }
  }
  if (!valid) {
    throw new Error("Credenciais inválidas.");
  }
  PropertiesService.getUserProperties().setProperty("loggedUser", formData.username.toString().trim());
  Logger.log("processLogin: Login efetuado para " + formData.username);
  updateMenus();
  return "Login efetuado com sucesso!";
}

/**
 * getLoggedUser: Retorna o usuário logado.
 */
function getLoggedUser() {
  return PropertiesService.getUserProperties().getProperty("loggedUser");
}

/**
 * showEstoqueSidebar: Abre o formulário de cadastro de estoque na sidebar.
 */
function showEstoqueSidebar() {
  var nextRow = updateUnprotectedRange();
  Logger.log("showEstoqueSidebar: Próxima linha para cadastro: " + nextRow);

  // OTIMIZADO: 1 busca em vez de 4
  var autocompleteData = getAllAutocompleteData();

  var template = HtmlService.createTemplateFromFile("DialogEstoque");
  template.itemList = JSON.stringify(autocompleteData.items);
  template.groupList = JSON.stringify(autocompleteData.groups);
  template.nfList = JSON.stringify(autocompleteData.nfs);
  template.obsList = JSON.stringify(autocompleteData.obs);
  template.currentRow = nextRow;

  var htmlOutput = template.evaluate().setTitle("CADASTRO DE ESTOQUE");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * updateUnprotectedRange: Retorna a próxima linha livre na aba ESTOQUE.
 */
function updateUnprotectedRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  var nextRow = sheet.getLastRow() + 1;
  return nextRow;
}

/**
 * setActiveNextEmptyCell: Seleciona a célula da coluna A que está 15 linhas abaixo da última preenchida.
 */
function setActiveNextEmptyCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet) {
    var nextRow = sheet.getLastRow() + 15;
    sheet.activate();
    sheet.setActiveSelection("A" + nextRow);
    Logger.log("setActiveNextEmptyCell: Célula A" + nextRow + " selecionada.");
  }
}

/**
 * select4RowsBelow: Seleciona a célula da coluna A que está 4 linhas abaixo da última linha preenchida.
 */
function select4RowsBelow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet) {
    var nextRow = sheet.getLastRow() + 4;
    sheet.activate();
    sheet.setActiveSelection("A" + nextRow);
    Logger.log("select4RowsBelow: Célula A" + nextRow + " selecionada.");
  }
}

/**
 * select10RowsBelow: Seleciona a célula da coluna A que está 10 linhas abaixo da última linha preenchida.
 */
function select10RowsBelow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet) {
    var nextRow = sheet.getLastRow() + 10;
    sheet.activate();
    sheet.setActiveSelection("A" + nextRow);
    Logger.log("select10RowsBelow: Célula A" + nextRow + " selecionada.");
  }
}

/**
 * getItemList: Retorna a lista única de itens da aba DADOS (Coluna A).
 */
function getItemList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) return [];
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 1) return [];
  var values = sheetDados.getRange(1, 1, lastRow, 1).getValues().flat();
  var items = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i] && values[i].toString().trim() !== "") {
      items.push(values[i].toString().trim());
    }
  }
  return Array.from(new Set(items));
}

/**
 * getGroupList: Retorna a lista única de grupos da aba DADOS (Coluna D).
 */
function getGroupList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) return [];
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 1) return [];
  var values = sheetDados.getRange(1, 4, lastRow, 1).getValues().flat();
  var groups = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i] && values[i].toString().trim() !== "") {
      groups.push(values[i].toString().trim());
    }
  }
  return Array.from(new Set(groups));
}

/**
 * getNfList: Retorna a lista única de valores da coluna D da aba ESTOQUE (Nota Fiscal/Pedido).
 */
function getNfList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();
  var nfList = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  return Array.from(new Set(nfList));
}

/**
 * getObsList: Retorna a lista única de valores da coluna E da aba ESTOQUE (Cliente/Observações).
 */
function getObsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat();
  var obsList = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  return Array.from(new Set(obsList));
}

/**
 * getMedidasList: Retorna a lista de unidades de medida.
 * Combina opções da aba DADOS (coluna MEDIDAS) com opções já usadas na aba ESTOQUE (coluna C).
 */
function getMedidasList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var medidasSet = new Set();

  // 1. Busca na aba DADOS (coluna MEDIDAS)
  var sheetDados = ss.getSheetByName("DADOS");
  if (sheetDados) {
    var headers = sheetDados.getRange(1, 1, 1, sheetDados.getLastColumn()).getValues()[0];
    var medidasCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var headerUpper = headers[i].toString().toUpperCase().trim();
      if (headerUpper === "MEDIDAS" || headerUpper === "MEDIDA" || headerUpper === "UNIDADE" || headerUpper === "UNIDADES") {
        medidasCol = i + 1;
        break;
      }
    }

    if (medidasCol !== -1) {
      var lastRow = sheetDados.getLastRow();
      if (lastRow >= 2) {
        var values = sheetDados.getRange(2, medidasCol, lastRow - 1, 1).getValues().flat();
        values.forEach(function(v) {
          var val = v.toString().trim();
          if (val !== "") medidasSet.add(val);
        });
      }
    }
  }

  // 2. Busca na aba ESTOQUE (coluna C - Unidade) para incluir opções já usadas
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (sheetEstoque) {
    var lastRowEstoque = sheetEstoque.getLastRow();
    if (lastRowEstoque >= 2) {
      var valuesEstoque = sheetEstoque.getRange(2, 3, lastRowEstoque - 1, 1).getValues().flat();
      valuesEstoque.forEach(function(v) {
        var val = v.toString().trim();
        if (val !== "") medidasSet.add(val);
      });
    }
  }

  // Converte Set para Array e ordena
  var medidasList = Array.from(medidasSet).sort();
  return medidasList;
}

/**
 * getObservacoesList: Retorna a lista de observações.
 * Busca da aba DADOS, coluna F (OBSERVAÇÃO).
 */
function getObservacoesList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var obsSet = new Set();

  // Busca na aba DADOS, coluna F (índice 6)
  var sheetDados = ss.getSheetByName("DADOS");
  if (sheetDados) {
    var lastRow = sheetDados.getLastRow();
    if (lastRow >= 2) {
      // Coluna F = índice 6
      var values = sheetDados.getRange(2, 6, lastRow - 1, 1).getDisplayValues().flat();
      values.forEach(function(v) {
        var val = v.toString().trim();
        if (val !== "") obsSet.add(val);
      });
    }
  }

  // Converte Set para Array e ordena
  var obsList = Array.from(obsSet).sort();
  return obsList;
}

/**
 * normalize: Função auxiliar para normalizar texto.
 */
function normalize(text) {
  if (!text) return "";
  return text.toString().trim().toLowerCase().replace(/\s+/g, " ");
}

/* ================================
   FUNÇÕES DE CACHE E AUTOCOMPLETE
   ================================ */

/**
 * getCachedData: Busca dados no cache ou executa função e armazena no cache.
 */
function getCachedData(key, fetchFunction, ttl) {
  ttl = ttl || 120; // 2 minutos padrão
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log("Cache parse error: " + key);
    }
  }

  var data = fetchFunction();
  try {
    var jsonData = JSON.stringify(data);
    if (jsonData.length < 100000) {
      cache.put(key, jsonData, ttl);
    }
  } catch (e) {
    Logger.log("Cache save error: " + e.message);
  }

  return data;
}

/**
 * invalidateCache: Invalida caches.
 */
function invalidateCache(keys) {
  var cache = CacheService.getScriptCache();
  var keysToInvalidate = typeof keys === 'string' ? [keys] : (keys || []);
  keysToInvalidate.forEach(function(key) { cache.remove(key); });
  cache.remove("autocompleteData");
}

/**
 * invalidateAllAutocompleteCache: Invalida todos os caches de autocomplete.
 */
function invalidateAllAutocompleteCache() {
  invalidateCache(["itemList", "groupList", "nfList", "obsList", "autocompleteData"]);
}

/**
 * getAllAutocompleteData: Busca todos os dados de autocomplete em uma única operação.
 * OTIMIZADO: Usa cache de 10 minutos
 */
function getAllAutocompleteData() {
  return getCachedData("autocompleteData", function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1ª Leitura: DADOS (apenas grupos)
    var sheetDados = ss.getSheetByName("DADOS");
    var groups = [];
    if (sheetDados) {
      var lastRowDados = sheetDados.getLastRow();
      if (lastRowDados >= 1) {
        var dadosData = sheetDados.getRange(1, 4, lastRowDados, 1).getDisplayValues();
        for (var i = 0; i < dadosData.length; i++) {
          if (dadosData[i][0] && dadosData[i][0].toString().trim() !== "") {
            groups.push(dadosData[i][0].toString().trim());
          }
        }
      }
    }

    // 2ª Leitura: ESTOQUE (itens da coluna B e NFs da coluna E)
    // Estrutura: A=Grupo, B=Item, C=Unidade, D=Data, E=NF, F=Obs
    var sheetEstoque = ss.getSheetByName("ESTOQUE");
    var items = [], nfs = [];
    if (sheetEstoque) {
      var lastRowEstoque = sheetEstoque.getLastRow();
      if (lastRowEstoque >= 2) {
        // Lê colunas B até E (4 colunas: B, C, D, E)
        var estoqueData = sheetEstoque.getRange(2, 2, lastRowEstoque - 1, 4).getDisplayValues();
        for (var j = 0; j < estoqueData.length; j++) {
          // Coluna B (índice 0) = Item
          if (estoqueData[j][0] && estoqueData[j][0].toString().trim() !== "") {
            items.push(estoqueData[j][0].toString().trim());
          }
          // Coluna E (índice 3) = NF (já em formato texto com getDisplayValues)
          if (estoqueData[j][3] && estoqueData[j][3].toString().trim() !== "") {
            nfs.push(estoqueData[j][3].toString().trim());
          }
        }
      }
    }

    return {
      items: Array.from(new Set(items)),
      groups: Array.from(new Set(groups)),
      nfs: Array.from(new Set(nfs)),
      medidas: getMedidasList(),
      observacoes: getObservacoesList()
    };
  }, 120); // 2 minutos
}

/**
 * getLastRegistration: Retorna o último registro de um item (data, estoque e grupo).
 * OTIMIZADO: Lê apenas as últimas 2000 linhas da planilha ESTOQUE
 */
function getLastRegistration(item, currentRow) {
  Logger.log("=== getLastRegistration INICIADO ===");
  Logger.log("Item buscado: '" + item + "'");
  Logger.log("CurrentRow: " + currentRow);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) {
    Logger.log("ERRO: Aba ESTOQUE não encontrada!");
    return { lastDate: null, lastStock: 0, lastGroup: null };
  }

  var lastRow = sheetEstoque.getLastRow();
  Logger.log("Última linha da planilha: " + lastRow);
  if (lastRow < 2) {
    Logger.log("Planilha vazia - sem dados");
    return { lastDate: null, lastStock: 0, lastGroup: null };
  }

  // Lê TODA a planilha como TEXTO (getDisplayValues) para evitar problemas de formato
  var startRow = 2;
  var numRows = lastRow - startRow + 1;
  Logger.log("Lendo TODA a planilha - linhas de " + startRow + " até " + lastRow + " (" + numRows + " linhas)");

  // USA getDisplayValues() para forçar conversão para texto
  // Lê 10 colunas (A-J) para incluir o Saldo que está na coluna J
  var data = sheetEstoque.getRange(startRow, 1, numRows, 10).getDisplayValues();
  Logger.log("Usando getDisplayValues() para forçar formato de TEXTO");

  var result = { lastDate: null, lastStock: 0, lastGroup: null };
  var itemNormalized = normalize(item);
  Logger.log("Item normalizado: '" + itemNormalized + "'");

  var encontrados = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    var rowNum = startRow + i;
    if (rowNum >= currentRow) continue;

    var currentItem = data[i][1]; // Coluna B (Item) - agora em formato TEXTO
    if (currentItem && currentItem.toString().trim() !== "") {
      var currentItemNormalized = normalize(currentItem);

      if (currentItemNormalized === itemNormalized) {
        encontrados++;
        result.lastGroup = data[i][0];  // Coluna A (Grupo)
        result.lastDate = data[i][3];   // Coluna D (Data) - como texto
        result.lastStock = data[i][9];  // Coluna J (Saldo) - como texto
        Logger.log("✓ ENCONTRADO na linha " + rowNum);
        Logger.log("  Grupo: '" + result.lastGroup + "'");
        Logger.log("  Data: " + result.lastDate);
        Logger.log("  Estoque: " + result.lastStock);
        Logger.log("  Item raw: '" + currentItem + "'");
        break;
      }
    }
  }

  if (encontrados === 0) {
    Logger.log("✗ NENHUM REGISTRO ENCONTRADO para o item '" + item + "'");
  }

  Logger.log("=== getLastRegistration FINALIZADO ===");
  return result;
}

/**
 * getLastInfoFromDados: Retorna a última informação não vazia da coluna C da aba DADOS para um produto.
 */
function getLastInfoFromDados(product) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("DADOS");
  if (!sheetDados) return "";
  var lastRow = sheetDados.getLastRow();
  if (lastRow < 2) return "";
  var data = sheetDados.getRange(2, 1, lastRow - 1, sheetDados.getLastColumn()).getValues();
  var lastInfo = "";
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === product && data[i][2].toString().trim() !== "") {
      lastInfo = data[i][2];
    }
  }
  return lastInfo;
}

/**
 * showCustomDialog: Exibe um diálogo HTML customizado com uma mensagem.
 */
function showCustomDialog(message) {
  var template = HtmlService.createTemplateFromFile("CustomDialog");
  template.message = message;
  var html = template.evaluate().setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, "AVISO");
}

/**
 * localizarProduto: Abre o diálogo para localizar um produto.
 */
function localizarProduto() {
  var template = HtmlService.createTemplateFromFile("DialogLocalizarProduto");
  template.produtos = JSON.stringify(getProdutosEstoque());
  var htmlOutput = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "LOCALIZAR PRODUTO");
}

/**
 * getProdutosEstoque: Retorna a lista única de produtos da aba ESTOQUE (Coluna B).
 */
function getProdutosEstoque() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("produtosEstoque");
  if (cached) {
    return JSON.parse(cached);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var range = sheet.getRange("B2:B" + lastRow);
  var values = range.getDisplayValues().flat();
  var produtos = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  var unique = Array.from(new Set(produtos));
  cache.put("produtosEstoque", JSON.stringify(unique), 300);
  return unique;
}

/**
 * filtrarProduto: Aplica um filtro na aba ESTOQUE para exibir apenas as linhas cujo valor da coluna B seja igual ao produto.
 */
function filtrarProduto(produto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  var range = sheet.getDataRange();
  var filter = range.createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(produto).build();
  filter.setColumnFilterCriteria(2, criteria);
}

/**
 * mostrarTodos: Remove o filtro, ordena a aba ESTOQUE pela data (Coluna C) e seleciona uma célula.
 */
function mostrarTodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).sort({ column: 3, ascending: true });
  }
  setActiveNextEmptyCell();
}

/**
 * abrirDialogRelatorioEstoque: Abre o diálogo para definir a faixa de data do relatório.
 */
function abrirDialogRelatorioEstoque() {
  var html = HtmlService.createTemplateFromFile("DialogRelatorioEstoque")
      .evaluate()
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "RELATÓRIO DE ESTOQUE");
}

/**
 * gerarRelatorioEstoque: Gera o relatório geral para o período definido.
 */
function gerarRelatorioEstoque(dataInicio, dataFim) {
  Logger.log("gerarRelatorioEstoque: Início " + dataInicio + " - Fim " + dataFim);
  var partsInicio = dataInicio.split("/");
  var partsFim = dataFim.split("/");
  var startDate = new Date(partsInicio[2], partsInicio[1] - 1, partsInicio[0], 0, 0, 0);
  var endDate = new Date(partsFim[2], partsFim[1] - 1, partsFim[0], 23, 59, 59);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  var lastColumn = sheetEstoque.getLastColumn();
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, lastColumn);
  var dataValues = dataRange.getValues();
  
  var filtered = dataValues.filter(function(row) {
    var dt = new Date(row[2]);
    return dt >= startDate && dt <= endDate;
  });
  
  var grupos = {};
  filtered.forEach(function(row) {
    var prod = row[1];
    if (!grupos[prod]) {
      grupos[prod] = row;
    } else {
      var currentDate = new Date(row[2]);
      var storedDate = new Date(grupos[prod][2]);
      if (currentDate > storedDate) {
        grupos[prod] = row;
      }
    }
  });
  
  var reportData = [];
  for (var prod in grupos) {
    var row = grupos[prod];
    reportData.push([prod, row[8], row[4], row[2]]);
  }
  
  reportData.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });
  
  var sheetRelatorio = ss.getSheetByName("RELATORIO");
  if (!sheetRelatorio) {
    sheetRelatorio = ss.insertSheet("RELATORIO");
    sheetRelatorio.getRange("J1").setValue(0);
  }
  var threshold = parseFloat(sheetRelatorio.getRange("J1").getValue());
  if (isNaN(threshold)) {
    threshold = 0;
  }
  
  for (var i = 0; i < reportData.length; i++) {
    var novoSaldo = parseFloat(reportData[i][1]);
    if (novoSaldo < threshold) {
      reportData[i].push("URGENTE");
    } else {
      reportData[i].push("ESTOQUE");
    }
  }
  
  sheetRelatorio.clearContents();
  sheetRelatorio.getRange("J1").setValue(threshold);
  sheetRelatorio.getRange(1, 1, 1, 5).setValues([["PRODUTO", "NOVO SALDO", "OBS", "DATA/HORA", "STATUS"]]);
  if (reportData.length > 0) {
    sheetRelatorio.getRange(2, 1, reportData.length, 5).setValues(reportData);
  }
  
  var relFilter = sheetRelatorio.getFilter();
  if (relFilter) {
    relFilter.remove();
  }
  sheetRelatorio.getRange(1, 1, sheetRelatorio.getLastRow(), 5).createFilter();
  
  Logger.log("gerarRelatorioEstoque: Relatório gerado com " + reportData.length + " registros.");
  return "Relatório gerado com sucesso!";
}

/**
 * abrirDialogRelatorioPorGrupo: Abre o diálogo para definir o grupo do relatório.
 */
function abrirDialogRelatorioPorGrupo() {
  var template = HtmlService.createTemplateFromFile("DialogRelatorioPorGrupo");
  template.grupos = JSON.stringify(getGruposEstoque());
  var htmlOutput = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "RELATÓRIO POR GRUPO");
}

/**
 * getGruposEstoque: Retorna os grupos únicos da aba ESTOQUE (Coluna A).
 */
function getGruposEstoque() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  var grupos = values.filter(function(v) {
    return v.toString().trim() !== "";
  });
  return Array.from(new Set(grupos));
}

/**
 * gerarRelatorioPorGrupo: Gera o relatório para um grupo específico.
 */
function gerarRelatorioPorGrupo(grupoSelecionado) {
  Logger.log("gerarRelatorioPorGrupo: Grupo selecionado: " + grupoSelecionado);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  var lastColumn = sheetEstoque.getLastColumn();
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, lastColumn);
  var dataValues = dataRange.getValues();
  
  var filtered = dataValues.filter(function(row) {
    return row[0].toString().trim() === grupoSelecionado;
  });
  
  var gruposItens = {};
  filtered.forEach(function(row) {
    var item = row[1];
    if (!gruposItens[item]) {
      gruposItens[item] = row;
    } else {
      var currentDate = new Date(row[2]);
      var storedDate = new Date(gruposItens[item][2]);
      if (currentDate > storedDate) {
        gruposItens[item] = row;
      }
    }
  });
  
  var reportData = [];
  for (var item in gruposItens) {
    var row = gruposItens[item];
    reportData.push([row[0], row[1], row[8], row[2]]);
  }
  
  reportData.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });
  
  var sheetRelatorio = ss.getSheetByName("RELATORIO POR GRUPO DE ITEM");
  if (!sheetRelatorio) {
    sheetRelatorio = ss.insertSheet("RELATORIO POR GRUPO DE ITEM");
  }
  sheetRelatorio.clearContents();
  sheetRelatorio.getRange(1, 1, 1, 4).setValues([["GRUPO", "ITEM", "NOVO SALDO", "DATA/HORA"]]);
  if (reportData.length > 0) {
    sheetRelatorio.getRange(2, 1, reportData.length, 4).setValues(reportData);
  }
  
  Logger.log("gerarRelatorioPorGrupo: Relatório gerado com " + reportData.length + " registros.");
  return "Relatório por grupo gerado com sucesso!";
}

/**
 * showListagemEstoqueSidebar: Abre a sidebar para a listagem de estoque.
 */
function showListagemEstoqueSidebar() {
  var template = HtmlService.createTemplateFromFile("DialogListagemEstoque");
  template.produtos = JSON.stringify(getProdutosEstoque());
  var htmlOutput = template.evaluate().setTitle("LISTAGEM DE ESTOQUE");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/* ================================
   NOVAS FUNÇÕES: Estoque por Período e Limpar Filtro
   ================================ */

/**
 * abrirDialogEstoquePorPeriodo: Abre um diálogo para que o usuário informe as datas de início e fim.
 */
function abrirDialogEstoquePorPeriodo() {
  var html = HtmlService.createTemplateFromFile("DialogEstoquePorPeriodo")
    .evaluate()
    .setWidth(350)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, "Filtrar Estoque por Período");
}

/**
 * filtrarEstoquePorPeriodo: Copia as linhas da aba ESTOQUE, cuja data na coluna C
 * esteja entre dataInicio e dataFim (formato dd/mm/yyyy), e as cola na aba "FILTRO POR PERIODO".
 * Antes de colar, apaga o conteúdo anterior da aba.
 */
function filtrarEstoquePorPeriodo(dataInicio, dataFim) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  // Divide as datas informadas (formato dd/mm/yyyy) em partes e cria objetos Date.
  var partsInicio = dataInicio.split("/");
  var partsFim = dataFim.split("/");
  if (partsInicio.length !== 3 || partsFim.length !== 3) {
    throw new Error("Formato de data inválido. Use dd/mm/yyyy");
  }
  
  var startDate = new Date(
    parseInt(partsInicio[2], 10), 
    parseInt(partsInicio[1], 10) - 1, 
    parseInt(partsInicio[0], 10)
  );
  
  var endDate = new Date(
    parseInt(partsFim[2], 10), 
    parseInt(partsFim[1], 10) - 1, 
    parseInt(partsFim[0], 10),
    23, 59, 59, 999
  );
  
  // Obtém os dados da aba ESTOQUE (assumindo que a primeira linha é o cabeçalho)
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, sheetEstoque.getLastColumn());
  var dataValues = dataRange.getValues();
  
  // Prepara a aba de destino "FILTRO POR PERIODO"
  var sheetFiltro = ss.getSheetByName("FILTRO POR PERIODO");
  if (!sheetFiltro) {
    sheetFiltro = ss.insertSheet("FILTRO POR PERIODO");
  } else {
    // Apaga todo o conteúdo da aba, inclusive formatação e filtros antigos.
    sheetFiltro.clear();
  }
  
  // Copia o cabeçalho da aba ESTOQUE para a aba "FILTRO POR PERIODO"
  var header = sheetEstoque.getRange(1, 1, 1, sheetEstoque.getLastColumn()).getValues();
  sheetFiltro.getRange(1, 1, 1, header[0].length).setValues(header);
  
  var targetData = [];
  
  // Percorre cada linha e copia as que tiverem data na coluna C (índice 2) dentro do período
  for (var i = 0; i < dataValues.length; i++) {
    var row = dataValues[i];
    var dateValue = row[2]; // Coluna C
    // Verifica se o valor é uma data válida
    if (!(dateValue instanceof Date)) continue;
    if (dateValue >= startDate && dateValue <= endDate) {
      targetData.push(row);
    }
  }
  
  // Copia os dados filtrados para a aba "FILTRO POR PERIODO", a partir da linha 2
  if (targetData.length > 0) {
    sheetFiltro.getRange(2, 1, targetData.length, targetData[0].length).setValues(targetData);
  }
  
  // Ordena os dados (exceto o cabeçalho) pela coluna C em ordem crescente
  var targetLastRow = sheetFiltro.getLastRow();
  if (targetLastRow > 1) {
    sheetFiltro.getRange(2, 1, targetLastRow - 1, sheetFiltro.getLastColumn())
              .sort({ column: 3, ascending: true });
  }
  
  return "Dados do período de " + dataInicio + " a " + dataFim + " foram copiados para a aba 'FILTRO POR PERIODO'.";
}

/**
 * limparFiltroEstoque: Remove o filtro da aba ESTOQUE, ordena pela coluna C (datas) de forma ascendente
 * e seleciona a célula 4 linhas abaixo da última linha preenchida.
 */
function limparFiltroEstoque() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESTOQUE");
  if (!sheet) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 3, ascending: true });
  }
  
  select4RowsBelow();
  
  return "Filtro removido e planilha ordenada por data.";
}

/**
 * convertDateFormat: Converte uma data do formato dd/mm/yyyy para mm/dd/yyyy.
 */
function convertDateFormat(dateStr) {
  var parts = dateStr.split("/");
  if (parts.length !== 3) throw new Error("Data inválida: " + dateStr);
  return parts[1] + "/" + parts[0] + "/" + parts[2];
}

/* ================================
   NOVAS FUNÇÕES: Cores Desatualizadas
   ================================ */

/**
 * processCoresDesatualizadas: A partir de uma data informada (formato dd/mm/yyyy),
 * procura na aba ESTOQUE as linhas marcadas de vermelho (verificando a cor da primeira célula)
 * e, para cada item (coluna A), agrupa os registros cuja data (coluna C) seja >= data informada.
 * Para cada item, pega as últimas 5 linhas (baseadas na data) e copia somente o valor da coluna B
 * para a aba "CORES DESATUALIZADAS", a partir da coluna E, com o cabeçalho "CORES DESATUALIZADAS".
 * Além disso, na célula F1 da aba "CORES DESATUALIZADAS" é exibida a data informada.
 */
function processCoresDesatualizadas(startDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Converte a data informada (dd/mm/yyyy) para objeto Date
  var startDate = parseDateBR(startDateStr);
  if (!startDate || isNaN(startDate)) {
    throw new Error("Data de início inválida: use o formato dd/mm/yyyy.");
  }
  
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  var lastCol = sheetEstoque.getLastColumn();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  // Lê os dados e os backgrounds (exceto o cabeçalho)
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, lastCol);
  var dadosValues = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();
  
  // Cria um objeto para agrupar os registros por item (chave em lowercase)
  var grupos = {};
  
  // Percorre cada linha para verificar se está marcada de vermelho na primeira célula
  // e se a data na coluna C é >= startDate.
  for (var i = 0; i < dadosValues.length; i++) {
    var bg = backgrounds[i][0].toLowerCase();
    if (bg !== "red" && bg !== "#ff0000") continue;
    
    var row = dadosValues[i];
    var item = row[0] ? row[0].toString().trim() : "";
    if (item === "") continue;
    
    // Data na coluna C (índice 2)
    var dataCell = row[2];
    if (!(dataCell instanceof Date)) {
      dataCell = parseDateBR(dataCell);
    }
    if (!(dataCell instanceof Date) || isNaN(dataCell)) continue;
    if (dataCell < startDate) continue;
    
    var key = item.toLowerCase();
    if (!grupos[key]) {
      grupos[key] = [];
    }
    grupos[key].push({ row: row, date: dataCell });
  }
  
  // Para cada item, ordena os registros por data decrescente e pega as últimas 5 linhas
  var resultados = [];
  for (var key in grupos) {
    if (grupos.hasOwnProperty(key)) {
      var registros = grupos[key];
      registros.sort(function(a, b) {
        return b.date - a.date;
      });
      var ultimos5 = registros.slice(0, 5);
      ultimos5.reverse(); // Exibe do mais antigo para o mais recente
      for (var j = 0; j < ultimos5.length; j++) {
        // Copia somente o valor da coluna B (índice 1)
        resultados.push([ultimos5[j].row[1]]);
      }
    }
  }
  
  // --- Parte 3: Copia os resultados para a aba CORES DESATUALIZADAS na coluna E ---
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) {
    sheetCores = ss.insertSheet("CORES DESATUALIZADAS");
  } else {
    sheetCores.clear();
  }
  
  // Define o cabeçalho na coluna E
  var header = [["CORES DESATUALIZADAS"]];
  sheetCores.getRange(1, 5, 1, header[0].length).setValues(header);
  
  if (resultados.length > 0) {
    var resultRange = sheetCores.getRange(2, 5, resultados.length, resultados[0].length);
    resultRange.setValues(resultados);
    // Define o fundo de todas as linhas copiadas como vermelho
    resultRange.setBackground("red");
  }
  
  // Exibe a data informada na célula F1
  sheetCores.getRange("F1").setValue(startDateStr);
  
  return "Valores da coluna B para os itens em vermelho a partir de " + startDateStr + " foram copiados para a aba 'CORES DESATUALIZADAS' na coluna E, e a data foi registrada em F1.";
}

/**
 * parseDateBR: Converte uma string no formato dd/mm/yyyy para um objeto Date.
 */
function parseDateBR(dateStr) {
  if (typeof dateStr === 'string') {
    var parts = dateStr.split("/");
    if (parts.length === 3) {
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
  }
  return new Date(dateStr);
}

/* ================================
   NOVAS FUNÇÕES: Cores Desatualizadas - Diálogo
   ================================ */
function showCoresDesatualizadasDialog() {
  var html = HtmlService.createTemplateFromFile("DialogCoresDesatualizadas")
    .evaluate()
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, "Filtrar Cores Desatualizadas");
}
/**
 * processRepeticoesCoresDesatualizadas:
 * Lê a data definida na célula F1 da aba CORES DESATUALIZADAS e os cadastros inseridos na coluna E.
 * Em seguida, na aba ESPELHO DO ESTOQUE, para cada cadastro encontrado em CORES DESATUALIZADAS (coluna E),
 * busca todos os registros cujo cadastro (coluna A) corresponda (comparação case-insensitive)
 * e cuja data na coluna B seja menor ou igual à data definida.
 * Os registros encontrados são copiados para a aba CORES DESATUALIZADAS, onde:
 *  - Coluna A: Cadastro
 *  - Coluna B: Data (da ESPELHO DO ESTOQUE)
 *  - Coluna C: Informação (da ESPELHO DO ESTOQUE, coluna E)
 */
function processRepeticoesCoresDesatualizadas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtém a aba CORES DESATUALIZADAS
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) throw new Error("A aba CORES DESATUALIZADAS não foi encontrada.");
  
  // Lê a data definida na célula F1 (deve estar no formato dd/mm/yyyy)
  var dataF1Raw = sheetCores.getRange("F1").getValue();
  var dataF1 = parseDateBR(dataF1Raw.toString());
  if (!dataF1 || isNaN(dataF1)) {
    throw new Error("Data em F1 inválida. Certifique-se de que está no formato dd/mm/yyyy.");
  }
  
  // Lê os cadastros da coluna E (a partir da linha 2)
  var lastRowCores = sheetCores.getLastRow();
  if (lastRowCores < 2) throw new Error("Não há cadastros na coluna E da aba CORES DESATUALIZADAS.");
  var rangeCadastros = sheetCores.getRange(2, 5, lastRowCores - 1, 1);
  var cadastrosData = rangeCadastros.getValues();
  
  // Armazena os cadastros únicos (usando lowercase para comparação)
  var cadastros = {};
  for (var i = 0; i < cadastrosData.length; i++) {
    var val = cadastrosData[i][0];
    if (val && val.toString().trim() !== "") {
      cadastros[val.toString().trim().toLowerCase()] = val.toString().trim();
    }
  }
  
  // Acessa a aba ESPELHO DO ESTOQUE
  var sheetEspelho = ss.getSheetByName("ESPELHO DO ESTOQUE");
  if (!sheetEspelho) throw new Error("A aba ESPELHO DO ESTOQUE não foi encontrada.");
  var lastRowEspelho = sheetEspelho.getLastRow();
  var lastColEspelho = sheetEspelho.getLastColumn();
  if (lastRowEspelho < 2) throw new Error("Não há dados na aba ESPELHO DO ESTOQUE.");
  var rangeEspelho = sheetEspelho.getRange(2, 1, lastRowEspelho - 1, lastColEspelho);
  var espelhoValues = rangeEspelho.getValues();
  
  var resultados = [];
  
  // Percorre cada registro da aba ESPELHO DO ESTOQUE
  // Assumindo que:
  // - Coluna A: Cadastro
  // - Coluna B: Data
  // - Coluna E: Informação
  for (var j = 0; j < espelhoValues.length; j++) {
    var row = espelhoValues[j];
    var cadastroEspelho = row[0] ? row[0].toString().trim() : "";
    if (cadastroEspelho === "") continue;
    var key = cadastroEspelho.toLowerCase();
    if (!(key in cadastros)) continue;
    
    // Data na ESPELHO DO ESTOQUE (coluna B, índice 1)
    var dataEspelho = row[1];
    if (!(dataEspelho instanceof Date)) {
      dataEspelho = parseDateBR(dataEspelho);
    }
    if (!(dataEspelho instanceof Date) || isNaN(dataEspelho)) continue;
    // Considera registros cuja data seja <= data definida (F1)
    if (dataEspelho > dataF1) continue;
    
    // Extrai:
    // Coluna A: Cadastro, Coluna B: Data, Coluna C: Informação (da coluna E da aba ESPELHO)
    var info = row[4]; // Coluna E
    resultados.push([cadastroEspelho, dataEspelho, info]);
  }
  
  // Agora, vamos limpar a aba CORES DESATUALIZADAS e inserir esses resultados nas colunas A, B e C.
  sheetCores.clear();
  // Define o cabeçalho
  sheetCores.getRange(1, 1, 1, 3).setValues([["Cadastro", "Data", "Informação"]]);
  
  if (resultados.length > 0) {
    sheetCores.getRange(2, 1, resultados.length, 3).setValues(resultados);
  }
  
  return "Registros até " + dataF1Raw + " foram copiados para a aba 'CORES DESATUALIZADAS' (Colunas A:C).";
}

/**
 * parseDateBR: Converte uma string no formato dd/mm/yyyy para um objeto Date.
 */
function parseDateBR(dateStr) {
  if (typeof dateStr === 'string') {
    var parts = dateStr.split("/");
    if (parts.length === 3) {
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
  }
  return new Date(dateStr);
}
/**
 * processConsultaAtualizacoes:
 * Lê a data de corte da célula F1 da aba CORES DESATUALIZADAS e os itens (cadastros) da coluna E.
 * Em seguida, na aba ESPELHO DO ESTOQUE, para cada registro, se o cadastro (coluna B)
 * corresponder (case-insensitive) a um dos itens e se a data (coluna C) for menor ou igual à data de corte,
 * extrai:
 *    - Cadastro (da coluna B),
 *    - Data (da coluna C),
 *    - Valor (da coluna E).
 * Os resultados são escritos na aba CORES DESATUALIZADAS, a partir da linha 2, nas colunas A, B e C.
 */
function processConsultaAtualizacoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Acessa a aba CORES DESATUALIZADAS
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) throw new Error("A aba CORES DESATUALIZADAS não foi encontrada.");
  
  // Lê a data de corte da célula F1 (formato dd/mm/yyyy)
  var cutoffRaw = sheetCores.getRange("F1").getValue();
  var cutoffDate = parseDateBR(cutoffRaw.toString());
  if (!cutoffDate || isNaN(cutoffDate)) {
    throw new Error("Data de corte em F1 inválida. Certifique-se de que está no formato dd/mm/yyyy.");
  }
  
  // Lê os itens (cadastros) da coluna E, a partir da linha 2
  var lastRowCores = sheetCores.getLastRow();
  if (lastRowCores < 2) throw new Error("Não há cadastros na coluna E da aba CORES DESATUALIZADAS.");
  var rangeItems = sheetCores.getRange(2, 5, lastRowCores - 1, 1);
  var itemsData = rangeItems.getValues();
  
  var queryItems = [];
  for (var i = 0; i < itemsData.length; i++) {
    var val = itemsData[i][0];
    if (val && val.toString().trim() !== "") {
      queryItems.push(val.toString().trim().toLowerCase());
    }
  }
  
  // Acessa a aba ESPELHO DO ESTOQUE
  var sheetEspelho = ss.getSheetByName("ESPELHO DO ESTOQUE");
  if (!sheetEspelho) throw new Error("A aba ESPELHO DO ESTOQUE não foi encontrada.");
  var lastRowEspelho = sheetEspelho.getLastRow();
  var lastColEspelho = sheetEspelho.getLastColumn();
  if (lastRowEspelho < 2) throw new Error("Não há dados na aba ESPELHO DO ESTOQUE.");
  
  var rangeEspelho = sheetEspelho.getRange(2, 1, lastRowEspelho - 1, lastColEspelho);
  var espelhoValues = rangeEspelho.getValues();
  
  var resultados = [];
  
  // Percorre os registros da aba ESPELHO DO ESTOQUE
  // Assumindo que:
  // - Coluna B (índice 1) contém o cadastro (para comparação)
  // - Coluna C (índice 2) contém a data
  // - Coluna E (índice 4) contém o valor
  for (var j = 0; j < espelhoValues.length; j++) {
    var row = espelhoValues[j];
    var cadastroEspelho = row[1] ? row[1].toString().trim() : "";
    if (cadastroEspelho === "") continue;
    var cadastroKey = cadastroEspelho.toLowerCase();
    if (queryItems.indexOf(cadastroKey) === -1) continue;
    
    var dataEspelho = row[2];
    if (!(dataEspelho instanceof Date)) {
      dataEspelho = parseDateBR(dataEspelho);
    }
    if (!(dataEspelho instanceof Date) || isNaN(dataEspelho)) continue;
    // Só traz se a data for menor ou igual à data de corte
    if (dataEspelho > cutoffDate) continue;
    
    var valor = row[4]; // Valor da coluna E
    resultados.push([cadastroEspelho, dataEspelho, valor]);
  }
  
  // --- Escreve os resultados na aba CORES DESATUALIZADAS nas colunas A, B e C ---
  // Limpa as colunas A, B e C
  sheetCores.getRange("A:C").clearContent();
  // Define o cabeçalho
  sheetCores.getRange(1, 1, 1, 3).setValues([["Cadastro", "Data", "Valor"]]);
  
  if (resultados.length > 0) {
    sheetCores.getRange(2, 1, resultados.length, 3).setValues(resultados);
  }
  
  return "Consulta atualizações concluída. Registros encontrados: " + resultados.length;
}

/**
 * parseDateBR: Converte uma string no formato dd/mm/yyyy para um objeto Date.
 */
function parseDateBR(dateStr) {
  if (typeof dateStr === 'string') {
    var parts = dateStr.split("/");
    if (parts.length === 3) {
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
  }
  return new Date(dateStr);
}
/**
 * consultaAtualizacao: Para cada item informado via formulário (até 10),
 * filtra os registros da aba "ESPELHO DO ESTOQUE" dos últimos 20 dias e
 * retorna os 5 registros mais recentes individualmente, gravando os resultados
 * na aba "CORES DESATUALIZADAS".
 *
 * Supõe-se que na aba "ESPELHO DO ESTOQUE":
 *   - Coluna A: Item
 *   - Coluna B: Data
 *   - Coluna D: Valor
 *   - Coluna E: Valor Adicional
 *
 * Os resultados serão escritos na aba "CORES DESATUALIZADAS" com:
 *   - Coluna A: Item
 *   - Coluna B: Data
 *   - Coluna C: Valor (coluna D do ESPELHO)
 *   - Coluna D: Valor adicional (coluna E do ESPELHO)
 */
function consultaAtualizacao(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Calcula o período: hoje e 20 dias atrás
  var today = new Date();
  var twentyDaysAgo = new Date(today.getTime() - 20 * 24 * 3600 * 1000);
  Logger.log("Hoje: " + today);
  Logger.log("20 dias atrás: " + twentyDaysAgo);
  
  // Obtém os itens informados no formulário (até 10) e converte para lowercase
  var itemsSelecionados = [];
  for (var i = 1; i <= 10; i++) {
    var value = formData["item" + i];
    if (value && value.trim() !== "") {
      itemsSelecionados.push(value.trim().toLowerCase());
    }
  }
  Logger.log("Itens selecionados: " + itemsSelecionados.join(", "));
  
  // Acessa a aba "ESPELHO DO ESTOQUE"
  var sheetEspelho = ss.getSheetByName("ESPELHO DO ESTOQUE");
  if (!sheetEspelho) throw new Error("A aba 'ESPELHO DO ESTOQUE' não foi encontrada.");
  
  var lastRowEspelho = sheetEspelho.getLastRow();
  if (lastRowEspelho < 2) throw new Error("Não há dados na aba 'ESPELHO DO ESTOQUE'.");
  
  var dadosRange = sheetEspelho.getRange(2, 1, lastRowEspelho - 1, sheetEspelho.getLastColumn());
  var dadosValues = dadosRange.getValues();
  Logger.log("Total de registros lidos (Espelho): " + dadosValues.length);
  
  var results = [];
  
  // Para cada item informado, filtra os registros que atendem ao critério
  itemsSelecionados.forEach(function(item) {
    // Filtra registros onde:
    // - O valor da coluna A (item) é igual ao item informado (case-insensitive)
    // - A data (coluna B) está entre 20 dias atrás e hoje
    var registros = dadosValues.filter(function(row) {
      var itemNome = row[0] ? row[0].toString().trim().toLowerCase() : "";
      var dataCell = row[1];
      if (!(dataCell instanceof Date)) {
        dataCell = new Date(dataCell);
      }
      return itemNome === item && dataCell >= twentyDaysAgo && dataCell <= today;
    });
    
    // Ordena os registros por data do mais recente para o mais antigo
    registros.sort(function(a, b) {
      return new Date(b[1]) - new Date(a[1]);
    });
    
    // Pega os 5 registros mais recentes (se existirem) e inverte a ordem para exibição cronológica crescente
    var ultimos5 = registros.slice(0, 5);
    ultimos5.reverse();
    
    // Para cada registro, extrai:
    // - Item (coluna A), Data (coluna B), Valor (coluna D – índice 3) e Valor adicional (coluna E – índice 4)
    ultimos5.forEach(function(row) {
      results.push([ row[0], row[1], row[3], row[4] ]);
    });
  });
  
  // Escreve os resultados na aba "CORES DESATUALIZADAS"
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) {
    sheetCores = ss.insertSheet("CORES DESATUALIZADAS");
  } else {
    sheetCores.clear(); // Limpa todo o conteúdo existente
  }
  
  // Define o cabeçalho e insere os resultados
  sheetCores.getRange(1, 1, 1, 4).setValues([["Item", "Data", "Valor", "Valor Adicional"]]);
  if (results.length > 0) {
    sheetCores.getRange(2, 1, results.length, 4).setValues(results);
  }
  
  return "Consulta de atualização (20 dias) concluída na aba 'CORES DESATUALIZADAS'.";
}
/**
 * showConsultaAtualizacaoSidebar: Abre uma sidebar com um formulário para inserir 15 itens.
 */
function showConsultaAtualizacaoSidebar() {
  var template = HtmlService.createTemplateFromFile("DialogConsultaAtualizacao");
  var htmlOutput = template.evaluate().setTitle("Consulta Atualização (15 Itens)").setWidth(350).setHeight(500);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
/**
 * gerarListagemCoresDesatualizadas: Para cada item informado via formulário (até 10),
 * busca na aba ESTOQUE os últimos 5 registros onde:
 *   - Coluna B da aba ESTOQUE → Coluna A da aba CORES DESATUALIZADAS,
 *   - Coluna C da aba ESTOQUE → Coluna B,
 *   - Coluna E da aba ESTOQUE → Coluna C.
 * Os registros são inseridos um item abaixo do outro na aba CORES DESATUALIZADAS.
 * Além disso, a função retorna uma tabela HTML com os resultados para ser exibida na sidebar.
 */
function gerarListagemCoresDesatualizadas(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) throw new Error("A aba ESTOQUE não foi encontrada.");
  
  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) throw new Error("Não há dados na aba ESTOQUE.");
  
  // Lê os dados da aba ESTOQUE (considerando que a primeira linha é cabeçalho)
  var estoqueData = sheetEstoque.getRange(2, 1, lastRow - 1, sheetEstoque.getLastColumn()).getValues();
  
  // Obtém os itens do formulário (até 10 itens)
  var items = [];
  for (var i = 1; i <= 10; i++) {
    var campo = formData["item" + i];
    if (campo && campo.trim() !== "") {
      items.push(campo.trim().toLowerCase());
    }
  }
  if (items.length === 0) throw new Error("Informe pelo menos um item.");
  
  var resultados = [];
  
  // Para cada item, filtra os registros (baseado na coluna B da aba ESTOQUE)
  items.forEach(function(item) {
    var registros = estoqueData.filter(function(row) {
      return row[1] && row[1].toString().trim().toLowerCase() === item;
    });
    // Seleciona os últimos 5 registros e inverte para ordem cronológica crescente
    var ultimos5 = registros.slice(-5);
    ultimos5.reverse();
    ultimos5.forEach(function(row) {
      resultados.push([ row[1], row[2], row[4] ]);
    });
  });
  
  // Escreve os resultados na aba CORES DESATUALIZADAS
  var sheetCores = ss.getSheetByName("CORES DESATUALIZADAS");
  if (!sheetCores) {
    sheetCores = ss.insertSheet("CORES DESATUALIZADAS");
  } else {
    sheetCores.clear();
  }
  
  sheetCores.getRange(1, 1, 1, 3).setValues([["Produto (Coluna B)", "Data (Coluna C)", "Valor Extra (Coluna E)"]]);
  if (resultados.length > 0) {
    sheetCores.getRange(2, 1, resultados.length, 3).setValues(resultados);
  }
  
  // Constrói uma tabela HTML com os resultados para exibição na sidebar
  var htmlOutput = "<h3>Resultados:</h3>";
  if (resultados.length > 0) {
    htmlOutput += "<table border='1' style='border-collapse:collapse; width:100%;'>";
    htmlOutput += "<tr><th>Produto</th><th>Data</th><th>Valor Extra</th></tr>";
    resultados.forEach(function(row) {
      htmlOutput += "<tr>";
      htmlOutput += "<td>" + row[0] + "</td>";
      htmlOutput += "<td>" + row[1] + "</td>";
      htmlOutput += "<td>" + row[2] + "</td>";
      htmlOutput += "</tr>";
    });
    htmlOutput += "</table>";
  } else {
    htmlOutput += "<p>Nenhum registro encontrado.</p>";
  }
  
  return htmlOutput;
}
/**
 * Abre a sidebar para a listagem de cores desatualizadas.
 */
function abrirDialogListagemCores() {
  var template = HtmlService.createTemplateFromFile("DialogListagemCores");
  template.espelhoItems = JSON.stringify(getEspelhoItemList());
  var htmlOutput = template.evaluate()
    .setTitle("Listagem de Cores Desatualizadas")
    .setWidth(350)
    .setHeight(500);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getEspelhoItemList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ESPELHO DO ESTOQUE");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var values = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
  var items = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i] && values[i].toString().trim() !== "") {
      items.push(values[i].toString().trim());
    }
  }
  return Array.from(new Set(items));
}
/**
 * Abre a sidebar “Estoque 3 Meses” corretamente,
 * passando o HtmlOutput para showSidebar().
 */
function showEstoque3MesesSidebar() {
  var template = HtmlService.createTemplateFromFile("DialogEstoque3Meses");
  template.itemList    = JSON.stringify(getItemList());
  // formata a data de hoje em dd/MM/yyyy
  template.currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var html = template.evaluate()
                .setTitle("Consumo 3 Meses")
                .setWidth(500)     // largura maior
                .setHeight(600);   // altura maior
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Soma o consumo dos últimos 3 meses para até 20 itens,
 * atualiza a aba CONSUMO 3 MESES e retorna o HTML da tabela.
 */
/**
 * processEstoque3Meses: Soma o consumo dos últimos 3 meses para até 20 itens
 * e retorna HTML incluindo a última data de lançamento de cada um.
 */
function processEstoque3Meses(formData) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEsp  = ss.getSheetByName("ESPELHO DO ESTOQUE");
  var sheetBase = ss.getSheetByName("BASE TINGIMENTO");
  var sheetCons = ss.getSheetByName("CONSUMO 3 MESES") 
                  || ss.insertSheet("CONSUMO 3 MESES");

  // Verifica existência das abas
  if (!sheetEsp)  throw new Error("Aba 'ESPELHO DO ESTOQUE' não foi encontrada.");
  if (!sheetBase) throw new Error("Aba 'BASE TINGIMENTO' não foi encontrada.");

  // 1) Coleta até 20 itens do formulário
  var items = [];
  for (var i = 1; i <= 20; i++) {
    var v = formData['item' + i];
    if (v && v.trim()) items.push(v.trim());
  }
  if (items.length === 0) {
    throw new Error("Informe ao menos um item.");
  }

  // 2) Soma consumo dos últimos 3 meses (coluna D do espelho)
  var today         = new Date();
  var threeMonthsAgo = new Date(today.getFullYear(), today.getMonth() - 3, today.getDate());
  var espData = sheetEsp
                  .getRange(2, 1, sheetEsp.getLastRow() - 1, sheetEsp.getLastColumn())
                  .getValues();

  var consumos = items.map(function(it) {
    var total = 0;
    espData.forEach(function(r) {
      if (r[0] && r[0].toString().trim().toLowerCase() === it.toLowerCase()) {
        var dt = (r[1] instanceof Date) ? r[1] : new Date(r[1]);
        if (dt >= threeMonthsAgo && dt <= today) {
          total += parseFloat(r[3]) || 0;
        }
      }
    });
    return { item: it, total: total };
  });

  // 3) Carrega padrões da BASE TINGIMENTO e fallback (linha 3)
  var baseAll = sheetBase
                  .getRange(2, 1, sheetBase.getLastRow() - 1, sheetBase.getLastColumn())
                  .getValues();
  var base = baseAll.map(function(r) {
    return {
      pattern: String(r[0]).trim().toLowerCase(),
      values:  r.slice(2).map(function(c) { return parseFloat(c) || 0; })
    };
  });
  var fallbackValues = baseAll[1].slice(2).map(function(c) { return parseFloat(c) || 0; });

  // 4) Para cada consumo, encontra padrão e pega 8 valores mais próximos
  consumos.forEach(function(c) {
    var text = c.item.toLowerCase();
    var match = base
      .filter(function(b) { return b.pattern && text.indexOf(b.pattern) !== -1; })
      .sort(function(a, b) { return b.pattern.length - a.pattern.length; })[0];

    var source = match ? match.values : fallbackValues;
    var diffs = source
      .map(function(v) { return { v: v, d: Math.abs(v - c.total) }; })
      .sort(function(a, b) { return a.d - b.d; })
      .slice(0, 8)
      .map(function(o) { return o.v; });

    for (var j = 0; j < 8; j++) {
      c['s' + (j + 1)] = diffs[j] !== undefined ? diffs[j] : '';
    }
  });

  // 5) Grava na planilha CONSUMO 3 MESES (mantendo formatação)
  var headers = ['Item', 'Total últimos 3 MESES'];
  for (var k = 1; k <= 8; k++) headers.push('LOTE TINGIMENTO ' + k);

  // Limpa só o conteúdo, sem afetar formatação
  sheetCons.clearContents();
  sheetCons.getRange(1, 1, 1, headers.length).setValues([headers]);

  var out = consumos.map(function(c) {
    var row = [c.item, c.total];
    for (var m = 1; m <= 8; m++) row.push(c['s' + m]);
    return row;
  });
  if (out.length) {
    sheetCons.getRange(2, 1, out.length, headers.length).setValues(out);
  }

  // Centraliza e redimensiona colunas
  var lastRow = sheetCons.getLastRow();
  var lastCol = headers.length;
  sheetCons.getRange(1, 1, lastRow, lastCol).setHorizontalAlignment("center");
  sheetCons.autoResizeColumns(1, lastCol);

  // 6) Monta e retorna tabela HTML para a sidebar
  var html = '<table><tr>';
  headers.forEach(function(h) { html += '<th>' + h + '</th>'; });
  html += '</tr>';
  out.forEach(function(r) {
    html += '<tr>';
    r.forEach(function(cell) { html += '<td>' + cell + '</td>'; });
    html += '</tr>';
  });
  html += '</table>';
  return html;
}

/**
 * gerarListagemVermelho: Para cada produto (coluna B) na aba ESTOQUE,
 * coleta apenas o registro mais recente (não contendo 'ACERTO' ou 'ATUALIZAÇÃO')
 * varrendo de baixo para cima, e grava em "CORES DESATUALIZADAS" como texto simples.
 */
function gerarListagemVermelho() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ESTOQUE');
  if (!sheet) throw new Error("A aba 'ESTOQUE' não foi encontrada.");

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error("Não há dados na aba 'ESTOQUE'.");

  // Lê todo o intervalo de dados como texto
  var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getDisplayValues();

  // Regex para filtrar explicações indesejadas
  var regex = /acerto|atualiza[cç][ãa]o/;

  // Agrupa apenas o registro mais recente por produto (coluna B)
  var grupos = {};
  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var obs = row[4];
    var textoObs = obs ? obs.toLowerCase() : '';
    if (regex.test(textoObs)) continue; // ignora se conter keywords

    var item = row[1]; // coluna B
    // Se ainda não capturou o último registro desse item, faz push
    if (!grupos[item]) {
      grupos[item] = row; // armazena o único registro desejado
    }
  }

  // Prepara resultados: extrai cada registro único
  var resultados = [];
  for (var produto in grupos) {
    resultados.push(grupos[produto]);
  }

  // Grava na aba CORES DESATUALIZADAS
  var outSheet = ss.getSheetByName('CORES DESATUALIZADAS');
  if (!outSheet) throw new Error("A aba 'CORES DESATUALIZADAS' não foi encontrada.");
  outSheet.clearContents();

  // Cabeçalho conforme layout da aba ESTOQUE (texto)
  var header = ['Grupo','Item','Data','NF','Obs','Saldo Anterior','Entrada','Saída','Novo Saldo','Data/Hora','Usuário'];
  outSheet.getRange(1, 1, 1, header.length).setValues([header]);

  // Insere registros únicos, mantendo formato texto
  if (resultados.length) {
    var targetRange = outSheet.getRange(2, 1, resultados.length, lastCol);
    targetRange.setValues(resultados);
    targetRange.setNumberFormat('@');
  }

  return 'CORES DESATUALIZADAS atualizada com ' + resultados.length + ' registro(s) mais recente(s) por produto.';
}

/**
 * DEBUG: Busca e mostra TODOS os registros de um item na ESTOQUE
 * Use esta função para verificar se o item existe e como está escrito
 */
function debugBuscarItemNaEstoque(itemBuscado) {
  Logger.log("=== DEBUG: Buscando '" + itemBuscado + "' na ESTOQUE ===");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  var lastRow = sheetEstoque.getLastRow();

  Logger.log("Total de linhas na ESTOQUE: " + lastRow);

  // Lê TODA a coluna B (Item) - USA getDisplayValues para forçar TEXTO
  var data = sheetEstoque.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  Logger.log("Lendo colunas A (Grupo) e B (Item) como TEXTO");

  var itemNormalized = normalize(itemBuscado);
  Logger.log("Item normalizado buscado: '" + itemNormalized + "'");

  var encontrados = [];
  var semelhantes = [];

  for (var i = 0; i < data.length; i++) {
    var grupo = data[i][0];
    var item = data[i][1];

    if (item && item.toString().trim() !== "") {
      var itemStr = item.toString();
      var itemNorm = normalize(itemStr);

      // Exato
      if (itemNorm === itemNormalized) {
        encontrados.push({
          linha: i + 2,
          grupo: grupo,
          item: itemStr,
          itemNorm: itemNorm
        });
      }
      // Semelhante (contém parte do nome)
      else if (itemNorm.indexOf(itemNormalized) >= 0 || itemNormalized.indexOf(itemNorm) >= 0) {
        semelhantes.push({
          linha: i + 2,
          grupo: grupo,
          item: itemStr,
          itemNorm: itemNorm
        });
      }
    }
  }

  Logger.log("\n========================================");
  Logger.log("RESULTADOS:");
  Logger.log("========================================");

  if (encontrados.length > 0) {
    Logger.log("\n✓ ENCONTRADOS " + encontrados.length + " registros EXATOS:");
    for (var j = 0; j < Math.min(10, encontrados.length); j++) {
      var reg = encontrados[j];
      Logger.log("  Linha " + reg.linha + ": Grupo='" + reg.grupo + "' | Item='" + reg.item + "'");
    }
    if (encontrados.length > 10) {
      Logger.log("  ... e mais " + (encontrados.length - 10) + " registros");
    }
  } else {
    Logger.log("\n✗ NENHUM registro EXATO encontrado");
  }

  if (semelhantes.length > 0) {
    Logger.log("\n≈ ENCONTRADOS " + semelhantes.length + " registros SEMELHANTES:");
    for (var k = 0; k < Math.min(10, semelhantes.length); k++) {
      var sem = semelhantes[k];
      Logger.log("  Linha " + sem.linha + ": Grupo='" + sem.grupo + "' | Item='" + sem.item + "'");
    }
    if (semelhantes.length > 10) {
      Logger.log("  ... e mais " + (semelhantes.length - 10) + " registros");
    }
  }

  Logger.log("\n========================================");
  Logger.log("TOTAL: " + encontrados.length + " exatos, " + semelhantes.length + " semelhantes");
  Logger.log("========================================");

  return {
    exatos: encontrados.length,
    semelhantes: semelhantes.length,
    amostraExatos: encontrados.slice(0, 5),
    amostraSemelhantes: semelhantes.slice(0, 5)
  };
}

// ========================================
// WEB APP FUNCTIONS
// ========================================

/**
 * loginUser: Autentica usuário no Web App
 * Verifica usuário e senha na sheet USUÁRIOS
 * Formato da sheet: Coluna A = Usuário, Coluna B = Senha
 */
function loginUser(username, password) {
  try {
    // Verifica credenciais no sheet USUÁRIOS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetUsuarios = ss.getSheetByName("USUÁRIOS");

    if (!sheetUsuarios) {
      return { success: false, message: "Sheet USUÁRIOS não encontrada" };
    }

    var lastRow = sheetUsuarios.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: "Nenhum usuário cadastrado" };
    }

    var data = sheetUsuarios.getRange(2, 1, lastRow - 1, 2).getValues();

    // Remove espaços extras e converte para string
    var usernameClean = String(username).trim();
    var passwordClean = String(password).trim();

    Logger.log("Tentativa de login - Usuário: '" + usernameClean + "' | Senha: '" + passwordClean + "'");

    for (var i = 0; i < data.length; i++) {
      var dbUser = String(data[i][0]).trim();
      var dbPass = String(data[i][1]).trim();

      Logger.log("Comparando com linha " + (i+2) + " - Usuário: '" + dbUser + "' | Senha: '" + dbPass + "'");

      if (dbUser === usernameClean && dbPass === passwordClean) {
        // Login bem-sucedido
        Logger.log("Login bem-sucedido!");
        PropertiesService.getUserProperties().setProperty("loggedUser", usernameClean);
        return { success: true, user: usernameClean };
      }
    }

    Logger.log("Login falhou - credenciais não encontradas");
    return { success: false, message: "Usuário ou senha incorretos" };
  } catch (error) {
    Logger.log("Erro loginUser: " + error);
    return { success: false, message: "Erro ao fazer login: " + error.message };
  }
}

/**
 * debugUsuarios: Função para debugar usuários cadastrados
 * Execute esta função no Apps Script para ver os usuários na aba USUÁRIOS
 */
function debugUsuarios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsuarios = ss.getSheetByName("USUÁRIOS");

  if (!sheetUsuarios) {
    Logger.log("❌ Sheet USUÁRIOS não encontrada!");
    return;
  }

  var lastRow = sheetUsuarios.getLastRow();
  Logger.log("📊 Total de linhas: " + lastRow);

  if (lastRow < 2) {
    Logger.log("❌ Nenhum usuário cadastrado (sheet vazia)");
    return;
  }

  var data = sheetUsuarios.getRange(1, 1, lastRow, 2).getValues();

  Logger.log("\n=== USUÁRIOS CADASTRADOS ===");
  for (var i = 0; i < data.length; i++) {
    var user = String(data[i][0]);
    var pass = String(data[i][1]);
    Logger.log("Linha " + (i+1) + ":");
    Logger.log("  Usuário: '" + user + "' (length: " + user.length + ")");
    Logger.log("  Senha: '" + pass + "' (length: " + pass.length + ")");
  }
  Logger.log("=========================\n");
}

/**
 * logoutUser: Remove autenticação do usuário
 */
function logoutUser() {
  PropertiesService.getUserProperties().deleteProperty("loggedUser");
  return { success: true };
}

/**
 * getDashboardData: Retorna dados estatísticos para o dashboard
 */
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");
    var sheetDados = ss.getSheetByName("DADOS");

    if (!sheetEstoque || !sheetDados) {
      return { success: false };
    }

    // Total de itens únicos
    var dataEstoque = sheetEstoque.getRange(2, 2, sheetEstoque.getLastRow() - 1, 1).getValues();
    var uniqueItems = new Set();
    dataEstoque.forEach(function(row) {
      if (row[0]) uniqueItems.add(row[0]);
    });

    // Total de grupos
    var dataGrupos = sheetDados.getRange(2, 4, sheetDados.getLastRow() - 1, 1).getValues();
    var uniqueGroups = new Set();
    dataGrupos.forEach(function(row) {
      if (row[0]) uniqueGroups.add(row[0]);
    });

    // Entradas e saídas de hoje
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var dataMovimentos = sheetEstoque.getRange(2, 3, sheetEstoque.getLastRow() - 1, 5).getValues();

    var recentEntries = 0;
    var recentExits = 0;

    dataMovimentos.forEach(function(row) {
      var dataMovimento = new Date(row[0]);
      dataMovimento.setHours(0, 0, 0, 0);

      if (dataMovimento.getTime() === today.getTime()) {
        var entrada = parseFloat(row[3]) || 0;
        var saida = parseFloat(row[4]) || 0;

        if (entrada > 0) recentEntries++;
        if (saida > 0) recentExits++;
      }
    });

    return {
      success: true,
      totalItems: uniqueItems.size,
      totalGroups: uniqueGroups.size,
      recentEntries: recentEntries,
      recentExits: recentExits
    };
  } catch (error) {
    Logger.log("Erro getDashboardData: " + error);
    return { success: false };
  }
}

/**
 * insertGroup: Wrapper para inserir grupo via web app
 */
function insertGroup(grupo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetDados = ss.getSheetByName("DADOS");

    if (!sheetDados) {
      return { success: false, message: "Sheet DADOS não encontrada" };
    }

    // Verifica se grupo já existe
    var lastRow = sheetDados.getLastRow();
    var existingGroups = sheetDados.getRange(2, 4, Math.max(1, lastRow - 1), 1).getValues();

    for (var i = 0; i < existingGroups.length; i++) {
      if (normalize(existingGroups[i][0]) === normalize(grupo)) {
        return { success: false, message: "Grupo já existe" };
      }
    }

    // Adiciona grupo
    var nextRow = sheetDados.getLastRow() + 1;
    sheetDados.getRange(nextRow, 4).setValue(grupo);

    // Invalida cache
    invalidateCache();

    return { success: true, message: "Grupo adicionado com sucesso" };
  } catch (error) {
    Logger.log("Erro insertGroup: " + error);
    return { success: false, message: "Erro ao inserir grupo: " + error.message };
  }
}

/**
 * buscarProduto: Wrapper para localizar produto via web app
 * Retorna dados ordenados do mais novo para o mais antigo, com cores das linhas
 */
function buscarProduto(item, dataInicio, dataFim) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: "Nenhum dado encontrado" };
    }

    var data = sheetEstoque.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var dataValues = sheetEstoque.getRange(2, 1, lastRow - 1, 13).getValues(); // Para pegar datas como Date
    var backgrounds = sheetEstoque.getRange(2, 1, lastRow - 1, 13).getBackgrounds();
    var results = [];
    var itemNormalized = normalize(item);

    // Filtra por item e data (se fornecida)
    for (var i = 0; i < data.length; i++) {
      var currentItem = normalize(data[i][1]);
      if (currentItem.indexOf(itemNormalized) >= 0) {
        // Pega a data como objeto Date (não string) - Coluna D (índice 3)
        var dataMovimento = dataValues[i][3];

        // Verifica filtro de data
        if (dataInicio && dataFim) {
          var inicio = new Date(dataInicio);
          var fim = new Date(dataFim);
          inicio.setHours(0, 0, 0, 0);
          fim.setHours(23, 59, 59, 999);

          if (dataMovimento < inicio || dataMovimento > fim) {
            continue; // Pula este registro
          }
        }

        // Adiciona dados com cor de fundo
        results.push({
          row: data[i],
          background: backgrounds[i][0], // Cor da primeira coluna (toda linha tem mesma cor)
          date: dataMovimento // Para ordenação (usa Date object real)
        });
      }
    }

    if (results.length === 0) {
      return { success: false, message: "Produto não encontrado" };
    }

    // Ordena do mais novo para o mais antigo (descendente)
    results.sort(function(a, b) {
      return b.date.getTime() - a.date.getTime();
    });

    // Extrai apenas as linhas e cores
    var sortedRows = [];
    var rowColors = [];
    for (var j = 0; j < results.length; j++) {
      sortedRows.push(results[j].row);
      rowColors.push(results[j].background);
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: sortedRows,
        colors: rowColors
      }
    };
  } catch (error) {
    Logger.log("Erro buscarProduto: " + error);
    return { success: false, message: "Erro ao buscar produto: " + error.message };
  }
}

/**
 * testeConexao: Função de teste para verificar se a comunicação está funcionando
 */
function testeConexao() {
  return { success: true, message: "Conexão funcionando!", timestamp: new Date().toString() };
}

/**
 * carregarTodosOsDadosEstoque: Carrega TODOS os dados do estoque de uma vez
 * Para filtros instantâneos no lado do cliente
 */
function carregarTodosOsDadosEstoque() {
  // Garantia absoluta de que NUNCA retornará null
  var resultado = { success: false, message: "Erro desconhecido", data: [] };

  try {
    Logger.log("=== carregarTodosOsDadosEstoque INICIADO ===");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      Logger.log("ERRO: Sheet ESTOQUE não encontrada");
      resultado.message = "Sheet ESTOQUE não encontrada";
      return resultado;
    }

    var lastRow = sheetEstoque.getLastRow();
    var lastCol = sheetEstoque.getLastColumn();
    Logger.log("Total de linhas: " + lastRow + ", Total de colunas: " + lastCol);

    if (lastRow < 2) {
      Logger.log("Planilha vazia - retornando array vazio");
      return { success: true, data: [], headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"] };
    }

    // Detecta automaticamente o número de colunas (11 ou 13)
    var numCols = lastCol >= 13 ? 13 : 11;
    Logger.log("Usando estrutura de " + numCols + " colunas");

    // Limita a 2.000 registros mais recentes para evitar exceder limite de transferência
    // Com 43k+ linhas, 10k era grande demais. 2k é suficiente para buscas rápidas.
    var maxRowsToLoad = 2000;
    var totalDataRows = lastRow - 1; // Exclui header

    var startRow, numRowsToLoad;
    if (totalDataRows <= maxRowsToLoad) {
      // Se tem menos de 2000 linhas, carrega tudo
      startRow = 2;
      numRowsToLoad = totalDataRows;
    } else {
      // Se tem mais de 2000, carrega apenas as últimas 2000 (mais recentes)
      startRow = lastRow - maxRowsToLoad + 1;
      numRowsToLoad = maxRowsToLoad;
    }

    Logger.log("Carregando " + numRowsToLoad + " linhas (da linha " + startRow + " até " + (startRow + numRowsToLoad - 1) + ")");

    var data = sheetEstoque.getRange(startRow, 1, numRowsToLoad, numCols).getDisplayValues();
    var dataValues = sheetEstoque.getRange(startRow, 1, numRowsToLoad, numCols).getValues();
    var backgrounds = sheetEstoque.getRange(startRow, 1, numRowsToLoad, numCols).getBackgrounds();

    Logger.log("Dados carregados - processando...");
    var allData = [];

    // Define índice da coluna Data baseado no número de colunas
    var dateColIndex = (numCols === 13) ? 3 : 2;  // Coluna D (índice 3) para 13 cols, C (índice 2) para 11 cols

    for (var i = 0; i < data.length; i++) {
      allData.push({
        row: data[i],
        date: dataValues[i][dateColIndex],
        background: backgrounds[i][0]
      });
    }

    Logger.log("Processamento concluído - " + allData.length + " registros");

    // Headers baseados no número de colunas
    var headers = (numCols === 13)
      ? ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"]
      : ["Grupo", "Item", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Alterado Em", "Alterado Por"];

    Logger.log("=== carregarTodosOsDadosEstoque FINALIZADO COM SUCESSO ===");

    return {
      success: true,
      data: allData,
      headers: headers
    };
  } catch (error) {
    Logger.log("ERRO CRÍTICO em carregarTodosOsDadosEstoque: " + error);
    Logger.log("Stack trace: " + error.stack);
    resultado.message = "Erro ao carregar dados: " + error.message;
    resultado.error = error.toString();
    return resultado;
  }

  // Garantia final - nunca deve chegar aqui
  return resultado;
}

/**
 * mostrarTodosProdutos: Retorna todos os produtos do estoque
 * Ordenados do mais novo para o mais antigo, com cores
 */
function mostrarTodosProdutos(dataInicio, dataFim) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: "Nenhum dado encontrado" };
    }

    var data = sheetEstoque.getRange(2, 1, Math.min(5000, lastRow - 1), 11).getDisplayValues();
    var dataValues = sheetEstoque.getRange(2, 1, Math.min(5000, lastRow - 1), 11).getValues(); // Para pegar datas como Date
    var backgrounds = sheetEstoque.getRange(2, 1, Math.min(5000, lastRow - 1), 11).getBackgrounds();
    var results = [];

    // Filtra por data (se fornecida)
    for (var i = 0; i < data.length; i++) {
      var dataMovimento = dataValues[i][2]; // Usa Date object real

      if (dataInicio && dataFim) {
        var inicio = new Date(dataInicio);
        var fim = new Date(dataFim);
        inicio.setHours(0, 0, 0, 0);
        fim.setHours(23, 59, 59, 999);

        if (dataMovimento < inicio || dataMovimento > fim) {
          continue; // Pula este registro
        }
      }

      results.push({
        row: data[i],
        background: backgrounds[i][0],
        date: dataMovimento // Usa Date object real
      });
    }

    // Ordena do mais novo para o mais antigo (descendente)
    results.sort(function(a, b) {
      return b.date.getTime() - a.date.getTime();
    });

    // Extrai apenas as linhas e cores
    var sortedRows = [];
    var rowColors = [];
    for (var j = 0; j < results.length; j++) {
      sortedRows.push(results[j].row);
      rowColors.push(results[j].background);
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: sortedRows,
        colors: rowColors
      }
    };
  } catch (error) {
    Logger.log("Erro mostrarTodosProdutos: " + error);
    return { success: false, message: "Erro ao buscar produtos: " + error.message };
  }
}

/**
 * filtrarEstoquePorPeriodo: Aplica filtro por período na planilha
 */
function filtrarEstoquePorPeriodo(dataInicio, dataFim) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    // Remove filtro existente
    var filter = sheetEstoque.getFilter();
    if (filter) {
      filter.remove();
    }

    // Cria novo filtro
    var lastRow = sheetEstoque.getLastRow();
    var lastCol = sheetEstoque.getLastColumn();
    var range = sheetEstoque.getRange(1, 1, lastRow, lastCol);

    var newFilter = range.createFilter();

    // Aplica filtro na coluna C (Data)
    var inicio = new Date(dataInicio);
    var fim = new Date(dataFim);

    var criteria = SpreadsheetApp.newFilterCriteria()
      .whenDateAfter(inicio)
      .whenDateBefore(fim)
      .build();

    newFilter.setColumnFilterCriteria(3, criteria);

    return { success: true, message: "Filtro aplicado com sucesso" };
  } catch (error) {
    Logger.log("Erro filtrarEstoquePorPeriodo: " + error);
    return { success: false, message: "Erro ao filtrar: " + error.message };
  }
}

/**
 * getEstoque3Meses: Retorna estoque dos últimos 3 meses
 */
function getEstoque3Meses() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: "Nenhum dado encontrado" };
    }

    var data = sheetEstoque.getRange(2, 1, lastRow - 1, 11).getValues();
    var results = [];

    var today = new Date();
    var threeMonthsAgo = new Date();
    threeMonthsAgo.setMonth(today.getMonth() - 3);

    for (var i = 0; i < data.length; i++) {
      var dataMovimento = new Date(data[i][2]);
      if (dataMovimento >= threeMonthsAgo) {
        results.push(data[i]);
      }
    }

    if (results.length === 0) {
      return { success: false, message: "Nenhum movimento nos últimos 3 meses" };
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: results
      }
    };
  } catch (error) {
    Logger.log("Erro getEstoque3Meses: " + error);
    return { success: false, message: "Erro ao buscar estoque: " + error.message };
  }
}
