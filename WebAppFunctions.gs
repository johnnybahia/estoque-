// ========================================
// WEB APP ADDITIONAL FUNCTIONS
// Funções adicionais do Web App separadas para melhor organização
// ========================================

// ========================================
// OTIMIZAÇÕES DE CACHE - TTLs aumentados
// ========================================

/**
 * Constantes de cache - TTLs otimizados para melhor performance
 * TTLs maiores reduzem chamadas à planilha sem prejudicar limites do Google
 */
var CACHE_TTL_OPT = {
  AUTOCOMPLETE: 600,    // 10 minutos para dados de autocomplete
  DASHBOARD: 120,       // 2 minutos para dashboard
  ITEM_INDEX: 300,      // 5 minutos para índice de itens
  DEFAULT: 300          // 5 minutos padrão
};

/**
 * getCachedDataOpt: Versão otimizada de cache com TTL maior
 */
function getCachedDataOpt(key, fetchFunction, ttl) {
  ttl = ttl || CACHE_TTL_OPT.DEFAULT;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log("Cache parse error, regenerating: " + key);
    }
  }

  var data = fetchFunction();
  try {
    var jsonData = JSON.stringify(data);
    if (jsonData.length < 100000) {
      cache.put(key, jsonData, ttl);
      Logger.log("Cache saved: " + key + " (TTL: " + ttl + "s)");
    }
  } catch (e) {
    Logger.log("Cache save error: " + e.message);
  }

  return data;
}

/**
 * getDashboardDataCached: Dashboard com cache de 2 minutos
 */
function getDashboardDataCached() {
  return getCachedDataOpt("dashboardData", function() {
    return _fetchDashboardData();
  }, CACHE_TTL_OPT.DASHBOARD);
}

function _fetchDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");

  if (!sheetEstoque) {
    return { totalItems: 0, totalGroups: 0, recentEntries: 0, recentExits: 0 };
  }

  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) {
    return { totalItems: 0, totalGroups: 0, recentEntries: 0, recentExits: 0 };
  }

  // Lê apenas colunas necessárias
  var data = sheetEstoque.getRange(2, 1, lastRow - 1, 9).getValues();

  var items = new Set();
  var groups = new Set();
  var todayEntries = 0;
  var todayExits = 0;
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 0; i < data.length; i++) {
    if (data[i][1]) items.add(data[i][1].toString().trim());
    if (data[i][0]) groups.add(data[i][0].toString().trim());

    var dataMovimento = new Date(data[i][3]);
    dataMovimento.setHours(0, 0, 0, 0);

    if (dataMovimento.getTime() === today.getTime()) {
      var entrada = parseFloat(data[i][7]) || 0;
      var saida = parseFloat(data[i][8]) || 0;
      if (entrada > 0) todayEntries++;
      if (saida > 0) todayExits++;
    }
  }

  return {
    totalItems: items.size,
    totalGroups: groups.size,
    recentEntries: todayEntries,
    recentExits: todayExits
  };
}

/**
 * getItemIndexOpt: Índice otimizado para busca O(1)
 */
function getItemIndexOpt() {
  return getCachedDataOpt("itemIndexOpt", function() {
    return _buildItemIndexOpt();
  }, CACHE_TTL_OPT.ITEM_INDEX);
}

function _buildItemIndexOpt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) return {};

  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) return {};

  var data = sheetEstoque.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
  var index = {};

  for (var i = 0; i < data.length; i++) {
    var item = data[i][1] ? data[i][1].toString().trim().toLowerCase() : null;
    if (item) {
      index[item] = {
        row: i + 2,
        group: data[i][0],
        date: data[i][3],
        stock: data[i][9]
      };
    }
  }

  Logger.log("Item index built with " + Object.keys(index).length + " entries");
  return index;
}

/**
 * getLastRegistrationOpt: Busca O(1) usando índice em cache
 */
function getLastRegistrationOpt(item) {
  if (!item) return { lastDate: null, lastStock: 0, lastGroup: null };

  var itemNormalized = item.toString().trim().toLowerCase();
  var index = getItemIndexOpt();

  if (index[itemNormalized]) {
    var cached = index[itemNormalized];
    return {
      lastDate: cached.date,
      lastStock: cached.stock,
      lastGroup: cached.group
    };
  }

  return { lastDate: null, lastStock: 0, lastGroup: null };
}

/**
 * invalidateCacheOpt: Invalida caches otimizados
 */
function invalidateCacheOpt() {
  var cache = CacheService.getScriptCache();
  cache.remove("autocompleteData");
  cache.remove("itemIndexOpt");
  cache.remove("dashboardData");
  cache.remove("itemHistoryIndex");
  Logger.log("Optimized caches invalidated");
}

// ========================================
// ÍNDICE DE HISTÓRICO - Últimos 20 registros por item
// ========================================

/**
 * Constante: número máximo de registros por item no índice
 */
var MAX_HISTORY_PER_ITEM = 20;

/**
 * getItemHistoryIndex: Retorna o índice de histórico completo do cache ou reconstrói
 * O índice contém os últimos 20 registros de cada item para acesso instantâneo
 */
function getItemHistoryIndex() {
  return getCachedDataOpt("itemHistoryIndex", function() {
    return _buildItemHistoryIndex();
  }, CACHE_TTL_OPT.ITEM_INDEX);
}

/**
 * _buildItemHistoryIndex: Constrói o índice de histórico com últimos 20 registros por item
 * Estrutura: { "item_normalizado": { info: {...}, history: [{row, date, background}, ...] } }
 */
function _buildItemHistoryIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstoque = ss.getSheetByName("ESTOQUE");
  if (!sheetEstoque) return {};

  var lastRow = sheetEstoque.getLastRow();
  if (lastRow < 2) return {};

  // Lê todos os dados de uma vez (mais eficiente)
  var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, 13);
  var data = dataRange.getDisplayValues();
  var backgrounds = dataRange.getBackgrounds();

  var index = {};

  // Processa todos os registros
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var itemOriginal = row[1] ? row[1].toString().trim() : null;
    if (!itemOriginal) continue;

    var itemKey = itemOriginal.toLowerCase();
    var dateStr = row[3]; // Coluna D (Data)
    var rowDate = dateStr ? new Date(dateStr) : new Date(0);
    var background = backgrounds[i][0] || null; // Cor de fundo da linha

    // Inicializa o item no índice se não existir
    if (!index[itemKey]) {
      index[itemKey] = {
        info: {
          item: itemOriginal,
          group: row[0],
          lastDate: dateStr,
          lastStock: row[9], // Coluna J (Saldo)
          unidade: row[2]    // Coluna C (Unidade)
        },
        history: []
      };
    }

    // Adiciona ao histórico
    index[itemKey].history.push({
      row: row,
      date: rowDate.getTime(),
      background: background
    });

    // Atualiza info se este registro é mais recente
    var currentLastDate = new Date(index[itemKey].info.lastDate);
    if (rowDate > currentLastDate) {
      index[itemKey].info.lastDate = dateStr;
      index[itemKey].info.lastStock = row[9];
      index[itemKey].info.group = row[0];
      index[itemKey].info.unidade = row[2];
    }
  }

  // Para cada item, ordena por data (mais recente primeiro) e limita a 20 registros
  var itemCount = 0;
  for (var key in index) {
    index[key].history.sort(function(a, b) {
      return b.date - a.date;
    });

    // Limita ao máximo de registros
    if (index[key].history.length > MAX_HISTORY_PER_ITEM) {
      index[key].history = index[key].history.slice(0, MAX_HISTORY_PER_ITEM);
    }
    itemCount++;
  }

  Logger.log("Item history index built: " + itemCount + " items with up to " + MAX_HISTORY_PER_ITEM + " records each");
  return index;
}

/**
 * getItemHistory: Retorna o histórico dos últimos 20 registros de um item específico
 * @param {string} item - Nome do item para buscar
 * @return {object} - { success, data: { headers, rows, colors }, info: { lastDate, lastStock, group } }
 */
function getItemHistory(item) {
  if (!item || item.trim() === '') {
    return { success: false, message: "Item não informado" };
  }

  var itemKey = item.toString().trim().toLowerCase();
  var index = getItemHistoryIndex();

  if (!index[itemKey]) {
    return { success: false, message: "Item não encontrado no índice" };
  }

  var itemData = index[itemKey];
  var headers = ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"];

  // Ordena por data (mais novo primeiro) antes de retornar
  var sortedHistory = itemData.history.slice().sort(function(a, b) {
    return b.date - a.date;
  });

  var rows = [];
  var colors = [];

  for (var i = 0; i < sortedHistory.length; i++) {
    rows.push(sortedHistory[i].row);
    colors.push(sortedHistory[i].background);
  }

  return {
    success: true,
    data: {
      headers: headers,
      rows: rows,
      colors: colors
    },
    info: itemData.info
  };
}

/**
 * getItemHistoryForClient: Versão otimizada para retornar índice parcial ao cliente
 * Retorna apenas itens que correspondem à busca parcial (para autocomplete inteligente)
 * @param {string} searchTerm - Termo de busca parcial
 * @param {number} limit - Limite de itens a retornar (default 50)
 */
function getItemHistoryForClient(searchTerm, limit) {
  limit = limit || 50;
  var index = getItemHistoryIndex();
  var result = {};
  var count = 0;

  var searchNormalized = searchTerm ? searchTerm.toString().trim().toLowerCase() : '';

  for (var key in index) {
    if (searchNormalized === '' || key.indexOf(searchNormalized) >= 0) {
      result[key] = index[key];
      count++;
      if (count >= limit) break;
    }
  }

  return result;
}

/**
 * preloadItemHistoryIndex: Pré-carrega o índice de histórico (chamado após login)
 * Útil para garantir que o cache está pronto antes do primeiro uso
 */
function preloadItemHistoryIndex() {
  var startTime = new Date().getTime();
  var index = getItemHistoryIndex();
  var elapsed = new Date().getTime() - startTime;
  var itemCount = Object.keys(index).length;
  Logger.log("Item history index preloaded: " + itemCount + " items in " + elapsed + "ms");
  return { success: true, itemCount: itemCount, loadTime: elapsed };
}

// ========================================
// FUNÇÕES ORIGINAIS DO WEB APP
// ========================================

/**
 * processEstoqueWebApp: Wrapper da função processEstoque para o Web App
 * Retorna sucesso/erro em formato JSON
 */
function processEstoqueWebApp(formData) {
  try {
    // Validações
    var entrada = parseFloat(formData.entrada) || 0;
    var saida = parseFloat(formData.saida) || 0;

    if (entrada > 0 && saida > 0) {
      return { success: false, message: "Não é possível ter entrada e saída ao mesmo tempo" };
    }

    if (entrada === 0 && saida === 0) {
      return { success: false, message: "Informe uma entrada ou saída" };
    }

    // Chama a função original processEstoque
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");
    var now = new Date();
    var nextRow = sheetEstoque.getLastRow() + 1;

    PropertiesService.getScriptProperties().setProperty("editingViaScript", "true");

    // Recupera último registro para cálculo de saldo e data
    var lastReg = getLastRegistration(formData.item, nextRow);
    var previousSaldo = parseFloat(lastReg.lastStock) || 0;
    var newSaldo = previousSaldo + parseFloat(formData.entrada) - parseFloat(formData.saida);

    // Nova estrutura com Unidade de Medida (após Item) e Valor (após Saldo)
    var rowData = [
      formData.group,              // A: Grupo
      formData.item,               // B: Item
      formData.unidade || '',      // C: Unidade de Medida (NOVO)
      now,                         // D: Data
      formData.nf || '',           // E: NF
      formData.obs || '',          // F: Obs
      previousSaldo,               // G: Saldo Anterior
      formData.entrada,            // H: Entrada
      formData.saida,              // I: Saída
      newSaldo,                    // J: Saldo
      parseFloat(formData.valorUnitario) || 0,  // K: Valor Unitário (NOVO)
      now,                         // L: Alterado Em
      getLoggedUser()              // M: Alterado Por
    ];

    sheetEstoque.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    Logger.log("processEstoqueWebApp: Dados inseridos na linha " + nextRow);

    PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
    backupEstoqueData();

    var warningMessage = null;

    // Verifica se passou mais de 20 dias desde a última data de registro
    if (lastReg.lastDate) {
      var lastDate = new Date(lastReg.lastDate);
      var diffDays = (now.getTime() - lastDate.getTime()) / (1000 * 3600 * 24);
      if (diffDays > 20) {
        // Verifica coluna F (obs) por palavras-chave
        var textoObs = formData.obs ? formData.obs.toString().toLowerCase() : "";
        var temKeyword = /acerto|atualiza[cç][ãa]o/.test(textoObs);
        // Se não conter 'acerto' ou 'atualização', pinta de vermelho
        if (!temKeyword) {
          var lastColumn = sheetEstoque.getLastColumn();
          sheetEstoque.getRange(nextRow, 1, 1, lastColumn).setBackground("red");
          warningMessage = "⚠️ PRODUTO DESATUALIZADO (ÚLTIMA ATUALIZAÇÃO HÁ MAIS DE 20 DIAS). POR FAVOR, ATUALIZAR URGENTE.";
        }
      }
    }

    // Verifica se houve ENTRADA de estoque - aviso para atualização
    if (parseFloat(formData.entrada) > 0) {
      var lastColumn = sheetEstoque.getLastColumn();
      sheetEstoque.getRange(nextRow, 1, 1, lastColumn).setBackground("yellow");
      warningMessage = "⚠️ ENTRADA DE ESTOQUE REGISTRADA!\n\nÉ NECESSÁRIO ATUALIZAR O ESTOQUE DESTE ITEM PARA EVITAR FUROS DE ESTOQUE.\n\nRealize uma contagem física e registre uma atualização completa do saldo.";
    }

    // Invalida caches (padrão e otimizado)
    invalidateCache();
    invalidateCacheOpt();

    return {
      success: true,
      message: warningMessage || "Estoque processado com sucesso!",
      warning: warningMessage ? true : false
    };
  } catch (error) {
    PropertiesService.getScriptProperties().deleteProperty("editingViaScript");
    Logger.log("Erro processEstoqueWebApp: " + error);
    return { success: false, message: "Erro ao processar estoque: " + error.message };
  }
}

/**
 * gerarListagemEstoqueWebApp: Wrapper para gerar listagem de estoque via web app
 */
function gerarListagemEstoqueWebApp(formData) {
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
    var results = [];

    var filterGroup = formData.group ? normalize(formData.group) : null;
    var filterItem = formData.item ? normalize(formData.item) : null;

    for (var i = 0; i < data.length; i++) {
      var match = true;

      if (filterGroup) {
        if (normalize(data[i][0]).indexOf(filterGroup) < 0) {
          match = false;
        }
      }

      if (filterItem) {
        if (normalize(data[i][1]).indexOf(filterItem) < 0) {
          match = false;
        }
      }

      if (match) {
        results.push(data[i]);
      }
    }

    if (results.length === 0) {
      return { success: false, message: "Nenhum resultado encontrado com os filtros aplicados" };
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: results
      }
    };
  } catch (error) {
    Logger.log("Erro gerarListagemEstoqueWebApp: " + error);
    return { success: false, message: "Erro ao gerar listagem: " + error.message };
  }
}

/**
 * gerarRelatorioEstoqueWebApp: Wrapper para gerar relatório de estoque por período
 */
function gerarRelatorioEstoqueWebApp(dataInicio, dataFim) {
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

    var data = sheetEstoque.getRange(2, 1, lastRow - 1, 13).getValues();
    var results = [];

    var inicio = new Date(dataInicio);
    var fim = new Date(dataFim);
    inicio.setHours(0, 0, 0, 0);
    fim.setHours(23, 59, 59, 999);

    for (var i = 0; i < data.length; i++) {
      var dataMovimento = new Date(data[i][3]); // Coluna D (índice 3)
      if (dataMovimento >= inicio && dataMovimento <= fim) {
        results.push(data[i]);
      }
    }

    if (results.length === 0) {
      return { success: false, message: "Nenhum movimento encontrado no período" };
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: results
      }
    };
  } catch (error) {
    Logger.log("Erro gerarRelatorioEstoqueWebApp: " + error);
    return { success: false, message: "Erro ao gerar relatório: " + error.message };
  }
}

/**
 * gerarRelatorioPorGrupoWebApp: Wrapper para gerar relatório por grupo
 */
function gerarRelatorioPorGrupoWebApp(grupoSelecionado) {
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
    var results = [];
    var grupoNormalized = normalize(grupoSelecionado);

    for (var i = 0; i < data.length; i++) {
      if (normalize(data[i][0]) === grupoNormalized) {
        results.push(data[i]);
      }
    }

    if (results.length === 0) {
      return { success: false, message: "Nenhum item encontrado para o grupo selecionado" };
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: results
      }
    };
  } catch (error) {
    Logger.log("Erro gerarRelatorioPorGrupoWebApp: " + error);
    return { success: false, message: "Erro ao gerar relatório: " + error.message };
  }
}

/**
 * gerarListagemCoresDesatualizadasWebApp: Wrapper para cores desatualizadas
 */
function gerarListagemCoresDesatualizadasWebApp(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var mesesAtras = parseInt(formData.mesesAtras) || 3;
    var observacao = formData.observacao || "";

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: "Nenhum dado encontrado" };
    }

    var data = sheetEstoque.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var results = [];

    var today = new Date();
    var targetDate = new Date();
    targetDate.setMonth(today.getMonth() - mesesAtras);

    // Mapeia último registro de cada item
    var itemsMap = {};
    for (var i = 0; i < data.length; i++) {
      var item = data[i][1];
      var dataMovimento = new Date(data[i][3]); // Coluna D (índice 3)
      var obs = data[i][5] || ""; // Coluna F (índice 5)

      if (!itemsMap[item] || dataMovimento > new Date(itemsMap[item][2])) {
        itemsMap[item] = data[i];
      }
    }

    // Filtra itens desatualizados
    for (var item in itemsMap) {
      var row = itemsMap[item];
      var dataMovimento = new Date(row[2]);
      var obs = row[4] || "";

      var matchObs = !observacao || normalize(obs).indexOf(normalize(observacao)) >= 0;

      if (dataMovimento < targetDate && matchObs) {
        results.push(row);
      }
    }

    if (results.length === 0) {
      return { success: false, message: "Nenhuma cor desatualizada encontrada" };
    }

    return {
      success: true,
      data: {
        headers: ["Grupo", "Item", "Unidade", "Data", "NF", "Obs", "Saldo Anterior", "Entrada", "Saída", "Saldo", "Valor", "Alterado Em", "Alterado Por"],
        rows: results
      }
    };
  } catch (error) {
    Logger.log("Erro gerarListagemCoresDesatualizadasWebApp: " + error);
    return { success: false, message: "Erro ao buscar cores: " + error.message };
  }
}

/**
 * atualizarCompraDeFioEHistoricoWebApp: Wrapper para atualizar compra de fio
 */
function atualizarCompraDeFioEHistoricoWebApp() {
  try {
    // Chama a função original
    atualizarCompraDeFioEHistorico();
    return { success: true, message: "Compra de fio e histórico atualizados com sucesso" };
  } catch (error) {
    Logger.log("Erro atualizarCompraDeFioEHistoricoWebApp: " + error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

/**
 * atualizarTotalEmbarcadoWebApp: Wrapper para atualizar total embarcado
 */
function atualizarTotalEmbarcadoWebApp() {
  try {
    // Chama a função original
    atualizarTotalEmbarcado();
    return { success: true, message: "Total embarcado atualizado com sucesso" };
  } catch (error) {
    Logger.log("Erro atualizarTotalEmbarcadoWebApp: " + error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

/**
 * apagarUltimaLinhaWebApp: Wrapper para apagar última linha
 */
function apagarUltimaLinhaWebApp() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: "Nenhuma linha para apagar" };
    }

    sheetEstoque.deleteRow(lastRow);
    backupEstoqueData();
    invalidateCache();
    invalidateCacheOpt();

    return { success: true, message: "Última linha apagada com sucesso" };
  } catch (error) {
    Logger.log("Erro apagarUltimaLinhaWebApp: " + error);
    return { success: false, message: "Erro ao apagar linha: " + error.message };
  }
}

/**
 * limparFiltroEstoqueWebApp: Wrapper para limpar filtro
 */
function limparFiltroEstoqueWebApp() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var filter = sheetEstoque.getFilter();
    if (filter) {
      filter.remove();
    }

    return { success: true, message: "Filtro removido com sucesso" };
  } catch (error) {
    Logger.log("Erro limparFiltroEstoqueWebApp: " + error);
    return { success: false, message: "Erro ao remover filtro: " + error.message };
  }
}

// ========================================
// FUNÇÕES DE SINCRONIZAÇÃO - IndexedDB
// ========================================

/**
 * getAllDataForSync: Retorna TODOS os dados formatados para sincronização inicial
 * Usado apenas no primeiro carregamento ou refresh completo
 */
function getAllDataForSync() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [] };
    }

    var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, 13);
    var data = dataRange.getDisplayValues();
    var backgrounds = dataRange.getBackgrounds();

    var records = [];
    for (var i = 0; i < data.length; i++) {
      var dateStr = data[i][3]; // Coluna D (Data)
      var rowDate = dateStr ? new Date(dateStr) : new Date(0);

      records.push({
        row: data[i],
        date: rowDate.getTime(),
        background: backgrounds[i][0] || null
      });
    }

    Logger.log("getAllDataForSync: " + records.length + " registros");
    return { success: true, data: records };

  } catch (error) {
    Logger.log("Erro getAllDataForSync: " + error);
    return { success: false, message: "Erro ao buscar dados: " + error.message };
  }
}

/**
 * getNewRecordsSince: Retorna registros inseridos/modificados desde um timestamp
 * Usado para sincronização incremental (a cada 10 segundos)
 * @param {number} sinceTimestamp - Timestamp em milissegundos
 */
function getNewRecordsSince(sinceTimestamp) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEstoque = ss.getSheetByName("ESTOQUE");

    if (!sheetEstoque) {
      return { success: false, message: "Sheet ESTOQUE não encontrada" };
    }

    var lastRow = sheetEstoque.getLastRow();
    if (lastRow < 2) {
      return { success: true, newRecords: [] };
    }

    // Converte timestamp para Date
    var sinceDate = new Date(sinceTimestamp || 0);

    var dataRange = sheetEstoque.getRange(2, 1, lastRow - 1, 13);
    var data = dataRange.getDisplayValues();
    var backgrounds = dataRange.getBackgrounds();

    var newRecords = [];

    for (var i = 0; i < data.length; i++) {
      var dateStr = data[i][3]; // Coluna D (Data)
      var rowDate = dateStr ? new Date(dateStr) : new Date(0);

      // Também verifica coluna L (Alterado Em) para pegar edições
      var alteredStr = data[i][11]; // Coluna L (Alterado Em)
      var alteredDate = alteredStr ? new Date(alteredStr) : new Date(0);

      // Usa a data mais recente entre Data e Alterado Em
      var effectiveDate = alteredDate > rowDate ? alteredDate : rowDate;

      // Se o registro é mais recente que o último sync, inclui
      if (effectiveDate > sinceDate) {
        newRecords.push({
          row: data[i],
          date: rowDate.getTime(),
          background: backgrounds[i][0] || null
        });
      }
    }

    Logger.log("getNewRecordsSince: " + newRecords.length + " novos registros desde " + sinceDate.toLocaleString());
    return { success: true, newRecords: newRecords };

  } catch (error) {
    Logger.log("Erro getNewRecordsSince: " + error);
    return { success: false, message: "Erro ao buscar novos registros: " + error.message };
  }
}
