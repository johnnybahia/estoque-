// ========================================
// FORNECEDORES - Cadastro via CNPJ
// A consulta do CNPJ (>= 2 fontes) é feita no navegador (WebApp.html).
// Aqui apenas se grava o nome em DADOS!F (lista de Observações) e o CNPJ
// em coluna auxiliar ("CNPJ FORNECEDOR", criada automaticamente se não existir).
// ========================================

var FORN_COL_NOME = 6; // DADOS!F
var FORN_HEADER_CNPJ = "CNPJ FORNECEDOR";

/**
 * Remove máscara e normaliza o CNPJ (aceita o formato alfanumérico vigente desde 07/2026).
 */
function fornLimparCnpj_(valor) {
  return String(valor == null ? "" : valor).replace(/[^0-9A-Za-z]/g, "").toUpperCase();
}

function fornFormatarCnpj_(cnpj) {
  return cnpj.replace(/^(.{2})(.{3})(.{3})(.{4})(.{2})$/, "$1.$2.$3/$4-$5");
}

/**
 * Valida estrutura e dígitos verificadores (módulo 11; letras valem código ASCII - 48).
 */
function fornCnpjValido_(cnpj) {
  if (!/^[0-9A-Z]{12}[0-9]{2}$/.test(cnpj) || /^(.)\1{13}$/.test(cnpj)) return false;
  var digito = function(base) {
    var pesos = base.length === 12
      ? [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
      : [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
    var soma = 0;
    for (var i = 0; i < base.length; i++) soma += (base.charCodeAt(i) - 48) * pesos[i];
    var resto = soma % 11;
    return resto < 2 ? 0 : 11 - resto;
  };
  var d1 = digito(cnpj.substr(0, 12));
  var d2 = digito(cnpj.substr(0, 12) + d1);
  return d1 === parseInt(cnpj.charAt(12), 10) && d2 === parseInt(cnpj.charAt(13), 10);
}

/**
 * Chave de comparação: ignora acento, pontuação e espaços ("S.A." == "SA").
 */
function fornChave_(texto) {
  return String(texto == null ? "" : texto).toUpperCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

/**
 * Localiza a coluna do CNPJ na aba DADOS; cria cabeçalho ao final se não existir.
 */
function fornColunaCnpj_(sheet) {
  var ultimaColuna = Math.max(sheet.getLastColumn(), 1);
  var cabecalho = sheet.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];
  for (var i = 0; i < cabecalho.length; i++) {
    if (fornChave_(cabecalho[i]) === fornChave_(FORN_HEADER_CNPJ)) return i + 1;
  }
  var nova = ultimaColuna + 1;
  if (nova > sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());
  sheet.getRange(1, FORN_COL_NOME).copyTo(sheet.getRange(1, nova), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.getRange(1, nova).setValue(FORN_HEADER_CNPJ);
  return nova;
}

/**
 * cadastrarFornecedor: grava o nome em DADOS!F e o CNPJ na coluna auxiliar.
 * - CNPJ já cadastrado: não duplica, devolve o nome existente.
 * - Nome já existente sem CNPJ (cadastro manual antigo): apenas vincula o CNPJ.
 * - Nome já existente com outro CNPJ: recusa (homônimo).
 */
function cadastrarFornecedor(nomeBruto, cnpjBruto) {
  var lock = LockService.getScriptLock();
  try {
    var nome = String(nomeBruto == null ? "" : nomeBruto).replace(/^[=+\-@\s]+/, "").replace(/\s+/g, " ").trim();
    var cnpj = fornLimparCnpj_(cnpjBruto);
    if (!nome) return { success: false, message: "Informe o nome do fornecedor." };
    if (nome.length > 120) return { success: false, message: "Nome muito longo (máx. 120 caracteres)." };
    if (!fornCnpjValido_(cnpj)) return { success: false, message: "CNPJ inválido." };

    lock.waitLock(15000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DADOS");
    if (!sheet) return { success: false, message: "Aba DADOS não encontrada." };

    var colunaCnpj = fornColunaCnpj_(sheet);
    var totalLinhas = Math.max(sheet.getLastRow(), 1);
    var nomes = sheet.getRange(1, FORN_COL_NOME, totalLinhas, 1).getDisplayValues();
    var cnpjs = sheet.getRange(1, colunaCnpj, totalLinhas, 1).getDisplayValues();

    var nomeNormalizado = fornChave_(nome);
    var ultimaLinhaComNome = 1; // linha 1 = cabeçalho
    var linhaMesmoNome = -1;

    for (var i = 1; i < totalLinhas; i++) {
      var existente = String(nomes[i][0]).trim();
      if (!existente) continue;
      ultimaLinhaComNome = i + 1;

      if (fornLimparCnpj_(cnpjs[i][0]) === cnpj) {
        return { success: true, jaExistia: true, nome: existente, observacoes: getObservacoesList() };
      }
      if (linhaMesmoNome < 0 && fornChave_(existente) === nomeNormalizado) linhaMesmoNome = i + 1;
    }

    var formatado = fornFormatarCnpj_(cnpj);

    if (linhaMesmoNome > 0) {
      var cnpjDaLinha = fornLimparCnpj_(cnpjs[linhaMesmoNome - 1][0]);
      if (cnpjDaLinha) {
        return { success: false, message: 'Já existe "' + nomes[linhaMesmoNome - 1][0] + '" com outro CNPJ. Ajuste o nome (ex.: acrescente a cidade) e salve novamente.' };
      }
      sheet.getRange(linhaMesmoNome, colunaCnpj).setNumberFormat("@").setValue(formatado);
      fornInvalidarCaches_();
      return { success: true, jaExistia: true, nome: String(nomes[linhaMesmoNome - 1][0]).trim(), observacoes: getObservacoesList() };
    }

    var novaLinha = ultimaLinhaComNome + 1;
    if (novaLinha > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), novaLinha - sheet.getMaxRows());
    sheet.getRange(novaLinha, FORN_COL_NOME).setNumberFormat("@").setValue(nome);
    sheet.getRange(novaLinha, colunaCnpj).setNumberFormat("@").setValue(formatado);
    SpreadsheetApp.flush();

    fornInvalidarCaches_();
    return { success: true, jaExistia: false, nome: nome, observacoes: getObservacoesList() };
  } catch (error) {
    Logger.log("Erro cadastrarFornecedor: " + error);
    return { success: false, message: "Erro ao cadastrar fornecedor: " + error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function fornInvalidarCaches_() {
  try { invalidateCache(); } catch (e) {}
  try { invalidateCacheOpt(); } catch (e) {}
}
