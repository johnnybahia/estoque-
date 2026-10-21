// ========================================
// FORNECEDORES - Cadastro via CNPJ
// Consulta o CNPJ em várias fontes públicas, confirma por consenso (>= 2 fontes)
// e grava o nome na aba DADOS, coluna F (lista de Observações) e o CNPJ em
// coluna auxiliar ("CNPJ FORNECEDOR", criada automaticamente se não existir).
// ========================================

var FORN_COL_NOME = 6; // DADOS!F
var FORN_HEADER_CNPJ = "CNPJ FORNECEDOR";
var FORN_CACHE_TTL_OK = 21600; // 6 horas

var FORN_FONTES = [
  { id: "BrasilAPI", url: "https://brasilapi.com.br/api/cnpj/v1/" },
  { id: "MinhaReceita", url: "https://minhareceita.org/" },
  { id: "ReceitaWS", url: "https://receitaws.com.br/v1/cnpj/" },
  { id: "CNPJ.ws", url: "https://publica.cnpj.ws/cnpj/" }
];

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

function fornNormalizar_(texto) {
  return String(texto == null ? "" : texto).toUpperCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

/**
 * Chave de comparação: ignora acento, pontuação e espaços ("S.A." == "SA").
 */
function fornChave_(texto) {
  return fornNormalizar_(texto).replace(/ /g, "");
}

/**
 * Extrai razão social, nome fantasia e situação de qualquer uma das fontes.
 * Retorna null se a resposta for inválida ou for de outro CNPJ.
 */
function fornExtrairDados_(texto, cnpj) {
  var d;
  try { d = JSON.parse(texto); } catch (e) { return null; }
  if (!d || typeof d !== "object" || d.status === "ERROR") return null;

  var est = d.estabelecimento || {};
  var cnpjResposta = fornLimparCnpj_(d.cnpj || est.cnpj || "");
  if (cnpjResposta && cnpjResposta !== cnpj) return null;

  var razao = String(d.razao_social || d.nome || "").trim();
  if (!razao) return null;

  var fantasia = String(d.nome_fantasia || d.fantasia || est.nome_fantasia || "").trim();
  if (/^[*\s]*$/.test(fantasia)) fantasia = "";

  var situacao = String(d.descricao_situacao_cadastral || d.situacao || est.situacao_cadastral || "")
    .trim().toUpperCase();

  return { razaoSocial: razao, fantasia: fantasia, situacao: situacao };
}

/**
 * Entre as fontes que concordam, usa o nome fantasia mais frequente;
 * sem fantasia, usa a razão social.
 */
function fornEscolherNome_(grupo) {
  var contagem = {}, original = {}, melhor = "", max = 0;
  grupo.forEach(function(c) {
    if (!c.fantasia) return;
    var k = fornChave_(c.fantasia);
    contagem[k] = (contagem[k] || 0) + 1;
    if (!original[k]) original[k] = c.fantasia;
    if (contagem[k] > max) { max = contagem[k]; melhor = k; }
  });
  return melhor ? original[melhor] : grupo[0].razaoSocial;
}

/**
 * buscarCnpjFornecedor: consulta o CNPJ em todas as fontes em paralelo.
 * status: ok (>= 2 fontes coincidem) | parcial (1 fonte) | divergente | nao_encontrado | erro
 */
function buscarCnpjFornecedor(cnpjBruto) {
  try {
    var cnpj = fornLimparCnpj_(cnpjBruto);
    if (!fornCnpjValido_(cnpj)) {
      return { success: false, message: "CNPJ inválido. Confira os dígitos." };
    }

    var cache = CacheService.getScriptCache();
    var chave = "forn_cnpj_" + cnpj;
    var emCache = cache.get(chave);
    if (emCache) {
      try { return JSON.parse(emCache); } catch (e) {}
    }

    var opcoes = {
      method: "get",
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "Accept": "application/json", "User-Agent": "Mozilla/5.0 (compatible; EstoqueBahia/1.0)" }
    };
    var requisicoes = FORN_FONTES.map(function(f) {
      return { url: f.url + cnpj, method: opcoes.method, muteHttpExceptions: true,
               followRedirects: true, headers: opcoes.headers };
    });

    var respostas;
    try {
      respostas = UrlFetchApp.fetchAll(requisicoes);
    } catch (e) {
      // fetchAll falha por inteiro se uma fonte cair na rede; refaz uma a uma
      respostas = FORN_FONTES.map(function(f) {
        try { return UrlFetchApp.fetch(f.url + cnpj, opcoes); } catch (err) { return null; }
      });
    }

    var candidatos = [], falhas = [], naoEncontrado = 0;
    respostas.forEach(function(resp, i) {
      var fonte = FORN_FONTES[i].id;
      if (!resp) { falhas.push(fonte + " (sem resposta)"); return; }
      var codigo = resp.getResponseCode();
      if (codigo === 404) { naoEncontrado++; return; }
      if (codigo !== 200) { falhas.push(fonte + " (HTTP " + codigo + ")"); return; }
      var dados = fornExtrairDados_(resp.getContentText(), cnpj);
      if (!dados) { falhas.push(fonte + " (resposta inválida)"); return; }
      dados.fonte = fonte;
      candidatos.push(dados);
    });

    var resultado = {
      success: true,
      cnpj: cnpj,
      cnpjFormatado: fornFormatarCnpj_(cnpj),
      falhas: falhas,
      candidatos: candidatos.map(function(c) {
        return { fonte: c.fonte, razaoSocial: c.razaoSocial, fantasia: c.fantasia, nome: c.fantasia || c.razaoSocial };
      })
    };

    if (!candidatos.length) {
      resultado.status = (naoEncontrado >= 2 || (naoEncontrado >= 1 && !falhas.length)) ? "nao_encontrado" : "erro";
      return resultado;
    }

    var grupos = {}, base = null;
    candidatos.forEach(function(c) {
      var k = fornChave_(c.razaoSocial);
      (grupos[k] = grupos[k] || []).push(c);
    });
    Object.keys(grupos).forEach(function(k) {
      if (!base || grupos[k].length > base.length) base = grupos[k];
    });

    if (base.length >= 2) resultado.status = "ok";
    else if (candidatos.length === 1) resultado.status = "parcial";
    else resultado.status = "divergente";

    var situacao = "";
    base.forEach(function(c) { if (!situacao && c.situacao) situacao = c.situacao; });

    resultado.nome = fornEscolherNome_(base);
    resultado.razaoSocial = base[0].razaoSocial;
    resultado.situacao = situacao;
    resultado.ativa = situacao ? situacao === "ATIVA" : null;
    resultado.fontesConfirmadas = base.map(function(c) { return c.fonte; });

    if (resultado.status === "ok") {
      try { cache.put(chave, JSON.stringify(resultado), FORN_CACHE_TTL_OK); } catch (e) {}
    }
    return resultado;
  } catch (error) {
    Logger.log("Erro buscarCnpjFornecedor: " + error);
    return { success: false, message: "Erro ao consultar CNPJ: " + error.message };
  }
}

/**
 * Localiza a coluna do CNPJ na aba DADOS; cria cabeçalho ao final se não existir.
 */
function fornColunaCnpj_(sheet) {
  var ultimaColuna = Math.max(sheet.getLastColumn(), 1);
  var cabecalho = sheet.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];
  for (var i = 0; i < cabecalho.length; i++) {
    if (fornNormalizar_(cabecalho[i]) === FORN_HEADER_CNPJ) return i + 1;
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
