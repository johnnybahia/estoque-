# 🚀 OTIMIZAÇÕES DE PERFORMANCE - Sistema de Estoque

## 📊 Resumo das Melhorias

Este projeto foi otimizado para lidar com **40.000+ linhas** de dados com performance **instantânea** (< 1 segundo).

### Ganhos de Performance Estimados:

| Operação | Antes | Depois | Melhoria |
|----------|-------|--------|----------|
| **Consulta simples** | ~20-25s | < 0.5s | **~98% mais rápido** |
| **Batch 20 itens** | ~8-10 min | 2-5s | **~99% mais rápido** |
| **Autocomplete** | ~5-7s | < 0.3s | **~95% mais rápido** |
| **Dashboard** | ~3-5s | < 1s | **~90% mais rápido** |

---

## 🎯 O Que Foi Implementado

### **FASE 1: Cache Segmentado e Otimizações Imediatas**

1. **TTLs de Cache Aumentados**
   - Autocomplete: 2 min → **10 minutos**
   - Dashboard: 1 min → **5 minutos**
   - Índice de itens: 2 min → **30 minutos**
   - Índice completo: → **1 hora**

2. **Leitura Única em Operações Batch**
   - Antes: Para 20 itens = 40+ leituras da planilha
   - Depois: Para 20 itens = **1 leitura da planilha**
   - Redução de **97% nas chamadas API**

3. **Inserção em Batch**
   - Antes: 20 inserções individuais (sequencial)
   - Depois: **1 inserção única** com todas as linhas
   - Muito mais rápido para Google Sheets API

---

### **FASE 2: Sistema de Índice Permanente**

#### **Nova Aba: `ÍNDICE_ITENS`**

Uma aba especial que mantém um "índice" de todos os itens únicos:

```
┌─────────────────┬──────────────┬──────────────┬────────────┬────────────────┬────────────────────┐
│ Item            │ Saldo Atual  │ Última Data  │ Grupo      │ Linha ESTOQUE  │ Última Atualização │
├─────────────────┼──────────────┼──────────────┼────────────┼────────────────┼────────────────────┤
│ CIMENTO CP2 50KG│ 150.5        │ 2025-11-28   │ CONSTRUÇÃO │ 38547          │ 2025-11-28 10:30   │
│ AREIA FINA      │ 20.0         │ 2025-11-28   │ MAT PRIMA  │ 40123          │ 2025-11-28 11:15   │
│ ...             │ ...          │ ...          │ ...        │ ...            │ ...                │
└─────────────────┴──────────────┴──────────────┴────────────┴────────────────┴────────────────────┘
```

#### **Como Funciona:**

1. **Ao inserir um item**: A função `updateIndiceItem()` atualiza automaticamente o índice
2. **Ao consultar um item**: A função `getLastRegistrationFromIndex()` lê o índice (300 linhas) ao invés da aba ESTOQUE (40.000 linhas)
3. **Resultado**: Busca **O(1)** instantânea ao invés de **O(n)** linear

#### **Vantagens:**

- ✅ Consultas **133x mais rápidas** (40.000 → 300 linhas)
- ✅ Performance **constante** independente do tamanho da planilha
- ✅ Cache mais eficiente (300 linhas cabe facilmente no limite de 100KB)
- ✅ Índice atualizado **automaticamente** após cada inserção

---

## 🛠️ Como Usar

### **1️⃣ Primeira Vez: Construir o Índice**

**IMPORTANTE:** Antes de usar o sistema otimizado, você precisa construir o índice inicial.

#### **Opção A: Inicialização Automática**

Execute este script no Google Apps Script:

```javascript
function inicializarIndice() {
  var result = initializeIndiceIfNeeded();
  Logger.log(result.message);

  if (result.initialized) {
    Logger.log("✅ Índice construído com sucesso!");
    Logger.log("Total de itens: " + result.totalItems);
  } else {
    Logger.log("✅ Índice já existe, nada a fazer");
  }
}
```

**Tempo esperado:** 30-60 segundos para 40.000 linhas

#### **Opção B: Reconstrução Manual**

Se você já tem o índice mas quer reconstruí-lo do zero:

```javascript
function reconstruirIndice() {
  var result = reconstruirIndiceCompleto();
  Logger.log("Índice reconstruído: " + result.totalItems + " itens em " + result.duration + " segundos");
}
```

#### **Opção C: Verificação e Reparo**

Para verificar se o índice está OK e reparar se necessário:

```javascript
function verificarIndice() {
  var result = verificarERepararIndice();
  Logger.log(result.message);
}
```

---

### **2️⃣ Uso Normal**

Após construir o índice, **não é necessário fazer mais nada!**

Todas as funções já foram atualizadas para usar automaticamente:

- ✅ `processEstoqueWebApp()` - Inserção única
- ✅ `processMultipleEstoqueItems()` - Batch simples
- ✅ `processMultipleEstoqueItemsWithGroup()` - Batch com grupo
- ✅ `getLastRegistrationFromIndex()` - Consulta otimizada
- ✅ `getItemGroupFromIndex()` - Grupo otimizado

O índice é **atualizado automaticamente** após cada inserção.

---

### **3️⃣ Manutenção**

#### **Quando Reconstruir o Índice?**

Reconstrua o índice se:
- Você fez alterações manuais diretas na aba ESTOQUE (fora do sistema)
- Você importou dados antigos
- O índice ficou dessincronizado por algum motivo

#### **Como Saber se o Índice Está Dessincronizado?**

Execute `verificarERepararIndice()` periodicamente (ex: 1x por semana).

---

## 📈 Estrutura das Otimizações

### **Funções Principais:**

| Função | Propósito | Performance |
|--------|-----------|-------------|
| `buildIndiceItensInitial()` | Constrói índice inicial do zero | ~30-60s para 40k linhas |
| `getIndiceItensCache()` | Carrega índice em cache | < 0.1s (com cache) |
| `getLastRegistrationFromIndex()` | Busca último registro via índice | **< 0.01s** |
| `updateIndiceItem()` | Atualiza 1 item no índice | < 0.1s |
| `initializeIndiceIfNeeded()` | Inicializa se necessário | Auto-detecta |
| `reconstruirIndiceCompleto()` | Reconstrói manualmente | ~30-60s |
| `verificarERepararIndice()` | Verifica e repara | ~2-5s |

### **Cache Hierarchy:**

```
┌─────────────────────────────────────────────────────────────┐
│                    REQUEST (Web App)                         │
└────────────────────────┬────────────────────────────────────┘
                         ▼
                  ┌──────────────┐
                  │  Cache (1h)  │ ← indiceItensCache
                  └──────┬───────┘
                         ▼
                  ┌──────────────┐
                  │ Aba ÍNDICE   │ ← 300-500 linhas (itens únicos)
                  │  (~0.5s)     │
                  └──────┬───────┘
                         │ (fallback raro)
                         ▼
                  ┌──────────────┐
                  │ Aba ESTOQUE  │ ← 40.000 linhas (só se necessário)
                  │  (~20-30s)   │
                  └──────────────┘
```

---

## 🔧 Detalhes Técnicos

### **Antes da Otimização:**

```javascript
// ❌ LENTO: Lê 40k linhas para CADA item
for (var i = 0; i < 20; i++) {
  var lastReg = getLastRegistration(item);  // Lê 40k linhas
  var grupo = getItemGroup(item);           // Lê 40k linhas
  // ...
}
// Total: 40 leituras × 20s = ~13 minutos
```

### **Depois da Otimização:**

```javascript
// ✅ RÁPIDO: Lê índice UMA vez para todos os itens
var indice = getIndiceItensCache();  // Lê 300 linhas uma vez (0.5s)

for (var i = 0; i < 20; i++) {
  var lastReg = indice[item];  // Busca O(1) em memória (0.01s)
  var grupo = indice[item].grupo;
  // ...
}
// Total: 1 leitura + processamento = ~2-5 segundos
```

---

## 🚨 Troubleshooting

### **"Erro: Aba ÍNDICE_ITENS não encontrada"**

**Solução:** Execute `initializeIndiceIfNeeded()` para criar o índice.

### **"Consultas retornando saldo 0 ou dados errados"**

**Causa:** Índice dessincronizado.
**Solução:** Execute `reconstruirIndiceCompleto()`.

### **"Performance ainda lenta após otimizações"**

**Possíveis causas:**
1. Índice não foi construído → Execute `initializeIndiceIfNeeded()`
2. Cache vazio (primeiro acesso) → Aguarde 1-2 segundos para popular
3. Função antiga sendo usada → Verifique se está usando as funções `*FromIndex()`

### **"Timeout de 6 minutos ao construir índice"**

**Causa:** Planilha muito grande (>100k linhas) ou conexão lenta.
**Solução:**
1. Tente novamente (pode ser problema temporário)
2. Divida a planilha em múltiplas abas por ano/período
3. Considere migrar para banco de dados real (Firebase/SQL)

---

## 📝 Changelog

### **v2.1 - Correções de Busca, Lote e Performance (2026-07-08)**

**Correções de busca (itens que "não eram listados"):**
- ✅ Datas em formato brasileiro (dd/mm/aaaa) eram parseadas com `new Date()` → datas inválidas quebravam o índice de histórico e a sincronização incremental (registros novos nunca chegavam ao cliente). Agora usa `parseDateBR`.
- ✅ Busca de produto exigia nome EXATO — agora aceita busca parcial e ignora acentos (exato tem prioridade).
- ✅ Autocomplete só sugeria por prefixo — agora também encontra por trecho no meio do nome.
- ✅ Busca rápida podia exibir o histórico do item ERRADO (match parcial ambíguo) — agora só usa o cache local quando o match é inequívoco.

**Correções de lançamento:**
- ✅ Saldos com ruído de ponto flutuante (ex: `-5.000000000000064`) — arredondamento em todos os cálculos.
- ✅ Lote sem confirmação: leituras de 40k+ linhas POR ITEM estouravam o tempo de execução; agora a planilha é lida 1 vez por lote (apenas linhas recentes) e falhas pós-inserção não descartam mais a confirmação.
- ✅ Vários retornos de servidor continham objetos `Date` (ilegais no `google.script.run`), fazendo chamadas falharem silenciosamente — convertidos para string/timestamp.

**Performance:**
- ✅ `Código.gs` estava com o conteúdo praticamente duplicado (57 funções repetidas, edições na 1ª cópia eram ignoradas) — deduplicado mantendo as definições efetivas.
- ✅ Índice de histórico, últimos lançamentos e verificação de atualização agora leem só as últimas 5000 linhas em vez da planilha inteira.
- ✅ Atualização do índice em background limitada a 1x a cada 2 minutos (antes: a cada busca).

### **v2.0 - Otimizações Massivas (2025-11-28)**

- ✅ Implementado sistema de índice permanente (ÍNDICE_ITENS)
- ✅ Cache segmentado com TTLs inteligentes (10min - 1h)
- ✅ Leitura única em operações batch
- ✅ Inserção em batch otimizada
- ✅ Funções de manutenção do índice
- ✅ Busca O(1) ao invés de O(n)
- 🎯 **Resultado: 90-99% mais rápido**

---

## 💡 Próximos Passos (Futuro)

Se a planilha crescer para **> 100.000 linhas**, considere:

1. **Migrar para Firebase Firestore**
   - Banco NoSQL gratuito do Google
   - Queries instantâneas (< 100ms)
   - Suporta milhões de registros
   - Integração fácil com Google Sheets

2. **Migrar para Google Cloud SQL**
   - Banco SQL completo (MySQL/PostgreSQL)
   - Queries complexas com índices automáticos
   - Ideal para relatórios avançados

3. **Arquivamento Automático**
   - Mover registros > 1 ano para aba "HISTÓRICO"
   - Manter aba ESTOQUE com apenas registros recentes
   - Reduz tamanho da planilha principal

---

## 📧 Suporte

Se tiver dúvidas ou problemas:

1. Verifique os logs: `View > Logs` no Google Apps Script
2. Execute `verificarERepararIndice()` para diagnóstico
3. Reconstrua o índice com `reconstruirIndiceCompleto()` em último caso

---

**🎉 Sistema otimizado e pronto para uso com 40k+ linhas!**
