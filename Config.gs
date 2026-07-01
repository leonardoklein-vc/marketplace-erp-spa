// ==========================================
// ARQUIVO: Config.gs
// FOCO: Configurações Globais e Funções Utilitárias Base
// ==========================================

// --- CONFIGURAÇÕES GLOBAIS ---
const CONFIG = {
  LIMITE_PAGINA: 5000,
  MAX_TENTATIVAS_LOCK: 3,
  DATA_FORMAT: 'dd/MM/yyyy',
  HORA_FORMAT: 'dd/MM/yyyy HH:mm'
};

// --- FUNÇÕES BASE E HELPERS ---

// Essa função é a que puxa os arquivos separados de HTML/JS/CSS para o Index
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Retorna a planilha ativa para ser usada em todas as outras requisições
function getPlanilha() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Sistema de fila/espera para evitar conflitos de edição simultânea
function withLock(fn, ...args) {
  const lock = LockService.getScriptLock();
  for (let i = 0; i < CONFIG.MAX_TENTATIVAS_LOCK; i++) {
    try {
      lock.waitLock(10000);
      const resultado = fn(...args);
      lock.releaseLock();
      return resultado;
    } catch (e) {
      Utilities.sleep(1000);
    }
  }
  return { sucesso: false, erro: 'Timeout no processamento (LockService)' };
}

// Sistema centralizado de Logs de auditoria
function logAcao(acao, usuario, detalhes, idAfetado = '') {
  try {
    const sheet = getPlanilha().getSheetByName('Logs');
    sheet.appendRow([
      new Date(),
      acao,
      usuario || Session.getActiveUser().getEmail(),
      detalhes,
      idAfetado
    ]);
  } catch(e) {
    console.warn("Falha ao registrar log: " + e.toString());
  }
}

// Validador genérico de dados para formulários e requisições
function validarDados(dados, regras) {
  const erros = [];
  Object.entries(regras).forEach(([chave, validacoes]) => {
    const valor = dados[chave];
    if (validacoes.required && (!valor || valor.toString().trim() === '')) {
      erros.push(`${chave} é obrigatório`);
    }
    if (validacoes.min != null && Number(valor) < validacoes.min) {
      erros.push(`${chave} mínimo ${validacoes.min}`);
    }
  });
  return erros.length ? { valido: false, erros } : { valido: true };
}

// Gerador de IDs únicos padronizados (Ex: COMP-A1B2C3D4)
function gerarId(prefixo) {
  return `${prefixo}-${Utilities.getUuid().slice(0, 8).toUpperCase()}`;
}

// Retorna a data atual já formatada conforme a timezone do script
function obterDataAtual(formato = CONFIG.DATA_FORMAT) {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), formato);
}

// Busca a linha exata de um registro no banco de dados com base no ID (AGORA COM BLINDAGEM)
function encontrarLinha(aba, id) {
  const sheet = getPlanilha().getSheetByName(aba);
  if (!sheet) return null; // Previne erro caso a aba não exista
  
  const data = sheet.getDataRange().getValues();
  const idBusca = String(id).trim().toUpperCase(); // Limpa espaços e força maiúscula
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === idBusca) {
      return i + 1;
    }
  }
  return null;
}
