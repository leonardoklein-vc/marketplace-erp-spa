// ==========================================
// ARQUIVO: Inicializacao.gs
// FOCO: Menus, Instalação do Banco de Dados e Backups
// ==========================================

// --- MENU DA PLANILHA ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 FDTech OS')
    .addItem('Abrir Sistema', 'abrirSistema')
    .addSeparator()
    .addItem('Backup Automático', 'criarBackup')
    .addToUi();
}

// --- CONSTRUTOR DO BANCO DE DADOS ---
function garantirAbasECabecalhos() {
  const ss = getPlanilha();
  
  // A Estrutura reflete o ERP atual (Com Clientes, OS, Categoria Global e o novo Catálogo)
  const estrutura = {
    'Clientes': ['ID Cliente', 'Nome', 'Telefone', 'Canal', 'Observações'],
    'Compras': ['ID Compra', 'Data', 'Modelo/Item', 'Qtd Estoque', 'Custo Produto', 'Custo Frete', 'Status', 'Observações'],
    'Vendas': ['ID Venda', 'Itens Origem', 'Data', 'Qtd', 'Valor Cobrado', 'Valor Frete', 'Custo Frete', 'Canal', 'Peça/Serviço', 'Observações', 'Tipo', 'Cliente'],
    'Catalogo': ['ID Catálogo', 'Data', 'Nome do Item', 'Tipo', 'Valor', 'ID Origem', 'Canais Anunciados', 'Status', 'Link / Obs'],
    'Financeiro': ['ID Transação', 'Data', 'Tipo', 'Categoria', 'Descrição', 'Valor', 'Conta', 'ID Relacionado', 'Categoria Global'],
    'Contas': ['Nome Conta', 'Saldo Inicial', 'Cor', 'Última Atualização'],
    'Configuracoes': ['Canais Venda', 'Plataformas Anúncio', 'Cat. Despesa', 'Cat. Receita'],
    'Logs': ['Timestamp', 'Ação', 'Usuário', 'Detalhes', 'ID Afetado']
  };

  Object.entries(estrutura).forEach(([nomeAba, cabecalhos]) => {
    let sheet = ss.getSheetByName(nomeAba);
    if (!sheet) sheet = ss.insertSheet(nomeAba);
    
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, cabecalhos.length)
        .setValues([cabecalhos])
        .setFontWeight('bold')
        .setBackground('#111827')
        .setFontColor('white');
      sheet.setFrozenRows(1);
    }
  });

  inicializarSeed();
}

// --- DADOS INICIAIS PADRÃO ---
function inicializarSeed() {
  // Contas padrão (Com cores ajustadas para as marcas dos bancos)
  const sheetContas = getPlanilha().getSheetByName('Contas');
  if (sheetContas.getLastRow() === 1) {
    sheetContas.getRange(2, 1, 7, 4).setValues([
      ['Dinheiro', 0, '#10b981', obterDataAtual()],
      ['Nubank', 0, '#8b5cf6', obterDataAtual()],
      ['Neon', 0, '#0ea5e9', obterDataAtual()],
      ['PicPay', 0, '#22c55e', obterDataAtual()],
      ['BB', 0, '#eab308', obterDataAtual()],
      ['Banrisul', 0, '#3b82f6', obterDataAtual()],
      ['Itaú', 0, '#f97316', obterDataAtual()]
    ]);
  }

  // Configurações padrão organizadas na matriz exata (Preenchendo espaços vazios das colunas menores)
  const sheetConf = getPlanilha().getSheetByName('Configuracoes');
  if (sheetConf.getLastRow() === 1) {
    sheetConf.getRange(2, 1, 15, 4).setValues([
      ['Mercado Livre', 'Mercado Livre', 'Compra de Notes', 'Venda de Produto'],
      ['OLX', 'OLX', 'Frete/Transporte', 'Venda de Sucata'],
      ['Balcão', 'Shopee', 'Embalagem/Insumos', 'Serviço/Reparo'],
      ['Facebook Marketplace', 'Facebook Marketplace', 'Taxas/Impostos', 'Outras Receitas'],
      ['', 'Instagram', 'Geral', 'Salário'],
      ['', '', 'Consumo', ''],
      ['', '', 'Streaming', ''],
      ['', '', 'Saúde', ''],
      ['', '', 'Compras de Peças', ''],
      ['', '', 'Deslocamento', ''],
      ['', '', 'Alimentação', ''],
      ['', '', 'Mercado', ''],
      ['', '', 'Cartões de Crédito', ''],
      ['', '', 'Seguros', ''],
      ['', '', 'Cuidados/Beleza', '']
    ]);
  }
}

// --- INICIALIZAÇÃO DO FRONT-END ---
function abrirSistema() {
  try {
    garantirAbasECabecalhos();
    
    // Constrói o HTML a partir do arquivo Index.html
    const template = HtmlService.createTemplateFromFile('Index');
    
    const html = template.evaluate()
        .setTitle('FDTech ERP')
        .setWidth(1400)
        .setHeight(800); 
        
    SpreadsheetApp.getUi().showModalDialog(html, 'FDTech OS v2.0'); 
  } catch (e) { 
    SpreadsheetApp.getUi().alert('Erro ao iniciar: ' + e.toString()); 
  }
}

// --- UTILITÁRIO ADMINISTRATIVO ---
function criarBackup() {
  const ss = getPlanilha();
  const backupName = `FDTech_Backup_${obterDataAtual('yyyy-MM-dd_HH-mm')}`;
  const backup = ss.copy(backupName);
  
  // Opcional: Mover para uma pasta específica
  // DriveApp.getFileById(backup.getId()).moveTo(DriveApp.getFolderById('SEU_FOLDER_ID_BACKUPS')); 
  
  logAcao('BACKUP_CRIADO', 'SYSTEM', backupName);
  return { sucesso: true, url: backup.getUrl() };
}
