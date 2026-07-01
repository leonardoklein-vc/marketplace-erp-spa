// ==========================================
// ARQUIVO: Consultas.gs
// FOCO: Motor de Paginação, Relatórios, Configurações e Edição de Configs
// ==========================================

// --- CONFIGURAÇÕES DINÂMICAS ---

function adicionarItemConfiguracao(aba, coluna, valor, cor = null) {
  return withLock(() => {
    try {
      const sheet = getPlanilha().getSheetByName(aba);
      const ultimaLinha = sheet.getLastRow();
      const dadosExistentes = ultimaLinha > 0 ? sheet.getRange(1, coluna, ultimaLinha, 1).getValues() : [];
      
      let proxLinha = 2;
      for (let i = 1; i < dadosExistentes.length; i++) {
        if (!dadosExistentes[i][0]) { proxLinha = i + 1; break; }
        if (i === dadosExistentes.length - 1) proxLinha = i + 2;
      }
      
      sheet.getRange(proxLinha, coluna).setValue(valor);
      
      if(aba === 'Contas' && coluna === 1) {
        const rangeExtras = sheet.getRange(proxLinha, 3, 1, 2);
        rangeExtras.setValues([[cor || '#3b82f6', obterDataAtual()]]);
      }

      logAcao('CONFIG_ADICIONADA', '', `${aba}:${coluna}="${valor}"`);
      return { sucesso: true };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function excluirItemConfiguracao(aba, colunaIndex, valor) {
  return withLock(() => {
    try {
      const sheet = getPlanilha().getSheetByName(aba);
      const data = sheet.getDataRange().getValues();
      
      if (aba === 'Contas') {
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === valor) {
            sheet.deleteRow(i + 1);
            logAcao('CONFIG_REMOVIDA', '', `Contas: "${valor}"`);
            return { sucesso: true };
          }
        }
      } else {
        const ultimaLinha = sheet.getLastRow();
        if (ultimaLinha <= 1) return { sucesso: false, erro: 'Sem dados' };

        let colunaData = sheet.getRange(2, colunaIndex, ultimaLinha - 1, 1).getValues();
        let novaColuna = [];
        let encontrado = false;

        for (let i = 0; i < colunaData.length; i++) {
          if (colunaData[i][0] === valor && !encontrado) {
            encontrado = true; 
          } else {
            novaColuna.push([colunaData[i][0]]); 
          }
        }

        if (encontrado) {
          novaColuna.push(['']); 
          sheet.getRange(2, colunaIndex, novaColuna.length, 1).setValues(novaColuna);
          logAcao('CONFIG_REMOVIDA', '', `${aba} (Col ${colunaIndex}): "${valor}"`);
          return { sucesso: true };
        }
      }
      
      return { sucesso: false, erro: 'Item não encontrado' };
    } catch(e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function editarContaServidor(nomeAntigo, nomeNovo, novaCor) {
  return withLock(() => {
    try {
      const ss = getPlanilha();
      const sheetContas = ss.getSheetByName('Contas');
      const dadosContas = sheetContas.getDataRange().getValues();
      
      for (let i = 1; i < dadosContas.length; i++) {
        if (dadosContas[i][0] === nomeAntigo) {
          sheetContas.getRange(i + 1, 1).setValue(nomeNovo);
          if(novaCor) sheetContas.getRange(i + 1, 3).setValue(novaCor);
          break;
        }
      }

      if (nomeAntigo !== nomeNovo) {
        const sheetFin = ss.getSheetByName('Financeiro');
        if (sheetFin.getLastRow() > 1) {
          const dadosFin = sheetFin.getDataRange().getValues();
          for (let j = 1; j < dadosFin.length; j++) {
            if (dadosFin[j][6] === nomeAntigo) {
              sheetFin.getRange(j + 1, 7).setValue(nomeNovo);
            }
          }
        }
      }

      logAcao('CONFIG_EDITAR_CONTA', '', `${nomeAntigo} -> ${nomeNovo} (${novaCor})`);
      return { sucesso: true };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function renomearConfigGenericaServidor(aba, colunaIndex, nomeAntigo, nomeNovo) {
  return withLock(() => {
    try {
      const ss = getPlanilha();
      const sheet = ss.getSheetByName(aba);
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (data[i][colunaIndex - 1] === nomeAntigo) {
          sheet.getRange(i + 1, colunaIndex).setValue(nomeNovo);
          break;
        }
      }

      if (colunaIndex === 1) { 
        const sh = ss.getSheetByName('Vendas');
        if(sh.getLastRow() > 1){
           const d = sh.getDataRange().getValues();
           for(let j=1; j<d.length; j++){ if(d[j][6] === nomeAntigo) sh.getRange(j+1, 7).setValue(nomeNovo); }
        }
      } else if (colunaIndex === 2) { 
        // Atualiza a coluna de Canais no Catalogo se o nome do canal mudar
        const sh = ss.getSheetByName('Catalogo');
        if(sh && sh.getLastRow() > 1){
           const d = sh.getDataRange().getValues();
           for(let j=1; j<d.length; j++){ 
              if(d[j][6] && d[j][6].includes(nomeAntigo)) {
                 let novosCanais = d[j][6].replace(nomeAntigo, nomeNovo);
                 sh.getRange(j+1, 7).setValue(novosCanais); 
              }
           }
        }
      } else if (colunaIndex === 3 || colunaIndex === 4) { 
        const sh = ss.getSheetByName('Financeiro');
        if(sh.getLastRow() > 1){
           const d = sh.getDataRange().getValues();
           for(let j=1; j<d.length; j++){ if(d[j][3] === nomeAntigo) sh.getRange(j+1, 4).setValue(nomeNovo); }
        }
      }

      logAcao('CONFIG_RENOMEAR_GENERICO', '', `${nomeAntigo} -> ${nomeNovo}`);
      return { sucesso: true };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function obterConfiguracoes() {
  const ss = getPlanilha();
  const configData = ss.getSheetByName('Configuracoes').getDataRange().getValues();
  const contasData = ss.getSheetByName('Contas').getDataRange().getValues();
  
  const config = { canais: [], plataformas: [], catDespesas: [], catReceitas: [], contas: [] };
  
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0]) config.canais.push(configData[i][0]);
    if (configData[i][1]) config.plataformas.push(configData[i][1]);
    if (configData[i][2]) config.catDespesas.push(configData[i][2]);
    if (configData[i][3]) config.catReceitas.push(configData[i][3]);
  }
  
  for (let i = 1; i < contasData.length; i++) {
    if (contasData[i][0]) {
      config.contas.push({ 
        nome: contasData[i][0], 
        cor: contasData[i][2] || '#6c757d' 
      });
    }
  }
  return config;
}

// --- RESUMOS E DASHBOARD ---

function obterIdsCompras() {
  const sheet = getPlanilha().getSheetByName('Compras');
  let lista = [{ id: 'EST-0000', descricao: 'EST-0000 - Estoque Inicial Genérico', qtd: 999 }];
  if (sheet && sheet.getLastRow() >= 2) {
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    dados.forEach(linha => {
      if (linha[0]) {
        lista.push({
          id: linha[0],
          descricao: `${linha[0]} - ${linha[2]} (Est: ${linha[3] || 0})`,
          qtd: linha[3] || 0
        });
      }
    });
  }
  return lista;
}

function obterResumoFinanceiro(filtroGlobal = 'Todas', filtroData = '') {
  const ss = getPlanilha();
  let receitas = 0, despesas = 0;
  const controleCaixa = {};

  const contasData = ss.getSheetByName('Contas').getDataRange().getValues();
  for (let i = 1; i < contasData.length; i++) {
    controleCaixa[contasData[i][0]] = {
      saldo: Number(contasData[i][1]) || 0,
      cor: contasData[i][2] || '#6c757d'
    };
  }

  const finSheet = ss.getSheetByName('Financeiro');
  if (finSheet && finSheet.getLastRow() > 1) {
    let finData = finSheet.getDataRange().getValues().slice(1);
    
    if (filtroData) {
      const dataFiltro = new Date(filtroData);
      finData = finData.filter(row => new Date(row[1]) >= dataFiltro);
    }

    if (filtroGlobal !== 'Todas') {
      finData = finData.filter(row => (row[8] || 'Empresa') === filtroGlobal);
    }
    
    finData.forEach(row => {
      const tipo = row[2];
      const valor = Number(row[5]) || 0;
      const conta = row[6];
      const status = row[9] || 'Pago'; // Pega a coluna 10. Se vazia, assume 'Pago'.
      
      // MAGIA AQUI: Se for pendente, pula o laço, não soma em nada!
      if (status === 'Pendente') return; 
      
      if (tipo === 'Receita') {
        receitas += valor;
        if (controleCaixa[conta] !== undefined) controleCaixa[conta].saldo += valor;
      } else if (tipo === 'Despesa') {
        despesas += valor;
        if (controleCaixa[conta] !== undefined) controleCaixa[conta].saldo -= valor;
      } else if (tipo === 'Transferência Saída') {
        if (controleCaixa[conta] !== undefined) controleCaixa[conta].saldo -= valor;
      } else if (tipo === 'Transferência Entrada') {
        if (controleCaixa[conta] !== undefined) controleCaixa[conta].saldo += valor;
      }
    });
  }

  const lucro = receitas - despesas;
  return {
    custos: despesas.toFixed(2).replace('.', ','),
    receitas: receitas.toFixed(2).replace('.', ','),
    lucro: lucro.toFixed(2).replace('.', ','),
    margem: receitas > 0 ? ((lucro / receitas) * 100).toFixed(1) : '0.0',
    saldos: Object.entries(controleCaixa).map(([nome, dados]) => ({
      nome,
      saldo: 'R$ ' + dados.saldo.toFixed(2).replace('.', ','),
      cor: dados.cor
    }))
  };
}

// --- MOTOR DE TABELAS ---

function obterDadosTabelaPaginated(aba) {
  const sheet = getPlanilha().getSheetByName(aba);
  if (!sheet || sheet.getLastRow() < 2) return [];

  let data = sheet.getDataRange().getValues().slice(1);
  
  if (data.length > CONFIG.LIMITE_PAGINA) {
     data = data.slice(0, CONFIG.LIMITE_PAGINA);
  }

  return formatarLinhas(data.reverse(), aba); 
}

function formatarLinhas(dados, aba) {
  return dados.map(linha => linha.map((cel, idx) => {
    if (cel === null || cel === undefined || cel === '') return '-';
    
    if (cel instanceof Date) return Utilities.formatDate(cel, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    if (typeof cel === 'number') {
      const formatDinheiro = (v) => 'R$ ' + v.toFixed(2).replace('.', ',');
      if (aba === 'Compras' && (idx === 4 || idx === 5)) return formatDinheiro(cel);
      if (aba === 'Vendas' && (idx === 4 || idx === 5 || idx === 6)) return formatDinheiro(cel);
      if (aba === 'Financeiro' && idx === 5) return formatDinheiro(cel);
      if (aba === 'Contas' && idx === 1) return formatDinheiro(cel);
      
      // Formatação de dinheiro para a nova aba Catalogo na coluna 5 (índice 4)
      if (aba === 'Catalogo' && idx === 4) return formatDinheiro(cel);
    }
    return String(cel);
  }));
}

function obterClientes() {
  const sheet = getPlanilha().getSheetByName('Clientes');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.map(r => ({ id: r[0], nome: r[1] })).filter(c => c.id);
}

// Funções Expostas para o Front-end
function getTabelaClientes() { return obterDadosTabelaPaginated('Clientes'); }
function getTabelaCompras() { return obterDadosTabelaPaginated('Compras'); }
function getTabelaVendas() { return obterDadosTabelaPaginated('Vendas'); }
function getTabelaFinanceiro() { return obterDadosTabelaPaginated('Financeiro'); }

// Função do novo módulo
function getTabelaCatalogo() { return obterDadosTabelaPaginated('Catalogo'); }

function obterDetalhesLucroCatalogo(idCat) {
  const ss = getPlanilha();
  const catSheet = ss.getSheetByName('Catalogo');
  const compSheet = ss.getSheetByName('Compras');
  const catData = catSheet.getDataRange().getValues();
  const compData = compSheet.getDataRange().getValues();
  
  // Acha o produto no catálogo
  const produto = catData.find(r => r[0] === idCat);
  if (!produto) return { nome: "Não encontrado" };
  
  const idsOrigem = produto[5].split(', '); // Pega os IDs de compra vinculados
  let custoTotal = 0;
  let detalheItens = "";
  
  idsOrigem.forEach(id => {
    const compra = compData.find(c => c[0] === id.trim());
    if (compra) {
      custoTotal += (Number(compra[4]) + Number(compra[5])); // Custo + Frete da compra
      detalheItens += `<li>${compra[2]} (Custo: R$ ${compra[4].toFixed(2)})</li>`;
    }
  });
  
  const valorVenda = Number(produto[4]);
  const margemVal = valorVenda - custoTotal;
  const margemPerc = custoTotal > 0 ? (margemVal / valorVenda * 100).toFixed(1) : 0;

  // --- TRATAMENTO VISUAL PARA OS LINKS DE CANAIS ---
  let canaisHtml = "";
  if (produto[6]) {
    let canaisArray = produto[6].split(', ');
    canaisArray.forEach(c => {
      let partes = c.split('|');
      let nomeCanal = partes[0];
      let linkCanal = partes[1]; // Pode existir ou não
      
      if (linkCanal && linkCanal.trim() !== '') {
        // Se tem link, cria um botão clicável com ícone
        canaisHtml += `<a href="${linkCanal}" target="_blank" class="badge bg-secondary text-decoration-none me-1 mb-1 p-2"><i class="bi bi-box-arrow-up-right"></i> ${nomeCanal}</a>`;
      } else if (nomeCanal && nomeCanal.trim() !== '') {
        // Se não tem link, cria apenas uma etiqueta simples
        canaisHtml += `<span class="badge bg-light text-dark border me-1 mb-1 p-2">${nomeCanal}</span>`;
      }
    });
  }

  return {
    nome: produto[2],
    tipo: produto[3],
    status: produto[7],
    valorVenda: "R$ " + valorVenda.toLocaleString('pt-BR', {minimumFractionDigits:2}),
    valorCustoTotal: "R$ " + custoTotal.toLocaleString('pt-BR', {minimumFractionDigits:2}),
    margem: "R$ " + margemVal.toFixed(2) + " (" + margemPerc + "%)",
    detalheItens: detalheItens || "Nenhuma peça vinculada",
    canais: canaisHtml || "<span class='text-muted'>Não anunciado em nenhum canal.</span>"
  };
}

  function obterCatalogoAtivo() {
  const sheet = getPlanilha().getSheetByName('Catalogo');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  
  let ativos = [];
  for(let i = 1; i < data.length; i++) {
    if(data[i][7] === 'Ativo') { 
      ativos.push({
        id: data[i][0],
        nome: data[i][2],
        valor: Number(data[i][4]),
        idOrigem: data[i][5] // Pega os IDs de peças vinculadas
      });
    }
  }
  return ativos;
}
