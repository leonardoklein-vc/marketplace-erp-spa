// ==========================================
// ARQUIVO: Financeiro.gs
// FOCO: CRUD Completo (Salvar, Editar, Excluir) e Regras de Negócio
// ==========================================

// --- SALVAMENTO (CREATE) ---

function salvarTransacaoFinanceira(tipo, categoria, descricao, valor, conta, idRelacionado = '', recorrencia = 1, dataInput = '', catGlobal = 'Empresa', status = 'Pago') {
  return withLock(() => {
    try {
      const sheet = getPlanilha().getSheetByName('Financeiro');
      const meses = parseInt(recorrencia) || 1;
      const dataBase = dataInput ? new Date(dataInput + 'T12:00:00') : new Date();
      const idsGerados = [];

      for (let i = 0; i < meses; i++) {
        let dataTransacao = new Date(dataBase);
        dataTransacao.setMonth(dataTransacao.getMonth() + i);
        let dataFormatada = Utilities.formatDate(dataTransacao, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
        let descFinal = meses > 1 ? `${descricao} (${i + 1}/${meses})` : descricao;
        const idTransacao = gerarId('FIN');
        
        // Salvando agora com a 10ª coluna (status)
        sheet.appendRow([idTransacao, dataFormatada, tipo, categoria, descFinal, Number(valor || 0), conta, idRelacionado, catGlobal, status]);
        
        logAcao('FIN_LANCAMENTO', '', descFinal, idTransacao);
        idsGerados.push(idTransacao);
      }
      return { sucesso: true, id: idsGerados[0] };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function salvarFinanceiroGeral(dados) {
  if (dados.tipo === 'Transferência') {
    if (dados.conta === dados.contaDestino) return { sucesso: false, erro: 'As contas devem ser diferentes.' };
    return withLock(() => {
      try {
        const sheet = getPlanilha().getSheetByName('Financeiro');
        let dataFormatada = dados.data;
        if (dados.data && dados.data.includes('-')) {
            const d = new Date(dados.data + 'T12:00:00');
            dataFormatada = Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
        } else { dataFormatada = obterDataAtual(); }

        const idRelacionado = gerarId('TRF');
        const descSaida = dados.descricao + ` (Para: ${dados.contaDestino})`;
        const descEntrada = dados.descricao + ` (De: ${dados.conta})`;

        sheet.appendRow([gerarId('FIN'), dataFormatada, 'Transferência Saída', 'Transferência', descSaida, Number(dados.valor), dados.conta, idRelacionado, dados.catGlobal, 'Pago']);
        sheet.appendRow([gerarId('FIN'), dataFormatada, 'Transferência Entrada', 'Transferência', descEntrada, Number(dados.valor), dados.contaDestino, idRelacionado, dados.catGlobal, 'Pago']);

        logAcao('TRANSFERENCIA', '', `${dados.conta} -> ${dados.contaDestino}`, idRelacionado);
        return { sucesso: true };
      } catch (e) { return { sucesso: false, erro: e.toString() }; }
    });
  } else {
    const rec = dados.recorrencia || 1; 
    // Passando o status pro salvamento
    return salvarTransacaoFinanceira(dados.tipo, dados.categoria, dados.descricao, dados.valor, dados.conta, 'AVULSO', rec, dados.data, dados.catGlobal, dados.status);
  }
}

function salvarCompra(dados) {
  if (!dados.itens || dados.itens.length === 0) return { sucesso: false, erro: "Adicione ao menos um item na lista de compra." };

  return withLock(() => {
    try {
      const sheet = getPlanilha().getSheetByName('Compras');
      const dataBase = dados.data ? new Date(dados.data + 'T12:00:00') : new Date();
      const dataFormatada = Utilities.formatDate(dataBase, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);

      let custoProdutosTotal = 0;
      let totalQtd = 0;
      let modelosAnotados = [];

      dados.itens.forEach(i => {
        custoProdutosTotal += (Number(i.custo) * Number(i.qtd));
        totalQtd += Number(i.qtd);
        modelosAnotados.push(i.modelo);
      });

      const freteTotal = Number(dados.freteTotal) || 0;
      const fretePorUnidade = totalQtd > 0 ? (freteTotal / totalQtd) : 0; 
      const valorFinanceiroTotal = custoProdutosTotal + freteTotal;
      
      const idPrincipalFinanceiro = gerarId('COMP'); 

      dados.itens.forEach((item, index) => {
        const idItem = (index === 0) ? idPrincipalFinanceiro : gerarId('COMP');
        const freteRateadoItem = fretePorUnidade * item.qtd;
        
        sheet.appendRow([
          idItem, dataFormatada, item.modelo, item.qtd, 
          Number(item.custo), freteRateadoItem, 'Em Estoque', dados.obs || ''
        ]);
      });

      let descFin = `Nota/Lote: ${modelosAnotados.join(', ')}`;
      if (descFin.length > 60) descFin = descFin.substring(0, 57) + '...';

      // NOVO: Detector de data futura para definir status inicial
      const hoje = new Date();
      hoje.setHours(0,0,0,0);
      const isFuturo = dataBase > hoje;
      const statusCompra = isFuturo ? 'Pendente' : 'Pago';

      salvarTransacaoFinanceira('Despesa', 'Compra Estoque', descFin, valorFinanceiroTotal, dados.conta, idPrincipalFinanceiro, 1, dados.data, 'Empresa', statusCompra);
      logAcao('COMPRA_LOTE', '', `Itens: ${dados.itens.length}`, idPrincipalFinanceiro);
      return { sucesso: true };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function salvarVenda(dados) {
  return withLock(() => {
    try {
      const sheetEstoque = getPlanilha().getSheetByName('Compras');
      let origensAnotadas = [];

      if (dados.itensUsados && dados.itensUsados.length > 0) {
        dados.itensUsados.forEach(item => {
          if (item.idOrigem && item.idOrigem !== 'NENHUM') {
            const linhaEstoque = encontrarLinha('Compras', item.idOrigem);
            if (linhaEstoque) {
              const qtdAtual = Number(sheetEstoque.getRange(linhaEstoque, 4).getValue());
              sheetEstoque.getRange(linhaEstoque, 4).setValue(qtdAtual - Number(item.qtd));
              origensAnotadas.push(`${item.idOrigem} (${item.qtd}x)`);
            }
          }
        });
      }

      const strOrigens = origensAnotadas.length > 0 ? origensAnotadas.join(', ') : 'Sem peças de estoque';
      const idVenda = gerarId('VEND');
      const dataBase = dados.data ? new Date(dados.data + 'T12:00:00') : new Date();
      const dataFormatada = Utilities.formatDate(dataBase, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
      const valorTotalFin = Number(dados.valorProduto) + Number(dados.valorFrete || 0) - Number(dados.custoFrete || 0);
      
      getPlanilha().getSheetByName('Vendas').appendRow([
        idVenda, strOrigens, dataFormatada, 1, Number(dados.valorProduto), Number(dados.valorFrete || 0), 
        Number(dados.custoFrete || 0), dados.canal, dados.peca, dados.obs || '', dados.tipo, dados.cliente
      ]);

      const descFin = `${dados.tipo}: ${dados.peca} | Cli: ${dados.cliente}`;
      
      // NOVO: Detector de data futura para definir status inicial
      const hoje = new Date();
      hoje.setHours(0,0,0,0);
      const isFuturo = dataBase > hoje;
      const statusVenda = isFuturo ? 'Pendente' : 'Pago';

      salvarTransacaoFinanceira('Receita', 'Venda/Serviço', descFin, valorTotalFin, dados.conta, idVenda, 1, dados.data, 'Empresa', statusVenda);
      
      if (dados.idCatalogo) {
        const linhaCat = encontrarLinha('Catalogo', dados.idCatalogo);
        if (linhaCat) {
          getPlanilha().getSheetByName('Catalogo').getRange(linhaCat, 8).setValue('Vendido');
        }
      }

      logAcao('VENDA_SERVICO', '', dados.peca, idVenda);
      return { sucesso: true, id: idVenda };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function salvarCliente(dados) {
  return withLock(() => {
    try {
      const id = gerarId('CLI');
      getPlanilha().getSheetByName('Clientes').appendRow([id, dados.nome, dados.telefone, dados.canal, dados.obs || '']);
      logAcao('CLIENTE_CRIADO', '', dados.nome, id);
      return { sucesso: true, id: id };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

function salvarCatalogo(dados) {
  return withLock(() => {
    try {
      const sheet = getPlanilha().getSheetByName('Catalogo');
      const idCatalogo = gerarId('CAT');
      const dataBase = dados.data ? new Date(dados.data + 'T12:00:00') : new Date();
      const dataFormatada = Utilities.formatDate(dataBase, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);

      // Asseguramos que os dados não cheguem undefined antes de gravar
      const nome = dados.nome || 'Sem Nome';
      const tipo = dados.tipo || '-';
      const valor = Number(dados.valor) || 0;
      const idOrigem = dados.idOrigem || 'NENHUM';
      const canais = dados.canais || '';
      const status = dados.status || 'Ativo';
      const obs = dados.obs || '';

      sheet.appendRow([
        idCatalogo, 
        dataFormatada, 
        nome, 
        tipo, 
        valor, 
        idOrigem, 
        canais, 
        status, 
        obs
      ]);
      
      logAcao('CATALOGO_CRIADO', '', nome, idCatalogo);
      return { sucesso: true };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

// --- EDIÇÃO (UPDATE) ---

function editarRegistro(aba, dados) {
  return withLock(() => {
    if (aba === 'Compras') return editarCompra(dados);
    if (aba === 'Vendas') return editarVenda(dados);
    if (aba === 'Catalogo') return editarCatalogo(dados); // Mudou aqui
    if (aba === 'Financeiro') return editarFinanceiroAvulso(dados);
    if (aba === 'Clientes') return editarCliente(dados); 
    return { sucesso: false, erro: 'Aba não suportada para edição direta' };
  });
}

function editarCompra(dados) {
  try {
    const linha = encontrarLinha('Compras', dados.id);
    if(!linha) return {sucesso: false, erro: 'Registro não encontrado'};
    
    const item = dados.itens[0]; 
    const custoTotal = (Number(item.custo) * Number(item.qtd)) + Number(dados.freteTotal || 0);
    
    let dataFormatada = dados.data;
    if (dados.data && dados.data.includes('-')) {
        const d = new Date(dados.data + 'T12:00:00');
        dataFormatada = Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
    } else { dataFormatada = obterDataAtual(); }

    getPlanilha().getSheetByName('Compras').getRange(linha, 2, 1, 7).setValues([[
      dataFormatada, item.modelo, Number(item.qtd), Number(item.custo), Number(dados.freteTotal), 'Em Estoque', dados.obs || ''
    ]]);
    
    atualizarFinanceiroVinculado(dados.id, `Compra (Estoque): ${item.modelo}`, custoTotal, dados.conta);
    logAcao('COMPRA_EDITADA', '', item.modelo, dados.id);
    return {sucesso: true};
  } catch(e) { return {sucesso: false, erro: e.toString()}; }
}

function editarVenda(dados) {
  try {
    const linha = encontrarLinha('Vendas', dados.id);
    if(!linha) return {sucesso: false, erro: 'Registro não encontrado'};
    
    let dataFormatada = dados.data;
    if (dados.data && dados.data.includes('-')) {
        const d = new Date(dados.data + 'T12:00:00');
        dataFormatada = Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
    } else { dataFormatada = obterDataAtual(); }

    const valorTotal = Number(dados.valorProduto) + Number(dados.valorFrete || 0) - Number(dados.custoFrete || 0);
    const sheet = getPlanilha().getSheetByName('Vendas');
    const origensAtuais = sheet.getRange(linha, 2).getValue(); 
    
    sheet.getRange(linha, 2, 1, 11).setValues([[
      origensAtuais, dataFormatada, 1, Number(dados.valorProduto), Number(dados.valorFrete || 0), 
      Number(dados.custoFrete || 0), dados.canal, dados.peca, dados.obs || '', dados.tipo, dados.cliente
    ]]);
    
    const descFin = `${dados.tipo}: ${dados.peca} | Cli: ${dados.cliente}`;
    atualizarFinanceiroVinculado(dados.id, descFin, valorTotal, dados.conta);
    return {sucesso: true};
  } catch(e) { return {sucesso: false, erro: e.toString()}; }
}

function editarCliente(dados) {
  try {
    const linha = encontrarLinha('Clientes', dados.id);
    if(!linha) return {sucesso: false, erro: 'Cliente não encontrado no banco de dados.'};
    
    getPlanilha().getSheetByName('Clientes').getRange(linha, 2, 1, 4).setValues([[
      dados.nome, dados.telefone, dados.canal, dados.obs || ''
    ]]);
    
    return {sucesso: true};
  } catch(e) { return {sucesso: false, erro: e.toString()}; }
}

function editarCatalogo(dados) {
  try {
    const linha = encontrarLinha('Catalogo', dados.id);
    if(!linha) return {sucesso: false, erro: 'Registro não encontrado'};
    
    let dataFormatada = dados.data;
    if (dados.data && dados.data.includes('-')) {
        const d = new Date(dados.data + 'T12:00:00');
        dataFormatada = Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
    } else { dataFormatada = obterDataAtual(); }

    getPlanilha().getSheetByName('Catalogo').getRange(linha, 2, 1, 8).setValues([[
      dataFormatada, dados.nome, dados.tipo, Number(dados.valor || 0), 
      dados.idOrigem || 'NENHUM', dados.canais || '', dados.status || 'Ativo', dados.obs || ''
    ]]);
    
    return {sucesso: true};
  } catch(e) { return {sucesso: false, erro: e.toString()}; }
}

function editarFinanceiroAvulso(dados) {
  try {
    const sheet = getPlanilha().getSheetByName('Financeiro');
    const linha = encontrarLinha('Financeiro', dados.id);
    if(!linha) return {sucesso: false, erro: 'Registro não encontrado'};
    
    // NOVO: Puxa a Origem que já estava lá, para NÃO quebrar o link com as Vendas/Compras!
    const idOrigemOriginal = sheet.getRange(linha, 8).getValue();
    
    let dataFormatada = dados.data;
    if (dados.data && dados.data.includes('-')) {
        const d = new Date(dados.data + 'T12:00:00');
        dataFormatada = Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.DATA_FORMAT);
    } else { dataFormatada = obterDataAtual(); }

    sheet.getRange(linha, 2, 1, 9).setValues([[
      dataFormatada, dados.tipo, dados.categoria, dados.descricao, Number(dados.valor), dados.conta, idOrigemOriginal, dados.catGlobal, dados.status
    ]]);
    return {sucesso: true};
  } catch(e) { return {sucesso: false, erro: e.toString()}; }
}

// --- EXCLUSÃO (DELETE) ---

function excluirRegistro(aba, id) {
  return withLock(() => {
    try {
      const linha = encontrarLinha(aba, id);
      if (!linha) return { sucesso: false, erro: 'ID não encontrado' };
      
      const sheet = getPlanilha().getSheetByName(aba);

      if (aba === 'Financeiro') {
        const rowData = sheet.getRange(linha, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (rowData[2].includes('Transferência')) {
          const idRelacionado = rowData[7];
          const data = sheet.getDataRange().getValues();
          for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][7] === idRelacionado) sheet.deleteRow(i + 1);
          }
          logAcao('EXCLUSAO_TRANSFERENCIA', '', `Cancelada`, idRelacionado);
          return { sucesso: true };
        }
      }

      sheet.deleteRow(linha);
      if (aba === 'Compras' || aba === 'Vendas') cascadeDeleteFinanceiro(id);
      logAcao(`EXCLUSAO_${aba.toUpperCase()}`, '', `ID: ${id}`, id);
      return { sucesso: true };
    } catch (e) { return { sucesso: false, erro: e.toString() }; }
  });
}

// --- AUXILIARES E INTEGRIDADE ---

function validarEstoqueVenda(idOrigem, qtdVender) {
  if (idOrigem === 'EST-0000') return { ok: true };
  const linha = encontrarLinha('Compras', idOrigem);
  if (!linha) return { ok: false, msg: 'Estoque não encontrado' };
  const qtdAtual = Number(getPlanilha().getSheetByName('Compras').getRange(linha, 4).getValue());
  return qtdAtual < qtdVender ? { ok: false, msg: `Estoque insuficiente: ${qtdAtual}` } : { ok: true };
}

function atualizarFinanceiroVinculado(idRelacionado, novaDesc, novoValor, novaConta) {
  const finSheet = getPlanilha().getSheetByName('Financeiro');
  const finData = finSheet.getDataRange().getValues();
  for(let i=1; i < finData.length; i++) {
    if(String(finData[i][7]).trim().toUpperCase() === String(idRelacionado).trim().toUpperCase()) {
      finSheet.getRange(i + 1, 5, 1, 3).setValues([[novaDesc, Number(novoValor), novaConta]]);
      break;
    }
  }
}

function cascadeDeleteFinanceiro(idRelacionado) {
  const finSheet = getPlanilha().getSheetByName('Financeiro');
  const finData = finSheet.getDataRange().getValues();
  for (let i = finData.length - 1; i >= 1; i--) {
    if (String(finData[i][7]).trim().toUpperCase() === String(idRelacionado).trim().toUpperCase()) {
      finSheet.deleteRow(i + 1);
    }
  }
}

function obterRegistroParaEdicao(aba, id) {
  try {
    const sheet = getPlanilha().getSheetByName(aba);
    if (!sheet) throw new Error(`A aba ${aba} não existe no banco de dados.`);
    
    const data = sheet.getDataRange().getValues();
    let dataRow = null;
    let idBusca = String(id).trim().toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === idBusca) {
        dataRow = i + 1;
        break;
      }
    }
    
    if (!dataRow) return null;
    
    const rawValues = sheet.getRange(dataRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const safeValues = rawValues.map(celula => {
      if (celula instanceof Date) { return celula.toISOString(); }
      return celula;
    });
    
    let obj = { id: safeValues[0], aba: aba, raw: safeValues };
    
    if (aba === 'Compras' || aba === 'Vendas') {
      const finSheet = getPlanilha().getSheetByName('Financeiro');
      if(finSheet && finSheet.getLastRow() > 1) {
        const finData = finSheet.getDataRange().getValues();
        for(let i=1; i<finData.length; i++) {
          if(String(finData[i][7]).trim().toUpperCase() === idBusca) { 
            obj.conta = finData[i][6]; 
            break; 
          }
        }
      }
    }
    return obj;
  } catch(e) {
    throw new Error("Erro interno ao buscar registro: " + e.message);
  }
}
