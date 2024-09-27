function calcularOrcamentoCombinado() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrcamento = spreadsheet.getSheetByName('Orçamento');
  const sheetDados = spreadsheet.getSheetByName('Dados');

  const dadosOrcamento = sheetOrcamento.getDataRange().getValues();
  const dadosProdutos = sheetDados.getDataRange().getValues();
  
  const valorFrete = sheetOrcamento.getRange('F2').getValue();
  const brinde = sheetOrcamento.getRange('G2').getValue();
  const perda = sheetOrcamento.getRange('H2').getValue(); 
  const descontoQuantidade = sheetOrcamento.getRange('I2').getValue();
  
  // Log dos valores fixos
  Logger.log(`Valor do Frete: R$ ${valorFrete}`);
  Logger.log(`Brinde: ${brinde}`);
  Logger.log(`Perda: ${perda}`);
  Logger.log(`Desconto por Quantidade: ${descontoQuantidade}`);

  let orcamentoTexto = '*ORÇAMENTO*\n\n'; // Título do orçamento em itálico
  let valorTotalProdutos = 0;

  for (let i = 1; i < dadosOrcamento.length; i++) {
    const nomeProduto = dadosOrcamento[i][0];
    const quantidade = dadosOrcamento[i][1];
    const alturaLugar = dadosOrcamento[i][2];
    const larguraLugar = dadosOrcamento[i][3];

    // Log dos dados do orçamento
    Logger.log(`Produto: ${nomeProduto}, Quantidade: ${quantidade}, Altura Lugar: ${alturaLugar}, Largura Lugar: ${larguraLugar}`);

    // Procurar produto na aba "Dados"
    let alturaProduto = 0, larguraProduto = 0, valorUnitario = 0, descontoTotal = 0, quantidadePecas = 0;
    
    for (let j = 1; j < dadosProdutos.length; j++) {
      if (dadosProdutos[j][0] === nomeProduto) {
        alturaProduto = dadosProdutos[j][1];
        larguraProduto = dadosProdutos[j][2];
        valorUnitario = dadosProdutos[j][3];

        // Log dos dados do produto
        Logger.log(`Encontrado produto na aba Dados: ${nomeProduto} | Altura: ${alturaProduto} | Largura: ${larguraProduto} | Valor Unitário: R$ ${valorUnitario}`);

        // Se o produto tem dimensões (produto por área)
        if (alturaLugar && larguraLugar) {
          const quantidadePecasLargura = Math.ceil(larguraLugar / larguraProduto); // Calcular quantidade na largura
          const quantidadePecasAltura = Math.ceil(alturaLugar / alturaProduto);   // Calcular quantidade na altura
          quantidadePecas = quantidadePecasLargura * quantidadePecasAltura; // Total de peças

        } else if (quantidade) { // Se tem quantidade (produto unitário)
          quantidadePecas = quantidade;
        }

        // Log da quantidade de peças calculadas
        Logger.log(`Quantidade de peças para ${nomeProduto}: ${quantidadePecas}`);

        // Aplicar cálculo de perda se a condição for "Sim"
        if (perda === "Sim") {
          quantidadePecas = Math.ceil(quantidadePecas * 1.1); // Adiciona 10%
          Logger.log(`Quantidade de peças ajustada para perda (10% a mais): ${quantidadePecas}`);
        }

        // Calcular valor total do produto sem o desconto inicialmente
        let valorTotalProduto = valorUnitario * quantidadePecas;

        // Aplicar desconto por quantidade de peças se I2 for "Sim"
        if (descontoQuantidade === "Sim") {
          descontoTotal = 0; // Resetar desconto total
          const descontoColunas = dadosProdutos[0].slice(11); // Cabeçalhos de desconto (quantidades)
          const descontoValores = dadosProdutos[j].slice(11); // Valores de desconto para o produto

          // Calcular o desconto total baseado na quantidade de peças
          for (let k = 0; k < descontoColunas.length; k++) {
            if (quantidadePecas >= descontoColunas[k]) { // Se a quantidade de peças for maior ou igual ao cabeçalho
              descontoTotal = descontoValores[k]; // Pega o desconto correspondente
            }
          }

          // Log do desconto total
          Logger.log(`Desconto total para ${nomeProduto} com ${quantidadePecas} peças: R$ ${descontoTotal}`);

          // Aplicar o desconto ao valor total do produto
          valorTotalProduto -= descontoTotal * quantidadePecas;
        }

        // Log do valor total do produto após desconto
        Logger.log(`Valor total do produto ${nomeProduto}: R$ ${valorTotalProduto.toFixed(2)}`);

        // Adicionar ao texto do orçamento com formatação em itálico
        orcamentoTexto += `*${nomeProduto}*\nQuantidade: ${quantidadePecas} unidades\n\n`;

        valorTotalProdutos += valorTotalProduto; // Somar ao valor total dos produtos
        break;
      }
    }
  }

  // Log do valor total dos produtos
  Logger.log(`Valor Total dos Produtos (sem frete): R$ ${valorTotalProdutos.toFixed(2)}`);

  // Adicionar o frete ao valor total geral
  const valorTotalGeral = valorTotalProdutos + valorFrete;

  // Log do valor total geral antes do parcelamento
  Logger.log(`Valor Total Geral (com frete): R$ ${valorTotalGeral.toFixed(2)}`);

  // Lógica de parcelamento
  let numParcelas = Math.min(10, Math.floor(valorTotalGeral / 150)); // até 10 parcelas
  let valorParcela = valorTotalGeral / numParcelas;

  // Ajustar número de parcelas até que o valor de cada parcela seja pelo menos R$ 150
  while (valorParcela < 150 && numParcelas > 1) {
    numParcelas--;
    valorParcela = valorTotalGeral / numParcelas;
  }

  // Se o valor total for abaixo de R$ 300, só pode ser 1x
  if (valorTotalGeral < 300) {
    numParcelas = 1;
    valorParcela = valorTotalGeral;
  }

  // Log do número de parcelas e valor de cada parcela
  Logger.log(`Número de parcelas: ${numParcelas}, Valor da parcela: R$ ${valorParcela.toFixed(2)}`);

  // Calcular o valor com desconto no Pix (5% de desconto) apenas sobre o valor dos produtos, sem o frete
  const valorPix = (valorTotalProdutos * 0.95) + valorFrete;

  // Log do valor com desconto no Pix
  Logger.log(`Valor no Pix (5% de desconto sobre os produtos, sem o frete): R$ ${valorPix.toFixed(2)}`);

  // Formatar valores finais para o formato brasileiro
  const valorTotalGeralFormatado = valorTotalGeral.toFixed(2).replace('.', ',');
  const valorParcelaFormatado = valorParcela.toFixed(2).replace('.', ',');
  const valorPixFormatado = valorPix.toFixed(2).replace('.', ',');

  // Adicionar detalhes de pagamento ao orçamento
  orcamentoTexto += `*Valor Final:* R$ ${valorTotalGeralFormatado}\n`; // Alterado para "*Valor Final:*"
  orcamentoTexto += `${numParcelas}x de R$ ${valorParcelaFormatado} no cartão sem juros\n`;
  orcamentoTexto += `R$ ${valorPixFormatado} no Pix\n\n`;

  // Brinde
  if (brinde === "Sim") {
    orcamentoTexto += '*Brinde:* Relógio de parede exclusivo\n\n';
  }

  // Observação sobre a entrega inclusa
  orcamentoTexto += 'OBS: Entrega já inclusa no valor\n';

  // Mostrar o orçamento na célula F3
  sheetOrcamento.getRange('F3').setValue(orcamentoTexto);

  // Log do orçamento final
  Logger.log(`Orçamento final:\n${orcamentoTexto}`);
}
