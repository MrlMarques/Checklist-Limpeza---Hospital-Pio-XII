/**
 * SCRIPT MESTRE DE AUTOMAÇÃO E IDENTIDADE VISUAL - HOSPITAL PIO XII
 * 1. Atualiza as opções do Formulário (Lista de Responsáveis pelo Checklist)
 * 2. Injeta todas as fórmulas do Dashboard com delimitadores em PT-BR (;)
 * 3. Aplica a paleta institucional Pio XII (#9A2B2E) e preserva os cards da logo
 * 4. Formata a aba de Respostas (cabeçalho congelado, zebra striping e card AO1)
 */
function executarAutomacaoCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss) {
    Logger.log("Erro: O script precisa rodar dentro da planilha ativa!");
    return;
  }

  // =========================================================================
  // PALETA DE CORES INSTITUCIONAL - HOSPITAL PIO XII
  // =========================================================================
  const COR_BORDO_PRINCIPAL = "#9A2B2E"; // Bordô institucional Pio XII
  const COR_BORDO_DESTAQUE  = "#F6E6E6"; // Bordô suave para linhas de total
  const COR_CINZA_SUBTITULO = "#9E9E9E"; // Cinza institucional para rótulos
  const COR_CINZA_ZEBRA     = "#F5F5F5"; // Cinza gelo para linhas alternadas
  const COR_CINZA_BORDA     = "#E0E0E0"; // Cinza claro para contornos suaves
  const COR_BRANCO          = "#FFFFFF"; // Branco puro

  // =========================================================================
  // ETAPA 1: ATUALIZAR AS OPÇÕES NO GOOGLE FORMS VINCULADO
  // =========================================================================
  const novaListaResponsaveis = [
    "Fernanda",
    "Lúcia",
    "Rayane",
    "Adriana",
    "Maria"
  ];

  try {
    const formUrl = ss.getFormUrl();
    if (formUrl) {
      const form = FormApp.openByUrl(formUrl);
      const itens = form.getItems();
      
      for (let i = 0; i < itens.length; i++) {
        const item = itens[i];
        const titulo = item.getTitle().toLowerCase();
        
        if (titulo.indexOf("responsável pelo checklist") !== -1 || titulo.indexOf("responsavel pelo checklist") !== -1) {
          if (item.getType() === FormApp.ItemType.LIST) {
            item.asListItem().setChoiceValues(novaListaResponsaveis);
          } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
            item.asMultipleChoiceItem().setChoiceValues(novaListaResponsaveis);
          }
          Logger.log("✔ Formulário: Lista de responsáveis atualizada com sucesso!");
          break;
        }
      }
    } else {
      Logger.log("⚠️ Aviso: Formulário vinculado não detectado diretamente pela URL.");
    }
  } catch (erroForm) {
    Logger.log("⚠️ Nota no Formulário: " + erroForm.message);
  }

  // =========================================================================
  // ETAPA 2: RECONSTRUIR FÓRMULAS DO DASHBOARD (SINTAXE PT-BR COM ';')
  // =========================================================================
  const abaDashboard = ss.getSheetByName("Dashboard");
  
  if (!abaDashboard) {
    Logger.log("Erro: Aba 'Dashboard' não encontrada na planilha!");
    return;
  }

  const abaResp = "'Respostas ao formulário 2'";

  // 1. Cartões de Indicadores Globais - KPIs (Linha 3)
  abaDashboard.getRange("B3").setFormula(`=COUNTA(${abaResp}!A2:A)`);
  abaDashboard.getRange("D3").setFormula(`=IF((COUNTIF(${abaResp}!G2:AL; "Limpeza Conforme") + COUNTIF(${abaResp}!G2:AL; "Não Conforme"))=0; 0; COUNTIF(${abaResp}!G2:AL; "Limpeza Conforme") / (COUNTIF(${abaResp}!G2:AL; "Limpeza Conforme") + COUNTIF(${abaResp}!G2:AL; "Não Conforme")))`);
  abaDashboard.getRange("F3").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Não Conforme")`);
  abaDashboard.getRange("H3").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Limpeza Conforme")`);
  abaDashboard.getRange("J3").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Não se Aplica")`);

  // 2. Consolidação Geral de Conformidade (Linhas 6 a 9)
  abaDashboard.getRange("C6").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Limpeza Conforme")`);
  abaDashboard.getRange("C7").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Não Conforme")`);
  abaDashboard.getRange("C8").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Não se Aplica")`);
  abaDashboard.getRange("C9").setFormula(`=SUM(C6:C8)`);

  abaDashboard.getRange("D6").setFormula(`=IF(C9=0; 0; C6/C9)`);
  abaDashboard.getRange("D7").setFormula(`=IF(C9=0; 0; C7/C9)`);
  abaDashboard.getRange("D8").setFormula(`=IF(C9=0; 0; C8/C9)`);
  abaDashboard.getRange("D9").setFormula(`=SUM(D6:D8)`);

  // 3. Desempenho por Responsável pelo Checklist (Coluna AN + SUMPRODUCT)
  const responsaveisConfig = [
    { nome: "Fernanda", linha: 13 },
    { nome: "Lúcia",    linha: 14 },
    { nome: "Rayane",   linha: 15 },
    { nome: "Adriana",  linha: 16 },
    { nome: "Maria",    linha: 17 }
  ];

  responsaveisConfig.forEach(resp => {
    abaDashboard.getRange(`B${resp.linha}`).setValue(resp.nome);
    abaDashboard.getRange(`C${resp.linha}`).setFormula(`=COUNTIF(${abaResp}!AN2:AN; "${resp.nome}")`);
    abaDashboard.getRange(`D${resp.linha}`).setFormula(`=SUMPRODUCT((${abaResp}!AN2:AN="${resp.nome}") * (${abaResp}!G2:AL="Não Conforme"))`);
    abaDashboard.getRange(`E${resp.linha}`).setFormula(`=IF(C${resp.linha}=0; 0; D${resp.linha}/C${resp.linha})`);
  });

  // Outros / Novos e Total Consolidado
  abaDashboard.getRange("B18").setValue("Outros / Novos");
  abaDashboard.getRange("C18").setFormula(`=COUNTA(${abaResp}!AN2:AN) - SUM(C13:C17)`);
  abaDashboard.getRange("D18").setFormula(`=COUNTIF(${abaResp}!G2:AL; "Não Conforme") - SUM(D13:D17)`);
  abaDashboard.getRange("E18").setFormula(`=IF(C18<=0; 0; D18/C18)`);

  abaDashboard.getRange("B19").setValue("Total Consolidado");
  abaDashboard.getRange("C19").setFormula(`=SUM(C13:C18)`);
  abaDashboard.getRange("D19").setFormula(`=SUM(D13:D18)`);
  abaDashboard.getRange("E19").setFormula(`=IF(C19=0; 0; D19/C19)`);

  // 4. Detalhamento por Item Inspecionado (Linhas 26 a 57)
  const itens = [
    { linha: 26, col: "G" },  // Teto
    { linha: 27, col: "H" },  // Parede
    { linha: 28, col: "I" },  // Régua de Gases
    { linha: 29, col: "J" },  // Cama
    { linha: 30, col: "K" },  // Campainha
    { linha: 31, col: "L" },  // Escada beira leito
    { linha: 32, col: "M" },  // Suporte de soro
    { linha: 33, col: "N" },  // Mesa de refeição
    { linha: 34, col: "O" },  // Mesa de cabeceira
    { linha: 35, col: "P" },  // Poltrona / Sofá
    { linha: 36, col: "Q" },  // Televisão
    { linha: 37, col: "R" },  // Controle TV/Ar condicionado
    { linha: 38, col: "S" },  // Janelas / Telas
    { linha: 39, col: "T" },  // Persiana janela
    { linha: 40, col: "U" },  // Pia do quarto
    { linha: 41, col: "V" },  // Lixeiras
    { linha: 42, col: "W" },  // Cortina divisória de leitos
    { linha: 43, col: "X" },  // Frigobar
    { linha: 44, col: "Y" },  // Armários
    { linha: 45, col: "Z" },  // Mesa com cadeira paciente
    { linha: 46, col: "AA" }, // Dispenser
    { linha: 47, col: "AB" }, // Quadro
    { linha: 48, col: "AC" }, // Interruptor quarto
    { linha: 49, col: "AD" }, // Portas
    { linha: 50, col: "AE" }, // Maçaneta interna do quarto
    { linha: 51, col: "AF" }, // Maçaneta interna do banheiro
    { linha: 52, col: "AG" }, // Interruptor banheiro
    { linha: 53, col: "AH" }, // Pia do banheiro
    { linha: 54, col: "AI" }, // Vaso sanitário
    { linha: 55, col: "AJ" }, // Ducha higienica
    { linha: 56, col: "AK" }, // Barra de apoio
    { linha: 57, col: "AL" }  // Ralo
  ];

  itens.forEach(item => {
    abaDashboard.getRange(`C${item.linha}`).setFormula(`=COUNTIF(${abaResp}!$${item.col}2:$${item.col}; "Não Conforme")`);
    abaDashboard.getRange(`D${item.linha}`).setFormula(`=COUNTIF(${abaResp}!$${item.col}2:$${item.col}; "Limpeza Conforme")`);
    abaDashboard.getRange(`E${item.linha}`).setFormula(`=IF((C${item.linha}+D${item.linha})=0; 0; C${item.linha}/(C${item.linha}+D${item.linha}))`);
  });

  Logger.log("✔ Planilha: Fórmulas do Dashboard atualizadas com sucesso!");

  // =========================================================================
  // ETAPA 3: ESTILIZAÇÃO VISUAL INSTITUCIONAL DO DASHBOARD
  // =========================================================================
  abaDashboard.setHiddenGridlines(true);

  // Card da Logo (B1:B2) - Fundo branco para contraste com logo transparente (não apaga imagem)
  const cardLogoDash = abaDashboard.getRange("B1:B2");
  cardLogoDash.setBackground(COR_BRANCO);
  cardLogoDash.setHorizontalAlignment("center");
  cardLogoDash.setVerticalAlignment("middle");
  cardLogoDash.setBorder(true, true, true, true, false, false, COR_BORDO_PRINCIPAL, SpreadsheetApp.BorderStyle.SOLID);

  // Banner Institucional (C1:J2)
  const banner = abaDashboard.getRange("C1:J2");
  banner.setBackground(COR_BORDO_PRINCIPAL);
  banner.setFontColor(COR_BRANCO);
  banner.setFontFamily("Roboto");
  banner.setHorizontalAlignment("center");
  banner.setVerticalAlignment("middle");

  // Cartões KPI (Valores na Linha 3, Subtítulos na Linha 4)
  const valoresKPI = abaDashboard.getRange("B3:J3");
  valoresKPI.setFontSize(16).setFontWeight("bold").setFontColor(COR_BORDO_PRINCIPAL);
  
  const subtitulosKPI = abaDashboard.getRange("B4:J4");
  subtitulosKPI.setFontSize(9).setFontColor(COR_CINZA_SUBTITULO);

  // Cabeçalhos das Tabelas
  ["B5:E5", "B11:E12", "B24:E25"].forEach(intervalo => {
    const cabecalho = abaDashboard.getRange(intervalo);
    cabecalho.setBackground(COR_BORDO_PRINCIPAL);
    cabecalho.setFontColor(COR_BRANCO);
    cabecalho.setFontWeight("bold");
    cabecalho.setFontFamily("Roboto");
  });

  // Formatação de Números e Percentuais
  ["C6:C9", "C13:D19", "C26:D57"].forEach(intervalo => {
    abaDashboard.getRange(intervalo).setNumberFormat("#,##0");
  });
  
  ["D6:D9", "E13:E19", "E26:E57"].forEach(intervalo => {
    abaDashboard.getRange(intervalo).setNumberFormat("0.0%");
  });

  // Destaque das Linhas de Total
  ["B9:D9", "B19:E19", "B58:E58"].forEach(intervalo => {
    const linhaTotal = abaDashboard.getRange(intervalo);
    linhaTotal.setBackground(COR_BORDO_DESTAQUE);
    linhaTotal.setFontWeight("bold");
    linhaTotal.setFontColor(COR_BORDO_PRINCIPAL);
  });

  // Bordas Suaves na área de conteúdo
  abaDashboard.getRange("B3:J58").setBorder(true, true, true, true, true, true, COR_CINZA_BORDA, SpreadsheetApp.BorderStyle.SOLID);

  Logger.log("✔ Estilização institucional do Dashboard aplicada!");

  // =========================================================================
  // ETAPA 4: FORMATAÇÃO VISUAL DA ABA DE RESPOSTAS DO FORMULÁRIO
  // =========================================================================
  const abaRespostas = ss.getSheetByName("Respostas ao formulário 2");
  
  if (abaRespostas) {
    const ultimaLinhaResp = abaRespostas.getLastRow();
    // Garante leitura até a coluna AN (coluna 40), ignorando o card da logo em AO1
    const colFinalDados = 40; 
    
    if (ultimaLinhaResp >= 1) {
      // 1. Centralizar todo o conteúdo horizontal e verticalmente
      const intervaloGeralResp = abaRespostas.getRange(1, 1, abaRespostas.getMaxRows(), colFinalDados);
      intervaloGeralResp.setHorizontalAlignment("center");
      intervaloGeralResp.setVerticalAlignment("middle");
      
      // 2. Cores alternadas (Zebra striping) nas linhas de dados
      if (ultimaLinhaResp >= 2) {
        const intervaloDadosResp = abaRespostas.getRange(2, 1, ultimaLinhaResp - 1, colFinalDados);
        
        const fundosResp = [];
        for (let i = 0; i < (ultimaLinhaResp - 1); i++) {
          const cor = (i % 2 === 0) ? COR_BRANCO : COR_CINZA_ZEBRA;
          const linhaFundo = [];
          for (let j = 0; j < colFinalDados; j++) {
            linhaFundo.push(cor);
          }
          fundosResp.push(linhaFundo);
        }
        intervaloDadosResp.setBackgrounds(fundosResp);
      }

      // 3. Estilizar Cabeçalho da Aba de Respostas (A1:AN1)
      abaRespostas.setRowHeight(1, 50);
      abaRespostas.setFrozenRows(1);
      const cabecalhoResp = abaRespostas.getRange(1, 1, 1, colFinalDados);
      cabecalhoResp.setBackground(COR_BORDO_PRINCIPAL);
      cabecalhoResp.setFontColor(COR_BRANCO);
      cabecalhoResp.setFontWeight("bold");
      cabecalhoResp.setFontFamily("Roboto");
      cabecalhoResp.setWrap(true);

      // 4. Preservar Card Branco da Logo na célula AO1
      const cardLogoResp = abaRespostas.getRange("AO1");
      cardLogoResp.setBackground(COR_BRANCO);
      cardLogoResp.setHorizontalAlignment("center");
      cardLogoResp.setVerticalAlignment("middle");
      cardLogoResp.setBorder(true, true, true, true, false, false, COR_BORDO_PRINCIPAL, SpreadsheetApp.BorderStyle.SOLID);
      abaRespostas.setColumnWidth(41, 160); // Coluna AO = 41
    }
    Logger.log("✔ Aba de Respostas formatada com paleta Pio XII e card AO1 preservado!");
  }
}

/**
 * PADRONIZAÇÃO DE LAYOUT E TÍTULOS DO GOOGLE FORMS - HOSPITAL PIO XII
 * Atualiza o título geral, descrição e padroniza o formato de todas as perguntas.
 */
function atualizarLayoutFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formUrl = ss.getFormUrl();
  
  if (!formUrl) {
    Logger.log("⚠️ Nenhum formulário vinculado a esta planilha!");
    return;
  }
  
  const form = FormApp.openByUrl(formUrl);
  
  // 1. Título e Descrição Institucional Oficial
  form.setTitle("• Checklist de Limpeza e Inspeção Hospitalar •");
  form.setDescription("Hospital Pio XII • Governança e Higienização\nPreencha todos os campos obrigatórios da inspeção com atenção.");
  
  // 2. Padronização automática dos títulos das perguntas existentes
  const itens = form.getItems();
  
  itens.forEach(item => {
    let tituloOriginal = item.getTitle();
    
    // Remove marcadores antigos duplicados para evitar sujeira no texto
    let tituloLimpo = tituloOriginal.replace(/•/g, "").trim();
    
    // Reaplica o padrão visual institucional com bolinhas nas pontas
    if (tituloLimpo.length > 0) {
      item.setTitle(`• ${tituloLimpo} •`);
    }
  });
  
  Logger.log("✔ Layout do formulário padronizado com sucesso!");
}

/**
 * 1. OBRIGATÓRIA PARA ABRIR A PÁGINA: Renderiza o arquivo index.html no navegador com Viewport Mobile
 */
/**
 * ROTEADOR WEB APP: Abre o Checklist ou o Dashboard Executivo (?page=dashboard)
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Dashboard Executivo • Hospital Pio XII')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Checklist de Limpeza - Hospital Pio XII')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * MOTOR ANALÍTICO DO DASHBOARD: Processa todos os parâmetros da aba de respostas
 */
function obterDadosDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaResp = ss.getSheetByName('Respostas ao formulário 2');

  if (!abaResp) {
    throw new Error("Aba 'Respostas ao formulário 2' não encontrada na planilha.");
  }

  const ultimaLinha = abaResp.getLastRow();
  const nomesItens = [
    // 25 Itens do Quarto (índices 0 a 24 -> colunas G a AE)
    { nome: 'Teto', ambiente: 'Quarto' },
    { nome: 'Parede', ambiente: 'Quarto' },
    { nome: 'Régua de Gases', ambiente: 'Quarto' },
    { nome: 'Cama', ambiente: 'Quarto' },
    { nome: 'Campainha', ambiente: 'Quarto' },
    { nome: 'Escada beira leito', ambiente: 'Quarto' },
    { nome: 'Suporte de soro', ambiente: 'Quarto' },
    { nome: 'Mesa de refeição', ambiente: 'Quarto' },
    { nome: 'Mesa de cabeceira', ambiente: 'Quarto' },
    { nome: 'Poltrona / Sofá', ambiente: 'Quarto' },
    { nome: 'Televisão', ambiente: 'Quarto' },
    { nome: 'Controle TV/Ar', ambiente: 'Quarto' },
    { nome: 'Janelas / Telas', ambiente: 'Quarto' },
    { nome: 'Persiana janela', ambiente: 'Quarto' },
    { nome: 'Pia do quarto', ambiente: 'Quarto' },
    { nome: 'Lixeiras', ambiente: 'Quarto' },
    { nome: 'Cortina divisória', ambiente: 'Quarto' },
    { nome: 'Frigobar', ambiente: 'Quarto' },
    { nome: 'Armários', ambiente: 'Quarto' },
    { nome: 'Mesa / Cadeira', ambiente: 'Quarto' },
    { nome: 'Dispenser', ambiente: 'Quarto' },
    { nome: 'Quadro', ambiente: 'Quarto' },
    { nome: 'Interruptor quarto', ambiente: 'Quarto' },
    { nome: 'Portas', ambiente: 'Quarto' },
    { nome: 'Maçaneta quarto', ambiente: 'Quarto' },
    // 7 Itens do Banheiro (índices 25 a 31 -> colunas AF a AL)
    { nome: 'Maçaneta banheiro', ambiente: 'Banheiro' },
    { nome: 'Interruptor banheiro', ambiente: 'Banheiro' },
    { nome: 'Pia do banheiro', ambiente: 'Banheiro' },
    { nome: 'Vaso sanitário', ambiente: 'Banheiro' },
    { nome: 'Ducha higiênica', ambiente: 'Banheiro' },
    { nome: 'Barra de apoio', ambiente: 'Banheiro' },
    { nome: 'Ralo', ambiente: 'Banheiro' }
  ];

  // Estrutura inicial vazia caso não haja registros
  if (ultimaLinha < 2) {
    return {
      atualizadoEm: Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
      kpis: { totalInspecoes: 0, conformes: 0, naoConformes: 0, naoSeAplica: 0, taxaConformidade: 0, mediaFalhasPorInspecao: 0 },
      ambientes: { quarto: { conf: 0, nconf: 0, taxa: 0 }, banheiro: { conf: 0, nconf: 0, taxa: 0 } },
      tiposLimpeza: [],
      evolucaoDiaria: [],
      itensDetalhados: nomesItens.map(i => ({ ...i, conf: 0, nconf: 0, na: 0, taxaFalha: 0 })),
      auditores: [],
      executoresLimpeza: [],
      historico: []
    };
  }

  // Lê todas as 40 colunas (A até AN) a partir da linha 2
  const dados = abaResp.getRange(2, 1, ultimaLinha - 1, 40).getValues();

  let totalInspecoes = 0;
  let totalConformes = 0;
  let totalNaoConformes = 0;
  let totalNaoAplica = 0;

  let quartoConf = 0, quartoNConf = 0;
  let banhConf = 0, banhNConf = 0;

  const statsItens = nomesItens.map(item => ({
    nome: item.nome,
    ambiente: item.ambiente,
    conf: 0,
    nconf: 0,
    na: 0
  }));

  const mapTipos = {};
  const mapDatas = {};
  const mapAuditores = {};
  const mapExecutores = {};
  const historico = [];

  dados.forEach(linha => {
    // Ignora linhas completamente vazias
    if (!linha[1] && !linha[3]) return;

    totalInspecoes++;

    const dataFormatada = linha[1] instanceof Date
      ? Utilities.formatDate(linha[1], 'America/Sao_Paulo', 'dd/MM/yyyy')
      : String(linha[1] || '').trim();

    const dataCurta = dataFormatada.length >= 5 ? dataFormatada.substring(0, 5) : dataFormatada;

    const horarioFormatado = linha[2] instanceof Date
      ? Utilities.formatDate(linha[2], 'America/Sao_Paulo', 'HH:mm')
      : String(linha[2] || '').trim();

    const leito = String(linha[3] || 'Não informado').trim();
    const tipoLimpeza = String(linha[4] || 'Não especificado').trim();
    const executor = String(linha[5] || 'Não informado').trim();
    const observacao = String(linha[38] || '').trim();
    const auditor = String(linha[39] || 'Outros / Não informado').trim();

    let inspConf = 0;
    let inspNConf = 0;
    let inspNA = 0;
    const itensFalhosNaInspecao = [];

    // Percorre os 32 itens (colunas G [índice 6] até AL [índice 37])
    for (let i = 0; i < 32; i++) {
      const status = String(linha[6 + i] || '').trim();
      const ehQuarto = i < 25;

      if (status === 'Limpeza Conforme') {
        totalConformes++;
        inspConf++;
        statsItens[i].conf++;
        if (ehQuarto) quartoConf++; else banhConf++;
      } else if (status === 'Não Conforme') {
        totalNaoConformes++;
        inspNConf++;
        statsItens[i].nconf++;
        itensFalhosNaInspecao.push(nomesItens[i].nome);
        if (ehQuarto) quartoNConf++; else banhNConf++;
      } else if (status === 'Não se Aplica') {
        totalNaoAplica++;
        inspNA++;
        statsItens[i].na++;
      }
    }

    // Agrupa por Tipo de Limpeza (Leito vs Quarto)
    if (!mapTipos[tipoLimpeza]) mapTipos[tipoLimpeza] = { tipo: tipoLimpeza, inspecoes: 0, conf: 0, nconf: 0 };
    mapTipos[tipoLimpeza].inspecoes++;
    mapTipos[tipoLimpeza].conf += inspConf;
    mapTipos[tipoLimpeza].nconf += inspNConf;

    // Agrupa por Data (Evolução)
    if (dataCurta) {
      if (!mapDatas[dataCurta]) mapDatas[dataCurta] = { data: dataCurta, inspecoes: 0, conf: 0, nconf: 0 };
      mapDatas[dataCurta].inspecoes++;
      mapDatas[dataCurta].conf += inspConf;
      mapDatas[dataCurta].nconf += inspNConf;
    }

    // Agrupa por Auditora (Responsável pelo Checklist)
    if (!mapAuditores[auditor]) mapAuditores[auditor] = { nome: auditor, inspecoes: 0, conf: 0, nconf: 0 };
    mapAuditores[auditor].inspecoes++;
    mapAuditores[auditor].conf += inspConf;
    mapAuditores[auditor].nconf += inspNConf;

    // Agrupa por Colaborador da Limpeza
    if (!mapExecutores[executor]) mapExecutores[executor] = { nome: executor, inspecoes: 0, conf: 0, nconf: 0 };
    mapExecutores[executor].inspecoes++;
    mapExecutores[executor].conf += inspConf;
    mapExecutores[executor].nconf += inspNConf;

    // Adiciona ao histórico
    const validosInsp = inspConf + inspNConf;
    historico.push({
      data: dataFormatada,
      horario: horarioFormatado,
      leito,
      tipoLimpeza,
      executor,
      auditor,
      observacao,
      conf: inspConf,
      nconf: inspNConf,
      na: inspNA,
      taxaAprovacao: validosInsp > 0 ? Math.round((inspConf / validosInsp) * 100) : 100,
      falhas: itensFalhosNaInspecao
    });
  });

  const totalValidos = totalConformes + totalNaoConformes;
  const taxaConformidade = totalValidos > 0 ? (totalConformes / totalValidos) * 100 : 0;
  const mediaFalhasPorInspecao = totalInspecoes > 0 ? (totalNaoConformes / totalInspecoes) : 0;

  const quartoValidos = quartoConf + quartoNConf;
  const banhValidos = banhConf + banhNConf;

  const itensDetalhados = statsItens.map(item => {
    const validos = item.conf + item.nconf;
    return {
      ...item,
      taxaFalha: validos > 0 ? Number(((item.nconf / validos) * 100).toFixed(1)) : 0,
      taxaConformidade: validos > 0 ? Number(((item.conf / validos) * 100).toFixed(1)) : 100
    };
  });

  const tiposLimpeza = Object.values(mapTipos).map(t => {
    const v = t.conf + t.nconf;
    return { ...t, taxa: v > 0 ? Number(((t.conf / v) * 100).toFixed(1)) : 0 };
  });

  const evolucaoDiaria = Object.values(mapDatas).slice(-10).map(d => {
    const v = d.conf + d.nconf;
    return { ...d, taxa: v > 0 ? Number(((d.conf / v) * 100).toFixed(1)) : 0 };
  });

  const auditores = Object.values(mapAuditores)
    .map(a => {
      const v = a.conf + a.nconf;
      return { ...a, taxaConformidade: v > 0 ? Number(((a.conf / v) * 100).toFixed(1)) : 0 };
    })
    .sort((a, b) => b.inspecoes - a.inspecoes);

  const executoresLimpeza = Object.values(mapExecutores)
    .map(e => {
      const v = e.conf + e.nconf;
      return { ...e, taxaConformidade: v > 0 ? Number(((e.conf / v) * 100).toFixed(1)) : 0 };
    })
    .sort((a, b) => b.inspecoes - a.inspecoes);

  return {
    atualizadoEm: Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
    kpis: {
      totalInspecoes,
      conformes: totalConformes,
      naoConformes: totalNaoConformes,
      naoSeAplica: totalNaoAplica,
      taxaConformidade: Number(taxaConformidade.toFixed(1)),
      mediaFalhasPorInspecao: Number(mediaFalhasPorInspecao.toFixed(1))
    },
    ambientes: {
      quarto: {
        conf: quartoConf,
        nconf: quartoNConf,
        taxa: quartoValidos > 0 ? Number(((quartoConf / quartoValidos) * 100).toFixed(1)) : 0
      },
      banheiro: {
        conf: banhConf,
        nconf: banhNConf,
        taxa: banhValidos > 0 ? Number(((banhConf / banhValidos) * 100).toFixed(1)) : 0
      }
    },
    tiposLimpeza,
    evolucaoDiaria,
    itensDetalhados,
    auditores,
    executoresLimpeza,
    historico: historico.reverse()
  };
}
/**
 * FUNÇÃO doPost CORRIGIDA E ALINHADA COM AS COLUNAS DA PLANILHA
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaRespostas = ss.getSheetByName("Respostas ao formulário 2");
    
    if (!abaRespostas) {
      abaRespostas = ss.insertSheet("Respostas ao formulário 2");
    }

    // Lê os dados enviados em JSON pelo index.html
    const dados = JSON.parse(e.postData.contents);
    const carimboDataHora = new Date();
    
    // ORDEM EXATA DAS COLUNAS DA SUA PLANILHA:
    const novaLinha = [
      carimboDataHora,                          // 1. Carimbo de data/hora
      dados.data || "",                         // 2. Data
      dados.horario || "",                      // 3. Horário
      dados.leito || "",                        // 4. Leito
      dados.tipo_limpeza || "",                 // 5. Tipo de Limpeza
      dados.responsavel || "",                  // 6. Responsável pela Limpeza
      
      // Itens do Quarto (1 a 25)
      dados.teto || "",
      dados.parede || "",
      dados.regua_gases || "",
      dados.cama || "",
      dados.campainha || "",
      dados.escada_leito || "",
      dados.suporte_soro || "",
      dados.mesa_refeicao || "",
      dados.mesa_cabeceira || "",
      dados.poltrona_sofa || "",
      dados.televisao || "",
      dados.controle_tv_ar || "",
      dados.janelas_telas || "",
      dados.persiana_janela || "",
      dados.pia_quarto || "",
      dados.lixeiras || "",
      dados.cortina_divisoria || "",
      dados.frigobar || "",
      dados.armarios || "",
      dados.mesa_cadeira_paciente || "",
      dados.dispenser || "",
      dados.quadro || "",
      dados.interruptor_quarto || "",
      dados.portas || "",
      dados.macaneta_quarto || "",
      
      // Itens do Banheiro (1 a 7)
      dados.macaneta_banheiro || "",
      dados.interruptor_banheiro || "",
      dados.pia_banheiro || "",
      dados.vaso_sanitario || "",
      dados.ducha_higienica || "",
      dados.barra_apoio || "",
      dados.ralo || "",
      
      // Finalização
      dados.observacoes || "",                  // Observação
      dados.responsavel_checklist || ""         // Responsável pelo Checklist (Auditora)
    ];

    // Insere a linha organizada na planilha
    abaRespostas.appendRow(novaLinha);

    return ContentService
      .createTextOutput(JSON.stringify({ "status": "sucesso" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (erro) {
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "erro", "mensagem": erro.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
