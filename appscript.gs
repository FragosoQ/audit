// Google Apps Script - Cole este código no Editor de Apps Script do Google Sheet
// 1. Abra o Google Sheet
// 2. Menu: Extensões > Apps Script
// 3. Limpe o código predefinido
// 4. Cole todo este código
// 5. Clique em "Deploy" > "New Deployment"
// 6. Selecione "Web app"
// 7. Execute como: [sua conta]
// 8. Quem tem acesso: Qualquer um
// 9. Copie o URL de deployment
// 10. Cole o URL no HTML em APPSCRIPT_URL

function doPost(e) {
  try {
    let data;
    const contentType = (e.postData && e.postData.type ? e.postData.type : '').toLowerCase();
    const contents = e.postData && e.postData.contents ? e.postData.contents : '';
    
    // Lidar com JSON ou form data
    if (contentType.includes('application/json') || contentType.includes('text/plain')) {
      data = JSON.parse(contents);
    } else if (e.parameter.data) {
      // Dados vindos de um formulário
      data = JSON.parse(e.parameter.data);
    } else {
      throw new Error('Dados inválidos');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const operacao = data.operacao;
    
    if (operacao === 'salvarIntro') {
      adicionarAuditoriaPlanificada(spreadsheet, data);
    } else if (operacao === 'salvarAuditoria') {
      adicionarRegistosAuditoria(spreadsheet, data);
    } else if (operacao === 'adicionarQuestao') {
      adicionarQuestao(spreadsheet, data);
    } else if (operacao === 'apagarIntro') {
      apagarAuditoriaPlanificada(spreadsheet, data);
    } else if (operacao === 'obterQuestoes') {
      const questoes = obterQuestoes(spreadsheet);
      return HtmlService.createHtmlOutput(JSON.stringify(questoes));
    }
    
    // Retornar resposta silenciosa (sem abrir janela)
    return HtmlService.createHtmlOutput('<script>window.close();</script>').setWidth(1).setHeight(1);
    
  } catch (erro) {
    Logger.log('Erro em doPost: ' + erro);
    return HtmlService.createHtmlOutput('<script>window.close();</script>').setWidth(1).setHeight(1);
  }
}

function doGet(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const operacao = e.parameter.operacao;
    
    Logger.log('=== doGet iniciado ===');
    Logger.log('Operação:', operacao);
    
    if (operacao === 'obterQuestoes') {
      Logger.log('Iniciando obtenção de questões...');
      const questoes = obterQuestoes(spreadsheet);
      Logger.log('Questões obtidas:', JSON.stringify(questoes));
      return ContentService.createTextOutput(JSON.stringify(questoes))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (operacao === 'obterAuditorias') {
      Logger.log('Iniciando obtenção de auditorias...');
      const auditorias = obterAuditorias(spreadsheet);
      Logger.log('Auditorias obtidas:', JSON.stringify(auditorias));
      return ContentService.createTextOutput(JSON.stringify(auditorias))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    Logger.log('Operação desconhecida:', operacao);
    return ContentService.createTextOutput(JSON.stringify({ erro: 'Operação desconhecida: ' + operacao }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (erro) {
    Logger.log('ERRO em doGet:', erro.toString());
    Logger.log('Stack:', erro.stack);
    return ContentService.createTextOutput(JSON.stringify({ 
      erro: erro.message,
      stack: erro.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Obter todos os registos da folha AUDITORIA
function obterAuditorias(spreadsheet) {
  try {
    Logger.log('=== Iniciando obterAuditorias ===');

    const sheet = spreadsheet.getSheetByName('AUDITORIA');
    if (!sheet) {
      Logger.log('❌ Sheet AUDITORIA não encontrado');
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log('Dimensões AUDITORIA: ' + lastRow + ' linhas x ' + lastCol + ' colunas');

    if (lastRow <= 1) {
      Logger.log('Sheet AUDITORIA vazio ou só com headers');
      return [];
    }

    const values = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 7)).getValues();
    const registos = [];

    values.forEach((row) => {
      const temDados = row.some(cell => cell !== null && cell !== '');
      if (!temDados) return;

      registos.push({
        id: String(row[0] || ''),
        departamento: String(row[1] || ''),
        investigacao: String(row[2] || ''),
        foco: String(row[3] || ''),
        evidencia: String(row[4] || ''),
        om: String(row[5] || ''),
        nc: String(row[6] || '')
      });
    });

    Logger.log('✅ Total de registos AUDITORIA:', registos.length);
    return registos;
  } catch (erro) {
    Logger.log('❌ ERRO em obterAuditorias: ' + erro.toString());
    return [];
  }
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// Adicionar auditoria planeada à folha INTRO
function adicionarAuditoriaPlanificada(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('INTRO');
    
    // Headers esperados
    const headers = [
      'ID', 'Departamento', 'Auditor Coordenador', 'Auditor', 'Auditado',
      'Documentos de Referência da Empresa', 'Registado Por', 'Colaboradores Contactados',
      'Requisitos', 'Data de Auditoria', 'Hora Prevista', 'Hora da Auditoria'
    ];
    
    // Verificar se headers existem, se não, adicionar
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      for (let i = 0; i < headers.length; i++) {
        sheet.getRange(1, i + 1).setValue(headers[i]);
      }
    }
    
    // Adicionar nova linha de dados
    const novaLinha = sheet.getLastRow() + 1;
    
    const valores = [
      data.id,
      data.departamento,
      data.auditorCoord,
      data.auditor,
      data.auditado,
      data.documentosRef,
      data.registadoPor,
      data.colaboradores,
      data.requisitos,
      data.dataAuditoria,
      data.horaPrevista,
      data.horaAuditoria
    ];
    
    for (let i = 0; i < valores.length; i++) {
      sheet.getRange(novaLinha, i + 1).setValue(valores[i]);
    }
  } catch (erro) {
    Logger.log('Erro em adicionarAuditoriaPlanificada: ' + erro);
    throw erro;
  }
}

// Adicionar registos de auditoria à folha AUDITORIA
function adicionarRegistosAuditoria(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('AUDITORIA');
    
    // Headers esperados conforme especificado
    const headers = [
      'ID', 'DEPARTAMENTO', 'Investigação (Questão de Auditoria)',
      'Foco / Requisito', 'CONSTATAÇÃO / EVIDÊNCIA', 'OM', 'NC'
    ];
    
    // Verificar se headers existem
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      for (let i = 0; i < headers.length; i++) {
        sheet.getRange(1, i + 1).setValue(headers[i]);
      }
    }
    
    // Adicionar uma linha por cada questão respondida
    if (data.registos && Array.isArray(data.registos)) {
      data.registos.forEach((registo) => {
        const novaLinha = sheet.getLastRow() + 1;
        
        const valores = [
          registo.id,
          registo.departamento,
          registo.investigacao,
          registo.foco,
          registo.evidencia,
          registo.om,
          registo.nc
        ];
        
        for (let i = 0; i < valores.length; i++) {
          sheet.getRange(novaLinha, i + 1).setValue(valores[i]);
        }
      });
    }
  } catch (erro) {
    Logger.log('Erro em adicionarRegistosAuditoria: ' + erro);
    throw erro;
  }
}

// Adicionar questão à folha REGISTOS
function adicionarQuestao(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('REGISTOS');
    
    // Headers esperados conforme especificado pelo utilizador
    const headers = [
      'ID', 'DEPARTAMENTO', 'Investigação (Questão de Auditoria)', 'Foco / Requisito'
    ];
    
    // Verificar se headers existem
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      for (let i = 0; i < headers.length; i++) {
        sheet.getRange(1, i + 1).setValue(headers[i]);
      }
    }
    
    // Gerar ID da questão
    const questionId = 'Q' + (lastRow); // lastRow é o número da última linha
    
    // Adicionar nova questão
    const novaLinha = sheet.getLastRow() + 1;
    
    const valores = [
      questionId,
      data.departamento,
      data.investigacao,
      data.foco
    ];
    
    for (let i = 0; i < valores.length; i++) {
      sheet.getRange(novaLinha, i + 1).setValue(valores[i]);
    }
  } catch (erro) {
    Logger.log('Erro em adicionarQuestao: ' + erro);
    throw erro;
  }
}

// Obter todas as questões da folha REGISTOS
function obterQuestoes(spreadsheet) {
  try {
    Logger.log('=== Iniciando obterQuestoes ===');
    
    let sheet;
    try {
      sheet = spreadsheet.getSheetByName('REGISTOS');
      Logger.log('✅ Sheet REGISTOS encontrado');
    } catch(e) {
      Logger.log('❌ Sheet REGISTOS NÃO encontrado. Sheets disponíveis:');
      const sheets = spreadsheet.getSheets();
      sheets.forEach((s, i) => Logger.log('  ' + (i+1) + '. ' + s.getName()));
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log('Dimensões do sheet: ' + lastRow + ' linhas x ' + lastCol + ' colunas');
    
    // Se sheet estiver vazio, criar headers
    if (lastRow === 0) {
      Logger.log('Sheet vazio, criando headers');
      const headers = ['ID', 'DEPARTAMENTO', 'Investigação (Questão de Auditoria)', 'Foco / Requisito'];
      for (let i = 0; i < headers.length; i++) {
        sheet.getRange(1, i + 1).setValue(headers[i]);
      }
      Logger.log('Headers criados');
      return [];
    }
    
    // Ler headers da linha 1
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    const headers = headerRange.getValues()[0];
    Logger.log('Headers:', JSON.stringify(headers));
    
    // Se há apenas header (linha 1)
    if (lastRow === 1) {
      Logger.log('Sheet tem apenas headers, sem dados');
      return [];
    }
    
    // Ler TODOS os dados do sheet (de linha 2 até a última)
    Logger.log('Lendo todas as linhas de dados...');
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    Logger.log('Total de linhas de dados:', values.length);
    Logger.log('Dados brutos (primeiras 3 linhas):', JSON.stringify(values.slice(0, 3)));
    
    const questoes = [];
    values.forEach((row, rowIndex) => {
      // Verificar se a linha tem qualquer conteúdo
      const temDados = row.some(cell => cell !== null && cell !== '');
      
      if (temDados) {
        const questao = {
          id: String(row[0] || 'Q' + (rowIndex + 2)),
          departamento: String(row[1] || ''),
          investigacao: String(row[2] || ''),
          foco: String(row[3] || '')
        };
        
        Logger.log('Linha ' + (rowIndex + 2) + ':', JSON.stringify(questao));
        questoes.push(questao);
      }
    });
    
    Logger.log('✅ Total de questões processadas:', questoes.length);
    Logger.log('Resultado final:', JSON.stringify(questoes));
    
    return questoes;
  } catch (erro) {
    Logger.log('❌ ERRO FATAL em obterQuestoes:', erro.toString());
    Logger.log('Stack:', erro.stack);
    return [];
  }
}

// Apagar auditoria planificada da folha INTRO
function apagarAuditoriaPlanificada(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('INTRO');
    const auditId = data.id;
    
    // Procurar a linha com o ID
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) { // Começar em 2 para pular o header
      const cellValue = sheet.getRange(i, 1).getValue(); // Coluna A tem o ID
      if (cellValue === auditId) {
        sheet.deleteRow(i);
        Logger.log('Linha apagada: ' + i);
        return;
      }
    }
    Logger.log('ID não encontrado: ' + auditId);
  } catch (erro) {
    Logger.log('Erro em apagarAuditoriaPlanificada: ' + erro);
    throw erro;
  }
}
