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
    } else if (operacao === 'atualizarStatusIntro') {
      atualizarStatusAuditoriaPlanificada(spreadsheet, data);
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
    } else if (operacao === 'obterIntro') {
      Logger.log('Iniciando obtenção de planeamentos INTRO...');
      const planeamentos = obterPlaneamentosIntro(spreadsheet);
      Logger.log('Planeamentos INTRO obtidos:', JSON.stringify(planeamentos));
      return ContentService.createTextOutput(JSON.stringify(planeamentos))
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

    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    const values = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 12)).getValues();
    const registos = [];

    const col = (aliases, fallbackIndex) => {
      const idx = headerRow.findIndex(h => aliases.includes(h));
      return idx >= 0 ? idx : fallbackIndex;
    };

    const idxId = col(['id'], 0);
    const idxDept = col(['departamento', 'departamento '], 1);
    const idxInvestigacao = col(['investigação (questão de auditoria)', 'investigacao (questão de auditoria)', 'investigação', 'investigacao'], 2);
    const idxFoco = col(['foco / requisito', 'foco/requisito', 'foco'], 3);
    const idxEvidencia = col(['constatação / evidência', 'constatacao / evidencia', 'evidencia', 'constatação/evidência'], 4);
    const idxOm = col(['om'], 5);
    const idxNc = col(['nc'], 6);
    const idxAuditorCoord = col(['auditor coordenador'], 7);
    const idxAuditor = col(['auditor'], 8);
    const idxAuditado = col(['auditado'], 9);
    const idxRegistadoPor = col(['registado por'], 10);
    const idxData = col(['data'], 11);

    const valor = (row, idx) => (idx >= 0 ? String(row[idx] || '') : '');

    values.forEach((row) => {
      const temDados = row.some(cell => cell !== null && cell !== '');
      if (!temDados) return;

      registos.push({
        id: valor(row, idxId),
        departamento: valor(row, idxDept),
        auditorCoord: valor(row, idxAuditorCoord),
        auditor: valor(row, idxAuditor),
        auditado: valor(row, idxAuditado),
        registadoPor: valor(row, idxRegistadoPor),
        data: valor(row, idxData),
        investigacao: valor(row, idxInvestigacao),
        foco: valor(row, idxFoco),
        evidencia: valor(row, idxEvidencia),
        om: valor(row, idxOm),
        nc: valor(row, idxNc)
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

function obterPlaneamentosIntro(spreadsheet) {
  try {
    Logger.log('=== Iniciando obterPlaneamentosIntro ===');

    const sheet = spreadsheet.getSheetByName('INTRO');
    if (!sheet) {
      Logger.log('❌ Sheet INTRO não encontrado');
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log('Dimensões INTRO: ' + lastRow + ' linhas x ' + lastCol + ' colunas');

    if (lastRow <= 1) {
      Logger.log('Sheet INTRO vazio ou só com headers');
      return [];
    }

    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const col = (aliases, fallbackIndex) => {
      const idx = headerRow.findIndex(h => aliases.includes(h));
      return idx >= 0 ? idx : fallbackIndex;
    };

    const idxId = col(['id'], 0);
    const idxDepartamento = col(['departamento'], 1);
    const idxAuditorCoord = col(['auditor coordenador'], 2);
    const idxAuditor = col(['auditor'], 3);
    const idxAuditado = col(['auditado'], 4);
    const idxDocumentosRef = col(['documentos de referência da empresa', 'documentos de referencia da empresa'], 5);
    const idxRegistadoPor = col(['registado por'], 6);
    const idxColaboradores = col(['colaboradores contactados'], 7);
    const idxRequisitos = col(['requisitos'], 8);
    const idxDataAuditoria = col(['data de auditoria'], 9);
    const idxHoraPrevista = col(['hora prevista'], 10);
    const idxHoraAuditoria = col(['hora da auditoria'], 11);
    const idxRealizada = col(['realizada'], 12);
    const idxDataRealizacao = col(['data realização', 'data realizacao'], 13);

    const valor = (row, idx) => (idx >= 0 ? String(row[idx] || '').trim() : '');

    // Formatar campos de data: o Sheets pode entregar objetos Date em vez de texto
    const formatarDataYMD = (raw) => {
      if (!raw) return '';
      if (raw instanceof Date && !isNaN(raw.getTime())) {
        const y = raw.getFullYear();
        const m = String(raw.getMonth() + 1).padStart(2, '0');
        const d = String(raw.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
      }
      // Já é texto — garantir só YYYY-MM-DD (sem hora)
      return String(raw).trim().substring(0, 10);
    };

    const planeamentos = [];

    values.forEach((row) => {
      const id = valor(row, idxId);
      if (!id) return;

      const realizadaTexto = valor(row, idxRealizada).toUpperCase();
      const rawData = idxDataAuditoria >= 0 ? row[idxDataAuditoria] : '';

      planeamentos.push({
        id,
        departamento: valor(row, idxDepartamento),
        auditorCoord: valor(row, idxAuditorCoord),
        auditor: valor(row, idxAuditor),
        auditado: valor(row, idxAuditado),
        documentosRef: valor(row, idxDocumentosRef),
        registadoPor: valor(row, idxRegistadoPor),
        colaboradores: valor(row, idxColaboradores),
        requisitos: valor(row, idxRequisitos),
        dataAuditoria: formatarDataYMD(rawData),
        horaPrevista: valor(row, idxHoraPrevista),
        horaAuditoria: valor(row, idxHoraAuditoria),
        realizada: realizadaTexto === 'SIM' || realizadaTexto === 'TRUE' || realizadaTexto === '1',
        dataRealizacao: valor(row, idxDataRealizacao)
      });
    });

    Logger.log('✅ Total de planeamentos INTRO: ' + planeamentos.length);
    return planeamentos;
  } catch (erro) {
    Logger.log('❌ ERRO em obterPlaneamentosIntro: ' + erro.toString());
    return [];
  }
}

// Adicionar auditoria planeada à folha INTRO
function adicionarAuditoriaPlanificada(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('INTRO');
    
    // Headers esperados
    const headers = [
      'ID', 'Departamento', 'Auditor Coordenador', 'Auditor', 'Auditado',
      'Documentos de Referência da Empresa', 'Registado Por', 'Colaboradores Contactados',
      'Requisitos', 'Data de Auditoria', 'Hora Prevista', 'Hora da Auditoria', 'Realizada', 'Data Realização'
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
      data.horaAuditoria,
      data.realizada === true ? 'SIM' : 'NÃO',
      data.dataRealizacao || ''
    ];
    
    for (let i = 0; i < valores.length; i++) {
      sheet.getRange(novaLinha, i + 1).setValue(valores[i]);
    }
  } catch (erro) {
    Logger.log('Erro em adicionarAuditoriaPlanificada: ' + erro);
    throw erro;
  }
}

function atualizarStatusAuditoriaPlanificada(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('INTRO');
    const id = data.id;
    if (!sheet || !id) return;

    if (!sheet.getRange(1, 13).getValue()) sheet.getRange(1, 13).setValue('Realizada');
    if (!sheet.getRange(1, 14).getValue()) sheet.getRange(1, 14).setValue('Data Realização');

    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      if (String(cellValue) === String(id)) {
        sheet.getRange(i, 13).setValue(data.realizada === true ? 'SIM' : 'NÃO');
        sheet.getRange(i, 14).setValue(data.realizada === true ? (data.dataRealizacao || new Date().toISOString()) : '');
        Logger.log('Status atualizado para ID: ' + id);
        return;
      }
    }

    Logger.log('ID não encontrado para atualizar status: ' + id);
  } catch (erro) {
    Logger.log('Erro em atualizarStatusAuditoriaPlanificada: ' + erro);
    throw erro;
  }
}

// Adicionar registos de auditoria à folha AUDITORIA
function adicionarRegistosAuditoria(spreadsheet, data) {
  try {
    const sheet = spreadsheet.getSheetByName('AUDITORIA');
    
    // Headers esperados conforme especificado
    const headers = [
      'ID', 'DEPARTAMENTO', 'Investigação (Questão de Auditoria)', 'Foco / Requisito', 'CONSTATAÇÃO / EVIDÊNCIA', 'OM', 'NC',
      'Auditor Coordenador', 'Auditor', 'Auditado', 'Registado Por', 'Data'
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
          registo.nc,
          registo.auditorCoord,
          registo.auditor,
          registo.auditado,
          registo.registadoPor,
          registo.data
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
      'ID', 'DEPARTAMENTO', 'Investigação (Questão de Auditoria)', 'Foco / Requisito', 'Categoria'
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
      data.foco,
      data.categoria || ''
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
      const headers = ['ID', 'DEPARTAMENTO', 'Investigação (Questão de Auditoria)', 'Foco / Requisito', 'Categoria'];
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

    const headerNorm = headers.map(h => String(h || '').trim().toLowerCase());
    const col = (aliases, fallbackIndex) => {
      const idx = headerNorm.findIndex(h => aliases.includes(h));
      return idx >= 0 ? idx : fallbackIndex;
    };

    const idxId = col(['id'], 0);
    const idxDepartamento = col(['departamento', 'departamento '], 1);
    const idxInvestigacao = col(['investigação (questão de auditoria)', 'investigacao (questão de auditoria)', 'investigação', 'investigacao'], 2);
    const idxFoco = col(['foco / requisito', 'foco/requisito', 'foco'], 3);
    const idxCategoria = col(['categoria'], 4);
    
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
          id: String(row[idxId] || 'Q' + (rowIndex + 2)),
          departamento: String(row[idxDepartamento] || ''),
          investigacao: String(row[idxInvestigacao] || ''),
          foco: String(row[idxFoco] || ''),
          categoria: String(row[idxCategoria] || '').trim()
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
