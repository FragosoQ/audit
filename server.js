const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// Configurar credenciais do Google
const sheets = google.sheets('v4');

// Função para autenticar com Google Sheets
async function getAuthClient() {
  // Usar variáveis de ambiente ou arquivo service account
  const auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_SERVICE_ACCOUNT || './service-account.json',
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  
  return auth.getClient();
}

// Endpoint para salvar auditoria planificada
app.post('/salvarIntro', async (req, res) => {
  try {
    const auth = await getAuthClient();
    const data = req.body;
    
    const spreadsheetId = process.env.SHEET_ID || '1bMoWDD2mHdzgFtJijQld-GHCcmm9rYpZ996Z2JIOMgk';
    
    // Dados a inserir
    const values = [[
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
      new Date().toISOString()
    ]];
    
    // Inserir dados na folha INTRO
    const response = await sheets.spreadsheets.values.append({
      auth,
      spreadsheetId,
      range: 'INTRO!A:M',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: values
      }
    });
    
    res.json({
      status: 'success',
      message: 'Dados gravados com sucesso',
      updates: response.data.updates
    });
    
  } catch (error) {
    console.error('Erro:', error);
    res.status(500).json({
      status: 'error',
      message: error.message
    });
  }
});

// Endpoint para salvar auditoria (registos)
app.post('/salvarAuditoria', async (req, res) => {
  try {
    const auth = await getAuthClient();
    const data = req.body;
    
    const spreadsheetId = process.env.SHEET_ID || '1bMoWDD2mHdzgFtJijQld-GHCcmm9rYpZ996Z2JIOMgk';
    
    // Preparar múltiplas linhas para cada registo
    const values = data.registos.map(registo => [
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
    ]);
    
    // Inserir dados na folha AUDITORIA
    const response = await sheets.spreadsheets.values.append({
      auth,
      spreadsheetId,
      range: 'AUDITORIA!A:L',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: values
      }
    });
    
    res.json({
      status: 'success',
      message: 'Auditoria guardada com sucesso',
      updates: response.data.updates
    });
    
  } catch (error) {
    console.error('Erro:', error);
    res.status(500).json({
      status: 'error',
      message: error.message
    });
  }
});

// Endpoint para adicionar questão
app.post('/adicionarQuestao', async (req, res) => {
  try {
    const auth = await getAuthClient();
    const data = req.body;
    
    const spreadsheetId = process.env.SHEET_ID || '1bMoWDD2mHdzgFtJijQld-GHCcmm9rYpZ996Z2JIOMgk';
    
    const values = [[
      data.departamento,
      data.investigacao,
      data.foco
    ]];
    
    // Inserir dados na folha REGISTOS
    const response = await sheets.spreadsheets.values.append({
      auth,
      spreadsheetId,
      range: 'REGISTOS!A:C',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: values
      }
    });
    
    res.json({
      status: 'success',
      message: 'Questão adicionada com sucesso',
      updates: response.data.updates
    });
    
  } catch (error) {
    console.error('Erro:', error);
    res.status(500).json({
      status: 'error',
      message: error.message
    });
  }
});

const PORT = process.env.PORT || 3002;
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
  console.log('Endpoints disponíveis:');
  console.log('  POST /salvarIntro');
  console.log('  POST /salvarAuditoria');
  console.log('  POST /adicionarQuestao');
});
