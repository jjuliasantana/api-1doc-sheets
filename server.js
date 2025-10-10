const express = require('express');
const fs = require('fs');
const { google } = require('googleapis');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(express.urlencoded({ extended: true }));

const idPlanilha = '1BK9bb5_rHuTXfemDRm0pFFcucNDDT06ED7uIvT7Yezg';

const auth = new google.auth.GoogleAuth({
  keyFile: 'credentials.json',
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

async function escreverNaPlanilha(dadosDoWebhook) {
 try {
  console.log('Iniciando escrita...');

  const sheets = google.sheets({ version: 'v4', auth });
  
  const emissao = dadosDoWebhook.emissao;

    const novaLinha = [
      emissao.matricula_1h0pel1h || '',
      '',
      '',
      '',
      emissao.data_da_fa_1hi4ib1h || '',
      '',
      '',
      '',
      '',
      '',
      '',
      ''
    ];

  const request = {
    spreadsheetId: idPlanilha,
    range: 'Página1!A:L',
    valueInputOption: 'USER_ENTERED',
    resource: {
      values: [novaLinha],
    },
  }
  await sheets.spreadsheets.values.append(request);
  console.log('Dados escritos na planilha com sucesso!');
} catch (error) {
  console.error('Erro ao escrever na planilha:', error.message);
  throw error;
}
}

app.post('/webhook', async (req, res) => {
    try {
    console.log('--- Webhook Recebido! ---');
    
    const corpoDoFormulario = req.body;

    const stringBase64 = corpoDoFormulario.data;

    if (!stringBase64) {
      console.error('ERRO: O campo "data" não foi encontrado no corpo da requisição.');
      return res.status(400).send('Payload inválido.');
    }

    const dadosJsonString = Buffer.from(stringBase64, 'base64').toString('utf8');

    const dadosFinais = JSON.parse(dadosJsonString);

    const idSetor = 635;
    
    if (dadosFinais.emissao.destino_id_setor == idSetor) {
      console.log('O setor de destino é o correto. Prosseguindo com a escrita na planilha...');
      await escreverNaPlanilha(dadosFinais);
    } else {
      console.log(`O setor de destino (${dadosFinais.emissao.destino_id_setor}) não corresponde ao setor esperado (${idSetor}). Ignorando a escrita na planilha.`);
    }

    console.log('DADOS DECODIFICADOS:');
    console.log(dadosFinais);

    const logEntry = `------------------------------\nData: ${new Date().toISOString()}\nDados Decodificados: ${JSON.stringify(dadosFinais, null, 2)}\n\n`;
    fs.appendFile('webhook_logs.txt', logEntry, () => {});
    
    
    console.log('---------------------------');
    res.status(200).send('Dados recebidos com sucesso!');

  } catch (error) {
    console.error('Ocorreu um erro ao processar o webhook:', error.message);
    res.status(500).send('Erro interno no servidor.');
  }
});

app.listen(PORT, () => {
  console.log(`Servidor rodando e ouvindo na porta ${PORT}`);
});
