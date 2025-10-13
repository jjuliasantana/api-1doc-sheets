const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
const { sheets } = require('googleapis/build/src/apis/sheets');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.urlencoded({ extended: true }));

const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const SPREADSHEET_ID = '1q3stC_Z5kpeZ4oek-hzuZQonANJ7uaOau2QInCVo3I4';

const auth = new google.auth.JWT({
  keyFile: CREDENTIALS_PATH,
  scopes: 'https://www.googleapis.com/auth/spreadsheets',
});

function formatarData(data) {
  const partes = data.split('-');
  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

async function escreverNaPlanilha(dadosDoWebhook) {
  try {
    console.log('Iniciando a escrita na planilha...');
    await auth.authorize();
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const emissao = dadosDoWebhook.emissao;
    
    const novaLinha = [
      emissao.matricula_1h0pel1h || '',
      '1',
      '9',
      emissao.data_da_fa_1hi4ib1h || '',
      emissao.data_da_fa_1hi4ib1h || '',
      'd',
      '0',
      'importado',
    ];

    const request = {
      spreadsheetId: SPREADSHEET_ID,
      range: 'Página1!A:H',
      valueInputOption: 'RAW',
      resource: {
        values: [novaLinha],
      },
    };

    await sheets.spreadsheets.values.append(request);
    console.log('-> Nova linha adicionada à planilha com sucesso!');

  } catch (error) {
    console.error('ERRO ao escrever na planilha:', error.message);
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
      console.log(`O setor de destino (${dadosFinais.emissao.destino_id_setor}) não corresponde ao setor esperado (${idSetor}) e não será incluído.`);
    }

    console.log('DADOS DECODIFICADOS:');
    console.log(dadosFinais);

    const logEntry = `------------------------------\nData: ${new Date().toISOString()}\nDados Decodificados: ${JSON.stringify(dadosFinais, null, 2)}\n\n`;
    // fs.appendFile('webhook_logs.txt', logEntry, () => {});
    
    
    console.log('---------------------------');
    res.status(200).send('Dados recebidos com sucesso!');

  } catch (error) {
    console.error('Ocorreu um erro ao processar o webhook:', error.message);
    res.status(500).send('Erro interno no servidor.');
  }
});

app.get('/gerar-relatorio', async (req, res) => {
  try {
    console.log('Iniciando a geração do arquivo...');
    
    const respostaDoSheets = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Página1!A:H'
    })
    const linhasDaPlanilha = respostaDoSheets.data.values;

    if (linhasDaPlanilha && linhasDaPlanilha.length > 1) {
      console.log(`Encontradas ${linhasDaPlanilha.length} linhas para processar.`);
     
      let conteudoDoRelatorio = 'Matricula;Sequencia;Ocorrencia;dtinicio;dtfim;Situacao;branco;observacao\n';

      for (let i = 1; i < linhasDaPlanilha.length; i++) {
        const linha = linhasDaPlanilha[i];
        const matricula = linha[0] || '';
        const sequencia = linha[1] || '1';
        const ocorrencia = linha[2] || '9';
        const dtinicio = formatarData(linha[3]) || '';
        const dtfim = formatarData(linha[4]) || '';
        const situacao = linha[5] || 'd';
        const branco = linha[6] || '0';
        const observacao = linha[7] || 'importado';

      const novaLinha = `${matricula};${sequencia};${ocorrencia};${dtinicio};${dtfim};${situacao};${branco};${observacao}\n`;
      conteudoDoRelatorio += novaLinha;
    }
      
    const dataAtual = new Date().toISOString().slice(0, 10);
    const nomeDoArquivo = `relatorio_${dataAtual}.txt`;

    await fsPromises.writeFile(nomeDoArquivo, conteudoDoRelatorio);
    console.log('Relatório gerado com sucesso:', nomeDoArquivo);
    
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Página1!A2:H'
    })
    res.send('Relatório gerado e planilha limpa com sucesso.');

    } else {
      console.log('Nenhuma linha de dados encontrada na planilha.');
      res.send('A planilha está vazia, nenhum relatório foi gerado.');
    }

  } catch (error) {
    console.error('Erro ao ler a planilha:', error.message);
    res.status(500).send('Erro ao ler a planilha.');
  }

});


app.listen(PORT, () => {
  console.log(`Servidor rodando e ouvindo na porta ${PORT}`);
  console.log(`Webhook: http://localhost:${PORT}/webhook`);
  console.log(`Gerar Relatório: http://localhost:${PORT}/gerar-relatorio`);
});