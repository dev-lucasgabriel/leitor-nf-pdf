// Leitor.js
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const pdf = require('pdf-parse');

const app = express();
const PORT = 3000;

// Pasta onde os uploads temporários serão salvos
const upload = multer({ dest: 'uploads/' });

// Função auxiliar para extrair campos do texto do PDF
function extrairValor(texto, regex) {
  const match = texto.match(regex);
  return match ? match[1].trim().replace(/\s+/g, ' ') : 'Nao encontrado';
}

// Função principal para processar um único PDF
async function processarNotaDeTexto(caminhoDoPdf) {
  try {
    const dataBuffer = fs.readFileSync(caminhoDoPdf);
    const data = await pdf(dataBuffer);
    const texto = data.text;

    const dadosExtraidos = {
      arquivo: path.basename(caminhoDoPdf),
      numeroNota: extrairValor(texto, /Nº\.\s*([\d.]+)/),
      serie: extrairValor(texto, /Série\s*(\d+)/),
      dataEmissao: extrairValor(texto, /DATA DA EMISSÃO\s*([\d\/]+)/),
      dataEntradaSaida: extrairValor(texto, /DATA DA SAÍDA\/ENTRADA\s*([\d\/]+)/),
      naturezaOperacao: extrairValor(texto, /NATUREZA DA OPERAÇÃO\s*([\w\s./-]+?)\s*PROTOCOLO/),
      valorTotalNota: extrairValor(texto, /V\. TOTAL DA NOTA\s*([\d.,]+)/),
      valorIcms: extrairValor(texto, /VALOR DO ICMS\s*([\d.,]+)/),
      valorIcmsSt: extrairValor(texto, /BASE DE CÁLC\. ICMS S\.T\.\s*([\d.,]+)/),
      valorTotalProdutos: extrairValor(texto, /V\. TOTAL PRODUTOS\s*([\d.,]+)/),
      produtos: []
    };

    const inicioTabela = 'DADOS DOS PRODUTOS / SERVIÇOS';
    const fimTabela = 'DADOS ADICIONAIS';
    const textoTabela = texto.substring(texto.indexOf(inicioTabela), texto.indexOf(fimTabela));
    const regexProdutos = /^(\d{4,})(.+?)(\d{8})/gms;

    let match;
    while ((match = regexProdutos.exec(textoTabela)) !== null) {
      dadosExtraidos.produtos.push({
        codigo: match[1].trim(),
        descricao: match[2].trim().replace(/\s+/g, ' '),
        ncm: match[3].trim()
      });
    }

    return dadosExtraidos;
  } catch (error) {
    console.error('Erro ao processar PDF:', error);
    return null;
  }
}

// Rota de upload para múltiplos PDFs
app.post('/upload', upload.array('pdfs'), async (req, res) => {
  const resultados = [];

  for (const file of req.files) {
    const resultado = await processarNotaDeTexto(file.path);
    if (resultado) resultados.push(resultado);
    fs.unlinkSync(file.path); // remove arquivo temporário
  }

  res.json(resultados);
});

// Servir os arquivos estáticos do front-end
app.use(express.static('public'));

// Iniciar servidor
app.listen(PORT, () => {
  console.log(`✅ Servidor rodando em http://localhost:${PORT}`);
});
