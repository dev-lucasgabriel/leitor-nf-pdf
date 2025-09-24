// leitor.js (Versão Definitiva: Mudado para Formato Tabela Consolidada / Flat File)

import express from 'express';
import multer from 'multer';
import { GoogleGenAI } from '@google/genai';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import random from 'random'; 
import 'dotenv/config'; 
import { fileURLToPath } from 'url'; 
import cors from 'cors'; 

// --- 1. Configurações de Ambiente e Variáveis Globais ---
const app = express();
const PORT = process.env.PORT || 3000; 

// Correção para obter __dirname em módulos ES
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');
const PUBLIC_DIR = path.join(__dirname, 'public'); 

// Garante que os diretórios existem
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

// Inicializa a API Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Configuração do Multer
const upload = multer({ dest: UPLOAD_DIR });

// Armazenamento em memória para dados de sessão
const sessionData = {};

// --- 2. Middlewares e Funções Essenciais ---
app.use(cors()); 

function fileToGenerativePart(filePath, mimeType) {
    return {
        inlineData: {
            data: Buffer.from(fs.readFileSync(filePath)).toString("base64"),
            mimeType
        },
    };
}

/**
 * Lógica de Backoff Exponencial para lidar com o erro 429.
 */
async function callApiWithRetry(apiCall, maxRetries = 5) {
    let delay = 2; 
    for (let attempt = 0; attempt < maxRetries; attempt++) {
        try {
            return await apiCall();
        } catch (error) {
            if (error.status === 429 || (error.message && error.message.includes('Resource has been exhausted'))) {
                if (attempt === maxRetries - 1) {
                    throw new Error('Limite de taxa excedido (429) após múltiplas tentativas. Tente novamente mais tarde.');
                }
                
                const jitter = random.uniform(0, 2)(); 
                const waitTime = (delay * (2 ** attempt)) + jitter;
                
                console.log(`[429] Tentando novamente em ${waitTime.toFixed(2)}s. Tentativa ${attempt + 1}/${maxRetries}`);
                await new Promise(resolve => setTimeout(resolve, waitTime * 1000));
            } else {
                throw error; 
            }
        }
    }
}

// --- 3. Função de Exportação FINAL (Implementa o Modelo Flat File para Tabela) ---

/**
 * Cria o arquivo Excel no formato Tabela Consolidada (Flat File)
 * Estrutura: | Arquivo Original | Chave 1 | Chave 2 | Chave 3 (Item 1) | Chave 4 (Item 2) | ...
 */
async function createExcelFile(allExtractedData, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Consolidados');

    if (allExtractedData.length === 0) return;

    const consolidatedRows = [];
    let allFileKeys = new Set(['arquivo_original']); // Chaves que serão as colunas

    // 1. APLANA E CONSOLIDA: Itera sobre todos os arquivos
    allExtractedData.forEach(fileData => {
        
        // Encontra a primeira chave cujo valor é um ARRAY DE OBJETOS (indica dados recorrentes - Ex: Itens)
        const recurrentKey = Object.keys(fileData).find(key => Array.isArray(fileData[key]) && typeof fileData[key][0] === 'object');
        
        // Coleta Metadados Estáticos (todos os campos que NÃO SÃO o array recorrente)
        const staticData = {};
        Object.keys(fileData).forEach(key => {
            if (!Array.isArray(fileData[key]) && key !== recurrentKey && key !== 'arquivo_original') {
                staticData[key] = fileData[key];
                allFileKeys.add(key);
            }
        });
        staticData['arquivo_original'] = fileData.arquivo_original;
        
        // Se houver dados recorrentes (Ex: Itens de NF, Registros de Ponto)
        if (recurrentKey && fileData[recurrentKey].length > 0) {
            // Cria MULTIPLAS LINHAS (uma por item/registro)
            fileData[recurrentKey].forEach(item => {
                const newRow = { ...staticData }; // Começa com os dados estáticos

                // Adiciona os dados do item/registro, renomeando as chaves para evitar conflito com dados estáticos, 
                // e garantindo que o nome da chave seja único (Ex: item_descricao, item_valor)
                Object.keys(item).forEach(key => {
                    const uniqueKey = `${recurrentKey.replace(/s$/, '')}_${key}`;
                    newRow[uniqueKey] = item[key];
                    allFileKeys.add(uniqueKey);
                });
                
                consolidatedRows.push(newRow);
            });
        
        // Se for um documento simples (apenas metadados)
        } else {
            // Adiciona uma ÚNICA LINHA
            consolidatedRows.push(staticData);
        }
    });

    // Filtra as chaves de rastreamento e garante 'arquivo_original' no início
    let finalHeaders = Array.from(allFileKeys);
    if (finalHeaders.includes('arquivo_original')) {
         finalHeaders.splice(finalHeaders.indexOf('arquivo_original'), 1);
    }
    finalHeaders.unshift('arquivo_original');
    
    // 2. Cria as colunas e injeta os dados
    worksheet.columns = finalHeaders.map(header => ({ 
        // Formatação do Header para melhor visualização (Ex: 'valor_total' -> 'Valor Total')
        header: header.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
        key: header, 
        width: 25 
    }));
    
    worksheet.addRows(consolidatedRows); // Adiciona todas as linhas da tabela

    // 3. Aplicação da Formatação Visual
    worksheet.getRow(1).eachCell(cell => {
        // Cor do Cabeçalho (Vermelho)
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } }; 
        cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 12 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { 
            row.eachCell(cell => {
                const header = worksheet.getRow(1).getCell(cell.col).value.toString().toLowerCase();
                // Formata colunas que contêm 'valor', 'total', 'icms', etc., como moeda
                if (header.includes('valor') || header.includes('total') || header.includes('icms') || header.includes('ipi') || header.includes('pis') || header.includes('cofins') || header.includes('fcp')) {
                    if (typeof cell.value === 'number') {
                        cell.numFmt = 'R$ #,##0.00'; 
                    }
                }
            });
        }
    });

    await workbook.xlsx.writeFile(outputPath);
}

// --- 4. Endpoint Principal de Upload e Processamento ---

app.post('/upload', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }

    const fileCleanupPromises = [];
    const allResultsForClient = [];
    const allResultsForExcel = [];
    
    // Prompt ajustado para pedir explicitamente ARRAY DE OBJETOS para dados recorrentes (necessário para Flat File)
    const prompt = `
        Você é um assistente especialista em extração de dados estruturados para análise em formato de tabela (Flat File). Sua tarefa é analisar o documento e extrair todas as informações.

        REGRAS CRÍTICAS para o JSON:
        1. Se houver dados recorrentes (como uma lista de itens, produtos, ou registros de ponto), crie uma chave (ex: "itens" ou "registros_diarios") cujo valor é um **ARRAY DE OBJETOS**. Isso é CRÍTICO para o formato de exportação.
        2. Dados estáticos (nome da empresa, data, ID, totais) devem ser campos simples no objeto principal.
        3. Formate datas como 'DD/MM/AAAA' e valores monetários/quantias como números (ex: 123.45).
        4. Inclua uma chave chamada 'resumo_executivo' com uma frase concisa descrevendo o documento, seguida por um resumo de todas as informações importantes extraídas.

        Retorne **APENAS** o objeto JSON completo.
    `;
    
    try {
        for (const file of req.files) {
            console.log(`[PROCESSANDO] ${file.originalname}`);
            
            const filePart = fileToGenerativePart(file.path, file.mimetype);
            
            const apiCall = () => ai.models.generateContent({
                model: 'gemini-1.5-flash', 
                contents: [filePart, { text: prompt }],
                config: {
                    responseMimeType: "application/json", 
                    temperature: 0.1
                }
            });

            try {
                const response = await callApiWithRetry(apiCall); 
                const dynamicData = JSON.parse(response.text);
                
                // 1. Obtém as chaves dinâmicas do JSON
                const keys = Object.keys(dynamicData);

                // 2. Mapeamento para o Front-end (Visualização Genérica):
                const clientResult = {
                    arquivo_original: file.originalname,
                    chave1: keys.length > 0 ? keys[0] : 'N/A',
                    valor1: keys.length > 0 ? dynamicData[keys[0]] : 'N/A',
                    chave2: keys.length > 1 ? keys[1] : 'N/A',
                    valor2: keys.length > 1 ? dynamicData[keys[1]] : 'N/A',
                    resumo: dynamicData.resumo_executivo || 'Resumo não fornecido.',
                };

                allResultsForClient.push(clientResult);
                
                // Adiciona o nome original do arquivo ao objeto dinâmico para rastreamento no Excel
                dynamicData.arquivo_original = file.originalname; 
                allResultsForExcel.push(dynamicData); 
                
                await new Promise(resolve => setTimeout(resolve, 5000)); 

            } catch (err) {
                console.error(`Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ erro: `Falha na API: ${err.message.substring(0, 100)}` });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        const sessionId = Date.now().toString();
        sessionData[sessionId] = allResultsForExcel;

        return res.json({ 
            results: allResultsForClient,
            sessionId: sessionId
        });

    } catch (error) {
        console.error('Erro fatal no processamento:', error);
        return res.status(500).send({ error: 'Erro interno do servidor.' });
    } finally {
        await Promise.all(fileCleanupPromises).catch(e => console.error("Erro ao limpar arquivos temporários:", e));
    }
});

// --- 5. Endpoint para Download do Excel ---

// Este endpoint agora executa a exportação no formato Flat File (Tabela Consolidada)
app.get('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const data = sessionData[sessionId];

    if (!data) {
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada.' });
    }

    const excelFileName = `extracao_consolidada_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        await createExcelFile(data, excelPath); // Chama a função que cria a tabela consolidada

        res.download(excelPath, excelFileName, async (err) => {
            if (err) {
                console.error("Erro ao enviar o Excel:", err);
            }
            await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel:", e));
            delete sessionData[sessionId];
        });
    } catch (error) {
        console.error('Erro ao gerar Excel:', error);
        res.status(500).send({ error: 'Falha ao gerar o arquivo Excel.' });
    }
});

// --- 6. Servir front-end e iniciar servidor ---

app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});