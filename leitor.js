// leitor.js (Versão Final: Assíncrono, Estável para Render, e Exportação Flat File)

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
import { setTimeout as delay } from 'timers/promises'; // Importa a função delay para uso em async/await

// --- 1. Configurações de Ambiente e Variáveis Globais ---
const app = express();
const PORT = process.env.PORT || 3000; 

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');
const PUBLIC_DIR = path.join(__dirname, 'public'); 

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const upload = multer({ dest: UPLOAD_DIR });

// Estrutura de sessão para rastrear o progresso assíncrono
const sessionData = {}; 

// --- 2. Funções Essenciais de Utilidade e Segurança ---
app.use(cors()); 
app.use(express.json()); // NECESSÁRIO para ler o body POST no download-excel

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
    let delayTime = 2; 
    for (let attempt = 0; attempt < maxRetries; attempt++) {
        try {
            return await apiCall();
        } catch (error) {
            if (error.status === 429 || (error.message && error.message.includes('Resource has been exhausted'))) {
                if (attempt === maxRetries - 1) {
                    throw new Error('Limite de taxa excedido (429) após múltiplas tentativas. Tente novamente mais tarde.');
                }
                
                const jitter = random.uniform(0, 2)(); 
                const waitTime = (delayTime * (2 ** attempt)) + jitter;
                
                console.log(`[429] Tentando novamente em ${waitTime.toFixed(2)}s. Tentativa ${attempt + 1}/${maxRetries}`);
                await delay(waitTime * 1000);
            } else {
                throw error; 
            }
        }
    }
}

// --- 3. Função de Exportação (Tabela Consolidada / Flat File) ---

/**
 * Cria o arquivo Excel no formato Tabela Consolidada (Flat File),
 * aplicando a curadoria (selectedKeys) no momento da exportação.
 */
async function createExcelFile(allExtractedData, selectedKeys, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Consolidados');

    if (allExtractedData.length === 0 || selectedKeys.length === 0) return;

    const consolidatedRows = [];
    let allFileKeys = new Set(['arquivo_original']);

    allExtractedData.forEach(fileData => {
        const recurrentKey = Object.keys(fileData).find(key => Array.isArray(fileData[key]) && typeof fileData[key][0] === 'object');
        
        const staticData = {};
        Object.keys(fileData).forEach(key => {
            if (!Array.isArray(fileData[key]) && key !== recurrentKey && key !== 'arquivo_original') {
                staticData[key] = fileData[key];
                allFileKeys.add(key);
            }
        });
        staticData['arquivo_original'] = fileData.arquivo_original;
        
        if (recurrentKey && fileData[recurrentKey].length > 0) {
            fileData[recurrentKey].forEach(item => {
                const newRow = { ...staticData };
                Object.keys(item).forEach(key => {
                    const uniqueKey = `${recurrentKey.replace(/s$/, '')}_${key}`; 
                    newRow[uniqueKey] = item[key];
                    allFileKeys.add(uniqueKey);
                });
                consolidatedRows.push(newRow);
            });
        } else {
            consolidatedRows.push(staticData);
        }
    });
    
    // --- Aplicação da Curadoria ---
    let finalHeaders = Array.from(allFileKeys).filter(key => selectedKeys.includes(key));
    
    if (finalHeaders.includes('arquivo_original')) {
         finalHeaders.splice(finalHeaders.indexOf('arquivo_original'), 1);
    }
    finalHeaders.unshift('arquivo_original');
    
    // Configura as Colunas e Adiciona as Linhas
    worksheet.columns = finalHeaders.map(header => ({ 
        header: header.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
        key: header, 
        width: 25 
    }));
    worksheet.addRows(consolidatedRows); 

    // Aplicação da Formatação Visual
    worksheet.getRow(1).eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } }; 
        cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 12 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { 
            row.eachCell(cell => {
                const header = worksheet.getRow(1).getCell(cell.col).value.toString().toLowerCase();
                // Formata Moeda (Flat File)
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

// --- 4. Função de Processamento Assíncrono (Background) ---

async function processFilesInBackground(sessionId, files) {
    const session = sessionData[sessionId];
    const allResultsForClient = [];
    const allResultsForExcel = [];
    const fileCleanupPromises = [];

    const prompt = `
        Você é um assistente especialista em extração de dados estruturados para análise em formato de tabela (Flat File). Sua tarefa é analisar o documento e extrair todas as informações.

        REGRAS CRÍTICAS para o JSON:
        1. Se houver dados recorrentes (como uma lista de itens, produtos, ou registros de ponto), crie uma chave (ex: "itens" ou "registros_diarios") cujo valor é um **ARRAY DE OBJETOS**.
        2. Dados estáticos (nome da empresa, data, ID, totais) devem ser campos simples no objeto principal.
        3. Formate datas como 'DD/MM/AAAA' e valores monetários/quantias como números (ex: 123.45).
        4. Inclua uma chave chamada 'resumo_executivo' com uma frase concisa descrevendo o documento.

        Retorne **APENAS** o objeto JSON completo.
    `;

    try {
        for (const file of files) {
            console.log(`[BACKGROUND] Processando: ${file.originalname}`);
            
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
                
                // Mapeamento para o Front-end (Visualização na tela de status)
                const keys = Object.keys(dynamicData);
                allResultsForClient.push({
                    arquivo_original: file.originalname,
                    chave1: keys.length > 0 ? keys[0] : 'N/A',
                    valor1: keys.length > 0 ? dynamicData[keys[0]] : 'N/A',
                    chave2: keys.length > 1 ? keys[1] : 'N/A',
                    valor2: keys.length > 1 ? dynamicData[keys[1]] : 'N/A',
                    resumo: dynamicData.resumo_executivo || 'Resumo não fornecido.',
                });
                
                dynamicData.arquivo_original = file.originalname; 
                allResultsForExcel.push(dynamicData); 

                await delay(3000); // Pequeno atraso para evitar esgotamento de taxa (429)

            } catch (err) {
                console.error(`[BACKGROUND] Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ erro: `Falha na IA: ${err.message.substring(0, 100)}`, arquivo_original: file.originalname });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        session.data = allResultsForExcel;
        session.clientResults = allResultsForClient;
        session.status = 'CONCLUIDO';
        console.log(`[BACKGROUND] Sessão ${sessionId} CONCLUÍDA.`);

    } catch (error) {
        console.error(`[BACKGROUND] ERRO FATAL na Sessão ${sessionId}:`, error);
        session.status = 'ERRO';
        session.error = error.message;
    } finally {
        await Promise.all(fileCleanupPromises).catch(e => console.error("Erro ao limpar uploads:", e));
    }
}

// --- 5. Endpoint Principal de Upload (Inicia o processo e responde rápido) ---

app.post('/upload', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }

    const sessionId = Date.now().toString();
    
    sessionData[sessionId] = { 
        status: 'PROCESSANDO', 
        data: [], 
        clientResults: [], 
        error: null 
    };

    // Inicia o processamento pesado EM SEGUNDO PLANO
    processFilesInBackground(sessionId, req.files); 

    // Resposta RÁPIDA (202 Accepted) para o Render
    return res.status(202).json({ 
        message: 'Upload aceito. O processamento da IA iniciou em segundo plano. Consulte o status.',
        sessionId: sessionId
    });
});

// --- 6. Endpoint de Status (Para o Frontend Consultar) ---

app.get('/status/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session) {
        return res.status(404).json({ status: 'EXPIRADA', message: 'Sessão não encontrada.' });
    }

    // Retorna o status, resultados parciais e os dados completos (para curadoria no front)
    return res.json({
        status: session.status,
        results: session.clientResults,
        data: session.data, // Importante para permitir a curadoria no frontend
        error: session.error
    });
});


// --- 7. Endpoint para Download do Excel (CURADORIA E DOWNLOAD) ---

// Este endpoint aceita POST, lê as chaves de curadoria (selectedKeys), gera o Excel, e envia o arquivo.
app.post('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const { selectedKeys } = req.body; // Recebe as chaves do frontend

    const session = sessionData[sessionId];

    if (!session || session.status !== 'CONCLUIDO') {
        return res.status(409).send({ error: 'Processamento não concluído. Tente novamente mais tarde.' });
    }

    if (!selectedKeys || selectedKeys.length === 0) {
        return res.status(400).send({ error: 'Nenhuma chave selecionada para exportação.' });
    }

    const excelFileName = `extracao_curada_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        // Usa a função createExcelFile com os dados completos E as chaves selecionadas
        await createExcelFile(session.data, selectedKeys, excelPath); 

        res.download(excelPath, excelFileName, async (err) => {
            if (err) {
                console.error("Erro ao enviar o Excel:", err);
            }
            await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel:", e));
            delete sessionData[sessionId]; // Limpa a memória após o download
        });
    } catch (error) {
        console.error('Erro ao gerar Excel Curado:', error);
        res.status(500).send({ error: 'Falha ao gerar o arquivo Excel curado.' });
    }
});

// --- 8. Servir front-end e iniciar servidor ---

app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

const server = app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});

// Aumenta o timeout do servidor para que as conexões de download não sejam cortadas
server.setTimeout(600000);