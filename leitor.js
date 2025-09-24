// leitor.js (Versão Definitiva: Curadoria de Dados com Formatação de Planilha Standard)

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

// Middleware para processar JSON (necessário para receber as chaves selecionadas)
app.use(express.json()); 

// Corrige o problema do __dirname em módulos ES
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');
const PUBLIC_DIR = path.join(__dirname, 'public'); 

// Cria diretórios e inicializa API
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const upload = multer({ dest: UPLOAD_DIR });
const sessionData = {}; // Armazena dados brutos da IA por sessão (Step 1)

// --- 2. Funções Essenciais de Utilidade e Segurança ---

function fileToGenerativePart(filePath, mimeType) {
    return {
        inlineData: {
            data: Buffer.from(fs.readFileSync(filePath)).toString("base64"),
            mimeType
        },
    };
}

/**
 * Lógica de Backoff Exponencial para lidar com o erro 429 (Proteção).
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

// --- 3. Função de Exportação FINAL (Cria o Excel no Formato Standard) ---

async function createFilteredExcel(allExtractedData, selectedKeys, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Curados');

    if (allExtractedData.length === 0 || selectedKeys.length === 0) return;

    // 1. O Cabeçalho é EXATAMENTE a lista que o usuário selecionou
    const finalHeaders = selectedKeys;

    worksheet.columns = finalHeaders.map(header => ({ 
        // Formata para cabeçalhos legíveis (Nome Cliente -> Nome Cliente)
        header: header.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
        key: header, 
        width: 30 
    }));
    
    // 2. Mapeia os dados: CADA OBJETO (ARQUIVO) VIRA UMA NOVA LINHA HORIZONTAL
    const filteredRows = allExtractedData.map(data => {
        const row = {};
        finalHeaders.forEach(key => {
            row[key] = data[key] || ''; // Valor alinhado à coluna
        });
        return row;
    });

    worksheet.addRows(filteredRows);
    
    // 3. Aplicação da Formatação Visual (Cores e Moeda)
    worksheet.getRow(1).eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } }; 
        cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 12 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { 
            row.eachCell(cell => {
                const header = worksheet.getRow(1).getCell(cell.col).value.toString().toLowerCase();
                if (header.includes('valor') || header.includes('total')) {
                    cell.numFmt = 'R$ #,##0.00'; 
                }
            });
        }
    });

    await workbook.xlsx.writeFile(outputPath);
}

// --- 4. Endpoint de ANÁLISE (Step 1: Upload e Extração Bruta) ---

app.post('/api/analyze', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }
    
    const sessionId = Date.now().toString();
    const allExtractedData = []; 
    let allUniqueKeys = new Set();
    const fileCleanupPromises = [];

    // Prompt Agnostico
    const prompt = `
        Você é um assistente especialista em extração de dados estruturados. Sua tarefa é analisar o documento anexado (PDF ou IMAGEM) e extrair **TODAS** as informações relevantes. Crie um objeto JSON plano onde cada chave é o nome da informação extraída. Inclua uma chave chamada 'resumo_executivo' com um resumo do documento. Retorne APENAS o JSON.
    `;

    for (const file of req.files) {
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
            
            allExtractedData.push(dynamicData);
            Object.keys(dynamicData).forEach(key => allUniqueKeys.add(key));
        } catch (err) {
            console.error(`Erro na análise de ${file.originalname}:`, err);
            // Adiciona um erro ao dado para informar o usuário
            allExtractedData.push({ 
                erro_processamento: `Falha na IA. ${err.message.substring(0, 50)}...`, 
                arquivo_original: file.originalname 
            });
            allUniqueKeys.add('erro_processamento');
        } finally {
            fileCleanupPromises.push(fs.promises.unlink(file.path));
        }
    }

    // Armazena todos os dados brutos e as chaves únicas para o próximo passo
    sessionData[sessionId] = { 
        data: allExtractedData, 
        keys: Array.from(allUniqueKeys)
    };

    await Promise.all(fileCleanupPromises);

    // Retorna a lista de chaves (opções) para o frontend
    return res.json({ 
        sessionId: sessionId, 
        availableKeys: Array.from(allUniqueKeys)
    });
});

// --- 5. Endpoint de EXPORTAÇÃO (Step 2: Recebe Chaves Selecionadas) ---

app.post('/api/export-excel', async (req, res) => {
    const { sessionId, selectedKeys } = req.body;

    if (!sessionId || !selectedKeys || selectedKeys.length === 0) {
        return res.status(400).send({ error: 'Sessão ou campos selecionados ausentes.' });
    }

    const session = sessionData[sessionId];
    if (!session) {
        return res.status(404).send({ error: 'Sessão expirada ou não encontrada.' });
    }

    const excelFileName = `extracao_curada_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        await createFilteredExcel(session.data, selectedKeys, excelPath);

        // Envia o arquivo para download e limpa a sessão
        res.download(excelPath, excelFileName, async (err) => {
            if (err) console.error("Erro ao enviar o Excel:", err);
            await fs.promises.unlink(excelPath).catch(() => {});
            delete sessionData[sessionId]; // Limpa a memória
        });
    } catch (error) {
        console.error('Erro ao gerar Excel Curado:', error);
        res.status(500).send({ error: 'Falha ao gerar o arquivo Excel curado.' });
    }
});

// --- 6. Servir front-end e iniciar servidor ---

app.use(cors()); 
app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});