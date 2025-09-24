// leitor.js (Versão Definitiva: Formato Flat File - Múltiplas Linhas por Documento Recorrente)

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

app.use(express.json()); 

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');
const PUBLIC_DIR = path.join(__dirname, 'public'); 

// Cria diretórios
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

// --- 3. Função de Exportação FINAL (Implementa o Modelo Flat File) ---

async function createFilteredExcel(allExtractedData, selectedKeys, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Consolidados');

    if (allExtractedData.length === 0 || selectedKeys.length === 0) return;

    const consolidatedRows = [];
    let allFileKeys = new Set(['arquivo_original']); // Começa com a chave de rastreamento

    // 1. APLANA E CONSOLIDA: Itera sobre todos os arquivos
    allExtractedData.forEach(fileData => {
        // Encontra a primeira chave cujo valor é um ARRAY DE OBJETOS (indica dados recorrentes)
        const recurrentKey = Object.keys(fileData).find(key => Array.isArray(fileData[key]) && typeof fileData[key][0] === 'object');
        
        // Coleta Metadados Estáticos (todos os campos que NÃO SÃO o array recorrente)
        const staticData = {};
        Object.keys(fileData).forEach(key => {
            if (!Array.isArray(fileData[key]) && key !== recurrentKey) {
                staticData[key] = fileData[key];
                allFileKeys.add(key);
            }
        });
        staticData['arquivo_original'] = fileData.arquivo_original;

        // Se houver dados recorrentes (Folha de Ponto)
        if (recurrentKey && fileData[recurrentKey].length > 0) {
            // Cria MULTIPLAS LINHAS (uma por dia/item)
            fileData[recurrentKey].forEach(item => {
                const newRow = { ...staticData, ...item }; // Combina metadados + dados do dia
                Object.keys(item).forEach(key => allFileKeys.add(key)); // Adiciona headers do array
                consolidatedRows.push(newRow);
            });
        
        // Se for um documento simples (Nota Fiscal)
        } else {
            // Adiciona uma ÚNICA LINHA
            consolidatedRows.push(staticData);
        }
    });

    // Filtra os headers baseados na seleção do usuário (e garante 'arquivo_original' no início)
    const finalHeaders = Array.from(allFileKeys).filter(key => selectedKeys.includes(key));
    
    if (finalHeaders.includes('arquivo_original')) {
         finalHeaders.splice(finalHeaders.indexOf('arquivo_original'), 1);
    }
    finalHeaders.unshift('arquivo_original');
    
    // 2. Cria as colunas e injeta os dados (Formato Standard)
    worksheet.columns = finalHeaders.map(header => ({ 
        header: header.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
        key: header, 
        width: 25 
    }));
    
    worksheet.addRows(consolidatedRows); // Linhas multi-colunas

    // 3. Aplicação da Formatação Visual
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

    // Prompt AGORA DEVE PEDIR DADOS RECORRENTES COMO UM ARRAY DE OBJETOS
    const prompt = `
        Você é um assistente especialista em extração de dados estruturados para análise. Sua tarefa é analisar o documento e extrair todas as informações.

        REGRAS CRÍTICAS para o JSON:
        1. Se houver dados recorrentes (como uma lista de itens, produtos ou dias de ponto), crie uma chave (ex: "itens_comprados" ou "registros_diarios") cujo valor é um **ARRAY DE OBJETOS**.
        2. Dados estáticos (nome da empresa, data, ID) devem ser campos simples no objeto principal.
        3. Formate valores como números. Inclua 'resumo_executivo'. Retorne APENAS o JSON.
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
            allExtractedData.push({ 
                erro_processamento: `Falha na IA. ${err.message.substring(0, 50)}...`, 
                arquivo_original: file.originalname 
            });
            allUniqueKeys.add('erro_processamento');
        } finally {
            fileCleanupPromises.push(fs.promises.unlink(file.path));
        }
    }

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