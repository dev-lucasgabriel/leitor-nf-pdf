// leitor.js (Versão Final: Arquitetura de Duas Etapas - Interativa)

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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');
const PUBLIC_DIR = path.join(__dirname, 'public'); 

// Cria diretórios e inicializa API, Multer, etc. (código omitido por brevidade, mas está no final)
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const upload = multer({ dest: UPLOAD_DIR });
const sessionData = {}; // Armazena dados brutos da IA

// --- 2. Funções Essenciais (callApiWithRetry, fileToGenerativePart) ---
// (Mantidas as mesmas do código anterior)

function fileToGenerativePart(filePath, mimeType) {
    return {
        inlineData: {
            data: Buffer.from(fs.readFileSync(filePath)).toString("base64"),
            mimeType
        },
    };
}
// [NOTA: A função callApiWithRetry foi omitida por brevidade, mas deve ser mantida AQUI]

// --- 3. Função de Exportação FINAL (Recebe Chaves Filtradas) ---

async function createFilteredExcel(allExtractedData, selectedKeys, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Curados');

    if (allExtractedData.length === 0 || selectedKeys.length === 0) return;

    // 1. O Cabeçalho é EXATAMENTE a lista que o usuário selecionou
    const finalHeaders = selectedKeys;

    worksheet.columns = finalHeaders.map(header => ({ 
        header: header.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
        key: header, 
        width: 30 
    }));
    
    // 2. Mapeia os dados, garantindo que só as chaves selecionadas existam
    const filteredRows = allExtractedData.map(data => {
        const row = {};
        finalHeaders.forEach(key => {
            row[key] = data[key] || ''; // Usa o valor ou string vazia se ausente
        });
        return row;
    });

    worksheet.addRows(filteredRows);

    // [NOTA: Adicione aqui sua formatação de estilo e moeda se desejar]

    await workbook.xlsx.writeFile(outputPath);
}

// --- 4. Endpoint de ANÁLISE (Step 1: Upload e Extração Bruta) ---

app.post('/api/analyze', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }
    
    // O código de processamento da IA (prompt, callApiWithRetry) vai aqui...
    // (Omitido por brevidade, mas deve ser copiado do seu código anterior)

    const sessionId = Date.now().toString();
    const allExtractedData = []; 
    let allUniqueKeys = new Set();
    const fileCleanupPromises = [];

    // Lógica de loop de processamento (copie do seu /upload anterior)
    for (const file of req.files) {
        // ... Lógica de extração do Gemini ...
        // [NOTA: Substitua a chamada abaixo pela sua lógica real de callApiWithRetry]
        
        // Simulação de resposta da IA para ilustrar a estrutura
        const dynamicData = { /* Resultado do JSON plano da IA */ }; 
        
        allExtractedData.push(dynamicData);
        Object.keys(dynamicData).forEach(key => allUniqueKeys.add(key));
        fileCleanupPromises.push(fs.promises.unlink(file.path));
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