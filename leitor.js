// leitor.js (Versão Definitiva: Relatório Chave-Valor CONSOLIDADO)

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

// --- 3. Função de Exportação FINAL (Implementa o Relatório Chave-Valor CONSOLIDADO) ---

async function createFilteredExcel(allExtractedData, selectedKeys, outputPath) {
    const workbook = new ExcelJS.Workbook();
    // Aba ÚNICA e Consolidada
    const worksheet = workbook.addWorksheet('Relatório Consolidado');

    const finalHeaders = selectedKeys.filter(key => typeof key === 'string' && key.trim() !== '');

    if (allExtractedData.length === 0 || finalHeaders.length === 0) return;

    // 1. Configura colunas no formato VERTICAL (Chave | Valor | Arquivo Original)
    worksheet.columns = [
        { header: 'Campo Extraído', key: 'key', width: 35 },
        { header: 'Valor', key: 'value', width: 25 },
        { header: 'Arquivo Original', key: 'original_file', width: 45 }
    ];

    const validData = allExtractedData.filter(data => 
        data && typeof data === 'object' && !data.hasOwnProperty('erro_processamento')
    );

    if (validData.length === 0) {
        worksheet.addRow(['Nenhum dado válido para exportar.']);
        await workbook.xlsx.writeFile(outputPath);
        return; 
    }

    // 2. Loop para consolidar dados de TODOS os documentos na mesma aba
    validData.forEach((data) => {
        const originalName = data.arquivo_original || 'Nome do Arquivo Não Disponível';

        // Mapeia os dados (curados) para LINHAS VERTICAIS
        const filteredEntries = Object.entries(data)
            .filter(([key, value]) => finalHeaders.includes(key)); 

        const verticalRows = filteredEntries.map(([key, value]) => ({
            // Formata a chave para o cabeçalho "Campo Extraído"
            key: String(key || '').replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
            value: value,
            original_file: originalName
        }));
        
        worksheet.addRows(verticalRows);

        // Adiciona uma linha em branco para separar visualmente os documentos
        worksheet.addRow({});
    });


    // 3. Aplica Formatação (Estilo da Imagem)

    // Estilo do Cabeçalho (Fixo)
    worksheet.getRow(1).eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } }; 
        cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 12 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // Formatação de Valores, Quebra de Texto e Destaque
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { 
            const cellKey = row.getCell(1);
            const cellValue = row.getCell(2);

            if (cellKey.value) {
                 const keyText = cellKey.value.toString().toLowerCase();

                 // Destaque de linha para valores financeiros (Verde da Imagem)
                 if (keyText.includes('valor') || keyText.includes('total') || keyText.includes('icms') || keyText.includes('ipi') || keyText.includes('pis') || keyText.includes('cofins')) {
                    cellKey.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAD3' } }; // Verde claro
                    cellValue.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAD3' } };
                    
                    if (typeof cellValue.value === 'number') {
                        cellValue.numFmt = 'R$ #,##0.00'; 
                    }
                 }

                 // Quebra de Texto para Campos Longos
                 if (typeof cellValue.value === 'string' && cellValue.value.length > 50) {
                     cellValue.alignment = { wrapText: true, vertical: 'top' };
                 }
            }
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

    // Prompt Agnostico (Agora mais simples, já que a exportação trata o formato)
    const prompt = `
        Você é um assistente especialista em extração de dados estruturados. Sua tarefa é analisar o documento anexado (PDF ou IMAGEM) e extrair **TODAS** as informações relevantes. Crie um objeto JSON plano onde cada chave é o nome da informação extraída. Não inclua arrays (listas de itens); se houver, consolide-os em um campo de resumo. Retorne APENAS o JSON.
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
            // Adiciona o nome original do arquivo para rastreamento
            const dynamicData = { ...JSON.parse(response.text), arquivo_original: file.originalname };
            
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

    const excelFileName = `relatorio_curado_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        // Usa a função atualizada para gerar o relatório consolidado
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