// leitor.js (Versão Definitiva FINAL: Excel CONSOLIDADO, Horizontal, com Agrupamento Interativo)

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

// Garante que os diretórios existem
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

// Inicializa a API Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Configuração do Multer
const upload = multer({ dest: UPLOAD_DIR });

// Armazenamento em memória para dados de sessão. 
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
 * Lógica de Backoff Exponencial para lidar com o erro 429. (Mantida)
 */
async function callApiWithRetry(apiCall, maxRetries = 5) {
    let delay = 2; 
    for (let attempt = 0; attempt < maxRetries; attempt++) {
        try {
            return await apiCall();
        } catch (error) {
            const isRateLimitError = error.status === 429 ||
                                     (error.response && error.response.status === 429) ||
                                     (error.message && error.message.includes('Resource has been exhausted'));

            if (isRateLimitError) {
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

/**
 * AGRUPAMENTO OTIMIZADO: Agrupa chaves sequenciais/numéricas para o frontend. (Mantida)
 */
function groupKeys(keys) {
    const groupedKeys = new Set();
    
    // Regex aprimorado para cobrir a maioria dos padrões sequenciais
    const regex = /(\w+_\w+_\d+$)|(\w+_\d+$)|(_\d+$)|(_[a-z]$)|(\d+$)/i; 
    
    keys.forEach(key => {
        if (key === 'arquivo_original' || key === 'resumo_executivo') {
            groupedKeys.add(key);
            return;
        }

        const match = key.match(regex);

        if (match) {
            const groupedKey = key.replace(regex, '').replace(/_$/, ''); 
            
            if (groupedKey) {
                 groupedKeys.add(groupedKey);
            } else {
                 groupedKeys.add(key); 
            }
        } else {
            groupedKeys.add(key);
        }
    });

    return Array.from(groupedKeys).sort();
}

/**
 * **MUDANÇA FINAL:** Cria o arquivo Excel no formato HORIZONTAL, em UMA ÚNICA ABA, 
 * utilizando Agrupamento (Outline) para organizar os documentos.
 * @param {Array<Object>} allExtractedData - Dados extraídos de todos os documentos (DETALHADOS).
 * @param {string} outputPath - Caminho para salvar o arquivo.
 * @param {Array<string>} allDetailedKeys - Lista de TODAS as chaves detalhadas únicas de todos os documentos.
 */
async function createExcelFile(allExtractedData, outputPath, allDetailedKeys) {
    const workbook = new ExcelJS.Workbook();
    
    if (allExtractedData.length === 0) return;

    // Apenas UMA ABA para a consolidação horizontal
    const worksheet = workbook.addWorksheet('Dados Consolidados (Interativo)');
    
    // --- Configuração das Chaves (para garantir Arquivo e Resumo sempre na frente) ---
    const mandatoryKeys = ['arquivo_original', 'resumo_executivo'];
    const dynamicKeys = allDetailedKeys.filter(key => !mandatoryKeys.includes(key)).sort();
    const finalKeys = [...mandatoryKeys, ...dynamicKeys];

    // --- Funções de Ajuda ---
    const defineColumns = (keys) => {
        return keys.map(key => ({
            header: key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()),
            key: key,
            width: key.includes('resumo') ? 60 : 25, 
            style: { alignment: { wrapText: true, vertical: 'top' } }
        }));
    };
    
    const applyHeaderFormatting = (worksheet) => {
        worksheet.getRow(1).eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } }; 
            cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 12 };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        });
    };

    const formatDataCell = (cell, headerKey) => {
        // Formato de Moeda
        if (headerKey.includes('valor') || headerKey.includes('total')) {
            const numericValue = typeof cell.value === 'string' ? parseFloat(cell.value) : cell.value;
            if (!isNaN(numericValue) && numericValue !== null) {
                 cell.value = numericValue; 
                 cell.numFmt = 'R$ #,##0.00'; 
            }
        }
        // Formato de Data (Apenas alinhamento visual)
        if (headerKey.includes('data') || headerKey.includes('vencimento')) {
            if (typeof cell.value === 'string' && cell.value.match(/\d{2}\/\d{2}\/\d{4}/)) {
                cell.alignment = { horizontal: 'center' };
            }
        }
    };
    
    // 1. Define as Colunas
    worksheet.columns = defineColumns(finalKeys);
    
    // Inicializa o contador de linhas
    let rowIndex = 2; // Começa na linha 2, após o cabeçalho
    
    // 2. Preenche as Linhas e Adiciona Interatividade (Agrupamento/Outline)
    allExtractedData.forEach(data => {
        
        // --- Linha 1: Marcador de Documento (Layout Criativo/Separador) ---
        
        // Célula que serve como "cabeçalho" do grupo e identificador
        const markerRowData = {};
        markerRowData[finalKeys[0]] = `► DOCUMENTO: ${data.arquivo_original}`; // Usamos a primeira coluna como ID
        
        const markerRow = worksheet.addRow(markerRowData);
        markerRow.height = 20;
        
        // Estilo de linha separadora (vermelho escuro)
        markerRow.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBE3E3' } }; // Cinza/vermelho claro
            cell.font = { bold: true, size: 10, color: { argb: 'C62828' } }; 
            cell.border = { bottom: { style: 'thin', color: { argb: 'C62828' } } };
        });
        
        // --- Linhas 2-N: Dados Detalhados (Linha única por documento) ---
        
        // Adiciona a linha de dados reais
        const dataRow = worksheet.addRow(getExcelRow(data, finalKeys));
        dataRow.height = 20;

        // Aplica o agrupamento (Outline): A linha de dados está aninhada abaixo do marcador
        dataRow.outlineLevel = 1; 

        // Aplica formatação aos dados
        dataRow.eachCell((cell, colNumber) => {
            const headerKey = worksheet.getColumn(colNumber).key.toLowerCase();
            formatDataCell(cell, headerKey);

            // Adiciona um fundo levemente cinza nas linhas de dados para contraste
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FAFAFA' } }; 
        });

        // Aumenta o contador para a próxima iteração
        rowIndex += 2;
    });

    // 3. Aplica Formatação Final
    
    // Formatação do Cabeçalho (Fica em Row 1)
    applyHeaderFormatting(worksheet);
    
    // Define o nível de agrupamento padrão (expande ou colapsa tudo por padrão)
    worksheet.properties.outlineLevelCol = 0; // Coluna de agrupamento padrão
    worksheet.properties.outlineLevelRow = 1; // Colapsa todas as linhas de outline (docs individuais)


    // --- Finaliza o Arquivo ---
    await workbook.xlsx.writeFile(outputPath);
}

// --- 5. Endpoint Principal de Upload e Processamento (Mantido) ---
// ... (O código do endpoint /upload permanece inalterado)

app.post('/upload', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }

    const fileCleanupPromises = [];
    const allResultsForClient = [];
    const allResultsForExcel = [];
    const fieldLists = []; 
    const allDetailedKeys = new Set(); 

    // Prompt (Mantido)
    const prompt = `
        Você é um assistente especialista em extração de dados estruturados. Sua tarefa é analisar o documento anexado (que pode ser qualquer tipo de PDF ou IMAGEM) e extrair **TODAS** as informações relevantes.

        O objetivo é criar um objeto JSON plano onde cada chave é o nome da informação extraída e o valor é o dado correspondente.

        REGRAS CRÍTICAS para o JSON:
        1.  O resultado deve ser um objeto JSON **plano** (sem aninhamento).
        2.  Crie chaves JSON **dinamicamente** que sejam o nome mais descritivo para a informação (Ex: 'valor_total', 'nome_do_cliente').
        3.  Formate datas como 'DD/MM/AAAA' e valores monetários/quantias como números (ex: 123.45).
        4.  Inclua uma chave chamada 'resumo_executivo' com uma frase concisa descrevendo o documento, seguida por um resumo de todas as informações importantes extraídas.

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
                
                const keys = Object.keys(dynamicData);
                keys.forEach(key => allDetailedKeys.add(key)); 
                
                const groupedKeys = groupKeys(keys);

                fieldLists.push({
                    filename: file.originalname,
                    keys: groupedKeys
                });

                const clientResult = {
                    arquivo_original: file.originalname,
                    chave1: keys.length > 0 ? keys[0] : 'N/A',
                    valor1: keys.length > 0 ? dynamicData[keys[0]] : 'N/A',
                    chave2: keys.length > 1 ? keys[1] : 'N/A',
                    valor2: keys.length > 1 ? dynamicData[keys[1]] : 'N/A',
                    resumo: dynamicData.resumo_executivo || 'Resumo não fornecido.',
                };

                allResultsForClient.push(clientResult);
                
                dynamicData.arquivo_original = file.originalname;
                allResultsForExcel.push(dynamicData);
                
                await new Promise(resolve => setTimeout(resolve, 5000));

            } catch (err) {
                console.error(`Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ 
                    arquivo_original: file.originalname,
                    erro: `Falha na API: ${err.message.substring(0, 100)}` 
                });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        allDetailedKeys.add('arquivo_original');
        allDetailedKeys.add('resumo_executivo'); 
        
        const sessionId = Date.now().toString();
        sessionData[sessionId] = {
            data: allResultsForExcel, 
            fieldLists: fieldLists, 
            allDetailedKeys: Array.from(allDetailedKeys).sort() 
        };

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

// --- Endpoint para buscar as chaves AGRUPADAS POR ARQUIVO (Mantido) ---
app.get('/fields/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session) {
        return res.status(404).json({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    res.json({ fieldLists: session.fieldLists });
});

// --- Endpoint para Download do Excel (GET Simples) ---
app.get('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.data) {
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    const data = session.data;
    const allDetailedKeys = session.allDetailedKeys; // Pega o cabeçalho unificado

    const excelFileName = `extracao_gemini_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        // CHAMA createExcelFile com a lista de TODAS as chaves para criar o cabeçalho HORIZONTAL na ABA ÚNICA
        await createExcelFile(data, excelPath, allDetailedKeys); 

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

// --- 8. Servir front-end e iniciar servidor (Mantido) ---

app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});