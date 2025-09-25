// leitor.js (Foco: Leitura de Ponto, Multi-Aba Horizontal por Arquivo, SEM CÁLCULO NO BACKEND)

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

// --- Configurações de Ambiente e Variáveis Globais ---
const app = express();
const PORT = process.env.PORT || 3000; 

app.use(express.json()); 

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');
const PUBLIC_DIR = path.join(__dirname, 'public'); 

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

// Inicializa a API Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

const upload = multer({ dest: UPLOAD_DIR });

// Armazenamento de Sessão (Agora armazena dados de ponto e resumo para o Front-end)
const sessionData = {};

// --- 2. Funções Essenciais ---

function fileToGenerativePart(filePath, mimeType) {
    return {
        inlineData: {
            data: Buffer.from(fs.readFileSync(filePath)).toString("base64"),
            mimeType
        },
    };
}

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
 * Funções de Agregação de Dados de Ponto para o Resumo Mensal do Front-end.
 * Agora soma os campos extraídos DIRETAMENTE pela IA (se existirem).
 */
function aggregatePointData(dataList) {
    const monthlySummary = {};

    dataList.forEach(data => {
        const nome = data.nome_colaborador || 'Desconhecido';
        // Agora usa o campo extraído pela IA
        const horasDiarias = parseFloat(data.total_horas_trabalhadas) || 0; 
        const horasExtras = parseFloat(data.horas_extra_diarias) || 0;

        if (!monthlySummary[nome]) {
            monthlySummary[nome] = { totalHoras: 0, totalExtras: 0 };
        }
        
        monthlySummary[nome].totalHoras += horasDiarias;
        monthlySummary[nome].totalExtras += horasExtras;
    });

    // Formata os totais para o frontend
    Object.keys(monthlySummary).forEach(nome => {
        monthlySummary[nome].totalHoras = monthlySummary[nome].totalHoras.toFixed(2);
        monthlySummary[nome].totalExtras = monthlySummary[nome].totalExtras.toFixed(2);
    });

    return monthlySummary;
}

/**
 * AGRUPAMENTO OTIMIZADO: Agrupa chaves sequenciais/numéricas para o frontend. (Mantida)
 */
function groupKeys(keys) {
    const groupedKeys = new Set();
    
    const regex = /(\w+_\d+$)|(_\d+$)/i; 
    
    keys.forEach(key => {
        if (key === 'arquivo_original' || key === 'resumo_executivo_mensal') {
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
 * Cria o arquivo Excel no formato HORIZONTAL, com UMA ABA POR ARQUIVO, 
 * incluindo uma Linha de Resumo com Fórmulas para o usuário.
 */
async function createExcelFile(allExtractedData, outputPath, allDetailedKeys) {
    const workbook = new ExcelJS.Workbook();
    
    if (allExtractedData.length === 0) return;

    // Agrupa todos os registros diários pelo nome do arquivo de origem
    const dataByFile = allExtractedData.reduce((acc, data) => {
        const filename = data.arquivo_original;
        if (!acc[filename]) {
            acc[filename] = [];
        }
        acc[filename].push(data);
        return acc;
    }, {});

    // --- Configuração das Chaves ---
    // Inclui as chaves de totais, esperando que a IA as forneça
    const orderedKeys = ['nome_colaborador', 'data_registro', 'entrada_1', 'saida_1', 'total_horas_trabalhadas', 'horas_extra_diarias', 'horas_falta_diarias', 'resumo_executivo_mensal', 'arquivo_original'];
    const dynamicKeys = allDetailedKeys.filter(key => !orderedKeys.includes(key)).sort();
    
    // Lista final de chaves para o cabeçalho
    const finalKeys = Array.from(new Set([...orderedKeys.filter(key => allDetailedKeys.includes(key)), ...dynamicKeys]));


    // --- Funções de Ajuda e Formatação ---
    const defineColumns = (keys) => {
        return keys.map(key => ({
            header: key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()),
            key: key,
            width: key.includes('resumo') ? 60 : (key.includes('total_horas') || key.includes('horas') ? 18 : 15), 
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

    const formatDataCell = (cell, headerKey, isSummaryRow = false) => {
        if (headerKey.includes('total_horas') || headerKey.includes('horas_extra') || headerKey.includes('horas_falta')) {
            // O valor é extraído diretamente da IA, formatamos como número
            if (isSummaryRow) {
                 cell.numFmt = '0.00'; 
            } else {
                const numericValue = typeof cell.value === 'string' ? parseFloat(cell.value) : cell.value;
                if (!isNaN(numericValue) && numericValue !== null) {
                     cell.value = numericValue; 
                     cell.numFmt = '0.00'; 
                }
            }
        } else if (headerKey.includes('data') || headerKey.includes('entrada') || headerKey.includes('saida')) {
            cell.alignment = { horizontal: 'center' };
        }
    };
    
    // --- Criação das Abas ---
    const usedSheetNames = new Set();

    for (const filename in dataByFile) {
        const records = dataByFile[filename];
        
        // --- Lógica de Nome da Aba ---
        const unsafeName = filename.replace(/\.[^/.]+$/, "");
        let baseName = unsafeName.substring(0, 28).replace(/[\[\]\*\:\/\?\\\,]/g, ' ');
        baseName = baseName.trim().replace(/\.$/, ''); 
        
        let worksheetName = baseName;
        let counter = 1;
        while (usedSheetNames.has(worksheetName)) {
            worksheetName = `${baseName.substring(0, 25)} (${counter})`; 
            counter++;
        }
        usedSheetNames.add(worksheetName); 
        // -----------------------------

        const worksheet = workbook.addWorksheet(worksheetName || `Arquivo ${filename}`);
        
        // 1. Define as Colunas
        worksheet.columns = defineColumns(finalKeys);
        
        // 2. Preenche as Linhas de Detalhe Diário
        records.forEach((record, index) => {
            const row = worksheet.addRow(record);
            row.height = 18;
            
            // Fundo da linha diária para contraste (listras)
            const fillColor = index % 2 === 0 ? 'F0F0F0' : 'FFFFFF';
            row.eachCell((cell, colNumber) => {
                const headerKey = worksheet.getColumn(colNumber).key.toLowerCase();
                formatDataCell(cell, headerKey, false);
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            });
        });

        // --- 3. Linha de Resumo com Fórmulas (Automação de Cálculo) ---
        const resumoRow = worksheet.addRow({});
        resumoRow.height = 25;
        
        const firstDataRow = 2; // Linha 2 é onde os dados começam
        const lastDataRow = worksheet.lastRow.number; 
        
        // Colunas onde as somas serão inseridas
        const totalCol = finalKeys.indexOf('total_horas_trabalhadas') + 1;
        const extraCol = finalKeys.indexOf('horas_extra_diarias') + 1;
        const faltaCol = finalKeys.indexOf('horas_falta_diarias') + 1;
        
        // Insere as fórmulas para SOMAR os valores extraídos/calculados pela IA
        resumoRow.getCell(1).value = 'RESUMO MENSAL / FÓRMULAS:';
        if (totalCol > 0) resumoRow.getCell(totalCol).value = { formula: `SUM(${worksheet.getColumn(totalCol).letter}${firstDataRow}:${worksheet.getColumn(totalCol).letter}${lastDataRow})` };
        if (extraCol > 0) resumoRow.getCell(extraCol).value = { formula: `SUM(${worksheet.getColumn(extraCol).letter}${firstDataRow}:${worksheet.getColumn(extraCol).letter}${lastDataRow})` };
        if (faltaCol > 0) resumoRow.getCell(faltaCol).value = { formula: `SUM(${worksheet.getColumn(faltaCol).letter}${firstDataRow}:${worksheet.getColumn(faltaCol).letter}${lastDataRow})` };
        
        // Formatação do resumo
        resumoRow.eachCell((cell, colNumber) => {
            const column = worksheet.getColumn(colNumber);
            if (!column || !column.key) return; 

            const headerKey = column.key.toLowerCase();
            
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } }; // Verde Escuro
            cell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 10 };
            formatDataCell(cell, headerKey, true);
        });

        // 4. Aplica Formatação Final
        applyHeaderFormatting(worksheet);
    }

    // --- Finaliza o Arquivo ---
    await workbook.xlsx.writeFile(outputPath);
}

// --- 5. Endpoint Principal de Upload e Processamento ---
app.post('/upload', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }

    const fileCleanupPromises = []; 
    const allResultsForClient = [];
    const allResultsForExcel = [];
    const fieldLists = []; 
    const allDetailedKeys = new Set(); 
    
    // PROMPT REVISADO: Pede todos os campos, incluindo os totais de horas, que agora são responsabilidade da IA ou do documento.
    const prompt = `
        Você é um assistente especialista em extração de registros de ponto.
        Sua tarefa é analisar o documento anexado (cartão de ponto, espelho ou folha de registro) e extrair os registros diários de forma estruturada.

        REGRAS CRÍTICAS para o JSON:
        1. O resultado deve ser um **array de objetos JSON**. Cada objeto no array representa **UM ÚNICO REGISTRO DIÁRIO**.
        2. Para CADA REGISTRO DIÁRIO, extraia todas as chaves de horário e totais disponíveis no documento. As chaves mais comuns são:
           - **nome_colaborador**
           - **data_registro** (Formato: 'DD/MM/AAAA')
           - **entrada_1** (Horário: 'HH:MM' - 24h)
           - **saida_1** (Horário: 'HH:MM' - 24h)
           - **total_horas_trabalhadas** (Valor numérico extraído)
           - **horas_extra_diarias** (Valor numérico extraído)
           - **horas_falta_diarias** (Valor numérico extraído)
        3. Se houver mais de um par de entrada/saída (Ex: almoço), use **entrada_2**, **saida_2**, etc.
        4. Se o documento NÃO fornecer os totais (como total_horas_trabalhadas), a IA DEVE tentar calculá-los. Caso contrário, extraia o valor do documento.
        5. **O resultado DEVE ser encapsulado em um bloco de código JSON Markdown.**
        6. Inclua um objeto no final do array com a chave 'resumo_executivo_mensal' contendo uma string informativa.

        Retorne **APENAS** o bloco de código JSON completo, sem formatação extra.
    `;
    
    try {
        for (const file of req.files) {
            console.log(`[PROCESSANDO] ${file.originalname}`);
            
            const filePart = fileToGenerativePart(file.path, file.mimetype);
            
            const apiCall = () => ai.models.generateContent({
                model: 'gemini-1.5-flash',
                contents: [filePart, { text: prompt }],
            });

            try {
                const response = await callApiWithRetry(apiCall);
                let responseText = response.text.trim();
                
                // Lógica de limpeza mais robusta do bloco de código Markdown
                if (responseText.startsWith('```')) {
                    const firstLineEnd = responseText.indexOf('\n');
                    responseText = responseText.substring(firstLineEnd).trim();
                }
                if (responseText.endsWith('```')) {
                    const lastBlockEnd = responseText.lastIndexOf('```');
                    if (lastBlockEnd > 0) {
                        responseText = responseText.substring(0, lastBlockEnd).trim();
                    }
                }
                
                const results = JSON.parse(responseText);

                const dailyRecords = Array.isArray(results) ? results.filter(r => !r.resumo_executivo_mensal) : [results];
                const monthlySummaryPlaceholder = Array.isArray(results) ? results.find(r => r.resumo_executivo_mensal) : null;
                
                dailyRecords.forEach(record => {
                    // CÁLCULO REMOVIDO: A IA/Documento fornecem os totais.
                    const finalRecord = { ...record };
                    
                    finalRecord.arquivo_original = file.originalname;
                    Object.keys(finalRecord).forEach(key => allDetailedKeys.add(key));
                    allResultsForExcel.push(finalRecord);
                });

                // Prepara dados para o Front-end
                if (dailyRecords.length > 0) {
                    const firstRecord = dailyRecords[0];
                    const keys = Object.keys(firstRecord);
                    const groupedKeys = groupKeys(keys);

                    fieldLists.push({ filename: file.originalname, keys: groupedKeys });
                    
                    // Usa os records extraídos (agora com os totais extraídos pela IA)
                    allResultsForClient.push({
                        arquivo_original: file.originalname,
                        nome_colaborador: firstRecord.nome_colaborador || 'N/A',
                        total_horas: firstRecord.total_horas_trabalhadas || '0.00 (Extr. IA)',
                        horas_extra: firstRecord.horas_extra_diarias || '0.00 (Extr. IA)',
                        resumo: monthlySummaryPlaceholder ? monthlySummaryPlaceholder.resumo_executivo_mensal : 'Extração de dados brutos concluída.',
                    });
                }
                
            } catch (err) {
                console.error(`Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ 
                    arquivo_original: file.originalname,
                    erro: `Falha na API: ${err.message}. Verifique a formatação do JSON.`
                });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        // Inclui chaves essenciais na lista de cabeçalho do Excel
        const orderedKeys = ['nome_colaborador', 'data_registro', 'entrada_1', 'saida_1', 'total_horas_trabalhadas', 'horas_extra_diarias', 'horas_falta_diarias', 'resumo_executivo_mensal', 'arquivo_original'];
        orderedKeys.forEach(key => allDetailedKeys.add(key));
        
        const sessionId = Date.now().toString();
        sessionData[sessionId] = {
            data: allResultsForExcel, 
            fieldLists: fieldLists, 
            allDetailedKeys: Array.from(allDetailedKeys).sort(),
            monthlySummary: aggregatePointData(allResultsForExcel) 
        };

        return res.json({ 
            results: allResultsForClient,
            sessionId: sessionId,
            summary: sessionData[sessionId].monthlySummary 
        });

    } catch (error) {
        console.error('Erro fatal no processamento:', error);
        return res.status(500).send({ error: 'Erro interno do servidor.' });
    } finally {
        await Promise.all(fileCleanupPromises).catch(e => console.error("Erro ao limpar arquivos temporários:", e));
    }
});

// --- Outros Endpoints (Mantidos) ---
app.get('/fields/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.fieldLists) {
        return res.status(404).json({ error: 'Dados da sessão não encontrados.' });
    }
    
    res.json({ fieldLists: session.fieldLists });
});

app.get('/summary/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.monthlySummary) {
        return res.status(404).json({ error: 'Resumo da sessão não encontrado.' });
    }
    
    res.json({ summary: session.monthlySummary });
});

app.get('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.data) {
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    const data = session.data;
    const allDetailedKeys = session.allDetailedKeys; 

    const excelFileName = `extracao_ponto_detalhado_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
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

// --- Servir front-end e iniciar servidor (Mantido) ---
app.use(cors()); 
app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});