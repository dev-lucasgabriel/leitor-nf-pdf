// leitor.js (Foco: Leitura de Ponto, Multi-Aba Horizontal por Arquivo, EXTRAÇÃO DIRETA VIA CSV)

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
 * Funções de Agregação de Dados de Ponto para o Resumo Mensal do Front-end. (Mantida)
 */
function aggregatePointData(dataList) {
    const monthlySummary = {};

    dataList.forEach(data => {
        const nome = data.nome_colaborador || 'Desconhecido';
        const horasDiarias = parseFloat(data.total_horas_trabalhadas) || 0; 
        const horasExtras = parseFloat(data.horas_extra_diarias) || 0;

        if (!monthlySummary[nome]) {
            monthlySummary[nome] = { totalHoras: 0, totalExtras: 0 };
        }
        
        monthlySummary[nome].totalHoras += horasDiarias;
        monthlySummary[nome].totalExtras += horasExtras;
    });

    Object.keys(monthlySummary).forEach(nome => {
        monthlySummary[nome].totalHoras = monthlySummary[nome].totalHoras.toFixed(2);
        monthlySummary[nome].totalExtras = monthlySummary[nome].totalExtras.toFixed(2);
    });

    return monthlySummary;
}

/**
 * NOVO: Função para converter a string CSV dos registros diários em objetos JavaScript (Mais Robusto)
 */
function parseCsvRecords(csvString, filename) {
    const lines = csvString.trim().split('\n').filter(line => line.trim().length > 0);
    if (lines.length < 2) return [];

    // O cabeçalho é a primeira linha
    const rawHeaders = lines[0].split(';');
    // Sanitiza cabeçalhos: remove espaços, caracteres especiais, e deixa em snake_case
    const headers = rawHeaders.map(h => h.trim().toLowerCase().replace(/[^a-z0-9_]/g, ''));
    
    const records = [];
    let nomeColaboradorGlobal = 'Desconhecido';

    // Processa as linhas de dados (a partir da segunda linha)
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(';');
        
        // Verifica se a linha de dados é muito curta (possível lixo ou linha vazia)
        if (values.length === 0) {
             continue; 
        }

        const record = {
            arquivo_original: filename,
        };
        
        let foundValidData = false;

        headers.forEach((header, index) => {
            if (header && values[index] !== undefined) {
                let value = values[index].trim();
                
                // Trata conversão de HH:MM para decimal no caso de horas (para o Excel)
                if (header.includes('horas') && value.includes(':')) {
                    const [h, m] = value.split(':');
                    // Garante que o valor não é lixo antes de tentar o parse
                    if (!isNaN(parseInt(h)) && !isNaN(parseInt(m))) {
                        value = (parseInt(h) + (parseInt(m) / 60)).toFixed(2);
                    }
                }
                
                record[header] = value || 'N/A';
                
                // Verifica se encontramos pelo menos um dado útil
                if (value && value !== 'N/A' && !header.includes('arquivo')) {
                    foundValidData = true;
                }
                
                // Lógica para capturar e propagar o nome do colaborador
                if (header === 'nome_colaborador' && value && value !== 'N/A') {
                    nomeColaboradorGlobal = value;
                }
            }
        });
        
        // Se a linha tiver dados válidos E tiver uma data, adicionamos.
        record.nome_colaborador = record.nome_colaborador && record.nome_colaborador !== 'N/A' ? record.nome_colaborador : nomeColaboradorGlobal;
        
        // A lógica mais robusta: se encontramos dados válidos E há uma data (ou é a primeira linha, para o caso de a IA falhar na data no primeiro registro)
        if (foundValidData && (record.data_registro && record.data_registro !== 'N/A' || i === 1)) {
             records.push(record);
        }
    }
    return records.filter(r => r.nome_colaborador !== 'Desconhecido');
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
 * Cria o arquivo Excel no formato HORIZONTAL, com UMA ABA POR ARQUIVO. (Mantida)
 */
async function createExcelFile(allExtractedData, outputPath, allDetailedKeys) {
    const workbook = new ExcelJS.Workbook();
    
    if (allExtractedData.length === 0) {
        // Lançar um erro para que o bloco try/catch do download o capture!
        throw new Error("Não há dados válidos para gerar o arquivo Excel.");
    }

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

// --- 3. Rotas e Endpoints ---

// Endpoint Principal de Upload e Processamento
app.post('/upload', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo enviado.' });
    }

    const fileCleanupPromises = []; 
    const allResultsForClient = [];
    const allResultsForExcel = [];
    const fieldLists = []; 
    const allDetailedKeys = new Set(); 
    
    // PROMPT OTIMIZADO PARA CSV: Pede uma tabela separada por ponto e vírgula
    const prompt = `
        Você é um assistente especialista em extração de registros de ponto.
        Sua tarefa é analisar o documento anexado (cartão de ponto, espelho ou folha de registro) e extrair os registros diários de forma estruturada.

        REGRAS CRÍTICAS DE EXTRAÇÃO:
        1. **GARANTIA DE COMPLETUDE (CRÍTICO):** Analise a estrutura visual do documento e garanta que você extraiu **TODAS** as linhas de registro. Não pule ou ignore nenhuma data, mesmo que os campos de horário estejam vazios (registre como 'N/A'). Se o documento for de quinzena, garanta 15 registros.
        2. **FORMATO DE SAÍDA (CRÍTICO):** Retorne os registros diários como uma tabela **CSV** separada por **ponto e vírgula (;)**, seguida por um objeto JSON de resumo.
        3. **FORMATO CSV:** A primeira linha deve ser o cabeçalho. As colunas devem incluir: Nome_Colaborador, Data_Registro (Formato: DD/MM/AAAA), Entrada_1 (HH:MM ou N/A), Saida_1 (HH:MM ou N/A), Total_Horas_Trabalhadas (Número extraído), Horas_Extra_Diarias (Número extraído), Horas_Falta_Diarias (Número extraído).
        4. **RESUMO:** Após a tabela CSV, inclua o resumo mensal em um bloco de código JSON.

        Retorne APENAS a tabela CSV (texto puro) seguida pelo bloco de código JSON do resumo.
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
                
                // Tenta encontrar a divisão entre CSV e JSON
                const lastBraceIndex = responseText.lastIndexOf('}');
                let csvPart = responseText;
                let jsonPart = null;

                if (lastBraceIndex !== -1) {
                    const potentialJsonStart = responseText.lastIndexOf('{', lastBraceIndex);
                    if (potentialJsonStart !== -1) {
                        csvPart = responseText.substring(0, potentialJsonStart).trim();
                        jsonPart = responseText.substring(potentialJsonStart).trim();
                    }
                }
                
                // Lógica de limpeza do JSON (se a IA usou Markdown)
                if (jsonPart && jsonPart.endsWith('```')) {
                    jsonPart = jsonPart.substring(0, jsonPart.lastIndexOf('```')).trim();
                }

                // 1. Processa os Registros Diários (CSV)
                const dailyRecords = parseCsvRecords(csvPart, file.originalname);
                
                // 2. Processa o Resumo (JSON)
                let monthlySummaryPlaceholder = null;
                if (jsonPart) {
                    try {
                        monthlySummaryPlaceholder = JSON.parse(jsonPart);
                    } catch (e) {
                         console.warn(`[ERRO JSON] Falha ao parsear resumo JSON para ${file.originalname}: ${e.message}`);
                    }
                }
                
                // Usa dailyRecords para popular allResultsForExcel e allDetailedKeys
                dailyRecords.forEach(record => {
                    const finalRecord = { ...record };
                    
                    Object.keys(finalRecord).forEach(key => allDetailedKeys.add(key));
                    allResultsForExcel.push(finalRecord);
                });

                // Prepara dados para o Front-end
                if (dailyRecords.length > 0) {
                    const firstRecord = dailyRecords[0];
                    const keys = Object.keys(firstRecord);
                    const groupedKeys = groupKeys(keys);

                    fieldLists.push({ filename: file.originalname, keys: groupedKeys });
                    
                    allResultsForClient.push({
                        arquivo_original: file.originalname,
                        nome_colaborador: firstRecord.nome_colaborador || 'N/A',
                        total_horas: firstRecord.total_horas_trabalhadas || '0.00 (Extr. AI)',
                        horas_extra: firstRecord.horas_extra_diarias || '0.00 (Extr. AI)',
                        resumo: monthlySummaryPlaceholder ? JSON.stringify(monthlySummaryPlaceholder) : 'Extração de dados brutos concluída.',
                    });
                }
                
            } catch (err) {
                console.error(`Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ 
                    arquivo_original: file.originalname,
                    erro: `Falha na API: ${err.message}. Verifique o formato de retorno da IA.`
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

// Endpoint para buscar FieldLists
app.get('/fields/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.fieldLists) {
        return res.status(404).json({ error: 'Dados da sessão não encontrados.' });
    }
    
    res.json({ fieldLists: session.fieldLists });
});


// Endpoint para buscar Resumo Mensal do Colaborador
app.get('/summary/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.monthlySummary) {
        return res.status(404).json({ error: 'Resumo da sessão não encontrado.' });
    }
    
    res.json({ summary: session.monthlySummary });
});

// Endpoint para Download do Excel
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
    
    let excelCreated = false;

    try {
        await createExcelFile(data, excelPath, allDetailedKeys); 
        excelCreated = true;

        res.download(excelPath, excelFileName, async (err) => {
            // Se houver um erro no download (após a criação), ainda tentamos limpar
            if (err) {
                console.error("Erro ao enviar o Excel:", err);
            }
            if (excelCreated) {
                 await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel:", e));
            }
            delete sessionData[sessionId]; 
        });
    } catch (error) {
        // Se a criação do Excel falhou (erro de lógica/ENOENT), capturamos aqui
        console.error('Erro ao gerar Excel:', error);
        res.status(500).send({ error: `Falha ao gerar o arquivo Excel: ${error.message}` });
        
        // Limpeza adicional em caso de falha de criação
        if (excelCreated) {
            await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel após falha de criação:", e));
        }
    }
});

// --- 4. Iniciar Servidor ---
app.use(cors()); 

app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});