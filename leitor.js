// leitor.js (VERSÃO FINAL COMPLETA E ROBUSTA)

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

// Garante que os diretórios existam
[UPLOAD_DIR, TEMP_DIR, PUBLIC_DIR].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// Inicializa a API Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY }); 

const upload = multer({ dest: UPLOAD_DIR });

// Armazenamento de Sessão
const sessionData = {};

// --- 2. Funções Essenciais ---

function fileToGenerativePart(filePath, mimeType) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`Arquivo não encontrado no caminho: ${filePath}`);
    }
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
            const isRateLimitError = (error.status === 429) || 
                                     (error.message && error.message.includes('Resource has been exhausted'));

            if (isRateLimitError) {
                if (attempt === maxRetries - 1) {
                    throw new Error('Limite de taxa excedido (429) após múltiplas tentativas. Tente novamente mais tarde.');
                }
                const jitter = random.uniform(0, 1)(); 
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
 * Funções de Agregação de Dados de Ponto para o Resumo Mensal.
 */
function aggregatePointData(dataList) {
    const monthlySummary = {};
    dataList.forEach(data => {
        const nome = data.nome_colaborador || 'Desconhecido';
        const horasDiarias = typeof data.total_horas_trabalhadas === 'number' ? data.total_horas_trabalhadas : parseFloat(String(data.total_horas_trabalhadas).replace(',', '.')) || 0; 
        const horasExtras = typeof data.horas_extra_diarias === 'number' ? data.horas_extra_diarias : parseFloat(String(data.horas_extra_diarias).replace(',', '.')) || 0;

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
 * REFORÇADO: Converte a string CSV da IA em objetos JavaScript.
 * Blindagem contra a IA errando o separador e o formato numérico.
 */
function parseCsvRecords(csvString, filename) {
    const lines = csvString.trim().split('\n').filter(line => line.trim().length > 0);
    if (lines.length < 2) {
        console.warn(`[PARSE CSV] Documento ${filename} retornou menos de 2 linhas.`);
        return [];
    }

    // 1. Detecção de Separador Robusta
    const headerLine = lines[0];
    const separatorCandidates = [';', ','];
    let bestSeparator = ',';
    let maxCount = 0;

    for (const sep of separatorCandidates) {
        const count = headerLine.split(sep).length;
        if (count > maxCount) {
            maxCount = count;
            bestSeparator = sep;
        }
    }
    const separator = maxCount > 1 ? bestSeparator : ','; 
    
    // 2. Processamento e Limpeza do Cabeçalho
    const rawHeaders = headerLine.split(separator);
    const headers = rawHeaders.map(h => {
        return h.trim()
                .toLowerCase()
                .replace(/\s+/g, '_')     
                .replace(/[^a-z0-9_]/g, '') 
                .replace(/_+/g, '_');    
    }).filter(h => h.length > 0);
    
    if (headers.length < 3) {
        console.error(`[PARSE CSV] Cabeçalhos inválidos após limpeza para ${filename}.`);
        return [];
    }

    const records = [];
    let nomeColaboradorGlobal = 'Desconhecido';

    // 3. Processamento das Linhas de Dados
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(separator);
        
        // Se a linha tiver menos colunas que o esperado, pula.
        if (values.length < headers.length) {
            continue; 
        }

        const record = { arquivo_original: filename };
        let foundValidData = false;

        headers.forEach((header, index) => {
            if (header && values[index] !== undefined) {
                let value = values[index].trim();
                
                // Conversão HH:MM para decimal (PARA AS HORAS TOTAIS)
                if (header.includes('horas') && value.includes(':')) {
                    const parts = value.split(':').map(p => parseInt(p.trim(), 10));
                    if (parts.length === 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
                        // CRÍTICO: Salva o número decimal para soma no Excel
                        value = (parts[0] + (parts[1] / 60)); 
                    } else {
                        value = value; 
                    }
                } else if (header.includes('horas') || header.includes('total')) {
                    // CRÍTICO: Tenta garantir que campos de hora/total sejam números, aceitando vírgula ou ponto
                    const numValue = parseFloat(value.replace(',', '.'));
                    value = isNaN(numValue) ? value : numValue;
                }
                
                record[header] = value === '' ? 'N/A' : value;
                
                if (value !== '' && value !== 'N/A' && !header.includes('arquivo')) {
                    foundValidData = true;
                }
                
                // Propagação do nome do colaborador
                if (header.includes('nome_colaborador') && typeof value === 'string' && value.length > 3) {
                    nomeColaboradorGlobal = value;
                }
            }
        });
        
        // Propaga o nome do colaborador para linhas que não o têm
        if (!record.nome_colaborador || record.nome_colaborador === 'N/A' || record.nome_colaborador.length < 3) {
             record.nome_colaborador = nomeColaboradorGlobal;
        }

        if (foundValidData && record.nome_colaborador !== 'Desconhecido') {
             records.push(record);
        }
    }
    
    return records.filter(r => r.data_registro || r.nome_colaborador !== 'Desconhecido');
}


/**
 * AGRUPAMENTO OTIMIZADO: Agrupa chaves sequenciais/numéricas para o frontend. (MANTIDA)
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
 * Cria o arquivo Excel no formato HORIZONTAL, com UMA ABA POR ARQUIVO. (SEM ALTERAÇÕES ESTRUTURAIS)
 */
async function createExcelFile(allExtractedData, outputPath, allDetailedKeys) {
    const workbook = new ExcelJS.Workbook();
    
    if (allExtractedData.length === 0) {
        throw new Error("Não há dados válidos para gerar o arquivo Excel.");
    }

    const dataByFile = allExtractedData.reduce((acc, data) => {
        const filename = data.arquivo_original;
        if (!acc[filename]) {
            acc[filename] = [];
        }
        acc[filename].push(data);
        return acc;
    }, {});

    // Adição de entrada_2 e saida_2 na lista de ordenação de chaves
    const orderedKeys = ['nome_colaborador', 'data_registro', 'entrada_1', 'saida_1', 'entrada_2', 'saida_2', 'total_horas_trabalhadas', 'horas_extra_diarias', 'horas_falta_diarias', 'resumo_executivo_mensal', 'arquivo_original'];
    const dynamicKeys = allDetailedKeys.filter(key => !orderedKeys.includes(key)).sort();
    
    const finalKeys = Array.from(new Set([...orderedKeys.filter(key => allDetailedKeys.includes(key)), ...dynamicKeys]));


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
                // Se o valor já foi processado pelo parseCsvRecords como número, ele é tratado aqui
                const numericValue = typeof cell.value === 'string' ? parseFloat(String(cell.value).replace(',', '.')) : cell.value;
                if (!isNaN(numericValue) && numericValue !== null) {
                     cell.value = numericValue; 
                     cell.numFmt = '0.00'; 
                }
            }
        } else if (headerKey.includes('data') || headerKey.includes('entrada') || headerKey.includes('saida')) {
            cell.alignment = { horizontal: 'center' };
        }
    };
    
    const usedSheetNames = new Set();

    for (const filename in dataByFile) {
        const records = dataByFile[filename];
        
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

        const worksheet = workbook.addWorksheet(worksheetName || `Arquivo ${filename}`);
        
        worksheet.columns = defineColumns(finalKeys);
        
        records.forEach((record, index) => {
            const row = worksheet.addRow(record);
            row.height = 18;
            
            const fillColor = index % 2 === 0 ? 'F0F0F0' : 'FFFFFF';
            row.eachCell((cell, colNumber) => {
                const headerKey = worksheet.getColumn(colNumber).key.toLowerCase();
                formatDataCell(cell, headerKey, false);
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            });
        });

        const resumoRow = worksheet.addRow({});
        resumoRow.height = 25;
        
        const firstDataRow = 2; 
        const lastDataRow = worksheet.lastRow.number - 1; 
        
        const totalCol = finalKeys.indexOf('total_horas_trabalhadas') + 1;
        const extraCol = finalKeys.indexOf('horas_extra_diarias') + 1;
        const faltaCol = finalKeys.indexOf('horas_falta_diarias') + 1;
        
        if (lastDataRow >= firstDataRow) {
             resumoRow.getCell(1).value = 'RESUMO MENSAL / FÓRMULAS:';
             if (totalCol > 0) resumoRow.getCell(totalCol).value = { formula: `SUM(${worksheet.getColumn(totalCol).letter}${firstDataRow}:${worksheet.getColumn(totalCol).letter}${lastDataRow})` };
             if (extraCol > 0) resumoRow.getCell(extraCol).value = { formula: `SUM(${worksheet.getColumn(extraCol).letter}${firstDataRow}:${worksheet.getColumn(extraCol).letter}${lastDataRow})` };
             if (faltaCol > 0) resumoRow.getCell(faltaCol).value = { formula: `SUM(${worksheet.getColumn(faltaCol).letter}${firstDataRow}:${worksheet.getColumn(faltaCol).letter}${lastDataRow})` };
        } else {
             resumoRow.getCell(1).value = 'RESUMO MENSAL: SEM DADOS VÁLIDOS PARA SOMA';
        }

        resumoRow.eachCell((cell, colNumber) => {
            const column = worksheet.getColumn(colNumber);
            if (!column || !column.key) return; 

            const headerKey = column.key.toLowerCase();
            
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } }; 
            cell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 10 };
            formatDataCell(cell, headerKey, true);
        });

        applyHeaderFormatting(worksheet);
    }

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
    
    // PROMPT OTIMIZADO PARA EXTRAÇÃO SEM ERRO
    const prompt = `
        Você é um assistente especialista em extração de registros de ponto.
        Sua única e crítica tarefa é analisar o documento anexado (cartão de ponto, espelho ou folha de registro) e extrair os registros diários de forma estruturada.

        REGRAS CRÍTICAS DE EXTRAÇÃO:
        1. **GARANTIA DE COMPLETUDE (CRÍTICO):** Extraia **TODAS** as linhas de registro.
        2. **FORMATO DE SAÍDA (CRÍTICO):** Retorne estritamente o seguinte formato:
            a) Tabela **CSV** separada por **ponto e vírgula (;)** (texto puro).
            b) Bloco de código **JSON** com o resumo.

        3. **FORMATO CSV OBRIGATÓRIO (NÃO MUDAR OS CABEÇALHOS):**
           Cabeçalho: Nome_Colaborador;Data_Registro;Entrada_1;Saida_1;Entrada_2;Saida_2;Total_Horas_Trabalhadas;Horas_Extra_Diarias;Horas_Falta_Diarias
           Valores Ausentes: Use **N/A** para horários e datas que não puderem ser extraídos.
           Formato de Horas/Minutos: Use **HH:MM** (Ex: 08:30) para as entradas e saídas, e **decimal** (Ex: 8.5) para os totais de horas.

        4. **RESUMO JSON:** Após a tabela CSV, inclua o resumo em um bloco de código JSON.

        Retorne APENAS a tabela CSV (texto puro), imediatamente seguida pelo bloco de código JSON do resumo. Não adicione nenhum texto introdutório, explicações ou marcadores de código Markdown ('\`\`\`csv' ou '\`\`\`json').
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
                
                // Separação de CSV e JSON
                const lastBraceIndex = responseText.lastIndexOf('}');
                let csvPart = responseText;
                let jsonPart = null;

                if (lastBraceIndex !== -1) {
                    const potentialJsonStart = responseText.lastIndexOf('{', lastBraceIndex);
                    if (potentialJsonStart !== -1) {
                        let jsonStart = potentialJsonStart;
                        while(jsonStart > 0 && responseText[jsonStart-1] !== '\n') {
                            jsonStart--;
                        }

                        csvPart = responseText.substring(0, jsonStart).trim();
                        jsonPart = responseText.substring(jsonStart).trim();
                    }
                }
                
                jsonPart = jsonPart ? jsonPart.replace(/^```(json)?\s*/i, '').replace(/\s*```$/, '') : null;


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
                
                // 3. Verifica a Qualidade da Extração
                if (dailyRecords.length === 0) {
                     throw new Error("A IA não conseguiu extrair registros válidos. Verifique o documento de origem.");
                }

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
                        total_dias_registrados: dailyRecords.length,
                        resumo_parcial: 'Extração de dados brutos concluída.',
                    });
                }
                
            } catch (err) {
                console.error(`Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ 
                    arquivo_original: file.originalname,
                    erro: `Falha na extração de dados: ${err.message}.`
                });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        const orderedKeys = ['nome_colaborador', 'data_registro', 'entrada_1', 'saida_1', 'entrada_2', 'saida_2', 'total_horas_trabalhadas', 'horas_extra_diarias', 'horas_falta_diarias', 'resumo_executivo_mensal', 'arquivo_original'];
        orderedKeys.forEach(key => allDetailedKeys.add(key));
        
        const sessionId = Date.now().toString();
        
        if (allResultsForExcel.length === 0) {
             throw new Error("Nenhum registro válido foi extraído de todos os arquivos. A planilha não será criada.");
        }

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
        return res.status(500).send({ error: `Erro interno do servidor: ${error.message}` });
    } finally {
        await Promise.all(fileCleanupPromises).catch(e => console.error("Erro ao limpar arquivos temporários:", e));
    }
});


// Endpoint para Download do Excel
app.get('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session || !session.data) {
        // Se a sessão não existe mais, ela pode ter sido limpa por uma tentativa anterior.
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada. Tente fazer o upload novamente.' });
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
            if (err) {
                 // Erro no envio, mas o arquivo temporário existe
                 console.error("Erro ao enviar o Excel:", err);
            }
            
            // Limpa o arquivo temporário e a sessão APÓS a tentativa de download
            if (excelCreated && fs.existsSync(excelPath)) {
                 await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel:", e));
            }
            delete sessionData[sessionId]; 
        });
    } catch (error) {
        console.error('Erro ao gerar Excel:', error);
        res.status(500).send({ error: `Falha ao gerar o arquivo Excel: ${error.message}` });
        
        // Limpeza adicional em caso de falha de criação
        if (excelCreated && fs.existsSync(excelPath)) {
            await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel após falha de criação:", e));
        }
        delete sessionData[sessionId]; 
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