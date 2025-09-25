// leitor.js (VERSÃO FINAL COM CÁLCULOS NO EXCEL)

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
 * REFORÇADO: Converte a string CSV da IA em objetos JavaScript.
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
        
        if (values.length < headers.length) {
            continue; 
        }

        const record = { arquivo_original: filename };
        let foundValidData = false;

        headers.forEach((header, index) => {
            if (header && values[index] !== undefined) {
                let value = values[index].trim();
                
                // CRÍTICO: Não fazemos conversão de horas, mantendo apenas HH:MM strings
                // O Excel fará o cálculo com as fórmulas.
                
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
 * Cria o arquivo Excel, injetando as FÓRMULAS para os cálculos.
 */
async function createExcelFile(allExtractedData, outputPath, allDetailedKeys) {
    const workbook = new ExcelJS.Workbook();
    const standardDailyHours = 8; // Jornada padrão de 8 horas para cálculo de extra/falta
    
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

    // Chaves FINAIS que o Excel terá (RAW + CALCULATED)
    const finalKeys = ['nome_colaborador', 'data_registro', 'entrada_1', 'saida_1', 'entrada_2', 'saida_2', 'total_horas_trabalhadas', 'horas_extra_diarias', 'horas_falta_diarias', 'arquivo_original'];
    
    // Mapeamento de Chave para Letra de Coluna (Ex: 'nome_colaborador' -> 'A', 'total_horas_trabalhadas' -> 'G')
    const keyToCol = {};
    finalKeys.forEach((key, index) => {
        // Assume até 26 colunas (A-Z) para simplificação.
        keyToCol[key] = String.fromCharCode('A'.charCodeAt(0) + index);
    });
    
    const defineColumns = (keys) => {
        return keys.map(key => ({
            header: key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()),
            key: key,
            width: key.includes('horas') || key.includes('total') ? 18 : 15, 
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
            // Fórmulas ou resultados de soma devem ser formatados como números com 2 decimais
            cell.numFmt = '0.00'; 
        } else if (headerKey.includes('data') || headerKey.includes('entrada') || headerKey.includes('saida')) {
            // Dados de hora/data centralizados
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
            const rowNum = index + 2; 

            // --- INJEÇÃO DE FÓRMULAS CRÍTICAS ---
            const c_e1 = keyToCol['entrada_1'];
            const c_s1 = keyToCol['saida_1'];
            const c_e2 = keyToCol['entrada_2'];
            const c_s2 = keyToCol['saida_2'];
            
            const totalFormula = `((IFERROR(TIMEVALUE(${c_s1}${rowNum}),0) - IFERROR(TIMEVALUE(${c_e1}${rowNum}),0) + IFERROR(TIMEVALUE(${c_s2}${rowNum}),0) - IFERROR(TIMEVALUE(${c_e2}${rowNum}),0)) * 24)`;
            
            const extraFormula = `MAX(0, ${totalFormula} - ${standardDailyHours})`;
            const faltaFormula = `MAX(0, ${standardDailyHours} - ${totalFormula})`;
            
            // Adiciona as FÓRMULAS ao registro que será inserido na linha
            record['total_horas_trabalhadas'] = { formula: totalFormula };
            record['horas_extra_diarias'] = { formula: extraFormula };
            record['horas_falta_diarias'] = { formula: faltaFormula };
            // -------------------------------------

            const row = worksheet.addRow(record);
            row.height = 18;
            
            const fillColor = index % 2 === 0 ? 'F0F0F0' : 'FFFFFF';
            row.eachCell((cell, colNumber) => {
                const headerKey = worksheet.getColumn(colNumber).key.toLowerCase();
                formatDataCell(cell, headerKey, false);
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            });
        });

        // --- Linha de Resumo/Soma ---
        const resumoRow = worksheet.addRow({});
        resumoRow.height = 25;
        
        const firstDataRow = 2; 
        const lastDataRow = worksheet.lastRow.number - 1; 
        
        const totalCol = keyToCol['total_horas_trabalhadas'].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        const extraCol = keyToCol['horas_extra_diarias'].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        const faltaCol = keyToCol['horas_falta_diarias'].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        
        if (lastDataRow >= firstDataRow) {
             resumoRow.getCell(1).value = 'RESUMO MENSAL / SOMAS:';
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
    const allDetailedKeys = new Set(); 
    
    // PROMPT REVISADO: Pede APENAS os campos brutos para que o Excel calcule o restante.
    const prompt = `
        Você é um assistente especialista em extração de registros de ponto.
        Sua única e crítica tarefa é analisar o documento anexado (cartão de ponto, espelho ou folha de registro) e extrair os registros diários de forma estruturada.

        REGRAS CRÍTICAS DE EXTRAÇÃO:
        1. **GARANTIA DE COMPLETUDE (CRÍTICO):** Extraia **TODAS** as linhas de registro.
        2. **FORMATO DE SAÍDA (CRÍTICO):** Retorne estritamente o seguinte formato: Tabela **CSV** separada por **ponto e vírgula (;)**.
        3. **CAMPOS OBRIGATÓRIOS (APENAS ESTES):**
           Cabeçalho: Nome_Colaborador;Data_Registro;Entrada_1;Saida_1;Entrada_2;Saida_2
           Valores Ausentes: Use **N/A** para horários e datas que não puderem ser extraídos.
           Formato de Horas/Minutos: Use **HH:MM** (Ex: 08:30) para todas as entradas e saídas.

        Retorne APENAS a tabela CSV (texto puro), **não** adicione resumo JSON, explicações ou blocos de código markdown (\`\`\`).
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
                let csvPart = response.text.trim();
                
                // Processa os Registros Diários (CSV)
                const dailyRecords = parseCsvRecords(csvPart, file.originalname);
                
                if (dailyRecords.length === 0) {
                     throw new Error("A IA não conseguiu extrair registros válidos. Verifique o documento de origem.");
                }

                dailyRecords.forEach(record => {
                    Object.keys(record).forEach(key => allDetailedKeys.add(key));
                    allResultsForExcel.push(record);
                });

                // Prepara dados para o Front-end
                if (dailyRecords.length > 0) {
                    const firstRecord = dailyRecords[0];
                    
                    allResultsForClient.push({
                        arquivo_original: file.originalname,
                        nome_colaborador: firstRecord.nome_colaborador || 'N/A',
                        total_dias_registrados: dailyRecords.length,
                        resumo_parcial: 'Extração de dados brutos concluída. Cálculos de horas serão feitos no Excel.',
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
        
        // Define as chaves finais para a sessão
        ['nome_colaborador', 'data_registro', 'entrada_1', 'saida_1', 'entrada_2', 'saida_2', 'total_horas_trabalhadas', 'horas_extra_diarias', 'horas_falta_diarias', 'arquivo_original'].forEach(key => allDetailedKeys.add(key));
        
        const sessionId = Date.now().toString();
        
        if (allResultsForExcel.length === 0) {
             throw new Error("Nenhum registro válido foi extraído de todos os arquivos. A planilha não será criada.");
        }

        sessionData[sessionId] = {
            data: allResultsForExcel, 
            allDetailedKeys: Array.from(allDetailedKeys).sort(),
        };

        return res.json({ 
            results: allResultsForClient,
            sessionId: sessionId,
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
                 console.error("Erro ao enviar o Excel:", err);
            }
            
            if (excelCreated && fs.existsSync(excelPath)) {
                 await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel:", e));
            }
            delete sessionData[sessionId]; 
        });
    } catch (error) {
        console.error('Erro ao gerar Excel:', error);
        res.status(500).send({ error: `Falha ao gerar o arquivo Excel: ${error.message}` });
        
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