// leitor.js (Versão Definitiva: Multimodal, Dinâmico, Excel Vertical, Sem Nomes Duplicados, Agrupamento Otimizado E Seleção por Arquivo)

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

// Middleware para receber JSON (necessário para o novo endpoint de download)
app.use(express.json()); 

// Correção para obter __dirname em módulos ES
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
// sessionData: { sessionId: { 
//    data: [{... doc 1 ...}, {... doc 2 ...}], 
//    fieldLists: [{ filename: "doc1.pdf", keys: ["campo1", "campo2_agrupado"] }, ...] 
// }}
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
 * Lógica de Backoff Exponencial para lidar com o erro 429.
 */
async function callApiWithRetry(apiCall, maxRetries = 5) {
    let delay = 2; 
    for (let attempt = 0; attempt < maxRetries; attempt++) {
        try {
            return await apiCall();
        } catch (error) {
            // Condição aprimorada para erros 429
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
 * **MELHORIA NO AGRUPAMENTO:** Agrupa chaves sequenciais/numéricas (ex: info_item_1, dia_1, total_2 -> info_item, dia, total).
 * @param {Array<string>} keys - Lista de chaves detalhadas extraídas da IA.
 * @returns {Array<string>} Lista de chaves agrupadas/resumidas.
 */
function groupKeys(keys) {
    const groupedKeys = new Set();
    
    // Regex aprimorado: 
    // 1. Padrões de 3 ou mais palavras com número no final (ex: info_item_1)
    // 2. Padrões de 2 palavras com número no final (ex: dia_1)
    // 3. Padrões com underline + número no final (ex: _1)
    // 4. Padrões com underline + letra no final (ex: _a)
    // 5. Padrões com número no final (ex: total1)
    const regex = /(\w+_\w+_\d+$)|(\w+_\d+$)|(_\d+$)|(_[a-z]$)|(\d+$)/i; 
    
    keys.forEach(key => {
        // Chaves que são exceções ou essenciais
        if (key === 'arquivo_original' || key === 'resumo_executivo') {
            groupedKeys.add(key);
            return;
        }

        const match = key.match(regex);

        if (match) {
            // Remove o sufixo numérico/sequencial para criar a chave agrupada
            const groupedKey = key.replace(regex, '').replace(/_$/, ''); // Remove _ residual
            
            // Adiciona a chave agrupada (ex: 'dia', 'item_lista', 'total')
            if (groupedKey) {
                 groupedKeys.add(groupedKey);
            } else {
                 groupedKeys.add(key); // Caso falhe, usa a chave original
            }
        } else {
            // Se não houver padrão (ex: 'nome_cliente'), a chave é adicionada como está
            groupedKeys.add(key);
        }
    });

    return Array.from(groupedKeys).sort();
}


/**
 * Cria o arquivo Excel com abas individuais e formato vertical (Key|Value), tratando nomes duplicados.
 * @param {Array<Object>} allExtractedData - Dados extraídos de todos os documentos (DETALHADOS).
 * @param {string} outputPath - Caminho para salvar o arquivo.
 * @param {Object} selectedFieldsMap - **NOVO:** Mapa de { filename: [selectedKeysAgrupadas] }
 */
async function createExcelFile(allExtractedData, outputPath, selectedFieldsMap) {
    const workbook = new ExcelJS.Workbook();
    
    if (allExtractedData.length === 0) return;

    const usedSheetNames = new Set();
    
    const regex = /(\w+_\w+_\d+$)|(\w+_\d+$)|(_\d+$)|(_[a-z]$)|(\d+$)/i; 
    
    // 1. Loop para criar uma aba para CADA DOCUMENTO
    for (let i = 0; i < allExtractedData.length; i++) {
        const data = allExtractedData[i];
        const filename = data.arquivo_original;
        const selectedKeysAgrupadas = selectedFieldsMap[filename] || [];
        
        // --- Lógica para expandir as chaves agrupadas para chaves detalhadas (filtros reais) ---
        const detailedKeysToInclude = new Set();
        
        // Se NENHUMA chave foi selecionada para este arquivo, pulamos o filtro
        if (selectedKeysAgrupadas.length === 0) {
            // Não faz sentido gerar uma aba vazia, mas se for o caso, pode-se decidir pular
            // Por segurança, vamos incluir as chaves padrão se nada foi selecionado.
            detailedKeysToInclude.add('arquivo_original');
            detailedKeysToInclude.add('resumo_executivo');
        } else {
            // 2. Expande as chaves agrupadas para as chaves detalhadas reais
            for (const groupKey of selectedKeysAgrupadas) {
                
                // Se a chave não parece agrupável ou é essencial, adiciona diretamente
                if (groupKey === 'arquivo_original' || groupKey === 'resumo_executivo' || !regex.test(groupKey)) {
                    detailedKeysToInclude.add(groupKey);
                } 
                
                // 3. Se for um grupo, encontra todos os membros detalhados correspondentes NESTE documento
                Object.keys(data).forEach(detailedKey => {
                    const baseName = detailedKey.replace(regex, '').replace(/_$/, ''); 
                    
                    // Se o nome base (ex: 'dia') corresponde ao grupo selecionado (ex: 'dia')
                    // Ou se a chave detalhada for igual ao nome do grupo (match exato sem sufixo)
                    if (baseName === groupKey || detailedKey === groupKey) {
                        detailedKeysToInclude.add(detailedKey);
                    }
                });
            }
        }
        
        // Define o nome da aba (máximo de 31 caracteres, sanitizando e evitando duplicatas)
        const unsafeName = filename.replace(/\.[^/.]+$/, "");
        let baseName = unsafeName.substring(0, 28).replace(/[\[\]\*\:\/\?\\\,]/g, ' ');
        baseName = baseName.trim().replace(/\.$/, ''); 
        
        // Lógica de desambiguação: se o nome já foi usado, adiciona um contador
        let worksheetName = baseName;
        let counter = 1;
        while (usedSheetNames.has(worksheetName)) {
            worksheetName = `${baseName.substring(0, 25)} (${counter})`; 
            counter++;
        }
        usedSheetNames.add(worksheetName); 

        const worksheet = workbook.addWorksheet(worksheetName || `Documento ${i + 1}`);
        
        // 4. Configura colunas no formato VERTICAL (Propriedade | Valor)
        worksheet.columns = [
            { header: 'Campo Extraído', key: 'key', width: 35 },
            { header: 'Valor', key: 'value', width: 60 }
        ];

        // 5. Mapeia o objeto JSON dinâmico para LINHAS VERTICAIS, aplicando FILTRO
        const verticalRows = Object.entries(data)
            .filter(([key, value]) => detailedKeysToInclude.has(key)) 
            .map(([key, value]) => ({
                key: key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
                value: value
            }));
        
        if (verticalRows.length > 0) {
            worksheet.addRows(verticalRows);

            // 6. Aplica Formatação (Estilo Profissional) - Somente se houver linhas para formatar
            
            // Formatação do Cabeçalho
            worksheet.getRow(1).eachCell(cell => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } }; 
                cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 12 };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Formatação de Valores e Quebra de Texto
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { 
                    const cellKey = row.getCell(1).value.toString().toLowerCase();
                    const cellValue = row.getCell(2);

                    if (cellKey.includes('valor') || cellKey.includes('total')) {
                        cellValue.numFmt = 'R$ #,##0.00'; 
                    }
                    
                    if (typeof cellValue.value === 'string' && cellValue.value.length > 50) {
                        cellValue.alignment = { wrapText: true, vertical: 'top' };
                    }
                }
            });
        }
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
    // Novo: Array para armazenar as chaves agrupadas POR ARQUIVO
    const fieldLists = [];
    
    const prompt = `
        Você é um assistente especialista em extração de dados estruturados. Sua tarefa é analisar o documento anexado (que pode ser qualquer tipo de PDF ou IMAGEM) e extrair **TODAS** as informações relevantes.
        ... (REGRAS CRÍTICAS para o JSON)
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
                
                // 1. Obtém as chaves DETALHADAS
                const keys = Object.keys(dynamicData);
                
                // 2. Agrupa as chaves DETALHADAS para o Frontend
                const groupedKeys = groupKeys(keys);

                // 3. Armazena a lista de campos agrupados para este arquivo
                fieldLists.push({
                    filename: file.originalname,
                    keys: groupedKeys
                });

                // 4. Mapeamento para o Front-end (Visualização Genérica):
                const clientResult = {
                    arquivo_original: file.originalname,
                    chave1: keys.length > 0 ? keys[0] : 'N/A',
                    valor1: keys.length > 0 ? dynamicData[keys[0]] : 'N/A',
                    chave2: keys.length > 1 ? keys[1] : 'N/A',
                    valor2: keys.length > 1 ? dynamicData[keys[1]] : 'N/A',
                    resumo: dynamicData.resumo_executivo || 'Resumo não fornecido.',
                };

                allResultsForClient.push(clientResult);
                
                // Adiciona o nome original do arquivo ao objeto dinâmico para rastreamento no Excel
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
        
        const sessionId = Date.now().toString();
        // Armazena dados originais (detalhados) e a lista de campos agrupados POR ARQUIVO
        sessionData[sessionId] = {
            data: allResultsForExcel, 
            fieldLists: fieldLists 
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

// --- Endpoint para buscar as chaves AGRUPADAS POR ARQUIVO ---
app.get('/fields/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session) {
        return res.status(404).json({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    // Retorna a lista de chaves agrupadas POR ARQUIVO para o frontend
    res.json({ fieldLists: session.fieldLists });
});

// --- Endpoint para Download do Excel (recebe mapa de campos selecionados) ---
app.post('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const { selectedFieldsMap } = req.body; // Recebe o mapa de { filename: [selectedKeysAgrupadas] }
    const session = sessionData[sessionId];

    if (!session || !session.data) {
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    const data = session.data;

    const excelFileName = `extracao_gemini_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        // Passa o mapa de seleção para createExcelFile
        await createExcelFile(data, excelPath, selectedFieldsMap);

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

// --- 8. Servir front-end e iniciar servidor ---

app.use(express.static(PUBLIC_DIR)); 

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, 'index.html')); 
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});