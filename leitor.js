// leitor.js (Versão Definitiva: Multimodal, Dinâmico, Excel Vertical e Sem Nomes Duplicados + Agrupamento de Chaves)

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

// Armazenamento em memória para dados de sessão (Adicionamos uniqueKeys)
// sessionData: { sessionId: { data: [/* ... */], uniqueKeys: [/* ... */] } }
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
 * **NOVA FUNÇÃO:** Agrupa chaves sequenciais/numéricas (ex: dia1, dia2, item_1 -> dia, item)
 * @param {Array<string>} keys - Lista de chaves detalhadas extraídas da IA.
 * @returns {Array<string>} Lista de chaves agrupadas/resumidas.
 */
function groupKeys(keys) {
    const groupedKeys = new Set();
    
    // Regex para identificar padrões numéricos/sequenciais no final da string
    // Ex: _1, _2, _A, _B, 1, 2 (se não estiver precedido por letra)
    const regex = /(_\d+$)|(_[a-z]$)|(\d+$)/i; 
    
    keys.forEach(key => {
        // Chaves que são exceções ou essenciais
        if (key === 'arquivo_original') {
            groupedKeys.add(key);
            return;
        }

        const match = key.match(regex);

        if (match) {
            // Remove o sufixo numérico/sequencial para criar a chave agrupada
            const groupedKey = key.replace(regex, '');
            
            // Adiciona a chave agrupada (ex: 'dia', 'item_lista')
            if (groupedKey) {
                 groupedKeys.add(groupedKey);
            } else {
                 // Caso a chave seja apenas um número ou um caracter (o que é improvável, mas seguro)
                 groupedKeys.add(key);
            }
        } else {
            // Se não houver padrão (ex: 'nome_cliente'), a chave é adicionada como está
            groupedKeys.add(key);
        }
    });

    // Filtra as chaves agrupadas para remover duplicatas e retornar a lista final
    return Array.from(groupedKeys);
}

/**
 * Cria o arquivo Excel com abas individuais e formato vertical (Key|Value), tratando nomes duplicados.
 * @param {Array<Object>} allExtractedData - Dados extraídos de todos os documentos (DETALHADOS).
 * @param {string} outputPath - Caminho para salvar o arquivo.
 * @param {Array<string>} [selectedKeys=[]] - Chaves AGRUPADAS selecionadas pelo usuário.
 */
async function createExcelFile(allExtractedData, outputPath, selectedKeys = []) {
    const workbook = new ExcelJS.Workbook();
    
    if (allExtractedData.length === 0) return;

    const usedSheetNames = new Set();
    
    // --- Lógica para expandir as chaves agrupadas para chaves detalhadas (filtros reais) ---
    const detailedKeysToInclude = new Set();
    const regex = /(_\d+$)|(_[a-z]$)|(\d+$)/i;

    // Se nenhuma chave foi selecionada (erro ou teste), inclua tudo (segurança)
    if (selectedKeys.length === 0) {
        allExtractedData.forEach(data => Object.keys(data).forEach(key => detailedKeysToInclude.add(key)));
    } else {
        // 1. Processa os grupos de chaves selecionados
        for (const groupKey of selectedKeys) {
            
            // Se a chave não parece agrupável ou é essencial (resumo, nome original), adiciona diretamente
            if (groupKey === 'arquivo_original' || groupKey === 'resumo_executivo' || !regex.test(groupKey)) {
                detailedKeysToInclude.add(groupKey);
            } 
            
            // 2. Se for um grupo, encontra todos os membros detalhados correspondentes
            // Devemos iterar sobre TODAS as chaves detalhadas de TODOS os documentos para garantir que pegamos todos os padrões.
            allExtractedData.forEach(data => {
                Object.keys(data).forEach(detailedKey => {
                    const baseName = detailedKey.replace(regex, '');
                    
                    // Se o nome base (ex: 'dia') corresponde ao grupo selecionado (ex: 'dia')
                    // Ou se a chave detalhada for igual ao nome do grupo (ex: 'valor_total' é igual a 'valor_total')
                    if (baseName === groupKey || detailedKey === groupKey) {
                        detailedKeysToInclude.add(detailedKey);
                    }
                });
            });
        }
    }
    
    // 1. Loop para criar uma aba para CADA DOCUMENTO
    for (let i = 0; i < allExtractedData.length; i++) {
        const data = allExtractedData[i];
        
        // Define o nome da aba (máximo de 31 caracteres, sanitizando e evitando duplicatas)
        const unsafeName = data.arquivo_original.replace(/\.[^/.]+$/, "");
        let baseName = unsafeName.substring(0, 28).replace(/[\[\]\*\:\/\?\\\,]/g, ' ');
        baseName = baseName.trim().replace(/\.$/, ''); 
        
        // Lógica de desambiguação: se o nome já foi usado, adiciona um contador
        let worksheetName = baseName;
        let counter = 1;
        while (usedSheetNames.has(worksheetName)) {
            worksheetName = `${baseName.substring(0, 25)} (${counter})`; 
            counter++;
        }
        usedSheetNames.add(worksheetName); // Marca o nome como usado

        const worksheet = workbook.addWorksheet(worksheetName || `Documento ${i + 1}`);
        
        // 2. Configura colunas no formato VERTICAL (Propriedade | Valor)
        worksheet.columns = [
            { header: 'Campo Extraído', key: 'key', width: 35 },
            { header: 'Valor', key: 'value', width: 60 }
        ];

        // 3. Mapeia o objeto JSON dinâmico para LINHAS VERTICAIS, aplicando FILTRO
        const verticalRows = Object.entries(data)
            .filter(([key, value]) => detailedKeysToInclude.has(key)) // FILTRO AQUI
            .map(([key, value]) => ({
                key: key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()), 
                value: value
            }));
        
        worksheet.addRows(verticalRows);

        // 4. Aplica Formatação (Estilo Profissional)
        
        // Formatação do Cabeçalho (Estilo Profissional)
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

                // Aplica formato de moeda para colunas com 'valor' ou 'total'
                if (cellKey.includes('valor') || cellKey.includes('total')) {
                    cellValue.numFmt = 'R$ #,##0.00'; 
                }
                
                // Quebra o texto na coluna de Valor (ex: Resumo)
                if (typeof cellValue.value === 'string' && cellValue.value.length > 50) {
                     cellValue.alignment = { wrapText: true, vertical: 'top' };
                }
            }
        });
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
    // Conjunto para coletar todas as chaves dinâmicas únicas (DETALHADAS)
    const detailedKeys = new Set();
    
    // Prompt Agnostico e Dinâmico (Lê PDF ou Imagem)
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
                
                // 1. Obtém as chaves dinâmicas do JSON e armazena as detalhadas
                const keys = Object.keys(dynamicData);
                keys.forEach(key => detailedKeys.add(key));

                // 2. Mapeamento para o Front-end (Visualização Genérica):
                const clientResult = {
                    arquivo_original: file.originalname,
                    // Pega as 4 primeiras chaves e valores para a visualização na tabela
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
                allResultsForClient.push({ erro: `Falha na API: ${err.message.substring(0, 100)}` });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        // --- Aplica o agrupamento de chaves para o Frontend ---
        const finalUniqueKeys = groupKeys(Array.from(detailedKeys));
        
        const sessionId = Date.now().toString();
        // Armazena dados originais (detalhados) e a lista de chaves AGRUPADAS na sessão
        sessionData[sessionId] = {
            data: allResultsForExcel, 
            uniqueKeys: finalUniqueKeys 
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

// --- Endpoint para buscar as chaves únicas (AGRUPADAS) ---
app.get('/fields/:sessionId', (req, res) => {
    const { sessionId } = req.params;
    const session = sessionData[sessionId];

    if (!session) {
        return res.status(404).json({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    // Retorna a lista de chaves únicas AGRUPADAS para o frontend
    res.json({ uniqueKeys: session.uniqueKeys });
});

// --- Endpoint para Download do Excel (recebe campos AGRUPADOS) ---
app.post('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const { selectedFields } = req.body; // Recebe os campos AGRUPADOS selecionados
    const session = sessionData[sessionId];

    if (!session || !session.data) {
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada.' });
    }
    
    const data = session.data;

    const excelFileName = `extracao_gemini_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        // Passa os campos AGRUPADOS selecionados. A função createExcelFile os expandirá.
        await createExcelFile(data, excelPath, selectedFields);

        res.download(excelPath, excelFileName, async (err) => {
            if (err) {
                console.error("Erro ao enviar o Excel:", err);
            }
            await fs.promises.unlink(excelPath).catch(e => console.error("Erro ao limpar arquivo Excel:", e));
            delete sessionData[sessionId]; // Limpa a sessão após o download
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