// leitor.js (Servidor Node.js Final - Corrigido para Render/Deploy)

import express from 'express';
import multer from 'multer';
import { GoogleGenAI } from '@google/genai';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import random from 'random'; 
import 'dotenv/config'; 
import { fileURLToPath } from 'url'; // Necessário para simular __dirname

// --- Configurações de Ambiente ---
const app = express();
// O Render define a porta para 10000, mas usamos a variável PORT para flexibilidade
const PORT = process.env.PORT || 3000; 

// 1. Correção para obter __dirname em módulos ES
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR = path.join(__dirname, 'temp');

// Garante que os diretórios existem
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

// Inicializa a API Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Configuração do Multer (para upload de arquivos)
const upload = multer({ dest: UPLOAD_DIR });

// Armazenamento em memória para dados de sessão
const sessionData = {};

// --- Funções Auxiliares de IA e Tratamento de Erro ---

/**
 * Converte o arquivo local em Base64 para a API Gemini.
 */
function fileToGenerativePart(filePath, mimeType) {
    return {
        inlineData: {
            data: Buffer.from(fs.readFileSync(filePath)).toString("base64"),
            mimeType
        },
    };
}

/**
 * Lógica de Backoff Exponencial para lidar com o erro 429 (Too Many Requests).
 */
async function callApiWithRetry(apiCall, maxRetries = 5) {
    let delay = 2; 
    for (let attempt = 0; attempt < maxRetries; attempt++) {
        try {
            return await apiCall();
        } catch (error) {
            // Verifica se o erro é 429 (ResourceExhaustedError)
            if (error.status === 429 || (error.message && error.message.includes('Resource has been exhausted'))) {
                if (attempt === maxRetries - 1) {
                    throw new Error('Limite de taxa excedido (429) após múltiplas tentativas. Tente novamente mais tarde.');
                }
                
                // Backoff Exponencial com Jitter
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
 * Cria a planilha Excel a partir dos dados extraídos.
 */
async function createExcelFile(allExtractedData, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Extraídos Gemini');

    if (allExtractedData.length === 0) return;

    const headers = [
        'Arquivo', 'Tipo_Documento', 'Resumo_Executivo', 'Remetente', 'Destinatario', 
        'numero_documento', 'data_emissao', 'valor_total', 'assunto'
    ];
    worksheet.columns = headers.map(header => ({ header, key: header, width: 20 }));

    worksheet.addRows(allExtractedData.map(d => ({
        Arquivo: d.arquivo_original,
        Tipo_Documento: d.tipo_documento,
        Resumo_Executivo: d.resumo_executivo,
        Remetente: d.dados_chave.remetente,
        Destinatario: d.dados_chave.destinatario,
        numero_documento: d.dados_chave.numero_documento,
        data_emissao: d.dados_chave.data_emissao,
        valor_total: d.dados_chave.valor_total,
        assunto: d.dados_chave.assunto
    })));

    await workbook.xlsx.writeFile(outputPath);
}

// --- 4. Endpoint Principal de Upload e Processamento ---

app.post('/upload', upload.array('pdfs'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send({ error: 'Nenhum arquivo PDF enviado.' });
    }

    const fileCleanupPromises = [];
    const allResultsForClient = [];
    const allResultsForExcel = [];
    
    const prompt = `
        Você é um assistente especialista em análise de documentos brasileiros. Leia o documento (pode ser Nota Fiscal, Folha de Ponto, Recibo, etc.) e extraia os dados mais relevantes.
        Sua resposta deve ser APENAS o objeto JSON, seguindo estritamente o 'responseSchema'.
        
        REGRAS DE EXTRAÇÃO:
        1. Para valores monetários, use o formato de número (ex: 123.45).
        2. Para datas, use o formato 'DD/MM/AAAA'.
        3. Se um campo não for encontrado, use 'N/A' (para texto) ou 0.00 (para valores).
    `;

    const responseSchema = {
        type: "object",
        properties: {
            arquivo_original: { type: "string" },
            tipo_documento: { type: "string", description: "O tipo de documento detectado." },
            resumo_executivo: { type: "string", description: "Um resumo conciso do conteúdo." },
            dados_chave: { 
                type: "object", 
                properties: {
                    remetente: { type: "string", description: "Nome/Razão Social do Emissor principal." },
                    destinatario: { type: "string", description: "Nome/Razão Social do Receptor principal." },
                    numero_documento: { type: "string", description: "Número de identificação único." },
                    data_emissao: { type: "string", description: "Data de emissão (DD/MM/AAAA)." },
                    valor_total: { type: "number", description: "Valor Total da transação/documento." },
                    assunto: { type: "string", description: "Breve descrição do serviço ou produto." }
                },
                required: ["remetente", "destinatario", "numero_documento", "data_emissao", "valor_total", "assunto"],
                additionalProperties: true 
            }
        },
        required: ["arquivo_original", "tipo_documento", "resumo_executivo", "dados_chave"]
    };

    try {
        for (const file of req.files) {
            console.log(`[PROCESSANDO] ${file.originalname}`);
            
            const filePart = fileToGenerativePart(file.path, file.mimetype);
            
            const apiCall = () => ai.models.generateContent({
                model: 'gemini-1.5-flash', 
                contents: [filePart, { text: prompt }],
                config: {
                    responseMimeType: "application/json",
                    responseSchema: responseSchema,
                    temperature: 0.1
                }
            });

            try {
                const response = await callApiWithRetry(apiCall);
                const fullJsonResponse = JSON.parse(response.text);
                
                const clientResult = {
                    ...fullJsonResponse.dados_chave,
                    remetente: fullJsonResponse.dados_chave.remetente || 'N/A',
                    destinatario: fullJsonResponse.dados_chave.destinatario || 'N/A',
                };

                allResultsForClient.push(clientResult);
                
                allResultsForExcel.push({
                    arquivo_original: file.originalname,
                    tipo_documento: fullJsonResponse.tipo_documento,
                    resumo_executivo: fullJsonResponse.resumo_executivo,
                    dados_chave: fullJsonResponse.dados_chave
                });
                
                await new Promise(resolve => setTimeout(resolve, 5000)); 

            } catch (err) {
                console.error(`Erro ao processar ${file.originalname}: ${err.message}`);
                allResultsForClient.push({ erro: `Falha na API: ${err.message.substring(0, 100)}` });
            } finally {
                fileCleanupPromises.push(fs.promises.unlink(file.path));
            }
        }
        
        const sessionId = Date.now().toString();
        sessionData[sessionId] = allResultsForExcel;

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

// --- 5. Endpoint para Download do Excel ---

app.get('/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const data = sessionData[sessionId];

    if (!data) {
        return res.status(404).send({ error: 'Sessão de dados não encontrada ou expirada.' });
    }

    const excelFileName = `extracao_gemini_${sessionId}.xlsx`;
    const excelPath = path.join(TEMP_DIR, excelFileName);

    try {
        await createExcelFile(data, excelPath);

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

// --- 6. Servir front-end e iniciar servidor (CORREÇÃO FINAL DE PATH) ---

// Serve arquivos estáticos a partir do diretório raiz onde leitor.js está.
app.use(express.static(__dirname)); 

// Rota principal (servir o index.html)
app.get('/', (req, res) => {
    // Usa __dirname para garantir que o arquivo seja encontrado em qualquer ambiente.
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
    console.log(`✅ Servidor rodando na porta ${PORT}`);
});