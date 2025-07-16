import { STATUS_MAP, APP_TEXTS } from './config.js';

/**
 * Lê o arquivo Excel e o converte para um array de dados.
 * @param {File} file - O arquivo carregado pelo usuário.
 * @returns {Promise<Array>} - Uma promessa que resolve com os dados do Excel.
 */
export async function readFileData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                resolve(jsonData);
            } catch (error) {
                reject(new Error(APP_TEXTS.fileError));
            }
        };
        reader.onerror = () => reject(new Error(APP_TEXTS.fileError));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Analisa os dados brutos e calcula os KPIs e estatísticas.
 * @param {Array} jsonData - Dados extraídos do Excel.
 * @returns {Object} - Um objeto com todos os dados processados.
 */
export function analyzeOeeData(jsonData) {
    if (!jsonData || jsonData.length < 2) {
        throw new Error(APP_TEXTS.invalidDataError);
    }

    const header = jsonData[0];
    const rows = jsonData.slice(1).filter(row => row.length > 0 && row[0]); // Filtra linhas vazias
    const circuitHeader = header[0];

    if (!circuitHeader) {
        throw new Error(APP_TEXTS.columnError);
    }
    
    // Converte para objetos para facilitar a leitura do código
    const fullData = rows.map(row =>
        header.reduce((obj, h, i) => ({ ...obj, [h]: row[i] }), {})
    );

    const analysis = {
        totalCircuits: 0,
        statusTotals: { UP: 0, SD: 0, PP: 0 },
        circuitStats: {}
    };

    fullData.forEach(row => {
        const circuitId = row[circuitHeader];
        if (!circuitId) return;

        if (!analysis.circuitStats[circuitId]) {
            analysis.circuitStats[circuitId] = { UP: 0, SD: 0, PP: 0, totalDays: 0 };
            analysis.totalCircuits++;
        }

        header.slice(1).forEach(dateCol => {
            const status = row[dateCol];
            const statusKey = STATUS_MAP[status];
            if (statusKey) {
                analysis.statusTotals[statusKey]++;
                analysis.circuitStats[circuitId][statusKey]++;
                analysis.circuitStats[circuitId].totalDays++;
            }
        });
    });

    return { header, fullData, ...analysis };
}