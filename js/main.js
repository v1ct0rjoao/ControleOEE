import { DOM_SELECTORS, APP_TEXTS, STATUS_COLORS } from './config.js';
import { readFileData, analyzeOeeData } from './data.js';
import { ui } from './ui.js';

// Mapeamento dos elementos do DOM para fácil acesso
const dom = Object.fromEntries(
    Object.entries(DOM_SELECTORS).map(([key, selector]) => [key, document.querySelector(selector)])
);

// Estado da aplicação
const appState = {
    dashboardVisible: false
};

/**
 * Mostra a seção do dashboard e oculta a página inicial.
 */
function showDashboardView() {
    dom.landingPage.classList.add('hidden');
    dom.dashboardView.classList.remove('hidden');
    appState.dashboardVisible = true;
}

/**
 * Função principal que é acionada ao selecionar um arquivo.
 * @param {Event} event - O evento de 'change' do input de arquivo.
 */
async function handleFileSelect(event) {
    const file = event.target.files?.[0]; // Use optional chaining
    if (!file || !appState.dashboardVisible) return;

    ui.showLoading(dom.loadingOverlay);
    dom.fileName.textContent = file.name;
    dom.alertContainer.innerHTML = ''; // Limpa alertas antigos

    try {
        const rawData = await readFileData(file);
        const processedData = analyzeOeeData(rawData);

        // Renderiza todos os componentes da UI com os dados processados
        ui.renderKPIs(dom.kpiGrid, processedData, STATUS_COLORS); // Passa STATUS_COLORS
        ui.renderCharts(dom.chartsGrid, processedData, STATUS_COLORS); // Passa STATUS_COLORS
        ui.renderTable(dom.tableWrapper, processedData);

        // Adiciona o listener de busca APÓS a tabela ser renderizada
        const searchInput = document.getElementById('search-input');
        if (searchInput) {
            searchInput.addEventListener('keyup', (e) => ui.filterTable(e.target.value));
        }

        ui.showDashboard(dom.dashboardContent, dom.emptyState);

    } catch (error) {
        console.error("Falha no processamento:", error);
        ui.renderAlert(dom.alertContainer, error.message);
        ui.resetDashboard(dom.dashboardContent, dom.emptyState, dom.kpiGrid, dom.chartsGrid, dom.tableWrapper);
    } finally {
        ui.hideLoading(dom.loadingOverlay);
        if (event.target) {
            event.target.value = ''; // Permite recarregar o mesmo arquivo
        }
    }
}

// Inicializador da aplicação
function init() {
    dom.enterDashboardButton.addEventListener('click', showDashboardView);
    dom.fileUpload.addEventListener('change', handleFileSelect);
    ui.resetDashboard(dom.dashboardContent, dom.emptyState, dom.kpiGrid, dom.chartsGrid, dom.tableWrapper); // Garante que o dashboard está oculto inicialmente
    console.log('Dashboard OEE Moura - Aplicação iniciada.');
}

// Roda a aplicação quando o DOM estiver pronto
document.addEventListener('DOMContentLoaded', init);