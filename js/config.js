// Mapeamento de Status para consistência e abreviação
export const STATUS_MAP = {
    'Uso Programado': 'UP',
    'Sem Demanda': 'SD',
    'Parada Planejada': 'PP'
};

// Cores para os gráficos e UI, alinhadas com o CSS (Tons mais escuros)
export const STATUS_COLORS = {
    UP: '#14532D',    // Verde escuro
    SD: '#A16207',    // Amarelo escuro/Ocre
    PP: '#1E3A8A'     // Azul escuro
};

// Seletores do DOM para evitar repetição no código
export const DOM_SELECTORS = {
    appContainer: '#app',           // Novo container principal
    landingPage: '#landing-page',    // Nova seção para a página inicial
    dashboardView: '#dashboard-view',  // Nova seção para o dashboard principal
    enterDashboardButton: '#enter-dashboard', // Botão para entrar no dashboard

    fileUpload: '#file-upload',
    fileName: '#file-name',
    dashboardContent: '#dashboard-content',
    emptyState: '#empty-state',
    loadingOverlay: '#loading-overlay',
    kpiGrid: '#kpi-grid',
    chartsGrid: '#charts-grid',
    tableWrapper: '#table-wrapper',
    alertContainer: '#alert-container'
};

// Textos da Aplicação
export const APP_TEXTS = {
    fileError: 'Ocorreu um erro ao ler o arquivo. Verifique o formato.',
    invalidDataError: 'O arquivo parece estar vazio ou em um formato inesperado.',
    columnError: 'Não foi possível encontrar a coluna de circuitos. Verifique o cabeçalho do arquivo.'
};