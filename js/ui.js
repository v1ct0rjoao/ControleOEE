import { STATUS_MAP, STATUS_COLORS } from './config.js';

let statusChart = null;
let topCircuitsChart = null;

/**
 * Formata uma célula de status com abreviação e cores.
 * @param {string} status - O status completo (ex: "Uso Programado").
 * @returns {string} - O HTML da célula formatada.
 */
function formatStatusCell(status) {
    const abbreviation = STATUS_MAP[status];
    if (!abbreviation) {
        return `<td>${status || '-'}</td>`; // Retorna o status original ou '-' se não houver
    }
    
    const color = STATUS_COLORS[abbreviation];
    const style = `background-color: ${color}20; color: ${color}; font-weight: 600;`; // Fundo com 20% de opacidade

    return `<td style="${style}">${abbreviation}</td>`;
}


export const ui = {
    // Funções de controle de estado visual
    showLoading: (loader) => loader.classList.remove('hidden'),
    hideLoading: (loader) => loader.classList.add('hidden'),
    showDashboard: (content, empty) => {
        content.classList.remove('hidden');
        empty.classList.add('hidden');
    },
    resetDashboard: (content, empty, kpiGrid, chartsGrid, tableWrapper) => {
        content.classList.add('hidden');
        empty.classList.remove('hidden');
        kpiGrid.innerHTML = '';
        chartsGrid.innerHTML = '';
        tableWrapper.innerHTML = '';
    },

    // Funções de renderização
    renderKPIs: (kpiGrid, { totalCircuits, statusTotals }) => {
        const totalDays = Object.values(statusTotals).reduce((a, b) => a + b, 0);
        if (totalCircuits === 0) return;

        const kpis = [
            { label: 'Total Circuitos', value: totalCircuits, color: '#64748b' },
            { label: 'Média Dias UP', value: (statusTotals.UP / totalCircuits).toFixed(1), color: STATUS_COLORS.UP },
            { label: 'Média Dias SD', value: (statusTotals.SD / totalCircuits).toFixed(1), color: STATUS_COLORS.SD },
            { label: 'Média Dias PP', value: (statusTotals.PP / totalCircuits).toFixed(1), color: STATUS_COLORS.PP },
        ];

        kpiGrid.innerHTML = kpis.map(kpi => `
            <div class="card kpi-card border-l-4" style="border-color: ${kpi.color}">
                <p class="text-sm font-medium text-slate-500">${kpi.label}</p>
                <p class="text-3xl font-bold mt-1" style="color: ${kpi.color}">${kpi.value}</p>
            </div>
        `).join('');
    },

    renderCharts: (chartsGrid, { statusTotals, circuitStats }) => {
        chartsGrid.innerHTML = `
            <div class="card lg:col-span-1">
                <h3 class="text-lg font-semibold mb-4 text-slate-700">Distribuição de Status</h3>
                <div class="relative h-64"><canvas id="status-pie-chart"></canvas></div>
            </div>
            <div class="card lg:col-span-2">
                <h3 class="text-lg font-semibold mb-4 text-slate-700">Top 10 Circuitos por Dias em UP</h3>
                <div class="relative h-64"><canvas id="top-circuits-chart"></canvas></div>
            </div>
        `;
        // Render Pie Chart
        if (statusChart) statusChart.destroy();
        statusChart = new Chart(document.getElementById('status-pie-chart'), {
            type: 'doughnut',
            data: {
                labels: ['UP', 'SD', 'PP'],
                datasets: [{
                    data: [statusTotals.UP, statusTotals.SD, statusTotals.PP],
                    backgroundColor: [STATUS_COLORS.UP, STATUS_COLORS.SD, STATUS_COLORS.PP],
                    borderColor: '#fff',
                    borderWidth: 4,
                }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: true, position: 'bottom' } } }
        });

        // Render Bar Chart
        const sortedCircuits = Object.entries(circuitStats).sort(([, a], [, b]) => b.UP - a.UP).slice(0, 10);
        if (topCircuitsChart) topCircuitsChart.destroy();
        topCircuitsChart = new Chart(document.getElementById('top-circuits-chart'), {
            type: 'bar',
            data: {
                labels: sortedCircuits.map(([circuit]) => circuit),
                datasets: [{
                    label: 'Dias em UP',
                    data: sortedCircuits.map(([, stats]) => stats.UP),
                    backgroundColor: STATUS_COLORS.UP,
                    borderRadius: 4,
                }]
            },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true } }, plugins: { legend: { display: false } } }
        });
    },

    // ===== FUNÇÃO MODIFICADA =====
    renderTable: (tableWrapper, { header, fullData }) => {
        const circuitHeader = header[0];

        const tableHeaderHTML = header.map(h => `<th>${h}</th>`).join('');

        const tableBodyHTML = fullData.map(row => {
            // A primeira célula é sempre o nome do circuito
            let rowHTML = `<td>${row[circuitHeader]}</td>`;
            
            // Itera sobre as colunas de data para formatar o status
            header.slice(1).forEach(dateHeader => {
                const status = row[dateHeader];
                rowHTML += formatStatusCell(status); // Usa a nova função auxiliar
            });

            return `<tr>${rowHTML}</tr>`;
        }).join('');

        tableWrapper.innerHTML = `
            <div class="flex flex-col md:flex-row md:items-center md:justify-between mb-4">
                <h3 class="text-lg font-semibold text-slate-700">Relatório Detalhado</h3>
                <input type="text" id="search-input" class="w-full md:w-auto p-2 border border-slate-300 rounded-md" placeholder="Pesquisar por circuito...">
            </div>
            <div class="table-container">
                <table id="details-table">
                    <thead><tr>${tableHeaderHTML}</tr></thead>
                    <tbody>${tableBodyHTML}</tbody>
                </table>
            </div>
        `;
    },
    // =============================

    filterTable: (searchTerm) => {
        const rows = document.querySelectorAll('#details-table tbody tr');
        rows.forEach(row => {
            const circuitCell = row.querySelector('td:first-child');
            if (circuitCell) {
                const isVisible = circuitCell.textContent.toLowerCase().includes(searchTerm.toLowerCase());
                row.style.display = isVisible ? '' : 'none';
            }
        });
    },

    renderAlert: (alertContainer, message, type = 'danger') => {
        alertContainer.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
        setTimeout(() => alertContainer.innerHTML = '', 5000); // O alerta some após 5 segundos
    }
};