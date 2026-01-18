// ===================================
// CONFIGURAÇÃO DOS ARQUIVOS JSON (8 DISTRITOS + GERAL)
// ===================================
const JSON_FILES = [
  { name: "ELDORADO", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_ELDORADO" },
  { name: "INDUSTRIAL", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_INDUSTRIAL" },
  { name: "NACIONAL", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_NACIONAL" },
  { name: "PETROLÂNDIA", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_PETROLANDIA" },
  { name: "RESSACA", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_RESSACA" },
  { name: "RIACHO", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_RIACHO" },
  { name: "SEDE", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_SEDE" },
  { name: "VARGEM DAS FLORES", url: "https://drive.google.com/uc?export=download&id=ID_DO_JSON_VARGEM" },
];

// ===================================
// VARIÁVEIS GLOBAIS
// ===================================
let allData = [];
let filteredData = [];

// ===================================
// FUNÇÃO AUXILIAR PARA VERIFICAR SE USUÁRIO ESTÁ PREENCHIDO
// ===================================
function hasUsuarioPreenchido(item) {
  return item['usuario'] && item['usuario'].trim() !== '';
}

// ===================================
// INICIALIZAÇÃO
// ===================================
document.addEventListener('DOMContentLoaded', () => {
  loadData();
});

// ===================================
// CARREGAR DADOS DE TODOS OS JSONS
// ===================================
async function loadData() {
  showLoading(true);
  allData = [];

  try {
    const promises = JSON_FILES.map(file =>
      fetch(file.url)
        .then(resp => resp.ok ? resp.json() : [])
        .then(jsonData => jsonData.map(item => ({ ...item, _distrito: file.name })))
        .catch(() => [])
    );

    const results = await Promise.all(promises);
    results.forEach(arr => allData.push(...arr));

    if (allData.length === 0) throw new Error("Nenhum dado foi carregado dos JSONs");

    filteredData = [...allData];
    populateFilters();
    updateDashboard();

  } catch (error) {
    alert(`Erro ao carregar os dados: ${error.message}`);
  } finally {
    showLoading(false);
  }
}

// ===================================
// MOSTRAR/OCULTAR LOADING
// ===================================
function showLoading(show) {
  const overlay = document.getElementById('loadingOverlay');
  if (!overlay) return;
  if (show) overlay.classList.add('active');
  else overlay.classList.remove('active');
}

// ===================================
// POPULAR FILTROS
// ===================================
function populateFilters() {
  const distritos = [...new Set(allData.map(item => item._distrito))].sort();
  renderMultiSelect('msDistritoPanel', distritos, applyFilters);
  setMultiSelectText('msDistritoText', [], 'Todos os Distritos');

  const statusList = [...new Set(allData.map(item => item.status))].sort();
  renderMultiSelect('msStatusPanel', statusList, applyFilters);
  setMultiSelectText('msStatusText', [], 'Todos');
}

// ===================================
// FILTRAR DADOS
// ===================================
function applyFilters() {
  const distritoSel = getSelectedFromPanel('msDistritoPanel');
  const statusSel = getSelectedFromPanel('msStatusPanel');

  setMultiSelectText('msDistritoText', distritoSel, 'Todos os Distritos');
  setMultiSelectText('msStatusText', statusSel, 'Todos');

  filteredData = allData.filter(item => {
    const okDistrito = distritoSel.length === 0 || distritoSel.includes(item._distrito);
    const okStatus = statusSel.length === 0 || statusSel.includes(item.status);
    return okDistrito && okStatus;
  });

  updateDashboard();
}

// ===================================
// FUNÇÕES AUXILIARES DE MULTISELECT
// ===================================
function renderMultiSelect(panelId, values, onChange) {
  const panel = document.getElementById(panelId);
  if (!panel) return;
  panel.innerHTML = '';
  values.forEach(v => {
    const item = document.createElement('label');
    item.innerHTML = `<input type="checkbox" value="${v}"> ${v}`;
    item.querySelector('input').addEventListener('change', onChange);
    panel.appendChild(item);
  });
}

function getSelectedFromPanel(panelId) {
  const panel = document.getElementById(panelId);
  if (!panel) return [];
  return [...panel.querySelectorAll('input[type="checkbox"]:checked')].map(cb => cb.value);
}

function setMultiSelectText(textId, selected, fallbackLabel) {
  const el = document.getElementById(textId);
  if (!el) return;
  if (!selected || selected.length === 0) el.textContent = fallbackLabel;
  else if (selected.length === 1) el.textContent = selected[0];
  else el.textContent = `${selected.length} selecionados`;
}

// ===================================
// ATUALIZAR DASHBOARD
// ===================================
function updateDashboard() {
  updateCards();
  updateCharts();
}

function updateCards() {
  const total = filteredData.filter(item => hasUsuarioPreenchido(item)).length;
  document.getElementById('totalPendencias').textContent = total;
}

function updateCharts() {
  // Exemplo simples de gráfico por status
  const statusCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const status = item.status || 'Não informado';
    statusCount[status] = (statusCount[status] || 0) + 1;
  });

  const labels = Object.keys(statusCount);
  const data = labels.map(l => statusCount[l]);

  const ctx = document.getElementById('chartStatus');
  if (!ctx) return;

  new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ data, backgroundColor: '#f97316' }] },
    options: { responsive: true }
  });
}

// ===================================
// ✅ ATUALIZAR GRÁFICOS (COM LEGENDAS DENTRO DAS BARRAS - NO MEIO)
// ===================================
function updateCharts() {
  // DISTRITOS - todos
  const distritosCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const distrito = item['_distrito'] || 'Não informado';
    distritosCount[distrito] = (distritosCount[distrito] || 0) + 1;
  });

  const distritosLabels = Object.keys(distritosCount).sort((a, b) => distritosCount[b] - distritosCount[a]);
  const distritosValues = distritosLabels.map(label => distritosCount[label]);
  createDistritoChart('chartDistritos', distritosLabels, distritosValues);

  // DISTRITOS - pendentes
  const distritosCountPendentes = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    if (item['_tipo'] !== 'PENDENTE') return;
    const distrito = item['_distrito'] || 'Não informado';
    distritosCountPendentes[distrito] = (distritosCountPendentes[distrito] || 0) + 1;
  });

  const distritosLabelsPendentes = Object.keys(distritosCountPendentes).sort((a, b) => distritosCountPendentes[b] - distritosCountPendentes[a]);
  const distritosValuesPendentes = distritosLabelsPendentes.map(label => distritosCountPendentes[label]);
  createDistritoPendenteChart('chartDistritosPendentes', distritosLabelsPendentes, distritosValuesPendentes);

  createResolutividadeDistritoChart();

  // STATUS
  const statusCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const status = getColumnValue(item, ['Status', 'STATUS', 'status'], 'Não informado');
    statusCount[status] = (statusCount[status] || 0) + 1;
  });

  const statusLabels = Object.keys(statusCount).sort((a, b) => statusCount[b] - statusCount[a]);
  const statusValues = statusLabels.map(label => statusCount[label]);
  createStatusChart('chartStatus', statusLabels, statusValues);

  // PRESTADOR - todos
  const prestadoresCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const prestador = item['Prestador'] || 'Não informado';
    prestadoresCount[prestador] = (prestadoresCount[prestador] || 0) + 1;
  });

  const prestadoresLabels = Object.keys(prestadoresCount).sort((a, b) => prestadoresCount[b] - prestadoresCount[a]).slice(0, 50);
  const prestadoresValues = prestadoresLabels.map(label => prestadoresCount[label]);
  createPrestadorChart('chartPrestadores', prestadoresLabels, prestadoresValues);

  // PRESTADOR - pendentes
  const prestadoresCountPendentes = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    if (item['_tipo'] !== 'PENDENTE') return;
    const prestador = item['Prestador'] || 'Não informado';
    prestadoresCountPendentes[prestador] = (prestadoresCountPendentes[prestador] || 0) + 1;
  });

  const prestadoresLabelsPendentes = Object.keys(prestadoresCountPendentes).sort((a, b) => prestadoresCountPendentes[b] - prestadoresCountPendentes[a]).slice(0, 50);
  const prestadoresValuesPendentes = prestadoresLabelsPendentes.map(label => prestadoresCountPendentes[label]);
  createPrestadorPendenteChart('chartPrestadoresPendentes', prestadoresLabelsPendentes, prestadoresValuesPendentes);

  createResolutividadePrestadorChart();

  // PIZZA
  createPieChart('chartPizzaStatus', statusLabels, statusValues);

  // POR MÊS
  const mesCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const dataInicio = parseDate(getColumnValue(item, [
      'Data Início da Pendência',
      'Data Inicio da Pendencia',
      'Data Início Pendência',
      'Data Inicio Pendencia'
    ]));
    if (!dataInicio) return;

    const y = dataInicio.getFullYear();
    const m = String(dataInicio.getMonth() + 1).padStart(2, '0');
    const key = `${y}-${m}`;
    mesCount[key] = (mesCount[key] || 0) + 1;
  });

  const mesKeys = Object.keys(mesCount).sort();
  const mesLabels = mesKeys.map(key => {
    const [ano, mes] = key.split('-');
    const nomeMes = new Date(Number(ano), Number(mes) - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
    return nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
  });
  const mesValues = mesKeys.map(k => mesCount[k]);

  createPendenciasPorMesChart('chartPendenciasPorMes', mesLabels, mesValues);
}

// ===================================
// ✅ GRÁFICO: Total de Pendências por Mês (BARRAS HORIZONTAIS)
// ===================================
function createPendenciasPorMesChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPendenciasPorMes) chartPendenciasPorMes.destroy();

  chartPendenciasPorMes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#1e3a8a',
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.7,
        categoryPercentage: 0.8
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { display: false },
          grid: { display: false },
          border: { display: false }
        },
        y: {
          ticks: {
            font: { size: 13, weight: 'bold' },
            color: '#1e3a8a'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'pendenciasMesInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const xPos = bar.x - 8;
          ctx.fillText(text, xPos, bar.y);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// ✅ GRÁFICO: Registros Geral de Pendências por Distrito (LEGENDAS NO MEIO DA BARRA - BRANCO E NEGRITO)
// ===================================
function createDistritoChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartDistritos) chartDistritos.destroy();

  chartDistritos = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#1e3a8a',
        borderWidth: 0,
        borderRadius: 8,
        barPercentage: 0.65,
        categoryPercentage: 0.75
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: { 
          ticks: { 
            font: { size: 13, weight: 'bold' }, 
            color: '#1e3a8a', 
            maxRotation: 45, 
            minRotation: 0 
          }, 
          grid: { display: false } 
        },
        y: { 
          beginAtZero: true, 
          ticks: { display: false }, 
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'distritosInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const xPos = bar.x;
          // ✅ AJUSTE: Posição Y no meio da barra
          const yPos = bar.y + (bar.height / 2);
          ctx.fillText(text, xPos, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// ✅ GRÁFICO: Pendências Não Resolvidas por Distrito (LEGENDAS NO MEIO DA BARRA - BRANCO E NEGRITO)
// ===================================
function createDistritoPendenteChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartDistritosPendentes) chartDistritosPendentes.destroy();

  chartDistritosPendentes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#dc2626',
        borderWidth: 0,
        borderRadius: 8,
        barPercentage: 0.65,
        categoryPercentage: 0.75
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: { 
          ticks: { 
            font: { size: 13, weight: 'bold' }, 
            color: '#dc2626', 
            maxRotation: 45, 
            minRotation: 0 
          }, 
          grid: { display: false } 
        },
        y: { 
          beginAtZero: true, 
          ticks: { display: false }, 
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'distritoPendenteInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const xPos = bar.x;
          // ✅ AJUSTE: Posição Y no meio da barra
          const yPos = bar.y + (bar.height / 2);
          ctx.fillText(text, xPos, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

function createResolutividadeDistritoChart() {
  const ctx = document.getElementById('chartResolutividadeDistrito');

  const distritosStats = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const distrito = item['_distrito'] || 'Não informado';
    if (!distritosStats[distrito]) distritosStats[distrito] = { total: 0, resolvidos: 0 };

    distritosStats[distrito].total++;
    if (item['_tipo'] === 'RESOLVIDO') distritosStats[distrito].resolvidos++;
  });

  const labels = Object.keys(distritosStats).sort((a, b) => {
    const percA = (distritosStats[a].resolvidos / distritosStats[a].total) * 100;
    const percB = (distritosStats[b].resolvidos / distritosStats[b].total) * 100;
    return percB - percA;
  });

  const percentuais = labels.map(d => {
    const s = distritosStats[d];
    return s.total > 0 ? Number(((s.resolvidos / s.total) * 100).toFixed(1)) : 0;
  });

  if (chartResolutividadeDistrito) chartResolutividadeDistrito.destroy();

  chartResolutividadeDistrito = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Taxa de Resolutividade (%)',
        data: percentuais,
        backgroundColor: '#059669',
        borderWidth: 0,
        borderRadius: 8,
        barPercentage: 0.65,
        categoryPercentage: 0.75
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { font: { size: 14, weight: 'bold' }, color: '#059669' } },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(5, 150, 105, 0.9)',
          titleFont: { size: 16, weight: 'bold' },
          bodyFont: { size: 14 },
          padding: 14,
          cornerRadius: 8,
          callbacks: {
            label: function(context) {
              const distrito = context.label;
              const stats = distritosStats[distrito];
              return [
                `Resolutividade: ${context.parsed.x}%`,
                `Resolvidos: ${stats.resolvidos}`,
                `Total: ${stats.total}`
              ];
            }
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          max: 100,
          ticks: {
            font: { size: 12, weight: '600' },
            color: '#4a5568',
            callback: function(value) { return value + '%'; }
          },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        y: { ticks: { font: { size: 13, weight: 'bold' }, color: '#059669' }, grid: { display: false } }
      }
    }
  });
}

// ===================================
// ✅ GRÁFICO: Registros Geral de Pendências por Status (LEGENDAS NO MEIO DA BARRA - BRANCO E NEGRITO)
// ===================================
function createStatusChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartStatus) chartStatus.destroy();

  chartStatus = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#f97316',
        borderWidth: 0,
        borderRadius: 8,
        barPercentage: 0.65,
        categoryPercentage: 0.75
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: { 
          ticks: { 
            font: { size: 13, weight: 'bold' }, 
            color: '#f97316', 
            maxRotation: 45, 
            minRotation: 0 
          }, 
          grid: { display: false } 
        },
        y: { 
          beginAtZero: true, 
          ticks: { display: false }, 
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'statusInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const xPos = bar.x;
          // ✅ AJUSTE: Posição Y no meio da barra
          const yPos = bar.y + (bar.height / 2);
          ctx.fillText(text, xPos, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// ✅ GRÁFICO: Registros Geral de Pendências por Prestador (BARRAS HORIZONTAIS)
// ===================================
function createPrestadorChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPrestadores) chartPrestadores.destroy();

  chartPrestadores = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#8b5cf6',
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.7,
        categoryPercentage: 0.8
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { display: false },
          grid: { display: false },
          border: { display: false }
        },
        y: {
          ticks: {
            font: { size: 13, weight: 'bold' },
            color: '#8b5cf6'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'prestadorInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const xPos = bar.x - 8;
          ctx.fillText(text, xPos, bar.y);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// ✅ GRÁFICO: Pendências Não Resolvidas por Prestador (BARRAS HORIZONTAIS)
// ===================================
function createPrestadorPendenteChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPrestadoresPendentes) chartPrestadoresPendentes.destroy();

  chartPrestadoresPendentes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#065f46',
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.7,
        categoryPercentage: 0.8
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { display: false },
          grid: { display: false },
          border: { display: false }
        },
        y: {
          ticks: {
            font: { size: 13, weight: 'bold' },
            color: '#065f46'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'prestadorPendInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const xPos = bar.x - 8;
          ctx.fillText(text, xPos, bar.y);
        });

        ctx.restore();
      }
    }]
  });
}

function createResolutividadePrestadorChart() {
  const ctx = document.getElementById('chartResolutividadePrestador');

  const prestadoresStats = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const prestador = item['Prestador'] || 'Não informado';
    if (!prestadoresStats[prestador]) prestadoresStats[prestador] = { total: 0, resolvidos: 0 };

    prestadoresStats[prestador].total++;
    if (item['_tipo'] === 'RESOLVIDO') prestadoresStats[prestador].resolvidos++;
  });

  const labels = Object.keys(prestadoresStats)
    .sort((a, b) => {
      const percA = (prestadoresStats[a].resolvidos / prestadoresStats[a].total) * 100;
      const percB = (prestadoresStats[b].resolvidos / prestadoresStats[b].total) * 100;
      return percB - percA;
    })
    .slice(0, 15);

  const percentuais = labels.map(p => {
    const s = prestadoresStats[p];
    return s.total > 0 ? Number(((s.resolvidos / s.total) * 100).toFixed(1)) : 0;
  });

  if (chartResolutividadePrestador) chartResolutividadePrestador.destroy();

  chartResolutividadePrestador = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Taxa de Resolutividade (%)',
        data: percentuais,
        backgroundColor: '#059669',
        borderWidth: 0,
        borderRadius: 8,
        barPercentage: 0.65,
        categoryPercentage: 0.75
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { font: { size: 14, weight: 'bold' }, color: '#059669' } },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(5, 150, 105, 0.9)',
          titleFont: { size: 16, weight: 'bold' },
          bodyFont: { size: 14 },
          padding: 14,
          cornerRadius: 8,
          callbacks: {
            label: function(context) {
              const prestador = context.label;
              const stats = prestadoresStats[prestador];
              return [
                `Resolutividade: ${context.parsed.x}%`,
                `Resolvidos: ${stats.resolvidos}`,
                `Total: ${stats.total}`
              ];
            }
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          max: 100,
          ticks: {
            font: { size: 12, weight: '600' },
            color: '#4a5568',
            callback: function(value) { return value + '%'; }
          },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        y: { ticks: { font: { size: 13, weight: 'bold' }, color: '#059669' }, grid: { display: false } }
      }
    }
  });
}

function createPieChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPizzaStatus) chartPizzaStatus.destroy();

  const colors = ['#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6','#ec4899','#06b6d4','#f97316','#6366f1','#84cc16'];
  const total = data.reduce((sum, val) => sum + val, 0);

  chartPizzaStatus = new Chart(ctx, {
    type: 'pie',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colors.slice(0, labels.length),
        borderWidth: 3,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          position: 'right',
          labels: {
            font: { size: 14, weight: 'bold', family: 'Arial, sans-serif' },
            color: '#000000',
            padding: 15,
            usePointStyle: true,
            pointStyle: 'circle',
            boxWidth: 20,
            boxHeight: 20,
            generateLabels: function(chart) {
              const datasets = chart.data.datasets;
              const labels = chart.data.labels;
              return labels.map((label, i) => {
                const value = datasets[0].data[i];
                const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                return {
                  text: `${label} (${percentage}%)`,
                  fillStyle: datasets[0].backgroundColor[i],
                  strokeStyle: datasets[0].backgroundColor[i],
                  lineWidth: 2,
                  hidden: false,
                  index: i
                };
              });
            }
          }
        },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(0, 0, 0, 0.8)',
          titleFont: { size: 14, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8,
          callbacks: {
            label: function(context) {
              const value = context.parsed;
              const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
              return `${context.label}: ${percentage}% (${value} registros)`;
            }
          }
        }
      }
    },
    plugins: [{
      id: 'customPieLabelsInside',
      afterDatasetsDraw: function(chart) {
        const ctx = chart.ctx;
        const dataset = chart.data.datasets[0];
        ctx.save();
        ctx.font = 'bold 14px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        chart.getDatasetMeta(0).data.forEach(function(element, index) {
          const value = dataset.data[index];
          const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
          if (parseFloat(percentage) > 5) {
            ctx.fillStyle = '#ffffff';
            const position = element.tooltipPosition();
            ctx.fillText(`${percentage}%`, position.x, position.y);
          }
        });

        ctx.restore();
      }
    }]
  });
}

function parseDate(dateString) {
  if (!dateString || dateString === '-') return null;

  let match = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return new Date(match[3], match[2] - 1, match[1]);

  match = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(match[1], match[2] - 1, match[3]);

  return null;
}

function formatDate(dateString) {
  if (!dateString || dateString === '-') return '-';

  const date = parseDate(dateString);
  if (!date || isNaN(date.getTime())) return dateString;

  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();

  return `${day}/${month}/${year}`;
}

function refreshData() {
  loadData();
}

function downloadExcel() {
  const dataParaExportar = filteredData.filter(item => hasUsuarioPreenchido(item));

  if (dataParaExportar.length === 0) {
    alert('Não há dados com Usuário preenchido para exportar.');
    return;
  }

  const exportData = dataParaExportar.map(item => ({
    'Nº Prontuário': getColumnValue(item, ['Nº Prontuário','N° Prontuário','Numero Prontuário','Prontuário','Prontuario'], ''),
    'Telefone': item['Telefone'] || '',
    'Distrito': item['_distrito'] || '',
    'Origem': item['_origem'] || '',
    'Data Solicitação': getColumnValue(item, ['Data da Solicitação','Data Solicitação','Data da Solicitacao','Data Solicitacao'], ''),
    'Unidade Solicitante': item['Unidade Solicitante'] || '',
    'CBO Especialidade': getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO'], ''),
    'Data Início Pendência': getColumnValue(item, ['Data Início da Pendência','Data Início Pendência','Data Inicio da Pendencia','Data Inicio Pendencia'], ''),
    'Status': item['Status'] || '',
    'Prestador': item['Prestador'] || '',
    'Usuário': getColumnValue(item, ['Usuário','Usuario'], ''),
    'Data Final Prazo 15d': getColumnValue(item, ['Data Final do Prazo (Pendência com 15 dias)','Data Final do Prazo (Pendencia com 15 dias)','Data Final Prazo 15d','Prazo 15 dias'], ''),
    'Data Envio Email 15d': getColumnValue(item, ['Data do envio do Email (Prazo: Pendência com 15 dias)','Data do envio do Email (Prazo: Pendencia com 15 dias)','Data Envio Email 15d','Email 15 dias'], ''),
    'Data Final Prazo 30d': getColumnValue(item, ['Data Final do Prazo (Pendência com 30 dias)','Data Final do Prazo (Pendencia com 30 dias)','Data Final Prazo 30d','Prazo 30 dias'], ''),
    'Data Envio Email 30d': getColumnValue(item, ['Data do envio do Email (Prazo: Pendência com 30 dias)','Data do envio do Email (Prazo: Pendencia com 30 dias)','Data Envio Email 30d','Email 30 dias'], '')
  }));

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Dados Completos');

  ws['!cols'] = [
    { wch: 18 },
    { wch: 18 },
    { wch: 16 },
    { wch: 20 },
    { wch: 30 },
    { wch: 18 },
    { wch: 30 },
    { wch: 30 },
    { wch: 18 },
    { wch: 20 },
    { wch: 25 },
    { wch: 20 },
    { wch: 18 },
    { wch: 20 },
    { wch: 18 },
    { wch: 20 }
  ];

  const hoje = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Dados_Todos_Distritos_${hoje}.xlsx`);
}

/* =========================================================
   TABELA "Todas as Demandas"
========================================================= */

const TABLE_PAGE_SIZE = 100;
let tableCurrentPage = 1;
let tableSearchQuery = '';
let tableColumnFilters = {};

const TABLE_COLUMNS = [
  { key: 'origem', label: 'Origem da planilha' },
  { key: 'dataSolicitacao', label: 'Data Solicitação' },
  { key: 'prontuario', label: 'Nº Prontuário' },
  { key: 'telefone', label: 'Telefone' },
  { key: 'unidadeSolicitante', label: 'Unidade Solicitante' },
  { key: 'cboEspecialidade', label: 'CBO Especialidade' },
  { key: 'dataInicioPendencia', label: 'Data Início Pendência' },
  { key: 'status', label: 'Status' },
  { key: 'dataFinalPrazo15', label: 'Data Final Prazo (15d)' },
  { key: 'dataEnvioEmail15', label: 'Data Envio Email (15d)' },
  { key: 'dataFinalPrazo30', label: 'Data Final Prazo (30d)' },
  { key: 'dataEnvioEmail30', label: 'Data Envio Email (30d)' }
];

function onTableSearch() {
  const el = document.getElementById('tableSearchInput');
  tableSearchQuery = (el.value || '').trim().toLowerCase();
  tableCurrentPage = 1;
  updateDemandasTable();
}

function tablePrevPage() {
  if (tableCurrentPage > 1) {
    tableCurrentPage--;
    updateDemandasTable();
  }
}

function tableNextPage() {
  tableCurrentPage++;
  updateDemandasTable();
}

function normalizeText(s) {
  return String(s || '').toLowerCase();
}

function getTableRowObject(item) {
  return {
    origem: item['_origem'] || '-',
    dataSolicitacao: formatDate(getColumnValue(item, ['Data da Solicitação','Data Solicitação','Data da Solicitacao','Data Solicitacao'], '-')),
    prontuario: getColumnValue(item, ['Nº Prontuário','N° Prontuário','Numero Prontuário','Prontuário','Prontuario'], '-'),
    telefone: getColumnValue(item, ['Telefone','TELEFONE'], '-'),
    unidadeSolicitante: getColumnValue(item, ['Unidade Solicitante','Unidade','UNIDADE SOLICITANTE'], '-'),
    cboEspecialidade: getColumnValue(item, ['Cbo Especialidade','CBO Especialidade','CBO','Especialidade','Especialidade CBO'], '-'),
    dataInicioPendencia: formatDate(getColumnValue(item, ['Data Início Pendência','Data Início da Pendência','Data Inicio da Pendencia','Data Inicio Pendencia'], '-')),
    status: getColumnValue(item, ['Status','STATUS'], '-'),
    dataFinalPrazo15: formatDate(getColumnValue(item, ['Data Final Prazo (15d)','Data Final do Prazo (Pendência com 15 dias)','Data Final do Prazo (Pendencia com 15 dias)','Data Final Prazo 15d','Prazo 15 dias'], '-')),
    dataEnvioEmail15: formatDate(getColumnValue(item, ['Data Envio Email (15d)','Data do envio do Email (Prazo: Pendência com 15 dias)','Data do envio do Email (Prazo: Pendencia com 15 dias)','Data Envio Email 15d','Email 15 dias'], '-')),
    dataFinalPrazo30: formatDate(getColumnValue(item, ['Data Final Prazo (30d)','Data Final do Prazo (Pendência com 30 dias)','Data Final do Prazo (Pendencia com 30 dias)','Data Final Prazo 30d','Prazo 30 dias'], '-')),
    dataEnvioEmail30: formatDate(getColumnValue(item, ['Data Envio Email (30d)','Data do envio do Email (Prazo: Pendência com 30 dias)','Data do envio do Email (Prazo: Pendencia com 30 dias)','Data Envio Email 30d','Email 30 dias'], '-'))
  };
}

function buildTableHeaderWithFilters(rows) {
  const thead = document.getElementById('demandasThead');
  thead.innerHTML = '';

  const tr = document.createElement('tr');

  TABLE_COLUMNS.forEach(col => {
    const th = document.createElement('th');
    th.className = 'th-cell';

    const title = document.createElement('div');
    title.className = 'th-title';
    title.textContent = col.label;

    const filterBox = document.createElement('div');
    filterBox.className = 'th-filter';
    filterBox.setAttribute('data-key', col.key);

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'th-filter-btn';
    btn.innerHTML = `<i class="fas fa-filter"></i>`;
    btn.onclick = (e) => {
      e.stopPropagation();
      filterBox.classList.toggle('open');
    };

    const panel = document.createElement('div');
    panel.className = 'th-filter-panel';

    const values = [...new Set(rows.map(r => r[col.key]).filter(v => v && v !== '-'))]
      .sort((a,b)=> String(a).localeCompare(String(b)));

    if (!tableColumnFilters[col.key]) tableColumnFilters[col.key] = new Set();

    const actions = document.createElement('div');
    actions.className = 'th-filter-actions';
    actions.innerHTML = `
      <button type="button" class="th-ms-all">Todos</button>
      <button type="button" class="th-ms-none">Limpar</button>
    `;

    const list = document.createElement('div');
    list.className = 'th-filter-list';

    actions.querySelector('.th-ms-all').onclick = () => {
      tableColumnFilters[col.key] = new Set(values);
      tableCurrentPage = 1;
      updateDemandasTable();
    };

    actions.querySelector('.th-ms-none').onclick = () => {
      tableColumnFilters[col.key] = new Set();
      tableCurrentPage = 1;
      updateDemandasTable();
    };

    values.forEach(v => {
      const lab = document.createElement('label');
      lab.className = 'th-filter-item';
      const checked = tableColumnFilters[col.key].has(v);
      lab.innerHTML = `
        <input type="checkbox" ${checked ? 'checked' : ''}>
        <span>${escapeHtml(v)}</span>
      `;
      lab.querySelector('input').onchange = (ev) => {
        const isChecked = ev.target.checked;
        const set = tableColumnFilters[col.key] || new Set();
        if (isChecked) set.add(v);
        else set.delete(v);
        tableColumnFilters[col.key] = set;
        tableCurrentPage = 1;
        updateDemandasTable();
      };
      list.appendChild(lab);
    });

    panel.appendChild(actions);
    panel.appendChild(list);

    filterBox.appendChild(btn);
    filterBox.appendChild(panel);

    th.appendChild(title);
    th.appendChild(filterBox);
    tr.appendChild(th);
  });

  thead.appendChild(tr);
}

function applyTableFilters(rows) {
  let out = [...rows];

  if (tableSearchQuery) {
    out = out.filter(r => {
      const blob = TABLE_COLUMNS.map(c => r[c.key]).join(' | ');
      return normalizeText(blob).includes(tableSearchQuery);
    });
  }

  TABLE_COLUMNS.forEach(col => {
    const set = tableColumnFilters[col.key];
    if (set && set.size > 0) {
      out = out.filter(r => r[col.key] && set.has(r[col.key]));
    }
  });

  return out;
}

function paginate(rows) {
  const total = rows.length;
  const totalPages = Math.max(1, Math.ceil(total / TABLE_PAGE_SIZE));

  if (tableCurrentPage > totalPages) tableCurrentPage = totalPages;
  if (tableCurrentPage < 1) tableCurrentPage = 1;

  const start = (tableCurrentPage - 1) * TABLE_PAGE_SIZE;
  const end = start + TABLE_PAGE_SIZE;
  const pageRows = rows.slice(start, end);

  return { pageRows, total, totalPages };
}

function updateDemandasTable() {
  const baseItems = filteredData.filter(item => hasUsuarioPreenchido(item));
  const rows = baseItems.map(getTableRowObject);

  buildTableHeaderWithFilters(rows);

  const filteredRows = applyTableFilters(rows);
  const { pageRows, total, totalPages } = paginate(filteredRows);

  const tbody = document.getElementById('demandasTbody');
  tbody.innerHTML = '';

  pageRows.forEach(r => {
    const tr = document.createElement('tr');
    TABLE_COLUMNS.forEach(col => {
      const td = document.createElement('td');
      td.textContent = r[col.key] ?? '-';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  const tableInfo = document.getElementById('tableInfo');
  tableInfo.textContent = `${total} registros`;

  const pageIndicator = document.getElementById('pageIndicator');
  pageIndicator.textContent = `Página ${tableCurrentPage} de ${totalPages}`;

  const btns = document.querySelectorAll('.table-pagination .btn-page');
  const btnPrev = btns[0];
  const btnNext = btns[1];

  if (btnPrev) btnPrev.disabled = (tableCurrentPage <= 1);
  if (btnNext) btnNext.disabled = (tableCurrentPage >= totalPages);
}

