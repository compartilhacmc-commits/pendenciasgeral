// ===================================
// CONFIGURAÇÃO DAS PLANILHAS (8 DISTRITOS)
// ===================================

// helper para padronizar URL CSV do Google Sheets COM CACHE BUSTING
function gvizCsvUrl(spreadsheetId, gid) {
  const timestamp = new Date().getTime();
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&gid=${gid}&_=${timestamp}`;
}

const SHEETS = [
  // DISTRITO ELDORADO 
  {
    name: 'PENDÊNCIAS ELDORADO',
    url: gvizCsvUrl('1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4', '278071504'),
    distrito: 'ELDORADO',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS ELDORADO',
    url: gvizCsvUrl('1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4', '2142054254'),
    distrito: 'ELDORADO',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO INDUSTRIAL 
  {
    name: 'PENDÊNCIAS INDUSTRIAL',
    url: gvizCsvUrl('14eUVIsWPubMve4DhVjVwlh7gin-qVyN3PspkwQ1PZMg', '278071504'),
    distrito: 'INDUSTRIAL',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS INDUSTRIAL',
    url: gvizCsvUrl('14eUVIsWPubMve4DhVjVwlh7gin-qVyN3PspkwQ1PZMg', '1086207100'),
    distrito: 'INDUSTRIAL',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO NACIONAL
  {
    name: 'PENDÊNCIAS NACIONAL',
    url: gvizCsvUrl('1lMGO9Hh_qL9OKI270fPL7lxadr-BZN9x_ZtmQeX6OcA', '278071504'),
    distrito: 'NACIONAL',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS NACIONAL',
    url: gvizCsvUrl('1lMGO9Hh_qL9OKI270fPL7lxadr-BZN9x_ZtmQeX6OcA', '150768142'),
    distrito: 'NACIONAL',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO PETROLÂNDIA
  {
    name: 'PENDÊNCIAS PETROLÂNDIA',
    url: gvizCsvUrl('1Z9Uf5MGm5tClVDR95SUpwOjivAdqEVUfDj7mIuRLf4s', '278071504'),
    distrito: 'PETROLÂNDIA',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS PETROLÂNDIA',
    url: gvizCsvUrl('1Z9Uf5MGm5tClVDR95SUpwOjivAdqEVUfDj7mIuRLf4s', '1067061018'),
    distrito: 'PETROLÂNDIA',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO RESSACA 
  {
    name: 'PENDÊNCIAS RESSACA',
    url: gvizCsvUrl('1aIsq1a8Lb90M19TQdiJG_WyX7wzzC2WRohelJY6A-u8', '278071504'),
    distrito: 'RESSACA',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS RESSACA',
    url: gvizCsvUrl('1aIsq1a8Lb90M19TQdiJG_WyX7wzzC2WRohelJY6A-u8', '699447584'),
    distrito: 'RESSACA',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO RIACHO
  {
    name: 'PENDÊNCIAS RIACHO',
    url: gvizCsvUrl('1367XyjVDYyDWo3vUz6Hd_zEqLAJkH_c1MwlvtZnpmUc', '278071504'),
    distrito: 'RIACHO',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS RIACHO',
    url: gvizCsvUrl('1367XyjVDYyDWo3vUz6Hd_zEqLAJkH_c1MwlvtZnpmUc', '1996983614'),
    distrito: 'RIACHO',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO SEDE
  {
    name: 'PENDÊNCIAS SEDE',
    url: gvizCsvUrl('1RPf2bfQVoM1FqnyA-0P8uPTJ_PG4I2Ce6lXnk54ixfc', '278071504'),
    distrito: 'SEDE',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS SEDE',
    url: gvizCsvUrl('1RPf2bfQVoM1FqnyA-0P8uPTJ_PG4I2Ce6lXnk54ixfc', '626867102'),
    distrito: 'SEDE',
    tipo: 'RESOLVIDO'
  },

  // DISTRITO VARGEM DAS FLORES
  {
    name: 'PENDÊNCIAS VARGEM DAS FLORES',
    url: gvizCsvUrl('1IHknmxe3xAnfy5Bju_23B5ivIL-qMaaE6q_HuPaLBpk', '278071504'),
    distrito: 'VARGEM DAS FLORES',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS VARGEM DAS FLORES',
    url: gvizCsvUrl('1IHknmxe3xAnfy5Bju_23B5ivIL-qMaaE6q_HuPaLBpk', '451254610'),
    distrito: 'VARGEM DAS FLORES',
    tipo: 'RESOLVIDO'
  }
];

// ===================================
// VARIÁVEIS GLOBAIS
// ===================================
let allData = [];
let filteredData = [];

let chartDistritos = null;
let chartDistritosPendentes = null;
let chartStatus = null;
let chartPrestadores = null;
let chartPrestadoresPendentes = null;
let chartPizzaStatus = null;
let chartResolutividadeDistrito = null;
let chartResolutividadePrestador = null;
let chartPendenciasPorMes = null;

// ===================================
// ✅ FUNÇÃO AUXILIAR PARA VERIFICAR SE USUÁRIO ESTÁ PREENCHIDO
// ===================================
function hasUsuarioPreenchido(item) {
  const usuario = getColumnValue(item, ['Usuário', 'Usuario', 'USUÁRIO', 'USUARIO']);
  return usuario && usuario !== '-' && usuario.trim() !== '';
}

// ===================================
// ✅ FUNÇÃO AUXILIAR PARA BUSCAR VALOR DE COLUNA (MELHORADA)
// ===================================
function getColumnValue(item, possibleNames, defaultValue = '-') {
  for (let name of possibleNames) {
    // Busca exata
    if (item.hasOwnProperty(name) && item[name]) return item[name];
    
    // Busca com trim (remove espaços)
    const trimmedName = name.trim();
    if (item.hasOwnProperty(trimmedName) && item[trimmedName]) return item[trimmedName];
    
    // Busca case-insensitive
    const keys = Object.keys(item);
    const foundKey = keys.find(k => k.toLowerCase().trim() === name.toLowerCase().trim());
    if (foundKey && item[foundKey]) return item[foundKey];
  }
  return defaultValue;
}

// ===================================
// MULTISELECT (CHECKBOX) HELPERS
// ===================================
function toggleMultiSelect(id) {
  document.getElementById(id).classList.toggle('open');
}

// fecha dropdown ao clicar fora
document.addEventListener('click', (e) => {
  document.querySelectorAll('.multi-select').forEach(ms => {
    if (!ms.contains(e.target)) ms.classList.remove('open');
  });

  document.querySelectorAll('.th-filter').forEach(box => {
    if (!box.contains(e.target)) box.classList.remove('open');
  });
});

function escapeHtml(str) {
  return String(str)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function renderMultiSelect(panelId, values, onChange) {
  const panel = document.getElementById(panelId);
  panel.innerHTML = '';

  const actions = document.createElement('div');
  actions.className = 'ms-actions';
  actions.innerHTML = `
    <button type="button" class="ms-all">Marcar todos</button>
    <button type="button" class="ms-none">Limpar</button>
  `;
  panel.appendChild(actions);

  const btnAll = actions.querySelector('.ms-all');
  const btnNone = actions.querySelector('.ms-none');

  btnAll.addEventListener('click', () => {
    panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
    onChange();
  });

  btnNone.addEventListener('click', () => {
    panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
    onChange();
  });

  values.forEach(v => {
    const item = document.createElement('label');
    item.className = 'ms-item';
    item.innerHTML = `
      <input type="checkbox" value="${escapeHtml(v)}">
      <span>${escapeHtml(v)}</span>
    `;
    item.querySelector('input').addEventListener('change', onChange);
    panel.appendChild(item);
  });
}

function getSelectedFromPanel(panelId) {
  const panel = document.getElementById(panelId);
  return [...panel.querySelectorAll('input[type="checkbox"]:checked')].map(cb => cb.value);
}

function setMultiSelectText(textId, selected, fallbackLabel) {
  const el = document.getElementById(textId);
  if (!selected || selected.length === 0) el.textContent = fallbackLabel;
  else if (selected.length === 1) el.textContent = selected[0];
  else el.textContent = `${selected.length} selecionados`;
}

// ===================================
// INICIALIZAÇÃO
// ===================================
document.addEventListener('DOMContentLoaded', function() {
  loadData();
});

// ===================================
// ✅ CARREGAR DADOS DE TODAS AS PLANILHAS
// ===================================
async function loadData() {
  showLoading(true);
  allData = [];

  try {
    const promises = SHEETS.map(sheet =>
      fetch(sheet.url, {
        cache: 'no-store', // Força não usar cache
        headers: {
          'Cache-Control': 'no-cache',
          'Pragma': 'no-cache'
        }
      })
        .then(response => response.ok ? response.text() : null)
        .then(csvText => {
          if (!csvText) return null;
          return { name: sheet.name, csv: csvText, distrito: sheet.distrito, tipo: sheet.tipo };
        })
        .catch(() => null)
    );

    const results = await Promise.all(promises);

    results.forEach(result => {
      if (!result) return;

      const rows = parseCSV(result.csv);
      if (rows.length < 2) return;

      const headers = rows[0];
      
      // ✅ DEBUG: Mostrar todas as colunas disponíveis
      console.log(`📋 Planilha: ${result.name}`);
      console.log('Colunas disponíveis:', headers);

      const sheetData = rows.slice(1)
        .filter(row => row.length > 1 && row[0])
        .map(row => {
          const obj = {
            _origem: result.name,
            _distrito: result.distrito,
            _tipo: result.tipo
          };
          headers.forEach((header, index) => {
            obj[header.trim()] = (row[index] || '').trim();
          });
          return obj;
        });

      allData.push(...sheetData);
    });

    if (allData.length === 0) throw new Error('Nenhum dado foi carregado das planilhas');

    // ✅ DEBUG: Mostrar exemplo de CBO Especialidade
    if (allData.length > 0) {
      console.log('🔍 Exemplo de registro completo:', allData[0]);
      console.log('🔍 CBO Especialidade encontrado:', getColumnValue(allData[0], ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO']));
    }

    filteredData = [...allData];
    populateFilters();
    updateDashboard();

  } catch (error) {
    alert(`Erro ao carregar dados das planilhas: ${error.message}`);
  } finally {
    showLoading(false);
  }
}

// ===================================
// PARSE CSV (COM SUPORTE A ASPAS)
// ===================================
function parseCSV(text) {
  const rows = [];
  let currentRow = [];
  let currentCell = '';
  let insideQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (char === '"') {
      if (insideQuotes && nextChar === '"') {
        currentCell += '"';
        i++;
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (char === ',' && !insideQuotes) {
      currentRow.push(currentCell.trim());
      currentCell = '';
    } else if ((char === '\n' || char === '\r') && !insideQuotes) {
      if (currentCell || currentRow.length > 0) {
        currentRow.push(currentCell.trim());
        rows.push(currentRow);
        currentRow = [];
        currentCell = '';
      }
      if (char === '\r' && nextChar === '\n') i++;
    } else {
      currentCell += char;
    }
  }

  if (currentCell || currentRow.length > 0) {
    currentRow.push(currentCell.trim());
    rows.push(currentRow);
  }

  return rows;
}

// ===================================
// MOSTRAR/OCULTAR LOADING
// ===================================
function showLoading(show) {
  const overlay = document.getElementById('loadingOverlay');
  if (show) overlay.classList.add('active');
  else overlay.classList.remove('active');
}

// ===================================
// ✅ POPULAR FILTROS (COM CBO ESPECIALIDADE)
// ===================================
function populateFilters() {
  // Distrito
  const distritos = [...new Set(allData.map(item => item['_distrito']))].filter(Boolean).sort();
  renderMultiSelect('msDistritoPanel', distritos, applyFilters);
  setMultiSelectText('msDistritoText', [], 'Todos os Distritos');

  // Unidade Solicitante
  const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
  renderMultiSelect('msUnidadePanel', unidades, applyFilters);
  setMultiSelectText('msUnidadeText', [], 'Todas');

  // Prestador
  const prestadores = [...new Set(allData.map(item => item['Prestador']))].filter(Boolean).sort();
  renderMultiSelect('msPrestadorPanel', prestadores, applyFilters);
  setMultiSelectText('msPrestadorText', [], 'Todos');

  // ✅ CBO Especialidade
  const cboEspecialidades = [...new Set(allData.map(item => getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO'])))].filter(v => v && v !== '-').sort();
  
  console.log('✅ Especialidades encontradas:', cboEspecialidades);
  
  renderMultiSelect('msCboEspecialidadePanel', cboEspecialidades, applyFilters);
  setMultiSelectText('msCboEspecialidadeText', [], 'Todas');

  // Status
  const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
  renderMultiSelect('msStatusPanel', statusList, applyFilters);
  setMultiSelectText('msStatusText', [], 'Todos');

  // Mês
  populateMonthFilter();
}

function populateMonthFilter() {
  const mesesSet = new Set();

  allData.forEach(item => {
    const dataInicio = parseDate(getColumnValue(item, [
      'Data Início da Pendência',
      'Data Inicio da Pendencia',
      'Data Início Pendência',
      'Data Inicio Pendencia'
    ]));

    if (dataInicio) {
      const mesAno = `${dataInicio.getFullYear()}-${String(dataInicio.getMonth() + 1).padStart(2, '0')}`;
      mesesSet.add(mesAno);
    }
  });

  const mesesOrdenados = Array.from(mesesSet).sort().reverse();
  const mesesFormatados = mesesOrdenados.map(mesAno => {
    const [ano, mes] = mesAno.split('-');
    const nomeMes = new Date(ano, mes - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
    return nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
  });

  renderMultiSelect('msMesPanel', mesesFormatados, applyFilters);
  setMultiSelectText('msMesText', [], 'Todos os Meses');
}

function applyFilters() {
  const distritoSel = getSelectedFromPanel('msDistritoPanel');
  const unidadeSel = getSelectedFromPanel('msUnidadePanel');
  const prestadorSel = getSelectedFromPanel('msPrestadorPanel');
  const cboEspecialidadeSel = getSelectedFromPanel('msCboEspecialidadePanel');
  const statusSel = getSelectedFromPanel('msStatusPanel');
  const mesSel = getSelectedFromPanel('msMesPanel');

  setMultiSelectText('msDistritoText', distritoSel, 'Todos os Distritos');
  setMultiSelectText('msUnidadeText', unidadeSel, 'Todas');
  setMultiSelectText('msPrestadorText', prestadorSel, 'Todos');
  setMultiSelectText('msCboEspecialidadeText', cboEspecialidadeSel, 'Todas');
  setMultiSelectText('msStatusText', statusSel, 'Todos');
  setMultiSelectText('msMesText', mesSel, 'Todos os Meses');

  filteredData = allData.filter(item => {
    const okDistrito = (distritoSel.length === 0) || distritoSel.includes(item['_distrito'] || '');
    const okUnidade = (unidadeSel.length === 0) || unidadeSel.includes(item['Unidade Solicitante'] || '');
    const okPrest = (prestadorSel.length === 0) || prestadorSel.includes(item['Prestador'] || '');
    
    const cboValue = getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO']);
    const okCbo = (cboEspecialidadeSel.length === 0) || cboEspecialidadeSel.includes(cboValue);
    
    const okStatus = (statusSel.length === 0) || statusSel.includes(item['Status'] || '');

    let okMes = true;
    if (mesSel.length > 0) {
      const dataInicio = parseDate(getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ]));
      if (dataInicio) {
        const nomeMes = new Date(dataInicio.getFullYear(), dataInicio.getMonth()).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
        const mesFormatado = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
        okMes = mesSel.includes(mesFormatado);
      } else okMes = false;
    }

    return okDistrito && okUnidade && okPrest && okCbo && okStatus && okMes;
  });

  updateDashboard();
}

function clearFilters() {
  ['msDistritoPanel','msUnidadePanel','msPrestadorPanel','msCboEspecialidadePanel','msStatusPanel','msMesPanel'].forEach(panelId => {
    const panel = document.getElementById(panelId);
    if (!panel) return;
    panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  });

  setMultiSelectText('msDistritoText', [], 'Todos os Distritos');
  setMultiSelectText('msUnidadeText', [], 'Todas');
  setMultiSelectText('msPrestadorText', [], 'Todos');
  setMultiSelectText('msCboEspecialidadeText', [], 'Todas');
  setMultiSelectText('msStatusText', [], 'Todos');
  setMultiSelectText('msMesText', [], 'Todos os Meses');

  filteredData = [...allData];
  updateDashboard();
}

// ===================================
// ATUALIZAR DASHBOARD
// ===================================
function updateDashboard() {
  updateCards();
  updateCharts();
  updateDemandasTable();
}

function updateCards() {
  const totalComUsuario = allData.filter(item => hasUsuarioPreenchido(item)).length;
  const filtradoComUsuario = filteredData.filter(item => hasUsuarioPreenchido(item)).length;

  const hoje = new Date();
  let pendencias15 = 0;
  let pendencias30 = 0;

  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const dataInicio = parseDate(getColumnValue(item, [
      'Data Início da Pendência',
      'Data Inicio da Pendencia',
      'Data Início Pendência',
      'Data Inicio Pendencia'
    ]));

    if (dataInicio) {
      const diasDecorridos = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
      if (diasDecorridos >= 15 && diasDecorridos < 30) pendencias15++;
      if (diasDecorridos >= 30) pendencias30++;
    }
  });

  document.getElementById('totalPendencias').textContent = totalComUsuario;
  document.getElementById('pendencias15').textContent = pendencias15;
  document.getElementById('pendencias30').textContent = pendencias30;

  const percentFiltrados = totalComUsuario > 0 ? ((filtradoComUsuario / totalComUsuario) * 100).toFixed(1) : '100.0';
  document.getElementById('percentFiltrados').textContent = percentFiltrados + '%';
}

// ===================================
// PLUGIN (NÚMEROS BRANCOS EM NEGRITO)
// ===================================
function addValueLabelsPlugin({
  id,
  color = '#FFFFFF',
  font = 'bold 16px Arial',
  mode = 'vertical',
  suffix = ''
} = {}) {
  return {
    id,
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      const dataset = chart.data.datasets[0];
      if (!meta || !meta.data) return;

      ctx.save();
      ctx.fillStyle = color;
      ctx.font = font;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';

      meta.data.forEach((bar, i) => {
        const value = dataset.data[i];
        const text = `${value}${suffix}`;

        if (mode === 'horizontal') {
          const xPos = bar.x + (bar.width / 2);
          ctx.fillText(text, xPos, bar.y);
        } else {
          const yPos = bar.y + (bar.height / 2);
          ctx.fillText(text, bar.x, yPos);
        }
      });

      ctx.restore();
    }
  };
}

function addOutsideValueLabelsPlugin({
  id,
  color = '#000000',
  font = 'bold 20px Arial',
  offset = 14,
  suffix = ''
} = {}) {
  return {
    id,
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      const dataset = chart.data.datasets[0];
      if (!meta || !meta.data) return;

      const isHorizontal = chart.options?.indexAxis === 'y';

      ctx.save();
      ctx.fillStyle = color;
      ctx.font = font;
      ctx.textBaseline = 'middle';

      meta.data.forEach((elem, i) => {
        const value = dataset.data[i];
        const text = `${value}${suffix}`;

        if (isHorizontal) {
          ctx.textAlign = 'left';
          ctx.fillText(text, elem.x + offset, elem.y);
        } else {
          ctx.textAlign = 'center';
          ctx.fillText(text, elem.x, elem.y - offset);
        }
      });

      ctx.restore();
    }
  };
}

// ===================================
// ✅ ATUALIZAR GRÁFICOS
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

function createPendenciasPorMesChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPendenciasPorMes) chartPendenciasPorMes.destroy();

  chartPendenciasPorMes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Pendências por Mês',
        data,
        backgroundColor: '#1e3a8a',
        borderWidth: 0,
        borderRadius: 10,
        barThickness: 18,
        maxBarThickness: 22
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { right: 26 } }, // espaço pro número fora
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(30, 58, 138, 0.9)',
          titleFont: { size: 14, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8
        }
      },
      scales: {
        y: {
          ticks: {
            font: { size: 11, weight: '600' },
            color: '#374151'
          },
          grid: { display: false },
          border: { display: false }
        },
        x: {
          beginAtZero: true,
          ticks: {
            font: { size: 12, weight: '600' },
            color: '#6b7280'
          },
          grid: { display: false }, // estilo limpo como a foto
          border: { display: false }
        }
      }
    },
    plugins: [
      addOutsideValueLabelsPlugin({
        id: 'pendenciasMesOutsideLabels',
        color: '#111827',
        font: 'bold 14px Arial',
        offset: 10
      })
    ]
  });
}

// (Demais gráficos de distrito/status/resolutividade/pizza permanecem iguais ao seu código)
// ... [MANTIDOS SEM ALTERAÇÃO] ...

/* =========================================================
   ✅✅✅ ALTERAÇÃO 2/3: Prestador (Geral) (MODELO DA FOTO)
   - barra horizontal
   - número fora na ponta direita
   - MANTÉM cor original #8b5cf6
========================================================= */
function createPrestadorChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPrestadores) chartPrestadores.destroy();

  chartPrestadores = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Pendências',
        data,
        backgroundColor: '#8b5cf6',
        borderWidth: 0,
        borderRadius: 10,
        barThickness: 18,
        maxBarThickness: 22
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { right: 26 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(139, 92, 246, 0.9)',
          titleFont: { size: 14, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8
        }
      },
      scales: {
        y: {
          ticks: {
            font: { size: 10, weight: '600' },
            color: '#374151'
          },
          grid: { display: false },
          border: { display: false }
        },
        x: {
          beginAtZero: true,
          ticks: {
            font: { size: 12, weight: '600' },
            color: '#6b7280'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [
      addOutsideValueLabelsPlugin({
        id: 'prestadorOutsideLabels',
        color: '#111827',
        font: 'bold 14px Arial',
        offset: 10
      })
    ]
  });
}

/* =========================================================
   ✅✅✅ ALTERAÇÃO 3/3: Prestador (Pendentes) (MODELO DA FOTO)
   - barra horizontal
   - número fora na ponta direita
   - MANTÉM cor original #065f46
========================================================= */
function createPrestadorPendenteChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPrestadoresPendentes) chartPrestadoresPendentes.destroy();

  chartPrestadoresPendentes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Pendências Não Resolvidas',
        data,
        backgroundColor: '#065f46',
        borderWidth: 0,
        borderRadius: 10,
        barThickness: 18,
        maxBarThickness: 22
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { right: 26 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(6, 95, 70, 0.9)',
          titleFont: { size: 14, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8
        }
      },
      scales: {
        y: {
          ticks: {
            font: { size: 10, weight: '600' },
            color: '#374151'
          },
          grid: { display: false },
          border: { display: false }
        },
        x: {
          beginAtZero: true,
          ticks: {
            font: { size: 12, weight: '600' },
            color: '#6b7280'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [
      addOutsideValueLabelsPlugin({
        id: 'prestadorPendOutsideLabels',
        color: '#111827',
        font: 'bold 14px Arial',
        offset: 10
      })
    ]
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

