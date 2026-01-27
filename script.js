// ===================================
// PLANILHAS (8 DISTRITOS)
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
let chartEvolucaoTemporal = null;

// ===================================
// TABELA: paginação e filtros
// ===================================
let TABLE_PAGE_SIZE = 50;
let tableCurrentPage = 1;
let tableSearchQuery = '';
let tableColumnFilters = {}; // (mantido, se usar depois)

// ===================================
// FUNÇÃO AUXILIAR PARA VERIFICAR SE USUÁRIO ESTÁ PREENCHIDO
// ===================================
function hasUsuarioPreenchido(item) {
  const usuario = getColumnValue(item, ['Usuário', 'Usuario', 'USUÁRIO', 'USUARIO']);
  return usuario && usuario !== '-' && usuario.trim() !== '';
}

// ===================================
// FUNÇÃO AUXILIAR PARA BUSCAR VALOR DE COLUNA (MELHORADA)
// ===================================
function getColumnValue(item, possibleNames, defaultValue = '-') {
  for (let name of possibleNames) {
    if (item.hasOwnProperty(name) && item[name]) return item[name];

    const trimmedName = name.trim();
    if (item.hasOwnProperty(trimmedName) && item[trimmedName]) return item[trimmedName];

    const keys = Object.keys(item);
    const foundKey = keys.find(k => k.toLowerCase().trim() === name.toLowerCase().trim());
    if (foundKey && item[foundKey]) return item[foundKey];
  }
  return defaultValue;
}

// ===================================
// CALCULAR PRAZOS
// ===================================
function calcularPrazos(dataInicio) {
  if (!dataInicio) return {
    prazo15: '-',
    email15: '-',
    prazo30: '-',
    email30: '-'
  };

  const dataInicioObj = parseDate(dataInicio);
  if (!dataInicioObj || isNaN(dataInicioObj.getTime())) {
    return {
      prazo15: '-',
      email15: '-',
      prazo30: '-',
      email30: '-'
    };
  }

  // Calcular Data Final do Prazo (15 dias)
  const prazo15Obj = new Date(dataInicioObj);
  prazo15Obj.setDate(prazo15Obj.getDate() + 15);

  // Calcular Data do Envio do Email (13 dias - 2 dias antes do prazo de 15)
  const email15Obj = new Date(dataInicioObj);
  email15Obj.setDate(email15Obj.getDate() + 13);

  // Calcular Data Final do Prazo (30 dias)
  const prazo30Obj = new Date(dataInicioObj);
  prazo30Obj.setDate(prazo30Obj.getDate() + 30);

  // Calcular Data do Envio do Email (28 dias - 2 dias antes do prazo de 30)
  const email30Obj = new Date(dataInicioObj);
  email30Obj.setDate(email30Obj.getDate() + 28);

  return {
    prazo15: formatDateFromObj(prazo15Obj),
    email15: formatDateFromObj(email15Obj),
    prazo30: formatDateFromObj(prazo30Obj),
    email30: formatDateFromObj(email30Obj)
  };
}

// ===================================
// FORMATAR DATA A PARTIR DE OBJETO DATE
// ===================================
function formatDateFromObj(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return '-';
  const day = String(dateObj.getDate()).padStart(2, '0');
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const year = dateObj.getFullYear();
  return `${day}/${month}/${year}`;
}

// ===================================
// MULTISELECT (CHECKBOX) HELPERS
// ===================================
function toggleMultiSelect(id) {
  document.getElementById(id).classList.toggle('open');
}

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
// CARREGAR DADOS DE TODAS AS PLANILHAS
// ===================================
async function loadData() {
  showLoading(true);
  allData = [];

  try {
    const promises = SHEETS.map(sheet =>
      fetch(sheet.url, {
        cache: 'no-store',
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
// POPULAR FILTROS (COM CBO ESPECIALIDADE)
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

  // CBO Especialidade
  const cboEspecialidades = [...new Set(allData.map(item => getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO'])))].filter(v => v && v !== '-').sort();
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

// ===================================
// CARDS
// CONTA PENDÊNCIAS QUE JÁ PASSARAM DE 15/30 DIAS
// ===================================
function updateCards() {
  const totalComUsuario = allData.filter(item => hasUsuarioPreenchido(item)).length;
  const filtradoComUsuario = filteredData.filter(item => hasUsuarioPreenchido(item)).length;

  // Somente pendências (aba PENDÊNCIAS) + Usuário preenchido
  const pendentesBaseAll = allData.filter(item =>
    item['_tipo'] === 'PENDENTE' && hasUsuarioPreenchido(item)
  );

  // Cards devem respeitar filtros atuais
  const pendentesBaseFiltered = filteredData.filter(item =>
    item['_tipo'] === 'PENDENTE' && hasUsuarioPreenchido(item)
  );

  const totalPendenciasResponder = pendentesBaseAll.length;

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  let vencendo15 = 0;
  let vencendo30 = 0;

  pendentesBaseFiltered.forEach(item => {
    const dataInicio = parseDate(getColumnValue(item, [
      'Data Início da Pendência',
      'Data Inicio da Pendencia',
      'Data Início Pendência',
      'Data Inicio Pendencia'
    ]));

    if (!dataInicio || isNaN(dataInicio.getTime())) return;

    // Calcular quantos dias se passaram desde o início
    const diasDecorridos = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));

    // Conta se já passou 15 dias ou mais
    if (diasDecorridos >= 15) {
      vencendo15++;
    }

    // Conta se já passou 30 dias ou mais
    if (diasDecorridos >= 30) {
      vencendo30++;
    }
  });

  document.getElementById('totalPendencias').textContent = totalComUsuario;
  document.getElementById('totalPendenciasResponder').textContent = totalPendenciasResponder;
  document.getElementById('pendencias15').textContent = vencendo15;
  document.getElementById('pendencias30').textContent = vencendo30;

  const percentFiltrados = totalComUsuario > 0
    ? ((filtradoComUsuario / totalComUsuario) * 100).toFixed(1)
    : '100.0';

  document.getElementById('percentFiltrados').textContent = percentFiltrados + '%';
}

// ===================================
// ATUALIZAR GRÁFICOS
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

  /* Evolução Temporal (por mês, Data Início da Pendência) */
  const evoCount = {};
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
    evoCount[key] = (evoCount[key] || 0) + 1;
  });

  const evoKeys = Object.keys(evoCount).sort();
  const evoLabels = evoKeys.map(key => {
    const [ano, mes] = key.split('-');
    const nome = new Date(Number(ano), Number(mes) - 1).toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' });
    return nome;
  });
  const evoValues = evoKeys.map(k => evoCount[k]);

  createEvolucaoTemporalChart('chartEvolucaoTemporal', evoLabels, evoValues);

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

  // POR MÊS (Barras horizontais)
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

/* =========================================================
   Evolução Temporal (linha + área)
========================================================= */
function createEvolucaoTemporalChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  if (chartEvolucaoTemporal) chartEvolucaoTemporal.destroy();

  chartEvolucaoTemporal = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Pendências Registradas',
        data,
        borderColor: '#f97316',
        backgroundColor: 'rgba(249,115,22,0.15)',
        fill: true,
        tension: 0.35,
        pointRadius: 4,
        pointHoverRadius: 5,
        pointBackgroundColor: '#f97316',
        pointBorderColor: '#ffffff',
        pointBorderWidth: 2,
        borderWidth: 3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          position: 'top',
          labels: {
            usePointStyle: true,
            pointStyle: 'circle',
            padding: 16,
            font: { size: 13, weight: 'bold' },
            color: '#111827'
          }
        },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(17,24,39,0.9)',
          titleFont: { size: 14, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: '#4b5563',
            font: { size: 12, weight: '600' },
            maxRotation: 0,
            minRotation: 0
          }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: {
            color: '#4b5563',
            font: { size: 12, weight: '600' }
          }
        }
      }
    }
  });
}

// ===================================
// GRÁFICO: Total de Pendências por Mês (BARRAS HORIZONTAIS)
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
// GRÁFICO: Registros Geral de Pendências por Distrito
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
          const yPos = bar.y + (bar.height / 2);
          ctx.fillText(text, xPos, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// GRÁFICO: Pendências Não Resolvidas por Distrito
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
// GRÁFICO: Registros Geral de Pendências por Status
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
          const yPos = bar.y + (bar.height / 2);
          ctx.fillText(text, xPos, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// GRÁFICO: Registros Geral de Pendências por Prestador
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
// GRÁFICO: Pendências Não Resolvidas por Prestador
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
        backgroundColor: '#dc2626',
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
            color: '#dc2626'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'prestadorPendenteInsideLabels',
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

  const prestadorStats = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const prestador = item['Prestador'] || 'Não informado';
    if (!prestadorStats[prestador]) prestadorStats[prestador] = { total: 0, resolvidos: 0 };

    prestadorStats[prestador].total++;
    if (item['_tipo'] === 'RESOLVIDO') prestadorStats[prestador].resolvidos++;
  });

  const labels = Object.keys(prestadorStats)
    .sort((a, b) => prestadorStats[b].total - prestadorStats[a].total)
    .slice(0, 20);

  const percentuais = labels.map(p => {
    const s = prestadorStats[p];
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
        legend: { display: true, labels: { font: { size: 14, weight: 'bold' }, color: '#059669' } },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(5, 150, 105, 0.9)',
          titleFont: { size: 14, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8,
          callbacks: {
            label: function(context) {
              const prestador = context.label;
              const stats = prestadorStats[prestador];
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
        y: { ticks: { font: { size: 12, weight: 'bold' }, color: '#059669' }, grid: { display: false } }
      }
    }
  });
}

// ===================================
// GRÁFICO DE PIZZA
// ===================================
function createPieChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (chartPizzaStatus) chartPizzaStatus.destroy();

  const colorMap = {
    'PENDENTE': '#3b82f6',
    'Pendente': '#3b82f6',
    'CANCELADO': '#ef4444',
    'Cancelado': '#ef4444',
    'RESOLVIDO': '#10b981',
    'Resolvido': '#10b981',
    'AGENDADO': '#f59e0b',
    'Agendado': '#f59e0b'
  };

  const colors = labels.map(label => colorMap[label] || '#8b5cf6');

  const total = data.reduce((acc, v) => acc + v, 0);

  chartPizzaStatus = new Chart(ctx, {
    type: 'pie',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colors,
        borderWidth: 3,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { size: 14, weight: 'bold' },
            color: '#111827',
            padding: 18,
            boxWidth: 22,
            usePointStyle: true,
            pointStyle: 'circle',
            generateLabels: function(chart) {
              const data = chart.data;
              if (data.labels.length && data.datasets.length) {
                return data.labels.map((label, i) => {
                  const value = data.datasets[0].data[i];
                  const percent = ((value / total) * 100).toFixed(1);
                  return {
                    text: `${label} (${percent}%)`,
                    fillStyle: data.datasets[0].backgroundColor[i],
                    hidden: false,
                    index: i
                  };
                });
              }
              return [];
            }
          }
        },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(17,24,39,0.95)',
          titleFont: { size: 15, weight: 'bold' },
          bodyFont: { size: 14 },
          padding: 14,
          cornerRadius: 10,
          callbacks: {
            label: function(context) {
              const label = context.label || '';
              const value = context.parsed || 0;
              const percent = ((value / total) * 100).toFixed(1);
              return `${label}: ${value} (${percent}%)`;
            }
          }
        }
      }
    },
    plugins: [{
      id: 'pizzaInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.font = 'bold 14px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((slice, i) => {
          const value = dataset.data[i];
          const percent = ((value / total) * 100).toFixed(1);

          const midAngle = (slice.startAngle + slice.endAngle) / 2;
          const radius = (slice.outerRadius + slice.innerRadius) / 2;

          const x = slice.x + Math.cos(midAngle) * radius * 0.7;
          const y = slice.y + Math.sin(midAngle) * radius * 0.7;

          ctx.fillStyle = '#ffffff';
          ctx.shadowColor = 'rgba(0,0,0,0.5)';
          ctx.shadowBlur = 4;
          ctx.shadowOffsetX = 1;
          ctx.shadowOffsetY = 1;

          ctx.fillText(`${percent}%`, x, y);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// FUNÇÃO DOWNLOAD EXCEL
// ===================================
function downloadExcel() {
  const dataToExport = filteredData
    .filter(item => hasUsuarioPreenchido(item))
    .map(item => {
      const dataInicioPendencia = getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ], '');

      const prazos = calcularPrazos(dataInicioPendencia);

      return {
        'Distrito': item['_distrito'] || '',
        'Tipo': item['_tipo'] || '',

        'Nº Solicitação': getColumnValue(item, [
          'Solicitação',
          'SOLICITAÇÃO',
          'Nº Solicitação',
          'Numero Solicitação'
        ], ''),

        'Data Solicitação': formatDate(getColumnValue(item, [
          'Data da Solicitação',
          'DATA DA SOLICITAÇÃO',
          'Data Solicitação',
          'Data Solicitacao'
        ], '')),

        'Nº Prontuário': getColumnValue(item, ['Nº Prontuário', 'Numero Prontuário'], ''),

        'Prestador': item['Prestador'] || '',

        'Unidade Solicitante': item['Unidade Solicitante'] || '',

        'CBO Especialidade': getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade'], ''),

        'Data Início da Pendência': formatDate(dataInicioPendencia),

        'Data Final do Prazo (Pendência com 15 dias)': prazos.prazo15,
        'Data do envio do Email (Prazo: Pendência com 15 dias)': prazos.email15,
        'Data Final do Prazo (Pendência com 30 dias)': prazos.prazo30,
        'Data do envio do Email (Prazo: Pendência com 30 dias)': prazos.email30,

        'Status': item['Status'] || ''
      };
    });

  const ws = XLSX.utils.json_to_sheet(dataToExport);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Pendências');

  const fileName = `Pendencias_${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

// ===================================
// TABELA
// ===================================
function updateDemandasTable() {
  const baseItems = filteredData.filter(item => hasUsuarioPreenchido(item));

  let rows = baseItems.map(item => {
    const dataInicioPendencia = getColumnValue(item, [
      'Data Início da Pendência',
      'Data Inicio da Pendencia',
      'Data Início Pendência',
      'Data Inicio Pendencia'
    ]);

    const prazos = calcularPrazos(dataInicioPendencia);

    return {
      _item: item,
      _dataInicio: parseDate(dataInicioPendencia),

      origem: item['_origem'] || '-',

      numeroSolicitacao: getColumnValue(item, [
        'Solicitação',
        'Solicitacâo',
        'Solicitaçâo',
        'SOLICITAÇÃO',
        'Nº Solicitação',
        'Numero Solicitação',
        'N Solicitação',
        'N Solicitao',
        'Número Solicitação',
        'Numero da Solicitação'
      ], '-'),

      dataSolicitacao: formatDate(getColumnValue(item, [
        'Data da Solicitação',
        'DATA DA SOLICITAÇÃO',
        'Data Solicitação',
        'Data Solicitacao',
        'Data da Solicitacao'
      ], '')),

      prontuario: getColumnValue(item, ['Nº Prontuário', 'Numero Prontuário'], '-'),

      prestador: getColumnValue(item, ['Prestador'], '-'),

      unidadeSolicitante: getColumnValue(item, ['Unidade Solicitante'], '-'),

      cboEspecialidade: getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade'], '-'),

      dataInicioPendencia: formatDate(dataInicioPendencia),

      prazo15: prazos.prazo15,
      email15: prazos.email15,
      prazo30: prazos.prazo30,
      email30: prazos.email30,

      status: getColumnValue(item, ['Status'], '-')
    };
  });

  if (tableSearchQuery) {
    rows = rows.filter(r => {
      return Object.values(r).some(val =>
        String(val).toLowerCase().includes(tableSearchQuery)
      );
    });
  }

  const { pageRows, total, totalPages } = paginate(rows);

  const thead = document.getElementById('demandasThead');
  thead.innerHTML = `
    <tr>
      <th>Origem</th>
      <th>Solicitação</th>
      <th>Data Solicitação</th>
      <th>Nº Prontuário</th>
      <th>Prestador</th>
      <th>Unidade Solicitante</th>
      <th>CBO Especialidade</th>
      <th>Data Início da Pendência</th>
      <th>Data Final do Prazo (15 dias)</th>
      <th>Data Envio Email (15 dias)</th>
      <th>Data Final do Prazo (30 dias)</th>
      <th>Data Envio Email (30 dias)</th>
      <th>Status</th>
    </tr>
  `;

  const tbody = document.getElementById('demandasTbody');
  tbody.innerHTML = '';

  const hoje = new Date();

  pageRows.forEach(r => {
    const tr = document.createElement('tr');

    if (r._item['_tipo'] === 'PENDENTE' && r._dataInicio) {
      const diasDecorridos = Math.floor((hoje - r._dataInicio) / (1000 * 60 * 60 * 24));
      if (diasDecorridos >= 26) {
        tr.style.backgroundColor = '#fefce8';
        tr.style.boxShadow = 'inset 4px 0 0 #fde68a';
      }
    }

    [
      'origem',
      'numeroSolicitacao',
      'dataSolicitacao',
      'prontuario',
      'prestador',
      'unidadeSolicitante',
      'cboEspecialidade',
      'dataInicioPendencia',
      'prazo15',
      'email15',
      'prazo30',
      'email30',
      'status'
    ].forEach(key => {
      const td = document.createElement('td');
      td.textContent = r[key] ?? '-';
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  document.getElementById('tableInfo').textContent = `${total} registros`;
  document.getElementById('pageIndicator').textContent = `Página ${tableCurrentPage} de ${totalPages}`;

  const btns = document.querySelectorAll('.table-pagination .btn-page');
  const btnPrev = btns[0];
  const btnNext = btns[1];

  if (btnPrev) btnPrev.disabled = (tableCurrentPage <= 1);
  if (btnNext) btnNext.disabled = (tableCurrentPage >= totalPages);
}

// ===================================
// PARSE DATE
// ===================================
function parseDate(dateStr) {
  if (!dateStr || dateStr === '-') return null;

  const s = String(dateStr).trim();

  let match = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?/);
  if (match) {
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const year = parseInt(match[3], 10);
    return new Date(year, month, day);
  }

  match = s.match(/(\d{4})-(\d{2})-(\d{2})(?:[T\s]\d{2}:\d{2}(?::\d{2})?)?/);
  if (match) {
    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const day = parseInt(match[3], 10);
    return new Date(year, month, day);
  }

  return null;
}

// ===================================
// FORMAT DATE
// ===================================
function formatDate(dateStr) {
  const d = parseDate(dateStr);
  if (!d || isNaN(d.getTime())) return dateStr;

  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

// ===================================
// PAGINAÇÃO
// ===================================
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

function tablePrevPage() {
  if (tableCurrentPage > 1) {
    tableCurrentPage--;
    updateDemandasTable();
  }
}

function tableNextPage() {
  const rows = filteredData.filter(item => hasUsuarioPreenchido(item));
  const totalPages = Math.ceil(rows.length / TABLE_PAGE_SIZE);
  if (tableCurrentPage < totalPages) {
    tableCurrentPage++;
    updateDemandasTable();
  }
}

function changeRecordsPerPage() {
  const select = document.getElementById('recordsPerPage');
  TABLE_PAGE_SIZE = parseInt(select.value, 10);
  tableCurrentPage = 1;
  document.getElementById('displayedRecords').textContent = TABLE_PAGE_SIZE;
  updateDemandasTable();
}

function onTableSearch() {
  const input = document.getElementById('tableSearchInput');
  tableSearchQuery = input.value.toLowerCase();
  tableCurrentPage = 1;
  updateDemandasTable();
}

// ===================================
// REFRESH DATA
// ===================================
function refreshData() {
  loadData();
}

