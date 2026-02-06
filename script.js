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

let chartDistritosPendentes = null;
let chartDistritosResolvidas = null;
let chartStatus = null;
let chartPrestadores = null;
let chartPrestadoresPendentes = null;
let chartEspecialidades = null;
let chartEspecialidadesPendentes = null;
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
let tableColumnFilters = {};

// ===================================
// FUNÇÃO AUXILIAR PARA VERIFICAR SE USUÁRIO ESTÁ PREENCHIDO
// ===================================
function hasUsuarioPreenchido(item) {
  const usuario = getColumnValue(item, ['Usuário', 'Usuario', 'USUÁRIO', 'USUARIO'], '');
  return usuario && usuario !== '-' && String(usuario).trim() !== '';
}

// ===================================
// FUNÇÃO PARA CANCELADOS POR VENCIMENTO
// STATUS = "CANCELADO/VENCIMENTO DO PRAZO"
// ===================================
function getCanceladoPorVencimentoInfo(item) {
  // Deve estar na aba RESOLVIDOS
  if (item['_tipo'] !== 'RESOLVIDO') return { isCancelado: false, dataVencimento: null };

  // Deve ter usuário preenchido
  if (!hasUsuarioPreenchido(item)) return { isCancelado: false, dataVencimento: null };

  // Verifica se o STATUS é exatamente "CANCELADO/VENCIMENTO DO PRAZO"
  const status = getColumnValue(item, ['Status', 'STATUS', 'status'], '').trim().toUpperCase();
  
  if (status === 'CANCELADO/VENCIMENTO DO PRAZO') {
    // Busca a data de vencimento (se houver)
    const dataEmail30 = getColumnValue(item, [
      'Data do envio do Email (Prazo: Pendência com 30 dias)',
      'Data do envio do Email (Prazo Pendência com 30 dias)',
      'Data envio Email 30 dias',
      'Email 30 dias',
      'Data Envio Email (30 dias)',
      'Data Envio Email 30 dias'
    ], '');

    const dataEmail30Parsed = parseDate(dataEmail30);

    return { 
      isCancelado: true, 
      dataVencimento: dataEmail30Parsed || parseDate(getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ]))
    };
  }

  return { isCancelado: false, dataVencimento: null };
}

// Função auxiliar para compatibilidade com código existente
function isCanceladoPorVencimentoPrazo(item) {
  return getCanceladoPorVencimentoInfo(item).isCancelado;
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

  const prazo15Obj = new Date(dataInicioObj);
  prazo15Obj.setDate(prazo15Obj.getDate() + 15);

  const email15Obj = new Date(dataInicioObj);
  email15Obj.setDate(email15Obj.getDate() + 13);

  const prazo30Obj = new Date(dataInicioObj);
  prazo30Obj.setDate(prazo30Obj.getDate() + 30);

  const email30Obj = new Date(dataInicioObj);
  email30Obj.setDate(email30Obj.getDate() + 28);

  return {
    prazo15: formatDateFromObj(prazo15Obj),
    email15: formatDateFromObj(email15Obj),
    prazo30: formatDateFromObj(prazo30Obj),
    email30: formatDateFromObj(email30Obj)
  };
}

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
  const distritos = [...new Set(allData.map(item => item['_distrito']))].filter(Boolean).sort();
  renderMultiSelect('msDistritoPanel', distritos, applyFilters);
  setMultiSelectText('msDistritoText', [], 'Todos os Distritos');

  const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
  renderMultiSelect('msUnidadePanel', unidades, applyFilters);
  setMultiSelectText('msUnidadeText', [], 'Todas');

  const prestadores = [...new Set(allData.map(item => item['Prestador']))].filter(Boolean).sort();
  renderMultiSelect('msPrestadorPanel', prestadores, applyFilters);
  setMultiSelectText('msPrestadorText', [], 'Todos');

  const cboEspecialidades = [...new Set(allData.map(item => getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO'])))].filter(v => v && v !== '-').sort();
  renderMultiSelect('msCboEspecialidadePanel', cboEspecialidades, applyFilters);
  setMultiSelectText('msCboEspecialidadeText', [], 'Todas');

  const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
  renderMultiSelect('msStatusPanel', statusList, applyFilters);
  setMultiSelectText('msStatusText', [], 'Todos');

  populateMonthFilter();
}

function populateMonthFilter() {
  const mesesSet = new Set();

  allData.forEach(item => {
    const canceladoInfo = getCanceladoPorVencimentoInfo(item);

    let dataParaMes = null;

    if (canceladoInfo.isCancelado) {
      dataParaMes = canceladoInfo.dataVencimento;
    } else {
      dataParaMes = parseDate(getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ]));
    }

    if (dataParaMes) {
      const mesAno = `${dataParaMes.getFullYear()}-${String(dataParaMes.getMonth() + 1).padStart(2, '0')}`;
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
      const canceladoInfo = getCanceladoPorVencimentoInfo(item);

      let dataParaFiltro = null;

      if (canceladoInfo.isCancelado) {
        dataParaFiltro = canceladoInfo.dataVencimento;
      } else {
        dataParaFiltro = parseDate(getColumnValue(item, [
          'Data Início da Pendência',
          'Data Inicio da Pendencia',
          'Data Início Pendência',
          'Data Inicio Pendencia'
        ]));
      }

      if (dataParaFiltro) {
        const nomeMes = new Date(dataParaFiltro.getFullYear(), dataParaFiltro.getMonth()).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
        const mesFormatado = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
        okMes = mesSel.includes(mesFormatado);
      } else {
        okMes = false;
      }
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
// ===================================
function updateCards() {
  const allComUsuario = allData.filter(item => hasUsuarioPreenchido(item));
  const filteredComUsuario = filteredData.filter(item => hasUsuarioPreenchido(item));

  const totalPendenciasGeral = allComUsuario.length;

  const totalPendenciasResponder = allData.filter(item =>
    item['_tipo'] === 'PENDENTE' && hasUsuarioPreenchido(item)
  ).length;

  //CONTA APENAS STATUS = "CANCELADO/VENCIMENTO DO PRAZO"
  const totalCanceladosVencimento = filteredComUsuario.filter(item =>
    isCanceladoPorVencimentoPrazo(item)
  ).length;

  const totalResolvidas = filteredComUsuario.filter(item =>
    item['_tipo'] === 'RESOLVIDO'
  ).length;

  const totalAgendadas = filteredComUsuario.filter(item => {
    const status = getColumnValue(item, ['Status', 'STATUS', 'status'], '');
    return String(status).trim().toLowerCase() === 'agendado' || String(status).trim().toLowerCase() === 'agendada';
  }).length;

  const totalCanceladosGeral = filteredComUsuario.filter(item => {
    const status = getColumnValue(item, ['Status', 'STATUS', 'status'], '');
    return String(status).trim().toLowerCase() === 'cancelado' || String(status).trim().toLowerCase() === 'cancelada';
  }).length;

  const percentFiltrados = totalPendenciasGeral > 0
    ? ((filteredComUsuario.length / totalPendenciasGeral) * 100).toFixed(1)
    : '100.0';

  document.getElementById('totalPendencias').textContent = totalPendenciasGeral;
  document.getElementById('totalPendenciasResponder').textContent = totalPendenciasResponder;
  document.getElementById('totalCanceladosVencimento').textContent = totalCanceladosVencimento;
  document.getElementById('totalResolvidas').textContent = totalResolvidas;
  document.getElementById('totalAgendadas').textContent = totalAgendadas;
  document.getElementById('totalCanceladosGeral').textContent = totalCanceladosGeral;
  document.getElementById('percentFiltrados').textContent = percentFiltrados + '%';
}

// ===================================
// ATUALIZAR GRÁFICOS
// ===================================
function updateCharts() {
  // Pendências Não Resolvidas por Distrito
  const distritosCountPendentes = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    if (item['_tipo'] !== 'PENDENTE') return;
    const distrito = item['_distrito'] || 'Não informado';
    distritosCountPendentes[distrito] = (distritosCountPendentes[distrito] || 0) + 1;
  });

  const distritosLabelsPendentes = Object.keys(distritosCountPendentes)
    .sort((a, b) => distritosCountPendentes[b] - distritosCountPendentes[a]);

  const distritosValuesPendentes = distritosLabelsPendentes.map(label => distritosCountPendentes[label]);
  createDistritoPendenteChart('chartDistritosPendentes', distritosLabelsPendentes, distritosValuesPendentes);

  // Registros de Pendências Resolvidas por Distrito
  const distritosCountResolvidas = {};
  filteredData.forEach(item => {
    if (item['_tipo'] !== 'RESOLVIDO') return;
    if (!hasUsuarioPreenchido(item)) return;
    const distrito = item['_distrito'] || 'Não informado';
    distritosCountResolvidas[distrito] = (distritosCountResolvidas[distrito] || 0) + 1;
  });

  const distritosLabelsResolvidas = Object.keys(distritosCountResolvidas)
    .sort((a, b) => distritosCountResolvidas[b] - distritosCountResolvidas[a]);

  const distritosValuesResolvidas = distritosLabelsResolvidas.map(label => distritosCountResolvidas[label]);
  createDistritoResolvidasChart('chartDistritos', distritosLabelsResolvidas, distritosValuesResolvidas);

  createResolutividadeDistritoChart();

  // GRÁFICO DE STATUS
  const statusCount = {
    'RESOLVIDOS': 0,
    'PENDENTES': 0,
    'CANCELADO': 0,
    'AGENDADO': 0,
    'CANCELADO/VENCIMENTO DO PRAZO': 0
  };

  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const status = getColumnValue(item, ['Status', 'STATUS', 'status'], '').trim().toUpperCase();

    if (item['_tipo'] === 'RESOLVIDO') {
      statusCount['RESOLVIDOS']++;
    }

    if (item['_tipo'] === 'PENDENTE') {
      statusCount['PENDENTES']++;
    }

    if (item['_tipo'] === 'RESOLVIDO' && status === 'CANCELADO') {
      statusCount['CANCELADO']++;
    }

    if (item['_tipo'] === 'RESOLVIDO' && (status === 'AGENDADO' || status === 'AGENDADA')) {
      statusCount['AGENDADO']++;
    }

    if (item['_tipo'] === 'RESOLVIDO' && status === 'CANCELADO/VENCIMENTO DO PRAZO') {
      statusCount['CANCELADO/VENCIMENTO DO PRAZO']++;
    }
  });

  const statusLabels = Object.keys(statusCount);
  const statusValues = statusLabels.map(label => statusCount[label]);
  createStatusChart('chartStatus', statusLabels, statusValues);

  // Evolução temporal
  const evoCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const canceladoInfo = getCanceladoPorVencimentoInfo(item);
    let dataParaGrafico = null;

    if (canceladoInfo.isCancelado) dataParaGrafico = canceladoInfo.dataVencimento;
    else {
      dataParaGrafico = parseDate(getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ]));
    }

    if (!dataParaGrafico) return;

    const y = dataParaGrafico.getFullYear();
    const m = String(dataParaGrafico.getMonth() + 1).padStart(2, '0');
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

  // ESPECIALIDADES - TOP 10
  const especialidadesCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const especialidade = getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO']);
    if (especialidade && especialidade !== '-') {
      especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    }
  });

  const especialidadesLabels = Object.keys(especialidadesCount).sort((a, b) => especialidadesCount[b] - especialidadesCount[a]).slice(0, 10);
  const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);
  createEspecialidadeChart('chartEspecialidades', especialidadesLabels, especialidadesValues);

  const especialidadesCountPendentes = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    if (item['_tipo'] !== 'PENDENTE') return;
    const especialidade = getColumnValue(item, ['Cbo Especialidade', 'CBO Especialidade', 'CBO', 'Especialidade', 'Especialidade CBO']);
    if (especialidade && especialidade !== '-') {
      especialidadesCountPendentes[especialidade] = (especialidadesCountPendentes[especialidade] || 0) + 1;
    }
  });

  const especialidadesLabelsPendentes = Object.keys(especialidadesCountPendentes).sort((a, b) => especialidadesCountPendentes[b] - especialidadesCountPendentes[a]).slice(0, 10);
  const especialidadesValuesPendentes = especialidadesLabelsPendentes.map(label => especialidadesCountPendentes[label]);
  createEspecialidadePendenteChart('chartEspecialidadesPendentes', especialidadesLabelsPendentes, especialidadesValuesPendentes);

  // Prestadores
  const prestadoresCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;
    const prestador = item['Prestador'] || 'Não informado';
    prestadoresCount[prestador] = (prestadoresCount[prestador] || 0) + 1;
  });

  const prestadoresLabels = Object.keys(prestadoresCount).sort((a, b) => prestadoresCount[b] - prestadoresCount[a]).slice(0, 50);
  const prestadoresValues = prestadoresLabels.map(label => prestadoresCount[label]);
  createPrestadorChart('chartPrestadores', prestadoresLabels, prestadoresValues);

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

  createPieChart('chartPizzaStatus', statusLabels, statusValues);

  // Pendências por mês
  const mesCount = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const canceladoInfo = getCanceladoPorVencimentoInfo(item);
    let dataParaMes = null;

    if (canceladoInfo.isCancelado) dataParaMes = canceladoInfo.dataVencimento;
    else {
      dataParaMes = parseDate(getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ]));
    }

    if (!dataParaMes) return;

    const y = dataParaMes.getFullYear();
    const m = String(dataParaMes.getMonth() + 1).padStart(2, '0');
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
// GRÁFICO: Pendências Não Resolvidas por Distrito (ESTILO VERTICAL)
// ===================================
function createDistritoPendenteChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartDistritosPendentes) chartDistritosPendentes.destroy();

  chartDistritosPendentes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#dc2626', // Vermelho similar à imagem
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.7,
        categoryPercentage: 0.8
      }]
    },
    options: {
      indexAxis: 'x', // ✅ BARRAS VERTICAIS
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: {
          ticks: {
            font: { size: 12, weight: 'bold' },
            color: '#991b1b'
          },
          grid: { display: false },
          border: { display: false }
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
        ctx.fillStyle = '#ffffff'; // ✅ RÓTULOS BRANCOS
        ctx.font = 'bold 18px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          if (value <= 0) return;
          
          const text = `${value}`;
          const yPos = bar.y + (bar.height / 2); // ✅ NO MEIO DA BARRA
          
          ctx.fillText(text, bar.x, yPos);
        });

        ctx.restore();
      }
    }]
  });
}


// ===================================
// GRÁFICO: Pendências Resolvidas por Distrito (VERDE MAIS ESCURO)
// ===================================
function createDistritoResolvidasChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartDistritosResolvidas) chartDistritosResolvidas.destroy();

  chartDistritosResolvidas = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#059669', // VERDE ESCURO
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.7,
        categoryPercentage: 0.8
      }]
    },
    options: {
      indexAxis: 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: {
          ticks: {
            font: { size: 12, weight: 'bold' },
            color: '#059669' // VERDE ESCURO
          },
          grid: { display: false },
          border: { display: false }
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
      id: 'distritoResolvidasInsideLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 18px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          if (value <= 0) return;
          
          const text = `${value}`;
          const yPos = bar.y + (bar.height / 2);
          
          ctx.fillText(text, bar.x, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// GRÁFICO: Registros de Pendências por Status
// ===================================
function createStatusChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartStatus) chartStatus.destroy();

  chartStatus = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#f97316', // LARANJA
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.7,
        categoryPercentage: 0.8
      }]
    },
    options: {
      indexAxis: 'x', // BARRAS VERTICAIS
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false }
      },
      scales: {
        x: {
          ticks: {
            font: { size: 12, weight: 'bold' },
            color: '#f97316'
          },
          grid: { display: false },
          border: { display: false }
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
        ctx.fillStyle = '#ffffff'; // RÓTULOS BRANCOS
        ctx.font = 'bold 18px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          if (value <= 0) return;
          
          const text = `${value}`;
          const yPos = bar.y + (bar.height / 2); // NO MEIO DA BARRA
          
          ctx.fillText(text, bar.x, yPos);
        });

        ctx.restore();
      }
    }]
  });
}
// ===================================
// GRÁFICO: Evolução Temporal
// ===================================
function createEvolucaoTemporalChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartEvolucaoTemporal) chartEvolucaoTemporal.destroy();

  chartEvolucaoTemporal = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Pendências',
        data,
        backgroundColor: 'rgba(59, 130, 246, 0.2)',
        borderColor: '#3b82f6',
        borderWidth: 3,
        fill: true,
        tension: 0.4,
        pointRadius: 5,
        pointBackgroundColor: '#3b82f6',
        pointBorderColor: '#ffffff',
        pointBorderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { font: { size: 14, weight: 'bold' }, color: '#1e3a8a' } },
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
        x: {
          ticks: { font: { size: 12 }, color: '#4a5568' },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { size: 12 }, color: '#4a5568' },
          grid: { color: 'rgba(0,0,0,0.06)' }
        }
      }
    }
  });
}

// ===================================
// Total de Pendências por Mês (AZUL ESCURO, RÓTULOS BRANCOS NO MEIO)
// ===================================
function createPendenciasPorMesChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartPendenciasPorMes) chartPendenciasPorMes.destroy();

  chartPendenciasPorMes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#1e3a8a', // AZUL ESCURO
        borderWidth: 0,
        borderRadius: 6
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
          ticks: { font: { size: 11 }, color: '#4a5568' }, 
          grid: { display: false } 
        },
        y: { 
          beginAtZero: true, 
          ticks: { display: false },
          grid: { display: false } 
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
        ctx.fillStyle = '#ffffff'; // RÓTULOS BRANCOS
        ctx.font = 'bold 16px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((bar, i) => {
          const value = dataset.data[i];
          const text = `${value}`;
          const yPos = bar.y + (bar.height / 2); // NO MEIO DA BARRA
          ctx.fillText(text, bar.x, yPos);
        });

        ctx.restore();
      }
    }]
  });
}

// ===================================
// GRÁFICO: Registros Geral de Pendências por Especialidade TOP 10
// ===================================
function createEspecialidadeChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartEspecialidades) chartEspecialidades.destroy();

  chartEspecialidades = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#1e3a8a',
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.85,
        categoryPercentage: 0.9
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
      id: 'especialidadeInsideLabels',
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
// GRÁFICO: Pendências Não Resolvidas por Especialidade (TOP 10)
// ===================================
function createEspecialidadePendenteChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartEspecialidadesPendentes) chartEspecialidadesPendentes.destroy();

  chartEspecialidadesPendentes = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '',
        data,
        backgroundColor: '#991b1b',
        borderWidth: 0,
        borderRadius: 6,
        barPercentage: 0.85,
        categoryPercentage: 0.9
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
            color: '#991b1b'
          },
          grid: { display: false },
          border: { display: false }
        }
      }
    },
    plugins: [{
      id: 'especialidadePendenteInsideLabels',
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
// GRÁFICO: Registros Geral de Pendências por Prestador
// ===================================
function createPrestadorChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
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
  if (!ctx) return;
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

// ===================================
// GRÁFICO: Resolutividade por Distrito
// ===================================
function createResolutividadeDistritoChart() {
  const ctx = document.getElementById('chartResolutividadeDistrito');
  if (!ctx) return;

  const distritoStats = {};
  filteredData.forEach(item => {
    if (!hasUsuarioPreenchido(item)) return;

    const distrito = item['_distrito'] || 'Não informado';
    if (!distritoStats[distrito]) distritoStats[distrito] = { total: 0, resolvidos: 0 };

    distritoStats[distrito].total++;
    if (item['_tipo'] === 'RESOLVIDO') distritoStats[distrito].resolvidos++;
  });

  const labels = Object.keys(distritoStats).sort();
  const percentuais = labels.map(d => {
    const s = distritoStats[d];
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
              const distrito = context.label;
              const stats = distritoStats[distrito];
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
// GRÁFICO: Resolutividade por Prestador
// ===================================
function createResolutividadePrestadorChart() {
  const ctx = document.getElementById('chartResolutividadePrestador');
  if (!ctx) return;

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
// GRÁFICO DE ROSCA (DOUGHNUT)
// ===================================
function createPieChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  if (chartPizzaStatus) chartPizzaStatus.destroy();

  const colorMap = {
    'RESOLVIDOS': '#10b981',
    'PENDENTES': '#3b82f6',
    'CANCELADO': '#ef4444',
    'AGENDADO': '#f59e0b',
    'CANCELADO/VENCIMENTO DO PRAZO': '#9333ea'
  };

  const colors = labels.map(label => colorMap[label] || '#8b5cf6');
  const total = data.reduce((acc, v) => acc + v, 0);

  chartPizzaStatus = new Chart(ctx, {
    type: 'doughnut',
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
                  const percent = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
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
              const percent = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
              return `${label}: ${value} (${percent}%)`;
            }
          }
        }
      },
      cutout: '62%'
    },
    plugins: [{
      id: 'doughnutInsideLabelsCentered',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        const dataset = chart.data.datasets[0];
        if (!meta || !meta.data) return;

        ctx.save();
        ctx.font = 'bold 14px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        meta.data.forEach((arc, i) => {
          const value = dataset.data[i];
          if (!value || value <= 0) return;

          const percent = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';

          if (typeof arc.getCenterPoint === 'function') {
            const p = arc.getCenterPoint();
            ctx.fillStyle = '#ffffff';
            ctx.shadowColor = 'rgba(0,0,0,0.45)';
            ctx.shadowBlur = 4;
            ctx.shadowOffsetX = 1;
            ctx.shadowOffsetY = 1;
            ctx.fillText(`${percent}%`, p.x, p.y);
            return;
          }

          const midAngle = (arc.startAngle + arc.endAngle) / 2;
          const radius = (arc.outerRadius + arc.innerRadius) / 2;
          const x = arc.x + Math.cos(midAngle) * radius;
          const y = arc.y + Math.sin(midAngle) * radius;

          ctx.fillStyle = '#ffffff';
          ctx.shadowColor = 'rgba(0,0,0,0.45)';
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

    // "Data Envio Email (15/30)" deve vir APENAS da planilha
    const email15Planilha = getColumnValue(item, [
      'Data Envio Email (15 dias)',
      'Data Envio Email 15 dias',
      'Data do envio do Email (Prazo: Pendência com 15 dias)',
      'Data do envio do Email (Prazo Pendência com 15 dias)',
      'Data envio Email 15 dias',
      'Email 15 dias'
    ], '-');

    const email30Planilha = getColumnValue(item, [
      'Data Envio Email (30 dias)',
      'Data Envio Email 30 dias',
      'Data do envio do Email (Prazo: Pendência com 30 dias)',
      'Data do envio do Email (Prazo Pendência com 30 dias)',
      'Data envio Email 30 dias',
      'Email 30 dias'
    ], '-');

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
      email15: formatDate(email15Planilha), // apenas planilha
      prazo30: prazos.prazo30,
      email30: formatDate(email30Planilha), // apenas planilha

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

    // Destacar em amarelo APENAS:
    // - registros da aba RESOLVIDOS
    // - com Usuário preenchido (baseItems já filtra por isso)
    // - com 26 dias ou mais desde a "Data Início da Pendência"
    if (r._item['_tipo'] === 'RESOLVIDO' && r._dataInicio) {
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

