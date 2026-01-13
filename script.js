// ===================================
// CONFIGURAÇÃO DAS PLANILHAS (8 DISTRITOS)
// ===================================
const SHEETS = [
    // DISTRITO ELDORADO
    {
        name: 'PENDÊNCIAS ELDORADO',
        url: 'https://docs.google.com/spreadsheets/d/1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'ELDORADO',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS ELDORADO',
        url: 'https://docs.google.com/spreadsheets/d/1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4/gviz/tq?tqx=out:csv&gid=2142054254',
        distrito: 'ELDORADO',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO INDUSTRIAL
    {
        name: 'PENDÊNCIAS INDUSTRIAL',
        url: 'https://docs.google.com/spreadsheets/d/14eUVIsWPubMve4DhVjVwlh7gin-qVyN3PspkwQ1PZMg/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'INDUSTRIAL',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS INDUSTRIAL',
        url: 'https://docs.google.com/spreadsheets/d/14eUVIsWPubMve4DhVjVwlh7gin-qVyN3PspkwQ1PZMg/gviz/tq?tqx=out:csv&gid=1086207100',
        distrito: 'INDUSTRIAL',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO NACIONAL
    {
        name: 'PENDÊNCIAS NACIONAL',
        url: 'https://docs.google.com/spreadsheets/d/1lMGO9Hh_qL9OKI270fPL7lxadr-BZN9x_ZtmQeX6OcA/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'NACIONAL',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS NACIONAL',
        url: 'https://docs.google.com/spreadsheets/d/1lMGO9Hh_qL9OKI270fPL7lxadr-BZN9x_ZtmQeX6OcA/gviz/tq?tqx=out:csv&gid=150768142',
        distrito: 'NACIONAL',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO PETROLÂNDIA
    {
        name: 'PENDÊNCIAS PETROLÂNDIA',
        url: 'https://docs.google.com/spreadsheets/d/1Z9Uf5MGm5tClVDR95SUpwOjivAdqEVUfDj7mIuRLf4s/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'PETROLÂNDIA',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS PETROLÂNDIA',
        url: 'https://docs.google.com/spreadsheets/d/1Z9Uf5MGm5tClVDR95SUpwOjivAdqEVUfDj7mIuRLf4s/gviz/tq?tqx=out:csv&gid=1067061018',
        distrito: 'PETROLÂNDIA',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO RESSACA
    {
        name: 'PENDÊNCIAS RESSACA',
        url: 'https://docs.google.com/spreadsheets/d/1aIsq1a8Lb90M19TQdiJG_WyX7wzzC2WRohelJY6A-u8/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'RESSACA',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS RESSACA',
        url: 'https://docs.google.com/spreadsheets/d/1aIsq1a8Lb90M19TQdiJG_WyX7wzzC2WRohelJY6A-u8/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'RESSACA',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO RIACHO
    {
        name: 'PENDÊNCIAS RIACHO',
        url: 'https://docs.google.com/spreadsheets/d/1367XyjVDYyDWo3vUz6Hd_zEqLAJkH_c1MwlvtZnpmUc/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'RIACHO',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS RIACHO',
        url: 'https://docs.google.com/spreadsheets/d/1367XyjVDYyDWo3vUz6Hd_zEqLAJkH_c1MwlvtZnpmUc/gviz/tq?tqx=out:csv&gid=1996983614',
        distrito: 'RIACHO',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO SEDE
    {
        name: 'PENDÊNCIAS SEDE',
        url: 'https://docs.google.com/spreadsheets/d/1RPf2bfQVoM1FqnyA-0P8uPTJ_PG4I2Ce6lXnk54ixfc/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'SEDE',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS SEDE',
        url: 'https://docs.google.com/spreadsheets/d/1RPf2bfQVoM1FqnyA-0P8uPTJ_PG4I2Ce6lXnk54ixfc/gviz/tq?tqx=out:csv&gid=626867102',
        distrito: 'SEDE',
        tipo: 'RESOLVIDO'
    },
    // DISTRITO VARGEM DAS FLORES
    {
        name: 'PENDÊNCIAS VARGEM DAS FLORES',
        url: 'https://docs.google.com/spreadsheets/d/1IHknmxe3xAnfy5Bju_23B5ivIL-qMaaE6q_HuPaLBpk/gviz/tq?tqx=out:csv&gid=278071504',
        distrito: 'VARGEM DAS FLORES',
        tipo: 'PENDENTE'
    },
    {
        name: 'RESOLVIDOS VARGEM DAS FLORES',
        url: 'https://docs.google.com/spreadsheets/d/1IHknmxe3xAnfy5Bju_23B5ivIL-qMaaE6q_HuPaLBpk/gviz/tq?tqx=out:csv&gid=451254610',
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
let chartEspecialidades = null;
let chartStatus = null;
let chartPrestadores = null;
let chartPrestadoresPendentes = null;
let chartPizzaStatus = null;
let chartResolutividadeDistrito = null;
let chartResolutividadePrestador = null;

// ===================================
// ✅ FUNÇÃO AUXILIAR PARA VERIFICAR SE USUÁRIO ESTÁ PREENCHIDO
// ===================================
function hasUsuarioPreenchido(item) {
    const usuario = getColumnValue(item, ['Usuário', 'Usuario', 'USUÁRIO', 'USUARIO']);
    return usuario && usuario !== '-' && usuario.trim() !== '';
}

// ===================================
// FUNÇÃO AUXILIAR PARA BUSCAR VALOR DE COLUNA
// ===================================
function getColumnValue(item, possibleNames, defaultValue = '-') {
    for (let name of possibleNames) {
        if (item.hasOwnProperty(name) && item[name]) {
            return item[name];
        }
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
    console.log('Iniciando carregamento de dados...');
    loadData();
});

// ===================================
// ✅ CARREGAR DADOS DE TODAS AS PLANILHAS
// ===================================
async function loadData() {
    showLoading(true);
    allData = [];
    
    try {
        console.log(`Carregando dados de ${SHEETS.length} abas (8 distritos)...`);

        // ✅ CARREGAR TODAS AS ABAS EM PARALELO
        const promises = SHEETS.map(sheet => 
            fetch(sheet.url)
                .then(response => {
                    if (!response.ok) {
                        console.warn(`Erro HTTP na aba "${sheet.name}": ${response.status}`);
                        return null;
                    }
                    return response.text();
                })
                .then(csvText => {
                    if (!csvText) return null;
                    console.log(`Dados CSV da aba "${sheet.name}" recebidos`);
                    return { name: sheet.name, csv: csvText, distrito: sheet.distrito, tipo: sheet.tipo };
                })
                .catch(error => {
                    console.warn(`Erro ao carregar "${sheet.name}":`, error);
                    return null;
                })
        );

        const results = await Promise.all(promises);

        // ✅ PROCESSAR CADA ABA
        results.forEach(result => {
            if (!result) return;

            const rows = parseCSV(result.csv);

            if (rows.length < 2) {
                console.warn(`Aba "${result.name}" está vazia ou sem dados`);
                return;
            }

            const headers = rows[0];
            console.log(`Cabeçalhos da aba "${result.name}":`, headers);

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

            console.log(`${sheetData.length} registros carregados da aba "${result.name}"`);
            allData.push(...sheetData);
        });

        console.log(`✅ Total de registros carregados (todos os distritos): ${allData.length}`);

        if (allData.length === 0) {
            throw new Error('Nenhum dado foi carregado das planilhas');
        }

        filteredData = [...allData];

        populateFilters();
        updateDashboard();

        console.log('✅ Dados carregados com sucesso!');

    } catch (error) {
        console.error('❌ Erro ao carregar dados:', error);
        alert(`Erro ao carregar dados das planilhas: ${error.message}\n\nVerifique:\n1. As planilhas estão públicas?\n2. Os IDs das abas estão corretos?\n3. Há dados nas planilhas?`);
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
            if (char === '\r' && nextChar === '\n') {
                i++;
            }
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
// ✅ POPULAR FILTROS (MULTISELECT + MÊS + DISTRITO)
// ===================================
function populateFilters() {
    const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
    renderMultiSelect('msStatusPanel', statusList, applyFilters);

    const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
    renderMultiSelect('msUnidadePanel', unidades, applyFilters);

    const especialidades = [...new Set(allData.map(item => item['Cbo Especialidade']))].filter(Boolean).sort();
    renderMultiSelect('msEspecialidadePanel', especialidades, applyFilters);

    const prestadores = [...new Set(allData.map(item => item['Prestador']))].filter(Boolean).sort();
    renderMultiSelect('msPrestadorPanel', prestadores, applyFilters);

    setMultiSelectText('msStatusText', [], 'Todos');
    setMultiSelectText('msUnidadeText', [], 'Todas');
    setMultiSelectText('msEspecialidadeText', [], 'Todas');
    setMultiSelectText('msPrestadorText', [], 'Todos');

    // ✅ POPULAR FILTRO DE MÊS (MULTI-SELECT)
    populateMonthFilter();

    // ✅ POPULAR FILTRO DE DISTRITO (MULTI-SELECT)
    populateDistritoFilter();
}

// ===================================
// ✅ POPULAR FILTRO DE MÊS (MULTI-SELECT)
// ===================================
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

// ===================================
// ✅ POPULAR FILTRO DE DISTRITO (MULTI-SELECT)
// ===================================
function populateDistritoFilter() {
    const distritos = [...new Set(allData.map(item => item['_distrito']))].filter(Boolean).sort();
    renderMultiSelect('msDistritoPanel', distritos, applyFilters);
    setMultiSelectText('msDistritoText', [], 'Todos os Distritos');
}

// ===================================
// ✅ APLICAR FILTROS (INCLUINDO MÊS E DISTRITO)
// ===================================
function applyFilters() {
    const statusSel = getSelectedFromPanel('msStatusPanel');
    const unidadeSel = getSelectedFromPanel('msUnidadePanel');
    const especialidadeSel = getSelectedFromPanel('msEspecialidadePanel');
    const prestadorSel = getSelectedFromPanel('msPrestadorPanel');
    const mesSel = getSelectedFromPanel('msMesPanel');
    const distritoSel = getSelectedFromPanel('msDistritoPanel');

    setMultiSelectText('msStatusText', statusSel, 'Todos');
    setMultiSelectText('msUnidadeText', unidadeSel, 'Todas');
    setMultiSelectText('msEspecialidadeText', especialidadeSel, 'Todas');
    setMultiSelectText('msPrestadorText', prestadorSel, 'Todos');
    setMultiSelectText('msMesText', mesSel, 'Todos os Meses');
    setMultiSelectText('msDistritoText', distritoSel, 'Todos os Distritos');

    filteredData = allData.filter(item => {
        const okStatus = (statusSel.length === 0) || statusSel.includes(item['Status'] || '');
        const okUnidade = (unidadeSel.length === 0) || unidadeSel.includes(item['Unidade Solicitante'] || '');
        const okEsp = (especialidadeSel.length === 0) || especialidadeSel.includes(item['Cbo Especialidade'] || '');
        const okPrest = (prestadorSel.length === 0) || prestadorSel.includes(item['Prestador'] || '');

        // ✅ FILTRO POR MÊS
        let okMes = true;
        if (mesSel.length > 0) {
            const dataInicio = parseDate(getColumnValue(item, [
                'Data Início da Pendência',
                'Data Inicio da Pendencia',
                'Data Início Pendência',
                'Data Inicio Pendencia'
            ]));
            if (dataInicio) {
                const [ano, mes] = [dataInicio.getFullYear(), dataInicio.getMonth()];
                const nomeMes = new Date(ano, mes).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
                const mesFormatado = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
                okMes = mesSel.includes(mesFormatado);
            } else {
                okMes = false;
            }
        }

        // ✅ FILTRO POR DISTRITO
        const okDistrito = (distritoSel.length === 0) || distritoSel.includes(item['_distrito'] || '');

        return okStatus && okUnidade && okEsp && okPrest && okMes && okDistrito;
    });

    updateDashboard();
}

// ===================================
// ✅ LIMPAR FILTROS (INCLUINDO MÊS E DISTRITO)
// ===================================
function clearFilters() {
    ['msStatusPanel','msUnidadePanel','msEspecialidadePanel','msPrestadorPanel','msMesPanel','msDistritoPanel'].forEach(panelId => {
        const panel = document.getElementById(panelId);
        if (!panel) return;
        panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
    });

    setMultiSelectText('msStatusText', [], 'Todos');
    setMultiSelectText('msUnidadeText', [], 'Todas');
    setMultiSelectText('msEspecialidadeText', [], 'Todas');
    setMultiSelectText('msPrestadorText', [], 'Todos');
    setMultiSelectText('msMesText', [], 'Todos os Meses');
    setMultiSelectText('msDistritoText', [], 'Todos os Distritos');

    filteredData = [...allData];
    updateDashboard();
}

// ===================================
// ATUALIZAR DASHBOARD
// ===================================
function updateDashboard() {
    updateCards();
    updateCharts();
}

// ===================================
// ✅ ATUALIZAR CARDS - BASEADO EM "USUÁRIO" PREENCHIDO
// ===================================
function updateCards() {
    // ✅ CONTA APENAS REGISTROS COM COLUNA "USUÁRIO" PREENCHIDA
    const totalComUsuario = allData.filter(item => hasUsuarioPreenchido(item)).length;
    const filtradoComUsuario = filteredData.filter(item => hasUsuarioPreenchido(item)).length;

    const hoje = new Date();
    let pendencias15 = 0;
    let pendencias30 = 0;

    filteredData.forEach(item => {
        // ✅ SÓ CONTA SE USUÁRIO ESTIVER PREENCHIDO
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
// ✅ ATUALIZAR GRÁFICOS - BASEADO EM "USUÁRIO" PREENCHIDO
// ===================================
function updateCharts() {
    // ✅ GRÁFICO PENDÊNCIAS POR DISTRITO (TODOS OS REGISTROS)
    const distritosCount = {};
    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;

        const distrito = item['_distrito'] || 'Não informado';
        distritosCount[distrito] = (distritosCount[distrito] || 0) + 1;
    });

    const distritosLabels = Object.keys(distritosCount)
        .sort((a, b) => distritosCount[b] - distritosCount[a]);
    const distritosValues = distritosLabels.map(label => distritosCount[label]);

    createDistritoChart('chartDistritos', distritosLabels, distritosValues);

    // ✅ GRÁFICO PENDÊNCIAS NÃO RESOLVIDAS POR DISTRITO (BARRAS VERMELHAS)
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

    // ✅ GRÁFICO DE RESOLUTIVIDADE POR DISTRITO
    createResolutividadeDistritoChart();

    // ✅ GRÁFICO DE STATUS (BARRAS LARANJA VERTICAIS)
    const statusCount = {};
    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;

        const status = item['Status'] || 'Não informado';
        statusCount[status] = (statusCount[status] || 0) + 1;
    });

    const statusLabels = Object.keys(statusCount)
        .sort((a, b) => statusCount[b] - statusCount[a]);
    const statusValues = statusLabels.map(label => statusCount[label]);

    createStatusChart('chartStatus', statusLabels, statusValues);

    // ✅ GRÁFICO DE ESPECIALIDADES (BARRAS HORIZONTAIS - ALTERADO)
    const especialidadesCount = {};
    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;

        const especialidade = item['Cbo Especialidade'] || 'Não informado';
        especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    });

    const especialidadesLabels = Object.keys(especialidadesCount)
        .sort((a, b) => especialidadesCount[b] - especialidadesCount[a])
        .slice(0, 50);
    const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);

    createEspecialidadeChart('chartEspecialidades', especialidadesLabels, especialidadesValues);

    // ✅ GRÁFICO PENDÊNCIAS POR PRESTADOR (TODOS OS REGISTROS)
    const prestadoresCount = {};
    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;

        const prestador = item['Prestador'] || 'Não informado';
        prestadoresCount[prestador] = (prestadoresCount[prestador] || 0) + 1;
    });

    const prestadoresLabels = Object.keys(prestadoresCount)
        .sort((a, b) => prestadoresCount[b] - prestadoresCount[a])
        .slice(0, 50);
    const prestadoresValues = prestadoresLabels.map(label => prestadoresCount[label]);

    createPrestadorChart('chartPrestadores', prestadoresLabels, prestadoresValues);

    // ✅ GRÁFICO PENDÊNCIAS NÃO RESOLVIDAS POR PRESTADOR (BARRAS VERDE ESCURO)
    const prestadoresCountPendentes = {};
    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;
        if (item['_tipo'] !== 'PENDENTE') return;

        const prestador = item['Prestador'] || 'Não informado';
        prestadoresCountPendentes[prestador] = (prestadoresCountPendentes[prestador] || 0) + 1;
    });

    const prestadoresLabelsPendentes = Object.keys(prestadoresCountPendentes)
        .sort((a, b) => prestadoresCountPendentes[b] - prestadoresCountPendentes[a])
        .slice(0, 50);
    const prestadoresValuesPendentes = prestadoresLabelsPendentes.map(label => prestadoresCountPendentes[label]);

    createPrestadorPendenteChart('chartPrestadoresPendentes', prestadoresLabelsPendentes, prestadoresValuesPendentes);

    // ✅ GRÁFICO DE RESOLUTIVIDADE POR PRESTADOR
    createResolutividadePrestadorChart();

    // ✅ GRÁFICO DE PIZZA (TODOS OS REGISTROS)
    createPieChart('chartPizzaStatus', statusLabels, statusValues);
}

// ===================================
// ✅ CRIAR GRÁFICO PENDÊNCIAS POR DISTRITO (AZUL ESCURO, VERTICAL)
// ===================================
function createDistritoChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);

    if (chartDistritos) chartDistritos.destroy();

    chartDistritos = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Pendências',
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
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#1e3a8a'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(30, 58, 138, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8
                }
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
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'distritoValueLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO PENDÊNCIAS NÃO RESOLVIDAS POR DISTRITO (BARRAS VERMELHAS)
// ===================================
function createDistritoPendenteChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);

    if (chartDistritosPendentes) chartDistritosPendentes.destroy();

    chartDistritosPendentes = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Pendências Não Resolvidas',
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
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#dc2626'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(220, 38, 38, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8
                }
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
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'distritoPendenteValueLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO DE RESOLUTIVIDADE POR DISTRITO (PORCENTAGEM DENTRO DA BARRA - BRANCO E NEGRITO)
// ===================================
function createResolutividadeDistritoChart() {
    const ctx = document.getElementById('chartResolutividadeDistrito');

    // Calcular resolutividade por distrito
    const distritosStats = {};

    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;

        const distrito = item['_distrito'] || 'Não informado';
        
        if (!distritosStats[distrito]) {
            distritosStats[distrito] = { total: 0, resolvidos: 0 };
        }

        distritosStats[distrito].total++;
        
        if (item['_tipo'] === 'RESOLVIDO') {
            distritosStats[distrito].resolvidos++;
        }
    });

    // Calcular porcentagens
    const labels = Object.keys(distritosStats).sort((a, b) => {
        const percA = (distritosStats[a].resolvidos / distritosStats[a].total) * 100;
        const percB = (distritosStats[b].resolvidos / distritosStats[b].total) * 100;
        return percB - percA;
    });

    const percentuais = labels.map(distrito => {
        const stats = distritosStats[distrito];
        return stats.total > 0 ? ((stats.resolvidos / stats.total) * 100).toFixed(1) : 0;
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
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#059669'
                    }
                },
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
                        callback: function(value) {
                            return value + '%';
                        }
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                },
                y: {
                    ticks: {
                        font: { size: 13, weight: 'bold' },
                        color: '#059669'
                    },
                    grid: { display: false }
                }
            }
        },
        plugins: [{
            id: 'resolutividadeDistritoLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();

                meta.data.forEach((bar, i) => {
                    const percentValue = dataset.data[i];

                    // ✅ VALOR PERCENTUAL FIXO DENTRO DA BARRA (BRANCO E NEGRITO)
                    ctx.fillStyle = '#FFFFFF';
                    ctx.font = 'bold 16px Arial';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    
                    const xPos = bar.x + (bar.width / 2);
                    ctx.fillText(`${percentValue}%`, xPos, bar.y);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO DE STATUS (BARRAS LARANJA VERTICAIS)
// ===================================
function createStatusChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);

    if (chartStatus) chartStatus.destroy();

    chartStatus = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Registros',
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
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#f97316'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(249, 115, 22, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8
                }
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
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'statusValueLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO ESPECIALIDADES (BARRAS HORIZONTAIS)
// ===================================
function createEspecialidadeChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);

    if (chartEspecialidades) chartEspecialidades.destroy();

    chartEspecialidades = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Registros',
                data,
                backgroundColor: '#ef4444',
                borderWidth: 0,
                borderRadius: 8,
                barPercentage: 0.65,
                categoryPercentage: 0.75
            }]
        },
        options: {
            indexAxis: 'y', // ✅ BARRAS HORIZONTAIS
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#ef4444'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(239, 68, 68, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                },
                y: {
                    ticks: {
                        font: { size: 13, weight: 'bold' },
                        color: '#ef4444'
                    },
                    grid: { display: false }
                }
            }
        },
        plugins: [{
            id: 'especialidadeValueLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const xPos = bar.x + (bar.width / 2);
                    ctx.fillText(String(value), xPos, bar.y);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO PRESTADOR (ROXO - TODOS OS REGISTROS)
// ===================================
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
                borderRadius: 8,
                barPercentage: 0.65,
                categoryPercentage: 0.75
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#8b5cf6'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(139, 92, 246, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8
                }
            },
            scales: {
                x: {
                    ticks: {
                        font: { size: 13, weight: 'bold' },
                        color: '#8b5cf6',
                        maxRotation: 45,
                        minRotation: 0
                    },
                    grid: { display: false }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'prestadorValueLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO PENDÊNCIAS NÃO RESOLVIDAS POR PRESTADOR (BARRAS VERDE ESCURO)
// ===================================
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
                borderRadius: 8,
                barPercentage: 0.65,
                categoryPercentage: 0.75
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#065f46'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(6, 95, 70, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8
                }
            },
            scales: {
                x: {
                    ticks: {
                        font: { size: 13, weight: 'bold' },
                        color: '#065f46',
                        maxRotation: 45,
                        minRotation: 0
                    },
                    grid: { display: false }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'prestadorPendenteValueLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO DE RESOLUTIVIDADE POR PRESTADOR
// ===================================
function createResolutividadePrestadorChart() {
    const ctx = document.getElementById('chartResolutividadePrestador');

    // Calcular resolutividade por prestador
    const prestadoresStats = {};

    filteredData.forEach(item => {
        if (!hasUsuarioPreenchido(item)) return;

        const prestador = item['Prestador'] || 'Não informado';
        
        if (!prestadoresStats[prestador]) {
            prestadoresStats[prestador] = { total: 0, resolvidos: 0 };
        }

        prestadoresStats[prestador].total++;
        
        if (item['_tipo'] === 'RESOLVIDO') {
            prestadoresStats[prestador].resolvidos++;
        }
    });

    // Calcular porcentagens e ordenar
    const labels = Object.keys(prestadoresStats)
        .sort((a, b) => {
            const percA = (prestadoresStats[a].resolvidos / prestadoresStats[a].total) * 100;
            const percB = (prestadoresStats[b].resolvidos / prestadoresStats[b].total) * 100;
            return percB - percA;
        })
        .slice(0, 15); // Top 15 prestadores

    const percentuais = labels.map(prestador => {
        const stats = prestadoresStats[prestador];
        return stats.total > 0 ? ((stats.resolvidos / stats.total) * 100).toFixed(1) : 0;
    });

    if (chartResolutividadePrestador) chartResolutividadePrestador.destroy();

    chartResolutividadePrestador = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Taxa de Resolutividade (%)',
                data: percentuais,
                backgroundColor: '#0284c7',
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
                legend: { 
                    display: true,
                    labels: {
                        font: { size: 14, weight: 'bold' },
                        color: '#0284c7'
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(2, 132, 199, 0.9)',
                    titleFont: { size: 16, weight: 'bold' },
                    bodyFont: { size: 14 },
                    padding: 14,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const prestador = context.label;
                            const stats = prestadoresStats[prestador];
                            return [
                                `Resolutividade: ${context.parsed.y}%`,
                                `Resolvidos: ${stats.resolvidos}`,
                                `Total: ${stats.total}`
                            ];
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        font: { size: 12, weight: 'bold' },
                        color: '#0284c7',
                        maxRotation: 45,
                        minRotation: 0
                    },
                    grid: { display: false }
                },
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568',
                        callback: function(value) {
                            return value + '%';
                        }
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'resolutividadePrestadorLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 14px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(value + '%', bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ CRIAR GRÁFICO DE PIZZA
// ===================================
function createPieChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);

    if (chartPizzaStatus) chartPizzaStatus.destroy();

    const colors = [
        '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
        '#ec4899', '#06b6d4', '#f97316', '#6366f1', '#84cc16'
    ];

    const total = data.reduce((sum, val) => sum + val, 0);

    chartPizzaStatus = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: data,
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
                        font: { 
                            size: 14,
                            weight: 'bold',
                            family: 'Arial, sans-serif'
                        },
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
                const meta = chart.getDatasetMeta(0);
                
                ctx.save();
                ctx.font = 'bold 14px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                
                meta.data.forEach(function(element, index) {
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

// ===================================
// FUNÇÕES AUXILIARES
// ===================================
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

// ===================================
// ATUALIZAR DADOS
// ===================================
function refreshData() {
    loadData();
}

// ===================================
// ✅ DOWNLOAD EXCEL - MODIFICADO PARA PRIORIZAR AS 3 COLUNAS SOLICITADAS
// ===================================
function downloadExcel() {
    // ✅ EXPORTA APENAS REGISTROS COM USUÁRIO PREENCHIDO
    const dataParaExportar = filteredData.filter(item => hasUsuarioPreenchido(item));

    if (dataParaExportar.length === 0) {
        alert('Não há dados com Usuário preenchido para exportar.');
        return;
    }

    // ✅ ESTRUTURA COM AS 3 COLUNAS PRIORITÁRIAS NO INÍCIO
    const exportData = dataParaExportar.map(item => ({
        'Nº Prontuário': getColumnValue(item, ['Nº Prontuário', 'N° Prontuário', 'Numero Prontuário', 'Prontuário', 'Prontuario'], ''),
        'Nº Solicitação': getColumnValue(item, ['Solicitação', 'Solicitacao', 'N° Solicitação', 'Nº Solicitação'], ''),
        'Telefone': item['Telefone'] || '',
        'Distrito': item['_distrito'] || '',
        'Origem': item['_origem'] || '',
        'Data Solicitação': getColumnValue(item, ['Data da Solicitação', 'Data Solicitação', 'Data da Solicitacao', 'Data Solicitacao'], ''),
        'Unidade Solicitante': item['Unidade Solicitante'] || '',
        'CBO Especialidade': item['Cbo Especialidade'] || '',
        'Data Início Pendência': getColumnValue(item, ['Data Início da Pendência','Data Início Pendência','Data Inicio da Pendencia','Data Inicio Pendencia'], ''),
        'Status': item['Status'] || '',
        'Prestador': item['Prestador'] || '',
        'Usuário': getColumnValue(item, ['Usuário', 'Usuario'], ''),
        'Data Final Prazo 15d': getColumnValue(item, ['Data Final do Prazo (Pendência com 15 dias)','Data Final do Prazo (Pendencia com 15 dias)','Data Final Prazo 15d','Prazo 15 dias'], ''),
        'Data Envio Email 15d': getColumnValue(item, ['Data do envio do Email (Prazo: Pendência com 15 dias)','Data do envio do Email (Prazo: Pendencia com 15 dias)','Data Envio Email 15d','Email 15 dias'], ''),
        'Data Final Prazo 30d': getColumnValue(item, ['Data Final do Prazo (Pendência com 30 dias)','Data Final do Prazo (Pendencia com 30 dias)','Data Final Prazo 30d','Prazo 30 dias'], ''),
        'Data Envio Email 30d': getColumnValue(item, ['Data do envio do Email (Prazo: Pendência com 30 dias)','Data do envio do Email (Prazo: Pendencia com 30 dias)','Data Envio Email 30d','Email 30 dias'], '')
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dados Completos');

    // ✅ AJUSTAR LARGURA DAS COLUNAS (PRIORIZANDO AS 3 PRIMEIRAS)
    ws['!cols'] = [
        { wch: 18 },  // Nº Prontuário
        { wch: 18 },  // Nº Solicitação
        { wch: 16 },  // Telefone
        { wch: 20 },  // Distrito
        { wch: 30 },  // Origem
        { wch: 18 },  // Data Solicitação
        { wch: 30 },  // Unidade Solicitante
        { wch: 30 },  // CBO Especialidade
        { wch: 18 },  // Data Início Pendência
        { wch: 20 },  // Status
        { wch: 25 },  // Prestador
        { wch: 20 },  // Usuário
        { wch: 18 },  // Data Final Prazo 15d
        { wch: 20 },  // Data Envio Email 15d
        { wch: 18 },  // Data Final Prazo 30d
        { wch: 20 }   // Data Envio Email 30d
    ];

    const hoje = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Dados_Todos_Distritos_${hoje}.xlsx`);
}
