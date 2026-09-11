/* ==========================================================================
   PAINEL DE GESTÃO DE DEMANDAS - SEGES 2026
   ========================================================================== */

/* --- ⚙️ ESTADO GLOBAL DA APLICAÇÃO --- */
window.state = {
  rawData: [],
  filteredData: [],
  // Filtros Principais (Arrays de opções selecionadas)
  filters: {
    Unidade: [],
    Area: [],
    Status: [],
    Categoria: [],
    Contratacao: []
  },
  // Filtros do Cabeçalho da Tabela Analítica (Arrays)
  colFilters: {
    Ano: [],
    UnidadeCol: [],
    AreaCol: [],
    StatusCol: [],
    CodDemanda: [],
    Objeto: [],
    CategoriaCol: [],
    ContratacaoCol: [],
    ItemContrato: [],
    Empenho: [],
    OrdemServico: []
  }
};

// Mapeamento de chaves para os nomes de colunas no Excel
const FIELD_MAP = {
  Unidade: 'Unidade',
  Area: 'Área',
  Status: 'Status / Orçamento',
  Categoria: 'Categoria Econômica',
  Contratacao: 'Contratação',
  // Colunas Analíticas
  Ano: 'Ano Orçamento',
  UnidadeCol: 'Unidade',
  AreaCol: 'Área',
  StatusCol: 'Status / Orçamento',
  CodDemanda: 'Cód. Demanda (Siged / clarity)',
  Objeto: 'Objeto/Sistema',
  CategoriaCol: 'Categoria Econômica',
  ContratacaoCol: 'Contratação',
  ItemContrato: 'Item do Contrato',
  Empenho: 'Nº Empenho',
  OrdemServico: 'Ordem de Serviço'
};

/* --- 🛠️ FUNÇÕES UTILITÁRIAS --- */
function parseCurrency(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleanStr = val.toString().replace(/R\$\s?/g, '').replace(/\./g, '').replace(',', '.').trim();
  const parsed = parseFloat(cleanStr);
  return isNaN(parsed) ? 0 : parsed;
}

function formatBRL(valor) {
  const num = typeof valor === 'number' ? valor : parseCurrency(valor);
  return num.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

/* --- 📥 CARREGAMENTO DE DADOS (EXCEL) --- */
async function loadExcelFromRepo() {
  const filterCounter = document.getElementById('filterCounter');
  if (filterCounter) {
    filterCounter.textContent = 'Carregando dados do Excel...';
    filterCounter.className = 'text-xs font-medium text-amber-600 bg-amber-50 px-2.5 py-1 rounded-full';
  }

  try {
    const response = await fetch('./demandas.xlsx?t=' + new Date().getTime());
    if (!response.ok) {
      throw new Error(`Arquivo 'demandas.xlsx' não encontrado (Status ${response.status}).`);
    }

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonRecords = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    if (!jsonRecords || jsonRecords.length === 0) {
      throw new Error("A planilha 'demandas.xlsx' está vazia.");
    }

    // Normaliza chaves removendo espaços extras nas propriedades do Excel
    const cleanRecords = jsonRecords.filter(row => {
      const ano = (row['Ano Orçamento'] || '').toString();
      return ano !== 'Total' && !ano.includes('Filtros aplicados');
    }).map((row, index) => {
      const cleanRow = {};
      Object.keys(row).forEach(k => {
        cleanRow[k.trim()] = row[k];
      });
      return {
        ...cleanRow,
        __id: index,
        'Valor Estimado': parseCurrency(cleanRow['Valor Estimado'])
      };
    });

    initData(cleanRecords);

  } catch (error) {
    console.error("Erro ao carregar Excel:", error);
    if (filterCounter) {
      filterCounter.textContent = 'Erro ao carregar dados do demandas.xlsx';
      filterCounter.className = 'text-xs font-semibold text-red-700 bg-red-100 px-2.5 py-1 rounded-full';
    }
  }
}

/* --- 🔄 INICIALIZAÇÃO E POVOAMENTO DOS FILTROS --- */
function initData(records) {
  window.state.rawData = records;
  populateTopDropdowns();
  populateHeaderFilters();
  applyFilters();
}

/* Popula as opções das listas suspensas principais */
function populateTopDropdowns() {
  const filterKeys = ['Unidade', 'Area', 'Status', 'Categoria', 'Contratacao'];

  filterKeys.forEach(key => {
    const fieldName = FIELD_MAP[key];
    const container = document.getElementById(`options${key}`);
    if (!container) return;

    const uniqueValues = [...new Set(window.state.rawData.map(r => {
      const v = r[fieldName];
      return (v !== undefined && v !== null && v !== '') ? v.toString().trim() : '(Não Informado)';
    }))].sort();

    container.innerHTML = uniqueValues.map(val => `
      <label class="flex items-center gap-2 px-1 py-1 hover:bg-gray-50 rounded cursor-pointer text-xs text-gray-700">
        <input type="checkbox" value="${val}" onchange="handleTopFilterChange('${key}')"
          ${window.state.filters[key].includes(val) ? 'checked' : ''}
          class="rounded text-blue-600 focus:ring-blue-500 h-3.5 w-3.5">
        <span class="truncate">${val}</span>
      </label>
    `).join('');

    updateTopFilterLabel(key);
  });
}

/* Popula as opções dos filtros no cabeçalho da Tabela Analítica */
function populateHeaderFilters() {
  const colKeys = Object.keys(window.state.colFilters);

  colKeys.forEach(key => {
    const fieldName = FIELD_MAP[key];
    const container = document.getElementById(`colList_${key}`);
    if (!container) return;

    const uniqueValues = [...new Set(window.state.rawData.map(r => {
      const v = r[fieldName];
      return (v !== undefined && v !== null && v !== '') ? v.toString().trim() : '-';
    }))].sort();

    container.innerHTML = uniqueValues.map(val => `
      <label class="flex items-center gap-2 px-1 py-0.5 hover:bg-gray-100 rounded cursor-pointer text-xs text-gray-700">
        <input type="checkbox" value="${val}" onchange="handleColFilterChange('${key}')"
          ${window.state.colFilters[key].includes(val) ? 'checked' : ''}
          class="rounded text-blue-600 focus:ring-blue-500 h-3 w-3">
        <span class="truncate">${val}</span>
      </label>
    `).join('');
  });
}

/* --- 🎯 INTERAÇÕES E MANIPULAÇÃO DE FILTROS --- */

// Alterna abertura de menus dropdown superiores
function toggleDropdown(key) {
  const menu = document.getElementById(`menu${key}`);
  if (!menu) return;

  const isHidden = menu.classList.contains('hidden');
  closeAllDropdowns();
  if (isHidden) {
    menu.classList.remove('hidden');
  }
}

// Alterna filtros de cabeçalho da tabela analítica
function toggleHeaderFilter(key, event) {
  if (event) event.stopPropagation();
  const menu = document.getElementById(`colFilter_${key}`);
  if (!menu) return;

  const isHidden = menu.classList.contains('hidden');
  closeAllDropdowns();
  if (isHidden) {
    menu.classList.remove('hidden');
  }
}

// Fecha todos os menus abertos
function closeAllDropdowns() {
  document.querySelectorAll('.filter-dropdown-container [id^="menu"]').forEach(m => m.classList.add('hidden'));
  document.querySelectorAll('[id^="colFilter_"]').forEach(m => m.classList.add('hidden'));
}

// Fecha menus ao clicar fora
document.addEventListener('click', (e) => {
  if (!e.target.closest('.filter-dropdown-container') && !e.target.closest('th')) {
    closeAllDropdowns();
  }
});

// Pesquisa dentro do dropdown principal
function filterDropdownOptions(key) {
  const searchInput = document.getElementById(`search${key}`);
  if (!searchInput) return;
  const filter = searchInput.value.toLowerCase();
  const options = document.querySelectorAll(`#options${key} label`);

  options.forEach(opt => {
    const text = opt.textContent.toLowerCase();
    opt.style.display = text.includes(filter) ? 'flex' : 'none';
  });
}

// Pesquisa dentro do dropdown do cabeçalho
function filterColOptions(key) {
  const searchInput = document.getElementById(`colSearch_${key}`);
  if (!searchInput) return;
  const filter = searchInput.value.toLowerCase();
  const options = document.querySelectorAll(`#colList_${key} label`);

  options.forEach(opt => {
    const text = opt.textContent.toLowerCase();
    opt.style.display = text.includes(filter) ? 'flex' : 'none';
  });
}

// Atualizar seleção no estado
function handleTopFilterChange(key) {
  const container = document.getElementById(`options${key}`);
  if (!container) return;

  const checked = Array.from(container.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
  window.state.filters[key] = checked;

  updateTopFilterLabel(key);
  applyFilters();
}

function updateTopFilterLabel(key) {
  const labelEl = document.getElementById(`label${key}`);
  if (!labelEl) return;

  const selected = window.state.filters[key];
  const defaultLabels = {
    Unidade: 'Todas as Unidades',
    Area: 'Todas as Áreas',
    Status: 'Todos os Status',
    Categoria: 'Todas as Categorias',
    Contratacao: 'Todas as Contratações'
  };

  if (selected.length === 0) {
    labelEl.textContent = defaultLabels[key];
  } else if (selected.length === 1) {
    labelEl.textContent = selected[0];
  } else {
    labelEl.textContent = `${selected.length} selecionados`;
  }
}

function handleColFilterChange(key) {
  const container = document.getElementById(`colList_${key}`);
  if (!container) return;

  const checked = Array.from(container.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
  window.state.colFilters[key] = checked;

  applyFilters();
}

// Marcar / Desmarcar Todos para Filtros Principais
function selectAllOptions(key) {
  const container = document.getElementById(`options${key}`);
  if (!container) return;
  container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
  handleTopFilterChange(key);
}

function deselectAllOptions(key) {
  const container = document.getElementById(`options${key}`);
  if (!container) return;
  container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  handleTopFilterChange(key);
}

// Marcar / Desmarcar Todos para Colunas da Tabela Analítica
function selectAllCol(key) {
  const container = document.getElementById(`colList_${key}`);
  if (!container) return;
  container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
  handleColFilterChange(key);
}

function deselectAllCol(key) {
  const container = document.getElementById(`colList_${key}`);
  if (!container) return;
  container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  handleColFilterChange(key);
}

/* --- 🔍 APLICAÇÃO DE FILTROS NA BASE --- */
function applyFilters() {
  window.state.filteredData = window.state.rawData.filter(row => {
    // 1. Validação de Filtros Principais
    const matchTop = Object.keys(window.state.filters).every(key => {
      const selected = window.state.filters[key];
      if (!selected || selected.length === 0) return true;
      const fieldName = FIELD_MAP[key];
      const val = (row[fieldName] !== undefined && row[fieldName] !== null && row[fieldName] !== '')
        ? row[fieldName].toString().trim()
        : '(Não Informado)';
      return selected.includes(val);
    });

    if (!matchTop) return false;

    // 2. Validação de Filtros das Colunas da Tabela
    const matchCol = Object.keys(window.state.colFilters).every(key => {
      const selected = window.state.colFilters[key];
      if (!selected || selected.length === 0) return true;
      const fieldName = FIELD_MAP[key];
      const val = (row[fieldName] !== undefined && row[fieldName] !== null && row[fieldName] !== '')
        ? row[fieldName].toString().trim()
        : '-';
      return selected.includes(val);
    });

    return matchCol;
  });

  updateFilterCounter();
  notifyStateChange();
}

function updateFilterCounter() {
  const filterCounter = document.getElementById('filterCounter');
  if (!filterCounter) return;

  const activeTopCount = Object.values(window.state.filters).filter(a => a.length > 0).length;
  const activeColCount = Object.values(window.state.colFilters).filter(a => a.length > 0).length;
  const totalActive = activeTopCount + activeColCount;

  if (totalActive === 0) {
    filterCounter.textContent = `Mostrando todos os ${window.state.filteredData.length} registros`;
    filterCounter.className = "text-xs font-medium text-gray-500 bg-gray-100 px-2.5 py-1 rounded-full";
  } else {
    filterCounter.textContent = `${totalActive} filtro(s) ativo(s) | ${window.state.filteredData.length} registro(s)`;
    filterCounter.className = "text-xs font-semibold text-blue-700 bg-blue-100 px-2.5 py-1 rounded-full";
  }
}

function resetFilters() {
  Object.keys(window.state.filters).forEach(k => {
    window.state.filters[k] = [];
    const container = document.getElementById(`options${k}`);
    if (container) container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
    updateTopFilterLabel(k);
  });

  Object.keys(window.state.colFilters).forEach(k => {
    window.state.colFilters[k] = [];
    const container = document.getElementById(`colList_${k}`);
    if (container) container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  });

  applyFilters();
}

/* --- 📊 RENDERIZAÇÃO DA INTERFACE --- */
function notifyStateChange() {
  renderKPIs();
  renderSyntheticTable();
  renderAnalyticalTable();
}

function renderKPIs() {
  const data = window.state.filteredData;
  const totalRecords = data.length;

  const totalValue = data.reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const capitalValue = data
    .filter(row => row['Categoria Econômica'] === 'Despesas de Capital')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const currentValue = data
    .filter(row => row['Categoria Econômica'] === 'Despesas Correntes')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);

  const totalRecEl = document.getElementById('kpiTotalRecords');
  const totalValEl = document.getElementById('kpiTotalValue');
  const capValEl = document.getElementById('kpiCapitalValue');
  const curValEl = document.getElementById('kpiCurrentValue');

  if (totalRecEl) totalRecEl.textContent = totalRecords;
  if (totalValEl) totalValEl.textContent = formatBRL(totalValue);
  if (capValEl) capValEl.textContent = formatBRL(capitalValue);
  if (curValEl) curValEl.textContent = formatBRL(currentValue);
}

function renderSyntheticTable() {
  const tbody = document.getElementById('syntheticTableBody');
  const tfoot = document.getElementById('syntheticTableFooter');
  if (!tbody || !tfoot) return;

  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const data = window.state.filteredData;
  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">Nenhum dado disponível para os filtros selecionados.</td></tr>`;
    return;
  }

  const grandTotal = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);

  // Agrupamento hierárquico por Contratação e Categoria Econômica
  const grouped = {};

  data.forEach(item => {
    const contratacao = (item['Contratação'] && item['Contratação'].toString().trim()) ? item['Contratação'].toString().trim() : '(Não Informado)';
    const categoria = (item['Categoria Econômica'] && item['Categoria Econômica'].toString().trim()) ? item['Categoria Econômica'].toString().trim() : '(Não Informado)';
    const itemContrato = (item['Item do Contrato'] && item['Item do Contrato'].toString().trim()) ? item['Item do Contrato'].toString().trim() : '-';

    if (!grouped[contratacao]) grouped[contratacao] = {};
    if (!grouped[contratacao][categoria]) grouped[contratacao][categoria] = { items: {}, count: 0, total: 0 };

    if (!grouped[contratacao][categoria].items[itemContrato]) {
      grouped[contratacao][categoria].items[itemContrato] = { count: 0, total: 0 };
    }

    grouped[contratacao][categoria].items[itemContrato].count += 1;
    grouped[contratacao][categoria].items[itemContrato].total += (item['Valor Estimado'] || 0);

    grouped[contratacao][categoria].count += 1;
    grouped[contratacao][categoria].total += (item['Valor Estimado'] || 0);
  });

  // Renderiza linhas agrupadas com subtotais por categoria
  Object.keys(grouped).forEach(contratacao => {
    const categorias = grouped[contratacao];

    Object.keys(categorias).forEach(categoria => {
      const catData = categorias[categoria];
      const items = catData.items;

      // Linhas detalhadas de itens do contrato
      Object.keys(items).forEach(itemContrato => {
        const row = items[itemContrato];
        const pct = grandTotal > 0 ? (row.total / grandTotal) * 100 : 0;
        const tr = document.createElement('tr');
        tr.className = "hover:bg-gray-50 border-b border-gray-100 transition-colors";

        const badgeClass = categoria === 'Despesas de Capital' ? 'bg-purple-100 text-purple-700' :
                           categoria === 'Despesas Correntes' ? 'bg-amber-100 text-amber-700' :
                           'bg-gray-100 text-gray-600';

        tr.innerHTML = `
          <td class="px-4 py-2.5 font-medium text-gray-800">${contratacao}</td>
          <td class="px-4 py-2.5">
            <span class="inline-block px-2 py-0.5 text-xs rounded font-medium ${badgeClass}">
              ${categoria}
            </span>
          </td>
          <td class="px-4 py-2.5 text-gray-600 text-xs">${itemContrato}</td>
          <td class="px-4 py-2.5 text-center font-medium">${row.count}</td>
          <td class="px-4 py-2.5 text-right font-semibold text-gray-900">${formatBRL(row.total)}</td>
          <td class="px-4 py-2.5 text-right text-xs font-semibold text-gray-500">${pct.toFixed(2)}%</td>
        `;
        tbody.appendChild(tr);
      });

      // Linha de Subtotal por Categoria Econômica
      const subPct = grandTotal > 0 ? (catData.total / grandTotal) * 100 : 0;
      const trSub = document.createElement('tr');
      trSub.className = "bg-blue-50/40 font-semibold border-b border-blue-200 text-xs text-blue-900";
      trSub.innerHTML = `
        <td colspan="3" class="px-4 py-2 text-right">Subtotal (${contratacao} - ${categoria}):</td>
        <td class="px-4 py-2 text-center text-blue-700">${catData.count}</td>
        <td class="px-4 py-2 text-right text-blue-700">${formatBRL(catData.total)}</td>
        <td class="px-4 py-2 text-right text-blue-700">${subPct.toFixed(2)}%</td>
      `;
      tbody.appendChild(trSub);
    });
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="3" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL SINTÉTICO (Visível):</td>
      <td class="px-4 py-3 text-center font-bold text-blue-700">${data.length}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">${formatBRL(grandTotal)}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">100,00%</td>
    </tr>
  `;
}

function renderAnalyticalTable() {
  const tbody = document.getElementById('analyticalTableBody');
  const tfoot = document.getElementById('analyticalTableFooter');
  if (!tbody || !tfoot) return;

  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const data = window.state.filteredData;
  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" class="px-4 py-8 text-center text-gray-400">Nenhum registro encontrado para os filtros aplicados.</td></tr>`;
    return;
  }

  data.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-blue-50/50 transition-colors border-b border-gray-100 text-xs";

    const statusBadgeClass = row['Status / Orçamento'] === 'A empenhar' ? 'bg-emerald-100 text-emerald-800' :
                             row['Status / Orçamento'] === 'Em elaboração' ? 'bg-amber-100 text-amber-800' :
                             'bg-gray-100 text-gray-600';

    tr.innerHTML = `
      <td class="px-3 py-2.5">${row['Ano Orçamento'] || '-'}</td>
      <td class="px-3 py-2.5 font-semibold text-gray-800">${row['Unidade'] || '-'}</td>
      <td class="px-3 py-2.5">${row['Área'] || '-'}</td>
      <td class="px-3 py-2.5">
        <span class="inline-block px-2 py-0.5 text-[11px] rounded-full font-medium ${statusBadgeClass}">
          ${row['Status / Orçamento'] || 'Não definido'}
        </span>
      </td>
      <td class="px-3 py-2.5 text-right font-semibold text-gray-900">${formatBRL(row['Valor Estimado'])}</td>
      <td class="px-3 py-2.5 font-mono text-gray-500">${row['Cód. Demanda (Siged / clarity)'] || '-'}</td>
      <td class="px-3 py-2.5 text-gray-800 font-medium">${row['Objeto/Sistema'] || '-'}</td>
      <td class="px-3 py-2.5">${row['Categoria Econômica'] || '-'}</td>
      <td class="px-3 py-2.5 text-gray-700">${row['Contratação'] || '-'}</td>
      <td class="px-3 py-2.5 text-gray-500">${row['Item do Contrato'] || '-'}</td>
      <td class="px-3 py-2.5 text-gray-700 font-mono">${row['Nº Empenho'] || '-'}</td>
      <td class="px-3 py-2.5 text-gray-700">${row['Ordem de Serviço'] || '-'}</td>
    `;
    tbody.appendChild(tr);
  });

  const totalVal = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);

  tfoot.innerHTML = `
    <tr>
      <td colspan="4" class="px-3 py-3 text-right font-bold text-gray-800">TOTAL ANALÍTICO (${data.length} ITENS):</td>
      <td class="px-3 py-3 text-right font-bold text-blue-700 text-xs">${formatBRL(totalVal)}</td>
      <td colspan="7" class="px-3 py-3"></td>
    </tr>
  `;
}

/* --- 📤 EXPORTAÇÃO PARA EXCEL --- */
function exportToExcel() {
  if (!window.state.filteredData || window.state.filteredData.length === 0) {
    alert("Não há dados visíveis para exportar.");
    return;
  }

  const exportData = window.state.filteredData.map(row => {
    const cleanRow = { ...row };
    delete cleanRow.__id;
    return cleanRow;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Demandas Analíticas");

  const dataSufixo = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `Demandas_SEGES_2026_${dataSufixo}.xlsx`);
}

// Vinculação ao escopo global para eventos HTML
window.toggleDropdown = toggleDropdown;
window.toggleHeaderFilter = toggleHeaderFilter;
window.filterDropdownOptions = filterDropdownOptions;
window.filterColOptions = filterColOptions;
window.handleTopFilterChange = handleTopFilterChange;
window.handleColFilterChange = handleColFilterChange;
window.selectAllOptions = selectAllOptions;
window.deselectAllOptions = deselectAllOptions;
window.selectAllCol = selectAllCol;
window.deselectAllCol = deselectAllCol;
window.resetFilters = resetFilters;
window.exportToExcel = exportToExcel;

/* --- 🚀 DISPARO INICIAL --- */
document.addEventListener('DOMContentLoaded', () => {
  loadExcelFromRepo();
});