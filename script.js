/* ==========================================================================
   PAINEL DE GESTÃO DE DEMANDAS - SEGES 2026
   Fonte de dados: demanda.xlsx (carregado dinamicamente via repositório)
   ========================================================================== */

/* --- ⚙️ ESTADO GLOBAL DA APLICAÇÃO --- */
window.state = {
  rawData: [],
  filteredData: [],
  activeFilters: {}
};

/* --- 🛠️ FUNÇÕES UTILITÁRIAS DE FORMATAÇÃO --- */
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

/* --- 📥 CARREGAMENTO EXCLUSIVO VIA EXCEL (GITHUB) --- */
async function loadExcelFromRepo() {
  const filterCounter = document.getElementById('filterCounter');
  if (filterCounter) {
    filterCounter.textContent = 'Carregando dados do Excel...';
    filterCounter.className = 'text-xs font-medium text-amber-600 bg-amber-50 px-2.5 py-1 rounded-full';
  }

  try {
    // Adiciona timestamp (?t=...) para evitar cache do navegador e pegar sempre o arquivo mais recente
    const response = await fetch('./demandas.xlsx?t=' + new Date().getTime());
    
    if (!response.ok) {
      throw new Error(`O arquivo 'demandas.xlsx' não foi encontrado na raiz do repositório (Status ${response.status}).`);
    }

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    // Pega a primeira aba da planilha
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Converte a planilha em objetos JSON
    const jsonRecords = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    if (!jsonRecords || jsonRecords.length === 0) {
      throw new Error("A planilha 'demandas.xlsx' foi encontrada, mas está vazia.");
    }

    initData(jsonRecords);

  } catch (error) {
    console.error("Erro ao carregar o arquivo Excel:", error);
    if (filterCounter) {
      filterCounter.textContent = 'Erro ao carregar dados do demandas.xlsx';
      filterCounter.className = 'text-xs font-semibold text-red-700 bg-red-100 px-2.5 py-1 rounded-full';
    }
  }
}

/* --- 🔄 INICIALIZAÇÃO DOS DADOS --- */
function initData(jsonRecords) {
  window.state.rawData = jsonRecords.map((row, index) => ({
    ...row,
    __id: index,
    'Valor Estimado': parseCurrency(row['Valor Estimado'])
  }));

  populateSelectOptions();
  applyFilters();
}

/* --- 🔍 GERENCIAMENTO DE FILTROS DINÂMICOS --- */
function populateSelectOptions() {
  const filterFields = [
    { id: 'filterUnidade', key: 'Unidade', defaultLabel: 'Todas as Unidades' },
    { id: 'filterStatus', key: 'Status / Orçamento', defaultLabel: 'Todos os Status' },
    { id: 'filterCategoria', key: 'Categoria Econômica', defaultLabel: 'Todas as Categorias' },
    { id: 'filterContratacao', key: 'Contratação', defaultLabel: 'Todas as Contratações' }
  ];

  filterFields.forEach(({ id, key, defaultLabel }) => {
    const select = document.getElementById(id);
    if (!select) return;

    const currentValue = select.value;
    
    // Mapeia valores únicos diretamente dos dados do Excel
    const values = [...new Set(window.state.rawData.map(item => {
      const val = item[key];
      return (val !== undefined && val !== null && val !== '') ? val.toString().trim() : '(Não Informado)';
    }))].sort();

    select.innerHTML = `<option value="">${defaultLabel}</option>`;
    
    values.forEach(val => {
      const option = document.createElement('option');
      option.value = val;
      option.textContent = val;
      select.appendChild(option);
    });
    
    select.value = currentValue;
  });
}

function handleFilterChange(field, value) {
  if (!value) {
    delete window.state.activeFilters[field];
  } else {
    window.state.activeFilters[field] = value;
  }
  applyFilters();
}

function applyFilters() {
  window.state.filteredData = window.state.rawData.filter(row => {
    return Object.entries(window.state.activeFilters).every(([key, val]) => {
      const rowVal = (row[key] !== undefined && row[key] !== null && row[key] !== '') 
        ? row[key].toString().trim() 
        : '(Não Informado)';
      return rowVal === val;
    });
  });

  const filterCounter = document.getElementById('filterCounter');
  const activeCount = Object.keys(window.state.activeFilters).length;
  if (filterCounter) {
    if (activeCount === 0) {
      filterCounter.textContent = `Mostrando todos os ${window.state.filteredData.length} registros`;
      filterCounter.className = "text-xs font-medium text-gray-500 bg-gray-100 px-2.5 py-1 rounded-full";
    } else {
      filterCounter.textContent = `${activeCount} filtro(s) ativo(s) | ${window.state.filteredData.length} registro(s)`;
      filterCounter.className = "text-xs font-semibold text-blue-700 bg-blue-100 px-2.5 py-1 rounded-full";
    }
  }

  notifyStateChange();
}

function resetFilters() {
  window.state.activeFilters = {};
  ['filterUnidade', 'filterStatus', 'filterCategoria', 'filterContratacao'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  applyFilters();
}

/* --- 📊 RENDERIZAÇÃO DA INTERFACE (KPIs E TABELAS) --- */
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
    tbody.innerHTML = `<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">Nenhum dado disponível para os filtros selecionados.</td></tr>`;
    return;
  }

  const totalValAll = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);
  const summaryMap = {};

  data.forEach(item => {
    const contratacao = (item['Contratação'] && item['Contratação'].toString().trim()) ? item['Contratação'].toString().trim() : '(Não Informado)';
    const categoria = (item['Categoria Econômica'] && item['Categoria Econômica'].toString().trim()) ? item['Categoria Econômica'].toString().trim() : '(Não Informado)';
    const key = `${contratacao}||${categoria}`;

    if (!summaryMap[key]) {
      summaryMap[key] = { contratacao, categoria, count: 0, totalValue: 0 };
    }
    summaryMap[key].count += 1;
    summaryMap[key].totalValue += (item['Valor Estimado'] || 0);
  });

  const summaryArray = Object.values(summaryMap).sort((a, b) => b.totalValue - a.totalValue);

  summaryArray.forEach(row => {
    const pct = totalValAll > 0 ? (row.totalValue / totalValAll) * 100 : 0;
    const tr = document.createElement('tr');
    tr.className = "hover:bg-gray-50 border-b border-gray-100 transition-colors";
    
    const badgeClass = row.categoria === 'Despesas de Capital' ? 'bg-purple-100 text-purple-700' :
                       row.categoria === 'Despesas Correntes' ? 'bg-amber-100 text-amber-700' : 
                       'bg-gray-100 text-gray-600';

    tr.innerHTML = `
      <td class="px-4 py-2.5 font-medium text-gray-800">${row.contratacao}</td>
      <td class="px-4 py-2.5">
        <span class="inline-block px-2 py-0.5 text-xs rounded font-medium ${badgeClass}">
          ${row.categoria}
        </span>
      </td>
      <td class="px-4 py-2.5 text-center font-medium">${row.count}</td>
      <td class="px-4 py-2.5 text-right font-bold text-gray-900">${formatBRL(row.totalValue)}</td>
      <td class="px-4 py-2.5 text-right text-xs font-semibold text-gray-500">${pct.toFixed(2)}%</td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="2" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL SINTÉTICO (Visível):</td>
      <td class="px-4 py-3 text-center font-bold text-blue-700">${data.length}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">${formatBRL(totalValAll)}</td>
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
    tbody.innerHTML = `<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">Nenhum registro encontrado para os filtros aplicados.</td></tr>`;
    return;
  }

  data.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-blue-50/50 transition-colors border-b border-gray-100";
    
    const statusBadgeClass = row['Status / Orçamento'] === 'A empenhar' ? 'bg-emerald-100 text-emerald-800' :
                             row['Status / Orçamento'] === 'Em elaboração' ? 'bg-amber-100 text-amber-800' : 
                             'bg-gray-100 text-gray-600';

    tr.innerHTML = `
      <td class="px-4 py-3 text-xs">${row['Ano Orçamento'] || '-'}</td>
      <td class="px-4 py-3 font-semibold text-gray-800 text-xs">${row['Unidade'] || '-'}</td>
      <td class="px-4 py-3 text-xs">${row['Área'] || '-'}</td>
      <td class="px-4 py-3 text-xs">
        <span class="inline-block px-2 py-0.5 text-xs rounded-full font-medium ${statusBadgeClass}">
          ${row['Status / Orçamento'] || 'Não definido'}
        </span>
      </td>
      <td class="px-4 py-3 text-right font-semibold text-gray-900 text-xs">${formatBRL(row['Valor Estimado'])}</td>
      <td class="px-4 py-3 text-xs font-mono text-gray-500">${row['Cód. Demanda (Siged / clarity)'] || '-'}</td>
      <td class="px-4 py-3 text-xs text-gray-800 font-medium">${row['Objeto/Sistema'] || '-'}</td>
      <td class="px-4 py-3 text-xs">${row['Categoria Econômica'] || '-'}</td>
      <td class="px-4 py-3 text-xs text-gray-700">${row['Contratação'] || '-'}</td>
      <td class="px-4 py-3 text-xs text-gray-500">${row['Item do Contrato'] || '-'}</td>
    `;
    tbody.appendChild(tr);
  });

  const totalVal = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);

  tfoot.innerHTML = `
    <tr>
      <td colspan="4" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL ANALÍTICO (${data.length} ITENS):</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700 text-sm">${formatBRL(totalVal)}</td>
      <td colspan="5" class="px-4 py-3"></td>
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

// Vinculação com o escopo global do HTML (onClick/onChange)
window.handleFilterChange = handleFilterChange;
window.resetFilters = resetFilters;
window.exportToExcel = exportToExcel;

/* --- 🚀 DISPARO AO CARREGAR A PÁGINA --- */
document.addEventListener('DOMContentLoaded', () => {
  loadExcelFromRepo();
});