window.state = {
  rawData: [],
  filteredData: [],
  activeFilters: {},
  isInitialized: false
};

// Configurações do Repositório GitHub
const REPO_OWNER = 'marcosangelo83-coder';
const REPO_NAME = 'Relatorio';
const FILE_PATH = 'demandas.json';
const GITHUB_TOKEN = ''; 

// URL do arquivo de dados público
const GITHUB_DATA_URL = `https://raw.githubusercontent.com/${REPO_OWNER}/${REPO_NAME}/main/${FILE_PATH}`;

/* --- SINCRONIZAÇÃO AUTOMÁTICA VIA GITHUB --- */
async function loadDataFromGithub() {
  updateSyncStatus('loading', 'Sincronizando dados com o GitHub...');
  try {
    const urlWithCacheBuster = `${GITHUB_DATA_URL}?t=${new Date().getTime()}`;
    const response = await fetch(urlWithCacheBuster);

    if (!response.ok) {
      throw new Error(`Erro na conexão (${response.status}: ${response.statusText})`);
    }

    let jsonRecords = [];
    const isExcel = GITHUB_DATA_URL.toLowerCase().includes('.xlsx') || GITHUB_DATA_URL.toLowerCase().includes('.xls');

    if (isExcel) {
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      jsonRecords = XLSX.utils.sheet_to_json(worksheet);
    } else {
      jsonRecords = await response.json();
    }

    if (Array.isArray(jsonRecords) && jsonRecords.length > 0) {
      initData(jsonRecords);
      const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      updateSyncStatus('success', `Atualizado via GitHub (${hora})`);
    } else {
      throw new Error('Arquivo no GitHub não possui registros válidos.');
    }
  } catch (err) {
    console.error("Falha ao sincronizar com GitHub:", err);
    updateSyncStatus('error', 'Falha ao buscar dados do GitHub');
  }
}

/* --- SALVAR ALTERAÇÕES DIRETAMENTE NO GITHUB --- */
async function saveToGithub() {
  let token = GITHUB_TOKEN;

  if (!token || token === 'COLE_SEU_TOKEN_GHP_AQUI') {
    token = prompt("Insira seu Personal Access Token do GitHub para salvar:");
    if (!token) {
      alert("Operação cancelada: Token não fornecido.");
      return;
    }
  }

  updateSyncStatus('loading', 'Enviando alterações para o GitHub...');
  const apiUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;

  try {
    // 1. Obter o SHA do arquivo atual
    const getFileResponse = await fetch(apiUrl, {
      headers: { 'Authorization': `token ${token}` }
    });

    let sha = '';
    if (getFileResponse.ok) {
      const fileData = await getFileResponse.json();
      sha = fileData.sha;
    }

    // 2. Limpar metadados e preparar o JSON
    const cleanData = window.state.rawData.map(row => {
      const copy = { ...row };
      delete copy.__id;
      return copy;
    });

    const jsonString = JSON.stringify(cleanData, null, 2);
    const contentBase64 = btoa(unescape(encodeURIComponent(jsonString)));

    // 3. Enviar requisição PUT para salvar no repositório
    const putResponse = await fetch(apiUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `token ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        message: 'Atualização realizada via Painel Web SEGES',
        content: contentBase64,
        sha: sha
      })
    });

    if (putResponse.ok) {
      const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      updateSyncStatus('success', `Salvo no GitHub (${hora})`);
      alert('Dados salvos com sucesso diretamente no seu repositório GitHub!');
    } else {
      throw new Error(`Erro na API (${putResponse.status})`);
    }
  } catch (error) {
    console.error('Erro ao salvar no GitHub:', error);
    updateSyncStatus('error', 'Falha ao salvar no GitHub');
    alert('Erro ao salvar no GitHub. Verifique as permissões do seu token.');
  }
}

/* --- FUNÇÕES UTILITÁRIAS --- */
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

function updateSyncStatus(type, message) {
  const statusEl = document.getElementById('syncStatus');
  if (!statusEl) return;

  const styles = {
    loading: 'bg-blue-100 text-blue-800 border-blue-200',
    success: 'bg-emerald-100 text-emerald-800 border-emerald-200',
    error: 'bg-amber-100 text-amber-800 border-amber-200',
    idle: 'bg-gray-100 text-gray-600 border-gray-200'
  };

  const icons = {
    loading: 'fa-solid fa-spinner fa-spin text-blue-600',
    success: 'fa-solid fa-circle-check text-emerald-600',
    error: 'fa-solid fa-triangle-exclamation text-amber-600',
    idle: 'fa-solid fa-circle-dot text-gray-400'
  };

  statusEl.className = `inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-medium border ${styles[type] || styles.idle}`;
  statusEl.innerHTML = `<i class="${icons[type] || icons.idle} text-[11px]"></i> ${message}`;
}

/* --- UPLOAD LOCAL --- */
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  updateSyncStatus('loading', 'Importando arquivo local...');
  const reader = new FileReader();

  if (file.name.endsWith('.json')) {
    reader.onload = function (e) {
      try {
        const jsonRecords = JSON.parse(e.target.result);
        initData(jsonRecords);
        const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        updateSyncStatus('success', `Importado local: ${file.name} (${hora})`);
      } catch (err) {
        alert("Erro ao ler arquivo JSON.");
        updateSyncStatus('error', 'Erro no JSON');
      }
    };
    reader.readAsText(file);
  } else {
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonRecords = XLSX.utils.sheet_to_json(worksheet);

        if (jsonRecords && jsonRecords.length > 0) {
          initData(jsonRecords);
          const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
          updateSyncStatus('success', `Importado local: ${file.name} (${hora})`);
        } else {
          alert("O arquivo selecionado não contém dados válidos.");
          updateSyncStatus('error', 'Arquivo sem dados');
        }
      } catch (err) {
        console.error("Erro ao ler arquivo:", err);
        alert("Erro ao ler planilha Excel.");
        updateSyncStatus('error', 'Falha na importação');
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

/* --- INICIALIZAÇÃO DE DADOS --- */
function initData(jsonRecords) {
  if (!Array.isArray(jsonRecords) || jsonRecords.length === 0) return;

  window.state.rawData = jsonRecords.map((row, index) => ({
    ...row,
    __id: index,
    'Valor Estimado': parseCurrency(row['Valor Estimado'])
  }));

  window.state.isInitialized = true;
  populateSelectOptions();
  applyFilters();
}

function notifyStateChange() {
  renderKPIs();
  renderSyntheticTable();
  renderAnalyticalTable();
}

/* --- FILTROS --- */
function populateSelectOptions() {
  const filterFields = [
    { id: 'filterUnidade', key: 'Unidade' },
    { id: 'filterStatus', key: 'Status / Orçamento' },
    { id: 'filterCategoria', key: 'Categoria Econômica' },
    { id: 'filterContratacao', key: 'Contratação' }
  ];

  filterFields.forEach(({ id, key }) => {
    const select = document.getElementById(id);
    if (!select) return;

    const currentValue = select.value;
    const values = [...new Set(window.state.rawData.map(item => (item[key] || '(Vazio)').toString().trim()))].sort();

    select.innerHTML = `<option value="">Todas as Opções (${key})</option>`;
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
      const rowVal = (row[key] || '(Vazio)').toString().trim();
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

/* --- RENDERIZAÇÃO DE PAINÉIS E TABELAS --- */
function renderKPIs() {
  const totalRecords = window.state.filteredData.length;
  const totalValue = window.state.filteredData.reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const capitalValue = window.state.filteredData
    .filter(row => row['Categoria Econômica'] === 'Despesas de Capital')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const currentValue = window.state.filteredData
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
    tbody.innerHTML = `<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">Nenhum dado disponível.</td></tr>`;
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
    tr.innerHTML = `
      <td class="px-4 py-2.5 font-medium text-gray-800">${row.contratacao}</td>
      <td class="px-4 py-2.5">
        <span class="inline-block px-2 py-0.5 text-xs rounded font-medium ${
          row.categoria === 'Despesas de Capital' ? 'bg-purple-100 text-purple-700' :
          row.categoria === 'Despesas Correntes' ? 'bg-amber-100 text-amber-700' : 'bg-gray-100 text-gray-600'
        }">
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
      <td colspan="2" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL SINTÉTICO:</td>
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
    tbody.innerHTML = `<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">Nenhum registro encontrado.</td></tr>`;
    return;
  }

  data.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-blue-50/50 transition-colors border-b border-gray-100";
    tr.innerHTML = `
      <td class="px-4 py-3 text-xs">${row['Ano Orçamento'] || '-'}</td>
      <td class="px-4 py-3 font-semibold text-gray-800 text-xs">${row['Unidade'] || '-'}</td>
      <td class="px-4 py-3 text-xs">${row['Área'] || '-'}</td>
      <td class="px-4 py-3 text-xs">
        <span class="inline-block px-2 py-0.5 text-xs rounded-full font-medium ${
          row['Status / Orçamento'] === 'A empenhar' ? 'bg-emerald-100 text-emerald-800' :
          row['Status / Orçamento'] === 'Em elaboração' ? 'bg-amber-100 text-amber-800' : 'bg-gray-100 text-gray-600'
        }">
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

/* --- EXPORTAÇÃO EXCEL --- */
function exportToExcel() {
  if (!window.state.filteredData || window.state.filteredData.length === 0) {
    alert("Não há dados para exportar.");
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

  XLSX.writeFile(workbook, `Demandas_SEGES_2026_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* --- INICIALIZAÇÃO NA CARGA DA PÁGINA --- */
document.addEventListener('DOMContentLoaded', () => {
  loadDataFromGithub();
});