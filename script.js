window.state = {
  rawData: [],
  filteredData: [],
  activeFilters: {},
  isInitialized: false
};

// ⚙️ Configurações do Repositório GitHub
// Mantenha estas configurações corretas para o seu repositório privado
const REPO_OWNER = 'marcosangelo83-coder';
const REPO_NAME = 'Relatorio';
const FILE_PATH = 'demandas.json';
const GITHUB_TOKEN = ''; // Opcional: Cole aqui ou informe no prompt (ficará salvo no navegador)

/* --- 🔑 GERENCIAMENTO DE TOKEN (PERSISTÊNCIA LOCAL) --- */
// Obtém o token do localStorage ou solicita ao usuário via prompt
function getGithubToken() {
  let token = localStorage.getItem('gh_token') || GITHUB_TOKEN;

  if (!token || token === 'COLE_SEU_TOKEN_GHP_AQUI' || token.trim() === '') {
    token = prompt("Seu repositório é privado. Insira seu Personal Access Token do GitHub para autenticar:");
    if (token && token.trim() !== '') {
      token = token.trim();
      localStorage.setItem('gh_token', token);
    } else {
      return null;
    }
  }
  return token;
}

// Remove o token salvo no navegador (útil para troca de usuário ou segurança)
function clearGithubToken() {
  localStorage.removeItem('gh_token');
  alert('Token do GitHub removido da memória do navegador.');
}

/* --- 🔑 CARREGAMENTO INICIAL VIA GITHUB (PULL) --- */
// Esta função roda apenas na carga da página para buscar o estado atual da nuvem
async function loadDataFromGithub() {
  const token = getGithubToken();
  if (!token) {
    updateSyncStatus('error', 'Token não fornecido');
    // Se não houver token, o painel fica vazio aguardando importação local
    return;
  }

  updateSyncStatus('loading', 'Sincronizando dados com o GitHub...');
  // API URL para obter o conteúdo do arquivo
  const apiUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;

  try {
    const response = await fetch(apiUrl, {
      headers: {
        'Authorization': `token ${token}`,
        // Pede o conteúdo bruto do arquivo (raw) para evitar parsing de Base64 manual aqui
        'Accept': 'application/vnd.github.v3.raw'
      }
    });

    // Trata erro 404 (Arquivo não existe) amigavelmente
    if (response.status === 404) {
      updateSyncStatus('error', 'Arquivo demandas.json ainda não criado no GitHub');
      console.warn('O arquivo demandas.json não existe no repositório. Importe o Excel local e clique em "Salvar no GitHub" para criá-lo.');
      return;
    }

    if (!response.ok) {
      throw new Error(`Erro na conexão (${response.status}: ${response.statusText})`);
    }

    let jsonRecords = [];
    // Suporte a JSON ou Excel (baseado na extensão do FILE_PATH definido no topo)
    const isExcel = FILE_PATH.toLowerCase().endsWith('.xlsx') || FILE_PATH.toLowerCase().endsWith('.xls');

    if (isExcel) {
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      jsonRecords = XLSX.utils.sheet_to_json(worksheet);
    } else {
      // Padrão JSON
      jsonRecords = await response.json();
    }

    if (Array.isArray(jsonRecords) && jsonRecords.length > 0) {
      initData(jsonRecords); // Inicializa o painel com os dados da nuvem
      const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      updateSyncStatus('success', `Sincronizado via GitHub (${hora})`);
    } else {
      throw new Error('O arquivo no GitHub não contém registros válidos.');
    }
  } catch (err) {
    console.error("Falha ao sincronizar com GitHub:", err);
    updateSyncStatus('error', 'Falha ao buscar dados do GitHub');
  }
}

/* --- 💾 SALVAR ALTERAÇÕES NO GITHUB (PUSH) --- */
// Função vinculada ao botão "Sincronizar" (no HTML como "Atualizar do GitHub")
// Ela envia o estado da TELA para a nuvem, SEM alterar a tela.
async function saveToGithub() {
  const token = getGithubToken();
  if (!token) {
    alert("Operação cancelada: Token não fornecido.");
    return;
  }

  // Validação: Não permite salvar se a tela estiver vazia
  if (!window.state.rawData || window.state.rawData.length === 0) {
    alert("Não há dados na tela para salvar no GitHub. Importe um arquivo local (Excel) primeiro.");
    return;
  }

  updateSyncStatus('loading', 'Enviando alterações para o GitHub...');
  const apiUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;

  try {
    // 1. Verificar se o arquivo já existe e obter seu SHA
    // Lógica CORRIGIDA: Trata explicitamente o status 404 Not Found
    let sha = null;
    let fileResponse;
    try {
      fileResponse = await fetch(apiUrl, {
        headers: { 'Authorization': `token ${token}` }
      });
    } catch (fetchError) {
      throw new Error(`Erro ao buscar metadados do arquivo: ${fetchError.message}`);
    }

    if (fileResponse.ok) {
      // Arquivo existe -> Pega o SHA para atualização (update)
      const fileData = await fileResponse.json();
      sha = fileData.sha;
    } else if (fileResponse.status === 404) {
      // Arquivo não existe -> Trata como criação de novo arquivo (create)
      sha = null; // Já é null, mas deixa explícito
    } else {
      // Outro erro de API (403, 401, 500, etc.) -> Não podemos prosseguir com PUT
      const errData = await fileResponse.json().catch(() => ({}));
      throw new Error(`Erro ao verificar existência do arquivo (${fileResponse.status}): ${errData.message || ''}`);
    }

    // 2. Preparar e Limpar dados para envio
    // Removemos o '__id' temporário que a função initData adiciona para controle interno da tabela
    const cleanData = window.state.rawData.map(row => {
      const copy = { ...row };
      delete copy.__id; // Remove metadado local
      return copy;
    });

    // Converte o array de objetos para string JSON formatada
    const jsonString = JSON.stringify(cleanData, null, 2);
    // Codifica para Base64 (padrão exigido pela API de Conteúdos do GitHub)
    const contentBase64 = btoa(unescape(encodeURIComponent(jsonString)));

    // 3. Montar o Payload da requisição PUT
    const payload = {
      message: 'Atualização de dados via Painel SEGES 2026 (Botão Sincronizar)',
      content: contentBase64
    };

    // Se o arquivo já existia (sha não é null), precisamos incluir o SHA para confirmar a sobrescrita
    if (sha) {
      payload.sha = sha;
    }

    // 4. Executar a requisição PUT para salvar/sobrescrever o arquivo
    const putResponse = await fetch(apiUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `token ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    if (putResponse.ok) {
      const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      // ATENÇÃO: NÃO chamamos initData() ou loadDataFromGithub() aqui.
      // Mantemos os dados importados localmente visíveis na tela.
      updateSyncStatus('success', `Salvo no GitHub (${hora})`);
      alert('Dados salvos com sucesso no repositório GitHub (demandas.json atualizado)!');
    } else {
      const errData = await putResponse.json().catch(() => ({}));
      throw new Error(`Erro na API (${putResponse.status}): ${errData.message || ''}`);
    }
  } catch (error) {
    console.error('Erro ao salvar no GitHub:', error);
    updateSyncStatus('error', 'Falha ao salvar no GitHub');
    alert(`Erro ao salvar no GitHub. Verifique as permissões do seu token (escopo 'repo').\nErro: ${error.message}`);
  }
}

/* --- 🛠️ FUNÇÕES UTILITÁRIAS DE FORMATAÇÃO --- */
// Converte string monetária "R$ 1.000,00" ou número para float puro
function parseCurrency(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  // Remove R$, pontos de milhar e substitui vírgula por ponto decimal
  const cleanStr = val.toString().replace(/R\$\s?/g, '').replace(/\./g, '').replace(',', '.').trim();
  const parsed = parseFloat(cleanStr);
  return isNaN(parsed) ? 0 : parsed;
}

// Converte float puro para string monetária formatada "R$ 1.000,00"
function formatBRL(valor) {
  const num = typeof valor === 'number' ? valor : parseCurrency(valor);
  return num.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// Atualiza o indicador visual de status de sincronização no cabeçalho
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

/* --- 📄 UPLOAD E PARSING DE ARQUIVO LOCAL (EXCEL/JSON) --- */
// Função chamada quando o usuário seleciona um arquivo no input file
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  updateSyncStatus('loading', 'Importando arquivo local...');
  const reader = new FileReader();

  // Parsing baseado na extensão do arquivo local
  if (file.name.endsWith('.json')) {
    reader.onload = function (e) {
      try {
        const jsonRecords = JSON.parse(e.target.result);
        initData(jsonRecords); // Atualiza a TELA (window.state)
        const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        updateSyncStatus('idle', `Importado local: ${file.name} (${hora})`);
        // Nota: O status fica 'idle' (cinza) indicando que NÃO está salvo na nuvem ainda.
      } catch (err) {
        alert("Erro ao ler arquivo JSON local.");
        updateSyncStatus('error', 'Erro no JSON');
      }
    };
    reader.readAsText(file);
  } else {
    // Padrão Excel (.xlsx, .xls)
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        // Usa a biblioteca XLSX para ler o ArrayBuffer
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // Converte a planilha para array de objetos JSON
        const jsonRecords = XLSX.utils.sheet_to_json(worksheet);

        if (jsonRecords && jsonRecords.length > 0) {
          initData(jsonRecords);
          const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
          updateSyncStatus('idle', `Importado local: ${file.name} (${hora})`);
        } else {
          alert("O arquivo selecionado não contém dados válidos.");
          updateSyncStatus('error', 'Arquivo sem dados');
        }
      } catch (err) {
        console.error("Erro ao ler arquivo Excel local:", err);
        alert("Erro ao ler planilha Excel.");
        updateSyncStatus('error', 'Falha na importação');
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

/* --- 🔄 INICIALIZAÇÃO DO ESTADO E RENDERIZAÇÃO --- */
// Recebe registros JSON puros, processa moedas e atualiza o estado global e a UI
function initData(jsonRecords) {
  if (!Array.isArray(jsonRecords) || jsonRecords.length === 0) return;

  // Processamento e limpeza de dados
  window.state.rawData = jsonRecords.map((row, index) => ({
    ...row,
    __id: index, // Adiciona ID único temporário para controle da tabela analítica
    // Garante que o campo 'Valor Estimado' seja um float puro para cálculos
    'Valor Estimado': parseCurrency(row['Valor Estimado'])
  }));

  window.state.isInitialized = true;
  // Recarrega as opções dos filtros baseados nos novos dados
  populateSelectOptions();
  // Aplica filtros vazios para renderizar tudo inicialmente
  applyFilters();
}

// Função centralizadora chamada sempre que o estado dos dados filtrados muda
function notifyStateChange() {
  // Se não houver dados, não renderiza (as funções de render tratam o estado vazio)
  renderKPIs();
  renderSyntheticTable();
  renderAnalyticalTable();
}

/* --- 🔍 GERENCIAMENTO DE FILTROS DINÂMICOS --- */
// Preenche os elementos <select> HTML com opções únicas baseadas nos dados brutos
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

    // Guarda o valor que estava selecionado antes de recarregar
    const currentValue = select.value;
    
    // Obtém valores únicos, remove nulos, formata e ordena
    const values = [...new Set(window.state.rawData.map(item => 
      (item[key] || '(Vazio)').toString().trim()
    ))].sort();

    // Reinicia o select com a opção padrão
    select.innerHTML = `<option value="">Todas as Opções (${key})</option>`;
    
    // Adiciona as novas opções dinâmicas
    values.forEach(val => {
      const option = document.createElement('option');
      option.value = val;
      option.textContent = val;
      select.appendChild(option);
    });
    
    // Tenta restaurar o valor selecionado anteriormente
    select.value = currentValue;
  });
}

// Chamada pelo HTML (onchange) quando um filtro é alterado
function handleFilterChange(field, value) {
  if (!value) {
    // Se selecionou a opção vazia, remove o filtro ativo
    delete window.state.activeFilters[field];
  } else {
    // Adiciona ou atualiza o filtro ativo
    window.state.activeFilters[field] = value;
  }
  applyFilters();
}

// Executa a lógica de filtragem combinada sobre os dados brutos (window.state.rawData)
function applyFilters() {
  // Gera window.state.filteredData a partir do rawData
  window.state.filteredData = window.state.rawData.filter(row => {
    // Retorna true apenas se a linha passar em TODOS os filtros ativos (lógica AND)
    return Object.entries(window.state.activeFilters).every(([key, val]) => {
      // Formata o valor da linha para comparação (tratando vazios)
      const rowVal = (row[key] || '(Vazio)').toString().trim();
      return rowVal === val;
    });
  });

  // Atualiza o contador de registros visíveis na UI
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

  // Notifica que os dados filtrados mudaram para engatilhar a renderização
  notifyStateChange();
}

// Limpa todos os filtros ativos e reseta os selects na UI
function resetFilters() {
  window.state.activeFilters = {};
  // Reseta visualmente os selects
  ['filterUnidade', 'filterStatus', 'filterCategoria', 'filterContratacao'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  applyFilters(); // Re-aplica (vai mostrar tudo)
}

/* --- 📊 RENDERIZAÇÃO DE UI (KPIs E TABELAS) --- */
// Calcula e renderiza os cards de KPI baseados nos dados filtrados
function renderKPIs() {
  const data = window.state.filteredData;
  const totalRecords = data.length;
  
  // Somatórios usando reduce sobre o float puro
  const totalValue = data.reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const capitalValue = data
    .filter(row => row['Categoria Econômica'] === 'Despesas de Capital')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const currentValue = data
    .filter(row => row['Categoria Econômica'] === 'Despesas Correntes')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);

  // Seleciona elementos HTML
  const totalRecEl = document.getElementById('kpiTotalRecords');
  const totalValEl = document.getElementById('kpiTotalValue');
  const capValEl = document.getElementById('kpiCapitalValue');
  const curValEl = document.getElementById('kpiCurrentValue');

  // Atualiza conteúdo formatado
  if (totalRecEl) totalRecEl.textContent = totalRecords;
  if (totalValEl) totalValEl.textContent = formatBRL(totalValue);
  if (capValEl) capValEl.textContent = formatBRL(capitalValue);
  if (curValEl) curValEl.textContent = formatBRL(currentValue);
}

// Agrupa dados e renderiza a Tabela Sintética (Resumo)
function renderSyntheticTable() {
  const tbody = document.getElementById('syntheticTableBody');
  const tfoot = document.getElementById('syntheticTableFooter');
  if (!tbody || !tfoot) return;

  // Limpa conteúdo anterior
  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const data = window.state.filteredData;
  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">Nenhum dado disponível para os filtros selecionados.</td></tr>`;
    return;
  }

  // Lógica de Agrupamento (Combinação: Contratação + Categoria Econômica)
  const totalValAll = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);
  const summaryMap = {};

  data.forEach(item => {
    const contratacao = (item['Contratação'] && item['Contratação'].toString().trim()) ? item['Contratação'].toString().trim() : '(Não Informado)';
    const categoria = (item['Categoria Econômica'] && item['Categoria Econômica'].toString().trim()) ? item['Categoria Econômica'].toString().trim() : '(Não Informado)';
    // Chave composta para agrupamento
    const key = `${contratacao}||${categoria}`;

    if (!summaryMap[key]) {
      summaryMap[key] = { contratacao, categoria, count: 0, totalValue: 0 };
    }
    summaryMap[key].count += 1;
    summaryMap[key].totalValue += (item['Valor Estimado'] || 0);
  });

  // Converte mapa para array e ordena por valor total descrescente
  const summaryArray = Object.values(summaryMap).sort((a, b) => b.totalValue - a.totalValue);

  // Renderiza linhas do corpo (tbody)
  summaryArray.forEach(row => {
    // Calcula porcentagem do grupo em relação ao total visível
    const pct = totalValAll > 0 ? (row.totalValue / totalValAll) * 100 : 0;
    const tr = document.createElement('tr');
    tr.className = "hover:bg-gray-50 border-b border-gray-100 transition-colors";
    
    // Classes CSS dinâmicas para o badge da categoria
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

  // Renderiza linha de total no rodapé (tfoot)
  tfoot.innerHTML = `
    <tr>
      <td colspan="2" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL SINTÉTICO (Visível):</td>
      <td class="px-4 py-3 text-center font-bold text-blue-700">${data.length}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">${formatBRL(totalValAll)}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">100,00%</td>
    </tr>
  `;
}

// Renderiza a Tabela Analítica (Detalhamento item a item)
function renderAnalyticalTable() {
  const tbody = document.getElementById('analyticalTableBody');
  const tfoot = document.getElementById('analyticalTableFooter');
  if (!tbody || !tfoot) return;

  // Limpa conteúdo anterior
  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const data = window.state.filteredData;
  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">Nenhum registro encontrado para os filtros aplicados.</td></tr>`;
    return;
  }

  // Renderiza linhas do corpo (tbody)
  data.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-blue-50/50 transition-colors border-b border-gray-100";
    
    // Classes CSS dinâmicas para o badge do status
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

  // Calcula total analítico visível
  const totalVal = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);

  tfoot.innerHTML = `
    <tr>
      <td colspan="4" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL ANALÍTICO (${data.length} ITENS):</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700 text-sm">${formatBRL(totalVal)}</td>
      <td colspan="5" class="px-4 py-3"></td>
    </tr>
  `;
}

/* --- 📤 EXPORTAÇÃO DE DADOS PARA EXCEL LOCAL --- */
// Gera um arquivo .xlsx contendo exatamente os dados visíveis (filtrados) na tela analítica
function exportToExcel() {
  if (!window.state.filteredData || window.state.filteredData.length === 0) {
    alert("Não há dados visíveis para exportar. Remova alguns filtros.");
    return;
  }

  // Prepara dados limpando metadados internos (__id) antes de gerar a planilha
  const exportData = window.state.filteredData.map(row => {
    const cleanRow = { ...row };
    delete cleanRow.__id; // Remove ID local
    return cleanRow;
  });

  // Usa biblioteca XLSX para criar workbook e worksheet
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Demandas Analíticas Filtradas");

  // Gera download do arquivo no navegador
  const dataSufixo = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  XLSX.writeFile(workbook, `Demandas_SEGES_2026_Export_${dataSufixo}.xlsx`);
}

// 🌐 Exportar funções cruciais para o escopo global (window)
// Isso é necessário porque o HTML Tailwind usa atributos onclick="funcao()" diretamente
window.loadDataFromGithub = loadDataFromGithub;
window.saveToGithub = saveToGithub; // Agora vinculado ao botão "Sincronizar"
window.handleFileUpload = handleFileUpload;
window.handleFilterChange = handleFilterChange;
window.resetFilters = resetFilters;
window.exportToExcel = exportToExcel;
window.clearGithubToken = clearGithubToken;

/* --- 🚀 INICIALIZAÇÃO AUTOMÁTICA NA CARGA DA PÁGINA --- */
document.addEventListener('DOMContentLoaded', () => {
  // Executa o PULL inicial do GitHub para preencher o painel
  loadDataFromGithub();

/* --- 📄 UPLOAD E PARSING DE ARQUIVO LOCAL (EXCEL/JSON) --- */
// Função chamada quando o usuário seleciona um arquivo no input file
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  updateSyncStatus('loading', 'Importando arquivo local...');
  const reader = new FileReader();

  // Parsing baseado na extensão do arquivo local
  if (file.name.endsWith('.json')) {
    reader.onload = function (e) {
      try {
        const jsonRecords = JSON.parse(e.target.result);
        if (!Array.isArray(jsonRecords) || jsonRecords.length === 0) {
          throw new Error("Arquivo JSON vazio ou em formato inválido.");
        }
        initData(jsonRecords); // Atualiza a TELA (window.state)
        const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        updateSyncStatus('idle', `Importado local: ${file.name} (${hora})`);
      } catch (err) {
        alert(`Erro ao ler arquivo JSON local: ${err.message}`);
        updateSyncStatus('error', 'Erro no JSON');
      }
    };
    reader.onerror = () => {
        alert("Erro na leitura do arquivo local.");
        updateSyncStatus('error', 'Falha na leitura');
    };
    reader.readAsText(file);
  } else {
    // Padrão Excel (.xlsx, .xls) - PARSER ROBUSTO ATUALIZADO
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Lógica de parsing robusta: Lemos como matriz pura (header: 1)
        // para mapear as colunas manualmente pelos nomes exatos, ignorando espaços extras ou formatações.
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        if (!rawRows || rawRows.length < 2) {
          throw new Error("A planilha selecionada está vazia ou não contém dados válidos.");
        }

        // 1. Identifica os cabeçalhos na primeira linha válida
        const headers = rawRows[0].map(h => (h ? h.toString().trim() : ''));

        // 2. Mapeia os índices das colunas necessárias baseados nos cabeçalhos esperados pelo sistema
        const colIndices = {
          'Ano Orçamento': headers.indexOf('Ano Orçamento'),
          'Unidade': headers.indexOf('Unidade'),
          'Área': headers.indexOf('Área'),
          'Status / Orçamento': headers.indexOf('Status / Orçamento'),
          'Valor Estimado': headers.indexOf('Valor Estimado'),
          'Cód. Demanda (Siged / clarity)': headers.indexOf('Cód. Demanda (Siged / clarity)'),
          'Objeto/Sistema': headers.indexOf('Objeto/Sistema'),
          'Categoria Econômica': headers.indexOf('Categoria Econômica'),
          'Contratação': headers.indexOf('Contratação'),
          'Item do Contrato': headers.indexOf('Item do Contrato')
        };

        // Verificação crítica: Garante que as colunas essenciais existem na planilha
        const colunasObrigatorias = ['Unidade', 'Status / Orçamento', 'Valor Estimado', 'Objeto/Sistema', 'Contratação'];
        const colunasFaltantes = colunasObrigatorias.filter(col => colIndices[col] === -1);

        if (colunasFaltantes.length > 0) {
          throw new Error(`A planilha Excel está incompleta. Coluna(s) obrigatória(s) não encontrada(s): ${colunasFaltantes.join(', ')}. Verifique os cabeçalhos da planilha.`);
        }

        // 3. Processa as linhas de dados (começando da linha 1, após os cabeçalhos)
        const jsonRecords = [];
        for (let i = 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          // Pula linhas que estão completamente vazias na planilha Excel
          if (!row || row.every(cell => cell === null || cell === '')) continue;

          // Cria o objeto JSON mapeando os índices encontrados para os nomes de propriedade esperados
          const record = {
            'Ano Orçamento': colIndices['Ano Orçamento'] !== -1 ? row[colIndices['Ano Orçamento']] : null,
            'Unidade': row[colIndices['Unidade']], // Obrigatória
            'Área': colIndices['Área'] !== -1 ? row[colIndices['Área']] : null,
            'Status / Orçamento': row[colIndices['Status / Orçamento']], // Obrigatória
            'Valor Estimado': row[colIndices['Valor Estimado']], // Obrigatória
            'Cód. Demanda (Siged / clarity)': colIndices['Cód. Demanda (Siged / clarity)'] !== -1 ? row[colIndices['Cód. Demanda (Siged / clarity)']] : null,
            'Objeto/Sistema': row[colIndices['Objeto/Sistema']], // Obrigatória
            'Categoria Econômica': colIndices['Categoria Econômica'] !== -1 ? row[colIndices['Categoria Econômica']] : null,
            'Contratação': row[colIndices['Contratação']], // Obrigatória
            'Item do Contrato': colIndices['Item do Contrato'] !== -1 ? row[colIndices['Item do Contrato']] : null
          };
          jsonRecords.push(record);
        }

        if (jsonRecords.length > 0) {
          initData(jsonRecords); // Atualiza o estado da tela com os dados lidos
          const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
          updateSyncStatus('idle', `Importado local: ${file.name} (${jsonRecords.length} registros) em ${hora}`);
          console.info('Dados importados localmente. Clique em "Atualizar do GitHub" para salvar permanentemente na nuvem.');
        } else {
          throw new Error("O arquivo Excel foi lido, mas não contém linhas de dados válidas abaixo dos cabeçalhos.");
        }
      } catch (err) {
        console.error("Erro crítico no parser Excel local:", err);
        alert(`Falha na importação do Excel: ${err.message}`);
        updateSyncStatus('error', 'Falha na importação');
      }
    };
    reader.onerror = () => {
        alert("Erro na leitura física do arquivo Excel.");
        updateSyncStatus('error', 'Falha na leitura');
    };
    reader.readAsArrayBuffer(file);
  }
}

/* --- 🔄 INICIALIZAÇÃO DO ESTADO E RENDERIZAÇÃO --- */
// Recebe registros JSON puros, processa moedas e atualiza o estado global e a UI
function initData(jsonRecords) {
  if (!Array.isArray(jsonRecords) || jsonRecords.length === 0) return;

  // Processamento e limpeza de dados
  window.state.rawData = jsonRecords.map((row, index) => ({
    ...row,
    __id: index, // Adiciona ID único temporário para controle da tabela analítica
    // Garante que o campo 'Valor Estimado' seja um float puro para cálculos
    'Valor Estimado': parseCurrency(row['Valor Estimado'])
  }));

  window.state.isInitialized = true;
  // Recarrega as opções dos filtros baseados nos novos dados
  populateSelectOptions();
  // Aplica filtros vazios para renderizar tudo inicialmente
  applyFilters();
}

// Função centralizadora chamada sempre que o estado dos dados filtrados muda
function notifyStateChange() {
  // Se não houver dados, não renderiza (as funções de render tratam o estado vazio)
  renderKPIs();
  renderSyntheticTable();
  renderAnalyticalTable();
}

/* --- 🔍 GERENCIAMENTO DE FILTROS DINÂMICOS --- */
// Preenche os elementos <select> HTML com opções únicas baseadas nos dados brutos
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

    // Guarda o valor que estava selecionado antes de recarregar
    const currentValue = select.value;
    
    // Obtém valores únicos, remove nulos, formata e ordena
    const values = [...new Set(window.state.rawData.map(item => 
      (item[key] || '(Vazio)').toString().trim()
    ))].sort();

    // Reinicia o select com a opção padrão
    select.innerHTML = `<option value="">Todas as Opções (${key})</option>`;
    
    // Adiciona as novas opções dinâmicas
    values.forEach(val => {
      const option = document.createElement('option');
      option.value = val;
      option.textContent = val;
      select.appendChild(option);
    });
    
    // Tenta restaurar o valor selecionado anteriormente
    select.value = currentValue;
  });
}

// Chamada pelo HTML (onchange) quando um filtro é alterado
function handleFilterChange(field, value) {
  if (!value) {
    // Se selecionou a opção vazia, remove o filtro ativo
    delete window.state.activeFilters[field];
  } else {
    // Adiciona ou atualiza o filtro ativo
    window.state.activeFilters[field] = value;
  }
  applyFilters();
}

// Executa a lógica de filtragem combinada sobre os dados brutos (window.state.rawData)
function applyFilters() {
  // Gera window.state.filteredData a partir do rawData
  window.state.filteredData = window.state.rawData.filter(row => {
    // Retorna true apenas se a linha passar em TODOS os filtros ativos (lógica AND)
    return Object.entries(window.state.activeFilters).every(([key, val]) => {
      // Formata o valor da linha para comparação (tratando vazios)
      const rowVal = (row[key] || '(Vazio)').toString().trim();
      return rowVal === val;
    });
  });

  // Atualiza o contador de registros visíveis na UI
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

  // Notifica que os dados filtrados mudaram para engatilhar a renderização
  notifyStateChange();
}

// Limpa todos os filtros ativos e reseta os selects na UI
function resetFilters() {
  window.state.activeFilters = {};
  // Reseta visualmente os selects
  ['filterUnidade', 'filterStatus', 'filterCategoria', 'filterContratacao'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  applyFilters(); // Re-aplica (vai mostrar tudo)
}

/* --- 📊 RENDERIZAÇÃO DE UI (KPIs E TABELAS) --- */
// Calcula e renderiza os cards de KPI baseados nos dados filtrados
function renderKPIs() {
  const data = window.state.filteredData;
  const totalRecords = data.length;
  
  // Somatórios usando reduce sobre o float puro
  const totalValue = data.reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const capitalValue = data
    .filter(row => row['Categoria Econômica'] === 'Despesas de Capital')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);
  const currentValue = data
    .filter(row => row['Categoria Econômica'] === 'Despesas Correntes')
    .reduce((acc, row) => acc + (row['Valor Estimado'] || 0), 0);

  // Seleciona elementos HTML
  const totalRecEl = document.getElementById('kpiTotalRecords');
  const totalValEl = document.getElementById('kpiTotalValue');
  const capValEl = document.getElementById('kpiCapitalValue');
  const curValEl = document.getElementById('kpiCurrentValue');

  // Atualiza conteúdo formatado
  if (totalRecEl) totalRecEl.textContent = totalRecords;
  if (totalValEl) totalValEl.textContent = formatBRL(totalValue);
  if (capValEl) capValEl.textContent = formatBRL(capitalValue);
  if (curValEl) curValEl.textContent = formatBRL(currentValue);
}

// Agrupa dados e renderiza a Tabela Sintética (Resumo)
function renderSyntheticTable() {
  const tbody = document.getElementById('syntheticTableBody');
  const tfoot = document.getElementById('syntheticTableFooter');
  if (!tbody || !tfoot) return;

  // Limpa conteúdo anterior
  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const data = window.state.filteredData;
  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="px-4 py-8 text-center text-gray-400">Nenhum dado disponível para os filtros selecionados.</td></tr>`;
    return;
  }

  // Lógica de Agrupamento (Combinação: Contratação + Categoria Econômica)
  const totalValAll = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);
  const summaryMap = {};

  data.forEach(item => {
    const contratacao = (item['Contratação'] && item['Contratação'].toString().trim()) ? item['Contratação'].toString().trim() : '(Não Informado)';
    const categoria = (item['Categoria Econômica'] && item['Categoria Econômica'].toString().trim()) ? item['Categoria Econômica'].toString().trim() : '(Não Informado)';
    // Chave composta para agrupamento
    const key = `${contratacao}||${categoria}`;

    if (!summaryMap[key]) {
      summaryMap[key] = { contratacao, categoria, count: 0, totalValue: 0 };
    }
    summaryMap[key].count += 1;
    summaryMap[key].totalValue += (item['Valor Estimado'] || 0);
  });

  // Converte mapa para array e ordena por valor total descrescente
  const summaryArray = Object.values(summaryMap).sort((a, b) => b.totalValue - a.totalValue);

  // Renderiza linhas do corpo (tbody)
  summaryArray.forEach(row => {
    // Calcula porcentagem do grupo em relação ao total visível
    const pct = totalValAll > 0 ? (row.totalValue / totalValAll) * 100 : 0;
    const tr = document.createElement('tr');
    tr.className = "hover:bg-gray-50 border-b border-gray-100 transition-colors";
    
    // Classes CSS dinâmicas para o badge da categoria
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

  // Renderiza linha de total no rodapé (tfoot)
  tfoot.innerHTML = `
    <tr>
      <td colspan="2" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL SINTÉTICO (Visível):</td>
      <td class="px-4 py-3 text-center font-bold text-blue-700">${data.length}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">${formatBRL(totalValAll)}</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700">100,00%</td>
    </tr>
  `;
}

// Renderiza a Tabela Analítica (Detalhamento item a item)
function renderAnalyticalTable() {
  const tbody = document.getElementById('analyticalTableBody');
  const tfoot = document.getElementById('analyticalTableFooter');
  if (!tbody || !tfoot) return;

  // Limpa conteúdo anterior
  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const data = window.state.filteredData;
  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="px-4 py-8 text-center text-gray-400">Nenhum registro encontrado para os filtros aplicados.</td></tr>`;
    return;
  }

  // Renderiza linhas do corpo (tbody)
  data.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-blue-50/50 transition-colors border-b border-gray-100";
    
    // Classes CSS dinâmicas para o badge do status
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

  // Calcula total analítico visível
  const totalVal = data.reduce((sum, item) => sum + (item['Valor Estimado'] || 0), 0);

  tfoot.innerHTML = `
    <tr>
      <td colspan="4" class="px-4 py-3 text-right font-bold text-gray-800">TOTAL ANALÍTICO (${data.length} ITENS):</td>
      <td class="px-4 py-3 text-right font-bold text-blue-700 text-sm">${formatBRL(totalVal)}</td>
      <td colspan="5" class="px-4 py-3"></td>
    </tr>
  `;
}

/* --- 📤 EXPORTAÇÃO DE DADOS PARA EXCEL LOCAL --- */
// Gera um arquivo .xlsx contendo exatamente os dados visíveis (filtrados) na tela analítica
function exportToExcel() {
  if (!window.state.filteredData || window.state.filteredData.length === 0) {
    alert("Não há dados visíveis para exportar. Remova alguns filtros.");
    return;
  }

  // Prepara dados limpando metadados internos (__id) antes de gerar a planilha
  const exportData = window.state.filteredData.map(row => {
    const cleanRow = { ...row };
    delete cleanRow.__id; // Remove ID local
    return cleanRow;
  });

  // Usa biblioteca XLSX para criar workbook e worksheet
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Demandas Analíticas Filtradas");

  // Gera download do arquivo no navegador
  const dataSufixo = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  XLSX.writeFile(workbook, `Demandas_SEGES_2026_Export_${dataSufixo}.xlsx`);
}

// 🌐 Exportar funções cruciais para o escopo global (window)
// Isso é necessário porque o HTML Tailwind usa atributos onclick="funcao()" diretamente
window.loadDataFromGithub = loadDataFromGithub;
window.saveToGithub = saveToGithub; // Agora vinculado ao botão "Sincronizar"
window.handleFileUpload = handleFileUpload;
window.handleFilterChange = handleFilterChange;
window.resetFilters = resetFilters;
window.exportToExcel = exportToExcel;
window.clearGithubToken = clearGithubToken;

/* --- 🚀 INICIALIZAÇÃO AUTOMÁTICA NA CARGA DA PÁGINA --- */
document.addEventListener('DOMContentLoaded', () => {
  // Executa o PULL inicial do GitHub para preencher o painel
  loadDataFromGithub();
});

});```