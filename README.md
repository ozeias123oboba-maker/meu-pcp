[index.html](https://github.com/user-attachments/files/28226758/index.html)
<!DOCTYPE html>
<html lang="pt-BR" data-theme="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PCP Flow</title>
  <link rel="preconnect" href="https://api.fontshare.com">
  <link href="https://api.fontshare.com/v2/css?f[]=satoshi@400,500,700,900&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js" defer></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js" defer></script>
  <style>
    :root, [data-theme="light"] {
      --text-xs: 0.7rem;
      --text-sm: 0.8rem;
      --text-base: 0.9rem;
      --text-lg: 1.1rem;
      --text-xl: 1.5rem;
      --space-1: 0.2rem; --space-2: 0.4rem; --space-3: 0.6rem; --space-4: 0.8rem; --space-5: 1rem; --space-6: 1.2rem;
      --color-bg: #f7f6f2; --color-surface: #f9f8f5; --color-surface-2: #fbfbf9; --color-surface-offset: #f3f0ec; --color-border: #d4d1ca;
      --color-text: #28251d; --color-text-muted: #6b6963; --color-text-faint: #9a9892; --color-text-inverse: #f9f8f4;
      --color-primary: #01696f; --color-primary-hover: #0c4e54; --color-primary-highlight: #cedcd8;
      --color-success: #437a22; --color-warning: #da7101; --color-error: #a13544; --color-blue: #006494;
      --radius-sm: 0.25rem; --radius-md: 0.4rem; --radius-lg: 0.6rem; --radius-xl: 0.8rem; --radius-full: 9999px;
      --shadow-sm: 0 1px 2px rgb(40 37 29 / 0.06); --shadow-md: 0 4px 12px rgb(40 37 29 / 0.08);
      --transition: 150ms ease-out;
      --sidebar-width: 240px;
      --gantt-row-height: 36px;
    }
    [data-theme="dark"] {
      --color-bg: #171614; --color-surface: #1c1b19; --color-surface-2: #201f1d; --color-surface-offset: #22211f; --color-border: #393836;
      --color-text: #e4e1dc; --color-text-muted: #aaa7a2; --color-text-faint: #7b7874; --color-text-inverse: #1c1b19;
      --color-primary: #4f98a3; --color-primary-hover: #227f8b; --color-primary-highlight: #313b3b;
      --color-success: #6daa45; --color-warning: #fdab43; --color-error: #dd6974; --color-blue: #5591c7;
    }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    html, body { min-height: 100%; overflow-x: hidden; }
    body {
      font-family: 'Satoshi', sans-serif;
      font-size: var(--text-base);
      line-height: 1.4;
      color: var(--color-text);
      background: var(--color-bg);
    }
    button, input, select { font: inherit; color: inherit; }
    button { cursor: pointer; border: 0; }
    input, select { border: 1px solid var(--color-border); background: var(--color-surface-2); border-radius: var(--radius-md); padding: 0.5rem 0.7rem; }
    .app { display: grid; grid-template-columns: var(--sidebar-width) minmax(0, 1fr); min-height: 100dvh; }
    .sidebar {
      grid-row: 1 / -1; overflow-y: auto; background: var(--color-surface);
      border-right: 1px solid var(--color-border); padding: var(--space-5);
    }
    .brand { display: flex; align-items: center; gap: var(--space-2); margin-bottom: var(--space-6); }
    .brand svg { width: 32px; height: 32px; color: var(--color-primary); }
    .brand h1 { font-size: var(--text-lg); }
    .nav-group { display: grid; gap: var(--space-2); margin-bottom: var(--space-4); }
    .nav-chip { display: flex; justify-content: space-between; padding: 0.6rem 0.8rem; border-radius: var(--radius-md); background: var(--color-surface-2); border: 1px solid var(--color-border); }
    .nav-chip strong { font-size: var(--text-xs); }
    .header {
      display: flex; justify-content: space-between; align-items: center; padding: var(--space-4) var(--space-6);
      background: var(--color-surface); border-bottom: 1px solid var(--color-border);
    }
    .header-title h2 { font-size: var(--text-lg); }
    .header-actions { display: flex; gap: var(--space-2); align-items: center; flex-wrap: wrap; justify-content: flex-end; }
    .btn { padding: 0.5rem 0.8rem; border-radius: var(--radius-md); font-size: var(--text-xs); font-weight: 700; transition: all var(--transition); }
    .btn-primary { background: var(--color-primary); color: var(--color-text-inverse); }
    .btn-secondary { background: var(--color-surface-offset); border: 1px solid var(--color-border); }
    .main { overflow-y: auto; padding: var(--space-4); }
    .panel-grid { display: grid; gap: var(--space-4); }
    .kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: var(--space-3); }
    .card { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius-lg); padding: var(--space-4); box-shadow: var(--shadow-sm); }
    .kpi-label { color: var(--color-text-muted); font-size: 0.65rem; text-transform: uppercase; margin-bottom: 2px; }
    .kpi-value { font-size: var(--text-lg); font-weight: 900; }
    .charts { display: grid; grid-template-columns: 1fr 1fr; gap: var(--space-3); }
    .chart-card canvas { height: 200px !important; }
    .toolbar { display: grid; grid-template-columns: 1.5fr 2fr 1fr 1fr 1fr auto; gap: var(--space-2); align-items: end; margin-bottom: var(--space-3); }
    .field { display: grid; gap: 2px; }
    .field label { font-size: 0.65rem; color: var(--color-text-muted); text-transform: uppercase; }
    
    .filter-section { 
      background: var(--color-surface-offset); 
      padding: var(--space-3); 
      border-radius: var(--radius-lg); 
      margin-bottom: var(--space-3);
      display: grid;
      grid-template-columns: 1fr 1fr 1fr auto;
      gap: var(--space-2);
      align-items: end;
    }
    
    .table-wrap { overflow: auto; max-height: 400px; border: 1px solid var(--color-border); border-radius: var(--radius-md); }
    table { width: 100%; border-collapse: collapse; font-size: var(--text-sm); }
    thead th { position: sticky; top: 0; background: var(--color-surface-offset); padding: 0.6rem; text-align: left; font-size: 0.65rem; color: var(--color-text-muted); border-bottom: 1px solid var(--color-border); }
    td { padding: 0.5rem 0.6rem; border-bottom: 1px solid var(--color-border); white-space: nowrap; }
    .status-pill { padding: 0.2rem 0.5rem; border-radius: var(--radius-full); font-size: 0.65rem; font-weight: 700; }
    .status-em-producao { background: #fff4e6; color: #d9480f; }
    .status-a-produzir { background: #e7f5ff; color: #1971c2; }
    .status-concluido { background: #ebfbee; color: #2b8a3e; }
    .status-atrasado { background: #fff5f5; color: #c92a2a; border: 1px solid #ffa8a8; }
    .gantt-wrap { overflow: auto; padding-top: var(--space-2); }
    .gantt-row { display: grid; grid-template-columns: 200px 80px 1fr; align-items: center; margin-bottom: 4px; }
    .timeline { height: 24px; background: var(--color-surface-offset); border-radius: 4px; position: relative; }
    .bar { position: absolute; top: 4px; height: 16px; border-radius: 8px; font-size: 0.6rem; padding: 0 6px; display: flex; align-items: center; color: white; font-weight: bold; }
    .bar-em-producao { background: var(--color-warning); }
    .bar-a-produzir { background: var(--color-blue); }
    .bar-concluido { background: var(--color-success); }
    .bar.atrasado { border: 2px solid var(--color-error); }
    
    .filter-badge {
      display: inline-block;
      background: var(--color-primary);
      color: white;
      padding: 0.3rem 0.6rem;
      border-radius: var(--radius-full);
      font-size: 0.65rem;
      margin-right: 0.4rem;
      margin-bottom: 0.3rem;
    }
    
    .filter-badge button {
      background: none;
      border: none;
      color: white;
      cursor: pointer;
      margin-left: 0.4rem;
      font-weight: bold;
    }
    
    @media (max-width: 1024px) { 
      .charts, .kpis { grid-template-columns: 1fr 1fr; } 
      .toolbar { grid-template-columns: 1fr 1fr; }
      .filter-section { grid-template-columns: 1fr 1fr; }
    }

    .page-shell { width: min(1400px, 100%); margin: 0 auto; }
    .metrics, .filters, .form-grid { min-width: 0; }
    .metrics { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: var(--space-4); }
    .card, .chart-card, .table-card, .gantt-card { min-width: 0; }
    canvas { max-width: 100% !important; }
    table { width: 100%; min-width: 760px; }
    @media (max-width: 1100px) {
      .app { grid-template-columns: 1fr; }
      .sidebar { grid-row: auto; border-right: 0; border-bottom: 1px solid var(--color-border); }
      .charts { grid-template-columns: 1fr; }
      .metrics { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
    @media (max-width: 720px) {
      .header { flex-direction: column; align-items: flex-start; gap: var(--space-3); }
      .header-actions { width: 100%; justify-content: flex-start; }
      .metrics { grid-template-columns: 1fr; }
      .form-grid, .filters { grid-template-columns: 1fr !important; }
      .content { padding: var(--space-4); }
      .sidebar { padding: var(--space-4); }
    }

  </style>
</head>
<body>
  <div class="app">
    <aside class="sidebar">
      <div class="brand">
        <svg viewBox="0 0 64 64" fill="none"><rect x="8" y="10" width="20" height="44" rx="10" stroke="currentColor" stroke-width="6"></rect><path d="M36 14h18v10H36zM36 28h12v10H36zM36 42h18v10H36z" fill="currentColor"></path></svg>
        <h1>PCP Flow</h1>
      </div>
      <div class="nav-group">
        <div class="nav-chip"><strong>Pedidos Ativos</strong><span id="navAtivos">0</span></div>
        <div class="nav-chip"><strong>Atrasados</strong><span id="navAtrasados">0</span></div>
        <div class="nav-chip"><strong>Concluídos</strong><span id="navConcluidos">0</span></div>
      </div>
      <div class="card" style="padding: var(--space-3);">
        <p class="kpi-label">Importação</p>
        <p style="font-size: 0.7rem; color: var(--color-text-muted);">Colunas: cliente, produto, início, previsão, etapa.</p>
        <input id="excelInput" type="file" accept=".xlsx,.xls,.csv" style="width: 100%; margin-top: 0.5rem; font-size: 0.7rem;">
      </div>
      <button class="btn btn-secondary" id="clearDataBtn" style="width: 100%; margin-top: 1rem; color: var(--color-error);">Limpar Tudo</button>
    </aside>

    <div style="display: flex; flex-direction: column; overflow: hidden;">
      <header class="header">
        <div class="header-title"><h2>Controle de Produção</h2></div>
        <div class="header-actions">
          <button class="btn btn-secondary" id="exportBtn">Exportar Excel</button>
          <button class="btn btn-secondary" id="themeBtn">Tema</button>
        </div>
      </header>

      <main class="main">
        <div class="panel-grid">
          <div class="kpis">
            <div class="card"><p class="kpi-label">Total</p><div class="kpi-value" id="kpiTotal">0</div></div>
            <div class="card"><p class="kpi-label">% Atraso</p><div class="kpi-value" id="kpiAtraso">0%</div></div>
            <div class="card"><p class="kpi-label">% Produção</p><div class="kpi-value" id="kpiProducao">0%</div></div>
            <div class="card"><p class="kpi-label">% Concluído</p><div class="kpi-value" id="kpiConcluidas">0%</div></div>
          </div>

          <div class="charts">
            <div class="card chart-card"><canvas id="statusChart"></canvas></div>
            <div class="card chart-card"><canvas id="mixChart"></canvas></div>
          </div>

          <div class="card">
            <div class="toolbar">
              <div class="field"><label>Cliente</label><input id="cliente" type="text" placeholder="Nome do cliente"></div>
              <div class="field"><label>Produto</label><input id="produto" type="text" placeholder="Descrição"></div>
              <div class="field"><label>Início</label><input id="inicio" type="date"></div>
              <div class="field"><label>Previsão</label><input id="previsao" type="date"></div>
              <div class="field"><label>Etapa</label>
                <select id="etapa">
                  <option value="A produzir">A produzir</option>
                  <option value="Em produção">Em produção</option>
                  <option value="Concluído">Concluído</option>
                </select>
              </div>
              <button class="btn btn-primary" id="addRowBtn">Adicionar</button>
            </div>

            <div class="filter-section">
              <div class="field"><label>Filtrar Cliente</label><input id="filterCliente" type="text" placeholder="Digite para filtrar..."></div>
              <div class="field"><label>Filtrar Produto</label><input id="filterProduto" type="text" placeholder="Digite para filtrar..."></div>
              <div class="field"><label>Filtrar Etapa</label>
                <select id="filterEtapa">
                  <option value="">Todas as etapas</option>
                  <option value="A produzir">A produzir</option>
                  <option value="Em produção">Em produção</option>
                  <option value="Concluído">Concluído</option>
                </select>
              </div>
              <button class="btn btn-secondary" id="clearFiltersBtn">Limpar Filtros</button>
            </div>

            <div id="activeFilters" style="margin-bottom: var(--space-2); min-height: 20px;"></div>

            <div class="table-wrap">
              <table>
                <thead>
                  <tr><th>Cliente</th><th>Produto</th><th>Início</th><th>Previsão</th><th>Etapa</th><th>Atraso</th><th>Ações</th></tr>
                </thead>
                <tbody id="tableBody"></tbody>
              </table>
            </div>
            <div id="resultCount" style="font-size: 0.7rem; color: var(--color-text-muted); margin-top: 0.5rem;"></div>
          </div>

          <div class="card">
            <p class="kpi-label">Cronograma Gantt</p>
            <div id="gantt" class="gantt-wrap"></div>
          </div>
        </div>
      </main>
    </div>
  </div>

  <script>
    let data = [];
    let statusChart, mixChart;
    const STORAGE_KEY = 'pcp_flow_data';
    
    // Estado dos filtros
    let filters = {
      cliente: '',
      produto: '',
      etapa: ''
    };

    const today = new Date();
    today.setHours(0,0,0,0);

    function save() { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }
    function load() { 
      const saved = localStorage.getItem(STORAGE_KEY);
      data = saved ? JSON.parse(saved) : [];
    }

    const iso = (d) => d instanceof Date ? d.toISOString().split('T')[0] : d;
    const fmt = (s) => new Date(s + 'T00:00:00');
    const human = (s) => s ? fmt(s).toLocaleDateString('pt-BR') : '-';
    const diffDays = (a, b) => Math.ceil((fmt(a) - fmt(b)) / 86400000);

    function delayDays(item) {
      if (!item.previsao || item.etapa === 'Concluído') return 0;
      const delay = diffDays(iso(today), item.previsao);
      return delay > 0 ? delay : 0;
    }

    // Função para filtrar dados
    function getFilteredData() {
      return data.filter(item => {
        const matchCliente = item.cliente.toLowerCase().includes(filters.cliente.toLowerCase());
        const matchProduto = item.produto.toLowerCase().includes(filters.produto.toLowerCase());
        const matchEtapa = !filters.etapa || item.etapa === filters.etapa;
        return matchCliente && matchProduto && matchEtapa;
      });
    }

    // Atualizar display de filtros ativos
    function updateActiveFilters() {
      const container = document.getElementById('activeFilters');
      const active = [];
      
      if (filters.cliente) active.push(`Cliente: ${filters.cliente}`);
      if (filters.produto) active.push(`Produto: ${filters.produto}`);
      if (filters.etapa) active.push(`Etapa: ${filters.etapa}`);
      
      if (active.length === 0) {
        container.innerHTML = '';
      } else {
        container.innerHTML = active.map(f => {
          const key = f.split(':')[0].toLowerCase();
          return `<span class="filter-badge">${f} <button onclick="clearFilter('${key}')">✕</button></span>`;
        }).join('');
      }
    }

    window.clearFilter = (key) => {
      filters[key] = '';
      document.getElementById('filter' + key.charAt(0).toUpperCase() + key.slice(1)).value = '';
      render();
    };

    function render() {
      save();
      const filteredData = getFilteredData();
      const tbody = document.getElementById('tableBody');
      const metrics = { total: data.length, atraso: 0, producao: 0, concluido: 0, aProduzir: 0 };

      // Calcular métricas com base em TODOS os dados (não apenas filtrados)
      data.forEach(item => {
        const atraso = delayDays(item);
        if (item.etapa === 'Concluído') metrics.concluido++;
        else {
          if (atraso > 0) metrics.atraso++;
          if (item.etapa === 'Em produção') metrics.producao++;
          else metrics.aProduzir++;
        }
      });

      // Renderizar apenas dados filtrados
      tbody.innerHTML = filteredData.map((item, idx) => {
        const realIdx = data.indexOf(item);
        const atraso = delayDays(item);
        
        return `<tr>
          <td title="${item.cliente}">${item.cliente.substring(0, 20)}${item.cliente.length > 20 ? '...' : ''}</td>
          <td title="${item.produto}">${item.produto.substring(0, 30)}${item.produto.length > 30 ? '...' : ''}</td>
          <td>${human(item.inicio)}</td>
          <td>${human(item.previsao)}</td>
          <td><span class="status-pill ${atraso > 0 && item.etapa !== 'Concluído' ? 'status-atrasado' : 'status-' + item.etapa.toLowerCase().replace(/ /g, '-')}">${item.etapa}</span></td>
          <td>${atraso} d</td>
          <td>
            <button onclick="editRow(${realIdx})" style="background:none; border:none; cursor:pointer;">✎</button>
            <button onclick="deleteRow(${realIdx})" style="background:none; border:none; cursor:pointer; color:red; margin-left:8px;">✕</button>
          </td>
        </tr>`;
      }).join('');

      // Mostrar contagem de resultados
      const resultCount = document.getElementById('resultCount');
      if (filteredData.length === 0 && Object.values(filters).some(f => f)) {
        resultCount.textContent = 'Nenhum resultado encontrado com os filtros aplicados.';
      } else if (filteredData.length !== data.length) {
        resultCount.textContent = `Mostrando ${filteredData.length} de ${data.length} registros.`;
      } else {
        resultCount.textContent = '';
      }

      document.getElementById('kpiTotal').innerText = metrics.total;
      document.getElementById('kpiAtraso').innerText = metrics.total ? Math.round((metrics.atraso / metrics.total) * 100) + '%' : '0%';
      document.getElementById('kpiProducao').innerText = metrics.total ? Math.round((metrics.producao / metrics.total) * 100) + '%' : '0%';
      document.getElementById('kpiConcluidas').innerText = metrics.total ? Math.round((metrics.concluido / metrics.total) * 100) + '%' : '0%';
      
      document.getElementById('navAtivos').innerText = metrics.total - metrics.concluido;
      document.getElementById('navAtrasados').innerText = metrics.atraso;
      document.getElementById('navConcluidos').innerText = metrics.concluido;

      updateCharts(metrics);
      renderGantt();
      updateActiveFilters();
    }

    function updateCharts(m) {
      const ctx1 = document.getElementById('statusChart').getContext('2d');
      const ctx2 = document.getElementById('mixChart').getContext('2d');
      const textColor = getComputedStyle(document.documentElement).getPropertyValue('--color-text');

      if (statusChart) statusChart.destroy();
      statusChart = new Chart(ctx1, {
        type: 'bar',
        data: {
          labels: ['Atraso', 'Produção', 'Concluído'],
          datasets: [{ data: [m.atraso, m.producao, m.concluido], backgroundColor: ['#a13544', '#da7101', '#437a22'] }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, title: { display: true, text: 'Status das Ordens', color: textColor } } }
      });

      if (mixChart) mixChart.destroy();
      mixChart = new Chart(ctx2, {
        type: 'doughnut',
        data: {
          labels: ['A produzir', 'Em produção', 'Concluído'],
          datasets: [{ data: [m.aProduzir, m.producao, m.concluido], backgroundColor: ['#006494', '#da7101', '#437a22'] }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: textColor } }, title: { display: true, text: 'Mix de Produção', color: textColor } } }
      });
    }

    function renderGantt() {
      const el = document.getElementById('gantt');
      const filteredData = getFilteredData();
      const valid = filteredData.filter(i => i.inicio && i.previsao).slice(0, 15);
      if (!valid.length) { el.innerHTML = '<p style="text-align:center; font-size:0.7rem; color:var(--color-text-faint);">Sem dados para o Gantt</p>'; return; }

      const minDate = new Date(Math.min(...valid.map(i => fmt(i.inicio))));
      const maxDate = new Date(Math.max(...valid.map(i => fmt(i.previsao))));
      const totalRange = Math.max(1, (maxDate - minDate) / 86400000);

      el.innerHTML = valid.map(item => {
        const left = ((fmt(item.inicio) - minDate) / 86400000 / totalRange) * 100;
        const width = ((fmt(item.previsao) - fmt(item.inicio)) / 86400000 / totalRange) * 100;
        const atraso = delayDays(item);
        const type = item.etapa.toLowerCase().replace(/ /g, '-');
        return `<div class="gantt-row">
          <div style="font-size:0.7rem; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;"><strong>${item.cliente}</strong></div>
          <div style="font-size:0.6rem; color:var(--color-text-muted);">${atraso} d</div>
          <div class="timeline">
            <div class="bar bar-${type} ${atraso > 0 && item.etapa !== 'Concluído' ? 'atrasado' : ''}" style="left:${left}%; width:${Math.max(width, 2)}%;"></div>
          </div>
        </div>`;
      }).join('');
    }

    window.deleteRow = (i) => { data.splice(i, 1); render(); };
    window.editRow = (i) => {
      const item = data[i];
      document.getElementById('cliente').value = item.cliente;
      document.getElementById('produto').value = item.produto;
      document.getElementById('inicio').value = item.inicio;
      document.getElementById('previsao').value = item.previsao;
      document.getElementById('etapa').value = item.etapa;
      data.splice(i, 1);
      render();
    };

    // Event listeners para filtros
    document.getElementById('filterCliente').addEventListener('input', (e) => {
      filters.cliente = e.target.value;
      render();
    });

    document.getElementById('filterProduto').addEventListener('input', (e) => {
      filters.produto = e.target.value;
      render();
    });

    document.getElementById('filterEtapa').addEventListener('change', (e) => {
      filters.etapa = e.target.value;
      render();
    });

    document.getElementById('clearFiltersBtn').addEventListener('click', () => {
      filters = { cliente: '', produto: '', etapa: '' };
      document.getElementById('filterCliente').value = '';
      document.getElementById('filterProduto').value = '';
      document.getElementById('filterEtapa').value = '';
      render();
    });

    document.getElementById('addRowBtn').addEventListener('click', () => {
      const item = {
        cliente: document.getElementById('cliente').value,
        produto: document.getElementById('produto').value,
        inicio: document.getElementById('inicio').value,
        previsao: document.getElementById('previsao').value,
        etapa: document.getElementById('etapa').value
      };
      if (!item.cliente || !item.produto) return alert('Preencha os campos!');
      data.unshift(item);
      render();
      ['cliente', 'produto', 'inicio', 'previsao'].forEach(id => document.getElementById(id).value = '');
    });

    document.getElementById('excelInput').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws);

      const mapped = json.map(row => {
        const find = (keys) => {
          const key = Object.keys(row).find(k => keys.some(s => k.toLowerCase().includes(s)));
          return key ? row[key] : '';
        };

        const parseDate = (val) => {
          if (!val) return '';
          if (typeof val === 'number') return iso(new Date((val - 25569) * 86400 * 1000));
          if (val.includes('/')) {
            const [d, m, y] = val.split('/');
            return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
          }
          return val;
        };

        return {
          cliente: find(['cliente', 'empresa']),
          produto: find(['produto', 'descri', 'item']),
          inicio: parseDate(find(['inicio', 'começo'])),
          previsao: parseDate(find(['previs', 'conclusão', 'fim'])),
          etapa: find(['etapa', 'status']) || 'A produzir'
        };
      }).filter(i => i.cliente);

      data = [...mapped, ...data];
      render();
    });

    document.getElementById('exportBtn').addEventListener('click', () => {
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "PCP");
      XLSX.writeFile(wb, "PCP_Flow_Export.xlsx");
    });

    document.getElementById('themeBtn').addEventListener('click', () => {
      const html = document.documentElement;
      html.setAttribute('data-theme', html.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
      render();
    });

    document.getElementById('clearDataBtn').addEventListener('click', () => {
      if(confirm('Deseja apagar todos os dados salvos?')) {
        data = [];
        filters = { cliente: '', produto: '', etapa: '' };
        render();
      }
    });

    window.onload = () => { load(); render(); };
  </script>
</body>
</html>
