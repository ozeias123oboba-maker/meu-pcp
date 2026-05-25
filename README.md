[index.html](https://github.com/user-attachments/files/28228940/index.html)
<!DOCTYPE html>
<html lang="pt-BR" data-theme="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PCP Flow - Painel</title>
  <link rel="preconnect" href="https://api.fontshare.com">
  <link href="https://api.fontshare.com/v2/css?f[]=satoshi@400,500,700,900&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js" defer></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js" defer></script>
  <style>
    :root,[data-theme="light"]{--bg:#f7f6f2;--surface:#fffdf9;--surface2:#faf7f2;--border:#d6d1c8;--text:#26231d;--muted:#6e695f;--primary:#01696f;--primary2:#0c4e54;--success:#437a22;--warning:#da7101;--error:#a13544;--blue:#006494;--shadow:0 4px 12px rgb(40 37 29/.08);--r:14px;--sidebar:240px;--gap:14px}
    [data-theme="dark"]{--bg:#171614;--surface:#1c1b19;--surface2:#21201e;--border:#383734;--text:#e7e3dc;--muted:#aba79f;--primary:#4f98a3;--primary2:#227f8b;--success:#6daa45;--warning:#fdab43;--error:#dd6974;--blue:#5591c7}
    *{box-sizing:border-box} html,body{margin:0;min-height:100%;background:var(--bg);color:var(--text);font-family:'Satoshi',system-ui,sans-serif} button,input,select{font:inherit} button{cursor:pointer;border:0}
    .shell{min-height:100dvh;display:grid;grid-template-columns:var(--sidebar) minmax(0,1fr)}
    .sidebar{background:var(--surface);border-right:1px solid var(--border);padding:18px;display:grid;gap:18px;align-content:start}
    .brand{display:flex;gap:10px;align-items:center}.brand h1{margin:0;font-size:1.15rem}.brand p{margin:2px 0 0;color:var(--muted);font-size:.86rem}
    .menu,.stats{display:grid;gap:10px}.menu button{padding:12px 14px;border-radius:12px;background:var(--surface2);border:1px solid var(--border);text-align:left;font-weight:700}.menu button.active{background:var(--primary);color:#fff;border-color:var(--primary)}
    .stats .mini{padding:12px 14px;border-radius:12px;background:var(--surface2);border:1px solid var(--border);display:flex;justify-content:space-between;gap:10px}
    .topbar{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:18px;border-bottom:1px solid var(--border);background:var(--surface);position:sticky;top:0;z-index:5;flex-wrap:wrap}
    .topbar h2{margin:0;font-size:1.25rem}.actions{display:flex;gap:10px;flex-wrap:wrap}.btn{padding:10px 14px;border-radius:12px;border:1px solid var(--border);background:var(--surface2);font-weight:700}.btn.primary{background:var(--primary);border-color:var(--primary);color:#fff}
    .main{min-width:0;overflow:auto}.content{padding:18px;display:grid;gap:16px}.screen{display:none;gap:16px}.screen.active{display:grid}
    .hero{padding:18px;border:1px solid var(--border);border-radius:var(--r);background:linear-gradient(180deg,var(--surface),var(--surface2))}.hero p{margin:8px 0 0;color:var(--muted);max-width:80ch}
    .kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:var(--gap)} .card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:16px;box-shadow:var(--shadow);min-width:0}
    .kpi-label{font-size:.66rem;text-transform:uppercase;color:var(--muted)} .kpi-value{font-size:1.5rem;font-weight:900;margin-top:10px}
    .charts{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(280px,.9fr);gap:var(--gap)} .chart-card canvas{height:240px!important;max-width:100%}
    .section-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:10px}.section-head h3{margin:0;font-size:1.05rem}.section-head p{margin:4px 0 0;color:var(--muted);font-size:.92rem}
    .grid-2{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(280px,.65fr);gap:var(--gap)}
    .toolbar,.filter-section{display:grid;gap:10px}.toolbar{grid-template-columns:1.3fr 1.8fr 1fr 1fr 1fr auto;align-items:end}.filter-section{grid-template-columns:1fr 1fr 1fr auto;align-items:end;background:var(--surface2);padding:14px;border-radius:var(--r)}
    .field{display:grid;gap:4px}.field label{font-size:.68rem;text-transform:uppercase;color:var(--muted)} input,select{border:1px solid var(--border);background:var(--surface2);border-radius:10px;padding:10px 12px;min-width:0}
    .table-wrap{overflow:auto;max-height:440px;border:1px solid var(--border);border-radius:12px}.table-wrap table{min-width:920px;width:100%;border-collapse:collapse;font-size:.92rem} thead th{position:sticky;top:0;background:var(--surface2);padding:10px;border-bottom:1px solid var(--border);text-align:left;font-size:.66rem;text-transform:uppercase;color:var(--muted)} td{padding:10px;border-bottom:1px solid var(--border);white-space:nowrap}
    .pill{padding:4px 8px;border-radius:999px;font-size:.72rem;font-weight:800}.p-em{background:#fff4e6;color:#d9480f}.p-a{background:#e7f5ff;color:#1971c2}.p-c{background:#ebfbee;color:#2b8a3e}.p-t{background:#fff5f5;color:#c92a2a;border:1px solid #ffa8a8}
    .gantt-wrap{overflow:auto;padding-top:4px}.gantt-row{display:grid;grid-template-columns:220px 80px 1fr;align-items:center;gap:10px;margin-bottom:6px}.timeline{height:24px;background:var(--surface2);border-radius:999px;position:relative}.bar{position:absolute;top:4px;height:16px;border-radius:999px;color:#fff;font-size:.62rem;font-weight:800;padding:0 8px;display:flex;align-items:center}.bar-p{background:var(--warning)} .bar-a{background:var(--blue)} .bar-c{background:var(--success)} .bar.delay{outline:2px solid var(--error)}
    .screen-nav{display:grid;gap:10px}.screen-link{padding:12px 14px;border:1px solid var(--border);background:var(--surface2);border-radius:12px;text-align:left;font-weight:800}.screen-link.active{background:var(--primary);color:#fff;border-color:var(--primary)}
    .report-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:var(--gap)} .report-card{padding:16px;border:1px solid var(--border);border-radius:var(--r);background:var(--surface)} .report-card span{display:block;color:var(--muted);font-size:.9rem}.report-card strong{display:block;font-size:1.6rem;margin-top:8px}
    .list{display:grid;gap:10px}.item{padding:12px 14px;border:1px solid var(--border);border-radius:12px;background:var(--surface)} .item strong{display:block}.item span{color:var(--muted);font-size:.92rem}
    .hidden{display:none!important}
    @media (max-width:1180px){.shell{grid-template-columns:1fr}.sidebar{border-right:0;border-bottom:1px solid var(--border)}.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.charts,.grid-2,.report-grid{grid-template-columns:1fr}.toolbar{grid-template-columns:1fr 1fr}.filter-section{grid-template-columns:1fr 1fr}.screen-nav{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @media (max-width:720px){.toolbar,.filter-section,.screen-nav{grid-template-columns:1fr}.kpis{grid-template-columns:1fr}}
  </style>
</head>
<body>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand"><div style="width:40px;height:40px;border-radius:12px;background:var(--primary);display:grid;place-items:center;color:#fff;font-weight:900">P</div><div><h1>PCP Flow</h1><p>Painel de produção</p></div></div>
      <div class="stats">
        <div class="mini"><strong>Pedidos Ativos</strong><span id="navAtivos">0</span></div>
        <div class="mini"><strong>Atrasados</strong><span id="navAtrasados">0</span></div>
        <div class="mini"><strong>Concluídos</strong><span id="navConcluidos">0</span></div>
      </div>
      <div class="screen-nav">
        <button class="screen-link active" data-screen="dashboard">Dashboard</button>
        <button class="screen-link" data-screen="pedidos">Pedidos</button>
        <button class="screen-link" data-screen="gantt">Gantt</button>
        <button class="screen-link" data-screen="relatorios">Relatórios</button>
      </div>
      <div class="card">
        <div class="field"><label>Importação</label><span style="color:var(--muted);font-size:.9rem">Colunas: cliente, produto, início, previsão, etapa.</span><input id="excelInput" type="file" accept=".xlsx,.xls,.csv"></div>
        <button class="btn" id="clearDataBtn" style="width:100%;margin-top:12px;color:var(--error)">Limpar tudo</button>
      </div>
    </aside>
    <div class="main">
      <div class="topbar"><h2>Controle de Produção</h2><div class="actions"><button class="btn" id="exportBtn">Exportar Excel</button><button class="btn" id="themeBtn">Tema</button></div></div>
      <div class="content">
        <section class="screen active" data-screen-panel="dashboard">
          <div class="hero"><strong>Visão geral do PCP</strong><p>Indicadores e gráficos ficam nesta tela. As demais áreas foram separadas para ficar mais fácil navegar e visualizar os dados.</p></div>
          <div class="kpis">
            <div class="card"><div class="kpi-label">Total</div><div class="kpi-value" id="kpiTotal">0</div></div>
            <div class="card"><div class="kpi-label">% Atraso</div><div class="kpi-value" id="kpiAtraso">0%</div></div>
            <div class="card"><div class="kpi-label">% Produção</div><div class="kpi-value" id="kpiProducao">0%</div></div>
            <div class="card"><div class="kpi-label">% Concluído</div><div class="kpi-value" id="kpiConcluidas">0%</div></div>
          </div>
          <div class="charts">
            <div class="card chart-card"><canvas id="statusChart"></canvas></div>
            <div class="card chart-card"><canvas id="mixChart"></canvas></div>
          </div>
        </section>
        <section class="screen" data-screen-panel="pedidos">
          <div class="hero"><strong>Pedidos e clientes</strong><p>Aqui você cadastra, importa e filtra os pedidos sem misturar com o Gantt.</p></div>
          <div class="grid-2">
            <div class="card">
              <div class="toolbar">
                <div class="field"><label>Cliente</label><input id="cliente" type="text" placeholder="Nome do cliente"></div>
                <div class="field"><label>Produto</label><input id="produto" type="text" placeholder="Descrição"></div>
                <div class="field"><label>Início</label><input id="inicio" type="date"></div>
                <div class="field"><label>Previsão</label><input id="previsao" type="date"></div>
                <div class="field"><label>Etapa</label><select id="etapa"><option value="A produzir">A produzir</option><option value="Em produção">Em produção</option><option value="Concluído">Concluído</option></select></div>
                <button class="btn primary" id="addRowBtn">Adicionar</button>
              </div>
              <div class="filter-section" style="margin-top:14px">
                <div class="field"><label>Filtrar cliente</label><input id="filterCliente" type="text" placeholder="Digite para filtrar..."></div>
                <div class="field"><label>Filtrar produto</label><input id="filterProduto" type="text" placeholder="Digite para filtrar..."></div>
                <div class="field"><label>Filtrar etapa</label><select id="filterEtapa"><option value="">Todas as etapas</option><option value="A produzir">A produzir</option><option value="Em produção">Em produção</option><option value="Concluído">Concluído</option></select></div>
                <button class="btn" id="clearFiltersBtn">Limpar filtros</button>
              </div>
              <div id="activeFilters" style="margin:10px 0;min-height:20px"></div>
              <div class="table-wrap"><table><thead><tr><th>Cliente</th><th>Produto</th><th>Início</th><th>Previsão</th><th>Etapa</th><th>Atraso</th><th>Ações</th></tr></thead><tbody id="tableBody"></tbody></table></div>
              <div id="resultCount" style="margin-top:8px;color:var(--muted);font-size:.86rem"></div>
            </div>
            <div class="card">
              <div class="section-head"><div><h3>Lista de clientes</h3><p>Resumo rápido dos clientes com pedidos.</p></div></div>
              <div class="list" id="clientList"></div>
            </div>
          </div>
        </section>
        <section class="screen" data-screen-panel="gantt">
          <div class="hero"><strong>Cronograma Gantt</strong><p>Esta tela ficou separada para facilitar o acompanhamento de prazos e atrasos.</p></div>
          <div class="card">
            <div class="section-head"><div><h3>Gantt</h3><p>Visualização horizontal dos períodos de cada pedido.</p></div></div>
            <div id="gantt" class="gantt-wrap"></div>
          </div>
        </section>
        <section class="screen" data-screen-panel="relatorios">
          <div class="hero"><strong>Relatórios rápidos</strong><p>Resumo executivo para reunião e análise rápida do cenário atual.</p></div>
          <div class="report-grid" id="reportGrid"></div>
          <div class="card"><div class="section-head"><div><h3>Observações automáticas</h3><p>Leitura resumida dos dados atuais.</p></div></div><div class="list" id="reportNotes"></div></div>
        </section>
      </div>
    </div>
  </div>
  <script>
    let data=[];let statusChart,mixChart;const STORAGEKEY='pcpflowdata';let filters={cliente:'',produto:'',etapa:''};const today=new Date();today.setHours(0,0,0,0);
    const normalize=s=>(s||'').toString().trim(); const safeDate=s=>{if(!s) return ''; const d=new Date(s+'T00:00:00'); return isNaN(d)?'':d.toISOString().slice(0,10)};
    const fmt=s=>s?new Date(s+'T00:00:00').toLocaleDateString('pt-BR'):''; const diffDays=(a,b)=>Math.ceil((new Date(a+'T00:00:00')-new Date(b+'T00:00:00'))/86400000);
    const delayDays=item=>{if(!item.previsao||item.etapa==='Concluído') return 0; const d=diffDays(today.toISOString().slice(0,10),item.previsao); return d>0?d:0};
    function save(){localStorage.setItem(STORAGEKEY,JSON.stringify(data));}
    function load(){const saved=localStorage.getItem(STORAGEKEY); data=saved?JSON.parse(saved):[]; render();}
    function getFilteredData(){return data.filter(item=>normalize(item.cliente).toLowerCase().includes(filters.cliente.toLowerCase())&&normalize(item.produto).toLowerCase().includes(filters.produto.toLowerCase())&&(!filters.etapa||item.etapa===filters.etapa));}
    function activePills(){const c=document.getElementById('activeFilters');const a=[]; if(filters.cliente) a.push(`Cliente: ${filters.cliente}`); if(filters.produto) a.push(`Produto: ${filters.produto}`); if(filters.etapa) a.push(`Etapa: ${filters.etapa}`); c.innerHTML=a.length? a.map(x=>`<span style='display:inline-block;background:var(--primary);color:#fff;padding:4px 10px;border-radius:999px;margin-right:6px;margin-bottom:6px;font-size:.78rem'>${x}</span>`).join(''):'';}
    function updateCharts(m){const c1=document.getElementById('statusChart').getContext('2d');const c2=document.getElementById('mixChart').getContext('2d');const tc=getComputedStyle(document.documentElement).getPropertyValue('--text'); if(statusChart) statusChart.destroy(); if(mixChart) mixChart.destroy(); statusChart=new Chart(c1,{type:'bar',data:{labels:['Atraso','Produção','Concluído'],datasets:[{data:[m.atraso,m.producao,m.concluido],backgroundColor:['#a13544','#da7101','#437a22']}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},title:{display:true,text:'Status das Ordens',color:tc}},scales:{x:{ticks:{color:tc}},y:{ticks:{color:tc}}}}}); mixChart=new Chart(c2,{type:'doughnut',data:{labels:['A produzir','Em produção','Concluído'],datasets:[{data:[m.aProduzir,m.producao,m.concluido],backgroundColor:['#006494','#da7101','#437a22']}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:tc}},title:{display:true,text:'Mix de Produção',color:tc}}}});}
    function renderGantt(){const el=document.getElementById('gantt');const rows=getFilteredData().filter(i=>i.inicio&&i.previsao); if(!rows.length){el.innerHTML='<p style="text-align:center;color:var(--muted)">Sem dados para o Gantt.</p>';return;} const min=new Date(Math.min(...rows.map(i=>new Date(i.inicio+'T00:00:00')))); const max=new Date(Math.max(...rows.map(i=>new Date(i.previsao+'T00:00:00')))); const range=Math.max(1,(max-min)/86400000); el.innerHTML=rows.map(item=>{const left=((new Date(item.inicio+'T00:00:00')-min)/86400000)/range*100; const width=(diffDays(item.previsao,item.inicio)/range)*100; const d=delayDays(item); const cls=item.etapa==='Concluído'?'c':item.etapa==='Em produção'?'p':'a'; return `<div class='gantt-row'><div style='overflow:hidden;text-overflow:ellipsis;white-space:nowrap'><strong>${item.cliente}</strong></div><div style='font-size:.72rem;color:var(--muted)'>${d}d</div><div class='timeline'><div class='bar bar-${cls} ${d>0&&item.etapa!=='Concluído'?'delay':''}' style='left:${left}%;width:${Math.max(width,2)}%'>${item.etapa}</div></div></div>`;}).join('');}
    function renderClientList(rows){const el=document.getElementById('clientList'); const map=new Map(); rows.forEach(r=>map.set(r.cliente,(map.get(r.cliente)||0)+1)); const arr=[...map.entries()].sort((a,b)=>b[1]-a[1]); el.innerHTML=arr.length?arr.map(([n,t])=>`<div class='item'><strong>${n}</strong><span>${t} pedido(s)</span></div>`).join(''):'<div class="item"><strong>Nenhum cliente</strong><span>Sem pedidos cadastrados.</span></div>';}
    function renderReports(rows,m){const eg=document.getElementById('reportGrid'); const notes=document.getElementById('reportNotes'); eg.innerHTML=[['Pedidos ativos',m.total],['Concluídos',m.concluido],['Atrasados',m.atrasoPedidos],['Dias de atraso',m.atrasoDias],['Em produção',m.producao],['A produzir',m.aProduzir]].map(([l,v])=>`<div class='report-card'><span>${l}</span><strong>${v}</strong></div>`).join(''); notes.innerHTML=[`Há ${m.total} pedido(s) cadastrados.`, `${m.concluido} concluído(s), ${m.producao} em produção e ${m.aProduzir} a produzir.`, m.atrasoPedidos?`${m.atrasoPedidos} pedido(s) em atraso, totalizando ${m.atrasoDias} dia(s).`:'Nenhum pedido em atraso no momento.'].map(x=>`<div class='item'>${x}</div>`).join('');}
    function renderTable(rows){const tbody=document.getElementById('tableBody');tbody.innerHTML=rows.map((item,i)=>{const atraso=delayDays(item);const p=item.etapa==='Concluído'?'p-c':item.etapa==='Em produção'?'p-em':'p-a'; const late=atraso>0&&item.etapa!=='Concluído'; return `<tr><td title='${item.cliente}'>${item.cliente}</td><td title='${item.produto}'>${item.produto}</td><td>${fmt(item.inicio)}</td><td>${fmt(item.previsao)}</td><td><span class='pill ${late?'p-t':p}'>${late?'Atrasado':item.etapa}</span></td><td>${atraso} d</td><td><button onclick='editRow(${i})'>✎</button> <button onclick='deleteRow(${i})' style='color:#c92a2a'>🗑</button></td></tr>`;}).join('');}
    function render(){save(); const filtered=getFilteredData(); const metrics={total:data.length, atraso:0, producao:0, concluido:0, aProduzir:0, atrasoPedidos:0, atrasoDias:0}; data.forEach(item=>{const a=delayDays(item); if(item.etapa==='Concluído') metrics.concluido++; else if(item.etapa==='Em produção') metrics.producao++; else metrics.aProduzir++; if(a>0&&item.etapa!=='Concluído'){metrics.atraso++; metrics.atrasoPedidos++; metrics.atrasoDias+=a}}); document.getElementById('kpiTotal').innerText=metrics.total; document.getElementById('kpiAtraso').innerText=metrics.total?Math.round(metrics.atraso/metrics.total*100):0+'%'; document.getElementById('kpiProducao').innerText=metrics.total?Math.round(metrics.producao/metrics.total*100):0+'%'; document.getElementById('kpiConcluidas').innerText=metrics.total?Math.round(metrics.concluido/metrics.total*100):0+'%'; document.getElementById('navAtivos').innerText=metrics.total-metrics.concluido; document.getElementById('navAtrasados').innerText=metrics.atrasoPedidos; document.getElementById('navConcluidos').innerText=metrics.concluido; document.getElementById('resultCount').innerText=filtered.length===data.length?`${data.length} registro(s).`:`Mostrando ${filtered.length} de ${data.length} registro(s).`; renderTable(filtered); updateCharts(metrics); renderGantt(); renderClientList(filtered); renderReports(filtered,metrics); activePills();}
    function addRow(){const item={cliente:document.getElementById('cliente').value.trim(),produto:document.getElementById('produto').value.trim(),inicio:safeDate(document.getElementById('inicio').value),previsao:safeDate(document.getElementById('previsao').value),etapa:document.getElementById('etapa').value}; if(!item.cliente||!item.produto){alert('Preencha cliente e produto.');return} data.unshift(item); ['cliente','produto','inicio','previsao'].forEach(id=>document.getElementById(id).value=''); render();}
    window.editRow=i=>{const it=data[i]; document.getElementById('cliente').value=it.cliente; document.getElementById('produto').value=it.produto; document.getElementById('inicio').value=it.inicio; document.getElementById('previsao').value=it.previsao; document.getElementById('etapa').value=it.etapa; data.splice(i,1); render();}
    window.deleteRow=i=>{data.splice(i,1); render();}
    function setupScreens(){const buttons=[...document.querySelectorAll('[data-screen]')]; const panels=[...document.querySelectorAll('[data-screen-panel]')]; const act=s=>{buttons.forEach(b=>b.classList.toggle('active',b.dataset.screen===s)); panels.forEach(p=>p.classList.toggle('active',p.dataset.screenPanel===s));}; buttons.forEach(b=>b.onclick=()=>act(b.dataset.screen));}
    document.getElementById('addRowBtn').addEventListener('click',addRow); document.getElementById('clearFiltersBtn').addEventListener('click',()=>{filters={cliente:'',produto:'',etapa:''}; document.getElementById('filterCliente').value=''; document.getElementById('filterProduto').value=''; document.getElementById('filterEtapa').value=''; render();}); document.getElementById('filterCliente').addEventListener('input',e=>{filters.cliente=e.target.value; render()}); document.getElementById('filterProduto').addEventListener('input',e=>{filters.produto=e.target.value; render()}); document.getElementById('filterEtapa').addEventListener('change',e=>{filters.etapa=e.target.value; render()});
    document.getElementById('clearDataBtn').addEventListener('click',()=>{if(confirm('Deseja apagar todos os dados?')){data=[]; render();}}); document.getElementById('themeBtn').addEventListener('click',()=>{document.documentElement.setAttribute('data-theme',document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark'); render();}); document.getElementById('exportBtn').addEventListener('click',()=>{const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'PCP'); XLSX.writeFile(wb,'PCPFlowExport.xlsx')});
    document.getElementById('excelInput').addEventListener('change',async e=>{const file=e.target.files[0]; if(!file) return; const buffer=await file.arrayBuffer(); const wb=XLSX.read(buffer,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; const json=XLSX.utils.sheet_to_json(ws); const find=(row,keys)=>{const k=Object.keys(row).find(x=>keys.some(s=>x.toLowerCase().includes(s))); return k?row[k]:''}; const parseDate=v=>{if(!v)return''; if(typeof v==='number') return new Date((v-25569)*86400*1000).toISOString().slice(0,10); const s=String(v); if(s.includes('/')){const [d,m,y]=s.split('/'); return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;} return safeDate(s)}; data=[...json.map(r=>({cliente:find(r,['cliente','empresa']),produto:find(r,['produto','descricao','item']),inicio:parseDate(find(r,['inicio','começo','comeco','start'])),previsao:parseDate(find(r,['previsao','previsão','fim','conclusao','conclusão'])),etapa:find(r,['etapa','status'])||'A produzir'})).filter(i=>i.cliente||i.produto), ...data]; render();});
    const seeds=[{cliente:'Cliente A',produto:'Peça 1',inicio:'2026-05-01',previsao:'2026-05-09',etapa:'Em produção'},{cliente:'Cliente B',produto:'Peça 2',inicio:'2026-05-02',previsao:'2026-05-10',etapa:'A produzir'},{cliente:'Cliente C',produto:'Peça 3',inicio:'2026-05-04',previsao:'2026-05-07',etapa:'Concluído'}]; if(!localStorage.getItem(STORAGEKEY)) data=seeds; setupScreens(); load();
  </script>
</body>
</html>
