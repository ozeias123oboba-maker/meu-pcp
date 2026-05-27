[index.html..html](https://github.com/user-attachments/files/28312466/index.html.html)
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PCP Flow - V5</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
:root{--bg:#f7f6f2;--surface:#fffdf9;--surface2:#faf7f2;--border:#d6d1c8;--text:#26231d;--muted:#6e695f;--primary:#01696f;--success:#437a22;--warning:#da7101;--error:#a13544;--blue:#006494;--shadow:0 4px 12px rgb(40 37 29/.08);--r:14px}
*{box-sizing:border-box}body{margin:0;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--text)}button,input,select{font:inherit}.app{display:grid;grid-template-columns:250px 1fr;min-height:100vh}.sidebar{background:var(--surface);border-right:1px solid var(--border);padding:16px;display:grid;gap:14px}.brand h1{margin:0;font-size:1.2rem}.mini{padding:10px 12px;background:var(--surface2);border:1px solid var(--border);border-radius:12px;display:flex;justify-content:space-between}.nav button,.btn{width:100%;padding:11px 12px;border:1px solid var(--border);background:var(--surface2);border-radius:12px;font-weight:700;cursor:pointer}.nav button.active{background:var(--primary);color:#fff;border-color:var(--primary)}.main{min-width:0}.top{display:flex;justify-content:space-between;gap:12px;align-items:center;padding:16px 18px;border-bottom:1px solid var(--border);background:var(--surface);position:sticky;top:0;z-index:2;flex-wrap:wrap}.top h2{margin:0}.content{padding:18px;display:grid;gap:16px}.screen{display:none}.screen.active{display:grid}.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--shadow);padding:16px}.hero{padding:16px;border:1px solid var(--border);border-radius:var(--r);background:linear-gradient(180deg,var(--surface),var(--surface2))}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.kpi{padding:14px;border:1px solid var(--border);border-radius:var(--r);background:var(--surface)}.kpi span{color:var(--muted);font-size:.8rem}.kpi strong{display:block;font-size:1.5rem;margin-top:8px}.grid2{display:grid;grid-template-columns:1.35fr .65fr;gap:12px}.toolbar,.filters{display:grid;gap:10px}.toolbar{grid-template-columns:1.2fr 1.4fr 1fr 1fr 1fr 1fr 1fr auto;align-items:end}.filters{grid-template-columns:1fr 1fr 1fr auto;align-items:end;background:var(--surface2);padding:12px;border-radius:var(--r)}.field{display:grid;gap:4px}.field label{font-size:.68rem;text-transform:uppercase;color:var(--muted)}input,select{width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:10px;background:var(--surface2)}.table-wrap{overflow:auto;border:1px solid var(--border);border-radius:12px}.table-wrap table{min-width:1200px;width:100%;border-collapse:collapse;font-size:.92rem}th,td{padding:10px;border-bottom:1px solid var(--border);white-space:nowrap;text-align:left}thead th{position:sticky;top:0;background:var(--surface2);text-transform:uppercase;font-size:.68rem;color:var(--muted)}.pill{display:inline-block;padding:4px 8px;border-radius:999px;font-size:.72rem;font-weight:800}.a{background:#e7f5ff;color:#1971c2}.t{background:#fff5f5;color:#c92a2a}.e{background:#fff4e6;color:#d9480f}.g{background:#f3e8ff;color:#7e22ce}.p{background:#fee2e2;color:#b91c1c}.pa{background:#fef3c7;color:#a16207}.c{background:#ebfbee;color:#2b8a3e}.d{background:#dbeafe;color:#1d4ed8}.urgent{border-color:var(--error)}.urgent::before{content:'⚠';position:absolute;top:8px;right:10px;color:var(--error)}.kanban{display:grid;grid-template-columns:repeat(8,minmax(220px,1fr));gap:12px;overflow:auto}.lane{min-width:220px;background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:12px;display:grid;gap:10px}.lane h4{margin:0}.lane small{color:var(--muted)}.lane-list{display:grid;gap:10px}.cardjob{position:relative;background:var(--surface2);border:1px solid var(--border);border-radius:12px;padding:12px;display:grid;gap:5px;cursor:grab}.cardjob strong{font-size:.92rem}.cardjob span{font-size:.86rem;color:var(--muted)}.overdue{border-color:var(--error);background:#fff5f5}.alert{padding:12px 14px;border:1px solid var(--error);background:#fff5f5;color:var(--error);border-radius:12px;font-weight:800}.report-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.report-card{padding:14px;border:1px solid var(--border);border-radius:var(--r);background:var(--surface)}.report-card span{color:var(--muted);font-size:.9rem}.report-card strong{display:block;font-size:1.5rem;margin-top:8px}.hidden{display:none!important}@media(max-width:1180px){.app{grid-template-columns:1fr}.kpis,.report-grid{grid-template-columns:repeat(2,1fr)}.grid2{grid-template-columns:1fr}.toolbar{grid-template-columns:1fr 1fr}.filters{grid-template-columns:1fr 1fr 1fr auto}.kanban{grid-template-columns:repeat(4,minmax(220px,1fr))}}@media(max-width:720px){.kpis,.report-grid,.toolbar,.filters{grid-template-columns:1fr}.kanban{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
<aside class="sidebar">
<div class="brand"><h1>PCP Flow</h1><div style="color:var(--muted);font-size:.9rem">Controle de produção</div></div>
<div class="mini"><strong>Ativos</strong><span id="navAtivos">0</span></div>
<div class="mini"><strong>Atrasados</strong><span id="navAtrasados">0</span></div>
<div class="mini"><strong>Concluídos</strong><span id="navConcluidos">0</span></div>
<div class="nav">
<button data-screen="dashboard" class="active">Dashboard</button>
<button data-screen="pedidos">Pedidos</button>
<button data-screen="kanban">Kanban</button>
<button data-screen="gantt">Gantt</button>
<button data-screen="relatorios">Relatórios</button>
</div>
<div class="card">
<div class="field"><label>Importar planilha</label><input id="excelInput" type="file" accept=".xlsx,.xls,.csv"></div>
<button class="btn" id="clearDataBtn">Limpar tudo</button>
</div>
</aside>
<main class="main">
<div class="top"><h2>Controle de Produção</h2><div style="display:flex;gap:10px;flex-wrap:wrap"><button class="btn" id="exportBtn">Exportar Excel</button><button class="btn" id="themeBtn">Tema</button></div></div>
<div class="content">
<section class="screen active" data-panel="dashboard">
<div class="hero"><strong>Visão geral do PCP</strong><p>Indicadores, atrasos e acompanhamento das ordens em um painel único.</p></div>
<div class="kpis"><div class="kpi"><span>Total</span><strong id="kpiTotal">0</strong></div><div class="kpi"><span>% Atraso</span><strong id="kpiAtraso">0%</strong></div><div class="kpi"><span>% Produção</span><strong id="kpiProducao">0%</strong></div><div class="kpi"><span>% Concluído</span><strong id="kpiConcluidas">0%</strong></div></div>
<div class="card"><canvas id="statusChart" height="120"></canvas></div>
</section>
<section class="screen" data-panel="pedidos">
<div class="hero"><strong>Pedidos</strong><p>Cadastro manual e importação por planilha funcionando em um fluxo simples.</p></div>
<div class="grid2">
<div class="card">
<div class="toolbar">
<div class="field"><label>Cliente</label><input id="cliente" type="text"></div>
<div class="field"><label>Produto</label><input id="produto" type="text"></div>
<div class="field"><label>Responsável</label><input id="responsavel" type="text"></div>
<div class="field"><label>Prioridade</label><select id="prioridade"><option>Baixa</option><option>Média</option><option>Alta</option><option>Urgente</option></select></div>
<div class="field"><label>Pedido</label><input id="pedido" type="text"></div>
<div class="field"><label>Entrada</label><input id="entrada" type="date"></div>
<div class="field"><label>Entrega</label><input id="entrega" type="date"></div>
<button class="btn" id="addRowBtn">Adicionar</button>
</div>
<div class="filters" style="margin-top:12px">
<div class="field"><label>Filtrar cliente</label><input id="filterCliente" type="text"></div>
<div class="field"><label>Filtrar produto</label><input id="filterProduto" type="text"></div>
<div class="field"><label>Filtrar etapa</label><select id="filterEtapa"><option value="">Todas</option><option>A produzir</option><option>Atrasado</option><option>Em produção</option><option>Galvanização</option><option>Pintura</option><option>Pátio</option><option>Concluído</option><option>Entregue</option></select></div>
<button class="btn" id="clearFiltersBtn">Limpar</button>
</div>
<div class="filters" style="margin-top:10px;grid-template-columns:repeat(4,1fr)">
<button class="btn" id="showOverdueBtn">Só atrasados</button>
<button class="btn" id="showPendingBtn">Pendências</button>
<button class="btn" id="showAllBtn">Mostrar tudo</button>
<div id="resultCount" style="align-self:center;color:var(--muted);font-size:.9rem"></div>
</div>
<div id="topAlert"></div>
<div class="table-wrap"><table><thead><tr><th>Cliente</th><th>Produto</th><th>Resp.</th><th>Prioridade</th><th>Pedido</th><th>Entrada</th><th>Entrega</th><th>Prazo proc.</th><th>Etapa</th><th>Atraso</th><th>Ações</th></tr></thead><tbody id="tableBody"></tbody></table></div>
</div>
<div class="card"><strong>Clientes</strong><div id="clientList" style="display:grid;gap:10px;margin-top:10px"></div></div>
</div>
</section>
<section class="screen" data-panel="kanban">
<div class="hero"><strong>Kanban</strong><p>Controle por etapa com alerta visual de atraso, prioridade e contador de dias.</p></div>
<div class="report-grid" id="stageCounts"></div>
<div id="kanbanBoard" class="kanban"></div>
</section>
<section class="screen" data-panel="gantt">
<div class="hero"><strong>Gantt</strong><p>Acompanhamento dos prazos de entrada, processo e entrega.</p></div>
<div class="card"><div id="gantt" style="overflow:auto"></div></div>
</section>
<section class="screen" data-panel="relatorios">
<div class="hero"><strong>Relatórios</strong><p>Resumo rápido da operação atual.</p></div>
<div class="report-grid" id="reportGrid"></div>
<div class="card"><div id="reportNotes" style="display:grid;gap:8px;margin-top:10px"></div></div>
</section>
</div>
</main>
</div>
<script>
let data=[];let chart;const stages=['A produzir','Atrasado','Em produção','Galvanização','Pintura','Pátio','Concluído','Entregue'];let filters={cliente:'',produto:'',etapa:''};let pendingOnly=false, overdueOnly=false;const today=()=>new Date(new Date().toISOString().slice(0,10)+'T00:00:00');
const d0=s=>s?new Date(s+'T00:00:00'):null;const fmt=s=>s?new Date(s+'T00:00:00').toLocaleDateString('pt-BR'):'';const days=(a,b)=>Math.ceil((d0(a)-d0(b))/86400000);
const getDelay=x=>{const ref=x.entrega||x.previsao; if(!ref||x.etapa==='Concluído'||x.etapa==='Entregue') return 0; return Math.max(0,days(today().toISOString().slice(0,10),ref));};
const save=()=>localStorage.setItem('pcpflowv5',JSON.stringify(data));
function load(){try{data=JSON.parse(localStorage.getItem('pcpflowv5')||'[]')}catch(e){data=[]} render();}
function getRows(){let r=[...data]; if(filters.cliente) r=r.filter(x=>(x.cliente||'').toLowerCase().includes(filters.cliente.toLowerCase())); if(filters.produto) r=r.filter(x=>(x.produto||'').toLowerCase().includes(filters.produto.toLowerCase())); if(filters.etapa) r=r.filter(x=>x.etapa===filters.etapa); if(overdueOnly) r=r.filter(x=>getDelay(x)>0); if(pendingOnly) r=r.filter(x=>getDelay(x)>0||!['Concluído','Entregue'].includes(x.etapa)); return r;}
function labelStage(x){if(getDelay(x)>0&& !['Concluído','Entregue'].includes(x.etapa)) return 'Atrasado'; return x.etapa||'A produzir';}
function render(){save(); const rows=getRows(); const m={total:data.length, atraso:0, producao:0, concluidas:0, entregues:0, emprod:0, aprod:0, gg:0, pt:0, patio:0, atrasoDias:0}; data.forEach(x=>{const a=getDelay(x); if(a>0 && !['Concluído','Entregue'].includes(x.etapa)){m.atraso++; m.atrasoDias+=a} if(x.etapa==='Concluído') m.concluidas++; else if(x.etapa==='Entregue') m.entregues++; else if(x.etapa==='Em produção') m.emprod++; else if(x.etapa==='Galvanização') m.gg++; else if(x.etapa==='Pintura') m.pt++; else if(x.etapa==='Pátio') m.patio++; else m.aprod++;}); document.getElementById('kpiTotal').textContent=m.total; document.getElementById('kpiAtraso').textContent=(m.total?Math.round(m.atraso/m.total*100):0)+'%'; document.getElementById('kpiProducao').textContent=(m.total?Math.round(m.emprod/m.total*100):0)+'%'; document.getElementById('kpiConcluidas').textContent=(m.total?Math.round((m.concluidas+m.entregues)/m.total*100):0)+'%'; document.getElementById('navAtivos').textContent=m.total-(m.concluidas+m.entregues); document.getElementById('navAtrasados').textContent=m.atraso; document.getElementById('navConcluidos').textContent=m.concluidas+m.entregues; document.getElementById('resultCount').textContent=overdueOnly?'Mostrando somente atrasados.':pendingOnly?'Mostrando pendências.':`Mostrando ${rows.length} de ${data.length}`; document.getElementById('topAlert').innerHTML=m.atraso?`<div class="alert">⚠ ${m.atraso} pedido(s) em atraso, totalizando ${m.atrasoDias} dia(s).</div>`:'';
const tbody=document.getElementById('tableBody'); tbody.innerHTML=rows.map((x,i)=>{const a=getDelay(x); const late=a>0&&!['Concluído','Entregue'].includes(x.etapa); const cls=late?'t':x.etapa==='Em produção'?'e':x.etapa==='Galvanização'?'g':x.etapa==='Pintura'?'p':x.etapa==='Pátio'?'pa':x.etapa==='Concluído'?'c':x.etapa==='Entregue'?'d':'a'; return `<tr><td>${x.cliente||''}</td><td>${x.produto||''}</td><td>${x.responsavel||''}</td><td>${x.prioridade||''}</td><td>${x.pedido||''}</td><td>${fmt(x.entrada)}</td><td>${fmt(x.entrega||x.previsao)}</td><td>${fmt(x.prazo)}</td><td><span class="pill ${late?'t':cls}">${late?'Atrasado':(x.etapa||'A produzir')}</span></td><td>${a} d</td><td><button data-act="edit" data-idx="${data.indexOf(x)}">Editar</button> <button data-act="del" data-idx="${data.indexOf(x)}">Excluir</button></td></tr>`}).join('');
const stageCounts=document.getElementById('stageCounts'); stageCounts.innerHTML=stages.map(s=>`<div class="report-card"><span>${s}</span><strong>${rows.filter(x=>labelStage(x)===s).length}</strong></div>`).join('');
const kanban=document.getElementById('kanbanBoard'); kanban.innerHTML=stages.map(s=>{const items=rows.filter(x=>labelStage(x)===s); return `<div class="lane" ondragover="event.preventDefault()" data-stage="${s}"><h4>${s}</h4><small>${items.length} item(ns)</small><div class="lane-list">${items.length?items.map(x=>`<div class="cardjob ${getDelay(x)>0&&!['Concluído','Entregue'].includes(x.etapa)?'overdue':''} ${x.prioridade==='Urgente'?'urgent':''}" draggable="true" data-id="${data.indexOf(x)}"><strong>${x.cliente||''}</strong><span>${x.produto||''}</span><span>Resp.: ${x.responsavel||'-'} • Pri.: ${x.prioridade||'-'}</span><span>Pedido: ${x.pedido||'-'}</span><span>Entrada: ${fmt(x.entrada)} • Entrega: ${fmt(x.entrega||x.previsao)}</span><span>${getDelay(x)} dia(s) atraso</span></div>`).join(''):'<div class="cardjob"><strong>Sem itens</strong><span>Nenhum pedido nesta etapa.</span></div>'}</div></div>`}).join('');
const gantt=document.getElementById('gantt'); gantt.innerHTML=rows.filter(x=>x.entrada&& (x.entrega||x.prazo||x.previsao)).map(x=>`<div style="padding:6px 0;border-bottom:1px solid var(--border)"><strong>${x.cliente||''}</strong> — ${x.etapa||'A produzir'} — Entrega ${fmt(x.entrega||x.previsao)}</div>`).join('')||'<div style="color:var(--muted)">Sem dados para o Gantt.</div>';
const rep=document.getElementById('reportGrid'); rep.innerHTML=[['Pedidos',m.total],['Atrasados',m.atraso],['Dias atraso',m.atrasoDias],['Concluídos',m.concluidas+m.entregues]].map(([a,b])=>`<div class="report-card"><span>${a}</span><strong>${b}</strong></div>`).join(''); document.getElementById('reportNotes').innerHTML=[`Há ${m.total} pedido(s) cadastrados.`, m.atraso?`${m.atraso} pedido(s) em atraso.`:'Nenhum pedido em atraso.', `${m.concluidas+m.entregues} concluído(s)/entregue(s).`].map(t=>`<div>${t}</div>`).join('');
if(chart) chart.destroy(); chart=new Chart(document.getElementById('statusChart'),{type:'bar',data:{labels:['Atraso','Produção','Concluído'],datasets:[{data:[m.atraso,m.emprod,m.concluidas+m.entregues],backgroundColor:['#a13544','#da7101','#437a22']}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}});
}
function add(){const x={cliente:cliente.value.trim(),produto:produto.value.trim(),responsavel:responsavel.value.trim(),prioridade:prioridade.value,pedido:pedido.value.trim(),entrada:entrada.value,entrega:entrega.value,prazo:entrada.value,previsao:entrega.value,etapa:'A produzir'}; if(!x.cliente||!x.produto){alert('Preencha cliente e produto.'); return} data.unshift(x); ['cliente','produto','responsavel','pedido','entrada','entrega'].forEach(id=>document.getElementById(id).value=''); prioridade.value='Baixa'; render();}
function bind(){document.querySelectorAll('[data-screen]').forEach(b=>b.addEventListener('click',()=>{document.querySelectorAll('[data-screen]').forEach(x=>x.classList.remove('active')); b.classList.add('active'); const s=b.dataset.screen; document.querySelectorAll('.screen').forEach(p=>p.classList.toggle('active',p.dataset.panel===s));})); document.getElementById('addRowBtn').onclick=add; document.getElementById('clearDataBtn').onclick=()=>{if(confirm('Apagar todos os dados?')){data=[]; render();}}; document.getElementById('clearFiltersBtn').onclick=()=>{filters={cliente:'',produto:'',etapa:''}; filterCliente.value=''; filterProduto.value=''; filterEtapa.value=''; overdueOnly=false; pendingOnly=false; render();}; document.getElementById('showOverdueBtn').onclick=()=>{overdueOnly=true; pendingOnly=false; render();}; document.getElementById('showPendingBtn').onclick=()=>{pendingOnly=true; overdueOnly=false; render();}; document.getElementById('showAllBtn').onclick=()=>{pendingOnly=false; overdueOnly=false; render();}; filterCliente.oninput=e=>{filters.cliente=e.target.value; render();}; filterProduto.oninput=e=>{filters.produto=e.target.value; render();}; filterEtapa.onchange=e=>{filters.etapa=e.target.value; render();}; exportBtn.onclick=()=>{const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'PCP'); XLSX.writeFile(wb,'PCPFlowExport.xlsx');}; themeBtn.onclick=()=>{document.documentElement.classList.toggle('dark');}; document.getElementById('excelInput').onchange=async e=>{const file=e.target.files[0]; if(!file) return; const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; const json=XLSX.utils.sheet_to_json(ws); const f=(r,keys)=>{const k=Object.keys(r).find(x=>keys.some(s=>x.toLowerCase().includes(s))); return k?r[k]:''}; data=[...json.map(r=>({cliente:f(r,['cliente','empresa']),produto:f(r,['produto','descricao','item']),responsavel:f(r,['responsavel','responsável']),prioridade:f(r,['prioridade'])||'Baixa',pedido:f(r,['pedido','op','numero','número']),entrada:(f(r,['entrada','inicio'])||''),entrega:(f(r,['entrega','previsao','previsão','fim'])||''),prazo:(f(r,['prazo'])||''),previsao:(f(r,['entrega','previsao','previsão','fim'])||''),etapa:f(r,['etapa','status'])||'A produzir'})).filter(x=>x.cliente||x.produto), ...data]; render();}; document.getElementById('tableBody').onclick=e=>{const btn=e.target.closest('button'); if(!btn) return; const idx=Number(btn.dataset.idx); if(btn.dataset.act==='del'){data.splice(idx,1); render();} if(btn.dataset.act==='edit'){const x=data[idx]; cliente.value=x.cliente||''; produto.value=x.produto||''; responsavel.value=x.responsavel||''; prioridade.value=x.prioridade||'Baixa'; pedido.value=x.pedido||''; entrada.value=x.entrada||''; entrega.value=x.entrega||x.previsao||''; data.splice(idx,1); render(); document.querySelector('[data-panel="pedidos"]').scrollIntoView({behavior:'smooth'});}};}
cliente.onkeydown=produto.onkeydown=responsavel.onkeydown=pedido.onkeydown=entrada.onkeydown=entrega.onkeydown=(e)=>{if(e.key==='Enter'){e.preventDefault(); add();}};
bind(); load();
</script>
</body>
</html>
