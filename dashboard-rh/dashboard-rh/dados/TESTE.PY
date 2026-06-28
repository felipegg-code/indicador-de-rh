from pathlib import Path

html = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TechSolutions — Dashboard RH · Filial 58</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
:root{
  --bg:#f4f5f7;--white:#fff;--border:#e3e6ed;--text:#1e2235;--muted:#7b83a0;
  --orange:#f47c20;--orange-l:#fff4ec;--orange-bd:#ffd4ae;
  --blue:#2563b8;--blue-l:#edf3fc;--blue-bd:#bad0f7;
  --green:#1e9e5e;--green-l:#eaf7f1;--green-bd:#a8e4c8;
  --amber:#c97b00;--amber-l:#fef8ec;--amber-bd:#f5d98a;
  --red:#c0392b;--red-l:#fdf1f0;--red-bd:#f5bcb7;
  --purple:#7c3aed;--purple-l:#f5f3ff;--purple-bd:#ddd6fe;
  --font:'Inter',sans-serif;--r:10px;--r-sm:7px;--sh:0 1px 3px rgba(0,0,0,0.07);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:var(--text);font-family:var(--font);font-size:13px;line-height:1.5;}
.topbar{background:var(--orange);padding:14px 28px 12px;}
.topbar-title{font-size:18px;font-weight:700;color:#fff;letter-spacing:.2px;}
.topbar-sub{font-size:11px;color:rgba(255,255,255,.82);margin-top:2px;}
.nav{background:var(--white);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;padding:0 24px;flex-wrap:wrap;gap:8px;}
.tabs{display:flex;}
.tab{background:none;border:none;border-bottom:3px solid transparent;padding:13px 16px;font-size:12px;font-weight:500;color:var(--muted);cursor:pointer;font-family:var(--font);transition:all .15s;white-space:nowrap;}
.tab:hover{color:var(--orange);}
.tab.active{color:var(--orange);border-bottom-color:var(--orange);}
.nav-chips{display:flex;gap:6px;padding:8px 0;}
.chip{background:var(--bg);border:1px solid var(--border);border-radius:20px;padding:4px 11px;font-size:10px;color:var(--muted);}
.wrap{max-width:1340px;margin:0 auto;padding:20px 20px 48px;}
.panel{display:none;}
.panel.active{display:block;animation:up .22s ease both;}
@keyframes up{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
.kpi-strip{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;}
.kpi{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:15px 18px;box-shadow:var(--sh);border-top:3px solid var(--kc,var(--border));}
.kpi-lbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:var(--kc,var(--muted));margin-bottom:7px;}
.kpi-val{font-size:28px;font-weight:700;color:var(--kc,var(--text));line-height:1;}
.kpi-sub{font-size:10px;color:var(--muted);margin-top:5px;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;}
.g3{display:grid;grid-template-columns:2fr 1fr;gap:14px;margin-bottom:14px;}
.card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;box-shadow:var(--sh);margin-bottom:14px;}
.card-hdr{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px;gap:8px;}
.card-title{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;display:flex;align-items:center;gap:7px;}
.card-title::before{content:'';width:3px;height:13px;background:var(--orange);border-radius:2px;flex-shrink:0;}
.card-tag{font-size:9px;font-weight:600;padding:3px 9px;border-radius:10px;background:var(--amber-l);color:var(--amber);border:1px solid var(--amber-bd);white-space:nowrap;}
.chart-box{position:relative;width:100%;}
.tbl-wrap{overflow-x:auto;}
table{width:100%;border-collapse:collapse;}
thead th{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;padding:7px 10px;border-bottom:2px solid var(--border);text-align:left;white-space:nowrap;background:var(--bg);}
tbody td{padding:7px 10px;border-bottom:1px solid var(--border);font-size:12px;color:var(--text);vertical-align:middle;}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover td{background:#f9fafc;}
.mono{font-family:'Courier New',monospace;font-size:11px;}
.pill{display:inline-block;font-size:9px;font-weight:600;padding:2px 8px;border-radius:10px;white-space:nowrap;}
.p-red{background:var(--red-l);color:var(--red);border:1px solid var(--red-bd);}
.p-amber{background:var(--amber-l);color:var(--amber);border:1px solid var(--amber-bd);}
.p-blue{background:var(--blue-l);color:var(--blue);border:1px solid var(--blue-bd);}
.p-green{background:var(--green-l);color:var(--green);border:1px solid var(--green-bd);}
.p-orange{background:var(--orange-l);color:var(--orange);border:1px solid var(--orange-bd);}
.p-gray{background:var(--bg);color:var(--muted);border:1px solid var(--border);}
.p-purple{background:var(--purple-l);color:var(--purple);border:1px solid var(--purple-bd);}
.legend{display:flex;flex-wrap:wrap;gap:12px;margin-top:10px;}
.leg{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--muted);}
.leg-sq{width:9px;height:9px;border-radius:2px;flex-shrink:0;}
.donut-wrap{position:relative;display:flex;align-items:center;justify-content:center;}
.donut-center{position:absolute;text-align:center;pointer-events:none;}
.donut-num{font-size:22px;font-weight:700;}
.donut-lbl{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;}
.filter-row{display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap;}
.filter-lbl{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;}
.filter-sel{background:var(--white);border:1px solid var(--border);border-radius:var(--r-sm);padding:5px 11px;font-size:11px;color:var(--text);cursor:pointer;font-family:var(--font);outline:none;transition:border-color .15s;}
.filter-sel:focus,.filter-sel:hover{border-color:var(--orange);}
.footer{border-top:1px solid var(--border);padding-top:14px;margin-top:4px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;}
.footer-txt{font-size:10px;color:var(--muted);}
.sec-div{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin:20px 0 12px;padding-bottom:8px;border-bottom:1px solid var(--border);}
.pbar{display:flex;align-items:center;gap:6px;}
.pbar-track{flex:1;height:5px;background:var(--border);border-radius:3px;overflow:hidden;}
.pbar-fill{height:100%;border-radius:3px;}
.pbar-val{font-size:10px;font-weight:600;min-width:34px;text-align:right;}
.alert-banner{background:var(--red-l);border:1px solid var(--red-bd);border-radius:var(--r);padding:13px 18px;margin-bottom:16px;display:flex;align-items:flex-start;gap:12px;}
.alert-icon{font-size:20px;margin-top:1px;}
.alert-title{font-size:12px;font-weight:700;color:var(--red);}
.alert-sub{font-size:11px;color:var(--muted);margin-top:3px;}
.grid-outer{overflow-x:auto;-webkit-overflow-scrolling:touch;border-radius:6px;border:1px solid var(--border);}
.esc-tbl{border-collapse:collapse;font-size:10px;}
.esc-tbl th,.esc-tbl td{padding:0;border:none;}
.col-n{position:sticky;left:0;z-index:3;background:var(--white);width:155px;min-width:155px;}
.col-s{position:sticky;left:155px;z-index:3;background:var(--white);width:95px;min-width:95px;border-right:2px solid var(--border);}
.esc-tbl thead th{background:var(--bg);border-bottom:2px solid var(--border);}
.esc-tbl thead .col-n,.esc-tbl thead .col-s{background:var(--bg);z-index:4;}
.th-n{padding:8px 10px;font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;text-align:left;}
.th-day{width:26px;min-width:26px;text-align:center;vertical-align:bottom;padding:4px 1px 3px;}
.th-day .dn{font-size:9px;font-weight:700;display:block;line-height:1.2;}
.th-day .dw{font-size:8px;font-weight:500;display:block;opacity:.7;line-height:1.2;}
.th-day.dom{background:#fff0f0;}
.th-day.dom .dn,.th-day.dom .dw{color:var(--red);}
.sec-hdr-row td{background:var(--bg);padding:5px 10px;font-size:9px;font-weight:700;color:var(--orange);text-transform:uppercase;letter-spacing:.08em;border-top:2px solid var(--border);}
.sec-hdr-row .col-n{background:var(--bg);}
.sec-hdr-row .col-s{background:var(--bg);}
.emp-row:hover .col-n,.emp-row:hover .col-s,.emp-row:hover td{background:#f5f7ff !important;}
.td-n{padding:5px 10px;font-size:10px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-bottom:1px solid var(--border);}
.td-s{padding:5px 8px;font-size:9px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-bottom:1px solid var(--border);}
.td-day{width:26px;min-width:26px;text-align:center;padding:3px 1px;border-bottom:1px solid var(--border);vertical-align:middle;}
.td-day.dom{background:#fff8f8;}
.dc{width:22px;height:20px;border-radius:3px;display:inline-flex;align-items:center;justify-content:center;font-size:7px;font-weight:800;letter-spacing:0;}
.dc-T{background:#dbeafe;color:#1e40af;}
.dc-F{background:#dcfce7;color:#15803d;}
.dc-DT{background:#fee2e2;color:#991b1b;}
.dc-DS{background:#bbf7d0;color:#14532d;}
.sum-row td{background:var(--orange-l);border-top:2px solid var(--orange-bd);border-bottom:2px solid var(--orange-bd);font-weight:700;}
.sum-row .col-n{background:var(--orange-l);}
.sum-row .col-s{background:var(--orange-l);}
.vbadge{display:inline-block;font-size:7px;font-weight:700;padding:1px 4px;border-radius:6px;margin-left:5px;vertical-align:middle;}
.vb-r{background:var(--red-l);color:var(--red);}
.vb-a{background:var(--amber-l);color:var(--amber);}
.esc-legend{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:12px;}
.esc-leg{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--muted);}
@media(max-width:900px){
  .kpi-strip{grid-template-columns:repeat(2,1fr);}
  .g2,.g3{grid-template-columns:1fr;}
}
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-title">TECHSOLUTIONS — DASHBOARD RH</div>
  <div class="topbar-sub">Filial 58 · Operações de Varejo · Gestão de Escalas, BH e Jornada · Abr/2026</div>
</div>

<div class="nav">
  <div class="tabs">
    <button class="tab active" onclick="show('esc')">📅 Escalas 6×1</button>
    <button class="tab" onclick="show('bh')">Banco de Horas</button>
    <button class="tab" onclick="show('intra')">Intrajornada</button>
    <button class="tab" onclick="show('inter')">Interjornada</button>
  </div>
  <div class="nav-chips">
    <div class="chip">101 colaboradores (BH)</div>
    <div class="chip">18 na escala 6×1</div>
    <div class="chip">Validação CLT automática</div>
    <div class="chip">Abr/2026</div>
  </div>
</div>

<div class="wrap">
  <div id="panel-esc" class="panel active">
    <div class="kpi-strip">
      <div class="kpi" style="--kc:var(--blue)">
        <div class="kpi-lbl">Total escalados</div>
        <div class="kpi-val">18</div>
        <div class="kpi-sub">colaboradores na escala Abr/2026</div>
      </div>
      <div class="kpi" style="--kc:var(--orange)">
        <div class="kpi-lbl">Domingos trabalhados</div>
        <div class="kpi-val" id="kpi-dt">—</div>
        <div class="kpi-sub">ocorrências no mês (4 domingos)</div>
      </div>
      <div class="kpi" style="--kc:var(--red)">
        <div class="kpi-lbl">Sem descanso dominical</div>
        <div class="kpi-val" id="kpi-sd">—</div>
        <div class="kpi-sub">colaboradores · 0 dom. de folga</div>
      </div>
      <div class="kpi" style="--kc:var(--amber)">
        <div class="kpi-lbl">Alertas CLT</div>
        <div class="kpi-val" id="kpi-al">—</div>
        <div class="kpi-sub">jornada contínua &gt; 6 dias</div>
      </div>
    </div>

    <div class="alert-banner" id="alertBanner" style="display:none">
      <div class="alert-icon">⚠️</div>
      <div>
        <div class="alert-title" id="alertTitle"></div>
        <div class="alert-sub">Ajuste as escalas antes do fechamento do período para evitar passivo trabalhista.</div>
      </div>
    </div>

    <div class="card">
      <div class="card-hdr">
        <div class="card-title">Escala mensal — Abril 2026</div>
        <div class="card-tag" id="escTag">18 colaboradores</div>
      </div>
      <div class="filter-row">
        <span class="filter-lbl">Seção:</span>
        <select class="filter-sel" id="fSec" onchange="renderEsc()"><option value="">Todas</option></select>
        <span class="filter-lbl" style="margin-left:8px;">Exibir:</span>
        <select class="filter-sel" id="fViol" onchange="renderEsc()">
          <option value="">Todos</option>
          <option value="viol">Apenas com violação</option>
          <option value="dom">Sem descanso dominical</option>
          <option value="ok">Conformes CLT</option>
        </select>
      </div>
      <div class="esc-legend">
        <div class="esc-leg"><div class="dc dc-T">T</div> Trabalhado</div>
        <div class="esc-leg"><div class="dc dc-F">F</div> Folga (dia útil)</div>
        <div class="esc-leg"><div class="dc dc-DT">D</div> Domingo trabalhado</div>
        <div class="esc-leg"><div class="dc dc-DS">D</div> Domingo de folga ✓</div>
      </div>
      <div class="grid-outer">
        <table class="esc-tbl" id="escTbl">
          <thead id="escHead"></thead>
          <tbody id="escBody"></tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="card-hdr">
        <div class="card-title">Registro de violações CLT — Abril 2026</div>
        <div class="card-tag" id="violTag" style="background:var(--red-l);color:var(--red);border-color:var(--red-bd);">—</div>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>#</th><th>Colaborador</th><th>Seção</th><th>Tipo de violação</th><th>Detalhes</th><th>Severidade</th></tr></thead>
          <tbody id="violBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div id="panel-bh" class="panel">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">Banco de Horas</div>
        <div class="card-tag">Exemplo</div>
      </div>
      <p style="font-size:12px;color:#7b83a0;">Estrutura pronta para expandir com seus dados reais.</p>
    </div>
  </div>

  <div id="panel-intra" class="panel">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">Intrajornada</div>
        <div class="card-tag">Exemplo</div>
      </div>
      <p style="font-size:12px;color:#7b83a0;">Estrutura pronta para expandir com seus dados reais.</p>
    </div>
  </div>

  <div id="panel-inter" class="panel">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">Interjornada</div>
        <div class="card-tag">Exemplo</div>
      </div>
      <p style="font-size:12px;color:#7b83a0;">Estrutura pronta para expandir com seus dados reais.</p>
    </div>
  </div>

  <div class="footer">
    <div class="footer-txt">Dashboard RH · Filial 58</div>
    <div class="footer-txt">Gerado automaticamente por Python</div>
  </div>
</div>

<script>
function show(name){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('panel-' + name).classList.add('active');
  event.target.classList.add('active');
}

const SUNS = new Set([4,11,18,25]);
const WD = ['Q','Q','X','B','D','S','T','Q','Q','X','B','D','S','T','Q','Q','X','B','D','S','T','Q','Q','X','B','D','S','T','Q','Q'];

const EMPS = [
  {n:"ANDRESA SANTANA DANTAS", s:"FRENTE DE LOJA", off:0},
  {n:"WENDSSON VICTOR DA SILVA LIMA", s:"FRENTE DE LOJA", off:1},
  {n:"VILMARA CARVALHO ANDRADE", s:"FRENTE DE LOJA", off:2},
  {n:"HIAGO SANTOS SOUZA", s:"FRENTE DE LOJA", custom:[1,1,1,1,1,1,1,1,0,1,1,1,1,1,1,0,1,1,1,1,1,1,0,1,1,1,1,1,1,0]},
  {n:"WILLIAM DA SILVA RODRIGUES", s:"AÇOUGUE", off:3},
  {n:"EGINALDO SANTOS DE ALMEIDA", s:"AÇOUGUE", off:4},
  {n:"MOISES DE JESUS FONSECA", s:"PADARIA", off:5},
  {n:"CLEITON DE SOUZA SILVA", s:"PREV. PERDAS", off:0},
  {n:"MAXWELL MESSIAS DOS SANTOS", s:"PREV. PERDAS", off:3},
  {n:"MAXWELL DA SILVA COSTA", s:"HORTIFRUTI", off:4},
  {n:"SERGIO REZENDE DE SOUZA", s:"HORTIFRUTI", off:1},
  {n:"JIMMY WESLEY DOS ANJOS", s:"SERV. GERAIS", off:5},
  {n:"GILTON ALVES LINO", s:"SERV. GERAIS", off:2},
  {n:"ADEMILSON DOS SANTOS", s:"CONGELADOS", off:6},
  {n:"ADRIANO FELIX DOS SANTOS", s:"MERCEARIA SECA", off:3},
  {n:"BIANCA DE ANDRADE OLIVEIRA", s:"MERCEARIA SECA", off:6},
  {n:"MARIA VITORIA DE JESUS", s:"CONFEITARIA", off:4},
  {n:"VIVIANE SANTOS DE JESUS", s:"ROTISSERIA", off:1}
];

let allData = [];

function getPat(emp){
  if(emp.custom) return emp.custom;
  return Array.from({length:30}, (_, d) => ((d + emp.off) % 7 === 6 ? 0 : 1));
}

function getViol(emp, pat){
  const viols = [];
  const sunWork = [4,11,18,25].filter(d => pat[d] === 1).length;

  if(sunWork === 4){
    viols.push({
      type:'SEM_DOM',
      label:'Sem descanso dominical',
      detail:'4/4 domingos trabalhados — nenhum domingo de folga no mês',
      sev:'ALERTA'
    });
  }

  let mx = 0, cur = 0;
  for(let d = 0; d < 30; d++){
    if(pat[d] === 1){
      cur++;
      mx = Math.max(mx, cur);
    } else {
      cur = 0;
    }
  }

  if(mx > 6){
    viols.push({
      type:'JORNADA',
      label:'Jornada contínua',
      detail:`${mx} dias consecutivos sem folga (limite CLT: 6 dias)`,
      sev:'GRAVE'
    });
  }

  return viols;
}

function computeData(){
  allData = EMPS.map(emp => {
    const pat = getPat(emp);
    const viols = getViol(emp, pat);
    const sunWorked = [4,11,18,25].filter(d => pat[d] === 1).length;
    return {...emp, pat, viols, sunWorked};
  });
}

function initEscFilters(){
  const secs = [...new Set(EMPS.map(e => e.s))].sort();
  const sel = document.getElementById('fSec');
  secs.forEach(s => {
    const o = document.createElement('option');
    o.value = s;
    o.textContent = s;
    sel.appendChild(o);
  });
}

function buildEscKPIs(){
  const totalDT = allData.reduce((s,e) => s + e.sunWorked, 0);
  const semDom = allData.filter(e => e.sunWorked === 4).length;
  const alertas = allData.filter(e => e.viols.some(v => v.type === 'JORNADA')).length;

  document.getElementById('kpi-dt').textContent = totalDT;
  document.getElementById('kpi-sd').textContent = semDom;
  document.getElementById('kpi-al').textContent = alertas;

  const msgs = [];
  if(semDom > 0) msgs.push(`${semDom} colaborador(es) sem nenhum domingo de folga em abril`);
  if(alertas > 0) msgs.push(`${alertas} violação(ões) de jornada contínua acima de 6 dias`);

  if(msgs.length > 0){
    document.getElementById('alertTitle').textContent = msgs.join(' · ');
    document.getElementById('alertBanner').style.display = 'flex';
  }
}

function cellType(d, worked){
  const sun = SUNS.has(d);
  if(worked && sun) return 'DT';
  if(!worked && sun) return 'DS';
  if(worked) return 'T';
  return 'F';
}

function renderEsc(){
  const fSec = document.getElementById('fSec').value;
  const fViol = document.getElementById('fViol').value;

  let data = [...allData];
  if(fSec) data = data.filter(e => e.s === fSec);
  if(fViol === 'viol') data = data.filter(e => e.viols.length > 0);
  if(fViol === 'dom') data = data.filter(e => e.sunWorked === 4);
  if(fViol === 'ok') data = data.filter(e => e.viols.length === 0);

  document.getElementById('escTag').textContent = `${data.length} colaboradores`;

  const head = document.getElementById('escHead');
  let hRow = `<tr><th class="col-n th-n">Colaborador</th><th class="col-s th-n">Seção</th>`;
  for(let d = 0; d < 30; d++){
    const sun = SUNS.has(d);
    hRow += `<th class="th-day${sun ? ' dom' : ''}"><span class="dn">${d+1}</span><span class="dw">${WD[d]}</span></th>`;
  }
  head.innerHTML = hRow + '</tr>';

  const body = document.getElementById('escBody');
  let rows = '';
  const sections = [...new Set(data.map(e => e.s))];
  const domCount = Array(30).fill(0);

  sections.forEach(sec => {
    const grp = data.filter(e => e.s === sec);
    rows += `<tr class="sec-hdr-row"><td class="col-n" colspan="2">${sec}</td>${Array(30).fill('<td></td>').join('')}</tr>`;

    grp.forEach(emp => {
      const hasGrave = emp.viols.some(v => v.sev === 'GRAVE');
      const hasAlerta = emp.viols.some(v => v.sev === 'ALERTA');
      let badge = '';
      if(hasGrave) badge = '<span class="vbadge vb-r">GRAVE</span>';
      else if(hasAlerta) badge = '<span class="vbadge vb-a">ALERTA</span>';

      rows += `<tr class="emp-row"><td class="col-n td-n" title="${emp.n}">${emp.n}${badge}</td><td class="col-s td-s">${emp.s}</td>`;
      for(let d = 0; d < 30; d++){
        const worked = emp.pat[d] === 1;
        const ct = cellType(d, worked);
        const sun = SUNS.has(d);
        if(worked && sun) domCount[d]++;
        rows += `<td class="td-day${sun ? ' dom' : ''}"><div class="dc dc-${ct}">${ct === 'T' || ct === 'DT' ? 'T' : 'F'}</div></td>`;
      }
      rows += '</tr>';
    });
  });

  if(data.length > 0){
    rows += `<tr class="sum-row"><td class="col-n td-n">TOTAL DOMINGOS TRABALHADOS</td><td class="col-s td-s">Filial 58</td>`;
    for(let d = 0; d < 30; d++){
      const sun = SUNS.has(d);
      if(sun){
        const cnt = domCount[d];
        rows += `<td class="td-day dom" style="text-align:center;font-size:9px;font-weight:700;color:${cnt > 0 ? 'var(--red)' : 'var(--green)'}">${cnt > 0 ? cnt : '✓'}</td>`;
      } else {
        rows += `<td class="td-day"></td>`;
      }
    }
    rows += '</tr>';
  }

  body.innerHTML = rows;

  const violBody = document.getElementById('violBody');
  let violRows = '', violCount = 0;

  data.forEach(emp => {
    emp.viols.forEach(v => {
      violCount++;
      const pill = v.sev === 'GRAVE' ? 'p-red' : 'p-amber';
      violRows += `<tr>
        <td class="mono">${violCount}</td>
        <td style="font-weight:500;font-size:11px;">${emp.n}</td>
        <td style="color:var(--muted);font-size:11px;">${emp.s}</td>
        <td><span class="pill p-blue">${v.label}</span></td>
        <td style="font-size:11px;color:var(--muted);">${v.detail}</td>
        <td><span class="pill ${pill}">${v.sev}</span></td>
      </tr>`;
    });
  });

  if(violCount === 0){
    violRows = `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:16px;">Nenhuma violação encontrada na seleção atual ✓</td></tr>`;
  }

  violBody.innerHTML = violRows;
  document.getElementById('violTag').textContent = `${violCount} violação${violCount !== 1 ? 'ões' : ''}`;
}

computeData();
initEscFilters();
buildEscKPIs();
renderEsc();
</script>

</body>
</html>
"""

arquivo_saida = Path("index.html")
arquivo_saida.write_text(html, encoding="utf-8")

print(f"Arquivo HTML gerado com sucesso em: {arquivo_saida.resolve()}")