!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Panel de Despachos</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --bg:#0E1316;
    --panel:#161D22;
    --panel-alt:#1C242A;
    --line:#2B353C;
    --line-soft:#212A30;
    --text:#EAF1F4;
    --text-dim:#8CA0AC;
    --text-faint:#5E717B;
    --claro:#FF6B4A;
    --vtr:#2DD4BF;
    --amber:#F5A623;
    --ok:#4ADE80;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background:
      radial-gradient(circle at 15% 0%, #16202480 0%, transparent 45%),
      var(--bg);
    color:var(--text);
    font-family:'Inter',sans-serif;
    -webkit-font-smoothing:antialiased;
  }
  .mono{font-family:'IBM Plex Mono',monospace;}
  .wrap{max-width:1280px;margin:0 auto;padding:28px 28px 64px;}

  header{
    display:flex;justify-content:space-between;align-items:flex-end;
    border-bottom:1px solid var(--line);
    padding-bottom:20px;margin-bottom:28px;
    flex-wrap:wrap;gap:16px;
  }
  .brand-eyebrow{
    font-family:'IBM Plex Mono',monospace;font-size:11px;letter-spacing:.18em;
    color:var(--text-faint);text-transform:uppercase;margin-bottom:6px;
  }
  h1{
    font-family:'Barlow Condensed',sans-serif;font-weight:700;
    font-size:40px;line-height:1;margin:0;letter-spacing:.01em;
    text-transform:uppercase;
  }
  h1 span{color:var(--vtr);}
  .header-right{text-align:right;}
  .week-pill{
    display:inline-flex;align-items:center;gap:8px;
    background:var(--panel-alt);border:1px solid var(--line);
    border-radius:100px;padding:7px 16px;font-family:'IBM Plex Mono',monospace;
    font-size:12px;color:var(--text-dim);
  }
  .week-pill .dot{width:7px;height:7px;border-radius:50%;background:var(--amber);
    box-shadow:0 0 0 4px rgba(245,166,35,.15);animation:pulse 2s infinite;}
  @keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
  .header-right .date{font-size:11px;color:var(--text-faint);margin-top:8px;font-family:'IBM Plex Mono',monospace;}

  .section-label{display:flex;align-items:center;gap:12px;margin:44px 0 16px;}
  .section-label .num{
    font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--text-faint);
    border:1px solid var(--line);border-radius:4px;padding:2px 7px;
  }
  .section-label h2{
    font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:600;
    text-transform:uppercase;letter-spacing:.03em;margin:0;color:var(--text);
  }
  .section-label .rule{flex:1;height:1px;background:var(--line);}

  .loader-panel{
    background:var(--panel);border:1px solid var(--line);border-radius:10px;
    padding:20px 22px;display:flex;align-items:center;gap:22px;flex-wrap:wrap;
    transition:border-color .15s, background .15s;
  }
  .loader-panel.drag-over{border-color:var(--vtr);background:var(--panel-alt);}
  .loader-icon{
    width:46px;height:46px;border-radius:8px;flex-shrink:0;
    background:var(--panel-alt);border:1px dashed var(--line);
    display:flex;align-items:center;justify-content:center;
    font-family:'IBM Plex Mono',monospace;font-size:18px;color:var(--vtr);
  }
  .loader-text{flex:1;min-width:220px;}
  .loader-text .title{font-family:'Barlow Condensed',sans-serif;font-size:17px;font-weight:600;text-transform:uppercase;letter-spacing:.02em;}
  .loader-text .sub{font-size:12px;color:var(--text-dim);margin-top:3px;}
  #sourceStatus{font-family:'IBM Plex Mono',monospace;font-size:11px;margin-top:6px;}
  #sourceStatus.ok{color:var(--ok);}
  #sourceStatus.err{color:#FF6B6B;}
  #sourceStatus.demo{color:var(--text-faint);}
  .loader-actions{display:flex;gap:10px;flex-shrink:0;}
  .btn{
    font-family:'IBM Plex Mono',monospace;font-size:11.5px;letter-spacing:.03em;
    text-transform:uppercase;padding:10px 18px;border-radius:7px;border:1px solid var(--line);
    background:var(--panel-alt);color:var(--text);cursor:pointer;transition:.15s;
  }
  .btn:hover{border-color:var(--vtr);color:var(--vtr);}
  .btn.primary{background:var(--vtr);color:#08201C;border-color:var(--vtr);font-weight:600;}
  .btn.primary:hover{filter:brightness(1.08);color:#08201C;}
  .btn.ghost{background:transparent;}
  #fileInput{display:none;}

  .kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:18px;}
  .ticket{position:relative;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:20px 20px 18px;overflow:hidden;}
  .ticket::before{content:'';position:absolute;left:0;top:0;bottom:0;width:5px;}
  .ticket.t-total::before{background:var(--vtr);}
  .ticket.t-avance::before{background:var(--amber);}
  .ticket.t-claro::before{background:var(--claro);}
  .ticket.t-vtr::before{background:var(--vtr);}
  .ticket .barcode{position:absolute;right:14px;top:16px;display:flex;gap:2px;opacity:.35;}
  .ticket .barcode span{width:2px;background:var(--text-dim);}
  .ticket .label{font-family:'IBM Plex Mono',monospace;font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--text-faint);margin-bottom:10px;}
  .ticket .value{font-family:'Barlow Condensed',sans-serif;font-weight:700;font-size:38px;line-height:1;color:var(--text);}
  .ticket .value small{font-size:16px;color:var(--text-dim);font-weight:500;}
  .ticket .sub{margin-top:8px;font-size:12px;color:var(--text-dim);}
  .ticket .sub b{color:var(--text);}

  .grid-2{display:grid;grid-template-columns:1.6fr 1fr;gap:20px;}
  .grid-2b{display:grid;grid-template-columns:1fr 1.3fr;gap:20px;}
  .panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:22px;}
  .panel h3{font-family:'Barlow Condensed',sans-serif;font-size:18px;text-transform:uppercase;letter-spacing:.02em;margin:0 0 4px;font-weight:600;}
  .panel .desc{font-size:12px;color:var(--text-dim);margin:0 0 18px;}
  .legend{display:flex;gap:18px;margin-bottom:14px;flex-wrap:wrap;}
  .legend .item{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--text-dim);font-family:'IBM Plex Mono',monospace;}
  .legend .sw{width:10px;height:10px;border-radius:2px;}

  .radial-wrap{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;}
  .radial{position:relative;width:172px;height:172px;}
  .radial svg{transform:rotate(-90deg);}
  .radial-center{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;}
  .radial-center .pct{font-family:'Barlow Condensed',sans-serif;font-size:42px;font-weight:700;}
  .radial-center .lbl{font-size:10px;color:var(--text-faint);font-family:'IBM Plex Mono',monospace;letter-spacing:.08em;text-transform:uppercase;}
  .radial-foot{margin-top:16px;text-align:center;font-size:12px;color:var(--text-dim);}
  .radial-foot b{color:var(--text);font-family:'IBM Plex Mono',monospace;}

  table{width:100%;border-collapse:collapse;font-size:12.5px;}
  th{text-align:left;font-family:'IBM Plex Mono',monospace;font-weight:500;font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--text-faint);border-bottom:1px solid var(--line);padding:0 10px 10px;}
  td{padding:11px 10px;border-bottom:1px solid var(--line-soft);vertical-align:middle;}
  tr:last-child td{border-bottom:none;}
  .neg-tag{display:inline-flex;align-items:center;gap:6px;font-family:'IBM Plex Mono',monospace;font-size:10.5px;font-weight:600;padding:3px 8px;border-radius:4px;letter-spacing:.03em;}
  .neg-tag.claro{background:rgba(255,107,74,.12);color:var(--claro);}
  .neg-tag.vtr{background:rgba(45,212,191,.12);color:var(--vtr);}
  .neg-tag.other{background:rgba(140,160,172,.12);color:var(--text-dim);}
  .prod-name{color:var(--text);font-weight:500;}
  .bar-cell{min-width:150px;}
  .bar-track{height:7px;background:var(--line-soft);border-radius:4px;overflow:hidden;position:relative;}
  .bar-fill{height:100%;border-radius:4px;}
  .num-r{text-align:right;font-family:'IBM Plex Mono',monospace;color:var(--text-dim);}
  .pend-pos{color:var(--amber);font-family:'IBM Plex Mono',monospace;font-weight:600;}
  .pend-neg{color:var(--ok);font-family:'IBM Plex Mono',monospace;font-weight:600;}

  .guia-list{display:flex;flex-direction:column;gap:9px;max-height:330px;overflow-y:auto;padding-right:4px;}
  .guia-list::-webkit-scrollbar{width:5px;}
  .guia-list::-webkit-scrollbar-thumb{background:var(--line);border-radius:3px;}
  .guia{display:flex;align-items:center;gap:12px;background:var(--panel-alt);border:1px solid var(--line);border-radius:7px;padding:10px 12px;position:relative;}
  .guia::after{content:'';position:absolute;left:58px;top:6px;bottom:6px;width:1px;background:repeating-linear-gradient(var(--line) 0 4px, transparent 4px 8px);}
  .guia .gid{font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--text-dim);width:46px;flex-shrink:0;}
  .guia .gfam{flex:1;font-size:12.5px;}
  .guia .gfam .neg{font-size:10px;color:var(--text-faint);display:block;font-family:'IBM Plex Mono',monospace;letter-spacing:.05em;}
  .guia .gunits{font-family:'IBM Plex Mono',monospace;font-size:14px;font-weight:600;}
  .guia .gstate{font-family:'IBM Plex Mono',monospace;font-size:9.5px;padding:3px 7px;border-radius:4px;text-transform:uppercase;letter-spacing:.04em;flex-shrink:0;}
  .gstate.val{background:rgba(74,222,128,.14);color:var(--ok);}
  .gstate.nov{background:rgba(245,166,35,.14);color:var(--amber);}
  .empty-msg{color:var(--text-faint);font-size:12.5px;text-align:center;padding:30px 0;font-family:'IBM Plex Mono',monospace;}

  .sched-table{width:100%;border-collapse:collapse;font-size:11.5px;}
  .sched-table th{background:var(--panel-alt);text-align:center;padding:8px 4px;font-family:'IBM Plex Mono',monospace;font-size:9.5px;color:var(--text-dim);text-transform:uppercase;border-bottom:1px solid var(--line);}
  .sched-table td{text-align:center;padding:9px 4px;border-bottom:1px solid var(--line-soft);color:var(--text-dim);}
  .sched-table td.wk{font-family:'Barlow Condensed',sans-serif;font-weight:700;font-size:16px;color:var(--vtr);background:var(--panel-alt);}
  .sched-table .slot{display:inline-block;font-family:'IBM Plex Mono',monospace;font-size:10.5px;background:rgba(245,166,35,.12);color:var(--amber);border-radius:4px;padding:3px 6px;}
  .sched-table .date{display:block;font-size:9.5px;color:var(--text-faint);margin-bottom:3px;}

  footer{margin-top:50px;padding-top:20px;border-top:1px solid var(--line);display:flex;justify-content:space-between;color:var(--text-faint);font-size:11px;font-family:'IBM Plex Mono',monospace;flex-wrap:wrap;gap:10px;}

  @media(max-width:900px){
    .kpi-row{grid-template-columns:repeat(2,1fr);}
    .grid-2, .grid-2b{grid-template-columns:1fr;}
  }
</style>
</head>
<body>
<div class="wrap">

  <header>
    <div>
      <div class="brand-eyebrow">Centro J011 · Almacén A009 · Control de despachos</div>
      <h1>Panel de <span>Despachos</span></h1>
    </div>
    <div class="header-right">
      <span class="week-pill"><span class="dot"></span> Semana activa <span id="hdrWeek">—</span></span>
      <div class="date" id="hdrDate">Sin datos cargados</div>
    </div>
  </header>

  <div class="loader-panel" id="loaderPanel">
    <div class="loader-icon">⇪</div>
    <div class="loader-text">
      <div class="title">Cargar datos</div>
      <div class="sub">Arrastra tu archivo <b>despachos.xlsx</b> aquí, o selecciónalo. Debe incluir las hojas <b>DESPACHOS</b> y <b>PLANIFICACION</b>.</div>
      <div id="sourceStatus" class="demo">Mostrando datos de ejemplo (WK28, julio 2026)</div>
    </div>
    <div class="loader-actions">
      <button class="btn primary" onclick="document.getElementById('fileInput').click()">Cargar archivo</button>
      <button class="btn ghost" id="resetBtn">Restaurar ejemplo</button>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.csv">
    </div>
  </div>

  <div class="kpi-row" style="margin-top:24px;">
    <div class="ticket t-total">
      <div class="barcode"><span style="height:14px"></span><span style="height:20px"></span><span style="height:10px"></span><span style="height:18px"></span><span style="height:12px"></span></div>
      <div class="label">Plan mensual</div>
      <div class="value" id="kpiMonthTotal">—</div>
      <div class="sub" id="kpiMonthSub">—</div>
    </div>
    <div class="ticket t-avance">
      <div class="barcode"><span style="height:18px"></span><span style="height:10px"></span><span style="height:16px"></span><span style="height:20px"></span><span style="height:8px"></span></div>
      <div class="label" id="kpiAvanceLabel">Avance semana vs meta</div>
      <div class="value" id="kpiAvancePct">—</div>
      <div class="sub" id="kpiAvanceSub">—</div>
    </div>
    <div class="ticket t-claro">
      <div class="barcode"><span style="height:12px"></span><span style="height:20px"></span><span style="height:14px"></span><span style="height:8px"></span><span style="height:18px"></span></div>
      <div class="label">Negocio Claro</div>
      <div class="value" id="kpiClaro">—</div>
      <div class="sub" id="kpiClaroSub">—</div>
    </div>
    <div class="ticket t-vtr">
      <div class="barcode"><span style="height:20px"></span><span style="height:8px"></span><span style="height:16px"></span><span style="height:12px"></span><span style="height:20px"></span></div>
      <div class="label">Negocio VTR</div>
      <div class="value" id="kpiVtr">—</div>
      <div class="sub" id="kpiVtrSub">—</div>
    </div>
  </div>

  <div class="section-label"><span class="num">01</span><h2>Plan semanal de despachos</h2><span class="rule"></span></div>
  <div class="grid-2">
    <div class="panel">
      <h3>Unidades planificadas por semana</h3>
      <p class="desc" id="weeklyDesc">Distribución por negocio sobre las semanas de facturación del mes</p>
      <div class="legend" id="weeklyLegend"></div>
      <canvas id="weeklyChart" height="230"></canvas>
    </div>
    <div class="panel">
      <h3>Avance semana en curso</h3>
      <p class="desc" id="radialDesc">Despachos reales registrados vs. meta planificada</p>
      <div class="radial-wrap">
        <div class="radial">
          <svg width="172" height="172">
            <circle cx="86" cy="86" r="74" stroke="var(--line-soft)" stroke-width="14" fill="none"/>
            <circle id="radialArc" cx="86" cy="86" r="74" stroke="var(--amber)" stroke-width="14" fill="none"
              stroke-dasharray="464.9" stroke-dashoffset="464.9" stroke-linecap="round"/>
          </svg>
          <div class="radial-center">
            <div class="pct" id="radialPct">—</div>
            <div class="lbl">Avance</div>
          </div>
        </div>
        <div class="radial-foot" id="radialFoot">—</div>
      </div>
    </div>
  </div>

  <div class="section-label"><span class="num">02</span><h2>Estado de preparación por producto</h2><span class="rule"></span></div>
  <div class="panel">
    <p class="desc">Meta mensual por SKU vs. unidades preparadas (diagnóstico, etiquetado, limpieza y pintado). Pendiente negativo = preparación adelantada a la meta.</p>
    <table>
      <thead>
        <tr>
          <th>Negocio</th><th>Producto</th><th id="weekColHeader">Meta semanal</th><th class="num-r">Meta mes</th>
          <th class="bar-cell">Preparado</th><th class="num-r">Pendiente</th>
        </tr>
      </thead>
      <tbody id="planTableBody"></tbody>
    </table>
  </div>

  <div class="section-label"><span class="num">03</span><h2>Detalle real de despachos</h2><span class="rule"></span></div>
  <div class="grid-2b">
    <div class="panel">
      <h3>Unidades por familia</h3>
      <p class="desc" id="familyDesc">—</p>
      <canvas id="familyChart" height="260"></canvas>
    </div>
    <div class="panel">
      <h3>Guías de despacho</h3>
      <p class="desc" id="guiaDesc">—</p>
      <div class="guia-list" id="guiaList"></div>
    </div>
  </div>

  <div class="section-label"><span class="num">04</span><h2>Horarios de entrega semanal</h2><span class="rule"></span></div>
  <div class="panel">
    <p class="desc">Ventanas de recepción confirmadas por semana</p>
    <table class="sched-table">
      <thead><tr><th>Semana</th><th>Lun</th><th>Mar</th><th>Mié</th><th>Jue</th><th>Vie</th></tr></thead>
      <tbody id="schedBody"></tbody>
    </table>
  </div>

  <footer>
    <span id="footerSource">Fuente: datos de ejemplo · hojas DESPACHOS, PLANIFICACION</span>
    <span>Panel generado automáticamente</span>
  </footer>
</div>

const scheduleRows = [
  {wk:28, dates:['06 jul','07 jul','08 jul','09 jul','10 jul'], slots:[null,'10:00',null,'10:00','12:00']},
  {wk:29, dates:['13 jul','14 jul','15 jul','16 jul','17 jul'], slots:[null,'10:00',null,'10:00','12:00']},
  {wk:30, dates:['20 jul','21 jul','22 jul','23 jul','24 jul'], slots:[null,'10:00',null,'10:00','12:00']},
  {wk:31, dates:['27 jul','28 jul','29 jul','30 jul','31 jul'], slots:[null,'10:00',null,'10:00','12:00']},
];

function findSheet(wb, re){
  const key = wb.SheetNames.find(n => re.test(n));
  return key ? wb.Sheets[key] : null;
}

function buildStateFromWorkbook(wb){
  const despSheet = findSheet(wb, /despacho/i);
  const planSheet = findSheet(wb, /planificaci/i);
  if(!despSheet) throw new Error("No se encontró una hoja 'DESPACHOS' en el archivo.");
  if(!planSheet) throw new Error("No se encontró una hoja 'PLANIFICACION' en el archivo.");

  const despRows = XLSX.utils.sheet_to_json(despSheet, {header:1, raw:true, defval:null});
  const headers = (despRows[0]||[]).map(h => (h||'').toString().trim());
  const idx = name => headers.indexOf(name);
  const iNeg = idx('NEGOCIO'), iFam = idx('FAMILIA'), iGuia = idx('GUIA'),
        iEstadoSap = idx('ESTADO SAP'), iEstado = idx('ESTADO'), iSemana = idx('SEMANA'),
        iFecha = idx('FECHA');
  if(iNeg<0 || iFam<0 || iGuia<0) throw new Error("La hoja DESPACHOS no tiene las columnas esperadas (NEGOCIO, FAMILIA, GUIA).");

  const dataRows = despRows.slice(1).filter(r => r && r[iNeg]);
  const byNeg = {}, byFamNeg = {}, guiaMap = {}, weekCounts = {};
  let fechaMin=null, fechaMax=null;

  dataRows.forEach(r=>{
    const neg = (r[iNeg]||'').toString().trim().toUpperCase();
    const fam = (r[iFam]||'').toString().trim();
    const guia = r[iGuia];
    const estadoRaw = (iEstadoSap>=0 ? r[iEstadoSap] : r[iEstado]) || '';
    const wk = iSemana>=0 ? (r[iSemana]||'').toString().trim().toUpperCase() : null;
    byNeg[neg] = (byNeg[neg]||0)+1;
    const famKey = fam+'||'+neg;
    byFamNeg[famKey] = (byFamNeg[famKey]||0)+1;
    if(wk) weekCounts[wk] = (weekCounts[wk]||0)+1;
    if(guia!=null && guia!==''){
      const gk = neg+'-'+guia;
      if(!guiaMap[gk]) guiaMap[gk] = {id:guia, neg, fam, u:0, estados:{}};
      guiaMap[gk].u++;
      guiaMap[gk].estados[estadoRaw] = (guiaMap[gk].estados[estadoRaw]||0)+1;
    }
    if(iFecha>=0 && r[iFecha]){
      const d = new Date(r[iFecha]);
      if(!isNaN(d)){ if(!fechaMin||d<fechaMin) fechaMin=d; if(!fechaMax||d>fechaMax) fechaMax=d; }
    }
  });

  if(dataRows.length===0) throw new Error("La hoja DESPACHOS no tiene filas de datos.");

  let currentWeek=null, maxCount=-1;
  Object.entries(weekCounts).forEach(([wk,c])=>{ if(c>maxCount){maxCount=c; currentWeek=wk;} });

  const guias = Object.values(guiaMap).map(g=>{
    const estadoFinal = Object.entries(g.estados).sort((a,b)=>b[1]-a[1])[0][0];
    return {id:g.id, neg:g.neg, fam:g.fam, u:g.u, valorado:/^valorado$/i.test((estadoFinal||'').trim())};
  }).sort((a,b)=> a.neg.localeCompare(b.neg) || (Number(a.id)-Number(b.id)));

  const familyBreakdown = Object.entries(byFamNeg).map(([k,v])=>{
    const [fam,neg]=k.split('||'); return {fam,neg,units:v};
  }).sort((a,b)=>b.units-a.units);

  const planRowsRaw = XLSX.utils.sheet_to_json(planSheet, {header:1, raw:true, defval:null});
  const headerRowIdx = planRowsRaw.findIndex(r => r.some(c => (c||'').toString().trim().toLowerCase()==='negocio'));
  if(headerRowIdx<0) throw new Error("No se encontró el encabezado 'Negocio' en la hoja PLANIFICACION.");
  const pHeaders = planRowsRaw[headerRowIdx].map(h=>(h||'').toString().trim());
  const pNegIdx = pHeaders.findIndex(h=>h.toLowerCase()==='negocio');
  const pProdIdx = pNegIdx+1;
  const weekIdxs = [];
  pHeaders.forEach((h,i)=>{ if(/^wk\d+$/i.test(h)) weekIdxs.push(i); });
  if(weekIdxs.length===0) throw new Error("No se encontraron columnas semanales (WKxx) en PLANIFICACION.");
  const lastWeekIdx = weekIdxs[weekIdxs.length-1];
  let metaTotalIdx=-1;
  for(let i=lastWeekIdx+1;i<pHeaders.length;i++){ if(pHeaders[i].toLowerCase()==='total'){ metaTotalIdx=i; break; } }
  let pendienteIdx=-1;
  for(let i=pHeaders.length-1;i>Math.max(metaTotalIdx,lastWeekIdx);i--){ if(/pendiente/i.test(pHeaders[i])){ pendienteIdx=i; break; } }
  let prepTotalIdx=-1;
  if(pendienteIdx>0){
    for(let i=pendienteIdx-1;i>metaTotalIdx;i--){ if(pHeaders[i].toLowerCase()==='total'){ prepTotalIdx=i; break; } }
  }

  const planDataRows = planRowsRaw.slice(headerRowIdx+1).filter(r=> r && r[pNegIdx] && r[pProdIdx]);
  if(planDataRows.length===0) throw new Error("La hoja PLANIFICACION no tiene filas de producto.");

  const planRows = planDataRows.map(r=>({
    neg: (r[pNegIdx]||'').toString().trim().toUpperCase(),
    prod: (r[pProdIdx]||'').toString().trim(),
    weeks: weekIdxs.map(i=> Number(r[i])||0),
    meta: metaTotalIdx>=0 ? (Number(r[metaTotalIdx])||0) : weekIdxs.reduce((s,i)=>s+(Number(r[i])||0),0),
    prep: prepTotalIdx>=0 ? (Number(r[prepTotalIdx])||0) : null,
    pend: pendienteIdx>=0 ? (Number(r[pendienteIdx])||0) : null,
  }));

  const weekLabels = weekIdxs.map(i=>pHeaders[i].toUpperCase());
  const weeklyByNegocio = {};
  planRows.forEach(r=>{
    if(!weeklyByNegocio[r.neg]) weeklyByNegocio[r.neg] = weekLabels.map(()=>0);
    r.weeks.forEach((v,i)=> weeklyByNegocio[r.neg][i]+=v);
  });
  const monthTotal = planRows.reduce((s,r)=>s+r.meta,0);
  const negTotals = {};
  planRows.forEach(r=> negTotals[r.neg]=(negTotals[r.neg]||0)+r.meta);

  let wkPlanIdx = weekLabels.indexOf(currentWeek);
  if(wkPlanIdx<0) wkPlanIdx = 0;
  const wkPlanTotal = Object.values(weeklyByNegocio).reduce((s,arr)=>s+arr[wkPlanIdx],0);
  const wkActual = iSemana>=0 ? dataRows.filter(r=>(r[iSemana]||'').toString().trim().toUpperCase()===currentWeek).length : dataRows.length;

  return {
    monthTotal, negTotals, weekLabels, weeklyByNegocio, planRows,
    currentWeek: currentWeek || weekLabels[wkPlanIdx],
    wkPlanTotal, wkActual, byNeg, familyBreakdown, guias,
    fechaMin, fechaMax, totalDespachos: dataRows.length, isDemo:false,
  };
}
function fmtDateRange(min,max){
  if(!min && !max) return '';
  const opts = {day:'2-digit', month:'short'};
  const a = min ? min.toLocaleDateString('es-CL', opts) : '';
  const b = max ? max.toLocaleDateString('es-CL', opts) : '';
  return (a===b || !b) ? a : (a + ' – ' + b);
}

function applyState(state){
  // ---- header ----
  document.getElementById('hdrWeek').textContent = state.currentWeek || '—';
  document.getElementById('hdrDate').textContent = state.isDemo
    ? ('Datos de ejemplo · ' + fmtDateRange(state.fechaMin, state.fechaMax) + ' 2026')
    : ('Actualizado ahora · ' + fmtDateRange(state.fechaMin, state.fechaMax));

  const negKeys = Object.keys(state.negTotals);

  // ---- KPI: plan mensual ----
  document.getElementById('kpiMonthTotal').innerHTML = fmt(state.monthTotal) + ' <small>und.</small>';
  document.getElementById('kpiMonthSub').innerHTML = 'Total planificado sobre ' + state.weekLabels.length + ' semanas (' + state.weekLabels[0] + '–' + state.weekLabels[state.weekLabels.length-1] + ')';

  // ---- KPI: avance semana ----
  const pct = state.wkPlanTotal>0 ? (state.wkActual/state.wkPlanTotal*100) : 0;
  document.getElementById('kpiAvanceLabel').textContent = 'Avance ' + (state.currentWeek||'') + ' vs meta';
  document.getElementById('kpiAvancePct').innerHTML = pct.toFixed(1).replace('.',',') + '<small>%</small>';
  document.getElementById('kpiAvanceSub').innerHTML = fmt(state.wkActual) + ' despachadas de ' + fmt(state.wkPlanTotal) + ' planificadas';
  document.getElementById('radialPct').textContent = pct.toFixed(1).replace('.',',') + '%';
  const circumference = 2*Math.PI*74;
  const offset = circumference * (1 - Math.min(1, pct/100));
  document.getElementById('radialArc').setAttribute('stroke-dasharray', circumference.toFixed(1));
  document.getElementById('radialArc').setAttribute('stroke-dashoffset', offset.toFixed(1));
  const restante = Math.max(0, state.wkPlanTotal - state.wkActual);
  document.getElementById('radialFoot').innerHTML = 'Despachadas <b>' + fmt(state.wkActual) + '</b> · Meta <b>' + fmt(state.wkPlanTotal) + '</b><br>Restan <b>' + fmt(restante) + '</b> und. para cerrar ' + (state.currentWeek||'la semana');

  // ---- KPI: por negocio (hasta 2 tarjetas fijas Claro/VTR, si existen) ----
  function fillNegTicket(prefix, negName){
    const meta = state.negTotals[negName]||0;
    const actual = state.byNeg[negName]||0;
    const nGuias = state.guias.filter(g=>g.neg===negName).length;
    document.getElementById('kpi'+prefix).innerHTML = fmt(meta) + ' <small>und./mes</small>';
    document.getElementById('kpi'+prefix+'Sub').innerHTML = (state.currentWeek||'Semana')+': <b>'+fmt(actual)+'</b> und. reales · '+nGuias+' guías';
  }
  fillNegTicket('Claro','CLARO');
  fillNegTicket('Vtr','VTR');

  // ---- SECTION 1: weekly chart ----
  const legend = document.getElementById('weeklyLegend');
  legend.innerHTML = negKeys.map((n,i)=>
    '<span class="item"><span class="sw" style="background:'+colorFor(n,i)+'"></span>'+n+'</span>'
  ).join('');
  document.getElementById('weeklyDesc').textContent = 'Distribución por negocio sobre las semanas de facturación ('+state.weekLabels.join('–')+')';

  if(weeklyChartInst) weeklyChartInst.destroy();
  weeklyChartInst = new Chart(document.getElementById('weeklyChart'), {
    type:'bar',
    data:{
      labels: state.weekLabels,
      datasets: negKeys.map((n,i)=>({
        label:n, data: state.weeklyByNegocio[n], backgroundColor: colorFor(n,i), borderRadius:4, maxBarThickness:46
      }))
    },
    options:{
      responsive:true,
      plugins:{ legend:{display:false}, tooltip:{backgroundColor:'#1C242A', borderColor:'#2B353C', borderWidth:1, padding:10} },
      scales:{
        x:{ stacked:true, grid:{display:false}, ticks:{font:{family:"'IBM Plex Mono', monospace", size:11}} },
        y:{ stacked:true, grid:{color:'#212A30'}, ticks:{font:{family:"'IBM Plex Mono', monospace", size:10}} }
      }
    }
  });

  // ---- SECTION 2: plan table ----
  document.getElementById('weekColHeader').textContent = 'Meta semanal ('+state.weekLabels[0]+'→'+state.weekLabels[state.weekLabels.length-1]+')';
  const planBody = document.getElementById('planTableBody');
  planBody.innerHTML = '';
  state.planRows.forEach(r=>{
    const hasPrep = r.prep!=null;
    const pct2 = hasPrep ? Math.min(100, Math.round((r.prep/r.meta)*100)) : 0;
    const color = colorFor(r.neg, negKeys.indexOf(r.neg));
    const tr = document.createElement('tr');
    let prepCell;
    if(hasPrep){
      prepCell = '<div class="bar-track"><div class="bar-fill" style="width:'+pct2+'%;background:'+color+'"></div></div>'
        + '<div style="font-size:10.5px;color:var(--text-faint);margin-top:4px;font-family:\'IBM Plex Mono\',monospace;">'+fmt(r.prep)+' und. · '+pct2+'%</div>';
    } else {
      prepCell = '<span style="color:var(--text-faint);font-family:\'IBM Plex Mono\',monospace;font-size:11px;">s/d</span>';
    }
    let pendCell = '—';
    if(r.pend!=null){
      const pendClass = r.pend<0 ? 'pend-neg' : 'pend-pos';
      const pendLabel = r.pend<0 ? (Math.abs(r.pend)+' adelanto') : (''+r.pend);
      pendCell = '<span class="'+pendClass+'">'+pendLabel+'</span>';
    }
    tr.innerHTML =
      '<td><span class="neg-tag '+tagClass(r.neg)+'">'+r.neg+'</span></td>'
      + '<td class="prod-name">'+r.prod+'</td>'
      + '<td class="mono" style="color:var(--text-faint);font-size:11px;">'+r.weeks.join(' · ')+'</td>'
      + '<td class="num-r">'+fmt(r.meta)+'</td>'
      + '<td class="bar-cell">'+prepCell+'</td>'
      + '<td class="num-r">'+pendCell+'</td>';
    planBody.appendChild(tr);
  });

  // ---- SECTION 3: family donut ----
  document.getElementById('familyDesc').textContent = fmt(state.totalDespachos)+' unidades despachadas'+(state.currentWeek?(' · '+state.currentWeek):'')+(state.fechaMin?(' · '+fmtDateRange(state.fechaMin,state.fechaMax)):'');
  const famData = state.familyBreakdown.slice(0,8);
  const famPalette = ['#2DD4BF','#FF6B4A','#1B6F65','#F5A623','#5E9E96','#8CA0AC','#B18CD9','#E8735A'];
  if(familyChartInst) familyChartInst.destroy();
  if(famData.length>0){
    familyChartInst = new Chart(document.getElementById('familyChart'), {
      type:'doughnut',
      data:{
        labels: famData.map(f=> f.fam+' ('+f.neg+')'),
        datasets:[{ data: famData.map(f=>f.units), backgroundColor: famPalette, borderColor:'#161D22', borderWidth:3 }]
      },
      options:{ cutout:'62%', plugins:{ legend:{ position:'bottom', labels:{ boxWidth:10, font:{size:11}, padding:14 } } } }
    });
  }

  // ---- SECTION 3: guias ----
  document.getElementById('guiaDesc').textContent = state.guias.length+' guías activas · estado de valorización SAP';
  const guiaList = document.getElementById('guiaList');
  guiaList.innerHTML = '';
  if(state.guias.length===0){
    guiaList.innerHTML = '<div class="empty-msg">Sin guías detectadas en los datos cargados</div>';
  } else {
    state.guias.forEach(g=>{
      const div = document.createElement('div');
      div.className='guia';
      div.innerHTML =
        '<span class="gid">#'+g.id+'</span>'
        + '<span class="gfam">'+g.fam+'<span class="neg">'+g.neg+'</span></span>'
        + '<span class="gunits">'+g.u+'</span>'
        + '<span class="gstate '+(g.valorado?'val':'nov')+'">'+(g.valorado?'Valorado':'No valorado')+'</span>';
      guiaList.appendChild(div);
    });
  }

  // ---- footer ----
  document.getElementById('footerSource').textContent = state.isDemo
    ? 'Fuente: datos de ejemplo · hojas DESPACHOS, PLANIFICACION'
    : 'Fuente: archivo cargado por el usuario · hojas DESPACHOS, PLANIFICACION';
}

function renderSchedule(){
  const schedBody = document.getElementById('schedBody');
  schedBody.innerHTML = '';
  scheduleRows.forEach(w=>{
    const tr = document.createElement('tr');
    let cells = '<td class="wk">'+w.wk+'</td>';
    w.dates.forEach((d,i)=>{
      cells += '<td><span class="date">'+d+'</span>'+(w.slots[i] ? ('<span class="slot">'+w.slots[i]+'</span>') : '—')+'</td>';
    });
    tr.innerHTML = cells;
    schedBody.appendChild(tr);
  });
}

// ---------------------------------------------------------
// EVENTS
// ---------------------------------------------------------
function setStatus(msg, cls){
  const el = document.getElementById('sourceStatus');
  el.textContent = msg;
  el.className = cls;
}

function handleFile(file){
  if(!file) return;
  setStatus('Leyendo '+file.name+'…', 'demo');
  const reader = new FileReader();
  reader.onload = function(evt){
    try{
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, {type:'array', cellDates:true});
      const state = buildStateFromWorkbook(wb);
      applyState(state);
      setStatus('✓ Datos cargados desde '+file.name+' · '+fmt(state.totalDespachos)+' registros', 'ok');
    }catch(err){
      console.error(err);
      setStatus('✗ '+err.message, 'err');
    }
  };
  reader.onerror = function(){ setStatus('✗ No se pudo leer el archivo.', 'err'); };
  reader.readAsArrayBuffer(file);
}

document.getElementById('fileInput').addEventListener('change', e=> handleFile(e.target.files[0]));
document.getElementById('resetBtn').addEventListener('click', ()=>{
  applyState(demoState);
  setStatus('Mostrando datos de ejemplo (WK28, julio 2026)', 'demo');
});

const loaderPanel = document.getElementById('loaderPanel');
['dragover','dragenter'].forEach(evt=> loaderPanel.addEventListener(evt, e=>{ e.preventDefault(); loaderPanel.classList.add('drag-over'); }));
['dragleave','drop'].forEach(evt=> loaderPanel.addEventListener(evt, e=>{ e.preventDefault(); loaderPanel.classList.remove('drag-over'); }));
loaderPanel.addEventListener('drop', e=>{
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if(file) handleFile(file);
});

renderSchedule();
applyState(demoState);
</script>
</body>
</html>
