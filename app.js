// ══════════════════════════════════════════════════════
//  Cascos Analytics v2 — PezzaliApp
//  Engine completo: filtri, drill-down, confronto, sconti
// ══════════════════════════════════════════════════════
'use strict';

const SCONTO_MAX = 0.60;
const MESI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
const QNAMES = ['Q1','Q2','Q3','Q4'];
const PORTO_MAP = {1:'Franco',2:'Assegnato',3:'Franco+Add.',6:'Altro'};
const PAL = ['#00e5a0','#4db8ff','#ffb547','#b47aff','#ff5f72','#00d4e8','#ff8c42','#7fff6e','#ff79a8','#6ee7ff'];
const PAL2 = ['rgba(0,229,160,.7)','rgba(77,184,255,.7)','rgba(255,181,71,.7)','rgba(180,122,255,.7)','rgba(255,95,114,.7)'];

let G = {}; // global data store
let F = {}; // filtered data cache
let CMP = false; // compare mode
let charts = {};
let sortState = {}; // table sort state

// ── file upload wiring ──────────────────────────────
['v','o','l'].forEach(id => {
  document.getElementById('fi-'+id).addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    document.getElementById('uc-'+id).classList.add('ok');
    document.getElementById('fn-'+id).textContent = f.name;
    checkReady();
  });
});
function checkReady() {
  const ok = ['v','o'].every(id => document.getElementById('uc-'+id).classList.contains('ok'));
  document.getElementById('btn-go').classList.toggle('on', ok);
}
// drag & drop
document.querySelectorAll('.upc').forEach(card => {
  card.addEventListener('dragover', e => { e.preventDefault(); card.classList.add('drag'); });
  card.addEventListener('dragleave', () => card.classList.remove('drag'));
  card.addEventListener('drop', e => {
    e.preventDefault(); card.classList.remove('drag');
    const inp = card.querySelector('input[type=file]');
    if (!inp || !e.dataTransfer.files[0]) return;
    const dt = new DataTransfer(); dt.items.add(e.dataTransfer.files[0]); inp.files = dt.files;
    inp.dispatchEvent(new Event('change'));
  });
});

// ── MAIN RUN ───────────────────────────────────────
async function runAnalysis() {
  showLoad('Lettura file vendite...');
  await sleep(30);
  try {
    const vRaw = await readXLSX('fi-v');
    setLoad('Lettura ordini...'); await sleep(20);
    const oRaw = await readXLSX('fi-o');
    setLoad('Caricamento listino...'); await sleep(20);
    let lRaw = null;
    const lInp = document.getElementById('fi-l');
    if (lInp.files[0]) {
      const ext = lInp.files[0].name.split('.').pop().toLowerCase();
      lRaw = ext === 'csv' ? await readCSV('fi-l') : await readXLSX('fi-l');
    }
    setLoad('Elaborazione dati...'); await sleep(30);
    processData(vRaw, oRaw, lRaw);
    setLoad('Rendering cruscotto...'); await sleep(30);
    initDashboard();
    hideLoad();
    document.getElementById('upload-screen').style.display = 'none';
    document.getElementById('top-filters').removeAttribute('hidden');
    document.getElementById('btn-reset').removeAttribute('hidden');
    document.getElementById('status-pill').textContent = 'ATTIVO';
    document.getElementById('status-pill').className = 'pill pill-g';
  } catch(err) {
    hideLoad(); alert('Errore: ' + err.message); console.error(err);
  }
}

// ── READ HELPERS ────────────────────────────────────
function readXLSX(id) {
  return new Promise((res, rej) => {
    const f = document.getElementById(id).files[0];
    if (!f) return rej(new Error('File mancante: ' + id));
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
        res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval:'' }));
      } catch(err) { rej(err); }
    };
    r.onerror = () => rej(new Error('Lettura fallita'));
    r.readAsArrayBuffer(f);
  });
}
function readCSV(id) {
  return new Promise((res, rej) => {
    Papa.parse(document.getElementById(id).files[0], {
      header:true, skipEmptyLines:true, dynamicTyping:true,
      complete: r => res(r.data), error: e => rej(e)
    });
  });
}

// ── PROCESS DATA ────────────────────────────────────
function processData(vRaw, oRaw, lRaw) {
  // listino map
  const listinoMap = {};
  if (lRaw) {
    lRaw.forEach(r => {
      const cod = str(r.Codice || r.codice || r.CODICE || '').replace(/^0+/, '');
      const pl = num(r.PrezzoLordo || r.prezzo_lordo || r.PREZZO_LORDO);
      if (cod && pl > 0) listinoMap[cod] = { pl, inst: num(r.CostoInstallazione||0), trasp: num(r.CostoTrasporto||0) };
    });
  }

  // detect vendite columns
  const keys = Object.keys(vRaw[0] || {});
  const col = (cands) => cands.find(c => keys.includes(c)) || cands[0];
  const CV = {
    anno: col(['ANNO SPEDIZIONE','ANNO']),
    data: col(['DATA SPEDIZIONE','DATA']),
    importo: col(['IMPORTO CONSEGNATO','IMPORTO']),
    pz: col(['PZ NETTO VENDITA','PZ NETTO']),
    qty: col(['QTA CONSEGNATA','QTA']),
    trasp: col(['SPESE DI TRASPORTO','TRASPORTO']),
    causale: col(['CAUSALE MAGAZZINO','CAUSALE']),
    cat: col(['DESCRIZIONE ELEMENTO.5','CATEGORIA']),
    agente: col(['DESCRIZIONE ELEMENTO.2','AGENTE']),
    cliente: col(['RAGIONE SOCIALE 1','CLIENTE']),
    articolo: col(['ARTICOLO']),
    desc: col(['DESCRIZIONE']),
    porto: col(['PORTO']),
    ordine: col(['NUMERO ORDINE','NUM.ORDINE']),
  };

  // filter sales
  const VEND = vRaw
    .filter(r => str(r[CV.causale]).toUpperCase().startsWith('V'))
    .map(r => {
      const cod = str(r[CV.articolo]).replace(/^0+/,'');
      const li = listinoMap[cod] || null;
      const pz = num(r[CV.pz]);
      const importo = num(r[CV.importo]);
      const trasp = num(r[CV.trasp]);
      const lordo = li ? li.pl : null;
      const sconto = lordo && lordo > 0 && pz > 0 ? Math.max(0, Math.min(1, 1 - pz/lordo)) : null;
      const anno = parseInt(r[CV.anno]) || 0;
      let date = null, mese = -1, trim = -1;
      try {
        const d = r[CV.data];
        date = d instanceof Date ? d : new Date(d);
        if (!isNaN(date)) { mese = date.getMonth(); trim = Math.floor(mese/3); }
      } catch(e){}
      return {
        anno, date, mese, trim,
        importo, pz, qty: num(r[CV.qty]), trasp,
        cat: str(r[CV.cat]), agente: str(r[CV.agente]),
        cliente: str(r[CV.cliente]),
        articolo: str(r[CV.articolo]), desc: str(r[CV.desc]),
        porto: r[CV.porto], lordo, sconto,
        incTrasp: importo > 0 ? trasp/importo : 0,
        _raw: r
      };
    })
    .filter(r => r.anno > 2000);

  const RESI = vRaw.filter(r => ['H1','X2','X4','X5'].includes(str(r[CV.causale]).toUpperCase()));
  const anni = [...new Set(VEND.map(r => r.anno))].sort();
  const agenti = [...new Set(VEND.map(r => r.agente).filter(Boolean))].sort();

  // ordini
  const OK = Object.keys(oRaw[0] || {});
  const CO = {
    cliente: ['CLIENTE.1','CLIENTE','RAGIONE SOCIALE'].find(c=>OK.includes(c)),
    desc: ['DESCRIZIONE'].find(c=>OK.includes(c)),
    qtyI: ['QTA INEVASA'].find(c=>OK.includes(c)),
    importoI: ['IMPORTO INEVASO'].find(c=>OK.includes(c)),
    imponibile: ['IMPONIBILE ORD. VAL','IMPONIBILE ORD'].find(c=>OK.includes(c)),
    trasp: ['SPESE DI TRASPORTO'].find(c=>OK.includes(c)),
    consegna: ['DATA CONSEGNA'].find(c=>OK.includes(c)),
    creazione: ['DATA CREAZIONE'].find(c=>OK.includes(c)),
    porto: ['PORTO'].find(c=>OK.includes(c)),
    num: ['NUM.'].find(c=>OK.includes(c)),
  };
  const ORDINI = oRaw.map(r => {
    let consegna = null;
    try { const d = r[CO.consegna]; consegna = d instanceof Date ? d : new Date(d); if(isNaN(consegna)) consegna=null; } catch(e){}
    return {
      cliente: str(r[CO.cliente]), desc: str(r[CO.desc]),
      qtyI: num(r[CO.qtyI]), importoI: num(r[CO.importoI]),
      imponibile: num(r[CO.imponibile]), trasp: num(r[CO.trasp]),
      consegna, porto: r[CO.porto],
      num: str(r[CO.num]),
    };
  });

  // store
  G = { VEND, RESI, ORDINI, anni, agenti, listinoMap, CV };

  // populate filter dropdowns
  const da = document.getElementById('f-da');
  const fa = document.getElementById('f-a');
  const fc = document.getElementById('f-cmp-a');
  const ta1 = document.getElementById('tr-a1');
  const ta2 = document.getElementById('tr-a2');
  const fagt = document.getElementById('f-agt');
  const cliSel = document.getElementById('cli-sel');

  da.innerHTML = fa.innerHTML = fc.innerHTML = ta1.innerHTML = ta2.innerHTML = '';
  anni.forEach(a => {
    [da, fa, fc, ta1, ta2].forEach(s => s.insertAdjacentHTML('beforeend', `<option value="${a}">${a}</option>`));
  });
  da.value = anni[0]; fa.value = anni[anni.length-1];
  fc.value = anni[anni.length-2] || anni[0];
  ta1.value = anni[anni.length-2] || anni[0];
  ta2.value = anni[anni.length-1];

  fagt.innerHTML = '<option value="">Tutti</option>';
  agenti.forEach(a => fagt.insertAdjacentHTML('beforeend', `<option value="${a}">${a}</option>`));

  cliSel.innerHTML = '';
  const topCli = [...new Set(VEND.map(r=>r.cliente).filter(Boolean))].slice(0,50);
  topCli.forEach(c => cliSel.insertAdjacentHTML('beforeend', `<option value="${c}">${c}</option>`));

  applyFilters();
}

// ── FILTER ENGINE ───────────────────────────────────
function applyFilters() {
  const annoDa = parseInt(document.getElementById('f-da').value) || 0;
  const annoA  = parseInt(document.getElementById('f-a').value) || 9999;
  const perVal = document.getElementById('f-per').value;
  const agente = document.getElementById('f-agt').value;

  // resolve mesi
  let mesiOk = null;
  if (perVal.startsWith('q')) {
    const q = parseInt(perVal[1])-1;
    mesiOk = [q*3, q*3+1, q*3+2];
  } else if (perVal.startsWith('m')) {
    mesiOk = [parseInt(perVal.slice(1))-1];
  }

  F.vend = G.VEND.filter(r => {
    if (r.anno < annoDa || r.anno > annoA) return false;
    if (mesiOk && !mesiOk.includes(r.mese)) return false;
    if (agente && r.agente !== agente) return false;
    return true;
  });

  // compare period
  if (CMP) {
    const cmpAnno = parseInt(document.getElementById('f-cmp-a').value) || 0;
    F.vendCmp = G.VEND.filter(r => {
      if (r.anno !== cmpAnno) return false;
      if (mesiOk && !mesiOk.includes(r.mese)) return false;
      if (agente && r.agente !== agente) return false;
      return true;
    });
  } else F.vendCmp = null;

  // period label
  let plabel = `${annoDa}–${annoA}`;
  if (perVal) {
    if (perVal.startsWith('q')) plabel += ` Q${perVal[1]}`;
    else if (perVal.startsWith('m')) plabel += ` ${MESI[parseInt(perVal.slice(1))-1]}`;
  }
  if (agente) plabel += ` · ${agente}`;
  F.label = plabel;
  document.getElementById('sb-period').textContent = plabel;

  renderAll();
}

function toggleCompare() {
  CMP = !CMP;
  document.getElementById('chip-cmp').classList.toggle('on', CMP);
  document.getElementById('f-cmp-a').style.display = CMP ? 'inline-block' : 'none';
  applyFilters();
}

// ── RENDER ALL ──────────────────────────────────────
function renderAll() {
  renderOverview();
  renderTrend();
  renderVendite();
  renderClienti();
  renderAgenti();
  renderSconti();
  renderMargine();
  renderTrasporti();
  renderOrdini();
  renderCriticita();
}

function initDashboard() {
  renderAll();
  go('overview');
}

// ═══════════════════════════════════════════════════
//  PANEL RENDERERS
// ═══════════════════════════════════════════════════

// ── OVERVIEW ────────────────────────────────────────
function renderOverview() {
  const V = F.vend;
  const anni = G.anni;
  const fattTot = sum(V, r=>r.importo);
  const traspTot = sum(V, r=>r.trasp);
  const annoMax = Math.max(...anni);

  // Δ vs previous year (same period)
  const aF = parseInt(document.getElementById('f-a').value);
  const vCur = V.filter(r=>r.anno===aF);
  const vPrev = V.filter(r=>r.anno===aF-1);
  const fCur = sum(vCur, r=>r.importo);
  const fPrev = sum(vPrev, r=>r.importo);
  const delta = fPrev > 0 ? (fCur-fPrev)/fPrev : null;

  const scontiRows = V.filter(r=>r.sconto!==null);
  const scontoMed = scontiRows.length ? avg(scontiRows, r=>r.sconto) : null;
  const over60n = scontiRows.filter(r=>r.sconto>SCONTO_MAX).length;

  document.getElementById('ov-sub').textContent = `${F.label} · ${V.length.toLocaleString('it')} righe vendita`;

  kpi('kr-ov', [
    { l:'Fatturato Periodo', v:fmt(fattTot), cls:'g', sub:F.label },
    { l:`Fatturato ${aF}`, v:fmt(fCur), cls:'g', delta, sub:`vs ${aF-1}` },
    { l:'Spese Trasporto', v:fmt(traspTot), cls:'p', sub:`${pct(fattTot>0?traspTot/fattTot:0)} su fatturato` },
    { l:'Sconto Medio', v:scontoMed!==null?pct(scontoMed):'N/D', cls:scontoMed>SCONTO_MAX?'r':'g', sub:'Su articoli in listino' },
    { l:'Righe >60%', v:over60n.toLocaleString('it'), cls:over60n>0?'r':'g', sub:'Oltre soglia contrattuale' },
    { l:'Ordini Inevasi', v:fmt(sum(G.ORDINI,r=>r.importoI)), cls:'a', sub:`${G.ORDINI.length} righe` },
  ]);

  // annual chart
  const fattAnnuo = groupBy(G.VEND, r=>r.anno, rows=>sum(rows,r=>r.importo));
  const traspAnnuo = groupBy(G.VEND, r=>r.anno, rows=>sum(rows,r=>r.trasp));
  lineBar('ch-annual', anni, anni.map(a=>fattAnnuo[a]||0), anni.map(a=>traspAnnuo[a]||0),
    'Fatturato','Trasporto','rgba(0,229,160,.7)','#ff5f72');

  // pie
  const catFatt = groupBy(V.filter(r=>r.cat), r=>r.cat, rows=>sum(rows,r=>r.importo));
  const catSorted = Object.entries(catFatt).sort((a,b)=>b[1]-a[1]).slice(0,8);
  doPie('ch-pie', catSorted.map(([k])=>k.split(' - ')[0]), catSorted.map(([,v])=>v));

  // quarterly
  const trimFatt = {};
  anni.forEach(a => { QNAMES.forEach((_,q) => { trimFatt[`${a}-Q${q+1}`]=0; }); });
  G.VEND.forEach(r => { if(r.trim>=0) trimFatt[`${r.anno}-Q${r.trim+1}`]=(trimFatt[`${r.anno}-Q${r.trim+1}`]||0)+r.importo; });
  const trimLabels = []; const trimData = [];
  anni.forEach(a => QNAMES.forEach((_,q) => { trimLabels.push(`${a} Q${q+1}`); trimData.push(trimFatt[`${a}-Q${q+1}`]||0); }));
  doBar('ch-qtr', trimLabels, trimData, 'rgba(77,184,255,.7)');

  // top clients
  const cliF = groupBy(V, r=>r.cliente, rows=>sum(rows,r=>r.importo));
  const topCli = Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,10);
  doHBar('ch-top-cli', topCli.map(([k])=>trunc(k,22)), topCli.map(([,v])=>v), '#00e5a0');
}

// ── TREND ────────────────────────────────────────────
function renderTrend() {
  const a1 = parseInt(document.getElementById('tr-a1').value);
  const a2 = parseInt(document.getElementById('tr-a2').value);
  const view = document.getElementById('tr-view').value; // m | q
  const met = document.getElementById('tr-met').value;   // f t s n

  const labels = view==='m' ? MESI : QNAMES;
  const n = labels.length;

  function serie(anno) {
    return labels.map((_, i) => {
      const rows = G.VEND.filter(r => r.anno===anno && (view==='m' ? r.mese===i : r.trim===i));
      if (met==='f') return sum(rows,r=>r.importo);
      if (met==='t') return sum(rows,r=>r.trasp);
      if (met==='s') { const sr=rows.filter(r=>r.sconto!==null); return sr.length?avg(sr,r=>r.sconto)*100:0; }
      if (met==='n') return rows.length;
      return 0;
    });
  }

  const d1 = serie(a1); const d2 = serie(a2);
  const metLbl = met==='f'?'Fatturato':met==='t'?'Trasporto':met==='s'?'Sconto %':'N° Righe';
  document.getElementById('tr-title').textContent = `Confronto ${view==='m'?'Mensile':'Trimestrale'} — ${metLbl}`;
  document.getElementById('tr-sub').textContent = `${a1} vs ${a2}`;

  // delta KPI
  const tot1=d1.reduce((a,b)=>a+b,0), tot2=d2.reduce((a,b)=>a+b,0);
  const dPct = tot1>0?(tot2-tot1)/tot1:null;
  document.getElementById('tr-kpi-delta').innerHTML = dPct!==null
    ? `<span class="dt ${dPct>=0?'up':'dn'}">${dPct>=0?'↑':'↓'} ${pct(Math.abs(dPct))}</span>` : '';

  // trend chart
  dc('ch-trend'); const ctx = document.getElementById('ch-trend').getContext('2d');
  charts['ch-trend'] = new Chart(ctx, {
    data: { labels,
      datasets: [
        { type:'bar', label:`${a1}`, data:d1, backgroundColor:'rgba(77,184,255,.6)', borderRadius:3 },
        { type:'line', label:`${a2}`, data:d2, borderColor:'#00e5a0', tension:.3, pointRadius:5, fill:false }
      ]
    },
    options: chartOpts({ callbackY: met==='f'||met==='t' ? v=>fmtShort(v) : met==='s' ? v=>v.toFixed(1)+'%' : v=>v })
  });

  // cumulative
  const cum1=[]; const cum2=[];
  d1.reduce((acc,v,i)=>{ cum1[i]=acc+v; return acc+v; }, 0);
  d2.reduce((acc,v,i)=>{ cum2[i]=acc+v; return acc+v; }, 0);
  dc('ch-cumul'); const ctx2=document.getElementById('ch-cumul').getContext('2d');
  charts['ch-cumul'] = new Chart(ctx2, {
    data: { labels, datasets:[
      { type:'line', label:`${a1}`, data:cum1, borderColor:'rgba(77,184,255,.8)', fill:true, backgroundColor:'rgba(77,184,255,.08)', tension:.3, pointRadius:3 },
      { type:'line', label:`${a2}`, data:cum2, borderColor:'#00e5a0', fill:true, backgroundColor:'rgba(0,229,160,.08)', tension:.3, pointRadius:3 }
    ]},
    options: chartOpts({ legend:true })
  });

  // delta chart
  const deltas = d1.map((v,i) => d1[i]>0 ? ((d2[i]-d1[i])/d1[i])*100 : 0);
  dc('ch-delta'); const ctx3=document.getElementById('ch-delta').getContext('2d');
  charts['ch-delta'] = new Chart(ctx3, {
    type:'bar',
    data: { labels, datasets:[{ data:deltas,
      backgroundColor: deltas.map(d=>d>=0?'rgba(0,229,160,.7)':'rgba(255,95,114,.7)'),
      borderRadius:3 }]},
    options: chartOpts({ callbackY: v=>v.toFixed(1)+'%' })
  });
}

// ── VENDITE ──────────────────────────────────────────
function renderVendite() {
  const V = F.vend;
  const catFatt = groupBy(V.filter(r=>r.cat), r=>r.cat, rows=>({
    f:sum(rows,r=>r.importo), q:sum(rows,r=>r.qty),
    n:rows.length, sc:avg(rows.filter(r=>r.sconto!==null),r=>r.sconto)
  }));
  const cats = Object.entries(catFatt).sort((a,b)=>b[1].f-a[1].f);

  kpi('kr-vend', [
    { l:'Fatturato Totale', v:fmt(sum(V,r=>r.importo)), cls:'g' },
    { l:'Righe Vendita', v:V.length.toLocaleString('it'), cls:'b' },
    { l:'Prezzo Netto Medio', v:fmt(avg(V,r=>r.pz)), cls:'g', sub:'Per unità' },
    { l:'Categorie Attive', v:cats.length, cls:'p' },
    { l:'Ticket Medio Riga', v:fmt(V.length?sum(V,r=>r.importo)/V.length:0), cls:'b' },
  ]);

  const catLabels = cats.map(([k])=>k.split(' - ')[0]);
  const catVals = cats.map(([,v])=>v.f);
  const catQty = cats.map(([,v])=>v.q);

  // bar with click drill-down
  dc('ch-cat-bar');
  const ctx = document.getElementById('ch-cat-bar').getContext('2d');
  charts['ch-cat-bar'] = new Chart(ctx, {
    type:'bar',
    data: { labels:catLabels, datasets:[{data:catVals, backgroundColor:PAL2, borderRadius:4}]},
    options: {
      ...chartOpts({callbackY:v=>fmtShort(v)}),
      onClick: (evt, els) => {
        if (!els.length) return;
        const idx = els[0].index;
        const catKey = cats[idx][0];
        showDrill(catKey, V);
      }
    }
  });

  doBar('ch-cat-qty', catLabels, catQty, 'rgba(180,122,255,.7)');

  // summary table
  tbl('tbl-cat',
    ['Categoria','Fatturato','%','Pezzi','Sconto Medio'],
    cats.map(([k,v]) => [
      trunc(k,30),
      `<span class="mo">${fmt(v.f)}</span>`,
      `<span class="bdg bg">${pct(sum(F.vend,r=>r.importo)>0?v.f/sum(F.vend,r=>r.importo):0)}</span>`,
      `<span class="mo">${v.q.toLocaleString('it',{maximumFractionDigits:0})}</span>`,
      v.sc>0 ? `<span class="bdg ${v.sc>SCONTO_MAX?'br':'bg'}">${pct(v.sc)}</span>` : '—'
    ])
  );
}

function showDrill(catKey, V) {
  document.getElementById('drill-cc').removeAttribute('hidden');
  document.getElementById('drill-t').textContent = `Clienti → ${catKey.split(' - ')[0]}`;
  const rows = V.filter(r=>r.cat===catKey);
  const cliF = groupBy(rows, r=>r.cliente, r=>sum(r,x=>x.importo));
  const sorted = Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,20);
  const totCat = sum(rows,r=>r.importo);
  tbl('tbl-drill',
    ['Cliente','Fatturato','%'],
    sorted.map(([k,v]) => [k, `<span class="mo">${fmt(v)}</span>`, `<span class="bdg bg">${pct(totCat>0?v/totCat:0)}</span>`])
  );
}
function closeDrill() { document.getElementById('drill-cc').setAttribute('hidden',''); }

// ── CLIENTI ──────────────────────────────────────────
let cliAllData = [];
function renderClienti() {
  const V = F.vend;
  const cliF = groupBy(V, r=>r.cliente, rows=>({
    f:sum(rows,r=>r.importo), n:rows.length,
    sc:avg(rows.filter(r=>r.sconto!==null),r=>r.sconto),
    tr:sum(rows,r=>r.trasp)
  }));
  const sorted = Object.entries(cliF).sort((a,b)=>b[1].f-a[1].f);
  cliAllData = sorted;

  const top12 = sorted.slice(0,12);
  doHBar('ch-cli-rank', top12.map(([k])=>trunc(k,22)), top12.map(([,v])=>v.f), '#00e5a0');

  renderCliStorico();
  filterCliTbl();
}

function renderCliStorico() {
  const sel = document.getElementById('cli-sel').value;
  if (!sel) return;
  const rows = G.VEND.filter(r=>r.cliente===sel && r.mese>=0);
  // monthly across all years
  const byAnnoMese = {};
  rows.forEach(r => {
    const k = `${r.anno}-${r.mese}`;
    byAnnoMese[k] = (byAnnoMese[k]||0) + r.importo;
  });
  const anni = [...new Set(rows.map(r=>r.anno))].sort();
  const datasets = anni.map((a,i) => ({
    type:'line', label:`${a}`,
    data: MESI.map((_,m) => byAnnoMese[`${a}-${m}`]||0),
    borderColor: PAL[i%PAL.length],
    backgroundColor: PAL[i%PAL.length]+'20',
    tension:.3, pointRadius:3, fill:false
  }));
  dc('ch-cli-st');
  const ctx = document.getElementById('ch-cli-st').getContext('2d');
  charts['ch-cli-st'] = new Chart(ctx, {
    data: { labels:MESI, datasets },
    options: chartOpts({ legend:true, callbackY:v=>fmtShort(v) })
  });
}

function filterCliTbl() {
  const q = document.getElementById('cli-srch').value.toLowerCase();
  const rows = cliAllData.filter(([k])=>k.toLowerCase().includes(q)).slice(0,60);
  const totF = sum(F.vend,r=>r.importo);
  tbl('tbl-cli',
    ['Cliente','Fatturato','%','N° Righe','Sconto Medio','Trasporto'],
    rows.map(([k,v]) => [
      k,
      `<span class="mo">${fmt(v.f)}</span>`,
      `<span class="bdg bg">${pct(totF>0?v.f/totF:0)}</span>`,
      v.n,
      v.sc>0?`<span class="bdg ${v.sc>SCONTO_MAX?'br':'bg'}">${pct(v.sc)}</span>`:'—',
      `<span class="mo">${fmt(v.tr)}</span>`
    ])
  );
}

// ── AGENTI ───────────────────────────────────────────
function renderAgenti() {
  const V = F.vend;
  const agtF = groupBy(V.filter(r=>r.agente), r=>r.agente, rows=>({
    f:sum(rows,r=>r.importo), n:rows.length, tr:sum(rows,r=>r.trasp),
    sc:avg(rows.filter(r=>r.sconto!==null),r=>r.sconto)
  }));
  const sorted = Object.entries(agtF).sort((a,b)=>b[1].f-a[1].f);
  const totF = sum(V,r=>r.importo);

  kpi('kr-agt', sorted.slice(0,4).map(([k,v],i) => ({
    l:`#${i+1} ${k}`, v:fmt(v.f), cls:['g','b','a','p'][i]||'p',
    sub:`${pct(totF>0?v.f/totF:0)} del totale`
  })));

  doHBar('ch-agt-bar', sorted.map(([k])=>k), sorted.map(([,v])=>v.f), null,
    sorted.map((_,i)=>PAL[i%PAL.length]+'bb'));

  // evolution per anno
  const anni = G.anni;
  const datasets = sorted.map(([ k],i) => ({
    type:'line', label:k,
    data: anni.map(a => {
      const r2 = G.VEND.filter(r=>r.anno===a && r.agente===k);
      return sum(r2,r=>r.importo);
    }),
    borderColor: PAL[i%PAL.length], tension:.3, pointRadius:3, fill:false
  }));
  dc('ch-agt-evol');
  const ctx = document.getElementById('ch-agt-evol').getContext('2d');
  charts['ch-agt-evol'] = new Chart(ctx, {
    data: { labels:anni, datasets },
    options: chartOpts({ legend:true, callbackY:v=>fmtShort(v) })
  });

  // agent rank cards
  const agEl = document.getElementById('agr');
  agEl.innerHTML = '';
  const maxF = sorted[0]?.[1]?.f || 1;
  sorted.forEach(([k,v],i) => {
    agEl.insertAdjacentHTML('beforeend', `
      <div class="agrow">
        <div class="agrow-top">
          <span class="agno">${String(i+1).padStart(2,'0')}</span>
          <span class="agname">${k}</span>
          <span class="agval">${fmt(v.f)}</span>
        </div>
        <div class="agbar"><div class="agfill" style="width:${Math.round(v.f/maxF*100)}%"></div></div>
      </div>`);
  });
}

// ── SCONTI ────────────────────────────────────────────
let scontiData = [];
function renderSconti() {
  const V = F.vend;
  const sR = V.filter(r=>r.sconto!==null);
  const sMedia = sR.length ? avg(sR,r=>r.sconto) : 0;
  const over60 = sR.filter(r=>r.sconto>SCONTO_MAX);

  kpi('kr-sc', [
    { l:'Sconto Medio Periodo', v:pct(sMedia), cls:sMedia>SCONTO_MAX?'r':'g', sub:`Su ${sR.length} righe con listino` },
    { l:'Righe >60%', v:over60.length.toLocaleString('it'), cls:'r', sub:fmt(sum(over60,r=>r.importo)) },
    { l:'Soglia Contrattuale', v:pct(SCONTO_MAX), cls:'a', sub:'Max rivenditori' },
    { l:'Articoli in Listino', v:sR.length.toLocaleString('it'), cls:'b', sub:`${pct(V.length?sR.length/V.length:0)} del totale` },
  ]);

  // annual sconto
  const anni = G.anni;
  const scAnno = anni.map(a => {
    const r2 = G.VEND.filter(r=>r.anno===a && r.sconto!==null);
    return r2.length ? avg(r2,r=>r.sconto)*100 : 0;
  });
  dc('ch-sc-yr');
  const ctx = document.getElementById('ch-sc-yr').getContext('2d');
  charts['ch-sc-yr'] = new Chart(ctx, {
    type:'line',
    data: { labels:anni, datasets:[
      { label:'Sconto %', data:scAnno, borderColor:'#b47aff', tension:.3, fill:true, backgroundColor:'rgba(180,122,255,.1)', pointRadius:5 },
      { label:'Soglia 60%', data:anni.map(()=>60), borderColor:'#ff5f72', borderDash:[6,4], pointRadius:0, borderWidth:1.5 }
    ]},
    options: chartOpts({ legend:true, callbackY:v=>v.toFixed(1)+'%' })
  });

  // distribution
  const buckets = Array(10).fill(0);
  sR.forEach(r => { buckets[Math.min(9, Math.floor(r.sconto*100/10))]++; });
  doBar('ch-sc-dist',
    ['0-10%','10-20%','20-30%','30-40%','40-50%','50-60%','60-70%','70-80%','80-90%','90-100%'],
    buckets, null, buckets.map((_,i)=>i>=6?'rgba(255,95,114,.75)':'rgba(0,229,160,.65)')
  );

  // sconti per prodotto
  const prodSconto = groupBy(sR, r=>r.desc, rows=>({
    sc:avg(rows,r=>r.sconto), n:rows.length,
    f:sum(rows,r=>r.importo), pz:avg(rows,r=>r.pz),
    lordo:rows[0].lordo
  }));
  scontiData = Object.entries(prodSconto).sort((a,b)=>b[1].f-a[1].f);
  renderScontiTbl();

  // top clienti over60
  const over60cli = groupBy(over60, r=>r.cliente, rows=>({ n:rows.length, f:sum(rows,r=>r.importo), sc:avg(rows,r=>r.sconto) }));
  const over60sorted = Object.entries(over60cli).sort((a,b)=>b[1].n-a[1].n).slice(0,20);
  tbl('tbl-over60-cli',
    ['Cliente','Righe >60%','Valore','Sconto Medio'],
    over60sorted.map(([k,v])=>[
      k, `<span class="bdg br">${v.n}</span>`,
      `<span class="mo">${fmt(v.f)}</span>`,
      `<span class="bdg ${v.sc>SCONTO_MAX?'br':'bg'}">${pct(v.sc)}</span>`
    ])
  );
}

function renderScontiTbl() {
  const flt = document.getElementById('sc-flt').value;
  const q = document.getElementById('sc-srch').value.toLowerCase();
  let rows = scontiData.filter(([k]) => k.toLowerCase().includes(q));
  if (flt==='over') rows = rows.filter(([,v])=>v.sc>SCONTO_MAX);
  if (flt==='ok') rows = rows.filter(([,v])=>v.sc<=SCONTO_MAX);
  tbl('tbl-sc',
    ['Prodotto/Macchina','Sconto Medio','N°','Fatturato','PZ Netto Medio','Lordo Listino'],
    rows.slice(0,80).map(([k,v])=>[
      trunc(k,32),
      `<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${v.sc>SCONTO_MAX?'⚠ ':''}${pct(v.sc)}</span>`,
      v.n, `<span class="mo">${fmt(v.f)}</span>`,
      `<span class="mo">${fmt(v.pz)}</span>`,
      v.lordo?`<span class="mo">${fmt(v.lordo)}</span>`:'—'
    ])
  );
}

// ── MARGINE ───────────────────────────────────────────
function renderMargine() {
  const V = F.vend;
  const catMarg = groupBy(V.filter(r=>r.cat && r.sconto!==null), r=>r.cat, rows=>({
    sc:avg(rows,r=>r.sconto), tr:avg(rows,r=>r.incTrasp),
    f:sum(rows,r=>r.importo), n:rows.length
  }));
  const sorted = Object.entries(catMarg).map(([k,v])=>[k,{...v,erosione:v.sc+v.tr}])
    .filter(([k])=>k).sort((a,b)=>b[1].erosione-a[1].erosione);

  // heatmap
  const hmEl = document.getElementById('hm-margine'); hmEl.innerHTML='';
  const maxEr = sorted[0]?.[1]?.erosione||1;
  sorted.forEach(([k,v]) => {
    const col = v.erosione>0.8?'#ff5f72':v.erosione>0.65?'#ffb547':'#00e5a0';
    hmEl.insertAdjacentHTML('beforeend', `
      <div class="hmc">
        <div class="hmg" style="background:${col}"></div>
        <div class="hmn">${k.split(' - ')[0]}</div>
        <div class="hmv" style="color:${col}">${pct(v.erosione)}</div>
        <div class="hms">Sc ${pct(v.sc)} + Tr ${pct(v.tr)}</div>
        <div class="hmbar"><div class="hmfill" style="width:${Math.min(100,v.erosione*100)}%;background:${col}"></div></div>
      </div>`);
  });

  // erosione bar
  doHBar('ch-eros',
    sorted.map(([k])=>trunc(k.split(' - ')[0],20)),
    sorted.map(([,v])=>v.erosione*100), null,
    sorted.map(([,v])=>v.erosione>0.8?'rgba(255,95,114,.75)':v.erosione>0.65?'rgba(255,181,71,.75)':'rgba(0,229,160,.65)')
  );

  // scatter
  dc('ch-scatter');
  const ctx = document.getElementById('ch-scatter').getContext('2d');
  charts['ch-scatter'] = new Chart(ctx, {
    type:'scatter',
    data: { datasets:[{
      data: sorted.map(([,v])=>({ x:v.sc*100, y:v.tr*100 })),
      backgroundColor: sorted.map(([,v])=>v.erosione>0.8?'rgba(255,95,114,.8)':v.erosione>0.65?'rgba(255,181,71,.8)':'rgba(0,229,160,.8)'),
      pointRadius:7
    }]},
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: {
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>`${sorted[ctx.dataIndex]?.[0]?.split(' - ')[0]||''}: Sc ${ctx.parsed.x.toFixed(1)}% + Tr ${ctx.parsed.y.toFixed(1)}%`}}
      },
      scales: {
        x:{title:{display:true,text:'Sconto %',color:'#3d5a7a'},grid:{color:'rgba(255,255,255,.05)'},ticks:{color:'#3d5a7a',callback:v=>v+'%'}},
        y:{title:{display:true,text:'Incidenza Trasp %',color:'#3d5a7a'},grid:{color:'rgba(255,255,255,.05)'},ticks:{color:'#3d5a7a',callback:v=>v+'%'}}
      }
    }
  });

  tbl('tbl-marg',
    ['Categoria','Erosione Tot','Sconto %','Trasp %','Fatturato','N° Righe'],
    sorted.map(([k,v])=>[
      trunc(k,30),
      `<span class="bdg ${v.erosione>0.8?'br':v.erosione>0.65?'ba':'bg'}">${pct(v.erosione)}</span>`,
      `<span class="mo">${pct(v.sc)}</span>`,
      `<span class="mo">${pct(v.tr)}</span>`,
      `<span class="mo">${fmt(v.f)}</span>`,
      v.n
    ])
  );
}

// ── TRASPORTI ─────────────────────────────────────────
function renderTrasporti() {
  const V = F.vend;
  const anni = G.anni;
  const fattAn = groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const trAn = groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const inc = anni.map(a=>(trAn[a]||0)/(fattAn[a]||1));

  const trTot = sum(V,r=>r.trasp);
  const fTot = sum(V,r=>r.importo);
  const incMed = fTot>0?trTot/fTot:0;
  const picco = anni.reduce((mx,a)=>inc[anni.indexOf(a)]>inc[anni.indexOf(mx)]?a:mx, anni[0]);

  kpi('kr-tr',[
    {l:'Spese Trasporto Periodo',v:fmt(trTot),cls:'p',sub:`${pct(incMed)} su fatturato`},
    {l:'Incidenza Media',v:pct(incMed),cls:incMed>0.05?'r':'g',sub:'% su fatturato'},
    {l:'Anno Picco',v:picco,cls:'a',sub:`${pct(inc[anni.indexOf(picco)])} incidenza`},
    {l:'Ordini — Trasp. Pianif.',v:fmt(sum(G.ORDINI,r=>r.trasp)),cls:'p',sub:`${G.ORDINI.length} ordini aperti`},
  ]);

  lineBar('ch-tr-yr', anni, anni.map(a=>trAn[a]||0), anni.map(a=>inc[anni.indexOf(a)]*100),
    'Trasporto €','Incidenza %','rgba(180,122,255,.7)','#ffb547', true);

  // porto
  const portD={};
  V.forEach(r=>{ const p=PORTO_MAP[r.porto]||`Porto ${r.porto}`; portD[p]=(portD[p]||0)+1; });
  doPie('ch-porto', Object.keys(portD), Object.values(portD));

  tbl('tbl-tr',
    ['Anno','Fatturato','Trasporto','Incidenza','Δ'],
    anni.map((a,i)=>{
      const dlt=i>0?inc[i]-inc[i-1]:null;
      return [a, `<span class="mo">${fmt(fattAn[a]||0)}</span>`, `<span class="mo">${fmt(trAn[a]||0)}</span>`,
        `<span class="bdg ${inc[i]>0.05?'br':'bg'}">${pct(inc[i])}</span>`,
        dlt!==null?`<span class="bdg ${dlt>0?'br':'bg'}">${dlt>0?'+':''}${pct(dlt)}</span>`:'—'];
    })
  );
}

// ── ORDINI ────────────────────────────────────────────
let ordiniAllData = [];
function renderOrdini() {
  const O = G.ORDINI;
  const today = new Date();
  const scaduti = O.filter(r=>r.consegna&&r.consegna<today);
  const in30 = O.filter(r=>{ if(!r.consegna||r.consegna<today) return false; return (r.consegna-today)/86400000<=30; });
  ordiniAllData = [...O].sort((a,b)=>(a.consegna||new Date(9999,0))-(b.consegna||new Date(9999,0)));

  kpi('kr-ord',[
    {l:'Valore Inevaso',v:fmt(sum(O,r=>r.importoI)),cls:'a',sub:`${O.length} righe`},
    {l:'Clienti con Ordini',v:[...new Set(O.map(r=>r.cliente).filter(Boolean))].length,cls:'b'},
    {l:'Scaduti',v:scaduti.length,cls:'r',sub:fmt(sum(scaduti,r=>r.importoI))},
    {l:'In Scadenza 30gg',v:in30.length,cls:'a',sub:fmt(sum(in30,r=>r.importoI))},
    {l:'Trasp. Pianificato',v:fmt(sum(O,r=>r.trasp)),cls:'p'},
  ]);

  if (scaduti.length>0) document.getElementById('ord-scad').textContent=`${scaduti.length} SCADUTI`;

  const cliOrd = groupBy(O,r=>r.cliente,rows=>sum(rows,r=>r.importoI));
  const top12 = Object.entries(cliOrd).sort((a,b)=>b[1]-a[1]).slice(0,12);
  doHBar('ch-ord-cli', top12.map(([k])=>trunc(k,22)), top12.map(([,v])=>v), '#ffb547');

  const dateOrd = groupBy(O.filter(r=>r.consegna&&!isNaN(r.consegna)), r=>r.consegna.toISOString().split('T')[0], rows=>sum(rows,r=>r.importoI));
  const dateSorted = Object.entries(dateOrd).sort(([a],[b])=>a.localeCompare(b)).slice(0,20);
  doBar('ch-ord-date', dateSorted.map(([k])=>k.slice(5)), dateSorted.map(([,v])=>v),
    null, dateSorted.map(([k])=>new Date(k)<today?'rgba(255,95,114,.7)':'rgba(77,184,255,.7)'));

  filterOrdTbl();
}

function filterOrdTbl() {
  const q = document.getElementById('ord-srch').value.toLowerCase();
  const today = new Date();
  const rows = ordiniAllData.filter(r => (r.cliente+r.desc).toLowerCase().includes(q)).slice(0,80);
  tbl('tbl-ord',
    ['Cliente','Prodotto','Qtà','Importo','Consegna','Stato'],
    rows.map(r=>{
      const late = r.consegna && r.consegna < today;
      const soon = r.consegna && !late && (r.consegna-today)/86400000<=30;
      const stato = late?'<span class="bdg br">SCADUTO</span>':soon?'<span class="bdg ba">30gg</span>':'<span class="bdg bb">OK</span>';
      return [
        trunc(r.cliente,25), trunc(r.desc,28), r.qtyI,
        `<span class="mo">${fmt(r.importoI)}</span>`,
        r.consegna&&!isNaN(r.consegna)?r.consegna.toLocaleDateString('it'):'—',
        stato
      ];
    })
  );
}

// ── CRITICITÀ ─────────────────────────────────────────
function renderCriticita() {
  const V = G.VEND; const O = G.ORDINI;
  const sR = V.filter(r=>r.sconto!==null);
  const over60 = sR.filter(r=>r.sconto>SCONTO_MAX);
  const sMedia = sR.length?avg(sR,r=>r.sconto):0;
  const anni = G.anni; const annoMax = Math.max(...anni);
  const fAnnuo = groupBy(V,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAnnuo = groupBy(V,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const today = new Date();
  const scaduti = O.filter(r=>r.consegna&&r.consegna<today);
  const in30 = O.filter(r=>{ if(!r.consegna||r.consegna<today)return false; return (r.consegna-today)/86400000<=30; });

  const alerts = [];

  // 1. Over 60
  if (over60.length) {
    const pOver = sR.length?over60.length/sR.length:0;
    alerts.push({ type:pOver>0.1?'danger':'warn', icon:'🏷️',
      t:`Sconti >60%: ${over60.length} righe (${pct(pOver)}) — valore €${fmt(sum(over60,r=>r.importo))}`,
      b:`Verificare se trattasi di promozioni Autopromotec o azioni speciali autorizzate. I prodotti più colpiti: ${[...new Set(over60.sort((a,b)=>b.sconto-a.sconto).slice(0,3).map(r=>r.desc))].join(', ')}.`
    });
  }

  // 2. Sconto medio globale
  const diffS = sMedia - SCONTO_MAX;
  alerts.push({ type:diffS>0?'warn':'ok', icon:diffS>0?'⚠️':'✅',
    t:`Sconto medio globale: ${pct(sMedia)} (soglia: ${pct(SCONTO_MAX)})`,
    b:diffS>0?`Supera la soglia contrattuale di ${pct(Math.abs(diffS))}. Consigliare revisione policy con la rete vendita.`:`Dentro i limiti contrattuali (margine residuo: ${pct(Math.abs(diffS))}).`
  });

  // 3. Trend fatturato
  const f25=fAnnuo[2025]||0, f24=fAnnuo[2024]||0;
  if (f24&&f25) {
    const d=(f25-f24)/f24;
    alerts.push({ type:d<-0.05?'warn':d>0.05?'ok':'info', icon:d>0?'📈':'📉',
      t:`Trend 2025 vs 2024: ${d>0?'+':''}${pct(d)}`,
      b:`2024: €${fmt(f24)} → 2025: €${fmt(f25)}. ${d<0?'Attenzione alla contrazione: analizzare categorie e clienti in calo.':'Crescita positiva confermata.'}`
    });
  }

  // 4. Picco trasporti 2022
  const incidenze = anni.map(a=>({a,inc:(tAnnuo[a]||0)/(fAnnuo[a]||1)}));
  const picco = incidenze.reduce((m,i)=>i.inc>m.inc?i:m,incidenze[0]);
  if (picco.inc>0.07) alerts.push({ type:'warn', icon:'🚚',
    t:`Picco trasporti ${picco.a}: ${pct(picco.inc)} di incidenza (€${fmt(tAnnuo[picco.a]||0)})`,
    b:`Analizzare i fattori scatenanti (carburante, volumi, nuove zone) per prevenire anomalie future.`
  });

  // 5. Ordini scaduti
  if (scaduti.length) alerts.push({ type:'danger', icon:'⏰',
    t:`Ordini con data consegna scaduta: ${scaduti.length} righe — €${fmt(sum(scaduti,r=>r.importoI))}`,
    b:`Clienti: ${[...new Set(scaduti.map(r=>r.cliente).filter(Boolean))].slice(0,4).join(', ')}. Aggiornare date o contattare i clienti.`
  });

  // 6. In scadenza 30gg
  if (in30.length) alerts.push({ type:'warn', icon:'📅',
    t:`${in30.length} righe in scadenza entro 30 giorni — €${fmt(sum(in30,r=>r.importoI))}`,
    b:'Pianificare priorità logistica per l\'evasione nei tempi previsti.'
  });

  // 7. Concentrazione clienti
  const cli2425 = groupBy(V.filter(r=>r.anno>=2024), r=>r.cliente, rows=>sum(rows,r=>r.importo));
  const fTot2425 = sum(V.filter(r=>r.anno>=2024),r=>r.importo);
  const top3 = Object.entries(cli2425).sort((a,b)=>b[1]-a[1]).slice(0,3);
  const top3pct = fTot2425>0?sum(top3,([,v])=>v)/fTot2425:0;
  if (top3pct>0.35) alerts.push({ type:'warn', icon:'🏢',
    t:`Alta concentrazione: top 3 clienti = ${pct(top3pct)} del fatturato 2024–${annoMax}`,
    b:top3.map(([k,v])=>`${trunc(k,20)} ${pct(fTot2425>0?v/fTot2425:0)}`).join(' · ')+'. Rischio dipendenza — diversificare il portafoglio.'
  });

  // 8. Resi
  const nResi = G.RESI.length;
  const resiPct = nResi/(V.length+nResi);
  if (resiPct>0.02) alerts.push({ type:'info', icon:'↩️',
    t:`Resi: ${nResi} righe (${pct(resiPct)} sul totale movimenti)`,
    b:'Monitorare motivazioni per categoria. Tasso >5% può indicare problemi qualitativi o logistici.'
  });

  // update badge
  const dangerCount = alerts.filter(a=>a.type==='danger'||a.type==='warn').length;
  document.getElementById('nbadge').textContent = dangerCount;

  const el = document.getElementById('alerts');
  el.innerHTML = alerts.length===0
    ? '<div class="al ok"><div class="al-ic">✅</div><div class="al-b"><strong>Nessuna criticità rilevata</strong><p>Tutti gli indicatori sono nei range normali.</p></div></div>'
    : alerts.map(a=>`
      <div class="al ${a.type}">
        <div class="al-ic">${a.icon}</div>
        <div class="al-b"><strong>${a.t}</strong><p>${a.b}</p></div>
      </div>`).join('');
}

// ═══════════════════════════════════════════════════
//  CHART FACTORY
// ═══════════════════════════════════════════════════
function chartOpts({ legend=false, callbackY=null, y2=false }={}) {
  const yTick = callbackY || (v=>fmtShort(v));
  return {
    responsive:true, maintainAspectRatio:false,
    plugins: {
      legend: legend ? { display:true, labels:{ color:'#7a9cc0', font:{size:10,family:'Syne'}, boxWidth:10, padding:8 } } : { display:false },
      tooltip: { backgroundColor:'#141e30', borderColor:'#1e2d47', borderWidth:1, titleColor:'#e2eaf8', bodyColor:'#7a9cc0', padding:10 }
    },
    scales: {
      x:{ grid:{color:'rgba(255,255,255,.04)'}, ticks:{color:'#3d5a7a',font:{size:9,family:'Syne'},maxRotation:45} },
      y:{ grid:{color:'rgba(255,255,255,.04)'}, ticks:{color:'#3d5a7a',font:{size:9,family:'Syne'},callback:yTick} }
    }
  };
}

function lineBar(id, labels, barData, lineData, bLbl, lLbl, bCol, lCol, secondAxis=false) {
  dc(id);
  const ctx = document.getElementById(id).getContext('2d');
  const scales = secondAxis ? {
    x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#3d5a7a',font:{size:9,family:'Syne'}}},
    y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#3d5a7a',font:{size:9},callback:v=>fmtShort(v)}},
    y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:'#3d5a7a',font:{size:9},callback:v=>v.toFixed(1)+'%'}}
  } : {
    x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#3d5a7a',font:{size:9,family:'Syne'}}},
    y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#3d5a7a',font:{size:9},callback:v=>fmtShort(v)}}
  };
  charts[id] = new Chart(ctx, {
    data: { labels, datasets:[
      { type:'bar', label:bLbl, data:barData, backgroundColor:bCol||'rgba(0,229,160,.7)', borderRadius:4, yAxisID:'y' },
      { type:'line', label:lLbl, data:lineData, borderColor:lCol||'#ff5f72', tension:.3, pointRadius:4, fill:false, yAxisID:secondAxis?'y2':'y' }
    ]},
    options: { responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:true,labels:{color:'#7a9cc0',font:{size:10},boxWidth:10,padding:8}},
        tooltip:{backgroundColor:'#141e30',borderColor:'#1e2d47',borderWidth:1,titleColor:'#e2eaf8',bodyColor:'#7a9cc0'}},
      scales }
  });
}

function doPie(id, labels, data) {
  dc(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, { type:'doughnut',
    data:{ labels, datasets:[{ data, backgroundColor:PAL, borderWidth:0, hoverOffset:6 }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'right',labels:{color:'#7a9cc0',font:{size:9,family:'Syne'},boxWidth:8,padding:6}},
        tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${fmt(ctx.raw)}`}} } }
  });
}

function doBar(id, labels, data, color, colors) {
  dc(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, { type:'bar',
    data:{ labels, datasets:[{ data, backgroundColor:colors||color||'rgba(77,184,255,.7)', borderRadius:3 }]},
    options: chartOpts({ callbackY:v=>fmtShort(v) })
  });
}

function doHBar(id, labels, data, color, colors) {
  dc(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, { type:'bar',
    data:{ labels, datasets:[{ data, backgroundColor:colors||color||'rgba(0,229,160,.7)', borderRadius:3 }]},
    options: { ...chartOpts({ callbackY:v=>fmtShort(v) }), indexAxis:'y' }
  });
}

function dc(id) { if(charts[id]){ charts[id].destroy(); delete charts[id]; } }

// ═══════════════════════════════════════════════════
//  TABLE ENGINE (sortable)
// ═══════════════════════════════════════════════════
function tbl(id, headers, rows) {
  const el = document.getElementById(id); if(!el) return;
  const state = sortState[id] || { col:-1, asc:true };
  let sortedRows = [...rows];
  if (state.col >= 0) {
    sortedRows.sort((a,b) => {
      const va = stripHtml(a[state.col]), vb = stripHtml(b[state.col]);
      const na = parseFloat(va.replace(/[€%.,]/g,'')), nb = parseFloat(vb.replace(/[€%.,]/g,''));
      let cmp = !isNaN(na)&&!isNaN(nb) ? na-nb : va.localeCompare(vb);
      return state.asc ? cmp : -cmp;
    });
  }
  el.innerHTML = `
    <thead><tr>${headers.map((h,i)=>`<th class="${state.col===i?(state.asc?'sa':'sd'):''}" onclick="sortTbl('${id}',${i})">${h}</th>`).join('')}</tr></thead>
    <tbody>${sortedRows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>`;
}

function sortTbl(id, col) {
  const s = sortState[id] || { col:-1, asc:true };
  sortState[id] = { col, asc: s.col===col ? !s.asc : true };
  // re-render — find current panel and call appropriate render
  const panelMap = {
    'tbl-cat':'vendite','tbl-drill':'vendite','tbl-cli':'clienti',
    'tbl-sc':'sconti','tbl-over60-cli':'sconti','tbl-marg':'margine',
    'tbl-tr':'trasporti','tbl-ord':'ordini'
  };
  if (panelMap[id]==='vendite') renderVendite();
  else if (panelMap[id]==='clienti') filterCliTbl();
  else if (panelMap[id]==='sconti') renderScontiTbl();
  else if (panelMap[id]==='margine') renderMargine();
  else if (panelMap[id]==='trasporti') renderTrasporti();
  else if (panelMap[id]==='ordini') filterOrdTbl();
}
function stripHtml(s) { return String(s).replace(/<[^>]+>/g,'').trim(); }

// ═══════════════════════════════════════════════════
//  UI HELPERS
// ═══════════════════════════════════════════════════
function kpi(id, items) {
  document.getElementById(id).innerHTML = items.map(i => `
    <div class="kk ${i.cls||'g'}">
      <div class="kl">${i.l}</div>
      <div class="kv">${i.v}</div>
      <div class="ka">
        ${i.sub?`<span class="ks">${i.sub}</span>`:''}
        ${i.delta!==undefined&&i.delta!==null?`<span class="dt ${i.delta>=0?'up':'dn'}">${i.delta>=0?'↑':'↓'}${pct(Math.abs(i.delta))}</span>`:''}
      </div>
    </div>`).join('');
}

function go(name) {
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.ni').forEach(n=>n.classList.remove('on'));
  document.getElementById('panel-'+name).classList.add('on');
  document.querySelectorAll('.ni').forEach(n=>{ if(n.getAttribute('onclick')===`go('${name}')`) n.classList.add('on'); });
  document.getElementById('main').scrollTop=0;
}

function showLoad(m) { document.getElementById('loading').style.display='flex'; setLoad(m); }
function setLoad(m,s) { document.getElementById('lmsg').textContent=m; if(s) document.getElementById('lsub').textContent=s; }
function hideLoad() { document.getElementById('loading').style.display='none'; }
function resetApp() { location.reload(); }
function sleep(ms) { return new Promise(r=>setTimeout(r,ms)); }

// ═══════════════════════════════════════════════════
//  MATH HELPERS
// ═══════════════════════════════════════════════════
function fmt(v) {
  if (v===null||v===undefined||isNaN(v)) return '—';
  return '€'+Number(v).toLocaleString('it',{minimumFractionDigits:0,maximumFractionDigits:0});
}
function fmtShort(v) {
  if (!v&&v!==0) return '—';
  if (Math.abs(v)>=1000000) return '€'+(v/1000000).toFixed(1)+'M';
  if (Math.abs(v)>=1000) return '€'+(v/1000).toFixed(0)+'k';
  return '€'+Math.round(v);
}
function pct(v) {
  if (v===null||v===undefined||isNaN(v)) return '—';
  return (Number(v)*100).toFixed(1)+'%';
}
function num(v) { return parseFloat(v)||0; }
function str(v) { return String(v||'').trim(); }
function sum(arr, fn) { return (arr||[]).reduce((a,r)=>a+(parseFloat(fn(r))||0),0); }
function avg(arr, fn) { if(!arr||!arr.length) return 0; return sum(arr,fn)/arr.length; }
function groupBy(arr, kFn, vFn) {
  const r={};
  (arr||[]).forEach(x=>{ const k=kFn(x); if(!r[k]) r[k]=[]; r[k].push(x); });
  if (vFn) Object.keys(r).forEach(k=>{ r[k]=vFn(r[k]); });
  return r;
}
function trunc(s, n) { return s&&s.length>n?s.slice(0,n-1)+'…':s||''; }
