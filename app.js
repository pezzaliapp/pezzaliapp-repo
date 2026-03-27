// ══════════════════════════════════════════════════
//  PezzaliApp — Cormach Dashboard v3
//  Font: DM Sans + JetBrains Mono
//  Theme: dark/light toggle
// ══════════════════════════════════════════════════
'use strict';

const SCONTO_MAX = 0.60;
const MESI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
const QNAMES = ['Q1','Q2','Q3','Q4'];
const PORTO_MAP = {1:'Franco',2:'Assegnato',3:'Franco+Add.',6:'Altro'};

// Cascos brand detection via descrizione (C-series only)
const CASCOS_PATTERN = /\bC\s*3[.,]2|\bC\s*3[.,]5|\bC\s*4\s*(XL|s|S|\b)|\bC\s*5\s*(XL|s|S|[.,]5|\b)|\bC\s*5[.,]5|\bC\s*7s\b|\bC\s*4[3-9]\d\b|\bC\s*4\d{2}\b|\bC\s*125\b|\bPARKING\s*2x1\b/i;
const CASCOS_EXCLUDE = /PFA\s*\d+|STD\s*72|FORBICE|SCHEDA|CENTRALINA|POMPA|MICRO\s*FIN|FOTOCELLULA|PERNO|BRACCIO\s*(SX|DX)|USATO|CARTER|TASTATORE/i;

// ── theme-aware chart colors ──
function tc() {
  const dark = document.documentElement.getAttribute('data-theme') !== 'light';
  return {
    green:  dark ? '#00b894' : '#059669',
    red:    dark ? '#e17055' : '#dc2626',
    amber:  dark ? '#fdcb6e' : '#d97706',
    blue:   dark ? '#74b9ff' : '#2563eb',
    purple: dark ? '#a29bfe' : '#7c3aed',
    cyan:   dark ? '#00cec9' : '#0891b2',
    text2:  dark ? '#94a3b8' : '#475569',
    text3:  dark ? '#4a5568' : '#94a3b8',
    border: dark ? 'rgba(255,255,255,.06)' : 'rgba(0,0,0,.06)',
    bg1:    dark ? '#1f2937' : '#ffffff',
    bg2:    dark ? '#1a2233' : '#f8fafc',
    tooltip:{ bg: dark ? '#1a2233' : '#ffffff', border: dark ? 'rgba(255,255,255,.1)' : 'rgba(0,0,0,.1)' }
  };
}

let G = {}, F = {}, CMP = false, charts = {}, sortState = {};

// ── FILE UPLOAD ──
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

// ── MAIN ──
async function runAnalysis() {
  showLoad('Lettura vendite...'); await sleep(30);
  try {
    const vRaw = await readXLSX('fi-v');
    setLoad('Lettura ordini...'); await sleep(20);
    const oRaw = await readXLSX('fi-o');
    setLoad('Caricamento listino...'); await sleep(20);
    let lRaw = null;
    const lInp = document.getElementById('fi-l');
    if (lInp.files[0]) {
      lRaw = lInp.files[0].name.endsWith('.csv') ? await readCSV('fi-l') : await readXLSX('fi-l');
    }
    setLoad('Elaborazione dati...'); await sleep(40);
    processData(vRaw, oRaw, lRaw);
    setLoad('Rendering grafici...'); await sleep(30);
    initDashboard();
    hideLoad();
    document.getElementById('upload-screen').style.display = 'none';
    document.getElementById('top-filters').removeAttribute('hidden');
    document.getElementById('btn-reset').removeAttribute('hidden');
    document.getElementById('btn-print').removeAttribute('hidden');
    document.getElementById('status-pill').textContent = 'ATTIVO';
  } catch(err) {
    hideLoad(); alert('Errore: ' + err.message); console.error(err);
  }
}

function readXLSX(id) {
  return new Promise((res, rej) => {
    const f = document.getElementById(id).files[0]; if (!f) return rej(new Error('File mancante'));
    const r = new FileReader();
    r.onload = e => { try { const wb = XLSX.read(e.target.result, {type:'array', cellDates:true}); res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''})); } catch(e){ rej(e); } };
    r.onerror = () => rej(new Error('Lettura fallita')); r.readAsArrayBuffer(f);
  });
}
function readCSV(id) {
  return new Promise((res, rej) => Papa.parse(document.getElementById(id).files[0], {header:true,skipEmptyLines:true,dynamicTyping:true,complete:r=>res(r.data),error:e=>rej(e)}));
}

// ── PROCESS DATA ──
function processData(vRaw, oRaw, lRaw) {
  // listino
  const listinoMap = {};
  if (lRaw) {
    lRaw.forEach(r => {
      const cod = str(r.Codice||r.codice||r.CODICE||'').replace(/^0+/,'');
      const pl = num(r.PrezzoLordo||r.prezzo_lordo||0);
      if (cod && pl > 0) listinoMap[cod] = { pl, inst: num(r.CostoInstallazione||0), trasp: num(r.CostoTrasporto||0) };
    });
  }

  const keys = Object.keys(vRaw[0]||{});
  const col = cs => cs.find(c=>keys.includes(c))||cs[0];
  const CV = {
    anno:    col(['ANNO SPEDIZIONE','ANNO']),
    data:    col(['DATA SPEDIZIONE','DATA']),
    importo: col(['IMPORTO CONSEGNATO','IMPORTO']),
    pz:      col(['PZ NETTO VENDITA','PZ NETTO']),
    qty:     col(['QTA CONSEGNATA','QTA']),
    trasp:   col(['SPESE DI TRASPORTO','TRASPORTO']),
    causale: col(['CAUSALE MAGAZZINO','CAUSALE']),
    cat:     col(['DESCRIZIONE ELEMENTO.5','CATEGORIA']),
    sottocat:col(['DESCRIZIONE ELEMENTO.4','SOTTOCATEGORIA']),
    agente:  col(['DESCRIZIONE ELEMENTO.2','DESCRIZIONE ELEMENTO.6']),
    cliente: col(['RAGIONE SOCIALE 1','RAGIONE SOCIALE','CLIENTE']),
    dest:    col(['DESC. DESTINAZIONE 1','DESTINAZIONE']),
    articolo:col(['ARTICOLO']),
    desc:    col(['DESCRIZIONE']),
    porto:   col(['PORTO']),
    zona:    col(['ZONA']),
    nazione: col(['NAZIONE']),
    regione: col(['DESCRIZIONE ELEMENTO.3','REGIONE']),
    citta:   col(['DESCRIZIONE ELEMENTO.1','CITTA']),
  };

  const VEND = vRaw.filter(r => str(r[CV.causale]).toUpperCase().startsWith('V')).map(r => {
    const cod = str(r[CV.articolo]).replace(/^0+/,'');
    const li = listinoMap[cod]||null;
    const pz = num(r[CV.pz]), importo = num(r[CV.importo]), trasp = num(r[CV.trasp]);
    const lordo = li?li.pl:null;
    const sconto = lordo&&lordo>0&&pz>0 ? Math.max(0,Math.min(1,1-pz/lordo)) : null;
    const anno = parseInt(r[CV.anno])||0;
    const desc = str(r[CV.desc]);
    const brand = (CASCOS_PATTERN.test(desc)&&!CASCOS_EXCLUDE.test(desc)) ? 'Cascos' : (li?'Cormach':'__no');
    let date=null,mese=-1,trim=-1;
    try{ const d=r[CV.data]; date=d instanceof Date?d:new Date(d); if(!isNaN(date)){mese=date.getMonth();trim=Math.floor(mese/3);} }catch(e){}
    const PORTO_LABELS={1:'Franco',2:'Assegnato',3:'Franco+Add.',6:'Altro',29:'Altro',168:'Altro'};
    const portoRaw=r[CV.porto];
    const porto_desc=PORTO_LABELS[portoRaw]||'Altro';
    const sconto_eff = sconto!==null ? sconto : 0.60; // fallback 60% soglia max
    return {
      anno,date,mese,trim,importo,pz,qty:num(r[CV.qty]),trasp,
      cat: (_cat==='nan'||_cat===''||!isNaN(Number(_cat))) ? '' : _cat,
      sottocat:str(r[CV.sottocat]||''),
      agente: (_agente===''||_agente==='nan'||!isNaN(Number(_agente))) ? '' : _agente,
      cliente:str(r[CV.cliente]),dest:str(r[CV.dest]),
      nazione:str(r[CV.nazione]),
      regione:str(r[CV.regione]),citta:str(r[CV.citta]),
      articolo:str(r[CV.articolo]),desc,lordo,sconto,sconto_eff,brand,
      incTrasp:importo>0?trasp/importo:0,
      porto_desc,
    };
  }).filter(r=>r.anno>2000);

  const RESI = vRaw.filter(r=>['H1','X2','X4','X5'].includes(str(r[CV.causale]).toUpperCase()));

  // ordini
  const OK = Object.keys(oRaw[0]||{});
  const CO = {
    cliente: ['CLIENTE.1','CLIENTE','RAGIONE SOCIALE'].find(c=>OK.includes(c)),
    desc:    ['DESCRIZIONE'].find(c=>OK.includes(c)),
    qtyI:    ['QTA INEVASA'].find(c=>OK.includes(c)),
    importoI:['IMPORTO INEVASO'].find(c=>OK.includes(c)),
    trasp:   ['SPESE DI TRASPORTO'].find(c=>OK.includes(c)),
    consegna:['DATA CONSEGNA'].find(c=>OK.includes(c)),
    porto:   ['PORTO'].find(c=>OK.includes(c)),
    num:     ['NUM.'].find(c=>OK.includes(c)),
    cat3:    ['CLASSE 3 ARTICOLO'].find(c=>OK.includes(c)),
  };
  const ORDINI = oRaw.map(r => {
    let consegna=null;
    try{const d=r[CO.consegna];consegna=d instanceof Date?d:new Date(d);if(isNaN(consegna))consegna=null;}catch(e){}
    return {
      cliente:str(r[CO.cliente]),desc:str(r[CO.desc]),
      qtyI:num(r[CO.qtyI]),importoI:num(r[CO.importoI]),
      trasp:num(r[CO.trasp]),consegna,porto:r[CO.porto],num:str(r[CO.num]),
    };
  });

  const anni = [...new Set(VEND.map(r=>r.anno))].sort();
  const agenti = [...new Set(VEND.map(r=>r.agente).filter(Boolean))].sort();

  // populate filter selects
  const selDa=document.getElementById('f-da'), selA=document.getElementById('f-a');
  const selCmp=document.getElementById('f-cmp-a');
  const selT1=document.getElementById('tr-a1'), selT2=document.getElementById('tr-a2');
  [selDa,selA,selCmp,selT1,selT2].forEach(s=>s.innerHTML='');
  anni.forEach(a=>[selDa,selA,selCmp,selT1,selT2].forEach(s=>s.insertAdjacentHTML('beforeend',`<option value="${a}">${a}</option>`)));
  selDa.value=anni[0]; selA.value=anni[anni.length-1];
  selCmp.value=anni[anni.length-2]||anni[0];
  selT1.value=anni[anni.length-2]||anni[0]; selT2.value=anni[anni.length-1];

  const fagt=document.getElementById('f-agt');
  fagt.innerHTML='<option value="">Tutti</option>';
  agenti.forEach(a=>fagt.insertAdjacentHTML('beforeend',`<option value="${a}">${a}</option>`));

  // build client list for storico select
  const cliRank = groupBy(VEND, r=>r.cliente, rows=>sum(rows,r=>r.importo));
  const topCli = Object.entries(cliRank).sort((a,b)=>b[1]-a[1]).slice(0,60).map(([k])=>k);
  const cliSel=document.getElementById('cli-sel');
  cliSel.innerHTML='';
  topCli.forEach(c=>cliSel.insertAdjacentHTML('beforeend',`<option value="${c}">${c}</option>`));

  G = { VEND, RESI, ORDINI, anni, agenti, listinoMap, CV };
  applyFilters();
}

// ── FILTER ENGINE ──
function applyFilters() {
  const annoDa = parseInt(document.getElementById('f-da').value)||0;
  const annoA  = parseInt(document.getElementById('f-a').value)||9999;
  const perVal = document.getElementById('f-per').value;
  const agente = document.getElementById('f-agt').value;
  const brandF = (document.getElementById('f-brand')||{}).value||'';
  let mesiOk=null;
  if(perVal.startsWith('q')){const q=parseInt(perVal[1])-1;mesiOk=[q*3,q*3+1,q*3+2];}
  else if(perVal.startsWith('m')){mesiOk=[parseInt(perVal.slice(1))-1];}

  const filt = r => {
    if(r.anno<annoDa||r.anno>annoA) return false;
    if(mesiOk&&!mesiOk.includes(r.mese)) return false;
    if(agente&&r.agente!==agente) return false;
    if(brandF){
      if(brandF==='__no'&&r.brand!=='__no') return false;
      if(brandF!=='__no'&&r.brand.toLowerCase()!==brandF.toLowerCase()) return false;
    }
    return true;
  };
  F.vend = G.VEND.filter(filt);
  if(CMP){
    const cmpAnno=parseInt(document.getElementById('f-cmp-a').value)||0;
    F.vendCmp=G.VEND.filter(r=>{if(r.anno!==cmpAnno)return false;if(mesiOk&&!mesiOk.includes(r.mese))return false;if(agente&&r.agente!==agente)return false;if(brandF){if(brandF==='__no'&&r.brand!=='__no')return false;if(brandF!=='__no'&&r.brand.toLowerCase()!==brandF.toLowerCase())return false;}return true;});
  } else F.vendCmp=null;

  let plabel=`${annoDa}–${annoA}`;
  if(perVal.startsWith('q'))plabel+=` Q${perVal[1]}`;
  else if(perVal.startsWith('m'))plabel+=` ${MESI[parseInt(perVal.slice(1))-1]}`;
  if(agente)plabel+=` · ${agente}`;
  if(brandF&&brandF!=='__no')plabel+=` · ${brandF}`;
  if(brandF==='__no')plabel+=` · Senza listino`;
  F.label=plabel;
  document.getElementById('sb-period').textContent=plabel;
  renderAll();
}
function toggleCompare(){CMP=!CMP;document.getElementById('chip-cmp').classList.toggle('on',CMP);document.getElementById('f-cmp-a').style.display=CMP?'inline-block':'none';applyFilters();}

// ── RENDER ALL ──
function renderAll(){renderOverview();renderTrend();renderVendite();renderClienti();renderAgenti();renderSconti();renderMargine();renderTrasporti();renderOrdini();renderCriticita();}
function initDashboard(){renderAll();go('overview');}

// ══════════════════════════════════════════════════
//  PANEL RENDERERS
// ══════════════════════════════════════════════════

// ── OVERVIEW ──
function renderOverview(){
  const V=F.vend, C=tc();
  const fattTot=sum(V,r=>r.importo), traspTot=sum(V,r=>r.trasp);
  const aF=parseInt(document.getElementById('f-a').value);
  const fCur=sum(V.filter(r=>r.anno===aF),r=>r.importo);
  const fPrev=sum(V.filter(r=>r.anno===aF-1),r=>r.importo);
  const delta=fPrev>0?(fCur-fPrev)/fPrev:null;
  const sR=V.filter(r=>r.sconto!==null);
  const sMed=sR.length?avg(sR,r=>r.sconto):null;
  const over60=sR.filter(r=>r.sconto>SCONTO_MAX);
  document.getElementById('ov-sub').textContent=`${F.label} · ${V.length.toLocaleString('it')} righe vendita`;

  kpi('kr-ov',[
    {l:'Fatturato Periodo',v:fmt(fattTot),col:'g',sub:F.label},
    {l:`Fatturato ${aF}`,v:fmt(fCur),col:'g',delta,sub:`vs ${aF-1}`},
    {l:'Spese Trasporto',v:fmt(traspTot),col:'p',sub:pct(fattTot>0?traspTot/fattTot:0)+' del fatturato'},
    {l:'Sconto Medio',v:sMed!==null?pct(sMed):'N/D',col:sMed>SCONTO_MAX?'r':'g',sub:'su articoli in listino'},
    {l:'Righe >60%',v:over60.length.toLocaleString('it'),col:over60.length>0?'r':'g',sub:'oltre soglia contrattuale'},
    {l:'Ordini Inevasi',v:fmt(sum(G.ORDINI,r=>r.importoI)),col:'a',sub:`${G.ORDINI.length} righe aperte`},
  ]);

  // annual chart with transport line + peak annotation
  const anni=G.anni;
  const fAnnuo=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAnnuo=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const incArr=anni.map(a=>(tAnnuo[a]||0)/(fAnnuo[a]||1)*100);
  const peakAnno=anni[incArr.indexOf(Math.max(...incArr))];

  dc('ch-annual');
  const ctx=document.getElementById('ch-annual').getContext('2d');
  charts['ch-annual']=new Chart(ctx,{
    data:{labels:anni,datasets:[
      {type:'bar',label:'Fatturato €',data:anni.map(a=>fAnnuo[a]||0),backgroundColor:anni.map(a=>a===aF?C.green+'cc':C.green+'55'),borderRadius:4,yAxisID:'y'},
      {type:'line',label:'Incidenza Trasp %',data:incArr,borderColor:C.red,backgroundColor:C.red+'15',tension:.3,pointRadius:5,fill:false,yAxisID:'y2',
       pointBackgroundColor:anni.map(a=>a===peakAnno?C.red:C.red+'80'),
       pointRadius:anni.map(a=>a===peakAnno?7:4)}
    ]},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:true,labels:{color:C.text2,font:{size:10,family:'DM Sans'},boxWidth:10,padding:10}},
        tooltip:{backgroundColor:C.tooltip.bg,borderColor:C.tooltip.border,borderWidth:1,titleColor:C.text2,bodyColor:C.text2},
        annotation:{annotations:{
          peak:{type:'point',xValue:peakAnno,yValue:incArr[anni.indexOf(peakAnno)],yScaleID:'y2',
            backgroundColor:C.red,radius:6,borderColor:'#fff',borderWidth:2,
            label:{content:`Picco ${peakAnno}: ${incArr[anni.indexOf(peakAnno)].toFixed(1)}%`,enabled:true,position:'top',backgroundColor:C.red,color:'#fff',font:{size:9}}}
        }}
      },
      scales:{
        x:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9,family:'DM Sans'}}},
        y:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>fmtShort(v)}},
        y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:C.red,font:{size:9},callback:v=>v.toFixed(1)+'%'}}
      }
    }
  });

  // pie
  const catFatt=groupBy(V.filter(r=>r.cat&&r.cat!=='nan'&&r.cat.length>1),r=>r.cat,rows=>sum(rows,r=>r.importo));
  const catS=Object.entries(catFatt).sort((a,b)=>b[1]-a[1]).slice(0,8);
  doPie('ch-pie',catS.map(([k])=>k.split(' - ')[0]),catS.map(([,v])=>v));

  // quarterly
  const trimMap={};
  G.VEND.forEach(r=>{if(r.trim>=0){const k=`${r.anno}-Q${r.trim+1}`;trimMap[k]=(trimMap[k]||0)+r.importo;}});
  const tLbls=[],tData=[];
  G.anni.forEach(a=>QNAMES.forEach((_,q)=>{tLbls.push(`${a} Q${q+1}`);tData.push(trimMap[`${a}-Q${q+1}`]||0);}));
  doBar('ch-qtr',tLbls,tData,[C.blue+'aa'],null,{ticks:{maxRotation:45,autoSkip:true,maxTicksLimit:12}});

  // top clients
  const cliF=groupBy(V,r=>r.cliente,rows=>sum(rows,r=>r.importo));
  const top10=Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,10);
  doHBar('ch-top-cli',top10.map(([k])=>trunc(k,24)),top10.map(([,v])=>v),C.green);
}

// ── TREND ──
function renderTrend(){
  const C=tc();
  const a1=parseInt(document.getElementById('tr-a1').value);
  const a2=parseInt(document.getElementById('tr-a2').value);
  const view=document.getElementById('tr-view').value;
  const met=document.getElementById('tr-met').value;
  const labels=view==='m'?MESI:QNAMES;
  const metLbl={f:'Fatturato',t:'Trasporto €',s:'Sconto %',n:'N° Righe'}[met];

  const serie=anno=>labels.map((_,i)=>{
    const rows=G.VEND.filter(r=>r.anno===anno&&(view==='m'?r.mese===i:r.trim===i));
    if(met==='f')return sum(rows,r=>r.importo);
    if(met==='t')return sum(rows,r=>r.trasp);
    if(met==='s'){const sr=rows.filter(r=>r.sconto!==null);return sr.length?avg(sr,r=>r.sconto)*100:0;}
    return rows.length;
  });
  const d1=serie(a1),d2=serie(a2);
  const tot1=d1.reduce((a,b)=>a+b,0),tot2=d2.reduce((a,b)=>a+b,0);
  const dPct=tot1>0?(tot2-tot1)/tot1:null;

  document.getElementById('tr-title').textContent=`Confronto ${view==='m'?'Mensile':'Trimestrale'} — ${metLbl}`;
  document.getElementById('tr-sub').textContent=`${a1} vs ${a2}`;
  document.getElementById('tr-delta').innerHTML=dPct!==null
    ?`<span class="dt ${dPct>=0?'up':'dn'}">${dPct>=0?'↑':'↓'} ${pct(Math.abs(dPct))}</span>`:'';

  const fmtY=met==='f'||met==='t'?v=>fmtShort(v):met==='s'?v=>v.toFixed(1)+'%':v=>v;
  dc('ch-trend');
  const ctx=document.getElementById('ch-trend').getContext('2d');
  charts['ch-trend']=new Chart(ctx,{
    data:{labels,datasets:[
      {type:'bar',label:`${a1}`,data:d1,backgroundColor:C.blue+'66',borderRadius:3},
      {type:'line',label:`${a2}`,data:d2,borderColor:C.green,tension:.3,pointRadius:5,fill:false,
       pointBackgroundColor:d2.map((v,i)=>v>d1[i]?C.green:C.red),
       pointRadius:5}
    ]},
    options:{...chartOpts({callbackY:fmtY,legend:true,C})}
  });

  // cumulative
  const cum1=[],cum2=[];
  d1.reduce((acc,v,i)=>{cum1[i]=acc+v;return acc+v;},0);
  d2.reduce((acc,v,i)=>{cum2[i]=acc+v;return acc+v;},0);
  dc('ch-cumul');
  const ctx2=document.getElementById('ch-cumul').getContext('2d');
  charts['ch-cumul']=new Chart(ctx2,{
    data:{labels,datasets:[
      {type:'line',label:`${a1}`,data:cum1,borderColor:C.blue,fill:true,backgroundColor:C.blue+'15',tension:.3,pointRadius:3},
      {type:'line',label:`${a2}`,data:cum2,borderColor:C.green,fill:true,backgroundColor:C.green+'15',tension:.3,pointRadius:3}
    ]},
    options:{...chartOpts({legend:true,callbackY:v=>fmtShort(v),C})}
  });

  // delta %
  const deltas=d1.map((v,i)=>d1[i]>0?((d2[i]-d1[i])/d1[i])*100:0);
  dc('ch-delta');
  const ctx3=document.getElementById('ch-delta').getContext('2d');
  charts['ch-delta']=new Chart(ctx3,{type:'bar',
    data:{labels,datasets:[{data:deltas,backgroundColor:deltas.map(d=>d>=0?C.green+'aa':C.red+'aa'),borderRadius:3}]},
    options:{...chartOpts({callbackY:v=>v.toFixed(1)+'%',C})}
  });
}

// ── VENDITE ──
function renderVendite(){
  const V=F.vend, C=tc();
  const catFatt=groupBy(V.filter(r=>r.cat&&r.cat!=='nan'&&r.cat.length>1),r=>r.cat,rows=>({f:sum(rows,r=>r.importo),q:sum(rows,r=>r.qty),n:rows.length,sc:avg(rows.filter(r=>r.sconto!==null),r=>r.sconto)}));
  const cats=Object.entries(catFatt).sort((a,b)=>b[1].f-a[1].f);
  const totF=sum(V,r=>r.importo);

  kpi('kr-vend',[
    {l:'Fatturato',v:fmt(totF),col:'g'},
    {l:'Righe Vendita',v:V.length.toLocaleString('it'),col:'b'},
    {l:'Prezzo Netto Medio',v:fmt(avg(V,r=>r.pz)),col:'g',sub:'per unità'},
    {l:'Categorie Attive',v:cats.length,col:'p'},
    {l:'Ticket Medio Riga',v:fmt(V.length?totF/V.length:0),col:'b'},
  ]);

  const catColors=[C.green+'aa',C.blue+'aa',C.amber+'aa',C.purple+'aa',C.red+'aa',C.cyan+'aa',C.green+'66',C.blue+'66',C.amber+'66',C.purple+'66',C.red+'66',C.cyan+'66'];
  dc('ch-cat-bar');
  const ctx=document.getElementById('ch-cat-bar').getContext('2d');
  charts['ch-cat-bar']=new Chart(ctx,{type:'bar',
    data:{labels:cats.map(([k])=>trunc(k.split(' - ')[0],20)),datasets:[{data:cats.map(([,v])=>v.f),backgroundColor:catColors,borderRadius:4}]},
    options:{...chartOpts({callbackY:v=>fmtShort(v),C}),
      onClick:(_,els)=>{if(!els.length)return;showDrill(cats[els[0].index][0],V);}}
  });

  doBar('ch-cat-qty',cats.map(([k])=>trunc(k.split(' - ')[0],20)),cats.map(([,v])=>v.q),[C.purple+'aa'],null);

  tbl('tbl-cat',
    ['Categoria','Fatturato','%','Pezzi','Sconto Medio'],
    cats.map(([k,v])=>[trunc(k,32),`<span class="mono">${fmt(v.f)}</span>`,`<span class="bdg bg">${pct(totF>0?v.f/totF:0)}</span>`,`<span class="mono">${Math.round(v.q).toLocaleString('it')}</span>`,v.sc>0?`<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${pct(v.sc)}</span>`:'—'])
  );
}

function showDrill(catKey,V){
  document.getElementById('drill-cc').removeAttribute('hidden');
  document.getElementById('drill-t').textContent=`Clienti → ${catKey.split(' - ')[0]}`;
  const rows=V.filter(r=>r.cat===catKey);
  const cliF=groupBy(rows,r=>r.cliente,r=>sum(r,x=>x.importo));
  const sorted=Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,25);
  const totCat=sum(rows,r=>r.importo);
  tbl('tbl-drill',['Cliente','Fatturato','%'],
    sorted.map(([k,v])=>[k,`<span class="mono">${fmt(v)}</span>`,`<span class="bdg bg">${pct(totCat>0?v/totCat:0)}</span>`]));
}
function closeDrill(){document.getElementById('drill-cc').setAttribute('hidden','');}

// ── CLIENTI & DESTINAZIONI ──
let cliAllData=[];
function renderClienti(){
  const V=F.vend, C=tc();
  // group by client: fatturato + destinazioni uniche
  const cliData={};
  V.forEach(r=>{
    if(!cliData[r.cliente]) cliData[r.cliente]={f:0,n:0,dests:new Set(),sc:[],tr:0};
    cliData[r.cliente].f+=r.importo;
    cliData[r.cliente].n++;
    cliData[r.cliente].tr+=r.trasp;
    if(r.dest&&r.dest!==r.cliente&&r.dest.length>1) cliData[r.cliente].dests.add(r.dest);
    if(r.sconto!==null) cliData[r.cliente].sc.push(r.sconto);
  });
  cliAllData=Object.entries(cliData).map(([k,v])=>({
    nome:k,f:v.f,n:v.n,tr:v.tr,
    dests:[...v.dests],
    sc:v.sc.length?v.sc.reduce((a,b)=>a+b,0)/v.sc.length:null
  })).sort((a,b)=>b.f-a.f);

  const top12=cliAllData.slice(0,12);
  doHBar('ch-cli-rank',top12.map(c=>trunc(c.nome,24)),top12.map(c=>c.f),C.green);
  renderCliStorico();
  filterCliTbl();
}

function renderCliStorico(){
  const sel=document.getElementById('cli-sel').value; if(!sel) return;
  const C=tc();
  const rows=G.VEND.filter(r=>r.cliente===sel&&r.mese>=0);
  const anni=[...new Set(rows.map(r=>r.anno))].sort();
  const colors=[C.green,C.blue,C.amber,C.purple,C.red,C.cyan];
  const datasets=anni.map((a,i)=>({
    type:'line',label:`${a}`,
    data:MESI.map((_,m)=>rows.filter(r=>r.anno===a&&r.mese===m).reduce((s,r)=>s+r.importo,0)),
    borderColor:colors[i%colors.length],backgroundColor:colors[i%colors.length]+'15',
    tension:.3,pointRadius:3,fill:false
  }));
  dc('ch-cli-st');
  const ctx=document.getElementById('ch-cli-st').getContext('2d');
  charts['ch-cli-st']=new Chart(ctx,{data:{labels:MESI,datasets},options:{...chartOpts({legend:true,callbackY:v=>fmtShort(v),C})}});
}

function filterCliTbl(){
  const q=document.getElementById('cli-srch').value.toLowerCase();
  const destFilter=document.getElementById('cli-dest-filter').value;
  let rows=cliAllData.filter(c=>c.nome.toLowerCase().includes(q)||c.dests.some(d=>d.toLowerCase().includes(q)));
  if(destFilter==='multi') rows=rows.filter(c=>c.dests.length>0);
  if(destFilter==='single') rows=rows.filter(c=>c.dests.length===0);
  const totF=sum(F.vend,r=>r.importo);
  tbl('tbl-cli',
    ['Cliente','Fatturato','%','Destinazioni diverse','Sconto Medio','Trasporto'],
    rows.slice(0,80).map(c=>{
      const destPills=c.dests.length>0
        ?`<div class="dest-list">${c.dests.slice(0,4).map(d=>`<span class="dest-pill">${trunc(d,22)}</span>`).join('')}${c.dests.length>4?`<span class="dest-pill">+${c.dests.length-4}</span>`:''}</div>`
        :`<span class="mono" style="color:var(--text3)">sede unica</span>`;
      return [
        c.nome,
        `<span class="mono">${fmt(c.f)}</span>`,
        `<span class="bdg bg">${pct(totF>0?c.f/totF:0)}</span>`,
        destPills,
        c.sc!==null?`<span class="bdg ${c.sc>SCONTO_MAX?'br':c.sc>0.5?'ba':'bg'}">${pct(c.sc)}</span>`:'—',
        `<span class="mono">${fmt(c.tr)}</span>`
      ];
    })
  );
}

// ── AGENTI ──
function renderAgenti(){
  const V=F.vend, C=tc();
  const agtF=groupBy(V.filter(r=>r.agente),r=>r.agente,rows=>({f:sum(rows,r=>r.importo),n:rows.length,tr:sum(rows,r=>r.trasp)}));
  const sorted=Object.entries(agtF).sort((a,b)=>b[1].f-a[1].f);
  const totF=sum(V,r=>r.importo);

  kpi('kr-agt',sorted.slice(0,4).map(([k,v],i)=>({l:`#${i+1} ${k}`,v:fmt(v.f),col:['g','b','a','p'][i]||'p',sub:`${pct(totF>0?v.f/totF:0)} del totale`})));

  const colors=[C.green+'bb',C.blue+'bb',C.amber+'bb',C.purple+'bb',C.red+'bb',C.cyan+'bb'];
  doHBar('ch-agt-bar',sorted.map(([k])=>k),sorted.map(([,v])=>v.f),null,sorted.map((_,i)=>colors[i%colors.length]));

  const anni=G.anni;
  const datasets=sorted.map(([k],i)=>({
    type:'line',label:k,
    data:anni.map(a=>sum(G.VEND.filter(r=>r.anno===a&&r.agente===k),r=>r.importo)),
    borderColor:colors[i%colors.length].replace('bb',''),tension:.3,pointRadius:3,fill:false
  }));
  dc('ch-agt-evol');
  const ctx=document.getElementById('ch-agt-evol').getContext('2d');
  charts['ch-agt-evol']=new Chart(ctx,{data:{labels:anni,datasets},options:{...chartOpts({legend:true,callbackY:v=>fmtShort(v),C})}});

  const agEl=document.getElementById('agr'); agEl.innerHTML='';
  const maxF=sorted[0]?.[1]?.f||1;
  sorted.forEach(([k,v],i)=>{
    agEl.insertAdjacentHTML('beforeend',`
      <div class="agrow">
        <div class="agrow-top">
          <span class="agno">${String(i+1).padStart(2,'0')}</span>
          <span class="agname">${k}</span>
          <span class="agval">${fmt(v.f)}</span>
          <span class="bdg bg" style="font-size:8px">${pct(totF>0?v.f/totF:0)}</span>
        </div>
        <div class="agbar"><div class="agfill" style="width:${Math.round(v.f/maxF*100)}%"></div></div>
      </div>`);
  });
}

// ── SCONTI ──
let scontiData=[];
function renderSconti(){
  const V=F.vend, C=tc();
  const sR=V.filter(r=>r.sconto!==null);
  const sMed=sR.length?avg(sR,r=>r.sconto):0;
  const over60=sR.filter(r=>r.sconto>SCONTO_MAX);

  kpi('kr-sc',[
    {l:'Sconto Medio Periodo',v:pct(sMed),col:sMed>SCONTO_MAX?'r':'g',sub:`su ${sR.length} righe con listino`},
    {l:'Righe >60%',v:over60.length.toLocaleString('it'),col:'r',sub:fmt(sum(over60,r=>r.importo))},
    {l:'Soglia Contrattuale',v:pct(SCONTO_MAX),col:'a',sub:'max rivenditori Cormach'},
    {l:'Copertura Listino',v:pct(V.length?sR.length/V.length:0),col:'b',sub:`${sR.length} righe su ${V.length}`},
  ]);

  const anni=G.anni;
  const scAnno=anni.map(a=>{const r2=G.VEND.filter(r=>r.anno===a&&r.sconto!==null);return r2.length?avg(r2,r=>r.sconto)*100:0;});
  dc('ch-sc-yr');
  const ctx=document.getElementById('ch-sc-yr').getContext('2d');
  charts['ch-sc-yr']=new Chart(ctx,{
    data:{labels:anni,datasets:[
      {type:'line',label:'Sconto %',data:scAnno,borderColor:C.purple,backgroundColor:C.purple+'15',tension:.3,fill:true,pointRadius:5,
       pointBackgroundColor:scAnno.map(v=>v>60?C.red:C.purple),pointRadius:scAnno.map(v=>v>60?7:4)},
      {type:'line',label:'Soglia 60%',data:anni.map(()=>60),borderColor:C.red,borderDash:[6,4],pointRadius:0,borderWidth:1.5}
    ]},
    options:{...chartOpts({legend:true,callbackY:v=>v.toFixed(1)+'%',C})}
  });

  const buckets=Array(10).fill(0);
  sR.forEach(r=>{buckets[Math.min(9,Math.floor(r.sconto*100/10))]++;});
  doBar('ch-sc-dist',
    ['0–10%','10–20%','20–30%','30–40%','40–50%','50–60%','60–70%','70–80%','80–90%','90–100%'],
    buckets,null,buckets.map((_,i)=>i>=6?C.red+'aa':i>=5?C.amber+'aa':C.green+'88')
  );

  const prodSconto=groupBy(sR,r=>r.desc,rows=>({sc:avg(rows,r=>r.sconto),n:rows.length,f:sum(rows,r=>r.importo),pz:avg(rows,r=>r.pz),lordo:rows[0].lordo}));
  scontiData=Object.entries(prodSconto).sort((a,b)=>b[1].f-a[1].f);
  renderScontiTbl();

  const over60cli=groupBy(over60,r=>r.cliente,rows=>({n:rows.length,f:sum(rows,r=>r.importo),sc:avg(rows,r=>r.sconto)}));
  tbl('tbl-over60-cli',['Cliente','Righe >60%','Valore','Sconto Medio'],
    Object.entries(over60cli).sort((a,b)=>b[1].n-a[1].n).slice(0,20).map(([k,v])=>[
      k,`<span class="bdg br">${v.n}</span>`,`<span class="mono">${fmt(v.f)}</span>`,
      `<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${pct(v.sc)}</span>`
    ]));
}

function renderScontiTbl(){
  const flt=document.getElementById('sc-flt').value;
  const q=document.getElementById('sc-srch').value.toLowerCase();
  let rows=scontiData.filter(([k])=>k.toLowerCase().includes(q));
  if(flt==='over')rows=rows.filter(([,v])=>v.sc>SCONTO_MAX);
  if(flt==='ok')rows=rows.filter(([,v])=>v.sc<=SCONTO_MAX);
  tbl('tbl-sc',['Prodotto/Macchina','Sconto Medio','N°','Fatturato','PZ Netto','Lordo Listino'],
    rows.slice(0,80).map(([k,v])=>[
      trunc(k,34),
      `<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${v.sc>SCONTO_MAX?'⚠ ':''}${pct(v.sc)}</span>`,
      v.n,`<span class="mono">${fmt(v.f)}</span>`,
      `<span class="mono">${fmt(v.pz)}</span>`,v.lordo?`<span class="mono">${fmt(v.lordo)}</span>`:'—'
    ]));
}

// ── MARGINE ──
function renderMargine(){
  const V=F.vend, C=tc();
  // Marginalità: include TUTTE le categorie
  // sc_real = sconto reale da listino se disponibile, null altrimenti
  // sc_eff  = sconto reale O 60% (soglia contrattuale) come fallback
  const catMarg=groupBy(V.filter(r=>r.cat&&r.cat!=='nan'&&r.cat.length>1),r=>r.cat,rows=>({
    sc:     rows.filter(r=>r.sconto!==null).length>0 ? avg(rows.filter(r=>r.sconto!==null),r=>r.sconto) : null,
    sc_eff: avg(rows,r=>r.sconto_eff),
    tr:     avg(rows,r=>r.incTrasp),
    f:      sum(rows,r=>r.importo),
    n:      rows.length
  }));
  const sorted=Object.entries(catMarg).map(([k,v])=>[k,{...v,erosione:v.sc_eff+v.tr}]).filter(([k])=>k&&k!=='nan').sort((a,b)=>b[1].f-a[1].f);

  const hmEl=document.getElementById('hm-margine'); hmEl.innerHTML='';
  sorted.forEach(([k,v])=>{
    const col=v.erosione>0.8?C.red:v.erosione>0.65?C.amber:C.green;
    hmEl.insertAdjacentHTML('beforeend',`
      <div class="hmc">
        <div class="hmn">${trunc(k.split(' - ')[0],22)}</div>
        <div class="hmv" style="color:${col}">${pct(v.erosione)}</div>
        <div class="hms">${v.sc!==null?'Sc '+pct(v.sc)+' + ':'Tr solo: '}${v.sc!==null?'Tr '+pct(v.tr):pct(v.tr)}</div>
        <div class="hmbar"><div class="hmfill" style="width:${Math.min(100,v.erosione*100)}%;background:${col}"></div></div>
      </div>`);
  });

  const erosColors=sorted.map(([,v])=>v.erosione>0.8?C.red+'bb':v.erosione>0.65?C.amber+'bb':v.sc===null?C.blue+'88':C.green+'88');
  doHBar('ch-eros',sorted.map(([k])=>trunc(k.split(' - ')[0],20)),sorted.map(([,v])=>v.erosione*100),null,erosColors);

  // scatter with size proportional to fatturato
  const maxF=Math.max(...sorted.map(([,v])=>v.f));
  dc('ch-scatter');
  const ctx=document.getElementById('ch-scatter').getContext('2d');
  charts['ch-scatter']=new Chart(ctx,{type:'bubble',
    data:{datasets:[{
      data:sorted.map(([,v])=>({x:v.sc_eff*100,y:v.tr*100,r:Math.max(4,Math.min(20,v.f/maxF*18))})),
      backgroundColor:sorted.map(([,v])=>v.erosione>0.8?C.red+'bb':v.erosione>0.65?C.amber+'bb':C.green+'88'),
      borderColor:'transparent'
    }]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{backgroundColor:C.tooltip.bg,borderColor:C.tooltip.border,borderWidth:1,titleColor:C.text2,bodyColor:C.text2,
          callbacks:{label:ctx=>`${sorted[ctx.dataIndex]?.[0]?.split(' - ')[0]||''}: Sc ${ctx.parsed.x.toFixed(1)}% + Tr ${ctx.parsed.y.toFixed(1)}%`}}},
      scales:{
        x:{title:{display:true,text:'Sconto %',color:C.text3,font:{size:9}},grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>v+'%'}},
        y:{title:{display:true,text:'Incidenza Trasp %',color:C.text3,font:{size:9}},grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>v+'%'}}
      }
    }
  });

  tbl('tbl-marg',['Categoria','Erosione Totale','Sc. Reale','Sc. Usato','Trasp %','Fatturato','N°'],
    sorted.map(([k,v])=>[trunc(k,30),`<span class="bdg ${v.erosione>0.8?'br':v.erosione>0.65?'ba':'bg'}">${pct(v.erosione)}</span>`,v.sc!==null?`<span class="mono">${pct(v.sc)}</span>`:'<span class="ba bdg" title="Stima: soglia max 60%">~60%</span>',`<span class="mono">${pct(v.sc_eff)}</span>`,`<span class="mono">${pct(v.tr)}</span>`,`<span class="mono">${fmt(v.f)}</span>`,v.n]));
}

// ── TRASPORTI ──
function renderTrasporti(){
  const V=F.vend, C=tc(), anni=G.anni;
  const fAnnuo=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAnnuo=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const inc=anni.map(a=>(tAnnuo[a]||0)/(fAnnuo[a]||1));
  const trTot=sum(V,r=>r.trasp), fTot=sum(V,r=>r.importo);
  const incMed=fTot>0?trTot/fTot:0;
  const piccoIdx=inc.indexOf(Math.max(...inc)); const picco=anni[piccoIdx];

  kpi('kr-tr',[
    {l:'Spese Trasporto Periodo',v:fmt(trTot),col:'p',sub:`${pct(incMed)} sul fatturato`},
    {l:'Incidenza Media',v:pct(incMed),col:incMed>0.05?'r':'g',sub:'% sul fatturato totale'},
    {l:`Anno Picco`,v:`${picco}`,col:'a',sub:`${pct(inc[piccoIdx])} incidenza`},
    {l:'Trasp. Ordini Inevasi',v:fmt(sum(G.ORDINI,r=>r.trasp)),col:'p',sub:`${G.ORDINI.length} ordini aperti`},
  ]);

  dc('ch-tr-yr');
  const ctx=document.getElementById('ch-tr-yr').getContext('2d');
  charts['ch-tr-yr']=new Chart(ctx,{
    data:{labels:anni,datasets:[
      {type:'bar',label:'Trasporto €',data:anni.map(a=>tAnnuo[a]||0),backgroundColor:anni.map(a=>a===picco?C.red+'cc':C.purple+'88'),borderRadius:4,yAxisID:'y'},
      {type:'line',label:'Incidenza %',data:inc.map(v=>v*100),borderColor:C.amber,tension:.3,pointRadius:4,fill:false,yAxisID:'y2',
       pointBackgroundColor:inc.map((v,i)=>anni[i]===picco?C.red:C.amber)}
    ]},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:true,labels:{color:C.text2,font:{size:10},boxWidth:10,padding:8}},
        tooltip:{backgroundColor:C.tooltip.bg,borderColor:C.tooltip.border,borderWidth:1,titleColor:C.text2,bodyColor:C.text2},
        annotation:{annotations:{
          peak:{type:'label',xValue:picco,yValue:inc[piccoIdx]*100,yScaleID:'y2',
            backgroundColor:C.red+'dd',color:'#fff',font:{size:9},
            content:[`Picco ${picco}`,`${(inc[piccoIdx]*100).toFixed(1)}%`],padding:4,borderRadius:4}
        }}
      },
      scales:{
        x:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9}}},
        y:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>fmtShort(v)}},
        y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:C.amber,font:{size:9},callback:v=>v.toFixed(1)+'%'}}
      }
    }
  });

  const portD={};
  V.forEach(r=>{const p=r.porto_desc||'Altro';portD[p]=(portD[p]||0)+1;});
  doPie('ch-porto',Object.keys(portD),Object.values(portD));

  tbl('tbl-tr',['Anno','Fatturato','Trasporto','Incidenza','Δ'],
    anni.map((a,i)=>{
      const dlt=i>0?inc[i]-inc[i-1]:null;
      return [a,`<span class="mono">${fmt(fAnnuo[a]||0)}</span>`,`<span class="mono">${fmt(tAnnuo[a]||0)}</span>`,
        `<span class="bdg ${inc[i]>0.05?'br':'bg'}">${pct(inc[i])}</span>`,
        dlt!==null?`<span class="bdg ${dlt>0?'br':'bg'}">${dlt>0?'+':''}${pct(dlt)}</span>`:'—'];
    }));

  // ── Trasporti per Regione
  const trReg=groupBy(V.filter(r=>r.regione&&r.regione!=='nan'),r=>r.regione,rows=>({f:sum(rows,r=>r.importo),tr:sum(rows,r=>r.trasp)}));
  const regSorted=Object.entries(trReg).map(([k,v])=>({k,f:v.f,tr:v.tr,inc:v.f>0?v.tr/v.f:0})).sort((a,b)=>b.tr-a.tr);
  dc('ch-tr-reg');
  const ctxReg=document.getElementById('ch-tr-reg');
  if(ctxReg){
    charts['ch-tr-reg']=new Chart(ctxReg.getContext('2d'),{type:'bar',
      data:{labels:regSorted.map(r=>r.k),datasets:[
        {label:'Trasporto €',data:regSorted.map(r=>r.tr),backgroundColor:regSorted.map(r=>r.inc>0.08?C.red+'aa':r.inc>0.05?C.amber+'aa':C.purple+'88'),borderRadius:3}
      ]},
      options:{...chartOpts({callbackY:v=>fmtShort(v),C}),indexAxis:'y'}
    });
  }

  // ── Trasporti per Agente
  const trAgt=groupBy(V.filter(r=>r.agente&&r.agente!=='nan'),r=>r.agente,rows=>({f:sum(rows,r=>r.importo),tr:sum(rows,r=>r.trasp)}));
  const agtTrSorted=Object.entries(trAgt).map(([k,v])=>({k,f:v.f,tr:v.tr,inc:v.f>0?v.tr/v.f:0})).sort((a,b)=>b.tr-a.tr);

  tbl('tbl-tr-reg',['Regione','Fatturato','Trasporto','Incidenza'],
    regSorted.slice(0,20).map(r=>[r.k,`<span class="mono">${fmt(r.f)}</span>`,`<span class="mono">${fmt(r.tr)}</span>`,
      `<span class="bdg ${r.inc>0.08?'br':r.inc>0.05?'ba':'bg'}">${pct(r.inc)}</span>`]));

  tbl('tbl-tr-agt',['Agente','Fatturato','Trasporto','Incidenza'],
    agtTrSorted.map(r=>[r.k,`<span class="mono">${fmt(r.f)}</span>`,`<span class="mono">${fmt(r.tr)}</span>`,
      `<span class="bdg ${r.inc>0.06?'br':r.inc>0.04?'ba':'bg'}">${pct(r.inc)}</span>`]));
}

// ── ORDINI ──
let ordiniAll=[];
function renderOrdini(){
  const O=G.ORDINI, C=tc(), today=new Date();
  const scaduti=O.filter(r=>r.consegna&&r.consegna<today);
  const in30=O.filter(r=>{if(!r.consegna||r.consegna<today)return false;return(r.consegna-today)/86400000<=30;});
  ordiniAll=[...O].sort((a,b)=>(a.consegna||new Date(9999,0))-(b.consegna||new Date(9999,0)));

  kpi('kr-ord',[
    {l:'Valore Inevaso',v:fmt(sum(O,r=>r.importoI)),col:'a',sub:`${O.length} righe aperte`},
    {l:'Clienti con Ordini',v:[...new Set(O.map(r=>r.cliente).filter(Boolean))].length,col:'b'},
    {l:'Ordini Scaduti',v:scaduti.length,col:'r',sub:fmt(sum(scaduti,r=>r.importoI))},
    {l:'Scadenza 30gg',v:in30.length,col:'a',sub:fmt(sum(in30,r=>r.importoI))},
    {l:'Trasp. Pianificato',v:fmt(sum(O,r=>r.trasp)),col:'p'},
  ]);

  if(scaduti.length>0)document.getElementById('ord-scad').textContent=`${scaduti.length} SCADUTI`;

  const cliOrd=groupBy(O,r=>r.cliente,rows=>sum(rows,r=>r.importoI));
  const top12=Object.entries(cliOrd).sort((a,b)=>b[1]-a[1]).slice(0,12);
  doHBar('ch-ord-cli',top12.map(([k])=>trunc(k,24)),top12.map(([,v])=>v),C.amber);

  const dateOrd=groupBy(O.filter(r=>r.consegna&&!isNaN(r.consegna)),r=>r.consegna.toISOString().split('T')[0],rows=>sum(rows,r=>r.importoI));
  const dateSorted=Object.entries(dateOrd).sort(([a],[b])=>a.localeCompare(b)).slice(0,20);
  doBar('ch-ord-date',dateSorted.map(([k])=>k.slice(5)),dateSorted.map(([,v])=>v),null,
    dateSorted.map(([k])=>new Date(k)<today?C.red+'aa':C.blue+'aa'));

  filterOrdTbl();
}
function filterOrdTbl(){
  const q=document.getElementById('ord-srch').value.toLowerCase(), today=new Date();
  const rows=ordiniAll.filter(r=>(r.cliente+r.desc).toLowerCase().includes(q)).slice(0,80);
  tbl('tbl-ord',['Cliente','Prodotto','Qtà','Importo','Consegna','Stato'],
    rows.map(r=>{
      const late=r.consegna&&r.consegna<today;
      const soon=r.consegna&&!late&&(r.consegna-today)/86400000<=30;
      const stato=late?'<span class="bdg br">SCADUTO</span>':soon?'<span class="bdg ba">≤30gg</span>':'<span class="bdg bb">OK</span>';
      return [trunc(r.cliente,26),trunc(r.desc,28),r.qtyI,`<span class="mono">${fmt(r.importoI)}</span>`,
        r.consegna&&!isNaN(r.consegna)?r.consegna.toLocaleDateString('it'):'—',stato];
    }));
}

// ── CRITICITÀ ──
function renderCriticita(){
  const V=G.VEND, O=G.ORDINI;
  const sR=V.filter(r=>r.sconto!==null);
  const over60=sR.filter(r=>r.sconto>SCONTO_MAX);
  const sMed=sR.length?avg(sR,r=>r.sconto):0;
  const anni=G.anni, annoMax=Math.max(...anni);
  const fAnnuo=groupBy(V,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAnnuo=groupBy(V,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const today=new Date();
  const scaduti=O.filter(r=>r.consegna&&r.consegna<today);
  const in30=O.filter(r=>{if(!r.consegna||r.consegna<today)return false;return(r.consegna-today)/86400000<=30;});

  const alerts=[];

  if(over60.length){
    const pOver=sR.length?over60.length/sR.length:0;
    alerts.push({type:pOver>0.1?'danger':'warn',icon:'🏷️',
      t:`Sconti >60%: ${over60.length} righe (${pct(pOver)}) — valore €${fmt(sum(over60,r=>r.importo))}`,
      b:`Verificare se trattasi di promozioni autorizzate. Prodotti più colpiti: ${[...new Set(over60.sort((a,b)=>b.sconto-a.sconto).slice(0,3).map(r=>r.desc))].join(', ')}.`});
  }

  const diffS=sMed-SCONTO_MAX;
  alerts.push({type:diffS>0?'warn':'ok',icon:diffS>0?'⚠️':'✅',
    t:`Sconto medio globale: ${pct(sMed)} (soglia contrattuale: ${pct(SCONTO_MAX)})`,
    b:diffS>0?`Supera la soglia di ${pct(Math.abs(diffS))}. Consigliare revisione policy con la rete vendita.`:`Dentro i limiti contrattuali (margine residuo: ${pct(Math.abs(diffS))}).`});

  const f25=fAnnuo[2025]||0, f24=fAnnuo[2024]||0;
  if(f24&&f25){const d=(f25-f24)/f24;alerts.push({type:d<-0.05?'warn':d>0.05?'ok':'info',icon:d>0?'📈':'📉',
    t:`Trend 2025 vs 2024: ${d>0?'+':''}${pct(d)}`,
    b:`2024: €${fmt(f24)} → 2025: €${fmt(f25)}. ${d<0?'Attenzione alla contrazione del fatturato.':'Crescita positiva confermata.'}`});}

  const incidenze=anni.map(a=>({a,inc:(tAnnuo[a]||0)/(fAnnuo[a]||1)}));
  const picco=incidenze.reduce((m,i)=>i.inc>m.inc?i:m,incidenze[0]);
  if(picco.inc>0.07)alerts.push({type:'warn',icon:'🚚',
    t:`Picco trasporti ${picco.a}: ${pct(picco.inc)} — €${fmt(tAnnuo[picco.a]||0)}`,
    b:`Analizzare i fattori scatenanti per prevenire anomalie future.`});

  if(scaduti.length)alerts.push({type:'danger',icon:'⏰',
    t:`${scaduti.length} ordini con data consegna scaduta — €${fmt(sum(scaduti,r=>r.importoI))}`,
    b:`Clienti: ${[...new Set(scaduti.map(r=>r.cliente).filter(Boolean))].slice(0,4).join(', ')}. Aggiornare date o contattare i clienti.`});

  if(in30.length)alerts.push({type:'warn',icon:'📅',
    t:`${in30.length} righe in scadenza entro 30 giorni — €${fmt(sum(in30,r=>r.importoI))}`,
    b:'Pianificare priorità logistica per evasione nei tempi previsti.'});

  const cliF=groupBy(V.filter(r=>r.anno>=2024),r=>r.cliente,rows=>sum(rows,r=>r.importo));
  const fTot2425=sum(V.filter(r=>r.anno>=2024),r=>r.importo);
  const top3=Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,3);
  const top3pct=fTot2425>0?sum(top3,([,v])=>v)/fTot2425:0;
  if(top3pct>0.35)alerts.push({type:'warn',icon:'🏢',
    t:`Alta concentrazione: top 3 clienti = ${pct(top3pct)} del fatturato 2024–${annoMax}`,
    b:top3.map(([k,v])=>`${trunc(k,22)} ${pct(fTot2425>0?v/fTot2425:0)}`).join(' · ')+'. Diversificare il portafoglio.'});

  document.getElementById('nbadge').textContent=alerts.filter(a=>a.type==='danger'||a.type==='warn').length;
  document.getElementById('alerts').innerHTML=alerts.length===0
    ?'<div class="al ok"><div class="al-ic">✅</div><div class="al-b"><strong>Nessuna criticità rilevata</strong><p>Tutti gli indicatori sono nei range normali.</p></div></div>'
    :alerts.map(a=>`<div class="al ${a.type}"><div class="al-ic">${a.icon}</div><div class="al-b"><strong>${a.t}</strong><p>${a.b}</p></div></div>`).join('');
}

// ══════════════════════════════════════════════════
//  CHART FACTORY
// ══════════════════════════════════════════════════
function chartOpts({legend=false,callbackY=null,C}={}){
  const c=C||tc();
  const yFn=callbackY||(v=>fmtShort(v));
  return{
    responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:legend?{display:true,labels:{color:c.text2,font:{size:10,family:'DM Sans'},boxWidth:10,padding:8}}:{display:false},
      tooltip:{backgroundColor:c.tooltip.bg,borderColor:c.tooltip.border,borderWidth:1,titleColor:c.text2,bodyColor:c.text2,padding:10}
    },
    scales:{
      x:{grid:{color:c.border},ticks:{color:c.text3,font:{size:9,family:'DM Sans'},maxRotation:45}},
      y:{grid:{color:c.border},ticks:{color:c.text3,font:{size:9,family:'DM Sans'},callback:yFn}}
    }
  };
}
function doPie(id,labels,data){
  const C=tc(), PAL=[C.green,C.blue,C.amber,C.purple,C.red,C.cyan,C.green+'88',C.blue+'88',C.amber+'88'];
  dc(id);
  const ctx=document.getElementById(id).getContext('2d');
  charts[id]=new Chart(ctx,{type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:PAL,borderWidth:0,hoverOffset:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'right',labels:{color:C.text2,font:{size:9,family:'DM Sans'},boxWidth:8,padding:6}},
        tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${fmt(ctx.raw)}`}}}}
  });
}
function doBar(id,labels,data,colors,colorsArr,extraScales={}){
  const C=tc(); dc(id);
  const ctx=document.getElementById(id).getContext('2d');
  charts[id]=new Chart(ctx,{type:'bar',
    data:{labels,datasets:[{data,backgroundColor:colorsArr||colors||[C.blue+'aa'],borderRadius:3}]},
    options:{...chartOpts({callbackY:v=>fmtShort(v),C}),
      scales:{...chartOpts({C}).scales,x:{...chartOpts({C}).scales.x,...extraScales}}}
  });
}
function doHBar(id,labels,data,color,colors){
  const C=tc(); dc(id);
  const ctx=document.getElementById(id).getContext('2d');
  charts[id]=new Chart(ctx,{type:'bar',
    data:{labels,datasets:[{data,backgroundColor:colors||(color||C.green),borderRadius:3}]},
    options:{...chartOpts({callbackY:v=>fmtShort(v),C}),indexAxis:'y'}
  });
}
function dc(id){if(charts[id]){charts[id].destroy();delete charts[id];}}

// ══════════════════════════════════════════════════
//  TABLE ENGINE (sortable)
// ══════════════════════════════════════════════════
function tbl(id,headers,rows){
  const el=document.getElementById(id);if(!el)return;
  const s=sortState[id]||{col:-1,asc:true};
  let sr=[...rows];
  if(s.col>=0){sr.sort((a,b)=>{const va=sh(a[s.col]),vb=sh(b[s.col]);const na=parseFloat(va.replace(/[€%., ]/g,'')),nb=parseFloat(vb.replace(/[€%., ]/g,''));let c=!isNaN(na)&&!isNaN(nb)?na-nb:va.localeCompare(vb,'it');return s.asc?c:-c;});}
  el.innerHTML=`<thead><tr>${headers.map((h,i)=>`<th class="${s.col===i?(s.asc?'sa':'sd'):''}" onclick="sortTbl('${id}',${i})">${h}</th>`).join('')}</tr></thead><tbody>${sr.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>`;
}
function sortTbl(id,col){
  const s=sortState[id]||{col:-1,asc:true};
  sortState[id]={col,asc:s.col===col?!s.asc:true};
  const map={'tbl-cat':'vendite','tbl-drill':'vendite','tbl-cli':'clienti','tbl-sc':'sconti','tbl-over60-cli':'sconti','tbl-marg':'margine','tbl-tr':'trasporti','tbl-tr-reg':'trasporti','tbl-tr-agt':'trasporti','tbl-ord':'ordini'};
  if(map[id]==='vendite')renderVendite();
  else if(map[id]==='clienti')filterCliTbl();
  else if(map[id]==='sconti')renderScontiTbl();
  else if(map[id]==='margine')renderMargine();
  else if(map[id]==='trasporti')renderTrasporti();
  else if(map[id]==='ordini')filterOrdTbl();
}
function sh(s){return String(s).replace(/<[^>]+>/g,'').trim();}

// ══════════════════════════════════════════════════
//  UI HELPERS
// ══════════════════════════════════════════════════
function kpi(elId,items){
  document.getElementById(elId).innerHTML=items.map(i=>`
    <div class="kk">
      <div class="kk-bar ${i.col||'g'}"></div>
      <div class="kl">${i.l}</div>
      <div class="kv ${i.col||'def'}">${i.v}</div>
      <div class="ka">
        ${i.sub?`<span class="ks">${i.sub}</span>`:''}
        ${i.delta!==undefined&&i.delta!==null?`<span class="dt ${i.delta>=0?'up':'dn'}">${i.delta>=0?'↑':'↓'}${pct(Math.abs(i.delta))}</span>`:''}
      </div>
    </div>`).join('');
}

function go(name){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.ni').forEach(n=>n.classList.remove('on'));
  document.getElementById('panel-'+name).classList.add('on');
  document.querySelectorAll('.ni').forEach(n=>{if(n.getAttribute('onclick')===`go('${name}')`)n.classList.add('on');});
  document.getElementById('main').scrollTop=0;
}

function showLoad(m){document.getElementById('loading').style.display='flex';setLoad(m);}
function setLoad(m){document.getElementById('lmsg').textContent=m;}
function hideLoad(){document.getElementById('loading').style.display='none';}
function resetApp(){location.reload();}
function sleep(ms){return new Promise(r=>setTimeout(r,ms));}

// ══════════════════════════════════════════════════
//  MATH & STRING HELPERS
// ══════════════════════════════════════════════════
function fmt(v){if(v===null||v===undefined||isNaN(v))return'—';return'€'+Number(v).toLocaleString('it',{minimumFractionDigits:0,maximumFractionDigits:0});}
function fmtShort(v){if(!v&&v!==0)return'—';const a=Math.abs(v);if(a>=1e6)return'€'+(v/1e6).toFixed(1)+'M';if(a>=1000)return'€'+(v/1000).toFixed(0)+'k';return'€'+Math.round(v);}
function pct(v){if(v===null||v===undefined||isNaN(v))return'—';return(Number(v)*100).toFixed(1)+'%';}
function num(v){return parseFloat(v)||0;}
function str(v){return String(v||'').trim();}
function sum(arr,fn){return(arr||[]).reduce((a,r)=>a+(parseFloat(fn(r))||0),0);}
function avg(arr,fn){if(!arr||!arr.length)return 0;return sum(arr,fn)/arr.length;}
function groupBy(arr,kFn,vFn){const r={};(arr||[]).forEach(x=>{const k=kFn(x);if(!r[k])r[k]=[];r[k].push(x);});if(vFn)Object.keys(r).forEach(k=>{r[k]=vFn(r[k]);});return r;}
function trunc(s,n){return s&&s.length>n?s.slice(0,n-1)+'…':s||'';}
