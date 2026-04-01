// PezzaliApp — Cormach Dashboard v3
// Agenti: CABASSI, PEZZALI, MARABELLI, BRUNO ecc (col M = DESCRIZIONE ELEMENTO.2)
// Categorie: SMONTAGOMME, EQUILIBRATRICI ecc (col V = DESCRIZIONE ELEMENTO.5)
// Linea prodotto: F 536GT, MEC 10 ecc (col S = DESCRIZIONE ELEMENTO.4)
// Macchina: descrizione specifica (col W = DESCRIZIONE)
'use strict';

const SCONTO_MAX = 0.60;
const MESI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
const QNAMES = ['Q1','Q2','Q3','Q4'];
const PORTO_LABELS = {1:'Franco',2:'Assegnato',3:'Franco+Add.',6:'Altro',29:'Altro',168:'Altro'};

// Cascos C-series detection via descrizione
const CASCOS_PAT = /\bC\s*3[.,]2|\bC\s*3[.,]5|\bC\s*4\s*(XL|s|S|\b)|\bC\s*5\s*(XL|s|S|[.,]5|\b)|\bC\s*5[.,]5|\bC\s*7s\b|\bC\s*4[3-9]\d\b|\bC\s*4\d{2}\b|\bC\s*125\b|\bPARKING\s*2x1\b/i;
const CASCOS_EX  = /PFA\s*\d+|STD\s*72|FORBICE|SCHEDA|CENTRALINA|POMPA|FOTOCELLULA|USATO|CARTER/i;

let G={}, F={}, CMP=false, charts={}, sortState={};
const BUDGET_KEY='pza-budgets-v1';

function tc(){
  const dark = document.documentElement.getAttribute('data-theme') !== 'light';
  return {
    green:  dark?'#00b894':'#059669',
    red:    dark?'#e17055':'#dc2626',
    amber:  dark?'#fdcb6e':'#d97706',
    blue:   dark?'#74b9ff':'#2563eb',
    purple: dark?'#a29bfe':'#7c3aed',
    cyan:   dark?'#00cec9':'#0891b2',
    text2:  dark?'#94a3b8':'#475569',
    text3:  dark?'#4a5568':'#94a3b8',
    border: dark?'rgba(255,255,255,.06)':'rgba(0,0,0,.06)',
    tip:    {bg:dark?'#1a2233':'#fff', border:dark?'rgba(255,255,255,.1)':'rgba(0,0,0,.1)'}
  };
}



function loadBudgets(){
  try{
    const raw=JSON.parse(localStorage.getItem(BUDGET_KEY)||'[]');
    return Array.isArray(raw)?raw.map(r=>({
      year:parseInt(r.year)||0,
      agente:str(r.agente).toUpperCase(),
      budget:num(r.budget)
    })).filter(r=>r.year>2000&&r.agente&&r.budget>=0):[];
  }catch(e){return [];}
}
function saveBudgets(list){
  localStorage.setItem(BUDGET_KEY,JSON.stringify((list||[]).sort((a,b)=>a.year-b.year||a.agente.localeCompare(b.agente,'it'))));
}
function periodMonthFactor(){
  const perVal=(document.getElementById('f-per')||{}).value||'';
  if(perVal.startsWith('q')) return 3/12;
  if(perVal.startsWith('m')) return 1/12;
  return 1;
}
function activeYears(){
  const annoDa=parseInt((document.getElementById('f-da')||{}).value)||0;
  const annoA=parseInt((document.getElementById('f-a')||{}).value)||0;
  const years=[];
  for(let y=annoDa;y<=annoA;y++) years.push(y);
  return years;
}
function getActiveBudgetEntries(filteredAgent=''){
  const years=activeYears();
  const factor=periodMonthFactor();
  return (G.budgets||[])
    .filter(r=>years.includes(r.year) && (!filteredAgent || r.agente===filteredAgent))
    .map(r=>({...r,budgetActive:r.budget*factor}));
}
function getActiveBudgetMap(filteredAgent=''){
  const out={};
  getActiveBudgetEntries(filteredAgent).forEach(r=>{out[r.agente]=(out[r.agente]||0)+r.budgetActive;});
  return out;
}
function upsertBudget(year,agente,budget){
  const rows=loadBudgets();
  const y=parseInt(year)||0,a=str(agente).toUpperCase(),b=num(budget);
  if(!y||!a) return;
  const idx=rows.findIndex(r=>r.year===y&&r.agente===a);
  if(idx>=0) rows[idx].budget=b; else rows.push({year:y,agente:a,budget:b});
  saveBudgets(rows);
  G.budgets=rows;
}
function deleteBudget(year,agente){
  const y=parseInt(year)||0,a=str(agente).toUpperCase();
  const rows=loadBudgets().filter(r=>!(r.year===y&&r.agente===a));
  saveBudgets(rows);
  G.budgets=rows;
}
function seedBudgetAgentSelect(){
  const sel=document.getElementById('bud-agent');
  if(!sel) return;
  const curr=sel.value;
  const names=[...new Set([...(G.agenti||[]),...(G.budgets||[]).map(r=>r.agente)])].filter(Boolean).sort();
  sel.innerHTML='<option value="">Seleziona agente esistente</option>';
  names.forEach(n=>sel.insertAdjacentHTML('beforeend',`<option value="${n}">${n}</option>`));
  if(names.includes(curr)) sel.value=curr;
}
function renderBudgetManager(){
  const ySel=document.getElementById('bud-year');
  if(!ySel) return;
  const years=[...new Set([...(G.anni||[]),...(G.budgets||[]).map(r=>r.year)])].sort((a,b)=>a-b);
  ySel.innerHTML='';
  years.forEach(y=>ySel.insertAdjacentHTML('beforeend',`<option value="${y}">${y}</option>`));
  if(!ySel.value && years.length) ySel.value=years[years.length-1];
  seedBudgetAgentSelect();
  const y=parseInt(ySel.value)||years[years.length-1]||new Date().getFullYear();
  const rows=(G.budgets||[]).filter(r=>r.year===y).sort((a,b)=>b.budget-a.budget||a.agente.localeCompare(b.agente,'it'));
  tbl('tbl-budget',['Anno','Agente/Commerciale','Budget','Azioni'],
    rows.map(r=>[
      `<span class="mono">${r.year}</span>`,
      r.agente,
      `<span class="mono">${fmt(r.budget)}</span>`,
      `<button class="bsm" onclick="loadBudgetIntoForm(${r.year},'${r.agente.replace(/'/g,"\\'")}')">Modifica</button> <button class="bsm" onclick="removeBudget(${r.year},'${r.agente.replace(/'/g,"\\'")}')">Elimina</button>`
    ])
  );
  const total=rows.reduce((a,r)=>a+r.budget,0);
  const el=document.getElementById('budget-total');
  if(el) el.textContent=`Budget anno ${y}: ${fmt(total)} · ${rows.length} voci`;
}
function loadBudgetIntoForm(year,agente){
  document.getElementById('bud-year').value=year;
  if([...(document.getElementById('bud-agent').options||[])].some(o=>o.value===agente)) document.getElementById('bud-agent').value=agente;
  else document.getElementById('bud-agent-new').value=agente;
  const row=(G.budgets||[]).find(r=>r.year===parseInt(year)&&r.agente===agente);
  document.getElementById('bud-value').value=row?Math.round(row.budget):'';
}
function saveBudgetFromForm(){
  const year=parseInt(document.getElementById('bud-year').value)||0;
  const existing=document.getElementById('bud-agent').value;
  const custom=document.getElementById('bud-agent-new').value;
  const agente=str(existing||custom).toUpperCase();
  const budget=num(document.getElementById('bud-value').value);
  if(!year||!agente){alert('Inserisci anno e nome commerciale.');return;}
  upsertBudget(year,agente,budget);
  document.getElementById('bud-agent').value='';
  document.getElementById('bud-agent-new').value='';
  document.getElementById('bud-value').value='';
  renderBudgetManager();
  renderAll();
}
function removeBudget(year,agente){
  deleteBudget(year,agente);
  renderBudgetManager();
  renderAll();
}

// ── FILE UPLOAD
['v','o','l'].forEach(id=>{
  document.getElementById('fi-'+id).addEventListener('change',e=>{
    const f=e.target.files[0]; if(!f) return;
    document.getElementById('uc-'+id).classList.add('ok');
    document.getElementById('fn-'+id).textContent=f.name;
    checkReady();
  });
});
function checkReady(){
  const ok=['v','o'].every(id=>document.getElementById('uc-'+id).classList.contains('ok'));
  document.getElementById('btn-go').classList.toggle('on',ok);
}
document.querySelectorAll('.upc').forEach(card=>{
  card.addEventListener('dragover',e=>{e.preventDefault();card.classList.add('drag');});
  card.addEventListener('dragleave',()=>card.classList.remove('drag'));
  card.addEventListener('drop',e=>{
    e.preventDefault();card.classList.remove('drag');
    const inp=card.querySelector('input[type=file]');
    if(!inp||!e.dataTransfer.files[0])return;
    const dt=new DataTransfer();dt.items.add(e.dataTransfer.files[0]);inp.files=dt.files;
    inp.dispatchEvent(new Event('change'));
  });
});

async function runAnalysis(){
  showLoad('Lettura vendite...'); await sleep(30);
  try{
    const vRaw=await readXLSX('fi-v');
    setLoad('Lettura ordini...'); await sleep(20);
    const oRaw=await readXLSX('fi-o');
    setLoad('Caricamento listino...'); await sleep(20);
    let lRaw=null;
    const lInp=document.getElementById('fi-l');
    if(lInp.files[0]) lRaw=lInp.files[0].name.endsWith('.csv')?await readCSV('fi-l'):await readXLSX('fi-l');
    setLoad('Elaborazione dati...'); await sleep(40);
    processData(vRaw,oRaw,lRaw);
    setLoad('Rendering...'); await sleep(30);
    initDashboard();
    hideLoad();
    document.getElementById('upload-screen').style.display='none';
    document.getElementById('top-filters').removeAttribute('hidden');
    document.getElementById('btn-reset').removeAttribute('hidden');
    document.getElementById('btn-print').removeAttribute('hidden');
    document.getElementById('status-pill').textContent='ATTIVO';
  }catch(err){hideLoad();alert('Errore: '+err.message);console.error(err);}
}
function readXLSX(id){
  return new Promise((res,rej)=>{
    const f=document.getElementById(id).files[0]; if(!f)return rej(new Error('File mancante'));
    const r=new FileReader();
    r.onload=e=>{try{
      const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      // Leggi con header:1 per usare lettere colonna (A,B,C...) invece di nomi duplicati
      const rawRows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      if(rawRows.length<2){res([]);return;}
      // Costruisci oggetti con chiavi = lettera colonna (A,B,C..AA,AB...)
      const toKey=i=>i<26?String.fromCharCode(65+i):String.fromCharCode(64+Math.floor(i/26))+String.fromCharCode(65+(i%26));
      const rows=rawRows.slice(1).map(row=>{
        const obj={};
        row.forEach((v,i)=>{obj[toKey(i)]=v;});
        return obj;
      });
      res(rows);
    }catch(e){rej(e);}};
    r.onerror=()=>rej(new Error('Lettura fallita'));r.readAsArrayBuffer(f);
  });
}
function readCSV(id){
  return new Promise((res,rej)=>Papa.parse(document.getElementById(id).files[0],{header:true,skipEmptyLines:true,dynamicTyping:true,complete:r=>res(r.data),error:e=>rej(e)}));
}

function processData(vRaw,oRaw,lRaw){
  // Listino
  const listinoMap={};
  if(lRaw) lRaw.forEach(r=>{
    const cod=str(r.Codice||r.codice||'').replace(/^0+/,'');
    const pl=num(r.PrezzoLordo||r.prezzo_lordo||0);
    if(cod&&pl>0) listinoMap[cod]={pl,inst:num(r.CostoInstallazione||0),trasp:num(r.CostoTrasporto||0)};
  });

  // MAPPATURA PER LETTERA COLONNA (file Excel ha "DESCRIZIONE ELEMENTO" duplicato)
  // SheetJS con header:1 usa lettere A,B,C... come chiavi
  // Vendite Excel: V=cat, W=desc macchina, M=agente, O=regione, S=sottocat, J=città
  const CV={
    anno:    'A',   // ANNO SPEDIZIONE
    data:    'H',   // DATA SPEDIZIONE
    importo: 'Z',   // IMPORTO CONSEGNATO
    pz:      'Y',   // PZ NETTO VENDITA
    qty:     'X',   // QTA CONSEGNATA
    trasp:   'D',   // SPESE DI TRASPORTO
    causale: 'AA',  // CAUSALE MAGAZZINO
    cat:     'V',   // col V = DESCRIZIONE ELEMENTO (categoria: SMONTAGOMME, EQUILIBRATRICI...)
    sottocat:'S',   // col S = DESCRIZIONE ELEMENTO (linea: F 536GT, MEC 10...)
    agente:  'M',   // col M = DESCRIZIONE ELEMENTO (agente: CABASSI, PEZZALI...)
    cliente: 'Q',   // RAGIONE SOCIALE 1
    dest:    'R',   // DESC. DESTINAZIONE 1
    articolo:'T',   // ARTICOLO
    desc:    'W',   // col W = DESCRIZIONE (macchina specifica)
    porto:   'E',   // PORTO
    nazione: 'P',   // NAZIONE
    regione: 'O',   // col O = DESCRIZIONE ELEMENTO (regione: LOMBARDIA, VENETO...)
    citta:   'J',   // col J = DESCRIZIONE ELEMENTO (città: MILANO, PADOVA...)
  };

  console.log('[PezzaliApp v3] Lettura per lettera colonna: V=cat, M=agente, O=regione');

  const VEND=vRaw.filter(r=>str(r[CV.causale]).toUpperCase().startsWith('V')).map(r=>{
    const cod=str(r[CV.articolo]).replace(/^0+/,'');
    const li=listinoMap[cod]||null;
    const pz=num(r[CV.pz]),importo=num(r[CV.importo]),trasp=num(r[CV.trasp]);
    const lordo=li?li.pl:null;
    const sconto=lordo&&lordo>0&&pz>0?Math.max(0,Math.min(1,1-pz/lordo)):null;
    const sconto_eff=sconto!==null?sconto:0.60;
    const anno=parseInt(r[CV.anno])||0;
    const desc=str(r[CV.desc]);
    const brand=(CASCOS_PAT.test(desc)&&!CASCOS_EX.test(desc))?'Cascos':(li?'Cormach':'__no');
    // Normalizza cat e agente — rimuovi NaN letterale e valori numerici puri
    const _cat=str(r[CV.cat]||'');
    const _agente=str(r[CV.agente]||'');
    const cat=(_cat===''||_cat==='nan'||_cat==='None'||(!isNaN(parseFloat(_cat))&&isFinite(_cat)))?'':_cat;
    const agente=(_agente===''||_agente==='nan'||_agente==='None'||(!isNaN(parseFloat(_agente))&&isFinite(_agente)))?'':_agente;
    const sottocat=str(r[CV.sottocat]||'');
    const regione=str(r[CV.regione]||'');
    const citta=str(r[CV.citta]||'');
    let date=null,mese=-1,trim=-1;
    try{const d=r[CV.data];date=d instanceof Date?d:new Date(d);if(!isNaN(date)){mese=date.getMonth();trim=Math.floor(mese/3);}}catch(e){}
    return {
      anno,date,mese,trim,importo,pz,qty:num(r[CV.qty]),trasp,
      cat,sottocat,agente,cliente:str(r[CV.cliente]),dest:str(r[CV.dest]),
      nazione:str(r[CV.nazione]),regione,citta,
      articolo:str(r[CV.articolo]),desc,lordo,sconto,sconto_eff,brand,
      incTrasp:importo>0?trasp/importo:0,
      porto_desc:PORTO_LABELS[r[CV.porto]]||'Altro',
    };
  }).filter(r=>r.anno>2000);

  const RESI=vRaw.filter(r=>['H1','X2','X4','X5'].includes(str(r[CV.causale]).toUpperCase()));

  // Verifica dati caricati correttamente
  const catCheck=VEND.filter(r=>r.cat&&r.cat.length>1).length;
  const agtCheck=VEND.filter(r=>r.agente&&r.agente.length>1).length;
  const regCheck=VEND.filter(r=>r.regione&&r.regione.length>1).length;
  console.log('[PezzaliApp v3] VEND:',VEND.length,'con cat:',catCheck,'con agente:',agtCheck,'con regione:',regCheck);
  if(catCheck===0) console.warn('[PezzaliApp] ATTENZIONE: nessuna categoria trovata! Col V =',CV.cat,'esempio valore:',vRaw[0]?.[CV.cat]);

  // Ordini
  // Ordini: stessa logica per lettera colonna
  // Ordini Excel: M=cliente, U=descrizione, Y=qtyInevasa, Z=importoInevaso, F=trasp, AB=consegna
  const CO={
    cliente: 'M',   // CLIENTE
    desc:    'U',   // DESCRIZIONE
    qtyI:    'Y',   // QTA INEVASA
    importoI:'Z',   // IMPORTO INEVASO
    trasp:   'F',   // SPESE DI TRASPORTO
    consegna:'AB',  // DATA CONSEGNA
  };
  const ORDINI=oRaw.map(r=>{
    let consegna=null;
    try{const d=r[CO.consegna];consegna=d instanceof Date?d:new Date(d);if(isNaN(consegna))consegna=null;}catch(e){}
    return {cliente:str(r[CO.cliente]),desc:str(r[CO.desc]),qtyI:num(r[CO.qtyI]),
      importoI:num(r[CO.importoI]),trasp:num(r[CO.trasp]),consegna};
  });

  const anni=[...new Set(VEND.map(r=>r.anno))].sort();
  const agenti=[...new Set(VEND.filter(r=>r.agente).map(r=>r.agente))].sort();

  // Popola select filtri
  const selDa=document.getElementById('f-da'),selA=document.getElementById('f-a');
  const selCmp=document.getElementById('f-cmp-a');
  const selT1=document.getElementById('tr-a1'),selT2=document.getElementById('tr-a2');
  [selDa,selA,selCmp,selT1,selT2].forEach(s=>s.innerHTML='');
  anni.forEach(a=>[selDa,selA,selCmp,selT1,selT2].forEach(s=>s.insertAdjacentHTML('beforeend',`<option value="${a}">${a}</option>`)));
  selDa.value=anni[0];selA.value=anni[anni.length-1];
  selCmp.value=anni[anni.length-2]||anni[0];
  selT1.value=anni[anni.length-2]||anni[0];selT2.value=anni[anni.length-1];

  document.getElementById('f-agt').innerHTML='<option value="">Tutti</option>';
  agenti.forEach(a=>document.getElementById('f-agt').insertAdjacentHTML('beforeend',`<option value="${a}">${a}</option>`));

  const cliRank=groupBy(VEND,r=>r.cliente,rows=>sum(rows,r=>r.importo));
  const topCli=Object.entries(cliRank).sort((a,b)=>b[1]-a[1]).slice(0,60).map(([k])=>k);
  const cliSel=document.getElementById('cli-sel');
  cliSel.innerHTML='';
  topCli.forEach(c=>cliSel.insertAdjacentHTML('beforeend',`<option value="${c}">${c}</option>`));

  G={VEND,RESI,ORDINI,anni,agenti,listinoMap,CV,budgets:loadBudgets()};
  renderBudgetManager();
  applyFilters();
}

function applyFilters(){
  const annoDa=parseInt(document.getElementById('f-da').value)||0;
  const annoA=parseInt(document.getElementById('f-a').value)||9999;
  const perVal=document.getElementById('f-per').value;
  const agente=document.getElementById('f-agt').value;
  const brandF=(document.getElementById('f-brand')||{}).value||'';
  let mesiOk=null;
  if(perVal.startsWith('q')){const q=parseInt(perVal[1])-1;mesiOk=[q*3,q*3+1,q*3+2];}
  else if(perVal.startsWith('m')){mesiOk=[parseInt(perVal.slice(1))-1];}

  const filt=r=>{
    if(r.anno<annoDa||r.anno>annoA)return false;
    if(mesiOk&&!mesiOk.includes(r.mese))return false;
    if(agente&&r.agente!==agente)return false;
    if(brandF){
      if(brandF==='__no'&&r.brand!=='__no')return false;
      if(brandF!=='__no'&&r.brand.toLowerCase()!==brandF.toLowerCase())return false;
    }
    return true;
  };
  F.vend=G.VEND.filter(filt);
  F.vendCmp=null;
  if(CMP){
    const cmpAnno=parseInt(document.getElementById('f-cmp-a').value)||0;
    F.vendCmp=G.VEND.filter(r=>{if(r.anno!==cmpAnno)return false;if(mesiOk&&!mesiOk.includes(r.mese))return false;if(agente&&r.agente!==agente)return false;return true;});
  }
  let plabel=`${annoDa}–${annoA}`;
  if(perVal.startsWith('q'))plabel+=` Q${perVal[1]}`;
  else if(perVal.startsWith('m'))plabel+=` ${MESI[parseInt(perVal.slice(1))-1]}`;
  if(agente)plabel+=` · ${agente}`;
  if(brandF&&brandF!=='__no')plabel+=` · ${brandF}`;
  F.label=plabel;
  document.getElementById('sb-period').textContent=plabel;
  renderAll();
}
function toggleCompare(){CMP=!CMP;document.getElementById('chip-cmp').classList.toggle('on',CMP);document.getElementById('f-cmp-a').style.display=CMP?'inline-block':'none';applyFilters();}
function renderAll(){renderOverview();renderTrend();renderVendite();renderClienti();renderAgenti();renderSconti();renderMargine();renderTrasporti();renderOrdini();renderCriticita();}
function initDashboard(){renderAll();go('overview');}

// ═══════════════════════════════════════════════════════
//  RENDERERS
// ═══════════════════════════════════════════════════════

function renderOverview(){
  const V=F.vend,C=tc();
  const fattTot=sum(V,r=>r.importo),traspTot=sum(V,r=>r.trasp);
  const aF=parseInt(document.getElementById('f-a').value);
  const fCur=sum(V.filter(r=>r.anno===aF),r=>r.importo);
  const fPrev=sum(V.filter(r=>r.anno===aF-1),r=>r.importo);
  const delta=fPrev>0?(fCur-fPrev)/fPrev:null;
  const sR=V.filter(r=>r.sconto!==null);
  const sMed=sR.length?avg(sR,r=>r.sconto):null;
  const over60=sR.filter(r=>r.sconto>SCONTO_MAX);
  document.getElementById('ov-sub').textContent=`${F.label} · ${V.length.toLocaleString('it')} righe vendita`;
  const budgetMap=getActiveBudgetMap();
  const budgetTot=Object.values(budgetMap).reduce((a,b)=>a+b,0);
  const budDelta=budgetTot>0?(fattTot-budgetTot)/budgetTot:null;
  kpi('kr-ov',[
    {l:'Fatturato Periodo',v:fmt(fattTot),col:'g',sub:F.label},
    {l:`Fatturato ${aF}`,v:fmt(fCur),col:'g',delta,sub:`vs ${aF-1}`},
    {l:'Budget Periodo',v:budgetTot?fmt(budgetTot):'—',col:'b',sub:budgetTot?`${Object.keys(budgetMap).length} commerciali`:'nessun budget'},
    {l:'Scostamento vs Budget',v:budgetTot?fmt(fattTot-budgetTot):'—',col:budDelta!==null&&budDelta<0?'r':'g',sub:budDelta!==null?pct(Math.abs(budDelta))+(budDelta>=0?' sopra':' sotto'):'—'},
    {l:'Spese Trasporto',v:fmt(traspTot),col:'p',sub:pct(fattTot>0?traspTot/fattTot:0)+' del fatturato'},
    {l:'Sconto Medio',v:sMed!==null?pct(sMed):'N/D',col:sMed>SCONTO_MAX?'r':'g',sub:'su articoli listino'},
    {l:'Righe >60%',v:over60.length.toLocaleString('it'),col:over60.length>0?'r':'g',sub:'oltre soglia'},
    {l:'Ordini Inevasi',v:fmt(sum(G.ORDINI,r=>r.importoI)),col:'a',sub:`${G.ORDINI.length} righe`},
  ]);
  // Annual chart
  const anni=G.anni;
  const fAn=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAn=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const incArr=anni.map(a=>(tAn[a]||0)/(fAn[a]||1)*100);
  const peakAnno=anni[incArr.indexOf(Math.max(...incArr))];
  dc('ch-annual');
  const ctx=document.getElementById('ch-annual').getContext('2d');
  charts['ch-annual']=new Chart(ctx,{
    data:{labels:anni,datasets:[
      {type:'bar',label:'Fatturato €',data:anni.map(a=>fAn[a]||0),backgroundColor:anni.map(a=>a===aF?C.green+'cc':C.green+'55'),borderRadius:4,yAxisID:'y'},
      {type:'line',label:'Incidenza Trasp %',data:incArr,borderColor:C.red,backgroundColor:C.red+'15',tension:.3,pointRadius:anni.map(a=>a===peakAnno?7:4),pointBackgroundColor:anni.map(a=>a===peakAnno?C.red:C.red+'80'),fill:false,yAxisID:'y2'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,labels:{color:C.text2,font:{size:10,family:'DM Sans'},boxWidth:10,padding:8}},
        tooltip:{backgroundColor:C.tip.bg,borderColor:C.tip.border,borderWidth:1,titleColor:C.text2,bodyColor:C.text2},
        annotation:{annotations:{peak:{type:'label',xValue:peakAnno,yValue:Math.max(...incArr),yScaleID:'y2',
          backgroundColor:C.red+'dd',color:'#fff',font:{size:9},content:[`Picco ${peakAnno}`,`${Math.max(...incArr).toFixed(1)}%`],padding:4,borderRadius:4}}}},
      scales:{x:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9,family:'DM Sans'}}},
        y:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>fmtS(v)}},
        y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:C.red,font:{size:9},callback:v=>v.toFixed(1)+'%'}}}}
  });
  // Pie categorie
  const catFatt=groupBy(V.filter(r=>r.cat&&r.cat.length>1),r=>r.cat,rows=>sum(rows,r=>r.importo));
  const catS=Object.entries(catFatt).sort((a,b)=>b[1]-a[1]).slice(0,9);
  doPie('ch-pie',catS.map(([k])=>k.split(' - ')[0]),catS.map(([,v])=>v));
  // Quarterly
  const trimMap={};
  G.VEND.forEach(r=>{if(r.trim>=0){const k=`${r.anno}-Q${r.trim+1}`;trimMap[k]=(trimMap[k]||0)+r.importo;}});
  const tLbls=[],tData=[];
  G.anni.forEach(a=>QNAMES.forEach((_,q)=>{tLbls.push(`${a} Q${q+1}`);tData.push(trimMap[`${a}-Q${q+1}`]||0);}));
  doBar('ch-qtr',tLbls,tData,[C.blue+'aa'],null,{maxTicksLimit:16});
  // Top clients
  const cliF=groupBy(V,r=>r.cliente,rows=>sum(rows,r=>r.importo));
  const top10=Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,10);
  doHBar('ch-top-cli',top10.map(([k])=>trunc(k,24)),top10.map(([,v])=>v),C.green);
}

function renderTrend(){
  const C=tc();
  const a1=parseInt(document.getElementById('tr-a1').value);
  const a2=parseInt(document.getElementById('tr-a2').value);
  const view=document.getElementById('tr-view').value;
  const met=document.getElementById('tr-met').value;
  const labels=view==='m'?MESI:QNAMES;
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
  document.getElementById('tr-title').textContent=`Confronto ${view==='m'?'Mensile':'Trimestrale'} — ${{f:'Fatturato',t:'Trasporto',s:'Sconto %',n:'N° Righe'}[met]}`;
  document.getElementById('tr-sub').textContent=`${a1} vs ${a2}`;
  document.getElementById('tr-delta').innerHTML=dPct!==null?`<span class="dt ${dPct>=0?'up':'dn'}">${dPct>=0?'↑':'↓'}${pct(Math.abs(dPct))}</span>`:'';
  const fmtY=met==='f'||met==='t'?v=>fmtS(v):met==='s'?v=>v.toFixed(1)+'%':v=>v;
  dc('ch-trend');
  charts['ch-trend']=new Chart(document.getElementById('ch-trend').getContext('2d'),{
    data:{labels,datasets:[
      {type:'bar',label:`${a1}`,data:d1,backgroundColor:C.blue+'66',borderRadius:3},
      {type:'line',label:`${a2}`,data:d2,borderColor:C.green,tension:.3,pointRadius:5,fill:false,
       pointBackgroundColor:d2.map((v,i)=>v>d1[i]?C.green:C.red)}
    ]},
    options:{...chartOpts({callbackY:fmtY,legend:true,C})}
  });
  const cum1=[],cum2=[];
  d1.reduce((acc,v,i)=>{cum1[i]=acc+v;return acc+v;},0);
  d2.reduce((acc,v,i)=>{cum2[i]=acc+v;return acc+v;},0);
  dc('ch-cumul');
  charts['ch-cumul']=new Chart(document.getElementById('ch-cumul').getContext('2d'),{
    data:{labels,datasets:[
      {type:'line',label:`${a1}`,data:cum1,borderColor:C.blue,fill:true,backgroundColor:C.blue+'15',tension:.3,pointRadius:3},
      {type:'line',label:`${a2}`,data:cum2,borderColor:C.green,fill:true,backgroundColor:C.green+'15',tension:.3,pointRadius:3}
    ]},options:{...chartOpts({legend:true,callbackY:v=>fmtS(v),C})}
  });
  const deltas=d1.map((v,i)=>d1[i]>0?((d2[i]-d1[i])/d1[i])*100:0);
  dc('ch-delta');
  charts['ch-delta']=new Chart(document.getElementById('ch-delta').getContext('2d'),{type:'bar',
    data:{labels,datasets:[{data:deltas,backgroundColor:deltas.map(d=>d>=0?C.green+'aa':C.red+'aa'),borderRadius:3}]},
    options:{...chartOpts({callbackY:v=>v.toFixed(1)+'%',C})}
  });
}

function renderVendite(){
  const V=F.vend,C=tc();
  // Filtra righe con categoria valida
  const VwCat=V.filter(r=>r.cat&&r.cat.length>1);
  const catFatt=groupBy(VwCat,r=>r.cat,rows=>({
    f:sum(rows,r=>r.importo),q:sum(rows,r=>r.qty),n:rows.length,
    sc:avg(rows.filter(r=>r.sconto!==null),r=>r.sconto),
    // Top linea prodotto per questa categoria
    topLinea:Object.entries(groupBy(rows,r=>r.sottocat,r=>sum(r,x=>x.importo))).sort((a,b)=>b[1]-a[1])[0]?.[0]||'',
  }));
  const cats=Object.entries(catFatt).sort((a,b)=>b[1].f-a[1].f);
  const totF=sum(V,r=>r.importo);
  kpi('kr-vend',[
    {l:'Fatturato',v:fmt(totF),col:'g'},
    {l:'Righe Vendita',v:V.length.toLocaleString('it'),col:'b'},
    {l:'Prezzo Netto Medio',v:fmt(avg(V,r=>r.pz)),col:'g',sub:'per unità'},
    {l:'Categorie Attive',v:cats.length,col:'p'},
    {l:'Ticket Medio Riga',v:fmt(V.length?totF/V.length:0),col:'b'},
  ]);
  const PAL=[C.green+'aa',C.blue+'aa',C.amber+'aa',C.purple+'aa',C.red+'aa',C.cyan+'aa',C.green+'66',C.blue+'66',C.amber+'66',C.purple+'66',C.red+'66',C.cyan+'66'];
  dc('ch-cat-bar');
  charts['ch-cat-bar']=new Chart(document.getElementById('ch-cat-bar').getContext('2d'),{type:'bar',
    data:{labels:cats.map(([k])=>trunc(k.split(' - ')[0],20)),datasets:[{data:cats.map(([,v])=>v.f),backgroundColor:PAL,borderRadius:4}]},
    options:{...chartOpts({callbackY:v=>fmtS(v),C}),onClick:(_,els)=>{if(!els.length)return;showDrill(cats[els[0].index][0],V);}}
  });
  doBar('ch-cat-qty',cats.map(([k])=>trunc(k.split(' - ')[0],20)),cats.map(([,v])=>v.q),[C.purple+'aa'],null);
  tbl('tbl-cat',['Categoria','Fatturato','%','Pezzi','Top Linea Prodotto','Sconto Medio'],
    cats.map(([k,v])=>[trunc(k,30),`<span class="mono">${fmt(v.f)}</span>`,
      `<span class="bdg bg">${pct(totF>0?v.f/totF:0)}</span>`,
      `<span class="mono">${Math.round(v.q).toLocaleString('it')}</span>`,
      trunc(v.topLinea,22),
      v.sc>0?`<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${pct(v.sc)}</span>`:'—'])
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

let cliAllData=[];
function renderClienti(){
  const V=F.vend,C=tc();
  const cliData={};
  V.forEach(r=>{
    if(!cliData[r.cliente])cliData[r.cliente]={f:0,n:0,tr:0,dests:new Set(),sc:[]};
    cliData[r.cliente].f+=r.importo;cliData[r.cliente].n++;cliData[r.cliente].tr+=r.trasp;
    if(r.dest&&r.dest!==r.cliente&&r.dest.length>1)cliData[r.cliente].dests.add(r.dest);
    if(r.sconto!==null)cliData[r.cliente].sc.push(r.sconto);
  });
  cliAllData=Object.entries(cliData).map(([k,v])=>({
    nome:k,f:v.f,n:v.n,tr:v.tr,dests:[...v.dests],
    sc:v.sc.length?v.sc.reduce((a,b)=>a+b,0)/v.sc.length:null
  })).sort((a,b)=>b.f-a.f);
  doHBar('ch-cli-rank',cliAllData.slice(0,12).map(c=>trunc(c.nome,24)),cliAllData.slice(0,12).map(c=>c.f),C.green);
  renderCliStorico();
  filterCliTbl();
}
function renderCliStorico(){
  const sel=document.getElementById('cli-sel').value;if(!sel)return;
  const C=tc();
  const rows=G.VEND.filter(r=>r.cliente===sel&&r.mese>=0);
  const anni=[...new Set(rows.map(r=>r.anno))].sort();
  const colors=[C.green,C.blue,C.amber,C.purple,C.red,C.cyan];
  dc('ch-cli-st');
  charts['ch-cli-st']=new Chart(document.getElementById('ch-cli-st').getContext('2d'),{
    data:{labels:MESI,datasets:anni.map((a,i)=>({type:'line',label:`${a}`,
      data:MESI.map((_,m)=>rows.filter(r=>r.anno===a&&r.mese===m).reduce((s,r)=>s+r.importo,0)),
      borderColor:colors[i%colors.length],backgroundColor:colors[i%colors.length]+'15',tension:.3,pointRadius:3,fill:false}))},
    options:{...chartOpts({legend:true,callbackY:v=>fmtS(v),C})}
  });
}
function filterCliTbl(){
  const q=document.getElementById('cli-srch').value.toLowerCase();
  const df=document.getElementById('cli-dest-filter').value;
  let rows=cliAllData.filter(c=>c.nome.toLowerCase().includes(q)||c.dests.some(d=>d.toLowerCase().includes(q)));
  if(df==='multi')rows=rows.filter(c=>c.dests.length>0);
  if(df==='single')rows=rows.filter(c=>c.dests.length===0);
  const totF=sum(F.vend,r=>r.importo);
  tbl('tbl-cli',['Cliente','Fatturato','%','Destinazioni diverse','Sconto Medio','Trasporto'],
    rows.slice(0,80).map(c=>{
      const destPills=c.dests.length>0
        ?`<div class="dest-list">${c.dests.slice(0,4).map(d=>`<span class="dest-pill">${trunc(d,22)}</span>`).join('')}${c.dests.length>4?`<span class="dest-pill">+${c.dests.length-4}</span>`:''}</div>`
        :`<span class="mono" style="color:var(--text3)">sede unica</span>`;
      return [c.nome,`<span class="mono">${fmt(c.f)}</span>`,`<span class="bdg bg">${pct(totF>0?c.f/totF:0)}</span>`,
        destPills,c.sc!==null?`<span class="bdg ${c.sc>SCONTO_MAX?'br':c.sc>0.5?'ba':'bg'}">${pct(c.sc)}</span>`:'—',`<span class="mono">${fmt(c.tr)}</span>`];
    })
  );
}

function renderAgenti(){
  const V=F.vend,C=tc();
  const budgetMap=getActiveBudgetMap();
  const VwAgt=V.filter(r=>r.agente&&r.agente.length>1);
  const agtF=groupBy(VwAgt,r=>r.agente,rows=>({f:sum(rows,r=>r.importo),n:rows.length,tr:sum(rows,r=>r.trasp)}));
  const names=[...new Set([...Object.keys(agtF),...Object.keys(budgetMap)])].sort((a,b)=>(agtF[b]?.f||0)-(agtF[a]?.f||0)||a.localeCompare(b,'it'));
  const rows=names.map(k=>{
    const actual=agtF[k]?.f||0, budget=budgetMap[k]||0, n=agtF[k]?.n||0, tr=agtF[k]?.tr||0;
    const delta=actual-budget, attain=budget>0?actual/budget:null;
    return {agente:k,actual,budget,delta,attain,n,tr};
  }).sort((a,b)=>b.actual-a.actual||b.budget-a.budget||a.agente.localeCompare(b.agente,'it'));
  const totF=sum(V,r=>r.importo), budgetTot=rows.reduce((a,r)=>a+r.budget,0);
  const attainTot=budgetTot>0?totF/budgetTot:null;
  kpi('kr-agt',[
    {l:'Fatturato Totale',v:fmt(totF),col:'g',sub:`${rows.length} commerciali`},
    {l:'Budget Totale',v:budgetTot?fmt(budgetTot):'—',col:'b',sub:budgetTot?'periodo filtrato':'nessun budget'},
    {l:'Scostamento',v:budgetTot?fmt(totF-budgetTot):'—',col:budgetTot&&totF<budgetTot?'r':'g',sub:attainTot!==null?pct(Math.abs(attainTot-1))+(attainTot>=1?' sopra':' sotto'):'—'},
    {l:'Raggiungimento',v:attainTot!==null?pct(attainTot):'—',col:attainTot!==null&&attainTot<1?'a':'g',sub:'fatturato / budget'},
  ]);
  dc('ch-agt-bar');
  charts['ch-agt-bar']=new Chart(document.getElementById('ch-agt-bar').getContext('2d'),{
    data:{labels:rows.map(r=>r.agente),datasets:[
      {type:'bar',label:'Fatturato',data:rows.map(r=>r.actual),backgroundColor:C.green+'bb',borderRadius:4},
      {type:'bar',label:'Budget',data:rows.map(r=>r.budget),backgroundColor:C.blue+'88',borderRadius:4}
    ]},
    options:{...chartOpts({legend:true,callbackY:v=>fmtS(v),C})}
  });
  const anni=[...new Set([...(G.anni||[]),...(G.budgets||[]).map(r=>r.year)])].sort((a,b)=>a-b);
  const colors=[C.green,C.blue,C.amber,C.purple,C.red,C.cyan];
  dc('ch-agt-evol');
  charts['ch-agt-evol']=new Chart(document.getElementById('ch-agt-evol').getContext('2d'),{
    data:{labels:anni,datasets:rows.slice(0,6).flatMap((r,i)=>([
      {type:'line',label:`${r.agente} fatt.`,data:anni.map(a=>sum(G.VEND.filter(x=>x.anno===a&&x.agente===r.agente),x=>x.importo)),
      borderColor:colors[i%colors.length],tension:.3,pointRadius:3,fill:false},
      {type:'line',label:`${r.agente} budget`,data:anni.map(a=>(G.budgets||[]).filter(x=>x.year===a&&x.agente===r.agente).reduce((s,x)=>s+x.budget,0)),
      borderColor:colors[i%colors.length],borderDash:[6,4],tension:.3,pointRadius:2,fill:false}
    ]))},
    options:{...chartOpts({legend:true,callbackY:v=>fmtS(v),C})}
  });
  const agEl=document.getElementById('agr');agEl.innerHTML='';
  const maxBase=Math.max(1,...rows.map(r=>Math.max(r.actual,r.budget,1)));
  rows.forEach((r,i)=>{
    const tone=r.budget>0&&r.actual<r.budget?'ba':'bg';
    agEl.insertAdjacentHTML('beforeend',`
      <div class="agrow"><div class="agrow-top">
        <span class="agno">${String(i+1).padStart(2,'0')}</span>
        <span class="agname">${r.agente}</span>
        <span class="agval">${fmt(r.actual)}</span>
        <span class="bdg bb" style="font-size:8px">B ${r.budget?fmtS(r.budget):'—'}</span>
        <span class="bdg ${tone}" style="font-size:8px">${r.attain!==null?pct(r.attain):'—'}</span>
      </div>
      <div class="agbar"><div class="agfill" style="width:${Math.round(Math.max(r.actual,r.budget)/maxBase*100)}%"></div></div>
      </div>`);
  });
  tbl('tbl-agt-report',['Agente','Fatturato','Budget','Scostamento','Raggiungimento','Righe','Trasporto'],
    rows.map(r=>[
      r.agente,
      `<span class="mono">${fmt(r.actual)}</span>`,
      r.budget?`<span class="mono">${fmt(r.budget)}</span>`:'—',
      r.budget?`<span class="mono">${fmt(r.delta)}</span>`:'—',
      r.attain!==null?`<span class="bdg ${r.attain<1?'ba':'bg'}">${pct(r.attain)}</span>`:'—',
      `<span class="mono">${r.n}</span>`,
      `<span class="mono">${fmt(r.tr)}</span>`
    ])
  );
  renderBudgetManager();
}

let scontiData=[];
function renderSconti(){
  const V=F.vend,C=tc();
  const sR=V.filter(r=>r.sconto!==null);
  const sMed=sR.length?avg(sR,r=>r.sconto):0;
  const over60=sR.filter(r=>r.sconto>SCONTO_MAX);
  kpi('kr-sc',[
    {l:'Sconto Medio',v:pct(sMed),col:sMed>SCONTO_MAX?'r':'g',sub:`su ${sR.length} righe listino`},
    {l:'Righe >60%',v:over60.length.toLocaleString('it'),col:'r',sub:fmt(sum(over60,r=>r.importo))},
    {l:'Soglia Max',v:pct(SCONTO_MAX),col:'a',sub:'contrattuale rivenditori'},
    {l:'Copertura Listino',v:pct(V.length?sR.length/V.length:0),col:'b',sub:`${sR.length}/${V.length} righe`},
  ]);
  const anni=G.anni;
  const scAnno=anni.map(a=>{const r2=G.VEND.filter(r=>r.anno===a&&r.sconto!==null);return r2.length?avg(r2,r=>r.sconto)*100:0;});
  dc('ch-sc-yr');
  charts['ch-sc-yr']=new Chart(document.getElementById('ch-sc-yr').getContext('2d'),{
    data:{labels:anni,datasets:[
      {type:'line',label:'Sconto %',data:scAnno,borderColor:C.purple,backgroundColor:C.purple+'15',tension:.3,fill:true,
       pointRadius:scAnno.map(v=>v>60?7:4),pointBackgroundColor:scAnno.map(v=>v>60?C.red:C.purple)},
      {type:'line',label:'Soglia 60%',data:anni.map(()=>60),borderColor:C.red,borderDash:[6,4],pointRadius:0,borderWidth:1.5}
    ]},
    options:{...chartOpts({legend:true,callbackY:v=>v.toFixed(1)+'%',C})}
  });
  const bkts=Array(10).fill(0);
  sR.forEach(r=>{bkts[Math.min(9,Math.floor(r.sconto*100/10))]++;});
  doBar('ch-sc-dist',['0–10%','10–20%','20–30%','30–40%','40–50%','50–60%','60–70%','70–80%','80–90%','90–100%'],
    bkts,null,bkts.map((_,i)=>i>=6?C.red+'aa':i>=5?C.amber+'aa':C.green+'88'));
  const prodSc=groupBy(sR,r=>r.desc,rows=>({sc:avg(rows,r=>r.sconto),n:rows.length,f:sum(rows,r=>r.importo),pz:avg(rows,r=>r.pz),lordo:rows[0].lordo}));
  scontiData=Object.entries(prodSc).sort((a,b)=>b[1].f-a[1].f);
  renderScontiTbl();
  const ov60c=groupBy(over60,r=>r.cliente,rows=>({n:rows.length,f:sum(rows,r=>r.importo),sc:avg(rows,r=>r.sconto)}));
  tbl('tbl-over60-cli',['Cliente','Righe >60%','Valore','Sconto Medio'],
    Object.entries(ov60c).sort((a,b)=>b[1].n-a[1].n).slice(0,20).map(([k,v])=>[
      k,`<span class="bdg br">${v.n}</span>`,`<span class="mono">${fmt(v.f)}</span>`,
      `<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${pct(v.sc)}</span>`]));
}
function renderScontiTbl(){
  const flt=document.getElementById('sc-flt').value;
  const q=document.getElementById('sc-srch').value.toLowerCase();
  let rows=scontiData.filter(([k])=>k.toLowerCase().includes(q));
  if(flt==='over')rows=rows.filter(([,v])=>v.sc>SCONTO_MAX);
  if(flt==='ok')rows=rows.filter(([,v])=>v.sc<=SCONTO_MAX);
  tbl('tbl-sc',['Prodotto/Macchina','Sconto Medio','N°','Fatturato','PZ Netto','Lordo'],
    rows.slice(0,80).map(([k,v])=>[trunc(k,34),
      `<span class="bdg ${v.sc>SCONTO_MAX?'br':v.sc>0.5?'ba':'bg'}">${v.sc>SCONTO_MAX?'⚠ ':''}${pct(v.sc)}</span>`,
      v.n,`<span class="mono">${fmt(v.f)}</span>`,`<span class="mono">${fmt(v.pz)}</span>`,
      v.lordo?`<span class="mono">${fmt(v.lordo)}</span>`:'—']));
}

function renderMargine(){
  const V=F.vend,C=tc();
  // Tutte le categorie — sc_eff usa 60% come fallback per categorie senza listino
  const catMarg=groupBy(V.filter(r=>r.cat&&r.cat.length>1),r=>r.cat,rows=>({
    sc:rows.filter(r=>r.sconto!==null).length>0?avg(rows.filter(r=>r.sconto!==null),r=>r.sconto):null,
    sc_eff:avg(rows,r=>r.sconto_eff),
    tr:avg(rows,r=>r.incTrasp),
    f:sum(rows,r=>r.importo),n:rows.length
  }));
  const sorted=Object.entries(catMarg).map(([k,v])=>[k,{...v,erosione:v.sc_eff+v.tr}])
    .filter(([k])=>k&&k.length>1).sort((a,b)=>b[1].f-a[1].f);
  // Heatmap
  const hmEl=document.getElementById('hm-margine');hmEl.innerHTML='';
  sorted.forEach(([k,v])=>{
    const col=v.erosione>0.8?C.red:v.erosione>0.65?C.amber:v.sc===null?C.blue:C.green;
    hmEl.insertAdjacentHTML('beforeend',`
      <div class="hmc">
        <div class="hmn">${trunc(k.split(' - ')[0],22)}</div>
        <div class="hmv" style="color:${col}">${pct(v.erosione)}</div>
        <div class="hms">${v.sc!==null?'Sc '+pct(v.sc):'Sc ~60%*'} + Tr ${pct(v.tr)}</div>
        <div class="hmbar"><div class="hmfill" style="width:${Math.min(100,v.erosione*100)}%;background:${col}"></div></div>
      </div>`);
  });
  const erosColors=sorted.map(([,v])=>v.erosione>0.8?C.red+'bb':v.erosione>0.65?C.amber+'bb':v.sc===null?C.blue+'88':C.green+'88');
  doHBar('ch-eros',sorted.map(([k])=>trunc(k.split(' - ')[0],20)),sorted.map(([,v])=>v.erosione*100),null,erosColors);
  const maxF=Math.max(...sorted.map(([,v])=>v.f));
  dc('ch-scatter');
  charts['ch-scatter']=new Chart(document.getElementById('ch-scatter').getContext('2d'),{type:'bubble',
    data:{datasets:[{data:sorted.map(([,v])=>({x:v.sc_eff*100,y:v.tr*100,r:Math.max(4,Math.min(20,v.f/maxF*18))})),
      backgroundColor:erosColors,borderColor:'transparent'}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{backgroundColor:C.tip.bg,borderColor:C.tip.border,borderWidth:1,
        titleColor:C.text2,bodyColor:C.text2,
        callbacks:{label:ctx=>`${sorted[ctx.dataIndex]?.[0]?.split(' - ')[0]||''}: Sc ${ctx.parsed.x.toFixed(1)}% + Tr ${ctx.parsed.y.toFixed(1)}%`}}},
      scales:{x:{title:{display:true,text:'Sconto %',color:C.text3,font:{size:9}},grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>v+'%'}},
        y:{title:{display:true,text:'Incidenza Trasp %',color:C.text3,font:{size:9}},grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>v+'%'}}}}
  });
  tbl('tbl-marg',['Categoria','Erosione','Sc. Reale','Sc. Usato','Trasp %','Fatturato','N°'],
    sorted.map(([k,v])=>[trunc(k,28),
      `<span class="bdg ${v.erosione>0.8?'br':v.erosione>0.65?'ba':'bg'}">${pct(v.erosione)}</span>`,
      v.sc!==null?`<span class="mono">${pct(v.sc)}</span>`:'<span class="ba bdg">~60%</span>',
      `<span class="mono">${pct(v.sc_eff)}</span>`,`<span class="mono">${pct(v.tr)}</span>`,
      `<span class="mono">${fmt(v.f)}</span>`,v.n]));
}

function renderTrasporti(){
  const V=F.vend,C=tc(),anni=G.anni;
  const fAn=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAn=groupBy(G.VEND,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const inc=anni.map(a=>(tAn[a]||0)/(fAn[a]||1));
  const trTot=sum(V,r=>r.trasp),fTot=sum(V,r=>r.importo),incMed=fTot>0?trTot/fTot:0;
  const piccoIdx=inc.indexOf(Math.max(...inc));const picco=anni[piccoIdx];
  kpi('kr-tr',[
    {l:'Spese Trasporto',v:fmt(trTot),col:'p',sub:`${pct(incMed)} sul fatturato`},
    {l:'Incidenza Media',v:pct(incMed),col:incMed>0.05?'r':'g'},
    {l:'Anno Picco',v:`${picco}`,col:'a',sub:`${pct(inc[piccoIdx])} incidenza`},
    {l:'Trasp. Ordini',v:fmt(sum(G.ORDINI,r=>r.trasp)),col:'p',sub:`${G.ORDINI.length} ordini`},
  ]);
  dc('ch-tr-yr');
  charts['ch-tr-yr']=new Chart(document.getElementById('ch-tr-yr').getContext('2d'),{
    data:{labels:anni,datasets:[
      {type:'bar',label:'Trasporto €',data:anni.map(a=>tAn[a]||0),backgroundColor:anni.map(a=>a===picco?C.red+'cc':C.purple+'88'),borderRadius:4,yAxisID:'y'},
      {type:'line',label:'Incidenza %',data:inc.map(v=>v*100),borderColor:C.amber,tension:.3,pointRadius:4,fill:false,yAxisID:'y2',
       pointBackgroundColor:inc.map((v,i)=>anni[i]===picco?C.red:C.amber)}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,labels:{color:C.text2,font:{size:10},boxWidth:10,padding:8}},
        tooltip:{backgroundColor:C.tip.bg,borderColor:C.tip.border,borderWidth:1,titleColor:C.text2,bodyColor:C.text2},
        annotation:{annotations:{peak:{type:'label',xValue:picco,yValue:inc[piccoIdx]*100,yScaleID:'y2',
          backgroundColor:C.red+'dd',color:'#fff',font:{size:9},content:[`Picco ${picco}`,`${(inc[piccoIdx]*100).toFixed(1)}%`],padding:4,borderRadius:4}}}},
      scales:{x:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9}}},
        y:{grid:{color:C.border},ticks:{color:C.text3,font:{size:9},callback:v=>fmtS(v)}},
        y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:C.amber,font:{size:9},callback:v=>v.toFixed(1)+'%'}}}}
  });
  // Porto
  const portD={};V.forEach(r=>{const p=r.porto_desc||'Altro';portD[p]=(portD[p]||0)+1;});
  doPie('ch-porto',Object.keys(portD),Object.values(portD));
  // Anno table
  tbl('tbl-tr',['Anno','Fatturato','Trasporto','Incidenza','Δ'],
    anni.map((a,i)=>{const dlt=i>0?inc[i]-inc[i-1]:null;
      return [a,`<span class="mono">${fmtS(fAn[a]||0)}</span>`,`<span class="mono">${fmtS(tAn[a]||0)}</span>`,
        `<span class="bdg ${inc[i]>0.05?'br':'bg'}">${pct(inc[i])}</span>`,
        dlt!==null?`<span class="bdg ${dlt>0?'br':'bg'}">${dlt>0?'+':''}${pct(dlt)}</span>`:'—'];
    }));
  // Trasporti per Regione
  const trReg=groupBy(V.filter(r=>r.regione&&r.regione.length>1),r=>r.regione,rows=>({f:sum(rows,r=>r.importo),tr:sum(rows,r=>r.trasp)}));
  const regS=Object.entries(trReg).map(([k,v])=>({k,f:v.f,tr:v.tr,inc:v.f>0?v.tr/v.f:0})).sort((a,b)=>b.tr-a.tr);
  dc('ch-tr-reg');
  if(regS.length>0){
    charts['ch-tr-reg']=new Chart(document.getElementById('ch-tr-reg').getContext('2d'),{type:'bar',
      data:{labels:regS.map(r=>r.k),datasets:[{data:regS.map(r=>r.tr),
        backgroundColor:regS.map(r=>r.inc>0.08?C.red+'aa':r.inc>0.05?C.amber+'aa':C.purple+'88'),borderRadius:3}]},
      options:{...chartOpts({callbackY:v=>fmtS(v),C}),indexAxis:'y'}
    });
  }
  tbl('tbl-tr-reg',['Regione','Fatturato','Trasporto','Incidenza'],
    regS.slice(0,20).map(r=>[r.k,`<span class="mono">${fmtS(r.f)}</span>`,`<span class="mono">${fmtS(r.tr)}</span>`,
      `<span class="bdg ${r.inc>0.08?'br':r.inc>0.05?'ba':'bg'}">${pct(r.inc)}</span>`]));
  // Trasporti per Agente
  const trAgt=groupBy(V.filter(r=>r.agente&&r.agente.length>1),r=>r.agente,rows=>({f:sum(rows,r=>r.importo),tr:sum(rows,r=>r.trasp)}));
  const agtTrS=Object.entries(trAgt).map(([k,v])=>({k,f:v.f,tr:v.tr,inc:v.f>0?v.tr/v.f:0})).sort((a,b)=>b.tr-a.tr);
  tbl('tbl-tr-agt',['Agente','Fatturato','Trasporto','Incidenza'],
    agtTrS.map(r=>[r.k,`<span class="mono">${fmtS(r.f)}</span>`,`<span class="mono">${fmtS(r.tr)}</span>`,
      `<span class="bdg ${r.inc>0.06?'br':r.inc>0.04?'ba':'bg'}">${pct(r.inc)}</span>`]));
}

let ordiniAll=[];
function renderOrdini(){
  const O=G.ORDINI,C=tc(),today=new Date();
  const scaduti=O.filter(r=>r.consegna&&r.consegna<today);
  const in30=O.filter(r=>{if(!r.consegna||r.consegna<today)return false;return(r.consegna-today)/86400000<=30;});
  ordiniAll=[...O].sort((a,b)=>(a.consegna||new Date(9999,0))-(b.consegna||new Date(9999,0)));
  kpi('kr-ord',[
    {l:'Valore Inevaso',v:fmt(sum(O,r=>r.importoI)),col:'a',sub:`${O.length} righe`},
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
  const q=document.getElementById('ord-srch').value.toLowerCase(),today=new Date();
  tbl('tbl-ord',['Cliente','Prodotto','Qtà','Importo','Consegna','Stato'],
    ordiniAll.filter(r=>(r.cliente+r.desc).toLowerCase().includes(q)).slice(0,80).map(r=>{
      const late=r.consegna&&r.consegna<today;
      const soon=r.consegna&&!late&&(r.consegna-today)/86400000<=30;
      return [trunc(r.cliente,26),trunc(r.desc,28),r.qtyI,`<span class="mono">${fmt(r.importoI)}</span>`,
        r.consegna&&!isNaN(r.consegna)?r.consegna.toLocaleDateString('it'):'—',
        late?'<span class="bdg br">SCADUTO</span>':soon?'<span class="bdg ba">≤30gg</span>':'<span class="bdg bb">OK</span>'];
    }));
}

function renderCriticita(){
  const V=G.VEND,O=G.ORDINI;
  const sR=V.filter(r=>r.sconto!==null);
  const over60=sR.filter(r=>r.sconto>SCONTO_MAX);
  const sMed=sR.length?avg(sR,r=>r.sconto):0;
  const anni=G.anni,annoMax=Math.max(...anni);
  const fAn=groupBy(V,r=>r.anno,rows=>sum(rows,r=>r.importo));
  const tAn=groupBy(V,r=>r.anno,rows=>sum(rows,r=>r.trasp));
  const today=new Date();
  const scaduti=O.filter(r=>r.consegna&&r.consegna<today);
  const in30=O.filter(r=>{if(!r.consegna||r.consegna<today)return false;return(r.consegna-today)/86400000<=30;});
  const alerts=[];
  if(over60.length){const pOver=sR.length?over60.length/sR.length:0;
    alerts.push({type:pOver>0.1?'danger':'warn',icon:'🏷️',
      t:`Sconti >60%: ${over60.length} righe (${pct(pOver)}) — €${fmtS(sum(over60,r=>r.importo))}`,
      b:`Prodotti: ${[...new Set(over60.sort((a,b)=>b.sconto-a.sconto).slice(0,3).map(r=>r.desc))].join(', ')}`});}
  const diffS=sMed-SCONTO_MAX;
  alerts.push({type:diffS>0?'warn':'ok',icon:diffS>0?'⚠️':'✅',
    t:`Sconto medio: ${pct(sMed)} (soglia: ${pct(SCONTO_MAX)})`,
    b:diffS>0?`Supera la soglia di ${pct(Math.abs(diffS))}.`:`Dentro i limiti (residuo: ${pct(Math.abs(diffS))}).`});
  const f25=fAn[2025]||0,f24=fAn[2024]||0;
  if(f24&&f25){const d=(f25-f24)/f24;alerts.push({type:d<-0.05?'warn':d>0.05?'ok':'info',icon:d>0?'📈':'📉',
    t:`Trend 2025 vs 2024: ${d>0?'+':''}${pct(d)}`,b:`2024: €${fmtS(f24)} → 2025: €${fmtS(f25)}`});}
  const incid=anni.map(a=>({a,inc:(tAn[a]||0)/(fAn[a]||1)}));
  const picco=incid.reduce((m,i)=>i.inc>m.inc?i:m,incid[0]);
  if(picco.inc>0.07)alerts.push({type:'warn',icon:'🚚',t:`Picco trasporti ${picco.a}: ${pct(picco.inc)} — €${fmtS(tAn[picco.a]||0)}`,b:'Analizzare i fattori scatenanti.'});
  if(scaduti.length)alerts.push({type:'danger',icon:'⏰',t:`${scaduti.length} ordini scaduti — €${fmtS(sum(scaduti,r=>r.importoI))}`,
    b:`Clienti: ${[...new Set(scaduti.map(r=>r.cliente).filter(Boolean))].slice(0,4).join(', ')}`});
  if(in30.length)alerts.push({type:'warn',icon:'📅',t:`${in30.length} ordini in scadenza 30gg — €${fmtS(sum(in30,r=>r.importoI))}`,b:'Pianificare priorità logistica.'});
  const budgetMap=getActiveBudgetMap();
  const budgetRows=Object.entries(budgetMap).map(([agente,budget])=>({agente,budget,actual:sum(F.vend.filter(r=>r.agente===agente),r=>r.importo)})).filter(r=>r.budget>0);
  const lowBudget=budgetRows.filter(r=>r.actual/r.budget<0.8).sort((a,b)=>(a.actual/a.budget)-(b.actual/b.budget));
  if(lowBudget.length) alerts.push({type:'warn',icon:'🎯',t:`${lowBudget.length} commerciali sotto budget`,b:lowBudget.slice(0,4).map(r=>`${r.agente} ${pct(r.actual/r.budget)}`).join(' · ')});
    const cliF=groupBy(V.filter(r=>r.anno>=2024),r=>r.cliente,rows=>sum(rows,r=>r.importo));
  const fTot2425=sum(V.filter(r=>r.anno>=2024),r=>r.importo);
  const top3=Object.entries(cliF).sort((a,b)=>b[1]-a[1]).slice(0,3);
  const top3pct=fTot2425>0?sum(top3,([,v])=>v)/fTot2425:0;
  if(top3pct>0.35)alerts.push({type:'warn',icon:'🏢',t:`Concentrazione: top 3 clienti = ${pct(top3pct)} del fatturato 2024–${annoMax}`,
    b:top3.map(([k,v])=>`${trunc(k,22)} ${pct(fTot2425>0?v/fTot2425:0)}`).join(' · ')});
  document.getElementById('nbadge').textContent=alerts.filter(a=>a.type==='danger'||a.type==='warn').length;
  document.getElementById('alerts').innerHTML=alerts.length===0
    ?'<div class="al ok"><div class="al-ic">✅</div><div class="al-b"><strong>Nessuna criticità rilevata</strong><p>Tutti gli indicatori nei range normali.</p></div></div>'
    :alerts.map(a=>`<div class="al ${a.type}"><div class="al-ic">${a.icon}</div><div class="al-b"><strong>${a.t}</strong><p>${a.b}</p></div></div>`).join('');
}

// ═══════════════════════════════════════════════════════
//  CHART FACTORY
// ═══════════════════════════════════════════════════════
function chartOpts({legend=false,callbackY=null,C}={}){
  const c=C||tc();
  const yFn=callbackY||(v=>fmtS(v));
  return{responsive:true,maintainAspectRatio:false,
    plugins:{legend:legend?{display:true,labels:{color:c.text2,font:{size:10,family:'DM Sans'},boxWidth:10,padding:8}}:{display:false},
      tooltip:{backgroundColor:c.tip.bg,borderColor:c.tip.border,borderWidth:1,titleColor:c.text2,bodyColor:c.text2,padding:10}},
    scales:{x:{grid:{color:c.border},ticks:{color:c.text3,font:{size:9,family:'DM Sans'},maxRotation:45}},
      y:{grid:{color:c.border},ticks:{color:c.text3,font:{size:9,family:'DM Sans'},callback:yFn}}}};
}
function doPie(id,labels,data){
  const C=tc(),PAL=[C.green,C.blue,C.amber,C.purple,C.red,C.cyan,C.green+'88',C.blue+'88',C.amber+'88'];
  dc(id);
  charts[id]=new Chart(document.getElementById(id).getContext('2d'),{type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:PAL,borderWidth:0,hoverOffset:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'right',labels:{color:C.text2,font:{size:9,family:'DM Sans'},boxWidth:8,padding:6}},
        tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${fmt(ctx.raw)}`}}}}});
}
function doBar(id,labels,data,colors,colorsArr,extraTick={}){
  const C=tc();dc(id);
  charts[id]=new Chart(document.getElementById(id).getContext('2d'),{type:'bar',
    data:{labels,datasets:[{data,backgroundColor:colorsArr||colors||[C.blue+'aa'],borderRadius:3}]},
    options:{...chartOpts({callbackY:v=>fmtS(v),C}),
      scales:{...chartOpts({C}).scales,x:{...chartOpts({C}).scales.x,...extraTick}}}});
}
function doHBar(id,labels,data,color,colors){
  const C=tc();dc(id);
  charts[id]=new Chart(document.getElementById(id).getContext('2d'),{type:'bar',
    data:{labels,datasets:[{data,backgroundColor:colors||(color||C.green),borderRadius:3}]},
    options:{...chartOpts({callbackY:v=>fmtS(v),C}),indexAxis:'y'}});
}
function dc(id){if(charts[id]){charts[id].destroy();delete charts[id];}}

// ═══════════════════════════════════════════════════════
//  TABLE ENGINE
// ═══════════════════════════════════════════════════════
function tbl(id,headers,rows){
  const el=document.getElementById(id);if(!el)return;
  const s=sortState[id]||{col:-1,asc:true};
  let sr=[...rows];
  if(s.col>=0){sr.sort((a,b)=>{const va=sh(a[s.col]),vb=sh(b[s.col]);
    const na=parseFloat(va.replace(/[€%., ]/g,'')),nb=parseFloat(vb.replace(/[€%., ]/g,''));
    return (s.asc?1:-1)*(!isNaN(na)&&!isNaN(nb)?na-nb:va.localeCompare(vb,'it'));});}
  el.innerHTML=`<thead><tr>${headers.map((h,i)=>`<th class="${s.col===i?(s.asc?'sa':'sd'):''}" onclick="sortTbl('${id}',${i})">${h}</th>`).join('')}</tr></thead>`+
    `<tbody>${sr.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>`;
}
function sortTbl(id,col){
  const s=sortState[id]||{col:-1,asc:true};
  sortState[id]={col,asc:s.col===col?!s.asc:true};
  const map={'tbl-cat':'vendite','tbl-drill':'vendite','tbl-cli':'clienti',
    'tbl-sc':'sconti','tbl-over60-cli':'sconti','tbl-marg':'margine',
    'tbl-tr':'trasporti','tbl-tr-reg':'trasporti','tbl-tr-agt':'trasporti','tbl-ord':'ordini','tbl-agt-report':'agenti','tbl-budget':'agenti'};
  if(map[id]==='vendite')renderVendite();
  else if(map[id]==='clienti')filterCliTbl();
  else if(map[id]==='agenti')renderAgenti();
  else if(map[id]==='sconti')renderScontiTbl();
  else if(map[id]==='margine')renderMargine();
  else if(map[id]==='trasporti')renderTrasporti();
  else if(map[id]==='ordini')filterOrdTbl();
}
function sh(s){return String(s).replace(/<[^>]+>/g,'').trim();}

// ═══════════════════════════════════════════════════════
//  UI
// ═══════════════════════════════════════════════════════
function kpi(elId,items){
  document.getElementById(elId).innerHTML=items.map(i=>`
    <div class="kk"><div class="kk-bar ${i.col||'g'}"></div>
      <div class="kl">${i.l}</div>
      <div class="kv ${i.col||'def'}">${i.v}</div>
      <div class="ka">
        ${i.sub?`<span class="ks">${i.sub}</span>`:''}
        ${i.delta!=null?`<span class="dt ${i.delta>=0?'up':'dn'}">${i.delta>=0?'↑':'↓'}${pct(Math.abs(i.delta))}</span>`:''}
      </div></div>`).join('');
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

// ═══════════════════════════════════════════════════════
//  MATH & STRING
// ═══════════════════════════════════════════════════════
function fmt(v){if(v==null||isNaN(v))return'—';return'€'+Number(v).toLocaleString('it',{minimumFractionDigits:0,maximumFractionDigits:0});}
function fmtS(v){if(!v&&v!==0)return'—';const a=Math.abs(v);if(a>=1e6)return'€'+(v/1e6).toFixed(1)+'M';if(a>=1000)return'€'+(v/1000).toFixed(0)+'k';return'€'+Math.round(v);}
function pct(v){if(v==null||isNaN(v))return'—';return(Number(v)*100).toFixed(1)+'%';}
function num(v){return parseFloat(v)||0;}
function str(v){return String(v==null?'':v).trim();}
function sum(arr,fn){return(arr||[]).reduce((a,r)=>a+(parseFloat(fn(r))||0),0);}
function avg(arr,fn){if(!arr||!arr.length)return 0;return sum(arr,fn)/arr.length;}
function groupBy(arr,kFn,vFn){
  const r={};(arr||[]).forEach(x=>{const k=kFn(x);if(!r[k])r[k]=[];r[k].push(x);});
  if(vFn)Object.keys(r).forEach(k=>{r[k]=vFn(r[k]);});return r;
}
function trunc(s,n){return s&&s.length>n?s.slice(0,n-1)+'…':s||'';}
