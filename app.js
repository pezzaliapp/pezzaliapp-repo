// ══════════════════════════════════════════════════════════════════
//  Cascos Analytics — PezzaliApp
//  app.js — Engine di analisi + rendering dashboard
// ══════════════════════════════════════════════════════════════════

const SCONTO_MAX = 0.60; // soglia contrattuale rivenditori
const MESI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
const PORTO_MAP = { 1: 'Franco', 2: 'Assegnato', 3: 'Franco+Add.', 6: 'Altro' };

// ── stato globale
let G = { vendite: null, ordini: null, listino: null };
let charts = {};

// ════════════════ FILE UPLOAD ════════════════

['vendite','ordini','listino'].forEach(id => {
  const inp = document.getElementById('file-' + id);
  if (!inp) return;
  inp.addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    document.getElementById('card-' + id).classList.add('loaded');
    document.getElementById(id + '-name').textContent = f.name;
    checkReady();
  });
});

function checkReady() {
  const ok = ['vendite','ordini'].every(id =>
    document.getElementById('card-' + id).classList.contains('loaded')
  );
  document.getElementById('btn-analyze').classList.toggle('active', ok);
}

// drag & drop
document.querySelectorAll('.upload-card').forEach(card => {
  card.addEventListener('dragover', e => { e.preventDefault(); card.classList.add('dragover'); });
  card.addEventListener('dragleave', () => card.classList.remove('dragover'));
  card.addEventListener('drop', e => {
    e.preventDefault(); card.classList.remove('dragover');
    const input = card.querySelector('input[type=file]');
    if (!input || !e.dataTransfer.files[0]) return;
    const dt = new DataTransfer();
    dt.items.add(e.dataTransfer.files[0]);
    input.files = dt.files;
    input.dispatchEvent(new Event('change'));
  });
});

// ════════════════ MAIN ANALYSIS ════════════════

async function runAnalysis() {
  showLoading('Lettura file Excel...');
  await sleep(50);

  try {
    G.vendite = await readXLSX('file-vendite');
    setLoadingMsg('Parsing ordini...');
    G.ordini  = await readXLSX('file-ordini');
    setLoadingMsg('Caricamento listino...');

    const listinoInput = document.getElementById('file-listino');
    if (listinoInput.files[0]) {
      const ext = listinoInput.files[0].name.split('.').pop().toLowerCase();
      if (ext === 'csv') G.listino = await readCSV('file-listino');
      else               G.listino = await readXLSX('file-listino');
    }

    setLoadingMsg('Elaborazione dati...');
    await sleep(80);
    buildDashboard();
  } catch(err) {
    hideLoading();
    alert('Errore durante il caricamento: ' + err.message);
    console.error(err);
  }
}

// ════════════════ READ HELPERS ════════════════

function readXLSX(inputId) {
  return new Promise((res, rej) => {
    const f = document.getElementById(inputId).files[0];
    if (!f) return rej(new Error('File non trovato: ' + inputId));
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        res(rows);
      } catch(err) { rej(err); }
    };
    reader.onerror = () => rej(new Error('Errore lettura file'));
    reader.readAsArrayBuffer(f);
  });
}

function readCSV(inputId) {
  return new Promise((res, rej) => {
    const f = document.getElementById(inputId).files[0];
    Papa.parse(f, {
      header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: r => res(r.data),
      error: e => rej(e)
    });
  });
}

// ════════════════ BUILD DASHBOARD ════════════════

function buildDashboard() {
  setLoadingMsg('Rendering grafici...');

  // ── normalizza colonne vendite
  const vRaw = G.vendite;
  const oRaw = G.ordini;

  // colonne vendite
  const COL_V = detectCols(vRaw[0], {
    anno: ['ANNO SPEDIZIONE','ANNO'],
    data: ['DATA SPEDIZIONE','DATA'],
    importo: ['IMPORTO CONSEGNATO','IMPORTO'],
    pz: ['PZ NETTO VENDITA','PZ NETTO'],
    qty: ['QTA CONSEGNATA','QTA'],
    trasporto: ['SPESE DI TRASPORTO','SPESE TRASPORTO','TRASPORTO'],
    causale: ['CAUSALE MAGAZZINO','CAUSALE'],
    cat: ['DESCRIZIONE ELEMENTO.5','CATEGORIA','CAT'],
    agente: ['DESCRIZIONE ELEMENTO.2','AGENTE'],
    cliente: ['RAGIONE SOCIALE 1','RAGIONE SOCIALE','CLIENTE'],
    articolo: ['ARTICOLO'],
    porto: ['PORTO']
  });

  // colonne ordini
  const COL_O = detectCols(oRaw[0], {
    anno: ['ANNO'],
    data: ['DATA CREAZIONE','DATA'],
    cliente: ['CLIENTE.1','CLIENTE','RAGIONE SOCIALE'],
    articolo: ['ARTICOLO'],
    desc: ['DESCRIZIONE'],
    qtyInevasa: ['QTA INEVASA','QTA_INEVASA'],
    importoInevaso: ['IMPORTO INEVASO','IMPORTO_INEVASO'],
    imponibile: ['IMPONIBILE ORD. VAL','IMPONIBILE ORD','IMPONIBILE'],
    porto: ['PORTO'],
    trasporto: ['SPESE DI TRASPORTO','SPESE_TRASPORTO'],
    dataConsegna: ['DATA CONSEGNA','CONSEGNA'],
    cat3: ['CLASSE 3 ARTICOLO','CLASSE 3'],
    cat1: ['CLASSE 1 ARTICOLO','CLASSE 1']
  });

  // ── filtra solo vendite (causale V)
  const VEND = vRaw.filter(r => String(r[COL_V.causale]||'').toUpperCase().startsWith('V'));
  const RESI = vRaw.filter(r => ['H1','X2','X4','X5'].includes(String(r[COL_V.causale]||'').toUpperCase()));

  // ── normalizza listino
  const listinoMap = {};
  if (G.listino) {
    G.listino.forEach(r => {
      const cod = String(r.Codice||r.CODICE||r.codice||'').replace(/^0+/,'');
      const pl  = parseFloat(r.PrezzoLordo||r.PREZZO_LORDO||r.prezzo_lordo||0);
      if (cod && pl > 0) listinoMap[cod] = pl;
    });
  }

  // ── attach listino prices & sconti
  VEND.forEach(r => {
    const cod = String(r[COL_V.articolo]||'').replace(/^0+/,'');
    r.__lordo = listinoMap[cod] || null;
    r.__pz = parseFloat(r[COL_V.pz]) || 0;
    r.__qty = parseFloat(r[COL_V.qty]) || 0;
    r.__importo = parseFloat(r[COL_V.importo]) || 0;
    r.__trasporto = parseFloat(r[COL_V.trasporto]) || 0;
    r.__anno = parseInt(r[COL_V.anno]) || 0;
    r.__cat = String(r[COL_V.cat]||'').trim();
    r.__agente = String(r[COL_V.agente]||'').trim();
    r.__cliente = String(r[COL_V.cliente]||'').trim();
    r.__porto = r[COL_V.porto];
    if (r.__lordo && r.__lordo > 0 && r.__pz > 0) {
      r.__sconto = Math.max(0, Math.min(1, 1 - r.__pz / r.__lordo));
    } else {
      r.__sconto = null;
    }
    // data
    try {
      const d = r[COL_V.data];
      r.__date = d instanceof Date ? d : new Date(d);
    } catch(e) { r.__date = null; }
    r.__mese = r.__date instanceof Date && !isNaN(r.__date) ? r.__date.getMonth() : -1;
  });

  // ── anni presenti
  const anni = [...new Set(VEND.map(r => r.__anno))].filter(a => a > 2000).sort();
  const annoMax = Math.max(...anni);
  const anniStorici = anni.filter(a => a < annoMax); // escludiamo anno corrente parziale se < 12 mesi

  // ── periodo
  document.getElementById('dash-period').textContent =
    `Periodo: ${anni[0]} – ${annoMax} | ${VEND.length.toLocaleString('it')} righe vendita | ${oRaw.length} ordini inevasi`;

  // ════ KPI DATI ════

  // Fatturato per anno
  const fattAnnuo = groupBy(VEND, r => r.__anno, rows => sum(rows, r => r.__importo));

  // Trasporto per anno
  const traspAnnuo = groupBy(VEND, r => r.__anno, rows => sum(rows, r => r.__trasporto));

  // Sconto medio per anno
  const scontoAnnuo = groupBy(
    VEND.filter(r => r.__sconto !== null), r => r.__anno,
    rows => rows.length > 0 ? avg(rows, r => r.__sconto) : null
  );

  // Categorie
  const fattCat = groupBy(VEND.filter(r => r.__cat), r => r.__cat,
    rows => ({ fatturato: sum(rows, r => r.__importo), pezzi: sum(rows, r => r.__qty) })
  );

  // Agenti 2024+
  const fattAgente = groupBy(VEND.filter(r => r.__anno >= 2024 && r.__agente),
    r => r.__agente, rows => sum(rows, r => r.__importo)
  );

  // Clienti 2024+
  const fattClienti = groupBy(VEND.filter(r => r.__anno >= 2024 && r.__cliente),
    r => r.__cliente, rows => sum(rows, r => r.__importo)
  );

  // Mensile 2024 e 2025/2026
  const mensile = {};
  [2024, 2025, annoMax].forEach(a => {
    mensile[a] = Array(12).fill(0);
    VEND.filter(r => r.__anno === a && r.__mese >= 0).forEach(r => {
      mensile[a][r.__mese] += r.__importo;
    });
  });

  // Porto distribution
  const portoDist = groupBy(VEND, r => r.__porto, rows => rows.length);

  // Over 60%
  const over60 = VEND.filter(r => r.__sconto !== null && r.__sconto > SCONTO_MAX)
    .sort((a,b) => b.__sconto - a.__sconto);

  // Ordini
  oRaw.forEach(r => {
    r.__clienteO  = String(r[COL_O.cliente]||'').trim();
    r.__importoI  = parseFloat(r[COL_O.importoInevaso]) || 0;
    r.__trasportoO= parseFloat(r[COL_O.trasporto]) || 0;
    r.__desc      = String(r[COL_O.desc]||'').trim();
    r.__qty       = parseFloat(r[COL_O.qtyInevasa]) || 0;
    r.__pz        = parseFloat(r[COL_O.imponibile]) || 0;
    try {
      const d = r[COL_O.dataConsegna];
      r.__consegna = d instanceof Date ? d : new Date(d);
    } catch(e) { r.__consegna = null; }
  });

  const ordClienti = groupBy(oRaw, r => r.__clienteO,
    rows => ({ importo: sum(rows, r => r.__importoI), n: rows.length })
  );
  const ordByDate = groupBy(oRaw.filter(r => r.__consegna && !isNaN(r.__consegna)),
    r => r.__consegna.toISOString().split('T')[0],
    rows => sum(rows, r => r.__importoI)
  );

  const totOrdini = sum(oRaw, r => r.__importoI);
  const totTrasportoOrdini = sum(oRaw, r => r.__trasportoO);

  // ════════════════ RENDER TABS ════════════════

  // ── OVERVIEW KPIs
  const fatt2025 = fattAnnuo[2025] || 0;
  const fatt2024 = fattAnnuo[2024] || 0;
  const fatt2026 = fattAnnuo[annoMax] || 0;
  const deltaPct = fatt2024 > 0 ? (fatt2025 - fatt2024) / fatt2024 : 0;
  const scontoMedio = Object.values(scontoAnnuo).filter(v => v !== null);
  const scontoGlobale = scontoMedio.length > 0 ? avg(VEND.filter(r => r.__sconto !== null), r => r.__sconto) : null;
  const trasportoTot = sum(VEND, r => r.__trasporto);
  const fattTot = sum(VEND, r => r.__importo);

  renderKPIs('kpi-overview', [
    { label:'Fatturato Totale', value: fmt(fattTot), sub: `${anni[0]}–${annoMax}`, cls:'green' },
    { label:'Fatturato 2025', value: fmt(fatt2025), delta: deltaPct, sub: 'vs 2024', cls:'green' },
    { label:'Fatturato 2026 YTD', value: fmt(fatt2026), sub: 'Anno in corso', cls:'blue' },
    { label:'Ordini Inevasi', value: fmt(totOrdini), sub: `${oRaw.length} righe`, cls:'yellow' },
    { label:'Incidenza Trasporti', value: pct(trasportoTot/fattTot), sub: `€${fmt(trasportoTot)} tot`, cls:'purple' },
    { label:'Sconto Medio Globale', value: scontoGlobale !== null ? pct(scontoGlobale) : 'N/D', sub: 'Su articoli listino', cls: scontoGlobale > SCONTO_MAX ? 'red':'green' },
  ]);

  // ── Chart: Annual
  renderLineBar('chart-annual', anni, anni.map(a => fattAnnuo[a]||0), anni.map(a => traspAnnuo[a]||0));

  // ── Chart: Cat Pie
  const catSorted = Object.entries(fattCat).sort((a,b) => b[1].fatturato - a[1].fatturato).slice(0,8);
  renderPie('chart-cat-pie', catSorted.map(([k]) => k.split(' - ')[0]), catSorted.map(([,v]) => v.fatturato));

  // ── Chart: Monthly comparison
  const annoComp1 = annoMax === 2026 ? 2025 : Math.max(...anni.filter(a=>a<annoMax));
  const annoComp2 = annoMax === 2026 ? 2026 : annoMax;
  renderMonthly('chart-monthly', annoComp1, mensile[annoComp1]||Array(12).fill(0), annoComp2, mensile[annoComp2]||Array(12).fill(0));

  // ── Chart: Top clienti
  const topCli = Object.entries(fattClienti).sort((a,b)=>b[1]-a[1]).slice(0,10);
  renderHBar('chart-top-cli', topCli.map(([k])=>k.length>22?k.slice(0,20)+'…':k), topCli.map(([,v])=>v), '#00b894');

  // ── VENDITE TAB
  renderKPIs('kpi-vendite', [
    { label:'Righe Vendita', value: VEND.length.toLocaleString('it'), sub:'Causale V', cls:'green' },
    { label:'Ticket Medio', value: fmt(VEND.length > 0 ? fattTot/VEND.length : 0), sub:'Per riga spedizione', cls:'blue' },
    { label:'Prezzo Netto Medio', value: fmt(avg(VEND, r=>r.__pz)), sub:'Per unità', cls:'green' },
    { label:'Resi Totali', value: RESI.length.toLocaleString('it'), sub:'Causale H/X', cls:'red' },
    { label:'Categorie Attive', value: Object.keys(fattCat).length, sub:'Famiglie prodotto', cls:'purple' },
    { label:'Clienti Attivi 24-26', value: Object.keys(fattClienti).length, sub:'Con almeno 1 vendita', cls:'blue' },
  ]);

  // Cat bar
  const catAll = Object.entries(fattCat).sort((a,b)=>b[1].fatturato-a[1].fatturato).slice(0,12);
  renderVBar('chart-cat-bar', catAll.map(([k])=>k.split(' - ')[0]), catAll.map(([,v])=>v.fatturato), '#74b9ff');

  // Agent bar
  const agSorted = Object.entries(fattAgente).sort((a,b)=>b[1]-a[1]).slice(0,8);
  renderHBar('chart-agent', agSorted.map(([k])=>k), agSorted.map(([,v])=>v), '#a29bfe');

  // Table clienti
  const topCli15 = Object.entries(fattClienti).sort((a,b)=>b[1]-a[1]).slice(0,15);
  renderTable('table-clienti',
    ['Cliente','Fatturato 2024–2026','%'],
    topCli15.map(([k,v]) => [
      k,
      `<span class="td-mono">${fmt(v)}</span>`,
      `<span class="badge badge-green">${pct(v/fattTot)}</span>`
    ])
  );

  // ── SCONTI TAB
  const scontiRows = VEND.filter(r => r.__sconto !== null);
  const scontoGlobal2 = scontiRows.length > 0 ? avg(scontiRows, r => r.__sconto) : 0;
  const over60Count = over60.length;
  const over60Val = sum(over60, r => r.__importo);

  renderKPIs('kpi-sconti', [
    { label:'Sconto Medio', value: pct(scontoGlobal2), sub:`Su ${scontiRows.length} righe listino`, cls: scontoGlobal2 > SCONTO_MAX ? 'red':'green' },
    { label:'Righe >60% Sconto', value: over60Count.toLocaleString('it'), sub: `€${fmt(over60Val)} di valore`, cls:'red' },
    { label:'Soglia Max Contrattuale', value: pct(SCONTO_MAX), sub:'Salvo promozioni/azioni speciali', cls:'yellow' },
    { label:'Righe con Listino', value: scontiRows.length.toLocaleString('it'), sub:`${pct(scontiRows.length/VEND.length)} del totale`, cls:'blue' },
  ]);

  // Disc meter by year
  const discMeterEl = document.getElementById('disc-meter');
  discMeterEl.innerHTML = '';
  [...anni].reverse().slice(0,6).reverse().forEach(a => {
    const sc = scontoAnnuo[a];
    if (sc === null) return;
    const isOver = sc > SCONTO_MAX;
    const col = isOver ? '#e17055' : '#00b894';
    discMeterEl.innerHTML += `
      <div class="disc-card">
        <div class="disc-name">${a}</div>
        <div class="disc-value" style="color:${col}">${pct(sc)}</div>
        <div class="disc-gauge">
          <div class="disc-fill" style="width:${Math.min(100,sc*100)}%;background:${col}"></div>
        </div>
      </div>`;
  });

  // Chart: disc annual
  renderDiscAnnual('chart-disc-annual', anni, anni.map(a => (scontoAnnuo[a]||0)*100));

  // Chart: disc distribution
  const buckets = [0,10,20,30,40,50,60,70,80,90,100];
  const discDist = Array(10).fill(0);
  scontiRows.forEach(r => {
    const b = Math.min(9, Math.floor(r.__sconto*100/10));
    discDist[b]++;
  });
  renderVBar('chart-disc-dist',
    buckets.slice(0,10).map((b,i) => `${b}–${buckets[i+1]}%`),
    discDist,
    null,
    discDist.map((_, i) => i >= 6 ? '#e17055' : '#00b894')
  );

  // Table over60
  renderTable('table-over60',
    ['Anno','Cliente','Prodotto','PZ Netto','Lordo','Sconto'],
    over60.slice(0,50).map(r => [
      r.__anno,
      (r.__cliente||'').slice(0,30),
      (r[detectColName(r, ['DESCRIZIONE'])]||'').slice(0,30),
      `<span class="td-mono">${fmt(r.__pz)}</span>`,
      `<span class="td-mono">${fmt(r.__lordo)}</span>`,
      `<span class="badge badge-red">${pct(r.__sconto)}</span>`
    ])
  );

  // ── TRASPORTI TAB
  const trasportoGlobal = sum(VEND, r => r.__trasporto);
  const incidenzaMedia = fattTot > 0 ? trasportoGlobal/fattTot : 0;
  const anniTr = anni.filter(a => fattAnnuo[a] > 0);
  const incidenze = anniTr.map(a => {
    const f = fattAnnuo[a]||1; const t = traspAnnuo[a]||0;
    return { anno: a, fatt: f, trasp: t, inc: t/f };
  });
  const maxInc = Math.max(...incidenze.map(i=>i.inc));

  renderKPIs('kpi-trasporti', [
    { label:'Spese Trasporto Totali', value: fmt(trasportoGlobal), sub:`${anni[0]}–${annoMax}`, cls:'purple' },
    { label:'Incidenza Media', value: pct(incidenzaMedia), sub:'Su fatturato totale', cls: incidenzaMedia > 0.05 ? 'red':'green' },
    { label:'Anno Picco', value: (incidenze.find(i=>i.inc===maxInc)||{}).anno||'—', sub: `${pct(maxInc)} incidenza`, cls:'yellow' },
    { label:'Spese Trasporti 2025', value: fmt(traspAnnuo[2025]||0), sub: pct((traspAnnuo[2025]||0)/(fattAnnuo[2025]||1)), cls:'blue' },
    { label:'Trasp. Ordini Inevasi', value: fmt(totTrasportoOrdini), sub:`${oRaw.length} ordini`, cls:'purple' },
  ]);

  renderTransportChart('chart-transport', anniTr,
    anniTr.map(a=>traspAnnuo[a]||0),
    anniTr.map(a=>((traspAnnuo[a]||0)/(fattAnnuo[a]||1))*100)
  );

  const portoCounts = {};
  VEND.forEach(r => {
    const p = PORTO_MAP[r.__porto] || `Porto ${r.__porto}`;
    portoCounts[p] = (portoCounts[p]||0) + 1;
  });
  renderPie('chart-porto', Object.keys(portoCounts), Object.values(portoCounts));

  renderTable('table-transport',
    ['Anno','Fatturato','Spese Trasporto','Incidenza %','Δ Incidenza'],
    incidenze.map((item, i) => {
      const delta = i > 0 ? item.inc - incidenze[i-1].inc : 0;
      const cls = item.inc > 0.05 ? 'badge-red' : 'badge-green';
      return [
        item.anno,
        `<span class="td-mono">${fmt(item.fatt)}</span>`,
        `<span class="td-mono">${fmt(item.trasp)}</span>`,
        `<span class="badge ${cls}">${pct(item.inc)}</span>`,
        i > 0 ? `<span class="${delta>0?'badge badge-red':'badge badge-green'}">${delta>0?'+':''}${pct(delta)}</span>` : '—'
      ];
    })
  );

  // ── ORDINI TAB
  renderKPIs('kpi-ordini', [
    { label:'Valore Inevaso Totale', value: fmt(totOrdini), sub:`${oRaw.length} righe ordine`, cls:'yellow' },
    { label:'Clienti in Portafoglio', value: Object.keys(ordClienti).length, sub:'Con ordini aperti', cls:'blue' },
    { label:'Ticket Medio Ordine', value: oRaw.length > 0 ? fmt(totOrdini/oRaw.length) : '—', sub:'Per riga ordine', cls:'green' },
    { label:'Spese Trasporto Ordini', value: fmt(totTrasportoOrdini), sub:'Già pianificate', cls:'purple' },
  ]);

  const ordCliTop = Object.entries(ordClienti).sort((a,b)=>b[1].importo-a[1].importo).slice(0,12);
  renderHBar('chart-ordini-cli',
    ordCliTop.map(([k])=>k.length>22?k.slice(0,20)+'…':k),
    ordCliTop.map(([,v])=>v.importo), '#fdcb6e'
  );

  const dateOrd = Object.entries(ordByDate).sort(([a],[b])=>a.localeCompare(b)).slice(0,20);
  renderVBar('chart-ordini-date',
    dateOrd.map(([k])=>k.slice(5)),
    dateOrd.map(([,v])=>v), '#74b9ff'
  );

  renderTable('table-ordini',
    ['Cliente','Prodotto','Q.tà','Importo Inevaso','Data Consegna'],
    [...oRaw].sort((a,b)=>b.__importoI-a.__importoI).slice(0,60).map(r => {
      const today = new Date();
      const late = r.__consegna && r.__consegna < today;
      return [
        (r.__clienteO||'').slice(0,28),
        (r.__desc||'').slice(0,30),
        r.__qty,
        `<span class="td-mono">${fmt(r.__importoI)}</span>`,
        r.__consegna && !isNaN(r.__consegna)
          ? `<span class="badge ${late?'badge-red':'badge-green'}">${r.__consegna.toLocaleDateString('it')}</span>`
          : '—'
      ];
    })
  );

  // ── CRITICITÀ TAB
  buildCriticita({ VEND, over60, scontoGlobal2, incidenzaMedia, incidenze, oRaw, totOrdini, fattAnnuo, anni, annoMax, traspAnnuo });

  // ── SHOW
  setLoadingMsg('Quasi pronto...');
  setTimeout(() => {
    hideLoading();
    document.getElementById('upload-screen').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';
    document.getElementById('status-badge').textContent = 'ATTIVO';
  }, 400);
}

// ════════════════ CRITICITÀ ════════════════

function buildCriticita({ VEND, over60, scontoGlobal2, incidenzaMedia, incidenze, oRaw, totOrdini, fattAnnuo, anni, annoMax, traspAnnuo }) {
  const alerts = [];

  // 1. Sconti oltre soglia
  if (over60.length > 0) {
    const pctOver = over60.length / VEND.filter(r=>r.__sconto!==null).length;
    alerts.push({
      type: pctOver > 0.1 ? 'danger' : 'warn',
      icon: '🏷️',
      title: `Sconti oltre soglia 60%: ${over60.length} righe (${pct(pctOver)})`,
      msg: `Valore totale transazioni: €${fmt(sum(over60, r=>r.__importo))}. Verificare se trattasi di promozioni autorizzate o azioni speciali. I casi più critici riguardano accessori e componenti.`
    });
  }

  // 2. Sconto medio globale
  if (scontoGlobal2 > 0) {
    const diff = scontoGlobal2 - SCONTO_MAX;
    alerts.push({
      type: diff > 0 ? 'warn' : 'ok',
      icon: diff > 0 ? '⚠️' : '✅',
      title: `Sconto medio globale: ${pct(scontoGlobal2)} (max contrattuale: ${pct(SCONTO_MAX)})`,
      msg: diff > 0
        ? `Il livello medio supera la soglia contrattuale di ${pct(Math.abs(diff))}. Consigliabile revisione policy sconti con la rete vendita.`
        : `Il livello medio è entro i limiti contrattuali (margine residuo: ${pct(Math.abs(diff))}).`
    });
  }

  // 3. Incidenza trasporti 2022 picco
  const picco = incidenze.reduce((m,i)=>i.inc>m.inc?i:m, incidenze[0]||{inc:0});
  if (picco && picco.inc > 0.07) {
    alerts.push({
      type: 'warn',
      icon: '🚚',
      title: `Picco incidenza trasporti ${picco.anno}: ${pct(picco.inc)}`,
      msg: `Nel ${picco.anno} le spese di trasporto hanno raggiunto €${fmt(traspAnnuo[picco.anno]||0)}, il ${pct(picco.inc)} del fatturato. Analizzare i fattori (aumento carburante, volumi extra, nuove zone) per prevenire future anomalie.`
    });
  }

  // 4. Trend fatturato 2025 vs 2024
  const f25 = fattAnnuo[2025]||0; const f24 = fattAnnuo[2024]||0;
  if (f24 > 0 && f25 > 0) {
    const delta = (f25-f24)/f24;
    alerts.push({
      type: delta < -0.05 ? 'warn' : delta > 0.05 ? 'ok' : 'info',
      icon: delta > 0 ? '📈' : '📉',
      title: `Trend 2025 vs 2024: ${delta > 0 ? '+' : ''}${pct(delta)}`,
      msg: `Fatturato 2024: €${fmt(f24)} → 2025: €${fmt(f25)}. ${delta < 0 ? 'Attenzione alla contrazione del fatturato: analizzare categorie e clienti in calo.' : 'Crescita positiva confermata. Consolidare i segmenti in crescita.'}`
    });
  }

  // 5. Ordini scaduti
  const today = new Date();
  const scaduti = oRaw.filter(r => r.__consegna && !isNaN(r.__consegna) && r.__consegna < today);
  if (scaduti.length > 0) {
    alerts.push({
      type: 'danger',
      icon: '⏰',
      title: `Ordini con data consegna scaduta: ${scaduti.length} righe`,
      msg: `Valore inevaso scaduto: €${fmt(sum(scaduti, r=>r.__importoI))}. Clienti coinvolti: ${[...new Set(scaduti.map(r=>r.__clienteO))].slice(0,5).join(', ')}. Aggiornare le date di consegna o contattare i clienti.`
    });
  }

  // 6. Ordini in scadenza 30gg
  const in30 = oRaw.filter(r => {
    if (!r.__consegna || isNaN(r.__consegna)) return false;
    const diff = (r.__consegna - today) / 86400000;
    return diff >= 0 && diff <= 30;
  });
  if (in30.length > 0) {
    alerts.push({
      type: 'warn',
      icon: '📅',
      title: `${in30.length} righe ordine in scadenza entro 30 giorni`,
      msg: `Valore: €${fmt(sum(in30, r=>r.__importoI))}. Priorità logistica da pianificare.`
    });
  }

  // 7. Concentrazione clienti
  const cli2425 = Object.entries(groupBy(VEND.filter(r=>r.__anno>=2024), r=>r.__cliente, rows=>sum(rows,r=>r.__importo)));
  const fatt2425 = sum(VEND.filter(r=>r.__anno>=2024), r=>r.__importo);
  const top3 = cli2425.sort((a,b)=>b[1]-a[1]).slice(0,3);
  const top3pct = fatt2425 > 0 ? sum(top3, ([,v])=>v) / fatt2425 : 0;
  if (top3pct > 0.4) {
    alerts.push({
      type: 'warn',
      icon: '🏢',
      title: `Alta concentrazione clienti: top 3 = ${pct(top3pct)} del fatturato 2024-2026`,
      msg: `${top3.map(([k,v])=>`${k.slice(0,25)} (${pct(v/fatt2425)})`).join(', ')}. Rischio di dipendenza da pochi clienti. Diversificare il portafoglio.`
    });
  }

  // 8. Resi
  const RESI_count = VEND.filter(r=>['H1','X2','X4','X5'].includes(String(r[detectColNameRaw(r,'CAUSALE MAGAZZINO','CAUSALE')]||'').toUpperCase())).length;
  const resiPct = VEND.length > 0 ? RESI_count / (VEND.length + RESI_count) : 0;
  if (resiPct > 0.02) {
    alerts.push({
      type: 'info',
      icon: '↩️',
      title: `Resi: ${RESI_count} righe (${pct(resiPct)} sul totale movimenti)`,
      msg: `Monitorare le motivazioni dei resi per categoria. Un tasso sopra il 5% può indicare problemi qualitativi o logistici.`
    });
  }

  // render
  const el = document.getElementById('alert-list');
  if (alerts.length === 0) {
    el.innerHTML = `<div class="alert-item ok"><div class="alert-icon">✅</div><div class="alert-text"><strong>Nessuna criticità rilevata</strong><p>Tutti gli indicatori sono nei range normali.</p></div></div>`;
  } else {
    el.innerHTML = alerts.map(a => `
      <div class="alert-item ${a.type}">
        <div class="alert-icon">${a.icon}</div>
        <div class="alert-text">
          <strong>${a.title}</strong>
          <p>${a.msg}</p>
        </div>
      </div>`).join('');
  }
}

// ════════════════ CHART BUILDERS ════════════════

const PALETTE = ['#00b894','#74b9ff','#fdcb6e','#a29bfe','#e17055','#55efc4','#fd79a8','#6c5ce7','#00cec9'];

function renderLineBar(id, labels, barData, lineData) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    data: {
      labels,
      datasets: [
        {
          type: 'bar', label: 'Fatturato (€)',
          data: barData, backgroundColor: 'rgba(0,184,148,0.7)',
          borderRadius: 4, yAxisID: 'y'
        },
        {
          type: 'line', label: 'Trasporto (€)',
          data: lineData, borderColor: '#e17055', backgroundColor: 'rgba(225,112,85,0.1)',
          tension: 0.3, pointRadius: 4, fill: true, yAxisID: 'y2'
        }
      ]
    },
    options: chartOptions({ y: { title: 'Fatturato', position: 'left' }, y2: { title: 'Trasporto', position: 'right', grid: { drawOnChartArea: false } } })
  });
}

function renderMonthly(id, a1, d1, a2, d2) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    data: {
      labels: MESI,
      datasets: [
        { type: 'bar', label: `${a1}`, data: d1, backgroundColor: 'rgba(116,185,255,0.6)', borderRadius: 3 },
        { type: 'line', label: `${a2}`, data: d2, borderColor: '#00b894', tension: 0.3, pointRadius: 4, fill: false }
      ]
    },
    options: chartOptions()
  });
}

function renderPie(id, labels, data) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data, backgroundColor: PALETTE, borderWidth: 0, hoverOffset: 6 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: { color: '#888', font: { size: 10, family: 'Space Grotesk' }, boxWidth: 10, padding: 8 }
        },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: €${fmt(ctx.raw)}` } }
      }
    }
  });
}

function renderHBar(id, labels, data, color) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ data, backgroundColor: color || '#00b894', borderRadius: 3 }]
    },
    options: {
      ...chartOptions(),
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` €${fmt(ctx.raw)}` } }
      }
    }
  });
}

function renderVBar(id, labels, data, color, colors) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ data, backgroundColor: colors || color || '#74b9ff', borderRadius: 3 }]
    },
    options: {
      ...chartOptions(),
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` €${fmt(ctx.raw)}` } }
      }
    }
  });
}

function renderDiscAnnual(id, labels, data) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Sconto medio %', data, borderColor: '#a29bfe', backgroundColor: 'rgba(162,155,254,0.1)', tension: 0.3, fill: true, pointRadius: 5 },
        { label: 'Soglia 60%', data: labels.map(()=>60), borderColor: '#e17055', borderDash: [6,4], pointRadius: 0, borderWidth: 1.5 }
      ]
    },
    options: {
      ...chartOptions(),
      plugins: {
        legend: { display: true, labels: { color: '#888', font: { size: 11 } } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.raw.toFixed(1)}%` } }
      }
    }
  });
}

function renderTransportChart(id, labels, trData, incData) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    data: {
      labels,
      datasets: [
        { type: 'bar', label: 'Trasporto (€)', data: trData, backgroundColor: 'rgba(162,155,254,0.7)', borderRadius: 4, yAxisID: 'y' },
        { type: 'line', label: 'Incidenza %', data: incData, borderColor: '#fdcb6e', tension: 0.3, pointRadius: 4, fill: false, yAxisID: 'y2' }
      ]
    },
    options: chartOptions({ y: { title: '€', position: 'left' }, y2: { title: '%', position: 'right', grid: { drawOnChartArea: false } } })
  });
}

function chartOptions(extraAxes = {}) {
  const scales = {
    x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#555', font: { size: 10, family: 'Space Grotesk' } } },
    y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#555', font: { size: 10, family: 'Space Grotesk' }, callback: v => fmt(v) } }
  };
  Object.entries(extraAxes).forEach(([k, v]) => {
    scales[k] = {
      position: v.position, grid: v.grid || { color: 'rgba(255,255,255,0.05)' },
      ticks: { color: '#555', font: { size: 10 } }
    };
  });
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false }, tooltip: { backgroundColor: '#1a1a1a', borderColor: '#2a2a2a', borderWidth: 1, titleColor: '#e8e8e8', bodyColor: '#888' } },
    scales
  };
}

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

// ════════════════ UI HELPERS ════════════════

function renderKPIs(elId, items) {
  document.getElementById(elId).innerHTML = items.map(i => `
    <div class="kpi-card ${i.cls}">
      <div class="kpi-label">${i.label}</div>
      <div class="kpi-value">${i.value}</div>
      ${i.sub ? `<div class="kpi-sub">${i.sub}</div>` : ''}
      ${i.delta !== undefined ? `<div class="kpi-delta ${i.delta >= 0 ? 'up':'down'}">${i.delta >= 0 ? '↑':'↓'} ${pct(Math.abs(i.delta))}</div>` : ''}
    </div>`).join('');
}

function renderTable(elId, headers, rows) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = `
    <thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
    <tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>`;
}

function showTab(name) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  event.currentTarget.classList.add('active');
}

function showLoading(msg) {
  document.getElementById('loading-screen').style.display = 'flex';
  document.getElementById('loading-msg').textContent = msg;
}
function setLoadingMsg(msg) {
  document.getElementById('loading-msg').textContent = msg;
}
function hideLoading() {
  document.getElementById('loading-screen').style.display = 'none';
}
function resetApp() {
  location.reload();
}

// ════════════════ MATH HELPERS ════════════════

function fmt(v) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return '€' + Number(v).toLocaleString('it', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
function pct(v) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return (Number(v)*100).toFixed(1) + '%';
}
function sum(arr, fn) {
  return arr.reduce((a, r) => a + (parseFloat(fn(r)) || 0), 0);
}
function avg(arr, fn) {
  if (!arr.length) return 0;
  return sum(arr, fn) / arr.length;
}
function groupBy(arr, keyFn, valFn) {
  const res = {};
  arr.forEach(r => {
    const k = keyFn(r);
    if (!res[k]) res[k] = [];
    res[k].push(r);
  });
  if (valFn) Object.keys(res).forEach(k => { res[k] = valFn(res[k]); });
  return res;
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ════════════════ COLUMN DETECTION ════════════════

function detectCols(firstRow, map) {
  const keys = firstRow ? Object.keys(firstRow) : [];
  const result = {};
  Object.entries(map).forEach(([alias, candidates]) => {
    result[alias] = candidates.find(c => keys.includes(c)) || candidates[0];
  });
  return result;
}
function detectColName(row, candidates) {
  return candidates.find(c => c in row) || candidates[0];
}
function detectColNameRaw(row, ...candidates) {
  return candidates.find(c => c in row) || candidates[0];
}
