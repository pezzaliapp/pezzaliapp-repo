# 📊 Cascos Analytics — PezzaliApp

Dashboard professionale per l'analisi di **vendite**, **ordini** e KPI commerciali per **Cormach Srl** / Cascos.

## 🚀 Funzionalità

| Sezione | Contenuto |
|---|---|
| **Overview** | KPI sintetici, trend annuale fatturato, mix categorie, mensile comparativo |
| **Vendite** | Analisi per categoria, agente, top clienti 2024–2026 |
| **Sconti** | Sconto medio per anno, distribuzione, alert sconto >60% (soglia contrattuale) |
| **Trasporti** | Costo trasporto, incidenza %, picchi anomali, confronto annuale |
| **Ordini** | Portafoglio inevaso, scadenze, clienti, pianificazione consegne |
| **Criticità** | Alert automatici su anomalie sconti, ordini scaduti, concentrazione clienti |

## 📁 File da caricare

- **Vendite** — Export Excel `SPEDIZIONI E RESI CLIENTI` (causale V = vendite, H/X = resi)
- **Ordini** — Export Excel `ORDINI CLIENTI DETTAGLIO`
- **Listino Prezzi** *(opzionale)* — CSV con colonne: `Codice`, `Descrizione`, `PrezzoLordo`, `CostoInstallazione`, `CostoTrasporto`

> ⚠️ **Privacy**: tutti i file vengono elaborati **localmente nel browser**. Nessun dato viene inviato a server esterni.

## 🏷️ Logica Sconti

- **Soglia massima contrattuale rivenditori: 60%**
- Salvo promozioni in corso (`Autopromotec 2025`) o azioni speciali autorizzate
- Il listino CSV viene usato come base lordo per il calcolo degli sconti effettivi
- Alert automatico su ogni transazione con sconto > 60%

## 🌐 Deploy su GitHub Pages

```bash
# 1. Crea il repository su GitHub: pezzaliapp_repo
git init
git add .
git commit -m "Initial release — Cascos Analytics v1"
git remote add origin https://github.com/PezzaliStack/pezzaliapp_repo.git
git push -u origin main

# 2. Abilita GitHub Pages:
#    Settings → Pages → Branch: main / folder: / (root)
```

L'app sarà disponibile su: `https://pezzalistack.github.io/pezzaliapp_repo/`

## 🛠️ Tecnologie

- **Vanilla JS + HTML/CSS** — zero framework, massima velocità
- **Chart.js 4** — grafici professionali
- **SheetJS (xlsx)** — lettura file Excel nel browser
- **PapaParse** — parsing CSV
- **PWA** — installabile su desktop e mobile con service worker

## 📐 Design System PezzaliApp

- Background: `#0d0d0d` / `#141414`
- Accent: `#00b894` (verde Cascos)
- Font: Space Grotesk + JetBrains Mono
- Compatibile con il design system pezzaliapp.com

---

*Sviluppato da [PezzaliApp](https://pezzaliapp.com) — © 2026 Cormach Srl*
