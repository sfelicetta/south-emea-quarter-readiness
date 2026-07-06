# South EMEA FY27 — Quarter Readiness Dashboard

## File Structure

```
readiness-dashboard/
├── index.html               ← Dashboard frontend (open locally o deploy su Heroku/GCS)
├── apps-script-exporter.js  ← Google Apps Script da incollare nel foglio
└── README.md
```

## Setup in 3 passi

### 1. Google Apps Script (backend dati)

1. Apri il Google Sheet "South EMEA FY27 Quarter Readiness"
2. **Extensions → Apps Script**
3. Incolla il contenuto di `apps-script-exporter.js`
4. Sostituisci `YOUR_GOOGLE_SHEET_ID_HERE` con l'ID del tuo sheet (dalla URL)
5. **Deploy → New deployment → Web App**
   - Execute as: **Me**
   - Who has access: **Anyone within Salesforce** (o Anyone with link)
6. Copia l'URL del deployment

### 2. Collegare la dashboard al live data

Nell'`index.html`, cerca il commento `// Uncomment in production:` e aggiorna:

```js
async function loadData() {
  const res = await fetch('https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec');
  const json = await res.json();
  // Merge json.CQ e json.NQ nel DATA object
  Object.assign(DATA, { CQ: json.CQ, NQ: json.NQ });
  Object.assign(TIMESTAMPS, json.timestamps);
}
```

### 3. Deploy su Heroku (già usato per AI Deck Tool)

```bash
# nella cartella readiness-dashboard/
heroku create south-emea-readiness
git init && git add . && git commit -m "init"
heroku buildpacks:set heroku/static
git push heroku main
```

Oppure semplicemente aprire `index.html` direttamente nel browser per uso locale.

## Tab Google Sheet attesi

| Tab Name | Contenuto |
|----------|-----------|
| `CQ Dealband (Opportunity LVL)` | Dati opp singole CQ |
| `CQ Dealband (Combo LVL)` | Dati combo CQ |
| `NQ Dealband (Opportunity LVL)` | Dati opp singole NQ |
| `NQ Dealband (Combo LVL)` | Dati combo NQ |
| `_meta` | Timestamps refresh: A2=Snowflake, B2=Org62, C2=Finplan, D2=HC |

### Header obbligatori in ogni tab Dealband (row 1)

`Band | AE | Forecast | Pipe | Y/Y | Pipe Coverage | BCO | Commit | Finplan`

## Refresh automatico

La dashboard si aggiorna ogni **5 minuti** in autonomia. Il countdown è visibile in basso a destra. Click su "Last refresh" per forzare un refresh manuale.
