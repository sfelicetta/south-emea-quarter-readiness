/**
 * Google Apps Script — South EMEA FY27 Quarter Readiness — Backend
 *
 * SETUP:
 *   1. Incolla questo file in Extensions → Apps Script del tuo Sheet
 *   2. Sostituisci SHEET_ID con l'ID del tuo Google Sheet
 *   3. Deploy → New deployment → Web App
 *      - Execute as: Me
 *      - Who has access: Anyone (within Salesforce o Anyone with link)
 *   4. Copia l'URL del deployment → incollalo in index.html su APPS_SCRIPT_URL
 *
 * ENDPOINT:
 *   GET  ?action=readiness          → restituisce tutti i dati dei 4 tab Dealband
 *   POST { action:'writeForecast', key:'CQ_OPP_SOUTH', value:'$44.2M' }
 *        → scrive il valore nella cella Forecast del tab _meta
 *
 * TAB ATTESI NEL SHEET:
 *   "CQ Dealband (Opportunity LVL)"
 *   "CQ Dealband (Combo LVL)"
 *   "NQ Dealband (Opportunity LVL)"
 *   "NQ Dealband (Combo LVL)"
 *   "_meta"   → A1:header, A2:ts_snowflake, B2:ts_org62, C2:ts_finplan, D2:ts_hc
 *              → righe 4+ per i Forecast override: A=key, B=value
 *
 * HEADER OBBLIGATORI IN OGNI TAB DEALBAND (riga 1, case-insensitive):
 *   Band | OU | AE | Forecast | Pipe | Y/Y | Pipe Coverage | BCO | Commit | Finplan | Closed QTD | Closed QTD YoY | Deals in Commit | ACV PY | Pipe Growth YoY | Hist Pipe Coverage
 */

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE'; // ← CAMBIA QUESTO

const TABS = {
  CQ_OPP:   'CQ Dealband (Opportunity LVL)',
  CQ_COMBO: 'CQ Dealband (Combo LVL)',
  NQ_OPP:   'NQ Dealband (Opportunity LVL)',
  NQ_COMBO: 'NQ Dealband (Combo LVL)',
  META:     '_meta',
};

const OU_KEYS = ['SOUTH', 'IBERIA', 'ITALY', 'EGM', 'MIDEAST'];
const OU_NAMES = { SOUTH:['emea south','south'], IBERIA:['iberia','es'], ITALY:['italy','it'], EGM:['egm','emerging'], MIDEAST:['middle east','me','mideast'] };

// ── CORS helper ──────────────────────────────────────────────────────────────
function setCors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET handler ──────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'readiness';
    if (action === 'readiness') {
      return setCors(ContentService.createTextOutput(JSON.stringify(buildPayload())));
    }
    return setCors(ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' })));
  } catch(err) {
    return setCors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

// ── POST handler (write Forecast back to Sheet) ───────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'writeForecast') {
      writeForecast(body.key, body.value);
      return setCors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
    }
    return setCors(ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' })));
  } catch(err) {
    return setCors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

// ── Build full payload ────────────────────────────────────────────────────────
function buildPayload() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const meta = readMeta(ss);
  return {
    generated_at: new Date().toISOString(),
    timestamps: meta.timestamps,
    forecast_overrides: meta.forecasts,
    CQ: {
      OPP:   parseTab(ss, TABS.CQ_OPP),
      COMBO: parseTab(ss, TABS.CQ_COMBO),
    },
    NQ: {
      OPP:   parseTab(ss, TABS.NQ_OPP),
      COMBO: parseTab(ss, TABS.NQ_COMBO),
    },
  };
}

// ── Read _meta tab ────────────────────────────────────────────────────────────
function readMeta(ss) {
  const sheet = ss.getSheetByName(TABS.META);
  if (!sheet) return { timestamps: {}, forecasts: {} };
  const all = sheet.getDataRange().getValues();
  const ts = all[1] || [];
  const forecasts = {};
  for (let i = 3; i < all.length; i++) {
    if (all[i][0]) forecasts[String(all[i][0])] = String(all[i][1]);
  }
  return {
    timestamps: {
      snowflake: ts[0] ? String(ts[0]) : null,
      org62:     ts[1] ? String(ts[1]) : null,
      finplan:   ts[2] ? String(ts[2]) : null,
      hc:        ts[3] ? String(ts[3]) : null,
    },
    forecasts,
  };
}

// ── Write forecast override ───────────────────────────────────────────────────
function writeForecast(key, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TABS.META);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // Key non trovata → aggiunge nuova riga
  sheet.appendRow([key, value]);
}

// ── Parse a Dealband tab ──────────────────────────────────────────────────────
function parseTab(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log('Tab not found: ' + tabName); return emptyResult(); }
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return emptyResult();

  const hdr = rows[0].map(h => String(h).trim().toLowerCase());
  const data = rows.slice(1).filter(r => r.some(c => c !== ''));

  // Column index helper
  const ci = (...names) => {
    for (const n of names) {
      const idx = hdr.indexOf(n.toLowerCase());
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const iOU   = ci('ou','operating unit');
  const iAE   = ci('ae','account executive');
  const iBand = ci('band','deal band','dealband');
  const iFcst = ci('forecast','fcst');
  const iPipe = ci('pipe','pipeline');
  const iYoY  = ci('y/y','yoy','y-o-y');
  const iCov  = ci('pipe coverage','coverage','pipe cov');
  const iBCO  = ci('bco');
  const iCmt  = ci('commit','commitment');
  const iFP   = ci('finplan','fin plan');
  const iCQTD = ci('closed qtd','closed qtd fy27');
  const iCYoY = ci('closed qtd yoy','closed yoy');
  const iDeals= ci('deals in commit','# deals','deals in commitment');
  const iACV  = ci('acv py','q2 fy26 acv','acv fy26');
  const iPGrw = ci('pipe growth','pipe growth yoy');
  const iHCov = ci('hist pipe coverage','historical pipe coverage','hist cov');

  // Aggregate per OU + per Band
  const ouMap   = {};
  const bandMap = {};
  const aeRows  = [];

  data.forEach(row => {
    const ou   = mapOU(String(row[iOU] || ''));
    const band = String(row[iBand] || '').trim();
    const fcst = n(row[iFcst]);
    const pipe = n(row[iPipe]);
    const yoy  = row[iYoY] != null ? row[iYoY] : 0;
    const cmt  = iCmt  >= 0 ? n(row[iCmt])  : 0;
    const fp   = iFP   >= 0 ? n(row[iFP])   : 0;
    const cqtd = iCQTD >= 0 ? n(row[iCQTD]) : 0;
    const cyoy = iCYoY >= 0 ? row[iCYoY]    : 0;
    const deals= iDeals>= 0 ? n(row[iDeals]): 0;
    const acvpy= iACV  >= 0 ? n(row[iACV])  : 0;
    const pgrw = iPGrw >= 0 ? row[iPGrw]    : 0;
    const hcov = iHCov >= 0 ? n(row[iHCov]) : 0;
    const cov  = iCov  >= 0 ? n(row[iCov])  : 0;
    const bco  = iBCO  >= 0 ? n(row[iBCO])  : 0;

    // OU aggregation
    if (!ouMap[ou]) ouMap[ou] = { fcst:0, pipe:0, yoy_sum:0, yoy_cnt:0, cmt:0, fp:0, cqtd:0, cyoy_sum:0, cyoy_cnt:0, deals:0, acvpy:0, pgrw_sum:0, pgrw_cnt:0, hcov_sum:0, hcov_cnt:0, cov_sum:0, cov_cnt:0, bco_sum:0, ae_cnt:0, bands:{} };
    const o = ouMap[ou];
    o.fcst += fcst; o.pipe += pipe;
    o.yoy_sum += n(yoy); o.yoy_cnt++;
    o.cmt += cmt; o.fp += fp;
    o.cqtd += cqtd;
    o.cyoy_sum += n(cyoy); o.cyoy_cnt++;
    o.deals += deals;
    o.acvpy += acvpy;
    o.pgrw_sum += n(pgrw); o.pgrw_cnt++;
    o.hcov_sum += hcov; o.hcov_cnt++;
    o.cov_sum  += cov;  o.cov_cnt++;
    o.bco_sum  += bco;  o.ae_cnt++;

    // Band within OU
    if (!o.bands[band]) o.bands[band] = { pipe:0, fcst:0, yoy_sum:0, yoy_cnt:0, cmt:0, fp:0 };
    o.bands[band].pipe     += pipe;
    o.bands[band].fcst     += fcst;
    o.bands[band].yoy_sum  += n(yoy);
    o.bands[band].yoy_cnt  ++;
    o.bands[band].cmt      += cmt;
    o.bands[band].fp       += fp;

    // Global band map
    if (!bandMap[band]) bandMap[band] = { pipe:0, fcst:0, yoy_sum:0, yoy_cnt:0, cmt:0, fp:0 };
    bandMap[band].pipe    += pipe;
    bandMap[band].fcst    += fcst;
    bandMap[band].yoy_sum += n(yoy);
    bandMap[band].yoy_cnt ++;
    bandMap[band].cmt     += cmt;
    bandMap[band].fp      += fp;

    // AE row
    if (iAE >= 0 && row[iAE]) {
      aeRows.push({
        ou,
        name:   String(row[iAE]).trim(),
        fcst:   fmtM(fcst),
        pipe:   fmtM(pipe),
        cov:    cov > 0 ? r1(cov).toFixed(1) + 'x' : '—',
        bco:    bco > 0 ? fmtM(bco) : '—',
        deals:  deals,
        closed: cqtd > 0 ? fmtM(cqtd) : '—',
        yoy:    fmtPct(cyoy),
        dir:    n(cyoy) >= 0 ? 'pos' : 'neg',
      });
    }
  });

  // SOUTH = aggregation of all OUs
  const allOUs = Object.values(ouMap);
  ouMap['SOUTH'] = ouMap['SOUTH'] || aggregateAll(allOUs);

  // Build per-OU structured output
  const result = {};
  OU_KEYS.forEach(ouKey => {
    const o = ouMap[ouKey];
    if (!o) { result[ouKey] = emptyOU(); return; }
    const avgCov  = o.cov_cnt  > 0 ? r1(o.cov_sum  / o.cov_cnt)  : 0;
    const avgHCov = o.hcov_cnt > 0 ? r1(o.hcov_sum / o.hcov_cnt) : 0;
    const avgYoY  = o.yoy_cnt  > 0 ? r1(o.yoy_sum  / o.yoy_cnt)  : 0;
    const avgCYoY = o.cyoy_cnt > 0 ? r1(o.cyoy_sum / o.cyoy_cnt) : 0;
    const avgPGrw = o.pgrw_cnt > 0 ? r1(o.pgrw_sum / o.pgrw_cnt) : 0;
    const bcoAE   = o.ae_cnt   > 0 ? r1(o.bco_sum  / o.ae_cnt)   : 0;

    result[ouKey] = {
      kpis: {
        acv_py:       { val: fmtM(o.acvpy),   lbl: 'Q2 FY26 ACV (PY)' },
        forecast:     { val: fmtM(o.fcst),    lbl: 'Q2 FY27 Forecast', editable: true },
        yoy:          { val: fmtPct(avgYoY),  lbl: 'Y/Y Forecast',     dir: avgYoY  >= 0 ? 'pos' : 'neg' },
        pipeline:     { val: fmtM(o.pipe),    lbl: 'Pipeline Q2 FY27' },
        pipe_growth:  { val: fmtPct(avgPGrw), lbl: 'Pipeline Growth Y/Y', dir: avgPGrw >= 0 ? 'pos' : 'neg' },
        pipe_cov:     { val: avgCov.toFixed(1) + 'x',  lbl: 'Pipe Coverage' },
        hist_pipe_cov:{ val: avgHCov.toFixed(1) + 'x', lbl: 'Hist. Pipe Coverage' },
        deals_cmt:    { val: o.deals,         lbl: '# Deals in Commitment' },
        bco_ae:       { val: fmtM(bcoAE),     lbl: 'BCO per AE' },
        closed_qtd:   { val: fmtM(o.cqtd),   lbl: 'Closed QTD FY27', delta: fmtPct(avgCYoY) + ' Y/Y', dir: avgCYoY >= 0 ? 'pos' : 'neg' },
      },
      bands: bandOrder().map(band => {
        const b = o.bands[band] || { pipe:0, fcst:0, yoy_sum:0, yoy_cnt:1, cmt:0, fp:0 };
        const yoy = b.yoy_cnt > 0 ? b.yoy_sum / b.yoy_cnt : 0;
        return { band, fcst: r1(b.fcst/1e6), pipe: r1(b.pipe/1e6), yoy: fmtPct(yoy), dir: yoy >= 0 && band !== '<$0' ? 'pos' : 'neg' };
      }),
      aes: aeRows.filter(a => ouKey === 'SOUTH' || a.ou === ouKey),
      coverage: bandOrder().map(band => {
        const b = o.bands[band] || { cmt:0, fcst:0, fp:0 };
        return { band, commit: r1(b.cmt/1e6), fcst: r1(b.fcst/1e6), fp: r1(b.fp/1e6) };
      }),
    };
  });

  return result;
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function emptyResult() { return OU_KEYS.reduce((a, k) => { a[k] = emptyOU(); return a; }, {}); }
function emptyOU() { return { kpis: {}, bands: [], aes: [], coverage: [] }; }
function bandOrder() { return ['<$0', '$0–$100K', '$100–$500K', '$500K–$1M', '$1M+']; }

function mapOU(raw) {
  const v = raw.toLowerCase().trim();
  for (const [key, aliases] of Object.entries(OU_NAMES)) {
    if (aliases.some(a => v.includes(a))) return key;
  }
  return 'SOUTH';
}

function aggregateAll(ous) {
  const base = { fcst:0, pipe:0, yoy_sum:0, yoy_cnt:0, cmt:0, fp:0, cqtd:0, cyoy_sum:0, cyoy_cnt:0, deals:0, acvpy:0, pgrw_sum:0, pgrw_cnt:0, hcov_sum:0, hcov_cnt:0, cov_sum:0, cov_cnt:0, bco_sum:0, ae_cnt:0, bands:{} };
  ous.forEach(o => {
    Object.keys(base).forEach(k => { if (typeof base[k] === 'number') base[k] += (o[k] || 0); });
    Object.entries(o.bands || {}).forEach(([b, v]) => {
      if (!base.bands[b]) base.bands[b] = { pipe:0, fcst:0, yoy_sum:0, yoy_cnt:0, cmt:0, fp:0 };
      Object.keys(base.bands[b]).forEach(k => base.bands[b][k] += (v[k] || 0));
    });
  });
  return base;
}

function n(v)      { if (typeof v === 'number') return v; const s = String(v).replace(/[$,%\s]/g,''); return parseFloat(s) || 0; }
function r1(v)     { return Math.round(v * 10) / 10; }
function fmtM(v)   { const a = Math.abs(v/1e6); const s = v < 0 ? '-' : ''; return s + '$' + r1(a) + 'M'; }
function fmtPct(v) {
  const num = n(v);
  if (String(v).includes('%')) return String(v);
  const sign = num >= 0 ? '+' : '';
  return sign + Math.round(num * (Math.abs(num) < 1 ? 100 : 1)) + '%';
}

function testBuild() { Logger.log(JSON.stringify(buildPayload(), null, 2)); }
