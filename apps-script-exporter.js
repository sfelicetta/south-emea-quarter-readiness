/**
 * Google Apps Script — South EMEA FY27 Quarter Readiness — Backend
 *
 * STRUTTURA REALE DEL TAB (0-indexed):
 *   col 0:  sempre vuota
 *   col 1:  Deal Band  (header "Deal Band", poi "<$0", "<$100K", "$100K - $500K", ...)
 *   col 2:  Q2 FY26 ACV          es. "$33.3"
 *   col 3:  Q2 FY27 Forecast     es. "$19.0"  (SOLO nella riga Total)
 *   col 4:  Y/Y forecast          es. "12%", "-100%", "-"
 *   col 5:  Q2 FY27 Pipe         es. "$24.7"
 *   col 6:  Pipe Growth Y/Y       es. "-14%"
 *   col 7:  Pipe Coverage (Hist)  es. "-1.3x  (1.9x)" o "2x  (1.7x)"
 *   col 8:  # Deals in Commit     es. "#Q2 Open Opps: 930"
 *   col 9:  BCO per AE            es. "BCO per AE $57K, -18% Y/Y"
 *   col 10: FY27 Closed QTD       es. "$19.6"
 *   col 11: Y/Y closed QTD        es. "6%"
 *   col 15: OU name               es. "South", "Iberia", "Italy", "EGM", "Middle East"
 */

const SHEET_ID = '1DdcjTIC-w26qBgYYdIMHbLQIOthfxnHAMv7AC27QLpc';

const TABS = {
  CQ_OPP:   'CQ Dealband (Opportunity LVL)',
  CQ_COMBO: 'CQ Dealband (Combo LVL)',
  NQ_OPP:   'NQ Dealband (Opportunity LVL)',
  NQ_COMBO: 'NQ Dealband (Combo LVL)',
  META:     '_meta',
};

const OU_KEYS    = ['SOUTH', 'IBERIA', 'ITALY', 'EGM', 'MIDEAST'];
const BAND_ORDER = ['<$0', '$0–$100K', '$100–$500K', '$500K–$1M', '$1M+'];

// Normalizza i nomi banda dal foglio → chiave canonica
const BAND_NORM = {
  '<$0':           '<$0',
  '<$100k':        '$0–$100K',
  '$100k - $500k': '$100–$500K',
  '$500k - $1m':   '$500K–$1M',
  '$1m+':          '$1M+',
};

function mapOU(raw) {
  const v = String(raw || '').trim().toLowerCase();
  if (!v) return null;
  if (v === 'south' || v === 'emea south') return 'SOUTH';
  if (v === 'iberia')                       return 'IBERIA';
  if (v === 'italy')                        return 'ITALY';
  if (v === 'egm')                          return 'EGM';
  if (v === 'middle east')                  return 'MIDEAST';
  return null;
}

// ── GET ───────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const cb     = (e && e.parameter && e.parameter.callback) || '';
    const action = (e && e.parameter && e.parameter.action)   || 'readiness';
    var payload;
    if (action === 'readiness') {
      payload = buildPayload();
    } else if (action === 'lookup') {
      const q  = (e && e.parameter && e.parameter.q)  || '';
      const qtr = (e && e.parameter && e.parameter.qtr) || 'CQ';
      payload = lookupAccounts(q, qtr);
    } else if (action === 'tdb') {
      const qtr   = (e && e.parameter && e.parameter.qtr)   || 'CQ';
      const logic = (e && e.parameter && e.parameter.logic) || 'Opportunity';
      payload = buildTDB(qtr, logic);
    } else {
      payload = { error: 'unknown action' };
    }
    const json = JSON.stringify(payload);
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST ──────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'writeForecast') {
      writeForecast(body.key, body.value);
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Payload principale ────────────────────────────────────────────────────────
function buildPayload() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const meta = readMeta(ss);
  return {
    generated_at:       new Date().toISOString(),
    timestamps:         meta.timestamps,
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

// ── _meta tab ─────────────────────────────────────────────────────────────────
function readMeta(ss) {
  const sheet = ss.getSheetByName(TABS.META);
  if (!sheet) return { timestamps: {}, forecasts: {} };
  const all = sheet.getDataRange().getValues();
  const ts  = all[1] || [];
  const forecasts = {};
  for (var i = 3; i < all.length; i++) {
    if (all[i][0]) forecasts[String(all[i][0])] = String(all[i][1]);
  }
  return {
    timestamps: {
      snowflake: ts[0] ? String(ts[0]) : null,
      org62:     ts[1] ? String(ts[1]) : null,
      finplan:   ts[2] ? String(ts[2]) : null,
      hc:        ts[3] ? String(ts[3]) : null,
    },
    forecasts: forecasts,
  };
}

function writeForecast(key, value) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TABS.META);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (var i = 3; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ── Parser principale ─────────────────────────────────────────────────────────
function parseTab(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log('Tab not found: ' + tabName); return emptyResult(); }

  const rows = sheet.getDataRange().getValues();

  // Per ogni OU: { total: {...}, bands: { bandKey: {...} } }
  const ouData = {};
  OU_KEYS.forEach(function(k) { ouData[k] = { total: null, bands: {} }; });

  var inSection = false;

  rows.forEach(function(row) {
    // IMPORTANTE: i dati iniziano dalla colonna 1 (col 0 sempre vuota)
    const cell = String(row[1] || '').trim();
    const cellLo = cell.toLowerCase();

    // Rileva riga header di ogni sezione
    if (cellLo === 'deal band') {
      inSection = true;
      return;
    }
    if (!inSection) return;

    // OU dalla colonna 15
    const ou = mapOU(String(row[15] || '').trim());
    if (!ou) return;

    // Riga Total → KPI aggregati
    if (cellLo === 'total') {
      const cov = parseCov(row[7]);
      ouData[ou].total = {
        acvPY:      nM(row[2]),
        forecast:   nM(row[3]),
        yoyFcst:    nPct(row[4]),
        pipe:       nM(row[5]),
        pipeGrowth: nPct(row[6]),
        pipeCov:    cov.current,
        histCov:    cov.hist,
        deals:      parseDeals(row[8]),
        bco:        parseBCO(row[9]),
        closedQTD:  nM(row[10]),
        yoyClosed:  nPct(row[11]),
      };
      return;
    }

    // Righe banda → dati per il grafico distribution
    const canonBand = BAND_NORM[cellLo];
    if (!canonBand) return;

    const cov = parseCov(row[7]);
    ouData[ou].bands[canonBand] = {
      acvPY:   nM(row[2]),
      pipe:    nM(row[5]),
      yoyPipe: nPct(row[6]),
      pipeCov: cov.current,
      histCov: cov.hist,
    };
  });

  // Costruisci risultato per OU
  const result = {};
  OU_KEYS.forEach(function(ouKey) {
    const d = ouData[ouKey];
    const t = d.total || {};

    const acvPY      = t.acvPY      || 0;
    const forecast   = t.forecast   || 0;
    const yoyFcst    = t.yoyFcst    || 0;
    const pipe       = t.pipe       || 0;
    const pipeGrowth = t.pipeGrowth || 0;
    const pipeCov    = t.pipeCov    || 0;
    const histCov    = t.histCov    || 0;
    const deals      = t.deals      || 0;
    const bco        = t.bco        || 0;
    const closedQTD  = t.closedQTD  || 0;
    const yoyClosed  = t.yoyClosed  || 0;

    result[ouKey] = {
      kpis: {
        acv_py:        { val: fmtM(acvPY),                     lbl: 'Q2 FY26 ACV (PY)' },
        forecast:      { val: fmtM(forecast),                  lbl: 'Q2 FY27 Forecast', editable: true },
        yoy:           { val: fmtPct(yoyFcst),                 lbl: 'Y/Y Forecast',         dir: yoyFcst    >= 0 ? 'pos' : 'neg' },
        pipeline:      { val: fmtM(pipe),                      lbl: 'Pipeline Q2 FY27' },
        pipe_growth:   { val: fmtPct(pipeGrowth),              lbl: 'Pipeline Growth Y/Y',  dir: pipeGrowth >= 0 ? 'pos' : 'neg' },
        pipe_cov:      { val: Math.abs(pipeCov).toFixed(1) + 'x', lbl: 'Pipe Coverage' },
        hist_pipe_cov: { val: Math.abs(histCov).toFixed(1) + 'x', lbl: 'Hist. Pipe Coverage' },
        deals_cmt:     { val: deals,                            lbl: '# Deals in Commitment' },
        bco_ae:        { val: bco > 0 ? fmtK(bco) : '—',      lbl: 'BCO per AE' },
        closed_qtd:    { val: fmtM(closedQTD),
                         delta: fmtPct(yoyClosed) + ' Y/Y',
                         lbl: 'Closed QTD FY27',               dir: yoyClosed  >= 0 ? 'pos' : 'neg' },
      },
      bands: BAND_ORDER.map(function(band) {
        const b = d.bands[band];
        if (!b) return { band: band, fcst: 0, pipe: 0, yoy: '—', dir: 'neu' };
        return {
          band: band,
          fcst: r1(b.acvPY),       // ACV PY come riferimento per-banda
          pipe: r1(b.pipe),
          yoy:  fmtPct(b.yoyPipe),
          dir:  band === '<$0' ? 'neg' : (b.yoyPipe >= 0 ? 'pos' : 'neg'),
        };
      }),
      aes: [],
      coverage: BAND_ORDER.map(function(band) {
        const b = d.bands[band] || {};
        return { band: band, commit: 0, fcst: r1(b.pipe || 0), fp: 0 };
      }),
    };
  });

  return result;
}

// ── Value parsers ─────────────────────────────────────────────────────────────

// nM: "$33.3" → 33.3  oppure  numero diretto → passthrough
function nM(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[$,\s]/g, '');  // togli $ , spazi (ma NON M/K per evitare strip su "EGM")
  return parseFloat(s) || 0;
}

// nPct: "12%" → 0.12 | "-100%" → -1.0 | "-" → 0 | 0.12 (già decimale) → 0.12
function nPct(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') {
    return Math.abs(v) <= 5 ? v : v / 100;
  }
  var s = String(v).replace(/[%\s]/g, '').trim();
  if (!s || s === '-') return 0;
  var n = parseFloat(s);
  if (isNaN(n)) return 0;
  return Math.abs(n) > 5 ? n / 100 : n;
}

// parseCov: "-1.3x  (1.9x)" → { current: -1.3, hist: 1.9 }
//           "2x  (1.7x)"    → { current: 2,    hist: 1.7 }
//           "-  (-)"        → { current: 0,    hist: 0   }
function parseCov(raw) {
  var s     = String(raw || '').trim();
  var curM  = s.match(/^([+-]?\d+\.?\d*)/);
  var histM = s.match(/\(([+-]?\d+\.?\d*)x?\)/);
  var cur   = curM  ? parseFloat(curM[1])  : 0;
  var hist  = histM ? parseFloat(histM[1]) : 0;
  if (isNaN(cur))  cur  = 0;
  if (isNaN(hist)) hist = 0;
  return { current: cur, hist: hist };
}

// parseDeals: "#Q2 Open Opps: 1,073" → 1073
function parseDeals(raw) {
  var s = String(raw || '');
  var m = s.match(/(\d[\d,]*)/);
  return m ? parseInt(m[1].replace(/,/g, ''), 10) : 0;
}

// parseBCO: "BCO per AE $57K, -18% Y/Y" → 57000
function parseBCO(raw) {
  var s = String(raw || '');
  if (s.indexOf('BCO') < 0) return 0;
  var kM = s.match(/\$(\d+\.?\d*)[Kk]/);
  var mM = s.match(/\$(\d+\.?\d*)[Mm]/);
  if (kM) return parseFloat(kM[1]) * 1000;
  if (mM) return parseFloat(mM[1]) * 1e6;
  return 0;
}

// ── Format helpers ────────────────────────────────────────────────────────────
function r1(v) { return Math.round(v * 10) / 10; }

function fmtM(v) {
  var abs = Math.abs(v);
  return (v < 0 ? '-' : '') + '$' + r1(abs) + 'M';
}

function fmtK(dollars) {
  if (dollars >= 1e6)  return '$' + r1(dollars / 1e6) + 'M';
  if (dollars >= 1000) return '$' + Math.round(dollars / 1000) + 'K';
  return '$' + Math.round(dollars);
}

// fmtPct: 0.12 → "+12%"  |  -0.03 → "-3%"
function fmtPct(v) {
  if (typeof v !== 'number' || isNaN(v)) return '—';
  var pct  = Math.abs(v) <= 5 ? v * 100 : v;
  var sign = pct >= 0 ? '+' : '';
  return sign + Math.round(pct) + '%';
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function emptyResult() {
  return OU_KEYS.reduce(function(a, k) { a[k] = emptyOU(); return a; }, {});
}
function emptyOU() {
  return {
    kpis: {},
    bands:    BAND_ORDER.map(function(b) { return { band: b, fcst: 0, pipe: 0, yoy: '—', dir: 'neu' }; }),
    aes:      [],
    coverage: BAND_ORDER.map(function(b) { return { band: b, commit: 0, fcst: 0, fp: 0 }; }),
  };
}

// ── Account Lookup ────────────────────────────────────────────────────────────
// Cerca account in Openpipe e [Org62] Commit per il quarter selezionato.
// qtr: 'CQ' → Q2 FY27 | 'NQ' → Q3 FY27
// Ritorna: { results: [{ name, pipe, commit }] } — max 20 risultati
function lookupAccounts(q, qtr) {
  if (!q || q.length < 2) return { results: [] };
  const qLow = q.toLowerCase();

  // Map quarter code → valori Sheet
  var fiscalYear = 'FY 2027';
  var fiscalQtr  = qtr === 'NQ' ? 'FQ 3' : 'FQ 2';

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Openpipe: aggrega OPENPIPE per COMBO_GLOBAL_COMPANY_NAME ──────────────
  var opSheet = ss.getSheetByName('Openpipe');
  var opRows  = opSheet.getDataRange().getValues();
  // col index (0-based): H=7 name, E=4 year, F=5 qtr, N=13 openpipe, L=11 mgr_fcst
  var pipeMap   = {};  // name → pipe $
  var commitMap = {};  // name → commit $
  for (var i = 1; i < opRows.length; i++) {
    var row = opRows[i];
    var name = String(row[7] || '').trim();
    if (!name || name.toLowerCase().indexOf(qLow) < 0) continue;
    var year = String(row[4] || '').trim();
    var fqtr = String(row[5] || '').trim();
    if (year !== fiscalYear || fqtr !== fiscalQtr) continue;
    var pipe = parseFloat(String(row[13] || '0').replace(/,/g, '')) || 0;
    var mgr  = String(row[11] || '').trim().toUpperCase();
    pipeMap[name] = (pipeMap[name] || 0) + pipe;
    if (mgr === 'IN') commitMap[name] = (commitMap[name] || 0) + pipe;
  }

  // ── [Org62] Commit: aggrega Forecast Amount per Global Company ────────────
  // col index: A=0 role, D=3 quarter (es. "Q2 FY27"), F=5 amount
  // Qui non abbiamo global company name direttamente — skip, usiamo solo openpipe
  // (il commit già aggregato sopra via MGR_FORCST_JUDG_TXT = IN)

  // ── Costruisci risultati ──────────────────────────────────────────────────
  var names = Object.keys(pipeMap);
  var results = names.map(function(n) {
    return {
      name:   n,
      pipe:   Math.round(pipeMap[n]) / 1000000,   // in $M
      commit: Math.round(commitMap[n] || 0) / 1000000,
    };
  }).sort(function(a, b) { return b.pipe - a.pipe; }).slice(0, 20);

  return { results: results };
}

// ── Top Deals by Dealband ─────────────────────────────────────────────────────
// Legge Openpipe e aggrega per OU + dealband.
// logic: 'Opportunity' | 'GlobalCompany' | 'ComboCompany'
// qtr:   'CQ' (FQ 2) | 'NQ' (FQ 3)
// Ritorna: { quarter, logic, ous: [ { key, label, bands: { lt100, '100to500', '500to1m', gt1m } } ] }
// Ogni banda: { open: [{name, pipe}], totalOpen, closed: [{name, closed}], totalClosed }
function buildTDB(qtr, logic) {
  var fiscalYear = 'FY 2027';
  var fiscalQtr  = qtr === 'NQ' ? 'FQ 3' : 'FQ 2';
  var quarter    = qtr === 'NQ' ? 'Q3 FY27' : 'Q2 FY27';

  // NQ + GlobalCompany non esiste nel Sheet
  if (qtr === 'NQ' && logic === 'GlobalCompany') {
    return { quarter: quarter, logic: logic, ous: [], notAvailable: true };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var opSheet = ss.getSheetByName('Openpipe');
  var rows = opSheet.getDataRange().getValues();

  // Colonne (0-based):
  // A=0 lvl2 (OU head), E=4 year, F=5 qtr
  // H=7 COMBO_GLOBAL_COMPANY_NAME, J=9 COMBO_COMPANY_NAME
  // L=11 MGR_FORCST, N=13 OPENPIPE
  // P=15 COMBO_GLOBAL_DEALBAND, Q=16 COMBO_COMPANY_DEALBAND

  // Mappa OU head → chiave
  var OU_MAP = {
    'marco hernansanz':          'SOUTH',
    'vanessa fortarezza':        'ITALY',
    'ana alonso muñumer':        'EGM',
    'mohammed alkhotani':        'MIDEAST',
  };
  var IB_PATTERN = /emea\s*-\s*south\s*-\s*ib/i;

  // Mappa banda Sheet → chiave JS
  var BAND_MAP = {
    'up to 100': 'lt100',
    '100-500':   '100to500',
    '500-1m':    '500to1m',
    '1m+':       'gt1m',
  };
  var BANDS = ['lt100', '100to500', '500to1m', 'gt1m'];
  var OU_KEYS_LIST = ['SOUTH', 'IBERIA', 'ITALY', 'EGM', 'MIDEAST'];
  var OU_LABELS = { SOUTH:'EMEA South', IBERIA:'Iberia', ITALY:'Italy', EGM:'EGM', MIDEAST:'Middle East' };

  // Struttura accumulatori: ouKey → bandKey → { pipe:{name→$}, commit:{name→$} }
  var acc = {};
  OU_KEYS_LIST.forEach(function(k) {
    acc[k] = {};
    BANDS.forEach(function(b) { acc[k][b] = { pipe:{}, commit:{} }; });
  });

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var year = String(row[4] || '').trim();
    var fq   = String(row[5] || '').trim();
    if (year !== fiscalYear || fq !== fiscalQtr) continue;

    // Determina OU da colonna B (OVA_LVL_03_USR_NM = OU head)
    var lvl3 = String(row[1] || '').trim();
    var ou;
    if (IB_PATTERN.test(lvl3)) {
      ou = 'IBERIA';
    } else {
      ou = OU_MAP[lvl3.toLowerCase()] || null;
      if (!ou) continue; // salta righe non mappabili (es. headers)
    }

    // Determina nome account e banda in base alla logica
    // Opportunity → OPPORTUNITY_DEALBAND (col K=10), COMBO_GLOBAL_COMPANY_NAME (col H=7)
    // GlobalCompany → COMBO_GLOBAL_DEALBAND (col P=15), COMBO_GLOBAL_COMPANY_NAME (col H=7)
    // ComboCompany  → COMBO_COMPANY_DEALBAND (col Q=16), COMBO_COMPANY_NAME (col J=9)
    var name, bandRaw;
    if (logic === 'Opportunity') {
      name    = String(row[7]  || '').trim();
      bandRaw = String(row[10] || '').trim().toLowerCase(); // OPPORTUNITY_DEALBAND
    } else if (logic === 'GlobalCompany') {
      name    = String(row[7]  || '').trim();
      bandRaw = String(row[15] || '').trim().toLowerCase(); // COMBO_GLOBAL_DEALBAND
    } else {
      // ComboCompany
      name    = String(row[9]  || '').trim();
      bandRaw = String(row[16] || '').trim().toLowerCase(); // COMBO_COMPANY_DEALBAND
    }
    if (!name) continue;
    var band = BAND_MAP[bandRaw];
    if (!band) continue;

    var pipe = parseFloat(String(row[13] || '0').replace(/,/g,'')) || 0;
    var mgr  = String(row[11] || '').trim().toUpperCase();

    // Accumula per OU specifico
    acc[ou][band].pipe[name] = (acc[ou][band].pipe[name] || 0) + pipe;
    if (mgr === 'IN') {
      acc[ou][band].commit[name] = (acc[ou][band].commit[name] || 0) + pipe;
    }
    // SOUTH aggrega sempre tutti gli OU figli
    acc['SOUTH'][band].pipe[name] = (acc['SOUTH'][band].pipe[name] || 0) + pipe;
    if (mgr === 'IN') acc['SOUTH'][band].commit[name] = (acc['SOUTH'][band].commit[name] || 0) + pipe;
  }

  // Costruisci output
  var ous = OU_KEYS_LIST.map(function(ouKey) {
    var bands = {};
    BANDS.forEach(function(b) {
      var pipeMap   = acc[ouKey][b].pipe;
      var commitMap = acc[ouKey][b].commit;

      // Top 10 open per pipe desc
      var openArr = Object.keys(pipeMap).map(function(n) {
        return { name: n, pipe: Math.round(pipeMap[n]) / 1000000 };
      }).sort(function(a,b){ return b.pipe - a.pipe; }).slice(0, 10);

      var totalOpen = Object.keys(pipeMap).reduce(function(s,n){ return s + pipeMap[n]; }, 0) / 1000000;

      // Top 10 commit per commit desc
      var closedArr = Object.keys(commitMap).map(function(n) {
        return { name: n, closed: Math.round(commitMap[n]) / 1000000 };
      }).sort(function(a,b){ return b.closed - a.closed; }).slice(0, 10);

      var totalClosed = Object.keys(commitMap).reduce(function(s,n){ return s + commitMap[n]; }, 0) / 1000000;

      bands[b] = {
        open:        openArr,
        totalOpen:   Math.round(totalOpen * 10) / 10,
        closed:      closedArr,
        totalClosed: Math.round(totalClosed * 10) / 10,
      };
    });
    return { key: ouKey, label: OU_LABELS[ouKey], bands: bands };
  });

  return { quarter: quarter, logic: logic, ous: ous };
}

// ── Test ──────────────────────────────────────────────────────────────────────
function testBuild() {
  const payload = buildPayload();
  Logger.log(JSON.stringify(payload.CQ.OPP, null, 2));
}

function testTDB() {
  Logger.log(JSON.stringify(buildTDB('CQ', 'Opportunity'), null, 2));
}

function testTDBCombo() {
  Logger.log(JSON.stringify(buildTDB('NQ', 'ComboCompany'), null, 2));
}

function testLookup() {
  Logger.log(JSON.stringify(lookupAccounts('omantel', 'NQ'), null, 2));
}
