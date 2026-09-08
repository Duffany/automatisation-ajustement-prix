/* ============================================================
   Ajustement des Prix — Marjane Mall
   Client-side processing with SheetJS (xlsx)
   All 11 source files processed entirely in the browser.
   ============================================================ */

// ── Constants (mirror app.py) ──────────────────────────────
const COST_MULTIPLIER = 1.2;
const MARGE_TOLERANCE = 0.003;
const TYPE_FILTER     = 'LOCAL B1';
const SAIS_PERMANENT  = 'PER';
const ETAT_ACTIF      = 'Actif';
const ACTION_UP       = 'Augmenter';
const ACTION_DOWN     = 'Baisser';
const MODIFIER_FLAG   = 'Oui';

// ── State ──────────────────────────────────────────────────
const files = {
  mmall: null, marj1: null, marj2: null, marj3: null,
  baseRetail: null, liquidation: null, exclus: null,
  margeMin: null, paReception: null, recapMarjane: null, recapMmall: null,
};
let resultBlob = null;

// ── DOM ready ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setupUploadZones();
  document.getElementById('btn-process').addEventListener('click', runProcess);
  document.getElementById('btn-download').addEventListener('click', downloadResult);
});

function setupUploadZones() {
  document.querySelectorAll('.upload-zone').forEach(zone => {
    const key   = zone.dataset.key;
    const input = zone.querySelector('.file-input');
    const label = zone.querySelector('.file-name');

    input.addEventListener('change', () => handleFile(input.files[0], key, zone, label));
    zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
    zone.addEventListener('dragleave', ()  => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
      e.preventDefault(); zone.classList.remove('drag-over');
      handleFile(e.dataTransfer.files[0], key, zone, label);
    });
  });
}

function handleFile(file, key, zone, label) {
  if (!file) return;
  files[key] = file;
  label.textContent = file.name;
  zone.classList.add('done');
  updateCounter();
}

function updateCounter() {
  const count = Object.values(files).filter(Boolean).length;
  document.getElementById('count').textContent = count;
  document.getElementById('btn-process').disabled = count < 11;
}

// ── Logging ────────────────────────────────────────────────
function log(msg, type = '') {
  const block = document.getElementById('log-block');
  const card  = document.getElementById('output-card');
  card.classList.remove('hidden');
  const line = document.createElement('div');
  line.className = 'log-line' + (type ? ' ' + type : '');
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  block.appendChild(line);
  block.scrollTop = block.scrollHeight;
}

// ── Excel reading ──────────────────────────────────────────
function readWorkbook(file, opts) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        resolve(XLSX.read(e.target.result, { type: 'array', cellDates: false, ...(opts || {}) }));
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sheetFromWb(wb, sheetHint) {
  const target = sheetHint.trim().toLowerCase();
  const sName  = wb.SheetNames.find(s => s.trim().toLowerCase() === target) ?? wb.SheetNames[0];
  return wb.Sheets[sName];
}

// Certain exports (ex. Recap Marjane) placent des lignes de bandeau au-dessus
// de la vraie ligne d'en-tête. On cherche dans les ~25 premières lignes celle
// qui contient toutes les colonnes repères (comparaison insensible à la casse).
function findHeaderRow(ws, markers) {
  if (!markers || !markers.length || !ws['!ref']) return 0;
  const full  = XLSX.utils.decode_range(ws['!ref']);
  const probe = XLSX.utils.sheet_to_json(ws, {
    header: 1, defval: null, blankrows: true,
    range: { s: { r: full.s.r, c: full.s.c },
             e: { r: Math.min(full.s.r + 24, full.e.r), c: full.e.c } },
  });
  const want = markers.map(m => m.trim().toLowerCase());
  for (let i = 0; i < probe.length; i++) {
    const cells = probe[i].map(v => v === null ? '' : String(v).trim().toLowerCase());
    if (want.every(m => cells.includes(m))) return full.s.r + i;
  }
  return 0;
}

function readSheet(file, sheetHint, headerMarkers) {
  return readWorkbook(file).then(wb => {
    const ws = sheetFromWb(wb, sheetHint);
    const headerRow = findHeaderRow(ws, headerMarkers);
    return XLSX.utils.sheet_to_json(ws, { defval: null, range: headerRow });
  });
}

// Recap mmall is enormous (~292k rows × 43 cols → a 513 MB sheet XML). Running
// sheet_to_json on it hangs the browser for many minutes. Instead we parse in
// SheetJS "dense" mode (fast, ~40 s) and walk only the 7 columns we actually
// use, building every lookup Map in a couple of passes — then let the giant
// cell store be garbage-collected. Returns the resolved column indexes + rows.
const RECAP_COLS = ['EAN', 'CODE_INTERNE', 'PA_NET', 'LIB_RAY', 'LIB_FOURNISSEUR', 'SAIS', 'ETAT', 'PV_TTC'];

async function readRecapDense(file, sheetHint, wantCols) {
  const wb = await readWorkbook(file, {
    dense: true, cellStyles: false, cellNF: false, cellHTML: false,
  });
  const ws    = sheetFromWb(wb, sheetHint);
  const range = XLSX.utils.decode_range(ws['!ref']);
  const nRows = range.e.r + 1, nCols = range.e.c + 1;
  // Dense rows live under ws['!data'] (SheetJS ≥ 0.20) or directly on ws[r] (older).
  const dataRows = ws['!data'] || ws;
  // L'en-tête n'est pas forcément en ligne 1 (lignes de bandeau possibles) :
  // on cherche dans les 25 premières lignes celle qui contient toutes les colonnes.
  let hdrRow = -1, idx = {};
  for (let r = 0; r < Math.min(25, nRows) && hdrRow === -1; r++) {
    const hdr = dataRows[r] || [];
    const cand = {};
    for (let c = 0; c < nCols; c++) {
      const cell = hdr[c];
      if (cell && cell.v != null) {
        const name = String(cell.v).trim();
        if (wantCols.includes(name) && !(name in cand)) cand[name] = c;
      }
    }
    if (wantCols.every(w => w in cand)) { hdrRow = r; idx = cand; }
  }
  if (hdrRow === -1) {
    throw new Error(`Colonnes ${wantCols.join(', ')} introuvables dans Recap mmall.`);
  }
  return { nRows, dataRows, idx, hdrRow };
}

// Build a Map(normKey(keyCol) -> valueCol) from a dense recap context, using
// the exact same skip rules as buildLookup (first non-null value wins).
function denseLookup(ctx, keyCol, valueCol) {
  const { nRows, dataRows, idx, hdrRow } = ctx;
  const ki = idx[keyCol], vi = idx[valueCol];
  const map = new Map();
  for (let r = hdrRow + 1; r < nRows; r++) {
    const row = dataRows[r];
    if (!row) continue;
    const kc = row[ki];
    const k  = kc ? normKey(kc.v) : null;
    if (k === null) continue;
    const vc = row[vi];
    const v  = vc ? vc.v : undefined;
    if (v === null || v === undefined) continue;
    if (typeof v === 'number' && (isNaN(v) || !isFinite(v))) continue;
    if (!map.has(k)) map.set(k, v);
  }
  return map;
}

// Load every Recap-mmall lookup Map in one place, then drop the dense cell
// store so ~2 GB of parsed cells can be reclaimed before the pipeline continues.
async function loadRecapMaps(file, sheetHint) {
  const ctx = await readRecapDense(file, sheetHint, RECAP_COLS);
  const maps = {
    paEan:   denseLookup(ctx, 'EAN',          'PA_NET'),
    paCi:    denseLookup(ctx, 'CODE_INTERNE', 'PA_NET'),
    rayEan:  denseLookup(ctx, 'EAN',          'LIB_RAY'),
    rayCi:   denseLookup(ctx, 'CODE_INTERNE', 'LIB_RAY'),
    fourEan: denseLookup(ctx, 'EAN',          'LIB_FOURNISSEUR'),
    fourCi:  denseLookup(ctx, 'CODE_INTERNE', 'LIB_FOURNISSEUR'),
    saisEan: denseLookup(ctx, 'EAN',          'SAIS'),
    saisCi:  denseLookup(ctx, 'CODE_INTERNE', 'SAIS'),
    etatEan: denseLookup(ctx, 'EAN',          'ETAT'),
    etatCi:  denseLookup(ctx, 'CODE_INTERNE', 'ETAT'),
    pvTtcEan: denseLookup(ctx, 'EAN',          'PV_TTC'),
    pvTtcCi:  denseLookup(ctx, 'CODE_INTERNE', 'PV_TTC'),
  };
  return maps;
}

// ── Helpers (mirror app.py) ────────────────────────────────
function normKey(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') {
    if (!isFinite(value) || isNaN(value)) return null;
    return String(Math.round(value));
  }
  const s = String(value).trim();
  if (s === '' || s.toLowerCase() === 'nan') return null;
  // Number(), not parseFloat(): parseFloat silently truncates "3046449-JAUNE"
  // to 3046449, which can spuriously collide with an unrelated CODE_INTERNE.
  // Excel's VLOOKUP treats such values as literal text, never matching a pure
  // numeric key — Number() rejects the whole string instead, matching that.
  const n = Number(s);
  if (!isNaN(n) && isFinite(n)) {
    if (Number.isInteger(n)) return String(n);
    if (Math.abs(n - Math.round(n)) < 1e-6) return String(Math.round(n));
  }
  return s;
}

function isNum(v) {
  return v !== null && v !== undefined && typeof v === 'number' && isFinite(v) && !isNaN(v);
}

function toNum(v) {
  if (v === null || v === undefined) return NaN;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(',', '.').trim());
  return isNaN(n) ? NaN : n;
}

function buildLookup(rows, keyCol, valueCol) {
  const map = new Map();
  for (const row of rows) {
    const k = normKey(row[keyCol]);
    if (k === null) continue;
    const v = row[valueCol];
    if (v === null || v === undefined) continue;
    if (typeof v === 'number' && (isNaN(v) || !isFinite(v))) continue;
    if (!map.has(k)) map.set(k, v);
  }
  return map;
}

function lookupFallback(row, pk, pm, fk, fm) {
  const k1 = normKey(row[pk]);
  if (k1 !== null && pm.has(k1)) return pm.get(k1);
  const k2 = normKey(row[fk]);
  if (k2 !== null && fm.has(k2)) return fm.get(k2);
  return null;
}


function findBarreCol(keys) {
  return keys.find(c => c.toUpperCase().includes('BARR')) ?? null;
}

// ── Main processing pipeline ───────────────────────────────
async function runProcess() {
  document.getElementById('btn-process').disabled = true;
  document.getElementById('result-block').classList.add('hidden');
  document.getElementById('log-block').innerHTML = '';
  resultBlob = null;

  try {
    await pipeline();
  } catch (err) {
    log('ERREUR : ' + err.message, 'error');
    console.error(err);
  } finally {
    document.getElementById('btn-process').disabled = false;
  }
}

async function pipeline() {
  log('Chargement des 10 fichiers standards...');

  const [mmall, marj1, marj2, marj3, baseRetail, liquidation, exclus,
         margeMin, paReception, recapMarjane] = await Promise.all([
    readSheet(files.mmall,        'catalogue',       ['GTIN', 'PRIX']),
    readSheet(files.marj1,        'catalogue',       ['GTIN', 'PRIX']),
    readSheet(files.marj2,        'catalogue',       ['GTIN', 'PRIX']),
    readSheet(files.marj3,        'catalogue',       ['GTIN', 'PRIX']),
    readSheet(files.baseRetail,   'Feuil1',          ['GTIN_octopia', 'productid']),
    readSheet(files.liquidation,  'pour ajustement', ['SKU']),
    readSheet(files.exclus,       'Feuil1',          ['EAN']),
    readSheet(files.margeMin,     'Feuil1',          ['RAYON']),
    readSheet(files.paReception,  'Feuil2',          ['Code article']),
    readSheet(files.recapMarjane, 'Sheet1',          ['code interne', 'Gencode', 'Qte stock']),
  ]);

  log(`Actif mmall chargé : ${mmall.length} lignes.`);

  // Recap mmall is huge — parse it separately in dense mode (see loadRecapMaps).
  log('Analyse du fichier Recap mmall (volumineux, ~1 min)...', 'warn');
  await new Promise(r => setTimeout(r, 30));   // let the log paint before the parse blocks the UI
  const recapMaps = await loadRecapMaps(files.recapMmall, 'Sheet 1');
  log('Recap mmall analysé.');

  // Find PRIX BARRÉ column in mmall
  const barreCol = mmall.length > 0 ? findBarreCol(Object.keys(mmall[0])) : null;
  if (!barreCol) throw new Error("Colonne 'PRIX BARRE' introuvable dans Actif mmall.");

  // Work array — add internal columns
  let df = mmall.map(r => ({
    ...r,
    __orig_prix_barre: toNum(r[barreCol]),
    __orig_prix:       toNum(r['PRIX']),
  }));

  // ── Step 3 : type + filter LOCAL B1 ───────────────────────
  log(`Étape 3 : ajout colonne type, filtre ${TYPE_FILTER}...`);

  let typeCol = null;
  if (baseRetail.length > 0) {
    if ('Type' in baseRetail[0]) {
      typeCol = 'Type';
    } else {
      const cols = Object.keys(baseRetail[0]);
      if (cols.length > 8) {
        typeCol = cols[8];
        log(`Attention : colonne 'Type' absente, utilisation position 8 ('${typeCol}').`, 'warn');
      }
    }
  }

  const mapTypeGtin = typeCol ? buildLookup(baseRetail, 'GTIN_octopia', typeCol) : new Map();
  const mapTypeSku  = typeCol ? buildLookup(baseRetail, 'productid',    typeCol) : new Map();

  df = df.map(r => ({
    ...r,
    type: lookupFallback(r, 'GTIN', mapTypeGtin, 'SKU', mapTypeSku),
  }));
  df = df.filter(r => r.type && String(r.type).toUpperCase().includes(TYPE_FILTER));
  log(`Après filtre ${TYPE_FILTER} : ${df.length} lignes.`);

  // ── Step 4 : liquidation → "Liste ne pas toucher" flag ────
  // These SKUs stay in the catalogue and remain eligible for price decreases;
  // they're only blocked from price increases (see SKU augmentation filter below).
  log("Étape 4 : indicateur 'Liste ne pas toucher' (liquidation)...");
  const liqSkus = new Set(
    liquidation.map(r => normKey(r['SKU'])).filter(Boolean)
  );
  df = df.map(r => {
    const k = normKey(r['SKU']);
    return { ...r, 'Liste ne pas toucher': (k && liqSkus.has(k)) ? r['SKU'] : null };
  });
  log(`Indicateur liquidation posé sur ${df.filter(r => r['Liste ne pas toucher']).length} lignes.`);

  // ── Step 5 : exclus (l'Oréal) ─────────────────────────────
  log("Étape 5 : indicateur l'Oréal...");
  const exclusEans = new Set(
    exclus.map(r => normKey(r['EAN'])).filter(Boolean)
  );
  df = df.map(r => {
    const k = normKey(r['GTIN']);
    return { ...r, "l'oreal": (k && exclusEans.has(k)) ? k : null };
  });

  // ── Steps 6-7 : prix 3 magasins Marjane ──────────────────
  log('Étapes 6-7 : prix des 3 magasins Marjane...');
  for (let i = 0; i < 3; i++) {
    const store    = [marj1, marj2, marj3][i];
    const mapPrix  = buildLookup(store, 'GTIN', 'PRIX');
    const mapStock = buildLookup(store, 'GTIN', 'STOCK');
    df = df.map(r => ({
      ...r,
      [`Prix Marjane ${i+1}`]:  toNum(mapPrix.get(normKey(r['GTIN'])) ?? NaN),
      [`Stock Marjane ${i+1}`]: mapStock.get(normKey(r['GTIN'])) ?? null,
    }));
  }
  // Step 9 (PV TTC Recap): 4th acceptable price source, both for the survival
  // filter and for Min Prix Marjane (step 8 below). Validated against the
  // manual reference once the normKey() REF-fallback bug above was fixed.
  const mapPvTtcEan = recapMaps.pvTtcEan;
  const mapPvTtcCi  = recapMaps.pvTtcCi;
  df = df.map(r => ({
    ...r,
    'PV TTC Recap': toNum(lookupFallback(r, 'GTIN', mapPvTtcEan, 'REF', mapPvTtcCi) ?? NaN),
  }));

  df = df.filter(r =>
    isNum(r['Prix Marjane 1']) || isNum(r['Prix Marjane 2']) || isNum(r['Prix Marjane 3']) ||
    isNum(r['PV TTC Recap'])
  );
  log(`Après filtre prix Marjane + PV TTC Recap : ${df.length} lignes.`);

  // ── Step 8 : Min Prix Marjane (3 magasins + PV TTC Recap) ─
  log('Étape 8 : Min Prix Marjane...');
  df = df.map(r => {
    const prices = [r['Prix Marjane 1'], r['Prix Marjane 2'], r['Prix Marjane 3'], r['PV TTC Recap']].filter(isNum);
    return { ...r, 'Min Prix Marjane': prices.length ? Math.min(...prices) : NaN };
  });

  // ── Step 9 : PA RECAP ─────────────────────────────────────
  log('Étape 9 : PA RECAP...');
  const mapPaRecapEan = recapMaps.paEan;
  const mapPaRecapCi  = recapMaps.paCi;
  df = df.map(r => ({
    ...r,
    'PA RECAP': toNum(lookupFallback(r, 'GTIN', mapPaRecapEan, 'REF', mapPaRecapCi) ?? NaN),
  }));

  // ── Step 10 : PA Réception ────────────────────────────────
  log('Étape 10 : PA Réception...');
  const mapPaRec = buildLookup(paReception, 'Code article', 'Prix revient unitaire');
  df = df.map(r => ({
    ...r,
    'PA Réception': toNum(lookupFallback(r, 'GTIN', mapPaRec, 'REF', mapPaRec) ?? NaN),
  }));

  // ── Step 11 : PA Moyen ────────────────────────────────────
  log('Étape 11 : PA Moyen...');
  const rmRows = recapMarjane.map(r => {
    const qte = toNum(r['Qte stock']);
    const val = toNum(r['Valeur du stock']);
    return { ...r, __pam: (isNum(qte) && qte !== 0 && isNum(val)) ? val / qte : NaN };
  });
  const mapMoyGencode = buildLookup(rmRows, 'Gencode',     '__pam');
  const mapMoyCi      = buildLookup(rmRows, 'code interne','__pam');
  df = df.map(r => ({
    ...r,
    'PA Moyen': toNum(lookupFallback(r, 'GTIN', mapMoyGencode, 'REF', mapMoyCi) ?? NaN),
  }));

  // ── Steps 13-15 : Marges ──────────────────────────────────
  log('Étapes 13-15 : calcul des marges...');
  df = df.map(r => {
    const minp = r['Min Prix Marjane'];
    const marge = (pa) => (isNum(minp) && isNum(pa)) ? (minp - pa * COST_MULTIPLIER) / minp : NaN;
    return {
      ...r,
      'MARGE RECAP':     marge(r['PA RECAP']),
      'MARGE Réception': marge(r['PA Réception']),
      'MARGE Moyen':     marge(r['PA Moyen']),
    };
  });

  // ── Step 16 : RAYON ───────────────────────────────────────
  log('Étape 16 : RAYON...');
  const mapRayEan = recapMaps.rayEan;
  const mapRayCi  = recapMaps.rayCi;
  df = df.map(r => ({ ...r, 'RAYON': lookupFallback(r, 'GTIN', mapRayEan, 'REF', mapRayCi) }));

  // ── Step 17 : Fournisseur ─────────────────────────────────
  log('Étape 17 : Fournisseur...');
  const mapFourEan = recapMaps.fourEan;
  const mapFourCi  = recapMaps.fourCi;
  df = df.map(r => ({ ...r, 'Fournisseur': lookupFallback(r, 'GTIN', mapFourEan, 'REF', mapFourCi) }));

  // ── Step 18 : MIN % ───────────────────────────────────────
  log('Étape 18 : MIN %...');
  const mapMinPct = new Map();
  if (margeMin.length > 0) {
    const keys      = Object.keys(margeMin[0]);
    const rayonCol  = keys.find(c => c.trim().toUpperCase() === 'RAYON');
    const minPctCol = keys.find(c => c.toUpperCase().includes('MIN') && c.includes('%'))
                   ?? keys.find(c => c.trim().toUpperCase() === 'MIN %');
    if (!rayonCol || !minPctCol) {
      log("Attention : colonnes RAYON ou MIN % absentes dans Marge min.", 'warn');
    } else {
      for (const r of margeMin) {
        const ray = r[rayonCol];
        if (ray === null || ray === undefined) continue;
        const key = String(ray).trim().toUpperCase();
        const pct = toNum(r[minPctCol]);
        if (!mapMinPct.has(key) && isNum(pct)) mapMinPct.set(key, pct);
      }
    }
  }
  df = df.map(r => {
    const ray = r['RAYON'];
    return { ...r, 'MIN %': ray ? (mapMinPct.get(String(ray).trim().toUpperCase()) ?? NaN) : NaN };
  });

  // ── Step 19 : Ecart Vs BO MM ──────────────────────────────
  log('Étape 19 : Ecart Vs BO MM...');
  df = df.map(r => {
    const m = r['Min Prix Marjane'];
    const p = toNum(r['PRIX']);
    return { ...r, 'Ecart Vs BO MM': (isNum(m) && isNum(p)) ? m - p : 0 };
  });

  // ── Step 20 : action ─────────────────────────────────────
  log("Étape 20 : calcul de l'action...");
  df = df.map(r => {
    const ecart = r['Ecart Vs BO MM'];
    if (!isNum(ecart) || ecart === 0) return { ...r, action: '' };
    if (ecart > 0) return { ...r, action: ACTION_UP };
    // ecart < 0 : check marges in priority order
    const minPct = r['MIN %'];
    for (const marge of [r['MARGE Moyen'], r['MARGE Réception'], r['MARGE RECAP']]) {
      if (isNum(marge)) {
        if (!isNum(minPct)) return { ...r, action: '' };
        if (marge >= minPct || (minPct - marge) < MARGE_TOLERANCE)
          return { ...r, action: ACTION_DOWN };
        return { ...r, action: '' };
      }
    }
    return { ...r, action: '' };
  });

  // ── Steps 21-22 : saise + etat ───────────────────────────
  // Lookup on GTIN (EAN) first, then fall back to CODE_INTERNE — same pattern
  // as RAYON / Fournisseur. No EAN-equality guard: an earlier attempt at one
  // was found to reject genuinely correct matches (confirmed by comparing
  // product titles — Récap frequently lists a different EAN than Actif mmall
  // for the exact same product under a shared Code Interne). The manual
  // reference leaving these blank looks like a gap in its own VLOOKUP, not a
  // rule to replicate.
  log('Étapes 21-22 : saise et etat...');
  const mapSaisEan = recapMaps.saisEan;
  const mapSaisCi  = recapMaps.saisCi;
  const mapEtatEan = recapMaps.etatEan;
  const mapEtatCi  = recapMaps.etatCi;

  df = df.map(r => ({
    ...r,
    saise: lookupFallback(r, 'GTIN', mapSaisEan, 'REF', mapSaisCi),
    etat:  lookupFallback(r, 'GTIN', mapEtatEan, 'REF', mapEtatCi),
  }));

  // ── Build catalogue sheet ─────────────────────────────────
  // PRIX and PRIX BARRÉ are stored as text in the source file; emit them as
  // numbers (matching the reference workbook). STOCK stays text, also as in
  // the reference. __orig_prix / __orig_prix_barre are the numeric versions.
  df = df.map(r => ({
    ...r,
    'PRIX':       isNum(r.__orig_prix)       ? r.__orig_prix       : null,
    'PRIX BARRÉ': isNum(r.__orig_prix_barre) ? r.__orig_prix_barre : null,
  }));

  const CAT_COLS = [
    'MODIFIER','GTIN','REF','SKU','TITLE',
    'ATTRIBUTE CLE','ATTRIBUTE VALEUR','STOCK','PRIX','PRIX BARRÉ',
    'type',"l'oreal",'Liste ne pas toucher',
    'Prix Marjane 1','Stock Marjane 1',
    'Prix Marjane 2','Stock Marjane 2',
    'Prix Marjane 3','Stock Marjane 3',
    'Min Prix Marjane','PV TTC Recap',
    'PA RECAP','PA Réception','PA Moyen',
    'MARGE RECAP','MARGE Réception','MARGE Moyen',
    'RAYON','Fournisseur','MIN %',
    'Ecart Vs BO MM','action','saise','etat',
  ];
  log(`Feuille catalogue : ${df.length} lignes.`);

  // ── Step 23 : SKU baisse ──────────────────────────────────
  log('Étape 23 : feuille SKU baisse...');
  const BAISSE_COLS = ['MODIFIER','GTIN','REF','SKU','TITLE',
    'ATTRIBUTE CLE','ATTRIBUTE VALEUR','STOCK','NV PRIX','PRIX BARRÉ'];

  // Rule : products on the "exclus" list (flagged in "l'oreal") are excluded
  // from destockage — removed from SKU baisse only, kept for augmentation.
  const dfBaisse = df
    .filter(r => r.action === ACTION_DOWN && !r["l'oreal"])
    .map(r => ({
      'MODIFIER':        MODIFIER_FLAG,
      'GTIN':            r['GTIN'],
      'REF':             r['REF'],
      'SKU':             r['SKU'],
      'TITLE':           r['TITLE'],
      'ATTRIBUTE CLE':   r['ATTRIBUTE CLE'],
      'ATTRIBUTE VALEUR':r['ATTRIBUTE VALEUR'],
      'STOCK':           r['STOCK'],
      'NV PRIX':         r['Min Prix Marjane'],
      'PRIX BARRÉ':      r.__orig_prix,
    }));
  log(`Feuille SKU baisse : ${dfBaisse.length} lignes.`);

  // ── Step 24 : SKU augmentation ────────────────────────────
  log('Étape 24 : feuille SKU augmentation...');
  const AUG_COLS = ['MODIFIER','GTIN','REF','SKU','TITLE',
    'ATTRIBUTE CLE','ATTRIBUTE VALEUR','STOCK','NV PRIX','NV PRIX BARRÉ'];

  // "Liste ne pas toucher" is informational only (matches the manual reference,
  // which flags these rows via a live VLOOKUP but does not exclude them here).
  const dfAugment = df
    .filter(r => r.action === ACTION_UP && r.saise === SAIS_PERMANENT && r.etat === ETAT_ACTIF)
    .map(r => {
      const mp = r['Min Prix Marjane'];
      const ob = r.__orig_prix_barre;
      const op = r.__orig_prix;
      const nvPB = (isNum(mp) && isNum(ob) && isNum(op) && op !== 0) ? mp * ob / op : null;
      return {
        'MODIFIER':        MODIFIER_FLAG,
        'GTIN':            r['GTIN'],
        'REF':             r['REF'],
        'SKU':             r['SKU'],
        'TITLE':           r['TITLE'],
        'ATTRIBUTE CLE':   r['ATTRIBUTE CLE'],
        'ATTRIBUTE VALEUR':r['ATTRIBUTE VALEUR'],
        'STOCK':           r['STOCK'],
        'NV PRIX':         mp,
        'NV PRIX BARRÉ':   nvPB,
      };
    });
  log(`Feuille SKU augmentation : ${dfAugment.length} lignes.`);

  // ── Build Excel output ────────────────────────────────────
  log('Génération du fichier Excel...');
  const wb = XLSX.utils.book_new();

  const toSheet = (rows, cols) => {
    const data = rows.map(r => cols.map(c => r[c] ?? null));
    return XLSX.utils.aoa_to_sheet([cols, ...data]);
  };

  XLSX.utils.book_append_sheet(wb, toSheet(df,        CAT_COLS),    'catalogue');
  XLSX.utils.book_append_sheet(wb, toSheet(dfBaisse,  BAISSE_COLS), 'SKU baisse');
  XLSX.utils.book_append_sheet(wb, toSheet(dfAugment, AUG_COLS),    'SKU augmentation');

  resultBlob = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  log('Traitement terminé avec succès !', 'info');

  // Show results
  document.getElementById('stat-cat').textContent    = df.length.toLocaleString('fr-FR');
  document.getElementById('stat-baisse').textContent = dfBaisse.length.toLocaleString('fr-FR');
  document.getElementById('stat-aug').textContent    = dfAugment.length.toLocaleString('fr-FR');
  document.getElementById('result-block').classList.remove('hidden');
}

function downloadResult() {
  if (!resultBlob) return;
  const blob = new Blob([resultBlob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'Ajustement_prix_resultat.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}
