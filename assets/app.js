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
function readSheet(file, sheetHint) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb     = XLSX.read(e.target.result, { type: 'array', cellDates: false });
        const target = sheetHint.trim().toLowerCase();
        const sName  = wb.SheetNames.find(s => s.trim().toLowerCase() === target)
                       ?? wb.SheetNames[0];
        const rows   = XLSX.utils.sheet_to_json(wb.Sheets[sName], { defval: null });
        resolve(rows);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
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
  const n = parseFloat(s);
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
  log('Chargement des 11 fichiers en parallèle...');

  const [mmall, marj1, marj2, marj3, baseRetail, liquidation, exclus,
         margeMin, paReception, recapMarjane, recapMmall] = await Promise.all([
    readSheet(files.mmall,        'catalogue'),
    readSheet(files.marj1,        'catalogue'),
    readSheet(files.marj2,        'catalogue'),
    readSheet(files.marj3,        'catalogue'),
    readSheet(files.baseRetail,   'Feuil1'),
    readSheet(files.liquidation,  'pour ajustement'),
    readSheet(files.exclus,       'Feuil1'),
    readSheet(files.margeMin,     'Feuil1'),
    readSheet(files.paReception,  'Feuil2'),
    readSheet(files.recapMarjane, 'Sheet1'),
    readSheet(files.recapMmall,   'Sheet 1'),
  ]);

  log(`Actif mmall chargé : ${mmall.length} lignes.`);

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

  // ── Step 4 : liquidation ──────────────────────────────────
  log('Étape 4 : filtre liquidation...');
  const liqSkus = new Set(
    liquidation.map(r => normKey(r['SKU'])).filter(Boolean)
  );
  df = df.filter(r => !liqSkus.has(normKey(r['SKU'])));
  log(`Après filtre liquidation : ${df.length} lignes.`);

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
  df = df.filter(r =>
    isNum(r['Prix Marjane 1']) || isNum(r['Prix Marjane 2']) || isNum(r['Prix Marjane 3'])
  );
  log(`Après filtre prix Marjane : ${df.length} lignes.`);

  // ── Step 8 : Min Prix Marjane ─────────────────────────────
  log('Étape 8 : Min Prix Marjane...');
  df = df.map(r => {
    const prices = [r['Prix Marjane 1'], r['Prix Marjane 2'], r['Prix Marjane 3']].filter(isNum);
    return { ...r, 'Min Prix Marjane': prices.length ? Math.min(...prices) : NaN };
  });

  // ── Step 9 : PA RECAP ─────────────────────────────────────
  log('Étape 9 : PA RECAP...');
  const mapPaRecapEan = buildLookup(recapMmall, 'EAN',          'PA_NET');
  const mapPaRecapCi  = buildLookup(recapMmall, 'CODE_INTERNE', 'PA_NET');
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
  const mapRayEan = buildLookup(recapMmall, 'EAN',          'LIB_RAY');
  const mapRayCi  = buildLookup(recapMmall, 'CODE_INTERNE', 'LIB_RAY');
  df = df.map(r => ({ ...r, 'RAYON': lookupFallback(r, 'GTIN', mapRayEan, 'REF', mapRayCi) }));

  // ── Step 17 : Fournisseur ─────────────────────────────────
  log('Étape 17 : Fournisseur...');
  const mapFourEan = buildLookup(recapMmall, 'EAN',          'LIB_FOURNISSEUR');
  const mapFourCi  = buildLookup(recapMmall, 'CODE_INTERNE', 'LIB_FOURNISSEUR');
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

  // ── Steps 21-22 : saise + etat (with EAN validation) ─────
  log('Étapes 21-22 : saise et etat...');
  const mapSaisEan = buildLookup(recapMmall, 'EAN',          'SAIS');
  const mapSaisCi  = buildLookup(recapMmall, 'CODE_INTERNE', 'SAIS');
  const mapEtatEan = buildLookup(recapMmall, 'EAN',          'ETAT');
  const mapEtatCi  = buildLookup(recapMmall, 'CODE_INTERNE', 'ETAT');
  const mapCiEan   = buildLookup(recapMmall, 'CODE_INTERNE', 'EAN');

  function lookupValidated(row, pm, fm) {
    const g = normKey(row['GTIN']);
    if (g && pm.has(g)) return pm.get(g);
    const ref = normKey(row['REF']);
    if (ref && fm.has(ref)) {
      if (normKey(mapCiEan.get(ref)) === g) return fm.get(ref);
    }
    return null;
  }

  df = df.map(r => ({
    ...r,
    saise: lookupValidated(r, mapSaisEan, mapSaisCi),
    etat:  lookupValidated(r, mapEtatEan, mapEtatCi),
  }));

  // ── Build catalogue sheet ─────────────────────────────────
  const prixBarre = df.length > 0 ? (findBarreCol(Object.keys(df[0])) ?? barreCol) : barreCol;
  df = df.map(r => ({ ...r, 'PRIX BARRÉ': r['PRIX BARRÉ'] ?? r[prixBarre] ?? null }));

  const CAT_COLS = [
    'MODIFIER','GTIN','REF','SKU','TITLE',
    'ATTRIBUTE CLE','ATTRIBUTE VALEUR','STOCK','PRIX','PRIX BARRÉ',
    'type',"l'oreal",'Liste ne pas toucher',
    'Prix Marjane 1','Stock Marjane 1',
    'Prix Marjane 2','Stock Marjane 2',
    'Prix Marjane 3','Stock Marjane 3',
    'Min Prix Marjane',
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

  const dfBaisse = df
    .filter(r => r.action === ACTION_DOWN)
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
