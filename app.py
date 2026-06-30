"""
Ajustement des Prix - Marjane Mall
Application d'automatisation du workflow d'alignement des prix.

Lance avec: python app.py
Dependances: gradio, pandas, numpy, openpyxl
"""

import os
import traceback
import tempfile

import numpy as np
import pandas as pd
import gradio as gr


# ---------------------------------------------------------------------------
# Configuration — update here when source files change sheet or column names
# ---------------------------------------------------------------------------

# ── Input sheet names (matched case- and whitespace-insensitively at load) ──
SHEET_ACTIF         = "catalogue"        # Actif mmall + 3 Actif marjane
SHEET_BASE_RETAIL   = "Feuil1"
SHEET_LIQUIDATION   = "pour ajustement"  # trailing space in source is stripped
SHEET_EXCLUS        = "Feuil1"
SHEET_MARGE_MIN     = "Feuil1"
SHEET_PA_RECEPTION  = "Feuil2"
SHEET_RECAP_MARJANE = "Sheet1"
SHEET_RECAP_MMALL   = "Sheet 1"

# ── Column names — Actif catalogue (mmall + marjane stores) ──
COL_MODIFIER = "MODIFIER"
COL_GTIN     = "GTIN"
COL_REF      = "REF"
COL_SKU      = "SKU"
COL_PRIX     = "PRIX"
COL_STOCK    = "STOCK"
COL_TITLE    = "TITLE"
COL_ATTR_CLE = "ATTRIBUTE CLE"
COL_ATTR_VAL = "ATTRIBUTE VALEUR"

# ── Column names — Exclus / Liquidation files ──
COL_EXCL_EAN = "EAN"
COL_LIQ_SKU  = COL_SKU          # same column name as in the catalogue

# ── Column names — Base retail ──
COL_BR_GTIN = "GTIN_octopia"
COL_BR_SKU  = "productid"
COL_BR_TYPE = "Type"             # previously referenced by positional index 8

# ── Column names — Recap mmall ──
COL_MM_EAN      = "EAN"
COL_MM_CI       = "CODE_INTERNE"
COL_MM_PA_NET   = "PA_NET"
COL_MM_RAYON    = "LIB_RAY"
COL_MM_FOURNISS = "LIB_FOURNISSEUR"
COL_MM_SAIS     = "SAIS"
COL_MM_ETAT     = "ETAT"

# ── Column names — PA Reception ──
COL_PA_CODE = "Code article"
COL_PA_PRIX = "Prix revient unitaire"

# ── Column names — Recap marjane ──
COL_MJ_GENCODE = "Gencode"
COL_MJ_CI      = "code interne"
COL_MJ_QTE     = "Qte stock"
COL_MJ_VAL     = "Valeur du stock"

# ── Business logic constants ──
COST_MULTIPLIER = 1.2    # PA × 1.2 : 20 % uplift applied before margin calc
MARGE_TOLERANCE = 0.003  # within 0.3 pp of MIN % still qualifies as Baisser
TYPE_FILTER     = "LOCAL B1"  # product type substring required for inclusion
SAIS_PERMANENT  = "PER"       # saison code meaning permanent range
ETAT_ACTIF      = "Actif"     # etat value meaning product is active
ACTION_UP       = "Augmenter"
ACTION_DOWN     = "Baisser"
MODIFIER_FLAG   = "Oui"

# ── Output sheet names ──
OUT_SHEET_CAT    = "catalogue"
OUT_SHEET_BAISSE = "SKU baisse"
OUT_SHEET_AUG    = "SKU augmentation"

# ── Output-only column names (written by this app, read by downstream) ──
OUT_COL_NV_PRIX       = "NV PRIX"
OUT_COL_PRIX_BARRE    = "PRIX BARRÉ"
OUT_COL_NV_PRIX_BARRE = "NV PRIX BARRÉ"

# ── Output column lists (defined once, shared by populated and empty-df branches) ──
OUT_COLS_BAISSE = [
    COL_MODIFIER, COL_GTIN, COL_REF, COL_SKU, COL_TITLE,
    COL_ATTR_CLE, COL_ATTR_VAL, COL_STOCK, OUT_COL_NV_PRIX, OUT_COL_PRIX_BARRE,
]
OUT_COLS_AUG = [
    COL_MODIFIER, COL_GTIN, COL_REF, COL_SKU, COL_TITLE,
    COL_ATTR_CLE, COL_ATTR_VAL, COL_STOCK, OUT_COL_NV_PRIX, OUT_COL_NV_PRIX_BARRE,
]

N_INPUT_FILES = 11   # total number of source files (reflected in UI copy)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _path(file_obj):
    """Return a usable filesystem path from a Gradio file upload object."""
    if file_obj is None:
        return None
    if isinstance(file_obj, str):
        return file_obj
    name = getattr(file_obj, "name", None)
    if name:
        return name
    return file_obj


def _read_sheet(path, sheet_hint):
    """Read a sheet matching sheet_hint case- and whitespace-insensitively.
    Falls back to the first sheet if the hint matches nothing."""
    resolved = _path(path)
    xf = pd.ExcelFile(resolved)
    target = sheet_hint.strip().lower()
    matched = next(
        (s for s in xf.sheet_names if s.strip().lower() == target),
        xf.sheet_names[0],
    )
    return pd.read_excel(xf, sheet_name=matched)


def norm_key(value):
    """Normalize an EAN/GTIN/code value to a clean string key.

    Excel reads these as floats (e.g. 3600542298599.0). We convert to int
    then to string so that the float artifact is removed. Non numeric values
    are stripped strings. NaN/empty -> None.
    """
    if value is None:
        return None
    if isinstance(value, float) and np.isnan(value):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (int, np.integer)):
        return str(int(value))

    if isinstance(value, (float, np.floating)):
        if float(value).is_integer():
            return str(int(value))
        return str(value)

    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return None
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except (TypeError, ValueError):
        pass
    return s


def build_lookup(df, key_col, value_col):
    """Build a dict mapping normalized(key) -> first non-NaN value.

    Returns an empty dict (silently) when a column is missing so callers
    receive NaN for every row rather than a hard crash; the caller logs a
    warning when the column is critical.
    """
    lookup = {}
    if key_col not in df.columns or value_col not in df.columns:
        return lookup
    for key_raw, val in zip(df[key_col], df[value_col]):
        k = norm_key(key_raw)
        if k is None:
            continue
        if val is None:
            continue
        try:
            if pd.isna(val):
                continue
        except (TypeError, ValueError):
            pass
        if k not in lookup:
            lookup[k] = val
    return lookup


def lookup_with_fallback(row, primary_col, primary_map, fallback_col, fallback_map):
    """Lookup using primary key first; fall back to a secondary key/map."""
    k1 = norm_key(row.get(primary_col))
    if k1 is not None and k1 in primary_map:
        return primary_map[k1]
    k2 = norm_key(row.get(fallback_col))
    if k2 is not None and k2 in fallback_map:
        return fallback_map[k2]
    return np.nan


def find_barre_col(df):
    """Find the 'PRIX BARRE' column dynamically (handles accents/casing)."""
    return next((c for c in df.columns if "BARR" in str(c).upper()), None)


def is_number(value):
    """True if value is a real number (not NaN, not None, not string)."""
    if value is None:
        return False
    if isinstance(value, (int, np.integer)):
        return True
    if isinstance(value, (float, np.floating)):
        return not np.isnan(value)
    return False


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def process(
    f_mmall, f_marj1, f_marj2, f_marj3,
    f_base_retail, f_liquidation, f_exclus,
    f_marge_min, f_pa_reception, f_recap_marjane, f_recap_mmall,
    progress=None,
):
    """Run the full price adjustment workflow and return (df_cat, df_baisse,
    df_augment, messages list)."""

    msgs = []

    def log(m):
        msgs.append(m)
        if progress is not None:
            try:
                progress(m)
            except Exception:
                pass

    # --- Load files ---------------------------------------------------------
    log("Chargement des fichiers...")

    mmall         = _read_sheet(f_mmall,         SHEET_ACTIF)
    marj1         = _read_sheet(f_marj1,         SHEET_ACTIF)
    marj2         = _read_sheet(f_marj2,         SHEET_ACTIF)
    marj3         = _read_sheet(f_marj3,         SHEET_ACTIF)
    base_retail   = _read_sheet(f_base_retail,   SHEET_BASE_RETAIL)
    liquidation   = _read_sheet(f_liquidation,   SHEET_LIQUIDATION)
    exclus        = _read_sheet(f_exclus,        SHEET_EXCLUS)
    marge_min     = _read_sheet(f_marge_min,     SHEET_MARGE_MIN)
    pa_reception  = _read_sheet(f_pa_reception,  SHEET_PA_RECEPTION)
    recap_marjane = _read_sheet(f_recap_marjane, SHEET_RECAP_MARJANE)
    recap_mmall   = _read_sheet(f_recap_mmall,   SHEET_RECAP_MMALL)

    # Working DataFrame (Step 1)
    df = mmall.copy()
    log(f"Actif mmall charge: {len(df)} lignes.")

    # Capture original PRIX and PRIX BARRE for later sheets
    barre_col = find_barre_col(df)
    if barre_col is None:
        raise ValueError("Colonne 'PRIX BARRE' introuvable dans Actif mmall.")
    df["__orig_prix_barre"] = pd.to_numeric(df[barre_col], errors="coerce")
    df["__orig_prix"]       = pd.to_numeric(df[COL_PRIX],  errors="coerce")

    # --- Step 3 : Add 'type' column -----------------------------------------
    log(f"Etape 3 : ajout de la colonne 'type' et filtre {TYPE_FILTER}...")

    # Resolve type column by name; fall back to positional index 8 with warning
    if COL_BR_TYPE in base_retail.columns:
        type_col = COL_BR_TYPE
    elif len(base_retail.columns) > 8:
        type_col = base_retail.columns[8]
        log(f"Attention : colonne '{COL_BR_TYPE}' introuvable dans Base retail, "
            f"utilisation de la colonne en position 8 ('{type_col}').")
    else:
        raise ValueError(
            f"Colonne '{COL_BR_TYPE}' introuvable dans Base retail "
            f"et le fichier contient moins de 9 colonnes."
        )

    map_type_gtin = build_lookup(base_retail, COL_BR_GTIN, type_col)
    map_type_sku  = build_lookup(base_retail, COL_BR_SKU,  type_col)

    if not map_type_gtin and not map_type_sku:
        log(f"Attention : aucune correspondance dans Base retail "
            f"(colonnes '{COL_BR_GTIN}' / '{COL_BR_SKU}' introuvables ?).")

    df["type"] = df.apply(
        lambda r: lookup_with_fallback(r, COL_GTIN, map_type_gtin, COL_SKU, map_type_sku),
        axis=1,
    )

    def has_local_b1(v):
        if not isinstance(v, str):
            return False
        return TYPE_FILTER in v.upper()

    df = df[df["type"].apply(has_local_b1)].copy()
    log(f"Apres filtre {TYPE_FILTER}: {len(df)} lignes.")

    # --- Step 4 : Filter liquidation ----------------------------------------
    log("Etape 4 : filtre liste liquidation...")
    liq_skus = set()
    for v in liquidation.get(COL_LIQ_SKU, pd.Series(dtype=object)):
        k = norm_key(v)
        if k is not None:
            liq_skus.add(k)

    df["Liste ne pas toucher"] = df[COL_SKU].apply(
        lambda s: norm_key(s) if norm_key(s) in liq_skus else np.nan
    )
    df = df[df["Liste ne pas toucher"].isna()].copy()
    log(f"Apres filtre liquidation: {len(df)} lignes.")

    # --- Step 5 : exclus indicator "l'oreal" --------------------------------
    log("Etape 5 : indicateur l'oreal (exclus)...")
    exclus_eans = set()
    for v in exclus.get(COL_EXCL_EAN, pd.Series(dtype=object)):
        k = norm_key(v)
        if k is not None:
            exclus_eans.add(k)

    def loreal_indicator(g):
        k = norm_key(g)
        return k if (k is not None and k in exclus_eans) else np.nan

    df["l'oreal"] = df[COL_GTIN].apply(loreal_indicator)

    # --- Steps 6-7 : prices from 3 Marjane stores ---------------------------
    log("Etapes 6-7 : ajout des prix des 3 magasins Marjane...")
    stores = [marj1, marj2, marj3]
    for i, store in enumerate(stores, start=1):
        map_prix  = build_lookup(store, COL_GTIN, COL_PRIX)
        map_stock = build_lookup(store, COL_GTIN, COL_STOCK)
        df[f"Prix Marjane {i}"]  = df[COL_GTIN].apply(lambda g: map_prix.get(norm_key(g),  np.nan))
        df[f"Stock Marjane {i}"] = df[COL_GTIN].apply(lambda g: map_stock.get(norm_key(g), np.nan))

    prix_cols = ["Prix Marjane 1", "Prix Marjane 2", "Prix Marjane 3"]
    for c in prix_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[~df[prix_cols].isna().all(axis=1)].copy()
    log(f"Apres filtre prix Marjane: {len(df)} lignes.")

    # --- Step 8 : Min Prix Marjane ------------------------------------------
    log("Etape 8 : Min Prix Marjane...")
    df["Min Prix Marjane"] = df[prix_cols].min(axis=1, skipna=True)

    # --- Step 9 : PA RECAP (Recap mmall) ------------------------------------
    log("Etape 9 : PA RECAP...")
    map_pa_recap_ean = build_lookup(recap_mmall, COL_MM_EAN, COL_MM_PA_NET)
    map_pa_recap_ci  = build_lookup(recap_mmall, COL_MM_CI,  COL_MM_PA_NET)
    df["PA RECAP"] = df.apply(
        lambda r: lookup_with_fallback(r, COL_GTIN, map_pa_recap_ean, COL_REF, map_pa_recap_ci),
        axis=1,
    )
    df["PA RECAP"] = pd.to_numeric(df["PA RECAP"], errors="coerce")

    # --- Step 10 : PA Reception (PA reception) ------------------------------
    log("Etape 10 : PA Reception...")
    map_pa_rec = build_lookup(pa_reception, COL_PA_CODE, COL_PA_PRIX)
    df["PA Réception"] = df.apply(
        lambda r: lookup_with_fallback(r, COL_GTIN, map_pa_rec, COL_REF, map_pa_rec),
        axis=1,
    )
    df["PA Réception"] = pd.to_numeric(df["PA Réception"], errors="coerce")

    # --- Step 11 : PA Moyen (Recap marjane) ---------------------------------
    log("Etape 11 : PA Moyen...")
    rm  = recap_marjane.copy()
    qte = pd.to_numeric(rm[COL_MJ_QTE], errors="coerce")
    val = pd.to_numeric(rm[COL_MJ_VAL], errors="coerce")
    rm["__pa_moyen_calc"] = np.where((qte == 0) | qte.isna(), np.nan, val / qte)
    map_moy_gencode = build_lookup(rm, COL_MJ_GENCODE, "__pa_moyen_calc")
    map_moy_ci      = build_lookup(rm, COL_MJ_CI,      "__pa_moyen_calc")
    df["PA Moyen"] = df.apply(
        lambda r: lookup_with_fallback(r, COL_GTIN, map_moy_gencode, COL_REF, map_moy_ci),
        axis=1,
    )
    df["PA Moyen"] = pd.to_numeric(df["PA Moyen"], errors="coerce")

    # --- Steps 13-15 : Marges -----------------------------------------------
    log("Etapes 13-15 : calcul des marges...")
    minp = df["Min Prix Marjane"]
    df["MARGE RECAP"]     = (minp - df["PA RECAP"]     * COST_MULTIPLIER) / minp
    df["MARGE Réception"] = (minp - df["PA Réception"] * COST_MULTIPLIER) / minp
    df["MARGE Moyen"]     = (minp - df["PA Moyen"]     * COST_MULTIPLIER) / minp

    # --- Step 16 : RAYON ----------------------------------------------------
    log("Etape 16 : RAYON...")
    map_ray_ean = build_lookup(recap_mmall, COL_MM_EAN, COL_MM_RAYON)
    map_ray_ci  = build_lookup(recap_mmall, COL_MM_CI,  COL_MM_RAYON)
    df["RAYON"] = df.apply(
        lambda r: lookup_with_fallback(r, COL_GTIN, map_ray_ean, COL_REF, map_ray_ci),
        axis=1,
    )

    # --- Step 17 : Fournisseur ----------------------------------------------
    log("Etape 17 : Fournisseur...")
    map_four_ean = build_lookup(recap_mmall, COL_MM_EAN, COL_MM_FOURNISS)
    map_four_ci  = build_lookup(recap_mmall, COL_MM_CI,  COL_MM_FOURNISS)
    df["Fournisseur"] = df.apply(
        lambda r: lookup_with_fallback(r, COL_GTIN, map_four_ean, COL_REF, map_four_ci),
        axis=1,
    )

    # --- Step 18 : MIN % (Marge min) ----------------------------------------
    log("Etape 18 : MIN %...")
    rayon_col = next(
        (c for c in marge_min.columns if str(c).strip().upper() == "RAYON"), None
    )
    minpct_col = next(
        (c for c in marge_min.columns if "MIN" in str(c).upper() and "%" in str(c)), None
    )
    if minpct_col is None:
        minpct_col = next(
            (c for c in marge_min.columns if str(c).strip().upper() == "MIN %"), None
        )
    if rayon_col is None or minpct_col is None:
        log("Attention : colonnes RAYON ou MIN % introuvables dans Marge min — MIN % sera NaN.")

    map_minpct = {}
    if rayon_col is not None and minpct_col is not None:
        for ray, pct in zip(marge_min[rayon_col], marge_min[minpct_col]):
            if ray is None:
                continue
            try:
                if pd.isna(ray):
                    continue
            except (TypeError, ValueError):
                pass
            key = str(ray).strip().upper()
            if key not in map_minpct and not (isinstance(pct, float) and np.isnan(pct)):
                map_minpct[key] = pct

    def min_pct_lookup(ray):
        if ray is None:
            return np.nan
        try:
            if pd.isna(ray):
                return np.nan
        except (TypeError, ValueError):
            pass
        return map_minpct.get(str(ray).strip().upper(), np.nan)

    df["MIN %"] = df["RAYON"].apply(min_pct_lookup)
    df["MIN %"] = pd.to_numeric(df["MIN %"], errors="coerce")

    # --- Step 19 : Ecart Vs BO MM -------------------------------------------
    log("Etape 19 : Ecart Vs BO MM...")
    prix_num = pd.to_numeric(df[COL_PRIX], errors="coerce")

    def ecart(row, prix_val):
        m = row["Min Prix Marjane"]
        if is_number(m) and is_number(prix_val):
            return m - prix_val
        return 0

    df["Ecart Vs BO MM"] = [
        ecart(r, p) for (_, r), p in zip(df.iterrows(), prix_num)
    ]

    # --- Step 20 : action ---------------------------------------------------
    log("Etape 20 : calcul de l'action...")

    def compute_action(row):
        ecart_v = row["Ecart Vs BO MM"]
        if ecart_v is None:
            return ""
        try:
            if pd.isna(ecart_v):
                return ""
        except (TypeError, ValueError):
            pass
        if ecart_v == 0:
            return ""
        if ecart_v > 0:
            return ACTION_UP
        # ecart < 0: check marges in priority order, stop at first valid number
        min_pct = row["MIN %"]
        for marge in [row["MARGE Moyen"], row["MARGE Réception"], row["MARGE RECAP"]]:
            if is_number(marge):
                if not is_number(min_pct):
                    return ""
                if marge >= min_pct:
                    return ACTION_DOWN
                elif (min_pct - marge) < MARGE_TOLERANCE:
                    return ACTION_DOWN
                else:
                    return ""
        return ""

    df["action"] = df.apply(compute_action, axis=1)

    # --- Step 21 : saise ----------------------------------------------------
    log("Etape 21 : saise...")
    map_sais_ean = build_lookup(recap_mmall, COL_MM_EAN, COL_MM_SAIS)
    map_sais_ci  = build_lookup(recap_mmall, COL_MM_CI,  COL_MM_SAIS)

    # When falling back on CODE_INTERNE, only accept the match if the row's EAN
    # in Recap mmall equals our item's GTIN — prevents inheriting attributes
    # from a different product that happens to share the same internal code.
    map_ci_to_ean_mmall = build_lookup(recap_mmall, COL_MM_CI, COL_MM_EAN)

    def lookup_mmall_validated(row, primary_map, fallback_map):
        g   = norm_key(row.get(COL_GTIN))
        if g is not None and g in primary_map:
            return primary_map[g]
        ref = norm_key(row.get(COL_REF))
        if ref is not None and ref in fallback_map:
            if norm_key(map_ci_to_ean_mmall.get(ref)) == g:
                return fallback_map[ref]
        return np.nan

    df["saise"] = df.apply(
        lambda r: lookup_mmall_validated(r, map_sais_ean, map_sais_ci), axis=1
    )

    # --- Step 22 : etat -----------------------------------------------------
    log("Etape 22 : etat...")
    map_etat_ean = build_lookup(recap_mmall, COL_MM_EAN, COL_MM_ETAT)
    map_etat_ci  = build_lookup(recap_mmall, COL_MM_CI,  COL_MM_ETAT)
    df["etat"] = df.apply(
        lambda r: lookup_mmall_validated(r, map_etat_ean, map_etat_ci), axis=1
    )

    # --- Build catalogue output ---------------------------------------------
    output_cols = [
        COL_MODIFIER, COL_GTIN, COL_REF, COL_SKU, COL_TITLE,
        COL_ATTR_CLE, COL_ATTR_VAL, COL_STOCK, COL_PRIX, OUT_COL_PRIX_BARRE,
        "type", "l'oreal", "Liste ne pas toucher",
        "Prix Marjane 1", "Stock Marjane 1",
        "Prix Marjane 2", "Stock Marjane 2",
        "Prix Marjane 3", "Stock Marjane 3",
        "Min Prix Marjane",
        "PA RECAP", "PA Réception", "PA Moyen",
        "MARGE RECAP", "MARGE Réception", "MARGE Moyen",
        "RAYON", "Fournisseur", "MIN %",
        "Ecart Vs BO MM", "action", "saise", "etat",
    ]

    if OUT_COL_PRIX_BARRE not in df.columns:
        df[OUT_COL_PRIX_BARRE] = df[barre_col]

    # Source columns missing from Actif mmall will be empty (NaN) in output
    for c in output_cols:
        if c not in df.columns:
            log(f"Attention : colonne '{c}' absente — remplie avec NaN.")
            df[c] = np.nan

    df_cat = df[output_cols].copy()
    log(f"Feuille catalogue: {len(df_cat)} lignes.")

    # --- Step 23 : SKU baisse -----------------------------------------------
    log("Etape 23 : feuille SKU baisse...")
    baisse_src = df[df["action"] == ACTION_DOWN].copy()
    if len(baisse_src) > 0:
        df_baisse = pd.DataFrame({
            COL_MODIFIER:  MODIFIER_FLAG,
            COL_GTIN:      baisse_src[COL_GTIN],
            COL_REF:       baisse_src[COL_REF],
            COL_SKU:       baisse_src[COL_SKU],
            COL_TITLE:     baisse_src[COL_TITLE],
            COL_ATTR_CLE:  baisse_src[COL_ATTR_CLE],
            COL_ATTR_VAL:  baisse_src[COL_ATTR_VAL],
            COL_STOCK:     baisse_src[COL_STOCK],
            OUT_COL_NV_PRIX:    baisse_src["Min Prix Marjane"],
            OUT_COL_PRIX_BARRE: baisse_src["__orig_prix"],
        })
    else:
        df_baisse = pd.DataFrame(columns=OUT_COLS_BAISSE)
    log(f"Feuille SKU baisse: {len(df_baisse)} lignes.")

    # --- Step 24 : SKU augmentation -----------------------------------------
    log("Etape 24 : feuille SKU augmentation...")
    augm_src = df[
        (df["action"] == ACTION_UP)
        & (df["saise"] == SAIS_PERMANENT)
        & (df["etat"] == ETAT_ACTIF)
    ].copy()

    def nv_prix_barre(row):
        mp = row["Min Prix Marjane"]
        ob = row["__orig_prix_barre"]
        op = row["__orig_prix"]
        if is_number(mp) and is_number(ob) and is_number(op) and op != 0:
            return mp * ob / op
        return np.nan

    if len(augm_src) > 0:
        df_augment = pd.DataFrame({
            COL_MODIFIER:        MODIFIER_FLAG,
            COL_GTIN:            augm_src[COL_GTIN],
            COL_REF:             augm_src[COL_REF],
            COL_SKU:             augm_src[COL_SKU],
            COL_TITLE:           augm_src[COL_TITLE],
            COL_ATTR_CLE:        augm_src[COL_ATTR_CLE],
            COL_ATTR_VAL:        augm_src[COL_ATTR_VAL],
            COL_STOCK:           augm_src[COL_STOCK],
            OUT_COL_NV_PRIX:          augm_src["Min Prix Marjane"],
            OUT_COL_NV_PRIX_BARRE:    augm_src.apply(nv_prix_barre, axis=1),
        })
    else:
        df_augment = pd.DataFrame(columns=OUT_COLS_AUG)
    log(f"Feuille SKU augmentation: {len(df_augment)} lignes.")

    return df_cat, df_baisse, df_augment, msgs


# ---------------------------------------------------------------------------
# Gradio glue
# ---------------------------------------------------------------------------

def run_workflow(
    f_mmall, f_marj1, f_marj2, f_marj3,
    f_base_retail, f_liquidation, f_exclus,
    f_marge_min, f_pa_reception, f_recap_marjane, f_recap_mmall,
):
    all_files = [
        f_mmall, f_marj1, f_marj2, f_marj3,
        f_base_retail, f_liquidation, f_exclus,
        f_marge_min, f_pa_reception, f_recap_marjane, f_recap_mmall,
    ]
    labels = [
        "Actif mmall", "Actif marjane 1", "Actif marjane 2", "Actif marjane 3",
        "Base retail", "Liste liquidation", "Liste exclus",
        "Marge min", "PA reception", "Recap marjane", "Recap mmall",
    ]

    missing = [labels[i] for i, f in enumerate(all_files) if f is None]
    if missing:
        return None, "Erreur : fichiers manquants -> " + ", ".join(missing)

    try:
        df_cat, df_baisse, df_augment, msgs = process(*all_files)

        tmp_dir  = tempfile.mkdtemp(prefix="ajustement_")
        out_path = os.path.join(tmp_dir, "Ajustement_prix_resultat.xlsx")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df_cat.to_excel(writer,     sheet_name=OUT_SHEET_CAT,    index=False)
            df_baisse.to_excel(writer,  sheet_name=OUT_SHEET_BAISSE, index=False)
            df_augment.to_excel(writer, sheet_name=OUT_SHEET_AUG,    index=False)

        summary = (
            "Traitement termine avec succes !\n\n"
            + "\n".join(msgs)
            + "\n\n--- Resume ---\n"
            + f"Catalogue          : {len(df_cat)} lignes\n"
            + f"SKU a baisser       : {len(df_baisse)} lignes\n"
            + f"SKU a augmenter     : {len(df_augment)} lignes\n"
        )
        return out_path, summary

    except Exception as e:
        tb = traceback.format_exc()
        return None, f"Erreur lors du traitement : {e}\n\n{tb}"


def build_ui():
    with gr.Blocks(title="Ajustement des Prix - Marjane Mall") as demo:
        gr.Markdown("# Ajustement des Prix - Marjane Mall")
        gr.Markdown(
            "Cet outil automatise l'alignement des prix. "
            f"Chargez les **{N_INPUT_FILES} fichiers Excel** repartis en 3 sections, "
            "puis cliquez sur **Generer l'ajustement**. "
            "Le resultat est un fichier Excel a 3 feuilles : "
            f"*{OUT_SHEET_CAT}*, *{OUT_SHEET_BAISSE}*, *{OUT_SHEET_AUG}*."
        )

        with gr.Row():
            with gr.Column():
                gr.Markdown(
                    "### 1. Fichiers Actif\n"
                    "Le catalogue mmall et les 3 catalogues concurrents Marjane "
                    f"(feuille `{SHEET_ACTIF}`)."
                )
                f_mmall = gr.File(label="Actif mmall")
                f_marj1 = gr.File(label="Actif marjane 1")
                f_marj2 = gr.File(label="Actif marjane 2")
                f_marj3 = gr.File(label="Actif marjane 3")

            with gr.Column():
                gr.Markdown(
                    "### 2. Listes de reference\n"
                    "Master produit (Base retail), SKU a exclure (liquidation), "
                    "et indicateur EAN (exclus)."
                )
                f_base_retail = gr.File(label="Base retail")
                f_liquidation = gr.File(label="Liste sku - liquidation")
                f_exclus      = gr.File(label="Liste sku - exclus")

            with gr.Column():
                gr.Markdown(
                    "### 3. Donnees financieres\n"
                    "Marge minimale par rayon, prix d'achat (PA reception), "
                    "et les deux recaps (marjane / mmall)."
                )
                f_marge_min     = gr.File(label="Marge min")
                f_pa_reception  = gr.File(label="PA reception")
                f_recap_marjane = gr.File(label="Recap marjane")
                f_recap_mmall   = gr.File(label="Recap mmall")

        btn = gr.Button("Generer l'ajustement", variant="primary", size="lg")

        status = gr.Textbox(
            label="Statut / journal",
            lines=14,
            interactive=False,
        )
        out_file = gr.File(label="Fichier resultat (telechargement)")

        btn.click(
            fn=run_workflow,
            inputs=[
                f_mmall, f_marj1, f_marj2, f_marj3,
                f_base_retail, f_liquidation, f_exclus,
                f_marge_min, f_pa_reception, f_recap_marjane, f_recap_mmall,
            ],
            outputs=[out_file, status],
        )

    return demo


if __name__ == "__main__":
    app = build_ui()
    app.launch()
