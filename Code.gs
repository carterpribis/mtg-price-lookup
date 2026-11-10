/**
 * ─── CONFIG ──────────────────────────────────────────────────────────────
 * A  Card            (exact or close)
 * B  Set             (optional set code, e.g. LEB)
 * C  Variant         (blank / nonfoil / foil / borderless / foil borderless)
 * D  Price (USD)     ← script writes here
 * E  Updated         ← script writes here
 */
const SHEET_NAME      = 'Sheet1';
const CACHE_TTL_SEC   = 60 * 60 * 24;   // 24 h
const THROTTLE_MS     = 120;            // keep <10 req/s to Scryfall

/* ── 1. Trigger handler: run whenever A, B, or C is edited ────────────── */
function onEditMTG(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;

  const startCol = e.range.getColumn();
  const endCol   = startCol + e.range.getNumColumns() - 1;
  if (endCol < 1 || startCol > 3) return;          // ignore edits outside A‑C

  const rows = e.range.getNumRows();
  for (let i = 0; i < rows; i++) {
    const r = e.range.getRow() + i;
    if (r === 1) continue;                         // skip header
    updateRow_(sh, r);
  }
}

/* ── 2. Update a single row ───────────────────────────────────────────── */
function updateRow_(sh, r) {
  const name = sh.getRange(r, 1).getValue().toString().trim();
  const set  = sh.getRange(r, 2).getValue().toString().trim();
  const { foil, borderless } = parseVariant_(sh.getRange(r, 3).getValue());

  if (!name) {
    sh.getRange(r, 4, 1, 2).clearContent();
    return;
  }

  const key   = `${name}|${set}|${foil}|${borderless}`;
  const cache = CacheService.getDocumentCache();
  //let price   = cache.get(key);

  //if (!price) {
    price = fetchPriceDetailed_(name, set, foil, borderless);
    cache.put(key, price, CACHE_TTL_SEC);
  //}
  sh.getRange(r, 4, 1, 2).setValues([[price, new Date()]]);
}

/* ── 3. Parse the VARIANT cell into flags ─────────────────────────────── */
function parseVariant_(raw) {
  const txt = raw.toString().toLowerCase();
  return {
    foil:       /\bfoil\b/.test(txt),
    borderless: /\bborderless\b/.test(txt)
  };
}

/* ── 4. Fetch the correct printing from Scryfall (FOIL ⨯ BORDERLESS) ─── */
function fetchPriceDetailed_(name, set, foil, borderless) {

  // Build URL
  let url;

  
  if (borderless) {
    // borderless search; unique:prints strips duplicates
    let q = `!"${name}" is:borderless unique:prints`;
    if (set) q += ` set:${set}`;
    url = `https://api.scryfall.com/cards/search?q=${encodeURIComponent(q)}`;
  } else {
    url = `https://api.scryfall.com/cards/named?exact=${encodeURIComponent(name)}${set?`&set=${set}`:''}`;
  }
  
  Logger.log('DEBUG url=' + url);
  Logger.log(
  `DEBUG name=${name}  set=${set || '(any)'}  foil=${foil}  borderless=${borderless}`
  );
  // Fetch
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return 'Not found';

  // Pick correct printing
  let card;
  if (borderless) {
    const list = JSON.parse(res.getContentText()).data || [];
    if (!list.length) return 'Not found';
    card = list.find(c => foil ? c.finishes.includes('foil')
                               : c.finishes.includes('nonfoil')) || list[0];
  } else {
    card = JSON.parse(res.getContentText());
  }

  // Return correct price column
  const price = foil ? card?.prices?.usd_foil : card?.prices?.usd;
  return price ?? 'N/A';

}

/* ── 6. UI menu + self‑installing trigger ─────────────────────────────── */
function onOpen() {
  ensureTrigger_();
  SpreadsheetApp.getUi()
    .createMenu('MTG Prices')
    .addItem('Refresh All Prices', 'refreshAllPrices')
    .addToUi();
}

function ensureTrigger_() {
  const fn = 'onEditMTG';
  const exists = ScriptApp.getProjectTriggers()
                  .some(t => t.getHandlerFunction() === fn);
  if (!exists) {
    ScriptApp.newTrigger(fn)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

function clearPriceCache() {
  const sh     = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const last   = sh.getLastRow();
  const cache  = CacheService.getDocumentCache();

  for (let r = 2; r <= last; r++) {
    const name = sh.getRange(r, 1).getValue().toString().trim();
    if (!name) continue;

    const set             = sh.getRange(r, 2).getValue().toString().trim();
    const { foil, borderless } = parseVariant_(sh.getRange(r, 3).getValue());
    const key = `${name}|${set}|${foil}|${borderless}`;
    cache.remove(key);
  }
  SpreadsheetApp.getUi().alert('Price cache cleared for rows on this sheet.');
}