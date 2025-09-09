/* Machines Grid Web App (clean single-file version) */

const CONFIG = {
  SPREADSHEET_ID: '1a_QsBtruUBjCAIhm6aYafqqtR9UqH56GGHUDhQ7WRLM',
  SHEET_NAME: 'Machines',
  CACHE_SECONDS: 300,
  BRAND: {
    pageGrey: '#747678',
    ok:       '#22C55E',
    warn:     '#F59E0B',
    danger:   '#EF4444',
    access:   '#00C6D7',
    text:     '#0f172a',
    card:     '#800080',
    cardEdge: '#e5e7eb'
  }
};

// optional HTML cache (kept small due to CacheService limit ~100 KB)
const HTML_CACHE_SECONDS = 300;

function getVersion_() {
  return PropertiesService.getScriptProperties().getProperty('BUMP') || '0';
}

function warmCache(){
  const t = HtmlService.createTemplateFromFile('Index');
  t.initialData = getMachines_(); t.brand = CONFIG.BRAND;
  const html = t.evaluate().getContent(); // renders + JIT warms
}

function createWarmTrigger(){ ScriptApp.newTrigger('warmCache').timeBased().everyMinutes(5).create(); }

function doGet(e) {
  const params = (e && e.parameter) || {};
  const ver = getVersion_();
  const cache = CacheService.getScriptCache();

  // JSON API
  if (params.format === 'json') {
    const dataKey = 'data_v1_' + ver;
    const hit = cache.get(dataKey);
    if (hit && !params.nocache) {
      return ContentService.createTextOutput(hit)
        .setMimeType(ContentService.MimeType.JSON);
    }
    const payload = JSON.stringify({ updated: new Date().toISOString(), machines: getMachines_() });
    cache.put(dataKey, payload, CONFIG.CACHE_SECONDS);
    return ContentService.createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // HTML (try cached)
  const htmlKey = 'html_v1_' + ver;
  const htmlHit = cache.get(htmlKey);
  if (htmlHit && !params.nocache) {
    return HtmlService.createHtmlOutput(htmlHit)
      .setTitle('Machines Grid')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Render fresh
  const t = HtmlService.createTemplateFromFile('Index');
  t.initialData = getMachines_();
  t.brand = CONFIG.BRAND;

  const out = t.evaluate()
    .setTitle('Machines Grid')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  const html = out.getContent();
  if (html.length < 95000) cache.put(htmlKey, html, HTML_CACHE_SECONDS);
  return out;
}

// install once: creates onEdit trigger to bust caches immediately on changes
function installEditTrigger() {
  ScriptApp.newTrigger('handleSheetEdit')
    .forSpreadsheet(CONFIG.SPREADSHEET_ID)
    .onEdit()
    .create();
}

function handleSheetEdit(e) {
  if (e && e.range && e.range.getSheet().getName() !== CONFIG.SHEET_NAME) return;
  PropertiesService.getScriptProperties().setProperty('BUMP', String(Date.now()));
}

// data loader (no internal cache; doGet caches the API/HTML by version)
function getMachines_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" not found.');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => normalizeKey_(h));
  const rows = values.slice(1).filter(r => r.some(c => c !== '' && c !== null));

  const items = rows.map(r => {
    const obj = {};
    headers.forEach((key, i) => { obj[key] = r[i]; });

    obj.name          = String(obj.name || '').trim();
    obj.category      = String(obj.category || '').trim();
    obj.location      = String(obj.location || '').trim();
    obj.status        = String(obj.status || '').trim();
    obj.hazard_class  = String(obj.hazard_class || obj.hazardclass || '').toString().trim();
    obj.access        = String(obj.access || '').trim();
    obj.access_chips  = splitToList_(obj.access_chips || obj.accesschips);
    obj.tags          = splitToList_(obj.tags);
    obj.thumbnail_url = String(obj.thumbnail_url || obj.thumbnailurl || '').trim();
    obj.detail_url    = String(obj.detail_url || obj.detailurl || '').trim();
    obj.description   = String(obj.description || '').trim();
    obj.specs         = String(obj.specs || '').trim();

    obj.status_color = pickStatusColor_(obj.status);
    obj.hazard_color = pickHazardColor_(obj.hazard_class);
    return obj;
  });

  return items;
}

function pickStatusColor_(status) {
  const s = String(status || '').toLowerCase();
  if (s.includes('down') || s.includes('offline') || s.includes('out of service')) return CONFIG.BRAND.danger;
  if (s.includes('restricted') || s.includes('limited')) return CONFIG.BRAND.warn;
  return CONFIG.BRAND.ok;
}

function pickHazardColor_(hazard) {
  const h = String(hazard || '').trim();
  if (h === '3') return CONFIG.BRAND.danger;
  if (h === '2') return CONFIG.BRAND.warn;
  if (h === '1') return CONFIG.BRAND.ok;
  return CONFIG.BRAND.cardEdge;
}

function splitToList_(val) {
  if (!val) return [];
  return String(val).split(/[;,]/).map(s => s.trim()).filter(Boolean);
}

function normalizeKey_(k) {
  return String(k).trim().toLowerCase().replace(/\s+/g, '_');
}

// for HtmlService includes (if you add more HTML files)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}