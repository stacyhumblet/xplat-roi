// ============================================================
//  Campaign ROI Dashboard
//  Google Apps Script -- Code.gs
//  Northwind Pulse -- Creator & Ecommerce
// ============================================================

const DATA_SHEET_ID  = '1c86yKsvHafajJSLYiHDrFQ6Evgx9rxGJTalRSEI46l0';
const ROI_TAB        = 'db_campaign_roi';
const CACHE_KEY      = 'campaign_roi_v1';
const CACHE_META_KEY = '__roi_chunks__';
const CACHE_TTL      = 21600; // 6 hours


// ── Entry point ────────────────────────────────────────────────────────────────
function doGet() {
  try {
    return ContentService
      .createTextOutput(JSON.stringify(getCampaignRoiData()))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ── Main data function ─────────────────────────────────────────────────────────
function getCampaignRoiData() {
  const cache  = CacheService.getScriptCache();
  const cached = _getChunks(cache);
  if (cached) return JSON.parse(cached);

  const ss    = SpreadsheetApp.openById(DATA_SHEET_ID);
  const sheet = ss.getSheetByName(ROI_TAB);
  if (!sheet) throw new Error('Tab not found: ' + ROI_TAB);

  const vals    = sheet.getDataRange().getValues();
  const headers = vals[0].map(String);
  function col(n) { return headers.indexOf(n); }

  const iId           = col('campaign_id');
  const iDate         = col('campaign_date');
  const iPlatform     = col('platform');
  const iCampaignName = col('campaign_name');
  const iCampaignType = col('campaign_type');
  const iImpressions  = col('impressions');
  const iClicks       = col('clicks');
  const iSpend        = col('spend');
  const iConversions  = col('conversions');
  const iRevenue      = col('revenue');

  const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun',
                       'Jul','Aug','Sep','Oct','Nov','Dec'];

  const rows          = [];
  const years         = new Set();
  const months        = new Set();
  const platforms     = new Set();
  const campaignTypes = new Set();
  const campaigns     = new Set();

  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (!row[iId]) continue;

    const rawDate = row[iDate];
    const d       = rawDate ? new Date(rawDate) : null;
    const yr      = d ? String(d.getFullYear())                   : '';
    const mo      = d ? String(d.getMonth() + 1).padStart(2,'0') : '';
    const mKey    = yr && mo ? `${yr}-${mo}` : '';
    const mLabel  = d ? `${MONTH_NAMES[d.getMonth()]} ${yr}`     : '';

    const platform     = String(row[iPlatform]     || '').trim().toLowerCase();
    const campaignName = String(row[iCampaignName] || '').trim();
    const campaignType = String(row[iCampaignType] || '').trim().toLowerCase();
    const impressions  = parseFloat(row[iImpressions]) || 0;
    const clicks       = parseFloat(row[iClicks])      || 0;
    const spend        = parseFloat(row[iSpend])        || 0;
    const conversions  = parseFloat(row[iConversions]) || 0;
    const revenue      = parseFloat(row[iRevenue])     || 0;

    rows.push({
      yr, mo, mKey, mLabel,
      platform, campaignName, campaignType,
      impressions, clicks, spend, conversions, revenue
    });

    if (yr)           years.add(yr);
    if (mo)           months.add(mo);
    if (platform)     platforms.add(platform);
    if (campaignType) campaignTypes.add(campaignType);
    if (campaignName) campaigns.add(campaignName);
  }

  const MONTH_ORDER = ['01','02','03','04','05','06','07','08','09','10','11','12'];

  const result = {
    rows,
    fo: {
      years:          [...years].sort().reverse(),
      months:         MONTH_ORDER.filter(m => months.has(m))
                        .map(m => ({ value: m, label: MONTH_NAMES[parseInt(m) - 1] })),
      platforms:      [...platforms].sort(),
      campaign_types: [...campaignTypes].sort(),
      campaigns:      [...campaigns].sort(),
    },
  };

  _putChunks(cache, JSON.stringify(result));
  return result;
}


// ── Cache helpers ──────────────────────────────────────────────────────────────
function _putChunks(cache, json) {
  try {
    const CHUNK = 90000;
    const total = Math.ceil(json.length / CHUNK);
    const pairs = { [CACHE_META_KEY]: String(total) };
    for (let i = 0; i < total; i++) {
      pairs[CACHE_KEY + '_' + i] = json.slice(i * CHUNK, (i + 1) * CHUNK);
    }
    cache.putAll(pairs, CACHE_TTL);
  } catch (e) { console.log('Cache write failed:', e); }
}

function _getChunks(cache) {
  try {
    const meta = cache.get(CACHE_META_KEY);
    if (!meta) return null;
    const total  = parseInt(meta);
    const keys   = Array.from({ length: total }, (_, i) => CACHE_KEY + '_' + i);
    const stored = cache.getAll(keys);
    if (Object.keys(stored).length !== total) return null;
    return keys.map(k => stored[k]).join('');
  } catch (e) { return null; }
}


// ── Setup sheet with Maple & Co. demo data ─────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  let sheet = ss.getSheetByName(ROI_TAB);
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(ROI_TAB);
  }

  const headers = [
    'campaign_id','campaign_date','platform','campaign_name',
    'campaign_type','impressions','clicks','spend','conversions','revenue'
  ];
  sheet.appendRow(headers);

  // Maple & Co. demo data -- 12 months March 2025 through March 2026
  // Story: summer growth, Aug IG collab spike, Oct TikTok viral spike,
  //        Nov-Dec holiday boom, Jan dip, Mar 2026 recovery
  const rows = [
    // ── MARCH 2025 ──
    ['c001','2025-03-01','instagram','Instagram Organic',           'organic', 82000,  1640,    0,   29,  1740],
    ['c002','2025-03-01','instagram','Instagram Paid - Retargeting','paid',    18000,   540,  180,   14,   630],
    ['c003','2025-03-01','tiktok',   'TikTok Organic',              'organic', 41000,   615,    0,    8,   480],
    ['c004','2025-03-01','google',   'Google Organic Search',       'organic',  9200,   460,    0,    6,   360],
    ['c005','2025-03-01','google',   'Google Paid - Spring',        'paid',    14000,   700,  310,   18,   810],
    ['c006','2025-03-01','pinterest','Pinterest Organic',           'organic', 22000,   330,    0,    5,   300],
    ['c007','2025-03-01','email',    'Email Newsletter',            'organic',  6800,  2380,   30,   36,  2160],

    // ── APRIL 2025 ──
    ['c008','2025-04-01','instagram','Instagram Organic',           'organic', 88000,  1760,    0,   31,  1860],
    ['c009','2025-04-01','instagram','Instagram Paid - Retargeting','paid',    20000,   600,  195,   16,   720],
    ['c010','2025-04-01','tiktok',   'TikTok Organic',              'organic', 44000,   660,    0,    9,   540],
    ['c011','2025-04-01','google',   'Google Organic Search',       'organic',  9800,   490,    0,    7,   420],
    ['c012','2025-04-01','google',   'Google Paid - Spring',        'paid',    15500,   775,  330,   21,   945],
    ['c013','2025-04-01','pinterest','Pinterest Organic',           'organic', 24000,   360,    0,    6,   360],
    ['c014','2025-04-01','email',    'Email Newsletter',            'organic',  7100,  2485,   30,   38,  2280],

    // ── MAY 2025 ──
    ['c015','2025-05-01','instagram','Instagram Organic',           'organic', 95000,  1900,    0,   34,  2040],
    ['c016','2025-05-01','instagram','Instagram Paid - Retargeting','paid',    22000,   660,  210,   18,   810],
    ['c017','2025-05-01','tiktok',   'TikTok Organic',              'organic', 52000,   780,    0,   11,   660],
    ['c018','2025-05-01','google',   'Google Organic Search',       'organic', 10500,   525,    0,    8,   480],
    ['c019','2025-05-01','google',   'Google Paid - Spring',        'paid',    18000,   900,  380,   25,  1125],
    ['c020','2025-05-01','pinterest','Pinterest Organic',           'organic', 27000,   405,    0,    7,   420],
    ['c021','2025-05-01','email',    'Email Newsletter',            'organic',  7400,  2590,   30,   41,  2460],

    // ── JUNE 2025 ──
    ['c022','2025-06-01','instagram','Instagram Organic',           'organic',102000,  2040,    0,   36,  2160],
    ['c023','2025-06-01','tiktok',   'TikTok Organic',              'organic', 58000,   870,    0,   13,   780],
    ['c024','2025-06-01','google',   'Google Organic Search',       'organic', 11200,   560,    0,    9,   540],
    ['c025','2025-06-01','pinterest','Pinterest Organic',           'organic', 29000,   435,    0,    8,   480],
    ['c026','2025-06-01','email',    'Email Newsletter',            'organic',  7800,  2730,   30,   44,  2640],

    // ── JULY 2025 ──
    ['c027','2025-07-01','instagram','Instagram Organic',           'organic',110000,  2200,    0,   39,  2340],
    ['c028','2025-07-01','tiktok',   'TikTok Organic',              'organic', 63000,   945,    0,   14,   840],
    ['c029','2025-07-01','google',   'Google Organic Search',       'organic', 11800,   590,    0,    9,   540],
    ['c030','2025-07-01','pinterest','Pinterest Organic',           'organic', 31000,   465,    0,    9,   540],
    ['c031','2025-07-01','email',    'Email Newsletter',            'organic',  8100,  2835,   30,   46,  2760],

    // ── AUGUST 2025 -- IG collab spike ──
    ['c032','2025-08-01','instagram','Instagram Organic',           'organic',168000,  3360,    0,   72,  4320],
    ['c033','2025-08-01','instagram','Instagram Paid - Retargeting','paid',    38000,  1140,  340,   42,  1890],
    ['c034','2025-08-01','tiktok',   'TikTok Organic',              'organic', 71000,  1065,    0,   16,   960],
    ['c035','2025-08-01','google',   'Google Organic Search',       'organic', 13400,   670,    0,   12,   720],
    ['c036','2025-08-01','pinterest','Pinterest Organic',           'organic', 38000,   570,    0,   13,   780],
    ['c037','2025-08-01','email',    'Email Newsletter',            'organic',  9200,  3220,   30,   58,  3480],

    // ── SEPTEMBER 2025 ──
    ['c038','2025-09-01','instagram','Instagram Organic',           'organic',118000,  2360,    0,   42,  2520],
    ['c039','2025-09-01','instagram','Instagram Paid - Retargeting','paid',    25000,   750,  215,   19,   855],
    ['c040','2025-09-01','tiktok',   'TikTok Organic',              'organic', 67000,  1005,    0,   15,   900],
    ['c041','2025-09-01','google',   'Google Organic Search',       'organic', 12600,   630,    0,   10,   600],
    ['c042','2025-09-01','pinterest','Pinterest Organic',           'organic', 33000,   495,    0,   10,   600],
    ['c043','2025-09-01','email',    'Email Newsletter',            'organic',  8500,  2975,   30,   48,  2880],

    // ── OCTOBER 2025 -- TikTok viral spike + holiday ramp ──
    ['c044','2025-10-01','instagram','Instagram Organic',           'organic',130000,  2600,    0,   47,  2820],
    ['c045','2025-10-01','instagram','Instagram Paid - Retargeting','paid',    32000,   960,  290,   26,  1170],
    ['c046','2025-10-01','tiktok',   'TikTok Organic',              'organic',820000, 12300,    0,  148,  8880],
    ['c047','2025-10-01','google',   'Google Organic Search',       'organic', 19800,   990,    0,   18,  1080],
    ['c048','2025-10-01','google',   'Google Paid - Holiday',       'paid',    22000,  1100,  420,   31,  1395],
    ['c049','2025-10-01','pinterest','Pinterest Organic',           'organic', 44000,   660,    0,   14,   840],
    ['c050','2025-10-01','email',    'Email Newsletter',            'organic',  9800,  3430,   30,   62,  3720],

    // ── NOVEMBER 2025 -- holiday boom ──
    ['c051','2025-11-01','instagram','Instagram Organic',           'organic',155000,  3100,    0,   65,  3900],
    ['c052','2025-11-01','instagram','Instagram Paid - Retargeting','paid',    48000,  1440,  380,   54,  2430],
    ['c053','2025-11-01','tiktok',   'TikTok Organic',              'organic',210000,  3150,    0,   58,  3480],
    ['c054','2025-11-01','google',   'Google Organic Search',       'organic', 22000,  1100,    0,   22,  1320],
    ['c055','2025-11-01','google',   'Google Paid - Holiday',       'paid',    35000,  1750,  560,   51,  2295],
    ['c056','2025-11-01','pinterest','Pinterest Organic',           'organic', 58000,   870,    0,   21,  1260],
    ['c057','2025-11-01','email',    'Email Newsletter',            'organic', 12400,  4340,   30,   88,  5280],

    // ── DECEMBER 2025 -- peak holiday ──
    ['c058','2025-12-01','instagram','Instagram Organic',           'organic',178000,  3560,    0,   82,  4920],
    ['c059','2025-12-01','instagram','Instagram Paid - Retargeting','paid',    62000,  1860,  400,   74,  3330],
    ['c060','2025-12-01','tiktok',   'TikTok Organic',              'organic',185000,  2775,    0,   51,  3060],
    ['c061','2025-12-01','google',   'Google Organic Search',       'organic', 26000,  1300,    0,   27,  1620],
    ['c062','2025-12-01','google',   'Google Paid - Holiday',       'paid',    42000,  2100,  600,   65,  2925],
    ['c063','2025-12-01','pinterest','Pinterest Organic',           'organic', 72000,  1080,    0,   28,  1680],
    ['c064','2025-12-01','email',    'Email Newsletter',            'organic', 15800,  5530,   30,  118,  7080],

    // ── JANUARY 2026 -- post-holiday dip ──
    ['c065','2026-01-01','instagram','Instagram Organic',           'organic', 92000,  1840,    0,   28,  1680],
    ['c066','2026-01-01','tiktok',   'TikTok Organic',              'organic', 74000,  1110,    0,   16,   960],
    ['c067','2026-01-01','google',   'Google Organic Search',       'organic', 11000,   550,    0,    8,   480],
    ['c068','2026-01-01','pinterest','Pinterest Organic',           'organic', 28000,   420,    0,    6,   360],
    ['c069','2026-01-01','email',    'Email Newsletter',            'organic',  8800,  3080,   30,   38,  2280],

    // ── FEBRUARY 2026 ──
    ['c070','2026-02-01','instagram','Instagram Organic',           'organic', 98000,  1960,    0,   32,  1920],
    ['c071','2026-02-01','instagram','Instagram Paid - Retargeting','paid',    19000,   570,  185,   15,   675],
    ['c072','2026-02-01','tiktok',   'TikTok Organic',              'organic', 82000,  1230,    0,   18,  1080],
    ['c073','2026-02-01','google',   'Google Organic Search',       'organic', 11800,   590,    0,    9,   540],
    ['c074','2026-02-01','pinterest','Pinterest Organic',           'organic', 30000,   450,    0,    7,   420],
    ['c075','2026-02-01','email',    'Email Newsletter',            'organic',  9100,  3185,   30,   41,  2460],

    // ── MARCH 2026 -- recovery ──
    ['c076','2026-03-01','instagram','Instagram Organic',           'organic',108000,  2160,    0,   38,  2280],
    ['c077','2026-03-01','instagram','Instagram Paid - Retargeting','paid',    24000,   720,  210,   20,   900],
    ['c078','2026-03-01','tiktok',   'TikTok Organic',              'organic', 90000,  1350,    0,   21,  1260],
    ['c079','2026-03-01','google',   'Google Organic Search',       'organic', 12400,   620,    0,   10,   600],
    ['c080','2026-03-01','google',   'Google Paid - Spring',        'paid',    16000,   800,  320,   22,   990],
    ['c081','2026-03-01','pinterest','Pinterest Organic',           'organic', 32000,   480,    0,    8,   480],
    ['c082','2026-03-01','email',    'Email Newsletter',            'organic',  9600,  3360,   30,   45,  2700],
  ];

  rows.forEach(r => sheet.appendRow(r));
  Logger.log('setupSheet complete -- ' + rows.length + ' rows written to ' + ROI_TAB);
}


// ── Utilities ──────────────────────────────────────────────────────────────────
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_META_KEY);
  Logger.log('Cache cleared.');
}

function warmCache() {
  clearCache();
  getCampaignRoiData();
  Logger.log('Cache warmed at ' + new Date().toLocaleString());
}

function createWarmCacheTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'warmCache')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warmCache').timeBased().everyHours(5).create();
  Logger.log('Trigger created -- fires every 5 hrs, cache TTL 6 hrs.');
}

function testDataAccess() {
  clearCache();
  const data = getCampaignRoiData();
  Logger.log('Total rows: '     + data.rows.length);
  Logger.log('Platforms: '      + JSON.stringify(data.fo.platforms));
  Logger.log('Campaign types: ' + JSON.stringify(data.fo.campaign_types));
  Logger.log('Campaigns: '      + JSON.stringify(data.fo.campaigns));
  Logger.log('Filter options: ' + JSON.stringify(data.fo));
}
