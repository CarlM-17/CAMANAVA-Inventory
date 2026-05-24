'use strict';

const express = require('express');
const { google } = require('googleapis');
const { parse } = require('csv-parse');
const XLSX = require('xlsx');
const cron = require('node-cron');
const { Readable } = require('stream');
const crypto = require('crypto');

const app = express();
app.use(express.json());

// ─── ENV CONFIG ───────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
const GOOGLE_CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;
const GOOGLE_PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
const GDRIVE_FOLDER_ID = process.env.GDRIVE_FOLDER_ID;
const INV_FILE_NAME = process.env.INV_FILE_NAME || 'InvData.csv';
const STORES_FILE_NAME = process.env.STORES_FILE_NAME || 'ListOfStores.xlsx';
const REFRESH_INTERVAL_MINUTES = parseInt(process.env.REFRESH_INTERVAL_MINUTES || '10');

// ─── IN-MEMORY CACHE ──────────────────────────────────────────────────────────
let cache = {
  ready: false,
  lastRefresh: null,
  lastFileHash: null,
  lastFileSize: null,
  lastModifiedTime: null,
  rows: [],
  storeMap: {},       // storeId -> { area, storeName, region }
  kpis: {},
  criticalItems: [],
  overstockItems: [],
  deadStockItems: [],
  storeAnalysis: [],
  supplierAnalysis: [],
  filterMeta: {},
  refreshing: false,
  error: null
};

// ─── GOOGLE DRIVE AUTH ────────────────────────────────────────────────────────
function getDriveClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: GOOGLE_CLIENT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY
    },
    scopes: ['https://www.googleapis.com/auth/drive.readonly']
  });
  return google.drive({ version: 'v3', auth });
}

// ─── FIND FILE IN FOLDER ──────────────────────────────────────────────────────
async function findFile(drive, fileName) {
  const res = await drive.files.list({
    q: `'${GDRIVE_FOLDER_ID}' in parents and name='${fileName}' and trashed=false`,
    fields: 'files(id,name,size,modifiedTime,md5Checksum)',
    pageSize: 5
  });
  const files = res.data.files;
  if (!files || files.length === 0) return null;
  return files[0];
}

// ─── DOWNLOAD FILE AS BUFFER ──────────────────────────────────────────────────
async function downloadFileBuffer(drive, fileId) {
  const res = await drive.files.get(
    { fileId, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  return Buffer.from(res.data);
}

// ─── PARSE CSV FROM BUFFER ────────────────────────────────────────────────────
function parseCSV(buffer) {
  return new Promise((resolve, reject) => {
    const rows = [];
    const stream = Readable.from(buffer.toString('utf8'));
    stream
      .pipe(parse({ columns: false, skip_empty_lines: true, trim: true, from_line: 1 }))
      .on('data', (row) => rows.push(row))
      .on('end', () => resolve(rows))
      .on('error', reject);
  });
}

// ─── PARSE STORES XLSX FROM BUFFER ───────────────────────────────────────────
function parseStoresXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const storeMap = {};
  // Skip header row (row 0), data starts row 1
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const region = (row[0] || '').toString().trim();
    const area = (row[1] || '').toString().trim();
    const storeId = (row[2] || '').toString().trim();
    const storeName = (row[3] || '').toString().trim();
    const remarks = (row[4] || '').toString().trim();
    if (storeId) {
      storeMap[storeId] = { region, area, storeName, remarks };
    }
  }
  return storeMap;
}

// ─── SAFE NUMBER ─────────────────────────────────────────────────────────────
function num(val) {
  const n = parseFloat((val || '').toString().replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

// ─── BUILD ANALYTICS FROM ROWS ────────────────────────────────────────────────
function buildAnalytics(rawRows, storeMap) {
  // rawRows[0] = header
  if (rawRows.length < 2) return null;

  const header = rawRows[0];
  const dataRows = rawRows.slice(1);

  // Column index map (0-based, A=0)
  const COL = {
    regionCode: 0,   // A
    regionName: 1,   // B
    storeNumber: 2,  // C
    storeName: 3,    // D
    dept: 4,         // E
    deptName: 5,     // F
    subDept: 6,      // G
    subDeptName: 7,  // H
    cls: 8,          // I
    clsName: 9,      // J
    subCls: 10,      // K
    subClsName: 11,  // L
    brand: 12,       // M
    skuCode: 13,     // N
    skuDesc: 14,     // O
    skuStatus: 15,   // P
    onHand: 16,      // Q
    poOrderNet: 17,  // R
    trfOrderNet: 18, // S
    xdockNet: 19,    // T
    totalPONet: 20,  // U
    poOrderGR: 21,   // V
    trfOrderGR: 22,  // W
    xdockGR: 23,     // X
    wtsGross: 24,    // Y
    wtsAfterDeliv: 25, // Z
    currentWkSales: 26, // AA
    wk1: 27, wk2: 28, wk3: 29, wk4: 30,
    wk5: 31, wk6: 32, wk7: 33, wk8: 34,
    p8aveGross: 35,  // AJ (index 35)
    wtsNet: 36,      // AK
    wtsAftDelive: 37,// AL
    wk1net: 38, wk2net: 39, wk3net: 40, wk4net: 41,
    wk5net: 42, wk6net: 43, wk7net: 44, wk8net: 45,
    wkAveNet: 46,    // AU
    supplierCode: 47,// AV
    supplierName: 48,// AW
    avgCost: 49,     // AX
    buyUM: 50,       // AY
    stdPack: 51,     // AZ
    skuType: 52,     // BA
    merchGro: 53,    // BB
    onHandValue: 54, // BC
    poValue: 55,     // BD
    trfValue: 56,    // BE
    xdockValue: 57,  // BF
    total8wksGross: 58, // BG
    total8wksNet: 59,   // BH
    skuTyp: 60,      // BI
    ico: 61,         // BJ
    poType: 62,      // BK
    delivMode: 63,   // BL
    stsBatch: 64,    // BM
    stsNumber: 65,   // BN
    replenishment: 74 // BW
  };

  // Enrich rows and map store info
  const enriched = [];
  for (const row of dataRows) {
    if (!row || row.length < 10) continue;
    const storeId = (row[COL.storeNumber] || '').toString().trim();
    const storeInfo = storeMap[storeId] || {};
    const area = storeInfo.area || '';
    const wtsNet = num(row[COL.wtsNet]);
    const onHand = num(row[COL.onHand]);
    const onHandValue = num(row[COL.onHandValue]);
    const p8ave = num(row[COL.p8aveGross]);
    const currentWkSales = num(row[COL.currentWkSales]);
    const totalPO = num(row[COL.totalPONet]);
    const poValue = num(row[COL.poValue]);
    const trfValue = num(row[COL.trfValue]);

    const isCritical = wtsNet > 0 && wtsNet < 2 && onHand > 0;
    const isOverstock = wtsNet > 12 && onHand > 0;
    const isDeadStock = onHand > 0 && p8ave === 0 && currentWkSales === 0;
    const isZeroStock = onHand === 0;

    enriched.push({
      regionCode: row[COL.regionCode] || '',
      regionName: row[COL.regionName] || '',
      storeNumber: storeId,
      storeName: row[COL.storeName] || storeInfo.storeName || '',
      area,
      dept: row[COL.dept] || '',
      deptName: row[COL.deptName] || '',
      subDept: row[COL.subDept] || '',
      subDeptName: row[COL.subDeptName] || '',
      cls: row[COL.cls] || '',
      clsName: row[COL.clsName] || '',
      subCls: row[COL.subCls] || '',
      subClsName: row[COL.subClsName] || '',
      brand: row[COL.brand] || '',
      skuCode: row[COL.skuCode] || '',
      skuDesc: row[COL.skuDesc] || '',
      skuStatus: row[COL.skuStatus] || '',
      onHand,
      onHandValue,
      avgCost: num(row[COL.avgCost]),
      totalPO,
      poValue,
      trfValue,
      xdockValue: num(row[COL.xdockValue]),
      currentWkSales,
      p8ave,
      wtsNet: wtsNet === 0 && onHand > 0 && p8ave === 0 ? 999 : wtsNet,
      wtsGross: num(row[COL.wtsGross]),
      wtsAfterDeliv: num(row[COL.wtsAfterDeliv]),
      supplierCode: row[COL.supplierCode] || '',
      supplierName: row[COL.supplierName] || '',
      delivMode: row[COL.delivMode] || '',
      isCritical,
      isOverstock,
      isDeadStock,
      isZeroStock
    });
  }

  // ── KPIs ──────────────────────────────────────────────────────────────────
  const totalOnHandValue = enriched.reduce((s, r) => s + r.onHandValue, 0);
  const totalOnHand = enriched.reduce((s, r) => s + r.onHand, 0);
  const criticalCount = enriched.filter(r => r.isCritical).length;
  const overstockCount = enriched.filter(r => r.isOverstock).length;
  const deadStockCount = enriched.filter(r => r.isDeadStock).length;
  const activeStores = new Set(enriched.map(r => r.storeNumber)).size;
  const activeSuppliers = new Set(enriched.map(r => r.supplierCode).filter(Boolean)).size;
  const totalPOValue = enriched.reduce((s, r) => s + r.poValue, 0);
  const totalTRFValue = enriched.reduce((s, r) => s + r.trfValue, 0);
  const validWts = enriched.filter(r => r.wtsNet > 0 && r.wtsNet < 999 && r.onHand > 0);
  const avgWts = validWts.length > 0 ? validWts.reduce((s, r) => s + r.wtsNet, 0) / validWts.length : 0;

  const kpis = {
    totalOnHandValue,
    totalOnHand,
    criticalCount,
    overstockCount,
    deadStockCount,
    activeStores,
    activeSuppliers,
    totalPOValue,
    totalTRFValue,
    avgWts,
    totalSKUs: enriched.length
  };

  // ── CRITICAL ITEMS ────────────────────────────────────────────────────────
  const criticalItems = enriched
    .filter(r => r.isCritical)
    .sort((a, b) => a.wtsNet - b.wtsNet)
    .slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      onHandValue: r.onHandValue,
      currentWkSales: r.currentWkSales,
      p8ave: r.p8ave,
      wtsNet: r.wtsNet,
      totalPO: r.totalPO,
      action: r.totalPO > 0 ? 'PO Incoming' : r.p8ave > 0 ? 'URGENT: Place PO' : 'Review'
    }));

  // ── OVERSTOCK ITEMS ───────────────────────────────────────────────────────
  const overstockItems = enriched
    .filter(r => r.isOverstock)
    .sort((a, b) => b.wtsNet - a.wtsNet)
    .slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      onHandValue: r.onHandValue,
      p8ave: r.p8ave,
      wtsNet: r.wtsNet === 999 ? 'Dead Stock' : r.wtsNet.toFixed(1),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
    }));

  // ── DEAD STOCK ────────────────────────────────────────────────────────────
  const deadStockItems = enriched
    .filter(r => r.isDeadStock)
    .sort((a, b) => b.onHandValue - a.onHandValue)
    .slice(0, 300)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      onHandValue: r.onHandValue,
      action: 'No Sales 8 Wks - Review/Markdown'
    }));

  // ── STORE ANALYSIS ────────────────────────────────────────────────────────
  const storeGroups = {};
  for (const r of enriched) {
    const key = r.storeNumber;
    if (!storeGroups[key]) {
      storeGroups[key] = {
        storeNumber: r.storeNumber,
        storeName: r.storeName,
        area: r.area,
        region: r.regionName,
        totalValue: 0, totalOnHand: 0,
        criticalCount: 0, overstockCount: 0, deadCount: 0,
        totalSKUs: 0, totalSales: 0
      };
    }
    const g = storeGroups[key];
    g.totalValue += r.onHandValue;
    g.totalOnHand += r.onHand;
    g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
  }
  const storeAnalysis = Object.values(storeGroups).sort((a, b) => b.totalValue - a.totalValue);

  // ── SUPPLIER ANALYSIS ─────────────────────────────────────────────────────
  const supplierGroups = {};
  for (const r of enriched) {
    if (!r.supplierCode) continue;
    const key = r.supplierCode;
    if (!supplierGroups[key]) {
      supplierGroups[key] = {
        supplierCode: r.supplierCode,
        supplierName: r.supplierName,
        totalValue: 0, totalOnHand: 0,
        criticalCount: 0, overstockCount: 0, totalSKUs: 0
      };
    }
    const g = supplierGroups[key];
    g.totalValue += r.onHandValue;
    g.totalOnHand += r.onHand;
    g.totalSKUs++;
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
  }
  const supplierAnalysis = Object.values(supplierGroups).sort((a, b) => b.totalValue - a.totalValue).slice(0, 100);

  // ── FILTER METADATA ───────────────────────────────────────────────────────
  const uniq = (arr) => [...new Set(arr.filter(Boolean))].sort();
  const filterMeta = {
    regions: uniq(enriched.map(r => r.regionName)),
    areas: uniq(enriched.map(r => r.area)),
    stores: [...new Set(enriched.map(r => r.storeNumber))].sort((a, b) => a - b)
      .map(id => {
        const r = enriched.find(x => x.storeNumber === id);
        return { id, name: r ? r.storeName : id, area: r ? r.area : '' };
      }),
    depts: uniq(enriched.map(r => r.deptName)).filter(d => d.length > 0),
    subDepts: uniq(enriched.map(r => r.subDeptName)).filter(d => d.length > 0),
    classes: uniq(enriched.map(r => r.clsName)).filter(d => d.length > 0),
    suppliers: uniq(enriched.map(r => r.supplierName)).filter(d => d.length > 0),
    brands: uniq(enriched.map(r => r.brand)).filter(d => d.length > 0),
    skuStatuses: uniq(enriched.map(r => r.skuStatus))
  };

  return { kpis, criticalItems, overstockItems, deadStockItems, storeAnalysis, supplierAnalysis, filterMeta, rows: enriched };
}

// ─── MAIN REFRESH FUNCTION ────────────────────────────────────────────────────
async function refreshData(force = false) {
  if (cache.refreshing) {
    console.log('[Cache] Refresh already in progress, skipping.');
    return;
  }
  cache.refreshing = true;
  console.log(`[Cache] Starting refresh at ${new Date().toISOString()}`);

  try {
    if (!GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY || !GDRIVE_FOLDER_ID) {
      throw new Error('Missing Google Drive credentials in environment variables.');
    }

    const drive = getDriveClient();

    // Find InvData.csv
    const invFile = await findFile(drive, INV_FILE_NAME);
    if (!invFile) throw new Error(`${INV_FILE_NAME} not found in folder.`);

    const modifiedTime = invFile.modifiedTime;
    const fileSize = invFile.size;

    if (!force && cache.ready &&
        cache.lastModifiedTime === modifiedTime &&
        cache.lastFileSize === fileSize) {
      console.log('[Cache] File unchanged, skipping reprocess.');
      cache.refreshing = false;
      return;
    }

    console.log(`[Cache] Downloading ${INV_FILE_NAME} (${Math.round(fileSize / 1024 / 1024)}MB)...`);
    const invBuffer = await downloadFileBuffer(drive, invFile.id);

    const hash = crypto.createHash('md5').update(invBuffer).digest('hex');
    if (!force && cache.ready && cache.lastFileHash === hash) {
      console.log('[Cache] File hash unchanged, skipping reprocess.');
      cache.lastModifiedTime = modifiedTime;
      cache.refreshing = false;
      return;
    }

    // Find and parse ListOfStores.xlsx
    let storeMap = {};
    try {
      const storesFile = await findFile(drive, STORES_FILE_NAME);
      if (storesFile) {
        console.log('[Cache] Downloading ListOfStores.xlsx...');
        const storesBuffer = await downloadFileBuffer(drive, storesFile.id);
        storeMap = parseStoresXLSX(storesBuffer);
        console.log(`[Cache] Loaded ${Object.keys(storeMap).length} stores.`);
      }
    } catch (e) {
      console.warn('[Cache] Could not load stores file:', e.message);
    }

    console.log('[Cache] Parsing CSV...');
    const rawRows = await parseCSV(invBuffer);
    console.log(`[Cache] Parsed ${rawRows.length} rows.`);

    console.log('[Cache] Building analytics...');
    const analytics = buildAnalytics(rawRows, storeMap);
    if (!analytics) throw new Error('Analytics build failed - no data.');

    // Atomic swap
    cache.rows = analytics.rows;
    cache.storeMap = storeMap;
    cache.kpis = analytics.kpis;
    cache.criticalItems = analytics.criticalItems;
    cache.overstockItems = analytics.overstockItems;
    cache.deadStockItems = analytics.deadStockItems;
    cache.storeAnalysis = analytics.storeAnalysis;
    cache.supplierAnalysis = analytics.supplierAnalysis;
    cache.filterMeta = analytics.filterMeta;
    cache.lastFileHash = hash;
    cache.lastFileSize = fileSize;
    cache.lastModifiedTime = modifiedTime;
    cache.lastRefresh = new Date().toISOString();
    cache.ready = true;
    cache.error = null;

    console.log(`[Cache] Ready. ${analytics.rows.length} SKU rows loaded. Critical: ${analytics.kpis.criticalCount}, Overstock: ${analytics.kpis.overstockCount}`);
  } catch (err) {
    console.error('[Cache] Refresh error:', err.message);
    cache.error = err.message;
  } finally {
    cache.refreshing = false;
  }
}

// ─── BACKGROUND SCHEDULER ─────────────────────────────────────────────────────
cron.schedule(`*/${REFRESH_INTERVAL_MINUTES} * * * *`, () => {
  refreshData(false);
});

// ─── FILTER HELPER ────────────────────────────────────────────────────────────
function applyFilters(rows, filters = {}) {
  return rows.filter(r => {
    if (filters.region && r.regionName !== filters.region) return false;
    if (filters.area && r.area !== filters.area) return false;
    if (filters.store && r.storeNumber !== filters.store) return false;
    if (filters.dept && r.deptName !== filters.dept) return false;
    if (filters.subDept && r.subDeptName !== filters.subDept) return false;
    if (filters.cls && r.clsName !== filters.cls) return false;
    if (filters.supplier && r.supplierName !== filters.supplier) return false;
    if (filters.brand && r.brand !== filters.brand) return false;
    if (filters.skuStatus && r.skuStatus !== filters.skuStatus) return false;
    return true;
  });
}

// ─── API ROUTES ───────────────────────────────────────────────────────────────

app.get('/api/status', (req, res) => {
  res.json({
    ready: cache.ready,
    refreshing: cache.refreshing,
    lastRefresh: cache.lastRefresh,
    error: cache.error,
    totalRows: cache.rows.length
  });
});

app.get('/api/kpis', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = req.query;
  if (Object.keys(filters).length === 0) return res.json(cache.kpis);
  const filtered = applyFilters(cache.rows, filters);
  const totalOnHandValue = filtered.reduce((s, r) => s + r.onHandValue, 0);
  const totalOnHand = filtered.reduce((s, r) => s + r.onHand, 0);
  const criticalCount = filtered.filter(r => r.isCritical).length;
  const overstockCount = filtered.filter(r => r.isOverstock).length;
  const deadStockCount = filtered.filter(r => r.isDeadStock).length;
  const validWts = filtered.filter(r => r.wtsNet > 0 && r.wtsNet < 999);
  const avgWts = validWts.length > 0 ? validWts.reduce((s, r) => s + r.wtsNet, 0) / validWts.length : 0;
  res.json({
    totalOnHandValue, totalOnHand, criticalCount, overstockCount, deadStockCount,
    activeStores: new Set(filtered.map(r => r.storeNumber)).size,
    activeSuppliers: new Set(filtered.map(r => r.supplierCode).filter(Boolean)).size,
    totalPOValue: filtered.reduce((s, r) => s + r.poValue, 0),
    totalTRFValue: filtered.reduce((s, r) => s + r.trfValue, 0),
    avgWts,
    totalSKUs: filtered.length
  });
});

app.get('/api/filters', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  res.json(cache.filterMeta);
});

app.get('/api/critical', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = req.query;
  if (Object.keys(filters).length === 0) return res.json(cache.criticalItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isCritical)
    .sort((a, b) => a.wtsNet - b.wtsNet).slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue,
      currentWkSales: r.currentWkSales, p8ave: r.p8ave,
      wtsNet: r.wtsNet, totalPO: r.totalPO,
      action: r.totalPO > 0 ? 'PO Incoming' : r.p8ave > 0 ? 'URGENT: Place PO' : 'Review'
    }));
  res.json(filtered);
});

app.get('/api/overstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = req.query;
  if (Object.keys(filters).length === 0) return res.json(cache.overstockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isOverstock)
    .sort((a, b) => b.wtsNet - a.wtsNet).slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue, p8ave: r.p8ave,
      wtsNet: r.wtsNet === 999 ? 'Dead Stock' : r.wtsNet.toFixed(1),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
    }));
  res.json(filtered);
});

app.get('/api/deadstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = req.query;
  if (Object.keys(filters).length === 0) return res.json(cache.deadStockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isDeadStock)
    .sort((a, b) => b.onHandValue - a.onHandValue).slice(0, 300)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue,
      action: 'No Sales 8 Wks - Review/Markdown'
    }));
  res.json(filtered);
});

app.get('/api/stores', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = req.query;
  if (Object.keys(filters).length === 0) return res.json(cache.storeAnalysis);
  const filtered = applyFilters(cache.rows, filters);
  const storeGroups = {};
  for (const r of filtered) {
    const key = r.storeNumber;
    if (!storeGroups[key]) storeGroups[key] = { storeNumber: r.storeNumber, storeName: r.storeName, area: r.area, region: r.regionName, totalValue: 0, totalOnHand: 0, criticalCount: 0, overstockCount: 0, deadCount: 0, totalSKUs: 0, totalSales: 0 };
    const g = storeGroups[key];
    g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
  }
  res.json(Object.values(storeGroups).sort((a, b) => b.totalValue - a.totalValue));
});

app.get('/api/suppliers', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = req.query;
  if (Object.keys(filters).length === 0) return res.json(cache.supplierAnalysis);
  const filtered = applyFilters(cache.rows, filters);
  const supplierGroups = {};
  for (const r of filtered) {
    if (!r.supplierCode) continue;
    const key = r.supplierCode;
    if (!supplierGroups[key]) supplierGroups[key] = { supplierCode: r.supplierCode, supplierName: r.supplierName, totalValue: 0, totalOnHand: 0, criticalCount: 0, overstockCount: 0, totalSKUs: 0 };
    const g = supplierGroups[key];
    g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
  }
  res.json(Object.values(supplierGroups).sort((a, b) => b.totalValue - a.totalValue).slice(0, 100));
});

app.post('/api/refresh', async (req, res) => {
  refreshData(true);
  res.json({ message: 'Refresh triggered' });
});

// ─── EXPORT CSV ───────────────────────────────────────────────────────────────
app.get('/api/export/:type', (req, res) => {
  if (!cache.ready) return res.status(503).send('Cache not ready');
  const type = req.params.type;
  const dataMap = { critical: cache.criticalItems, overstock: cache.overstockItems, deadstock: cache.deadStockItems, stores: cache.storeAnalysis, suppliers: cache.supplierAnalysis };
  const data = dataMap[type];
  if (!data) return res.status(404).send('Not found');
  if (data.length === 0) return res.status(204).send('No data');
  const headers = Object.keys(data[0]);
  const csv = [headers.join(','), ...data.map(row => headers.map(h => `"${(row[h] || '').toString().replace(/"/g, '""')}"`).join(','))].join('\n');
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="${type}_${new Date().toISOString().split('T')[0]}.csv"`);
  res.send(csv);
});

// ─── FRONTEND HTML ────────────────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>CAMANAVA Inventory Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
<style>
:root {
  --bg: #0d1117;
  --bg2: #161b22;
  --bg3: #1c2128;
  --border: #30363d;
  --text: #e6edf3;
  --text2: #8b949e;
  --green: #1a7f37;
  --green-light: #2ea043;
  --green-bright: #3fb950;
  --green-dim: #0d4a1f;
  --red: #da3633;
  --red-light: #f85149;
  --red-dim: #4a1a1a;
  --yellow: #d29922;
  --yellow-light: #e3b341;
  --yellow-dim: #4a3200;
  --blue: #1f6feb;
  --blue-light: #388bfd;
  --blue-dim: #0d2a5e;
  --accent: #2ea043;
  --radius: 8px;
  --shadow: 0 4px 24px rgba(0,0,0,0.4);
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'IBM Plex Sans',sans-serif; background:var(--bg); color:var(--text); min-height:100vh; }
a { color: var(--green-bright); }

/* HEADER */
.header {
  background: var(--bg2);
  border-bottom: 1px solid var(--border);
  padding: 0 24px;
  height: 60px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  position: sticky; top: 0; z-index: 100;
}
.header-logo {
  display: flex; align-items: center; gap: 12px;
}
.header-logo .dot {
  width: 10px; height: 10px; border-radius: 50%;
  background: var(--green-bright);
  box-shadow: 0 0 8px var(--green-bright);
  animation: pulse 2s infinite;
}
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
.header-logo h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; color: var(--text); }
.header-logo span { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--text2); }
.header-right { display:flex; align-items:center; gap:12px; }
.refresh-info { font-size:11px; color:var(--text2); font-family:'IBM Plex Mono',monospace; }
.btn { padding:6px 14px; border-radius:var(--radius); border:1px solid var(--border); background:var(--bg3); color:var(--text); font-size:12px; cursor:pointer; transition:all 0.2s; font-family:'IBM Plex Sans',sans-serif; }
.btn:hover { border-color:var(--green-bright); color:var(--green-bright); }
.btn-green { background:var(--green); border-color:var(--green-light); color:#fff; }
.btn-green:hover { background:var(--green-light); color:#fff; }
.btn-sm { padding:4px 10px; font-size:11px; }

/* LOADING OVERLAY */
#loading-overlay {
  position:fixed; inset:0; background:var(--bg); z-index:9999;
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:16px;
}
.loader-bar {
  width:300px; height:3px; background:var(--border); border-radius:2px; overflow:hidden;
}
.loader-bar-fill {
  height:100%; background:var(--green-bright); border-radius:2px;
  animation: loadbar 2s ease-in-out infinite;
}
@keyframes loadbar { 0%{width:0;margin-left:0} 50%{width:60%;margin-left:20%} 100%{width:0;margin-left:100%} }
.loader-text { font-family:'IBM Plex Mono',monospace; font-size:13px; color:var(--text2); }

/* LAYOUT */
.main { display:flex; height:calc(100vh - 60px); }

/* SIDEBAR */
.sidebar {
  width: 260px; min-width:260px;
  background: var(--bg2);
  border-right: 1px solid var(--border);
  overflow-y: auto;
  padding: 16px;
  display: flex; flex-direction: column; gap: 12px;
}
.sidebar-title { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px; color:var(--text2); margin-bottom:4px; }
.filter-group { display:flex; flex-direction:column; gap:6px; }
.filter-label { font-size:11px; color:var(--text2); }
.filter-select {
  width:100%; padding:6px 8px; border-radius:var(--radius);
  border:1px solid var(--border); background:var(--bg3); color:var(--text);
  font-size:12px; font-family:'IBM Plex Sans',sans-serif; cursor:pointer;
}
.filter-select:focus { outline:none; border-color:var(--green-bright); }
.filter-search {
  width:100%; padding:6px 8px; border-radius:var(--radius);
  border:1px solid var(--border); background:var(--bg3); color:var(--text);
  font-size:12px; font-family:'IBM Plex Sans',sans-serif;
}
.filter-search:focus { outline:none; border-color:var(--green-bright); }
.sidebar-divider { height:1px; background:var(--border); margin:4px 0; }
.active-filters { display:flex; flex-wrap:wrap; gap:4px; }
.filter-tag {
  display:inline-flex; align-items:center; gap:4px;
  background:var(--green-dim); border:1px solid var(--green);
  color:var(--green-bright); font-size:10px; padding:2px 6px; border-radius:4px;
}
.filter-tag button { background:none; border:none; color:var(--green-bright); cursor:pointer; font-size:10px; padding:0; }

/* CONTENT */
.content { flex:1; overflow-y:auto; padding:20px; display:flex; flex-direction:column; gap:20px; }

/* NAV TABS */
.tabs { display:flex; gap:2px; border-bottom:1px solid var(--border); padding-bottom:0; }
.tab {
  padding:8px 16px; font-size:13px; cursor:pointer;
  border-bottom:2px solid transparent; color:var(--text2);
  transition:all 0.2s; white-space:nowrap;
}
.tab:hover { color:var(--text); }
.tab.active { color:var(--green-bright); border-bottom-color:var(--green-bright); }

/* KPI CARDS */
.kpi-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(180px,1fr)); gap:12px; }
.kpi-card {
  background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius);
  padding:16px; display:flex; flex-direction:column; gap:6px;
  transition:border-color 0.2s;
}
.kpi-card:hover { border-color:var(--green-bright); }
.kpi-label { font-size:10px; text-transform:uppercase; letter-spacing:0.8px; color:var(--text2); font-weight:600; }
.kpi-value { font-size:22px; font-weight:700; font-family:'IBM Plex Mono',monospace; color:var(--text); }
.kpi-sub { font-size:11px; color:var(--text2); }
.kpi-card.red { border-left:3px solid var(--red); }
.kpi-card.yellow { border-left:3px solid var(--yellow); }
.kpi-card.green { border-left:3px solid var(--green-bright); }
.kpi-card.blue { border-left:3px solid var(--blue); }
.kpi-card.red .kpi-value { color:var(--red-light); }
.kpi-card.yellow .kpi-value { color:var(--yellow-light); }
.kpi-card.green .kpi-value { color:var(--green-bright); }
.kpi-card.blue .kpi-value { color:var(--blue-light); }

/* SECTION */
.section { background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius); overflow:hidden; }
.section-header {
  padding:12px 16px; border-bottom:1px solid var(--border);
  display:flex; align-items:center; justify-content:space-between;
}
.section-title { font-size:13px; font-weight:600; }
.section-actions { display:flex; gap:8px; align-items:center; }
.badge { padding:2px 8px; border-radius:12px; font-size:11px; font-weight:600; }
.badge-red { background:var(--red-dim); color:var(--red-light); }
.badge-yellow { background:var(--yellow-dim); color:var(--yellow-light); }
.badge-green { background:var(--green-dim); color:var(--green-bright); }
.badge-blue { background:var(--blue-dim); color:var(--blue-light); }

/* TABLE */
.table-wrap { overflow-x:auto; max-height:420px; overflow-y:auto; }
table { width:100%; border-collapse:collapse; font-size:12px; }
thead th {
  position:sticky; top:0; z-index:2;
  background:#1a4731; color:#fff;
  padding:8px 12px; text-align:left;
  font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:0.5px;
  white-space:nowrap; cursor:pointer; user-select:none;
}
thead th:hover { background:#1f5c3d; }
tbody tr { border-bottom:1px solid var(--border); transition:background 0.1s; }
tbody tr:hover { background:var(--bg3); }
tbody td { padding:7px 12px; white-space:nowrap; color:var(--text); }
tbody td.mono { font-family:'IBM Plex Mono',monospace; font-size:11px; }
.action-badge {
  padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; white-space:nowrap;
}
.action-urgent { background:var(--red-dim); color:var(--red-light); }
.action-po { background:var(--blue-dim); color:var(--blue-light); }
.action-monitor { background:var(--yellow-dim); color:var(--yellow-light); }
.action-review { background:var(--bg3); color:var(--text2); }
.action-markdown { background:#3a1a4a; color:#c084fc; }

/* SEARCH IN TABLE */
.table-search {
  padding:6px 10px; border-radius:var(--radius);
  border:1px solid var(--border); background:var(--bg3); color:var(--text);
  font-size:12px; width:180px; font-family:'IBM Plex Sans',sans-serif;
}
.table-search:focus { outline:none; border-color:var(--green-bright); }

/* CHARTS */
.charts-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
@media(max-width:900px){ .charts-grid { grid-template-columns:1fr; } }
.chart-card { background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius); padding:16px; }
.chart-title { font-size:12px; font-weight:600; color:var(--text2); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:12px; }
canvas { max-height:260px; }

/* STATUS BAR */
.status-bar {
  background:var(--bg3); border:1px solid var(--border); border-radius:var(--radius);
  padding:8px 16px; display:flex; align-items:center; gap:16px; font-size:11px; color:var(--text2);
}
.status-dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.status-dot.green { background:var(--green-bright); box-shadow:0 0 6px var(--green-bright); }
.status-dot.yellow { background:var(--yellow-light); }
.status-dot.red { background:var(--red-light); }

/* PAGINATION */
.pagination { display:flex; align-items:center; gap:8px; padding:10px 16px; border-top:1px solid var(--border); font-size:12px; }
.page-btn { padding:4px 10px; border-radius:4px; border:1px solid var(--border); background:var(--bg3); color:var(--text); cursor:pointer; font-size:12px; }
.page-btn:hover { border-color:var(--green-bright); }
.page-btn.active { background:var(--green); border-color:var(--green-light); color:#fff; }
.page-info { color:var(--text2); }

/* EMPTY STATE */
.empty { padding:40px; text-align:center; color:var(--text2); font-size:13px; }

/* WTS HEATMAP */
.heatmap-grid { display:grid; gap:3px; padding:16px; }
.heatmap-row { display:flex; align-items:center; gap:6px; font-size:11px; }
.heatmap-label { width:140px; text-align:right; color:var(--text2); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.heatmap-bar { flex:1; height:20px; border-radius:3px; position:relative; }
.heatmap-bar span { position:absolute; right:6px; top:50%; transform:translateY(-50%); font-size:10px; font-weight:600; font-family:'IBM Plex Mono',monospace; color:#fff; }

/* SCROLLBAR */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:var(--bg); }
::-webkit-scrollbar-thumb { background:var(--border); border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background:var(--text2); }
</style>
</head>
<body>

<div id="loading-overlay">
  <div style="display:flex;align-items:center;gap:10px;">
    <div class="dot" style="width:10px;height:10px;border-radius:50%;background:#3fb950;box-shadow:0 0 8px #3fb950;animation:pulse 1s infinite;"></div>
    <span style="font-family:'IBM Plex Mono',monospace;font-size:14px;font-weight:600;color:#e6edf3;">CAMANAVA INVENTORY</span>
  </div>
  <div class="loader-bar"><div class="loader-bar-fill"></div></div>
  <div class="loader-text" id="loading-msg">Connecting to data source...</div>
</div>

<div id="app" style="display:none;">

<header class="header">
  <div class="header-logo">
    <div class="dot"></div>
    <div>
      <h1>CAMANAVA INVENTORY DASHBOARD</h1>
      <span id="header-sub">Loading...</span>
    </div>
  </div>
  <div class="header-right">
    <div class="refresh-info" id="refresh-info">–</div>
    <button class="btn btn-green" onclick="triggerRefresh()">↺ Refresh</button>
  </div>
</header>

<div class="main">

  <!-- SIDEBAR FILTERS -->
  <aside class="sidebar">
    <div>
      <div class="sidebar-title">Filters</div>
    </div>
    <div class="filter-group">
      <div class="filter-label">Area</div>
      <select class="filter-select" id="f-area" onchange="applyFilter('area',this.value)">
        <option value="">All Areas</option>
      </select>
    </div>
    <div class="filter-group">
      <div class="filter-label">Store</div>
      <select class="filter-select" id="f-store" onchange="applyFilter('store',this.value)">
        <option value="">All Stores</option>
      </select>
    </div>
    <div class="sidebar-divider"></div>
    <div class="filter-group">
      <div class="filter-label">Department</div>
      <select class="filter-select" id="f-dept" onchange="applyFilter('dept',this.value)">
        <option value="">All Departments</option>
      </select>
    </div>
    <div class="filter-group">
      <div class="filter-label">Sub-Department</div>
      <select class="filter-select" id="f-subdept" onchange="applyFilter('subDept',this.value)">
        <option value="">All Sub-Depts</option>
      </select>
    </div>
    <div class="filter-group">
      <div class="filter-label">Class</div>
      <select class="filter-select" id="f-cls" onchange="applyFilter('cls',this.value)">
        <option value="">All Classes</option>
      </select>
    </div>
    <div class="sidebar-divider"></div>
    <div class="filter-group">
      <div class="filter-label">Supplier</div>
      <select class="filter-select" id="f-supplier" onchange="applyFilter('supplier',this.value)">
        <option value="">All Suppliers</option>
      </select>
    </div>
    <div class="filter-group">
      <div class="filter-label">Brand</div>
      <select class="filter-select" id="f-brand" onchange="applyFilter('brand',this.value)">
        <option value="">All Brands</option>
      </select>
    </div>
    <div class="filter-group">
      <div class="filter-label">SKU Status</div>
      <select class="filter-select" id="f-skustatus" onchange="applyFilter('skuStatus',this.value)">
        <option value="">All Statuses</option>
      </select>
    </div>
    <div class="sidebar-divider"></div>
    <div>
      <div class="sidebar-title">Active Filters</div>
      <div class="active-filters" id="active-filters"><span style="font-size:11px;color:var(--text2);">None</span></div>
    </div>
    <button class="btn" style="margin-top:4px;" onclick="clearFilters()">✕ Clear All Filters</button>
  </aside>

  <!-- MAIN CONTENT -->
  <main class="content">

    <!-- STATUS BAR -->
    <div class="status-bar" id="status-bar">
      <div class="status-dot green" id="status-dot"></div>
      <span id="status-text">Ready</span>
      <span style="margin-left:auto;" id="status-rows">–</span>
    </div>

    <!-- TABS -->
    <div class="tabs">
      <div class="tab active" onclick="showTab('overview')">Overview</div>
      <div class="tab" onclick="showTab('critical')">⚠ Critical</div>
      <div class="tab" onclick="showTab('overstock')">📦 Overstock</div>
      <div class="tab" onclick="showTab('deadstock')">💀 Dead Stock</div>
      <div class="tab" onclick="showTab('stores')">🏪 Stores</div>
      <div class="tab" onclick="showTab('suppliers')">🏭 Suppliers</div>
    </div>

    <!-- OVERVIEW TAB -->
    <div id="tab-overview">
      <div class="kpi-grid" id="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Loading...</div></div>
      </div>
      <div class="charts-grid" style="margin-top:4px;">
        <div class="chart-card">
          <div class="chart-title">Top 10 Stores by Inventory Value</div>
          <canvas id="chart-stores"></canvas>
        </div>
        <div class="chart-card">
          <div class="chart-title">Top 10 Suppliers by Inventory Value</div>
          <canvas id="chart-suppliers"></canvas>
        </div>
        <div class="chart-card">
          <div class="chart-title">Inventory Health by Area</div>
          <canvas id="chart-areas"></canvas>
        </div>
        <div class="chart-card">
          <div class="chart-title">Inventory Risk Distribution</div>
          <canvas id="chart-risk"></canvas>
        </div>
      </div>
    </div>

    <!-- CRITICAL TAB -->
    <div id="tab-critical" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">⚠ Critical Stock <span class="badge badge-red" id="critical-count">0</span></div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('critical-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('critical')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="critical-table">
            <thead><tr>
              <th onclick="sortTable('critical-table',0)">Store</th>
              <th onclick="sortTable('critical-table',1)">Area</th>
              <th onclick="sortTable('critical-table',2)">SKU Code</th>
              <th onclick="sortTable('critical-table',3)">Description</th>
              <th onclick="sortTable('critical-table',4)">Supplier</th>
              <th onclick="sortTable('critical-table',5)">On Hand</th>
              <th onclick="sortTable('critical-table',6)">Value</th>
              <th onclick="sortTable('critical-table',7)">Cur Wk Sales</th>
              <th onclick="sortTable('critical-table',8)">P8 Ave</th>
              <th onclick="sortTable('critical-table',9)">WTS Net</th>
              <th onclick="sortTable('critical-table',10)">Total PO</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="critical-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="critical-pagination"></div>
      </div>
    </div>

    <!-- OVERSTOCK TAB -->
    <div id="tab-overstock" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">📦 Overstock Items <span class="badge badge-yellow" id="overstock-count">0</span></div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('overstock-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('overstock')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="overstock-table">
            <thead><tr>
              <th onclick="sortTable('overstock-table',0)">Store</th>
              <th onclick="sortTable('overstock-table',1)">Area</th>
              <th onclick="sortTable('overstock-table',2)">SKU Code</th>
              <th onclick="sortTable('overstock-table',3)">Description</th>
              <th onclick="sortTable('overstock-table',4)">Supplier</th>
              <th onclick="sortTable('overstock-table',5)">On Hand</th>
              <th onclick="sortTable('overstock-table',6)">Value</th>
              <th onclick="sortTable('overstock-table',7)">P8 Ave</th>
              <th onclick="sortTable('overstock-table',8)">WTS Net</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="overstock-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="overstock-pagination"></div>
      </div>
    </div>

    <!-- DEAD STOCK TAB -->
    <div id="tab-deadstock" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">💀 Dead Stock <span class="badge badge-red" id="deadstock-count">0</span></div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('deadstock-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('deadstock')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="deadstock-table">
            <thead><tr>
              <th onclick="sortTable('deadstock-table',0)">Store</th>
              <th onclick="sortTable('deadstock-table',1)">Area</th>
              <th onclick="sortTable('deadstock-table',2)">SKU Code</th>
              <th onclick="sortTable('deadstock-table',3)">Description</th>
              <th onclick="sortTable('deadstock-table',4)">Supplier</th>
              <th onclick="sortTable('deadstock-table',5)">On Hand</th>
              <th onclick="sortTable('deadstock-table',6)">Value</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="deadstock-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="deadstock-pagination"></div>
      </div>
    </div>

    <!-- STORES TAB -->
    <div id="tab-stores" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">🏪 Store Analysis</div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search store..." oninput="searchTable('stores-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('stores')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="stores-table">
            <thead><tr>
              <th onclick="sortTable('stores-table',0)">Store #</th>
              <th onclick="sortTable('stores-table',1)">Store Name</th>
              <th onclick="sortTable('stores-table',2)">Area</th>
              <th onclick="sortTable('stores-table',3)">Inv Value</th>
              <th onclick="sortTable('stores-table',4)">On Hand</th>
              <th onclick="sortTable('stores-table',5)">SKUs</th>
              <th onclick="sortTable('stores-table',6)">Cur Wk Sales</th>
              <th onclick="sortTable('stores-table',7)">Critical</th>
              <th onclick="sortTable('stores-table',8)">Overstock</th>
              <th onclick="sortTable('stores-table',9)">Dead Stock</th>
            </tr></thead>
            <tbody id="stores-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="stores-pagination"></div>
      </div>
    </div>

    <!-- SUPPLIERS TAB -->
    <div id="tab-suppliers" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">🏭 Supplier Analysis</div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search supplier..." oninput="searchTable('suppliers-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('suppliers')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="suppliers-table">
            <thead><tr>
              <th onclick="sortTable('suppliers-table',0)">Supplier Code</th>
              <th onclick="sortTable('suppliers-table',1)">Supplier Name</th>
              <th onclick="sortTable('suppliers-table',2)">Inv Value</th>
              <th onclick="sortTable('suppliers-table',3)">On Hand</th>
              <th onclick="sortTable('suppliers-table',4)">SKUs</th>
              <th onclick="sortTable('suppliers-table',5)">Critical</th>
              <th onclick="sortTable('suppliers-table',6)">Overstock</th>
            </tr></thead>
            <tbody id="suppliers-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="suppliers-pagination"></div>
      </div>
    </div>

  </main>
</div><!-- end main -->
</div><!-- end app -->

<script>
// ─── STATE ────────────────────────────────────────────────────────────────────
let activeFilters = {};
let activeTab = 'overview';
let charts = {};
let tablePages = { critical:1, overstock:1, deadstock:1, stores:1, suppliers:1 };
const PAGE_SIZE = 50;

// ─── INIT ─────────────────────────────────────────────────────────────────────
async function init() {
  setLoadingMsg('Checking server status...');
  let ready = false;
  for (let i = 0; i < 3; i++) {
    try {
      const r = await fetch('/api/status?t=' + Date.now());
      const text = await r.text();
      const d = JSON.parse(text);
      if (d && d.ready === true) { ready = true; break; }
    } catch(e) {}
    await sleep(1000);
  }
  document.getElementById('loading-overlay').style.display = 'none';
  document.getElementById('app').style.display = 'flex';
  document.getElementById('app').style.flexDirection = 'column';
  await loadFilters();
  await loadAll();
  if (!ready) { pollAndRefresh(); }
}

async function pollAndRefresh() {
  while (true) {
    await sleep(3000);
    try {
      const r = await fetch('/api/status?t=' + Date.now());
      const text = await r.text();
      const d = JSON.parse(text);
      if (d && d.ready === true) { await loadFilters(); await loadAll(); return; }
    } catch(e) {}
  }
}

function setLoadingMsg(msg) {
  document.getElementById('loading-msg').textContent = msg;
}

async function pollUntilReady(maxWait = 300000) {
  const start = Date.now();
  while (Date.now() - start < maxWait) {
    try {
      const r = await fetch('/api/status?t=' + Date.now());
      if (!r.ok) { await sleep(2000); continue; }
      const text = await r.text();
      const d = JSON.parse(text);
      if (d && d.ready === true) return true;
      if (d && d.error && !d.refreshing) {
        setLoadingMsg('Error: ' + d.error);
        await sleep(5000);
      } else {
        setLoadingMsg(d && d.refreshing ? 'Processing inventory data... (120MB file, please wait)' : 'Waiting for data...');
        await sleep(3000);
      }
    } catch(e) {
      setLoadingMsg('Connecting... (' + e.message + ')');
      await sleep(2000);
    }
  }
  return false;
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function loadAll() {
  await Promise.all([loadKPIs(), loadTabData()]);
  updateStatus();
}

async function updateStatus() {
  const r = await fetch('/api/status');
  const d = await r.json();
  document.getElementById('refresh-info').textContent = d.lastRefresh ? 'Last refresh: ' + new Date(d.lastRefresh).toLocaleString() : '–';
  document.getElementById('header-sub').textContent = d.totalRows ? fmt(d.totalRows) + ' SKU rows loaded' : 'CAMANAVA';
  document.getElementById('status-rows').textContent = d.totalRows ? fmt(d.totalRows) + ' rows in cache' : '';
  const dot = document.getElementById('status-dot');
  const txt = document.getElementById('status-text');
  if (d.ready) { dot.className = 'status-dot green'; txt.textContent = 'Cache ready — Dashboard live'; }
  else if (d.error) { dot.className = 'status-dot red'; txt.textContent = 'Error: ' + d.error; }
  else { dot.className = 'status-dot yellow'; txt.textContent = 'Refreshing data...'; }
}

// ─── FILTERS ──────────────────────────────────────────────────────────────────
async function loadFilters() {
  const r = await fetch('/api/filters');
  const d = await r.json();
  if (d.error) return;

  populateSelect('f-area', d.areas, 'All Areas');
  populateSelect('f-dept', d.depts, 'All Departments');
  populateSelect('f-subdept', d.subDepts, 'All Sub-Depts');
  populateSelect('f-cls', d.classes, 'All Classes');
  populateSelect('f-supplier', d.suppliers, 'All Suppliers');
  populateSelect('f-brand', d.brands, 'All Brands');
  populateSelect('f-skustatus', d.skuStatuses, 'All Statuses');

  const storeSelect = document.getElementById('f-store');
  storeSelect.innerHTML = '<option value="">All Stores</option>';
  (d.stores || []).forEach(s => {
    const o = document.createElement('option');
    o.value = s.id; o.textContent = s.id + ' - ' + s.name;
    storeSelect.appendChild(o);
  });
}

function populateSelect(id, items, defaultText) {
  const sel = document.getElementById(id);
  sel.innerHTML = '<option value="">' + defaultText + '</option>';
  (items || []).forEach(v => {
    const o = document.createElement('option');
    o.value = v; o.textContent = v;
    sel.appendChild(o);
  });
}

function applyFilter(key, value) {
  if (value) activeFilters[key] = value;
  else delete activeFilters[key];
  renderActiveTags();
  loadAll();
}

function removeFilterTag(btn) {
  const key = btn.getAttribute('data-key');
  if (key) applyFilter(key, '');
}

function clearFilters() {
  activeFilters = {};
  ['f-area','f-store','f-dept','f-subdept','f-cls','f-supplier','f-brand','f-skustatus']
    .forEach(id => { document.getElementById(id).value = ''; });
  renderActiveTags();
  loadAll();
}

function renderActiveTags() {
  const cont = document.getElementById('active-filters');
  const entries = Object.entries(activeFilters);
  if (entries.length === 0) { cont.innerHTML = '<span style="font-size:11px;color:var(--text2);">None</span>'; return; }
  let html = '';
  entries.forEach(function(entry) {
    const k = entry[0]; const v = entry[1];
    html += '<span class="filter-tag">' + k + ': ' + v.substring(0,15) + '<button onclick="removeFilterTag(this)" data-key="' + k + '">x</button></span>';
  });
  cont.innerHTML = html;
}

function filterQuery() {
  const params = new URLSearchParams(activeFilters);
  return params.toString() ? '?' + params.toString() : '';
}

// ─── KPIs ─────────────────────────────────────────────────────────────────────
async function loadKPIs() {
  const r = await fetch('/api/kpis' + filterQuery());
  const d = await r.json();
  if (d.error) return;
  const grid = document.getElementById('kpi-grid');
  grid.innerHTML = [
    kpiCard('Total Inv Value', '₱' + fmtM(d.totalOnHandValue), 'w/ VAT', 'green'),
    kpiCard('On Hand Qty', fmt(d.totalOnHand), 'units', 'blue'),
    kpiCard('Critical SKUs', fmt(d.criticalCount), 'WTS < 2 weeks', 'red'),
    kpiCard('Overstock SKUs', fmt(d.overstockCount), 'WTS > 12 weeks', 'yellow'),
    kpiCard('Dead Stock SKUs', fmt(d.deadStockCount), 'No sales 8 wks', 'red'),
    kpiCard('Active Stores', fmt(d.activeStores), 'stores', 'green'),
    kpiCard('Suppliers', fmt(d.activeSuppliers), 'active', 'blue'),
    kpiCard('PO Value', '₱' + fmtM(d.totalPOValue), 'incoming', 'blue'),
    kpiCard('Transfer Value', '₱' + fmtM(d.totalTRFValue), 'incoming', 'yellow'),
    kpiCard('Avg WTS Net', (d.avgWts || 0).toFixed(1) + ' wks', 'weeks to sell', d.avgWts < 4 ? 'red' : d.avgWts > 12 ? 'yellow' : 'green'),
    kpiCard('Total SKUs', fmt(d.totalSKUs), 'in scope', 'blue'),
  ].join('');
}

function kpiCard(label, value, sub, type) {
  return '<div class="kpi-card ' + type + '"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>';
}

// ─── CHARTS ───────────────────────────────────────────────────────────────────
async function loadCharts() {
  const [storesRes, suppliersRes] = await Promise.all([
    fetch('/api/stores' + filterQuery()),
    fetch('/api/suppliers' + filterQuery())
  ]);
  const stores = await storesRes.json();
  const suppliers = await suppliersRes.json();
  if (!Array.isArray(stores) || !Array.isArray(suppliers)) return;

  // Destroy existing charts
  Object.values(charts).forEach(c => { try { c.destroy(); } catch(e){} });
  charts = {};

  // Top 10 Stores by Value
  const top10stores = stores.slice(0,10);
  charts.stores = new Chart(document.getElementById('chart-stores'), {
    type: 'bar',
    data: {
      labels: top10stores.map(s => s.storeNumber + '-' + s.storeName.substring(0,12)),
      datasets: [{ label: 'Inv Value', data: top10stores.map(s => s.totalValue), backgroundColor: '#2ea043', borderRadius: 4 }]
    },
    options: { responsive:true, plugins:{legend:{display:false}}, scales:{ y:{ ticks:{ color:'#8b949e', callback: v => '₱'+fmtM(v) }, grid:{color:'#30363d'} }, x:{ ticks:{color:'#8b949e',font:{size:9}}, grid:{display:false} } } }
  });

  // Top 10 Suppliers
  const top10sup = suppliers.slice(0,10);
  charts.suppliers = new Chart(document.getElementById('chart-suppliers'), {
    type: 'bar',
    data: {
      labels: top10sup.map(s => s.supplierName.substring(0,15)),
      datasets: [{ label: 'Inv Value', data: top10sup.map(s => s.totalValue), backgroundColor: '#1f6feb', borderRadius: 4 }]
    },
    options: { indexAxis:'y', responsive:true, plugins:{legend:{display:false}}, scales:{ x:{ ticks:{ color:'#8b949e', callback: v => '₱'+fmtM(v) }, grid:{color:'#30363d'} }, y:{ ticks:{color:'#8b949e',font:{size:9}}, grid:{display:false} } } }
  });

  // Inventory Health by Area
  const areaMap = {};
  stores.forEach(s => {
    if (!s.area) return;
    if (!areaMap[s.area]) areaMap[s.area] = { value:0, critical:0, overstock:0 };
    areaMap[s.area].value += s.totalValue;
    areaMap[s.area].critical += s.criticalCount;
    areaMap[s.area].overstock += s.overstockCount;
  });
  const areas = Object.keys(areaMap);
  charts.areas = new Chart(document.getElementById('chart-areas'), {
    type: 'bar',
    data: {
      labels: areas,
      datasets: [
        { label: 'Inv Value', data: areas.map(a => areaMap[a].value), backgroundColor: '#2ea043', borderRadius: 4, yAxisID: 'y' },
        { label: 'Critical', data: areas.map(a => areaMap[a].critical), backgroundColor: '#f85149', borderRadius: 4, yAxisID: 'y1' }
      ]
    },
    options: { responsive:true, plugins:{legend:{labels:{color:'#8b949e',font:{size:10}}}}, scales:{ y:{ ticks:{color:'#8b949e',callback: v=>'₱'+fmtM(v)}, grid:{color:'#30363d'} }, y1:{position:'right',ticks:{color:'#8b949e'},grid:{display:false}}, x:{ticks:{color:'#8b949e',font:{size:9}},grid:{display:false}} } }
  });

  // Risk Distribution
  const critTotal = stores.reduce((s,r) => s + r.criticalCount, 0);
  const ovTotal = stores.reduce((s,r) => s + r.overstockCount, 0);
  const deadTotal = stores.reduce((s,r) => s + r.deadCount, 0);
  const totalSKUs = stores.reduce((s,r) => s + r.totalSKUs, 0);
  const normal = Math.max(0, totalSKUs - critTotal - ovTotal - deadTotal);
  charts.risk = new Chart(document.getElementById('chart-risk'), {
    type: 'doughnut',
    data: {
      labels: ['Normal', 'Overstock', 'Critical', 'Dead Stock'],
      datasets: [{ data: [normal, ovTotal, critTotal, deadTotal], backgroundColor: ['#2ea043','#e3b341','#f85149','#8b949e'], borderWidth: 0 }]
    },
    options: { responsive:true, plugins:{ legend:{ position:'bottom', labels:{ color:'#8b949e', font:{size:10} } } } }
  });
}

// ─── TABLE LOADING ────────────────────────────────────────────────────────────
async function loadTabData() {
  if (activeTab === 'overview') { await loadCharts(); return; }
  if (activeTab === 'critical') await loadCritical();
  if (activeTab === 'overstock') await loadOverstock();
  if (activeTab === 'deadstock') await loadDeadstock();
  if (activeTab === 'stores') await loadStores();
  if (activeTab === 'suppliers') await loadSuppliers();
}

async function loadCritical() {
  const r = await fetch('/api/critical' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  document.getElementById('critical-count').textContent = fmt(data.length);
  renderTable('critical-body', data, renderCriticalRow, 'critical-pagination', 'critical', tablePages.critical);
}
async function loadOverstock() {
  const r = await fetch('/api/overstock' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  document.getElementById('overstock-count').textContent = fmt(data.length);
  renderTable('overstock-body', data, renderOverstockRow, 'overstock-pagination', 'overstock', tablePages.overstock);
}
async function loadDeadstock() {
  const r = await fetch('/api/deadstock' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  document.getElementById('deadstock-count').textContent = fmt(data.length);
  renderTable('deadstock-body', data, renderDeadstockRow, 'deadstock-pagination', 'deadstock', tablePages.deadstock);
}
async function loadStores() {
  const r = await fetch('/api/stores' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  renderTable('stores-body', data, renderStoreRow, 'stores-pagination', 'stores', tablePages.stores);
}
async function loadSuppliers() {
  const r = await fetch('/api/suppliers' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  renderTable('suppliers-body', data, renderSupplierRow, 'suppliers-pagination', 'suppliers', tablePages.suppliers);
}

function renderTable(bodyId, data, rowFn, paginationId, key, page) {
  const tbody = document.getElementById(bodyId);
  const total = data.length;
  const totalPages = Math.ceil(total / PAGE_SIZE);
  const currentPage = Math.min(page, Math.max(1, totalPages));
  tablePages[key] = currentPage;
  const start = (currentPage - 1) * PAGE_SIZE;
  const pageData = data.slice(start, start + PAGE_SIZE);
  if (pageData.length === 0) { tbody.innerHTML = '<tr><td colspan="20" class="empty">No data found</td></tr>'; }
  else { tbody.innerHTML = pageData.map(rowFn).join(''); }
  renderPagination(paginationId, currentPage, totalPages, key, data);
}

function renderPagination(id, current, total, key, data) {
  const cont = document.getElementById(id);
  if (total <= 1) { cont.innerHTML = '<span class="page-info">' + fmt(data.length) + ' rows</span>'; return; }
  let html = '<span class="page-info">' + fmt(data.length) + ' rows | Page ' + current + ' of ' + total + '</span>';
  html += '<button class="page-btn" data-key="' + key + '" data-page="1" onclick="pageBtnClick(this)">&laquo;</button>';
  html += '<button class="page-btn" data-key="' + key + '" data-page="' + Math.max(1,current-1) + '" onclick="pageBtnClick(this)">&lsaquo;</button>';
  const startP = Math.max(1, current - 2), endP = Math.min(total, current + 2);
  for (let i = startP; i <= endP; i++) {
    html += '<button class="page-btn' + (i===current?' active':'') + '" data-key="' + key + '" data-page="' + i + '" onclick="pageBtnClick(this)">' + i + '</button>';
  }
  html += '<button class="page-btn" data-key="' + key + '" data-page="' + Math.min(total,current+1) + '" onclick="pageBtnClick(this)">&rsaquo;</button>';
  html += '<button class="page-btn" data-key="' + key + '" data-page="' + total + '" onclick="pageBtnClick(this)">&raquo;</button>';
  cont.innerHTML = html;
}

function pageBtnClick(btn) {
  const key = btn.getAttribute('data-key');
  const page = parseInt(btn.getAttribute('data-page'));
  goPage(key, page);
}

function goPage(key, page) {
  tablePages[key] = page;
  loadTabData();
}

// ─── ROW RENDERERS ─────────────────────────────────────────────────────────────
function renderCriticalRow(r) {
  const ac = r.action === 'URGENT: Place PO' ? 'action-urgent' : r.action === 'PO Incoming' ? 'action-po' : 'action-review';
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono">' + fmt(r.currentWkSales) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono" style="color:var(--red-light);font-weight:600;">' + fmtN(r.wtsNet) + '</td>' +
    '<td class="mono">' + fmt(r.totalPO) + '</td>' +
    '<td><span class="action-badge ' + ac + '">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderOverstockRow(r) {
  const ac = r.action === 'Consider Markdown' ? 'action-markdown' : 'action-monitor';
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono" style="color:var(--yellow-light);font-weight:600;">' + r.wtsNet + '</td>' +
    '<td><span class="action-badge ' + ac + '">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderDeadstockRow(r) {
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td><span class="action-badge action-markdown">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderStoreRow(r) {
  const ci = r.criticalCount > 0 ? '<span style="color:var(--red-light);font-weight:600;">' + fmt(r.criticalCount) + '</span>' : '0';
  const ov = r.overstockCount > 0 ? '<span style="color:var(--yellow-light);">' + fmt(r.overstockCount) + '</span>' : '0';
  const dd = r.deadCount > 0 ? '<span style="color:var(--text2);">' + fmt(r.deadCount) + '</span>' : '0';
  return '<tr>' +
    '<td class="mono" style="font-weight:600;">' + esc(r.storeNumber) + '</td>' +
    '<td>' + esc(r.storeName) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(r.totalValue) + '</td>' +
    '<td class="mono">' + fmt(r.totalOnHand) + '</td>' +
    '<td class="mono">' + fmt(r.totalSKUs) + '</td>' +
    '<td class="mono">' + fmt(r.totalSales) + '</td>' +
    '<td>' + ci + '</td>' +
    '<td>' + ov + '</td>' +
    '<td>' + dd + '</td>' +
    '</tr>';
}
function renderSupplierRow(r) {
  const ci = r.criticalCount > 0 ? '<span style="color:var(--red-light);font-weight:600;">' + fmt(r.criticalCount) + '</span>' : '0';
  const ov = r.overstockCount > 0 ? '<span style="color:var(--yellow-light);">' + fmt(r.overstockCount) + '</span>' : '0';
  return '<tr>' +
    '<td class="mono">' + esc(r.supplierCode) + '</td>' +
    '<td>' + esc(r.supplierName) + '</td>' +
    '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(r.totalValue) + '</td>' +
    '<td class="mono">' + fmt(r.totalOnHand) + '</td>' +
    '<td class="mono">' + fmt(r.totalSKUs) + '</td>' +
    '<td>' + ci + '</td>' +
    '<td>' + ov + '</td>' +
    '</tr>';
}

// ─── TABS ─────────────────────────────────────────────────────────────────────
function showTab(name) {
  ['overview','critical','overstock','deadstock','stores','suppliers'].forEach(t => {
    document.getElementById('tab-' + t).style.display = t === name ? '' : 'none';
  });
  document.querySelectorAll('.tab').forEach((el, i) => {
    el.classList.toggle('active', ['overview','critical','overstock','deadstock','stores','suppliers'][i] === name);
  });
  activeTab = name;
  loadTabData();
}

// ─── TABLE SEARCH ─────────────────────────────────────────────────────────────
function searchTable(tableId, query) {
  const q = query.toLowerCase();
  document.querySelectorAll('#' + tableId + ' tbody tr').forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
}

// ─── TABLE SORT ───────────────────────────────────────────────────────────────
let sortState = {};
function sortTable(tableId, colIndex) {
  const table = document.getElementById(tableId);
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const key = tableId + '_' + colIndex;
  const asc = sortState[key] !== true;
  sortState[key] = asc;
  rows.sort((a, b) => {
    const av = a.cells[colIndex]?.textContent?.trim() || '';
    const bv = b.cells[colIndex]?.textContent?.trim() || '';
    const an = parseFloat(av.replace(/[₱,]/g, ''));
    const bn = parseFloat(bv.replace(/[₱,]/g, ''));
    if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  });
  rows.forEach(r => tbody.appendChild(r));
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function exportData(type) {
  window.open('/api/export/' + type + filterQuery(), '_blank');
}

// ─── REFRESH ──────────────────────────────────────────────────────────────────
async function triggerRefresh() {
  document.getElementById('status-text').textContent = 'Refresh triggered...';
  await fetch('/api/refresh', { method:'POST' });
  await sleep(2000);
  await updateStatus();
  await loadAll();
}

// ─── UTILS ────────────────────────────────────────────────────────────────────
function fmt(n) { return (n||0).toLocaleString(); }
function fmtN(n) { return (Math.round(n*100)/100).toLocaleString(undefined,{minimumFractionDigits:0,maximumFractionDigits:2}); }
function fmtM(n) {
  if (n >= 1e9) return (n/1e9).toFixed(2) + 'B';
  if (n >= 1e6) return (n/1e6).toFixed(2) + 'M';
  if (n >= 1e3) return (n/1e3).toFixed(1) + 'K';
  return n.toFixed(0);
}
function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── START ────────────────────────────────────────────────────────────────────
init();
</script>
</body>
</html>`);
});

// ─── START SERVER ─────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`[Server] CAMANAVA Inventory Dashboard running on port ${PORT}`);
  if (!GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY || !GDRIVE_FOLDER_ID) {
    console.warn('[Server] WARNING: Missing Google Drive env vars. Set GOOGLE_CLIENT_EMAIL, GOOGLE_PRIVATE_KEY, GDRIVE_FOLDER_ID');
  } else {
    console.log('[Server] Starting initial data load...');
    refreshData(true);
  }
});
