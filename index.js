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
const LOGS_SHEET_ID = process.env.LOGS_SHEET_ID || '';

// ─── IN-MEMORY CACHE ──────────────────────────────────────────────────────────
let cache = {
  ready: false,
  lastRefresh: null,
  lastFileHash: null,
  lastFileSize: null,
  lastModifiedTime: null,
  rows: [],
  storeMap: {},       // storeId -> { area, storeName, region }
  users: {},          // username -> { username, password, level, area }
  kpis: {},
  criticalItems: [],
  overstockItems: [],
  deadStockItems: [],
  outOfStockItems: [],
  storeAnalysis: [],
  supplierAnalysis: [],
  filterMeta: {},
  refreshing: false,
  error: null
};

// ─── GOOGLE DRIVE AUTH ────────────────────────────────────────────────────────
function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: GOOGLE_CLIENT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY
    },
    scopes: [
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/spreadsheets'
    ]
  });
}

function getDriveClient() {
  return google.drive({ version: 'v3', auth: getAuth() });
}

function getSheetsClient() {
  return google.sheets({ version: 'v4', auth: getAuth() });
}

// ─── ACTIVITY LOG (Google Sheets) ─────────────────────────────────────────────
// Appends a login row, returns the row number for later logout update
async function logLoginEvent(username, area) {
  if (!LOGS_SHEET_ID) { console.warn('[Logs] LOGS_SHEET_ID not set, skipping log'); return null; }
  try {
    const sheets = getSheetsClient();
    const loginTime = new Date().toISOString();
    const resp = await sheets.spreadsheets.values.append({
      spreadsheetId: LOGS_SHEET_ID,
      range: 'A:E',
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[username, loginTime, '', '', area || 'All']] }
    });
    // Extract row number from updatedRange like "Logs!A5:E5"
    const updatedRange = resp.data.updates && resp.data.updates.updatedRange;
    let rowNum = null;
    if (updatedRange) {
      const m = updatedRange.match(/!\w?(\d+):/);
      if (m) rowNum = parseInt(m[1]);
    }
    return { rowNum, loginTime };
  } catch (e) {
    console.error('[Logs] Login log error:', e.message);
    return null;
  }
}

// Updates the logout time + duration for a given row
async function logLogoutEvent(rowNum, loginTimeISO) {
  if (!LOGS_SHEET_ID || !rowNum) return;
  try {
    const sheets = getSheetsClient();
    const logoutTime = new Date();
    const loginTime = new Date(loginTimeISO);
    const durationMs = logoutTime - loginTime;
    const mins = Math.floor(durationMs / 60000);
    const hrs = Math.floor(mins / 60);
    const remMins = mins % 60;
    const durationStr = hrs > 0 ? (hrs + 'h ' + remMins + 'm') : (mins + 'm');
    await sheets.spreadsheets.values.update({
      spreadsheetId: LOGS_SHEET_ID,
      range: 'C' + rowNum + ':D' + rowNum,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[logoutTime.toISOString(), durationStr]] }
    });
  } catch (e) {
    console.error('[Logs] Logout log error:', e.message);
  }
}

// Reads all log rows
async function readLogs() {
  if (!LOGS_SHEET_ID) return [];
  try {
    const sheets = getSheetsClient();
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: LOGS_SHEET_ID,
      range: 'A2:E'
    });
    const rows = resp.data.values || [];
    return rows.map(r => ({
      user: r[0] || '',
      loginTime: r[1] || '',
      logoutTime: r[2] || '',
      duration: r[3] || '',
      area: r[4] || ''
    })).filter(r => r.user);
  } catch (e) {
    console.error('[Logs] Read error:', e.message);
    return [];
  }
}

// Clears all log rows (keeps header)
async function clearLogs() {
  if (!LOGS_SHEET_ID) return false;
  try {
    const sheets = getSheetsClient();
    await sheets.spreadsheets.values.clear({
      spreadsheetId: LOGS_SHEET_ID,
      range: 'A2:E'
    });
    return true;
  } catch (e) {
    console.error('[Logs] Clear error:', e.message);
    return false;
  }
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
    if (!row || row.length === 0) continue;
    const region = (row[0] || '').toString().trim();
    const area = (row[1] || '').toString().trim();
    let storeId = row[2];
    if (storeId === null || storeId === undefined) storeId = '';
    storeId = storeId.toString().trim();
    const storeName = (row[3] || '').toString().trim();
    const remarks = (row[4] || '').toString().trim();
    if (storeId) {
      const info = { region, area, storeName, remarks };
      storeMap[storeId] = info;
      // Also store without leading zeros and with padded zeros for fuzzy match
      const numStoreId = parseInt(storeId).toString();
      if (numStoreId !== storeId && numStoreId !== 'NaN') storeMap[numStoreId] = info;
    }
  }
  console.log('[Stores] Loaded ' + Object.keys(storeMap).length + ' store keys');
  // Log a sample
  const keys = Object.keys(storeMap).slice(0, 5);
  keys.forEach(k => console.log('[Stores] Sample: ' + k + ' -> ' + JSON.stringify(storeMap[k])));
  return storeMap;
}

// ─── PARSE USERS SHEET FROM XLSX ─────────────────────────────────────────────
function parseUsersXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  // Find the "Users" sheet (case-insensitive)
  const usersSheetName = wb.SheetNames.find(n => n.toLowerCase().trim() === 'users');
  if (!usersSheetName) {
    console.warn('[Users] No "Users" sheet found. Available sheets: ' + wb.SheetNames.join(', '));
    return {};
  }
  const ws = wb.Sheets[usersSheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const users = {};
  // Skip header row (row 0), data starts row 1
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    const username = (row[0] || '').toString().trim();
    const password = (row[1] || '').toString().trim();
    const level = (row[2] || '').toString().trim().toLowerCase();
    const area = (row[3] || '').toString().trim();
    if (username) {
      users[username.toLowerCase()] = { username, password, level, area };
    }
  }
  console.log('[Users] Loaded ' + Object.keys(users).length + ' users from "' + usersSheetName + '" sheet');
  return users;
}

// ─── SAFE NUMBER ─────────────────────────────────────────────────────────────
function num(val) {
  const n = parseFloat((val || '').toString().replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

// ─── DATE PARSER ─────────────────────────────────────────────────────────────
function parseDate(val) {
  if (!val) return null;
  const s = val.toString().trim();
  if (!s || s === '0' || s === '00000000') return null;
  // Try common formats: MM/DD/YYYY, YYYY-MM-DD, M/D/YYYY, YYYYMMDD
  let d = new Date(s);
  if (!isNaN(d.getTime()) && d.getFullYear() > 2000 && d.getFullYear() < 2100) return d;
  // Try YYYYMMDD
  if (/^\d{8}$/.test(s)) {
    const y = parseInt(s.substr(0, 4));
    const m = parseInt(s.substr(4, 2)) - 1;
    const dy = parseInt(s.substr(6, 2));
    d = new Date(y, m, dy);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function formatDate(d) {
  if (!d) return '';
  if (typeof d === 'string') d = parseDate(d);
  if (!d) return '';
  const m = (d.getMonth() + 1).toString().padStart(2, '0');
  const dy = d.getDate().toString().padStart(2, '0');
  return m + '/' + dy + '/' + d.getFullYear();
}

function daysSince(d) {
  if (!d) return null;
  if (typeof d === 'string') d = parseDate(d);
  if (!d) return null;
  const ms = Date.now() - d.getTime();
  return Math.floor(ms / (1000 * 60 * 60 * 24));
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
    dateLastReceived: 66, // BO
    qtyLastReceived: 67,  // BP
    dateLastOrdered: 68,  // BQ
    setCode: 69,          // BR
    dateLastAdjusted: 70, // BS
    dateLastSold: 71,     // BT
    lastXfer1: 72,        // BU
    lastXfer2: 73,        // BV
    replenishment: 74 // BW
  };

  // PRE-PASS: Build SKU price lookup (avg cost by SKU code, from rows that have stock/cost)
  const skuCostLookup = {};
  for (const row of dataRows) {
    if (!row || row.length < 10) continue;
    const skuCode = (row[COL.skuCode] || '').toString().trim();
    if (!skuCode) continue;
    const cost = num(row[COL.avgCost]);
    if (cost > 0 && !skuCostLookup[skuCode]) {
      skuCostLookup[skuCode] = cost;
    }
  }

  // Debug: sample of InvData store IDs
  const sampleStoreIds = new Set();
  for (let i = 0; i < Math.min(50, dataRows.length); i++) {
    const row = dataRows[i];
    if (row && row[2]) sampleStoreIds.add(row[2].toString().trim());
    if (sampleStoreIds.size >= 10) break;
  }
  console.log('[InvData] Sample store IDs from CSV:', [...sampleStoreIds].join(', '));
  console.log('[Match] Testing lookups:');
  [...sampleStoreIds].slice(0, 5).forEach(sid => {
    const info = storeMap[sid] || storeMap[parseInt(sid).toString()];
    console.log('  Store "' + sid + '" -> ' + (info ? ('Area: ' + info.area + ' | Name: ' + info.storeName) : 'NOT FOUND'));
  });

  // Enrich rows and map store info
  // GLOBAL FILTER: Only include rows with STS Number (column BN)
  const enriched = [];
  let skippedNoSTS = 0;
  for (const row of dataRows) {
    if (!row || row.length < 10) continue;
    // STS Number filter - skip blank/empty STS rows
    const stsNumber = (row[COL.stsNumber] || '').toString().trim();
    if (!stsNumber) { skippedNoSTS++; continue; }
    const storeIdRaw = (row[COL.storeNumber] || '').toString().trim();
    const storeId = parseInt(storeIdRaw).toString();
    const storeInfo = storeMap[storeIdRaw] || storeMap[storeId] || {};
    const wtsNet = num(row[COL.wtsNet]);
    const onHand = num(row[COL.onHand]);
    const onHandValue = num(row[COL.onHandValue]);
    const p8ave = num(row[COL.p8aveGross]);
    const wkAveNet = num(row[COL.wkAveNet]);
    const currentWkSales = num(row[COL.currentWkSales]);
    const totalPO = num(row[COL.totalPONet]);
    const poValue = num(row[COL.poValue]);
    const trfValue = num(row[COL.trfValue]);

    const isCritical = wtsNet > 0 && wtsNet < 2 && onHand > 0;
    const isOverstock = wtsNet > 12 && onHand > 0;
    const isDeadStock = onHand > 0 && p8ave === 0 && currentWkSales === 0;
    const isZeroStock = onHand === 0;
    const skuCode = (row[COL.skuCode] || '').toString().trim();
    // Get avg cost from this row, fallback to SKU lookup from other stores
    let avgCost = num(row[COL.avgCost]);
    if (avgCost === 0 && skuCode && skuCostLookup[skuCode]) {
      avgCost = skuCostLookup[skuCode];
    }
    // Out of Stock = no stock + was selling = LOST SALES
    const isOutOfStock = onHand === 0 && p8ave > 0;
    const lostSalesPerWeek = isOutOfStock ? p8ave * avgCost : 0;
    const dateLastSold = row[COL.dateLastSold] || '';
    const dateLastReceived = row[COL.dateLastReceived] || '';
    const daysNoSales = daysSince(dateLastSold);

    enriched.push({
      regionCode: row[COL.regionCode] || '',
      regionName: storeInfo.region || row[COL.regionName] || '',
      storeNumber: storeId,
      storeName: storeInfo.storeName || row[COL.storeName] || '',
      area: storeInfo.area || '',
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
      avgCost,
      stdPack: num(row[COL.stdPack]),
      // Pre-computed for sorting (numeric, null when "Per Piece")
      qtyCasesNum: (num(row[COL.stdPack]) > 0 && num(row[COL.stdPack]) !== onHand) ? (onHand / num(row[COL.stdPack])) : null,
      ico: (row[COL.ico] || '').toString().trim(),
      totalPO,
      poValue,
      trfValue,
      xdockValue: num(row[COL.xdockValue]),
      poOrderGR: num(row[COL.poOrderGR]),
      trfOrderGR: num(row[COL.trfOrderGR]),
      currentWkSales,
      p8ave,
      wkAveNet,
      wtsNet: wtsNet === 0 && onHand > 0 && p8ave === 0 ? 999 : wtsNet,
      wtsGross: num(row[COL.wtsGross]),
      wtsAfterDeliv: num(row[COL.wtsAfterDeliv]),
      // Per-SKU metrics for SKU Analysis tab
      skuWTS: p8ave > 0 ? onHand / p8ave : null,
      skuDaysCover: (wkAveNet > 0 && avgCost > 0) ? (onHandValue * 7) / (wkAveNet * avgCost) : null,
      supplierCode: row[COL.supplierCode] || '',
      supplierName: row[COL.supplierName] || '',
      delivMode: row[COL.delivMode] || '',
      dateLastSold,
      dateLastReceived,
      daysNoSales,
      lostSalesPerWeek,
      isCritical,
      isOverstock,
      isDeadStock,
      isZeroStock,
      isOutOfStock
    });
  }
  console.log('[Filter] STS Number filter: ' + enriched.length + ' rows kept, ' + skippedNoSTS + ' rows skipped (no STS Number)');

  // ── KPIs ──────────────────────────────────────────────────────────────────
  const totalOnHandValue = enriched.reduce((s, r) => s + r.onHandValue, 0);
  const totalOnHand = enriched.reduce((s, r) => s + r.onHand, 0);
  const criticalCount = enriched.filter(r => r.isCritical).length;
  const overstockCount = enriched.filter(r => r.isOverstock).length;
  const deadStockCount = enriched.filter(r => r.isDeadStock).length;
  const outOfStockCount = enriched.filter(r => r.isOutOfStock).length;
  const totalLostSalesPerWeek = enriched.reduce((s, r) => s + r.lostSalesPerWeek, 0);
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
    outOfStockCount,
    totalLostSalesPerWeek,
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
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
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
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
    }));

  // ── DEAD STOCK ────────────────────────────────────────────────────────────
  const deadStockItems = enriched
    .filter(r => r.isDeadStock)
    .sort((a, b) => b.onHandValue - a.onHandValue)
    .slice(0, 300)
    .map(r => {
      const wtsItem = r.p8ave > 0 ? r.onHand / r.p8ave : null;
      const dcItem = (r.wkAveNet > 0 && r.avgCost > 0) ? (r.onHandValue * 7) / (r.wkAveNet * r.avgCost) : null;
      return {
        store: `${r.storeNumber} - ${r.storeName}`,
        area: r.area,
        skuCode: r.skuCode,
        skuDesc: r.skuDesc,
        supplier: r.supplierName,
        onHand: r.onHand,
        onHandValue: r.onHandValue,
        weeksToSell: wtsItem,
        daysCover: dcItem,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        action: 'No Sales 8 Wks - Review/Markdown'
      };
    });

  // ── OUT OF STOCK ITEMS (Lost Sales) ───────────────────────────────────────
  const outOfStockItems = enriched
    .filter(r => r.isOutOfStock)
    .sort((a, b) => b.lostSalesPerWeek - a.lostSalesPerWeek)
    .slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      p8ave: r.p8ave,
      avgCost: r.avgCost,
      lostSalesPerWeek: r.lostSalesPerWeek,
      totalPO: r.totalPO,
      daysNoSales: r.daysNoSales != null ? r.daysNoSales : '',
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      action: r.totalPO > 0 ? 'PO Incoming' : 'URGENT: Place PO Now'
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
        criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0,
        totalSKUs: 0, totalSales: 0, totalLostSales: 0,
        totalWklSalesValue: 0  // sum of AU × AX (weekly sales value, net wholesale)
      };
    }
    const g = storeGroups[key];
    g.totalValue += r.onHandValue;
    g.totalOnHand += r.onHand;
    g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    g.totalLostSales += r.lostSalesPerWeek;
    g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock) g.oosCount++;
  }
  // Compute risk percentages and days cover
  // Days Cover = OnHand Value / (Weekly Sales Net WS × Avg Cost / 7) = BC / (AU × AX / 7)
  const storeAnalysis = Object.values(storeGroups).map(g => {
    const total = g.totalSKUs || 1;
    g.criticalPct = (g.criticalCount / total) * 100;
    g.oosPct = (g.oosCount / total) * 100;
    g.overstockPct = (g.overstockCount / total) * 100;
    g.deadPct = (g.deadCount / total) * 100;
    g.daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
    g.weeksToSell = g.daysCover != null ? g.daysCover / 7 : null;
    return g;
  }).sort((a, b) => b.totalValue - a.totalValue);

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
        criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0,
        totalSKUs: 0, totalSales: 0, totalLostSales: 0,
        totalWklSalesValue: 0
      };
    }
    const g = supplierGroups[key];
    g.totalValue += r.onHandValue;
    g.totalOnHand += r.onHand;
    g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    g.totalLostSales += r.lostSalesPerWeek;
    g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock) g.oosCount++;
  }
  const supplierAnalysis = Object.values(supplierGroups).map(g => {
    const total = g.totalSKUs || 1;
    g.criticalPct = (g.criticalCount / total) * 100;
    g.oosPct = (g.oosCount / total) * 100;
    g.overstockPct = (g.overstockCount / total) * 100;
    g.deadPct = (g.deadCount / total) * 100;
    g.daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
    g.weeksToSell = g.daysCover != null ? g.daysCover / 7 : null;
    return g;
  }).sort((a, b) => b.totalValue - a.totalValue).slice(0, 100);

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

  return { kpis, criticalItems, overstockItems, deadStockItems, outOfStockItems, storeAnalysis, supplierAnalysis, filterMeta, rows: enriched };
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
    let usersMap = {};
    try {
      console.log('[Cache] Looking for ' + STORES_FILE_NAME + ' in folder ' + GDRIVE_FOLDER_ID);
      const storesFile = await findFile(drive, STORES_FILE_NAME);
      if (storesFile) {
        console.log('[Cache] Found stores file ID: ' + storesFile.id + ', downloading...');
        const storesBuffer = await downloadFileBuffer(drive, storesFile.id);
        storeMap = parseStoresXLSX(storesBuffer);
        usersMap = parseUsersXLSX(storesBuffer);
        console.log(`[Cache] Loaded ${Object.keys(storeMap).length} stores, ${Object.keys(usersMap).length} users.`);
      } else {
        console.warn('[Cache] STORES FILE NOT FOUND! Searched name: "' + STORES_FILE_NAME + '" in folder: ' + GDRIVE_FOLDER_ID);
        // List all files in folder for debugging
        const allFiles = await drive.files.list({
          q: `'${GDRIVE_FOLDER_ID}' in parents and trashed=false`,
          fields: 'files(id,name,mimeType)',
          pageSize: 20
        });
        console.warn('[Cache] Files in folder:');
        (allFiles.data.files || []).forEach(f => console.warn('  - "' + f.name + '" (type: ' + f.mimeType + ')'));
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
    if (Object.keys(usersMap).length > 0) cache.users = usersMap;
    cache.kpis = analytics.kpis;
    cache.criticalItems = analytics.criticalItems;
    cache.overstockItems = analytics.overstockItems;
    cache.deadStockItems = analytics.deadStockItems;
    cache.outOfStockItems = analytics.outOfStockItems;
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

// Resolve filters with session area-lock enforced. Non-admin users are forced to their area.
function resolveFilters(req) {
  const filters = { ...req.query };
  delete filters.token;
  const token = req.query.token || (req.headers['x-auth-token']) || '';
  const s = sessions[token];
  if (s && !s.isAdmin && s.area) {
    // Force area lock - override any area filter the client sent
    filters.area = s.area;
  }
  return filters;
}

// ─── API ROUTES ───────────────────────────────────────────────────────────────

// ─── SESSION / AUTH (Simple) ──────────────────────────────────────────────────
const sessions = {}; // token -> { username, level, area, created }

function makeToken() {
  return crypto.randomBytes(24).toString('hex');
}

app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'Username and password required' });
  if (Object.keys(cache.users).length === 0) {
    return res.status(503).json({ error: 'User data not loaded yet. Please try again in a moment.' });
  }
  const user = cache.users[username.toLowerCase().trim()];
  if (!user || user.password !== password) {
    return res.status(401).json({ error: 'Invalid username or password' });
  }
  const token = makeToken();
  const isAdmin = user.level === 'admin' || (user.area || '').toLowerCase() === 'all';
  const areaLabel = isAdmin ? 'All' : user.area;
  sessions[token] = {
    username: user.username,
    level: user.level,
    area: isAdmin ? '' : user.area,  // empty = all access
    isAdmin,
    created: Date.now()
  };
  // Log the login event (async, don't block response)
  logLoginEvent(user.username, areaLabel).then(result => {
    if (result && sessions[token]) {
      sessions[token].logRow = result.rowNum;
      sessions[token].loginTimeISO = result.loginTime;
    }
  });
  res.json({
    token,
    username: user.username,
    level: user.level,
    area: isAdmin ? '' : user.area,
    isAdmin
  });
});

app.post('/api/logout', (req, res) => {
  const token = (req.body && req.body.token) || '';
  const s = sessions[token];
  if (s) {
    if (s.logRow && s.loginTimeISO) {
      logLogoutEvent(s.logRow, s.loginTimeISO);
    }
    delete sessions[token];
  }
  res.json({ ok: true });
});

app.get('/api/me', (req, res) => {
  const token = req.query.token || '';
  const s = sessions[token];
  if (!s) return res.status(401).json({ error: 'Not logged in' });
  res.json({ username: s.username, level: s.level, area: s.area, isAdmin: s.isAdmin });
});

// Activity logs — admin only
function requireAdmin(req, res) {
  const token = req.query.token || (req.body && req.body.token) || '';
  const s = sessions[token];
  if (!s || !s.isAdmin) { res.status(403).json({ error: 'Admin access required' }); return null; }
  return s;
}

app.get('/api/logs', async (req, res) => {
  if (!requireAdmin(req, res)) return;
  const logs = await readLogs();
  // Most recent first
  logs.reverse();
  res.json(logs);
});

app.post('/api/logs/clear', async (req, res) => {
  if (!requireAdmin(req, res)) return;
  const ok = await clearLogs();
  res.json({ ok });
});

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
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.kpis);
  const filtered = applyFilters(cache.rows, filters);
  const totalOnHandValue = filtered.reduce((s, r) => s + r.onHandValue, 0);
  const totalOnHand = filtered.reduce((s, r) => s + r.onHand, 0);
  const criticalCount = filtered.filter(r => r.isCritical).length;
  const overstockCount = filtered.filter(r => r.isOverstock).length;
  const deadStockCount = filtered.filter(r => r.isDeadStock).length;
  const outOfStockCount = filtered.filter(r => r.isOutOfStock).length;
  const totalLostSalesPerWeek = filtered.reduce((s, r) => s + r.lostSalesPerWeek, 0);
  const validWts = filtered.filter(r => r.wtsNet > 0 && r.wtsNet < 999);
  const avgWts = validWts.length > 0 ? validWts.reduce((s, r) => s + r.wtsNet, 0) / validWts.length : 0;
  res.json({
    totalOnHandValue, totalOnHand, criticalCount, overstockCount, deadStockCount,
    outOfStockCount, totalLostSalesPerWeek,
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
  // Determine area scope: locked area for non-admin, or area query param
  const token = req.query.token || '';
  const s = sessions[token];
  let area = req.query.area || '';
  if (s && !s.isAdmin && s.area) area = s.area; // force locked area

  if (!area) return res.json(cache.filterMeta);

  // Return stores filtered to the selected area
  const storesInArea = cache.filterMeta.stores.filter(st => st.area === area);
  res.json({ ...cache.filterMeta, stores: storesInArea });
});

app.get('/api/critical', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.criticalItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isCritical)
    .sort((a, b) => a.wtsNet - b.wtsNet).slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue,
      currentWkSales: r.currentWkSales, p8ave: r.p8ave,
      wtsNet: r.wtsNet, totalPO: r.totalPO,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      action: r.totalPO > 0 ? 'PO Incoming' : r.p8ave > 0 ? 'URGENT: Place PO' : 'Review'
    }));
  res.json(filtered);
});

app.get('/api/overstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.overstockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isOverstock)
    .sort((a, b) => b.wtsNet - a.wtsNet).slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue, p8ave: r.p8ave,
      wtsNet: r.wtsNet === 999 ? 'Dead Stock' : r.wtsNet.toFixed(1),
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
    }));
  res.json(filtered);
});

app.get('/api/deadstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.deadStockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isDeadStock)
    .sort((a, b) => b.onHandValue - a.onHandValue).slice(0, 300)
    .map(r => {
      const wtsItem = r.p8ave > 0 ? r.onHand / r.p8ave : null;
      const dcItem = (r.wkAveNet > 0 && r.avgCost > 0) ? (r.onHandValue * 7) / (r.wkAveNet * r.avgCost) : null;
      return {
        store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
        skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
        onHand: r.onHand, onHandValue: r.onHandValue,
        weeksToSell: wtsItem,
        daysCover: dcItem,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        action: 'No Sales 8 Wks - Review/Markdown'
      };
    });
  res.json(filtered);
});

app.get('/api/outofstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.outOfStockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isOutOfStock)
    .sort((a, b) => b.lostSalesPerWeek - a.lostSalesPerWeek).slice(0, 500)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      p8ave: r.p8ave, avgCost: r.avgCost,
      lostSalesPerWeek: r.lostSalesPerWeek, totalPO: r.totalPO,
      daysNoSales: r.daysNoSales != null ? r.daysNoSales : '',
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      action: r.totalPO > 0 ? 'PO Incoming' : 'URGENT: Place PO Now'
    }));
  res.json(filtered);
});

// SKU Analysis endpoint — server-side pagination/sorting/searching
app.get('/api/skus', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const { page = '1', pageSize = '100', sortBy = '', sortDir = 'asc', search = '', status = '', token, ...filters } = req.query;
  // Enforce area lock for non-admin users
  const s = sessions[token || ''];
  if (s && !s.isAdmin && s.area) filters.area = s.area;
  let rows = applyFilters(cache.rows, filters);

  // Status filter
  if (status === 'critical') rows = rows.filter(r => r.isCritical);
  else if (status === 'oos') rows = rows.filter(r => r.isOutOfStock);
  else if (status === 'overstock') rows = rows.filter(r => r.isOverstock);
  else if (status === 'deadstock') rows = rows.filter(r => r.isDeadStock);
  else if (status === 'normal') rows = rows.filter(r => !r.isCritical && !r.isOutOfStock && !r.isOverstock && !r.isDeadStock);

  // Search across SKU code + description + supplier + store name
  if (search) {
    const q = search.toLowerCase();
    rows = rows.filter(r =>
      (r.skuCode || '').toLowerCase().includes(q) ||
      (r.skuDesc || '').toLowerCase().includes(q) ||
      (r.supplierName || '').toLowerCase().includes(q) ||
      (r.storeName || '').toLowerCase().includes(q)
    );
  }

  // Sort
  if (sortBy) {
    // Map frontend sort keys to actual raw row fields where they differ
    const sortFieldMap = { invValue: 'onHandValue', weeksToSell: 'skuWTS', daysCover: 'skuDaysCover' };
    const field = sortFieldMap[sortBy] || sortBy;
    const dir = sortDir === 'desc' ? -1 : 1;
    rows = [...rows].sort((a, b) => {
      let av = a[field], bv = b[field];
      // Nulls always at bottom regardless of direction
      const aNull = (av == null || av === '');
      const bNull = (bv == null || bv === '');
      if (aNull && bNull) return 0;
      if (aNull) return 1;
      if (bNull) return -1;
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv)) * dir;
    });
  }

  const total = rows.length;
  const p = Math.max(1, parseInt(page));
  const ps = Math.min(500, Math.max(10, parseInt(pageSize)));
  const start = (p - 1) * ps;
  const pageRows = rows.slice(start, start + ps).map(r => {
    let qtyCases;
    if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
    else if (r.stdPack > 0 && r.onHand > 0) qtyCases = (r.onHand / r.stdPack).toFixed(2);
    else if (r.onHand === 0) qtyCases = '0';
    else qtyCases = 'Per Piece';
    return {
      storeNumber: r.storeNumber,
      storeName: r.storeName,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplierName: r.supplierName,
      onHand: r.onHand,
      qtyCases,
      qtyCasesNum: r.qtyCasesNum,
      stdPack: r.stdPack,
      invValue: r.onHandValue,
      weeksToSell: r.skuWTS,
      skuWTS: r.skuWTS,
      daysCover: r.skuDaysCover,
      skuDaysCover: r.skuDaysCover,
      p8ave: r.p8ave,
      status: r.isCritical ? 'Critical' : r.isOutOfStock ? 'OOS' : r.isOverstock ? 'Overstock' : r.isDeadStock ? 'Dead Stock' : 'Normal',
      lostSalesPerWeek: r.lostSalesPerWeek,
      ico: r.ico,
      poOrderGR: r.poOrderGR,
      trfOrderGR: r.trfOrderGR,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived)
    };
  });

  res.json({ total, page: p, pageSize: ps, totalPages: Math.ceil(total / ps), rows: pageRows });
});

app.get('/api/stores', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.storeAnalysis);
  const filtered = applyFilters(cache.rows, filters);
  const storeGroups = {};
  for (const r of filtered) {
    const key = r.storeNumber;
    if (!storeGroups[key]) storeGroups[key] = { storeNumber: r.storeNumber, storeName: r.storeName, area: r.area, region: r.regionName, totalValue: 0, totalOnHand: 0, criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0, totalSKUs: 0, totalSales: 0, totalLostSales: 0, totalWklSalesValue: 0 };
    const g = storeGroups[key];
    g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    g.totalLostSales += r.lostSalesPerWeek;
    g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock) g.oosCount++;
  }
  const result = Object.values(storeGroups).map(g => {
    const total = g.totalSKUs || 1;
    g.criticalPct = (g.criticalCount / total) * 100;
    g.oosPct = (g.oosCount / total) * 100;
    g.overstockPct = (g.overstockCount / total) * 100;
    g.deadPct = (g.deadCount / total) * 100;
    g.daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
    g.weeksToSell = g.daysCover != null ? g.daysCover / 7 : null;
    return g;
  }).sort((a, b) => b.totalValue - a.totalValue);
  res.json(result);
});

app.get('/api/suppliers', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.supplierAnalysis);
  const filtered = applyFilters(cache.rows, filters);
  const supplierGroups = {};
  for (const r of filtered) {
    if (!r.supplierCode) continue;
    const key = r.supplierCode;
    if (!supplierGroups[key]) supplierGroups[key] = { supplierCode: r.supplierCode, supplierName: r.supplierName, totalValue: 0, totalOnHand: 0, criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0, totalSKUs: 0, totalSales: 0, totalLostSales: 0, totalWklSalesValue: 0 };
    const g = supplierGroups[key];
    g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    g.totalLostSales += r.lostSalesPerWeek;
    g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
    if (r.isCritical) g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock) g.oosCount++;
  }
  const result = Object.values(supplierGroups).map(g => {
    const total = g.totalSKUs || 1;
    g.criticalPct = (g.criticalCount / total) * 100;
    g.oosPct = (g.oosCount / total) * 100;
    g.overstockPct = (g.overstockCount / total) * 100;
    g.deadPct = (g.deadCount / total) * 100;
    g.daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
    g.weeksToSell = g.daysCover != null ? g.daysCover / 7 : null;
    return g;
  }).sort((a, b) => b.totalValue - a.totalValue).slice(0, 100);
  res.json(result);
});

app.post('/api/refresh', async (req, res) => {
  refreshData(true);
  res.json({ message: 'Refresh triggered' });
});

// ─── EXPORT CSV ───────────────────────────────────────────────────────────────
app.get('/api/export/:type', (req, res) => {
  if (!cache.ready) return res.status(503).send('Cache not ready');
  const type = req.params.type;
  // Enforce area lock for non-admin
  const token = req.query.token || '';
  const s = sessions[token];
  const lockedArea = (s && !s.isAdmin && s.area) ? s.area : null;

  let data;
  if (lockedArea) {
    // Recompute from filtered rows so users only export their area
    const rows = applyFilters(cache.rows, { area: lockedArea });
    if (type === 'critical') data = rows.filter(r => r.isCritical).map(r => ({ store:`${r.storeNumber} - ${r.storeName}`, area:r.area, skuCode:r.skuCode, skuDesc:r.skuDesc, supplier:r.supplierName, onHand:r.onHand, onHandValue:r.onHandValue, wtsNet:r.wtsNet }));
    else if (type === 'overstock') data = rows.filter(r => r.isOverstock).map(r => ({ store:`${r.storeNumber} - ${r.storeName}`, area:r.area, skuCode:r.skuCode, skuDesc:r.skuDesc, supplier:r.supplierName, onHand:r.onHand, onHandValue:r.onHandValue }));
    else if (type === 'deadstock') data = rows.filter(r => r.isDeadStock).map(r => ({ store:`${r.storeNumber} - ${r.storeName}`, area:r.area, skuCode:r.skuCode, skuDesc:r.skuDesc, supplier:r.supplierName, onHand:r.onHand, onHandValue:r.onHandValue }));
    else if (type === 'outofstock') data = rows.filter(r => r.isOutOfStock).map(r => ({ store:`${r.storeNumber} - ${r.storeName}`, area:r.area, skuCode:r.skuCode, skuDesc:r.skuDesc, supplier:r.supplierName, p8ave:r.p8ave, lostSalesPerWeek:r.lostSalesPerWeek }));
    else data = [];
  } else {
    const dataMap = { critical: cache.criticalItems, overstock: cache.overstockItems, deadstock: cache.deadStockItems, outofstock: cache.outOfStockItems, stores: cache.storeAnalysis, suppliers: cache.supplierAnalysis };
    data = dataMap[type];
  }
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
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
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

/* LOGIN SCREEN */
#login-screen {
  position:fixed; inset:0; background:var(--bg); z-index:10000;
  display:flex; align-items:center; justify-content:center;
}
.login-box {
  background:var(--bg2); border:1px solid var(--border); border-radius:12px;
  padding:32px; width:340px; max-width:90vw;
  display:flex; flex-direction:column; gap:14px;
  box-shadow:0 8px 40px rgba(0,0,0,0.5);
}
.login-logo { display:flex; align-items:center; gap:10px; justify-content:center; }
.login-logo span { font-family:'IBM Plex Mono',monospace; font-size:16px; font-weight:700; color:var(--text); letter-spacing:0.5px; }
.login-subtitle { text-align:center; font-size:12px; color:var(--text2); margin-bottom:6px; }
.login-input {
  width:100%; padding:10px 12px; border-radius:var(--radius);
  border:1px solid var(--border); background:var(--bg3); color:var(--text);
  font-size:13px; font-family:'IBM Plex Sans',sans-serif;
}
.login-input:focus { outline:none; border-color:var(--green-bright); }
.login-btn {
  width:100%; padding:11px; border-radius:var(--radius); border:none;
  background:var(--green); color:#fff; font-size:14px; font-weight:600; cursor:pointer;
  font-family:'IBM Plex Sans',sans-serif; margin-top:4px; transition:background 0.2s;
}
.login-btn:hover { background:var(--green-light); }
.login-btn:disabled { opacity:0.6; cursor:not-allowed; }
.login-error { color:var(--red-light); font-size:12px; min-height:16px; text-align:center; }
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

/* SIDEBAR (deprecated - using top filter bar) */
.filter-select {
  width:100%; padding:6px 8px; border-radius:var(--radius);
  border:1px solid var(--border); background:var(--bg3); color:var(--text);
  font-size:12px; font-family:'IBM Plex Sans',sans-serif; cursor:pointer;
}
.filter-select:focus { outline:none; border-color:var(--green-bright); }

/* Top Filter Bar */
.filter-bar {
  background: var(--bg2);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 10px 14px;
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  align-items: flex-end;
}
.fb-group { display: flex; flex-direction: column; gap: 4px; flex: 1; min-width: 140px; max-width: 220px; }
.fb-group label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.8px; color: var(--text2); font-weight: 600; }
.fb-group .filter-select { width: 100%; }
.fb-actions { display: flex; align-items: center; gap: 10px; margin-left: auto; flex-wrap: wrap; }
.fb-actions .active-filters { display: flex; flex-wrap: wrap; gap: 4px; max-width: 400px; }

/* Searchable Dropdown */
.search-dropdown { position:relative; }
.sd-trigger {
  width:100%; padding:6px 8px; border-radius:var(--radius);
  border:1px solid var(--border); background:var(--bg3); color:var(--text);
  font-size:12px; font-family:'IBM Plex Sans',sans-serif; cursor:pointer;
  display:flex; align-items:center; justify-content:space-between; gap:4px;
  text-align:left;
}
.sd-trigger:hover { border-color:var(--green-bright); }
.sd-trigger .sd-arrow { font-size:10px; color:var(--text2); }
.sd-trigger.has-value { border-color:var(--green-bright); color:var(--green-bright); }
.sd-panel {
  position:absolute; top:100%; left:0; right:0; z-index:50;
  background:var(--bg3); border:1px solid var(--border); border-radius:var(--radius);
  margin-top:2px; box-shadow:0 6px 18px rgba(0,0,0,0.5);
  display:none; max-height:280px; overflow:hidden;
  display:flex; flex-direction:column;
}
.search-dropdown.open .sd-panel { display:flex; }
.sd-search {
  width:100%; padding:6px 8px; border:none; border-bottom:1px solid var(--border);
  background:var(--bg2); color:var(--text); font-size:12px; outline:none;
  font-family:'IBM Plex Sans',sans-serif;
}
.sd-options { overflow-y:auto; max-height:240px; }
.sd-option {
  padding:6px 10px; font-size:12px; cursor:pointer; color:var(--text);
  border-bottom:1px solid rgba(48,54,61,0.4);
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}
.sd-option:hover { background:var(--green-dim); color:var(--green-bright); }
.sd-option.selected { background:var(--green); color:#fff; }
.sd-option.empty-state { color:var(--text2); font-style:italic; cursor:default; }
.sd-option.empty-state:hover { background:transparent; color:var(--text2); }
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

/* RISK PILLS */
.risk-pill {
  display:inline-flex; align-items:center; gap:4px;
  padding:3px 8px; border-radius:12px; font-size:11px; font-weight:600;
  font-family:'IBM Plex Mono',monospace; min-width:60px; justify-content:center;
}
.risk-low { background:rgba(63,185,80,0.15); color:#3fb950; border:1px solid rgba(63,185,80,0.3); }
.risk-med { background:rgba(227,179,65,0.15); color:#e3b341; border:1px solid rgba(227,179,65,0.3); }
.risk-high { background:rgba(248,81,73,0.15); color:#f85149; border:1px solid rgba(248,81,73,0.3); }
.risk-none { background:var(--bg3); color:var(--text2); }

/* Sort indicator */
.sort-ind { font-size:9px; margin-left:2px; opacity:0.5; }
.sort-ind.asc::after { content:'▲'; opacity:1; }
.sort-ind.desc::after { content:'▼'; opacity:1; }

/* Status badges in SKU table */
.status-critical { background:rgba(248,81,73,0.15); color:#f85149; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; }
.status-oos { background:rgba(248,81,73,0.25); color:#f85149; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; }
.status-overstock { background:rgba(227,179,65,0.15); color:#e3b341; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; }
.status-dead { background:rgba(139,148,158,0.2); color:#8b949e; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; }
.status-normal { background:rgba(63,185,80,0.15); color:#3fb950; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; }

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

<div id="login-screen" style="display:none;">
  <div class="login-box">
    <div class="login-logo">
      <div class="dot" style="width:12px;height:12px;border-radius:50%;background:#3fb950;box-shadow:0 0 10px #3fb950;"></div>
      <span>CAMANAVA INVENTORY</span>
    </div>
    <div class="login-subtitle">Sign in to continue</div>
    <input type="text" id="login-username" class="login-input" placeholder="Username" autocomplete="username"/>
    <input type="password" id="login-password" class="login-input" placeholder="Password" autocomplete="current-password"/>
    <div class="login-error" id="login-error"></div>
    <button class="login-btn" id="login-btn" onclick="doLogin()">Sign In</button>
  </div>
</div>

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
    <span id="user-info" style="font-size:11px;color:var(--text2);font-family:'IBM Plex Mono',monospace;"></span>
    <div class="refresh-info" id="refresh-info">–</div>
    <button class="btn btn-green" onclick="triggerRefresh()">↺ Refresh</button>
    <button class="btn" onclick="doLogout()">Logout</button>
  </div>
</header>

<div class="main">

  <!-- SIDEBAR FILTERS -->
  <!-- MAIN CONTENT (full width now) -->
  <main class="content">

    <!-- TOP FILTER BAR -->
    <div class="filter-bar">
      <div class="fb-group">
        <label>Area</label>
        <select class="filter-select" id="f-area" onchange="applyFilter('area',this.value)">
          <option value="">All Areas</option>
        </select>
      </div>
      <div class="fb-group">
        <label>Store</label>
        <select class="filter-select" id="f-store" onchange="applyFilter('store',this.value)">
          <option value="">All Stores</option>
        </select>
      </div>
      <div class="fb-group">
        <label>Department</label>
        <select class="filter-select" id="f-dept" onchange="applyFilter('dept',this.value)">
          <option value="">All Departments</option>
        </select>
      </div>
      <div class="fb-group">
        <label>Sub-Department</label>
        <select class="filter-select" id="f-subdept" onchange="applyFilter('subDept',this.value)">
          <option value="">All Sub-Depts</option>
        </select>
      </div>
      <div class="fb-group">
        <label>Class</label>
        <select class="filter-select" id="f-cls" onchange="applyFilter('cls',this.value)">
          <option value="">All Classes</option>
        </select>
      </div>
      <div class="fb-group">
        <label>Supplier</label>
        <select class="filter-select" id="f-supplier" onchange="applyFilter('supplier',this.value)">
          <option value="">All Suppliers</option>
        </select>
      </div>
      <div class="fb-actions">
        <div class="active-filters" id="active-filters"></div>
        <button class="btn btn-sm" onclick="clearFilters()">✕ Clear</button>
      </div>
    </div>

    <!-- STATUS BAR -->
    <div class="status-bar" id="status-bar">
      <div class="status-dot green" id="status-dot"></div>
      <span id="status-text">Ready</span>
      <span style="margin-left:auto;" id="status-rows">–</span>
    </div>

    <!-- TABS -->
    <div class="tabs">
      <div class="tab active" onclick="showTab('overview')">Overview</div>
      <div class="tab" onclick="showTab('outofstock')">🚫 Out of Stock</div>
      <div class="tab" onclick="showTab('critical')">⚠ Critical</div>
      <div class="tab" onclick="showTab('overstock')">📦 Overstock</div>
      <div class="tab" onclick="showTab('deadstock')">💀 Dead Stock</div>
      <div class="tab" onclick="showTab('stores')">🏪 Stores</div>
      <div class="tab" onclick="showTab('suppliers')">🏭 Suppliers</div>
      <div class="tab" onclick="showTab('skus')">🔍 SKU Analysis</div>
      <div class="tab" id="tab-btn-logs" onclick="showTab('logs')" style="display:none;">🔐 Activity Log</div>
    </div>

    <!-- OVERVIEW TAB -->
    <div id="tab-overview">
      <div class="kpi-grid" id="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Loading...</div></div>
      </div>

      <!-- STORE RISK MATRIX -->
      <div class="section" style="margin-top:4px;">
        <div class="section-header">
          <div class="section-title">📊 Store Risk Matrix
            <span style="font-size:10px;color:var(--text2);margin-left:8px;">
              <span style="color:#3fb950;">🟢 Low</span> &nbsp;
              <span style="color:#e3b341;">🟡 Medium</span> &nbsp;
              <span style="color:#f85149;">🔴 High</span>
            </span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search store..." oninput="searchTable('risk-matrix-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('stores')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap" style="max-height:520px;">
          <table id="risk-matrix-table">
            <thead><tr>
              <th onclick="sortRiskTable(0)">Store #</th>
              <th onclick="sortRiskTable(1)">Store Name</th>
              <th onclick="sortRiskTable(2)">Area</th>
              <th onclick="sortRiskTable(3)">SKUs</th>
              <th onclick="sortRiskTable(4)" title="Avg days of inventory at current sales pace">Days Cover</th>
              <th onclick="sortRiskTable(5)" title="< 2 weeks supply / Total SKUs">Critical %</th>
              <th onclick="sortRiskTable(6)" title="Out of Stock / Total SKUs">OOS %</th>
              <th onclick="sortRiskTable(7)" title="> 12 weeks supply / Total SKUs">Overstock %</th>
              <th onclick="sortRiskTable(8)" title="No sales 8 weeks / Total SKUs">Dead %</th>
              <th onclick="sortRiskTable(9)">Lost Sales/Wk</th>
            </tr></thead>
            <tbody id="risk-matrix-body"></tbody>
          </table>
        </div>
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
              <th onclick="sortTable('critical-table',11)">Last Sold</th>
              <th onclick="sortTable('critical-table',12)">Last Received</th>
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
              <th onclick="sortTable('overstock-table',9)">Last Sold</th>
              <th onclick="sortTable('overstock-table',10)">Last Received</th>
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
              <th onclick="sortTable('deadstock-table',7)">WTS</th>
              <th onclick="sortTable('deadstock-table',8)">Days Cover</th>
              <th onclick="sortTable('deadstock-table',9)">Last Sold</th>
              <th onclick="sortTable('deadstock-table',10)">Last Received</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="deadstock-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="deadstock-pagination"></div>
      </div>
    </div>

    <!-- OUT OF STOCK TAB -->
    <div id="tab-outofstock" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">🚫 Out of Stock — Lost Sales <span class="badge badge-red" id="outofstock-count">0</span></div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('outofstock-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('outofstock')">⬇ Export CSV</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="outofstock-table">
            <thead><tr>
              <th onclick="sortTable('outofstock-table',0)">Store</th>
              <th onclick="sortTable('outofstock-table',1)">Area</th>
              <th onclick="sortTable('outofstock-table',2)">SKU Code</th>
              <th onclick="sortTable('outofstock-table',3)">Description</th>
              <th onclick="sortTable('outofstock-table',4)">Supplier</th>
              <th onclick="sortTable('outofstock-table',5)">P8 Ave/Wk</th>
              <th onclick="sortTable('outofstock-table',6)">Avg Cost</th>
              <th onclick="sortTable('outofstock-table',7)">Lost Sales/Wk</th>
              <th onclick="sortTable('outofstock-table',8)">Incoming PO</th>
              <th onclick="sortTable('outofstock-table',9)">Days No Sales</th>
              <th onclick="sortTable('outofstock-table',10)">Last Sold</th>
              <th onclick="sortTable('outofstock-table',11)">Last Received</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="outofstock-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="outofstock-pagination"></div>
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
              <th onclick="sortTable('stores-table',6)">WTS</th>
              <th onclick="sortTable('stores-table',7)">Days Cover</th>
              <th onclick="sortTable('stores-table',8)">Out of Stock</th>
              <th onclick="sortTable('stores-table',9)">Lost Sales/Wk</th>
              <th onclick="sortTable('stores-table',10)">Critical</th>
              <th onclick="sortTable('stores-table',11)">Overstock</th>
              <th onclick="sortTable('stores-table',12)">Dead Stock</th>
            </tr></thead>
            <tbody id="stores-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="stores-pagination"></div>
      </div>
    </div>

    <!-- SUPPLIERS TAB -->
    <div id="tab-suppliers" style="display:none;">
      <!-- SUPPLIER RISK MATRIX -->
      <div class="section">
        <div class="section-header">
          <div class="section-title">📊 Supplier Risk Matrix
            <span style="font-size:10px;color:var(--text2);margin-left:8px;">
              <span style="color:#3fb950;">🟢 Low</span> &nbsp;
              <span style="color:#e3b341;">🟡 Medium</span> &nbsp;
              <span style="color:#f85149;">🔴 High</span>
            </span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search supplier..." oninput="searchTable('supplier-risk-table',this.value)"/>
          </div>
        </div>
        <div class="table-wrap" style="max-height:480px;">
          <table id="supplier-risk-table">
            <thead><tr>
              <th onclick="sortSupplierRisk(0)">Supplier</th>
              <th onclick="sortSupplierRisk(1)">SKUs</th>
              <th onclick="sortSupplierRisk(2)">Inv Value</th>
              <th onclick="sortSupplierRisk(3)">Days Cover</th>
              <th onclick="sortSupplierRisk(4)">Critical %</th>
              <th onclick="sortSupplierRisk(5)">OOS %</th>
              <th onclick="sortSupplierRisk(6)">Overstock %</th>
              <th onclick="sortSupplierRisk(7)">Dead %</th>
              <th onclick="sortSupplierRisk(8)">Lost Sales/Wk</th>
            </tr></thead>
            <tbody id="supplier-risk-body"></tbody>
          </table>
        </div>
      </div>

      <!-- SUPPLIER ANALYSIS TABLE -->
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
              <th onclick="sortTable('suppliers-table',5)">WTS</th>
              <th onclick="sortTable('suppliers-table',6)">Days Cover</th>
              <th onclick="sortTable('suppliers-table',7)">Out of Stock</th>
              <th onclick="sortTable('suppliers-table',8)">Critical</th>
              <th onclick="sortTable('suppliers-table',9)">Overstock</th>
              <th onclick="sortTable('suppliers-table',10)">Dead Stock</th>
            </tr></thead>
            <tbody id="suppliers-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="suppliers-pagination"></div>
      </div>
    </div>

    <!-- SKU ANALYSIS TAB -->
    <div id="tab-skus" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">🔍 SKU Analysis
            <span class="badge badge-blue" id="skus-total-count" style="margin-left:8px;">0</span>
          </div>
          <div class="section-actions">
            <select class="filter-select" id="sku-status-filter" onchange="loadSKUs(1)" style="width:140px;">
              <option value="">All Status</option>
              <option value="critical">Critical</option>
              <option value="oos">Out of Stock</option>
              <option value="overstock">Overstock</option>
              <option value="deadstock">Dead Stock</option>
              <option value="normal">Normal</option>
            </select>
            <input type="text" class="table-search" id="sku-search-input" placeholder="Search SKU / Store / Supplier..." oninput="debouncedSKUSearch()"/>
          </div>
        </div>
        <div class="table-wrap" style="max-height:600px;">
          <table id="skus-table">
            <thead><tr>
              <th onclick="sortSKUs('storeName')">Store Name <span class="sort-ind" data-key="storeName"></span></th>
              <th onclick="sortSKUs('skuCode')">SKU <span class="sort-ind" data-key="skuCode"></span></th>
              <th onclick="sortSKUs('skuDesc')">Description <span class="sort-ind" data-key="skuDesc"></span></th>
              <th onclick="sortSKUs('supplierName')">Supplier <span class="sort-ind" data-key="supplierName"></span></th>
              <th onclick="sortSKUs('onHand')">On Hand <span class="sort-ind" data-key="onHand"></span></th>
              <th onclick="sortSKUs('qtyCasesNum')">Qty in Cases <span class="sort-ind" data-key="qtyCasesNum"></span></th>
              <th onclick="sortSKUs('invValue')">Inv Value <span class="sort-ind" data-key="invValue"></span></th>
              <th onclick="sortSKUs('skuWTS')">WTS <span class="sort-ind" data-key="skuWTS"></span></th>
              <th onclick="sortSKUs('skuDaysCover')">Days Cover <span class="sort-ind" data-key="skuDaysCover"></span></th>
              <th onclick="sortSKUs('p8ave')">P8 Ave/Wk <span class="sort-ind" data-key="p8ave"></span></th>
              <th onclick="sortSKUs('status')">Status <span class="sort-ind" data-key="status"></span></th>
              <th onclick="sortSKUs('lostSalesPerWeek')">Lost Sales/Wk <span class="sort-ind" data-key="lostSalesPerWeek"></span></th>
              <th onclick="sortSKUs('ico')">ICO <span class="sort-ind" data-key="ico"></span></th>
              <th onclick="sortSKUs('poOrderGR')">PO On Order <span class="sort-ind" data-key="poOrderGR"></span></th>
              <th onclick="sortSKUs('trfOrderGR')">Trf On Order <span class="sort-ind" data-key="trfOrderGR"></span></th>
              <th onclick="sortSKUs('dateLastSold')">Last Sold <span class="sort-ind" data-key="dateLastSold"></span></th>
              <th onclick="sortSKUs('dateLastReceived')">Last Received <span class="sort-ind" data-key="dateLastReceived"></span></th>
            </tr></thead>
            <tbody id="skus-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="skus-pagination"></div>
      </div>
    </div>

    <!-- ACTIVITY LOG TAB (admin only) -->
    <div id="tab-logs" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">🔐 Activity Log
            <span class="badge badge-blue" id="logs-count" style="margin-left:8px;">0</span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search user..." oninput="searchTable('logs-table',this.value)"/>
            <button class="btn btn-sm" onclick="loadLogs()">↺ Refresh</button>
            <button class="btn btn-sm" onclick="exportLogs()">⬇ Export Excel</button>
            <button class="btn btn-sm" style="border-color:var(--red);color:var(--red-light);" onclick="confirmClearLogs()">🗑 Delete All Logs</button>
          </div>
        </div>
        <div class="table-wrap" style="max-height:600px;">
          <table id="logs-table">
            <thead><tr>
              <th>User</th>
              <th>Login Time</th>
              <th>Logout Time</th>
              <th>Duration</th>
              <th>Area</th>
            </tr></thead>
            <tbody id="logs-body"></tbody>
          </table>
        </div>
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
let tablePages = { critical:1, overstock:1, deadstock:1, outofstock:1, stores:1, suppliers:1 };
const PAGE_SIZE = 50;

// ─── AUTH STATE ───────────────────────────────────────────────────────────────
let authToken = '';
let currentUser = null;

function tokenParam(prefix) {
  if (!authToken) return '';
  return (prefix || '?') + 'token=' + encodeURIComponent(authToken);
}

async function doLogin() {
  const btn = document.getElementById('login-btn');
  const errEl = document.getElementById('login-error');
  const username = document.getElementById('login-username').value.trim();
  const password = document.getElementById('login-password').value;
  if (!username || !password) { errEl.textContent = 'Enter username and password'; return; }
  btn.disabled = true; btn.textContent = 'Signing in...'; errEl.textContent = '';
  try {
    const r = await fetch('/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    const d = await r.json();
    if (!r.ok) { errEl.textContent = d.error || 'Login failed'; btn.disabled = false; btn.textContent = 'Sign In'; return; }
    authToken = d.token;
    currentUser = d;
    try { sessionStorage.setItem('camanava_token', authToken); sessionStorage.setItem('camanava_user', JSON.stringify(d)); } catch(e) {}
    document.getElementById('login-screen').style.display = 'none';
    startApp();
  } catch(e) {
    errEl.textContent = 'Connection error. Try again.';
    btn.disabled = false; btn.textContent = 'Sign In';
  }
}

async function doLogout() {
  try { await fetch('/api/logout', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ token: authToken }) }); } catch(e) {}
  authToken = ''; currentUser = null;
  try { sessionStorage.removeItem('camanava_token'); sessionStorage.removeItem('camanava_user'); } catch(e) {}
  location.reload();
}

// ─── ACTIVITY LOG ─────────────────────────────────────────────────────────────
let logsData = [];
async function loadLogs() {
  const r = await fetch('/api/logs?token=' + encodeURIComponent(authToken));
  if (!r.ok) { document.getElementById('logs-body').innerHTML = '<tr><td colspan="5" class="empty">Access denied or no logs sheet configured</td></tr>'; return; }
  const data = await r.json();
  if (!Array.isArray(data)) return;
  logsData = data;
  document.getElementById('logs-count').textContent = fmt(data.length);
  renderLogs(data);
}

function fmtLogTime(iso) {
  if (!iso) return '<span style="color:var(--text2);">—</span>';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return esc(iso);
  return d.toLocaleString('en-PH', { year:'numeric', month:'short', day:'numeric', hour:'2-digit', minute:'2-digit' });
}

function renderLogs(data) {
  const tbody = document.getElementById('logs-body');
  if (data.length === 0) { tbody.innerHTML = '<tr><td colspan="5" class="empty">No activity logs yet</td></tr>'; return; }
  tbody.innerHTML = data.map(r => {
    const logout = r.logoutTime ? fmtLogTime(r.logoutTime) : '<span style="color:var(--yellow-light);">Active / No logout</span>';
    const dur = r.duration ? r.duration : '<span style="color:var(--text2);">—</span>';
    return '<tr>' +
      '<td style="font-weight:600;">' + esc(r.user) + '</td>' +
      '<td class="mono">' + fmtLogTime(r.loginTime) + '</td>' +
      '<td class="mono">' + logout + '</td>' +
      '<td class="mono">' + dur + '</td>' +
      '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
      '</tr>';
  }).join('');
}

function exportLogs() {
  if (logsData.length === 0) { alert('No logs to export'); return; }
  const headers = ['User', 'Login Time', 'Logout Time', 'Duration', 'Area'];
  const rows = logsData.map(r => [r.user, r.loginTime, r.logoutTime, r.duration, r.area]);
  const csv = [headers.join(','), ...rows.map(row => row.map(c => '"' + (c || '').toString().replace(/"/g, '""') + '"').join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'activity_logs_' + new Date().toISOString().split('T')[0] + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}

async function confirmClearLogs() {
  if (!confirm('Delete ALL activity logs? This cannot be undone.')) return;
  const r = await fetch('/api/logs/clear', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token: authToken })
  });
  const d = await r.json();
  if (d.ok) { alert('Logs cleared.'); loadLogs(); }
  else alert('Failed to clear logs.');
}

// ─── INIT ─────────────────────────────────────────────────────────────────────
async function init() {
  console.log('[Init] Starting...');
  // Hide loading overlay immediately
  const loading = document.getElementById('loading-overlay');
  const loginEl = document.getElementById('login-screen');
  if (loading) loading.style.display = 'none';

  // Try restoring session, but never block on it
  let restored = false;
  try {
    let savedToken = null, savedUser = null;
    try { savedToken = sessionStorage.getItem('camanava_token'); } catch(e) {}
    try { savedUser = sessionStorage.getItem('camanava_user'); } catch(e) {}
    if (savedToken && savedUser) {
      console.log('[Init] Found saved session, verifying...');
      const r = await Promise.race([
        fetch('/api/me?token=' + encodeURIComponent(savedToken)),
        new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 5000))
      ]);
      if (r && r.ok) {
        authToken = savedToken;
        try { currentUser = JSON.parse(savedUser); } catch(e) { currentUser = null; }
        if (currentUser) { restored = true; startApp(); return; }
      }
    }
  } catch(e) { console.warn('[Init] Session restore failed:', e && e.message); }

  // Show login
  console.log('[Init] Showing login screen');
  if (loginEl) loginEl.style.display = 'flex';
  try {
    const pwInput = document.getElementById('login-password');
    const userInput = document.getElementById('login-username');
    if (pwInput) pwInput.addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
    if (userInput) userInput.addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
  } catch(e) {}
}

async function startApp() {
  document.getElementById('loading-overlay').style.display = 'flex';
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
  // Show user info + lock area filter if needed
  applyUserRestrictions();
  await loadFilters();
  await loadAll();
  if (!ready) { pollAndRefresh(); }
}

function applyUserRestrictions() {
  if (!currentUser) return;
  const info = document.getElementById('user-info');
  if (info) {
    const areaLabel = currentUser.isAdmin ? 'All Areas' : (currentUser.area || 'All Areas');
    info.textContent = '👤 ' + currentUser.username + ' (' + (currentUser.level || 'user') + ' · ' + areaLabel + ')';
  }
  // If user is locked to an area, set it in filters and disable the area dropdown
  if (!currentUser.isAdmin && currentUser.area) {
    activeFilters.area = currentUser.area;
  }
  // Show Activity Log tab only for admins
  const logsTabBtn = document.getElementById('tab-btn-logs');
  if (logsTabBtn) logsTabBtn.style.display = currentUser.isAdmin ? '' : 'none';
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
  // Include current area scope so stores list is filtered
  const params = new URLSearchParams();
  if (authToken) params.set('token', authToken);
  if (activeFilters.area) params.set('area', activeFilters.area);
  const r = await fetch('/api/filters?' + params.toString());
  const d = await r.json();
  if (d.error) return;

  populateSelect('f-area', d.areas, 'All Areas');
  populateSelect('f-dept', d.depts, 'All Departments');
  populateSelect('f-subdept', d.subDepts, 'All Sub-Depts');
  populateSelect('f-cls', d.classes, 'All Classes');
  populateSelect('f-supplier', d.suppliers, 'All Suppliers');

  populateStores(d.stores);

  // Lock area filter for non-admin users
  if (currentUser && !currentUser.isAdmin && currentUser.area) {
    const areaSel = document.getElementById('f-area');
    if (areaSel) {
      areaSel.value = currentUser.area;
      areaSel.disabled = true;
      areaSel.style.opacity = '0.6';
      areaSel.style.cursor = 'not-allowed';
      areaSel.title = 'Locked to your assigned area';
    }
  }
}

function populateStores(stores) {
  const storeSelect = document.getElementById('f-store');
  if (!storeSelect) return;
  storeSelect.innerHTML = '<option value="">All Stores</option>';
  (stores || []).forEach(s => {
    const o = document.createElement('option');
    o.value = s.id; o.textContent = s.id + ' - ' + s.name;
    storeSelect.appendChild(o);
  });
}

// Reload only the store dropdown based on current area
async function reloadStoresForArea() {
  const params = new URLSearchParams();
  if (authToken) params.set('token', authToken);
  if (activeFilters.area) params.set('area', activeFilters.area);
  const r = await fetch('/api/filters?' + params.toString());
  const d = await r.json();
  if (d.error) return;
  populateStores(d.stores);
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
  // When area changes, reset store selection and reload store list scoped to area
  if (key === 'area') {
    delete activeFilters.store;
    const storeSel = document.getElementById('f-store');
    if (storeSel) storeSel.value = '';
    reloadStoresForArea();
  }
  renderActiveTags();
  loadAll();
}

function removeFilterTag(btn) {
  const key = btn.getAttribute('data-key');
  if (!key) return;
  // Prevent non-admin from removing their locked area
  if (key === 'area' && currentUser && !currentUser.isAdmin && currentUser.area) return;
  const idMap = { area: 'f-area', store: 'f-store', dept: 'f-dept', subDept: 'f-subdept', cls: 'f-cls', supplier: 'f-supplier' };
  const sel = document.getElementById(idMap[key]);
  if (sel) sel.value = '';
  applyFilter(key, '');
}

function clearFilters() {
  activeFilters = {};
  ['f-store','f-dept','f-subdept','f-cls','f-supplier']
    .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  // Non-admin: keep area locked
  if (currentUser && !currentUser.isAdmin && currentUser.area) {
    activeFilters.area = currentUser.area;
  } else {
    const areaEl = document.getElementById('f-area');
    if (areaEl) areaEl.value = '';
  }
  reloadStoresForArea();
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
  if (authToken) params.set('token', authToken);
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
    kpiCard('Out of Stock', fmt(d.outOfStockCount || 0), 'SKUs losing sales', 'red'),
    kpiCard('Lost Sales/Wk', '₱' + fmtM(d.totalLostSalesPerWeek || 0), 'estimated/week', 'red'),
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

  // Register datalabels plugin globally (once)
  if (window.ChartDataLabels && !Chart._dlRegistered) {
    Chart.register(window.ChartDataLabels);
    Chart._dlRegistered = true;
  }

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
    options: {
      responsive:true,
      plugins:{
        legend:{display:false},
        datalabels:{ anchor:'end', align:'top', color:'#e6edf3', font:{size:10,weight:'600'}, formatter: v => '₱'+fmtM(v) }
      },
      layout:{ padding:{top:20} },
      scales:{ y:{ ticks:{ color:'#8b949e', callback: v => '₱'+fmtM(v) }, grid:{color:'#30363d'} }, x:{ ticks:{color:'#8b949e',font:{size:9}}, grid:{display:false} } }
    }
  });

  // Top 10 Suppliers
  const top10sup = suppliers.slice(0,10);
  charts.suppliers = new Chart(document.getElementById('chart-suppliers'), {
    type: 'bar',
    data: {
      labels: top10sup.map(s => s.supplierName.substring(0,15)),
      datasets: [{ label: 'Inv Value', data: top10sup.map(s => s.totalValue), backgroundColor: '#1f6feb', borderRadius: 4 }]
    },
    options: {
      indexAxis:'y', responsive:true,
      plugins:{
        legend:{display:false},
        datalabels:{ anchor:'end', align:'right', color:'#e6edf3', font:{size:10,weight:'600'}, formatter: v => '₱'+fmtM(v) }
      },
      layout:{ padding:{right:50} },
      scales:{ x:{ ticks:{ color:'#8b949e', callback: v => '₱'+fmtM(v) }, grid:{color:'#30363d'} }, y:{ ticks:{color:'#8b949e',font:{size:9}}, grid:{display:false} } }
    }
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
        { label: 'Inv Value', data: areas.map(a => areaMap[a].value), backgroundColor: '#2ea043', borderRadius: 4, yAxisID: 'y',
          datalabels: { anchor:'end', align:'top', color:'#3fb950', font:{size:9,weight:'600'}, formatter: v => '₱'+fmtM(v) } },
        { label: 'Critical', data: areas.map(a => areaMap[a].critical), backgroundColor: '#f85149', borderRadius: 4, yAxisID: 'y1',
          datalabels: { anchor:'end', align:'top', color:'#f85149', font:{size:9,weight:'600'}, formatter: v => fmt(v) } }
      ]
    },
    options: {
      responsive:true,
      plugins:{ legend:{labels:{color:'#8b949e',font:{size:10}}} },
      layout:{ padding:{top:20} },
      scales:{ y:{ ticks:{color:'#8b949e',callback: v=>'₱'+fmtM(v)}, grid:{color:'#30363d'} }, y1:{position:'right',ticks:{color:'#8b949e'},grid:{display:false}}, x:{ticks:{color:'#8b949e',font:{size:9}},grid:{display:false}} }
    }
  });

  // Risk Distribution
  const critTotal = stores.reduce((s,r) => s + r.criticalCount, 0);
  const ovTotal = stores.reduce((s,r) => s + r.overstockCount, 0);
  const deadTotal = stores.reduce((s,r) => s + r.deadCount, 0);
  const totalSKUs = stores.reduce((s,r) => s + r.totalSKUs, 0);
  const normal = Math.max(0, totalSKUs - critTotal - ovTotal - deadTotal);
  const grandTotal = normal + ovTotal + critTotal + deadTotal;
  charts.risk = new Chart(document.getElementById('chart-risk'), {
    type: 'doughnut',
    data: {
      labels: ['Normal', 'Overstock', 'Critical', 'Dead Stock'],
      datasets: [{ data: [normal, ovTotal, critTotal, deadTotal], backgroundColor: ['#2ea043','#e3b341','#f85149','#8b949e'], borderWidth: 0 }]
    },
    options: {
      responsive:true,
      plugins:{
        legend:{ position:'bottom', labels:{ color:'#8b949e', font:{size:10} } },
        datalabels:{
          color:'#fff', font:{size:11,weight:'700'},
          formatter: (v) => {
            if (!grandTotal) return '';
            const pct = (v / grandTotal) * 100;
            return pct < 3 ? '' : pct.toFixed(1) + '%';
          }
        }
      }
    }
  });
}

// ─── RISK MATRIX ──────────────────────────────────────────────────────────────
// Thresholds:
// Critical:  Low <5%, Med 5-10%, High >10%
// OOS:       Low <3%, Med 3-7%, High >7%
// Overstock: Low <10%, Med 10-20%, High >20%
// Dead:      Low <3%, Med 3-7%, High >7%
function riskPill(pct, type) {
  if (pct == null || isNaN(pct)) return '<span class="risk-pill risk-none">—</span>';
  let level = 'low';
  if (type === 'critical') {
    if (pct > 10) level = 'high';
    else if (pct >= 5) level = 'med';
  } else if (type === 'oos') {
    if (pct > 7) level = 'high';
    else if (pct >= 3) level = 'med';
  } else if (type === 'overstock') {
    if (pct > 20) level = 'high';
    else if (pct >= 10) level = 'med';
  } else if (type === 'dead') {
    if (pct > 7) level = 'high';
    else if (pct >= 3) level = 'med';
  }
  const icon = level === 'high' ? '🔴' : level === 'med' ? '🟡' : '🟢';
  return '<span class="risk-pill risk-' + level + '">' + icon + ' ' + pct.toFixed(1) + '%</span>';
}

// Days Cover pill: 🔴<7, 🟡7-14, 🟢15-60, 🟡61-90, 🔴>90
function daysCoverPill(days) {
  if (days == null || isNaN(days)) return '<span class="risk-pill risk-none">No Sales</span>';
  let level = 'low';
  if (days < 7) level = 'high';
  else if (days < 15) level = 'med';
  else if (days <= 60) level = 'low';
  else if (days <= 90) level = 'med';
  else level = 'high';
  const icon = level === 'high' ? '🔴' : level === 'med' ? '🟡' : '🟢';
  const display = days > 999 ? '999+' : days.toFixed(0);
  return '<span class="risk-pill risk-' + level + '">' + icon + ' ' + display + 'd</span>';
}

let riskMatrixData = [];
async function loadRiskMatrix() {
  const r = await fetch('/api/stores' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  riskMatrixData = data;
  renderRiskMatrix(data);
}

function renderRiskMatrix(data) {
  const tbody = document.getElementById('risk-matrix-body');
  if (!tbody) return;
  if (data.length === 0) { tbody.innerHTML = '<tr><td colspan="10" class="empty">No data</td></tr>'; return; }
  tbody.innerHTML = data.map(s => {
    return '<tr>' +
      '<td class="mono" style="font-weight:600;">' + esc(s.storeNumber) + '</td>' +
      '<td>' + esc(s.storeName) + '</td>' +
      '<td><span class="badge badge-blue">' + esc(s.area) + '</span></td>' +
      '<td class="mono">' + fmt(s.totalSKUs) + '</td>' +
      '<td>' + daysCoverPill(s.daysCover) + '</td>' +
      '<td>' + riskPill(s.criticalPct, 'critical') + '</td>' +
      '<td>' + riskPill(s.oosPct, 'oos') + '</td>' +
      '<td>' + riskPill(s.overstockPct, 'overstock') + '</td>' +
      '<td>' + riskPill(s.deadPct, 'dead') + '</td>' +
      '<td class="mono" style="color:var(--red-light);font-weight:600;">₱' + fmtN(s.totalLostSales || 0) + '</td>' +
      '</tr>';
  }).join('');
}

let riskSortState = {};
function sortRiskTable(colIndex) {
  const keys = ['storeNumber','storeName','area','totalSKUs','daysCover','criticalPct','oosPct','overstockPct','deadPct','totalLostSales'];
  const key = keys[colIndex];
  const asc = riskSortState[colIndex] !== true;
  riskSortState[colIndex] = asc;
  const sorted = [...riskMatrixData].sort((a, b) => {
    const av = a[key], bv = b[key];
    if (av == null && bv == null) return 0;
    if (av == null) return asc ? 1 : -1;
    if (bv == null) return asc ? -1 : 1;
    if (typeof av === 'number' && typeof bv === 'number') return asc ? av - bv : bv - av;
    return asc ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
  });
  renderRiskMatrix(sorted);
}

// ─── TABLE LOADING ────────────────────────────────────────────────────────────
async function loadTabData() {
  if (activeTab === 'overview') { await loadCharts(); await loadRiskMatrix(); return; }
  if (activeTab === 'critical') await loadCritical();
  if (activeTab === 'overstock') await loadOverstock();
  if (activeTab === 'deadstock') await loadDeadstock();
  if (activeTab === 'outofstock') await loadOutOfStock();
  if (activeTab === 'stores') await loadStores();
  if (activeTab === 'suppliers') await loadSuppliers();
  if (activeTab === 'skus') await loadSKUs(1);
  if (activeTab === 'logs') await loadLogs();
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
async function loadOutOfStock() {
  const r = await fetch('/api/outofstock' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  document.getElementById('outofstock-count').textContent = fmt(data.length);
  renderTable('outofstock-body', data, renderOutOfStockRow, 'outofstock-pagination', 'outofstock', tablePages.outofstock);
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
  supplierRiskData = data;
  renderSupplierRisk(data);
  renderTable('suppliers-body', data, renderSupplierRow, 'suppliers-pagination', 'suppliers', tablePages.suppliers);
}

let supplierRiskData = [];
function renderSupplierRisk(data) {
  const tbody = document.getElementById('supplier-risk-body');
  if (!tbody) return;
  if (data.length === 0) { tbody.innerHTML = '<tr><td colspan="9" class="empty">No data</td></tr>'; return; }
  tbody.innerHTML = data.map(s => {
    return '<tr>' +
      '<td>' + esc(s.supplierName) + '</td>' +
      '<td class="mono">' + fmt(s.totalSKUs) + '</td>' +
      '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(s.totalValue) + '</td>' +
      '<td>' + daysCoverPill(s.daysCover) + '</td>' +
      '<td>' + riskPill(s.criticalPct, 'critical') + '</td>' +
      '<td>' + riskPill(s.oosPct, 'oos') + '</td>' +
      '<td>' + riskPill(s.overstockPct, 'overstock') + '</td>' +
      '<td>' + riskPill(s.deadPct, 'dead') + '</td>' +
      '<td class="mono" style="color:var(--red-light);font-weight:600;">₱' + fmtN(s.totalLostSales || 0) + '</td>' +
      '</tr>';
  }).join('');
}

let supplierRiskSortState = {};
function sortSupplierRisk(colIndex) {
  const keys = ['supplierName','totalSKUs','totalValue','daysCover','criticalPct','oosPct','overstockPct','deadPct','totalLostSales'];
  const key = keys[colIndex];
  const asc = supplierRiskSortState[colIndex] !== true;
  supplierRiskSortState[colIndex] = asc;
  const sorted = [...supplierRiskData].sort((a, b) => {
    const av = a[key], bv = b[key];
    if (av == null && bv == null) return 0;
    if (av == null) return asc ? 1 : -1;
    if (bv == null) return asc ? -1 : 1;
    if (typeof av === 'number' && typeof bv === 'number') return asc ? av - bv : bv - av;
    return asc ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
  });
  renderSupplierRisk(sorted);
}

// ─── SKU ANALYSIS ─────────────────────────────────────────────────────────────
let skuState = { page: 1, pageSize: 100, sortBy: '', sortDir: 'asc', search: '', total: 0, totalPages: 1 };
let skuSearchTimer = null;
function debouncedSKUSearch() {
  clearTimeout(skuSearchTimer);
  skuSearchTimer = setTimeout(() => loadSKUs(1), 350);
}

async function loadSKUs(page) {
  if (page) skuState.page = page;
  const search = document.getElementById('sku-search-input').value;
  const status = document.getElementById('sku-status-filter').value;
  skuState.search = search;
  const params = new URLSearchParams({
    ...activeFilters,
    page: skuState.page,
    pageSize: skuState.pageSize,
    sortBy: skuState.sortBy,
    sortDir: skuState.sortDir,
    search,
    status
  });
  if (authToken) params.set('token', authToken);
  const r = await fetch('/api/skus?' + params.toString());
  const d = await r.json();
  if (!d || d.error) return;
  skuState.total = d.total;
  skuState.totalPages = d.totalPages;
  document.getElementById('skus-total-count').textContent = fmt(d.total);
  renderSKUTable(d.rows);
  renderSKUPagination();
  updateSortIndicators();
}

function renderSKUTable(rows) {
  const tbody = document.getElementById('skus-body');
  if (rows.length === 0) { tbody.innerHTML = '<tr><td colspan="17" class="empty">No data found</td></tr>'; return; }
  tbody.innerHTML = rows.map(r => {
    const wts = r.weeksToSell != null ? r.weeksToSell.toFixed(1) : '—';
    const dc = r.daysCover != null ? r.daysCover.toFixed(0) + 'd' : '—';
    let statusClass = 'status-normal';
    if (r.status === 'Critical') statusClass = 'status-critical';
    else if (r.status === 'OOS') statusClass = 'status-oos';
    else if (r.status === 'Overstock') statusClass = 'status-overstock';
    else if (r.status === 'Dead Stock') statusClass = 'status-dead';
    const qtyCasesDisplay = r.qtyCases === 'Per Piece'
      ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>'
      : r.qtyCases;
    const lostSales = r.lostSalesPerWeek > 0
      ? '<span style="color:var(--red-light);font-weight:600;">₱' + fmtN(r.lostSalesPerWeek) + '</span>'
      : '—';
    return '<tr>' +
      '<td>' + esc(r.storeName) + '</td>' +
      '<td class="mono">' + esc(r.skuCode) + '</td>' +
      '<td>' + esc(r.skuDesc) + '</td>' +
      '<td>' + esc(r.supplierName) + '</td>' +
      '<td class="mono">' + fmt(r.onHand) + '</td>' +
      '<td class="mono">' + qtyCasesDisplay + '</td>' +
      '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(r.invValue) + '</td>' +
      '<td class="mono">' + wts + '</td>' +
      '<td class="mono">' + dc + '</td>' +
      '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
      '<td><span class="' + statusClass + '">' + esc(r.status) + '</span></td>' +
      '<td class="mono">' + lostSales + '</td>' +
      '<td class="mono">' + esc(r.ico || '—') + '</td>' +
      '<td class="mono">' + fmt(r.poOrderGR) + '</td>' +
      '<td class="mono">' + fmt(r.trfOrderGR) + '</td>' +
      '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
      '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
      '</tr>';
  }).join('');
}

function renderSKUPagination() {
  const cont = document.getElementById('skus-pagination');
  const total = skuState.totalPages;
  const current = skuState.page;
  if (total <= 1) { cont.innerHTML = '<span class="page-info">' + fmt(skuState.total) + ' SKUs</span>'; return; }
  let html = '<span class="page-info">' + fmt(skuState.total) + ' SKUs | Page ' + current + ' of ' + total + '</span>';
  html += '<button class="page-btn" onclick="goSKUPage(1)">&laquo;</button>';
  html += '<button class="page-btn" onclick="goSKUPage(' + Math.max(1, current - 1) + ')">&lsaquo;</button>';
  const startP = Math.max(1, current - 2), endP = Math.min(total, current + 2);
  for (let i = startP; i <= endP; i++) {
    html += '<button class="page-btn' + (i === current ? ' active' : '') + '" onclick="goSKUPage(' + i + ')">' + i + '</button>';
  }
  html += '<button class="page-btn" onclick="goSKUPage(' + Math.min(total, current + 1) + ')">&rsaquo;</button>';
  html += '<button class="page-btn" onclick="goSKUPage(' + total + ')">&raquo;</button>';
  cont.innerHTML = html;
}

function goSKUPage(p) {
  loadSKUs(p);
}

function sortSKUs(key) {
  if (skuState.sortBy === key) {
    skuState.sortDir = skuState.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    skuState.sortBy = key;
    skuState.sortDir = 'asc';
  }
  loadSKUs(1);
}

function updateSortIndicators() {
  document.querySelectorAll('#skus-table .sort-ind').forEach(el => {
    el.classList.remove('asc', 'desc');
    if (el.getAttribute('data-key') === skuState.sortBy) {
      el.classList.add(skuState.sortDir);
    }
  });
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
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
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
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td><span class="action-badge ' + ac + '">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderDeadstockRow(r) {
  const wts = r.weeksToSell != null ? r.weeksToSell.toFixed(1) : 'No Sales';
  const dc = r.daysCover != null ? r.daysCover.toFixed(0) + 'd' : 'No Sales';
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono" style="color:var(--text2);">' + wts + '</td>' +
    '<td class="mono" style="color:var(--text2);">' + dc + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td><span class="action-badge action-markdown">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderOutOfStockRow(r) {
  const ac = r.action === 'URGENT: Place PO Now' ? 'action-urgent' : 'action-po';
  const days = r.daysNoSales !== '' && r.daysNoSales != null ? r.daysNoSales : '-';
  const daysColor = (typeof r.daysNoSales === 'number' && r.daysNoSales > 30) ? 'var(--red-light)' : (typeof r.daysNoSales === 'number' && r.daysNoSales > 14) ? 'var(--yellow-light)' : 'var(--text)';
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono">₱' + fmtN(r.avgCost) + '</td>' +
    '<td class="mono" style="color:var(--red-light);font-weight:600;">₱' + fmtN(r.lostSalesPerWeek) + '</td>' +
    '<td class="mono">' + fmt(r.totalPO) + '</td>' +
    '<td class="mono" style="color:' + daysColor + ';font-weight:600;">' + days + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td><span class="action-badge ' + ac + '">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderStoreRow(r) {
  const oo = r.oosCount > 0 ? '<span style="color:var(--red-light);font-weight:600;">' + fmt(r.oosCount) + '</span>' : '0';
  const ci = r.criticalCount > 0 ? '<span style="color:var(--red-light);font-weight:600;">' + fmt(r.criticalCount) + '</span>' : '0';
  const ov = r.overstockCount > 0 ? '<span style="color:var(--yellow-light);">' + fmt(r.overstockCount) + '</span>' : '0';
  const dd = r.deadCount > 0 ? '<span style="color:var(--text2);">' + fmt(r.deadCount) + '</span>' : '0';
  const wts = r.weeksToSell != null ? r.weeksToSell.toFixed(1) : '—';
  return '<tr>' +
    '<td class="mono" style="font-weight:600;">' + esc(r.storeNumber) + '</td>' +
    '<td>' + esc(r.storeName) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(r.totalValue) + '</td>' +
    '<td class="mono">' + fmt(r.totalOnHand) + '</td>' +
    '<td class="mono">' + fmt(r.totalSKUs) + '</td>' +
    '<td class="mono">' + wts + '</td>' +
    '<td>' + daysCoverPill(r.daysCover) + '</td>' +
    '<td>' + oo + '</td>' +
    '<td class="mono" style="color:var(--red-light);font-weight:600;">₱' + fmtN(r.totalLostSales || 0) + '</td>' +
    '<td>' + ci + '</td>' +
    '<td>' + ov + '</td>' +
    '<td>' + dd + '</td>' +
    '</tr>';
}
function renderSupplierRow(r) {
  const oo = r.oosCount > 0 ? '<span style="color:var(--red-light);font-weight:600;">' + fmt(r.oosCount) + '</span>' : '0';
  const ci = r.criticalCount > 0 ? '<span style="color:var(--red-light);font-weight:600;">' + fmt(r.criticalCount) + '</span>' : '0';
  const ov = r.overstockCount > 0 ? '<span style="color:var(--yellow-light);">' + fmt(r.overstockCount) + '</span>' : '0';
  const dd = r.deadCount > 0 ? '<span style="color:var(--text2);">' + fmt(r.deadCount) + '</span>' : '0';
  const wts = r.weeksToSell != null ? r.weeksToSell.toFixed(1) : '—';
  return '<tr>' +
    '<td class="mono">' + esc(r.supplierCode) + '</td>' +
    '<td>' + esc(r.supplierName) + '</td>' +
    '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(r.totalValue) + '</td>' +
    '<td class="mono">' + fmt(r.totalOnHand) + '</td>' +
    '<td class="mono">' + fmt(r.totalSKUs) + '</td>' +
    '<td class="mono">' + wts + '</td>' +
    '<td>' + daysCoverPill(r.daysCover) + '</td>' +
    '<td>' + oo + '</td>' +
    '<td>' + ci + '</td>' +
    '<td>' + ov + '</td>' +
    '<td>' + dd + '</td>' +
    '</tr>';
}

// ─── TABS ─────────────────────────────────────────────────────────────────────
function showTab(name) {
  ['overview','outofstock','critical','overstock','deadstock','stores','suppliers','skus','logs'].forEach(t => {
    const el = document.getElementById('tab-' + t);
    if (el) el.style.display = t === name ? '' : 'none';
  });
  document.querySelectorAll('.tab').forEach((el) => {
    el.classList.remove('active');
  });
  // Mark the clicked tab active by matching its onclick target
  document.querySelectorAll('.tab').forEach((el) => {
    if (el.getAttribute('onclick') && el.getAttribute('onclick').includes("'" + name + "'")) el.classList.add('active');
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
function startInit() {
  init().catch(function(e) {
    console.error('Init error:', e);
    const loading = document.getElementById('loading-overlay');
    const login = document.getElementById('login-screen');
    if (loading) loading.style.display = 'none';
    if (login) login.style.display = 'flex';
  });
  // Fallback: if after 3s nothing is visible, force show login
  setTimeout(function() {
    const login = document.getElementById('login-screen');
    const app = document.getElementById('app');
    const loading = document.getElementById('loading-overlay');
    if (login && app && loading) {
      const anyVisible = (login.style.display && login.style.display !== 'none')
        || (app.style.display && app.style.display !== 'none');
      // If loading still showing OR nothing visible, force show login
      if (!anyVisible && loading.style.display !== 'none') {
        // still loading is OK
      } else if (!anyVisible) {
        login.style.display = 'flex';
      }
    }
  }, 3000);
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', startInit);
} else {
  startInit();
}
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
