'use strict';

// Force IPv4 for DNS resolution — many cloud containers have broken IPv6 egress,
// which causes googleapis "Premature close" errors on the OAuth token endpoint.
// This must run before any network-facing modules are required.
const dns = require('dns');
if (typeof dns.setDefaultResultOrder === 'function') {
  dns.setDefaultResultOrder('ipv4first');
}

const express = require('express');
const { google } = require('googleapis');
const { parse } = require('csv-parse');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const cron = require('node-cron');
const { Readable } = require('stream');
const crypto = require('crypto');
const https = require('https');

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
  catMap: {},         // deptName (uppercase) -> catName
  upcMap: {},         // upc -> { sku, desc }
  top300: [],         // [{ area, storeNumber, storeName, rank, sku, desc }]
  storeSkuIndex: {},  // "storeNum_skuCode" -> enriched row (for fast lookup)
  kpis: {},
  criticalItems: [],
  overstockItems: [],
  agingItems: [],
  blackInventoryItems: [],
  negativeSkuItems: [],
  deadStockItems: [],
  outOfStockItems: [],
  storeAnalysis: [],
  supplierAnalysis: [],
  filterMeta: {},
  refreshing: false,
  error: null
};

// ─── GOOGLE DRIVE AUTH ────────────────────────────────────────────────────────
// Hand-rolled JWT → access_token exchange using ONLY Node built-ins (crypto + https).
// This bypasses gaxios/undici (whose transitive updates have caused
// "Premature close" errors on cloud containers). Forces IPv4 at the socket layer.
let tokenCache = { token: null, expiresAt: 0 };

function httpsRequest(opts, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(opts, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString('utf8') });
      });
      res.on('error', reject);
    });
    req.on('error', reject);
    req.setTimeout(30000, () => req.destroy(new Error('OAuth request timeout after 30s')));
    if (body) req.write(body);
    req.end();
  });
}

async function getAccessToken() {
  // Return cached token if still valid (60s safety window before real expiry)
  if (tokenCache.token && tokenCache.expiresAt > Date.now() + 60000) {
    return tokenCache.token;
  }
  if (!GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY) {
    throw new Error('Missing GOOGLE_CLIENT_EMAIL or GOOGLE_PRIVATE_KEY');
  }
  // 1. Build signed JWT (RS256)
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claims = {
    iss: GOOGLE_CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };
  const b64url = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const signingInput = b64url(header) + '.' + b64url(claims);
  const signature = crypto.createSign('RSA-SHA256').update(signingInput).sign(GOOGLE_PRIVATE_KEY).toString('base64url');
  const jwt = signingInput + '.' + signature;

  // 2. POST to Google's token endpoint (native https, IPv4)
  const body = 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + encodeURIComponent(jwt);
  const res = await httpsRequest({
    hostname: 'oauth2.googleapis.com',
    port: 443,
    path: '/token',
    method: 'POST',
    family: 4,
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(body),
      'Accept': 'application/json'
    }
  }, body);

  let data;
  try { data = JSON.parse(res.body); } catch (e) {
    throw new Error('OAuth response not JSON (HTTP ' + res.status + '): ' + res.body.slice(0, 200));
  }
  if (res.status !== 200 || !data.access_token) {
    throw new Error('OAuth token exchange failed (HTTP ' + res.status + '): ' + (data.error_description || data.error || res.body));
  }
  tokenCache.token = data.access_token;
  tokenCache.expiresAt = Date.now() + (data.expires_in * 1000);
  console.log('[Auth] Access token acquired via native OAuth (expires in ' + data.expires_in + 's)');
  return tokenCache.token;
}

async function getDriveClient() {
  const token = await getAccessToken();
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: token });
  return google.drive({ version: 'v3', auth });
}

async function getSheetsClient() {
  const token = await getAccessToken();
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: token });
  return google.sheets({ version: 'v4', auth });
}

// ─── ACTIVITY LOG (Google Sheets) ─────────────────────────────────────────────
// Appends a login row, returns the row number for later logout update
async function logLoginEvent(username, area) {
  if (!LOGS_SHEET_ID) { console.warn('[Logs] LOGS_SHEET_ID not set, skipping log'); return null; }
  try {
    const sheets = await getSheetsClient();
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
async function logLogoutEvent(rowNum, loginTimeISO, reason) {
  if (!LOGS_SHEET_ID || !rowNum) return;
  try {
    const sheets = await getSheetsClient();
    const logoutTime = new Date();
    const loginTime = new Date(loginTimeISO);
    const durationMs = logoutTime - loginTime;
    const mins = Math.floor(durationMs / 60000);
    const hrs = Math.floor(mins / 60);
    const remMins = mins % 60;
    let durationStr = hrs > 0 ? (hrs + 'h ' + remMins + 'm') : (mins + 'm');
    if (reason === 'auto-timeout') durationStr += ' (auto)';
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
    const sheets = await getSheetsClient();
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
    const sheets = await getSheetsClient();
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


// ─── NATIVE DRIVE API HELPERS ────────────────────────────────────────────────
// Hand-rolled Drive REST calls using native https + IPv4, bypassing gaxios/undici
// (which causes "Premature close" errors on some cloud containers like Railway).
async function driveList(q, fields, pageSize) {
  const token = await getAccessToken();
  const params = new URLSearchParams({
    q,
    fields: fields || 'files(id,name,size,modifiedTime,md5Checksum)',
    pageSize: String(pageSize || 5)
  });
  const res = await httpsRequest({
    hostname: 'www.googleapis.com',
    port: 443,
    path: '/drive/v3/files?' + params.toString(),
    method: 'GET',
    family: 4,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/json'
    }
  });
  if (res.status !== 200) {
    throw new Error('Drive list failed (HTTP ' + res.status + '): ' + res.body.slice(0, 300));
  }
  return JSON.parse(res.body);
}

function driveDownload(fileId) {
  return getAccessToken().then((token) => new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'www.googleapis.com',
      port: 443,
      path: '/drive/v3/files/' + encodeURIComponent(fileId) + '?alt=media',
      method: 'GET',
      family: 4,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': '*/*'
      }
    }, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        const buf = Buffer.concat(chunks);
        if (res.statusCode !== 200) {
          reject(new Error('Drive download failed (HTTP ' + res.statusCode + '): ' + buf.toString('utf8').slice(0, 300)));
        } else {
          resolve(buf);
        }
      });
      res.on('error', reject);
    });
    req.on('error', reject);
    req.setTimeout(180000, () => req.destroy(new Error('Drive download timeout after 180s')));
    req.end();
  }));
}

// ─── FIND FILE IN FOLDER ──────────────────────────────────────────────────────
async function findFile(_drive, fileName) {
  const q = `'${GDRIVE_FOLDER_ID}' in parents and name='${fileName}' and trashed=false`;
  const data = await driveList(q);
  const files = data.files;
  if (!files || files.length === 0) return null;
  return files[0];
}

// ─── DOWNLOAD FILE AS BUFFER ──────────────────────────────────────────────────
async function downloadFileBuffer(_drive, fileId) {
  return driveDownload(fileId);
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

// ─── PARSE CATCODE SHEET FROM XLSX ───────────────────────────────────────────
function parseCatCodeXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().trim() === 'catcode');
  if (!sheetName) {
    console.warn('[CatCode] No "CatCode" sheet found. Available sheets: ' + wb.SheetNames.join(', '));
    return {};
  }
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const catMap = {};
  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    const deptName = (row[0] || '').toString().trim().toUpperCase();
    const catName = (row[1] || '').toString().trim();
    if (deptName && catName) catMap[deptName] = catName;
  }
  console.log('[CatCode] Loaded ' + Object.keys(catMap).length + ' category mappings');
  return catMap;
}

// ─── PARSE UPC17 SHEET FROM XLSX ──────────────────────────────────────────────
function parseUPCXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().trim() === 'upc17');
  if (!sheetName) {
    console.warn('[UPC] No "UPC17" sheet found. Available sheets: ' + wb.SheetNames.join(', '));
    return {};
  }
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const upcMap = {};
  // Skip header row (row 0). One row per UPC (multiple UPCs per SKU possible).
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    const upc = (row[0] || '').toString().trim();
    const sku = (row[1] || '').toString().trim();
    const desc = (row[2] || '').toString().trim();
    if (upc && sku) upcMap[upc] = { sku, desc };
  }
  console.log('[UPC] Loaded ' + Object.keys(upcMap).length + ' UPC mappings');
  return upcMap;
}

// ─── PARSE TOP300 SHEET FROM XLSX ─────────────────────────────────────────────
// Columns: A=Area, B=Store Number, C=Store Name, D=Rank, E=SKU, F=Item_Description
function parseTop300XLSX(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().trim() === 'top300');
  if (!sheetName) {
    console.warn('[Top300] No "Top300" sheet found. Available sheets: ' + wb.SheetNames.join(', '));
    return [];
  }
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const items = [];
  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 5) continue;
    const area = (row[0] || '').toString().trim();
    const storeNumber = (row[1] || '').toString().trim();
    const storeName = (row[2] || '').toString().trim();
    const rankRaw = row[3];
    const rank = (rankRaw === '' || rankRaw == null) ? null : parseInt(rankRaw, 10);
    const sku = (row[4] || '').toString().trim();
    const desc = (row[5] || '').toString().trim();
    if (storeNumber && sku) items.push({ area, storeNumber, storeName, rank, sku, desc });
  }
  console.log('[Top300] Loaded ' + items.length + ' Top 300 entries');
  return items;
}


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
function buildAnalytics(rawRows, storeMap, catMap = {}) {
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
    lastTransferOut: 72,  // BU = Transfer Out
    lastTransferIn: 73    // BV = Transfer In
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
    // isOverstock computed below (after daysSince... calculations) to support 30-day stock-in exclusion
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
    const lastTransferIn = row[COL.lastTransferIn] || '';
    const lastTransferOut = row[COL.lastTransferOut] || '';
    const daysNoSales = daysSince(dateLastSold);
    const daysSinceReceived = daysSince(dateLastReceived);
    const daysSinceTransferIn = daysSince(lastTransferIn);
    // Most recent stock-in event (last received OR last transfer in, whichever is more recent)
    let daysSinceLastStockIn = null;
    if (daysSinceReceived != null && daysSinceTransferIn != null) daysSinceLastStockIn = Math.min(daysSinceReceived, daysSinceTransferIn);
    else if (daysSinceReceived != null) daysSinceLastStockIn = daysSinceReceived;
    else if (daysSinceTransferIn != null) daysSinceLastStockIn = daysSinceTransferIn;
    // Overstock: WTS > 12 weeks AND has stock AND last stock-in event was 30+ days ago (excludes fresh deliveries)
    const isOverstock = wtsNet > 12 && onHand > 0 && (daysSinceLastStockIn == null || daysSinceLastStockIn >= 30);
    // Per-SKU days cover (matches existing formula)
    const skuDaysCover = (wkAveNet > 0 && avgCost > 0) ? (onHandValue * 7) / (wkAveNet * avgCost) : null;
    // Aging: stock projected to last 180+ days AND last received 180+ days ago
    // (excludes recent deliveries which would otherwise look like aging due to high stock)
    const isAging = onHand > 0
      && skuDaysCover != null && skuDaysCover >= 180
      && daysSinceLastStockIn != null && daysSinceLastStockIn >= 180;
    // Black Inventory: on hand AND no sales 180+ days (or never sold) AND last received 180+ days ago
    const isBlackInventory = onHand > 0
      && (daysNoSales == null || daysNoSales >= 180)
      && daysSinceReceived != null && daysSinceReceived >= 180;
    // Negative SKU: on hand is negative (system error or data sync issue)
    const isNegativeStock = onHand < 0;

    enriched.push({
      regionCode: row[COL.regionCode] || '',
      regionName: storeInfo.region || row[COL.regionName] || '',
      storeNumber: storeId,
      storeName: storeInfo.storeName || row[COL.storeName] || '',
      area: storeInfo.area || '',
      dept: row[COL.dept] || '',
      deptName: row[COL.deptName] || '',
      catName: catMap[(row[COL.deptName] || '').toString().trim().toUpperCase()] || '',
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
      merchGro: (row[COL.merchGro] || '').toString().trim().toUpperCase(),
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
      skuDaysCover,
      supplierCode: row[COL.supplierCode] || '',
      supplierName: row[COL.supplierName] || '',
      delivMode: row[COL.delivMode] || '',
      dateLastSold,
      dateLastReceived,
      lastTransferIn,
      lastTransferOut,
      daysNoSales,
      lostSalesPerWeek,
      isCritical,
      isOverstock,
      isDeadStock,
      isAging,
      isBlackInventory,
      isNegativeStock,
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
  const agingCount = enriched.filter(r => r.isAging).length;
  const blackInventoryCount = enriched.filter(r => r.isBlackInventory).length;
  const negativeSkuCount = enriched.filter(r => r.isNegativeStock).length;
  const outOfStockCount = enriched.filter(r => r.isOutOfStock).length;
  // Value totals per category
  const overstockValue = enriched.filter(r => r.isOverstock).reduce((s, r) => s + r.onHandValue, 0);
  const agingValue = enriched.filter(r => r.isAging).reduce((s, r) => s + r.onHandValue, 0);
  const blackInventoryValue = enriched.filter(r => r.isBlackInventory).reduce((s, r) => s + r.onHandValue, 0);
  const deadStockValue = enriched.filter(r => r.isDeadStock).reduce((s, r) => s + r.onHandValue, 0);
  const totalLostSalesPerWeek = enriched.reduce((s, r) => s + r.lostSalesPerWeek, 0);
  const activeStores = new Set(enriched.map(r => r.storeNumber)).size;
  const activeSuppliers = new Set(enriched.map(r => r.supplierCode).filter(Boolean)).size;
  const totalPOValue = enriched.reduce((s, r) => s + r.poValue, 0);
  const totalTRFValue = enriched.reduce((s, r) => s + r.trfValue, 0);
  // Value-weighted Days Cover / Weeks-to-Sell (same formula used by Store & Supplier rollups)
  //   daysCover = (totalValue × 7) / Σ(wkAveNet × avgCost)
  //   avgWts    = daysCover / 7
  const totalWklSalesValue = enriched.reduce((s, r) => s + (r.wkAveNet * r.avgCost), 0);
  const daysCover = totalWklSalesValue > 0 ? (totalOnHandValue * 7) / totalWklSalesValue : 0;
  const avgWts = daysCover > 0 ? daysCover / 7 : 0;

  const kpis = {
    totalOnHandValue,
    totalOnHand,
    criticalCount,
    overstockCount,
    overstockValue,
    deadStockCount,
    deadStockValue,
    agingCount,
    agingValue,
    blackInventoryCount,
    blackInventoryValue,
    negativeSkuCount,
    outOfStockCount,
    totalLostSalesPerWeek,
    activeStores,
    activeSuppliers,
    totalPOValue,
    totalTRFValue,
    avgWts,
    daysCover,
    totalSKUs: enriched.length
  };

  // ── CRITICAL ITEMS ────────────────────────────────────────────────────────
  const criticalItems = enriched
    .filter(r => r.isCritical)
    .sort((a, b) => a.wtsNet - b.wtsNet)
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
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: r.totalPO > 0 ? 'PO Incoming' : r.p8ave > 0 ? 'URGENT: Place PO' : 'Review'
    }));

  // ── OVERSTOCK ITEMS ───────────────────────────────────────────────────────
  const overstockItems = enriched
    .filter(r => r.isOverstock)
    .sort((a, b) => b.wtsNet - a.wtsNet)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases,
      onHandValue: r.onHandValue,
      p8ave: r.p8ave,
      wtsNet: r.wtsNet === 999 ? 'Dead Stock' : r.wtsNet.toFixed(1),
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
      };
    });

  // ── AGING ─────────────────────────────────────────────────────────────────
  // Stock projected to last 180+ days AND sold within last 180 days
  const agingItems = enriched
    .filter(r => r.isAging)
    .sort((a, b) => (b.skuDaysCover || 0) - (a.skuDaysCover || 0))
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases,
      onHandValue: r.onHandValue,
      p8ave: r.p8ave,
      daysCover: r.skuDaysCover,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: 'For Stop Booking'
      };
    });

  // ── BLACK INVENTORY ───────────────────────────────────────────────────────
  // OnHand > 0, no sales 180+ days (or never sold), and last received 180+ days ago
  const blackInventoryItems = enriched
    .filter(r => r.isBlackInventory)
    .sort((a, b) => b.onHandValue - a.onHandValue)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
      store: `${r.storeNumber} - ${r.storeName}`,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases,
      onHandValue: r.onHandValue,
      p8ave: r.p8ave,
      daysCover: r.skuDaysCover,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: 'Investigate / Liquidate'
      };
    });

  // ── NEGATIVE SKU ──────────────────────────────────────────────────────────
  // On hand is negative (system error or data sync issue) — needs investigation
  const negativeSkuItems = enriched
    .filter(r => r.isNegativeStock)
    .sort((a, b) => a.onHand - b.onHand)  // most negative first
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = (r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        storeName: `${r.storeNumber} - ${r.storeName}`,
        area: r.area,
        catName: r.catName || 'Uncategorized',
        skuCode: r.skuCode,
        skuDesc: r.skuDesc,
        supplier: r.supplierName,
        onHand: r.onHand,
        qtyCases,
        onHandValue: r.onHandValue,
        p8ave: r.p8ave,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn),
        lastTransferOut: formatDate(r.lastTransferOut)
      };
    });

  // ── DEAD STOCK ────────────────────────────────────────────────────────────
  const deadStockItems = enriched
    .filter(r => r.isDeadStock)
    .sort((a, b) => b.onHandValue - a.onHandValue)
    .map(r => {
      const wtsItem = r.p8ave > 0 ? r.onHand / r.p8ave : null;
      const dcItem = (r.wkAveNet > 0 && r.avgCost > 0) ? (r.onHandValue * 7) / (r.wkAveNet * r.avgCost) : null;
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        store: `${r.storeNumber} - ${r.storeName}`,
        area: r.area,
        skuCode: r.skuCode,
        skuDesc: r.skuDesc,
        supplier: r.supplierName,
        onHand: r.onHand,
        qtyCases,
        onHandValue: r.onHandValue,
        weeksToSell: wtsItem,
        daysCover: dcItem,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn),
        lastTransferOut: formatDate(r.lastTransferOut),
        action: 'No Sales 8 Wks - Review/Markdown'
      };
    });

  // ── OUT OF STOCK ITEMS (Lost Sales) ───────────────────────────────────────
  const outOfStockItems = enriched
    .filter(r => r.isOutOfStock)
    .sort((a, b) => b.lostSalesPerWeek - a.lostSalesPerWeek)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        storeNumber: r.storeNumber,
        storeName: r.storeName,
        area: r.area,
        skuCode: r.skuCode,
        skuDesc: r.skuDesc,
        supplier: r.supplierName,
        onHand: r.onHand,
        stdPack: r.stdPack,
        qtyCases,
        invValue: r.onHandValue,
        p8ave: r.p8ave,
        weeksToSell: r.skuWTS != null ? +r.skuWTS.toFixed(2) : null,
        daysCover: r.skuDaysCover != null ? Math.round(r.skuDaysCover) : null,
        status: 'OOS',
        lostSalesPerWeek: r.lostSalesPerWeek,
        ico: r.ico,
        poOrderGR: r.poOrderGR,
        trfOrderGR: r.trfOrderGR,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn),
        lastTransferOut: formatDate(r.lastTransferOut)
      };
    });

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
    // Track non-P (excluded from OOS/Critical) SKU count for accurate percentages
    if (r.merchGro !== 'P') g.nonPCount = (g.nonPCount || 0) + 1;
    // OOS & Critical exclude items with Merchandise Gro = 'P' (per business rule)
    if (r.isCritical && r.merchGro !== 'P') g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock && r.merchGro !== 'P') g.oosCount++;
  }
  // Compute risk percentages and days cover
  // Days Cover = OnHand Value / (Weekly Sales Net WS × Avg Cost / 7) = BC / (AU × AX / 7)
  const storeAnalysis = Object.values(storeGroups).map(g => {
    const total = g.totalSKUs || 1;
    const nonP = g.nonPCount || 1; // for OOS/Critical %
    g.criticalPct = (g.criticalCount / nonP) * 100;
    g.oosPct = (g.oosCount / nonP) * 100;
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
        totalWklSalesValue: 0, totalP8Ave: 0
      };
    }
    const g = supplierGroups[key];
    g.totalValue += r.onHandValue;
    g.totalOnHand += r.onHand;
    g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    g.totalLostSales += r.lostSalesPerWeek;
    g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
    g.totalP8Ave += r.p8ave;
    if (r.merchGro !== 'P') g.nonPCount = (g.nonPCount || 0) + 1;
    // OOS & Critical exclude items with Merchandise Gro = 'P'
    if (r.isCritical && r.merchGro !== 'P') g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock && r.merchGro !== 'P') g.oosCount++;
  }
  const supplierAnalysis = Object.values(supplierGroups).map(g => {
    const total = g.totalSKUs || 1;
    const nonP = g.nonPCount || 1;
    g.criticalPct = (g.criticalCount / nonP) * 100;
    g.oosPct = (g.oosCount / nonP) * 100;
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
    skuStatuses: uniq(enriched.map(r => r.skuStatus)),
    categories: uniq(enriched.map(r => r.catName)).filter(d => d.length > 0)
  };

  return { kpis, criticalItems, overstockItems, agingItems, blackInventoryItems, negativeSkuItems, deadStockItems, outOfStockItems, storeAnalysis, supplierAnalysis, filterMeta, rows: enriched };
}

// Retry a network operation on transient errors (premature close, ECONNRESET, ETIMEDOUT, 5xx).
async function retryNet(label, fn, maxAttempts = 3) {
  let attempt = 0;
  let lastErr;
  while (attempt < maxAttempts) {
    attempt++;
    try {
      return await fn();
    } catch (e) {
      lastErr = e;
      const msg = (e && e.message) || '';
      const code = (e && e.code) || '';
      const status = e && e.response && e.response.status;
      const transient = /premature close|ECONNRESET|ETIMEDOUT|EAI_AGAIN|ENETUNREACH|socket hang up|network|EPIPE/i.test(msg)
        || /ECONNRESET|ETIMEDOUT|EAI_AGAIN|ENETUNREACH|EPIPE/i.test(code)
        || (status && status >= 500 && status < 600);
      if (!transient || attempt >= maxAttempts) throw e;
      const wait = 1000 * Math.pow(2, attempt - 1); // 1s, 2s, 4s
      console.warn(`[Cache] ${label} failed (attempt ${attempt}/${maxAttempts}): ${msg}. Retrying in ${wait}ms...`);
      await new Promise(r => setTimeout(r, wait));
    }
  }
  throw lastErr;
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

    const drive = await getDriveClient();

    // Find InvData.csv
    const invFile = await retryNet('findFile(' + INV_FILE_NAME + ')', () => findFile(drive, INV_FILE_NAME));
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
    const invBuffer = await retryNet('download(' + INV_FILE_NAME + ')', () => downloadFileBuffer(drive, invFile.id));

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
    let catMap = {};
    let upcMap = {};
    let top300 = [];
    try {
      console.log('[Cache] Looking for ' + STORES_FILE_NAME + ' in folder ' + GDRIVE_FOLDER_ID);
      const storesFile = await retryNet('findFile(' + STORES_FILE_NAME + ')', () => findFile(drive, STORES_FILE_NAME));
      if (storesFile) {
        console.log('[Cache] Found stores file ID: ' + storesFile.id + ', downloading...');
        const storesBuffer = await retryNet('download(' + STORES_FILE_NAME + ')', () => downloadFileBuffer(drive, storesFile.id));
        storeMap = parseStoresXLSX(storesBuffer);
        usersMap = parseUsersXLSX(storesBuffer);
        catMap = parseCatCodeXLSX(storesBuffer);
        upcMap = parseUPCXLSX(storesBuffer);
        top300 = parseTop300XLSX(storesBuffer);
        console.log(`[Cache] Loaded ${Object.keys(storeMap).length} stores, ${Object.keys(usersMap).length} users, ${Object.keys(catMap).length} categories, ${Object.keys(upcMap).length} UPCs, ${top300.length} Top300 entries.`);
      } else {
        console.warn('[Cache] STORES FILE NOT FOUND! Searched name: "' + STORES_FILE_NAME + '" in folder: ' + GDRIVE_FOLDER_ID);
        // List all files in folder for debugging
        const allFiles = await driveList(
          `'${GDRIVE_FOLDER_ID}' in parents and trashed=false`,
          'files(id,name,mimeType)',
          20
        );
        console.warn('[Cache] Files in folder:');
        (allFiles.files || []).forEach(f => console.warn('  - "' + f.name + '" (type: ' + f.mimeType + ')'));
      }
    } catch (e) {
      console.warn('[Cache] Could not load stores file:', e.message);
    }

    console.log('[Cache] Parsing CSV...');
    const rawRows = await parseCSV(invBuffer);
    console.log(`[Cache] Parsed ${rawRows.length} rows.`);

    console.log('[Cache] Building analytics...');
    const analytics = buildAnalytics(rawRows, storeMap, catMap);
    if (!analytics) throw new Error('Analytics build failed - no data.');

    // Atomic swap
    cache.rows = analytics.rows;
    cache.storeMap = storeMap;
    if (Object.keys(usersMap).length > 0) cache.users = usersMap;
    if (Object.keys(catMap).length > 0) cache.catMap = catMap;
    if (Object.keys(upcMap).length > 0) cache.upcMap = upcMap;
    if (top300.length > 0) cache.top300 = top300;
    // Build fast store_sku lookup index for Top 300 join (and any future cross-references)
    const storeSkuIndex = {};
    for (const r of analytics.rows) {
      const key = (r.storeNumber || '').toString().trim() + '_' + (r.skuCode || '').toString().trim();
      storeSkuIndex[key] = r;
    }
    cache.storeSkuIndex = storeSkuIndex;
    cache.kpis = analytics.kpis;
    cache.criticalItems = analytics.criticalItems;
    cache.overstockItems = analytics.overstockItems;
    cache.agingItems = analytics.agingItems;
    cache.blackInventoryItems = analytics.blackInventoryItems;
    cache.negativeSkuItems = analytics.negativeSkuItems;
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
    if (filters.category && r.catName !== filters.category) return false;
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
    created: Date.now(),
    lastActivity: Date.now()
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
  const reason = (req.body && req.body.reason) || 'manual';
  const s = sessions[token];
  if (s) {
    if (s.logRow && s.loginTimeISO) {
      logLogoutEvent(s.logRow, s.loginTimeISO, reason);
    }
    delete sessions[token];
  }
  res.json({ ok: true });
});

// Heartbeat - client pings every minute to keep session alive
app.post('/api/heartbeat', (req, res) => {
  const token = (req.body && req.body.token) || '';
  const s = sessions[token];
  if (!s) return res.status(401).json({ error: 'Not logged in' });
  s.lastActivity = Date.now();
  res.json({ ok: true });
});

// Background task: check for inactive sessions every 60 seconds and auto-logout
const INACTIVITY_TIMEOUT_MS = 10 * 60 * 1000; // 10 minutes
setInterval(() => {
  const now = Date.now();
  for (const token of Object.keys(sessions)) {
    const s = sessions[token];
    const lastActive = s.lastActivity || s.created;
    if (now - lastActive > INACTIVITY_TIMEOUT_MS) {
      console.log('[Session] Auto-logout for inactive user: ' + s.username);
      if (s.logRow && s.loginTimeISO) {
        logLogoutEvent(s.logRow, s.loginTimeISO, 'auto-timeout');
      }
      delete sessions[token];
    }
  }
}, 60 * 1000);

app.get('/api/me', (req, res) => {
  const token = req.query.token || '';
  const s = sessions[token];
  if (!s) return res.status(401).json({ error: 'Not logged in' });
  res.json({ username: s.username, level: s.level, area: s.area, isAdmin: s.isAdmin });
});

// UPC barcode lookup — returns the matching SKU & description
app.get('/api/upc-lookup', (req, res) => {
  const upc = (req.query.upc || '').toString().trim();
  if (!upc) return res.status(400).json({ error: 'UPC required' });
  if (!cache.upcMap || Object.keys(cache.upcMap).length === 0) {
    return res.status(503).json({ error: 'UPC database not loaded yet' });
  }
  const entry = cache.upcMap[upc];
  if (!entry) return res.status(404).json({ error: 'UPC not found', upc });
  res.json({ upc, sku: entry.sku, desc: entry.desc });
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

// Activity Log XLSX export (admin only)
app.get('/api/export-logs-xlsx', async (req, res) => {
  if (!requireAdmin(req, res)) return;
  const logs = await readLogs();
  logs.reverse();
  const data = logs.map(r => ({
    user: r.user,
    loginTime: r.loginTime,
    logoutTime: r.logoutTime,
    duration: r.duration,
    area: r.area
  }));
  const columns = [
    { header: 'User', key: 'user' },
    { header: 'Login Time', key: 'loginTime' },
    { header: 'Logout Time', key: 'logoutTime' },
    { header: 'Duration', key: 'duration' },
    { header: 'Area', key: 'area' }
  ];
  if (data.length === 0) return res.status(204).send('No logs');
  const filename = `Activity_Logs_${new Date().toISOString().split('T')[0]}.xlsx`;
  await writeStyledWorkbook(res, {
    sheetName: 'Activity Logs',
    title: 'Activity Logs — CAMANAVA Inventory Dashboard',
    columns, rows: data, filename
  });
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
  const agingCount = filtered.filter(r => r.isAging).length;
  const blackInventoryCount = filtered.filter(r => r.isBlackInventory).length;
  const outOfStockCount = filtered.filter(r => r.isOutOfStock).length;
  const overstockValue = filtered.filter(r => r.isOverstock).reduce((s, r) => s + r.onHandValue, 0);
  const agingValue = filtered.filter(r => r.isAging).reduce((s, r) => s + r.onHandValue, 0);
  const blackInventoryValue = filtered.filter(r => r.isBlackInventory).reduce((s, r) => s + r.onHandValue, 0);
  const deadStockValue = filtered.filter(r => r.isDeadStock).reduce((s, r) => s + r.onHandValue, 0);
  const totalLostSalesPerWeek = filtered.reduce((s, r) => s + r.lostSalesPerWeek, 0);
  // Value-weighted: daysCover = (totalValue × 7) / Σ(wkAveNet × avgCost); avgWts = daysCover / 7
  const totalWklSalesValue = filtered.reduce((s, r) => s + (r.wkAveNet * r.avgCost), 0);
  const daysCover = totalWklSalesValue > 0 ? (totalOnHandValue * 7) / totalWklSalesValue : 0;
  const avgWts = daysCover > 0 ? daysCover / 7 : 0;
  res.json({
    totalOnHandValue, totalOnHand, criticalCount, overstockCount, deadStockCount,
    agingCount, blackInventoryCount,
    overstockValue, agingValue, blackInventoryValue, deadStockValue,
    outOfStockCount, totalLostSalesPerWeek,
    activeStores: new Set(filtered.map(r => r.storeNumber)).size,
    activeSuppliers: new Set(filtered.map(r => r.supplierCode).filter(Boolean)).size,
    totalPOValue: filtered.reduce((s, r) => s + r.poValue, 0),
    totalTRFValue: filtered.reduce((s, r) => s + r.trfValue, 0),
    avgWts,
    daysCover,
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
    .sort((a, b) => a.wtsNet - b.wtsNet)
    .map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue,
      currentWkSales: r.currentWkSales, p8ave: r.p8ave,
      wtsNet: r.wtsNet, totalPO: r.totalPO,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: r.totalPO > 0 ? 'PO Incoming' : r.p8ave > 0 ? 'URGENT: Place PO' : 'Review'
    }));
  res.json(filtered);
});

app.get('/api/overstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.overstockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isOverstock)
    .sort((a, b) => b.wtsNet - a.wtsNet)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, qtyCases, onHandValue: r.onHandValue, p8ave: r.p8ave,
      wtsNet: r.wtsNet === 999 ? 'Dead Stock' : r.wtsNet.toFixed(1),
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
      };
    });
  res.json(filtered);
});

app.get('/api/aging', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.agingItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isAging)
    .sort((a, b) => (b.skuDaysCover || 0) - (a.skuDaysCover || 0))
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, qtyCases, onHandValue: r.onHandValue, p8ave: r.p8ave,
      daysCover: r.skuDaysCover,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: 'For Stop Booking'
      };
    });
  res.json(filtered);
});

app.get('/api/blackinventory', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.blackInventoryItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isBlackInventory)
    .sort((a, b) => b.onHandValue - a.onHandValue)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
      skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, qtyCases, onHandValue: r.onHandValue, p8ave: r.p8ave,
      daysCover: r.skuDaysCover,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut),
      action: 'Investigate / Liquidate'
      };
    });
  res.json(filtered);
});

app.get('/api/negativeskus', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.negativeSkuItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isNegativeStock)
    .sort((a, b) => a.onHand - b.onHand)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = (r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        storeName: `${r.storeNumber} - ${r.storeName}`,
        area: r.area,
        catName: r.catName || 'Uncategorized',
        skuCode: r.skuCode,
        skuDesc: r.skuDesc,
        supplier: r.supplierName,
        onHand: r.onHand,
        qtyCases,
        onHandValue: r.onHandValue,
        p8ave: r.p8ave,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn),
        lastTransferOut: formatDate(r.lastTransferOut)
      };
    });
  res.json(filtered);
});

// Top 300 SKUs — joined with InvData via storeNumber + skuCode
app.get('/api/top300skus', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  const top300 = cache.top300 || [];
  const idx = cache.storeSkuIndex || {};

  // Enforce area lock at the top300 level (since filters.area might be locked)
  const out = [];
  for (const t of top300) {
    // Apply area filter against the top300 entry's own area
    if (filters.area && t.area !== filters.area) continue;
    if (filters.store && t.storeNumber !== filters.store) continue;
    const key = (t.storeNumber || '').trim() + '_' + (t.sku || '').trim();
    const inv = idx[key];
    // Apply remaining filters against the joined InvData row
    if (filters.category) {
      if (!inv || inv.catName !== filters.category) continue;
    }
    if (filters.dept && (!inv || inv.deptName !== filters.dept)) continue;
    if (filters.subDept && (!inv || inv.subDeptName !== filters.subDept)) continue;
    if (filters.cls && (!inv || inv.clsName !== filters.cls)) continue;
    if (filters.supplier && (!inv || inv.supplierName !== filters.supplier)) continue;

    // Compute status
    let status = 'Normal';
    if (inv) {
      if (inv.isCritical) status = 'Critical';
      else if (inv.isOutOfStock) status = 'OOS';
      else if (inv.isOverstock) status = 'Overstock';
      else if (inv.isDeadStock) status = 'Dead Stock';
    } else {
      status = 'Not Found';
    }

    out.push({
      area: t.area,
      storeName: t.storeName ? (t.storeNumber + ' - ' + t.storeName) : (inv ? `${inv.storeNumber} - ${inv.storeName}` : t.storeNumber),
      rank: t.rank,
      sku: t.sku,
      itemDescription: t.desc || (inv ? inv.skuDesc : ''),
      supplier: inv ? inv.supplierName : '',
      onHand: inv ? inv.onHand : null,
      qtyCases: inv ? (inv.stdPack > 0 && inv.stdPack === inv.onHand ? 'Per Piece' : (inv.stdPack > 0 && inv.onHand !== 0 ? +(inv.onHand / inv.stdPack).toFixed(2) : 'Per Piece')) : null,
      p8ave: inv ? inv.p8ave : null,
      daysCover: inv ? inv.skuDaysCover : null,
      status,
      incomingPO: inv ? inv.poOrderGR : null,
      lostSalesPerWeek: inv ? inv.lostSalesPerWeek : null,
      ico: inv ? inv.ico : '',
      dateLastSold: inv ? formatDate(inv.dateLastSold) : '',
      dateLastReceived: inv ? formatDate(inv.dateLastReceived) : '',
      lastTransferIn: inv ? formatDate(inv.lastTransferIn) : '',
      lastTransferOut: inv ? formatDate(inv.lastTransferOut) : ''
    });
  }
  res.json(out);
});

// Top 300 SKU Store Metrics — KPI per store ranked best to worst
// Score formula: (OOS% × 0.6) + (Critical% × 0.4). Lower is better.
app.get('/api/top300-store-metrics', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  const top300 = cache.top300 || [];
  const idx = cache.storeSkuIndex || {};

  // Group by store
  const byStore = {};
  for (const t of top300) {
    if (filters.area && t.area !== filters.area) continue;
    if (filters.store && t.storeNumber !== filters.store) continue;
    const key = (t.storeNumber || '').trim() + '_' + (t.sku || '').trim();
    const inv = idx[key];
    if (filters.category && (!inv || inv.catName !== filters.category)) continue;
    if (filters.supplier && (!inv || inv.supplierName !== filters.supplier)) continue;

    const storeKey = t.storeNumber;
    if (!byStore[storeKey]) {
      byStore[storeKey] = {
        storeNumber: t.storeNumber,
        storeName: t.storeName || (inv ? inv.storeName : ''),
        area: t.area || (inv ? inv.area : ''),
        total: 0, oos: 0, critical: 0, overstock: 0, deadStock: 0, notFound: 0, healthy: 0
      };
    }
    const s = byStore[storeKey];
    s.total++;
    if (!inv) s.notFound++;
    else if (inv.isOutOfStock) s.oos++;
    else if (inv.isCritical) s.critical++;
    else if (inv.isDeadStock) s.deadStock++;
    else if (inv.isOverstock) s.overstock++;
    else s.healthy++;
  }

  // Calculate percentages and combined score (lower = better)
  const result = Object.values(byStore).map(s => {
    const oosPct = s.total > 0 ? (s.oos / s.total) * 100 : 0;
    const criticalPct = s.total > 0 ? (s.critical / s.total) * 100 : 0;
    const healthyPct = s.total > 0 ? (s.healthy / s.total) * 100 : 0;
    // Treat Not Found as missing data (not a stockout) — counts toward total but excluded from problem score
    const score = (oosPct * 0.6) + (criticalPct * 0.4);
    return { ...s, oosPct, criticalPct, healthyPct, score };
  });

  // Sort ascending (best stores first)
  result.sort((a, b) => a.score - b.score);
  result.forEach((r, i) => r.rank = i + 1);
  res.json(result);
});

app.get('/api/deadstock', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  if (Object.keys(filters).length === 0) return res.json(cache.deadStockItems);
  const filtered = applyFilters(cache.rows, filters).filter(r => r.isDeadStock)
    .sort((a, b) => b.onHandValue - a.onHandValue)
    .map(r => {
      const wtsItem = r.p8ave > 0 ? r.onHand / r.p8ave : null;
      const dcItem = (r.wkAveNet > 0 && r.avgCost > 0) ? (r.onHandValue * 7) / (r.wkAveNet * r.avgCost) : null;
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        store: `${r.storeNumber} - ${r.storeName}`, area: r.area,
        skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
        onHand: r.onHand, qtyCases, onHandValue: r.onHandValue,
        weeksToSell: wtsItem,
        daysCover: dcItem,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn),
        lastTransferOut: formatDate(r.lastTransferOut),
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
    .sort((a, b) => b.lostSalesPerWeek - a.lostSalesPerWeek)
    .map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        storeNumber: r.storeNumber, storeName: r.storeName, area: r.area,
        skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
        onHand: r.onHand, stdPack: r.stdPack, qtyCases, invValue: r.onHandValue,
        p8ave: r.p8ave,
        weeksToSell: r.skuWTS != null ? +r.skuWTS.toFixed(2) : null,
        daysCover: r.skuDaysCover != null ? Math.round(r.skuDaysCover) : null,
        status: 'OOS', lostSalesPerWeek: r.lostSalesPerWeek,
        ico: r.ico, poOrderGR: r.poOrderGR, trfOrderGR: r.trfOrderGR,
        dateLastSold: formatDate(r.dateLastSold),
        dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn),
        lastTransferOut: formatDate(r.lastTransferOut)
      };
    });
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
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut)
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
    if (r.merchGro !== 'P') g.nonPCount = (g.nonPCount || 0) + 1;
    if (r.isCritical && r.merchGro !== 'P') g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock && r.merchGro !== 'P') g.oosCount++;
  }
  const result = Object.values(storeGroups).map(g => {
    const total = g.totalSKUs || 1;
    const nonP = g.nonPCount || 1;
    g.criticalPct = (g.criticalCount / nonP) * 100;
    g.oosPct = (g.oosCount / nonP) * 100;
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
    if (!supplierGroups[key]) supplierGroups[key] = { supplierCode: r.supplierCode, supplierName: r.supplierName, totalValue: 0, totalOnHand: 0, criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0, totalSKUs: 0, totalSales: 0, totalLostSales: 0, totalWklSalesValue: 0, totalP8Ave: 0 };
    const g = supplierGroups[key];
    g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
    g.totalSales += r.currentWkSales;
    g.totalLostSales += r.lostSalesPerWeek;
    g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
    g.totalP8Ave += r.p8ave;
    if (r.merchGro !== 'P') g.nonPCount = (g.nonPCount || 0) + 1;
    if (r.isCritical && r.merchGro !== 'P') g.criticalCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
    if (r.isOutOfStock && r.merchGro !== 'P') g.oosCount++;
  }
  const result = Object.values(supplierGroups).map(g => {
    const total = g.totalSKUs || 1;
    const nonP = g.nonPCount || 1;
    g.criticalPct = (g.criticalCount / nonP) * 100;
    g.oosPct = (g.oosCount / nonP) * 100;
    g.overstockPct = (g.overstockCount / total) * 100;
    g.deadPct = (g.deadCount / total) * 100;
    g.daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
    g.weeksToSell = g.daysCover != null ? g.daysCover / 7 : null;
    return g;
  }).sort((a, b) => b.totalValue - a.totalValue).slice(0, 100);
  res.json(result);
});

// Category aggregates for the Overview chart
app.get('/api/categories', (req, res) => {
  if (!cache.ready) return res.json({ error: 'Cache not ready' });
  const filters = resolveFilters(req);
  // Don't filter by category for this endpoint (it would zero out other categories)
  delete filters.category;
  const rows = applyFilters(cache.rows, filters);
  const catGroups = {};
  for (const r of rows) {
    if (!r.catName) continue;
    if (!catGroups[r.catName]) catGroups[r.catName] = { catName: r.catName, totalValue: 0, totalSKUs: 0, criticalCount: 0, oosCount: 0, overstockCount: 0, deadCount: 0 };
    const g = catGroups[r.catName];
    g.totalValue += r.onHandValue;
    g.totalSKUs++;
    if (r.isCritical) g.criticalCount++;
    if (r.isOutOfStock) g.oosCount++;
    if (r.isOverstock) g.overstockCount++;
    if (r.isDeadStock) g.deadCount++;
  }
  res.json(Object.values(catGroups).sort((a, b) => b.totalValue - a.totalValue));
});

app.post('/api/refresh', async (req, res) => {
  refreshData(true);
  res.json({ message: 'Refresh triggered' });
});

// ─── EXPORT CSV ───────────────────────────────────────────────────────────────
// ─── STYLED XLSX HELPER ───────────────────────────────────────────────────────
// Generates a dark green header workbook and streams to response.
// columns: [{ header, key, width?, format? ('currency'|'integer'|'decimal'|'date'|'text'|'percent') }]
async function writeStyledWorkbook(res, opts) {
  const { sheetName, title, columns, rows, filename } = opts;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CAMANAVA Inventory Dashboard';
  wb.created = new Date();
  const ws = wb.addWorksheet(sheetName || 'Data');
  const DARK_GREEN = 'FF1B5E20';

  // Title row (merged)
  ws.mergeCells(1, 1, 1, columns.length);
  const titleCell = ws.getCell(1, 1);
  // Format: "{Sheet Name} Report as of {Month Day, Year}" — e.g., "Out of Stock Report as of June 27, 2026"
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const baseTitle = title || sheetName;
  // Strip any trailing "—" segment from passed-in title; use sheetName as the report subject
  const reportName = sheetName || baseTitle;
  titleCell.value = reportName + ' Report as of ' + today;
  titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: 'Calibri' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(1).height = 26;

  // Set columns + header row
  ws.columns = columns.map(c => ({ header: c.header, key: c.key, width: c.width || Math.max(12, c.header.length + 2) }));
  const headerRow = ws.getRow(2);
  columns.forEach((c, i) => { headerRow.getCell(i + 1).value = c.header; });
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  });
  headerRow.height = 22;

  // Number formats per column type
  const numFmtMap = {
    currency: '"₱"#,##0.00',
    decimal: '#,##0.00',
    integer: '#,##0',
    percent: '0.0"%"',
    date: 'mm/dd/yyyy',
    text: '@'
  };

  // Add data rows + auto-size columns based on content
  const maxWidths = columns.map(c => c.header.length);
  const dateRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/;
  for (const r of rows) {
    const rowObj = {};
    columns.forEach((c, i) => {
      let v = r[c.key];
      // Date type: parse "MM/DD/YYYY" string to Date so Excel formats it as date
      if (c.format === 'date' && typeof v === 'string' && v) {
        const m = v.match(dateRegex);
        if (m) {
          let [_, mm, dd, yy] = m;
          if (yy.length === 2) yy = '20' + yy;
          v = new Date(parseInt(yy), parseInt(mm) - 1, parseInt(dd));
        }
      } else if ((c.format === 'currency' || c.format === 'decimal' || c.format === 'integer' || c.format === 'percent') && v != null && v !== '') {
        v = Number(v);
        if (!Number.isFinite(v)) v = null;
      }
      rowObj[c.key] = v;
      const dispLen = (v == null) ? 0 : (v instanceof Date ? 10 : String(v).length);
      if (dispLen > maxWidths[i]) maxWidths[i] = dispLen;
    });
    ws.addRow(rowObj);
  }

  // Apply number formats and auto-adjust widths (capped at 50)
  ws.columns.forEach((col, i) => {
    const c = columns[i];
    if (c && c.format && numFmtMap[c.format]) col.numFmt = numFmtMap[c.format];
    if (!c.width) col.width = Math.min(50, Math.max(10, maxWidths[i] + 2));
  });

  // Freeze title + header, enable autoFilter
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: columns.length } };

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  await wb.xlsx.write(res);
  res.end();
}

// Apply tab search across common displayed fields
function applyTableSearch(rows, search) {
  if (!search) return rows;
  const q = search.toLowerCase().trim();
  return rows.filter(r => {
    return Object.values(r).some(v => v != null && String(v).toLowerCase().includes(q));
  });
}

// Unified styled XLSX export for all tabs (replaces the old CSV endpoint)
app.get('/api/export/:type', async (req, res) => {
  if (!cache.ready) return res.status(503).send('Cache not ready');
  const type = req.params.type;
  const filters = resolveFilters(req);
  const search = req.query.search || '';
  const today = new Date().toISOString().split('T')[0];
  const baseRows = applyFilters(cache.rows, filters);

  // Build (rows, columns, sheetName, title) per type
  let data, columns, sheetName, title;

  if (type === 'critical') {
    sheetName = 'Critical';
    title = 'Critical Stock — CAMANAVA Inventory';
    data = baseRows.filter(r => r.isCritical).sort((a, b) => a.wtsNet - b.wtsNet).map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area, skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand, onHandValue: r.onHandValue, currentWkSales: r.currentWkSales, p8ave: r.p8ave, wtsNet: r.wtsNet,
      totalPO: r.totalPO, dateLastSold: formatDate(r.dateLastSold), dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn), lastTransferOut: formatDate(r.lastTransferOut),
      action: r.poOrderGR > 0 || r.trfOrderGR > 0 ? 'PO Incoming' : 'URGENT: Place PO'
    }));
    columns = [
      { header: 'Store', key: 'store' }, { header: 'Area', key: 'area' },
      { header: 'SKU Code', key: 'skuCode' }, { header: 'Description', key: 'skuDesc' }, { header: 'Supplier', key: 'supplier' },
      { header: 'On Hand', key: 'onHand', format: 'integer' }, { header: 'Inv Value', key: 'onHandValue', format: 'currency' },
      { header: 'Cur Wk Sales', key: 'currentWkSales', format: 'integer' }, { header: 'P8 Ave', key: 'p8ave', format: 'decimal' },
      { header: 'WTS Net', key: 'wtsNet', format: 'decimal' }, { header: 'Total PO', key: 'totalPO', format: 'integer' },
      { header: 'Last Sold', key: 'dateLastSold', format: 'date' }, { header: 'Last Received', key: 'dateLastReceived', format: 'date' },
      { header: 'Transfer In', key: 'lastTransferIn', format: 'date' }, { header: 'Transfer Out', key: 'lastTransferOut', format: 'date' },
      { header: 'Action', key: 'action' }
    ];
  }
  else if (type === 'overstock') {
    sheetName = 'Overstock';
    title = 'Overstock Items — CAMANAVA Inventory';
    data = baseRows.filter(r => r.isOverstock).sort((a, b) => b.wtsNet - a.wtsNet).map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area, skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases: r.stdPack > 0 && r.stdPack === r.onHand ? 'Per Piece' : (r.stdPack > 0 && r.onHand !== 0 ? +(r.onHand / r.stdPack).toFixed(2) : 'Per Piece'),
      onHandValue: r.onHandValue, p8ave: r.p8ave, wtsNet: r.wtsNet,
      dateLastSold: formatDate(r.dateLastSold), dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn), lastTransferOut: formatDate(r.lastTransferOut),
      action: r.wtsNet > 26 ? 'Consider Markdown' : 'Monitor / Transfer'
    }));
    columns = [
      { header: 'Store', key: 'store' }, { header: 'Area', key: 'area' },
      { header: 'SKU Code', key: 'skuCode' }, { header: 'Description', key: 'skuDesc' }, { header: 'Supplier', key: 'supplier' },
      { header: 'On Hand', key: 'onHand', format: 'integer' }, { header: 'Qty in Cases', key: 'qtyCases' },
      { header: 'Inv Value', key: 'onHandValue', format: 'currency' },
      { header: 'P8 Ave', key: 'p8ave', format: 'decimal' }, { header: 'WTS Net', key: 'wtsNet', format: 'decimal' },
      { header: 'Last Sold', key: 'dateLastSold', format: 'date' }, { header: 'Last Received', key: 'dateLastReceived', format: 'date' },
      { header: 'Transfer In', key: 'lastTransferIn', format: 'date' }, { header: 'Transfer Out', key: 'lastTransferOut', format: 'date' },
      { header: 'Action', key: 'action' }
    ];
  }
  else if (type === 'aging') {
    sheetName = 'Aging';
    title = 'Aging Items — CAMANAVA Inventory';
    data = baseRows.filter(r => r.isAging).sort((a, b) => (b.skuDaysCover || 0) - (a.skuDaysCover || 0)).map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area, skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases: r.stdPack > 0 && r.stdPack === r.onHand ? 'Per Piece' : (r.stdPack > 0 && r.onHand !== 0 ? +(r.onHand / r.stdPack).toFixed(2) : 'Per Piece'),
      onHandValue: r.onHandValue, p8ave: r.p8ave, daysCover: r.skuDaysCover != null ? Math.round(r.skuDaysCover) : null,
      dateLastSold: formatDate(r.dateLastSold), dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn), lastTransferOut: formatDate(r.lastTransferOut),
      action: 'For Stop Booking'
    }));
    columns = [
      { header: 'Store', key: 'store' }, { header: 'Area', key: 'area' },
      { header: 'SKU Code', key: 'skuCode' }, { header: 'Description', key: 'skuDesc' }, { header: 'Supplier', key: 'supplier' },
      { header: 'On Hand', key: 'onHand', format: 'integer' }, { header: 'Qty in Cases', key: 'qtyCases' },
      { header: 'Inv Value', key: 'onHandValue', format: 'currency' },
      { header: 'P8 Ave', key: 'p8ave', format: 'decimal' }, { header: 'Days Cover', key: 'daysCover', format: 'integer' },
      { header: 'Last Sold', key: 'dateLastSold', format: 'date' }, { header: 'Last Received', key: 'dateLastReceived', format: 'date' },
      { header: 'Transfer In', key: 'lastTransferIn', format: 'date' }, { header: 'Transfer Out', key: 'lastTransferOut', format: 'date' },
      { header: 'Action', key: 'action' }
    ];
  }
  else if (type === 'blackinventory') {
    sheetName = 'Black Inventory';
    title = 'Black Inventory — CAMANAVA Inventory';
    data = baseRows.filter(r => r.isBlackInventory).sort((a, b) => b.onHandValue - a.onHandValue).map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area, skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases: r.stdPack > 0 && r.stdPack === r.onHand ? 'Per Piece' : (r.stdPack > 0 && r.onHand !== 0 ? +(r.onHand / r.stdPack).toFixed(2) : 'Per Piece'),
      onHandValue: r.onHandValue, p8ave: r.p8ave, daysCover: r.skuDaysCover != null ? Math.round(r.skuDaysCover) : null,
      dateLastSold: formatDate(r.dateLastSold), dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn), lastTransferOut: formatDate(r.lastTransferOut),
      action: 'Investigate / Liquidate'
    }));
    columns = [
      { header: 'Store', key: 'store' }, { header: 'Area', key: 'area' },
      { header: 'SKU Code', key: 'skuCode' }, { header: 'Description', key: 'skuDesc' }, { header: 'Supplier', key: 'supplier' },
      { header: 'On Hand', key: 'onHand', format: 'integer' }, { header: 'Qty in Cases', key: 'qtyCases' },
      { header: 'Inv Value', key: 'onHandValue', format: 'currency' },
      { header: 'P8 Ave', key: 'p8ave', format: 'decimal' }, { header: 'Days Cover', key: 'daysCover', format: 'integer' },
      { header: 'Last Sold', key: 'dateLastSold', format: 'date' }, { header: 'Last Received', key: 'dateLastReceived', format: 'date' },
      { header: 'Transfer In', key: 'lastTransferIn', format: 'date' }, { header: 'Transfer Out', key: 'lastTransferOut', format: 'date' },
      { header: 'Action', key: 'action' }
    ];
  }
  else if (type === 'deadstock') {
    sheetName = 'P8 Weeks No Sales';
    title = 'P8 Weeks No Sales — CAMANAVA Inventory';
    data = baseRows.filter(r => r.isDeadStock).sort((a, b) => b.onHandValue - a.onHandValue).map(r => ({
      store: `${r.storeNumber} - ${r.storeName}`, area: r.area, skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases: r.stdPack > 0 && r.stdPack === r.onHand ? 'Per Piece' : (r.stdPack > 0 && r.onHand !== 0 ? +(r.onHand / r.stdPack).toFixed(2) : 'Per Piece'),
      onHandValue: r.onHandValue,
      weeksToSell: r.p8ave > 0 ? r.onHand / r.p8ave : null,
      daysCover: r.skuDaysCover != null ? Math.round(r.skuDaysCover) : null,
      dateLastSold: formatDate(r.dateLastSold), dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn), lastTransferOut: formatDate(r.lastTransferOut),
      action: r.onHandValue > 10000 ? 'High-Value Review' : 'Markdown / Liquidate'
    }));
    columns = [
      { header: 'Store', key: 'store' }, { header: 'Area', key: 'area' },
      { header: 'SKU Code', key: 'skuCode' }, { header: 'Description', key: 'skuDesc' }, { header: 'Supplier', key: 'supplier' },
      { header: 'On Hand', key: 'onHand', format: 'integer' }, { header: 'Qty in Cases', key: 'qtyCases' },
      { header: 'Inv Value', key: 'onHandValue', format: 'currency' },
      { header: 'WTS', key: 'weeksToSell', format: 'decimal' }, { header: 'Days Cover', key: 'daysCover', format: 'integer' },
      { header: 'Last Sold', key: 'dateLastSold', format: 'date' }, { header: 'Last Received', key: 'dateLastReceived', format: 'date' },
      { header: 'Transfer In', key: 'lastTransferIn', format: 'date' }, { header: 'Transfer Out', key: 'lastTransferOut', format: 'date' },
      { header: 'Action', key: 'action' }
    ];
  }
  else if (type === 'outofstock') {
    sheetName = 'Out of Stock';
    title = 'Out of Stock';
    data = baseRows.filter(r => r.isOutOfStock).sort((a, b) => b.lostSalesPerWeek - a.lostSalesPerWeek).map(r => {
      let qtyCases;
      if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
      else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
      else qtyCases = 'Per Piece';
      return {
        storeNumber: r.storeNumber, storeName: r.storeName, area: r.area,
        skuCode: r.skuCode, skuDesc: r.skuDesc, supplier: r.supplierName,
        onHand: r.onHand, stdPack: r.stdPack, qtyCases, invValue: r.onHandValue,
        p8ave: r.p8ave, weeksToSell: r.skuWTS != null ? +r.skuWTS.toFixed(2) : null,
        daysCover: r.skuDaysCover != null ? Math.round(r.skuDaysCover) : null,
        status: 'OOS', lostSalesPerWeek: r.lostSalesPerWeek,
        ico: r.ico, poOrderGR: r.poOrderGR, trfOrderGR: r.trfOrderGR,
        dateLastSold: formatDate(r.dateLastSold), dateLastReceived: formatDate(r.dateLastReceived),
        lastTransferIn: formatDate(r.lastTransferIn), lastTransferOut: formatDate(r.lastTransferOut)
      };
    });
    columns = [
      { header: 'Store #', key: 'storeNumber' }, { header: 'Store Name', key: 'storeName' }, { header: 'Area', key: 'area' },
      { header: 'SKU Code', key: 'skuCode' }, { header: 'Description', key: 'skuDesc' }, { header: 'Supplier', key: 'supplier' },
      { header: 'On Hand', key: 'onHand', format: 'integer' }, { header: 'Std Pack', key: 'stdPack', format: 'integer' },
      { header: 'Qty Cases', key: 'qtyCases' }, { header: 'Inv Value', key: 'invValue', format: 'currency' },
      { header: 'P8 Ave/Wk', key: 'p8ave', format: 'decimal' }, { header: 'WTS', key: 'weeksToSell', format: 'decimal' },
      { header: 'Days Cover', key: 'daysCover', format: 'integer' }, { header: 'Status', key: 'status' },
      { header: 'Lost Sales/Wk', key: 'lostSalesPerWeek', format: 'currency' }, { header: 'ICO', key: 'ico' },
      { header: 'PO On Order', key: 'poOrderGR', format: 'integer' }, { header: 'Trf On Order', key: 'trfOrderGR', format: 'integer' },
      { header: 'Last Sold', key: 'dateLastSold', format: 'date' }, { header: 'Last Received', key: 'dateLastReceived', format: 'date' },
      { header: 'Transfer In', key: 'lastTransferIn', format: 'date' }, { header: 'Transfer Out', key: 'lastTransferOut', format: 'date' }
    ];
  }
  else if (type === 'stores') {
    sheetName = 'Stores';
    title = 'Store Analysis — CAMANAVA Inventory';
    // Aggregate from filtered rows (respects area lock)
    const groups = {};
    for (const r of baseRows) {
      const k = r.storeNumber;
      if (!groups[k]) groups[k] = { storeNumber: r.storeNumber, storeName: r.storeName, area: r.area, totalValue: 0, totalOnHand: 0, totalSKUs: 0, criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0, totalLostSales: 0, totalWklSalesValue: 0, nonPCount: 0 };
      const g = groups[k];
      g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
      g.totalLostSales += r.lostSalesPerWeek;
      g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
      if (r.merchGro !== 'P') g.nonPCount++;
      if (r.isCritical && r.merchGro !== 'P') g.criticalCount++;
      if (r.isOverstock) g.overstockCount++;
      if (r.isDeadStock) g.deadCount++;
      if (r.isOutOfStock && r.merchGro !== 'P') g.oosCount++;
    }
    data = Object.values(groups).map(g => {
      const daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
      return {
        storeNumber: g.storeNumber, storeName: g.storeName, area: g.area,
        totalValue: g.totalValue, totalOnHand: g.totalOnHand, totalSKUs: g.totalSKUs,
        weeksToSell: daysCover != null ? +(daysCover / 7).toFixed(2) : null,
        daysCover: daysCover != null ? Math.round(daysCover) : null,
        oosCount: g.oosCount, totalLostSales: g.totalLostSales,
        criticalCount: g.criticalCount, overstockCount: g.overstockCount, deadCount: g.deadCount
      };
    }).sort((a, b) => b.totalValue - a.totalValue);
    columns = [
      { header: 'Store #', key: 'storeNumber' }, { header: 'Name', key: 'storeName' }, { header: 'Area', key: 'area' },
      { header: 'Inv Value', key: 'totalValue', format: 'currency' }, { header: 'On Hand', key: 'totalOnHand', format: 'integer' },
      { header: 'SKUs', key: 'totalSKUs', format: 'integer' },
      { header: 'WTS', key: 'weeksToSell', format: 'decimal' }, { header: 'Days Cover', key: 'daysCover', format: 'integer' },
      { header: 'OOS', key: 'oosCount', format: 'integer' }, { header: 'Lost Sales/Wk', key: 'totalLostSales', format: 'currency' },
      { header: 'Critical', key: 'criticalCount', format: 'integer' }, { header: 'Overstock', key: 'overstockCount', format: 'integer' },
      { header: 'Dead', key: 'deadCount', format: 'integer' }
    ];
  }
  else if (type === 'suppliers') {
    sheetName = 'Suppliers';
    title = 'Supplier Analysis — CAMANAVA Inventory';
    const groups = {};
    for (const r of baseRows) {
      if (!r.supplierCode) continue;
      const k = r.supplierCode;
      if (!groups[k]) groups[k] = { supplierCode: r.supplierCode, supplierName: r.supplierName, totalValue: 0, totalOnHand: 0, totalSKUs: 0, criticalCount: 0, overstockCount: 0, deadCount: 0, oosCount: 0, totalWklSalesValue: 0, totalP8Ave: 0, nonPCount: 0 };
      const g = groups[k];
      g.totalValue += r.onHandValue; g.totalOnHand += r.onHand; g.totalSKUs++;
      g.totalWklSalesValue += (r.wkAveNet * r.avgCost);
      g.totalP8Ave += r.p8ave;
      if (r.merchGro !== 'P') g.nonPCount++;
      if (r.isCritical && r.merchGro !== 'P') g.criticalCount++;
      if (r.isOverstock) g.overstockCount++;
      if (r.isDeadStock) g.deadCount++;
      if (r.isOutOfStock && r.merchGro !== 'P') g.oosCount++;
    }
    data = Object.values(groups).map(g => {
      const daysCover = g.totalWklSalesValue > 0 ? (g.totalValue * 7) / g.totalWklSalesValue : null;
      return {
        supplierCode: g.supplierCode, supplierName: g.supplierName,
        totalValue: g.totalValue, totalP8Ave: +g.totalP8Ave.toFixed(2), totalOnHand: g.totalOnHand, totalSKUs: g.totalSKUs,
        weeksToSell: daysCover != null ? +(daysCover / 7).toFixed(2) : null,
        daysCover: daysCover != null ? Math.round(daysCover) : null,
        oosCount: g.oosCount,
        criticalCount: g.criticalCount, overstockCount: g.overstockCount, deadCount: g.deadCount
      };
    }).sort((a, b) => b.totalValue - a.totalValue);
    columns = [
      { header: 'Code', key: 'supplierCode' }, { header: 'Name', key: 'supplierName' },
      { header: 'Inv Value', key: 'totalValue', format: 'currency' },
      { header: 'P8 Ave/Wk', key: 'totalP8Ave', format: 'decimal' },
      { header: 'On Hand', key: 'totalOnHand', format: 'integer' },
      { header: 'SKUs', key: 'totalSKUs', format: 'integer' },
      { header: 'WTS', key: 'weeksToSell', format: 'decimal' }, { header: 'Days Cover', key: 'daysCover', format: 'integer' },
      { header: 'OOS', key: 'oosCount', format: 'integer' },
      { header: 'Critical', key: 'criticalCount', format: 'integer' }, { header: 'Overstock', key: 'overstockCount', format: 'integer' },
      { header: 'Dead', key: 'deadCount', format: 'integer' }
    ];
  }
  else {
    return res.status(404).send('Unknown export type');
  }

  // Apply tab search filter (so the export matches what the user sees)
  data = applyTableSearch(data, search);
  if (!data || data.length === 0) return res.status(204).send('No data');

  const filename = `${sheetName.replace(/[^a-zA-Z0-9]+/g, '_')}_${today}.xlsx`;
  await writeStyledWorkbook(res, { sheetName, title, columns, rows: data, filename });
});

// SKU Analysis Excel export — respects all filters, sort, search, and status
app.get('/api/export-skus-xlsx', async (req, res) => {
  if (!cache.ready) return res.status(503).send('Cache not ready');
  const { sortBy = '', sortDir = 'asc', search = '', status = '', token, ...filters } = req.query;
  const s = sessions[token || ''];
  if (s && !s.isAdmin && s.area) filters.area = s.area;

  let rows = applyFilters(cache.rows, filters);
  if (status === 'critical') rows = rows.filter(r => r.isCritical);
  else if (status === 'oos') rows = rows.filter(r => r.isOutOfStock);
  else if (status === 'overstock') rows = rows.filter(r => r.isOverstock);
  else if (status === 'deadstock') rows = rows.filter(r => r.isDeadStock);
  else if (status === 'normal') rows = rows.filter(r => !r.isCritical && !r.isOutOfStock && !r.isOverstock && !r.isDeadStock);

  if (search) {
    const q = search.toLowerCase();
    rows = rows.filter(r =>
      (r.skuCode || '').toLowerCase().includes(q) ||
      (r.skuDesc || '').toLowerCase().includes(q) ||
      (r.supplierName || '').toLowerCase().includes(q) ||
      (r.storeName || '').toLowerCase().includes(q)
    );
  }

  if (sortBy) {
    const sortFieldMap = { invValue: 'onHandValue', weeksToSell: 'skuWTS', daysCover: 'skuDaysCover' };
    const field = sortFieldMap[sortBy] || sortBy;
    const dir = sortDir === 'desc' ? -1 : 1;
    rows = [...rows].sort((a, b) => {
      let av = a[field], bv = b[field];
      const aNull = (av == null || av === '');
      const bNull = (bv == null || bv === '');
      if (aNull && bNull) return 0;
      if (aNull) return 1;
      if (bNull) return -1;
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv)) * dir;
    });
  }

  // Build workbook
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CAMANAVA Inventory Dashboard';
  wb.created = new Date();
  const ws = wb.addWorksheet('SKU Analysis');

  // Dark green brand color #1B5E20, white font for title + header
  const DARK_GREEN = 'FF1B5E20';
  const headers = [
    { header: 'Store #', key: 'storeNumber', width: 10 },
    { header: 'Store Name', key: 'storeName', width: 24 },
    { header: 'Area', key: 'area', width: 18 },
    { header: 'SKU Code', key: 'skuCode', width: 14 },
    { header: 'Description', key: 'skuDesc', width: 36 },
    { header: 'Supplier', key: 'supplierName', width: 28 },
    { header: 'On Hand', key: 'onHand', width: 10 },
    { header: 'Std Pack', key: 'stdPack', width: 10 },
    { header: 'Qty Cases', key: 'qtyCases', width: 10 },
    { header: 'Inv Value', key: 'invValue', width: 14 },
    { header: 'P8 Ave/Wk', key: 'p8ave', width: 12 },
    { header: 'WTS', key: 'weeksToSell', width: 10 },
    { header: 'Days Cover', key: 'daysCover', width: 12 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Lost Sales/Wk', key: 'lostSalesPerWeek', width: 14 },
    { header: 'ICO', key: 'ico', width: 8 },
    { header: 'PO On Order', key: 'poOrderGR', width: 14 },
    { header: 'Trf On Order', key: 'trfOrderGR', width: 14 },
    { header: 'Last Sold', key: 'dateLastSold', width: 14 },
    { header: 'Last Received', key: 'dateLastReceived', width: 14 },
    { header: 'Transfer In', key: 'lastTransferIn', width: 16 },
    { header: 'Transfer Out', key: 'lastTransferOut', width: 16 }
  ];

  // ── Title row (merged) ──
  const todayStr = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const titleText = 'SKU Analysis Report as of ' + todayStr;
  ws.mergeCells(1, 1, 1, headers.length);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = titleText;
  titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: 'Calibri' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(1).height = 26;

  // ── Header row ──
  ws.columns = headers;
  const headerRow = ws.getRow(2);
  headers.forEach((h, i) => { headerRow.getCell(i + 1).value = h.header; });
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  });
  headerRow.height = 20;

  // ── Data rows ──
  for (const r of rows) {
    let qtyCases;
    if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
    else if (r.stdPack > 0 && r.onHand > 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
    else if (r.onHand === 0) qtyCases = 0;
    else qtyCases = 'Per Piece';
    ws.addRow({
      storeNumber: r.storeNumber,
      storeName: r.storeName,
      area: r.area,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplierName: r.supplierName,
      onHand: r.onHand,
      stdPack: r.stdPack,
      qtyCases,
      invValue: +(r.onHandValue || 0).toFixed(2),
      p8ave: +(r.p8ave || 0).toFixed(2),
      weeksToSell: r.skuWTS != null ? +r.skuWTS.toFixed(2) : null,
      daysCover: r.skuDaysCover != null ? +r.skuDaysCover.toFixed(0) : null,
      status: r.isCritical ? 'Critical' : r.isOutOfStock ? 'OOS' : r.isOverstock ? 'Overstock' : r.isDeadStock ? 'Dead Stock' : 'Normal',
      lostSalesPerWeek: +(r.lostSalesPerWeek || 0).toFixed(2),
      ico: r.ico,
      poOrderGR: r.poOrderGR,
      trfOrderGR: r.trfOrderGR,
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut)
    });
  }

  // Freeze the title + header rows
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  // AutoFilter on the header row
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: headers.length } };

  // Number formats on currency/qty columns
  const formatMap = { invValue: '#,##0.00', p8ave: '#,##0.00', lostSalesPerWeek: '#,##0.00', weeksToSell: '#,##0.00', daysCover: '#,##0', onHand: '#,##0', poOrderGR: '#,##0', trfOrderGR: '#,##0' };
  ws.columns.forEach(col => { if (formatMap[col.key]) col.numFmt = formatMap[col.key]; });

  // Stream the workbook
  const filename = `SKU_Analysis_${new Date().toISOString().split('T')[0]}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  await wb.xlsx.write(res);
  res.end();
});

// Negative SKU Excel export — items with onHand < 0
app.get('/api/export-negativeskus-xlsx', async (req, res) => {
  if (!cache.ready) return res.status(503).send('Cache not ready');
  const filters = resolveFilters(req);
  let rows = applyFilters(cache.rows, filters).filter(r => r.isNegativeStock);
  rows.sort((a, b) => a.onHand - b.onHand);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'CAMANAVA Inventory Dashboard';
  wb.created = new Date();
  const ws = wb.addWorksheet('Negative SKU');
  const DARK_GREEN = 'FF1B5E20';

  const headers = [
    { header: 'Store Name', key: 'storeName', width: 26 },
    { header: 'SKU', key: 'skuCode', width: 14 },
    { header: 'Description', key: 'skuDesc', width: 36 },
    { header: 'Supplier', key: 'supplier', width: 28 },
    { header: 'On Hand', key: 'onHand', width: 12 },
    { header: 'Qty in Cases', key: 'qtyCases', width: 14 },
    { header: 'Inv Value', key: 'onHandValue', width: 14 },
    { header: 'P8 Ave/Wk', key: 'p8ave', width: 12 },
    { header: 'Last Sold', key: 'dateLastSold', width: 14 },
    { header: 'Last Received', key: 'dateLastReceived', width: 14 }
  ];

  ws.mergeCells(1, 1, 1, headers.length);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = 'Negative SKU Report as of ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: 'Calibri' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(1).height = 26;

  ws.columns = headers;
  const headerRow = ws.getRow(2);
  headers.forEach((h, i) => { headerRow.getCell(i + 1).value = h.header; });
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  });
  headerRow.height = 20;

  for (const r of rows) {
    let qtyCases;
    if (r.stdPack > 0 && r.stdPack === r.onHand) qtyCases = 'Per Piece';
    else if (r.stdPack > 0 && r.onHand !== 0) qtyCases = +(r.onHand / r.stdPack).toFixed(2);
    else qtyCases = 'Per Piece';
    ws.addRow({
      storeName: `${r.storeNumber} - ${r.storeName}`,
      skuCode: r.skuCode,
      skuDesc: r.skuDesc,
      supplier: r.supplierName,
      onHand: r.onHand,
      qtyCases,
      onHandValue: +(r.onHandValue || 0).toFixed(2),
      p8ave: +(r.p8ave || 0).toFixed(2),
      dateLastSold: formatDate(r.dateLastSold),
      dateLastReceived: formatDate(r.dateLastReceived),
      lastTransferIn: formatDate(r.lastTransferIn),
      lastTransferOut: formatDate(r.lastTransferOut)
    });
  }

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: headers.length } };

  const formatMap = { onHand: '#,##0', onHandValue: '#,##0.00', p8ave: '#,##0.00' };
  ws.columns.forEach(col => { if (formatMap[col.key]) col.numFmt = formatMap[col.key]; });

  const filename = `Negative_SKU_${new Date().toISOString().split('T')[0]}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  await wb.xlsx.write(res);
  res.end();
});

// Top 300 SKU Excel export — joined data, dark green styled
app.get('/api/export-top300-xlsx', async (req, res) => {
  if (!cache.ready) return res.status(503).send('Cache not ready');
  const filters = resolveFilters(req);
  const top300 = cache.top300 || [];
  const idx = cache.storeSkuIndex || {};
  const rows = [];
  for (const t of top300) {
    if (filters.area && t.area !== filters.area) continue;
    if (filters.store && t.storeNumber !== filters.store) continue;
    const key = (t.storeNumber || '').trim() + '_' + (t.sku || '').trim();
    const inv = idx[key];
    if (filters.category && (!inv || inv.catName !== filters.category)) continue;
    if (filters.supplier && (!inv || inv.supplierName !== filters.supplier)) continue;
    let status = 'Not Found';
    if (inv) {
      if (inv.isCritical) status = 'Critical';
      else if (inv.isOutOfStock) status = 'OOS';
      else if (inv.isOverstock) status = 'Overstock';
      else if (inv.isDeadStock) status = 'Dead Stock';
      else status = 'Normal';
    }
    rows.push({
      area: t.area,
      storeName: t.storeName ? (t.storeNumber + ' - ' + t.storeName) : (inv ? `${inv.storeNumber} - ${inv.storeName}` : t.storeNumber),
      rank: t.rank,
      sku: t.sku,
      itemDescription: t.desc || (inv ? inv.skuDesc : ''),
      supplier: inv ? inv.supplierName : '',
      onHand: inv ? inv.onHand : null,
      qtyCases: inv ? (inv.stdPack > 0 && inv.stdPack === inv.onHand ? 'Per Piece' : (inv.stdPack > 0 && inv.onHand !== 0 ? +(inv.onHand / inv.stdPack).toFixed(2) : 'Per Piece')) : null,
      p8ave: inv ? +(inv.p8ave || 0).toFixed(2) : null,
      daysCover: inv && inv.skuDaysCover != null ? +inv.skuDaysCover.toFixed(0) : null,
      status,
      incomingPO: inv ? inv.poOrderGR : null,
      lostSalesPerWeek: inv ? +(inv.lostSalesPerWeek || 0).toFixed(2) : null,
      ico: inv ? inv.ico : '',
      dateLastSold: inv ? formatDate(inv.dateLastSold) : '',
      dateLastReceived: inv ? formatDate(inv.dateLastReceived) : '',
      lastTransferIn: inv ? formatDate(inv.lastTransferIn) : '',
      lastTransferOut: inv ? formatDate(inv.lastTransferOut) : ''
    });
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'CAMANAVA Inventory Dashboard';
  wb.created = new Date();
  const ws = wb.addWorksheet('Top 300 SKU');
  const DARK_GREEN = 'FF1B5E20';
  const headers = [
    { header: 'Area', key: 'area', width: 18 },
    { header: 'Store Name', key: 'storeName', width: 24 },
    { header: 'Rank', key: 'rank', width: 8 },
    { header: 'SKU', key: 'sku', width: 14 },
    { header: 'Item Description', key: 'itemDescription', width: 36 },
    { header: 'Supplier', key: 'supplier', width: 28 },
    { header: 'On Hand', key: 'onHand', width: 10 },
    { header: 'Qty in Cases', key: 'qtyCases', width: 12 },
    { header: 'P8 Ave', key: 'p8ave', width: 12 },
    { header: 'Days Cover', key: 'daysCover', width: 12 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Incoming PO', key: 'incomingPO', width: 12 },
    { header: 'Lost Sales/Wk', key: 'lostSalesPerWeek', width: 14 },
    { header: 'ICO', key: 'ico', width: 8 },
    { header: 'Last Sold', key: 'dateLastSold', width: 14 },
    { header: 'Last Received', key: 'dateLastReceived', width: 14 },
    { header: 'Transfer In', key: 'lastTransferIn', width: 14 },
    { header: 'Transfer Out', key: 'lastTransferOut', width: 14 }
  ];
  ws.mergeCells(1, 1, 1, headers.length);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = 'Top 300 SKU Report as of ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: 'Calibri' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(1).height = 26;

  ws.columns = headers;
  const headerRow = ws.getRow(2);
  headers.forEach((h, i) => { headerRow.getCell(i + 1).value = h.header; });
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREEN } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  });
  headerRow.height = 20;

  for (const r of rows) ws.addRow(r);
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: headers.length } };

  const formatMap = { onHand: '#,##0', p8ave: '#,##0.00', daysCover: '#,##0', incomingPO: '#,##0', lostSalesPerWeek: '#,##0.00', rank: '#,##0' };
  ws.columns.forEach(col => { if (formatMap[col.key]) col.numFmt = formatMap[col.key]; });

  const filename = `Top_300_SKU_${new Date().toISOString().split('T')[0]}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  await wb.xlsx.write(res);
  res.end();
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
<script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"></script>
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

/* CAMERA MODAL */
#camera-modal {
  position:fixed; inset:0; background:rgba(0,0,0,0.85); z-index:10002;
  display:flex; align-items:center; justify-content:center;
}
.camera-box {
  background:var(--bg2); border:1px solid var(--border); border-radius:12px;
  padding:16px; width:min(400px, 95vw);
  display:flex; flex-direction:column; gap:10px;
}
.camera-header { display:flex; justify-content:space-between; align-items:center; }
#camera-reader { border-radius:8px; overflow:hidden; background:#000; min-height:240px; }
#camera-reader video { width:100% !important; border-radius:8px; }
.camera-status { font-size:12px; color:var(--text2); text-align:center; font-family:'IBM Plex Mono',monospace; }
.camera-hint { font-size:11px; color:var(--text2); text-align:center; }

/* INACTIVITY MODAL */
#inactivity-modal {
  position:fixed; inset:0; background:rgba(0,0,0,0.7); z-index:10001;
  display:flex; align-items:center; justify-content:center;
  backdrop-filter:blur(4px);
}
.inactivity-box {
  background:var(--bg2); border:1px solid var(--yellow); border-radius:12px;
  padding:32px; width:380px; max-width:90vw; text-align:center;
  display:flex; flex-direction:column; gap:14px;
  box-shadow:0 8px 40px rgba(0,0,0,0.6);
}
.inactivity-icon { font-size:48px; }
.inactivity-title { font-size:18px; font-weight:700; color:var(--yellow-light); }
.inactivity-text { font-size:13px; color:var(--text); line-height:1.5; }
.inactivity-text span { font-weight:700; color:var(--red-light); font-family:'IBM Plex Mono',monospace; font-size:16px; }
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

/* SUMMARY PANELS (Negative SKU) */
.summary-grid {
  display: grid; grid-template-columns: 1fr 1fr; gap: 12px;
  margin-bottom: 14px;
}
.summary-card {
  background: var(--bg2); border: 1px solid var(--border); border-radius: 8px;
  overflow: hidden; display: flex; flex-direction: column;
}
.summary-card-title {
  background: var(--bg3); padding: 8px 12px; font-size: 12px; font-weight: 600;
  color: var(--text2); text-transform: uppercase; letter-spacing: 0.5px;
  border-bottom: 1px solid var(--border);
}
.summary-card-body { max-height: 240px; overflow-y: auto; overflow-x: auto; }
.summary-table { width: 100%; border-collapse: collapse; font-size: 11px; }
.summary-table th {
  background: var(--bg3); color: var(--text2); text-align: left;
  padding: 6px 10px; font-weight: 600; font-size: 10px; text-transform: uppercase;
  letter-spacing: 0.3px; cursor: pointer; user-select: none;
  border-bottom: 1px solid var(--border); position: sticky; top: 0; z-index: 1;
}
.summary-table th:hover { color: var(--text); }
.summary-table td { padding: 6px 10px; border-bottom: 1px solid var(--border); }
.summary-table tbody tr:hover { background: var(--bg3); }
.summary-table .num { font-family: 'IBM Plex Mono', monospace; text-align: right; }
.summary-table .neg { color: var(--red-light); font-weight: 600; }
.summary-table tfoot td {
  background: var(--bg3); font-weight: 700; border-top: 2px solid var(--green-bright);
  color: var(--text);
}
@media (max-width: 900px) {
  .summary-grid { grid-template-columns: 1fr; gap: 10px; }
  .summary-card-body { max-height: 200px; }
}

/* RISK PILLS */
.totals-pill {
  display:inline-flex; align-items:center; gap:10px;
  margin-left:12px; padding:4px 12px;
  background:var(--bg3); border:1px solid var(--border); border-radius:6px;
  font-size:11px; font-weight:500; color:var(--text2);
  font-family:'IBM Plex Mono',monospace;
}
.totals-pill .tp-label { color:var(--text2); }
.totals-pill .tp-value { color:var(--text); font-weight:700; }
.totals-pill .tp-value.green { color:var(--green-bright); }
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

/* ============ MOBILE / TABLET RESPONSIVE ============ */
@media (max-width: 900px) {
  /* Prevent content area from scrolling horizontally; tabs handle their own */
  .content { overflow-x: hidden; min-width: 0; }
  .main { overflow-x: hidden; }

  /* Header: compact and stack-friendly */
  .header { padding: 8px 10px; flex-wrap: wrap; gap: 8px; }
  .header-logo h1 { font-size: 13px; }
  .header-logo span { font-size: 10px; }
  .header-right { gap: 6px; flex-wrap: wrap; }
  .refresh-info { font-size: 10px; }
  #user-info { font-size: 10px; max-width: 100%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }

  /* Buttons bigger touch target */
  .btn { padding: 8px 12px; font-size: 12px; min-height: 36px; }
  .btn-sm { padding: 6px 10px; font-size: 11px; min-height: 32px; }

  /* Content padding */
  .content { padding: 10px; gap: 12px; }
  .main { height: calc(100vh - 90px); }

  /* TABS: WRAP to multiple rows on mobile so all tabs are always visible */
  .tabs {
    display: flex !important;
    flex-wrap: wrap !important;
    gap: 4px !important;
    overflow: visible !important;
    border-bottom: 1px solid var(--border);
    padding-bottom: 4px;
    /* Remove any mask/gradient */
    -webkit-mask-image: none !important;
    mask-image: none !important;
  }
  .tab {
    white-space: nowrap;
    padding: 8px 12px;
    font-size: 12px;
    min-height: 36px;
    border: 1px solid var(--border);
    border-bottom: 1px solid var(--border) !important;
    border-radius: 6px;
    background: var(--bg3);
    flex: 0 0 auto;
  }
  .tab.active {
    background: var(--green);
    border-color: var(--green-light);
    color: #fff !important;
    border-bottom: 1px solid var(--green-light) !important;
  }

  /* FILTER BAR: stack as 2-column grid */
  .filter-bar { padding: 10px; gap: 8px; }
  .fb-group { flex: 1 1 calc(50% - 4px); min-width: 0; max-width: none; }
  .fb-actions { width: 100%; margin-left: 0; justify-content: space-between; }
  .filter-select { font-size: 13px; padding: 8px 10px; min-height: 38px; }

  /* KPI GRID: 2 columns on mobile */
  .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 8px; }
  .kpi-card { padding: 10px; }
  .kpi-label { font-size: 10px; }
  .kpi-value { font-size: 18px; }
  .kpi-sub { font-size: 9px; }

  /* SECTION HEADER: stack title + actions */
  .section-header {
    flex-direction: column; align-items: stretch; gap: 8px;
  }
  .section-title { font-size: 13px; flex-wrap: wrap; }
  .section-actions { flex-wrap: wrap; gap: 6px; }
  .section-actions .table-search { width: 100%; min-height: 38px; }
  .totals-pill { margin-left: 0; margin-top: 4px; font-size: 10px; padding: 4px 8px; }

  /* TABLES: smaller font, allow horizontal scroll */
  table { font-size: 11px; }
  table th, table td { padding: 6px 8px; }
  .table-wrap { -webkit-overflow-scrolling: touch; max-height: 70vh; }

  /* PAGINATION: bigger taps, wrap */
  .pagination { flex-wrap: wrap; gap: 4px; padding: 8px 0; }
  .page-btn { min-width: 36px; min-height: 36px; font-size: 12px; }
  .page-info { width: 100%; text-align: center; font-size: 11px; padding: 4px 0; }

  /* MODALS: full-width on small screens */
  .login-box, .inactivity-box { width: min(360px, 95vw); padding: 24px; }
  .camera-box { width: 95vw; padding: 12px; }
  #camera-reader { min-height: 60vw; max-height: 70vh; }

  /* SKU Analysis controls stack better */
  #sku-status-filter { width: 100% !important; }
  #sku-upc-input { width: 100% !important; }
  #sku-upc-msg { display: block; width: 100%; }
}

/* Phone (narrow) — extra tightening */
@media (max-width: 480px) {
  .header-logo h1 { font-size: 12px; }
  .kpi-grid { grid-template-columns: 1fr 1fr; gap: 6px; }
  .kpi-value { font-size: 16px; }
  .fb-group { flex: 1 1 100%; } /* one filter per row */
  table { font-size: 10px; }
  table th, table td { padding: 5px 6px; }
  .tab { padding: 10px 10px; font-size: 11px; }
  .charts-grid canvas { max-height: 220px; }
  .login-input { font-size: 16px; } /* avoid iOS zoom-on-focus */
  .filter-select, .table-search { font-size: 16px; } /* iOS no-zoom */
}

/* TOUCH DEVICE: remove hover-only effects */
@media (hover: none) {
  .kpi-card:hover { border-color: var(--border); } /* don't latch hover state */
  .btn:hover { border-color: var(--border); color: var(--text); }
}

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

<div id="camera-modal" style="display:none;">
  <div class="camera-box">
    <div class="camera-header">
      <div style="font-size:14px;font-weight:600;">📷 Scan Barcode</div>
      <button onclick="closeCameraScan()" style="background:none;border:none;color:var(--text);font-size:22px;cursor:pointer;line-height:1;">×</button>
    </div>
    <div id="camera-reader" style="width:100%;"></div>
    <div class="camera-status" id="camera-status">Initializing camera...</div>
    <div class="camera-hint">Point your camera at the barcode</div>
  </div>
</div>

<div id="inactivity-modal" style="display:none;">
  <div class="inactivity-box">
    <div class="inactivity-icon">⏰</div>
    <div class="inactivity-title">Session expiring soon</div>
    <div class="inactivity-text">You will be logged out in <span id="inactivity-countdown">60</span> seconds due to inactivity.</div>
    <button class="login-btn" onclick="stayLoggedIn()">Stay Logged In</button>
  </div>
</div>

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
        <label>Category</label>
        <select class="filter-select" id="f-category" onchange="applyFilter('category',this.value)">
          <option value="">All Categories</option>
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
      <div class="tab" onclick="showTab('aging')">📅 Aging</div>
      <div class="tab" onclick="showTab('blackinv')">⬛ Black Inventory</div>
      <div class="tab" onclick="showTab('negsku')">⚠️ Negative SKU</div>
      <div class="tab" onclick="showTab('deadstock')">💀 P8 Weeks No Sales</div>
      <div class="tab" onclick="showTab('stores')">🏪 Stores</div>
      <div class="tab" onclick="showTab('suppliers')">🏭 Suppliers</div>
      <div class="tab" onclick="showTab('skus')">🔍 SKU Analysis</div>
      <div class="tab" onclick="showTab('top300')">⭐ Top 300 SKU</div>
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
            <button class="btn btn-sm" onclick="exportData('stores')">⬇ Export Excel</button>
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

      <!-- TOP 300 SKU STORE PERFORMANCE -->
      <div class="section" style="margin-top:4px;">
        <div class="section-header">
          <div class="section-title">🎯 Top 300 SKU Store Performance
            <span style="font-size:10px;color:var(--text2);margin-left:8px;">
              Score = (OOS% × 0.6) + (Critical% × 0.4) &nbsp;|&nbsp;
              <span style="color:#3fb950;">🟢 Best</span> &nbsp;
              <span style="color:#e3b341;">🟡 Watch</span> &nbsp;
              <span style="color:#f85149;">🔴 Worst</span>
            </span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search store..." oninput="searchTable('top300-perf-table',this.value)"/>
          </div>
        </div>
        <div class="table-wrap" style="max-height:520px;">
          <table id="top300-perf-table">
            <thead><tr>
              <th onclick="sortTop300Perf(0)">Rank</th>
              <th onclick="sortTop300Perf(1)">Store #</th>
              <th onclick="sortTop300Perf(2)">Store Name</th>
              <th onclick="sortTop300Perf(3)">Area</th>
              <th onclick="sortTop300Perf(4)" title="Top 300 entries matched against InvData for this store">SKUs</th>
              <th onclick="sortTop300Perf(5)" title="Out of Stock in Top 300 / Total">OOS %</th>
              <th onclick="sortTop300Perf(6)" title="Critical (WTS < 2 wks) in Top 300 / Total">Critical %</th>
              <th onclick="sortTop300Perf(7)" title="Healthy in Top 300 / Total">Healthy %</th>
              <th onclick="sortTop300Perf(8)" title="(OOS% × 0.6) + (Critical% × 0.4)">Score</th>
            </tr></thead>
            <tbody id="top300-perf-body"></tbody>
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
          <div class="chart-title">Inventory Value by Category</div>
          <canvas id="chart-categories"></canvas>
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
            <button class="btn btn-sm" onclick="exportData('critical')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="critical-table">
            <thead><tr>
              <th data-field="store" onclick="sortTable('critical-table',0)">Store</th>
              <th data-field="area" onclick="sortTable('critical-table',1)">Area</th>
              <th data-field="skuCode" onclick="sortTable('critical-table',2)">SKU Code</th>
              <th data-field="skuDesc" onclick="sortTable('critical-table',3)">Description</th>
              <th data-field="supplier" onclick="sortTable('critical-table',4)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('critical-table',5)">On Hand</th>
              <th data-field="onHandValue" onclick="sortTable('critical-table',6)">Value</th>
              <th data-field="currentWkSales" onclick="sortTable('critical-table',7)">Cur Wk Sales</th>
              <th data-field="p8ave" onclick="sortTable('critical-table',8)">P8 Ave</th>
              <th data-field="wtsNet" onclick="sortTable('critical-table',9)">WTS Net</th>
              <th data-field="totalPO" onclick="sortTable('critical-table',10)">Total PO</th>
              <th data-field="dateLastSold" onclick="sortTable('critical-table',11)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('critical-table',12)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('critical-table',13)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('critical-table',14)">Transfer Out</th>
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
          <div class="section-title">📦 Overstock Items <span class="badge badge-yellow" id="overstock-count">0</span>
            <span class="totals-pill" id="overstock-totals"></span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('overstock-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('overstock')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="overstock-table">
            <thead><tr>
              <th data-field="store" onclick="sortTable('overstock-table',0)">Store</th>
              <th data-field="area" onclick="sortTable('overstock-table',1)">Area</th>
              <th data-field="skuCode" onclick="sortTable('overstock-table',2)">SKU Code</th>
              <th data-field="skuDesc" onclick="sortTable('overstock-table',3)">Description</th>
              <th data-field="supplier" onclick="sortTable('overstock-table',4)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('overstock-table',5)">On Hand</th>
              <th data-field="qtyCases" onclick="sortTable('overstock-table',6)">Qty in Cases</th>
              <th data-field="onHandValue" onclick="sortTable('overstock-table',7)">Value</th>
              <th data-field="p8ave" onclick="sortTable('overstock-table',8)">P8 Ave</th>
              <th data-field="wtsNet" onclick="sortTable('overstock-table',9)">WTS Net</th>
              <th data-field="dateLastSold" onclick="sortTable('overstock-table',10)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('overstock-table',11)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('overstock-table',12)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('overstock-table',13)">Transfer Out</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="overstock-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="overstock-pagination"></div>
      </div>
    </div>

    <!-- AGING TAB -->
    <div id="tab-aging" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">📅 Aging Items <span class="badge badge-yellow" id="aging-count">0</span>
            <span class="totals-pill" id="aging-totals"></span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('aging-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('aging')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="aging-table">
            <thead><tr>
              <th data-field="store" onclick="sortTable('aging-table',0)">Store</th>
              <th data-field="area" onclick="sortTable('aging-table',1)">Area</th>
              <th data-field="skuCode" onclick="sortTable('aging-table',2)">SKU Code</th>
              <th data-field="skuDesc" onclick="sortTable('aging-table',3)">Description</th>
              <th data-field="supplier" onclick="sortTable('aging-table',4)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('aging-table',5)">On Hand</th>
              <th data-field="qtyCases" onclick="sortTable('aging-table',6)">Qty in Cases</th>
              <th data-field="onHandValue" onclick="sortTable('aging-table',7)">Value</th>
              <th data-field="p8ave" onclick="sortTable('aging-table',8)">P8 Ave</th>
              <th data-field="daysCover" onclick="sortTable('aging-table',9)">Days Cover</th>
              <th data-field="dateLastSold" onclick="sortTable('aging-table',10)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('aging-table',11)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('aging-table',12)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('aging-table',13)">Transfer Out</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="aging-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="aging-pagination"></div>
      </div>
    </div>

    <!-- BLACK INVENTORY TAB -->
    <div id="tab-blackinv" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">⬛ Black Inventory <span class="badge badge-red" id="blackinv-count">0</span>
            <span class="totals-pill" id="blackinv-totals"></span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('blackinv-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('blackinventory')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="blackinv-table">
            <thead><tr>
              <th data-field="store" onclick="sortTable('blackinv-table',0)">Store</th>
              <th data-field="area" onclick="sortTable('blackinv-table',1)">Area</th>
              <th data-field="skuCode" onclick="sortTable('blackinv-table',2)">SKU Code</th>
              <th data-field="skuDesc" onclick="sortTable('blackinv-table',3)">Description</th>
              <th data-field="supplier" onclick="sortTable('blackinv-table',4)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('blackinv-table',5)">On Hand</th>
              <th data-field="qtyCases" onclick="sortTable('blackinv-table',6)">Qty in Cases</th>
              <th data-field="onHandValue" onclick="sortTable('blackinv-table',7)">Value</th>
              <th data-field="p8ave" onclick="sortTable('blackinv-table',8)">P8 Ave</th>
              <th data-field="daysCover" onclick="sortTable('blackinv-table',9)">Days Cover</th>
              <th data-field="dateLastSold" onclick="sortTable('blackinv-table',10)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('blackinv-table',11)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('blackinv-table',12)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('blackinv-table',13)">Transfer Out</th>
              <th>Action</th>
            </tr></thead>
            <tbody id="blackinv-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="blackinv-pagination"></div>
      </div>
    </div>

    <!-- NEGATIVE SKU TAB -->
    <div id="tab-negsku" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">⚠️ Negative SKU <span class="badge badge-red" id="negsku-count">0</span>
            <span class="totals-pill" id="negsku-totals"></span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('negsku-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportNegativeSKUsExcel()" id="negsku-export-btn">⬇ Export Excel</button>
          </div>
        </div>
        <!-- Summary panels: Store + Category -->
        <div class="summary-grid">
          <div class="summary-card">
            <div class="summary-card-title">📍 Summary by Store</div>
            <div class="summary-card-body">
              <table class="summary-table">
                <thead><tr>
                  <th onclick="sortSummary('store',0)">Store</th>
                  <th onclick="sortSummary('store',1)">On Hand</th>
                  <th onclick="sortSummary('store',2)">Inv Value</th>
                  <th onclick="sortSummary('store',3)">Count</th>
                </tr></thead>
                <tbody id="negsku-summary-store"></tbody>
              </table>
            </div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">🗂 Summary by Category</div>
            <div class="summary-card-body">
              <table class="summary-table">
                <thead><tr>
                  <th onclick="sortSummary('category',0)">Category</th>
                  <th onclick="sortSummary('category',1)">On Hand</th>
                  <th onclick="sortSummary('category',2)">Inv Value</th>
                  <th onclick="sortSummary('category',3)">Count</th>
                </tr></thead>
                <tbody id="negsku-summary-cat"></tbody>
              </table>
            </div>
          </div>
        </div>
        <div class="table-wrap">
          <table id="negsku-table">
            <thead><tr>
              <th data-field="storeName" onclick="sortTable('negsku-table',0)">Store Name</th>
              <th data-field="skuCode" onclick="sortTable('negsku-table',1)">SKU</th>
              <th data-field="skuDesc" onclick="sortTable('negsku-table',2)">Description</th>
              <th data-field="supplier" onclick="sortTable('negsku-table',3)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('negsku-table',4)">On Hand</th>
              <th data-field="qtyCases" onclick="sortTable('negsku-table',5)">Qty in Cases</th>
              <th data-field="onHandValue" onclick="sortTable('negsku-table',6)">Inv Value</th>
              <th data-field="p8ave" onclick="sortTable('negsku-table',7)">P8 Ave/Wk</th>
              <th data-field="dateLastSold" onclick="sortTable('negsku-table',8)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('negsku-table',9)">Last Received</th>
            </tr></thead>
            <tbody id="negsku-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="negsku-pagination"></div>
      </div>
    </div>

    <!-- DEAD STOCK TAB -->
    <div id="tab-deadstock" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">💀 P8 Weeks No Sales <span class="badge badge-red" id="deadstock-count">0</span>
            <span class="totals-pill" id="deadstock-totals"></span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('deadstock-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportData('deadstock')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="deadstock-table">
            <thead><tr>
              <th data-field="store" onclick="sortTable('deadstock-table',0)">Store</th>
              <th data-field="area" onclick="sortTable('deadstock-table',1)">Area</th>
              <th data-field="skuCode" onclick="sortTable('deadstock-table',2)">SKU Code</th>
              <th data-field="skuDesc" onclick="sortTable('deadstock-table',3)">Description</th>
              <th data-field="supplier" onclick="sortTable('deadstock-table',4)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('deadstock-table',5)">On Hand</th>
              <th data-field="qtyCases" onclick="sortTable('deadstock-table',6)">Qty in Cases</th>
              <th data-field="onHandValue" onclick="sortTable('deadstock-table',7)">Value</th>
              <th data-field="weeksToSell" onclick="sortTable('deadstock-table',8)">WTS</th>
              <th data-field="daysCover" onclick="sortTable('deadstock-table',9)">Days Cover</th>
              <th data-field="dateLastSold" onclick="sortTable('deadstock-table',10)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('deadstock-table',11)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('deadstock-table',12)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('deadstock-table',13)">Transfer Out</th>
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
            <button class="btn btn-sm" onclick="exportData('outofstock')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="outofstock-table">
            <thead><tr>
              <th data-field="storeNumber" onclick="sortTable('outofstock-table',0)">Store #</th>
              <th data-field="storeName" onclick="sortTable('outofstock-table',1)">Store Name</th>
              <th data-field="area" onclick="sortTable('outofstock-table',2)">Area</th>
              <th data-field="skuCode" onclick="sortTable('outofstock-table',3)">SKU Code</th>
              <th data-field="skuDesc" onclick="sortTable('outofstock-table',4)">Description</th>
              <th data-field="supplier" onclick="sortTable('outofstock-table',5)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('outofstock-table',6)">On Hand</th>
              <th data-field="stdPack" onclick="sortTable('outofstock-table',7)">Std Pack</th>
              <th data-field="qtyCases" onclick="sortTable('outofstock-table',8)">Qty Cases</th>
              <th data-field="invValue" onclick="sortTable('outofstock-table',9)">Inv Value</th>
              <th data-field="p8ave" onclick="sortTable('outofstock-table',10)">P8 Ave/Wk</th>
              <th data-field="weeksToSell" onclick="sortTable('outofstock-table',11)">WTS</th>
              <th data-field="daysCover" onclick="sortTable('outofstock-table',12)">Days Cover</th>
              <th data-field="status" onclick="sortTable('outofstock-table',13)">Status</th>
              <th data-field="lostSalesPerWeek" onclick="sortTable('outofstock-table',14)">Lost Sales/Wk</th>
              <th data-field="ico" onclick="sortTable('outofstock-table',15)">ICO</th>
              <th data-field="poOrderGR" onclick="sortTable('outofstock-table',16)">PO On Order</th>
              <th data-field="trfOrderGR" onclick="sortTable('outofstock-table',17)">Trf On Order</th>
              <th data-field="dateLastSold" onclick="sortTable('outofstock-table',18)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('outofstock-table',19)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('outofstock-table',20)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('outofstock-table',21)">Transfer Out</th>
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
            <button class="btn btn-sm" onclick="exportData('stores')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="stores-table">
            <thead><tr>
              <th data-field="storeNumber" onclick="sortTable('stores-table',0)">Store #</th>
              <th data-field="storeName" onclick="sortTable('stores-table',1)">Store Name</th>
              <th data-field="area" onclick="sortTable('stores-table',2)">Area</th>
              <th data-field="totalValue" onclick="sortTable('stores-table',3)">Inv Value</th>
              <th data-field="totalOnHand" onclick="sortTable('stores-table',4)">On Hand</th>
              <th data-field="totalSKUs" onclick="sortTable('stores-table',5)">SKUs</th>
              <th data-field="weeksToSell" onclick="sortTable('stores-table',6)">WTS</th>
              <th data-field="daysCover" onclick="sortTable('stores-table',7)">Days Cover</th>
              <th data-field="oosCount" onclick="sortTable('stores-table',8)">Out of Stock</th>
              <th data-field="totalLostSales" onclick="sortTable('stores-table',9)">Lost Sales/Wk</th>
              <th data-field="criticalCount" onclick="sortTable('stores-table',10)">Critical</th>
              <th data-field="overstockCount" onclick="sortTable('stores-table',11)">Overstock</th>
              <th data-field="deadCount" onclick="sortTable('stores-table',12)">Dead Stock</th>
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
              <th onclick="sortSupplierRisk(3)">P8 Ave/Wk</th>
              <th onclick="sortSupplierRisk(4)">Days Cover</th>
              <th onclick="sortSupplierRisk(5)">Critical %</th>
              <th onclick="sortSupplierRisk(6)">OOS %</th>
              <th onclick="sortSupplierRisk(7)">Overstock %</th>
              <th onclick="sortSupplierRisk(8)">Dead %</th>
              <th onclick="sortSupplierRisk(9)">Lost Sales/Wk</th>
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
            <button class="btn btn-sm" onclick="exportData('suppliers')">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="suppliers-table">
            <thead><tr>
              <th data-field="supplierCode" onclick="sortTable('suppliers-table',0)">Supplier Code</th>
              <th data-field="supplierName" onclick="sortTable('suppliers-table',1)">Supplier Name</th>
              <th data-field="totalValue" onclick="sortTable('suppliers-table',2)">Inv Value</th>
              <th data-field="totalP8Ave" onclick="sortTable('suppliers-table',3)">P8 Ave/Wk</th>
              <th data-field="totalOnHand" onclick="sortTable('suppliers-table',4)">On Hand</th>
              <th data-field="totalSKUs" onclick="sortTable('suppliers-table',5)">SKUs</th>
              <th data-field="weeksToSell" onclick="sortTable('suppliers-table',6)">WTS</th>
              <th data-field="daysCover" onclick="sortTable('suppliers-table',7)">Days Cover</th>
              <th data-field="oosCount" onclick="sortTable('suppliers-table',8)">Out of Stock</th>
              <th data-field="criticalCount" onclick="sortTable('suppliers-table',9)">Critical</th>
              <th data-field="overstockCount" onclick="sortTable('suppliers-table',10)">Overstock</th>
              <th data-field="deadCount" onclick="sortTable('suppliers-table',11)">Dead Stock</th>
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
            <input type="text" class="table-search" id="sku-upc-input" placeholder="📷 Scan/enter UPC" onkeydown="if(event.key==='Enter')lookupUPC()" style="width:180px;"/>
            <button class="btn btn-sm" onclick="lookupUPC()">Find</button>
            <button class="btn btn-sm btn-green" onclick="openCameraScan()">📷 Camera</button>
            <button class="btn btn-sm" onclick="exportSKUsExcel()" id="sku-export-btn">⬇ Export Excel</button>
            <span id="sku-upc-msg" style="font-size:11px;font-family:'IBM Plex Mono',monospace;"></span>
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
              <th onclick="sortSKUs('lastTransferIn')">Transfer In <span class="sort-ind" data-key="lastTransferIn"></span></th>
              <th onclick="sortSKUs('lastTransferOut')">Transfer Out <span class="sort-ind" data-key="lastTransferOut"></span></th>
            </tr></thead>
            <tbody id="skus-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="skus-pagination"></div>
      </div>
    </div>

    <!-- TOP 300 SKU TAB -->
    <div id="tab-top300" style="display:none;">
      <div class="section">
        <div class="section-header">
          <div class="section-title">⭐ Top 300 SKU <span class="badge badge-green" id="top300-count">0</span>
            <span class="totals-pill" id="top300-totals"></span>
          </div>
          <div class="section-actions">
            <input type="text" class="table-search" placeholder="Search..." oninput="searchTable('top300-table',this.value)"/>
            <button class="btn btn-sm" onclick="exportTop300Excel()" id="top300-export-btn">⬇ Export Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table id="top300-table">
            <thead><tr>
              <th data-field="area" onclick="sortTable('top300-table',0)">Area</th>
              <th data-field="storeName" onclick="sortTable('top300-table',1)">Store Name</th>
              <th data-field="rank" onclick="sortTable('top300-table',2)">Rank</th>
              <th data-field="sku" onclick="sortTable('top300-table',3)">SKU</th>
              <th data-field="itemDescription" onclick="sortTable('top300-table',4)">Item Description</th>
              <th data-field="supplier" onclick="sortTable('top300-table',5)">Supplier</th>
              <th data-field="onHand" onclick="sortTable('top300-table',6)">On Hand</th>
              <th data-field="qtyCases" onclick="sortTable('top300-table',7)">Qty in Cases</th>
              <th data-field="p8ave" onclick="sortTable('top300-table',8)">P8 Ave</th>
              <th data-field="daysCover" onclick="sortTable('top300-table',9)">Days Cover</th>
              <th data-field="status" onclick="sortTable('top300-table',10)">Status</th>
              <th data-field="incomingPO" onclick="sortTable('top300-table',11)">Incoming PO</th>
              <th data-field="lostSalesPerWeek" onclick="sortTable('top300-table',12)">Lost Sales/Wk</th>
              <th data-field="ico" onclick="sortTable('top300-table',13)">ICO</th>
              <th data-field="dateLastSold" onclick="sortTable('top300-table',14)">Last Sold</th>
              <th data-field="dateLastReceived" onclick="sortTable('top300-table',15)">Last Received</th>
              <th data-field="lastTransferIn" onclick="sortTable('top300-table',16)">Transfer In</th>
              <th data-field="lastTransferOut" onclick="sortTable('top300-table',17)">Transfer Out</th>
            </tr></thead>
            <tbody id="top300-body"></tbody>
          </table>
        </div>
        <div class="pagination" id="top300-pagination"></div>
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
let tablePages = { critical:1, overstock:1, aging:1, blackinv:1, negsku:1, deadstock:1, outofstock:1, stores:1, suppliers:1, top300:1 };
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
  stopActivityTracking();
  try { await fetch('/api/logout', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ token: authToken, reason: 'manual' }) }); } catch(e) {}
  authToken = ''; currentUser = null;
  try { sessionStorage.removeItem('camanava_token'); sessionStorage.removeItem('camanava_user'); } catch(e) {}
  location.reload();
}

// ─── INACTIVITY DETECTION & HEARTBEAT ──────────────────────────────────────────
const INACTIVITY_LIMIT_MS = 10 * 60 * 1000;   // 10 minutes total
const WARNING_BEFORE_MS = 60 * 1000;          // Show warning 60s before
const HEARTBEAT_INTERVAL_MS = 60 * 1000;      // Ping server every 60s

let lastActivityAt = Date.now();
let inactivityTimer = null;
let warningTimer = null;
let countdownTimer = null;
let heartbeatTimer = null;

function recordActivity() {
  lastActivityAt = Date.now();
  // If warning is showing, hide it (user came back)
  const modal = document.getElementById('inactivity-modal');
  if (modal && modal.style.display !== 'none') {
    modal.style.display = 'none';
    if (countdownTimer) { clearInterval(countdownTimer); countdownTimer = null; }
  }
}

function startActivityTracking() {
  if (!authToken) return;
  ['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart', 'click'].forEach(ev => {
    document.addEventListener(ev, recordActivity, { passive: true });
  });
  // Main inactivity check every 10 seconds
  if (inactivityTimer) clearInterval(inactivityTimer);
  inactivityTimer = setInterval(checkInactivity, 10 * 1000);
  // Heartbeat every 60 seconds (keeps server session alive)
  if (heartbeatTimer) clearInterval(heartbeatTimer);
  heartbeatTimer = setInterval(sendHeartbeat, HEARTBEAT_INTERVAL_MS);
}

function stopActivityTracking() {
  ['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart', 'click'].forEach(ev => {
    document.removeEventListener(ev, recordActivity);
  });
  if (inactivityTimer) { clearInterval(inactivityTimer); inactivityTimer = null; }
  if (warningTimer) { clearTimeout(warningTimer); warningTimer = null; }
  if (countdownTimer) { clearInterval(countdownTimer); countdownTimer = null; }
  if (heartbeatTimer) { clearInterval(heartbeatTimer); heartbeatTimer = null; }
}

function checkInactivity() {
  const elapsed = Date.now() - lastActivityAt;
  const modal = document.getElementById('inactivity-modal');
  if (elapsed >= INACTIVITY_LIMIT_MS) {
    // Time's up - auto logout
    autoLogout();
  } else if (elapsed >= (INACTIVITY_LIMIT_MS - WARNING_BEFORE_MS)) {
    // Show warning if not already shown
    if (modal && modal.style.display === 'none') {
      showInactivityWarning();
    }
  }
}

function showInactivityWarning() {
  const modal = document.getElementById('inactivity-modal');
  if (!modal) return;
  modal.style.display = 'flex';
  let remaining = Math.ceil((INACTIVITY_LIMIT_MS - (Date.now() - lastActivityAt)) / 1000);
  const cd = document.getElementById('inactivity-countdown');
  if (cd) cd.textContent = remaining;
  if (countdownTimer) clearInterval(countdownTimer);
  countdownTimer = setInterval(() => {
    remaining = Math.ceil((INACTIVITY_LIMIT_MS - (Date.now() - lastActivityAt)) / 1000);
    if (remaining <= 0) {
      clearInterval(countdownTimer);
      autoLogout();
    } else if (cd) {
      cd.textContent = remaining;
    }
  }, 1000);
}

function stayLoggedIn() {
  recordActivity();
  sendHeartbeat();
}

async function autoLogout() {
  stopActivityTracking();
  try {
    await fetch('/api/logout', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: authToken, reason: 'auto-timeout' })
    });
  } catch(e) {}
  authToken = ''; currentUser = null;
  try { sessionStorage.removeItem('camanava_token'); sessionStorage.removeItem('camanava_user'); } catch(e) {}
  alert('You have been logged out due to inactivity.');
  location.reload();
}

async function sendHeartbeat() {
  if (!authToken) return;
  try {
    await fetch('/api/heartbeat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: authToken })
    });
  } catch(e) {}
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
  const params = new URLSearchParams();
  if (authToken) params.set('token', authToken);
  const a = document.createElement('a');
  a.href = '/api/export-logs-xlsx?' + params.toString();
  a.download = '';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
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
  // Start inactivity detection & heartbeat
  lastActivityAt = Date.now();
  startActivityTracking();
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

  populateSelect('f-category', d.categories, 'All Categories');
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
  const idMap = { category: 'f-category', area: 'f-area', store: 'f-store', dept: 'f-dept', subDept: 'f-subdept', cls: 'f-cls', supplier: 'f-supplier' };
  const sel = document.getElementById(idMap[key]);
  if (sel) sel.value = '';
  applyFilter(key, '');
}

function clearFilters() {
  activeFilters = {};
  ['f-category','f-store','f-dept','f-subdept','f-cls','f-supplier']
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
  const wtsColor = d.avgWts < 4 ? 'red' : d.avgWts > 12 ? 'yellow' : 'green';
  const daysCover = (d.daysCover != null ? d.daysCover : (d.avgWts || 0) * 7);
  grid.innerHTML = [
    // GREEN
    kpiCard('Total Inv Value', '₱' + fmtM(d.totalOnHandValue), 'w/ VAT', 'green'),
    kpiCard('Active Stores', fmt(d.activeStores), 'stores', 'green'),
    kpiCard('Avg WTS Net', (d.avgWts || 0).toFixed(1) + ' wks', 'weeks to sell', wtsColor),
    kpiCard('Days Cover', daysCover.toFixed(0) + ' days', 'avg coverage', wtsColor),
    // BLUE
    kpiCard('On Hand Qty', fmt(d.totalOnHand), 'units', 'blue'),
    kpiCard('Suppliers', fmt(d.activeSuppliers), 'active', 'blue'),
    kpiCard('PO Value', '₱' + fmtM(d.totalPOValue), 'incoming', 'blue'),
    kpiCard('Total SKUs', fmt(d.totalSKUs), 'in scope', 'blue'),
    // YELLOW
    kpiCard('Overstock SKUs', fmt(d.overstockCount), 'WTS > 12 weeks', 'yellow'),
    kpiCard('Overstock Value', '₱' + fmtM(d.overstockValue || 0), fmt(d.overstockCount || 0) + ' SKUs', 'yellow'),
    kpiCard('Aging Value', '₱' + fmtM(d.agingValue || 0), fmt(d.agingCount || 0) + ' SKUs, 180+ days', 'yellow'),
    kpiCard('Transfer Value', '₱' + fmtM(d.totalTRFValue), 'incoming', 'yellow'),
    // RED
    kpiCard('Out of Stock', fmt(d.outOfStockCount || 0), 'SKUs losing sales', 'red'),
    kpiCard('Lost Sales/Wk', '₱' + fmtM(d.totalLostSalesPerWeek || 0), 'estimated/week', 'red'),
    kpiCard('Critical SKUs', fmt(d.criticalCount), 'WTS < 2 weeks', 'red'),
    kpiCard('Dead Stock SKUs', fmt(d.deadStockCount), 'No sales 8 wks', 'red'),
    kpiCard('Black Inv Value', '₱' + fmtM(d.blackInventoryValue || 0), fmt(d.blackInventoryCount || 0) + ' SKUs, stagnant', 'red'),
    kpiCard('P8 Wks No Sales Val', '₱' + fmtM(d.deadStockValue || 0), fmt(d.deadStockCount || 0) + ' SKUs', 'red'),
  ].join('');
}

function kpiCard(label, value, sub, type) {
  return '<div class="kpi-card ' + type + '"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>';
}

// ─── CHARTS ───────────────────────────────────────────────────────────────────
async function loadCharts() {
  const [storesRes, suppliersRes, categoriesRes] = await Promise.all([
    fetch('/api/stores' + filterQuery()),
    fetch('/api/suppliers' + filterQuery()),
    fetch('/api/categories' + filterQuery())
  ]);
  const stores = await storesRes.json();
  const suppliers = await suppliersRes.json();
  const categories = await categoriesRes.json();
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

  // Inventory Value by Category
  const cats = Array.isArray(categories) ? categories : [];
  const catNames = cats.map(c => c.catName);
  const catColors = ['#2ea043','#1f6feb','#e3b341','#f85149','#8b949e','#c084fc','#38bdf8','#fb923c','#a3e635'];
  charts.categories = new Chart(document.getElementById('chart-categories'), {
    type: 'bar',
    data: {
      labels: catNames,
      datasets: [{ label: 'Inv Value', data: cats.map(c => c.totalValue), backgroundColor: catColors.slice(0, catNames.length), borderRadius: 4 }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        datalabels: { anchor:'end', align:'top', color:'#e6edf3', font:{size:10,weight:'600'}, formatter: v => '₱'+fmtM(v) }
      },
      layout: { padding: { top: 20 } },
      scales: { y: { ticks: { color:'#8b949e', callback: v => '₱'+fmtM(v) }, grid: { color:'#30363d' } }, x: { ticks: { color:'#8b949e', font:{size:9} }, grid: { display: false } } }
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

// ─── TOP 300 SKU STORE PERFORMANCE (Overview) ─────────────────────────────────
let top300PerfData = [];
async function loadTop300Performance() {
  const r = await fetch('/api/top300-store-metrics' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  top300PerfData = data;
  renderTop300Perf(data);
}

function renderTop300Perf(data) {
  const tbody = document.getElementById('top300-perf-body');
  if (!tbody) return;
  if (data.length === 0) { tbody.innerHTML = '<tr><td colspan="9" class="empty">No data</td></tr>'; return; }
  const totalStores = data.length;
  const topThird = Math.ceil(totalStores / 3);
  const bottomThird = totalStores - topThird;
  tbody.innerHTML = data.map((s, i) => {
    // Color the rank based on position: top third green, mid yellow, bottom third red
    let rankColor = '#3fb950';      // green (best)
    if (i >= topThird && i < bottomThird) rankColor = '#e3b341'; // yellow (middle)
    if (i >= bottomThird) rankColor = '#f85149';                  // red (worst)
    // Score pill color follows the same scheme
    const scoreColor = s.score < 5 ? '#3fb950' : s.score < 15 ? '#e3b341' : '#f85149';
    return '<tr>' +
      '<td class="mono" style="font-weight:700;color:' + rankColor + ';font-size:14px;">#' + s.rank + '</td>' +
      '<td class="mono" style="font-weight:600;">' + esc(s.storeNumber) + '</td>' +
      '<td>' + esc(s.storeName) + '</td>' +
      '<td><span class="badge badge-blue">' + esc(s.area) + '</span></td>' +
      '<td class="mono">' + fmt(s.total) + '</td>' +
      '<td class="mono"><span style="color:#f85149;font-weight:600;">' + s.oosPct.toFixed(1) + '%</span> <span style="color:var(--text2);font-size:10px;">(' + fmt(s.oos) + ')</span></td>' +
      '<td class="mono"><span style="color:#e3b341;font-weight:600;">' + s.criticalPct.toFixed(1) + '%</span> <span style="color:var(--text2);font-size:10px;">(' + fmt(s.critical) + ')</span></td>' +
      '<td class="mono"><span style="color:#3fb950;font-weight:600;">' + s.healthyPct.toFixed(1) + '%</span> <span style="color:var(--text2);font-size:10px;">(' + fmt(s.healthy) + ')</span></td>' +
      '<td class="mono"><span style="background:rgba(' + (scoreColor === '#3fb950' ? '63,185,80' : scoreColor === '#e3b341' ? '227,179,65' : '248,81,73') + ',0.15);color:' + scoreColor + ';padding:3px 10px;border-radius:4px;font-weight:700;">' + s.score.toFixed(2) + '</span></td>' +
      '</tr>';
  }).join('');
}

let top300PerfSortState = {};
function sortTop300Perf(colIndex) {
  const keys = ['rank','storeNumber','storeName','area','total','oosPct','criticalPct','healthyPct','score'];
  const key = keys[colIndex];
  const asc = top300PerfSortState[colIndex] !== true;
  top300PerfSortState[colIndex] = asc;
  const sorted = [...top300PerfData].sort((a, b) => {
    const av = a[key], bv = b[key];
    if (av == null && bv == null) return 0;
    if (av == null) return asc ? 1 : -1;
    if (bv == null) return asc ? -1 : 1;
    if (typeof av === 'number' && typeof bv === 'number') return asc ? av - bv : bv - av;
    return asc ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
  });
  renderTop300Perf(sorted);
}

// ─── TABLE LOADING ────────────────────────────────────────────────────────────
async function loadTabData() {
  if (activeTab === 'overview') { await loadCharts(); await loadRiskMatrix(); await loadTop300Performance(); return; }
  if (activeTab === 'critical') await loadCritical();
  if (activeTab === 'overstock') await loadOverstock();
  if (activeTab === 'aging') await loadAging();
  if (activeTab === 'blackinv') await loadBlackInventory();
  if (activeTab === 'negsku') await loadNegativeSKUs();
  if (activeTab === 'deadstock') await loadDeadstock();
  if (activeTab === 'outofstock') await loadOutOfStock();
  if (activeTab === 'stores') await loadStores();
  if (activeTab === 'suppliers') await loadSuppliers();
  if (activeTab === 'skus') await loadSKUs(1);
  if (activeTab === 'top300') await loadTop300();
  if (activeTab === 'logs') await loadLogs();
}

// Stores the full unpaginated dataset for each table — used by sortTable to sort across all pages
const tableData = {};

async function loadCritical() {
  const r = await fetch('/api/critical' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.critical = data;
  document.getElementById('critical-count').textContent = fmt(data.length);
  renderTable('critical-body', data, renderCriticalRow, 'critical-pagination', 'critical', tablePages.critical);
}
async function loadOverstock() {
  const r = await fetch('/api/overstock' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.overstock = data;
  document.getElementById('overstock-count').textContent = fmt(data.length);
  setTotalsPill('overstock-totals', data);
  renderTable('overstock-body', data, renderOverstockRow, 'overstock-pagination', 'overstock', tablePages.overstock);
}
async function loadAging() {
  const r = await fetch('/api/aging' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.aging = data;
  document.getElementById('aging-count').textContent = fmt(data.length);
  setTotalsPill('aging-totals', data);
  renderTable('aging-body', data, renderAgingRow, 'aging-pagination', 'aging', tablePages.aging);
}
async function loadBlackInventory() {
  const r = await fetch('/api/blackinventory' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.blackinv = data;
  document.getElementById('blackinv-count').textContent = fmt(data.length);
  setTotalsPill('blackinv-totals', data);
  renderTable('blackinv-body', data, renderBlackInventoryRow, 'blackinv-pagination', 'blackinv', tablePages.blackinv);
}
async function loadNegativeSKUs() {
  const r = await fetch('/api/negativeskus' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.negsku = data;
  document.getElementById('negsku-count').textContent = fmt(data.length);
  setTotalsPill('negsku-totals', data);
  renderTable('negsku-body', data, renderNegativeSKURow, 'negsku-pagination', 'negsku', tablePages.negsku);
  // Build summaries
  buildNegSkuSummaries(data);
}

// Aggregates the Negative SKU data by store & category, renders both summary tables
let negSkuSummary = { store: [], category: [], sort: { store: { col:2, dir:'asc' }, category: { col:2, dir:'asc' } } };
function buildNegSkuSummaries(data) {
  const byStore = {}, byCat = {};
  for (const r of data) {
    if (!byStore[r.storeName]) byStore[r.storeName] = { key: r.storeName, onHand: 0, value: 0, count: 0 };
    byStore[r.storeName].onHand += Number(r.onHand) || 0;
    byStore[r.storeName].value += Number(r.onHandValue) || 0;
    byStore[r.storeName].count++;
    const cat = r.catName || 'Uncategorized';
    if (!byCat[cat]) byCat[cat] = { key: cat, onHand: 0, value: 0, count: 0 };
    byCat[cat].onHand += Number(r.onHand) || 0;
    byCat[cat].value += Number(r.onHandValue) || 0;
    byCat[cat].count++;
  }
  negSkuSummary.store = Object.values(byStore);
  negSkuSummary.category = Object.values(byCat);
  renderNegSkuSummary('store');
  renderNegSkuSummary('category');
}

function renderNegSkuSummary(which) {
  const tbodyId = which === 'store' ? 'negsku-summary-store' : 'negsku-summary-cat';
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  const rows = negSkuSummary[which] || [];
  // Apply current sort
  const sort = negSkuSummary.sort[which];
  const fields = ['key', 'onHand', 'value', 'count'];
  const f = fields[sort.col];
  const sorted = [...rows].sort((a, b) => {
    let av = a[f], bv = b[f];
    if (typeof av === 'number' && typeof bv === 'number') return sort.dir === 'asc' ? av - bv : bv - av;
    return sort.dir === 'asc' ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
  });
  let html = sorted.map(r =>
    '<tr>' +
      '<td>' + esc(r.key) + '</td>' +
      '<td class="num neg">' + fmt(r.onHand) + '</td>' +
      '<td class="num">₱' + fmtN(r.value) + '</td>' +
      '<td class="num">' + fmt(r.count) + '</td>' +
    '</tr>'
  ).join('');
  // Footer totals
  if (sorted.length > 0) {
    const tOnHand = sorted.reduce((s,r) => s + r.onHand, 0);
    const tValue = sorted.reduce((s,r) => s + r.value, 0);
    const tCount = sorted.reduce((s,r) => s + r.count, 0);
    html += '<tr style="font-weight:700;background:var(--bg3);border-top:2px solid var(--green-bright);">' +
      '<td>TOTAL</td>' +
      '<td class="num neg">' + fmt(tOnHand) + '</td>' +
      '<td class="num">₱' + fmtN(tValue) + '</td>' +
      '<td class="num">' + fmt(tCount) + '</td>' +
      '</tr>';
  } else {
    html = '<tr><td colspan="4" style="text-align:center;color:var(--text2);padding:12px;">No data</td></tr>';
  }
  tbody.innerHTML = html;
}

function sortSummary(which, colIdx) {
  const sort = negSkuSummary.sort[which];
  if (sort.col === colIdx) sort.dir = sort.dir === 'asc' ? 'desc' : 'asc';
  else { sort.col = colIdx; sort.dir = colIdx === 0 ? 'asc' : 'desc'; }
  renderNegSkuSummary(which);
}

async function loadTop300() {
  const r = await fetch('/api/top300skus' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.top300 = data;
  document.getElementById('top300-count').textContent = fmt(data.length);
  // Custom totals pill for Top 300 (uses onHand + invValue derived from onHand × no value)
  // We'll show On Hand total + a rough Inv Value if available
  const totals = document.getElementById('top300-totals');
  if (totals) {
    const tOnHand = data.reduce((s, x) => s + (Number(x.onHand) || 0), 0);
    const tIncomingPO = data.reduce((s, x) => s + (Number(x.incomingPO) || 0), 0);
    totals.innerHTML = '<span class="tp-label">Total On Hand:</span> <span class="tp-value">' + fmt(tOnHand) +
                       '</span> &nbsp;|&nbsp; <span class="tp-label">Total Incoming PO:</span> <span class="tp-value">' + fmt(tIncomingPO) + '</span>';
  }
  renderTable('top300-body', data, renderTop300Row, 'top300-pagination', 'top300', tablePages.top300);
}

// Sums onHand & onHandValue and shows them in the section header
function setTotalsPill(elId, data) {
  const el = document.getElementById(elId);
  if (!el || !Array.isArray(data)) return;
  let totalOnHand = 0, totalValue = 0;
  for (const r of data) {
    totalOnHand += Number(r.onHand) || 0;
    totalValue += Number(r.onHandValue) || 0;
  }
  el.innerHTML = '<span class="tp-label">Total On Hand:</span> <span class="tp-value">' + fmt(totalOnHand) +
                 '</span> &nbsp;|&nbsp; <span class="tp-label">Total Value:</span> <span class="tp-value green">₱' + fmtN(totalValue) + '</span>';
}
async function loadDeadstock() {
  const r = await fetch('/api/deadstock' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.deadstock = data;
  document.getElementById('deadstock-count').textContent = fmt(data.length);
  setTotalsPill('deadstock-totals', data);
  renderTable('deadstock-body', data, renderDeadstockRow, 'deadstock-pagination', 'deadstock', tablePages.deadstock);
}
async function loadOutOfStock() {
  const r = await fetch('/api/outofstock' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.outofstock = data;
  document.getElementById('outofstock-count').textContent = fmt(data.length);
  renderTable('outofstock-body', data, renderOutOfStockRow, 'outofstock-pagination', 'outofstock', tablePages.outofstock);
}
async function loadStores() {
  const r = await fetch('/api/stores' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  tableData.stores = data;
  renderTable('stores-body', data, renderStoreRow, 'stores-pagination', 'stores', tablePages.stores);
}
async function loadSuppliers() {
  const r = await fetch('/api/suppliers' + filterQuery());
  const data = await r.json();
  if (!Array.isArray(data)) return;
  supplierRiskData = data;
  tableData.suppliers = data;
  renderSupplierRisk(data);
  renderTable('suppliers-body', data, renderSupplierRow, 'suppliers-pagination', 'suppliers', tablePages.suppliers);
}

let supplierRiskData = [];
function renderSupplierRisk(data) {
  const tbody = document.getElementById('supplier-risk-body');
  if (!tbody) return;
  if (data.length === 0) { tbody.innerHTML = '<tr><td colspan="10" class="empty">No data</td></tr>'; return; }
  tbody.innerHTML = data.map(s => {
    return '<tr>' +
      '<td>' + esc(s.supplierName) + '</td>' +
      '<td class="mono">' + fmt(s.totalSKUs) + '</td>' +
      '<td class="mono" style="color:var(--green-bright);">₱' + fmtN(s.totalValue) + '</td>' +
      '<td class="mono">' + fmtN(s.totalP8Ave || 0) + '</td>' +
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
  const keys = ['supplierName','totalSKUs','totalValue','totalP8Ave','daysCover','criticalPct','oosPct','overstockPct','deadPct','totalLostSales'];
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

async function lookupUPC() {
  const input = document.getElementById('sku-upc-input');
  const msgEl = document.getElementById('sku-upc-msg');
  const upc = (input.value || '').trim();
  if (!upc) { msgEl.innerHTML = ''; return; }
  msgEl.innerHTML = '<span style="color:var(--text2);">Looking up...</span>';
  try {
    const r = await fetch('/api/upc-lookup?upc=' + encodeURIComponent(upc));
    if (!r.ok) {
      if (r.status === 404) {
        msgEl.innerHTML = '<span style="color:var(--red-light);">UPC ' + upc + ' not found</span>';
      } else if (r.status === 503) {
        msgEl.innerHTML = '<span style="color:var(--yellow-light);">UPC database not ready</span>';
      } else {
        msgEl.innerHTML = '<span style="color:var(--red-light);">Lookup error</span>';
      }
      return;
    }
    const d = await r.json();
    // Found - put SKU into the regular search box and trigger search
    const searchBox = document.getElementById('sku-search-input');
    searchBox.value = d.sku;
    msgEl.innerHTML = '<span style="color:var(--green-bright);">✓ ' + esc(d.sku) + (d.desc ? ' — ' + esc(d.desc) : '') + '</span>';
    input.value = '';
    // Clear status filter to ensure SKU shows regardless of status
    const statusEl = document.getElementById('sku-status-filter');
    if (statusEl) statusEl.value = '';
    loadSKUs(1);
    input.focus();
  } catch (e) {
    msgEl.innerHTML = '<span style="color:var(--red-light);">Connection error</span>';
  }
}

// ─── CAMERA BARCODE SCAN ──────────────────────────────────────────────────────
let camScanner = null;
let camScanning = false;

async function openCameraScan() {
  if (typeof Html5Qrcode === 'undefined') {
    alert('Camera scanner library not loaded. Check your internet connection.');
    return;
  }
  const modal = document.getElementById('camera-modal');
  const status = document.getElementById('camera-status');
  modal.style.display = 'flex';
  status.textContent = 'Initializing camera...';
  try {
    if (!camScanner) camScanner = new Html5Qrcode('camera-reader');
    camScanning = true;
    await camScanner.start(
      { facingMode: 'environment' }, // rear camera
      {
        fps: 10,
        qrbox: function(w, h) {
          // Wide rectangle for barcodes; sized for both landscape and portrait
          const minEdge = Math.min(w, h);
          const boxW = Math.floor(Math.min(w * 0.9, 320));
          const boxH = Math.floor(Math.min(boxW * 0.5, 160));
          return { width: boxW, height: boxH };
        }
      },
      onCameraScanSuccess,
      // onScanFailure - called every frame when nothing found; ignore quietly
      function() {}
    );
    status.textContent = '📷 Ready — point at barcode';
  } catch (e) {
    console.error('Camera start error:', e);
    let msg = 'Cannot open camera.';
    if (e && (e.toString().toLowerCase().includes('permission') || e.toString().toLowerCase().includes('notallowed'))) {
      msg = 'Camera permission denied. Allow camera access in browser settings.';
    } else if (e && e.toString().toLowerCase().includes('notfound')) {
      msg = 'No camera found on this device.';
    } else if (location.protocol !== 'https:') {
      msg = 'Camera requires HTTPS. Use the https:// URL.';
    }
    status.innerHTML = '<span style="color:var(--red-light);">' + msg + '</span>';
  }
}

function onCameraScanSuccess(decodedText, decodedResult) {
  if (!camScanning) return;
  camScanning = false; // prevent re-fire while we process
  playBeep();
  const status = document.getElementById('camera-status');
  status.innerHTML = '<span style="color:var(--green-bright);">✓ Detected: ' + esc(decodedText) + '</span>';
  // Put scanned code into UPC field and run lookup
  const upcInput = document.getElementById('sku-upc-input');
  if (upcInput) upcInput.value = decodedText;
  // Brief delay so user sees the "Detected" feedback
  setTimeout(() => {
    closeCameraScan();
    lookupUPC();
  }, 400);
}

// Short scanner-style beep using Web Audio API (no external audio file needed)
let beepCtx = null;
function playBeep() {
  try {
    if (!beepCtx) beepCtx = new (window.AudioContext || window.webkitAudioContext)();
    // Resume context if it was suspended (browser autoplay policy)
    if (beepCtx.state === 'suspended') beepCtx.resume();
    const osc = beepCtx.createOscillator();
    const gain = beepCtx.createGain();
    osc.type = 'sine';
    osc.frequency.value = 1000; // 1kHz — classic scanner beep
    gain.gain.setValueAtTime(0.0001, beepCtx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.3, beepCtx.currentTime + 0.01);
    gain.gain.exponentialRampToValueAtTime(0.0001, beepCtx.currentTime + 0.15);
    osc.connect(gain);
    gain.connect(beepCtx.destination);
    osc.start();
    osc.stop(beepCtx.currentTime + 0.15);
  } catch (e) {
    // Silently fail if audio not supported
  }
}

async function closeCameraScan() {
  const modal = document.getElementById('camera-modal');
  camScanning = false;
  if (camScanner) {
    try { await camScanner.stop(); } catch(e) {}
    try { await camScanner.clear(); } catch(e) {}
  }
  modal.style.display = 'none';
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
  if (rows.length === 0) { tbody.innerHTML = '<tr><td colspan="19" class="empty">No data found</td></tr>'; return; }
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
      '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
      '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
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
  // If we have cached data for this table, just re-render (preserves search/sort)
  if (Array.isArray(tableData[key]) && renderFromCache(key)) return;
  // Fallback for tables without cache
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
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
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
    '<td class="mono">' + (r.qtyCases === 'Per Piece' ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>' : fmtN(r.qtyCases)) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono" style="color:var(--yellow-light);font-weight:600;">' + r.wtsNet + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
    '<td><span class="action-badge ' + ac + '">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderAgingRow(r) {
  const dc = r.daysCover != null ? r.daysCover.toFixed(0) + 'd' : '—';
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">' + (r.qtyCases === 'Per Piece' ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>' : fmtN(r.qtyCases)) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono" style="color:var(--yellow-light);font-weight:600;">' + dc + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
    '<td><span class="action-badge action-markdown">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderBlackInventoryRow(r) {
  const dc = r.daysCover != null ? r.daysCover.toFixed(0) + 'd' : '—';
  const lastSold = r.dateLastSold || '<span style="color:var(--text2);font-style:italic;">Never</span>';
  return '<tr>' +
    '<td>' + esc(r.store) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">' + (r.qtyCases === 'Per Piece' ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>' : fmtN(r.qtyCases)) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono" style="color:var(--red-light);font-weight:600;">' + dc + '</td>' +
    '<td class="mono">' + lastSold + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
    '<td><span class="action-badge action-urgent">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderNegativeSKURow(r) {
  return '<tr>' +
    '<td>' + esc(r.storeName) + '</td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono" style="color:var(--red-light);font-weight:600;">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">' + esc(r.qtyCases) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono">' + fmtN(r.p8ave) + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '</tr>';
}
function renderTop300Row(r) {
  const dc = r.daysCover != null ? r.daysCover.toFixed(0) + 'd' : '—';
  const p8 = r.p8ave != null ? fmtN(r.p8ave) : '—';
  const onHand = r.onHand != null ? fmt(r.onHand) : '—';
  const incPO = r.incomingPO != null ? fmt(r.incomingPO) : '—';
  const lostSales = r.lostSalesPerWeek != null ? '₱' + fmtN(r.lostSalesPerWeek) : '—';
  let statusClass = '';
  if (r.status === 'Critical') statusClass = 'status-critical';
  else if (r.status === 'OOS') statusClass = 'status-oos';
  else if (r.status === 'Overstock') statusClass = 'status-overstock';
  else if (r.status === 'Dead Stock') statusClass = 'status-dead';
  else if (r.status === 'Not Found') statusClass = 'status-oos';
  else statusClass = 'status-normal';
  const onHandStyle = (r.onHand !== null && r.onHand < 0) ? ' style="color:var(--red-light);font-weight:600;"' : '';
  const dcColor = r.daysCover != null && r.daysCover < 7 ? 'color:var(--red-light);' :
                  r.daysCover != null && r.daysCover > 90 ? 'color:var(--yellow-light);' : '';
  const qtyCases = r.qtyCases == null ? '—' : (r.qtyCases === 'Per Piece' ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>' : fmtN(r.qtyCases));
  return '<tr>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td>' + esc(r.storeName) + '</td>' +
    '<td class="mono" style="font-weight:700;color:var(--green-bright);">' + (r.rank != null ? '#' + r.rank : '—') + '</td>' +
    '<td class="mono">' + esc(r.sku) + '</td>' +
    '<td>' + esc(r.itemDescription) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono"' + onHandStyle + '>' + onHand + '</td>' +
    '<td class="mono">' + qtyCases + '</td>' +
    '<td class="mono">' + p8 + '</td>' +
    '<td class="mono" style="' + dcColor + '">' + dc + '</td>' +
    '<td><span class="' + statusClass + '">' + esc(r.status) + '</span></td>' +
    '<td class="mono">' + incPO + '</td>' +
    '<td class="mono">' + lostSales + '</td>' +
    '<td class="mono">' + esc(r.ico || '—') + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
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
    '<td class="mono">' + (r.qtyCases === 'Per Piece' ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>' : fmtN(r.qtyCases)) + '</td>' +
    '<td class="mono">₱' + fmtN(r.onHandValue) + '</td>' +
    '<td class="mono" style="color:var(--text2);">' + wts + '</td>' +
    '<td class="mono" style="color:var(--text2);">' + dc + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
    '<td><span class="action-badge action-markdown">' + esc(r.action) + '</span></td>' +
    '</tr>';
}
function renderOutOfStockRow(r) {
  const qtyCases = r.qtyCases === 'Per Piece' ? '<span style="color:var(--text2);font-style:italic;">Per Piece</span>' : fmtN(r.qtyCases);
  const wts = r.weeksToSell != null ? r.weeksToSell.toFixed(2) : '0.00';
  const dc = r.daysCover != null ? r.daysCover.toFixed(0) + 'd' : '—';
  const p8 = r.p8ave != null ? fmtN(r.p8ave) : '—';
  return '<tr>' +
    '<td class="mono" style="font-weight:600;">' + esc(r.storeNumber) + '</td>' +
    '<td>' + esc(r.storeName) + '</td>' +
    '<td><span class="badge badge-blue">' + esc(r.area) + '</span></td>' +
    '<td class="mono">' + esc(r.skuCode) + '</td>' +
    '<td>' + esc(r.skuDesc) + '</td>' +
    '<td>' + esc(r.supplier) + '</td>' +
    '<td class="mono">' + fmt(r.onHand) + '</td>' +
    '<td class="mono">' + fmt(r.stdPack) + '</td>' +
    '<td class="mono">' + qtyCases + '</td>' +
    '<td class="mono">₱' + fmtN(r.invValue) + '</td>' +
    '<td class="mono">' + p8 + '</td>' +
    '<td class="mono">' + wts + '</td>' +
    '<td class="mono">' + dc + '</td>' +
    '<td><span class="status-oos">OOS</span></td>' +
    '<td class="mono" style="color:var(--red-light);font-weight:600;">₱' + fmtN(r.lostSalesPerWeek) + '</td>' +
    '<td class="mono">' + esc(r.ico || '—') + '</td>' +
    '<td class="mono">' + fmt(r.poOrderGR) + '</td>' +
    '<td class="mono">' + fmt(r.trfOrderGR) + '</td>' +
    '<td class="mono">' + esc(r.dateLastSold) + '</td>' +
    '<td class="mono">' + esc(r.dateLastReceived) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferIn) + '</td>' +
    '<td class="mono">' + esc(r.lastTransferOut) + '</td>' +
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
    '<td class="mono">' + fmtN(r.totalP8Ave || 0) + '</td>' +
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
  ['overview','outofstock','critical','overstock','aging','blackinv','negsku','deadstock','stores','suppliers','skus','top300','logs'].forEach(t => {
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
// Stores current search query per table key (preserved across sort/pagination)
const tableSearch = {};
const tableKeyToId = {
  critical: 'critical-table',
  overstock: 'overstock-table',
  aging: 'aging-table',
  blackinv: 'blackinv-table',
  negsku: 'negsku-table',
  deadstock: 'deadstock-table',
  outofstock: 'outofstock-table',
  stores: 'stores-table',
  suppliers: 'suppliers-table',
  top300: 'top300-table'
};

// Re-render a table from cached data, applying current search filter
function renderFromCache(key) {
  const tableId = tableKeyToId[key];
  if (!tableId) return false;
  const cfg = getTableConfig(tableId);
  if (!cfg) return false;
  const fullData = tableData[key];
  if (!Array.isArray(fullData)) return false;
  const q = (tableSearch[key] || '').toLowerCase().trim();
  let view = fullData;
  if (q) {
    // Prefer fields declared via data-field on the <th>s (resilient to column add/reorder)
    const ths = document.querySelectorAll('#' + tableId + ' thead th');
    let fields = Array.from(ths).map(th => th.dataset && th.dataset.field).filter(Boolean);
    if (fields.length === 0) fields = cfg.cols;
    view = fullData.filter(r => fields.some(field => {
      if (!field) return false;
      const v = r[field];
      if (v == null) return false;
      return String(v).toLowerCase().includes(q);
    }));
  }
  const bodyId = cfg.pagination.replace('-pagination', '-body');
  renderTable(bodyId, view, cfg.render, cfg.pagination, key, tablePages[key] || 1);
  // Update count badge if exists
  const countEl = document.getElementById(key + '-count');
  if (countEl) countEl.textContent = fmt(view.length);
  return true;
}

function searchTable(tableId, query) {
  const cfg = getTableConfig(tableId);
  if (cfg && Array.isArray(tableData[cfg.key])) {
    tableSearch[cfg.key] = query || '';
    tablePages[cfg.key] = 1;
    renderFromCache(cfg.key);
    return;
  }
  // Fallback for tables without config (Risk Matrix, Supplier Risk, Activity Log)
  const q = (query || '').toLowerCase();
  document.querySelectorAll('#' + tableId + ' tbody tr').forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
}

// ─── TABLE SORT ───────────────────────────────────────────────────────────────
let sortState = {};
// Maps tableId -> { key (matches tableData / pagination key), columns: [field,...], renderFn, paginationId }
// Field == '' means that column is not sortable. Column order MUST match the <th> order in the table.
function getTableConfig(tableId) {
  const configs = {
    'critical-table':   { key: 'critical',   render: renderCriticalRow,   pagination: 'critical-pagination',
      cols: ['store','area','skuCode','skuDesc','supplier','onHand','onHandValue','currentWkSales','p8ave','wtsNet','totalPO','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'overstock-table':  { key: 'overstock',  render: renderOverstockRow,  pagination: 'overstock-pagination',
      cols: ['store','area','skuCode','skuDesc','supplier','onHand','qtyCases','onHandValue','p8ave','wtsNet','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'aging-table':      { key: 'aging',      render: renderAgingRow,      pagination: 'aging-pagination',
      cols: ['store','area','skuCode','skuDesc','supplier','onHand','qtyCases','onHandValue','p8ave','daysCover','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'blackinv-table':   { key: 'blackinv',   render: renderBlackInventoryRow, pagination: 'blackinv-pagination',
      cols: ['store','area','skuCode','skuDesc','supplier','onHand','qtyCases','onHandValue','p8ave','daysCover','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'negsku-table':     { key: 'negsku',     render: renderNegativeSKURow,    pagination: 'negsku-pagination',
      cols: ['storeName','skuCode','skuDesc','supplier','onHand','qtyCases','onHandValue','p8ave','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'deadstock-table':  { key: 'deadstock',  render: renderDeadstockRow,  pagination: 'deadstock-pagination',
      cols: ['store','area','skuCode','skuDesc','supplier','onHand','qtyCases','onHandValue','weeksToSell','daysCover','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'outofstock-table': { key: 'outofstock', render: renderOutOfStockRow, pagination: 'outofstock-pagination',
      cols: ['storeNumber','storeName','area','skuCode','skuDesc','supplier','onHand','stdPack','qtyCases','invValue','p8ave','weeksToSell','daysCover','status','lostSalesPerWeek','ico','poOrderGR','trfOrderGR','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] },
    'stores-table':     { key: 'stores',     render: renderStoreRow,      pagination: 'stores-pagination',
      cols: ['storeNumber','storeName','area','totalValue','totalOnHand','totalSKUs','weeksToSell','daysCover','oosCount','totalLostSales','criticalCount','overstockCount','deadCount'] },
    'suppliers-table':  { key: 'suppliers',  render: renderSupplierRow,   pagination: 'suppliers-pagination',
      cols: ['supplierCode','supplierName','totalValue','totalP8Ave','totalOnHand','totalSKUs','weeksToSell','daysCover','oosCount','criticalCount','overstockCount','deadCount'] },
    'top300-table':     { key: 'top300',     render: renderTop300Row,     pagination: 'top300-pagination',
      cols: ['area','storeName','rank','sku','itemDescription','supplier','onHand','qtyCases','p8ave','daysCover','status','incomingPO','lostSalesPerWeek','ico','dateLastSold','dateLastReceived','lastTransferIn','lastTransferOut'] }
  };
  return configs[tableId];
}

function sortTable(tableId, colIndex) {
  const cfg = getTableConfig(tableId);
  if (!cfg) return;
  // Prefer data-field on the actual <th> (resilient to column add/reorder).
  // Falls back to cfg.cols[colIndex] for tables that haven't migrated yet.
  const ths = document.querySelectorAll('#' + tableId + ' thead th');
  const th = ths[colIndex];
  const field = (th && th.dataset && th.dataset.field) || cfg.cols[colIndex];
  if (!field) return;
  const data = tableData[cfg.key];
  if (!Array.isArray(data) || data.length === 0) return;
  const stateKey = tableId + '_' + colIndex;
  const asc = sortState[stateKey] !== true;
  sortState[stateKey] = asc;
  // Reset other column states for this table so only one column shows sorted
  Object.keys(sortState).forEach(k => { if (k.startsWith(tableId + '_') && k !== stateKey) delete sortState[k]; });
  // Sort full dataset
  data.sort((a, b) => {
    let av = a[field], bv = b[field];
    // Handle wtsNet "Dead Stock" string for overstock — treat as Infinity so it goes to extremes
    if (av === 'Dead Stock') av = Infinity;
    if (bv === 'Dead Stock') bv = Infinity;
    // Nulls always at bottom regardless of direction
    const aNull = (av == null || av === '');
    const bNull = (bv == null || bv === '');
    if (aNull && bNull) return 0;
    if (aNull) return 1;
    if (bNull) return -1;
    // Try numeric coercion
    const an = (typeof av === 'number') ? av : parseFloat(String(av).replace(/[₱,]/g, ''));
    const bn = (typeof bv === 'number') ? bv : parseFloat(String(bv).replace(/[₱,]/g, ''));
    const aNum = !isNaN(an);
    const bNum = !isNaN(bn);
    // If either side is numeric, sort numerically and push non-numeric values
    // (e.g. "Per Piece", "N/A", text labels) to the bottom regardless of direction.
    if (aNum && bNum) return asc ? an - bn : bn - an;
    if (aNum) return -1;
    if (bNum) return 1;
    return asc ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
  });
  // Reset to page 1 and re-render (renderFromCache will reapply current search)
  tablePages[cfg.key] = 1;
  renderFromCache(cfg.key);
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function exportData(type) {
  // Maps type → tab key for tableData/tableSearch lookups
  const typeToKey = {
    critical: 'critical', overstock: 'overstock', aging: 'aging',
    blackinventory: 'blackinv', deadstock: 'deadstock', outofstock: 'outofstock',
    stores: 'stores', suppliers: 'suppliers'
  };
  const key = typeToKey[type] || type;
  // Build URL with filters + tab-specific search
  const params = new URLSearchParams(activeFilters);
  if (authToken) params.set('token', authToken);
  const search = tableSearch && tableSearch[key] ? tableSearch[key] : '';
  if (search) params.set('search', search);
  const url = '/api/export/' + type + '?' + params.toString();
  // Trigger download via hidden link (better than window.open which can be blocked)
  const a = document.createElement('a');
  a.href = url;
  a.download = '';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

async function exportSKUsExcel() {
  const btn = document.getElementById('sku-export-btn');
  const orig = btn ? btn.innerHTML : null;
  if (btn) { btn.innerHTML = '⏳ Generating...'; btn.disabled = true; }
  try {
    // Merge all params: top filter bar + search + status + sort + token
    const params = new URLSearchParams(activeFilters);
    if (authToken) params.set('token', authToken);
    const search = document.getElementById('sku-search-input').value || '';
    const status = document.getElementById('sku-status-filter').value || '';
    if (search) params.set('search', search);
    if (status) params.set('status', status);
    if (skuState && skuState.sortBy) params.set('sortBy', skuState.sortBy);
    if (skuState && skuState.sortDir) params.set('sortDir', skuState.sortDir);
    const url = '/api/export-skus-xlsx?' + params.toString();
    // Trigger download via hidden link (handles browser pop-up blockers better than window.open)
    const a = document.createElement('a');
    a.href = url;
    a.download = '';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    setTimeout(() => { if (btn) { btn.innerHTML = orig; btn.disabled = false; } }, 800);
  }
}

async function exportNegativeSKUsExcel() {
  const btn = document.getElementById('negsku-export-btn');
  const orig = btn ? btn.innerHTML : null;
  if (btn) { btn.innerHTML = '⏳ Generating...'; btn.disabled = true; }
  try {
    const url = '/api/export-negativeskus-xlsx' + filterQuery();
    const a = document.createElement('a');
    a.href = url;
    a.download = '';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    setTimeout(() => { if (btn) { btn.innerHTML = orig; btn.disabled = false; } }, 800);
  }
}

async function exportTop300Excel() {
  const btn = document.getElementById('top300-export-btn');
  const orig = btn ? btn.innerHTML : null;
  if (btn) { btn.innerHTML = '⏳ Generating...'; btn.disabled = true; }
  try {
    const url = '/api/export-top300-xlsx' + filterQuery();
    const a = document.createElement('a');
    a.href = url;
    a.download = '';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    setTimeout(() => { if (btn) { btn.innerHTML = orig; btn.disabled = false; } }, 800);
  }
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
