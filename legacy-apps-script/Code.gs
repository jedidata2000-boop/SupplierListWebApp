// ============================================================
//  Supplier List Web App — Google Apps Script Backend
//  Spreadsheet ID is stored per-user via PropertiesService.
//  No hardcoding needed — set it from the web app UI.
// ============================================================
const SUPPLIERS_SHEET = '_suppliers';

function getStoredSpreadsheetId_() {
  const id = PropertiesService.getUserProperties().getProperty('spreadsheet_id');
  if (!id) throw new Error('NO_SPREADSHEET_CONFIGURED');
  return id;
}

// Called by the client to check whether a spreadsheet is already configured.
function getSpreadsheetId() {
  return PropertiesService.getUserProperties().getProperty('spreadsheet_id') || '';
}

// Accepts a full URL or a raw ID. Validates access before saving.
function setSpreadsheetId(idOrUrl) {
  let id = String(idOrUrl).trim();
  const match = id.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) id = match[1];
  if (!id) return { ok: false, error: 'Please enter a valid spreadsheet URL or ID.' };
  try {
    const title = SpreadsheetApp.openById(id).getName();
    PropertiesService.getUserProperties().setProperty('spreadsheet_id', id);
    return { ok: true, title };
  } catch (e) {
    return { ok: false, error: 'Could not open that spreadsheet. Check the URL/ID and make sure you have access.' };
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Supplier List')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── SUPPLIERS ──────────────────────────────────────────────────────────────

function getSuppliers() {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const sheet = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .map((row, i) => ({
      rowIndex: i + 2,
      name:    String(row[0] || ''),
      address: String(row[1] || ''),
      phone:   String(row[2] || ''),
      email:   String(row[3] || ''),
      wechat:  String(row[4] || ''),
      notes:   String(row[5] || '')
    }))
    .filter(s => s.name.trim());
}

function addSupplier(data) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  let sheet = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SUPPLIERS_SHEET, 0);
    sheet.hideSheet();
    sheet.appendRow(['Name', 'Address', 'Phone', 'Email', 'WeChat', 'Notes']);
    formatHeader_(sheet, 6);
  }
  sheet.appendRow([data.name, data.address, data.phone, data.email, data.wechat, data.notes]);
  ensureSupplierSheet_(ss, data.name);
}

function updateSupplier(data) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const sheet = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sheet) return;
  sheet.getRange(data.rowIndex, 1, 1, 6).setValues([[
    data.name, data.address, data.phone, data.email, data.wechat, data.notes
  ]]);
}

function deleteSupplier(data) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const sheet = ss.getSheetByName(SUPPLIERS_SHEET);
  if (sheet && data.rowIndex >= 2) sheet.deleteRow(data.rowIndex);
  const orderSheet = ss.getSheetByName(data.name);
  if (orderSheet) ss.deleteSheet(orderSheet);
}

// ── ORDERS ─────────────────────────────────────────────────────────────────

function getOrders(supplierName) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  ensureSupplierSheet_(ss, supplierName);
  const sheet = ss.getSheetByName(supplierName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .map((row, i) => ({
      rowIndex:       i + 2,
      sku:            String(row[0] || ''),
      itemName:       String(row[1] || ''),
      price:          parseFloat(row[2]) || 0,
      qtyPerCarton:   parseInt(row[3])   || 0,
      cbm:            parseFloat(row[4]) || 0,
      orderedCartons: parseInt(row[5])   || 0,
      orderedDate:    formatDateCell_(row[6]),
      itemCost:       parseInt(row[8])   || 0,
      sellingPrice:   parseInt(row[9])   || 0
    }))
    .filter(o => o.sku.trim())
    .sort((a, b) => b.orderedDate.localeCompare(a.orderedDate));
}

function addOrder(data) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  ensureSupplierSheet_(ss, data.supplierName);
  const sheet = ss.getSheetByName(data.supplierName);
  sheet.appendRow([
    data.sku, data.itemName, data.price, data.qtyPerCarton, data.cbm,
    data.orderedCartons, data.orderedDate, '',
    data.itemCost, data.sellingPrice
  ]);
}

function updateOrder(data) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const sheet = ss.getSheetByName(data.supplierName);
  if (!sheet) return;
  const photoCell = sheet.getRange(data.rowIndex, 8).getValue();
  sheet.getRange(data.rowIndex, 1, 1, 10).setValues([[
    data.sku, data.itemName, data.price, data.qtyPerCarton, data.cbm,
    data.orderedCartons, data.orderedDate, photoCell,
    data.itemCost, data.sellingPrice
  ]]);
}

function deleteOrder(data) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const sheet = ss.getSheetByName(data.supplierName);
  if (sheet && data.rowIndex >= 2) sheet.deleteRow(data.rowIndex);
}

function addOrdersBulk(orders) {
  if (!orders || !orders.length) return 0;
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const supplierName = orders[0].supplierName;
  ensureSupplierSheet_(ss, supplierName);
  const sheet = ss.getSheetByName(supplierName);
  const rows = orders.map(o => [
    o.sku, o.itemName, o.price, o.qtyPerCarton, o.cbm,
    o.orderedCartons, o.orderedDate, '',
    o.itemCost, o.sellingPrice
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  return rows.length;
}

// ── SKU SEARCH ─────────────────────────────────────────────────────────────

function searchBySku(sku) {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const supplierMap = {};
  getSuppliers().forEach(s => { supplierMap[s.name] = s; });

  const results = [];
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith('_')) return;
    const data = sheet.getDataRange().getValues();
    data.slice(1).forEach((row, i) => {
      if (String(row[0] || '').trim().toUpperCase() !== sku.trim().toUpperCase()) return;
      const sup = supplierMap[sheet.getName()] || {};
      results.push({
        rowIndex:       i + 2,
        supplierName:   sheet.getName(),
        supplierPhone:  sup.phone  || '',
        supplierEmail:  sup.email  || '',
        supplierWechat: sup.wechat || '',
        sku:            String(row[0] || ''),
        itemName:       String(row[1] || ''),
        price:          parseFloat(row[2]) || 0,
        qtyPerCarton:   parseInt(row[3])   || 0,
        cbm:            parseFloat(row[4]) || 0,
        orderedCartons: parseInt(row[5])   || 0,
        orderedDate:    formatDateCell_(row[6]),
        itemCost:       parseInt(row[8])   || 0,
        sellingPrice:   parseInt(row[9])   || 0
      });
    });
  });
  return results.sort((a, b) => b.orderedDate.localeCompare(a.orderedDate));
}

function getAllSkus() {
  const ss = SpreadsheetApp.openById(getStoredSpreadsheetId_());
  const skus = new Set();
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith('_')) return;
    sheet.getDataRange().getValues().slice(1).forEach(row => {
      const s = String(row[0] || '').trim().toUpperCase();
      if (s) skus.add(s);
    });
  });
  return Array.from(skus).sort();
}

// ── EXCHANGE RATE ───────────────────────────────────────────────────────────

function getExchangeRate() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('cny_idr_rate');
    if (cached) return parseFloat(cached);
    const res = UrlFetchApp.fetch('https://api.exchangerate-api.com/v4/latest/CNY', { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const rate = (json.rates && json.rates.IDR) ? parseFloat(json.rates.IDR) : 0;
    if (rate > 0) cache.put('cny_idr_rate', String(rate), 3600);
    return rate;
  } catch (e) {
    return 0;
  }
}

// ── PRIVATE HELPERS ─────────────────────────────────────────────────────────

function ensureSupplierSheet_(ss, name) {
  if (ss.getSheetByName(name)) return;
  const sheet = ss.insertSheet(name);
  sheet.appendRow(['SKU', 'Item Name', 'Price (¥)', 'Qty/Carton', 'CBM',
    'Ordered Cartons', 'Ordered Date', 'Photo', 'Item Cost (Rp)', 'Selling Price (Rp)']);
  formatHeader_(sheet, 10);
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.getRange(2, 7, maxRows - 1, 1).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(2, 9, maxRows - 1, 2).setNumberFormat('"Rp "#,##0');
    sheet.hideColumns(9, 2);
  }
}

function formatHeader_(sheet, colCount) {
  const r = sheet.getRange(1, 1, 1, colCount);
  r.setBackground('#00695C');
  r.setFontColor('#FFFFFF');
  r.setFontWeight('bold');
}

function formatDateCell_(value) {
  if (!value && value !== 0) return '';
  if (value instanceof Date) {
    const y  = value.getFullYear();
    const m  = String(value.getMonth() + 1).padStart(2, '0');
    const d  = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const n = Number(s);
  if (!isNaN(n) && n > 1) {
    const dt = new Date((n - 25569) * 86400000);
    const y  = dt.getUTCFullYear();
    const mo = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const d  = String(dt.getUTCDate()).padStart(2, '0');
    return `${y}-${mo}-${d}`;
  }
  return s;
}
