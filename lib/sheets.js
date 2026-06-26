const { google } = require('googleapis');

const SUPPLIERS_SHEET = '_suppliers';
let rateCache = { value: 0, ts: 0 };

// Column layout for supplier order sheets:
// A(0):Photo  B(1):SKU  C(2):ItemName  D(3):Price  E(4):Qty  F(5):CBM
// G(6):Cartons  H(7):Date  I(8):ItemCost  J(9):SellingPrice

function getAuth() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const key = (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || '').replace(/\\n/g, '\n');
  if (!email || !key) throw new Error('Service account credentials are not configured.');
  return new google.auth.JWT(email, null, key, [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
  ]);
}

function sheetsClient() { return google.sheets({ version: 'v4', auth: getAuth() }); }
function driveClient()  { return google.drive({ version: 'v3', auth: getAuth() }); }

async function getSheetMeta(spreadsheetId) {
  const sheets = sheetsClient();
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  return res.data.sheets.map(s => s.properties);
}

function findSheetProps(metaList, name) {
  return metaList.find(p => p.title === name) || null;
}

function headerFormatRequest(sheetId, colCount) {
  return {
    repeatCell: {
      range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: colCount },
      cell: {
        userEnteredFormat: {
          backgroundColor: { red: 0, green: 0x69 / 255, blue: 0x5c / 255 },
          textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true }
        }
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat)'
    }
  };
}

function extractImageUrl(formulaOrValue) {
  if (!formulaOrValue) return '';
  const m = String(formulaOrValue).match(/=IMAGE\("([^"]+)"\)/i);
  return m ? m[1] : '';
}

async function ensureSuppliersSheet_(spreadsheetId) {
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  if (findSheetProps(meta, SUPPLIERS_SHEET)) return;

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: SUPPLIERS_SHEET, index: 0, hidden: true } } }] }
  });
  const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;

  await sheets.spreadsheets.values.update({
    spreadsheetId, range: `${SUPPLIERS_SHEET}!A1:F1`, valueInputOption: 'RAW',
    requestBody: { values: [['Name', 'Address', 'Phone', 'Email', 'WeChat', 'Notes']] }
  });
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId, requestBody: { requests: [headerFormatRequest(sheetId, 6)] }
  });
}

async function ensureSupplierSheet_(spreadsheetId, name) {
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  if (findSheetProps(meta, name)) return;

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: name } } }] }
  });
  const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;

  await sheets.spreadsheets.values.update({
    spreadsheetId, range: `${name}!A1:J1`, valueInputOption: 'RAW',
    requestBody: {
      values: [['Photo', 'SKU', 'Item Name', 'Price (¥)', 'Qty/Carton', 'CBM',
        'Ordered Cartons', 'Ordered Date', 'Item Cost (Rp)', 'Selling Price (Rp)']]
    }
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        headerFormatRequest(sheetId, 10),
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 1, startColumnIndex: 7, endColumnIndex: 8 },
            cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-MM-dd' } } },
            fields: 'userEnteredFormat.numberFormat'
          }
        },
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 1, startColumnIndex: 8, endColumnIndex: 10 },
            cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"Rp "#,##0' } } },
            fields: 'userEnteredFormat.numberFormat'
          }
        },
        {
          updateDimensionProperties: {
            range: { sheetId, dimension: 'COLUMNS', startIndex: 8, endIndex: 10 },
            properties: { hiddenByUser: true },
            fields: 'hiddenByUser'
          }
        },
        {
          updateDimensionProperties: {
            range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
            properties: { pixelSize: 130 },
            fields: 'pixelSize'
          }
        }
      ]
    }
  });
}

function formatDateCell_(value) {
  if (!value && value !== 0) return '';
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const n = Number(s);
  if (!isNaN(n) && n > 1) {
    const dt = new Date((n - 25569) * 86400000);
    const y = dt.getUTCFullYear();
    const mo = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const d = String(dt.getUTCDate()).padStart(2, '0');
    return `${y}-${mo}-${d}`;
  }
  return s;
}

// ── SETUP ────────────────────────────────────────────────────────────────────

async function setSpreadsheetId(idOrUrl) {
  let id = String(idOrUrl).trim();
  const match = id.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) id = match[1];
  if (!id) return { ok: false, error: 'Please enter a valid spreadsheet URL or ID.' };
  try {
    const sheets = sheetsClient();
    const res = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'properties.title' });
    return { ok: true, id, title: res.data.properties.title };
  } catch (e) {
    const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || 'the service account';
    return { ok: false, error: `Could not open that spreadsheet. Check the URL/ID and make sure it's shared with ${email} (Editor access).` };
  }
}

// ── SUPPLIERS ────────────────────────────────────────────────────────────────

async function getSuppliers(spreadsheetId) {
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  if (!findSheetProps(meta, SUPPLIERS_SHEET)) return [];
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${SUPPLIERS_SHEET}!A2:F` });
  const rows = res.data.values || [];
  return rows
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

async function addSupplier(spreadsheetId, data) {
  await ensureSuppliersSheet_(spreadsheetId);
  const sheets = sheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${SUPPLIERS_SHEET}!A:F`,
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [[data.name, data.address, data.phone, data.email, data.wechat, data.notes]] }
  });
  await ensureSupplierSheet_(spreadsheetId, data.name);
}

async function updateSupplier(spreadsheetId, data) {
  const sheets = sheetsClient();
  await sheets.spreadsheets.values.update({
    spreadsheetId, range: `${SUPPLIERS_SHEET}!A${data.rowIndex}:F${data.rowIndex}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[data.name, data.address, data.phone, data.email, data.wechat, data.notes]] }
  });
}

async function deleteSupplier(spreadsheetId, data) {
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  const requests = [];
  const supProps = findSheetProps(meta, SUPPLIERS_SHEET);
  if (supProps && data.rowIndex >= 2) {
    requests.push({
      deleteDimension: {
        range: { sheetId: supProps.sheetId, dimension: 'ROWS', startIndex: data.rowIndex - 1, endIndex: data.rowIndex }
      }
    });
  }
  const orderProps = findSheetProps(meta, data.name);
  if (orderProps) requests.push({ deleteSheet: { sheetId: orderProps.sheetId } });
  if (requests.length) await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
}

// ── ORDERS ───────────────────────────────────────────────────────────────────

async function getOrders(spreadsheetId, supplierName) {
  await ensureSupplierSheet_(spreadsheetId, supplierName);
  const sheets = sheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId, range: `${supplierName}!A2:J`,
    valueRenderOption: 'FORMULA'
  });
  const rows = res.data.values || [];
  return rows
    .map((row, i) => ({
      rowIndex:       i + 2,
      supplierName:   supplierName,
      photo:          extractImageUrl(row[0]),
      sku:            String(row[1] || ''),
      itemName:       String(row[2] || ''),
      price:          parseFloat(row[3]) || 0,
      qtyPerCarton:   parseInt(row[4]) || 0,
      cbm:            parseFloat(row[5]) || 0,
      orderedCartons: parseInt(row[6]) || 0,
      orderedDate:    formatDateCell_(row[7]),
      itemCost:       parseInt(row[8]) || 0,
      sellingPrice:   parseInt(row[9]) || 0
    }))
    .filter(o => o.sku.trim())
    .sort((a, b) => b.orderedDate.localeCompare(a.orderedDate));
}

async function addOrder(spreadsheetId, data) {
  await ensureSupplierSheet_(spreadsheetId, data.supplierName);
  const sheets = sheetsClient();
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${data.supplierName}!A:J`,
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[
        '', data.sku, data.itemName, data.price, data.qtyPerCarton, data.cbm,
        data.orderedCartons, data.orderedDate, data.itemCost, data.sellingPrice
      ]]
    }
  });
  const updatedRange = res.data.updates && res.data.updates.updatedRange;
  const m = updatedRange && updatedRange.match(/!.*?(\d+):/);
  return { rowIndex: m ? parseInt(m[1]) : -1 };
}

async function updateOrder(spreadsheetId, data) {
  const sheets = sheetsClient();
  const photoRes = await sheets.spreadsheets.values.get({
    spreadsheetId, range: `${data.supplierName}!A${data.rowIndex}`,
    valueRenderOption: 'FORMULA'
  });
  const photoCell = (photoRes.data.values && photoRes.data.values[0] && photoRes.data.values[0][0]) || '';
  await sheets.spreadsheets.values.update({
    spreadsheetId, range: `${data.supplierName}!A${data.rowIndex}:J${data.rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        photoCell, data.sku, data.itemName, data.price, data.qtyPerCarton, data.cbm,
        data.orderedCartons, data.orderedDate, data.itemCost, data.sellingPrice
      ]]
    }
  });
}

async function deleteOrder(spreadsheetId, data) {
  if (data.rowIndex < 2) return;
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  const props = findSheetProps(meta, data.supplierName);
  if (!props) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId: props.sheetId, dimension: 'ROWS', startIndex: data.rowIndex - 1, endIndex: data.rowIndex }
        }
      }]
    }
  });
}

async function deleteOrdersBulk(spreadsheetId, data) {
  const { supplierName, rows } = data;
  if (!rows || !rows.length) return;
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  const props = findSheetProps(meta, supplierName);
  if (!props) return;
  const sorted = [...rows].filter(r => r.rowIndex >= 2).sort((a, b) => b.rowIndex - a.rowIndex);
  if (!sorted.length) return;
  const requests = sorted.map(r => ({
    deleteDimension: {
      range: { sheetId: props.sheetId, dimension: 'ROWS', startIndex: r.rowIndex - 1, endIndex: r.rowIndex }
    }
  }));
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
}

async function addOrdersBulk(spreadsheetId, orders) {
  if (!orders || !orders.length) return { count: 0, startRow: -1 };
  const supplierName = orders[0].supplierName;
  await ensureSupplierSheet_(spreadsheetId, supplierName);
  const sheets = sheetsClient();
  const rows = orders.map(o => [
    '', o.sku, o.itemName, o.price, o.qtyPerCarton, o.cbm,
    o.orderedCartons, o.orderedDate, o.itemCost, o.sellingPrice
  ]);
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${supplierName}!A:J`,
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rows }
  });
  const updatedRange = res.data.updates && res.data.updates.updatedRange;
  const m = updatedRange && updatedRange.match(/!.*?(\d+):/);
  return { count: rows.length, startRow: m ? parseInt(m[1]) : -1 };
}

// ── PHOTO UPLOAD ─────────────────────────────────────────────────────────────

async function uploadPhoto(spreadsheetId, data) {
  const { supplierName, rowIndex, imageBase64, mimeType, siteUrl: clientSiteUrl } = data;
  const drive = driveClient();
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

  if (!folderId) {
    throw new Error('GOOGLE_DRIVE_FOLDER_ID is not set. Create a Google Drive folder, share it with the service account, and add the folder ID as a Netlify env var.');
  }

  const buffer = Buffer.from(imageBase64, 'base64');
  const { PassThrough } = require('stream');
  const stream = new PassThrough();
  stream.end(buffer);

  const uploadRes = await drive.files.create({
    requestBody: {
      name: `${supplierName}_row${rowIndex}`,
      mimeType: mimeType || 'image/jpeg',
      parents: [folderId]
    },
    media: { mimeType: mimeType || 'image/jpeg', body: stream },
    fields: 'id',
    supportsAllDrives: true
  });
  const fileId = uploadRes.data.id;

  await drive.permissions.create({
    fileId,
    requestBody: { role: 'reader', type: 'anyone' },
    supportsAllDrives: true
  });

  const siteUrl = (clientSiteUrl || process.env.URL || '').replace(/\/$/, '');
  const imageUrl = `${siteUrl}/.netlify/functions/photo?id=${fileId}`;

  const sheets = sheetsClient();
  await sheets.spreadsheets.values.update({
    spreadsheetId, range: `${supplierName}!A${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[`=IMAGE("${imageUrl}")`]] }
  });

  const meta = await getSheetMeta(spreadsheetId);
  const props = findSheetProps(meta, supplierName);
  if (props) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          updateDimensionProperties: {
            range: { sheetId: props.sheetId, dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex },
            properties: { pixelSize: 130 },
            fields: 'pixelSize'
          }
        }]
      }
    });
  }

  return { url: imageUrl };
}

// ── SKU SEARCH ───────────────────────────────────────────────────────────────

async function searchBySku(spreadsheetId, sku) {
  const sheets = sheetsClient();
  const suppliers = await getSuppliers(spreadsheetId);
  const supplierMap = {};
  suppliers.forEach(s => { supplierMap[s.name] = s; });

  const meta = await getSheetMeta(spreadsheetId);
  const orderSheets = meta.filter(p => !p.title.startsWith('_'));
  if (!orderSheets.length) return [];

  const ranges = orderSheets.map(p => `${p.title}!A2:J`);
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId, ranges, valueRenderOption: 'FORMULA'
  });
  const target = String(sku || '').trim().toUpperCase();
  const results = [];

  res.data.valueRanges.forEach((vr, idx) => {
    const sheetName = orderSheets[idx].title;
    const rows = vr.values || [];
    rows.forEach((row, i) => {
      if (String(row[1] || '').trim().toUpperCase() !== target) return;
      const sup = supplierMap[sheetName] || {};
      results.push({
        rowIndex:       i + 2,
        supplierName:   sheetName,
        supplierPhone:  sup.phone  || '',
        supplierEmail:  sup.email  || '',
        supplierWechat: sup.wechat || '',
        photo:          extractImageUrl(row[0]),
        sku:            String(row[1] || ''),
        itemName:       String(row[2] || ''),
        price:          parseFloat(row[3]) || 0,
        qtyPerCarton:   parseInt(row[4])   || 0,
        cbm:            parseFloat(row[5]) || 0,
        orderedCartons: parseInt(row[6])   || 0,
        orderedDate:    formatDateCell_(row[7]),
        itemCost:       parseInt(row[8])   || 0,
        sellingPrice:   parseInt(row[9])   || 0
      });
    });
  });

  return results.sort((a, b) => b.orderedDate.localeCompare(a.orderedDate));
}

async function getAllSkus(spreadsheetId) {
  const sheets = sheetsClient();
  const meta = await getSheetMeta(spreadsheetId);
  const orderSheets = meta.filter(p => !p.title.startsWith('_'));
  if (!orderSheets.length) return [];

  const ranges = orderSheets.map(p => `${p.title}!B2:B`);
  const res = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges });
  const skus = new Set();
  res.data.valueRanges.forEach(vr => {
    (vr.values || []).forEach(row => {
      const s = String(row[0] || '').trim().toUpperCase();
      if (s) skus.add(s);
    });
  });
  return Array.from(skus).sort();
}

// ── EXCHANGE RATE ────────────────────────────────────────────────────────────

async function getExchangeRate() {
  const now = Date.now();
  if (rateCache.value > 0 && now - rateCache.ts < 3600000) return rateCache.value;
  try {
    const rate = await new Promise((resolve) => {
      const https = require('https');
      https.get('https://api.exchangerate-api.com/v4/latest/CNY', (res) => {
        let data = '';
        res.on('data', chunk => { data += chunk; });
        res.on('end', () => {
          try {
            const json = JSON.parse(data);
            resolve((json.rates && json.rates.IDR) ? parseFloat(json.rates.IDR) : 0);
          } catch (e) { resolve(0); }
        });
      }).on('error', () => resolve(0));
    });
    if (rate > 0) rateCache = { value: rate, ts: now };
    return rate;
  } catch (e) {
    return 0;
  }
}

module.exports = {
  setSpreadsheetId,
  getSuppliers, addSupplier, updateSupplier, deleteSupplier,
  getOrders, addOrder, updateOrder, deleteOrder, addOrdersBulk, deleteOrdersBulk,
  uploadPhoto,
  searchBySku, getAllSkus,
  getExchangeRate
};
