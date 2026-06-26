const sheetsLib = require('../../lib/sheets');

const HANDLERS = {
  setSpreadsheetId:  (spreadsheetId, arg) => sheetsLib.setSpreadsheetId(arg),
  getSuppliers:      (spreadsheetId)       => sheetsLib.getSuppliers(spreadsheetId),
  addSupplier:       (spreadsheetId, arg)  => sheetsLib.addSupplier(spreadsheetId, arg),
  updateSupplier:    (spreadsheetId, arg)  => sheetsLib.updateSupplier(spreadsheetId, arg),
  deleteSupplier:    (spreadsheetId, arg)  => sheetsLib.deleteSupplier(spreadsheetId, arg),
  getOrders:         (spreadsheetId, arg)  => sheetsLib.getOrders(spreadsheetId, arg),
  addOrder:          (spreadsheetId, arg)  => sheetsLib.addOrder(spreadsheetId, arg),
  updateOrder:       (spreadsheetId, arg)  => sheetsLib.updateOrder(spreadsheetId, arg),
  deleteOrder:       (spreadsheetId, arg)  => sheetsLib.deleteOrder(spreadsheetId, arg),
  addOrdersBulk:     (spreadsheetId, arg)  => sheetsLib.addOrdersBulk(spreadsheetId, arg),
  deleteOrdersBulk:  (spreadsheetId, arg)  => sheetsLib.deleteOrdersBulk(spreadsheetId, arg),
  uploadPhoto:       (spreadsheetId, arg)  => sheetsLib.uploadPhoto(spreadsheetId, arg),
  searchBySku:       (spreadsheetId, arg)  => sheetsLib.searchBySku(spreadsheetId, arg),
  getAllSkus:         (spreadsheetId)       => sheetsLib.getAllSkus(spreadsheetId),
  getExchangeRate:   ()                    => sheetsLib.getExchangeRate()
};

const NO_SPREADSHEET_NEEDED = new Set(['setSpreadsheetId', 'getExchangeRate']);

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ ok: false, error: 'Method not allowed' }) };
  }

  let body;
  try {
    body = JSON.parse(event.body || '{}');
  } catch (e) {
    return { statusCode: 400, body: JSON.stringify({ ok: false, error: 'Invalid JSON' }) };
  }

  const { fn, arg, spreadsheetId } = body;
  const handler = HANDLERS[fn];
  if (!handler) {
    return { statusCode: 400, body: JSON.stringify({ ok: false, error: 'Unknown function: ' + fn }) };
  }
  if (!NO_SPREADSHEET_NEEDED.has(fn) && !spreadsheetId) {
    return { statusCode: 200, body: JSON.stringify({ ok: false, error: 'No spreadsheet connected.' }) };
  }

  try {
    const data = await handler(spreadsheetId, arg);
    return { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ok: true, data }) };
  } catch (e) {
    return { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ok: false, error: e.message || 'Server error' }) };
  }
};
