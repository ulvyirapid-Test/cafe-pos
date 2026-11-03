/**
 * ☕ Cafee Ankor POS — Cloud Sync for Google Sheets (v1)
 *
 * The POS sends sales, voids, shifts, stock changes, menu, staff and settings here.
 * Every POS device that uses the same Web App URL + API Key shares the same data.
 *
 * SETUP (once):
 *   1. Paste this file into Extensions → Apps Script of a NEW Google Sheet.
 *   2. Select the function "setup" and press ▶ Run (allow the permissions).
 *   3. Open "Execution log" — copy the API KEY that is printed.
 *   4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → Deploy.
 *   5. Copy the Web app URL (ends with /exec) and paste URL + API Key into the POS
 *      (Settings → Cloud Sync), tick "បើក Sync" and press "រក្សាទុក & សាកល្បង".
 *
 * Forgot the key? Run "showApiKey". Want a new key? Run "newApiKey" (then update every POS).
 */

const TZ = 'Asia/Phnom_Penh';          // time zone for the Date/Time columns
const CHUNK = 40000;                    // Google Sheets cell limit is 50,000 characters

const SHEETS = {
  Sales: ['saleId', 'invId', 'device', 'date', 'time', 'table', 'orderType', 'cashier', 'items', 'subtotal', 'discount', 'vat',
    'redeem', 'total', 'paymentMethod', 'payments', 'paidUSD', 'changeUSD', 'ref', 'customer', 'phone', 'pointsEarned',
    'cogs', 'profit', 'voided', 'voidReason', 'voidBy', 'voidAt', 'fromHistory', 'receivedAt'],
  SaleItems: ['saleId', 'invId', 'date', 'time', 'item', 'size', 'sugar', 'ice', 'toppings', 'note', 'qty', 'unitPrice', 'lineTotal', 'device', 'voided'],
  Voids: ['date', 'time', 'type', 'invId', 'table', 'item', 'qty', 'amount', 'by', 'approvedBy', 'reason', 'device'],
  Shifts: ['start', 'end', 'closedBy', 'approvedBy', 'bills', 'net', 'cash', 'khqr', 'card', 'discount', 'vat', 'redeem',
    'itemVoids', 'itemVoidsAmt', 'billVoids', 'billVoidsAmt', 'float', 'expenses', 'expected', 'counted', 'variance', 'device'],
  Inventory: ['id', 'name', 'unit', 'stock', 'used', 'initialStock', 'alertThreshold', 'updated'],
  StockLog: ['date', 'time', 'type', 'rawId', 'name', 'qty', 'stockAfter', 'by', 'ref', 'device'],
  Config: ['key', 'updated', 'device', 'value'],
  Processed: ['eventId', 'type', 'device', 'receivedAt']
};

// ============================ WEB APP ============================
function doGet() {
  return ContentService.createTextOutput('Cafee Ankor POS Sync — OK. (POS uses POST)');
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_json' }); }
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key) return out({ ok: false, error: 'not_setup' });
  if (!body || body.key !== key) return out({ ok: false, error: 'bad_key' });
  const device = String(body.device || '?').slice(0, 3);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return out({ ok: false, error: 'busy' });
  try {
    ensureSheets();
    const processed = (body.action === 'sync' && Array.isArray(body.events)) ? processEvents(body.events, device) : [];
    const state = serverState(Number(body.configSince) || 0);
    return out(Object.assign({ ok: true, processed: processed }, state));
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err).slice(0, 200) });
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================ EVENTS ============================
function processEvents(events, device) {
  const ss = SpreadsheetApp.getActive();
  const done = new Set(columnValues(ss.getSheetByName('Processed'), 1).map(String));
  const inv = loadInventory();
  const add = {};                 // sheet name -> rows to append
  const push = (name, row) => (add[name] = add[name] || []).push(row);
  const saleVoids = [];           // bill voids to mark on the Sales sheet
  const configWrites = [];
  const processed = [];
  const now = new Date();

  events.forEach(ev => {
    if (!ev || !ev.id) return;
    const id = String(ev.id);
    processed.push(id);
    if (done.has(id)) return;   // already handled (the POS re-sent it after a timeout)
    done.add(id);
    const p = ev.payload || {};
    switch (ev.type) {
      case 'sale': {
        const dt = splitDate(p.timestamp);
        const items = (p.items || []).map(i => `${i.qty}× ${i.name}${i.size && i.size !== '-' ? ' (' + i.size + ')' : ''}`).join(', ');
        push('Sales', [p.id, p.invId, device, dt[0], dt[1], p.table, p.orderType, p.cashierName, items, num(p.subtotal), num(p.discount),
          num(p.vatUSD), num(p.redeemUSD), num(p.totalUSD), p.paymentMethod,
          (p.payments || []).map(x => `${x.method}:${num(x.amountUSD)}`).join(' + '), num(p.paidUSD), num(p.changeUSD), p.ref || '',
          p.customerName || '', p.customerPhone ? "'" + p.customerPhone : '', num(p.pointsEarned), num(p.cogs),
          round(num(p.totalUSD) - num(p.vatUSD) - num(p.cogs)), p.voided ? true : false, p.voidReason || '', p.voidBy || '',
          p.voidAt || '', p.uploadedHistory ? true : false, now]);
        (p.items || []).forEach(i => push('SaleItems', [p.id, p.invId, dt[0], dt[1], i.name, i.size || '-', i.sugar || '-', i.ice || '-',
          (i.toppings || []).map(t => t.name || t).join(', '), i.note || '', num(i.qty), num(i.unitPrice), round(num(i.qty) * num(i.unitPrice)), device, p.voided ? true : false]));
        if (!p.skipStock) (p.consumption || []).forEach(c => changeStock(inv, c.rawId, -num(c.qty), true));
        break;
      }
      case 'voidBill': {
        const dt = splitDate(p.ts);
        push('Voids', [dt[0], dt[1], 'bill', p.invId, p.table, '', '', num(p.amount), p.by, p.approvedBy, p.reason, device]);
        saleVoids.push(p);
        if (!p.skipStock) (p.consumption || []).forEach(c => changeStock(inv, c.rawId, num(c.qty), true));
        break;
      }
      case 'voidItem': {
        const dt = splitDate(p.ts);
        push('Voids', [dt[0], dt[1], 'item', '', p.table, p.itemName, num(p.qty), num(p.amount), p.by, p.approvedBy, p.reason, device]);
        break;
      }
      case 'shift': {
        const s = p;
        push('Shifts', [s.start, s.end, s.closedBy, s.approvedBy, num(s.bills), num(s.net), num(s.cash), num(s.khqr), num(s.card),
          num(s.discount), num(s.vat), num(s.redeem), num(s.itemVoids), num(s.itemVoidsAmt), num(s.billVoids), num(s.billVoidsAmt),
          num(s.float), num(s.expenses), num(s.expected), num(s.counted), num(s.variance), device]);
        break;
      }
      case 'restock': {
        const r = changeStock(inv, p.rawId, num(p.qty), false);
        const dt = splitDate(p.ts);
        push('StockLog', [dt[0], dt[1], 'restock', p.rawId, r ? r.name : '', num(p.qty), r ? r.stock : '', p.by || '', p.ref || '', device]);
        break;
      }
      case 'ingredientNew': {
        if (!inv.map[p.id]) {
          inv.map[p.id] = { id: Number(p.id), name: p.name, unit: p.unit, stock: num(p.stock), used: num(p.used),
            initialStock: num(p.initialStock), alertThreshold: num(p.alertThreshold) };
          inv.order.push(String(p.id));
          inv.changed = true;
        }
        const dt = splitDate(new Date().toISOString());
        push('StockLog', [dt[0], dt[1], 'new', p.id, p.name, num(p.stock), num(p.stock), '', '', device]);
        break;
      }
      case 'ingredientDelete': {
        if (inv.map[p.id]) { delete inv.map[p.id]; inv.order = inv.order.filter(k => k !== String(p.id)); inv.changed = true; }
        break;
      }
      case 'inventorySet': {
        inv.map = {}; inv.order = [];
        (p.items || []).forEach(r => {
          inv.map[r.id] = { id: Number(r.id), name: r.name, unit: r.unit, stock: num(r.stock), used: num(r.used),
            initialStock: num(r.initialStock), alertThreshold: num(r.alertThreshold) };
          inv.order.push(String(r.id));
        });
        inv.changed = true;
        const dt = splitDate(new Date().toISOString());
        push('StockLog', [dt[0], dt[1], 'count/set', '', `${(p.items || []).length} items`, '', '', '', '', device]);
        break;
      }
      case 'config': configWrites.push({ key: String(p.key), value: p.value }); break;
      case 'configDelete': configWrites.push({ key: String(p.key), remove: true }); break;
    }
    push('Processed', [id, ev.type, device, now]);
  });

  Object.keys(add).forEach(name => appendRows(ss.getSheetByName(name), add[name]));
  if (saleVoids.length) markVoided(ss, saleVoids);
  if (inv.changed) saveInventory(inv);
  if (configWrites.length) writeConfig(configWrites, device);
  return processed;
}

// ============================ STOCK ============================
function loadInventory() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Inventory');
  const inv = { map: {}, order: [], changed: false };
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 7).getValues().forEach(r => {
    if (r[0] === '' || r[0] === null) return;
    inv.map[r[0]] = { id: Number(r[0]), name: r[1], unit: r[2], stock: num(r[3]), used: num(r[4]), initialStock: num(r[5]), alertThreshold: num(r[6]) };
    inv.order.push(String(r[0]));
  });
  return inv;
}
function changeStock(inv, rawId, qty, isUse) {
  const r = inv.map[rawId];
  if (!r) return null;
  r.stock = round(r.stock + qty);
  if (isUse) r.used = round(Math.max(0, r.used - qty)); // a sale uses (qty < 0 → used goes up); a void gives back
  else r.initialStock = round(r.initialStock + qty);
  inv.changed = true;
  return r;
}
function saveInventory(inv) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Inventory');
  const now = new Date();
  const rows = inv.order.map(k => inv.map[k]).filter(Boolean)
    .map(r => [r.id, r.name, r.unit, r.stock, r.used, r.initialStock, r.alertThreshold, now]);
  const old = sh.getLastRow() - 1;
  if (old > 0) sh.getRange(2, 1, old, 8).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 8).setValues(rows);
}

// ============================ CONFIG (menu / staff / settings / KHQR) ============================
function writeConfig(writes, device) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const keys = columnValues(sh, 1).map(String);
  const now = Date.now();
  const toDelete = [];
  writes.forEach(w => {
    let row = keys.indexOf(w.key);
    if (w.remove) { if (row >= 0) toDelete.push(row + 2); return; }
    const json = JSON.stringify(w.value === undefined ? null : w.value);
    const chunks = [];
    for (let i = 0; i < json.length; i += CHUNK) chunks.push(json.slice(i, i + CHUNK));
    if (!chunks.length) chunks.push('null');
    if (row < 0) { keys.push(w.key); row = keys.length - 1; }
    const r = row + 2;
    const width = Math.max(sh.getLastColumn(), 3 + chunks.length);
    const vals = [w.key, now, device].concat(chunks);
    while (vals.length < width) vals.push('');
    sh.getRange(r, 1, 1, width).setNumberFormat('@').setValues([vals]);
    sh.getRange(r, 2).setNumberFormat('0');
  });
  toDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}
function readConfig() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, Math.max(4, sh.getLastColumn())).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      let value = null;
      try { value = JSON.parse(r.slice(3).join('') || 'null'); } catch (err) { }
      return { key: String(r[0]), updated: Number(r[1]) || 0, value: value };
    });
}

// ============================ STATE SENT BACK TO THE POS ============================
function serverState(since) {
  const inv = loadInventory();
  const config = readConfig();
  return {
    inventory: inv.order.map(k => inv.map[k]),
    configKeys: config.map(c => c.key),
    config: config.filter(c => c.updated > since)
  };
}

// ============================ HELPERS ============================
function ensureSheets() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold').setBackground('#78350f').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      if (name === 'Processed' || name === 'Config') sh.hideSheet();
    }
  });
}
function appendRows(sh, rows) {
  if (!rows.length) return;
  const width = Math.max.apply(null, rows.map(r => r.length));
  const data = rows.map(r => { const c = r.map(v => v === undefined || v === null ? '' : v); while (c.length < width) c.push(''); return c; });
  sh.getRange(sh.getLastRow() + 1, 1, data.length, width).setValues(data);
}
function markVoided(ss, voids) {
  const sales = ss.getSheetByName('Sales'), items = ss.getSheetByName('SaleItems');
  const ids = sales.getLastRow() > 1 ? sales.getRange(2, 1, sales.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const itemIds = items.getLastRow() > 1 ? items.getRange(2, 1, items.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const col = SHEETS.Sales.indexOf('voided') + 1, itemCol = SHEETS.SaleItems.indexOf('voided') + 1;
  voids.forEach(v => {
    const id = String(v.saleId);
    const i = ids.lastIndexOf(id);
    if (i >= 0) sales.getRange(i + 2, col, 1, 4).setValues([[true, v.reason || '', v.approvedBy || '', v.ts || '']]);
    itemIds.forEach((x, j) => { if (x === id) items.getRange(j + 2, itemCol).setValue(true); });
  });
}
function columnValues(sh, col) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, col, n, 1).getValues().map(r => r[0]).filter(v => v !== '') : [];
}
function splitDate(iso) {
  const d = iso ? new Date(iso) : new Date();
  if (isNaN(d)) return ['', ''];
  return [Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), Utilities.formatDate(d, TZ, 'HH:mm:ss')];
}
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function round(v) { return Math.round(v * 1000) / 1000; }

// ============================ SETUP (run from the editor) ============================
function setup() {
  ensureSheets();
  const ss = SpreadsheetApp.getActive();
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('សន្លឹក1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  buildSummary();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_KEY')) props.setProperty('API_KEY', makeKey());
  Logger.log('✅ Setup done.  API KEY = ' + props.getProperty('API_KEY'));
}
function showApiKey() { Logger.log('API KEY = ' + PropertiesService.getScriptProperties().getProperty('API_KEY')); }
function newApiKey() {
  const k = makeKey();
  PropertiesService.getScriptProperties().setProperty('API_KEY', k);
  Logger.log('🔑 NEW API KEY = ' + k + '   (update it on every POS device)');
}
function makeKey() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40); }

// Daily summary sheet built from the Sales sheet (voided bills are excluded)
function buildSummary() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Summary');
  if (!sh) sh = ss.insertSheet('Summary', 0);
  const c = name => columnLetter(SHEETS.Sales.indexOf(name) + 1);
  sh.clear();
  sh.getRange('A1').setValue('📊 សរុបប្រចាំថ្ងៃ (មិនរាប់ Void)').setFontWeight('bold').setFontSize(13);
  sh.getRange('A2').setFormula(`=IFERROR(QUERY(Sales!A:${c('receivedAt')},"select ${c('date')}, count(${c('saleId')}), sum(${c('total')}), sum(${c('discount')}), sum(${c('vat')}), sum(${c('cogs')}), sum(${c('profit')}) where ${c('voided')} = false and ${c('saleId')} is not null group by ${c('date')} order by ${c('date')} desc label ${c('date')} 'ថ្ងៃ', count(${c('saleId')}) 'វិក្កយបត្រ', sum(${c('total')}) 'ចំណូល $', sum(${c('discount')}) 'បញ្ចុះតម្លៃ $', sum(${c('vat')}) 'VAT $', sum(${c('cogs')}) 'ថ្លៃដើម $', sum(${c('profit')}) 'ចំណេញ $'",1),"មិនទាន់មានការលក់")`);
  sh.getRange('J1').setValue('🏆 មុខម្ហូបលក់ដាច់').setFontWeight('bold').setFontSize(13);
  sh.getRange('J2').setFormula(`=IFERROR(QUERY(SaleItems!A:O,"select E, sum(K), sum(M) where E is not null and O = false group by E order by sum(K) desc limit 20 label E 'មុខម្ហូប', sum(K) 'ចំនួន', sum(M) 'ចំណូល $'",1),"")`);
  sh.setFrozenRows(2);
}
function columnLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
