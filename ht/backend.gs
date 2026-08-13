/**
 * Henderson Transpacific — vehicle service-record system.
 * Google Apps Script backend. Google Sheet is the database.
 *
 * Deploy as Web App: Execute as "me", Access "Anyone".
 * Front end posts Content-Type: text/plain;charset=utf-8 (CORS-simple-request
 * workaround) with a JSON string body -> doPost reads e.postData.contents.
 * doGet is a fallback that reads the same JSON payload from ?p=<json>.
 */

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

var VERSION = 'ht-1';
var SESSION_DAYS = 14;

var HEADERS = {
  shops:    ['shopId', 'name', 'salt', 'passHash', 'active', 'created'],
  sessions: ['token', 'shopId', 'expires'],
  vehicles: ['id', 'shopId', 'plate', 'vin', 'model', 'year', 'owner', 'odo', 'notes', 'updated'],
  records:  ['id', 'shopId', 'vehicleId', 'date', 'odo', 'category', 'work', 'parts', 'cost', 'currency', 'vendor', 'tech', 'notes', 'created'],
  plan:     ['id', 'shopId', 'vehicleId', 'item', 'months', 'km', 'lastDate', 'lastOdo', 'cost', 'note', 'status']
};

var DEFAULT_PLAN_ITEMS = [
  { item: '冷氣性能檢查｜A/C performance check', months: 12, km: '', cost: '¥150–400', note: '' },
  { item: '冷氣乾燥劑包｜A/C desiccant bag', months: 48, km: 120000, cost: '¥600–1,200', note: '' },
  { item: '冷氣壓縮機（預防性）｜A/C compressor (preventive)', months: '', km: 150000, cost: '¥2,500–6,000', note: '' },
  { item: '傳動 ATF 齒輪油｜Drive-unit ATF', months: 41, km: 100000, cost: '¥300–600', note: '' },
  { item: '電池冷卻液（全量）｜Battery coolant (full)', months: 60, km: 150000, cost: '¥800–1,800', note: '' },
  { item: '散熱器／冷凝器清洗｜Radiator & condenser clean', months: 24, km: '', cost: '¥200–500', note: '' },
  { item: '制動皮｜Brake pads', months: 12, km: 60000, cost: '¥800–2,000', note: '' },
  { item: '制動碟｜Brake rotors', months: 48, km: 120000, cost: '¥1,200–3,000', note: '' },
  { item: '制動液｜Brake fluid', months: 48, km: '', cost: '¥200–400', note: '' },
  { item: '卡鉗清潔潤滑｜Caliper service', months: 12, km: 20000, cost: '¥200–500', note: '' },
  { item: '前懸掛擺臂｜Front control arms', months: '', km: 100000, cost: '¥600–1,500', note: '' },
  { item: '避震機｜Shock absorbers', months: '', km: 150000, cost: '¥800–2,500', note: '' },
  { item: '12V 副電池｜12V battery', months: 42, km: '', cost: '¥800–1,500', note: '' },
  { item: '輪胎換位｜Tyre rotation', months: '', km: 10000, cost: '¥50–150', note: '' },
  { item: '輪胎更換｜Tyre replacement', months: 19, km: 45000, cost: '¥700–1,500', note: '' },
  { item: '四輪定位｜Wheel alignment', months: 12, km: '', cost: '¥200–500', note: '' },
  { item: '車廂濾芯｜Cabin filter', months: 12, km: '', cost: '¥100–300', note: '' },
  { item: '雨刮｜Wipers', months: 12, km: '', cost: '¥60–200', note: '' }
];

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

// Spreadsheet that stores all data. Works both bound and standalone.
var SHEET_ID = '15ZvU6g8u5ke9yOrd1wVLEkeK_x-8bZoLl5tlPvOcokE';

function getSS() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  var ss = getSS();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function initSheets() {
  Object.keys(HEADERS).forEach(function (name) { getSheet(name); });
}

// Reads a sheet into an array of plain objects keyed by header row.
// Each object also carries a non-enumerable-ish __row (1-based sheet row).
function sheetToObjects(sh) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var blank = row.every(function (c) { return c === '' || c === null || c === undefined; });
    if (blank) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    obj.__row = i + 1;
    out.push(obj);
  }
  return out;
}

function appendObject(sh, headers, obj) {
  var row = headers.map(function (h) {
    var v = obj[h];
    return (v === undefined || v === null) ? '' : v;
  });
  sh.appendRow(row);
}

function updateRowByIndex(sh, rowIndex, headers, obj) {
  var row = headers.map(function (h) {
    var v = obj[h];
    return (v === undefined || v === null) ? '' : v;
  });
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function deleteWhere(sheetName, predicate) {
  var sh = getSheet(sheetName);
  var rows = sheetToObjects(sh);
  var toDelete = rows.filter(predicate).map(function (r) { return r.__row; })
    .sort(function (a, b) { return b - a; }); // bottom-up so indices stay valid
  toDelete.forEach(function (rowIdx) { sh.deleteRow(rowIdx); });
  return toDelete.length;
}

function stripRow(obj) {
  var o = {};
  for (var k in obj) { if (k !== '__row') o[k] = obj[k]; }
  return o;
}

// ---------------------------------------------------------------------------
// Crypto / auth helpers
// ---------------------------------------------------------------------------

function sha256Hex(str) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var v = bytes[i];
    if (v < 0) v += 256;
    var h = v.toString(16);
    hex += (h.length === 1 ? '0' + h : h);
  }
  return hex;
}

function makeSalt() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function findShopById(shopId) {
  var rows = sheetToObjects(getSheet('shops'));
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].shopId) === String(shopId)) return rows[i];
  }
  return null;
}

function createSession(shopId) {
  var sh = getSheet('sessions');
  var token = Utilities.getUuid();
  var expires = Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000;
  appendObject(sh, HEADERS.sessions, { token: token, shopId: shopId, expires: expires });
  return { token: token, expires: expires };
}

// Returns { shopId, shopName, __row } or null. Never leaks *why* it failed.
function validateToken(token) {
  if (!token) return null;
  var rows = sheetToObjects(getSheet('sessions'));
  var found = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].token) === String(token)) { found = rows[i]; break; }
  }
  if (!found) return null;
  if (!(Number(found.expires) > Date.now())) return null;
  var shop = findShopById(found.shopId);
  if (!shop) return null;
  if (shop.active === false || String(shop.active).toLowerCase() === 'false') return null;
  return { shopId: found.shopId, shopName: shop.name, __row: found.__row };
}

function findVehicle(id, shopId) {
  var rows = sheetToObjects(getSheet('vehicles'));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id === id && rows[i].shopId === shopId) return rows[i];
  }
  return null;
}

// ---------------------------------------------------------------------------
// Plan due-date computation
// ---------------------------------------------------------------------------

function isoDate(d) {
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return y + '-' + m + '-' + day;
}

// state: 'none' (never done) | 'overdue' | 'due' (<=60 days or <=2000km away) | 'ok'
function computePlanItem(item, vehicleOdo) {
  var lastDate = item.lastDate ? String(item.lastDate) : '';
  var lastOdoNum = (item.lastOdo !== '' && item.lastOdo !== undefined && item.lastOdo !== null && !isNaN(Number(item.lastOdo)))
    ? Number(item.lastOdo) : null;
  var hasLast = !!lastDate || lastOdoNum !== null;
  if (!hasLast) return { nextDate: '', nextOdo: '', state: 'none' };

  var nextDate = '';
  var dateRank = null; // 0 ok, 1 due, 2 overdue
  var months = Number(item.months);
  if (lastDate && months && !isNaN(months) && months > 0) {
    var d = new Date(lastDate);
    if (!isNaN(d.getTime())) {
      d.setMonth(d.getMonth() + months);
      nextDate = isoDate(d);
      var diffDays = (d.getTime() - Date.now()) / 86400000;
      dateRank = diffDays < 0 ? 2 : (diffDays <= 60 ? 1 : 0);
    }
  }

  var nextOdo = '';
  var odoRank = null;
  var km = Number(item.km);
  if (lastOdoNum !== null && km && !isNaN(km) && km > 0) {
    nextOdo = lastOdoNum + km;
    var remaining = nextOdo - (Number(vehicleOdo) || 0);
    odoRank = remaining < 0 ? 2 : (remaining <= 2000 ? 1 : 0);
  }

  var rank = Math.max(dateRank === null ? -1 : dateRank, odoRank === null ? -1 : odoRank);
  var state = (rank === -1) ? 'ok' : (rank === 2 ? 'overdue' : (rank === 1 ? 'due' : 'ok'));
  return { nextDate: nextDate, nextOdo: nextOdo, state: state };
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

function doLogin(payload) {
  var GENERIC = 'invalid shop or password';
  var shopIdRaw = payload.shop, pass = payload.pass;
  if (!shopIdRaw || !pass) return { ok: false, error: 'shop and password required' };
  var shopId = String(shopIdRaw).trim().toLowerCase();

  var rows = sheetToObjects(getSheet('shops'));
  var found = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].shopId).trim().toLowerCase() === shopId) { found = rows[i]; break; }
  }
  if (!found) return { ok: false, error: GENERIC };
  if (found.active === false || String(found.active).toLowerCase() === 'false') return { ok: false, error: GENERIC };

  var hash = sha256Hex(String(found.salt) + String(pass));
  var match = (hash === String(found.passHash));
  if (!match) return { ok: false, error: GENERIC };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var sess;
  try {
    sess = createSession(found.shopId);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, token: sess.token, shopName: found.name, expires: sess.expires };
}

function doLogout(payload) {
  var sh = getSheet('sessions');
  var rows = sheetToObjects(sh);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].token) === String(payload.token)) { sh.deleteRow(rows[i].__row); break; }
  }
  return { ok: true };
}

function doMe(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  return { ok: true, shopId: s.shopId, shopName: s.shopName };
}

function doVehicles(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var rows = sheetToObjects(getSheet('vehicles')).filter(function (r) { return r.shopId === s.shopId; });
  rows.sort(function (a, b) {
    var ua = String(a.updated), ub = String(b.updated);
    return ua < ub ? 1 : (ua > ub ? -1 : 0);
  });
  return { ok: true, vehicles: rows.map(stripRow) };
}

function doVehicleSave(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var v = payload.vehicle || {};
  var plate = String(v.plate || '').trim();
  if (!plate) return { ok: false, error: 'plate required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet('vehicles');
    var now = new Date().toISOString();
    var obj = {
      shopId: s.shopId,
      plate: plate,
      vin: String(v.vin || '').trim(),
      model: String(v.model || '').trim(),
      year: (v.year !== undefined && v.year !== null && v.year !== '') ? Number(v.year) : '',
      owner: String(v.owner || '').trim(),
      odo: (v.odo !== undefined && v.odo !== null && v.odo !== '') ? Number(v.odo) : 0,
      notes: String(v.notes || '').trim(),
      updated: now
    };
    if (v.id) {
      var existing = findVehicle(v.id, s.shopId);
      if (!existing) return { ok: false, error: 'vehicle not found' };
      obj.id = existing.id;
      updateRowByIndex(sh, existing.__row, HEADERS.vehicles, obj);
    } else {
      obj.id = 'v_' + Utilities.getUuid().slice(0, 8);
      appendObject(sh, HEADERS.vehicles, obj);
    }
    return { ok: true, id: obj.id, vehicle: obj };
  } finally {
    lock.releaseLock();
  }
}

function doVehicleDelete(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var id = payload.id;
  if (!id) return { ok: false, error: 'id required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var vsh = getSheet('vehicles');
    var target = findVehicle(id, s.shopId);
    if (!target) return { ok: false, error: 'vehicle not found' };
    vsh.deleteRow(target.__row);
    deleteWhere('records', function (r) { return r.vehicleId === id && r.shopId === s.shopId; });
    deleteWhere('plan', function (r) { return r.vehicleId === id && r.shopId === s.shopId; });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function doRecords(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var vehicleId = payload.vehicleId;
  if (!vehicleId) return { ok: false, error: 'vehicleId required' };
  var rows = sheetToObjects(getSheet('records')).filter(function (r) {
    return r.shopId === s.shopId && r.vehicleId === vehicleId;
  });
  rows.sort(function (a, b) {
    var da = String(a.date), db = String(b.date);
    return da < db ? 1 : (da > db ? -1 : 0);
  });
  return { ok: true, records: rows.map(stripRow) };
}

function doRecordSave(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var vehicleId = payload.vehicleId;
  var r = payload.record || {};
  if (!vehicleId) return { ok: false, error: 'vehicleId required' };
  if (!r.date) return { ok: false, error: 'date required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var vsh = getSheet('vehicles');
    var vehicle = findVehicle(vehicleId, s.shopId);
    if (!vehicle) return { ok: false, error: 'vehicle not found' };

    var rsh = getSheet('records');
    var obj = {
      shopId: s.shopId,
      vehicleId: vehicleId,
      date: String(r.date),
      odo: (r.odo !== undefined && r.odo !== null && r.odo !== '') ? Number(r.odo) : '',
      category: String(r.category || '').trim(),
      work: String(r.work || '').trim(),
      parts: String(r.parts || '').trim(),
      cost: (r.cost !== undefined && r.cost !== null && r.cost !== '') ? Number(r.cost) : 0,
      currency: String(r.currency || '').trim(),
      vendor: String(r.vendor || '').trim(),
      tech: String(r.tech || '').trim(),
      notes: String(r.notes || '').trim()
    };

    var isCreate = !r.id;
    if (!isCreate) {
      var existing = null;
      var rrows = sheetToObjects(rsh);
      for (var j = 0; j < rrows.length; j++) {
        if (rrows[j].id === r.id && rrows[j].shopId === s.shopId) { existing = rrows[j]; break; }
      }
      if (!existing) return { ok: false, error: 'record not found' };
      obj.id = existing.id;
      obj.created = existing.created;
      updateRowByIndex(rsh, existing.__row, HEADERS.records, obj);
    } else {
      obj.id = 'r_' + Utilities.getUuid().slice(0, 8);
      obj.created = new Date().toISOString();
      appendObject(rsh, HEADERS.records, obj);
    }

    if (isCreate && typeof obj.odo === 'number' && !isNaN(obj.odo)) {
      var curOdo = Number(vehicle.odo) || 0;
      if (obj.odo > curOdo) {
        var vobj = {};
        HEADERS.vehicles.forEach(function (h) { vobj[h] = vehicle[h]; });
        vobj.odo = obj.odo;
        vobj.updated = new Date().toISOString();
        updateRowByIndex(vsh, vehicle.__row, HEADERS.vehicles, vobj);
      }
    }

    return { ok: true, id: obj.id, record: obj };
  } finally {
    lock.releaseLock();
  }
}

function doRecordDelete(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var id = payload.id;
  if (!id) return { ok: false, error: 'id required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet('records');
    var rows = sheetToObjects(sh);
    var target = null;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].id === id && rows[i].shopId === s.shopId) { target = rows[i]; break; }
    }
    if (!target) return { ok: false, error: 'record not found' };
    sh.deleteRow(target.__row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function doPlan(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var vehicleId = payload.vehicleId;
  if (!vehicleId) return { ok: false, error: 'vehicleId required' };
  var vehicle = findVehicle(vehicleId, s.shopId);
  if (!vehicle) return { ok: false, error: 'vehicle not found' };

  var rows = sheetToObjects(getSheet('plan')).filter(function (r) {
    return r.shopId === s.shopId && r.vehicleId === vehicleId;
  });
  var items = rows.map(function (r) {
    var computed = computePlanItem(r, Number(vehicle.odo) || 0);
    var out = stripRow(r);
    out.nextDate = computed.nextDate;
    out.nextOdo = computed.nextOdo;
    out.state = computed.state;
    return out;
  });
  return { ok: true, items: items };
}

function doPlanTick(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var vehicleId = payload.vehicleId, itemId = payload.itemId;
  if (!vehicleId || !itemId) return { ok: false, error: 'vehicleId and itemId required' };
  var date = payload.date || isoDate(new Date());
  var odo = payload.odo;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var vehicle = findVehicle(vehicleId, s.shopId);
    if (!vehicle) return { ok: false, error: 'vehicle not found' };

    var sh = getSheet('plan');
    var rows = sheetToObjects(sh);
    var target = null;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].id === itemId && rows[i].vehicleId === vehicleId && rows[i].shopId === s.shopId) { target = rows[i]; break; }
    }
    if (!target) return { ok: false, error: 'plan item not found' };

    var obj = stripRow(target);
    obj.lastDate = date;
    if (odo !== undefined && odo !== null && odo !== '') obj.lastOdo = Number(odo);

    var computed = computePlanItem(obj, Number(vehicle.odo) || 0);
    obj.status = computed.state;
    updateRowByIndex(sh, target.__row, HEADERS.plan, obj);

    obj.nextDate = computed.nextDate;
    obj.nextOdo = computed.nextOdo;
    obj.state = computed.state;
    return { ok: true, item: obj };
  } finally {
    lock.releaseLock();
  }
}

function doPlanSeed(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var vehicleId = payload.vehicleId;
  if (!vehicleId) return { ok: false, error: 'vehicleId required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var vehicle = findVehicle(vehicleId, s.shopId);
    if (!vehicle) return { ok: false, error: 'vehicle not found' };

    var sh = getSheet('plan');
    var existing = sheetToObjects(sh).filter(function (r) {
      return r.shopId === s.shopId && r.vehicleId === vehicleId;
    });
    if (existing.length > 0) return { ok: true, count: 0 };

    DEFAULT_PLAN_ITEMS.forEach(function (d) {
      var obj = {
        id: 'p_' + Utilities.getUuid().slice(0, 8),
        shopId: s.shopId,
        vehicleId: vehicleId,
        item: d.item,
        months: d.months,
        km: d.km,
        lastDate: '',
        lastOdo: '',
        cost: d.cost,
        note: d.note,
        status: 'none'
      };
      appendObject(sh, HEADERS.plan, obj);
    });
    return { ok: true, count: DEFAULT_PLAN_ITEMS.length };
  } finally {
    lock.releaseLock();
  }
}

function doReport(payload) {
  var s = validateToken(payload.token);
  if (!s) return { ok: false, error: 'auth' };
  var vehicleId = payload.vehicleId;
  if (!vehicleId) return { ok: false, error: 'vehicleId required' };
  var vehicle = findVehicle(vehicleId, s.shopId);
  if (!vehicle) return { ok: false, error: 'vehicle not found' };

  var recRes = doRecords({ token: payload.token, vehicleId: vehicleId });
  if (!recRes.ok) return recRes;
  var planRes = doPlan({ token: payload.token, vehicleId: vehicleId });
  if (!planRes.ok) return planRes;

  var records = recRes.records;
  var totalCost = 0, currency = '', dates = [];
  records.forEach(function (r) {
    totalCost += Number(r.cost) || 0;
    if (r.currency) currency = r.currency;
    if (r.date) dates.push(String(r.date));
  });
  dates.sort();

  var summary = {
    totalCost: totalCost,
    currency: currency,
    recordCount: records.length,
    firstDate: dates.length ? dates[0] : '',
    lastDate: dates.length ? dates[dates.length - 1] : '',
    odo: vehicle.odo
  };

  return { ok: true, vehicle: stripRow(vehicle), records: records, plan: planRes.items, summary: summary };
}

// ---------------------------------------------------------------------------
// Dispatch / HTTP entry points
// ---------------------------------------------------------------------------

function handleAction(payload) {
  payload = payload || {};
  switch (payload.action) {
    case 'login': return doLogin(payload);
    case 'logout': return doLogout(payload);
    case 'me': return doMe(payload);
    case 'vehicles': return doVehicles(payload);
    case 'vehicleSave': return doVehicleSave(payload);
    case 'vehicleDelete': return doVehicleDelete(payload);
    case 'records': return doRecords(payload);
    case 'recordSave': return doRecordSave(payload);
    case 'recordDelete': return doRecordDelete(payload);
    case 'plan': return doPlan(payload);
    case 'planTick': return doPlanTick(payload);
    case 'planSeed': return doPlanSeed(payload);
    case 'report': return doReport(payload);
    default: return { ok: false, error: 'unknown action' };
  }
}

function respond(payload) {
  var res;
  try {
    res = handleAction(payload);
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  if (!res || typeof res !== 'object') res = { ok: false, error: 'bad response' };
  res.version = VERSION;
  return res;
}

function jsonOutput(res) {
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = {};
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    payload = {};
  }
  return jsonOutput(respond(payload));
}

function doGet(e) {
  e = e || {};
  var q = (e.parameter || {});

  // --- owner bootstrap helpers (safe: create-only / read the owner's own SETUP tab) ---
  if (q.setup === 'health') {
    var names = [];
    try { getSS().getSheets().forEach(function (sh) { names.push(sh.getName()); }); } catch (err2) {
      return jsonOutput({ ok: false, error: String(err2) });
    }
    return jsonOutput({ ok: true, service: 'HT backend', sheets: names, shops: countShopsSafe() });
  }
  if (q.setup === 'init') {
    try {
      var msg = setupStep1_CreateSheets();
      return jsonOutput({ ok: true, result: msg });
    } catch (err3) { return jsonOutput({ ok: false, error: String(err3) }); }
  }
  if (q.setup === 'accounts') {
    try {
      var msg2 = setupStep2_CreateAccounts();
      return jsonOutput({ ok: true, result: msg2 });
    } catch (err4) { return jsonOutput({ ok: false, error: String(err4) }); }
  }

  var payload = {};
  try {
    payload = JSON.parse(q.p);
  } catch (err) {
    payload = {};
  }
  return jsonOutput(respond(payload));
}

/** Count of shop accounts — never returns names, salts or hashes. */
function countShopsSafe() {
  try { return sheetToObjects(getSheet('shops')).length; } catch (e) { return -1; }
}

// ---------------------------------------------------------------------------
// Owner-only setup menu (Sheet UI only — never exposed over HTTP)
// ---------------------------------------------------------------------------


// ---------------------------------------------------------------------------
// STANDALONE SETUP — run these from the Apps Script editor (Run ▸ pick function)
// The HT Admin menu only exists for bound scripts; these work standalone.
// ---------------------------------------------------------------------------

/** STEP 1 — creates all data sheets plus a SETUP tab you type accounts into. */
function setupStep1_CreateSheets() {
  initSheets();
  var ss = getSS();
  var sh = ss.getSheetByName('SETUP');
  if (!sh) {
    sh = ss.insertSheet('SETUP', 0);
    sh.getRange('A1:D1').setValues([['shop_id', 'shop_name', 'new_password', 'status']]);
    sh.getRange('A1:D1').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A3').setValue('HOW TO USE 用法:');
    sh.getRange('A4').setValue('1. Type shop_id / shop_name / new_password in row 2 (and more rows for more shops).');
    sh.getRange('A5').setValue('2. In the Apps Script editor run: setupStep2_CreateAccounts');
    sh.getRange('A6').setValue('3. The password cell is WIPED automatically once hashed. Never keep it here.');
    sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 320);
  }
  Logger.log('Sheets ready. Now type an account into the SETUP tab, then run setupStep2_CreateAccounts.');
  return 'OK — sheets created. Fill the SETUP tab, then run setupStep2_CreateAccounts.';
}

/** STEP 2 — hashes each password from the SETUP tab, creates the account, wipes the plaintext. */
function setupStep2_CreateAccounts() {
  var ss = getSS();
  var sh = ss.getSheetByName('SETUP');
  if (!sh) return 'Run setupStep1_CreateSheets first.';
  var last = sh.getLastRow();
  var done = 0, msgs = [];
  for (var r = 2; r <= last; r++) {
    var shopId = String(sh.getRange(r, 1).getValue() || '').trim().toLowerCase();
    var name   = String(sh.getRange(r, 2).getValue() || '').trim();
    var pass   = String(sh.getRange(r, 3).getValue() || '');
    if (!shopId || !pass) continue;
    if (pass.length < 8) { sh.getRange(r, 4).setValue('REJECTED: password must be at least 8 characters'); continue; }
    var salt = makeSalt();
    var hash = sha256Hex(salt + pass);
    var shops = getSheet('shops');
    var existing = findShopById(shopId);
    if (existing) {
      var rows = sheetToObjects(shops);
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i].shopId) === shopId) {
          shops.getRange(i + 2, 3).setValue(salt);
          shops.getRange(i + 2, 4).setValue(hash);
          shops.getRange(i + 2, 5).setValue(true);
          if (name) shops.getRange(i + 2, 2).setValue(name);
          break;
        }
      }
      msgs.push(shopId + ': password updated');
    } else {
      shops.appendRow([shopId, name || shopId, salt, hash, true, new Date()]);
      msgs.push(shopId + ': account created');
    }
    sh.getRange(r, 3).clearContent();               // wipe plaintext password
    sh.getRange(r, 4).setValue('DONE ' + new Date().toISOString().slice(0, 16).replace('T', ' ') + ' — password stored hashed, plaintext wiped');
    done++;
  }
  var out = done ? ('Processed ' + done + ' account(s): ' + msgs.join('; ')) : 'Nothing to process — type shop_id and new_password in row 2 of the SETUP tab.';
  Logger.log(out);
  return out;
}

/** Optional check — lists shop IDs (never passwords). */
function setupStep3_ListShops() {
  var rows = sheetToObjects(getSheet('shops'));
  var out = rows.map(function (r) { return r.shopId + ' (' + r.name + ') active=' + r.active; }).join('\n') || 'no shops yet';
  Logger.log(out);
  return out;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HT Admin')
    .addItem('1. Initialise sheets', 'adminInitSheets')
    .addItem('2. Add shop account', 'adminAddShop')
    .addItem('3. Reset shop password', 'adminResetPassword')
    .addItem('4. Deactivate shop', 'adminDeactivateShop')
    .addItem('5. List shops', 'adminListShops')
    .addItem('6. Purge expired sessions', 'adminPurgeSessions')
    .addToUi();
}

function adminInitSheets() {
  initSheets();
  SpreadsheetApp.getUi().alert('Sheets initialised.');
}

function adminAddShop() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt('Add shop account', 'Shop ID (short, unique, e.g. "ht-central"):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var shopId = String(r1.getResponseText() || '').trim().toLowerCase();
  if (!shopId) { ui.alert('Shop ID required.'); return; }

  var r2 = ui.prompt('Add shop account', 'Shop name:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var name = String(r2.getResponseText() || '').trim();

  var r3 = ui.prompt('Add shop account', 'Password:', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var pass = r3.getResponseText();
  if (!pass) { ui.alert('Password required.'); return; }

  if (findShopById(shopId)) { ui.alert('Shop ID already exists.'); return; }

  var salt = makeSalt();
  var passHash = sha256Hex(salt + pass);
  pass = null; // never retain plaintext

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    appendObject(getSheet('shops'), HEADERS.shops, {
      shopId: shopId, name: name, salt: salt, passHash: passHash, active: true, created: new Date().toISOString()
    });
  } finally {
    lock.releaseLock();
  }
  ui.alert('Shop "' + shopId + '" added.');
}

function adminResetPassword() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt('Reset shop password', 'Shop ID:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var shopId = String(r1.getResponseText() || '').trim().toLowerCase();

  var r2 = ui.prompt('Reset shop password', 'New password:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var pass = r2.getResponseText();
  if (!pass) { ui.alert('Password required.'); return; }

  var target = findShopById(shopId);
  if (!target) { ui.alert('Shop not found.'); return; }

  var salt = makeSalt();
  var passHash = sha256Hex(salt + pass);
  pass = null;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var obj = {}; HEADERS.shops.forEach(function (h) { obj[h] = target[h]; });
    obj.salt = salt; obj.passHash = passHash;
    updateRowByIndex(getSheet('shops'), target.__row, HEADERS.shops, obj);
  } finally {
    lock.releaseLock();
  }
  ui.alert('Password reset for "' + shopId + '".');
}

function adminDeactivateShop() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Deactivate shop', 'Shop ID:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var shopId = String(r1.getResponseText() || '').trim().toLowerCase();

  var target = findShopById(shopId);
  if (!target) { ui.alert('Shop not found.'); return; }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var obj = {}; HEADERS.shops.forEach(function (h) { obj[h] = target[h]; });
    obj.active = false;
    updateRowByIndex(getSheet('shops'), target.__row, HEADERS.shops, obj);
  } finally {
    lock.releaseLock();
  }
  ui.alert('Shop "' + shopId + '" deactivated.');
}

function adminListShops() {
  var ui = SpreadsheetApp.getUi();
  var rows = sheetToObjects(getSheet('shops'));
  if (!rows.length) { ui.alert('No shops yet.'); return; }
  var lines = rows.map(function (r) {
    return r.shopId + '  |  ' + r.name + '  |  active=' + r.active + '  |  created=' + r.created;
  });
  ui.alert('Shops:\n\n' + lines.join('\n'));
}

function adminPurgeSessions() {
  var ui = SpreadsheetApp.getUi();
  var now = Date.now();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var purged = 0;
  try {
    purged = deleteWhere('sessions', function (r) { return Number(r.expires) <= now; });
  } finally {
    lock.releaseLock();
  }
  ui.alert('Purged ' + purged + ' expired session(s).');
}
