/**
 * AGW Quote Maker - server (Google Apps Script, bound to the "AGW Quote Maker Data" Google Sheet).
 * Every request is checked here: mobile number + PIN, then an approved device. Nothing is sent to an
 * unapproved device. Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 */
var SH = { DATA: 'Data', STAFF: 'Staff', DEVICES: 'Devices', REG: 'Quote register' };
var HEAD = {
  Data: ['Collection', 'ID', 'Updated (ms)', 'By', 'Deleted', 'JSON'],
  Staff: ['ID', 'Name', 'Mobile', 'PIN (hashed)', 'Role', 'Active', 'Created'],
  Devices: ['Device code', 'Staff ID', 'Name', 'Device', 'Status', 'Requested', 'Approved'],
  'Quote register': ['Quote no.', 'Date', 'Client', 'Client phone', 'Grand total', 'Status', 'Salesperson', 'Updated', 'ID']
};
var ADMIN_ONLY = ['blandox', 'glass', 'colours', 'catalog', 'settings', 'products'];
var CHUNK = 45000;           // a Sheets cell holds 50,000 characters; long JSON spreads across cells
var TOKEN_DAYS = 60;

function doGet() { return out_({ ok: true, app: 'AGW Quote Maker' }); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ error: 'Bad request' }); }
  try { return out_(handle_(req)); }
  catch (err) { return out_({ error: String(err && err.message || err) }); }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------------- sheets ---------------- */
function book_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || '172uu1L1HxpB7rY1k6jvgppc47ZMoY5nmKHgwuOhVvs4';
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  var b = book_(), s = b.getSheetByName(name);
  if (!s) {
    s = b.insertSheet(name); s.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold'); s.setFrozenRows(1);
    if (name === SH.DATA) { s.getRange('A:B').setNumberFormat('@'); s.getRange('D:Z').setNumberFormat('@'); }
    if (name === SH.STAFF || name === SH.DEVICES) s.getRange('A:D').setNumberFormat('@');
  }
  return s;
}
function rows_(name) {
  var s = sheet_(name), n = s.getLastRow();
  if (n < 2) return [];
  var w = Math.max(s.getLastColumn(), HEAD[name].length);
  return s.getRange(2, 1, n - 1, w).getValues().map(function (r, i) { r._row = i + 2; return r; });
}
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function sha_(t) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)); }
function hmac_(t) { return hex_(Utilities.computeHmacSha256Signature(t, secret_())); }
function phone_(p) { var d = String(p || '').replace(/\D/g, ''); return d.slice(-10); }
function pinHash_(staffId, pin) { return sha_(staffId + ':' + String(pin) + ':' + secret_()); }
function now_() { return Date.now(); }
function esc_(c) { return typeof c === 'string' && /^[=+\-@']/.test(c) ? "'" + c : c; }

function staffList_() {
  return rows_(SH.STAFF).map(function (r) {
    return { id: String(r[0]), name: String(r[1]), phone: phone_(r[2]), pinHash: String(r[3] || ''), role: String(r[4] || 'sales'), active: r[5] !== false && String(r[5]).toUpperCase() !== 'FALSE' && String(r[5]).toUpperCase() !== 'NO', _row: r._row };
  }).filter(function (s) { return s.id; });
}
function deviceList_() {
  return rows_(SH.DEVICES).map(function (r) {
    return { hash: String(r[0]), staffId: String(r[1]), name: String(r[2]), label: String(r[3]), status: String(r[4] || 'pending'), requestedAt: r[5] ? new Date(r[5]).toISOString() : '', approvedAt: r[6] ? new Date(r[6]).toISOString() : '', _row: r._row };
  }).filter(function (d) { return d.hash; });
}
function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'); }

/* ---------------- auth ---------------- */
function makeToken_(staffId, hash) {
  var exp = now_() + TOKEN_DAYS * 864e5, body = staffId + '|' + hash + '|' + exp;
  return Utilities.base64EncodeWebSafe(body) + '.' + hmac_(body);
}
function auth_(req) {
  var t = String(req.token || '').split('.');
  if (t.length !== 2) throw new Error('LOGIN');
  var body = Utilities.newBlob(Utilities.base64DecodeWebSafe(t[0])).getDataAsString();
  if (hmac_(body) !== t[1]) throw new Error('LOGIN');
  var p = body.split('|');
  if (+p[2] < now_() || p[1] !== String(req.device || '')) throw new Error('LOGIN');
  var st = staffList_().filter(function (s) { return s.id === p[0] && s.active; })[0];
  if (!st) throw new Error('LOGIN');
  var dev = deviceList_().filter(function (d) { return d.hash === p[1] && d.staffId === st.id; })[0];
  if (!dev || dev.status !== 'approved') throw new Error('LOGIN');
  return { staff: st, isAdmin: st.role === 'admin', device: dev };
}

function login_(req) {
  var ph = phone_(req.phone), pin = String(req.pin || ''), hash = String(req.device || '');
  if (ph.length !== 10 || !/^\d{4,6}$/.test(pin) || !/^[0-9a-f]{64}$/.test(hash)) throw new Error('Enter your 10-digit mobile number and PIN.');
  var cache = CacheService.getScriptCache(), key = 'fail:' + ph, fails = +(cache.get(key) || 0);
  if (fails >= 5) throw new Error('Too many wrong PINs. Try again after 15 minutes.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var st = staffList_().filter(function (s) { return s.phone === ph && s.active; })[0];
    if (st && !st.pinHash) throw new Error('Your PIN has not been set yet. Ask the owner for your PIN.');
    if (!st || st.pinHash !== pinHash_(st.id, pin)) { cache.put(key, String(fails + 1), 900); throw new Error('Wrong mobile number or PIN.'); }
    cache.remove(key);
    var devs = deviceList_(), dev = devs.filter(function (d) { return d.hash === hash && d.staffId === st.id; })[0];
    var label = String(req.label || 'Device').slice(0, 60), code = hash.slice(0, 6).toUpperCase();
    if (!dev) {
      var admins = staffList_().filter(function (s) { return s.role === 'admin'; }).map(function (s) { return s.id; });
      var firstOwnerDevice = st.role === 'admin' && !devs.some(function (d) { return d.status === 'approved' && admins.indexOf(d.staffId) >= 0; });
      sheet_(SH.DEVICES).appendRow([hash, st.id, st.name, label, firstOwnerDevice ? 'approved' : 'pending', stamp_(), firstOwnerDevice ? stamp_() : '']);
      if (!firstOwnerDevice) return { status: 'pending', code: code, name: st.name };
      return { status: 'ok', token: makeToken_(st.id, hash), me: me_(st) };
    }
    if (dev.status === 'blocked') return { status: 'blocked', code: code, name: st.name };
    if (dev.status !== 'approved') return { status: 'pending', code: code, name: st.name };
    return { status: 'ok', token: makeToken_(st.id, hash), me: me_(st) };
  } finally { lock.releaseLock(); }
}
function me_(st) { return { id: st.id, name: st.name, phone: st.phone, role: st.role }; }

/* ---------------- data ---------------- */
function readData_() {
  return rows_(SH.DATA).map(function (r) {
    var json = r.slice(5).filter(function (c) { return c !== '' && c !== null; }).join('');
    return { coll: String(r[0]), id: String(r[1]), updatedAt: +r[2] || 0, by: String(r[3] || ''), deleted: r[4] === true || String(r[4]).toUpperCase() === 'TRUE', json: json, _row: r._row };
  });
}
function visible_(d, data, a) {
  if (a.isAdmin) return true;
  if (d.coll === 'quotes') return d.by === a.staff.id || (data && data.sp && data.sp.id === a.staff.id);
  if (d.coll === 'clients') return d.by === a.staff.id;
  return true;
}
function sync_(req, a) {
  var since = +req.since || 0, out = [];
  readData_().forEach(function (d) {
    if (d.updatedAt <= since) return;
    var data = null; if (!d.deleted && d.json) { try { data = JSON.parse(d.json); } catch (x) { return; } }
    if (!visible_(d, data, a)) return;
    out.push({ coll: d.coll, id: d.id, updatedAt: d.updatedAt, deleted: d.deleted, data: data });
  });
  var staff = staffList_().filter(function (s) { return a.isAdmin || s.id === a.staff.id; })
    .map(function (s) { return { id: s.id, name: s.name, phone: s.phone, role: s.role, active: s.active, hasPin: !!s.pinHash }; });
  var res = { now: now_(), docs: out, staff: staff, me: me_(a.staff) };
  if (a.isAdmin) res.devices = deviceList_().map(function (d) { var x = {}; for (var k in d) if (k !== '_row') x[k] = d[k]; x.code = d.hash.slice(0, 6).toUpperCase(); return x; });
  return res;
}
function canWrite_(coll, existing, a) {
  if (!/^[a-zA-Z]{1,30}$/.test(coll)) throw new Error('Bad collection');
  if (ADMIN_ONLY.indexOf(coll) >= 0 && !a.isAdmin) throw new Error('Only the owner can change this.');
  if (existing && !a.isAdmin && (coll === 'quotes' || coll === 'clients') && existing.by && existing.by !== a.staff.id) throw new Error('This belongs to another salesperson.');
}
function put_(coll, id, data, deleted, a) {
  var s = sheet_(SH.DATA), all = readData_(), ex = all.filter(function (d) { return d.coll === coll && d.id === id; })[0];
  canWrite_(coll, ex, a);
  var by = ex && ex.by ? ex.by : a.staff.id, json = deleted ? '' : JSON.stringify(data || {}), parts = [];
  for (var i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  var t = now_(), row = [coll, id, t, by, !!deleted].concat(parts.length ? parts : ['']);
  var r = ex ? ex._row : s.getLastRow() + 1, oldW = ex ? Math.max(s.getLastColumn() - 5, 1) : 1;
  if (ex && oldW > parts.length) s.getRange(r, 6, 1, oldW).clearContent();
  s.getRange(r, 1, 1, row.length).setValues([row.map(esc_)]);
  if (coll === 'quotes') register_(id, deleted ? null : data);
  return t;
}
function register_(id, q) {
  var s = sheet_(SH.REG), rs = rows_(SH.REG), ex = rs.filter(function (r) { return String(r[8]) === id; })[0];
  if (!q) { if (ex) s.deleteRow(ex._row); return; }
  var c = q.client || {}, row = [q.number || '', q.date || '', c.name || '', c.phone || '', +q.grandTotal || 0, q.status || '', (q.sp && q.sp.name) || '', stamp_(), id];
  row = row.map(esc_);
  if (ex) s.getRange(ex._row, 1, 1, row.length).setValues([row]); else s.appendRow(row);
}

/* ---------------- admin ---------------- */
function saveStaff_(req, a) {
  if (!a.isAdmin) throw new Error('Only the owner can change this.');
  var s = sheet_(SH.STAFF), all = staffList_(), ph = phone_(req.phone);
  if (!String(req.name || '').trim() || ph.length !== 10) throw new Error('Enter a name and 10-digit mobile number.');
  var ex = req.id ? all.filter(function (x) { return x.id === req.id; })[0] : null;
  if (all.some(function (x) { return x.phone === ph && (!ex || x.id !== ex.id); })) throw new Error('This mobile number is already used by another person.');
  if (req.pin && !/^\d{4,6}$/.test(String(req.pin))) throw new Error('PIN must be 4 to 6 digits.');
  var role = req.role === 'admin' ? 'admin' : 'sales', active = req.active !== false;
  if (ex && ex.id === a.staff.id && (role !== 'admin' || !active)) throw new Error("You can't remove your own owner access.");
  var id = ex ? ex.id : 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  var pin = req.pin ? pinHash_(id, req.pin) : (ex ? ex.pinHash : '');
  var row = [id, String(req.name).trim(), ph, pin, role, active, ex ? s.getRange(ex._row, 7).getValue() : stamp_()];
  if (ex) s.getRange(ex._row, 1, 1, row.length).setValues([row]); else s.appendRow(row);
  return { id: id };
}
function setDevice_(req, a, status) {
  if (!a.isAdmin) throw new Error('Only the owner can change this.');
  var d = deviceList_().filter(function (x) { return x.hash === req.hash && x.staffId === req.staffId; })[0];
  if (!d) throw new Error('Device not found.');
  if (d.hash === a.device.hash && d.staffId === a.staff.id && status !== 'approved') throw new Error("You can't block the device you are using.");
  var s = sheet_(SH.DEVICES);
  if (status === 'remove') { s.deleteRow(d._row); return {}; }
  s.getRange(d._row, 5).setValue(status); if (status === 'approved') s.getRange(d._row, 7).setValue(stamp_());
  return {};
}
function sheetValues_(req, a) {
  if (!a.isAdmin) throw new Error('Only the owner can change this.');
  var b = SpreadsheetApp.openById(String(req.spreadsheetId));
  return { values: (req.ranges || []).map(function (rg) { return b.getRange(rg).getDisplayValues(); }) };
}

function handle_(req) {
  if (req.action === 'login') return login_(req);
  var a;
  try { a = auth_(req); } catch (e) { return { error: 'LOGIN' }; }
  switch (req.action) {
    case 'sync': return sync_(req, a);
    case 'set': case 'delete': {
      var lock = LockService.getScriptLock(); lock.waitLock(20000);
      try { return { updatedAt: put_(String(req.coll), String(req.id), req.data, req.action === 'delete', a) }; }
      finally { lock.releaseLock(); }
    }
    case 'saveStaff': { var l1 = LockService.getScriptLock(); l1.waitLock(20000); try { return saveStaff_(req, a); } finally { l1.releaseLock(); } }
    case 'approveDevice': return setDevice_(req, a, 'approved');
    case 'blockDevice': return setDevice_(req, a, 'blocked');
    case 'removeDevice': return setDevice_(req, a, 'remove');
    case 'sheetValues': return sheetValues_(req, a);
    default: return { error: 'Unknown action' };
  }
}

/* Run once from the editor (select setupOwner, click Run): creates your owner login and imports the data from the old app. */
var SEED_ID = '1XGA5Bwi5rYgqsup-aV6AmEZ2e-jXBIVab0fhhQz5DWY';
function importSeed_() {
  var data = sheet_(SH.DATA);
  if (data.getLastRow() > 1 || !SEED_ID) return 0;
  var src; try { src = SpreadsheetApp.openById(SEED_ID).getSheets()[0]; } catch (x) { return 0; }
  if (!src) return 0; var n = src.getLastRow();
  if (n < 1) return 0;
  var vals = src.getRange(1, 1, n, 25).getValues(), t = now_(), out = [];
  vals.forEach(function (r) {
    if (!r[0] || !r[1]) return;
    var parts = r.slice(5, 25).filter(function (c) { return c !== '' && c !== null; }).map(String);
    var row = [String(r[0]), String(r[1]), t, 'owner', false].concat(parts);
    data.getRange(data.getLastRow() + 1, 1, 1, row.length).setValues([row.map(esc_)]);
    if (r[0] === 'quotes') { try { register_(String(r[1]), JSON.parse(parts.join(''))); } catch (x) {} }
    out.push(r[1]);
  });
  return out.length;
}
function setupOwner() {
  var all = staffList_();
  if (all.some(function (s) { return s.role === 'admin'; })) { Logger.log('Owner already set up. Imported ' + importSeed_() + ' records.'); return; }
  var pin = String(1000 + Math.floor(Math.random() * 9000));
  sheet_(SH.STAFF).appendRow(['owner', 'Lakshay Jindal', '9814043755', pinHash_('owner', pin), 'admin', true, stamp_()]);
  [SH.DATA, SH.DEVICES, SH.REG].forEach(sheet_);
  var imported = importSeed_();
  var first = book_().getSheetByName('Sheet1'); if (first && first.getLastRow() === 0 && book_().getSheets().length > 1) book_().deleteSheet(first);
  Logger.log('Imported ' + imported + ' records from the old app.');
  Logger.log('Owner login ready. Mobile: 9814043755  PIN: ' + pin + '  (change it later in Team & devices). The first device you log in on is approved automatically.');
}
