/**
 * מערכת שריון רכבי מחלקה — צד שרת (Google Apps Script)
 * הנתונים נשמרים בגיליון שאליו הסקריפט מחובר.
 * הפעלה ראשונה: להריץ את הפונקציה setup פעם אחת מתוך העורך.
 */

var INITIAL_ADMIN_NAME = 'מנהל';
var INITIAL_ADMIN_PASSWORD = 'CHANGE_ME'; // סיסמה זמנית — אחרי שמוסיפים את המזכירה אפשר למחוק את המשתמש הזה
var TOKEN_DAYS = 30;
var MAX_FAILS = 5;

var SHEETS = {
  Users: ['name', 'passHash', 'role', 'createdAt'],
  Cars: ['id', 'name', 'plate', 'active'],
  Requests: ['id', 'userName', 'fromDate', 'toDate', 'departTime', 'destination', 'project', 'notes',
             'status', 'carId', 'adminNote', 'createdAt', 'decidedAt', 'decidedBy']
};

/* ---------- התקנה ---------- */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.getRange('A:Z').setNumberFormat('@');
    sh.setFrozenRows(1);
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('גיליון1');
  if (def && ss.getSheets().length > 3) ss.deleteSheet(def);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());

  if (readRows('Cars').length === 0) {
    for (var i = 1; i <= 5; i++) appendRow('Cars', { id: 'c' + i, name: 'רכב ' + i, plate: '', active: 'yes' });
  }
  if (!findUser(INITIAL_ADMIN_NAME)) {
    appendRow('Users', { name: INITIAL_ADMIN_NAME, passHash: hashPass(INITIAL_ADMIN_PASSWORD), role: 'admin', createdAt: nowStr() });
  }
  Logger.log('ההתקנה הושלמה');
}

/* ---------- נקודות כניסה ---------- */

function doGet() {
  return json({ ok: true, service: 'car-booking' });
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    out = handle(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return json(out);
}

function handle(req) {
  var a = req.action;
  if (a === 'login') return login(req.name, req.pass);

  var me = verifyToken(req.token);
  if (!me) return { ok: false, error: 'AUTH', message: 'יש להתחבר מחדש' };

  if (a === 'me') return { ok: true, name: me.name, role: me.role };
  if (a === 'myRequests') return { ok: true, requests: myRequests(me.name), cars: activeCars() };

  var res;
  if (a === 'createRequest') res = withLock(function () { return createRequest(me, req); });
  else if (a === 'cancelRequest') res = withLock(function () { return cancelRequest(me, req.id); });
  else if (a === 'cancelDays') res = withLock(function () { return cancelDays(me, req.id, req.days); });
  if (res) {
    // מחזירים את הרשימה המעודכנת באותה תשובה, כדי לחסוך פנייה נוספת לשרת
    if (res.ok) res.mine = { requests: myRequests(me.name), cars: activeCars() };
    return res;
  }

  if (me.role !== 'admin') return { ok: false, message: 'אין הרשאה' };
  if (a === 'adminData') return adminData();
  var fns = {
    decide: function () { return decide(me, req); },
    addUser: function () { return addUser(req); },
    updateUser: function () { return updateUser(me, req); },
    removeUser: function () { return removeUser(me, req.name); },
    addCar: function () { return addCar(req); },
    updateCar: function () { return updateCar(req); },
    removeCar: function () { return removeCar(req.id); },
    assignCars: function () { return assignCars(me, req.changes); },
    adminCreate: function () { return adminCreate(me, req); },
    adminEdit: function () { return adminEdit(me, req); }
  };
  if (!fns[a]) return { ok: false, message: 'פעולה לא מוכרת' };
  res = withLock(fns[a]);
  var d = adminData();
  res.data = { requests: d.requests, cars: d.cars, users: d.users };
  return res;
}

/* ---------- התחברות ---------- */

function login(name, pass) {
  name = normName(name);
  pass = String(pass || '').replace(/[\s-]/g, '');
  if (!name || !pass) return { ok: false, message: 'יש למלא שם וסיסמה' };
  var cache = CacheService.getScriptCache();
  var key = 'fail_' + Utilities.base64EncodeWebSafe(name);
  var fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) return { ok: false, message: 'יותר מדי ניסיונות. נסה שוב בעוד 15 דקות' };

  var u = findUser(name);
  if (!u || u.passHash !== hashPass(pass)) {
    cache.put(key, String(fails + 1), 900);
    return { ok: false, message: 'שם או סיסמה שגויים' };
  }
  cache.remove(key);
  return { ok: true, token: makeToken(u.name, u.role), name: u.name, role: u.role };
}

function makeToken(name, role) {
  var payload = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify({
    n: name, r: role, e: Date.now() + TOKEN_DAYS * 86400000
  })).getBytes());
  return payload + '.' + sign(payload);
}

function verifyToken(token) {
  if (!token || token.indexOf('.') < 0) return null;
  var parts = token.split('.');
  if (sign(parts[0]) !== parts[1]) return null;
  var p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if (p.e < Date.now()) return null;
  var u = findUser(p.n);
  if (!u) return null;
  return { name: u.name, role: u.role };
}

function sign(s) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret));
}

function hashPass(pass) {
  var salt = PropertiesService.getScriptProperties().getProperty('SALT');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pass, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/* ---------- בקשות ---------- */

function myRequests(name) {
  return readRows('Requests').filter(function (r) { return r.userName === name; })
    .sort(function (a, b) { return b.createdAt < a.createdAt ? -1 : 1; }).slice(0, 50);
}

function createRequest(me, q) {
  var from = cleanDate(q.fromDate), to = cleanDate(q.toDate) || from;
  if (!from) return { ok: false, message: 'יש לבחור תאריך' };
  if (to < from) return { ok: false, message: 'תאריך הסיום לפני תאריך ההתחלה' };
  var dest = clip(q.destination, 120);
  if (!dest) return { ok: false, message: 'יש למלא לאן הנסיעה' };
  appendRow('Requests', {
    id: 'r' + Date.now().toString(36) + Math.floor(Math.random() * 1000),
    userName: me.name, fromDate: from, toDate: to,
    departTime: clip(q.departTime, 5), destination: dest, project: clip(q.project, 40), notes: clip(q.notes, 300),
    status: 'pending', carId: '', adminNote: '', createdAt: nowStr(), decidedAt: '', decidedBy: ''
  });
  return { ok: true };
}

function cancelRequest(me, id) {
  var r = findRow('Requests', 'id', id);
  if (!r || (r.data.userName !== me.name && me.role !== 'admin')) return { ok: false, message: 'הבקשה לא נמצאה' };
  if (r.data.status !== 'pending' && r.data.status !== 'approved') return { ok: false, message: 'אי אפשר לבטל את הבקשה הזאת' };
  var today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  if (r.data.status === 'approved' && r.data.toDate < today && me.role !== 'admin') return { ok: false, message: 'אי אפשר לבטל נסיעה שכבר עברה' };
  // שומרים את הרכב שהיה משובץ, כדי שהמזכירה תראה איזה רכב התפנה
  updateRow('Requests', r.row, { status: 'cancelled', adminNote: r.data.status === 'approved' ? 'בוטל אחרי אישור' : r.data.adminNote, decidedAt: nowStr(), decidedBy: me.name });
  return { ok: true };
}

/* ביטול ימים מסוימים מתוך בקשה של כמה ימים. הימים שנשארו ממשיכים עם אותו סטטוס ואותו רכב */
function cancelDays(me, id, days) {
  var r = findRow('Requests', 'id', id);
  if (!r || (r.data.userName !== me.name && me.role !== 'admin')) return { ok: false, message: 'הבקשה לא נמצאה' };
  var q = r.data;
  if (q.status !== 'pending' && q.status !== 'approved') return { ok: false, message: 'אי אפשר לבטל את הבקשה הזאת' };
  var all = dateList(q.fromDate, q.toDate);
  var today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  var drop = {};
  (days || []).forEach(function (d) { if (all.indexOf(d) >= 0) drop[d] = 1; });
  var dropped = all.filter(function (d) { return drop[d]; });
  if (!dropped.length) return { ok: false, message: 'יש לבחור לפחות יום אחד' };
  if (me.role !== 'admin' && dropped.some(function (d) { return d < today; })) return { ok: false, message: 'אי אפשר לבטל יום שכבר עבר' };
  var keep = all.filter(function (d) { return !drop[d]; });
  if (!keep.length) return cancelRequest(me, id);

  var keepRuns = runs(keep), dropRuns = runs(dropped);
  updateRow('Requests', r.row, { fromDate: keepRuns[0][0], toDate: keepRuns[0][1] });
  for (var i = 1; i < keepRuns.length; i++) {
    var copy = JSON.parse(JSON.stringify(q));
    copy.id = newId(); copy.fromDate = keepRuns[i][0]; copy.toDate = keepRuns[i][1];
    appendRow('Requests', copy);
  }
  dropRuns.forEach(function (run) {
    var c = JSON.parse(JSON.stringify(q));
    c.id = newId(); c.fromDate = run[0]; c.toDate = run[1];
    c.status = 'cancelled'; c.carId = q.status === 'approved' ? q.carId : '';
    c.adminNote = 'בוטל חלקית'; c.decidedAt = nowStr(); c.decidedBy = me.name;
    appendRow('Requests', c);
  });
  return { ok: true };
}

function newId() { return 'r' + Date.now().toString(36) + Math.floor(Math.random() * 100000).toString(36); }

function pad2(n) { return ('0' + n).slice(-2); }
function dateList(from, to) {
  var p = from.split('-'), d = new Date(+p[0], p[1] - 1, +p[2]), out = [];
  for (var i = 0; i < 400; i++) {
    var s = d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    if (s > to) break;
    out.push(s); d.setDate(d.getDate() + 1);
  }
  return out;
}
/* הופך רשימת תאריכים לרצפים: [[מ, עד], ...] */
function runs(list) {
  var out = [];
  list.forEach(function (d) {
    var last = out[out.length - 1];
    if (last && dateList(last[1], d).length === 2) last[1] = d; else out.push([d, d]);
  });
  return out;
}

/* המזכירה מוסיפה נסיעה בשם עובד. עם רכב – נשמרת כמאושרת. בלי רכב – נשמרת כממתינה */
function adminCreate(me, q) {
  var user = findUser(q.userName);
  if (!user) return { ok: false, message: 'יש לבחור עובד מהרשימה' };
  var from = cleanDate(q.fromDate), to = cleanDate(q.toDate) || from;
  if (!from) return { ok: false, message: 'יש לבחור תאריך' };
  if (to < from) return { ok: false, message: 'תאריך הסיום לפני תאריך ההתחלה' };
  var dest = clip(q.destination, 120);
  if (!dest) return { ok: false, message: 'יש למלא לאן הנסיעה' };
  var carId = String(q.carId || '');
  if (carId) {
    var car = findRow('Cars', 'id', carId);
    if (!car || car.data.active !== 'yes') return { ok: false, message: 'הרכב לא נמצא' };
    var clash = readRows('Requests').filter(function (x) { return x.status === 'approved' && x.carId === carId && x.fromDate <= to && x.toDate >= from; });
    if (clash.length) return { ok: false, message: car.data.name + ' כבר משובץ ל' + clash[0].userName + ' בתאריכים האלה' };
  }
  appendRow('Requests', {
    id: newId(), userName: user.name, fromDate: from, toDate: to,
    departTime: clip(q.departTime, 5), destination: dest, project: clip(q.project, 40), notes: clip(q.notes, 300),
    status: carId ? 'approved' : 'pending', carId: carId, adminNote: clip(q.adminNote, 200),
    createdAt: nowStr(), decidedAt: carId ? nowStr() : '', decidedBy: carId ? me.name : ''
  });
  return { ok: true };
}

/* המזכירה עורכת נסיעה קיימת: תאריכים, שעה, יעד, פרויקט, הערות ורכב */
function adminEdit(me, q) {
  var r = findRow('Requests', 'id', q.id);
  if (!r) return { ok: false, message: 'הנסיעה לא נמצאה' };
  if (r.data.status !== 'approved' && r.data.status !== 'pending') return { ok: false, message: 'אפשר לערוך רק נסיעה מאושרת או ממתינה' };
  var from = cleanDate(q.fromDate), to = cleanDate(q.toDate) || from;
  if (!from) return { ok: false, message: 'יש לבחור תאריך' };
  if (to < from) return { ok: false, message: 'תאריך הסיום לפני תאריך ההתחלה' };
  var dest = clip(q.destination, 120);
  if (!dest) return { ok: false, message: 'יש למלא לאן הנסיעה' };
  var carId = String(q.carId || '');
  if (carId) {
    var car = findRow('Cars', 'id', carId);
    if (!car || car.data.active !== 'yes') return { ok: false, message: 'הרכב לא נמצא' };
    var clash = readRows('Requests').filter(function (x) { return x.id !== q.id && x.status === 'approved' && x.carId === carId && x.fromDate <= to && x.toDate >= from; });
    if (clash.length) return { ok: false, message: car.data.name + ' כבר משובץ ל' + clash[0].userName + ' בתאריכים האלה' };
  }
  updateRow('Requests', r.row, {
    fromDate: from, toDate: to, departTime: clip(q.departTime, 5), destination: dest,
    project: clip(q.project, 40), notes: clip(q.notes, 300), adminNote: clip(q.adminNote, 200),
    carId: carId, status: carId ? 'approved' : 'pending', decidedAt: nowStr(), decidedBy: me.name
  });
  return { ok: true };
}

/* שמירה של כמה שינויי רכב בבת אחת. נשמר רק אם אין רכב שמשובץ לשני אנשים באותו יום */
function assignCars(me, changes) {
  changes = changes || [];
  if (!changes.length) return { ok: false, message: 'אין שינויים לשמור' };
  var all = readRows('Requests'), byId = {};
  all.forEach(function (x) { byId[x.id] = x; });
  var cars = {}; readRows('Cars').forEach(function (c) { if (c.active === 'yes') cars[c.id] = c; });
  for (var i = 0; i < changes.length; i++) {
    var ch = changes[i], x = byId[ch.id];
    if (!x || x.status !== 'approved') return { ok: false, message: 'אחת הנסיעות כבר לא מאושרת. רענן את הדף' };
    if (!cars[ch.carId]) return { ok: false, message: 'יש לבחור רכב לנסיעה של ' + x.userName };
    x.carId = ch.carId;
  }
  var appr = all.filter(function (x) { return x.status === 'approved'; });
  for (var a = 0; a < appr.length; a++) for (var b = a + 1; b < appr.length; b++) {
    var p = appr[a], q = appr[b];
    if (p.carId === q.carId && p.fromDate <= q.toDate && p.toDate >= q.fromDate)
      return { ok: false, message: cars[p.carId].name + ' משובץ גם ל' + p.userName + ' וגם ל' + q.userName + ' באותם ימים. תקן ונסה שוב' };
  }
  changes.forEach(function (ch) {
    var r = findRow('Requests', 'id', ch.id);
    updateRow('Requests', r.row, { carId: ch.carId, decidedAt: nowStr(), decidedBy: me.name });
  });
  return { ok: true };
}

/* החלפת רכב לנסיעה מאושרת. אם הרכב החדש תפוס באותם ימים, הנהג השני מקבל את הרכב הקודם */
function swapCar(me, r, car) {
  var q = r.data, oldCar = q.carId, newCar = car.data.id;
  if (q.status !== 'approved') return { ok: false, message: 'אפשר להחליף רכב רק בנסיעה מאושרת' };
  if (oldCar === newCar) return { ok: false, message: 'זה כבר הרכב של הנסיעה הזאת' };
  var all = readRows('Requests');
  var overlap = function (a, b) { return a.fromDate <= b.toDate && a.toDate >= b.fromDate; };
  var others = all.filter(function (x) { return x.id !== q.id && x.status === 'approved' && x.carId === newCar && overlap(x, q); });
  var otherIds = {}; others.forEach(function (x) { otherIds[x.id] = 1; });
  for (var i = 0; i < others.length; i++) {
    var x = others[i];
    var blocked = all.filter(function (y) { return y.id !== q.id && !otherIds[y.id] && y.status === 'approved' && y.carId === oldCar && overlap(y, x); });
    if (blocked.length) {
      var oc = findRow('Cars', 'id', oldCar);
      return { ok: false, message: 'אי אפשר להחליף: ' + x.userName + ' צריך את ' + car.data.name + ' גם ב-' + x.fromDate + '–' + x.toDate + ', ו' + (oc ? oc.data.name : 'הרכב הקודם') + ' תפוס אז אצל ' + blocked[0].userName };
    }
  }
  updateRow('Requests', r.row, { carId: newCar, decidedAt: nowStr(), decidedBy: me.name });
  others.forEach(function (x) {
    var xr = findRow('Requests', 'id', x.id);
    updateRow('Requests', xr.row, { carId: oldCar, decidedAt: nowStr(), decidedBy: me.name });
  });
  return { ok: true, swapped: others.map(function (x) { return x.userName; }) };
}

function adminData() {
  var reqs = readRows('Requests');
  var cutoff = Utilities.formatDate(new Date(Date.now() - 60 * 86400000), tz(), 'yyyy-MM-dd');
  reqs = reqs.filter(function (r) { return r.status === 'pending' || r.toDate >= cutoff; });
  var users = readRows('Users').map(function (u) { return { name: u.name, role: u.role }; });
  return { ok: true, requests: reqs, cars: readRows('Cars'), users: users };
}

function decide(me, q) {
  var r = findRow('Requests', 'id', q.id);
  if (!r) return { ok: false, message: 'הבקשה לא נמצאה' };
  if (q.decision === 'reject') {
    updateRow('Requests', r.row, { status: 'rejected', carId: '', adminNote: clip(q.note, 200), decidedAt: nowStr(), decidedBy: me.name });
    return { ok: true };
  }
  if (q.decision === 'unassign') {
    updateRow('Requests', r.row, { status: 'pending', carId: '', decidedAt: nowStr(), decidedBy: me.name });
    return { ok: true };
  }
  var car = findRow('Cars', 'id', q.carId);
  if (!car || car.data.active !== 'yes') return { ok: false, message: 'יש לבחור רכב' };
  if (q.decision === 'swap') return swapCar(me, r, car);
  var clash = readRows('Requests').filter(function (x) {
    return x.id !== q.id && x.status === 'approved' && x.carId === q.carId &&
      x.fromDate <= r.data.toDate && x.toDate >= r.data.fromDate;
  });
  if (clash.length) return { ok: false, message: car.data.name + ' כבר משובץ ל' + clash[0].userName + ' בתאריכים האלה' };
  updateRow('Requests', r.row, { status: 'approved', carId: q.carId, adminNote: clip(q.note, 200), decidedAt: nowStr(), decidedBy: me.name });
  return { ok: true };
}

/* ---------- עובדים ורכבים ---------- */

function addUser(q) {
  var name = normName(q.name);
  var pass = String(q.pass || '').replace(/[\s-]/g, '');
  if (!name) return { ok: false, message: 'יש למלא שם' };
  if (pass.length < 5) return { ok: false, message: 'יש למלא סיסמה (לפחות 5 תווים)' };
  var role = q.role === 'admin' ? 'admin' : 'user';
  var existing = findRow('Users', 'name', name);
  if (existing) {
    updateRow('Users', existing.row, { passHash: hashPass(pass), role: role });
    return { ok: true, updated: true };
  }
  appendRow('Users', { name: name, passHash: hashPass(pass), role: role, createdAt: nowStr() });
  return { ok: true };
}

function updateUser(me, q) {
  var old = normName(q.oldName), name = normName(q.name);
  var r = findRow('Users', 'name', old);
  if (!r) return { ok: false, message: 'העובד לא נמצא' };
  if (!name) return { ok: false, message: 'יש למלא שם' };
  if (name !== old && findRow('Users', 'name', name)) return { ok: false, message: 'כבר קיים עובד בשם הזה' };
  var role = q.role === 'admin' ? 'admin' : 'user';
  if (old === me.name && role !== 'admin') return { ok: false, message: 'אי אפשר להוריד לעצמך את הרשאת המנהל' };
  var pass = String(q.pass || '').replace(/[\s-]/g, '');
  if (pass && pass.length < 5) return { ok: false, message: 'הסיסמה קצרה מדי (לפחות 5 תווים)' };
  var patch = { name: name, role: role };
  if (pass) patch.passHash = hashPass(pass);
  updateRow('Users', r.row, patch);
  if (name !== old) {
    var sh = sheet('Requests'), last = sh.getLastRow();
    if (last >= 2) {
      var rng = sh.getRange(2, 2, last - 1, 1);
      rng.setValues(rng.getValues().map(function (v) { return [String(v[0]) === old ? name : v[0]]; }));
    }
  }
  var out = { ok: true };
  if (old === me.name) { out.token = makeToken(name, role); out.name = name; }
  return out;
}

function removeUser(me, name) {
  if (normName(name) === me.name) return { ok: false, message: 'אי אפשר למחוק את המשתמש שאיתו אתה מחובר' };
  var r = findRow('Users', 'name', normName(name));
  if (!r) return { ok: false, message: 'העובד לא נמצא' };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').deleteRow(r.row);
  return { ok: true };
}

function activeCars() {
  return readRows('Cars').filter(function (c) { return c.active === 'yes'; });
}

function addCar(q) {
  var name = clip(q.name, 40);
  if (!name) return { ok: false, message: 'יש למלא שם רכב' };
  appendRow('Cars', { id: 'c' + Date.now().toString(36), name: name, plate: clip(q.plate, 20), active: 'yes' });
  return { ok: true };
}

function updateCar(q) {
  var r = findRow('Cars', 'id', q.id);
  if (!r) return { ok: false, message: 'הרכב לא נמצא' };
  updateRow('Cars', r.row, { name: clip(q.name, 40) || r.data.name, plate: clip(q.plate, 20) });
  return { ok: true };
}

function removeCar(id) {
  var r = findRow('Cars', 'id', id);
  if (!r) return { ok: false, message: 'הרכב לא נמצא' };
  updateRow('Cars', r.row, { active: 'no' });
  return { ok: true };
}

/* ---------- עזרי גיליון ---------- */

function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function cellStr(v, key) {
  if (v instanceof Date) {
    if (key === 'departTime') return Utilities.formatDate(v, tz(), 'HH:mm');
    return Utilities.formatDate(v, tz(), 'yyyy-MM-dd');
  }
  return String(v === null || v === undefined ? '' : v);
}

function readRows(name) {
  var sh = sheet(name), cols = SHEETS[name];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map(function (row) {
    var o = {};
    cols.forEach(function (c, i) { o[c] = cellStr(row[i], c); });
    return o;
  }).filter(function (o) { return o[cols[0]] !== ''; });
}

function findRow(name, key, value) {
  var sh = sheet(name), cols = SHEETS[name];
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var k = cols.indexOf(key);
  for (var i = 0; i < vals.length; i++) {
    if (cellStr(vals[i][k], key) === String(value)) {
      var o = {};
      cols.forEach(function (c, j) { o[c] = cellStr(vals[i][j], c); });
      return { row: i + 2, data: o };
    }
  }
  return null;
}

function findUser(name) {
  var r = findRow('Users', 'name', normName(name));
  return r ? r.data : null;
}

function appendRow(name, obj) {
  sheet(name).appendRow(SHEETS[name].map(function (c) { return obj[c] === undefined ? '' : String(obj[c]); }));
}

function updateRow(name, row, patch) {
  var sh = sheet(name), cols = SHEETS[name];
  Object.keys(patch).forEach(function (k) {
    var i = cols.indexOf(k);
    if (i >= 0) sh.getRange(row, i + 1).setValue(String(patch[k]));
  });
}

function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function normName(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
function clip(s, n) { return String(s || '').trim().slice(0, n); }
function cleanDate(s) { s = String(s || ''); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function tz() { return Session.getScriptTimeZone() || 'Asia/Jerusalem'; }
function nowStr() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd HH:mm'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
