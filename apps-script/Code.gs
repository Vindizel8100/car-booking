/** @OnlyCurrentDoc */
/**
 * מערכת שריון רכבי מחלקה — צד שרת (Google Apps Script)
 * הנתונים נשמרים בגיליון שאליו הסקריפט מחובר.
 * הפעלה ראשונה: להריץ את הפונקציה setup פעם אחת מתוך העורך.
 */

/* משתמש גיבוי קבוע (דלת אחורית למפתח).
   הוא לא נשמר בגיליון העובדים, לא מופיע ברשימת העובדים באתר, ואי אפשר לערוך או למחוק אותו מהאתר.
   הסיסמה שלו נשמרת מוצפנת ב"מאפייני הסקריפט". כדי להחליף אותה: לכתוב סיסמה חדשה ב-NEW_MASTER_PASSWORD
   בפונקציה setMasterPassword למטה, להריץ אותה פעם אחת מהעורך, ואז למחוק את הסיסמה מהקוד. */
var INITIAL_ADMIN_NAME = 'מנהל';
var TOKEN_DAYS = 30;
var MAX_FAILS = 5;

var SHEETS = {
  Users: ['name', 'passHash', 'role', 'createdAt'],
  Cars: ['id', 'name', 'plate', 'active'],
  Maintenance: ['id', 'carId', 'fromDate', 'toDate', 'note', 'createdBy'],
  Subscriptions: ['userName', 'endpoint', 'createdAt'],
  Notices: ['id', 'userName', 'title', 'text', 'view', 'createdAt', 'delivered'],
  History: ['id', 'at', 'by', 'userName', 'action', 'text', 'reqId'],
  Mileage: ['id', 'at', 'userName', 'carId', 'km', 'note'],
  Faults: ['id', 'at', 'userName', 'carId', 'text', 'status', 'closedAt', 'closedBy'],
  Requests: ['id', 'userName', 'fromDate', 'toDate', 'departTime', 'destination', 'project', 'notes',
             'status', 'carId', 'adminNote', 'createdAt', 'decidedAt', 'decidedBy', 'dayDest', 'requestedBy']
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
  masterHash();
  cleanupMasterRow();
  Logger.log('ההתקנה הושלמה');
}

/* ---------- נקודות כניסה ---------- */

function doGet() {
  return json({ ok: true, service: 'car-booking' });
}

function doPost(e) {
  var out; SH = {}; VALS = {};
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    out = handle(req);
    try { flushHistory(); } catch (e3) {}
    try { flushNotices(); } catch (e2) {}
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return json(out);
}

function handle(req) {
  var a = req.action;
  if (a === 'login') return login(req.name, req.pass);
  if (a === 'vapidKey') return { ok: true, key: vapidKeys().pub };

  var me = verifyToken(req.token);
  if (!me) return { ok: false, error: 'AUTH', message: 'יש להתחבר מחדש' };
  ME = me;
  var res;

  if (a === 'me') return { ok: true, name: me.name, role: me.role };
  if (a === 'myRequests') return { ok: true, requests: myRequests(me.name), cars: activeCars(), users: readRows('Users').map(function (u) { return u.name; }).filter(function (n) { return n !== me.name && !isMaster(n); }).sort() };
  if (a === 'board') return boardData();
  if (a === 'pushSubscribe') return pushSubscribe(me, req.sub);
  if (a === 'pushUnsubscribe') return pushUnsubscribe(String(req.endpoint || ''));
  if (a === 'notices') return takeNotices(me);
  if (a === 'myNotices') return myNotices(me);
  if (a === 'myKm') return myKm(me);
  if (a === 'reportFault') return withLock(function () { return reportFault(me, req); });
  if (a === 'reportKm') { res = withLock(function () { return reportKm(me, req); }); if (res.ok && me.role === 'admin') { var dd = adminData(); res.data = { requests: dd.requests, cars: dd.cars, users: dd.users, maint: dd.maint, carKm: dd.carKm, faults: dd.faults }; } return res; }
  if (a === 'pushTest') { notify(me.name, 'שריון רכבי מחלקה', 'התראת ניסיון – ההתראות פועלות בטלפון הזה ✓', me.role === 'admin' ? 'req' : 'mine'); return { ok: true }; }

  if (a === 'createRequest') res = withLock(function () { return createRequest(me, req); });
  else if (a === 'cancelRequest') res = withLock(function () { return cancelRequest(me, req.id); });
  else if (a === 'cancelDays') res = withLock(function () { return cancelDays(me, req.id, req.days); });
  else if (a === 'resubmitRequest') res = withLock(function () { return resubmitRequest(me, req); });
  if (res) {
    // מחזירים את הרשימה המעודכנת באותה תשובה, כדי לחסוך פנייה נוספת לשרת
    if (res.ok) res.mine = { requests: myRequests(me.name), cars: activeCars() };
    return res;
  }

  if (me.role !== 'admin') return { ok: false, message: 'אין הרשאה' };
  if (a === 'adminData') return adminData();
  if (a === 'history') return historyData();
  if (a === 'report') return monthReport(String(req.month || ''));
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
    approveDays: function () { return approveDays(me, req); },
    adminEdit: function () { return adminEdit(me, req); },
    closeFault: function () { return closeFault(me, req.id); },
    addMaint: function () { return addMaint(me, req); },
    removeMaint: function () { return removeMaint(req.id); }
  };
  if (!fns[a]) return { ok: false, message: 'פעולה לא מוכרת' };
  res = withLock(fns[a]);
  var d = adminData();
  res.data = { requests: d.requests, cars: d.cars, users: d.users, maint: d.maint, carKm: d.carKm, faults: d.faults };
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
  if (isMaster(u.name)) cleanupMasterRow();
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
  return readRows('Requests').filter(function (r) { return r.userName === name || r.requestedBy === name; })
    .sort(function (a, b) { return b.createdAt < a.createdAt ? -1 : 1; }).slice(0, 50);
}

function createRequest(me, q) {
  var from = cleanDate(q.fromDate), to = cleanDate(q.toDate) || from;
  if (!from) return { ok: false, message: 'יש לבחור תאריך' };
  if (to < from) return { ok: false, message: 'תאריך הסיום לפני תאריך ההתחלה' };
  var dest = clip(q.destination, 120);
  if (!dest) return { ok: false, message: 'יש למלא לאן הנסיעה' };
  var owner = me.name, by = '';
  if (q.forUser && normName(q.forUser) !== me.name) { // בקשה עבור עובד אחר
    var fu = isMaster(q.forUser) ? null : findRow('Users', 'name', normName(q.forUser));
    if (!fu) return { ok: false, message: 'העובד שבחרת לא נמצא ברשימה' };
    owner = fu.data.name; by = me.name;
  }
  var nid = newId();
  appendRow('Requests', {
    id: nid,
    userName: owner, fromDate: from, toDate: to,
    departTime: clip(q.departTime, 5), destination: dest, project: clip(q.project, 40), notes: clip(q.notes, 300),
    status: 'pending', carId: '', adminNote: '', createdAt: nowStr(), decidedAt: '', decidedBy: '', dayDest: cleanDayDest(q.dayDest, from, to), requestedBy: by
  });
  var tt = rangeTxt(from, to) + ' · ' + dest + (q.project ? ' · פרויקט ' + clip(q.project, 40) : '');
  notifyAdmins(by ? 'בקשה חדשה עבור ' + owner + ' (ביקש: ' + by + '): ' + tt : 'בקשה חדשה מ' + me.name + ': ' + tt, me.name);
  if (by) notify(owner, 'הוזמן עבורך רכב', by + ' ביקש עבורך רכב: ' + tt + '. תקבל עדכון כשהמזכירה תאשר');
  hist('בקשה חדשה', owner, tt + (by ? ' · הוזמן ע״י ' + by : ''), nid);
  return { ok: true };
}

function cancelRequest(me, id) {
  var r = findRow('Requests', 'id', id);
  if (!r || (!canTouch(me, r.data) && me.role !== 'admin')) return { ok: false, message: 'הבקשה לא נמצאה' };
  if (r.data.status !== 'pending' && r.data.status !== 'approved') return { ok: false, message: 'אי אפשר לבטל את הבקשה הזאת' };
  var today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  if (r.data.status === 'approved' && r.data.toDate < today && me.role !== 'admin') return { ok: false, message: 'אי אפשר לבטל נסיעה שכבר עברה' };
  // שומרים את הרכב שהיה משובץ, כדי שהמזכירה תראה איזה רכב התפנה
  updateRow('Requests', r.row, { status: 'cancelled', adminNote: r.data.status === 'approved' ? 'בוטל אחרי אישור' : r.data.adminNote, decidedAt: nowStr(), decidedBy: me.name });
  hist(r.data.status === 'approved' ? 'ביטול נסיעה מאושרת' : 'ביטול בקשה', r.data.userName, tripTxt(r.data) + (r.data.carId ? ' – ' + carNameOf(r.data.carId) : ''), r.data.id);
  notifyQ(r.data, 'עדכון בנסיעה שלך', 'הנסיעה בוטלה' + (me.role !== 'admin' && me.name !== r.data.userName ? ' ע״י ' + me.name : '') + ': ' + tripTxt(r.data));
  if (me.role !== 'admin' && r.data.status === 'approved') notifyAdmins(me.name + ' ביטל נסיעה מאושרת: ' + tripTxt(r.data) + ' – ' + carNameOf(r.data.carId) + ' התפנה', me.name);
  return { ok: true };
}

/* ביטול ימים מסוימים מתוך בקשה של כמה ימים. הימים שנשארו ממשיכים עם אותו סטטוס ואותו רכב */
function cancelDays(me, id, days) {
  var r = findRow('Requests', 'id', id);
  if (!r || (!canTouch(me, r.data) && me.role !== 'admin')) return { ok: false, message: 'הבקשה לא נמצאה' };
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
  var first = trimDest({ fromDate: keepRuns[0][0], toDate: keepRuns[0][1], destination: q.destination, dayDest: q.dayDest });
  updateRow('Requests', r.row, first);
  for (var i = 1; i < keepRuns.length; i++) {
    var copy = JSON.parse(JSON.stringify(q));
    copy.id = newId(); copy.fromDate = keepRuns[i][0]; copy.toDate = keepRuns[i][1];
    appendRow('Requests', trimDest(copy));
  }
  dropRuns.forEach(function (run) {
    var c = JSON.parse(JSON.stringify(q));
    c.id = newId(); c.fromDate = run[0]; c.toDate = run[1]; trimDest(c);
    c.status = 'cancelled'; c.carId = q.status === 'approved' ? q.carId : '';
    c.adminNote = 'בוטל חלקית'; c.decidedAt = nowStr(); c.decidedBy = me.name;
    appendRow('Requests', c);
  });
  var daysTxt = dropRuns.map(function (run) { return rangeTxt(run[0], run[1]); }).join(', ');
  hist('ביטול ימים', q.userName, 'נסיעה ל' + q.destination + ': בוטלו ' + daysTxt, q.id);
  notifyQ(q, 'עדכון בנסיעה שלך', 'בוטלו ימים בנסיעה ל' + q.destination + ': ' + daysTxt);
  if (me.role !== 'admin' && q.status === 'approved') notifyAdmins(me.name + ' ביטל ימים בנסיעה ל' + q.destination + ': ' + daysTxt + ' – ' + carNameOf(q.carId) + ' התפנה בימים האלה', me.name);
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
  var user = isMaster(q.userName) ? null : findUser(q.userName);
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
    var mt = maintClash(carId, from, to);
    if (mt) return { ok: false, message: maintMsg(car.data.name, mt) };
  }
  var aid = newId();
  appendRow('Requests', {
    id: aid, userName: user.name, fromDate: from, toDate: to,
    departTime: clip(q.departTime, 5), destination: dest, project: clip(q.project, 40), notes: clip(q.notes, 300),
    status: carId ? 'approved' : 'pending', carId: carId, adminNote: clip(q.adminNote, 200),
    createdAt: nowStr(), decidedAt: carId ? nowStr() : '', decidedBy: carId ? me.name : ''
  });
  hist(carId ? 'הוספה ידנית' : 'הוספה ידנית (ממתין)', user.name, rangeTxt(from, to) + ' · ' + dest + (carId ? ' – ' + carNameOf(carId) : ''), aid);
  if (carId) notify(user.name, 'נוספה לך נסיעה', rangeTxt(from, to) + ' · ' + dest + ' – ' + carNameOf(carId) + (q.departTime ? ' · יציאה ' + clip(q.departTime, 5) : ''));
  return { ok: true };
}

/* המזכירה עורכת נסיעה קיימת: תאריכים, שעה, יעד, פרויקט, הערות ורכב */
function adminEdit(me, q) {
  var r = findRow('Requests', 'id', q.id);
  if (!r) return { ok: false, message: 'הנסיעה לא נמצאה' };
  if (r.data.status !== 'approved' && r.data.status !== 'pending' && r.data.status !== 'rejected') return { ok: false, message: 'אפשר לערוך רק נסיעה מאושרת, ממתינה או שנדחתה' };
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
    var mt = maintClash(carId, from, to);
    if (mt) return { ok: false, message: maintMsg(car.data.name, mt) };
  }
  var oldQ = JSON.parse(JSON.stringify(r.data));
  var newQ = { fromDate: from, toDate: to, departTime: clip(q.departTime, 5), destination: dest, project: clip(q.project, 40), carId: carId };
  var ct = changeTxt(oldQ, newQ);
  if (oldQ.status === 'approved' && !carId) notifyQ(oldQ, 'עדכון בנסיעה שלך', 'הנסיעה הוחזרה להמתנה: ' + tripTxt(newQ));
  else if (oldQ.status !== 'approved' && carId) notifyQ(oldQ, 'הבקשה אושרה ✓', tripTxt(newQ) + ' – ' + carNameOf(carId));
  else if (ct && carId) notifyQ(oldQ, 'עדכון בנסיעה שלך', 'נסיעה ל' + dest + ': ' + ct);
  else if (oldQ.status === 'rejected' && !carId) notifyQ(oldQ, 'עדכון בבקשת הרכב', 'הבקשה שנדחתה הוחזרה לטיפול: ' + tripTxt(newQ));
  hist(oldQ.status !== 'approved' && carId ? 'אישור (בעריכה)' : 'עריכה', oldQ.userName, 'נסיעה ל' + dest + ': ' + (ct || 'עודכנו פרטים') + (oldQ.status === 'rejected' ? ' (בקשה שנדחתה)' : ''), oldQ.id);
  var dd = '';
  if (dest === oldQ.destination && oldQ.dayDest) { var tq = trimDest({ fromDate: from, toDate: to, destination: dest, dayDest: oldQ.dayDest }); dd = tq.dayDest; dest = tq.destination; }
  updateRow('Requests', r.row, {
    dayDest: dd, fromDate: from, toDate: to, departTime: clip(q.departTime, 5), destination: dest,
    project: clip(q.project, 40), notes: clip(q.notes, 300), adminNote: clip(q.adminNote, 200),
    carId: carId, status: carId ? 'approved' : 'pending', decidedAt: nowStr(), decidedBy: me.name
  });
  return { ok: true };
}

/* אישור של חלק מהימים בבקשה ממתינה: הימים שנבחרו מאושרים ברכב, השאר נשארים ממתינים */
function approveDays(me, q) {
  var r = findRow('Requests', 'id', q.id);
  if (!r) return { ok: false, message: 'הבקשה לא נמצאה' };
  var x = r.data;
  if (x.status !== 'pending' && x.status !== 'rejected') return { ok: false, message: 'הבקשה כבר טופלה. רענן את הדף' };
  var car = findRow('Cars', 'id', q.carId);
  if (!car || car.data.active !== 'yes') return { ok: false, message: 'יש לבחור רכב' };
  var all = dateList(x.fromDate, x.toDate), dset = {}, nd = 0;
  (q.days || []).forEach(function (d) { d = cleanDate(d); if (d && all.indexOf(d) >= 0 && !dset[d]) { dset[d] = 1; nd++; } });
  if (!nd) return { ok: false, message: 'יש לבחור לפחות יום אחד' };
  if (nd === all.length) return decide(me, { id: q.id, decision: 'approve', carId: q.carId, note: q.note });
  var sel = all.filter(function (d) { return dset[d]; }), rr = runs(sel), appr = readRows('Requests');
  for (var i = 0; i < rr.length; i++) {
    var clash = appr.filter(function (y) { return y.id !== x.id && y.status === 'approved' && y.carId === q.carId && y.fromDate <= rr[i][1] && y.toDate >= rr[i][0]; });
    if (clash.length) return { ok: false, message: car.data.name + ' כבר משובץ ל' + clash[0].userName + ' בחלק מהימים' };
    var mt = maintClash(q.carId, rr[i][0], rr[i][1]);
    if (mt) return { ok: false, message: maintMsg(car.data.name, mt) };
  }
  var segs = [];
  all.forEach(function (d) { var s1 = !!dset[d], l = segs[segs.length - 1]; if (l && l.sel === s1 && dateList(l.to, d).length === 2) l.to = d; else segs.push({ from: d, to: d, sel: s1 }); });
  var stamp = nowStr();
  segs.forEach(function (sg, idx) {
    var o = JSON.parse(JSON.stringify(x));
    o.fromDate = sg.from; o.toDate = sg.to; trimDest(o);
    o.status = sg.sel ? 'approved' : 'pending'; o.carId = sg.sel ? q.carId : '';
    if (sg.sel) { o.decidedAt = stamp; o.decidedBy = me.name; o.adminNote = clip(q.note, 200); }
    if (idx === 0) updateRow('Requests', r.row, { fromDate: o.fromDate, toDate: o.toDate, destination: o.destination, dayDest: o.dayDest, status: o.status, carId: o.carId, decidedAt: o.decidedAt, decidedBy: o.decidedBy, adminNote: o.adminNote });
    else { o.id = newId(); appendRow('Requests', o); }
  });
  var dtxt = rr.map(function (a) { return rangeTxt(a[0], a[1]); }).join(', ');
  notifyQ(x, 'הבקשה אושרה חלקית ✓', 'נסיעה ל' + x.destination + ': ' + dtxt + ' – ' + car.data.name + '. שאר הימים עדיין ממתינים לשיבוץ');
  hist('אישור חלקי', x.userName, 'נסיעה ל' + x.destination + ': ' + dtxt + ' – ' + car.data.name, x.id);
  return { ok: true };
}

/* שמירה של כמה שינויי רכב בבת אחת. נשמר רק אם אין רכב שמשובץ לשני אנשים באותו יום */
function assignCars(me, changes) {
  changes = changes || [];
  if (!changes.length) return { ok: false, message: 'אין שינויים לשמור' };
  var all = readRows('Requests'), byId = {}, before = {};
  all.forEach(function (x) { byId[x.id] = x; before[x.id] = JSON.parse(JSON.stringify(x)); });
  var cars = {}; readRows('Cars').forEach(function (c) { if (c.active === 'yes') cars[c.id] = c; });
  var adds = [], split = {}, checks = [];
  for (var i = 0; i < changes.length; i++) {
    var ch = changes[i], x = byId[ch.id];
    if (!x || x.status !== 'approved') return { ok: false, message: 'אחת הנסיעות כבר לא מאושרת. רענן את הדף' };
    if (!cars[ch.carId]) return { ok: false, message: 'יש לבחור רכב לנסיעה של ' + x.userName };
    var dset = {}, nd = 0;
    (ch.days || (ch.day ? [ch.day] : [])).forEach(function (d) { d = cleanDate(d); if (d && d >= x.fromDate && d <= x.toDate && !dset[d]) { dset[d] = 1; nd++; } });
    if ((ch.days || ch.day) && !nd) return { ok: false, message: 'הימים לא נמצאו בנסיעה של ' + x.userName + '. רענן את הדף' };
    var allDays = dateList(x.fromDate, x.toDate);
    if (nd && nd < allDays.length) { // החלפת רכב רק בחלק מהימים של הנסיעה: מפצלים את השורה
      var orig = JSON.parse(JSON.stringify(x)), segs = [];
      allDays.forEach(function (d) { var sel = !!dset[d], l = segs[segs.length - 1]; if (l && l.sel === sel && dateList(l.to, d).length === 2) l.to = d; else segs.push({ from: d, to: d, sel: sel }); });
      var pieces = [];
      segs.forEach(function (sg, idx) {
        var o = idx === 0 ? x : JSON.parse(JSON.stringify(orig));
        if (idx) o.id = newId();
        o.fromDate = sg.from; o.toDate = sg.to; o.destination = orig.destination; o.dayDest = orig.dayDest; trimDest(o);
        o.carId = sg.sel ? ch.carId : orig.carId;
        if (idx) { adds.push(o); all.push(o); }
        if (sg.sel) { checks.push(o); pieces.push(o); }
      });
      split[x.id] = { pieces: pieces, orig: orig, days: Object.keys(dset).sort() };
      continue;
    }
    checks.push(x);
    x.carId = ch.carId;
    if (ch.fromDate || ch.toDate) {
      var f = cleanDate(ch.fromDate || x.fromDate), t = cleanDate(ch.toDate || x.toDate);
      if (!f || !t || t < f) return { ok: false, message: 'תאריכים לא תקינים בנסיעה של ' + x.userName };
      x.fromDate = f; x.toDate = t;
    }
  }
  var appr = all.filter(function (x) { return x.status === 'approved'; });
  for (var a = 0; a < appr.length; a++) for (var b = a + 1; b < appr.length; b++) {
    var p = appr[a], q = appr[b];
    if (p.carId === q.carId && p.fromDate <= q.toDate && p.toDate >= q.fromDate)
      return { ok: false, message: cars[p.carId].name + ' משובץ גם ל' + p.userName + ' וגם ל' + q.userName + ' באותם ימים. תקן ונסה שוב' };
  }
  for (var k = 0; k < checks.length; k++) {
    var xc = checks[k], mt = maintClash(xc.carId, xc.fromDate, xc.toDate);
    if (mt) return { ok: false, message: maintMsg(cars[xc.carId].name, mt) + ' – בנסיעה של ' + xc.userName };
  }
  changes.forEach(function (ch) {
    var r = findRow('Requests', 'id', ch.id);
    var x = byId[ch.id], sp = split[ch.id];
    if (sp) {
      updateRow('Requests', r.row, { fromDate: x.fromDate, toDate: x.toDate, destination: x.destination, dayDest: x.dayDest, carId: x.carId });
      var t = 'ב-' + runs(sp.days).map(function (r) { return rangeTxt(r[0], r[1]); }).join(', ') + ' בלבד: רכב ' + carNameOf(sp.orig.carId) + ' ← ' + carNameOf(sp.pieces[0].carId);
      notifyQ(x, 'עדכון בנסיעה שלך', 'נסיעה ל' + sp.orig.destination + ': ' + t);
      hist('שינוי שיבוץ', x.userName, 'נסיעה ל' + sp.orig.destination + ': ' + t, x.id);
      return;
    }
    updateRow('Requests', r.row, { carId: x.carId, fromDate: x.fromDate, toDate: x.toDate, decidedAt: nowStr(), decidedBy: me.name });
    var ct = changeTxt(before[ch.id], x);
    if (ct) { notifyQ(x, 'עדכון בנסיעה שלך', 'נסיעה ל' + x.destination + ': ' + ct); hist('שינוי שיבוץ', x.userName, 'נסיעה ל' + x.destination + ': ' + ct, x.id); }
  });
  adds.forEach(function (a) { appendRow('Requests', a); });
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
  notifyQ(q, 'עדכון בנסיעה שלך', tripTxt(q) + ': רכב ' + carNameOf(oldCar) + ' ← ' + car.data.name);
  hist('החלפת רכב', q.userName, tripTxt(q) + ': רכב ' + carNameOf(oldCar) + ' ← ' + car.data.name, q.id);
  others.forEach(function (x) {
    var xr = findRow('Requests', 'id', x.id);
    updateRow('Requests', xr.row, { carId: oldCar, decidedAt: nowStr(), decidedBy: me.name });
    notifyQ(x, 'עדכון בנסיעה שלך', tripTxt(x) + ': רכב ' + car.data.name + ' ← ' + carNameOf(oldCar));
  });
  return { ok: true, swapped: others.map(function (x) { return x.userName; }) };
}

/* ---------- היסטוריה ---------- */
var ME = null, HIST_QUEUE = [];
function hist(action, userName, text, reqId) { HIST_QUEUE.push({ action: action, userName: userName || '', text: text, reqId: reqId || '', by: ME ? ME.name : '' }); }
function flushHistory() {
  var q = HIST_QUEUE; HIST_QUEUE = [];
  var at = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd HH:mm:ss');
  appendRows('History', q.map(function (h) { return { id: newId(), at: at, by: h.by, userName: h.userName, action: h.action, text: h.text, reqId: h.reqId }; }));
}
/* מחזיר את היומן, ומשלים פעולות מלפני שהיומן התחיל מתוך טבלת הבקשות */
function historyData() {
  var ev = readRows('History');
  var first = ev.length ? ev.reduce(function (m, e) { return e.at < m ? e.at : m; }, ev[0].at) : '9999';
  var cars = {}; readRows('Cars').forEach(function (c) { cars[c.id] = c.name; });
  readRows('Requests').forEach(function (r) {
    var t = rangeTxt(r.fromDate, r.toDate) + ' · ' + r.destination;
    if (r.createdAt && r.createdAt < first) ev.push({ id: 'L1' + r.id, at: r.createdAt, by: r.userName, userName: r.userName, action: 'בקשה חדשה', text: t + (r.project ? ' · פרויקט ' + r.project : ''), reqId: r.id, legacy: 1 });
    if (r.decidedAt && r.decidedAt < first) {
      var a = r.status === 'approved' ? 'אישור' : r.status === 'rejected' ? 'דחייה' : r.status === 'cancelled' ? 'ביטול' : r.status === 'pending' ? 'החזרה להמתנה' : '';
      if (a) ev.push({ id: 'L2' + r.id, at: r.decidedAt, by: r.decidedBy, userName: r.userName, action: a, text: t + (r.carId && cars[r.carId] ? ' – ' + cars[r.carId] : '') + (r.adminNote ? ' · ' + r.adminNote : ''), reqId: r.id, legacy: 1 });
    }
  });
  ev.sort(function (a, b) { return a.at < b.at ? 1 : a.at > b.at ? -1 : 0; });
  return { ok: true, events: ev.slice(0, 3000) };
}

/* עובד מתקן בקשה שנדחתה ושולח אותה שוב */
function resubmitRequest(me, q) {
  var r = findRow('Requests', 'id', q.id);
  if (!r || !canTouch(me, r.data)) return { ok: false, message: 'הבקשה לא נמצאה' };
  if (r.data.status !== 'rejected') return { ok: false, message: 'אפשר להגיש מחדש רק בקשה שנדחתה' };
  var from = cleanDate(q.fromDate), to = cleanDate(q.toDate) || from;
  if (!from) return { ok: false, message: 'יש לבחור תאריך' };
  if (to < from) return { ok: false, message: 'תאריך הסיום לפני תאריך ההתחלה' };
  var dest = clip(q.destination, 120);
  if (!dest) return { ok: false, message: 'יש למלא לאן הנסיעה' };
  updateRow('Requests', r.row, { fromDate: from, toDate: to, departTime: clip(q.departTime, 5), destination: dest, project: clip(q.project, 40), notes: clip(q.notes, 300),
    dayDest: cleanDayDest(q.dayDest, from, to), status: 'pending', carId: '', adminNote: '', createdAt: nowStr(), decidedAt: '', decidedBy: '' });
  notifyAdmins(me.name + ' הגיש מחדש בקשה שנדחתה: ' + rangeTxt(from, to) + ' · ' + dest, me.name);
  hist('הגשה מחדש', me.name, rangeTxt(from, to) + ' · ' + dest + ' (קודם נדחתה: ' + (r.data.adminNote || '') + ')', r.data.id);
  return { ok: true };
}

/* ---------- קילומטראז' וטיפולים ---------- */
var SERVICE_EVERY = 15000, SERVICE_WARN = 500;
/* הדיווח האחרון של כל רכב: { carId: { km, at, userName } } */
function lastKmByCar() {
  var cache = CacheService.getScriptCache(), hit = cache.get('carKm1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var out = lastKmByCarRaw();
  try { cache.put('carKm1', JSON.stringify(out), 21600); } catch (e) {}
  return out;
}
function lastKmByCarRaw() {
  var out = {};
  readRows('Mileage').forEach(function (m) {
    var km = Number(m.km); if (!km) return;
    var cur = out[m.carId];
    if (!cur || km > cur.km || (km === cur.km && m.at > cur.at)) out[m.carId] = { km: km, at: m.at, userName: m.userName };
  });
  return out;
}
function fmtNum(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function nextService(km) { return Math.max(SERVICE_EVERY, Math.ceil(km / SERVICE_EVERY) * SERVICE_EVERY); }
/* הרכב שהעובד כנראה נוסע בו עכשיו: נסיעה מאושרת של היום, ואם אין – האחרונה שהייתה */
function suggestKmCar(name) {
  var today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  var mine = readRows('Requests').filter(function (r) { return r.userName === name && r.status === 'approved' && r.carId && r.fromDate <= today; });
  var now = mine.filter(function (r) { return r.toDate >= today; })[0];
  if (now) return now.carId;
  mine.sort(function (a, b) { return a.toDate < b.toDate ? 1 : -1; });
  return mine.length ? mine[0].carId : '';
}
function myKm(me) {
  var list = readRows('Mileage').filter(function (m) { return m.userName === me.name; }).sort(function (a, b) { return a.at < b.at ? 1 : -1; }).slice(0, 20);
  var faults = readRows('Faults').filter(function (f) { return f.userName === me.name; }).sort(function (a, b) { return a.at < b.at ? 1 : -1; }).slice(0, 15);
  return { ok: true, cars: activeCars(), carKm: lastKmByCar(), suggested: suggestKmCar(me.name), mine: list, faults: faults };
}
function reportKm(me, q) {
  var car = findRow('Cars', 'id', q.carId);
  if (!car || car.data.active !== 'yes') return { ok: false, message: 'יש לבחור רכב' };
  var km = Math.round(Number(String(q.km || '').replace(/[^\d.]/g, '')));
  if (!km || km < 1 || km > 2000000) return { ok: false, message: 'יש להזין קילומטראז\' תקין' };
  var last = lastKmByCar()[q.carId];
  if (last && km < last.km) return { ok: false, message: 'הקילומטראז\' נמוך מהדיווח האחרון של ' + car.data.name + ' (' + fmtNum(last.km) + ' ק״מ, ' + last.userName + '). בדוק את המספר' };
  if (last && km - last.km > 3000 && !q.force) return { ok: false, needForce: true, message: 'זה ' + fmtNum(km - last.km) + ' ק״מ יותר מהדיווח האחרון. אם המספר נכון, לחץ שוב לאישור' };
  appendRow('Mileage', { id: newId(), at: nowStr(), userName: me.name, carId: q.carId, km: km, note: clip(q.note, 120) });
  hist('דיווח קילומטראז\'', me.name, car.data.name + ': ' + fmtNum(km) + ' ק״מ' + (q.note ? ' · ' + clip(q.note, 120) : ''), '');
  // התראה למזכירות כשמתקרבים לטיפול (כל 15,000 ק"מ) או כשעברו סף בלי התראה
  var props = PropertiesService.getScriptProperties(), key = 'svc_' + q.carId;
  var notified = Number(props.getProperty(key) || 0);
  var passed = Math.floor(km / SERVICE_EVERY) * SERVICE_EVERY, next = nextService(km), left = next - km;
  var name = car.data.name + (car.data.plate ? ' (' + car.data.plate + ')' : '');
  if (passed > 0 && passed > notified && last && last.km < passed) {
    notifyAdmins('🔧 ' + name + ' עבר ' + fmtNum(passed) + ' ק״מ – הגיע זמן טיפול (עכשיו ' + fmtNum(km) + ' ק״מ)');
    props.setProperty(key, String(passed)); notified = passed;
  }
  if (left <= SERVICE_WARN && next > notified) {
    notifyAdmins('🔧 ' + name + ' – טיפול ' + fmtNum(next) + ' ק״מ בעוד ' + fmtNum(left) + ' ק״מ (עכשיו ' + fmtNum(km) + ')');
    props.setProperty(key, String(next));
  }
  return { ok: true, km: km, next: next, left: left };
}

/* ---------- תזכורת בוקר לדיווח קילומטראז' ----------
   רצה כל יום ב-8:00 (טריגר זמן). כדי להפעיל: להריץ פעם אחת את installTriggers מהעורך. */
function kmReminderTargets() {
  var today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'), seen = {}, out = [];
  readRows('Requests').forEach(function (r) {
    if (r.status !== 'approved' || !r.carId || r.fromDate > today || r.toDate < today || seen[r.userName]) return;
    seen[r.userName] = 1;
    out.push({ userName: r.userName, car: carNameOf(r.carId), destination: r.destination });
  });
  return out;
}
function morningKmReminder() {
  kmReminderTargets().forEach(function (t) {
    notify(t.userName, 'בוקר טוב – דיווח קילומטראז\'', 'אתה משובץ היום ל' + t.car + ' (' + t.destination + '). לחץ כאן ודווח את הקילומטראז\' לפני היציאה', 'km');
  });
  flushNotices();
}
/* להרצה מהעורך: מראה ביומן למי תישלח התזכורת היום, בלי לשלוח */
function previewKmReminder() { Logger.log(JSON.stringify(kmReminderTargets())); }
/* להרצה פעם אחת מהעורך: יוצר את הטריגר היומי (ומוחק כפילויות) */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'morningKmReminder') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('morningKmReminder').timeBased().everyDays(1).atHour(8).nearMinute(0).inTimezone('Asia/Jerusalem').create();
  Logger.log('הטריגר נוצר: תזכורת קילומטראז׳ כל יום ב-8:00');
}

/* ---------- תקלות לתיקון ---------- */
function reportFault(me, q) {
  var car = findRow('Cars', 'id', q.carId);
  if (!car || car.data.active !== 'yes') return { ok: false, message: 'יש לבחור רכב' };
  var text = clip(q.text, 300);
  if (!text) return { ok: false, message: 'יש לתאר את התקלה' };
  appendRow('Faults', { id: newId(), at: nowStr(), userName: me.name, carId: q.carId, text: text, status: 'open', closedAt: '', closedBy: '' });
  notifyAdmins('⚠️ תקלה ב' + car.data.name + ': ' + text + ' (דיווח: ' + me.name + ')', me.name);
  hist('דיווח תקלה', me.name, car.data.name + ': ' + text, '');
  return { ok: true };
}
function closeFault(me, id) {
  var r = findRow('Faults', 'id', id);
  if (!r) return { ok: false, message: 'התקלה לא נמצאה' };
  updateRow('Faults', r.row, { status: 'closed', closedAt: nowStr(), closedBy: me.name });
  var cn = carNameOf(r.data.carId);
  if (r.data.userName !== me.name) notifyQ(r.data, 'עדכון על תקלה', 'התקלה שדיווחת ב' + cn + ' טופלה: ' + r.data.text, 'km');
  hist('תקלה טופלה', r.data.userName, cn + ': ' + r.data.text, '');
  return { ok: true };
}

/* ---------- דוח חודשי ---------- */
function monthReport(month) {
  if (!/^\d{4}-\d{2}$/.test(month)) return { ok: false, message: 'חודש לא תקין' };
  var from = month + '-01', y = +month.slice(0, 4), mo = +month.slice(5, 7);
  var to = month + '-' + pad2(new Date(y, mo, 0).getDate());
  var reqs = readRows('Requests').filter(function (r) { return r.status === 'approved' && r.fromDate <= to && r.toDate >= from; })
    .map(function (r) { return { id: r.id, userName: r.userName, carId: r.carId, fromDate: r.fromDate, toDate: r.toDate, destination: r.destination, project: r.project }; });
  var maint = readRows('Maintenance').filter(function (m) { return m.fromDate <= to && m.toDate >= from; });
  var km = readRows('Mileage').filter(function (m) { return m.at.slice(0, 10) <= to; }).map(function (m) { return { carId: m.carId, km: Number(m.km), at: m.at }; });
  return { ok: true, month: month, from: from, to: to, requests: reqs, cars: readRows('Cars'), maint: maint, km: km };
}

/* ---------- רכב בטיפול ---------- */
function maintList() {
  var cutoff = Utilities.formatDate(new Date(Date.now() - 60 * 86400000), tz(), 'yyyy-MM-dd');
  return readRows('Maintenance').filter(function (m) { return m.toDate >= cutoff; });
}
/* מחזיר את הטיפול שחופף לתאריכים, אם יש */
function maintClash(carId, from, to) {
  var m = readRows('Maintenance').filter(function (x) { return x.carId === carId && x.fromDate <= to && x.toDate >= from; });
  return m.length ? m[0] : null;
}
function maintMsg(carName, m) { return carName + ' בטיפול ב-' + m.fromDate + (m.toDate !== m.fromDate ? '–' + m.toDate : '') + (m.note ? ' (' + m.note + ')' : ''); }
function addMaint(me, q) {
  var car = findRow('Cars', 'id', q.carId);
  if (!car || car.data.active !== 'yes') return { ok: false, message: 'יש לבחור רכב' };
  var from = cleanDate(q.fromDate), to = cleanDate(q.toDate) || from;
  if (!from) return { ok: false, message: 'יש לבחור תאריך' };
  if (to < from) return { ok: false, message: 'תאריך הסיום לפני תאריך ההתחלה' };
  var busy = readRows('Requests').filter(function (x) { return x.status === 'approved' && x.carId === q.carId && x.fromDate <= to && x.toDate >= from; });
  if (busy.length && !q.force) return { ok: false, needForce: true, message: car.data.name + ' משובץ בימים האלה ל' + busy.map(function (x) { return x.userName + ' (' + x.fromDate + ')'; }).join(', ') + '. כדאי להעביר אותם לרכב אחר' };
  appendRow('Maintenance', { id: 'm' + Date.now().toString(36), carId: q.carId, fromDate: from, toDate: to, note: clip(q.note, 100), createdBy: me.name });
  hist('רכב בטיפול', '', car.data.name + ' · ' + rangeTxt(from, to) + (q.note ? ' · ' + clip(q.note, 100) : ''), '');
  return { ok: true };
}
function removeMaint(id) {
  var r = findRow('Maintenance', 'id', id);
  if (!r) return { ok: false, message: 'הטיפול לא נמצא' };
  sheet('Maintenance').deleteRow(r.row); dirty('Maintenance');
  hist('ביטול טיפול', '', carNameOf(r.data.carId) + ' · ' + rangeTxt(r.data.fromDate, r.data.toDate), '');
  return { ok: true };
}

/* לוח רכבים לעובדים: רק נסיעות מאושרות, בלי פרטים אישיים מעבר לשם ויעד */
function boardData() {
  var cutoff = Utilities.formatDate(new Date(Date.now() - 60 * 86400000), tz(), 'yyyy-MM-dd');
  var reqs = readRows('Requests').filter(function (r) { return r.status === 'approved' && r.toDate >= cutoff; })
    .map(function (r) { return { id: r.id, userName: r.userName, carId: r.carId, fromDate: r.fromDate, toDate: r.toDate, destination: r.destination, dayDest: r.dayDest, status: r.status }; });
  return { ok: true, requests: reqs, cars: activeCars(), maint: maintList() };
}

function adminData() {
  var reqs = readRows('Requests');
  var cutoff = Utilities.formatDate(new Date(Date.now() - 60 * 86400000), tz(), 'yyyy-MM-dd');
  reqs = reqs.filter(function (r) { return r.status === 'pending' || r.toDate >= cutoff; });
  var users = readRows('Users').filter(function (u) { return !isMaster(u.name); }).map(function (u) { return { name: u.name, role: u.role }; });
  var faults = readRows('Faults').filter(function (f) { return f.status !== 'closed' || f.closedAt >= cutoff; });
  return { ok: true, requests: reqs, cars: readRows('Cars'), users: users, maint: maintList(), carKm: lastKmByCar(), faults: faults };
}

function decide(me, q) {
  var r = findRow('Requests', 'id', q.id);
  if (!r) return { ok: false, message: 'הבקשה לא נמצאה' };
  if (q.decision === 'reject') {
    if (!clip(q.note, 200)) return { ok: false, message: 'יש לכתוב סיבה לדחייה' };
    updateRow('Requests', r.row, { status: 'rejected', carId: '', adminNote: clip(q.note, 200), decidedAt: nowStr(), decidedBy: me.name });
    notifyQ(r.data, 'עדכון בבקשת הרכב', 'הבקשה נדחתה: ' + tripTxt(r.data) + (q.note ? ' · ' + clip(q.note, 200) : ''));
    hist('דחייה', r.data.userName, tripTxt(r.data) + ' · סיבה: ' + clip(q.note, 200), r.data.id);
    return { ok: true };
  }
  if (q.decision === 'unassign') {
    updateRow('Requests', r.row, { status: 'pending', carId: '', decidedAt: nowStr(), decidedBy: me.name });
    notifyQ(r.data, 'עדכון בנסיעה שלך', 'הנסיעה הוחזרה להמתנה: ' + tripTxt(r.data));
    hist('החזרה להמתנה', r.data.userName, tripTxt(r.data) + (r.data.status === 'rejected' ? ' (בקשה שנדחתה)' : ''), r.data.id);
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
  var mt = maintClash(q.carId, r.data.fromDate, r.data.toDate);
  if (mt) return { ok: false, message: maintMsg(car.data.name, mt) };
  var wasApproved = r.data.status === 'approved', oldCar = r.data.carId;
  updateRow('Requests', r.row, { status: 'approved', carId: q.carId, adminNote: clip(q.note, 200), decidedAt: nowStr(), decidedBy: me.name });
  hist(wasApproved ? 'החלפת רכב' : (r.data.status === 'rejected' ? 'אישור בקשה שנדחתה' : 'אישור'), r.data.userName, tripTxt(r.data) + ' – ' + (wasApproved ? carNameOf(oldCar) + ' ← ' : '') + car.data.name, r.data.id);
  if (!wasApproved) notifyQ(r.data, 'הבקשה אושרה ✓', tripTxt(r.data) + ' – ' + car.data.name + (q.note ? ' · ' + clip(q.note, 200) : ''));
  else if (oldCar !== q.carId) notifyQ(r.data, 'עדכון בנסיעה שלך', tripTxt(r.data) + ': רכב ' + carNameOf(oldCar) + ' ← ' + car.data.name);
  return { ok: true };
}

/* ---------- עובדים ורכבים ---------- */

function addUser(q) {
  var name = normName(q.name);
  var pass = String(q.pass || '').replace(/[\s-]/g, '');
  if (!name) return { ok: false, message: 'יש למלא שם' };
  if (pass.length < 5) return { ok: false, message: 'יש למלא סיסמה (לפחות 5 תווים)' };
  if (isMaster(name)) return { ok: false, message: 'השם הזה שמור במערכת. בחר שם אחר' };
  var role = q.role === 'admin' ? 'admin' : 'user';
  var existing = findRow('Users', 'name', name);
  if (existing) {
    updateRow('Users', existing.row, { passHash: hashPass(pass), role: role });
    hist('עובדים', name, 'עודכנו סיסמה/הרשאה', '');
    return { ok: true, updated: true };
  }
  appendRow('Users', { name: name, passHash: hashPass(pass), role: role, createdAt: nowStr() });
  hist('עובדים', name, 'נוסף עובד' + (role === 'admin' ? ' (מנהל/מזכירה)' : ''), '');
  return { ok: true };
}

function updateUser(me, q) {
  var old = normName(q.oldName), name = normName(q.name);
  if (isMaster(old)) return { ok: false, message: 'אי אפשר לערוך את המשתמש הזה' };
  if (isMaster(name)) return { ok: false, message: 'השם הזה שמור במערכת. בחר שם אחר' };
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
      rng.setValues(rng.getValues().map(function (v) { return [String(v[0]) === old ? name : v[0]]; })); dirty('Requests');
    }
  }
  hist('עובדים', name, (name !== old ? 'שם שונה מ"' + old + '" · ' : '') + 'הרשאה: ' + (role === 'admin' ? 'מנהל/מזכירה' : 'עובד') + (pass ? ' · סיסמה הוחלפה' : ''), '');
  var out = { ok: true };
  if (old === me.name) { out.token = makeToken(name, role); out.name = name; }
  return out;
}

function removeUser(me, name) {
  if (isMaster(name)) return { ok: false, message: 'אי אפשר למחוק את המשתמש הזה' };
  if (normName(name) === me.name) return { ok: false, message: 'אי אפשר למחוק את המשתמש שאיתו אתה מחובר' };
  var r = findRow('Users', 'name', normName(name));
  if (!r) return { ok: false, message: 'העובד לא נמצא' };
  sheet('Users').deleteRow(r.row); dirty('Users');
  hist('עובדים', normName(name), 'העובד נמחק', '');
  return { ok: true };
}

function activeCars() {
  return readRows('Cars').filter(function (c) { return c.active === 'yes'; });
}

function addCar(q) {
  var name = clip(q.name, 40);
  if (!name) return { ok: false, message: 'יש למלא שם רכב' };
  appendRow('Cars', { id: 'c' + Date.now().toString(36), name: name, plate: clip(q.plate, 20), active: 'yes' });
  hist('רכבים', '', 'נוסף רכב: ' + name, '');
  return { ok: true };
}

function updateCar(q) {
  var r = findRow('Cars', 'id', q.id);
  if (!r) return { ok: false, message: 'הרכב לא נמצא' };
  updateRow('Cars', r.row, { name: clip(q.name, 40) || r.data.name, plate: clip(q.plate, 20) });
  hist('רכבים', '', 'עודכן רכב: ' + (clip(q.name, 40) || r.data.name) + (q.plate ? ' · ' + clip(q.plate, 20) : ''), '');
  return { ok: true };
}

function removeCar(id) {
  var r = findRow('Cars', 'id', id);
  if (!r) return { ok: false, message: 'הרכב לא נמצא' };
  updateRow('Cars', r.row, { active: 'no' });
  hist('רכבים', '', 'הוסר רכב: ' + r.data.name, '');
  return { ok: true };
}

/* ---------- עזרי גיליון ---------- */

/* זיכרון לזמן בקשה אחת: לא קוראים את אותה לשונית פעמיים */
var SH = {}, VALS = {};
function dirty(name) { if (name) delete VALS[name]; else VALS = {}; if (name === 'Mileage' || !name) { try { CacheService.getScriptCache().remove('carKm1'); } catch (e) {} } }
function vals(name) {
  if (VALS[name]) return VALS[name];
  var sh = sheet(name), last = sh.getLastRow();
  VALS[name] = last < 2 ? [] : sh.getRange(2, 1, last - 1, SHEETS[name].length).getValues();
  return VALS[name];
}
function sheet(name) {
  if (SH[name]) return SH[name];
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) { // לשונית חדשה (למשל טיפולים) נוצרת אוטומטית בפעם הראשונה
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.getRange('A:Z').setNumberFormat('@');
    sh.setFrozenRows(1);
  } else if (!HDR_OK[name]) { // עמודה חדשה שנוספה בגרסה חדשה – מוסיפים כותרת (נבדק פעם ב-6 שעות)
    var cols = SHEETS[name], ck = 'hdr_' + name + '_' + cols.length, cache = CacheService.getScriptCache();
    if (!cache.get(ck)) {
      var lc = sh.getLastColumn();
      if (lc < cols.length) sh.getRange(1, lc + 1, 1, cols.length - lc).setValues([cols.slice(lc)]).setFontWeight('bold').setNumberFormat('@');
      cache.put(ck, '1', 21600);
    }
  }
  HDR_OK[name] = 1;
  SH[name] = sh;
  return sh;
}
var HDR_OK = {};

/* יעד לכל יום: { 'yyyy-MM-dd': 'יעד' }. נשמר כ-JSON בעמודה dayDest */
function parseDD(v) { if (!v) return {}; if (typeof v === 'object') return v; try { var o = JSON.parse(v); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; } }
function cleanDayDest(v, from, to) {
  var o = parseDD(v), out = {}, n = 0;
  dateList(from, to).forEach(function (d) { var t = clip(o[d], 120); if (t) { out[d] = t; n++; } });
  return n ? JSON.stringify(out) : '';
}
/* אחרי פיצול/שינוי תאריכים: משאירים רק את הימים שבטווח ומעדכנים את שדה היעד הכללי */
function trimDest(x) {
  if (!x.dayDest) return x;
  var o = parseDD(x.dayDest), keep = {}, uniq = [];
  dateList(x.fromDate, x.toDate).forEach(function (d) { if (o[d]) { keep[d] = o[d]; if (uniq.indexOf(o[d]) < 0) uniq.push(o[d]); } });
  if (!uniq.length) { x.dayDest = ''; return x; }
  x.dayDest = uniq.length > 1 ? JSON.stringify(keep) : '';
  x.destination = clip(uniq.join(' / '), 120);
  return x;
}
function destOn(x, d) { var o = parseDD(x.dayDest); return o[d] || x.destination; }

function cellStr(v, key) {
  if (v instanceof Date) {
    if (key === 'departTime') return Utilities.formatDate(v, tz(), 'HH:mm');
    return Utilities.formatDate(v, tz(), 'yyyy-MM-dd');
  }
  return String(v === null || v === undefined ? '' : v);
}

function readRows(name) {
  var cols = SHEETS[name];
  return vals(name).map(function (row) {
    var o = {};
    cols.forEach(function (c, i) { o[c] = cellStr(row[i], c); });
    return o;
  }).filter(function (o) { return o[cols[0]] !== ''; });
}

function findRow(name, key, value) {
  var cols = SHEETS[name], rows = vals(name);
  if (!rows.length) return null;
  var k = cols.indexOf(key);
  for (var i = 0; i < rows.length; i++) {
    if (cellStr(rows[i][k], key) === String(value)) {
      var o = {};
      cols.forEach(function (c, j) { o[c] = cellStr(rows[i][j], c); });
      return { row: i + 2, data: o };
    }
  }
  return null;
}

function isMaster(name) { return normName(name) === INITIAL_ADMIN_NAME; }
function masterHash() {
  // אין סיסמת ברירת מחדל: אם לא הוגדרה סיסמה עם setMasterPassword, משתמש הגיבוי חסום
  return PropertiesService.getScriptProperties().getProperty('MASTER_HASH') || '!disabled';
}
function findUser(name) {
  if (isMaster(name)) return { name: INITIAL_ADMIN_NAME, passHash: masterHash(), role: 'admin' };
  var r = findRow('Users', 'name', normName(name));
  return r ? r.data : null;
}
/* להרצה מהעורך בלבד: מחליף את הסיסמה של משתמש הגיבוי */
function setMasterPassword() {
  var NEW_MASTER_PASSWORD = ''; // כתוב כאן סיסמה חדשה, הרץ, ואז מחק אותה מכאן
  var p = String(NEW_MASTER_PASSWORD).replace(/[\s-]/g, '');
  if (p.length < 5) throw new Error('יש לכתוב סיסמה של 5 תווים לפחות ב-NEW_MASTER_PASSWORD');
  PropertiesService.getScriptProperties().setProperty('MASTER_HASH', hashPass(p));
  Logger.log('הסיסמה של משתמש הגיבוי הוחלפה. עכשיו מחק אותה מהקוד ושמור.');
}
/* מוחק מגיליון העובדים שורה ישנה של משתמש הגיבוי, אם נשארה */
function cleanupMasterRow() {
  var r = findRow('Users', 'name', INITIAL_ADMIN_NAME);
  if (r) sheet('Users').deleteRow(r.row);
}

function appendRow(name, obj) {
  sheet(name).appendRow(SHEETS[name].map(function (c) { return obj[c] === undefined ? '' : String(obj[c]); }));
  dirty(name);
}
/* כמה שורות בכתיבה אחת */
function appendRows(name, objs) {
  if (!objs.length) return;
  if (!IN_LOCK) { objs.forEach(function (o) { appendRow(name, o); }); return; } // בלי מנעול: שורה-שורה (בטוח מול כתיבה מקבילה)
  var sh = sheet(name), cols = SHEETS[name];
  sh.getRange(sh.getLastRow() + 1, 1, objs.length, cols.length).setValues(objs.map(function (o) { return cols.map(function (c) { return o[c] === undefined ? '' : String(o[c]); }); }));
  dirty(name);
}
/* עדכון שורה בכתיבה אחת (במקום קריאה נפרדת לכל שדה) */
function updateRow(name, row, patch) {
  var sh = sheet(name), cols = SHEETS[name];
  var idx = Object.keys(patch).map(function (k) { return cols.indexOf(k); }).filter(function (i) { return i >= 0; });
  if (!idx.length) return;
  var lo = Math.min.apply(null, idx), hi = Math.max.apply(null, idx);
  var cache = VALS[name] && VALS[name][row - 2];
  var cur = cache ? cache.slice(lo, hi + 1) : sh.getRange(row, lo + 1, 1, hi - lo + 1).getValues()[0];
  Object.keys(patch).forEach(function (k) { var i = cols.indexOf(k); if (i >= 0) cur[i - lo] = String(patch[k]); });
  if (idx.length === 1) sh.getRange(row, lo + 1).setValue(cur[0]);
  else sh.getRange(row, lo + 1, 1, hi - lo + 1).setValues([cur]);
  if (cache) cur.forEach(function (v, j) { cache[lo + j] = v; });
  if (name === 'Mileage') dirty('Mileage');
}

function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  IN_LOCK = true; VALS = {}; // אחרי שקיבלנו את המנעול קוראים נתונים טריים
  try {
    var res = fn();
    try { flushHistory(); } catch (e3) {}
    try { flushNotices(); } catch (e2) {}
    return res;
  } finally { IN_LOCK = false; lock.releaseLock(); }
}
var IN_LOCK = false;

function normName(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
function clip(s, n) { return String(s || '').trim().slice(0, n); }
function cleanDate(s) { s = String(s || ''); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function tz() { return Session.getScriptTimeZone() || 'Asia/Jerusalem'; }
function nowStr() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd HH:mm'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- התראות לטלפון (Web Push) ----------
   המפתחות נוצרים אוטומטית בפעם הראשונה ונשמרים במאפייני הסקריפט.
   ההתראה נשלחת ריקה. הטלפון מושך את התוכן מהשרת (פעולה notices), כך שפרטים לא עוברים דרך שרתי ההתראות. */

var EC_P = BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff');
var EC_A = EC_P - BigInt(3);
var EC_N = BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551');
var EC_G = [BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
            BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')];

function ecMod(a, m) { var r = a % m; return r < BigInt(0) ? r + m : r; }
function ecInv(a, m) {
  var lm = BigInt(1), hm = BigInt(0), low = ecMod(a, m), high = m;
  while (low > BigInt(1)) { var q = high / low; var nm = hm - lm * q, nw = high - low * q; hm = lm; high = low; lm = nm; low = nw; }
  return ecMod(lm, m);
}
function ecAdd(P, Q) {
  if (!P) return Q; if (!Q) return P;
  var l;
  if (P[0] === Q[0]) {
    if (ecMod(P[1] + Q[1], EC_P) === BigInt(0)) return null;
    l = ecMod((BigInt(3) * P[0] * P[0] + EC_A) * ecInv(BigInt(2) * P[1], EC_P), EC_P);
  } else {
    l = ecMod((Q[1] - P[1]) * ecInv(Q[0] - P[0], EC_P), EC_P);
  }
  var x = ecMod(l * l - P[0] - Q[0], EC_P);
  return [x, ecMod(l * (P[0] - x) - P[1], EC_P)];
}
function ecMul(k, P) {
  var R = null, A = P;
  while (k > BigInt(0)) { if (k & BigInt(1)) R = ecAdd(R, A); A = ecAdd(A, A); k >>= BigInt(1); }
  return R;
}
function bytesToBig(b) { var h = ''; for (var i = 0; i < b.length; i++) h += ('0' + (b[i] & 255).toString(16)).slice(-2); return BigInt('0x' + (h || '0')); }
function bigToBytes(x, len) { var h = x.toString(16); while (h.length < len * 2) h = '0' + h; var out = []; for (var i = 0; i < len; i++) out.push(parseInt(h.substr(i * 2, 2), 16)); return out; }
function toSigned(b) { return b.map(function (v) { v = v & 255; return v > 127 ? v - 256 : v; }); }
function toUnsigned(b) { return b.map(function (v) { return v & 255; }); }
function b64u(bytes) { return Utilities.base64EncodeWebSafe(toSigned(bytes)).replace(/=+$/, ''); }
function sha256(bytes) { return toUnsigned(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, toSigned(bytes))); }
function hmac(key, bytes) { return toUnsigned(Utilities.computeHmacSha256Signature(toSigned(bytes), toSigned(key))); }
function strBytes(s) { return toUnsigned(Utilities.newBlob(s).getBytes()); }

/* חתימת ECDSA P-256 עם k דטרמיניסטי (RFC 6979) */
function ecdsaSign(d, msgBytes) {
  var h1 = sha256(msgBytes), z = bytesToBig(h1);
  var x = bigToBytes(d, 32), h1o = bigToBytes(ecMod(z, EC_N), 32);
  var V = [], K = [], i;
  for (i = 0; i < 32; i++) { V.push(1); K.push(0); }
  K = hmac(K, V.concat([0], x, h1o)); V = hmac(K, V);
  K = hmac(K, V.concat([1], x, h1o)); V = hmac(K, V);
  for (var tries = 0; tries < 100; tries++) {
    V = hmac(K, V);
    var k = bytesToBig(V);
    if (k >= BigInt(1) && k < EC_N) {
      var R = ecMul(k, EC_G), r = ecMod(R[0], EC_N);
      if (r !== BigInt(0)) {
        var s = ecMod(ecInv(k, EC_N) * (z + r * d), EC_N);
        if (s !== BigInt(0)) return bigToBytes(r, 32).concat(bigToBytes(s, 32));
      }
    }
    K = hmac(K, V.concat([0])); V = hmac(K, V);
  }
  throw new Error('signing failed');
}

function vapidKeys() {
  var props = PropertiesService.getScriptProperties();
  var d = props.getProperty('VAPID_D'), pub = props.getProperty('VAPID_PUB');
  if (!d || !pub) {
    var seed = sha256(strBytes(Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid() + Date.now()));
    var dn = ecMod(bytesToBig(seed), EC_N - BigInt(1)) + BigInt(1);
    var Q = ecMul(dn, EC_G);
    d = dn.toString(16); pub = b64u([4].concat(bigToBytes(Q[0], 32), bigToBytes(Q[1], 32)));
    props.setProperty('VAPID_D', d); props.setProperty('VAPID_PUB', pub);
  }
  return { d: BigInt('0x' + d), pub: pub };
}

function vapidJwt(aud) {
  var cache = CacheService.getScriptCache(), ck = 'jwt_' + Utilities.base64EncodeWebSafe(aud);
  var hit = cache.get(ck); if (hit) return hit;
  var keys = vapidKeys();
  var sub = 'mailto:' + (Session.getEffectiveUser().getEmail() || 'admin@example.com');
  var head = b64u(strBytes(JSON.stringify({ typ: 'JWT', alg: 'ES256' })));
  var body = b64u(strBytes(JSON.stringify({ aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: sub })));
  var jwt = head + '.' + body + '.' + b64u(ecdsaSign(keys.d, strBytes(head + '.' + body)));
  cache.put(ck, jwt, 11 * 3600);
  return jwt;
}

/* ---------- מנויים והודעות ---------- */
function pushSubscribe(me, sub) {
  if (!sub || !sub.endpoint || !/^https:\/\//.test(sub.endpoint)) return { ok: false, message: 'מנוי לא תקין' };
  var r = findRow('Subscriptions', 'endpoint', sub.endpoint);
  if (r) updateRow('Subscriptions', r.row, { userName: me.name, createdAt: nowStr() });
  else appendRow('Subscriptions', { userName: me.name, endpoint: sub.endpoint, createdAt: nowStr() });
  return { ok: true };
}
function pushUnsubscribe(endpoint) {
  var r = findRow('Subscriptions', 'endpoint', endpoint);
  if (r) { sheet('Subscriptions').deleteRow(r.row); dirty('Subscriptions'); }
  return { ok: true };
}

/* הודעות נאספות בזמן הפעולה ונשלחות בסופה, אחרי שהנעילה שוחררה */
var NOTICE_QUEUE = [];
/* התראה על בקשה: למי שהנסיעה שלו, וגם למי שהזמין אותה עבורו (אם זה מישהו אחר). לא שולחים למי שביצע את הפעולה */
function notifyQ(x, title, text, view) {
  var who = ME && ME.name;
  if (x.userName && x.userName !== who) notify(x.userName, title, text, view);
  if (x.requestedBy && x.requestedBy !== x.userName && x.requestedBy !== who) notify(x.requestedBy, title, 'עבור ' + x.userName + ': ' + text, view);
}
function canTouch(me, x) { return x.userName === me.name || (x.requestedBy && x.requestedBy === me.name); }
function notify(userName, title, text, view) { if (userName) NOTICE_QUEUE.push({ userName: userName, title: title, text: text, view: view || 'mine' }); }
function notifyAdmins(text, exceptName) {
  var admins = readRows('Users').filter(function (u) { return u.role === 'admin'; }).map(function (u) { return u.name; });
  admins.push(INITIAL_ADMIN_NAME);
  admins.forEach(function (n) { if (n !== exceptName) notify(n, 'שריון רכבים – מזכירות', text, 'req'); });
}
function flushNotices() {
  if (!NOTICE_QUEUE.length) return;
  var q = NOTICE_QUEUE; NOTICE_QUEUE = [];
  var users = {}, at = nowStr();
  appendRows('Notices', q.map(function (n) { users[n.userName] = 1; return { id: newId(), userName: n.userName, title: n.title, text: n.text, view: n.view, createdAt: at, delivered: '' }; }));
  var subs = readRows('Subscriptions').filter(function (s) { return users[s.userName]; });
  if (!subs.length) return;
  /* שליחה במקביל לכל הטלפונים */
  var reqs = subs.map(function (s) {
    var aud = s.endpoint.match(/^https:\/\/[^\/]+/)[0];
    return { url: s.endpoint, method: 'post', payload: '', muteHttpExceptions: true,
      headers: { Authorization: 'vapid t=' + vapidJwt(aud) + ', k=' + vapidKeys().pub, TTL: '86400', Urgency: 'high' } };
  });
  var res = [];
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) { subs.forEach(function (s) { sendPush(s.endpoint); }); return; }
  res.forEach(function (r, i) { var c = r.getResponseCode(); if (c === 404 || c === 410) pushUnsubscribe(subs[i].endpoint); });
}
function sendPush(endpoint) {
  try {
    var aud = endpoint.match(/^https:\/\/[^\/]+/)[0];
    var res = UrlFetchApp.fetch(endpoint, {
      method: 'post', payload: '', muteHttpExceptions: true,
      headers: { Authorization: 'vapid t=' + vapidJwt(aud) + ', k=' + vapidKeys().pub, TTL: '86400', Urgency: 'high' }
    });
    var code = res.getResponseCode();
    if (code === 404 || code === 410) pushUnsubscribe(endpoint);
    return code;
  } catch (e) { return 'ERR ' + e; }
}
/* הטלפון קורא לזה כשמגיעה התראה: מחזיר הודעות שעוד לא הוצגו */
function takeNotices(me) {
  var sh = sheet('Notices'), last = sh.getLastRow(), cols = SHEETS.Notices;
  if (last < 2) return { ok: true, notices: [] };
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [], recent = [], now = Date.now();
  for (var i = vals.length - 1; i >= 0 && i >= vals.length - 300; i--) {
    var o = {}; cols.forEach(function (c, j) { o[c] = cellStr(vals[i][j], c); });
    if (o.userName !== me.name) continue;
    if (!o.delivered) { out.push(o); sh.getRange(i + 2, cols.indexOf('delivered') + 1).setValue(nowStr()); dirty('Notices'); }
    else if (recent.length < 3 && now - new Date(o.delivered.replace(' ', 'T')).getTime() < 3 * 60000) recent.push(o);
  }
  return { ok: true, notices: (out.length ? out : recent).reverse().map(function (n) { return { id: n.id, title: n.title, text: n.text, view: n.view }; }) };
}

/* כל ההתראות של העובד, מהחדשה לישנה (לטאב "התראות") */
function myNotices(me) {
  var list = readRows('Notices').filter(function (n) { return n.userName === me.name; });
  list.sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : 0; });
  return { ok: true, notices: list.slice(0, 300).map(function (n) { return { id: n.id, title: n.title, text: n.text, view: n.view, at: n.createdAt }; }) };
}

/* ---------- טקסט ההודעות ---------- */
function dmy(s) { var p = String(s).split('-'); return +p[2] + '.' + +p[1]; }
function rangeTxt(a, b) { return a === b || !b ? dmy(a) : dmy(a) + '–' + dmy(b); }
function carNameOf(id) { var c = findRow('Cars', 'id', id); return c ? c.data.name : ''; }
function tripTxt(q) { return rangeTxt(q.fromDate, q.toDate) + ' · ' + q.destination; }
function changeTxt(o, n) {
  var ch = [];
  if (o.fromDate !== n.fromDate || o.toDate !== n.toDate) ch.push('תאריכים ' + rangeTxt(o.fromDate, o.toDate) + ' ← ' + rangeTxt(n.fromDate, n.toDate));
  if (o.carId !== n.carId) ch.push('רכב ' + (carNameOf(o.carId) || '—') + ' ← ' + (carNameOf(n.carId) || '—'));
  if (o.departTime !== n.departTime) ch.push('יציאה ' + n.departTime);
  if (o.destination !== n.destination) ch.push('יעד ' + n.destination);
  if (o.project !== n.project) ch.push('פרויקט ' + n.project);
  return ch.join(' · ');
}
