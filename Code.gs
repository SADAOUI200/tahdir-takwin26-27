/**
 * Code.gs — الخادم: التهيئة، المصادقة، وحفظ/قراءة البيانات
 * ربط المشروع: ملف Google Sheet ← الإضافات ← Apps Script (أو clasp).
 * شغّل الدالة setupAccounts() مرة واحدة فقط ثم انشر التطبيق كـ Web App.
 */

const SH = { USERS: 'Users', COUNTS: 'Counts', LISTS: 'Lists' };
const HEAD = {
  Users:  ['username', 'salt', 'hash', 'role', 'dirCode', 'dirName', 'mustChange', 'lastLogin', 'createdAt'],
  Counts: ['dirCode', 'table', 'center', 'level', 'formation', 'subject', 'value', 'updatedAt', 'updatedBy'],
  Lists:  ['dirCode', 'listKey', 'subject', 'lastName', 'firstName', 'updatedAt', 'updatedBy']
};
// أعمدة تُجبر على الصيغة النصية (حتى لا يتحول "01" إلى 1)
const TEXTCOLS = { Users: [1, 2, 3, 5], Counts: [1, 3, 4, 5, 6], Lists: [1, 2, 3, 4, 5] };
const TZ = 'Africa/Algiers';

/* ===================== نقطة الدخول ===================== */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/** واجهة JSON تستعملها صفحة GitHub Pages (قائمة بيضاء بالدوال المسموحة) */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const api = { login: login, logout: logout, changePassword: changePassword, getCounts: getCounts,
                  saveCounts: saveCounts, deleteCenter: deleteCenter, getList: getList, saveList: saveList,
                  getStatus: getStatus, resetDirPassword: resetDirPassword };
    if (!Object.prototype.hasOwnProperty.call(api, req.fn)) throw new Error('طلب غير صالح');
    out = { ok: true, result: api[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/* ===================== التهيئة (مرة واحدة) ===================== */
function setupAccounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());

  Object.keys(HEAD).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    TEXTCOLS[name].forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  });

  const users = ss.getSheetByName(SH.USERS);
  if (users.getLastRow() > 1) {
    Logger.log('الحسابات منشأة مسبقاً — لم يتم أي تغيير. (لإعادة التهيئة احذف صفوف ورقة Users يدوياً.)');
    return;
  }

  const now = new Date();
  const rows = directorates_().map(function (d) {
    const salt = Utilities.getUuid().slice(0, 12);
    return [d.username, salt, hash_(salt, d.password), 'dir', d.code, d.name, true, '', now];
  });

  const minPw = props.getProperty('MINISTRY_INITIAL_PASSWORD') || ('Mn-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10));
  const mSalt = Utilities.getUuid().slice(0, 12);
  rows.unshift([MINISTRY.username, mSalt, hash_(mSalt, minPw), 'ministry', '', MINISTRY.name, true, '', now]);

  users.getRange(2, 1, rows.length, HEAD.Users.length).setValues(rows);
  props.deleteProperty('MINISTRY_INITIAL_PASSWORD');

  const msg = 'تم إنشاء ' + rows.length + ' حساباً.\nحساب الوزارة: ' + MINISTRY.username + '\nكلمة المرور الأولية: ' + minPw +
    '\n(احفظها الآن — لن تظهر مرة أخرى؛ سيُطلب تغييرها عند أول دخول)';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* التشغيل من المحرر */ }
}

/** تعرض أسماء الحسابات (دون كلمات المرور) في السجل */
function listAccounts() {
  const v = sheet_(SH.USERS).getDataRange().getValues().slice(1);
  Logger.log(v.map(function (r) { return r[0] + ' — ' + r[5]; }).join('\n'));
}

/** إعادة تعيين كلمة مرور مديرية من المحرر: resetPassword('de05') */
function resetPassword(username) {
  const d = directorates_().filter(function (x) { return x.username === username; })[0];
  if (!d) throw new Error('اسم مستخدم غير معروف');
  setPassword_(username, d.password, true);
  Logger.log('أُعيدت كلمة المرور الأولية للحساب ' + username);
}

/* ===================== أدوات داخلية ===================== */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) { return ss_().getSheetByName(name); }
function hash_(salt, pw) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8));
}
function dc_(v) { const s = String(v); return /^\d$/.test(s) ? '0' + s : s; }
function clean_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, 60); }
function fmt_(d) { return d ? Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd HH:mm') : ''; }
function withLock_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}
function replaceRows_(sh, keep, newRows, nCols) {
  const last = sh.getLastRow();
  let kept = [];
  if (last > 1) kept = sh.getRange(2, 1, last - 1, nCols).getValues().filter(keep);
  const all = kept.concat(newRows);
  if (last > 1) sh.getRange(2, 1, last - 1, nCols).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, nCols).setValues(all);
}
function dirNames_() { const m = {}; directorates_().forEach(function (d) { m[d.code] = d.name; }); return m; }
function allowedKeys_() {
  const set = {};
  Object.keys(LEVELS).forEach(function (lv) {
    Object.keys(LEVELS[lv].formations).forEach(function (f) {
      LEVELS[lv].subjects.forEach(function (sj) { set[lv + '|' + f + '|' + sj] = true; });
    });
  });
  return set;
}
function checkListKey_(key) {
  const p = String(key).split('|');
  const t = LIST_TYPES.filter(function (x) { return x.id === p[0]; })[0];
  if (!t) throw new Error('قائمة غير معروفة');
  if (t.ranks.length > 1) { if (t.ranks.indexOf(p[1]) < 0) throw new Error('رتبة غير صالحة'); }
  else if (p[1]) throw new Error('مفتاح غير صالح');
  return t;
}
function clientConfig_(isMin) {
  return {
    app: { title: APP.TITLE, year: APP.YEAR },
    levels: LEVELS, listTypes: LIST_TYPES, tables: TABLES,
    directorates: isMin ? directorates_().map(function (d) { return { code: d.code, name: d.name }; }) : []
  };
}
function setPassword_(username, newPw, mustChange) {
  withLock_(function () {
    const sh = sheet_(SH.USERS), v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]).toLowerCase() === String(username).toLowerCase()) {
        const salt = Utilities.getUuid().slice(0, 12);
        sh.getRange(i + 1, 2, 1, 2).setValues([[salt, hash_(salt, newPw)]]);
        sh.getRange(i + 1, 7).setValue(!!mustChange);
        return;
      }
    }
    throw new Error('الحساب غير موجود');
  });
}

/* ===================== الجلسات ===================== */
function auth_(token, allowPending) {
  const raw = token && CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('AUTH');
  const s = JSON.parse(raw);
  if (s.mc && !allowPending) throw new Error('MUSTCHANGE');
  return s;
}
function scope_(s, dir) {
  if (s.role === 'dir') return s.dir;                       // المديرية لا ترى إلا بياناتها
  if (!dir || dir === 'ALL') return 'ALL';
  if (!directorates_().some(function (d) { return d.code === dir; })) throw new Error('مديرية غير معروفة');
  return dir;
}

function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  const cache = CacheService.getScriptCache(), fk = 'fail_' + username;
  if (Number(cache.get(fk) || 0) >= APP.MAX_FAILS) {
    return { ok: false, error: 'تم تجاوز عدد المحاولات المسموح بها. أعد المحاولة بعد 10 دقائق.' };
  }
  const sh = sheet_(SH.USERS), vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[0]).toLowerCase() !== username) continue;
    if (hash_(r[1], password) !== r[2]) break;
    const mc = r[6] === true || r[6] === 'TRUE';
    sh.getRange(i + 1, 8).setValue(new Date());
    const sess = { u: r[0], role: r[3], dir: r[3] === 'dir' ? dc_(r[4]) : '', name: r[5], mc: mc };
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put('sess_' + token, JSON.stringify(sess), APP.SESSION_TTL);
    cache.remove(fk);
    return {
      ok: true, token: token,
      user: { username: r[0], role: r[3], dirCode: sess.dir, dirName: r[5], mustChange: mc },
      config: clientConfig_(r[3] === 'ministry')
    };
  }
  cache.put(fk, String(Number(cache.get(fk) || 0) + 1), APP.LOCK_SECONDS);
  return { ok: false, error: 'اسم المستخدم أو كلمة المرور غير صحيحة.' };
}

function logout(token) { if (token) CacheService.getScriptCache().remove('sess_' + token); return true; }

function changePassword(token, oldPw, newPw) {
  const s = auth_(token, true);
  oldPw = String(oldPw || ''); newPw = String(newPw || '');
  if (newPw.length < APP.MIN_PW || !/[A-Za-z\u0600-\u06FF]/.test(newPw) || !/\d/.test(newPw)) {
    throw new Error('كلمة المرور الجديدة: ' + APP.MIN_PW + ' أحرف على الأقل وتحتوي على حروف وأرقام.');
  }
  if (newPw === oldPw) throw new Error('كلمة المرور الجديدة مطابقة للقديمة.');
  const def = directorates_().filter(function (d) { return d.username === s.u; })[0];
  if (def && def.password === newPw) throw new Error('لا يمكن استعمال كلمة المرور الأولية.');

  const v = sheet_(SH.USERS).getDataRange().getValues();
  let ok = false;
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === s.u) { ok = hash_(v[i][1], oldPw) === v[i][2]; break; }
  }
  if (!ok) throw new Error('كلمة المرور الحالية غير صحيحة.');
  setPassword_(s.u, newPw, false);
  s.mc = false;
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(s), APP.SESSION_TTL);
  return true;
}

/* ===================== الجداول التعدادية ===================== */
/** يعيد {T1:{key:n}, T2:{...}, T4:{مركز:{...}}} — مفتاح الخلية: المستوى|التكوين|المادة */
function getCounts(token, dir) {
  const s = auth_(token), sc = scope_(s, dir);
  const vals = sheet_(SH.COUNTS).getDataRange().getValues();
  const out = { T1: {}, T2: {}, T4: {} };
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (sc !== 'ALL' && dc_(r[0]) !== sc) continue;
    const v = Number(r[6]); if (!v) continue;
    const k = r[3] + '|' + r[4] + '|' + r[5];
    if (r[1] === 'T4') {
      const cn = sc === 'ALL' ? '__ALL__' : String(r[2]);
      out.T4[cn] = out.T4[cn] || {};
      out.T4[cn][k] = (out.T4[cn][k] || 0) + v;
    } else if (out[r[1]]) {
      out[r[1]][k] = (out[r[1]][k] || 0) + v;
    }
  }
  return out;
}

function saveCounts(token, table, center, data) {
  const s = auth_(token);
  if (s.role !== 'dir') throw new Error('الإدخال مخصص للمديريات فقط.');
  if (['T1', 'T2', 'T4'].indexOf(table) < 0) throw new Error('جدول غير صالح.');
  center = table === 'T4' ? clean_(center) : '';
  if (table === 'T4' && !center) throw new Error('أدخل اسم مركز التكوين.');
  const allowed = allowedKeys_(), now = new Date(), rows = [];
  Object.keys(data || {}).forEach(function (k) {
    if (!allowed[k]) throw new Error('خلية غير صالحة.');
    const n = Math.floor(Number(data[k]));
    if (!isFinite(n) || n < 0 || n > 9999) throw new Error('قيمة غير صالحة.');
    if (!n) return;
    const p = k.split('|');
    rows.push([s.dir, table, center, p[0], p[1], p[2], n, now, s.u]);
  });
  withLock_(function () {
    const sh = sheet_(SH.COUNTS);
    if (table === 'T4') {
      const cs = {}; sh.getDataRange().getValues().forEach(function (r, i) { if (i && dc_(r[0]) === s.dir && r[1] === 'T4') cs[r[2]] = 1; });
      if (!cs[center] && Object.keys(cs).length >= 12) throw new Error('الحد الأقصى 12 مركزاً لكل مديرية.');
    }
    replaceRows_(sh, function (r) { return !(dc_(r[0]) === s.dir && r[1] === table && String(r[2]) === center); }, rows, 9);
  });
  return { saved: rows.length };
}

function deleteCenter(token, center) {
  const s = auth_(token);
  if (s.role !== 'dir') throw new Error('غير مسموح.');
  center = clean_(center);
  withLock_(function () {
    replaceRows_(sheet_(SH.COUNTS), function (r) { return !(dc_(r[0]) === s.dir && r[1] === 'T4' && String(r[2]) === center); }, [], 9);
  });
  return true;
}

/* ===================== القوائم الاسمية ===================== */
function getList(token, listKey, dir) {
  const s = auth_(token); checkListKey_(listKey);
  const sc = scope_(s, dir), names = dirNames_(), order = {};
  directorates_().forEach(function (d, i) { order[d.code] = i; });
  const vals = sheet_(SH.LISTS).getDataRange().getValues(), out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i]; if (r[1] !== listKey) continue;
    const d = dc_(r[0]); if (sc !== 'ALL' && d !== sc) continue;
    out.push({ subject: String(r[2]), last: String(r[3]), first: String(r[4]), dirName: names[d] || d, _o: order[d] });
  }
  if (sc === 'ALL') out.sort(function (a, b) { return a._o - b._o; });
  return out;
}

function saveList(token, listKey, rows) {
  const s = auth_(token);
  if (s.role !== 'dir') throw new Error('الإدخال مخصص للمديريات فقط.');
  const type = checkListKey_(listKey);
  if (!Array.isArray(rows) || rows.length > 500) throw new Error('عدد الأسطر غير صالح (الحد 500).');
  const now = new Date(), out = [];
  rows.forEach(function (r) {
    const last = clean_(r.last), first = clean_(r.first);
    if (!last && !first) return;
    let subj = '';
    if (type.subjects.length) {
      subj = String(r.subject || '');
      if (type.subjects.indexOf(subj) < 0) throw new Error('مادة غير صالحة.');
    }
    out.push([s.dir, listKey, subj, last, first, now, s.u]);
  });
  withLock_(function () {
    replaceRows_(sheet_(SH.LISTS), function (r) { return !(dc_(r[0]) === s.dir && r[1] === listKey); }, out, 7);
  });
  return { saved: out.length };
}

/* ===================== لوحة متابعة الوزارة ===================== */
function getStatus(token) {
  const s = auth_(token);
  if (s.role !== 'ministry') throw new Error('FORBIDDEN');
  const st = {};
  sheet_(SH.USERS).getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[3] !== 'dir') return;
    const c = dc_(r[4]);
    st[c] = { code: c, name: r[5], username: r[0], lastLogin: fmt_(r[7]), changed: !(r[6] === true || r[6] === 'TRUE'),
              t1: 0, t2: 0, t4: 0, centers: {}, lists: 0, upd: '' };
  });
  const bump = function (o, ts) { const f = fmt_(ts); if (f > o.upd) o.upd = f; };
  sheet_(SH.COUNTS).getDataRange().getValues().slice(1).forEach(function (r) {
    const o = st[dc_(r[0])]; if (!o) return;
    const v = Number(r[6]) || 0;
    if (r[1] === 'T1') o.t1 += v; else if (r[1] === 'T2') o.t2 += v;
    else if (r[1] === 'T4') { o.t4 += v; o.centers[r[2]] = 1; }
    bump(o, r[7]);
  });
  sheet_(SH.LISTS).getDataRange().getValues().slice(1).forEach(function (r) {
    const o = st[dc_(r[0])]; if (!o) return;
    o.lists++; bump(o, r[5]);
  });
  return Object.keys(st).map(function (k) { const o = st[k]; o.centers = Object.keys(o.centers).length; return o; });
}

function resetDirPassword(token, dirCode) {
  const s = auth_(token);
  if (s.role !== 'ministry') throw new Error('FORBIDDEN');
  const d = directorates_().filter(function (x) { return x.code === dirCode; })[0];
  if (!d) throw new Error('مديرية غير معروفة');
  setPassword_(d.username, d.password, true);
  return true;
}
