/**
 * خادم الحضور - Google Apps Script
 * - المالك يبدأ محاضرة (اسم + تاريخ) من صفحة admin.html
 * - كل محاضرة لها ورقة (Sheet) مستقلة تجمع كل الطلاب
 * - عند الإغلاق (يدويًا أو تلقائيًا بعد مدة) يصلك إيميل واحد بملف Excel واحد
 */

// ====== الإعدادات (عدّلها قبل النشر) ======
const ADMIN_PASSWORD = 'admin1234';                 // باسورد المالك (لصفحة admin.html)
const DEFAULT_STUDENT_PASSWORD = '1234';            // باسورد الطلاب (يمكن تغييره من صفحة المالك)
const DEFAULT_OWNER_EMAIL = 'your-email@gmail.com'; // الإيميل الذي يصله الملف (يمكن تغييره من صفحة المالك)
const SPREADSHEET_ID = '';                          // معرّف الشيت: ضعه فقط إذا أنشأت السكربت من script.google.com (وليس من داخل الشيت)
const ALLOW_DUPLICATES = false;                     // false: الطالب نفسه (نفس الاسم+الفرقة+القسم) يُسجَّل مرة واحدة
// =========================================

const HEADERS = ['م', 'وقت الإرسال', 'اسم الطالب', 'الفرقة', 'القسم', 'المحاضرة', 'تاريخ المحاضرة',
                 'خط العرض', 'خط الطول', 'دقة الموقع (م)', 'رابط الخريطة', 'مفتاح'];
const KEY_COL = 12;

/** يفتح الشيت: بالمعرّف إن وُضع، وإلا الشيت المربوط بالسكربت. */
function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

/* ---------- نقاط الدخول ---------- */

function doGet(e) {
  const a = e && e.parameter && e.parameter.action;
  if (a === 'config') return out_(publicConfig_());
  return ContentService.createTextOutput('نظام الحضور يعمل');
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'bad-request' }); }
  try {
    if (d.action === 'admin') return out_(admin_(d));
    return out_(submit_(d));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- الحالة المخزنة ---------- */

function props_() { return PropertiesService.getScriptProperties(); }
function getSession_() { const s = props_().getProperty('SESSION'); return s ? JSON.parse(s) : null; }
function setSession_(s) {
  props_().setProperty('SESSION', JSON.stringify(s));
  CacheService.getScriptCache().remove('cfg');
}
function studentPass_() { return props_().getProperty('STUDENT_PASS') || DEFAULT_STUDENT_PASSWORD; }
function ownerEmail_() { return props_().getProperty('OWNER_EMAIL') || DEFAULT_OWNER_EMAIL; }

/** ما تراه صفحة الطلاب (مخزّن مؤقتًا 20 ثانية لتحمّل ضغط الدخول). */
function publicConfig_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('cfg');
  if (c) return JSON.parse(c);
  const s = getSession_();
  const cfg = s ? { ok: true, active: true, open: !!s.open, title: s.title, date: s.date }
                : { ok: true, active: false, open: false };
  cache.put('cfg', JSON.stringify(cfg), 20);
  return cfg;
}

/* ---------- استقبال الطلاب ---------- */

function clean_(v, max) {
  let t = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max || 150);
  if (/^[=+\-@]/.test(t)) t = "'" + t;   // منع تنفيذ أي نص كمعادلة
  return t;
}
function makeKey_(name, grade, dept) {
  return [name, grade, dept].map(function (x) {
    return String(x).replace(/[أإآ]/g, 'ا').replace(/\s+/g, ' ').trim().toLowerCase();
  }).join('|');
}

function submit_(d) {
  if (String(d.password) !== String(studentPass_())) return { ok: false, error: 'password' };
  const s = getSession_();
  if (!s) return { ok: false, error: 'no-session' };
  if (!s.open) return { ok: false, error: 'closed' };

  const name = clean_(d.name, 120), grade = clean_(d.grade, 60), dept = clean_(d.dept, 80);
  if (!name || !grade || !dept) return { ok: false, error: 'missing' };
  const lat = Number(d.lat), lng = Number(d.lng);
  if (!isFinite(lat) || !isFinite(lng)) return { ok: false, error: 'location' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'busy' };   // الصفحة تعيد المحاولة تلقائيًا
  try {
    const sh = ss_().getSheetByName(s.sheet);
    if (!sh) return { ok: false, error: 'no-sheet' };
    const key = makeKey_(name, grade, dept);
    const last = sh.getLastRow();

    if (!ALLOW_DUPLICATES && last > 1) {
      const f = sh.getRange(2, KEY_COL, last - 1, 1).createTextFinder(key).matchEntireCell(true).findNext();
      if (f) return { ok: true, duplicate: true };
    }
    sh.appendRow([
      last,                         // م (الصف الأول عنوان، فالرقم = عدد الصفوف الحالية)
      new Date(), name, grade, dept, s.title, "'" + s.date,
      lat, lng, Number(d.acc) || '',
      'https://www.google.com/maps?q=' + lat + ',' + lng,
      key
    ]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- أوامر المالك ---------- */

function admin_(d) {
  if (String(d.adminPassword) !== String(ADMIN_PASSWORD)) return { ok: false, error: 'admin-password' };
  switch (d.op) {
    case 'status': return status_();
    case 'start':  return start_(d);
    case 'open':   return setOpen_(true);
    case 'close':  return closeSession_(!!d.send);
    case 'send':   return sendOnly_();
    default:       return { ok: false, error: 'unknown-op' };
  }
}

function countRows_(s) {
  if (!s) return 0;
  const sh = ss_().getSheetByName(s.sheet);
  return sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
}

function status_() {
  const s = getSession_();
  return { ok: true, session: s, count: countRows_(s), email: ownerEmail_() };
}

function uniqueSheetName_(ss, title, date) {
  const t = String(title).replace(/[\[\]\*\?:\/\\]/g, ' ').replace(/\s+/g, ' ').trim();
  const dd = date.replace(/\//g, '-');
  let n = 1;
  while (true) {
    const suffix = ' ' + dd + (n > 1 ? ' #' + n : '');
    const name = t.slice(0, Math.max(31 - suffix.length, 1)) + suffix;
    if (!ss.getSheetByName(name)) return name;
    n++;
  }
}

function start_(d) {
  const title = clean_(d.title, 100);
  const date = String(d.date || '');
  if (!title) return { ok: false, error: 'اكتب اسم المحاضرة' };
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(date)) return { ok: false, error: 'التاريخ غير صحيح' };

  if (d.studentPassword) props_().setProperty('STUDENT_PASS', String(d.studentPassword));
  if (d.email) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.email))) return { ok: false, error: 'الإيميل غير صحيح' };
    props_().setProperty('OWNER_EMAIL', String(d.email));
  }

  cancelAutoClose_();
  const old = getSession_();
  if (old && old.open) { old.open = false; setSession_(old); }

  const ss = ss_();
  const name = uniqueSheetName_(ss, title, date);
  const sh = ss.insertSheet(name);
  sh.appendRow(HEADERS);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d9ece9');
  sh.setRightToLeft(true);
  sh.hideColumns(KEY_COL);
  if (sh.getMaxRows() < 2000) sh.insertRowsAfter(sh.getMaxRows(), 2000 - sh.getMaxRows());

  const minutes = Math.max(Number(d.minutes) || 0, 0);
  const s = {
    title: title, date: date, sheet: name, gid: sh.getSheetId(), open: true,
    startedAt: new Date().toISOString(),
    closeAt: minutes ? new Date(Date.now() + minutes * 60000).toISOString() : null
  };
  setSession_(s);
  if (minutes) ScriptApp.newTrigger('autoClose').timeBased().after(minutes * 60000).create();
  return status_();
}

function setOpen_(open) {
  const s = getSession_();
  if (!s) return { ok: false, error: 'لا توجد محاضرة' };
  s.open = open;
  if (open) { s.closeAt = null; cancelAutoClose_(); }
  setSession_(s);
  return status_();
}

function closeSession_(send) {
  const s = getSession_();
  if (!s) return { ok: false, error: 'لا توجد محاضرة' };
  s.open = false;
  setSession_(s);
  cancelAutoClose_();
  const r = status_();
  if (send) r.sent = sendExcel_(s);
  return r;
}

function sendOnly_() {
  const s = getSession_();
  if (!s) return { ok: false, error: 'لا توجد محاضرة' };
  const r = status_();
  r.sent = sendExcel_(s);
  return r;
}

/* ---------- الإغلاق التلقائي ---------- */

function autoClose() {
  const s = getSession_();
  cancelAutoClose_();
  if (!s) return;
  s.open = false;
  setSession_(s);
  sendExcel_(s);
}

function cancelAutoClose_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoClose') ScriptApp.deleteTrigger(t);
  });
}

/* ---------- ملف Excel والإيميل (إيميل واحد لكل محاضرة) ---------- */

function exportSheetXlsx_(s) {
  const ss = ss_();
  const sh = ss.getSheetByName(s.sheet);
  const tmp = SpreadsheetApp.create('tmp-attendance');
  try {
    const copy = sh.copyTo(tmp);
    copy.setName(s.sheet.slice(0, 31));
    tmp.deleteSheet(tmp.getSheets()[0]);
    copy.deleteColumn(KEY_COL);          // عمود المفتاح داخلي، لا حاجة له في الملف
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    const fname = (s.title + ' ' + s.date.replace(/\//g, '-')).replace(/[\\\/:*?"<>|]/g, ' ') + '.xlsx';
    return blob.setName(fname);
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}

function sendExcel_(s) {
  const count = countRows_(s);
  const blob = exportSheetXlsx_(s);
  const ss = ss_();
  MailApp.sendEmail({
    to: ownerEmail_(),
    subject: 'حضور: ' + s.title + ' - ' + s.date + ' (' + count + ' طالب)',
    htmlBody: '<div dir="rtl">المحاضرة: <b>' + s.title + '</b><br>التاريخ: <b>' + s.date + '</b><br>' +
              'عدد المسجّلين: <b>' + count + '</b><br>الملف المرفق يضم جميع الطلاب مع موقع كل طالب.<br>' +
              '<a href="' + ss.getUrl() + '">فتح الشيت على الإنترنت</a></div>',
    attachments: [blob]
  });
  return count;
}

/** شغّل هذه الدالة مرة واحدة من المحرر لمنح الصلاحيات (Drive / Mail / Triggers). */
function setup() {
  DriveApp.getFileById(ss_().getId());
  ScriptApp.getProjectTriggers();
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch('https://www.google.com');
}
