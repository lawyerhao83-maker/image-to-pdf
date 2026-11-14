/**
 * 案件管理簿：Google Apps Script 後端
 *
 * 資料：本試算表的各分頁（案件、期日、工作紀錄、人員、假日、操作記錄）
 * 權限：網頁以「存取者本人」身分執行，所以只有被共用這份試算表的人能讀寫。
 * 日曆：由擁有者的定時觸發器每 10 分鐘把「期日」同步到「事務所期日」日曆。
 */

var APP_NAME = '案件管理簿';
var CAL_NAME = '事務所期日';
var HOLIDAY_CAL = 'zh-tw.taiwan#holiday@group.v.calendar.google.com';
var KINDS = ['開庭', '書狀期限', '上訴期限', '抗告期限', '會議', '其他'];

var SCHEMA = {
  cases: { sheet: '案件', cols: [
    ['id', 'ID'], ['code', '本所編號'], ['title', '案由'], ['client', '委任人'], ['opponent', '對造'],
    ['type', '類型'], ['stage', '審級'], ['court', '法院'], ['courtNo', '法院案號'], ['division', '股別'],
    ['status', '狀態'], ['openedAt', '委任日'], ['lawyers', '承辦人'], ['driveUrl', '卷宗資料夾'], ['notes', '備註'],
    ['createdAt', '建立時間'], ['createdBy', '建立者'], ['updatedAt', '更新時間'], ['updatedBy', '更新者']] },
  events: { sheet: '期日', cols: [
    ['id', 'ID'], ['caseId', '案件ID'], ['kind', '類別'], ['date', '日期'], ['time', '時間'], ['place', '地點'],
    ['note', '說明'], ['done', '完成'], ['deleted', '已刪除'], ['calId', '日曆ID'], ['calSig', '同步碼'],
    ['createdAt', '建立時間'], ['createdBy', '建立者'], ['updatedAt', '更新時間'], ['updatedBy', '更新者']] },
  logs: { sheet: '工作紀錄', cols: [['id', 'ID'], ['caseId', '案件ID'], ['text', '內容'], ['by', '記錄者'], ['at', '時間']] },
  staff: { sheet: '人員', cols: [['email', 'Email'], ['name', '姓名'], ['title', '職稱'], ['active', '啟用']] },
  holidays: { sheet: '假日', cols: [['date', '日期'], ['name', '名稱'], ['kind', '類型']] },
  audit: { sheet: '操作記錄', cols: [['at', '時間'], ['by', '使用者'], ['action', '動作'], ['target', '對象'], ['detail', '內容']] }
};
var BOOL = { done: 1, deleted: 1, active: 1 };

/* ================= 共用工具 ================= */

function props_() { return PropertiesService.getScriptProperties(); }
function need_(k) {
  var v = props_().getProperty(k);
  if (!v) throw new Error('系統尚未初始化：請擁有者在試算表選單「案件系統 → 初始化」執行一次。');
  return v;
}
function ss_() {
  var id = need_('SHEET_ID');
  try { return SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('你沒有這份案件資料的存取權。請洽管理者用「案件系統 → 新增同事」把你的 Google 帳號加入。'); }
}
function email_() { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); }
function now_() { return new Date().toISOString(); }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Taipei'; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function isDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 20000); }
// 避免使用者輸入的文字被試算表當成公式
function safe_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }

function sheet_(name, ss) {
  var def = SCHEMA[name];
  var sh = (ss || ss_()).getSheetByName(def.sheet);
  if (!sh) throw new Error('找不到「' + def.sheet + '」分頁，請擁有者重新執行初始化。');
  return sh;
}

function norm_(k, v) {
  if (BOOL[k]) return v === true || String(v).toUpperCase() === 'TRUE';
  if (v instanceof Date) {
    var f = k === 'time' ? 'HH:mm' : (k === 'date' || k === 'openedAt') ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ssXXX";
    return Utilities.formatDate(v, tz_(), f);
  }
  return v == null ? '' : String(v);
}

/** 讀整個分頁；依標題列找欄位，所以欄位順序被調動也沒關係。 */
function read_(name, ss) {
  var def = SCHEMA[name];
  var sh = sheet_(name, ss);
  var values = sh.getDataRange().getValues();
  var head = values[0].map(String);
  var idx = def.cols.map(function (c) { return head.indexOf(c[1]); });
  var missing = def.cols.filter(function (c, i) { return idx[i] < 0; }).map(function (c) { return c[1]; });
  if (missing.length) throw new Error('「' + def.sheet + '」分頁缺少欄位：' + missing.join('、') + '。請擁有者重新執行初始化。');
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.every(function (v) { return v === '' || v === null; })) continue;
    var o = { _row: r + 1 };
    def.cols.forEach(function (c, i) { o[c[0]] = norm_(c[0], row[idx[i]]); });
    rows.push(o);
  }
  return { sh: sh, head: head, idx: idx, rows: rows, def: def };
}

function col_(t, key) {
  for (var i = 0; i < t.def.cols.length; i++) if (t.def.cols[i][0] === key) return t.idx[i] + 1;
  throw new Error('unknown column ' + key);
}

/** 寫入一列；rowNum 為空時新增在最後。只改 obj 裡有的欄位。 */
function write_(t, obj, rowNum) {
  var width = t.head.length;
  var arr = rowNum ? t.sh.getRange(rowNum, 1, 1, width).getValues()[0] : new Array(width).fill('');
  t.def.cols.forEach(function (c, i) {
    if (!(c[0] in obj)) return;
    arr[t.idx[i]] = BOOL[c[0]] ? !!obj[c[0]] : safe_(obj[c[0]]);
  });
  if (rowNum) t.sh.getRange(rowNum, 1, 1, width).setValues([arr]);
  else t.sh.appendRow(arr);
}

function locked_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中（可能正在同步日曆），請幾秒後再試。');
  try { var r = fn(); SpreadsheetApp.flush(); return r; }
  finally { lock.releaseLock(); }
}

function audit_(ss, action, target, detail) {
  var t = read_('audit', ss);
  write_(t, { at: now_(), by: email_(), action: action, target: target, detail: clean_(detail, 500) }, null);
}

function staffName_(ss, email) {
  var s = read_('staff', ss).rows.filter(function (r) { return r.email.toLowerCase() === email; })[0];
  return s && s.name ? s.name : (email || '其他同事');
}

function checkVer_(ss, row, expected) {
  if (expected && row.updatedAt && row.updatedAt !== expected) {
    throw new Error('CONFLICT:這筆資料剛被' + staffName_(ss, row.updatedBy) + '修改過。畫面已更新為最新內容，請確認後再改一次。');
  }
}

/* ================= 網頁 ================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function apiLoad() {
  var ss = ss_();
  var staff = read_('staff', ss).rows.map(function (s) {
    return { email: s.email.toLowerCase().trim(), name: s.name, title: s.title, active: s.active };
  });
  var cases = read_('cases', ss).rows.map(function (c) {
    var o = strip_(c);
    o.lawyers = c.lawyers ? c.lawyers.split(',').map(function (x) { return x.trim(); }).filter(String) : [];
    return o;
  });
  var events = read_('events', ss).rows.filter(function (e) { return !e.deleted; }).map(function (e) {
    var o = strip_(e); delete o.calId; delete o.calSig; delete o.deleted; return o;
  });
  return JSON.stringify({
    me: email_(), staff: staff, cases: cases, events: events, today: today_(),
    sheetUrl: ss.getUrl(), owner: props_().getProperty('OWNER') || ''
  });
}
function strip_(o) { var c = {}; for (var k in o) if (k !== '_row') c[k] = o[k]; return c; }

function apiSaveCase(c) {
  return locked_(function () {
    var ss = ss_(), t = read_('cases', ss), me = email_(), now = now_();
    var d = {};
    ['code', 'title', 'client', 'opponent', 'type', 'stage', 'court', 'courtNo', 'division', 'status', 'openedAt', 'driveUrl', 'notes']
      .forEach(function (k) { d[k] = clean_(c[k]); });
    d.lawyers = (Array.isArray(c.lawyers) ? c.lawyers : []).map(function (s) { return String(s).toLowerCase().trim(); }).filter(String).join(',');
    if (!d.title || !d.client) throw new Error('請填寫案由與委任人。');
    if (d.openedAt && !isDate_(d.openedAt)) throw new Error('委任日格式不正確。');
    if (d.driveUrl && !/^https:\/\/(drive|docs)\.google\.com\//.test(d.driveUrl)) throw new Error('卷宗資料夾請貼上 Google 雲端硬碟的連結。');
    d.updatedAt = now; d.updatedBy = me;
    var id = c.id;
    if (id) {
      var row = t.rows.filter(function (r) { return r.id === id; })[0];
      if (!row) throw new Error('這件案件已被刪除。');
      checkVer_(ss, row, c.updatedAt);
      write_(t, d, row._row);
      audit_(ss, '修改案件', id, d.code + ' ' + d.title);
    } else {
      id = uid_(); d.id = id; d.createdAt = now; d.createdBy = me;
      write_(t, d, null);
      audit_(ss, '新增案件', id, d.code + ' ' + d.title);
    }
    return JSON.stringify({ id: id });
  });
}

function apiDeleteCase(id) {
  return locked_(function () {
    var ss = ss_(), t = read_('cases', ss), me = email_(), now = now_();
    var row = t.rows.filter(function (r) { return r.id === id; })[0];
    if (!row) return JSON.stringify({ ok: true });
    var ev = read_('events', ss);
    ev.rows.forEach(function (e) {
      if (e.caseId === id && !e.deleted) write_(ev, { deleted: true, updatedAt: now, updatedBy: me }, e._row);
    });
    var lg = read_('logs', ss);
    lg.rows.filter(function (l) { return l.caseId === id; }).map(function (l) { return l._row; })
      .sort(function (a, b) { return b - a; }).forEach(function (r) { lg.sh.deleteRow(r); });
    t.sh.deleteRow(row._row);
    audit_(ss, '刪除案件', id, row.code + ' ' + row.title + '（委任人 ' + row.client + '）');
    return JSON.stringify({ ok: true });
  });
}

function apiSaveEvent(e) {
  return locked_(function () {
    var ss = ss_(), t = read_('events', ss), me = email_(), now = now_();
    var d = {
      caseId: clean_(e.caseId, 64), kind: clean_(e.kind, 20), date: clean_(e.date, 10), time: clean_(e.time, 5),
      place: clean_(e.place, 200), note: clean_(e.note, 1000), updatedAt: now, updatedBy: me
    };
    if (!isDate_(d.date)) throw new Error('請填寫正確的日期。');
    if (d.time && !/^\d{2}:\d{2}$/.test(d.time)) throw new Error('時間格式不正確。');
    if (KINDS.indexOf(d.kind) < 0) d.kind = '其他';
    var c = read_('cases', ss).rows.filter(function (r) { return r.id === d.caseId; })[0];
    if (!c) throw new Error('找不到這件案件，可能已被刪除。');
    var id = e.id;
    if (id) {
      var row = t.rows.filter(function (r) { return r.id === id && !r.deleted; })[0];
      if (!row) throw new Error('這筆期日已被刪除。');
      checkVer_(ss, row, e.updatedAt);
      write_(t, d, row._row);
      audit_(ss, '修改期日', id, c.title + '／' + d.kind + ' ' + d.date);
    } else {
      id = uid_(); d.id = id; d.done = false; d.deleted = false; d.createdAt = now; d.createdBy = me;
      write_(t, d, null);
      audit_(ss, '新增期日', id, c.title + '／' + d.kind + ' ' + d.date);
    }
    return JSON.stringify({ id: id });
  });
}

function apiDeleteEvent(id) {
  return locked_(function () {
    var ss = ss_(), t = read_('events', ss);
    var row = t.rows.filter(function (r) { return r.id === id; })[0];
    if (!row || row.deleted) return JSON.stringify({ ok: true });
    write_(t, { deleted: true, updatedAt: now_(), updatedBy: email_() }, row._row);
    audit_(ss, '刪除期日', id, row.kind + ' ' + row.date + ' ' + row.note);
    return JSON.stringify({ ok: true });
  });
}

function apiSetDone(id, done) {
  return locked_(function () {
    var ss = ss_(), t = read_('events', ss);
    var row = t.rows.filter(function (r) { return r.id === id && !r.deleted; })[0];
    if (!row) throw new Error('這筆期日已被刪除。');
    write_(t, { done: !!done, updatedAt: now_(), updatedBy: email_() }, row._row);
    audit_(ss, done ? '期日完成' : '取消完成', id, row.kind + ' ' + row.date);
    return JSON.stringify({ ok: true });
  });
}

function apiLogs(caseId) {
  var rows = read_('logs').rows.filter(function (l) { return l.caseId === caseId; }).map(strip_);
  rows.sort(function (a, b) { return b.at < a.at ? -1 : b.at > a.at ? 1 : 0; });
  return JSON.stringify(rows);
}

function apiAddLog(caseId, text) {
  return locked_(function () {
    var ss = ss_();
    text = clean_(text, 5000);
    if (!text) throw new Error('紀錄內容是空的。');
    var c = read_('cases', ss).rows.filter(function (r) { return r.id === caseId; })[0];
    if (!c) throw new Error('找不到這件案件，可能已被刪除。');
    var id = uid_();
    write_(read_('logs', ss), { id: id, caseId: caseId, text: text, by: email_(), at: now_() }, null);
    audit_(ss, '新增紀錄', caseId, c.title);
    return JSON.stringify({ id: id });
  });
}

function apiDeleteLog(id) {
  return locked_(function () {
    var ss = ss_(), t = read_('logs', ss);
    var row = t.rows.filter(function (r) { return r.id === id; })[0];
    if (!row) return JSON.stringify({ ok: true });
    if (row.by !== email_()) throw new Error('只能刪除自己寫的紀錄。');
    t.sh.deleteRow(row._row);
    audit_(ss, '刪除紀錄', row.caseId, row.text.slice(0, 80));
    return JSON.stringify({ ok: true });
  });
}

/* ================= 期限計算 ================= */

/**
 * 民法§120 始日不算入；§122 末日為休息日者以次日代之。
 * off：放假日 {日期: 名稱}；work：補行上班日 {日期: 名稱}。
 * 假日表不完整時，結果只會偏早（較保守），不會偏晚。
 */
function computeDeadline_(served, days, transit, off, work) {
  var p = served.split('-').map(Number);
  var dt = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  dt.setUTCDate(dt.getUTCDate() + days + transit);
  var fmt = function (d) { return d.toISOString().slice(0, 10); };
  var original = fmt(dt), skipped = [];
  for (var i = 0; i < 40; i++) {
    var k = fmt(dt), wd = dt.getUTCDay();
    var isOff = work[k] ? false : (off[k] || wd === 0 || wd === 6);
    if (!isOff) break;
    skipped.push(k + ' ' + (off[k] || (wd === 0 ? '週日' : '週六')));
    dt.setUTCDate(dt.getUTCDate() + 1);
  }
  var workNote = work[fmt(dt)] ? work[fmt(dt)] : '';
  return { date: fmt(dt), original: original, skipped: skipped, workNote: workNote };
}

function apiDeadline(served, days, transit) {
  if (!isDate_(served)) throw new Error('請輸入送達日。');
  days = Math.floor(Number(days)) || 0; transit = Math.floor(Number(transit)) || 0;
  if (days < 1 || days > 365) throw new Error('期間天數請填 1 到 365。');
  if (transit < 0 || transit > 60) throw new Error('在途期間請填 0 到 60 日。');
  var hol = read_('holidays').rows, off = {}, work = {}, years = {};
  hol.forEach(function (h) {
    if (!isDate_(h.date)) return;
    years[h.date.slice(0, 4)] = 1;
    if (h.kind === '補班') work[h.date] = h.name || '補行上班日';
    else if (h.kind === '放假' || h.kind === '') off[h.date] = h.name || '假日';
  });
  var r = computeDeadline_(served, days, transit, off, work);
  r.hasYear = !!years[r.date.slice(0, 4)];
  return JSON.stringify(r);
}

/* ================= 日曆同步（擁有者的觸發器執行） ================= */

function evInfo_(e, c) {
  var title = (e.done ? '✓ ' : '') + '[' + e.kind + '] ' + c.title + (c.client ? '（' + c.client + '）' : '');
  var desc = [
    c.code ? '本所編號：' + c.code : '', c.court ? '法院：' + c.court : '',
    c.courtNo ? '案號：' + c.courtNo + (c.division ? ' ' + c.division : '') : '',
    e.note ? '說明：' + e.note : '', '', '由「案件管理簿」自動建立，請在系統內修改，日曆上的修改會被覆蓋。'
  ].filter(function (x, i) { return x || i === 4; }).join('\n');
  return { title: title, date: e.date, time: e.time, loc: e.place || c.court || '', desc: desc };
}

function syncCalendar() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var ss = ss_();
    var cal = CalendarApp.getCalendarById(need_('CAL_ID'));
    if (!cal) throw new Error('找不到「' + CAL_NAME + '」日曆。');
    var cases = {};
    read_('cases', ss).rows.forEach(function (c) { cases[c.id] = c; });
    var t = read_('events', ss);
    var cId = col_(t, 'calId'), cSig = col_(t, 'calSig');
    var started = Date.now(), purge = [];
    t.rows.forEach(function (e) {
      if (Date.now() - started > 60000) return; // 一次最多 1 分鐘，剩下的下次再同步
      var c = cases[e.caseId];
      if (e.deleted || !c) {
        if (e.calId) delEvent_(cal, e.calId);
        if (e.deleted) purge.push(e._row);
        else if (e.calId) { t.sh.getRange(e._row, cId).setValue(''); t.sh.getRange(e._row, cSig).setValue(''); }
        return;
      }
      var info = evInfo_(e, c);
      var sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(info), Utilities.Charset.UTF_8));
      if (e.calId && e.calSig === sig) return;
      if (e.calId) delEvent_(cal, e.calId);
      var p = e.date.split('-').map(Number), ev;
      if (e.time) {
        var hm = e.time.split(':').map(Number);
        var start = new Date(p[0], p[1] - 1, p[2], hm[0], hm[1]);
        ev = cal.createEvent(info.title, start, new Date(start.getTime() + 3600000), { description: info.desc, location: info.loc });
      } else {
        ev = cal.createAllDayEvent(info.title, new Date(p[0], p[1] - 1, p[2]), { description: info.desc, location: info.loc });
      }
      t.sh.getRange(e._row, cId).setValue(ev.getId());
      t.sh.getRange(e._row, cSig).setValue(sig);
    });
    purge.sort(function (a, b) { return b - a; }).forEach(function (r) { t.sh.deleteRow(r); });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
function delEvent_(cal, id) { try { var ev = cal.getEventById(id); if (ev) ev.deleteEvent(); } catch (err) { /* 已不存在 */ } }

/* ================= 每日 Email 提醒（擁有者的觸發器執行） ================= */

function esc_(s) { return String(s || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function roc_(d) { var p = d.split('-'); return (Number(p[0]) - 1911) + '/' + p[1] + '/' + p[2]; }

function dailyDigest() {
  var ss = ss_(), today = today_();
  var lim = new Date(); lim.setDate(lim.getDate() + 7);
  var until = Utilities.formatDate(lim, tz_(), 'yyyy-MM-dd');
  var cases = {};
  read_('cases', ss).rows.forEach(function (c) { cases[c.id] = c; });
  var list = read_('events', ss).rows.filter(function (e) {
    return !e.deleted && !e.done && cases[e.caseId] && e.date <= until;
  }).sort(function (a, b) { return (a.date + a.time) < (b.date + b.time) ? -1 : 1; });
  if (!list.length) return;
  var overdue = list.filter(function (e) { return e.date < today; }).length;
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (err) { }
  read_('staff', ss).rows.filter(function (s) { return s.active && s.email; }).forEach(function (s) {
    var me = s.email.toLowerCase().trim();
    var rows = list.map(function (e) {
      var c = cases[e.caseId], mine = c.lawyers.split(',').indexOf(me) >= 0;
      var late = e.date < today;
      return '<tr><td style="padding:6px 10px;font-family:monospace;color:' + (late ? '#B02A20' : '#15202D') + '">' + roc_(e.date) + (e.time ? ' ' + esc_(e.time) : '') + (late ? '（逾期）' : '') + '</td>' +
        '<td style="padding:6px 10px">' + esc_(e.kind) + '</td><td style="padding:6px 10px">' + (mine ? '★ ' : '') + esc_(c.title) + '（' + esc_(c.client) + '）<br><span style="color:#5A6676;font-size:12px">' + esc_([c.courtNo, e.place, e.note].filter(String).join('・')) + '</span></td></tr>';
    }).join('');
    var html = '<div style="font-family:sans-serif;font-size:14px;color:#15202D"><p>' + esc_(s.name || '') + ' 您好，以下是今天起 7 日內的期日' + (overdue ? '，另有 <b style="color:#B02A20">' + overdue + ' 筆逾期未完成</b>' : '') + '。★ 為您承辦的案件。</p>' +
      '<table style="border-collapse:collapse;border:1px solid #D7DDE5">' + rows + '</table>' +
      (url ? '<p><a href="' + esc_(url) + '">開啟案件管理簿</a></p>' : '') + '<p style="color:#5A6676;font-size:12px">此信由案件管理簿每天早上自動寄出。</p></div>';
    MailApp.sendEmail({ to: me, subject: '【期日提醒】7 日內 ' + (list.length - overdue) + ' 筆' + (overdue ? '，逾期 ' + overdue + ' 筆' : ''), htmlBody: html, name: APP_NAME });
  });
}

/* ================= 試算表選單（擁有者使用） ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('案件系統')
    .addItem('初始化（第一次使用）', 'setup')
    .addSeparator()
    .addItem('新增同事…', 'addStaff')
    .addItem('停用同事…', 'removeStaff')
    .addSeparator()
    .addItem('匯入台灣假日…', 'importHolidays')
    .addItem('立即同步日曆', 'syncNow')
    .addItem('寄一次今日提醒（測試）', 'dailyDigest')
    .addToUi();
}

function ownerOnly_() {
  var owner = props_().getProperty('OWNER');
  if (owner && owner !== email_()) {
    SpreadsheetApp.getUi().alert('這個功能只有系統擁有者（' + owner + '）可以使用。');
    return false;
  }
  return true;
}

function setup() {
  var ui = SpreadsheetApp.getUi();
  if (!ownerOnly_()) return;
  var ss = SpreadsheetApp.getActive(), me = email_();
  Object.keys(SCHEMA).forEach(function (k) {
    var def = SCHEMA[k], sh = ss.getSheetByName(def.sheet) || ss.insertSheet(def.sheet);
    var lastCol = sh.getLastColumn();
    var head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
    var missing = def.cols.map(function (c) { return c[1]; }).filter(function (l) { return head.indexOf(l) < 0; });
    if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    lastCol = sh.getLastColumn();
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#E2EAF6');
    if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, lastCol).setNumberFormat('@');
  });
  ['工作表1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  var P = props_();
  P.setProperty('SHEET_ID', ss.getId());
  P.setProperty('OWNER', me);
  if (!P.getProperty('CAL_ID') || !CalendarApp.getCalendarById(P.getProperty('CAL_ID'))) {
    var cal = CalendarApp.createCalendar(CAL_NAME, { color: CalendarApp.Color.BLUE });
    P.setProperty('CAL_ID', cal.getId());
  }
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (['syncCalendar', 'dailyDigest'].indexOf(tr.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('syncCalendar').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(8).everyDays(1).inTimezone('Asia/Taipei').create();
  var st = read_('staff', ss);
  if (!st.rows.some(function (r) { return r.email.toLowerCase() === me; })) {
    var res = ui.prompt('你的姓名', '會顯示在承辦人與工作紀錄上（例如：王大明律師）', ui.ButtonSet.OK);
    write_(st, { email: me, name: clean_(res.getResponseText(), 50) || me, title: '管理者', active: true }, null);
  }
  ui.alert('初始化完成',
    '已建立資料分頁、「' + CAL_NAME + '」日曆、每 10 分鐘的日曆同步，以及每天早上 8 點的 Email 提醒。\n\n' +
    '下一步：\n1. 在 Apps Script 編輯器按「部署 → 新增部署作業 → 網頁應用程式」。\n2. 用「案件系統 → 新增同事」把同事加進來。\n3. 用「案件系統 → 匯入台灣假日」匯入今年和明年的假日。', ui.ButtonSet.OK);
}

function addStaff() {
  if (!ownerOnly_()) return;
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('新增同事（1/3）', '同事的 Google 帳號 Email', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var email = r1.getResponseText().trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { ui.alert('Email 格式不正確。'); return; }
  var r2 = ui.prompt('新增同事（2/3）', '姓名（例如：陳小華律師）', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var r3 = ui.prompt('新增同事（3/3）', '職稱（律師／助理／其他）', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var ss = SpreadsheetApp.getActive(), notes = [];
  ss.addEditor(email);
  try {
    Calendar.Acl.insert({ role: 'reader', scope: { type: 'user', value: email } }, need_('CAL_ID'));
  } catch (e) { notes.push('日曆共用失敗，請到 Google 日曆手動把「' + CAL_NAME + '」共用給對方（' + e.message + '）'); }
  var st = read_('staff', ss);
  var row = st.rows.filter(function (r) { return r.email.toLowerCase() === email; })[0];
  write_(st, { email: email, name: clean_(r2.getResponseText(), 50), title: clean_(r3.getResponseText(), 20), active: true }, row ? row._row : null);
  audit_(ss, '新增同事', email, r2.getResponseText());
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
  ui.alert('已加入 ' + email + '\n\n請把系統網址傳給對方：\n' + (url || '（尚未部署網頁應用程式）') + (notes.length ? '\n\n注意：' + notes.join('\n') : ''));
}

function removeStaff() {
  if (!ownerOnly_()) return;
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('停用同事', '要停用的 Google 帳號 Email（資料與紀錄會保留）', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var email = r.getResponseText().trim().toLowerCase();
  if (email === props_().getProperty('OWNER')) { ui.alert('不能停用系統擁有者。'); return; }
  var ss = SpreadsheetApp.getActive();
  try { ss.removeEditor(email); } catch (e) { }
  try { ss.removeViewer(email); } catch (e) { }
  try { Calendar.Acl.remove(need_('CAL_ID'), 'user:' + email); } catch (e) { }
  var st = read_('staff', ss);
  var row = st.rows.filter(function (x) { return x.email.toLowerCase() === email; })[0];
  if (row) write_(st, { active: false }, row._row);
  audit_(ss, '停用同事', email, '');
  ui.alert('已停用 ' + email + '，對方已無法開啟試算表與系統。');
}

function importHolidays() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('匯入台灣假日', '要匯入哪一年？（西元，例如 ' + (new Date().getFullYear() + 1) + '）', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var y = Number(r.getResponseText());
  if (!(y >= 2000 && y <= 2100)) { ui.alert('年份不正確。'); return; }
  var cal = CalendarApp.getCalendarById(HOLIDAY_CAL) || CalendarApp.subscribeToCalendar(HOLIDAY_CAL);
  var ss = SpreadsheetApp.getActive(), t = read_('holidays', ss), have = {}, n = 0;
  t.rows.forEach(function (h) { have[h.date] = 1; });
  cal.getEvents(new Date(y, 0, 1), new Date(y + 1, 0, 1)).forEach(function (ev) {
    if (!ev.isAllDayEvent()) return;
    var d = Utilities.formatDate(ev.getAllDayStartDate(), tz_(), 'yyyy-MM-dd');
    if (have[d]) return;
    var name = ev.getTitle(), desc = ev.getDescription() || '';
    var kind = /補.*班|上班/.test(name) ? '補班' : /紀念|Observance/i.test(desc) ? '請確認' : '放假';
    write_(t, { date: d, name: name, kind: kind }, null); have[d] = 1; n++;
  });
  ui.alert('已匯入 ' + y + ' 年 ' + n + ' 筆',
    '請對照行政院人事行政總處公布的辦公日曆表核對「假日」分頁：\n' +
    '・「類型」只有「放假」會延後期限末日，「補班」視為上班日。\n' +
    '・標為「請確認」的是 Google 認為可能不放假的紀念日，計算時不當作假日。\n' +
    '・颱風假等臨時停班請自行新增。\n\n假日表不完整時，算出的末日只會偏早，不會偏晚。', ui.ButtonSet.OK);
}

function syncNow() {
  syncCalendar();
  SpreadsheetApp.getUi().alert('日曆已同步。');
}
