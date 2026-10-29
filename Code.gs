/**
 * 國北詞曲創作社 — 社辦時段預約 後端 (Google Apps Script)
 *
 * 【資料結構】（執行 setup() 會自動建立）
 *  Users         userId | name | studentId | createdAt
 *  Bookings      bookingId | weekStart | day | slot | userId | name | createdAt
 *                  weekStart = 該週星期一 (yyyy-MM-dd)；day = 1~5；slot = 0~12（對應 SLOT_STARTS）
 *  Fixed         day | slot | label        ← 樂器班／固定時段，每週重置不會清除
 *  Announcements id | text | createdAt
 *  Admins        userId | note            ← 管理員 LINE userId 名單，請手動填寫
 *  Config        key | value              ← clubPassword（社辦密碼）、lineChannelId（LINE Login Channel ID）
 *
 * 【部署步驟】
 *  1. 建立 Google 試算表 → 擴充功能 → Apps Script，貼上本檔
 *  2. 專案設定：勾選「顯示 appsscript.json」，將 "timeZone" 設為 "Asia/Taipei"
 *  3. 執行 setup()（授權後會建立工作表與每週日 23:59 的觸發程序）
 *  4. 在 Config 填入 lineChannelId；在 Admins 填入管理員的 LINE userId
 *  5. 部署 → 新增部署作業 → 網頁應用程式：執行身分「我」、存取權「所有人」，複製 /exec 網址到 app.js
 *  修改程式後需「管理部署作業 → 編輯 → 新版本」才會生效。
 *
 * 【安全】前端傳來的是 LIFF ID Token，後端以 LINE verify API 驗證後才取得 userId，
 *        無法由使用者自行偽造他人身分。
 */

const TZ = 'Asia/Taipei';
const LIMIT = 3; // 每人每週上限（固定樂器班不計）
const SLOT_STARTS = ['08:00', '09:00', '10:00', '11:00', '12:00', '13:30', '14:30', '15:30', '16:30', '18:00', '19:00', '20:00', '21:00'];

const SCHEMA = {
  Users: { cols: ['userId', 'name', 'studentId', 'createdAt'], text: ['C'] },
  Bookings: { cols: ['bookingId', 'weekStart', 'day', 'slot', 'userId', 'name', 'createdAt'], text: ['B'] },
  Fixed: { cols: ['day', 'slot', 'label'], text: [] },
  Announcements: { cols: ['id', 'text', 'createdAt'], text: [] },
  Admins: { cols: ['userId', 'note'], text: [] },
  Config: { cols: ['key', 'value'], text: ['B'] },
};

/* ================= 初始化 / 排程 ================= */
function setup() {
  const ss = ss_();
  Object.keys(SCHEMA).forEach((name) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(SCHEMA[name].cols); sh.setFrozenRows(1); }
    SCHEMA[name].text.forEach((c) => sh.getRange(c + ':' + c).setNumberFormat('@')); // 避免日期/學號被轉格式
  });
  if (!rows_('Config').length) {
    const cfg = ss.getSheetByName('Config');
    cfg.appendRow(['clubPassword', '0000']);
    cfg.appendRow(['lineChannelId', '請填入 LINE Login Channel ID']);
  }
  // 每週日 23:59 重置（Google 觸發程序實際會在 23:45~00:15 之間執行；資料以 weekStart 區分，不影響正確性）
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'resetWeeklyBookings')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resetWeeklyBookings').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).nearMinute(59).create();
}

/** 清除一般預約；Fixed（樂器班）、Users 不受影響 */
function resetWeeklyBookings() {
  withLock_(() => {
    const sh = ss_().getSheetByName('Bookings');
    const n = sh.getLastRow() - 1;
    if (n > 0) sh.deleteRows(2, n);
    console.log('weekly reset: removed ' + Math.max(n, 0) + ' bookings');
  });
}

/* ================= Web App 進入點 ================= */
function doGet(e) { return handle_(e.parameter || {}, false); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, code: 'BAD', error: '請求格式錯誤' }); }
  return handle_(p, true);
}

const WRITE_ACTIONS = {
  register: register_, book: book_, cancel: cancel_,
  admin_setFixed: adminSetFixed_, admin_addAnnouncement: adminAddAnn_,
  admin_deleteAnnouncement: adminDelAnn_, admin_setPassword: adminSetPw_,
};

function handle_(p, isPost) {
  try {
    const user = verify_(p.idToken);
    if (!isPost) {
      if (p.action !== 'init') throw err_('BAD', '未知的操作');
      return out_({ ok: true, data: snapshot_(user) });
    }
    const fn = WRITE_ACTIONS[p.action];
    if (!fn) throw err_('BAD', '未知的操作');
    withLock_(() => fn(user, p));
    return out_({ ok: true, data: snapshot_(user) });
  } catch (e) {
    return out_({ ok: false, code: e.code || 'ERR', error: e.message });
  }
}

/* ================= 動作 ================= */
function register_(u, p) {
  const name = String(p.name || '').trim();
  const sid = String(p.studentId || '').trim();
  if (!name || name.length > 20) throw err_('BAD', '請輸入正確的姓名');
  if (!/^[A-Za-z0-9\-]{3,20}$/.test(sid)) throw err_('BAD', '學號／社員編號格式不正確');
  if (findUser_(u.userId)) return; // 已註冊，直接回傳最新資料
  ss_().getSheetByName('Users').appendRow([u.userId, name, sid, new Date().toISOString()]);
}

function book_(u, p) {
  const me = requireUser_(u);
  const day = Number(p.day), slot = Number(p.slot);
  checkCell_(day, slot);
  const ws = weekStart_();
  if (slotStart_(ws, day, slot) <= Date.now()) throw err_('PAST', '這個時段已過期');
  if (rows_('Fixed').some((f) => Number(f.day) === day && Number(f.slot) === slot)) throw err_('FIXED', '這是固定時段，無法預約');
  const bookings = rows_('Bookings').filter((b) => b.weekStart === ws);
  if (bookings.some((b) => Number(b.day) === day && Number(b.slot) === slot)) throw err_('TAKEN', '這個時段剛被預約了');
  if (bookings.filter((b) => b.userId === u.userId).length >= LIMIT) throw err_('LIMIT', '每週最多預約 ' + LIMIT + ' 個時段');
  ss_().getSheetByName('Bookings').appendRow([Utilities.getUuid(), ws, day, slot, u.userId, me.name, new Date().toISOString()]);
}

function cancel_(u, p) {
  requireUser_(u);
  const day = Number(p.day), slot = Number(p.slot);
  checkCell_(day, slot);
  const ws = weekStart_();
  const b = rows_('Bookings').find((x) => x.weekStart === ws && Number(x.day) === day && Number(x.slot) === slot && x.userId === u.userId);
  if (!b) throw err_('NONE', '找不到這筆預約');
  if (slotStart_(ws, day, slot) <= Date.now()) throw err_('PAST', '時間已過的時段無法取消');
  ss_().getSheetByName('Bookings').deleteRow(b._row);
}

function adminSetFixed_(u, p) {
  requireAdmin_(u);
  const day = Number(p.day), slot = Number(p.slot);
  checkCell_(day, slot);
  const ws = weekStart_();
  const fixedSh = ss_().getSheetByName('Fixed');
  // 先移除該格既有的固定設定
  rows_('Fixed').filter((f) => Number(f.day) === day && Number(f.slot) === slot)
    .map((f) => f._row).sort((a, b) => b - a).forEach((r) => fixedSh.deleteRow(r));
  if (p.clear) return;
  const label = String(p.label || '樂器班').trim().slice(0, 10);
  // 該格若已有一般預約則一併移除
  const bkSh = ss_().getSheetByName('Bookings');
  rows_('Bookings').filter((b) => b.weekStart === ws && Number(b.day) === day && Number(b.slot) === slot)
    .map((b) => b._row).sort((a, b) => b - a).forEach((r) => bkSh.deleteRow(r));
  fixedSh.appendRow([day, slot, label]);
}

function adminAddAnn_(u, p) {
  requireAdmin_(u);
  const text = String(p.text || '').trim().slice(0, 200);
  if (!text) throw err_('BAD', '公告內容不可為空');
  ss_().getSheetByName('Announcements').appendRow([Utilities.getUuid().slice(0, 8), text, new Date().toISOString()]);
}

function adminDelAnn_(u, p) {
  requireAdmin_(u);
  const a = rows_('Announcements').find((x) => x.id === p.id);
  if (a) ss_().getSheetByName('Announcements').deleteRow(a._row);
}

function adminSetPw_(u, p) {
  requireAdmin_(u);
  const pw = String(p.password || '').trim().slice(0, 50);
  if (!pw) throw err_('BAD', '密碼不可為空');
  const c = rows_('Config').find((r) => r.key === 'clubPassword');
  if (c) ss_().getSheetByName('Config').getRange(c._row, 2).setValue(pw);
  else ss_().getSheetByName('Config').appendRow(['clubPassword', pw]);
}

/* ================= 資料快照（回傳給前端） ================= */
function snapshot_(u) {
  const ws = weekStart_();
  const me = findUser_(u.userId);
  const data = {
    now: Date.now(),
    weekStart: ws,
    limit: LIMIT,
    registered: !!me,
    user: me ? { name: me.name, studentId: String(me.studentId) } : { name: u.displayName },
    isAdmin: isAdmin_(u),
    // 不回傳他人的 userId，只標示是否為自己
    bookings: rows_('Bookings').filter((b) => b.weekStart === ws)
      .map((b) => ({ day: Number(b.day), slot: Number(b.slot), name: b.name, mine: b.userId === u.userId })),
    fixed: rows_('Fixed').map((f) => ({ day: Number(f.day), slot: Number(f.slot), label: f.label })),
    announcements: rows_('Announcements').map((a) => ({ id: a.id, text: a.text })),
  };
  if (me) data.password = String(cfg_('clubPassword') || ''); // 僅已註冊社員可取得
  return data;
}

/* ================= 工具函式 ================= */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function err_(code, msg) { const e = new Error(msg); e.code = code; return e; }

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function rows_(name) {
  const v = ss_().getSheetByName(name).getDataRange().getValues();
  const head = v.shift();
  const out = [];
  v.forEach((r, i) => {
    if (r[0] === '' && r[1] === '') return;
    const o = { _row: i + 2 };
    head.forEach((k, j) => { o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], TZ, 'yyyy-MM-dd') : r[j]; });
    out.push(o);
  });
  return out;
}

function cfg_(key) { const r = rows_('Config').find((x) => x.key === key); return r ? r.value : ''; }
function findUser_(id) { return rows_('Users').find((r) => r.userId === id); }
function isAdmin_(u) { return rows_('Admins').some((r) => r.userId === u.userId); }
function requireUser_(u) { const m = findUser_(u.userId); if (!m) throw err_('REG', '請先完成註冊'); return m; }
function requireAdmin_(u) { if (!isAdmin_(u)) throw err_('ADMIN', '沒有管理員權限'); }
function checkCell_(day, slot) {
  if (!(day >= 1 && day <= 5) || !(slot >= 0 && slot < SLOT_STARTS.length) || day % 1 || slot % 1) throw err_('BAD', '時段參數錯誤');
}

/** 本週星期一 (yyyy-MM-dd, 台北時間) */
function weekStart_() {
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u')); // 1=Mon … 7=Sun
  return Utilities.formatDate(new Date(Date.now() - (dow - 1) * 864e5), TZ, 'yyyy-MM-dd');
}
/** 某週某天某節的開始時間（毫秒） */
function slotStart_(ws, day, slot) {
  return new Date(ws + 'T' + SLOT_STARTS[slot] + ':00+08:00').getTime() + (day - 1) * 864e5;
}

/** 以 LINE 驗證 ID Token，回傳 {userId, displayName}（快取 10 分鐘） */
function verify_(token) {
  if (!token) throw err_('AUTH', '尚未登入');
  const cache = CacheService.getScriptCache();
  const key = 't' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', muteHttpExceptions: true,
    payload: { id_token: token, client_id: String(cfg_('lineChannelId')) },
  });
  if (res.getResponseCode() !== 200) throw err_('AUTH', '登入已過期，請重新登入');
  const j = JSON.parse(res.getContentText());
  const user = { userId: j.sub, displayName: j.name || '' };
  cache.put(key, JSON.stringify(user), 600);
  return user;
}
