// Code.gs
const SCRIPT_PROP = PropertiesService.getScriptProperties();

// ================= 1. 初始化與資料表建立 =================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Users': ['userId', 'name', 'nickname', 'role'], // role: user, admin
    'Bookings': ['day', 'timeSlot', 'userId', 'nickname'], // day: 1-5 (Mon-Fri)
    'FixedBookings': ['day', 'timeSlot', 'label'], // label: 樂器班 etc.
    'Announcements': ['id', 'text'],
    'Settings': ['key', 'value']
  };

  for (let name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
    }
  }
  
  // 預設設定
  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['roomPassword', '1234']);
    settingsSheet.appendRow(['adminList', '']); // 填入管理員的 userId，用逗號分隔
  }
}

// 每週日 23:59 觸發清空一般預約
function resetWeeklyBookings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  }
}

// 自動建立 Trigger (手動執行一次即可)
function createWeeklyTrigger() {
  ScriptApp.newTrigger('resetWeeklyBookings')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(23)
    .nearMinute(59)
    .create();
}

// ================= 2. API 路由 =================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;
    let result = {};

    switch (action) {
      case 'getInitData':
        result = getInitData(payload.userId);
        break;
      case 'register':
        result = registerUser(payload);
        break;
      case 'book':
        result = bookSlot(payload);
        break;
      case 'cancel':
        result = cancelSlot(payload);
        break;
      case 'adminUpdate':
        result = adminUpdate(payload);
        break;
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ================= 3. 核心邏輯 =================
function getInitData(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 讀取設定與公告
  const settingsData = ss.getSheetByName('Settings').getDataRange().getValues();
  let password = "", adminString = "";
  settingsData.forEach(row => {
    if(row[0] === 'roomPassword') password = row[1];
    if(row[0] === 'adminList') adminString = row[1];
  });
  
  const isAdmin = adminString.includes(userId);
  const announcements = ss.getSheetByName('Announcements').getDataRange().getValues().slice(1).map(r => r[1]);
  
  // 讀取使用者
  const users = ss.getSheetByName('Users').getDataRange().getValues();
  let user = users.find(r => r[0] === userId);
  
  // 讀取課表
  const bookings = ss.getSheetByName('Bookings').getDataRange().getValues().slice(1);
  const fixedBookings = ss.getSheetByName('FixedBookings').getDataRange().getValues().slice(1);
  
  let schedule = [];
  fixedBookings.forEach(r => schedule.push({ day: r[0], timeSlot: r[1], type: 'fixed', label: r[2] }));
  bookings.forEach(r => {
    if(r[0]) schedule.push({ day: r[0], timeSlot: r[1], type: 'user', userId: r[2], nickname: r[3] });
  });

  return {
    isRegistered: !!user,
    userData: user ? { userId: user[0], name: user[1], nickname: user[2], isAdmin: isAdmin } : null,
    password: user ? password : null,
    announcements: announcements,
    schedule: schedule
  };
}

function registerUser({ userId, name, nickname }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  sheet.appendRow([userId, name, nickname, 'user']);
  return { registered: true };
}

function bookSlot({ userId, nickname, day, timeSlot }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  const data = sheet.getDataRange().getValues();
  
  // 檢查是否超過3次 (每週)
  let userBookingsCount = data.filter(r => r[2] === userId).length;
  if (userBookingsCount >= 3) throw new Error("每週最多只能預約 3 個時段！");
  
  // 檢查是否已被預約
  let isBooked = data.some(r => r[0] == day && r[1] == timeSlot);
  if (isBooked) throw new Error("該時段已被預約！");

  sheet.appendRow([day, timeSlot, userId, nickname]);
  return { success: true };
}

function cancelSlot({ userId, day, timeSlot }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == day && data[i][1] == timeSlot && data[i][2] === userId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error("找不到您的預約或無權限刪除。");
}

function adminUpdate({ userId, type, data }) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminString = ss.getSheetByName('Settings').getDataRange().getValues().find(r => r[0] === 'adminList')[1];
  if (!adminString.includes(userId)) throw new Error("無管理員權限");

  if (type === 'password') {
    const sheet = ss.getSheetByName('Settings');
    const rows = sheet.getDataRange().getValues();
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === 'roomPassword') {
        sheet.getRange(i + 1, 2).setValue(data.newPassword);
        return { success: true };
      }
    }
  } else if (type === 'addAnnouncement') {
    ss.getSheetByName('Announcements').appendRow([new Date().getTime(), data.text]);
    return { success: true };
  } else if (type === 'setFixedSlot') {
    ss.getSheetByName('FixedBookings').appendRow([data.day, data.timeSlot, data.label]);
    return { success: true };
  }
}
