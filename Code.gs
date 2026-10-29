const SHEET_ID = '你的_GOOGLE_SHEET_ID'; // 請替換為你的試算表ID

function getSheet(sheetName) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let result = {};

    if (action === 'init') {
      result = handleInit(params.userId);
    } else if (action === 'register') {
      result = handleRegister(params.userId, params.name, params.nickname);
    } else if (action === 'book') {
      result = handleBook(params.userId, params.day, params.timeIndex);
    } else if (action === 'cancel') {
      result = handleCancel(params.userId, params.day, params.timeIndex);
    } else if (action === 'adminUpdatePassword') {
      result = handleAdminUpdatePassword(params.userId, params.password);
    } else if (action === 'adminSetFixed') {
      result = handleAdminSetFixed(params.userId, params.day, params.timeIndex, params.label);
    } else if (action === 'adminAddAnnouncement') {
      result = handleAdminAddAnnouncement(params.userId, params.text);
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 取得初始資料
function handleInit(userId) {
  const user = getUser(userId);
  const bookings = getBookings();
  const announcements = getAnnouncements();
  
  if (!user) {
    return { registered: false, announcements };
  }
  
  const password = getPassword();
  return { registered: true, user, bookings, announcements, password };
}

function getUser(userId) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return { userId: data[i][0], name: data[i][1], nickname: data[i][2], role: data[i][3] };
    }
  }
  return null;
}

function handleRegister(userId, name, nickname) {
  const sheet = getSheet('Users');
  sheet.appendRow([userId, name, nickname, 'user']);
  return handleInit(userId); // 註冊後回傳完整資料
}

function handleBook(userId, day, timeIndex) {
  const user = getUser(userId);
  if (!user) return { success: false, error: 'User not found' };

  const sheet = getSheet('Bookings');
  const data = sheet.getDataRange().getValues();
  
  // 檢查本週預約數量 (上限3次)
  let userBookingsCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === userId && data[i][4] === 'normal') {
      userBookingsCount++;
    }
    if (data[i][0] == day && data[i][1] == timeIndex) {
      return { success: false, error: '該時段已被預約' };
    }
  }
  
  if (userBookingsCount >= 3) return { success: false, error: '每週最多只能預約 3 個時段' };
  
  sheet.appendRow([day, timeIndex, userId, user.nickname, 'normal', '']);
  return { success: true, bookings: getBookings() };
}

function handleCancel(userId, day, timeIndex) {
  const sheet = getSheet('Bookings');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    // 只能取消自己的 normal 預約
    if (data[i][0] == day && data[i][1] == timeIndex && data[i][2] === userId && data[i][4] === 'normal') {
      sheet.deleteRow(i + 1);
      return { success: true, bookings: getBookings() };
    }
  }
  return { success: false, error: '找不到可取消的預約' };
}

// === 管理員功能 ===
function isAdmin(userId) {
  const user = getUser(userId);
  return user && user.role === 'admin';
}

function handleAdminUpdatePassword(userId, password) {
  if (!isAdmin(userId)) return { success: false, error: '無權限' };
  const sheet = getSheet('Settings');
  sheet.getRange(2, 2).setValue(password);
  return { success: true, password: password };
}

function handleAdminSetFixed(userId, day, timeIndex, label) {
  if (!isAdmin(userId)) return { success: false, error: '無權限' };
  const sheet = getSheet('Bookings');
  const data = sheet.getDataRange().getValues();
  
  // 覆蓋或刪除舊有該時段設定
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == day && data[i][1] == timeIndex) {
      sheet.deleteRow(i + 1);
    }
  }
  
  if (label !== '') { // 若 label 不為空，則新增固定時段
    sheet.appendRow([day, timeIndex, 'ADMIN', '管理員', 'fixed', label]);
  }
  return { success: true, bookings: getBookings() };
}

function handleAdminAddAnnouncement(userId, text) {
  if (!isAdmin(userId)) return { success: false, error: '無權限' };
  const sheet = getSheet('Announcements');
  const id = new Date().getTime();
  sheet.appendRow([id, text]);
  return { success: true, announcements: getAnnouncements() };
}

// === 輔助讀取函數 ===
function getBookings() {
  const sheet = getSheet('Bookings');
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  for (let i = 1; i < data.length; i++) {
    bookings.push({
      day: data[i][0], timeIndex: data[i][1], userId: data[i][2], 
      nickname: data[i][3], type: data[i][4], label: data[i][5]
    });
  }
  return bookings;
}

function getAnnouncements() {
  const sheet = getSheet('Announcements');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({ id: row[0], text: row[1] }));
}

function getPassword() {
  const sheet = getSheet('Settings');
  return sheet.getRange(2, 2).getValue();
}

// === 每週排程清理 ===
// 請在 Apps Script 設定觸發條件 (Trigger)：每週日晚上執行此函式
function resetWeeklyBookings() {
  const sheet = getSheet('Bookings');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][4] === 'normal') {
      sheet.deleteRow(i + 1);
    }
  }
}
