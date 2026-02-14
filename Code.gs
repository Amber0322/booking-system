// @ts-nocheck
/**
 * 按摩預約系統 - 後端核心邏輯 (Code.gs)
 * 適用地區：加拿大 溫尼伯 (Winnipeg)
 */

// --- 全域設定 ---
const SS_ID = "17ja21_3zV74XFxND4Avy52ZwSyMdjBitK33cBpmlLpw"; // 您的試算表 ID

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('按摩預約系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 1. 基礎資料讀取
 */
// 基礎資料讀取
function getServices() {
  const data = SpreadsheetApp.openById(SS_ID).getSheetByName("Services").getDataRange().getValues();
  return data.slice(1).filter(row => row[0]).map(row => ({ name: row[0].toString(), duration: row[1] }));
}

function getTherapists() {
  const data = SpreadsheetApp.openById(SS_ID).getSheetByName("Staff").getDataRange().getValues();
  return data.slice(1).filter(row => row[0]).map(row => ({ name: row[0].toString() }));
}

/**
 * 2. 核心：3人2床時段計算引擎
 */
function getAvailableSlots(serviceName, targetTherapist, dateStr) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const settings = ss.getSheetByName("Settings").getDataRange().getValues()[1];
  const startTimeStr = settings[0]; // A2
  const endTimeStr   = settings[1]; // B2
  const maxBeds      = settings[2]; // C2: 2
  const buffer       = settings[3]; // D2: 20

  const service = ss.getSheetByName("Services").getDataRange().getValues().find(r => r[0] === serviceName);
  const duration = service ? parseInt(service[1]) : 60;
  const totalNeedMinutes = duration + buffer;

  const staffData = ss.getSheetByName("Staff").getDataRange().getValues().slice(1);
  const timeZone = ss.getSpreadsheetTimeZone(); // 自動適應溫尼伯時區

  // 效能優化：一次性抓取所有按摩師當天行程
  const dayStart = new Date(dateStr.replace(/-/g, "/") + " 00:00:00");
  const dayEnd = new Date(dateStr.replace(/-/g, "/") + " 23:59:59");
  const allStaffCals = staffData.map(row => ({
    name: row[0],
    events: CalendarApp.getCalendarById(row[1]) ? CalendarApp.getCalendarById(row[1]).getEvents(dayStart, dayEnd) : []
  }));

  let currentPos = new Date(dateStr.replace(/-/g, "/") + " " + startTimeStr);
  const endLimit = new Date(dateStr.replace(/-/g, "/") + " " + endTimeStr);
  const slots = [];

  while (currentPos.getTime() + (duration * 60000) <= endLimit.getTime()) {
    const slotStart = new Date(currentPos);
    const slotEndWithBuffer = new Date(slotStart.getTime() + (totalNeedMinutes * 60000));
    
    let availableStaff = [];
    let busyBeds = 0;

    allStaffCals.forEach(staff => {
      if (targetTherapist !== "none" && targetTherapist !== staff.name) return;

      const currentEvents = staff.events.filter(e => e.getStartTime() < slotEndWithBuffer && e.getEndTime() > slotStart);
      const isWorking = currentEvents.some(e => e.getTitle().includes("上班") && e.getStartTime() <= slotStart && e.getEndTime() >= slotEndWithBuffer);
      const hasBooking = currentEvents.some(e => !e.getTitle().includes("上班"));

      if (hasBooking) busyBeds++;
      if (isWorking && !hasBooking) availableStaff.push(staff.name);
    });

    if (busyBeds < maxBeds && availableStaff.length > 0) {
      slots.push(Utilities.formatDate(slotStart, timeZone, "HH:mm"));
    }
    currentPos.setMinutes(currentPos.getMinutes() + 30);
  }
  return slots;
}

/**
 * 3. 驗證與常客邏輯
 */
function sendVerificationCode(email, name) {
  const code = Math.floor(100000 + Math.random() * 900000).toString();
  PropertiesService.getScriptProperties().setProperty('VERIFY_' + email, code);
  const body = `親愛的 ${name} 您好：\n\n您的預約驗證碼為：${code}\n請於網頁輸入此代碼以完成操作。`;
  MailApp.sendEmail(email, "預約系統驗證碼", body);
  return "驗證碼已寄出";
}

function verifyCodeServerSide(email, inputCode) {
  const saved = PropertiesService.getScriptProperties().getProperty('VERIFY_' + email);
  return inputCode === saved;
}

function checkRegularCustomer(email, phone) {
  const data = SpreadsheetApp.openById(SS_ID).getSheetByName("Bookings").getDataRange().getValues();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  // 檢查是否有 姓名+Email 且狀態為「正常」的紀錄
  const isRegular = data.some(row => row[7] === phone && row[3] === email && row[4] === "正常");
  return { isRegular: isRegular };
}

/**
 * 3.1 檢查修改次數 (用於修改流程)
 */
function checkEditLimit(bookingId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  const data = sheet.getDataRange().getValues();
  const row = data.find(r => r[0] === bookingId);
  
  // 假設第 13 欄 (索引 12) 是修改次數
  const editCount = row[12] || 0;
  return editCount < 1; // 只能修改一次
}
/**
 * 4. 提交與修改邏輯
 */
function processBooking(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const bookingSheet = ss.getSheetByName("Bookings");
  const staffSheet = ss.getSheetByName("Staff");
  
  let finalTherapist = data.therapist;
  const staffRows = staffSheet.getDataRange().getValues();

  if (finalTherapist === "none") {
    const therapists = staffRows.slice(1).map(r => r[0]);
    const todayBookings = bookingSheet.getDataRange().getValues().filter(r => r[7] === data.date && r[10] !== "已取消");
    finalTherapist = therapists.sort((a,b) => todayBookings.filter(r => r[5] === a).length - todayBookings.filter(r => r[5] === b).length)[0];
  }

  const calId = staffRows.find(r => r[0] === finalTherapist)[1];
  const service = ss.getSheetByName("Services").getDataRange().getValues().find(r => r[0] === data.service);
  const start = new Date(data.date.replace(/-/g, "/") + " " + data.time);
  const end = new Date(start.getTime() + (service[1] * 60000));

  const event = CalendarApp.getCalendarById(calId).createEvent(`[預約] ${data.name} - ${data.service}`, start, end, {description: `電話: ${data.phone}\n備註: ${data.note}`});

  bookingSheet.appendRow([
    "ID-" + new Date().getTime(), new Date(), data.name, data.phone, data.email, 
    finalTherapist, data.service, data.date, data.time, Utilities.formatDate(end, ss.getSpreadsheetTimeZone(), "HH:mm"),
    "正常", event.getId(), 0 // 修改次數初始化為 0
  ]);

  MailApp.sendEmail(data.email, "預約成功通知", `您已預約成功！\n日期：${data.date}\n時間：${data.time}\n師傅：${finalTherapist}`);
  return { success: true };
}

/**
 * 搜尋預約資料：支援回傳多筆紀錄 (處理夫妻/同行者情境)，搜尋預約資料：增加錯誤捕捉與 null 值保護
 */
// 核心：搜尋預約 (增加強效比對)
function searchBookingForEdit(phoneDigits) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName("Bookings");
    if (!sheet) return { success: false, msg: "系統錯誤：找不到名為 'Bookings' 的工作表。" };

    const data = sheet.getDataRange().getValues();
    const searchClean = String(phoneDigits).replace(/\D/g, "").trim();
    
    // 過濾資料
    const results = data.slice(1).filter(row => {
      if (!row[3]) return false;
      const sheetPhone = String(row[3]).replace(/\D/g, "").trim();
      const sheetStatus = String(row[10] || "").trim();
      return sheetPhone === searchClean && sheetStatus === "正常";
    }).map(row => ({
      bookingId: row[0],
      name: row[2],
      service: row[6],
      date: row[7] instanceof Date ? Utilities.formatDate(row[7], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : String(row[7]),
      time: row[8],
      email: row[4],
      phone: row[3],
      editCount: parseInt(row[12]) || 0
    }));

    if (results.length === 0) {
      return { success: false, msg: "查無此電話(" + searchClean + ")的預約紀錄。" };
    }

    return { success: true, bookings: results, maskedEmail: results[0].email.replace(/(.{2})(.*)(?=@)/, (g1, g2, g3) => g2 + "*".repeat(g3.length)), fullEmail: results[0].email, name: results[0].name };

  } catch (e) {
    // 萬一出錯，強制回傳一個物件，防止前端接到 null
    return { success: false, msg: "程式執行出錯，請致電工作室：" + e.message };
  }
}


/**同步更新google 日曆 */
function updateBooking(payload) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName("Bookings");
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(r => r[0] === payload.bookingId);
    
    if (rowIndex === -1) return { success: false, message: "找不到紀錄" };

    const oldEventId = data[rowIndex][11]; // 取得原本的 Calendar Event ID
    const currentCount = parseInt(data[rowIndex][12] || 0);
    const timeZone = ss.getSpreadsheetTimeZone();

    // 1. 更新試算表資料 (師傅, 項目, 日期, 時間)
    // 注意：這裡假設 Service 執行時間沒變，若需考慮不同 Service 需另取 duration
    sheet.getRange(rowIndex + 1, 6, 1, 4).setValues([[payload.therapist, payload.service, payload.date, payload.time]]);
    sheet.getRange(rowIndex + 1, 13).setValue(currentCount + 1); // 修改次數 +1

    // 2. 更新 Google 日曆 (連動修正)
    const staffSheet = ss.getSheetByName("Staff");
    const staffRows = staffSheet.getDataRange().getValues();
    const calId = staffRows.find(r => r[0] === payload.therapist)[1];
    
    if (calId && oldEventId) {
      const calendar = CalendarApp.getCalendarById(calId);
      const event = calendar.getEventById(oldEventId);
      if (event) {
        // 重新計算結束時間
        const serviceData = ss.getSheetByName("Services").getDataRange().getValues().find(r => r[0] === payload.service);
        const duration = serviceData ? parseInt(serviceData[1]) : 60;
        const newStart = new Date(payload.date.replace(/-/g, "/") + " " + payload.time);
        const newEnd = new Date(newStart.getTime() + (duration * 60000));
        
        event.setTitle(`[修改] ${payload.name} - ${payload.service}`);
        event.setTime(newStart, newEnd);
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, message: "更新失敗: " + e.message };
  }
}

/**
 * 5. 每月自動維護
 */
function monthlyMaintenance() {
  const props = PropertiesService.getScriptProperties();
  props.getKeys().filter(k => k.startsWith('VERIFY_')).forEach(k => props.deleteProperty(k));
  
  const ss = SpreadsheetApp.openById(SS_ID);
  const source = ss.getSheetByName("Bookings");
  const archive = ss.getSheetByName("Archive") || ss.insertSheet("Archive");
  const data = source.getDataRange().getValues();
  const today = new Date(); 
  today.setHours(0,0,0,0);

  const toKeep = [data[0]], toArchive = [];

  data.slice(1).forEach(row => {
    if (new Date(row[7]) < today) {
      toArchive.push(row);
    } else {
      toKeep.push(row);
    }
  });

  if (toArchive.length > 0) {
    if (archive.getLastRow() === 0) {
      archive.appendRow(data[0]);
    }

    archive
      .getRange(archive.getLastRow() + 1, 1, toArchive.length, toArchive[0].length)
      .setValues(toArchive);

    source.clear();
    source
      .getRange(1, 1, toKeep.length, toKeep[0].length)
      .setValues(toKeep);
  }
}
