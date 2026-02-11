function doGet() {
  // 必須使用 createTemplateFromFile 才能解析 HTML 內的腳本
  var template = HtmlService.createTemplateFromFile('Index');
  
  return template.evaluate()
      .setTitle('按摩預約系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 測試用：在 GAS 編輯器選取此函式執行，看 log 是否有資料
function testConnection() {
  console.log("Services:", getServices());
  console.log("Staff:", getTherapists());
}
// 修改後的 Code.gs 片段
function getServices() {
  try {
    // 請替換為你試算表網址中的那串長 ID
    const ss = SpreadsheetApp.openById("17ja21_3zV74XFxND4Avy52ZwSyMdjBitK33cBpmlLpw"); 
    const sheet = ss.getSheetByName("Services");
    const data = sheet.getDataRange().getValues();
    
    // 轉換為物件陣列並過濾標題列與空行
    return data.slice(1).filter(row => row[0]).map(row => {
      return { name: row[0].toString(), duration: row[1] };
    });
  } catch (e) {
    console.error("getServices Error: " + e.message);
    return [];
  }
}

function getTherapists() {
  try {
    const ss = SpreadsheetApp.openById("17ja21_3zV74XFxND4Avy52ZwSyMdjBitK33cBpmlLpw");
    const sheet = ss.getSheetByName("Staff");
    const data = sheet.getDataRange().getValues();
    return data.slice(1).filter(row => row[0]).map(row => {
      return { name: row[0].toString() };
    });
  } catch (e) {
    return [];
  }
}
// 先給予一個空殼，防止前端 fetchSlots 報錯 -> 更新成真正的3人2床計算
/**
 * 取得特定日期、項目、按摩師的可預約時段
 */
function getAvailableSlots(serviceName, targetTherapist, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceSheet = ss.getSheetByName("Services");
  const staffSheet = ss.getSheetByName("Staff");
  const settingsSheet = ss.getSheetByName("Settings");
  
  // 1. 讀取 Settings 資料
  const settingsData = settingsSheet.getDataRange().getValues();
  const startTimeStr = settingsData[1][0]; 
  const endTimeStr   = settingsData[1][1]; 
  const maxBeds      = settingsData[1][2]; 
  const buffer       = settingsData[1][3];

  // 2. 取得服務時長
  const services = serviceSheet.getDataRange().getValues();
  const service = services.find(r => r[0] === serviceName);
  const duration = service ? parseInt(service[1]) : 60; 
  const totalNeedMinutes = duration + buffer;

  // 3. 取得所有按摩師資訊
  const staffData = staffSheet.getDataRange().getValues().slice(1);
  const slots = [];

  // --- 【效能優化關鍵：批次抓取】 ---
  // 先定義當天的起點與終點 (以溫尼伯時區為準)
  const dayStart = new Date(dateStr.replace(/-/g, "/") + " 00:00:00");
  const dayEnd = new Date(dateStr.replace(/-/g, "/") + " 23:59:59");
  
  // 一次性將所有按摩師當天的行程抓進記憶體 (In-Memory)
  const allStaffCalendars = staffData.map(row => {
    const cal = CalendarApp.getCalendarById(row[1]);
    return {
      name: row[0],
      events: cal ? cal.getEvents(dayStart, dayEnd) : [] // 這裡只呼叫一次 API
    };
  });
  // ---------------------------------

  // 4. 設定掃描區間
  let currentPos = new Date(dateStr.replace(/-/g, "/") + " " + startTimeStr);
  const endOfDay = new Date(dateStr.replace(/-/g, "/") + " " + endTimeStr);
  const timeZone = ss.getSpreadsheetTimeZone();

  while (currentPos.getTime() + (duration * 60000) <= endOfDay.getTime()) {
    const slotStart = new Date(currentPos);
    const slotEndWithBuffer = new Date(slotStart.getTime() + (totalNeedMinutes * 60000));

    let availableTherapists = [];
    let busyBedsCount = 0;

    // 改用記憶體內的 allStaffCalendars 進行過濾，不再呼叫 Google API
    allStaffCalendars.forEach(staff => {
      // 如果有指定人且不是該人，則跳過
      if (targetTherapist !== "none" && targetTherapist !== staff.name) return;

      // 從預抓的 events 中找出與目前時段重疊的事件
      const currentEvents = staff.events.filter(e => {
        return e.getStartTime() < slotEndWithBuffer && e.getEndTime() > slotStart;
      });
      
      const workingEvent = currentEvents.find(e => 
        e.getTitle().includes("上班") && 
        e.getStartTime() <= slotStart && 
        e.getEndTime() >= slotEndWithBuffer
      );
      
      const hasBooking = currentEvents.some(e => !e.getTitle().includes("上班"));

      if (hasBooking) busyBedsCount++; 

      if (workingEvent && !hasBooking) {
        availableTherapists.push(staff.name);
      }
    });

    // 資源判定：有床且有人
    if (busyBedsCount < maxBeds && availableTherapists.length > 0) {
      slots.push(Utilities.formatDate(slotStart, timeZone, "HH:mm"));
    }

    currentPos.setMinutes(currentPos.getMinutes() + 30);
  }

  return slots;
}

function processBooking(data) {
  const ss = SpreadsheetApp.openById("17ja21_3zV74XFxND4Avy52ZwSyMdjBitK33cBpmlLpw");
  const bookingSheet = ss.getSheetByName("Bookings");
  const staffSheet = ss.getSheetByName("Staff");
  
  let finalTherapist = data.therapist;

  // 1. 如果是不指定，執行負載平衡分派
  if (finalTherapist === "none") {
    const staffList = staffSheet.getDataRange().getValues().slice(1).map(r => r[0]);
    // 簡單邏輯：計算今日每位師傅已有幾張單
    const todayBookings = bookingSheet.getDataRange().getValues()
      .filter(r => r[7] === data.date && r[10] !== "已取消");
    
    let minCount = Infinity;
    staffList.forEach(name => {
      const count = todayBookings.filter(r => r[5] === name).length;
      if (count < minCount) {
        minCount = count;
        finalTherapist = name;
      }
    });
  }

  // 2. 寫入 Google Calendar
  const therapistId = staffSheet.getDataRange().getValues().find(r => r[0] === finalTherapist)[1];
  const cal = CalendarApp.getCalendarById(therapistId);
  const start = new Date(data.date + "T" + data.time);
  // 取得服務時長
  const services = ss.getSheetByName("Services").getDataRange().getValues();
  const duration = services.find(r => r[0] === data.service)[1];
  const end = new Date(start.getTime() + (duration * 60000));
  
  const event = cal.createEvent(`[預約] ${data.name} - ${data.service}`, start, end, {
    description: `電話: ${data.phone}\n備註: ${data.note}`
  });

  // 3. 寫入 Sheet 資料庫
  bookingSheet.appendRow([
    "ID-" + new Date().getTime(), // 預約 ID
    new Date(),                   // 提交時間
    data.name,
    data.phone,
    data.email,
    finalTherapist,
    data.service,
    data.date,
    data.time,
    Utilities.formatDate(end, "GMT-6", "HH:mm"),
    "正常",
    event.getId()                 // 存下 Event ID 供未來修改/取消
  ]);

  // 4. 發送 Email 通知
  MailApp.sendEmail({
    to: data.email,
    subject: "按摩預約確認通知",
    body: `親愛的 ${data.name} 您好：\n\n您的預約已完成！\n日期：${data.date}\n時間：${data.time}\n項目：${data.service}\n按摩師：${finalTherapist}\n\n期待您的光臨。`
  });

  return { message: "預約成功！確認信已寄出。" };
}

// 生成 6 位驗證碼並寄送
function sendVerificationCode(phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName("Bookings");
  const data = bookingSheet.getDataRange().getValues();
  
  // 找尋最後一筆該電話的預約
  const booking = data.reverse().find(r => r[3] === phone && r[10] === "正常");
  if (!booking) return { success: false, msg: "查無此電話的有效預約" };
  
  // 檢查修改次數 (假設第 13 欄儲存修改次數)
  const editCount = booking[12] || 0;
  if (editCount >= 1) return { success: false, msg: "您已修改過一次，請致電工作室處理。" };

  const code = Math.floor(100000 + Math.random() * 900000).toString();
  PropertiesService.getScriptProperties().setProperty('VERIFY_' + phone, code);
  
  MailApp.sendEmail(booking[4], "您的預約修改驗證碼", "您的驗證碼為：" + code);
  return { success: true, msg: "驗證碼已寄至您的 Email" };
}

// 驗證並回傳預約資料進入修改流程
function verifyAndGetBooking(phone, inputCode) {
  const savedCode = PropertiesService.getScriptProperties().getProperty('VERIFY_' + phone);
  if (inputCode !== savedCode) return { success: false, msg: "驗證碼錯誤" };
  
  // 回傳該筆預約資料供前端填入
  // ...回傳邏輯...
}

/**
 * 1. 寄送驗證碼
 */
function sendVerificationCode(email, name) {
  const code = Math.floor(100000 + Math.random() * 900000).toString();
  // 存入腳本屬性，有效期約 10 分鐘
  PropertiesService.getScriptProperties().setProperty('VERIFY_' + email, code);
  
  const body = `親愛的 ${name} 您好：\n\n您的預約驗證碼為：${code}\n請於網頁輸入此代碼以完成預約。`;
  MailApp.sendEmail(email, "預約系統驗證碼", body);
  return "驗證碼已寄出";
}

/**
 * 2. 檢查是否為常客
function checkRegularCustomer(email, phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  const data = sheet.getDataRange().getValues();
  
  // 建議在比對前先過濾掉 phone 的橫線，確保比對準確
  const cleanPhone = phone.replace(/\D/g, "");
  
  const isRegular = data.some(row => {
    const sheetPhone = row[3].toString().replace(/\D/g, "");
    return sheetPhone === cleanPhone && row[4] === email && row[10] === "正常";
  });
  
  return { isRegular: isRegular };
}

/**
 * 3. 檢查修改次數 (用於修改流程)
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
function verifyCodeServerSide(email, inputCode) {
  const savedCode = PropertiesService.getScriptProperties().getProperty('VERIFY_' + email);
  return inputCode === savedCode;
}

/**
 * 搜尋預約資料並初步檢查修改權限
 */
function searchBookingForEdit(phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  const data = sheet.getDataRange().getValues();
  
  // 找出該電話最新一筆「正常」狀態的預約
  const booking = data.slice().reverse().find(r => r[3] === phone && r[10] === "正常");
  
  if (!booking) return { success: false, msg: "查無此電話的有效預約紀錄。" };

  // 檢查修改次數 (假設在第 13 欄，索引 12)
  const editCount = booking[12] || 0;
  if (editCount >= 1) {
    return { success: false, msg: "此預約已線上修改過 1 次。如需再次更動，請致電工作室處理。", limitReached: true };
  }

  // 隱藏部分 Email 資訊增加安全性並回傳
  const email = booking[4];
  const maskedEmail = email.replace(/(.{2})(.*)(?=@)/, (gp1, gp2, gp3) => gp2 + "*".repeat(gp3.length));
  
  return { 
    success: true, 
    bookingId: booking[0],
    name: booking[2],
    maskedEmail: maskedEmail,
    fullEmail: email // 僅供後端寄送驗證碼使用
  };
}

/**
 * 執行修改存檔 (更新 Sheet 與 Calendar)
 */
function updateBooking(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[0] === payload.bookingId);
  
  if (rowIndex === -1) return "找不到原始預約紀錄。";

  // 取得舊紀錄的修改次數並 +1
  const currentEditCount = parseInt(data[rowIndex][12] || 0);
  
  // 更新資料 (欄位索引需對應您的試算表結構)
  sheet.getRange(rowIndex + 1, 6).setValue(payload.therapist); // F 欄
  sheet.getRange(rowIndex + 1, 7).setValue(payload.service);   // G 欄
  sheet.getRange(rowIndex + 1, 8).setValue(payload.date);      // H 欄
  sheet.getRange(rowIndex + 1, 9).setValue(payload.time);      // I 欄
  sheet.getRange(rowIndex + 1, 13).setValue(currentEditCount + 1); // M 欄：修改次數

  return { success: true, message: "修改成功，已更新您的預約" };
}
/**
 * 每月自動執行的主程式
 */
function monthlyMaintenance() {
  clearAllExpiredProperties();
  archiveOldBookings();
}

/**
 * 任務 A: 清理所有過期的腳本屬性 (驗證碼)
 */
function clearAllExpiredProperties() {
  const props = PropertiesService.getScriptProperties();
  // 取得所有存儲的屬性 key
  const allKeys = props.getKeys();
  // 過濾出以 VERIFY_ 開頭的 key 並刪除
  const verifyKeys = allKeys.filter(key => key.indexOf('VERIFY_') === 0);
  
  if (verifyKeys.length > 0) {
    verifyKeys.forEach(key => props.deleteProperty(key));
    console.log("已清理驗證碼數量: " + verifyKeys.length);
  }
}

/**
 * 任務 B: 封存舊資料至 Archive 分頁
 */
function archiveOldBookings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Bookings");
  const targetSheet = ss.getSheetByName("Archive") || ss.insertSheet("Archive");
  
  // 如果 Archive 是新建立的，加上標題列
  if (targetSheet.getLastRow() === 0) {
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues();
    targetSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) return; // 只有標題列則跳過

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rowsToMove = [];
  const rowsToKeep = [data[0]]; // 保留標題列

  // 從第二列開始判斷
  for (let i = 1; i < data.length; i++) {
    const bookingDate = new Date(data[i][7]); // 假設日期在第 8 欄 (Index 7)
    const status = data[i][10];              // 假設狀態在第 11 欄 (Index 10)
    
    // 判斷準則：日期早於今天
    if (bookingDate < today) {
      rowsToMove.push(data[i]);
    } else {
      rowsToKeep.push(data[i]);
    }
  }

  // 如果有舊資料需要搬移
  if (rowsToMove.length > 0) {
    // 寫入 Archive
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length)
               .setValues(rowsToMove);
    
    // 清空原 Bookings 分頁並重新寫入需要保留的資料
    sourceSheet.clearContents();
    sourceSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length)
               .setValues(rowsToKeep);
               
    console.log("已搬移舊預約數量: " + rowsToMove.length);
  }
}

/**
 * 最後一步驗證：核對暫存在腳本屬性中的代碼
 */
function verifyCodeServerSide(email, inputCode) {
  try {
    const savedCode = PropertiesService.getScriptProperties().getProperty('VERIFY_' + email);
    if (!savedCode) return false;
    // 去除空白並轉為字串比對，增加穩定性
    return inputCode.toString().trim() === savedCode.toString().trim();
  } catch (e) {
    console.error("驗證碼核對出錯: " + e.message);
    return false;
  }
}
