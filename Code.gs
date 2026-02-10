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

  // 1. 取得服務時長與緩衝時間
  const services = serviceSheet.getDataRange().getValues();
  const service = services.find(r => r[0] === serviceName);
  const duration = service ? parseInt(service[1]) : 60; // 預設 60 分鐘
  const buffer = 20; // 討論確定的 20 分鐘緩衝
  const totalNeedMinutes = duration + buffer;

  // 2. 取得所有按摩師資訊與總床位限制
  const staffData = staffSheet.getDataRange().getValues().slice(1);
  const maxBeds = 2; // 你設定的 2 張床限制

  const startTimeStr = "10:00"; // 假設營業開始時間
  const endTimeStr = "20:00";   // 假設營業結束時間
  const slots = [];
  
  // 3. 模擬時間軸，以 30 分鐘為一格進行掃描
  let currentPos = new Date(dateStr + "T" + startTimeStr);
  const endOfDay = new Date(dateStr + "T" + endTimeStr);

  while (currentPos.getTime() + (duration * 60000) <= endOfDay.getTime()) {
    const slotStart = new Date(currentPos);
    const slotEnd = new Date(slotStart.getTime() + (duration * 60000));
    const slotEndWithBuffer = new Date(slotStart.getTime() + (totalNeedMinutes * 60000));

    let availableTherapists = [];
    let busyBedsCount = 0;

    // 檢查每位按摩師在該時段的狀態
    staffData.forEach(row => {
      const name = row[0];
      const calId = row[1];
      const cal = CalendarApp.getCalendarById(calId);
      if (!cal) return;

      const events = cal.getEvents(slotStart, slotEndWithBuffer);
      
      // 判定邏輯：
      // A. 是否有「上班」事件覆蓋整個時段
      const isWorking = events.some(e => e.getTitle().includes("上班") && e.getStartTime() <= slotStart && e.getEndTime() >= slotEndWithBuffer);
      
      // B. 是否有其他「預約」事件衝突
      const hasBooking = events.some(e => !e.getTitle().includes("上班"));

      if (hasBooking) {
        busyBedsCount++; // 只要該按摩師有預約，就佔用一張床
      }

      if (isWorking && !hasBooking) {
        availableTherapists.push(name); // 沒預約且在上班的人才是「可選」
      }
    });

    // 4. 床位與人頭判定
    // 條件：床位沒滿 (busyBeds < 2) 且 (若是指定人則該人要有空，若不指定則至少一人要有空)
    const bedAvailable = busyBedsCount < maxBeds;
    let canBook = false;

    if (bedAvailable) {
      if (targetTherapist === "none") {
        canBook = availableTherapists.length > 0;
      } else {
        canBook = availableTherapists.includes(targetTherapist);
      }
    }

    if (canBook) {
      slots.push(Utilities.formatDate(slotStart, "GMT+8", "HH:mm"));
    }

    // 移動到下一個時段 (每 30 分鐘一跳)
    currentPos.setMinutes(currentPos.getMinutes() + 30);
  }

  return slots;
}

function processBooking(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
