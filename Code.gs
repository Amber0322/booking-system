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
// 先給予一個空殼，防止前端 fetchSlots 報錯
function getAvailableSlots(service, therapist, date) {
  // 這部分我們下一階段會寫入 3人2床 邏輯
  return ["09:00", "10:30", "13:00", "15:00"]; 
}
