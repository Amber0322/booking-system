// 取得試算表中的服務項目
function getServices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Services"); // 確保分頁名稱正確
  const data = sheet.getDataRange().getValues();
  
  // 假設表頭是：項目名稱 | 時長 | 價格
  // 跳過第一行表頭，回傳物件陣列
  return data.slice(1).map(row => {
    return { name: row[0], duration: row[1] };
  });
}

// 取得試算表中的按摩師名單
function getTherapists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Staff"); // 確保分頁名稱正確
  const data = sheet.getDataRange().getValues();
  
  // 假設表頭是：姓名 | Calendar_ID
  return data.slice(1).map(row => {
    return { name: row[0] };
  });
}
