/**
 * Веб-хук для приема данных из локального Python приложения.
 * Нужно развернуть как "Web App" с доступом "Anyone".
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = payload.targetSheet || "Import_Local";
    
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      // Добавляем заголовки, если лист новый
      sheet.appendRow(["Дата", "Номер", "Описание", "Сумма", "Источник"]);
    }
    
    const rows = payload.rows.map(r => [
      r.date,
      "'" + r.number, // Форсируем текст для номеров документов
      r.description,
      r.amount,
      payload.fileName
    ]);
    
    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Записано " + rows.length + " строк"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
