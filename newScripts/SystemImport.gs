function onOpenSystem() {
  SpreadsheetApp.getUi()
    .createMenu("Обработка")
    .addItem("Обработать учетные системы", "processSystemFiles")
    .addToUi();
}

function processSystemFiles() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const monthName = spreadsheet.getName();
  const sourceFolderId = getConfigValue(
    "SOURCE_FOLDER_ID",
    CONFIG.DEFAULT_SOURCE_FOLDER_ID
  );

  const sourceFolder = findMonthFolder(sourceFolderId, monthName);
  const filesIterator = sourceFolder.getFiles();

  while (filesIterator.hasNext()) {
    const file = filesIterator.next();
    const fileName = file.getName();
    if (!isExcelFile(fileName)) {
      continue;
    }

    const systemName = resolveSystemName(fileName);
    if (!systemName) {
      Logger.log("Не удалось определить систему по файлу: " + fileName);
      continue;
    }

    const convertedFileId = convertExcelToSheet(fileName, file.getId());
    try {
      const sourceSheet = SpreadsheetApp.openById(convertedFileId).getSheets()[0];
      const data = sourceSheet.getDataRange().getDisplayValues();
      if (!data.length) {
        Logger.log("Файл без данных: " + fileName);
        continue;
      }

      const schema = inferSchemaForSystem(data, systemName);
      const rows = buildOutputRows(data, schema, systemName, fileName);

      const targetSheet = ensureTargetSheet(spreadsheet, systemName);
      writeOutputRows(targetSheet, rows);
    } finally {
      DriveApp.getFileById(convertedFileId).setTrashed(true);
    }
  }
}

function inferSchemaForSystem(data, systemName) {
  const sampleRows = buildSampleRows(data);
  const prompt = buildSchemaPrompt(sampleRows, systemName);
  const response = callDeepSeek(prompt);
  const schema = extractSchemaFromResponse(response);

  if (!schema || !schema.docNumberCol || !schema.sumCol) {
    throw new Error(
      "DeepSeek не вернул обязательные колонки: " + JSON.stringify(schema)
    );
  }

  return schema;
}

function buildSampleRows(data) {
  const maxRows = Math.min(
    data.length,
    CONFIG.SAMPLE_HEADER_ROWS + CONFIG.SAMPLE_DATA_ROWS
  );
  return data.slice(0, maxRows);
}

function buildOutputRows(data, schema, systemName, fileName) {
  const startRowIndex = Math.max(schema.headerRowIndex || 1, 1);
  const rows = [];

  for (let i = startRowIndex; i < data.length; i += 1) {
    const row = data[i];
    if (row.join("").trim() === "") {
      continue;
    }

    const dateValue = getCellValue(row, schema.dateCol);
    const docNumber = getCellValue(row, schema.docNumberCol);
    const sumValue = normalizeSum(getCellValue(row, schema.sumCol));
    const partnerValue = getCellValue(row, schema.partnerCol);
    const commentValue = getCellValue(row, schema.commentCol);
    const ttValue = getCellValue(row, schema.ttCol);

    if (!docNumber && !sumValue) {
      continue;
    }

    rows.push([
      dateValue,
      docNumber,
      partnerValue,
      sumValue,
      commentValue,
      ttValue,
      systemName,
      fileName
    ]);
  }

  return rows;
}

function ensureTargetSheet(spreadsheet, systemName) {
  let sheet = spreadsheet.getSheetByName(systemName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(systemName);
    sheet.getRange(1, 1, 1, CONFIG.OUTPUT_HEADERS.length).setValues([
      CONFIG.OUTPUT_HEADERS
    ]);
  }

  return sheet;
}

function writeOutputRows(sheet, rows) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  if (!rows.length) {
    return;
  }

  const outputRange = sheet.getRange(2, 1, rows.length, rows[0].length);
  outputRange.setValues(rows);
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat("@");
}

function resolveSystemName(fileName) {
  const lowerName = fileName.toLowerCase();
  if (lowerName.indexOf("iiko") !== -1) {
    return "IIKO";
  }
  if (lowerName.indexOf("dxbx") !== -1) {
    return "DOCSINBOX";
  }
  if (lowerName.indexOf("sbis") !== -1) {
    return "SBIS";
  }
  if (lowerName.indexOf("sap") !== -1) {
    return "SAP";
  }

  return "";
}

function isExcelFile(fileName) {
  return /\.(xlsx?|xls)$/i.test(fileName);
}

function findMonthFolder(sourceFolderId, monthName) {
  const rootFolder = DriveApp.getFolderById(sourceFolderId);
  const folderIterator = rootFolder.getFoldersByName(monthName);
  if (!folderIterator.hasNext()) {
    throw new Error("Папка месяца не найдена: " + monthName);
  }

  return folderIterator.next();
}

function convertExcelToSheet(fileName, fileId) {
  const title = fileName.replace(/\.[^/.]+$/, "");
  const file = Drive.Files.copy({ title: title }, fileId, { convert: true });
  return file.id;
}

function normalizeSum(value) {
  if (!value) {
    return "";
  }
  return value.toString().replace(/\s/g, "").replace(",", ".");
}

function getCellValue(row, colIndex) {
  if (!colIndex || colIndex < 1) {
    return "";
  }
  return row[colIndex - 1] ? row[colIndex - 1].toString().trim() : "";
}
