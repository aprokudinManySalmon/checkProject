function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Обработка")
    .addItem("Обработать", "processSystemFiles")
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
      const rows = buildOutputRows(data, schema, systemName);

      const targetSheet = ensureTargetSheet(spreadsheet, systemName);
      writeOutputRows(targetSheet, rows, systemName);
    } finally {
      DriveApp.getFileById(convertedFileId).setTrashed(true);
    }
  }
}

function inferSchemaForSystem(data, systemName) {
  const systemConfig = getSystemConfig(systemName);
  const sampleRows = buildSampleRows(data);
  const prompt = buildSchemaPrompt(sampleRows, systemName, systemConfig.fields);
  Logger.log("DeepSeek prompt (%s): %s", systemName, prompt);
  const response = callDeepSeek(prompt);
  Logger.log("DeepSeek response (%s): %s", systemName, response);
  const schema = extractSchemaFromResponse(response);
  Logger.log("DeepSeek schema (%s): %s", systemName, JSON.stringify(schema));

  if (!schema || !schema.columns) {
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

function buildOutputRows(data, schema, systemName) {
  const systemConfig = getSystemConfig(systemName);
  const startRowIndex = Math.max(schema.headerRowIndex || 1, 1);
  const columns = schema.columns || {};
  const rows = [];

  for (let i = startRowIndex; i < data.length; i += 1) {
    const row = data[i];
    if (row.join("").trim() === "") {
      continue;
    }

    const values = systemConfig.fields.map((field) => {
      const colIndex = columns[field.key] || 0;
      const rawValue = getCellValue(row, colIndex);
      if (field.type === "sum") {
        return normalizeSum(rawValue);
      }
      return rawValue;
    });

    if (values.join("").trim() === "") {
      continue;
    }

    rows.push(values);
  }

  return rows;
}

function ensureTargetSheet(spreadsheet, systemName) {
  const systemConfig = getSystemConfig(systemName);
  let sheet = spreadsheet.getSheetByName(systemName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(systemName);
    sheet.getRange(1, 1, 1, systemConfig.fields.length).setValues([
      systemConfig.fields.map((field) => field.label)
    ]);
  }

  return sheet;
}

function writeOutputRows(sheet, rows, systemName) {
  const systemConfig = getSystemConfig(systemName);
  const headerCount = systemConfig.fields.length;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headerCount).clearContent();
  }

  if (!rows.length) {
    return;
  }

  const outputRange = sheet.getRange(2, 1, rows.length, headerCount);
  outputRange.setValues(rows);
  const docFieldIndex = systemConfig.fields.findIndex(
    (field) => field.key === "docNumber"
  );
  if (docFieldIndex >= 0) {
    sheet
      .getRange(2, docFieldIndex + 1, rows.length, 1)
      .setNumberFormat("@");
  }
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
