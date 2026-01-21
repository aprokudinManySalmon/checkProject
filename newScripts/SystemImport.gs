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
    let processed = false;
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
      processed = true;
    } finally {
      DriveApp.getFileById(convertedFileId).setTrashed(true);
      if (processed && CONFIG.DELETE_SOURCE_FILES) {
        file.setTrashed(true);
      }
    }
  }
}

function inferSchemaForSystem(data, systemName) {
  const systemConfig = getSystemConfig(systemName);
  const headerRowIndex = findBestHeaderRowIndex(data, systemConfig.fields);
  const sampleRows = buildSampleRows(data);
  const prompt = buildSchemaPrompt(sampleRows, systemName, systemConfig.fields);
  Logger.log("DeepSeek prompt (%s): %s", systemName, prompt);
  const response = callDeepSeek(prompt);
  Logger.log("DeepSeek response (%s): %s", systemName, response);
  const schema = extractSchemaFromResponse(response);
  if (headerRowIndex) {
    schema.headerRowIndex = headerRowIndex;
  }
  applyHeaderFallback(schema, data, systemName);
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
    if (field.type === "date") {
      return normalizeDateValue(rawValue);
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

function applyHeaderFallback(schema, data, systemName) {
  const systemConfig = getSystemConfig(systemName);
  const headerIndex = Math.max((schema.headerRowIndex || 1) - 1, 0);
  const headerRow = data[headerIndex] || [];
  const headerMap = buildHeaderIndexMap(headerRow);
  const columns = schema.columns || {};

  systemConfig.fields.forEach((field) => {
    const exactIndex = getHeaderExactIndex(headerMap, field);
    const aliasIndex = getHeaderAliasIndex(headerMap, field);
    const matchIndex = exactIndex || aliasIndex;
    if (!matchIndex) {
      return;
    }
    const currentIndex = columns[field.key] || 0;
    if (!currentIndex) {
      columns[field.key] = matchIndex;
      return;
    }
    if (exactIndex && !headerMatchesExact(headerRow[currentIndex - 1], field)) {
      columns[field.key] = exactIndex;
      return;
    }
    if (!exactIndex && !headerMatchesField(headerRow[currentIndex - 1], field)) {
      columns[field.key] = matchIndex;
    }
  });

  schema.columns = columns;
}

function findBestHeaderRowIndex(data, fields) {
  const maxRows = Math.min(data.length, CONFIG.SAMPLE_HEADER_ROWS);
  let bestIndex = 0;
  let bestScore = 0;
  for (let i = 0; i < maxRows; i += 1) {
    const headerMap = buildHeaderIndexMap(data[i] || []);
    const score = countHeaderMatches(headerMap, fields);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i + 1;
    }
  }
  return bestIndex;
}

function countHeaderMatches(headerMap, fields) {
  let score = 0;
  fields.forEach((field) => {
    if (getHeaderExactIndex(headerMap, field) || getHeaderAliasIndex(headerMap, field)) {
      score += 1;
    }
  });
  return score;
}

function getHeaderExactIndex(headerMap, field) {
  const normalized = normalizeHeader(field.label);
  return headerMap[normalized] || 0;
}

function getHeaderAliasIndex(headerMap, field) {
  const aliases = field.aliases || [];
  for (let i = 0; i < aliases.length; i += 1) {
    const normalized = normalizeHeader(aliases[i]);
    const matchIndex = headerMap[normalized];
    if (matchIndex) {
      return matchIndex;
    }
  }
  return 0;
}

function headerMatchesField(headerValue, field) {
  const candidates = [field.label].concat(field.aliases || []);
  const normalizedHeader = normalizeHeader(headerValue);
  for (let i = 0; i < candidates.length; i += 1) {
    if (normalizedHeader === normalizeHeader(candidates[i])) {
      return true;
    }
  }
  return false;
}

function headerMatchesExact(headerValue, field) {
  return normalizeHeader(headerValue) === normalizeHeader(field.label);
}

function buildHeaderIndexMap(headerRow) {
  const map = {};
  headerRow.forEach((cell, index) => {
    const normalized = normalizeHeader(cell);
    if (normalized) {
      map[normalized] = index + 1;
    }
  });
  return map;
}

function normalizeHeader(value) {
  if (!value) {
    return "";
  }
  return value
    .toString()
    .toLowerCase()
    .replace(/[«»"']/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function ensureTargetSheet(spreadsheet, systemName) {
  const systemConfig = getSystemConfig(systemName);
  let sheet = spreadsheet.getSheetByName(systemName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(systemName);
  }

  const headerValues = systemConfig.fields.map((field) => field.label);
  sheet.getRange(1, 1, 1, headerValues.length).setValues([headerValues]);
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > headerValues.length) {
    sheet
      .getRange(1, headerValues.length + 1, sheet.getMaxRows(), lastColumn - headerValues.length)
      .clearContent();
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

  const textColumnIndexes = systemConfig.fields
    .map((field, index) => ({ field, index }))
    .filter(
      ({ field }) => field.key === "docNumber" || field.type === "date"
    )
    .map(({ index }) => index + 1);
  textColumnIndexes.forEach((colIndex) => {
    sheet.getRange(2, colIndex, rows.length, 1).setNumberFormat("@");
  });
  const outputRange = sheet.getRange(2, 1, rows.length, headerCount);
  outputRange.setValues(rows);
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

function normalizeDateValue(value) {
  if (!value) {
    return "";
  }
  const trimmed = value.toString().trim();
  if (!trimmed) {
    return "";
  }
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(trimmed)) {
    return trimmed;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const parts = trimmed.split("-");
    return [parts[2], parts[1], parts[0]].join(".");
  }
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) {
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(
        parsed,
        Session.getScriptTimeZone(),
        "dd.MM.yyyy"
      );
    }
  }
  return trimmed;
}

function getCellValue(row, colIndex) {
  if (!colIndex || colIndex < 1) {
    return "";
  }
  return row[colIndex - 1] ? row[colIndex - 1].toString().trim() : "";
}
