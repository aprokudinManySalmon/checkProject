function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Обработка")
    .addItem("Обработать сверку поставщика", "processPartnerFile")
    .addToUi();
}

function processPartnerFile() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rootFolderId = getPartnerConfigValue(
    "PARTNER_ROOT_FOLDER_ID",
    PARTNER_CONFIG.DEFAULT_PARTNER_ROOT_FOLDER_ID
  );
  if (!rootFolderId) {
    throw new Error("Не задан PARTNER_ROOT_FOLDER_ID в Script Properties.");
  }

  const partnerFolder = DriveApp.getFolderById(rootFolderId);
  const files = collectExcelFiles(partnerFolder);
  Logger.log("Partner files found: %s", files.length);
  const outputSheet = ensurePartnerSheet(spreadsheet);

  let processedAny = false;
  const allRows = [];
  for (let i = 0; i < files.length; i += 1) {
    const file = files[i];
    const fileName = file.getName();
    Logger.log("Processing partner file: %s", fileName);
    const convertedFileId = convertExcelToSheet(fileName, file.getId());
    let processed = false;
    try {
      const sourceSheet = SpreadsheetApp.openById(convertedFileId).getSheets()[0];
      const data = sourceSheet.getDataRange().getDisplayValues();
      if (!data.length) {
        Logger.log("Файл без данных: " + fileName);
        continue;
      }

      const schema = inferPartnerSchema(data, fileName);
      const rows = buildPartnerRows(data, schema, fileName);
      Logger.log("Rows extracted from %s: %s", fileName, rows.length);
      if (rows.length) {
        allRows.push.apply(allRows, rows);
      }
      processed = true;
      processedAny = true;
    } finally {
      DriveApp.getFileById(convertedFileId).setTrashed(true);
      if (processed && PARTNER_CONFIG.DELETE_SOURCE_FILES) {
        file.setTrashed(true);
      }
    }
  }

  if (!processedAny) {
    Logger.log("Подходящих файлов не найдено в папке: " + partnerFolder.getName());
    return;
  }

  if (!allRows.length) {
    Logger.log("Нет строк для записи после фильтрации.");
    return;
  }

  writePartnerRows(outputSheet, allRows);
}

function inferPartnerSchema(data, fileName) {
  const sampleRows = buildPartnerSampleRows(data);
  const headerRowIndex = findBestPartnerHeaderRowIndex(data);
  const prompt = buildPartnerSchemaPrompt(sampleRows, fileName);
  Logger.log("DeepSeek prompt (partner): %s", prompt);
  const response = callDeepSeek(prompt);
  Logger.log("DeepSeek response (partner): %s", response);
  const schema = extractSchemaFromResponse(response);
  if (headerRowIndex) {
    schema.headerRowIndex = headerRowIndex;
  }

  if (!schema || !schema.columns) {
    throw new Error("DeepSeek не вернул колонки партнера: " + response);
  }

  applyPartnerHeaderFallback(schema, data);
  Logger.log("DeepSeek schema (partner): %s", JSON.stringify(schema));
  return schema;
}

function buildPartnerSampleRows(data) {
  const maxRows = Math.min(
    data.length,
    PARTNER_CONFIG.SAMPLE_HEADER_ROWS + PARTNER_CONFIG.SAMPLE_DATA_ROWS
  );
  return data.slice(0, maxRows);
}

function buildPartnerRows(data, schema, fileName) {
  const columns = schema.columns || {};
  const startRowIndex = Math.max(schema.headerRowIndex || 1, 1);
  const rows = [];
  let skippedMissing = 0;
  let skippedType = 0;

  for (let i = startRowIndex; i < data.length; i += 1) {
    const row = data[i];
    if (row.join("").trim() === "") {
      continue;
    }

    const dateValue = getCellValue(row, columns.date);
    const sumValue = normalizeSum(getCellValue(row, columns.sum));
    const docName = getCellValue(row, columns.docName);
    const docNumberRaw = getCellValue(row, columns.docNumber);
    const docNumber = normalizeDocNumber(docNumberRaw || docName);
    const docType = detectDocType(docName);

    if (!dateValue || !sumValue || !docNumber) {
      skippedMissing += 1;
      continue;
    }
    if (!isAllowedDocType(docType, docName)) {
      skippedType += 1;
      continue;
    }

    rows.push([
      dateValue,
      docNumber,
      docType,
      sumValue,
      docName,
      fileName
    ]);
  }

  if (skippedMissing || skippedType) {
    Logger.log(
      "Skipped rows in %s (missing=%s, type=%s)",
      fileName,
      skippedMissing,
      skippedType
    );
  }

  return rows;
}

function ensurePartnerSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(PARTNER_CONFIG.PARTNER_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PARTNER_CONFIG.PARTNER_SHEET_NAME);
  }

  sheet
    .getRange(1, 1, 1, PARTNER_CONFIG.OUTPUT_HEADERS.length)
    .setValues([PARTNER_CONFIG.OUTPUT_HEADERS]);

  const lastColumn = sheet.getLastColumn();
  if (lastColumn > PARTNER_CONFIG.OUTPUT_HEADERS.length) {
    sheet
      .getRange(
        1,
        PARTNER_CONFIG.OUTPUT_HEADERS.length + 1,
        sheet.getMaxRows(),
        lastColumn - PARTNER_CONFIG.OUTPUT_HEADERS.length
      )
      .clearContent();
  }

  return sheet;
}

function writePartnerRows(sheet, rows) {
  if (PARTNER_CONFIG.CLEAR_SHEET_BEFORE_RUN) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet
        .getRange(2, 1, lastRow - 1, PARTNER_CONFIG.OUTPUT_HEADERS.length)
        .clearContent();
    }
  }

  if (!rows.length) {
    return;
  }

  const outputRange = sheet.getRange(
    2,
    1,
    rows.length,
    PARTNER_CONFIG.OUTPUT_HEADERS.length
  );
  outputRange.setValues(rows);
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat("@");
}

function buildPartnerSchemaPrompt(sampleRows, fileName) {
  const payload = {
    task: "Определи колонки сверки поставщика",
    fileName: fileName,
    requirements: {
      headerRowIndex: "Индекс строки заголовка (1-based)",
      columns: {
        date: "Колонка даты документа",
        docNumber: "Колонка номера документа (накладной/реализации)",
        docName: "Колонка описания/наименования документа",
        sum: "Колонка суммы"
      }
    },
    output_format: {
      headerRowIndex: 1,
      columns: {
        date: 0,
        docNumber: 0,
        docName: 0,
        sum: 0
      }
    },
    rules: [
      "Верни только JSON без пояснений.",
      "Если колонка отсутствует, верни 0.",
      "Строки с поступлением не использовать для определения нужных колонок."
    ],
    sampleRows: sampleRows
  };

  return JSON.stringify(payload, null, 2);
}

function applyPartnerHeaderFallback(schema, data) {
  const headerIndex = Math.max((schema.headerRowIndex || 1) - 1, 0);
  const headerRow = data[headerIndex] || [];
  const headerMap = buildHeaderIndexMap(headerRow);
  const columns = schema.columns || {};

  columns.date = columns.date || headerMap[normalizeHeader("Дата")] || 0;
  columns.sum = columns.sum || headerMap[normalizeHeader("Сумма")] || 0;
  columns.docNumber =
    columns.docNumber ||
    headerMap[normalizeHeader("Номер")] ||
    headerMap[normalizeHeader("№")] ||
    0;
  columns.docName =
    columns.docName ||
    headerMap[normalizeHeader("Документ")] ||
    headerMap[normalizeHeader("Наименование")] ||
    headerMap[normalizeHeader("Описание")] ||
    0;

  schema.columns = columns;
}

function findBestPartnerHeaderRowIndex(data) {
  const maxRows = Math.min(data.length, PARTNER_CONFIG.SAMPLE_HEADER_ROWS);
  const headerCandidates = [
    normalizeHeader("Дата"),
    normalizeHeader("Номер"),
    normalizeHeader("Сумма"),
    normalizeHeader("Документ"),
    normalizeHeader("Описание")
  ];
  let bestIndex = 0;
  let bestScore = 0;
  for (let i = 0; i < maxRows; i += 1) {
    const row = data[i] || [];
    let score = 0;
    row.forEach((cell) => {
      const normalized = normalizeHeader(cell);
      if (headerCandidates.indexOf(normalized) !== -1) {
        score += 1;
      }
    });
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i + 1;
    }
  }
  return bestIndex;
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

function normalizeDocNumber(value) {
  if (!value) {
    return "";
  }
  const text = value.toString().trim();
  if (!text) {
    return "";
  }
  const match = text.match(/[A-Za-zА-Яа-я0-9/-]{3,}/);
  return match ? match[0] : text;
}

function detectDocType(docName) {
  const text = docName.toString().toLowerCase();
  if (text.includes("коррект") || text.includes("исправ")) {
    return "Корректировка";
  }
  return "Реализация";
}

function isAllowedDocType(docType, docName) {
  const text = docName.toString().toLowerCase();
  if (text.includes("поступлен")) {
    return false;
  }
  return docType === "Реализация" || docType === "Корректировка";
}

function collectExcelFiles(rootFolder) {
  const stack = [rootFolder];
  const files = [];
  while (stack.length) {
    const folder = stack.pop();
    const fileIterator = folder.getFiles();
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      if (isExcelFile(file.getName())) {
        files.push(file);
      }
    }
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      stack.push(subfolders.next());
    }
  }
  return files;
}
