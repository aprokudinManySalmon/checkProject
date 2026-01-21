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
  applyPartnerPatternFallback(schema, data);
  schema.blocks = detectPartnerHeaderBlocks(data);
  if (schema.blocks.length) {
    Logger.log("Partner blocks detected: %s", JSON.stringify(schema.blocks));
  }
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

  if (schema.blocks && schema.blocks.length) {
    return buildPartnerRowsFromBlocks(data, schema, fileName);
  }

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
    if (!dateValue || !sumValue || !docNumber) {
      skippedMissing += 1;
      continue;
    }

    rows.push([
      dateValue,
      docName,
      docNumber,
      sumValue
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

  return applySemanticFilter(rows, fileName);
}

function buildPartnerRowsFromBlocks(data, schema, fileName) {
  const rows = [];
  let skippedMissing = 0;
  let skippedType = 0;
  const startRowIndex = Math.max(schema.headerRowIndex || 1, 1);
  const blocks = schema.blocks || [];

  for (let i = startRowIndex; i < data.length; i += 1) {
    const row = data[i];
    if (row.join("").trim() === "") {
      continue;
    }

    blocks.forEach((block) => {
      const dateValue = getCellValue(row, block.dateCol);
      const docName = getCellValue(row, block.docCol);
      const debitValue = normalizeSum(getCellValue(row, block.debitCol));
      const creditValue = normalizeSum(getCellValue(row, block.creditCol));
      const sumValue = debitValue || creditValue;
      const docNumber = normalizeDocNumber(docName);
      if (!dateValue || !sumValue || !docNumber) {
        skippedMissing += 1;
        return;
      }

      rows.push([
        dateValue,
        docName,
      docNumber,
      sumValue
      ]);
    });
  }

  if (skippedMissing || skippedType) {
    Logger.log(
      "Skipped rows in %s (missing=%s, type=%s)",
      fileName,
      skippedMissing,
      skippedType
    );
  }

  return applySemanticFilter(rows, fileName);
}

function applySemanticFilter(rows, fileName) {
  if (!PARTNER_CONFIG.USE_SEMANTIC_FILTER || !rows.length) {
    return rows;
  }

  const batchSize = PARTNER_CONFIG.SEMANTIC_BATCH_SIZE || 30;
  Logger.log("Semantic batch size: %s", batchSize);
  const filtered = [];
  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    Logger.log(
      "Semantic batch %s: rows %s-%s",
      Math.floor(i / batchSize) + 1,
      i + 1,
      Math.min(i + batchSize, rows.length)
    );
    const decisions = classifyPartnerRows(batch, fileName);
    Logger.log(
      "Semantic decisions: %s",
      decisions.filter((d) => d && d.include).length
    );
    decisions.forEach((decision, idx) => {
      if (decision && decision.include) {
        filtered.push(batch[idx]);
      }
    });
  }

  Logger.log(
    "Semantic filter kept %s/%s rows for %s",
    filtered.length,
    rows.length,
    fileName
  );
  return filtered;
}

function classifyPartnerRows(rows, fileName) {
  const payload = {
    task: "Классификация строк сверки поставщика",
    fileName: fileName,
    labels: {
      include:
        "Оставить строку: отражает расход клиента (реализация, отгрузка) или корректировку",
      exclude:
        "Исключить строку: платежи, оплаты, банковские поручения, поступления денег"
    },
    rules: [
      "Не используй позицию суммы (дебет/кредит) как единственный признак.",
      "Фокус на смысле текста документа.",
      "Верни только JSON."
    ],
    rows: rows.map((row, index) => ({
      id: index,
      date: row[0],
      text: row[1],
      number: row[2],
      sum: row[3]
    })),
    output_format: [
      { id: 0, include: true, reason: "..." }
    ]
  };

  const prompt = JSON.stringify(payload, null, 2);
  Logger.log("Semantic request rows: %s for %s", rows.length, fileName);
  const started = Date.now();
  const response = callDeepSeek(prompt);
  Logger.log("Semantic request ms: %s for %s", Date.now() - started, fileName);
  const parsed = extractSemanticArray(response);
  if (!Array.isArray(parsed)) {
    throw new Error("DeepSeek вернул не массив для семантики: " + response);
  }

  const decisions = rows.map(() => ({ include: false }));
  parsed.forEach((item) => {
    if (item && typeof item.id === "number") {
      decisions[item.id] = { include: !!item.include };
    }
  });

  return decisions;
}

function extractSemanticArray(content) {
  const startBracket = content.indexOf("[");
  const endBracket = content.lastIndexOf("]");
  if (startBracket !== -1 && endBracket !== -1) {
    const arrayText = content.slice(startBracket, endBracket + 1);
    try {
      return JSON.parse(arrayText);
    } catch (err) {
      Logger.log("Не удалось распарсить массив, пробуем объект: %s", err);
    }
  }

  const startObject = content.indexOf("{");
  const endObject = content.lastIndexOf("}");
  if (startObject !== -1 && endObject !== -1) {
    const objectText = content.slice(startObject, endObject + 1);
    try {
      const obj = JSON.parse(objectText);
      if (Array.isArray(obj)) {
        return obj;
      }
      if (Array.isArray(obj.items)) {
        return obj.items;
      }
    } catch (err) {
      Logger.log("Не удалось распарсить объект: %s", err);
    }
  }

  return null;
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

function applyPartnerPatternFallback(schema, data) {
  const columns = schema.columns || {};
  const startRowIndex = Math.max(schema.headerRowIndex || 1, 1);
  const maxRows = Math.min(data.length, startRowIndex + 200);

  if (!columns.date) {
    columns.date = detectDateColumn(data, startRowIndex, maxRows);
  }
  if (!columns.sum) {
    columns.sum = detectSumColumn(data, startRowIndex, maxRows);
  }
  if (!columns.docName) {
    columns.docName = detectDocNameColumn(data, startRowIndex, maxRows);
  }

  schema.columns = columns;
}

function detectDateColumn(data, startRowIndex, endRowIndex) {
  return detectColumnByPattern(
    data,
    startRowIndex,
    endRowIndex,
    /^\d{1,2}[./]\d{1,2}[./]\d{2,4}$/
  );
}

function detectSumColumn(data, startRowIndex, endRowIndex) {
  return detectColumnByPattern(
    data,
    startRowIndex,
    endRowIndex,
    /^-?\d+([ \u00A0]\d{3})*(?:[.,]\d+)?$/
  );
}

function detectDocNameColumn(data, startRowIndex, endRowIndex) {
  return detectColumnByPattern(
    data,
    startRowIndex,
    endRowIndex,
    /(реализац|коррект|исправ|накладн|упд)/i
  );
}

function detectColumnByPattern(data, startRowIndex, endRowIndex, pattern) {
  let bestIndex = 0;
  let bestScore = 0;
  const columnCount = data[0] ? data[0].length : 0;
  for (let col = 0; col < columnCount; col += 1) {
    let score = 0;
    for (let row = startRowIndex; row < endRowIndex; row += 1) {
      const value = getCellValue(data[row] || [], col + 1);
      if (value && pattern.test(value.toString().trim())) {
        score += 1;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestIndex = col + 1;
    }
  }
  return bestIndex;
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

function detectPartnerHeaderBlocks(data) {
  const maxRows = Math.min(data.length, 20);
  let headerRowIndex = 0;

  for (let i = 0; i < maxRows; i += 1) {
    const row = data[i] || [];
    const normalized = row.map((cell) => normalizeHeader(cell));
    if (
      normalized.indexOf("дата") !== -1 &&
      normalized.indexOf("документ") !== -1 &&
      (normalized.indexOf("дебет") !== -1 || normalized.indexOf("кредит") !== -1)
    ) {
      headerRowIndex = i + 1;
      break;
    }
  }

  if (!headerRowIndex) {
    return [];
  }

  const headerRow = data[headerRowIndex - 1] || [];
  const blocks = [];
  for (let i = 0; i < headerRow.length; i += 1) {
    if (normalizeHeader(headerRow[i]) !== "дата") {
      continue;
    }
    const docCol = findHeaderOffset(headerRow, i + 1, "документ");
    const debitCol = findHeaderOffset(headerRow, i + 1, "дебет");
    const creditCol = findHeaderOffset(headerRow, i + 1, "кредит");
    if (docCol && (debitCol || creditCol)) {
      blocks.push({
        headerRowIndex: headerRowIndex,
        dateCol: i + 1,
        docCol: docCol,
        debitCol: debitCol,
        creditCol: creditCol
      });
    }
  }

  return blocks;
}

function findHeaderOffset(headerRow, startIndex, headerName) {
  for (let i = startIndex; i < headerRow.length; i += 1) {
    if (normalizeHeader(headerRow[i]) === headerName) {
      return i + 1;
    }
  }
  return 0;
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
  const withNumber = text.match(/№\s*([A-Za-zА-Яа-я0-9/-]+)/);
  if (withNumber && withNumber[1]) {
    return withNumber[1];
  }
  const digits = text.match(/\b\d{2,}\b/);
  if (digits) {
    return digits[0];
  }
  const fallback = text.match(/[A-Za-zА-Яа-я0-9/-]{3,}/);
  return fallback ? fallback[0] : text;
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
