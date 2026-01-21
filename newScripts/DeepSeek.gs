function buildSchemaPrompt(sampleRows, systemHint, fields) {
  const fieldMap = fields.reduce((acc, field) => {
    acc[field.key] = field.label;
    return acc;
  }, {});

  const outputFormat = fields.reduce((acc, field) => {
    acc[field.key] = 0;
    return acc;
  }, {});

  const payload = {
    task: "Определи колонки в таблице",
    system: systemHint,
    requirements: {
      headerRowIndex: "Индекс строки заголовка (1-based, в пределах sampleRows)",
      columns: fieldMap
    },
    output_format: {
      headerRowIndex: 1,
      columns: outputFormat
    },
    rules: [
      "Верни только JSON без пояснений.",
      "Если колонка отсутствует, верни 0.",
      "Если есть несколько заголовков, выбери строку, где видны названия колонок.",
      "Колонки должны соответствовать данным, а не пустым значениям."
    ],
    sampleRows: sampleRows
  };

  return JSON.stringify(payload, null, 2);
}

function callDeepSeek(prompt) {
  const apiKey = getConfigValue("DEEPSEEK_API_KEY", "");
  if (!apiKey) {
    throw new Error("Не задан DEEPSEEK_API_KEY в Script Properties.");
  }

  const apiUrl = getConfigValue("DEEPSEEK_API_URL", CONFIG.DEFAULT_DEEPSEEK_API_URL);
  const model = getConfigValue("DEEPSEEK_MODEL", CONFIG.DEFAULT_DEEPSEEK_MODEL);

  const payload = {
    model: model,
    temperature: 0,
    messages: [
      {
        role: "system",
        content: "Ты извлекаешь структуру таблиц. Возвращай только JSON."
      },
      {
        role: "user",
        content: prompt
      }
    ]
  };

  const response = UrlFetchApp.fetch(apiUrl, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const raw = response.getContentText();
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    throw new Error("DeepSeek вернул не JSON: " + raw);
  }

  const message = parsed.choices && parsed.choices[0] && parsed.choices[0].message;
  if (!message || !message.content) {
    throw new Error("DeepSeek ответ без content: " + raw);
  }

  return message.content;
}

function extractSchemaFromResponse(content) {
  const start = content.indexOf("{");
  const end = content.lastIndexOf("}");
  if (start === -1 || end === -1) {
    throw new Error("Не удалось найти JSON в ответе DeepSeek: " + content);
  }

  const jsonText = content.slice(start, end + 1);
  let schema;
  try {
    schema = JSON.parse(jsonText);
  } catch (err) {
    throw new Error("Ошибка парсинга JSON из DeepSeek: " + jsonText);
  }

  return schema;
}
