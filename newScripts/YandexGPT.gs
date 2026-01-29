function callYandexGPT(prompt, options) {
  const config = getYandexConfig();
  const systemText =
    (options && options.systemText) ||
    "Ты помощник для классификации. Отвечай строго JSON без комментариев.";
  const maxTokens = (options && options.maxTokens) || 800;

  const payload = {
    modelUri: "gpt://" + config.folderId + "/" + config.model,
    completionOptions: {
      stream: false,
      temperature: 0,
      maxTokens: maxTokens
    },
    messages: [
      { role: "system", text: systemText },
      { role: "user", text: prompt }
    ]
  };

  const response = UrlFetchApp.fetch(
    "https://llm.api.cloud.yandex.net/foundationModels/v1/completion",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Api-Key " + config.apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error("YandexGPT error " + status + ": " + text);
  }

  const json = JSON.parse(text);
  if (
    !json ||
    !json.result ||
    !json.result.alternatives ||
    !json.result.alternatives.length ||
    !json.result.alternatives[0].message
  ) {
    throw new Error("YandexGPT response without message: " + text);
  }

  return json.result.alternatives[0].message.text;
}

function getYandexConfig() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("YANDEX_API_KEY");
  const folderId = props.getProperty("YANDEX_FOLDER_ID");
  const model = props.getProperty("YANDEX_MODEL") || "yandexgpt-lite/latest";

  if (!apiKey) {
    throw new Error("YANDEX_API_KEY not set in Script Properties.");
  }
  if (!folderId) {
    throw new Error("YANDEX_FOLDER_ID not set in Script Properties.");
  }

  return { apiKey: apiKey, folderId: folderId, model: model };
}
