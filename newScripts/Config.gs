const CONFIG = {
  DEFAULT_SOURCE_FOLDER_ID: "1ExdP_KkK0ho9F5_HxvFUbSdCpgnyfrhr",
  DEFAULT_DEEPSEEK_API_URL: "https://api.deepseek.com/chat/completions",
  DEFAULT_DEEPSEEK_MODEL: "deepseek-chat",
  SAMPLE_HEADER_ROWS: 5,
  SAMPLE_DATA_ROWS: 50,
  OUTPUT_HEADERS: [
    "Дата",
    "Номер документа",
    "Поставщик",
    "Сумма",
    "Комментарий",
    "ТТ",
    "Источник",
    "Имя файла"
  ]
};

function getConfigValue(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value || fallback;
}
