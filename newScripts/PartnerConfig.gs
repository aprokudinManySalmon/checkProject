const PARTNER_CONFIG = {
  DEFAULT_PARTNER_ROOT_FOLDER_ID: "18_Zb8YpOY4Mj-QaotPDQCqeiZkOwajKV",
  PARTNER_SHEET_NAME: "PARTNER_IMPORT",
  DELETE_SOURCE_FILES: false,
  CLEAR_SHEET_BEFORE_RUN: true,
  USE_CLOUD_FUNCTION: true,
  FUNCTION_URL: "",
  USE_SEMANTIC_FILTER: true,
  SEMANTIC_PROVIDER: "yandex",
  SEMANTIC_BATCH_SIZE: 250,
  SEMANTIC_FAST_EXCLUDE: true,
  SEMANTIC_SEND_REASON: false,
  NUMBER_MODE: "regex_first",
  LLM_EXTRACT: true,
  LLM_MAX_CHARS: 120000,
  SEMANTIC_EXCLUDE_PATTERNS: [
    "оплата",
    "оплачено",
    "платеж",
    "платёж",
    "платежное поручение",
    "платёжное поручение",
    "п/п",
    "плат поруч",
    "банковская выписка",
    "поступление денежных средств"
  ],
  SAMPLE_HEADER_ROWS: 5,
  SAMPLE_DATA_ROWS: 50,
  OUTPUT_HEADERS: [
    "Дата",
    "Текст",
    "Номер",
    "Сумма"
  ]
};

function getPartnerConfigValue(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value || fallback;
}
