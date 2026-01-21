const PARTNER_CONFIG = {
  DEFAULT_PARTNER_ROOT_FOLDER_ID: "18_Zb8YpOY4Mj-QaotPDQCqeiZkOwajKV",
  PARTNER_SHEET_NAME: "PARTNER_IMPORT",
  DELETE_SOURCE_FILES: false,
  CLEAR_SHEET_BEFORE_RUN: true,
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
