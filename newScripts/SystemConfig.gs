const SYSTEM_CONFIG = {
  IIKO: {
    fields: [
      { key: "date", label: "Дата", type: "text" },
      { key: "docNumber", label: "Входящий номер", type: "text" },
      { key: "partner", label: "Поставщик/Покупатель", type: "text" },
      { key: "warehouse", label: "Склад", type: "text" },
      { key: "sum", label: "Сумма, р.", type: "sum" },
      { key: "comment", label: "Комментарий", type: "text" }
    ]
  },
  DOCSINBOX: {
    fields: [
      { key: "date", label: "Дата", type: "text" },
      { key: "docNumber", label: "Номер накладной поставщика", type: "text" },
      { key: "supplier", label: "Поставщик", type: "text" },
      { key: "buyer", label: "Покупатель", type: "text" },
      { key: "sum", label: "Сумма", type: "sum" },
      { key: "status", label: "Статус приемки", type: "text" }
    ]
  },
  SBIS: {
    fields: [
      { key: "eventDate", label: "Дата события", type: "text" },
      { key: "docNumber", label: "Номер", type: "text" },
      { key: "counterparty", label: "Контрагент", type: "text" },
      { key: "sum", label: "Сумма", type: "sum" },
      { key: "status", label: "Статус", type: "text" }
    ]
  },
  SAP: {
    fields: [
      { key: "docDate", label: "Дата документа", type: "text" },
      { key: "paymentDate", label: "Дата платежа", type: "text" },
      { key: "reference", label: "Ссылка", type: "text" },
      { key: "counterparty", label: "Наименование контрагента", type: "text" },
      { key: "sum", label: "Сумма в ВВ", type: "sum" },
      { key: "docType", label: "Вид документа", type: "text" }
    ]
  }
};

function getSystemConfig(systemName) {
  const config = SYSTEM_CONFIG[systemName];
  if (!config) {
    throw new Error("Неизвестная система: " + systemName);
  }
  return config;
}
