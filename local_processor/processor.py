import ollama
import json
import os
import re
import time
import psutil
import requests
from datetime import datetime
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()

from excel_preprocessor.cleaner import clean_excel

SYSTEM_CONFIG = {
    "IIKO": {
        "output_headers": [
            "Дата",
            "Входящий номер",
            "Поставщик/Покупатель",
            "Склад",
            "Сумма, р.",
            "Комментарий"
        ],
        "fields": [
            {"key": "date", "labels": ["Дата", "Дата документа", "Дата операции"]},
            {"key": "docNumber", "labels": ["Входящий номер", "Номер документа", "Вх. номер", "Входящий №"]},
            {"key": "partner", "labels": ["Поставщик/Покупатель", "Поставщик", "Покупатель", "Контрагент"]},
            {"key": "warehouse", "labels": ["Склад"]},
            {"key": "sum", "labels": ["Сумма, р.", "Сумма", "Итого"]},
            {"key": "comment", "labels": ["Комментарий"]},
        ]
    },
    "DOCSINBOX": {
        "output_headers": [
            "Дата",
            "Номер накладной поставщика",
            "Поставщик",
            "Покупатель",
            "Сумма",
            "Статус приемки"
        ],
        "fields": [
            {"key": "date", "labels": ["Дата", "Дата документа", "Дата операции"]},
            {"key": "docNumber", "labels": ["Номер накладной поставщика", "Номер накладной", "Номер ТТН", "Номер", "Номер документа"]},
            {"key": "supplier", "labels": ["Поставщик"]},
            {"key": "buyer", "labels": ["Покупатель"]},
            {"key": "sum", "labels": ["Сумма"]},
            {"key": "status", "labels": ["Статус приемки", "Статус"]},
        ]
    },
    "SBIS": {
        "output_headers": [
            "Дата события",
            "Номер",
            "Контрагент",
            "Сумма",
            "Статус"
        ],
        "fields": [
            {"key": "eventDate", "labels": ["Дата события", "Дата"]},
            {"key": "docNumber", "labels": ["Номер"]},
            {"key": "counterparty", "labels": ["Контрагент"]},
            {"key": "sum", "labels": ["Сумма"]},
            {"key": "status", "labels": ["Статус"]},
        ]
    },
    "SAP": {
        "output_headers": [
            "Дата документа",
            "Дата платежа",
            "Ссылка",
            "Наименование контрагента",
            "Сумма в ВВ",
            "Вид документа"
        ],
        "fields": [
            {"key": "docDate", "labels": ["Дата документа"]},
            {"key": "paymentDate", "labels": ["Дата платежа"]},
            {"key": "reference", "labels": ["Ссылка"]},
            {"key": "counterparty", "labels": ["Наименование контрагента"]},
            {"key": "sum", "labels": ["Сумма в ВВ", "Сумма"]},
            {"key": "docType", "labels": ["Вид документа"]},
        ]
    },
    "FB": {
        "output_headers": [
            "Номер",
            "Тип",
            "Привязан к поставке",
            "Поставщик",
            "Точка",
            "Дата документа",
            "Статус",
            "Статус поставки",
            "Сумма"
        ],
        "fields": [
            {"key": "docNumber", "labels": ["Номер"]},
            {"key": "type", "labels": ["Тип"]},
            {"key": "linked", "labels": ["Привязан к поставке"]},
            {"key": "partner", "labels": ["Поставщик"]},
            {"key": "point", "labels": ["Точка"]},
            {"key": "date", "labels": ["Дата документа"]},
            {"key": "status", "labels": ["Статус"]},
            {"key": "deliveryStatus", "labels": ["Статус поставки"]},
            {"key": "sum", "labels": ["Сумма"]},
        ]
    },
}

def normalize_header(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[«»\"']", "", text)
    text = re.sub(r"\s+", " ", text)
    return text

def find_header_row(rows, system_name, max_rows=100):
    config = SYSTEM_CONFIG.get(system_name)
    if not config:
        return None
    best_index = None
    best_score = 0
    for i, row in enumerate(rows[:max_rows]):
        normalized = [normalize_header(cell) for cell in row]
        if not any(normalized):
            continue
        score = 0
        for field in config["fields"]:
            for label in field["labels"]:
                if normalize_header(label) in normalized:
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_index = i
    return best_index

def build_column_map(header_row, system_name):
    config = SYSTEM_CONFIG.get(system_name)
    if not config:
        return None
    normalized = [normalize_header(cell) for cell in header_row]
    col_map = {}
    for field in config["fields"]:
        col_idx = None
        for label in field["labels"]:
            norm_label = normalize_header(label)
            if norm_label in normalized:
                col_idx = normalized.index(norm_label)
                break
        col_map[field["key"]] = col_idx
    return col_map

def get_header_match_score(header_row, system_name):
    config = SYSTEM_CONFIG.get(system_name)
    if not config:
        return 0
    normalized = [normalize_header(cell) for cell in header_row]
    score = 0
    for field in config["fields"]:
        for label in field["labels"]:
            if normalize_header(label) in normalized:
                score += 1
                break
    return score

def detect_system_by_header(raw_rows):
    if not raw_rows:
        return None
    best_system = None
    best_score = 0
    for system_name in SYSTEM_CONFIG.keys():
        header_idx = find_header_row(raw_rows, system_name)
        if header_idx is None:
            continue
        score = get_header_match_score(raw_rows[header_idx], system_name)
        if score > best_score:
            best_score = score
            best_system = system_name
    return best_system if best_score > 0 else None

def chunk_rows(rows, chunk_size):
    return [rows[i:i + chunk_size] for i in range(0, len(rows), chunk_size)]

def parse_llm_json(content):
    if not content:
        return None
    content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip()
    cleaned = content.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z]*\n", "", cleaned).strip()
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3].strip()
    try:
        return json.loads(cleaned)
    except Exception:
        pass
    start_candidates = [cleaned.find("["), cleaned.find("{")]
    start_candidates = [idx for idx in start_candidates if idx != -1]
    if not start_candidates:
        return None
    start = min(start_candidates)
    end = max(cleaned.rfind("]"), cleaned.rfind("}"))
    if end == -1 or end <= start:
        return None
    snippet = cleaned[start:end + 1]
    try:
        return json.loads(snippet)
    except Exception:
        return None

def extract_rows_from_parsed(parsed):
    if isinstance(parsed, list):
        return parsed
    if isinstance(parsed, dict):
        if "numbers" in parsed and isinstance(parsed["numbers"], list):
            return parsed["numbers"]
        for value in parsed.values():
            if isinstance(value, list):
                return value
        keys = list(parsed.keys())
        if all(k.isdigit() for k in keys):
            sorted_keys = sorted(keys, key=lambda x: int(x))
            return [parsed[k] for k in sorted_keys]
        if any(k in parsed for k in ("date", "text", "amount", "doc_number", "number")):
            return [parsed]
    return None

def normalize_doc_number(val):
    if val is None:
        return None
    if isinstance(val, dict):
        return str(list(val.values())[0]) if val.values() else None
    s_val = str(val).strip()
    if s_val.lower() in ("null", "none", "", "skip"):
        return None
    # Убираем лишние слова, если ИИ их добавил
    s_val = re.sub(r"^(номер|№|док|id)\s*", "", s_val, flags=re.IGNORECASE).strip()
    return s_val

class UniversalProcessor:
    def __init__(self, model_name="llama3.2:3b"):
        self.model_name = model_name
        
        # Сначала пробуем взять из секретов Streamlit (для Облака)
        try:
            import streamlit as st
            self.yandex_api_key = st.secrets.get("YANDEX_API_KEY")
            self.yandex_folder_id = st.secrets.get("YANDEX_FOLDER_ID")
        except:
            self.yandex_api_key = None
            self.yandex_folder_id = None

        # Если в секретах нет, берем из .env (для Mac)
        if not self.yandex_api_key:
            self.yandex_api_key = os.getenv("YANDEX_API_KEY")
        if not self.yandex_folder_id:
            self.yandex_folder_id = os.getenv("YANDEX_FOLDER_ID")
            
        self.is_yandex = "yandex" in model_name.lower()

    def log(self, message):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def call_yandex_gpt(self, system_prompt, user_prompt):
        if not self.yandex_api_key or not self.yandex_folder_id:
            raise ValueError("Yandex API Key or Folder ID not found in .env")

        url = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Api-Key {self.yandex_api_key}",
            "x-folder-id": self.yandex_folder_id
        }
        
        payload = {
            "modelUri": f"gpt://{self.yandex_folder_id}/yandexgpt-lite",
            "completionOptions": {
                "stream": False,
                "temperature": 0,
                "maxTokens": "2000"
            },
            "messages": [
                {"role": "system", "text": system_prompt},
                {"role": "user", "text": user_prompt}
            ]
        }

        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        result = response.json()
        return result["result"]["alternatives"][0]["message"]["text"]

    def extract_system_rows(self, raw_rows, system_name):
        config = SYSTEM_CONFIG.get(system_name)
        if not config:
            return [], None

        header_idx = find_header_row(raw_rows, system_name)
        if header_idx is None:
            self.log(f"!! Не найдена строка заголовков для системы {system_name}. Первые 5 строк:")
            for i, r in enumerate(raw_rows[:5]):
                self.log(f"  Row {i}: {r}")
            return [], None

        header_row = raw_rows[header_idx]
        col_map = build_column_map(header_row, system_name)
        if not col_map:
            return [], None

        output_headers = config["output_headers"]
        results = []
        for row in raw_rows[header_idx + 1:]:
            # Собираем значения строго по ТЗ, без конвертации
            values = []
            has_value = False
            for field in config["fields"]:
                col_idx = col_map.get(field["key"])
                value = row[col_idx] if col_idx is not None and col_idx < len(row) else ""
                if value:
                    has_value = True
                values.append(value or "")
            if has_value:
                results.append(values)

        return results, output_headers

    def enrich_with_doc_numbers(self, rows, max_rows_per_chunk=50, max_chunks=None, 
                               income_keywords=None, expense_keywords=None, extraction_mode="Авто (Приоритет С/Ф)"):
        if not rows:
            return []

        # 1. ПРЕ-ФИЛЬТРАЦИЯ: Убираем платежки на уровне кода
        if income_keywords is None:
            income_keywords = ["платежное", "поступление", "оплата", "списание", "перечислено", "приход"]
        if expense_keywords is None:
            expense_keywords = ["реализация", "упд", "продажа", "корректировка", "акт"]
        
        filtered_rows = []
        for row in rows:
            text_lc = row[1].lower()
            # Если это явно доход (платежка), пропускаем
            if any(k.strip().lower() in text_lc for k in income_keywords) and \
               not any(k.strip().lower() in text_lc for k in expense_keywords):
                continue
            filtered_rows.append(row)

        if not filtered_rows:
            return []

        results = []
        chunks = chunk_rows(filtered_rows, max_rows_per_chunk)
        if max_chunks is not None:
            chunks = chunks[:max_chunks]

        # Улучшенный промт для работы со словарем
        if "Авто" in extraction_mode:
            system_prompt = (
                "Ты — эксперт-бухгалтер. Твоя задача — извлечь номер документа для сверки.\n"
                "ПРИОРИТЕТ: Если есть номер Счета-Фактуры (обычно в скобках, или с дробью типа /DP, /K), бери ЕГО.\n"
                "Если нет — бери номер Акта/Накладной. Игнорируй слова 'от', '№'.\n"
                "Пример: 'Продажа №20 (сф 20/DP)' -> '20/DP'. 'Акт 5' -> '5'.\n"
                "Тебе дан JSON {id: текст}. Верни JSON {id: номер}. Иначе null.\n"
                "Не пиши ничего, кроме JSON."
            )
        else:
            # Режим "Строго первый номер (Акт)"
            system_prompt = (
                "Ты — эксперт-бухгалтер. Твоя задача — извлечь ПЕРВЫЙ номер документа (номер Акта/Накладной).\n"
                "Игнорируй номера счетов-фактур в скобках.\n"
                "Пример: 'Продажа №20 (сф 20/DP)' -> '20'. 'Акт 5' -> '5'.\n"
                "Тебе дан JSON {id: текст}. Верни JSON {id: номер}. Иначе null.\n"
                "Не пиши ничего, кроме JSON."
            )

        for idx, chunk in enumerate(chunks, start=1):
            mem = psutil.virtual_memory()
            self.log(f"LLM: чанк {idx}/{len(chunks)} (строк: {len(chunk)}) | RAM: {mem.percent}%")
            
            # Отправляем словарь {id: text} для защиты от смещения
            payload_dict = {str(i): row[1] for i, row in enumerate(chunk)}
            user_prompt = json.dumps(payload_dict, ensure_ascii=False)
            
            try:
                start_time = time.time()
                
                if self.is_yandex:
                    self.log(f"-> Отправка {len(chunk)} строк в YandexGPT ({len(user_prompt)} симв.)")
                    content = self.call_yandex_gpt(system_prompt, f"Тексты:\n{user_prompt}")
                else:
                    self.log(f"-> Отправка {len(chunk)} строк в {self.model_name}")
                    response = ollama.chat(
                        model=self.model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_prompt},
                        ],
                        format="json",
                        options={"temperature": 0}
                    )
                    content = response["message"]["content"]
                
                elapsed = time.time() - start_time
                self.log(f"<- Ответ получен за {elapsed:.1f}с.")
                
                # Логируем ответ для отладки (без обратных слешей)
                debug_info = content[:150].replace("\n", " ")
                self.log(f"RAW DEBUG: {debug_info}...")
                
                parsed = parse_llm_json(content)
                
                if not isinstance(parsed, dict):
                    self.log(f"!! НЕ УДАЛОСЬ РАСПАРСИТЬ JSON. Полный ответ: {content[:200]}")
                    parsed = {}

                for i, original_row in enumerate(chunk):
                    # Извлекаем по ключу, чтобы не было смещения
                    val = parsed.get(str(i))
                    doc_num = normalize_doc_number(val)
                    results.append(original_row + [doc_num])
                    
            except Exception as e:
                self.log(f"!! ОШИБКА ЧАНКА {idx}: {str(e)}")
                for original_row in chunk:
                    results.append(original_row + [None])
            
            if not self.is_yandex:
                time.sleep(0.5)

        return results

    def process_file(self, file_path, income_keywords=None, expense_keywords=None, extraction_mode="Авто (Приоритет С/Ф)"):
        file_name = os.path.basename(file_path)
        self.log(f"Файл: {file_name}")
        system_name = self.resolve_system_name(file_name)

        raw_rows = None
        # Если система не определена по имени (OTHER), пробуем по заголовкам
        # НО! Если по заголовкам определится что-то невнятное, всё равно будем считать это Актом
        if system_name == "OTHER":
            raw_rows = clean_excel(file_path, raw=True)
            if isinstance(raw_rows, str):
                return [], "error", system_name, []
            
            # Пытаемся определить систему по структуре колонок
            detected = detect_system_by_header(raw_rows)
            if detected:
                # Дополнительная проверка: действительно ли это системный файл?
                # Если совпадение слабое (мало колонок), лучше считать это Актом
                header_idx = find_header_row(raw_rows, detected)
                score = get_header_match_score(raw_rows[header_idx], detected) if header_idx is not None else 0
                
                # Порог уверенности: например, должно совпасть хотя бы 3 ключевых колонки
                if score >= 3:
                    system_name = detected
                    self.log(f"Авто-определение системы по заголовкам: {system_name} (score: {score})")
                else:
                    self.log(f"Похоже на {detected} (score: {score}), но недостаточно уверенно. Считаем Актом.")


        if system_name != "OTHER":
            if raw_rows is None:
                raw_rows = clean_excel(file_path, raw=True)
            if isinstance(raw_rows, str):
                return [], "error", system_name, []
            system_rows, headers = self.extract_system_rows(raw_rows, system_name)
            if not headers:
                return [], "error", system_name, []
            return system_rows, "enriched_system", system_name, headers

        rows = clean_excel(file_path)
        if isinstance(rows, str):
            return [], "error", "OTHER", []

        enriched = self.enrich_with_doc_numbers(rows, income_keywords=income_keywords, expense_keywords=expense_keywords, extraction_mode=extraction_mode)
        # Приводим к формату для актов: [Дата, Текст, Номер, Сумма]
        partner_rows = []
        for row in enriched:
            if len(row) >= 4:
                partner_rows.append([row[0], row[1], row[3], row[2]])
        headers = ["Дата", "Текст", "Номер", "Сумма"]
        return partner_rows, "enriched", "OTHER", headers

    def resolve_system_name(self, file_name):
        lower_name = file_name.lower()
        if "iiko" in lower_name or "иико" in lower_name:
            return "IIKO"
        if "dxbx" in lower_name or "docs" in lower_name or "inbox" in lower_name:
            return "DOCSINBOX"
        # Проверяем "sbis", но только если это не часть названия поставщика в Акте
        # (хотя обычно поставщик так не называется, но перестрахуемся)
        if "sbis" in lower_name or "сбис" in lower_name:
            return "SBIS"
        if "sap" in lower_name or "сап" in lower_name:
            return "SAP"
        if "fb" in lower_name or "фб" in lower_name:
            return "FB"
        
        # Если явных признаков системы нет, скорее всего это Акт
        return "OTHER"
