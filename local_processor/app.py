import streamlit as st
import pandas as pd
import os
import time
import json
import tempfile
from processor import UniversalProcessor
from gsheets import upload_to_gsheet, find_file_in_folder, create_spreadsheet_in_folder, get_service_account_quota, read_all_sheets_data, update_supplier_sheet
from rapidfuzz import process, fuzz, utils

st.set_page_config(page_title="Excel Document Processor", layout="wide")

APP_VERSION = "1.1.18"

# Константы Google Drive
SYSTEMS_FOLDER_ID = "1ijv4no6aI3E_le5zAe12QGssehrUETkh"
SUPPLIERS_FOLDER_ID = "1BroJAZivTEypJjFsAOU6uODMaeyw2K7n"
SUPPLIER_TEMPLATE_ID = "1hXBDqliS5rYgL3ZUe2Cg5PUnLAs7aPCwDn8Icivn9qc"

# Путь к файлу справочника ТУ
TU_MAPPING_FILE = "/Users/aleksandrprokudin/Documents/check/searchTU/Распределение РУ, ТУ и ТШП по точкам.xlsx"

# Путь к файлу настроек
SETTINGS_FILE = "settings.json"

@st.cache_data
def load_tu_mapping(file_path):
    """
    Загружает справочник ТУ из Excel.
    Возвращает два словаря:
    syrye_map: { "ADDRESS_PART": "FIO" } - для "Сырье /"
    regular_map: { "FULL_NAME": "FIO" } - для остальных
    """
    if not os.path.exists(file_path):
        return {}, {}
    
    try:
        # Лист "Точка-ТУ СП" (для Сырья)
        # Ищем в col index 3 (D), берем col index 18 (S)
        df_sp = pd.read_excel(file_path, sheet_name="Точка-ТУ СП", header=None)
        syrye_map = {}
        # Пропускаем первые строки (шапку), берем данные
        for idx, row in df_sp.iterrows():
            if idx < 2: continue # Skip header rows
            key = str(row[3]).strip()
            val = str(row[18]).strip()
            if key and key.lower() != "nan" and val and val.lower() != "nan":
                syrye_map[key] = val
                
        # Лист "Точка-ТУ" (для остальных)
        # Ищем в col index 2 (C), берем col index 12 (M)
        df_reg = pd.read_excel(file_path, sheet_name="Точка-ТУ", header=None)
        regular_map = {}
        for idx, row in df_reg.iterrows():
            if idx < 2: continue
            key = str(row[2]).strip()
            val = str(row[12]).strip()
            if key and key.lower() != "nan" and val and val.lower() != "nan":
                regular_map[key] = val
                
        return syrye_map, regular_map
    except Exception as e:
        print(f"[ERROR] Loading TU Mapping: {e}")
        return {}, {}

def find_tu_for_warehouse(warehouse_name, syrye_map, regular_map):
    """
    Ищет ТУ для склада IIKO.
    """
    if not warehouse_name:
        return ""
        
    w_name = str(warehouse_name).strip()
    
    # Стратегия 1: Сырье
    if w_name.lower().startswith("сырье"):
        # "Сырье / КРД Красная ул., 176" -> "КРД Красная ул., 176"
        parts = w_name.split("/", 1)
        if len(parts) > 1:
            target = parts[1].strip()
            # Fuzzy match against syrye_map keys
            match = process.extractOne(target, syrye_map.keys(), scorer=fuzz.token_set_ratio)
            if match and match[1] >= 85:
                return syrye_map[match[0]]
    
    # Стратегия 2: Обычный поиск (или если стратегия 1 не сработала)
    # Ищем полное название в regular_map
    match = process.extractOne(w_name, regular_map.keys(), scorer=fuzz.token_set_ratio)
    if match and match[1] >= 85:
        return regular_map[match[0]]
        
    return ""

def load_settings():
    defaults = {
        "income_k": "платежное, поступление, оплата, списание, перечислено, приход",
        "expense_k": "реализация, упд, продажа, корректировка, акт",
        "target_month": "Январь 26",
        "suppliers": {} # { "Supplier Name": "Spreadsheet ID" }
    }
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                # Удаляем старые ключи папок и шаблона, если они были в JSON
                for k in ["systems_folder_id", "suppliers_folder_id", "supplier_template_id"]:
                    saved.pop(k, None)
                return {**defaults, **saved}
        except:
            return defaults
    return defaults

def save_settings(income_k, expense_k, target_month, suppliers):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "income_k": income_k, 
            "expense_k": expense_k,
            "target_month": target_month,
            "suppliers": suppliers
        }, f, ensure_ascii=False)

settings = load_settings()

st.title("📄 Обработка актов сверки")

# Настройки в боковой панели
with st.sidebar:
    st.header("📅 Период")
    months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
    years = ["25", "26", "27"]
    
    # Пытаемся распарсить текущий месяц из настроек
    try:
        cur_m, cur_y = settings["target_month"].split(" ")
        m_idx = months.index(cur_m)
        y_idx = years.index(cur_y)
    except:
        m_idx, y_idx = 0, 1

    sel_m = st.selectbox("Месяц", months, index=m_idx)
    sel_y = st.selectbox("Год", years, index=y_idx)
    target_month = f"{sel_m} {sel_y}"

    # ИИ всегда Yandex, настройки скрыты
    model_name = "yandexgpt"
    # Ключи берем только из .env
    api_key = os.getenv("YANDEX_API_KEY", "")
    folder_id = os.getenv("YANDEX_FOLDER_ID", "")

    st.divider()
    st.header("🏢 Справочник Поставщиков")
    new_supplier_name = st.text_input("Название поставщика")
    new_supplier_id = st.text_input("ID таблицы (оставьте пустым для авто-создания)")
    
    current_suppliers = settings["suppliers"].copy()
    if st.button("➕ Добавить поставщика"):
        if new_supplier_name:
            if not new_supplier_id:
                with st.spinner(f"Создаем таблицу для {new_supplier_name} по шаблону..."):
                    new_id = create_spreadsheet_in_folder(new_supplier_name, SUPPLIERS_FOLDER_ID, SUPPLIER_TEMPLATE_ID)
                    if new_id:
                        current_suppliers[new_supplier_name] = new_id
                        st.success(f"Создана таблица ID: {new_id}")
                    else:
                        st.error("Ошибка при создании таблицы")
            else:
                current_suppliers[new_supplier_name] = new_supplier_id
                st.success(f"Добавлен: {new_supplier_name}")
        else:
            st.warning("Введите название поставщика")

    if current_suppliers:
        st.write("Список:")
        for s_name in list(current_suppliers.keys()):
            col1, col2 = st.columns([3, 1])
            col1.text(f"{s_name}")
            if col2.button("🗑️", key=f"del_{s_name}"):
                del current_suppliers[s_name]
                save_settings(settings["income_k"], settings["expense_k"], target_month, current_suppliers)
                st.rerun()

    # Проверка наличия прав (Файл или Секреты)
    has_creds = os.path.exists("credentials.json") or "gcp_service_account" in st.secrets
    if not has_creds:
        st.warning("⚠️ Авторизация Google не настроена. Отправка будет недоступна.")

    st.divider()
    st.header("🔍 Настройки фильтра")
    
    inc_k = st.text_area("Ключевые слова ДОХОДА", value=settings["income_k"])
    exp_k = st.text_area("Ключевые слова РАСХОДА", value=settings["expense_k"])
    
    # Сохраняем настройки
    if (inc_k != settings["income_k"] or 
        exp_k != settings["expense_k"] or 
        target_month != settings["target_month"] or
        current_suppliers != settings["suppliers"]):
        save_settings(inc_k, exp_k, target_month, current_suppliers)
        st.rerun()
    
    income_list = [x.strip() for x in inc_k.split(",") if x.strip()]
    expense_list = [x.strip() for x in exp_k.split(",") if x.strip()]

    st.divider()
    st.caption(f"Версия: {APP_VERSION}")

    # Показываем квоту сервисного аккаунта (для диагностики 403 ошибки)
    try:
        from gsheets import get_service_account_quota
        quota = get_service_account_quota()
        if quota:
            usage = int(quota.get('usage', 0)) / (1024**3)
            limit = int(quota.get('limit', 0)) / (1024**3) if 'limit' in quota else 15
            st.caption(f"Drive Storage: {usage:.2f} GB / {limit:.2f} GB used")
    except:
        pass

st.info(f"📅 Выбран период: **{target_month}**")

def normalize_doc_num_for_search(val):
    if not val:
        return ""
    # Оставляем только цифры и буквы, убираем нули в начале
    s = str(val).strip().lower()
    s = "".join(c for c in s if c.isalnum())
    s = s.lstrip("0")
    return s

def find_doc_in_index(target_doc, idx_map):
    """
    Пытается найти документ в индексе:
    1. Точное совпадение
    2. Если точного нет - ищет вхождение (если документ достаточно длинный)
    """
    if not target_doc:
        return []
        
    # 1. Exact match
    if target_doc in idx_map:
        return idx_map[target_doc]
    
    # 2. Substring match (Fallack)
    # Опасно для коротких номеров ("20" найдет "120"), поэтому ставим ограничение
    # Если номер содержит буквы (20dp), риск меньше.
    # Если только цифры, нужна длина хотя бы 3-4? Или доверимся?
    # Кейс пользователя: "20" vs "20dp". "20" (len 2) in "20dp".
    # Чтобы не зацепить "120", проверим, что оно начинается с этого номера
    
    # Оптимизация: не перебирать весь словарь, если он огромный. Но у нас ~500-1000 доков.
    # Ищем ключи, которые НАЧИНАЮТСЯ с target_doc (20 -> 20dp)
    # Или ключи, которые являются target_doc (20dp -> 20 - вряд ли LLM обрежет буквы)
    
    for key in idx_map:
        # Key = System Doc (e.g. "20dp")
        # Target = Act Doc (e.g. "20")
        
        # Проверка 1: Система "20dp", Акт "20". "20dp" starts with "20"
        if key.startswith(target_doc) and len(key) > len(target_doc):
             # Доп проверка: следующий символ после совпадения - буква?
             # 20dp -> 20 (ok), 205 -> 20 (bad)
             suffix = key[len(target_doc):]
             if suffix[0].isalpha(): # DP starts with D
                 return idx_map[key]
                 
    return []

def is_correction(text, amount=None):
    """
    Определяет, является ли запись корректировкой/возвратом.
    1. По тексту (содержит "корректировка", "возврат")
    2. По сумме (отрицательная)
    """
    t = str(text).lower()
    if "корректировка" in t or "возврат" in t:
        return True
    if amount is not None:
        try:
            if float(amount) < 0:
                return True
        except:
            pass
    return False

def perform_reconciliation(act_data, system_data_map, supplier_name):
    # Load TU Mapping
    syrye_map, regular_map = load_tu_mapping(TU_MAPPING_FILE)
    
    # act_data headers: ["Дата", "Текст", "Номер", "Сумма"]
    # system_data_map: { "IIKO": [records...], "SBIS": [records...], ... }
    
    # 1. Prepare fast lookups for systems
    # Filter each system by supplier name (fuzzy) and index by doc number
    
    # Config for system columns
    system_cols = {
        "IIKO": {"partner": "Поставщик/Покупатель", "doc": "Входящий номер", "sum": "Сумма, р.", "comment": "Комментарий"},
        "DOCSINBOX": {"partner": "Поставщик", "doc": "Номер накладной поставщика", "sum": "Сумма"},
        "SBIS": {"partner": "Контрагент", "doc": "Номер", "sum": "Сумма"},
        "SAP": {"partner": "Наименование контрагента", "doc": "Ссылка", "sum": "Сумма в ВВ", "docType": "Вид документа"},
        "FB": {"partner": "Поставщик", "doc": "Номер", "sum": "Сумма", "linked": "Привязан к поставке", "point": "Точка"}
    }
    
    # Pre-process system data: filter by supplier and index by normalized doc number
    system_indices = {} 
    
    # Counters for system docs (excluding corrections)
    system_stats = {
        "IIKO": {"total_sum": 0.0, "count": 0},
        "SAP": {"total_sum": 0.0, "count": 0},
        "FB": {"total_sum": 0.0, "count": 0}
    }
    
    for sys_name, records in system_data_map.items():
        if sys_name not in system_cols:
            continue
            
        cols = system_cols[sys_name]
        
        # Fuzzy match supplier name
        clean_supplier_name = supplier_name.split("(")[0].strip()
        
        # Collect all unique supplier names from system data
        unique_partners = set()
        for r in records:
            p = r.get(cols["partner"], "")
            if p:
                unique_partners.add(str(p).strip()) 
        
        print(f"[RECON]   System {sys_name} has {len(unique_partners)} unique partners.")
             
        # Find matches using clean name and partial_ratio (best for substrings)
        matches = process.extract(
            clean_supplier_name, 
            unique_partners, 
            scorer=fuzz.partial_ratio, 
            limit=5,
            processor=utils.default_process 
        )
        
        # Filter by cutoff
        matched_partners = {m[0] for m in matches if m[1] >= 85}
        
        if matched_partners:
            print(f"[RECON]   ACCEPTED matches in {sys_name}: {list(matched_partners)}")
        else:
            print(f"[RECON]   NO matches accepted in {sys_name} (threshold 65)")
        
        # Index records and calculate stats
        idx_map = {}
        for r in records:
            p = str(r.get(cols["partner"], ""))
            if p in matched_partners:
                doc = r.get(cols["doc"])
                norm_doc = normalize_doc_num_for_search(doc)
                
                amount_str = r.get(cols["sum"], "0")
                try:
                    amount_float = float(str(amount_str).replace(",", ".").replace("\xa0", "").strip() or 0)
                except:
                    amount_float = 0.0
                
                # Check for correction (to exclude from stats)
                # Check negative amount
                # For SAP, normal amounts are negative. Don't use negative sign as correction indicator.
                check_amount = amount_float if sys_name != "SAP" else None
                is_corr = is_correction("", check_amount)
                
                # Check specific fields text
                if not is_corr and sys_name == "IIKO":
                    comment = str(r.get("Комментарий", "")).lower()
                    if "корректировка" in comment or "возврат" in comment:
                        is_corr = True
                
                # Add to stats if NOT correction
                if not is_corr and sys_name in system_stats:
                    system_stats[sys_name]["total_sum"] += amount_float
                    system_stats[sys_name]["count"] += 1
                    
                # Store record for matching (we keep corrections in index to match against act corrections)
                if norm_doc not in idx_map:
                    idx_map[norm_doc] = []
                idx_map[norm_doc].append({"amount": amount_float, "raw": r})
                
        system_indices[sys_name] = idx_map

        # Check for duplicates in IIKO
        if sys_name == "IIKO":
            # 1. Duplicates
            dups = []
            for k, v in idx_map.items():
                if len(v) > 1:
                    # Collect original doc numbers
                    orig_doc = v[0]["raw"].get(cols["doc"], k)
                    dups.append(str(orig_doc))
            if dups:
                system_stats["IIKO"]["duplicates"] = ", ".join(dups)
            
            # 2. Missing in Act (present in IIKO but not in Act)
            # Initially, all docs are unmatched
            system_indices["IIKO_unmatched"] = set(idx_map.keys())

    # 3. Build Result Table & Act Stats
    results = []
    
    act_stats = {"total_sum": 0.0, "count": 0}
    act_missing_in_iiko = [] # Documents in Act but not in IIKO
    
    print(f"\n[RECON] Starting reconciliation for supplier: '{supplier_name}'")
    for sys_name, idx_map in system_indices.items():
        print(f"[RECON] System {sys_name}: {len(idx_map)} docs indexed for this supplier.")
    
    for row in act_data:
        # row: [Date, Text, DocNum, Amount]
        date = row[0]
        text = row[1] 
        doc_num = row[2]
        amount_act_raw = row[3]
        
        try:
            amount_act = float(str(amount_act_raw).replace(",", ".").replace("\xa0", "").strip() or 0)
        except:
            amount_act = 0.0
            
        # Check correction for Act stats
        if not is_correction(text):
            act_stats["total_sum"] += amount_act
            act_stats["count"] += 1
        
        # Основной блок (Поставщик)
        res_row = {
            "supplier_date": date,
            "supplier_doc": text, 
            "supplier_sum": amount_act
        }
        
        norm_doc = normalize_doc_num_for_search(doc_num)
        
        # IIKO
        iiko_idx = system_indices.get("IIKO", {})
        iiko_wh_found = "" # To store warehouse for TU lookup
        
        matches = find_doc_in_index(norm_doc, iiko_idx)
        if matches:
            m = matches[0]
            raw = m["raw"]
            res_row["iiko_date"] = raw.get("Дата", "")
            res_row["iiko_doc"] = raw.get("Входящий номер", "")
            res_row["iiko_partner"] = raw.get("Поставщик/Покупатель", "")
            
            wh = raw.get("Склад", "")
            res_row["iiko_warehouse"] = wh
            iiko_wh_found = wh
            
            res_row["iiko_sum"] = m["amount"]
            res_row["iiko_comment"] = raw.get("Комментарий", "")
            res_row["iiko_delta"] = amount_act - m["amount"]
            
            # Mark as found (remove from unmatched set)
            if "IIKO_unmatched" in system_indices:
                if norm_doc in system_indices["IIKO_unmatched"]:
                    system_indices["IIKO_unmatched"].discard(norm_doc)
                else:
                    # Try to find by prefix if it was a fuzzy match
                    to_remove = []
                    for k in system_indices["IIKO_unmatched"]:
                        if k.startswith(norm_doc) and len(k) > len(norm_doc):
                             to_remove.append(k)
                    for k in to_remove:
                        system_indices["IIKO_unmatched"].discard(k)
        else:
            res_row["iiko_delta"] = amount_act 
            # Добавляем в список "Лишние в Акте"
            # Только если есть номер документа
            if doc_num and str(doc_num).strip():
                act_missing_in_iiko.append(str(doc_num).strip())
        
        # FB (New System)
        fb_idx = system_indices.get("FB", {})
        matches = find_doc_in_index(norm_doc, fb_idx)
        if matches:
            m = matches[0]
            raw = m["raw"]
            res_row["fb_doc"] = raw.get("Номер", "")
            res_row["fb_type"] = raw.get("Тип", "")
            res_row["fb_linked"] = raw.get("Привязан к поставке", "")
            res_row["fb_partner"] = raw.get("Поставщик", "")
            res_row["fb_point"] = raw.get("Точка", "")
            res_row["fb_date"] = raw.get("Дата документа", "")
            res_row["fb_status"] = raw.get("Статус", "")
            res_row["fb_del_status"] = raw.get("Статус поставки", "")
            res_row["fb_sum"] = m["amount"]
            res_row["fb_delta"] = amount_act - m["amount"]
        else:
             res_row["fb_delta"] = amount_act
            
        # DOCSINBOX
        dxbx_idx = system_indices.get("DOCSINBOX", {})
        matches = find_doc_in_index(norm_doc, dxbx_idx)
        if matches:
            m = matches[0]
            raw = m["raw"]
            res_row["dxbx_buyer"] = raw.get("Покупатель", "")
            res_row["dxbx_status"] = raw.get("Статус приемки", "")
            
            # Lookup TU based on IIKO warehouse
            if iiko_wh_found:
                tu_name = find_tu_for_warehouse(iiko_wh_found, syrye_map, regular_map)
                res_row["dxbx_tu"] = tu_name
            
            # FALLBACK: If TU not found via IIKO, try to find via DXBX Buyer
            if not res_row.get("dxbx_tu") and res_row.get("dxbx_buyer"):
                buyer_raw = res_row["dxbx_buyer"]
                # Clean: remove content in brackets and extra spaces
                buyer_clean = buyer_raw.split("(")[0].strip()
                
                # Search in regular_map (addresses) with lower threshold
                # Use token_set_ratio to handle word reordering and extra words
                match = process.extractOne(buyer_clean, regular_map.keys(), scorer=fuzz.token_set_ratio)
                if match and match[1] >= 60:
                     print(f"[RECON] TU Fallback: '{buyer_clean}' -> '{match[0]}' ({match[1]}%)")
                     res_row["dxbx_tu"] = regular_map[match[0]]
        
        # SBIS
        sbis_idx = system_indices.get("SBIS", {})
        matches = find_doc_in_index(norm_doc, sbis_idx)
        if matches:
            m = matches[0]
            raw = m["raw"]
            res_row["sbis_status"] = raw.get("Статус", "")
            res_row["sbis_delta"] = amount_act - m["amount"]
        else:
            res_row["sbis_delta"] = amount_act
            
        # SAP
        sap_idx = system_indices.get("SAP", {})
        matches = find_doc_in_index(norm_doc, sap_idx)
        if matches:
            m = matches[0]
            raw = m["raw"]
            res_row["sap_doc_type"] = raw.get("Вид документа", "")
            # SAP amounts are negative. Delta = Act + SAP (e.g. 100 + (-100) = 0)
            res_row["sap_delta"] = amount_act + m["amount"]
        else:
            res_row["sap_delta"] = amount_act
            
        # Пользовательский комментарий
        res_row["manager_comment"] = ""
                
        results.append(res_row)
        
    # Collect unmatched IIKO docs names
    iiko_missing_in_act = []
    if "IIKO_unmatched" in system_indices and "IIKO" in system_indices:
        for k in system_indices["IIKO_unmatched"]:
            # Get original doc name from the first record in the list
            records = system_indices["IIKO"].get(k, [])
            if records:
                orig = records[0]["raw"].get("Входящий номер", k)
                iiko_missing_in_act.append(str(orig))
                
    summary = {
        "iiko_total": system_stats["IIKO"]["total_sum"],
        "sap_total": system_stats["SAP"]["total_sum"],
        "fb_total": system_stats["FB"]["total_sum"],
        "act_total": act_stats["total_sum"],
        
        "delta_act_iiko": act_stats["total_sum"] - system_stats["IIKO"]["total_sum"],
        # SAP amounts are negative. Sum them up to get delta.
        "delta_act_sap": act_stats["total_sum"] + system_stats["SAP"]["total_sum"],
        "delta_act_fb": act_stats["total_sum"] - system_stats["FB"]["total_sum"],
        
        "act_count": act_stats["count"],
        "iiko_count": system_stats["IIKO"]["count"],
        "delta_count": act_stats["count"] - system_stats["IIKO"]["count"],
        
        "iiko_duplicates": system_stats["IIKO"].get("duplicates", ""),
        "iiko_missing": ", ".join(iiko_missing_in_act),
        "act_missing": ", ".join(act_missing_in_iiko)
    }
        
    return {"rows": results, "summary": summary}

uploaded_files = st.file_uploader("Выберите Excel файлы", type=["xlsx", "xls"], accept_multiple_files=True)

# Опция для стратегии извлечения номера
extraction_mode = st.radio(
    "Strategia извлечения номера документа:",
    ("Авто (Приоритет С/Ф)", "Строго первый номер (Акт)"),
    horizontal=True,
    help="Выберите 'Авто', чтобы искать номер счета-фактуры (обычно в скобках). Выберите 'Акт', чтобы брать первый номер (номер накладной)."
)

if "results" not in st.session_state:
    st.session_state.results = {}

if uploaded_files:
    for file_index, uploaded_file in enumerate(uploaded_files):
        # Добавляем режим в ключ, чтобы при смене радио-кнопки пересчитывалось
        file_key = f"{uploaded_file.name}_{uploaded_file.size}_{file_index}_{extraction_mode}"
        
        if file_key not in st.session_state.results:
            with st.spinner(f"Обработка {uploaded_file.name}..."):
                # Используем правильный tempfile для облака
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                    tmp.write(uploaded_file.getbuffer())
                    temp_path = tmp.name
                
                try:
                    processor = UniversalProcessor(model_name=model_name)
                    data, status, system_name, headers = processor.process_file(
                        temp_path,
                        income_keywords=income_list,
                        expense_keywords=expense_list,
                        extraction_mode=extraction_mode
                    )
                    
                    if status in ("enriched", "enriched_system"):
                        if not data:
                            st.warning(f"В файле {uploaded_file.name} не найдено данных для выгрузки.")
                        else:
                            st.session_state.results[file_key] = {
                                "data": data,
                                "system": system_name,
                                "filename": uploaded_file.name,
                                "headers": headers
                            }
                    else:
                        st.error(f"Ошибка при обработке {uploaded_file.name}. Убедитесь, что это Excel файл с данными.")
                finally:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

        if file_key in st.session_state.results:
            res_obj = st.session_state.results[file_key]
            data = res_obj["data"]
            system = res_obj["system"]
            filename = res_obj["filename"]
            headers = res_obj.get("headers", [])
            
            with st.expander(f"📊 {filename} [Система: {system}]", expanded=True):
                if data and headers:
                    df = pd.DataFrame(data, columns=headers).astype(str)
                    st.dataframe(df, use_container_width=True)
                    
                    if system != "OTHER":
                        if st.button(f"🚀 Отправить в Системы ({system})", key=f"btn_{file_key}"):
                            with st.spinner("Ищем/создаем таблицу месяца..."):
                                gs_id = find_file_in_folder(SYSTEMS_FOLDER_ID, target_month)
                                if not gs_id:
                                    gs_id = create_spreadsheet_in_folder(target_month, SYSTEMS_FOLDER_ID)
                                
                                if gs_id:
                                    success, msg = upload_to_gsheet(gs_id, system, data, headers)
                                    if success:
                                        st.success(f"{msg} в таблицу '{target_month}'")
                                    else:
                                        st.error(msg)
                    else:
                        st.info("Это похоже на Акт сверки. Выберите поставщика для обработки:")
                        selected_supplier = st.selectbox(
                            "Выберите поставщика",
                            ["-- Выберите --"] + list(current_suppliers.keys()),
                            key=f"sel_{file_key}"
                        )
                        
                        if selected_supplier != "-- Выберите --":
                            supplier_id = current_suppliers[selected_supplier]
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                if st.button(f"🚀 Отправить Поставщику (Исходный)", key=f"btn_orig_{file_key}"):
                                    success, msg = upload_to_gsheet(supplier_id, target_month, data, headers)
                                    if success:
                                        st.success(f"{msg} (лист {target_month})")
                                    else:
                                        st.error(msg)
                                        
                            with col2:
                                if st.button(f"⚔️ Сравнить с данными систем", key=f"btn_recon_{file_key}"):
                                    with st.spinner("Загружаем данные систем для сверки..."):
                                        # 1. Find system spreadsheet
                                        sys_ss_id = find_file_in_folder(SYSTEMS_FOLDER_ID, target_month)
                                        if not sys_ss_id:
                                            st.error(f"Не найден файл системных данных за период {target_month}")
                                        else:
                                            st.toast(f"Файл систем найден: {sys_ss_id}")
                                            # 2. Read all sheets
                                            sys_data = read_all_sheets_data(sys_ss_id)
                                            if not sys_data:
                                                st.error("Не удалось прочитать данные систем")
                                            else:
                                                sheet_counts = {k: len(v) for k, v in sys_data.items()}
                                                st.toast(f"Данные загружены: {sheet_counts}")
                                                
                                                # 3. Perform reconciliation
                                                st.info(f"Сверка для поставщика: {selected_supplier}")
                                                recon_result_obj = perform_reconciliation(data, sys_data, selected_supplier)
                                                
                                                # Save results to session state to display
                                                st.session_state[f"recon_{file_key}"] = recon_result_obj
                                                st.toast(f"Сверка завершена. Найдено {len(recon_result_obj['rows'])} строк.")
                                                
                        # Display reconciliation results if available
                        if f"recon_{file_key}" in st.session_state:
                            recon_res_obj = st.session_state[f"recon_{file_key}"]
                            recon_rows = recon_res_obj["rows"]
                            summary = recon_res_obj.get("summary", {})
                            
                            st.divider()
                            st.write("### 🏁 Результаты сверки")
                            
                            # Отображаем саммари
                            c1, c2, c3, c4 = st.columns(4)
                            c1.metric("Оборот Акт", f"{summary.get('act_total',0):,.2f}")
                            c2.metric("Оборот IIKO", f"{summary.get('iiko_total',0):,.2f}")
                            c3.metric("Оборот SAP", f"{summary.get('sap_total',0):,.2f}")
                            c4.metric("Оборот FB", f"{summary.get('fb_total',0):,.2f}")
                            
                            c5, c6, c7 = st.columns(3)
                            c5.metric("Δ Акт-IIKO", f"{summary.get('delta_act_iiko',0):,.2f}")
                            c6.metric("Δ Акт-SAP", f"{summary.get('delta_act_sap',0):,.2f}")
                            c7.metric("Δ Акт-FB", f"{summary.get('delta_act_fb',0):,.2f}")
                            
                            # New metrics: document counts
                            d1, d2, d3 = st.columns(3)
                            d1.metric("Кол-во док. Акт", f"{summary.get('act_count',0):,.0f}")
                            d2.metric("Кол-во док. IIKO", f"{summary.get('iiko_count',0):,.0f}")
                            d3.metric("Δ Кол-во", f"{summary.get('delta_count',0):,.0f}")

                            if summary.get("iiko_duplicates"):
                                st.warning(f"⚠️ Найдены дубликаты накладных в IIKO: {summary['iiko_duplicates']}")
                            
                            if summary.get("iiko_missing"):
                                st.error(f"🚫 Найдены документы в IIKO, которых НЕТ в Акте: {summary['iiko_missing']}")
                            
                            if summary.get("act_missing"):
                                st.error(f"❓ Найдены документы в Акте, которых НЕТ в IIKO: {summary['act_missing']}")
                            
                            st.dataframe(pd.DataFrame(recon_rows))
                            
                            if st.button(f"💾 Сохранить результаты сверки", key=f"btn_save_recon_{file_key}"):
                                supplier_id = current_suppliers[selected_supplier]
                                # Create a special sheet name for reconciliation
                                recon_sheet_name = f"Сверка {target_month}"
                                success, msg = update_supplier_sheet(supplier_id, recon_sheet_name, recon_rows, summary)
                                if success:
                                    st.success(msg)
                                else:
                                    st.error(msg)
