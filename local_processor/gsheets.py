import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os

# Путь к файлу ключей сервисного аккаунта
CREDENTIALS_FILE = "credentials.json"

def get_creds():
    # 1. Пробуем взять из секретов Streamlit (для Облака)
    try:
        if "gcp_service_account" in st.secrets:
            return Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=[
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
            )
    except Exception:
        # Игнорируем ошибку отсутствия secrets.toml локально
        pass
    
    # 2. Пробуем локальный файл (для Mac)
    if os.path.exists(CREDENTIALS_FILE):
        return Credentials.from_service_account_file(
            CREDENTIALS_FILE, 
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
        )
    return None

def get_gsheets_client():
    creds = get_creds()
    if not creds:
        return None
    return gspread.authorize(creds)

def get_drive_service():
    creds = get_creds()
    if not creds:
        return None
    return build('drive', 'v3', credentials=creds)

def get_service_account_quota():
    try:
        service = get_drive_service()
        if not service:
            return None
        about = service.about().get(fields="storageQuota").execute()
        return about.get('storageQuota', {})
    except Exception:
        return None

def find_file_in_folder(folder_id, file_name):
    service = get_drive_service()
    if not service:
        return None
    
    query = f"'{folder_id}' in parents and name = '{file_name}' and trashed = false"
    # Добавляем supportsAllDrives=True для поиска в общих дисках
    results = service.files().list(
        q=query, 
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = results.get('files', [])
    return files[0]['id'] if files else None

def create_spreadsheet_in_folder(file_name, folder_id, template_id=None):
    service = get_drive_service()
    if not service:
        return None
    
    new_file_id = None
    if template_id:
        # Копируем из шаблона
        body = {'name': file_name, 'parents': [folder_id]}
        try:
            new_file = service.files().copy(
                fileId=template_id, 
                body=body,
                supportsAllDrives=True
            ).execute()
            new_file_id = new_file.get('id')
        except Exception:
            raise
    else:
        # Создаем пустую таблицу
        file_metadata = {
            'name': file_name,
            'parents': [folder_id],
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        new_file = service.files().create(
            body=file_metadata, 
            fields='id',
            supportsAllDrives=True
        ).execute()
        new_file_id = new_file.get('id')

    # Попытка удалить "Лист1" (или Sheet1), если файл создан с нуля, чтобы было чисто
    # Для шаблонных файлов это может быть не нужно, если там нет лишних листов
    # Но если попросили - можно попробовать
    if new_file_id and not template_id:
        try:
            client = get_gsheets_client()
            if client:
                ss = client.open_by_key(new_file_id)
                # Обычно при создании пустой таблицы там один лист "Лист1"
                # Мы его не можем удалить, если он единственный.
                # Поэтому удаление имеет смысл только ПОСЛЕ добавления данных.
                pass 
        except:
            pass
            
    return new_file_id

def upload_to_gsheet(spreadsheet_id, sheet_name, rows, headers, clear_sheet=True):
    client = get_gsheets_client()
    if not client:
        return False, "Файл credentials.json не найден"
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
            if clear_sheet:
                # Очищаем всё
                worksheet.clear()
                worksheet.append_row(headers)
        except gspread.exceptions.WorksheetNotFound:
            # Создаем новый лист, если его нет
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=len(headers))
            worksheet.append_row(headers)
        
        # Добавляем данные
        if rows:
            worksheet.append_rows(rows)
            
        # Попытка удалить "Лист1/Sheet1", если он пустой и мы только что создали другой лист
        try:
            default_sheet = spreadsheet.sheet1
            if default_sheet.title in ["Лист1", "Sheet1"] and default_sheet.title != sheet_name:
                # Проверяем, пустой ли он (необязательно, но безопаснее)
                if not default_sheet.get_all_values():
                    spreadsheet.del_worksheet(default_sheet)
        except:
            pass
        
        return True, f"Данные успешно обновлены в '{sheet_name}'"
    except Exception as e:
        return False, str(e)

def read_all_sheets_data(spreadsheet_id):
    """
    Читает все листы из таблицы и возвращает словарь {sheet_name: list_of_dicts}
    """
    client = get_gsheets_client()
    if not client:
        return None
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheets = spreadsheet.worksheets()
        
        result = {}
        for ws in worksheets:
            # get_all_records() возвращает список словарей, где ключи - заголовки
            # Если заголовки пустые или дублируются, это может вызвать ошибку.
            # Поэтому используем get_all_values() и обрабатываем сами.
            rows = ws.get_all_values()
            if not rows:
                result[ws.title] = []
                continue
                
            headers = rows[0]
            data = []
            for row in rows[1:]:
                # Создаем словарь, пропуская пустые заголовки
                record = {}
                for i, cell in enumerate(row):
                    if i < len(headers) and headers[i]:
                        record[headers[i]] = cell
                data.append(record)
            result[ws.title] = data
            
        return result
    except Exception as e:
        print(f"Error reading spreadsheet {spreadsheet_id}: {e}")
        return None

def update_supplier_sheet(spreadsheet_id, sheet_name, data, summary=None):
    """
    Обновляет данные на конкретном листе с сохранением комментариев пользователя.
    Использует жесткую привязку к шаблону TEMPLATE_NEW.xlsx.
    summary: словарь с итогами {"act_total": ..., "iiko_total": ..., "sap_total": ..., "delta_act_iiko": ..., "delta_act_sap": ...}
    """
    client = get_gsheets_client()
    if not client:
        return False, "Файл credentials.json не найден"
    
    # Карта колонок по шаблону TEMPLATE_NEW.xlsx (A=0, B=1, ...)
    COL_MAP_IDX = {
        "supplier_date": 4, # E
        "supplier_doc": 5,  # F
        "supplier_sum": 6,  # G
        
        "iiko_date": 8,     # I
        "iiko_doc": 9,      # J
        "iiko_partner": 10, # K
        "iiko_warehouse": 11, # L
        "iiko_sum": 12,     # M
        "iiko_comment": 13, # N
        "iiko_delta": 14,   # O
        
        "fb_doc": 16,       # Q
        "fb_type": 17,      # R
        "fb_linked": 18,    # S
        "fb_partner": 19,   # T
        "fb_point": 20,     # U
        "fb_date": 21,      # V
        "fb_status": 22,    # W
        "fb_del_status": 23, # X
        "fb_sum": 24,       # Y
        "fb_delta": 25,     # Z
        
        "dxbx_buyer": 26,   # AA
        "dxbx_status": 27,  # AB
        "dxbx_tu": 28,      # AC
        "dxbx_comment": 29, # AD
        
        "sbis_delta": 31,   # AF
        "sbis_status": 32,  # AG
        
        "sap_delta": 34,    # AI
        "sap_doc_type": 35, # AJ
        "manager_comment": 36 # AK
    }
    
    MAX_COL_IDX = 36 # AK
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = None
        
        # 1. Сначала ищем лист с точным именем "Сверка {Месяц}"
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
            
            # Проверка целостности шаблона:
            # Проверяем ячейку E1 (должно быть "ПОСТАВЩИК")
            val_e1 = worksheet.acell('E1').value
            if not val_e1 or "ПОСТАВЩИК" not in str(val_e1).upper():
                print(f"[DEBUG] Лист {sheet_name} существует, но выглядит поврежденным (E1='{val_e1}'). Удаляем.")
                spreadsheet.del_worksheet(worksheet)
                worksheet = None # Сброс, чтобы пойти по ветке создания
            else:
                 print(f"[DEBUG] Лист {sheet_name} найден и выглядит валидным.")

        except gspread.exceptions.WorksheetNotFound:
            pass

        if not worksheet:
            # Если листа нет (или удалили битый), ищем базу
            print(f"[DEBUG] Листа {sheet_name} нет. Ищем шаблон для копирования.")
            
            # 1. Пробуем найти лист с именем месяца (без "Сверка")
            base_sheet_name = sheet_name.replace("Сверка ", "")
            try:
                base_ws = spreadsheet.worksheet(base_sheet_name)
                print(f"[DEBUG] Нашли базовый лист {base_sheet_name}. Переименовываем в {sheet_name}.")
                base_ws.update_title(sheet_name)
                worksheet = base_ws
            except gspread.exceptions.WorksheetNotFound:
                # 2. Если нет, копируем первый лист (надеясь что это шаблон)
                print(f"[DEBUG] Базового листа нет. Копируем первый лист.")
                try:
                    first_sheet = spreadsheet.get_worksheet(0)
                    if first_sheet:
                        worksheet = first_sheet.duplicate(new_sheet_name=sheet_name)
                except:
                    # 3. Совсем беда
                    print(f"[DEBUG] Не удалось скопировать. Создаем пустой.")
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=40)

        # 2. Читаем старые комментарии (Col AK / Index 36)
        old_comments = {} 
        try:
            # Читаем только нужные колонки для скорости, но get_all_values проще
            all_values = worksheet.get_all_values()
            # Пропускаем заголовки (первые 2 строки)
            if len(all_values) > 2:
                for row in all_values[2:]:
                    # Безопасное чтение индекса
                    if len(row) > 5:
                        doc_num = str(row[5]).strip() # Col F
                        comment = ""
                        if len(row) > 36:
                            comment = str(row[36]).strip()
                        
                        if doc_num and comment:
                            old_comments[doc_num] = comment
        except Exception as e:
            print(f"[DEBUG] Ошибка чтения комментариев: {e}")
            pass
            
        # 3. Формируем новые строки
        new_rows_data = []
        for i, item in enumerate(data):
            row_list = [""] * (MAX_COL_IDX + 1)
            
            # Заполняем по карте
            for key, col_idx in COL_MAP_IDX.items():
                val = item.get(key, "")
                row_list[col_idx] = val
            
            # Восстанавливаем комментарий
            doc_num = item.get("supplier_doc", "").strip()
            if doc_num in old_comments:
                if not row_list[36]: 
                    row_list[36] = old_comments[doc_num]
            
            # === ВСТАВКА ИТОГОВ (SUMMARY) В КОЛОНКИ B и C (Индексы 1 и 2) ===
            # Вставляем только в первые 12 строк данных
            if summary:
                if i == 0:
                    row_list[1] = "Оборот IIKO"
                    row_list[2] = summary.get("iiko_total", 0)
                elif i == 1:
                    row_list[1] = "Оборот SAP"
                    row_list[2] = summary.get("sap_total", 0)
                elif i == 2:
                    row_list[1] = "Оборот FB"
                    row_list[2] = summary.get("fb_total", 0)
                elif i == 3:
                    row_list[1] = "Оборот Акт"
                    row_list[2] = summary.get("act_total", 0)
                elif i == 4:
                    row_list[1] = "Дельта (Акт - IIKO)"
                    row_list[2] = summary.get("delta_act_iiko", 0)
                elif i == 5:
                    row_list[1] = "Дельта (Акт - SAP)"
                    row_list[2] = summary.get("delta_act_sap", 0)
                elif i == 6:
                    row_list[1] = "Дельта (Акт - FB)"
                    row_list[2] = summary.get("delta_act_fb", 0)
                elif i == 7:
                    row_list[1] = "Кол-во док. Акт"
                    row_list[2] = summary.get("act_count", 0)
                elif i == 8:
                    row_list[1] = "Кол-во док. IIKO"
                    row_list[2] = summary.get("iiko_count", 0)
                elif i == 9:
                    row_list[1] = "Дельта кол-ва"
                    row_list[2] = summary.get("delta_count", 0)
                elif i == 10:
                    row_list[1] = "Дубли IIKO"
                    val = summary.get("iiko_duplicates", "")
                    row_list[2] = val if val else "Нет"
                elif i == 11:
                    row_list[1] = "Лишние в IIKO"
                    val = summary.get("iiko_missing", "")
                    row_list[2] = val if val else "Нет"
            
            new_rows_data.append(row_list)
            
        # Если строк данных меньше 12, нужно добить пустыми строками, чтобы вывести саммари
        while summary and len(new_rows_data) < 12:
            i = len(new_rows_data)
            row_list = [""] * (MAX_COL_IDX + 1)
            if i == 0:
                row_list[1] = "Оборот IIKO"
                row_list[2] = summary.get("iiko_total", 0)
            elif i == 1:
                row_list[1] = "Оборот SAP"
                row_list[2] = summary.get("sap_total", 0)
            elif i == 2:
                row_list[1] = "Оборот FB"
                row_list[2] = summary.get("fb_total", 0)
            elif i == 3:
                row_list[1] = "Оборот Акт"
                row_list[2] = summary.get("act_total", 0)
            elif i == 4:
                row_list[1] = "Дельта (Акт - IIKO)"
                row_list[2] = summary.get("delta_act_iiko", 0)
            elif i == 5:
                row_list[1] = "Дельта (Акт - SAP)"
                row_list[2] = summary.get("delta_act_sap", 0)
            elif i == 6:
                row_list[1] = "Дельта (Акт - FB)"
                row_list[2] = summary.get("delta_act_fb", 0)
            elif i == 7:
                row_list[1] = "Кол-во док. Акт"
                row_list[2] = summary.get("act_count", 0)
            elif i == 8:
                row_list[1] = "Кол-во док. IIKO"
                row_list[2] = summary.get("iiko_count", 0)
            elif i == 9:
                row_list[1] = "Дельта кол-ва"
                row_list[2] = summary.get("delta_count", 0)
            elif i == 10:
                row_list[1] = "Дубли IIKO"
                val = summary.get("iiko_duplicates", "")
                row_list[2] = val if val else "Нет"
            elif i == 11:
                row_list[1] = "Лишние в IIKO"
                val = summary.get("iiko_missing", "")
                row_list[2] = val if val else "Нет"
            new_rows_data.append(row_list)

        # DEBUG PRINT: Выводим первые 5 строк (для отладки)
        print("\n--- DEBUG: DATA TO WRITE (First 5 rows) ---")
        for i, r in enumerate(new_rows_data[:5]):
            print(f"Row {i+1}: {r}")
        print("-------------------------------------------\n")
            
        # 4. Записываем
        print(f"[DEBUG] Записываем {len(new_rows_data)} строк в {sheet_name} (start A3)")
        
        # Очистка диапазона данных (берем с запасом до 5000 строки)
        # Важно не стереть заголовки (1-2 строки)
        try:
             worksheet.batch_clear([f"A3:AK5000"])
        except Exception as e:
             print(f"[DEBUG] Ошибка очистки: {e}")

        range_start = "A3"
        if new_rows_data:
            worksheet.update(range_name=range_start, values=new_rows_data)
                
        return True, f"Результаты сверки обновлены на листе '{sheet_name}' (комментарии сохранены, шаблон соблюден)"
    except Exception as e:
        print(f"[ERROR] update_supplier_sheet: {e}")
        return False, str(e)
