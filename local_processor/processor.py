import polars as pl
import ollama
import json
import io

class UniversalProcessor:
    def __init__(self, model_name="deepseek-r1:8b"):
        self.model_name = model_name

    def read_excel_raw(self, file_path):
        """Читает Excel и возвращает первые 20 строк как текст для анализа ИИ."""
        df = pl.read_excel(file_path)
        # Берем первые 20 строк и превращаем в JSON для ИИ
        sample_data = df.head(20).to_dicts()
        return sample_data, df

    def analyze_structure(self, sample_data, file_name):
        """Просит локальный ИИ определить структуру таблицы."""
        prompt = f"""
        Ты — эксперт по анализу данных. Перед тобой фрагмент таблицы из файла "{file_name}".
        
        Твоя задача:
        1. Определи, какие индексы колонок (0, 1, 2...) соответствуют следующим данным:
           - Дата документа
           - Номер документа (накл, УПД, №)
           - Описание/Наименование (если есть)
           - Сумма (ищи колонки со значениями типа 1000.00)
        
        2. Определи, с какой строки начинаются реальные данные (пропусти шапку).
        
        Верни ответ СТРОГО в формате JSON:
        {{
            "mapping": {{
                "date_col": index_or_null,
                "doc_number_col": index_or_null,
                "description_col": index_or_null,
                "amount_col": index_or_null
            }},
            "start_row": row_index,
            "file_type": "partner_reconciliation" или "system_export"
        }}

        ДАННЫЕ:
        {json.dumps(sample_data, ensure_ascii=False)}
        """
        
        response = ollama.chat(model=self.model_name, messages=[
            {'role': 'system', 'content': 'Ты возвращаешь только чистый JSON без пояснений.'},
            {'role': 'user', 'content': prompt}
        ])
        
        # Извлекаем JSON из ответа (учитывая возможные теги <think> у DeepSeek-R1)
        content = response['message']['content']
        if "</think>" in content:
            content = content.split("</think>")[-1].strip()
            
        try:
            return json.loads(content)
        except:
            # Если ИИ вернул текст с мусором, пытаемся найти JSON внутри
            import re
            match = re.search(r'\{.*\}', content, re.DOTALL)
            if match:
                return json.loads(match.group())
            raise ValueError(f"Не удалось распарсить ответ ИИ: {content}")

    def process_file(self, file_path):
        """Полный цикл обработки файла."""
        print(f"Читаем файл: {file_path}")
        sample, full_df = self.read_excel_raw(file_path)
        
        print("Анализируем структуру через локальный ИИ...")
        structure = self.analyze_structure(sample, file_path)
        
        print(f"Структура определена: {structure}")
        
        # Начинаем обработку данных на основе структуры от ИИ
        mapping = structure['mapping']
        start_row = structure['start_row']
        
        # Фильтруем и чистим данные через Polars
        # (Здесь мы используем индексы, которые нашел ИИ)
        processed_data = []
        
        # Пример быстрой обработки через Polars
        # Для простоты превращаем в список словарей, но можно оптимизировать
        rows = full_df.to_dicts()
        for i in range(start_row, len(rows)):
            row = rows[i]
            # Динамически берем значения по индексам или ключам
            keys = list(row.keys())
            
            try:
                processed_data.append({
                    "date": row[keys[mapping['date_col']]] if mapping['date_col'] is not None else None,
                    "number": row[keys[mapping['doc_number_col']]] if mapping['doc_number_col'] is not None else None,
                    "description": row[keys[mapping['description_col']]] if mapping['description_col'] is not None else "",
                    "amount": row[keys[mapping['amount_col']]] if mapping['amount_col'] is not None else 0
                })
            except:
                continue
                
        return processed_data, structure['file_type']

if __name__ == "__main__":
    # Тестовый запуск
    # processor = UniversalProcessor()
    # data = processor.process_file("test.xlsx")
    # print(data)
    pass
