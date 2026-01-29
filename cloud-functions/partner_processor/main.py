import base64
import io
import json
import os
import re
import traceback
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests


DATE_RE = re.compile(r"^\d{1,2}[./]\d{1,2}[./]\d{2,4}$")
NUMERIC_RE = re.compile(r"^-?\d+([ \u00A0]\d{3})*(?:[.,]\d+)?$")


def handler(event, context):
    code_version = get_code_version()
    print(f"Code version: {code_version}")
    try:
        body = event.get("body") or ""
        if event.get("isBase64Encoded"):
            body = base64.b64decode(body).decode("utf-8")
        payload = json.loads(body)
    except Exception as exc:
        return _response(400, {"error": f"Invalid request body: {exc}"})

    file_name = payload.get("fileName")
    file_b64 = payload.get("fileBase64")
    options = payload.get("options") or {}
    if not file_b64:
        return _response(400, {"error": "fileBase64 is required"})

    try:
        file_bytes = base64.b64decode(file_b64)
    except Exception as exc:
        return _response(400, {"error": f"Invalid base64: {exc}"})

    try:
        rows = process_excel(file_bytes, file_name or "file", options)
    except Exception as exc:
        safe_error = str(exc).encode("ascii", "backslashreplace").decode("ascii")
        safe_traceback = traceback.format_exc().encode("ascii", "backslashreplace").decode("ascii")
        print(f"Processing failed: {safe_error}")
        print(f"Traceback: {safe_traceback}")
        return _response(500, {"error": f"Processing failed: {safe_error}"})

    return _response(200, {"rows": rows, "meta": {"rowCount": len(rows)}})


def process_excel(file_bytes: bytes, file_name: str, options: Dict[str, Any]):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    all_rows: List[List[str]] = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet_name=sheet, header=None, dtype=str)
        data = df.fillna("").values.tolist()
        if not data:
            continue
        if options.get("llmExtract"):
            rows = extract_rows_llm(data, file_name, sheet, options)
        else:
            rows = extract_rows(data, file_name, options)
        all_rows.extend(rows)

    if options.get("semantic", True):
        all_rows = semantic_filter(all_rows, options)

    return all_rows


def extract_rows(
    data: List[List[str]], file_name: str, options: Dict[str, Any]
) -> List[List[str]]:
    blocks = detect_blocks(data)
    if blocks:
        rows = extract_from_blocks(data, blocks)
    else:
        columns = detect_columns(data)
        rows = extract_from_columns(data, columns)

    number_mode = options.get("numberMode", "regex_first")
    rows = apply_number_extraction(rows, number_mode, options)
    return rows


def extract_rows_llm(
    data: List[List[str]],
    file_name: str,
    sheet_name: str,
    options: Dict[str, Any],
) -> List[List[str]]:
    api_key, folder_id, model = get_yandex_config()
    max_chars = int(options.get("llmMaxChars", 120000))
    max_rows = int(options.get("llmMaxRows", 500))
    header_rows = int(options.get("llmHeaderRows", 8))
    max_cell_len = int(options.get("llmCellMax", 120))

    rows_payload = build_rows_payload(data, max_cell_len)

    if not rows_payload:
        return []

    user_text = json.dumps(
        {"fileName": file_name, "sheetName": sheet_name, "rows": rows_payload},
        ensure_ascii=True,
    )
    print(
        "LLM extract input chars: %s rows: %s sheet: %s"
        % (len(user_text), len(rows_payload), sheet_name)
    )
    if len(user_text) > max_chars:
        rows_payload = compress_rows_for_llm(
            data,
            header_rows=header_rows,
            max_rows=max_rows,
            max_cell_len=max_cell_len,
        )
        user_text = json.dumps(
            {"fileName": file_name, "sheetName": sheet_name, "rows": rows_payload},
            ensure_ascii=True,
        )
        print(
            "LLM extract compressed chars: %s rows: %s sheet: %s"
            % (len(user_text), len(rows_payload), sheet_name)
        )
    if len(user_text) > max_chars:
        raise RuntimeError(
            "LLM extract payload too large for single request: "
            f"{len(user_text)} chars > {max_chars} chars"
        )

    payload = {
        "modelUri": f"gpt://{folder_id}/{model}",
        "completionOptions": {"stream": False, "temperature": 0, "maxTokens": 1200},
        "messages": [
            {
                "role": "system",
                "text": (
                    "Ты извлекаешь строки сверки поставщика из таблицы. "
                    "Верни только JSON массив объектов "
                    "{id:number, date:string, text:string, number:string, sum:string}. "
                    "Исключай строки без даты/суммы/текста. "
                    "sum верни числом в строке, точка как разделитель."
                ),
            },
            {"role": "user", "text": user_text},
        ],
    }
    json_payload_bytes = json.dumps(payload, ensure_ascii=True).encode("utf-8")
    response = requests.post(
        "https://llm.api.cloud.yandex.net/foundationModels/v1/completion",
        headers={
            "Authorization": f"Api-Key {api_key}",
            "Content-Type": "application/json; charset=utf-8",
        },
        data=json_payload_bytes,
        timeout=120,
    )
    response.raise_for_status()
    message = response.json()["result"]["alternatives"][0]["message"]["text"]
    items = parse_json_array(message)
    results: List[List[str]] = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        date_val = str(item.get("date") or "").strip()
        text_val = str(item.get("text") or "").strip()
        number_val = str(item.get("number") or "").strip()
        sum_val = normalize_sum(str(item.get("sum") or "").strip())
        if not (date_val and text_val and sum_val):
            continue
        results.append([date_val, text_val, number_val, sum_val])

    return results


def detect_blocks(data: List[List[str]]):
    header_row_index = 0
    for i in range(min(len(data), 20)):
        row = [normalize_header(x) for x in data[i]]
        if "дата" in row and "документ" in row and ("дебет" in row or "кредит" in row):
            header_row_index = i + 1
            break
    if not header_row_index:
        return []

    header_row = data[header_row_index - 1]
    blocks = []
    for idx, cell in enumerate(header_row):
        if normalize_header(cell) != "дата":
            continue
        doc_col = find_header_offset(header_row, idx + 1, "документ")
        debit_col = find_header_offset(header_row, idx + 1, "дебет")
        credit_col = find_header_offset(header_row, idx + 1, "кредит")
        if doc_col and (debit_col or credit_col):
            blocks.append(
                {
                    "dateCol": idx + 1,
                    "docCol": doc_col,
                    "debitCol": debit_col,
                    "creditCol": credit_col,
                    "headerRowIndex": header_row_index,
                }
            )
    return blocks


def extract_from_blocks(data: List[List[str]], blocks: List[Dict[str, int]]):
    start_row = blocks[0]["headerRowIndex"]
    rows = []
    for i in range(start_row, len(data)):
        row = data[i]
        if not "".join(str(x) for x in row).strip():
            continue
        for block in blocks:
            date_val = get_cell(row, block["dateCol"])
            doc_text = get_cell(row, block["docCol"])
            debit = normalize_sum(get_cell(row, block["debitCol"]))
            credit = normalize_sum(get_cell(row, block["creditCol"]))
            sum_val = debit or credit
            if not (is_date(date_val) and sum_val and doc_text):
                continue
            rows.append([date_val, doc_text, "", sum_val])
    return rows


def detect_columns(data: List[List[str]]):
    column_count = len(data[0]) if data else 0
    scores = []
    for col in range(column_count):
        date_score = sum_score = text_score = 0
        for row in data[1:200]:
            val = get_cell(row, col + 1)
            if not val:
                continue
            if is_date(val):
                date_score += 1
            if is_numeric(val):
                sum_score += 1
            if re.search(r"[A-Za-zА-Яа-я]", val) or "№" in val:
                text_score += 1
        scores.append((date_score, sum_score, text_score))

    date_col = pick_best(scores, 0)
    sum_col = pick_best(scores, 1, exclude=[date_col])
    text_col = pick_best(scores, 2, exclude=[date_col, sum_col])
    return {"date": date_col, "sum": sum_col, "text": text_col}


def extract_from_columns(data: List[List[str]], columns: Dict[str, int]):
    rows = []
    for row in data:
        date_val = get_cell(row, columns["date"])
        sum_val = normalize_sum(get_cell(row, columns["sum"]))
        text_val = get_cell(row, columns["text"])
        if not (is_date(date_val) and is_numeric(sum_val) and text_val):
            continue
        rows.append([date_val, text_val, "", sum_val])
    return rows


def apply_number_extraction(rows: List[List[str]], mode: str, options: Dict[str, Any]):
    if mode == "regex_only":
        for row in rows:
            row[2] = extract_number_regex(row[1])
        return rows
    if mode == "llm_only":
        numbers = extract_numbers_llm([r[1] for r in rows], options)
        for row, num in zip(rows, numbers):
            row[2] = num or ""
        return rows

    # regex_first
    missing_indexes = []
    for idx, row in enumerate(rows):
        num = extract_number_regex(row[1])
        row[2] = num or ""
        if not num:
            missing_indexes.append(idx)
    if missing_indexes:
        texts = [rows[i][1] for i in missing_indexes]
        numbers = extract_numbers_llm(texts, options)
        for idx, num in zip(missing_indexes, numbers):
            rows[idx][2] = num or ""
    return rows


def extract_numbers_llm(texts: List[str], options: Dict[str, Any]):
    if not texts:
        return []
    api_key, folder_id, model = get_yandex_config()
    payload = {
        "modelUri": f"gpt://{folder_id}/{model}",
        "completionOptions": {"stream": False, "temperature": 0, "maxTokens": 800},
        "messages": [
            {
                "role": "system",
                "text": (
                    "Извлеки номер документа из текста. Верни JSON массив "
                    "объектов {id:number, number:string}. Только JSON."
                ),
            },
            {
                "role": "user",
                "text": json.dumps(
                    [{"id": i, "text": t} for i, t in enumerate(texts)],
                    ensure_ascii=True,
                ),
            },
        ],
    }
    json_payload_bytes = json.dumps(payload, ensure_ascii=True).encode("utf-8")
    response = requests.post(
        "https://llm.api.cloud.yandex.net/foundationModels/v1/completion",
        headers={
            "Authorization": f"Api-Key {api_key}",
            "Content-Type": "application/json; charset=utf-8",
        },
        data=json_payload_bytes,
        timeout=120,
    )
    response.raise_for_status()
    message = response.json()["result"]["alternatives"][0]["message"]["text"]
    items = parse_json_array(message)
    results = [""] * len(texts)
    for item in items:
        if isinstance(item, dict) and "id" in item:
            results[int(item["id"])] = str(item.get("number") or "")
    return results


def semantic_filter(rows: List[List[str]], options: Dict[str, Any]):
    api_key, folder_id, model = get_yandex_config()
    batch_size = int(options.get("semanticBatch", 200))
    filtered = []
    for i in range(0, len(rows), batch_size):
        batch = rows[i : i + batch_size]
        payload = {
            "modelUri": f"gpt://{folder_id}/{model}",
            "completionOptions": {"stream": False, "temperature": 0, "maxTokens": 800},
            "messages": [
                {
                    "role": "system",
                    "text": (
                        "Классифицируй строки сверки. Оставь только расходы клиента "
                        "и корректировки. Исключи оплаты/платежи/поручения. "
                        "Верни JSON массив {id:number, include:boolean}."
                    ),
                },
                {
                    "role": "user",
                    "text": json.dumps(
                        [{"id": idx, "text": row[1]} for idx, row in enumerate(batch)],
                        ensure_ascii=True,
                    ),
                },
            ],
        }
        json_payload_bytes = json.dumps(payload, ensure_ascii=True).encode("utf-8")
        response = requests.post(
            "https://llm.api.cloud.yandex.net/foundationModels/v1/completion",
            headers={
                "Authorization": f"Api-Key {api_key}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=json_payload_bytes,
            timeout=120,
        )
        response.raise_for_status()
        message = response.json()["result"]["alternatives"][0]["message"]["text"]
        items = parse_json_array(message)
        allowed = {int(item["id"]) for item in items if item.get("include")}
        for idx, row in enumerate(batch):
            if idx in allowed:
                filtered.append(row)
    return filtered


def parse_json_array(text: str):
    trimmed = text.strip()
    fenced = re.search(r"```(?:json)?\s*([\s\S]*?)```", trimmed, re.IGNORECASE)
    if fenced:
        trimmed = fenced.group(1).strip()
    start = trimmed.find("[")
    end = trimmed.rfind("]")
    if start != -1 and end != -1:
        try:
            return json.loads(trimmed[start : end + 1])
        except Exception:
            pass
    items = []
    for match in re.finditer(r"\{[^}]*\}", trimmed):
        try:
            items.append(json.loads(match.group(0)))
        except Exception:
            continue
    return items


def normalize_sum(value: str) -> str:
    if not value:
        return ""
    return value.replace(" ", "").replace("\u00A0", "").replace(",", ".")


def build_row_text(row: List[str], max_cell_len: int = 120) -> str:
    parts = []
    for cell in row:
        if cell is None:
            continue
        text = str(cell).strip()
        if not text:
            continue
        if len(text) > max_cell_len:
            text = text[:max_cell_len] + "..."
        parts.append(text)
    return " | ".join(parts)


def build_rows_payload(data: List[List[str]], max_cell_len: int) -> List[Dict[str, Any]]:
    rows_payload = []
    for idx, row in enumerate(data):
        row_text = build_row_text(row, max_cell_len)
        if row_text:
            rows_payload.append({"id": idx, "text": row_text})
    return rows_payload


def compress_rows_for_llm(
    data: List[List[str]],
    header_rows: int,
    max_rows: int,
    max_cell_len: int,
) -> List[Dict[str, Any]]:
    candidates = []
    for idx, row in enumerate(data):
        if idx < header_rows:
            row_text = build_row_text(row, max_cell_len)
            if row_text:
                candidates.append((idx, row_text, 1000))
            continue
        score = row_signal_score(row)
        if score <= 0:
            continue
        row_text = build_row_text(row, max_cell_len)
        if row_text:
            candidates.append((idx, row_text, score))

    if max_rows and len(candidates) > max_rows:
        top = sorted(candidates, key=lambda x: x[2], reverse=True)[:max_rows]
        selected = sorted(top, key=lambda x: x[0])
    else:
        selected = sorted(candidates, key=lambda x: x[0])

    return [{"id": idx, "text": text} for idx, text, _ in selected]


def row_signal_score(row: List[str]) -> int:
    score = 0
    nonempty = 0
    joined_parts = []
    for cell in row:
        if cell is None:
            continue
        text = str(cell).strip()
        if not text:
            continue
        nonempty += 1
        joined_parts.append(text)
        if is_date(text):
            score += 3
        if is_numeric(text):
            score += 2
        if "№" in text:
            score += 1
    if not nonempty:
        return 0
    joined = " ".join(joined_parts).lower()
    for keyword in ("дата", "документ", "дебет", "кредит", "сумма", "итого", "сальдо"):
        if keyword in joined:
            score += 2
    score += min(nonempty, 5)
    return score


def is_date(value: str) -> bool:
    return bool(value and DATE_RE.match(value.strip()))


def is_numeric(value: str) -> bool:
    if not value:
        return False
    text = value.replace(" ", "").replace("\u00A0", "")
    return bool(NUMERIC_RE.match(text))


def extract_number_regex(text: str) -> str:
    if not text:
        return ""
    m = re.search(r"№\s*([A-Za-zА-Яа-я0-9/-]+)", text)
    if m:
        return m.group(1)
    m = re.search(r"\b\d{2,}\b", text)
    if m:
        return m.group(0)
    m = re.search(r"[A-Za-zА-Яа-я0-9/-]{3,}", text)
    return m.group(0) if m else ""


def get_cell(row: List[str], col: int) -> str:
    if not col or col < 1 or col > len(row):
        return ""
    return str(row[col - 1]).strip() if row[col - 1] is not None else ""


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower().replace("«", "").replace("»", "").replace('"', "").replace("'", "")


def find_header_offset(header_row: List[str], start_index: int, header_name: str) -> int:
    for idx in range(start_index, len(header_row)):
        if normalize_header(header_row[idx]) == header_name:
            return idx + 1
    return 0


def pick_best(scores: List[Tuple[int, int, int]], key_index: int, exclude=None):
    exclude = exclude or []
    best_col = 0
    best_score = -1
    for col, score_tuple in enumerate(scores, start=1):
        if col in exclude:
            continue
        score = score_tuple[key_index]
        if score > best_score:
            best_score = score
            best_col = col
    return best_col


def get_yandex_config():
    api_key = os.getenv("YANDEX_API_KEY")
    folder_id = os.getenv("YANDEX_FOLDER_ID")
    model = os.getenv("YANDEX_MODEL", "yandexgpt-lite/latest")
    if not api_key or not folder_id:
        raise RuntimeError("YANDEX_API_KEY and YANDEX_FOLDER_ID are required")
    ensure_ascii(api_key, "YANDEX_API_KEY")
    ensure_ascii(folder_id, "YANDEX_FOLDER_ID")
    ensure_ascii(model, "YANDEX_MODEL")
    return api_key, folder_id, model


def ensure_ascii(value: str, name: str) -> None:
    try:
        value.encode("ascii")
    except UnicodeEncodeError:
        offenders = [
            f"{idx}:U+{ord(ch):04X}"
            for idx, ch in enumerate(value)
            if ord(ch) > 127
        ]
        tail = "..." if len(offenders) > 10 else ""
        raise RuntimeError(
            f"{name} contains non-ASCII characters at {', '.join(offenders[:10])}{tail}"
        )


def _response(status: int, payload: Dict[str, Any]):
    payload = dict(payload)
    meta = payload.get("meta") or {}
    meta["codeVersion"] = get_code_version()
    payload["meta"] = meta
    return {
        "statusCode": status,
        "headers": {
            "Content-Type": "application/json",
            "X-Code-Version": get_code_version(),
        },
        "body": json.dumps(payload, ensure_ascii=True),
    }


def get_code_version() -> str:
    return os.getenv("CODE_VERSION", "unknown")
