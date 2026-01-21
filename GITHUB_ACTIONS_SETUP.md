# ✅ GitHub Actions: деплой в Google Apps Script (итоговая схема)

Этот документ фиксирует рабочую конфигурацию GitHub Actions для автоматического деплоя Apps Script в этом проекте.

## 1) Где лежит workflow

Основной workflow:

- `.github/workflows/deploy-to-apps-script.yml`

Есть альтернативные/тестовые:

- `.github/workflows/deploy-to-apps-script-v2.yml`
- `.github/workflows/deploy-to-apps-script-alt.yml`

Используйте **`deploy-to-apps-script.yml`** как “истину”.

## 2) Что деплоится

Workflow пушит код в Apps Script при изменении файлов.

Триггеры:

- push в соответствующие ветки проекта
- manual trigger (`workflow_dispatch`)

## 3) Секреты (GitHub → Settings → Secrets and variables → Actions)

Нужны два секрета:

1) `CLASP_SCRIPT_ID`  
   - Script ID проекта Apps Script (из URL скрипта)

2) `CLASP_TOKEN`  
   - **Base64** от файла `~/.clasprc.json`
   - Используется для авторизации `clasp push`

## 4) Как получить `CLASP_TOKEN`

1. Локально авторизоваться в clasp:
   ```bash
   clasp login
   ```
2. Взять `~/.clasprc.json` и закодировать в base64:
   ```bash
   base64 -i ~/.clasprc.json | pbcopy
   ```
3. Вставить результат в `CLASP_TOKEN` (Secrets).

## 5) Что делает workflow (кратко)

- Checkout кода
- Установка Node.js 18
- Установка `@google/clasp@3.1.3`
- Создание `.clasp.json` с `scriptId`
- Декодирование `CLASP_TOKEN` в `$HOME/.config/.clasprc.json`
- Копирование в `$HOME/.clasprc.json` для совместимости
- `clasp push --force`

## 6) Проверка деплоя

После пуша:

1. GitHub → Actions → “Deploy to Google Apps Script”
2. Убедиться, что шаг `clasp push --force` прошёл без ошибок.

В Apps Script можно проверить:

- **Deploy → Manage deployments** (код обновился)

## 7) Частые проблемы и решения

### ❌ "Invalid JSON" / "Missing access_token"
Причина: плохой `CLASP_TOKEN` (не base64 или битый).

Решение: пересоздать токен (см. раздел 4).

### ❌ "URL is not in the script manifest whitelist"
Причина: `appsscript.json` не задеплоился.

Решение:

- убедиться, что изменяется именно `appsscript.json`
- workflow должен запускаться (push + path filter)

### ❌ Changes не применились в Apps Script
Причина: workflow не запускался (нет изменения файлов из списка).

Решение: менять `appsScript.js`, `dialog.html` или `appsscript.json`, чтобы workflow триггерился.

## 8) Рекомендация на будущее

Держите `deploy-to-apps-script.yml` как основной.  
Альтернативные workflow используйте только для отладки и тестов.

