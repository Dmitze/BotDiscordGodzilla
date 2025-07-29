# 🔧 Налаштування Discord Bot з AI

## 📋 Необхідні змінні середовища

Створіть файл `.env` в корені проекту з наступними змінними:

```env
# Discord Bot Token
BOT_TOKEN=your_discord_bot_token_here

# Google Sheets Configuration
SHEET_ID=your_google_sheet_id_here
GOOGLE_API_KEY=your_google_api_key_here
APP_SCRIPT_URL=your_google_apps_script_url_here

# OpenAI Configuration (для AI-функціоналу)
OPENAI_API_KEY=your_openai_api_key_here

# Optional: Logging Configuration
LOG_LEVEL=info
```

## 🔑 Отримання API ключів

### 1. Discord Bot Token

1. Перейдіть на [Discord Developer Portal](https://discord.com/developers/applications)
2. Натисніть "New Application"
3. Введіть назву для вашого бота
4. Перейдіть в розділ "Bot" в лівому меню
5. Натисніть "Add Bot"
6. Скопіюйте токен (натисніть "Copy" біля "Token")
7. Увімкніть всі необхідні Intents:
   - Message Content Intent
   - Server Members Intent
   - Presence Intent

### 2. Google API Key

1. Перейдіть в [Google Cloud Console](https://console.cloud.google.com/)
2. Створіть новий проект або виберіть існуючий
3. Увімкніть Google Sheets API:
   - Перейдіть в "APIs & Services" > "Library"
   - Знайдіть "Google Sheets API"
   - Натисніть "Enable"
4. Створіть API ключ:
   - Перейдіть в "APIs & Services" > "Credentials"
   - Натисніть "Create Credentials" > "API Key"
   - Скопіюйте ключ

### 3. OpenAI API Key

1. Перейдіть на [OpenAI Platform](https://platform.openai.com/)
2. Створіть обліковий запис або увійдіть
3. Перейдіть в розділ "API Keys"
4. Натисніть "Create new secret key"
5. Скопіюйте ключ

## 📊 Налаштування Google Sheets

### 1. Створення таблиці

1. Створіть нову Google Sheet
2. Додайте заголовки в перший рядок (наприклад):
   - Найменування номенклатури
   - Серійний номер
   - Контрагент
   - Кількість
   - Ціна
   - Вартість

### 2. Отримання ID таблиці

1. Відкрийте вашу Google Sheet
2. Скопіюйте ID з URL:
   ```
   https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit
   ```

### 3. Налаштування доступу

1. Натисніть "Share" в правому верхньому куті
2. Додайте ваш Google API ключ як редактор
3. Або налаштуйте публічний доступ (тільки для читання)

## 🤖 Налаштування Discord Bot

### 1. Додавання бота на сервер

1. В Discord Developer Portal перейдіть в розділ "OAuth2" > "URL Generator"
2. Виберіть scopes: `bot`, `applications.commands`
3. Виберіть permissions:
   - Send Messages
   - Use Slash Commands
   - Embed Links
   - Attach Files
   - Read Message History
4. Скопіюйте згенерований URL
5. Відкрийте URL в браузері та додайте бота на ваш сервер

### 2. Налаштування команд

Бот автоматично зареєструє slash-команди при запуску. Якщо потрібно оновити команди:

```bash
node index.js
```

## 🚀 Запуск бота

### 1. Встановлення залежностей

```bash
npm install
```

### 2. Запуск

```bash
node index.js
```

Або використовуйте PowerShell скрипт:

```powershell
.\запуск-бота.ps1
```

## 🔍 Тестування

### Основні команди для тестування:

1. `/help` - перевірка роботи бота
2. `/залишки` - перевірка підключення до Google Sheets
3. `/ai-аналіз` - перевірка AI-функціоналу

### Перевірка логів

Перевіряйте файли в папці `logs/` для діагностики проблем.

## ❗ Поширені проблеми

### 1. "Invalid token"
- Перевірте правильність Discord Bot Token
- Переконайтеся, що токен не містить зайвих символів

### 2. "Google Sheets API error"
- Перевірте правильність Google API Key
- Увімкніть Google Sheets API в Google Cloud Console
- Перевірте доступ до таблиці

### 3. "OpenAI API error"
- Перевірте правильність OpenAI API Key
- Переконайтеся, що у вас є кредити на рахунку OpenAI

### 4. "Command not found"
- Перезапустіть бота після зміни команд
- Перевірте, що бот має права на використання slash-команд

## 📞 Підтримка

Якщо у вас виникли проблеми:
1. Перевірте логи в папці `logs/`
2. Переконайтеся, що всі змінні середовища налаштовані правильно
3. Перевірте підключення до інтернету
4. Створіть Issue в репозиторії з детальним описом проблеми 