# 🔧 Налаштування Godzilla Bot

## 📋 Обов'язкові змінні середовища

Створіть файл `.env` в корені проекту, використавши `env.example` як основу:

```env
# ===== ОСНОВНІ НАЛАШТУВАННЯ =====
NODE_ENV=development
LOG_LEVEL=info

# ===== DISCORD =====
DISCORD_TOKEN=your_discord_bot_token_here
DISCORD_CLIENT_ID=your_discord_client_id
DISCORD_GUILD_ID=your_guild_id  # Опційно, для розробки

# ===== БЕЗПЕКА =====
HMAC_SECRET=generate_secure_random_string
COMPONENT_TTL=300000  # 5 хвилин у мілісекундах

# ===== БАЗА ДАНИХ =====
DB_PATH=./data/godzilla.db
DB_MIGRATIONS_PATH=./src/db/migrations

# ===== ПОШУК ТА ІНДЕКСАЦІЯ =====
SEARCH_INDEX_PATH=./data/search-index
SEARCH_BATCH_SIZE=50
SEARCH_FTS_TOKENIZER=unicode61  # Або porter для англійської

# ===== AI ТА RAG =====
# Виберіть провайдера: openai, ollama, mock
AI_PROVIDER=ollama
AI_MODEL=llama3  # Або інша модель
AI_MAX_TOKENS=4096
AI_TEMPERATURE=0.7

# Налаштування RAG (Retrieval-Augmented Generation)
RAG_ENABLED=true
RAG_MAX_CONTEXT_TOKENS=2000
RETRIEVER_K=5
RETRIEVER_ALPHA=0.5  # Вага гібридного пошуку (0-1)

# ===== ВСТАНОВЛЕННЯ (для Ollama) =====
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=llama3  # Вкажіть бажану модель

# ===== МЕТРИКИ ТА МОНІТОРИНГ =====
METRICS_ENABLED=true
PROMETHEUS_ENABLED=true
PROMETHEUS_PORT=9091
```

## 🔑 Налаштування Discord Bot

### 1. Отримання Discord Token

1. Перейдіть на [Discord Developer Portal](https://discord.com/developers/applications)
2. Натисніть "New Application" та введіть назву
3. У розділі "Bot" створіть нового бота
4. Скопіюйте токен (натисніть "Copy" біля "Token")
5. Увімкніть необхідні Intents:
   - ✅ Message Content Intent
   - ✅ Server Members Intent
   - ✅ Presence Intent
   - ✅ Message Content

### 2. Додавання бота на сервер

1. У розділі "OAuth2" > "URL Generator" оберіть:
   - `bot`
   - `applications.commands`
   - Дозволи: `Send Messages`, `Embed Links`, `Read Message History`
2. Згенеруйте URL та запросіть адміністратора додати бота

## ⚙️ Налаштування бази даних

1. SQLite3 встановлюється автоматично
2. База даних створиться за шляхом `./data/godzilla.db`
3. Для міграцій використовуйте:

   ```bash
   npm run db:migrate
   ```

## 🔑 Отримання API ключів

### OpenAI API Key

1. Перейдіть на [OpenAI Platform](https://platform.openai.com/)
2. Увійдіть або створіть обліковий запис
3. Перейдіть у розділ "API Keys"
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

## Підтримка

Якщо у вас виникли проблеми:

1. Перевірте логи в папці `logs/`
2. Переконайтеся, що всі змінні середовища налаштовані правильно
3. Перевірте підключення до інтернету
4. Створіть Issue в репозиторії з детальним описом проблеми