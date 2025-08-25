# 🚀 Швидкий старт з Godzilla Bot

## ⚡ Швидке налаштування за 5 кроків

### 1. Встановлення залежностей

```bash
git clone https://github.com/Dmitze/BotDiscordGodzilla.git
cd BotDiscordGodzilla
npm install
```

### 2. Налаштування змінних середовища

Створіть файл `.env` в корені проекту, використовуючи приклад:

```bash
cp .env.example .env
```

Відредагуйте `.env` файл, додавши обов'язкові налаштування:

```env
# Основні налаштування
NODE_ENV=development
LOG_LEVEL=info

# Discord
DISCORD_TOKEN=your_discord_bot_token_here
DISCORD_CLIENT_ID=your_discord_client_id

# Безпека
HMAC_SECRET=generate_secure_random_string_here
COMPONENT_TTL=300000  # 5 хвилин

# База даних
DB_PATH=./data/godzilla.db

# AI (виберіть провайдера: ollama або openai)
AI_PROVIDER=ollama
AI_MODEL=llama3
```

### 3. Налаштування Discord Bot

1. Перейдіть на [Discord Developer Portal](https://discord.com/developers/applications)
2. Створіть новий додаток та бота
3. Увімкніть необхідні Intents:
   - Message Content
   - Server Members
   - Presence
   - Message Content
4. Додайте бота на сервер через OAuth2 URL Generator з дозволами:
   - `bot`
   - `applications.commands`
   - Дозволи: `Send Messages`, `Embed Links`, `Read Message History`

### 4. Налаштування локального AI (Ollama)

1. Встановіть [Ollama](https://ollama.ai/)
2. Завантажте модель:
   ```bash
   ollama pull llama3
   ```
3. Запустіть сервер Ollama:
   ```bash
   ollama serve
   ```

### 5. Запуск бота

```bash
# Розробка з автоматичним перезавантаженням
npm run dev

# Або звичайний запуск
npm start
```

## 🧩 Основні можливості

### Пошук та аналіз
- `/search [запит]` - пошук серед індексованих документів
- `/analyze [текст]` - аналіз тексту за допомогою AI

### Робота з документами
- `/index [url]` - індексувати документ за посиланням
- `/list` - перегляд індексованих документів

### Налаштування
- `/settings` - керування налаштуваннями бота
- `/help` - довідка за командами

## 🧪 Тестування роботи

Переконайтеся, що бот працює коректно:

1. Відправте `/help` у чат Discord
2. Перевірте наявність команд автодоповнення
3. Протестуйте базові команди:
   ```bash
   /search тестовий пошук
   /analyze Це тестовий аналіз тексту
   ```

## 🔍 Наступні кроки

- Додайте більше документів для індексації
- Налаштуйте канали для роботи з ботом
- Вивчіть розширені можливості у повній документації

## ❓ Допомога

Якщо виникли питання:
1. Перевірте логи у консолі
2. Перегляньте розділ "Поширені проблеми" у документації
3. Створіть Issue у репозиторії
- `/ai-пошук запит:покажи товари дешевше 1000` - природномовний пошук

## 📋 Мінімальні вимоги

### Для роботи без AI:
- Discord Bot Token
- Google API Key
- Google Sheet ID

### Для повного функціоналу:
- + OpenAI API Key

## 🔧 Швидке виправлення проблем

### Бот не відповідає:
```bash
# Перевірте логи
tail -f logs/bot.log

# Перевірте змінні середовища
echo $BOT_TOKEN
```

### Помилка Google Sheets:
```bash
# Перевірте доступ до таблиці
curl "https://sheets.googleapis.com/v4/spreadsheets/YOUR_SHEET_ID/values/A1?key=YOUR_API_KEY"
```

### AI не працює:
```bash
# Тест AI-функцій
node test-ai.js
```

## 📚 Детальна документація

- [README.md](README.md) - повна документація
- [SETUP.md](SETUP.md) - детальне налаштування
- [AI_EXAMPLES.md](AI_EXAMPLES.md) - приклади AI-функцій

## 🆘 Підтримка

Проблеми? Створіть Issue в репозиторії з:
- Описом проблеми
- Логами з папки `logs/`
- Версією Node.js (`node --version`) 