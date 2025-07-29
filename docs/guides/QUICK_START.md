# 🚀 Швидкий старт Discord Bot з AI

## ⚡ Швидка настройка за 5 кроків

### 1. Клонування та встановлення
```bash
git clone https://github.com/Dmitze/BotDiscordGodzilla.git
cd BotDiscordGodzilla
npm install
```

### 2. Створення файлу .env
Створіть файл `.env` в корені проекту:
```env
BOT_TOKEN=your_discord_bot_token
SHEET_ID=your_google_sheet_id
GOOGLE_API_KEY=your_google_api_key
APP_SCRIPT_URL=your_google_apps_script_url
OPENAI_API_KEY=your_openai_api_key
```

### 3. Отримання API ключів (5 хвилин)

#### Discord Bot Token:
1. [Discord Developer Portal](https://discord.com/developers/applications) → New Application
2. Bot → Add Bot → Copy Token

#### Google API Key:
1. [Google Cloud Console](https://console.cloud.google.com/) → New Project
2. APIs & Services → Library → Google Sheets API → Enable
3. Credentials → Create Credentials → API Key

#### OpenAI API Key:
1. [OpenAI Platform](https://platform.openai.com/) → API Keys
2. Create new secret key

### 4. Налаштування Google Sheets
1. Створіть Google Sheet з даними
2. Скопіюйте ID з URL: `https://docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit`
3. Надайте доступ для API ключа

### 5. Запуск бота
```bash
node index.js
```

## 🧪 Тестування

### Перевірка AI-функціоналу:
```bash
node test-ai.js
```

### Основні команди для тестування:
- `/help` - список всіх команд
- `/ai-аналіз` - AI-аналіз даних
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