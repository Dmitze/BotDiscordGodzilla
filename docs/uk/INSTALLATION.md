# 🛠️ Встановлення та налаштування
# Посібник з встановлення Discord AI Assistant Bot

## 📋 Системні вимоги

Перед встановленням бота переконайтесь, що у вашій системі встановлено:

- **Node.js** 18 або новіша версія
- **Git** для клонування репозиторію
- **Docker** (опційно, для контейнеризації)
- **Google Cloud обліковий запис** для інтеграції з Google Sheets/Drive
- **Discord Developer обліковий запис** для створення бота

## 🚀 Встановлення

### 1. Клонування проекту

Відкрийте термінал та виконайте команди:

```bash
git clone https://github.com/Dmitze/BotDiscordGodzilla.git
cd BotDiscordGodzilla
```

### 2. Встановлення залежностей

```bash
npm install
```

### 3. Налаштування змінних середовища

Скопіюйте приклад файлу конфігурації:

```bash
cp .env.example .env
```

Відредагуйте файл `.env` та заповніть необхідні змінні:

```env
# Discord налаштування
DISCORD_TOKEN=ваш_discord_токен
DISCORD_CLIENT_ID=ваш_client_id
DISCORD_GUILD_ID=ваш_guild_id

# Google API налаштування
GOOGLE_API_KEY=ваш_google_api_ключ
GOOGLE_SERVICE_ACCOUNT_KEY=шлях_до_файлу_сервісного_акаунта.json
GOOGLE_SPREADSHEET_ID=ідентифікатор_таблиці

# AI налаштування (опційно)
OPENAI_API_KEY=ваш_openai_api_ключ
AI_PROVIDER=ollama # або openai
OLLAMA_BASE_URL=http://localhost:11434

# Інші налаштування
REDIS_URL=redis://localhost:6379
DATABASE_URL=sqlite://data/database.sqlite
```

### 4. Налаштування Discord бота

1. Перейдіть на [Discord Developer Portal](https://discord.com/developers/applications)
2. Натисніть "New Application" та введіть назву вашого бота
3. У розділі "Bot" натисніть "Add Bot"
4. Скопіюйте токен та вставте його в `.env` файл
5. Увімкніть наступні "Privileged Gateway Intents":
   - Presence Intent
   - Server Members Intent
   - Message Content Intent

### 5. Налаштування Google API

1. Перейдіть на [Google Cloud Console](https://console.cloud.google.com/)
2. Створіть новий проект або виберіть існуючий
3. Увімкніть наступні API:
   - Google Sheets API
   - Google Drive API
4. Створіть Service Account:
   - Перейдіть до "IAM & Admin" → "Service Accounts"
   - Натисніть "Create Service Account"
   - Введіть ім'я та натисніть "Create"
   - Натисніть "Create Key" та виберіть тип "JSON"
   - Завантажте файл та вкажіть шлях до нього в `.env`
5. Надайте доступ до вашої Google Sheet:
   - Відкрийте таблицю
   - Натисніть "Share"
   - Додайте email вашого Service Account з правами редактора

### 6. Налаштування локального AI (Ollama)

Для локального використання AI моделей:

1. Встановіть Ollama з [офиційного сайту](https://ollama.com/)
2. Завантажте модель:
   ```bash
   ollama run llama3.1
   ```
3. Вкажіть налаштування в `.env`:
   ```env
   AI_PROVIDER=ollama
   OLLAMA_BASE_URL=http://localhost:11434
   ```

## ▶️ Запуск бота

### Розробницький режим

```bash
npm run dev
```

### Продакшн режим

```bash
npm run build
npm start
```

### Docker (опційно)

```bash
docker-compose up -d
```

## ✅ Перевірка встановлення

Після запуску бота:

1. Перевірте, що бот онлайн у вашому Discord сервері
2. Спробуйте виконати команду `/допомога`
3. Перевірте логи на наявність помилок

## 🔧 Усунення проблем

### Бот не відповідає

- Перевірте правильність Discord токена
- Переконайтесь, що бот доданий до сервера
- Перевірте налаштування intents

### Помилки з Google API

- Перевірте правильність шляху до файлу Service Account
- Переконайтесь, що Service Account має доступ до таблиці
- Перевірте, чи увімкнені необхідні API

### Проблеми з AI

- Перевірте, чи запущений Ollama (якщо використовуєте локальні моделі)
- Перевірте правильність API ключів
- Переконайтесь, що обрана правильна модель

## 🔄 Оновлення

Для оновлення бота до останньої версії:

```bash
git pull
npm install
npm run build
npm run deploy
```

## 📞 Підтримка

Якщо у вас виникли проблеми з встановленням:

1. Перевірте логи у папці `logs/`
2. Створіть issue у репозиторії GitHub
3. Зверніться до підтримки через Discord або Telegram

© 2025 Dmitry Shivachov (Dmitze). Всі права захищені.