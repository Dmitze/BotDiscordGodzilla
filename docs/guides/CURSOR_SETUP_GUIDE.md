# 🚀 Гід по налаштуванню Cursor AI для Discord Bot Project

## 📋 Що було налаштовано

Я створив повний набір налаштувань для вашого проекту Discord AI Bot:

### 1. 📁 Файли конфігурації VS Code

- `.vscode/settings.json` - основні налаштування редактора
- `.vscode/extensions.json` - рекомендовані розширення
- `.vscode/launch.json` - налаштування для дебагу
- `.vscode/tasks.json` - задачі для швидкого запуску

### 2. 🛠️ Інструменти якості коду

- `.prettierrc` - форматування коду
- `.eslintrc.json` - правила якості коду

### 3. 📖 Документація

- `CURSOR_CUSTOM_INSTRUCTIONS.md` - детальні інструкції для AI
- `CURSOR_SETUP_GUIDE.md` - цей гід

## 🎯 Як налаштувати Cursor AI

### Крок 1: Встановлення розширень

1. Відкрийте VS Code
2. Перейдіть до розділу Extensions (Ctrl+Shift+X)
3. Встановіть рекомендовані розширення з `.vscode/extensions.json`

### Крок 2: Налаштування Custom Instructions

1. Відкрийте Cursor AI
2. Перейдіть до налаштувань (Settings)
3. Знайдіть розділ "Custom Instructions"
4. Скопіюйте вміст файлу `CURSOR_CUSTOM_INSTRUCTIONS.md`

### Крок 3: Налаштування середовища

1. Скопіюйте `.env.example` в `.env`
2. Заповніть необхідні змінні середовища
3. Встановіть залежності: `npm install`

## 🔧 Швидкі команди

### Запуск проекту:

- **Ctrl+Shift+P** → "Tasks: Run Task" → "🚀 Запустити Discord Bot"
- Або **F5** для дебагу

### Deploy команд:

- **Ctrl+Shift+P** → "Tasks: Run Task" → "🔧 Deploy Commands"

### Тестування AI:

- **Ctrl+Shift+P** → "Tasks: Run Task" → "🧪 Тестувати AI"

### Docker команди:

- **Ctrl+Shift+P** → "Tasks: Run Task" → "🐳 Docker Compose Up"

## 🎨 Custom Instructions для Cursor AI

### Основні принципи:

```markdown
Ви - експерт з Node.js, Discord.js v14, Google Sheets API та інтеграції з LLM (Ollama/OpenAI).
Допомагайте розробляти Discord-бота з AI-функціоналом для роботи з Google Sheets.
```

### Специфіка проекту:

- **ПРОЕКТ:** Discord AI Assistant Bot з Google Sheets інтеграцією
- **ТЕХНОЛОГІЇ:** Discord.js v14, Google Sheets API, Ollama LLM, Redis кешування, Prometheus метрики
- **АРХІТЕКТУРА:** Модульна система з розділенням відповідальності
- **МОВА:** JavaScript (ES6+), Node.js 18+

### Стиль коду:

- Використовуйте **async/await** замість промісів
- Додавайте **JSDoc коментарі** для функцій
- Використовуйте **try-catch** для обробки помилок
- Логування через **Winston**
- **Валідація вхідних даних**
- **Rate limiting** для API запитів

## 📁 Структура проекту

```
BotDiscordGodzilla/
├── .vscode/                    # Налаштування VS Code
│   ├── settings.json          # Основні налаштування
│   ├── extensions.json        # Розширення
│   ├── launch.json           # Дебаг
│   └── tasks.json            # Задачі
├── commands/                  # Discord команди
├── config/                    # Конфігурація
├── utils/                     # Утиліти
├── metrics/                   # Моніторинг
├── logs/                      # Логи
├── .prettierrc               # Форматування
├── .eslintrc.json            # Якість коду
└── CURSOR_CUSTOM_INSTRUCTIONS.md  # Інструкції для AI
```

## 🚀 Корисні поради

### 1. Використання Cursor AI:

- Використовуйте **@** для посилання на файли
- Додавайте контекст про ваш проект
- Задавайте конкретні питання

### 2. Розробка:

- Завжди тестуйте код перед комітом
- Використовуйте логування для дебагу
- Дотримуйтесь принципів безпеки

### 3. Продуктивність:

- Використовуйте кешування Redis
- Моніторте метрики Prometheus
- Оптимізуйте запити до API

## 🔍 Приклади запитів до Cursor AI

### Створення нової команди:

```
Створи нову Discord slash-команду для аналізу даних з Google Sheets з використанням AI
```

### Оптимізація коду:

```
Проаналізуй цей код та запропонуй покращення для продуктивності та безпеки
```

### Документація:

```
Створи JSDoc документацію для цієї функції
```

### Тестування:

```
Створи тести для цієї функції з використанням моків
```

## 🛡️ Безпека

### Важливі моменти:

- Ніколи не комітьте `.env` файл
- Валідуйте всі вхідні дані
- Використовуйте rate limiting
- Логуйте важливі події

### Змінні середовища:

```env
# Discord Bot
DISCORD_TOKEN=your_discord_token
DISCORD_CLIENT_ID=your_client_id

# Google Sheets
GOOGLE_SERVICE_ACCOUNT_EMAIL=your_email
GOOGLE_PRIVATE_KEY=your_private_key

# AI/LLM
OPENAI_API_KEY=your_openai_key
OLLAMA_HOST=http://localhost:11434
```

## 📞 Підтримка

Якщо у вас виникнуть питання:

1. Перевірте документацію в папці проекту
2. Подивіться логи в папці `logs/`
3. Перевірте метрики Prometheus
4. Зверніться до README.md

---

**🎉 Вітаємо! Ваш проект тепер повністю налаштований для роботи з Cursor AI!**
