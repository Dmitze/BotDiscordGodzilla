# 🦖 **DISCORD AI ASSISTANT BOT - GODZILLA**

**Потужний Discord бот з AI функціоналом для Збройних Сил України**

[![Version](https://img.shields.io/badge/version-3.0.0-blue.svg)](https://github.com/Dmitze/BotDiscordGodzilla)
[![License](https://img.shields.io/badge/license-Godzilla%20Bot%20License%20v3.0-green.svg)](../../LICENSE.md)
[![Node.js](https://img.shields.io/badge/node.js-18+-yellow.svg)](https://nodejs.org/)
[![TypeScript](https://img.shields.io/badge/typescript-5.3+-blue.svg)](https://www.typescriptlang.org/)

---

## 🚀 Key Features

### 🤖 AI Assistant
- Natural language document search and analysis
- Context-aware responses with conversation history
- Multi-model support (OpenAI, Ollama)
- RAG-enhanced document retrieval

### 📄 Document Management
- Google Drive integration with real-time indexing
- Advanced search with hybrid vector/text retrieval
- Document analysis and summarization
- Compliance and audit tracking

### 🎨 Enhanced Discord Experience
- Rich markdown rendering with syntax highlighting
- Interactive buttons and menus for document navigation
- Thread-based conversations with context preservation
- Multi-language support (primarily Ukrainian)

### 🔧 Automation & Integration
- n8n workflow automation for document processing
- Scheduled tasks and notifications
- External API integrations (Jira, Trello, Notion)
- Real-time document change monitoring

---

## 📚 **ДОКУМЕНТАЦІЯ**

- **[🔗 Індекс документації](../INDEX.md)** — огляд і швидкі переходи
- **[🏗️ Архітектура](../ARCHITECTURE.md)** — карта системи, технології, потоки
- **[🛡️ Безпека](../security/SECURITY_GUIDE.md)** — політика і технічні захисти
- **[🧩 API та команди](../API_OVERVIEW.md)** — огляд публічних інтерфейсів
- **[🧭 Гайд розробника](../DEVELOPER_GUIDE.md)** — як працювати з кодом
- **[🧠 RAG та пошук](../guides/rag.md)** — гібридний пошук, embeddings, ENV

> Примітка: детальні README з підпапок коду переїхали в `docs/...` з тію самою структурою (див. нижче «README по папках коду»).


## 🎯 **ОСНОВНІ КОМАНДИ**

### **🔍 Пошук та аналіз**
```bash
/пошук запит:"особовий склад" тип_документа:"накази"
/розумний-пошук кількість_вище:100 ціна_нижче:1000
/ai запит:проаналізуй залишки та дай рекомендації
```

### **📄 Управління документами**
```bash
/документи особовий-склад список
/документи техніка додати назва:"Танк Т-72"
/файли завантажити файл:document.pdf
```

### **📊 Аналітика**
```bash
/статистика
/аналітика звіт тип:general
/продуктивність моніторинг
```

### **⚡ Операції**
```bash
/операції ситуація поточний_стан
/операції завдання створити опис:"Патрулювання"
/операції координація зв'язок_з_штабом
```

### **📋 Аналіз документів**
```bash
/analyze-doc file:"Наказ №123" type:"full"
/analyze-doc file:"Звіт про постачання" type:"structure"
/analyze-doc file:"План операції" type:"summary"
/analyze-doc file:"Договір" type:"actions"
/analyze-doc file:"Фінансовий звіт" type:"compliance"
/analyze-doc file:"Протокол" type:"quality"
```

---

## 🏗️ **АРХІТЕКТУРА**

### **📁 Структура проекту**

```
src/
├── commands/          # Команди бота
├── services/          # Бізнес-логіка
├── core/             # Ядро системи
├── config/           # Конфігурація
├── utils/            # Утиліти
└── tests/            # Тести
```

### **🔧 Основні компоненти**
- **Bot** - головний клас бота
- **CommandManager** - управління командами
- **ServiceContainer** - контейнер сервісів
- **ErrorHandler** - обробка помилок
- **EventManager** - управління подіями

### **🛠️ Сервіси**
- **AIService** - AI функціональність
- **GoogleService** - робота з Google API
- **CacheService** - кешування
- **MetricsService** - метрики
- **SchedulerService** - планувальник
- **DocumentAnalysisService** - аналіз документів

### **🧠 Функції аналізу документів**
- **Структурний аналіз** - визначення розділів, заголовків, типу документа
- **Підсумування** - короткий, середній та детальний опис
- **Витяг дій** - завдання, відповідальні особи, терміни
- **Генерація питань** - фактичні, аналітичні та оцінювальні запитання
- **Перевірка відповідності** - дотримання стандартів та політик
- **Переклад** - переклад документів українською мовою
- **Оцінка якості** - читабельність, структура, граматика
- **Аналіз даних** - витяг числових даних та статистики
- **Огляд безпеки** - визначення конфіденційної інформації
- **Аналіз змін** - порівняння версій документів
- **Прогнозування ефективності** - оцінка сприйняття аудиторією
- **Аналіз зацікавлених сторін** - визначення учасників та їх ролей
- **Аналіз бюджету** - витрати, джерела фінансування
- **Оцінка ризиків** - потенційні загрози та заходи з протидії
- **Сегментація аудиторії** - визначення цільових груп
- **Хронологічний аналіз** - витяг таймлайнів та подій
- **Кластеризація ключових слів** - групування термінів та концепцій
- **Візуалізація даних** - рекомендації щодо графіків та діаграм
- **Перевірка цитування** - верифікація джерел та посилань
- **Аналіз мовного стилю** - оцінка тону та формальності
- **Аналіз прогалин** - визначення відсутньої інформації
- **Переосмислення вмісту** - ідеї для репурпузингу контенту
- **Оцінка доступності** - перевірка відповідності стандартам
- **Оптимізація процесів** - вдосконалення описаних процедур
- **Персоналізація контенту** - адаптація під різні аудиторії
- **Бенчмаркинг** - встановлення показників ефективності

---

## 🚀 **ЗАПУСК**

### **⚡ Швидкий запуск**

```bash
# Клонування репозиторію
git clone https://github.com/Dmitze/BotDiscordGodzilla.git
cd BotDiscordGodzilla

# Встановлення залежностей
npm install

# Налаштування змінних середовища
cp .env.example .env
# Відредагуйте .env файл

# Запуск
npm run dev
```

### **🐳 Docker запуск**

```
# Збірка образу
docker build -t godzilla-bot .

# Запуск контейнера
docker run -d --name godzilla-bot godzilla-bot
```

---

## 🔧 **НАЛАШТУВАННЯ**

### **📋 Обов'язкові змінні середовища**

```
# Discord
DISCORD_TOKEN=your_discord_token
DISCORD_CLIENT_ID=your_client_id
DISCORD_GUILD_ID=your_guild_id

# Google
GOOGLE_API_KEY=your_google_api_key
GOOGLE_APP_SCRIPT_URL=your_app_script_url

# AI
# Провайдер AI: ollama (локальна LLM)
OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=llama3

# Пошук/RAG (уривки, індекси, ембеддінги)
SEARCH_INDEX_PATH=.data/search/index.sqlite
SEARCH_FTS_TOKENIZER=unicode61
RETRIEVER_K=8
RETRIEVER_ALPHA=0.7
EMBEDDINGS_ENABLE=true
EMBEDDINGS_PROVIDER=local
EMBEDDINGS_MODEL=nomic-embed-text
RAG_MAX_CONTEXT_TOKENS=3000
AI_MAX_TOKENS=1024

# OpenAI (опційно)
OPENAI_API_KEY=your_openai_api_key
```

### **🎯 Рекомендовані налаштування**
- **Node.js** 18+ 
- **RAM** 2GB+
- **Discord** Developer Portal налаштування
- **Google Cloud** проект

---

## 📊 **МЕТРИКИ ТА МОНІТОРИНГ**

### **📈 Ключові показники**
- **Час відповіді:** < 1.5 секунди
- **Доступність:** 99.9%
- **Покриття тестами:** 95%+
- **Використання пам'яті:** < 500MB

### **🔍 Моніторинг**
- **Prometheus** метрики
- **Winston** логування
- **Health checks** кожні 30 секунд
- **Memory monitoring** кожну хвилину

---