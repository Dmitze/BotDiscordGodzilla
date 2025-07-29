# 🚀 Гід по оновленню проекту - Версія 2.2.0

## 📋 Що нового в версії 2.2.0

### ✨ **Архітектурні покращення:**
- 🏗️ **Модульна архітектура** - код розділено на логічні модулі
- ⚙️ **Централізована конфігурація** - всі налаштування в одному місці
- 🔄 **Система повторних спроб** - автоматичні повторні спроби при помилках
- 📊 **Система метрик Prometheus** - детальний моніторинг роботи бота

### 🔍 **Покращений пошук:**
- 📈 **Пошук за діапазонами** - ціна від/до, кількість від/до
- 🔄 **Сортування результатів** - за будь-яким полем
- 🎯 **Розширені фільтри** - комбінований пошук
- 📋 **Покращена пагінація** - кнопки "Перша/Остання сторінка"

### 📤 **Покращений експорт:**
- 📄 **Метадані в файлах** - інформація про експорт
- 📊 **Підтримка CSV** - експорт в різних форматах
- 📋 **Звіти аналізу** - експорт результатів AI-аналізу
- 🧹 **Автоочищення** - видалення старих файлів

### 🛠️ **Утиліти та форматування:**
- 📝 **Форматування даних** - красиве відображення чисел, дат, валют
- 🔧 **Утиліти повторних спроб** - надійна робота з API
- 📊 **Статистика експорту** - відстеження використання

## 🚀 Кроки оновлення

### 1. **Оновлення залежностей**

```bash
# Встановлюємо нові залежності
npm install express prom-client

# Перевіряємо встановлення
npm list
```

### 2. **Оновлення змінних середовища**

Додайте нові змінні до `.env`:

```env
# Metrics Configuration
METRICS_ENABLED=true
METRICS_PORT=3000
METRICS_PATH=/metrics

# Performance Configuration
REQUEST_TIMEOUT=30000
MAX_RETRIES=3

# Export Configuration
MAX_FILE_SIZE=26214400
TEMP_FILE_TTL=60000
INCLUDE_METADATA=true

# Search Configuration
SEARCH_MAX_RESULTS=100
FUZZY_MATCH=true
CASE_SENSITIVE=false
ENABLE_STEMMING=false

# OpenAI Configuration (опціонально)
OPENAI_MODEL=gpt-3.5-turbo
OPENAI_MAX_TOKENS=800
OPENAI_TEMPERATURE=0.3

# Security Configuration (опціонально)
RATE_LIMIT_WINDOW=900000
RATE_LIMIT_MAX=100
ALLOWED_ROLES=
ADMIN_ROLES=
```

### 3. **Оновлення команд**

```bash
# Оновлюємо slash-команди
node deploy-commands.js
```

### 4. **Тестування нових функцій**

```bash
# Тестуємо AI-функції
node test-ai.js

# Перевіряємо метрики (якщо увімкнено)
curl http://localhost:3000/metrics
curl http://localhost:3000/health
```

## 🆕 Нові команди

### **Покращений пошук:**
```
/розумний-пошук-покращений номенклатура:стол ціна_від:100 ціна_до:1000 сортування:ціна порядок:desc
```

### **Нові параметри пошуку:**
- `ціна_від` - мінімальна ціна
- `ціна_до` - максимальна ціна  
- `кількість_від` - мінімальна кількість
- `кількість_до` - максимальна кількість
- `сортування` - поле для сортування
- `порядок` - asc/desc

## 📊 Система метрик

### **Доступні метрики:**
- `discord_bot_commands_total` - кількість команд
- `discord_bot_command_duration_seconds` - час виконання
- `discord_bot_api_requests_total` - API запити
- `discord_bot_cache_hits_total` - попадання в кеш
- `discord_bot_errors_total` - помилки
- `discord_bot_memory_usage_bytes` - використання пам'яті

### **Перегляд метрик:**
```bash
# Prometheus формат
curl http://localhost:3000/metrics

# Health check
curl http://localhost:3000/health
```

## 🔧 Нові файли та структура

### **Нові папки:**
```
utils/           # Утиліти
├── retry.js     # Повторні спроби
├── formatters.js # Форматування
└── exportHelpers.js # Експорт

config/          # Конфігурація
└── config.js    # Центральна конфігурація

metrics/         # Метрики
└── prometheus.js # Prometheus метрики

commands/        # Команди (майбутнє)
└── enhancedSearch.js # Покращений пошук
```

### **Оновлені файли:**
- `index.js` - інтеграція нових модулів
- `package.json` - нові залежності
- `env.example` - нові змінні середовища

## 🧪 Тестування

### **Тестування пошуку:**
```bash
# Запускаємо бота
node index.js

# Тестуємо нові команди в Discord
/розумний-пошук-покращений номенклатура:стол ціна_від:100
/статистика
/статус
```

### **Тестування експорту:**
```bash
# Експорт з метаданими
/пошук поле:назва запит:стол
# Натискаємо кнопку "📊 Експорт Excel"
```

### **Тестування метрик:**
```bash
# Перевіряємо доступність метрик
curl http://localhost:3000/health
curl http://localhost:3000/metrics | grep discord_bot
```

## 🔧 Налаштування

### **Вимкнення метрик:**
```env
METRICS_ENABLED=false
```

### **Налаштування експорту:**
```env
MAX_FILE_SIZE=52428800  # 50MB
TEMP_FILE_TTL=300000    # 5 хвилин
INCLUDE_METADATA=false  # Без метаданих
```

### **Налаштування пошуку:**
```env
SEARCH_MAX_RESULTS=200  # Більше результатів
FUZZY_MATCH=false       # Точний пошук
CASE_SENSITIVE=true     # Чутливий до регістру
```

## 🐛 Вирішення проблем

### **Проблема: "Модуль не знайдено"**
```bash
# Перевіряємо залежності
npm install
npm list

# Перевіряємо шляхи
node -e "console.log(require.resolve('./config/config.js'))"
```

### **Проблема: "Метрики не працюють"**
```bash
# Перевіряємо порт
netstat -an | grep 3000

# Перевіряємо конфігурацію
echo $METRICS_ENABLED
```

### **Проблема: "Експорт не працює"**
```bash
# Перевіряємо права доступу
ls -la tmp/

# Перевіряємо розмір файлу
du -sh tmp/*
```

## 📈 Моніторинг

### **Grafana Dashboard (опціонально):**
```yaml
# docker-compose.yml
services:
  grafana:
    image: grafana/grafana:latest
    ports:
      - "3001:3000"
    environment:
      - GF_SECURITY_ADMIN_PASSWORD=admin
    volumes:
      - grafana_data:/var/lib/grafana

  prometheus:
    image: prom/prometheus:latest
    ports:
      - "9090:9090"
    volumes:
      - ./prometheus.yml:/etc/prometheus/prometheus.yml
      - prometheus_data:/prometheus
```

### **Prometheus конфігурація:**
```yaml
# prometheus.yml
global:
  scrape_interval: 15s

scrape_configs:
  - job_name: 'discord-bot'
    static_configs:
      - targets: ['localhost:3000']
    metrics_path: '/metrics'
```

## 🔄 Відкат змін

### **Якщо щось не працює:**
```bash
# Відключаємо метрики
export METRICS_ENABLED=false

# Використовуємо старий пошук
# Команди /пошук та /розумний-пошук залишаються працювати

# Відновлюємо стару конфігурацію
cp env.example.backup .env
```

## 📚 Документація

### **Корисні посилання:**
- [Prometheus метрики](https://prometheus.io/docs/concepts/metric_types/)
- [Express.js документація](https://expressjs.com/)
- [Discord.js документація](https://discord.js.org/)

### **Приклади використання:**
- [AI_EXAMPLES.md](AI_EXAMPLES.md) - приклади AI-функцій
- [COMMANDS_REFERENCE.md](COMMANDS_REFERENCE.md) - всі команди
- [LAUNCH_INSTRUCTIONS.md](LAUNCH_INSTRUCTIONS.md) - запуск

---

**🎉 Вітаємо з успішним оновленням! Ваш бот тепер має професійну архітектуру та розширений функціонал!** 