# 🛡️ Гід по безпеці Discord AI Bot

## 📋 Огляд безпеки

Цей документ описує систему безпеки Discord AI Bot, яка забезпечує захист даних, контроль доступу та запобігання зловживанням.

## 🔐 Рівні безпеки

### 1. **Контроль доступу в Discord**

#### Ролі користувачів:
- **Адміністратор** - повний доступ до всіх функцій
- **Бот-Користувач** - базовий доступ до бота
- **Sheets-Доступ** - доступ до Google Sheets функцій
- **AI-Доступ** - доступ до AI-функцій
- **Експорт-Доступ** - доступ до експорту даних

#### Налаштування ролей:
```javascript
// Приклад перевірки ролі
const hasAccess = await checkPermission(
  interaction, 
  [ROLES.AI_ACCESS, ROLES.ADMIN], 
  'AI Assistant'
);
```

### 2. **Rate Limiting**

#### Ліміти запитів:
- **Пошук:** 10 запитів за хвилину
- **AI-аналіз:** 5 запитів за 2 хвилини
- **Експорт:** 3 запити за 5 хвилин
- **Загальні команди:** 20 запитів за хвилину

#### Налаштування:
```javascript
const RATE_LIMITS = {
  SEARCH: { max: 10, window: 60 },
  AI_ANALYSIS: { max: 5, window: 120 },
  EXPORT: { max: 3, window: 300 },
  GENERAL: { max: 20, window: 60 }
};
```

### 3. **Валідація вхідних даних**

#### Типи валідації:
- **search** - пошукові запити (макс. 200 символів)
- **ai_prompt** - AI запити (макс. 1000 символів)
- **filename** - назви файлів (макс. 100 символів)
- **general** - загальний текст (макс. 500 символів)

#### Приклад валідації:
```javascript
const validationSchema = {
  query: {
    required: true,
    type: 'string',
    maxLength: 1000,
    sanitize: 'ai_prompt'
  }
};
```

## 🔒 Захист даних

### 1. **Змінні середовища**

#### Обов'язкові змінні:
```env
# Discord
DISCORD_TOKEN=your_discord_bot_token
DISCORD_CLIENT_ID=your_client_id

# Google Services
GOOGLE_APPLICATION_CREDENTIALS=path/to/service-account.json
GOOGLE_SPREADSHEET_ID=your_spreadsheet_id

# AI
OPENAI_API_KEY=your_openai_key
AI_PROVIDER=openai

# Redis
REDIS_HOST=localhost
REDIS_PORT=6379
```

#### Рекомендації:
- Ніколи не комітьте `.env` файл
- Використовуйте різні ключі для різних середовищ
- Регулярно ротуйте API ключі

### 2. **Google API безпека**

#### Service Account:
- Створіть окремий Service Account для бота
- Надайте мінімально необхідні права
- Обмежте доступ до конкретних папок/файлів

#### Налаштування прав:
```json
{
  "type": "service_account",
  "project_id": "your-project",
  "private_key_id": "key-id",
  "private_key": "-----BEGIN PRIVATE KEY-----\n...",
  "client_email": "bot@your-project.iam.gserviceaccount.com",
  "client_id": "client-id",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token"
}
```

### 3. **Шифрування та кешування**

#### Redis безпека:
- Використовуйте пароль для Redis
- Обмежте доступ до Redis порту
- Регулярно очищайте кеш

```javascript
// Приклад підключення до Redis
const redis = require('redis');
const client = redis.createClient({
  url: `redis://:${process.env.REDIS_PASSWORD}@${process.env.REDIS_HOST}:${process.env.REDIS_PORT}`,
  database: process.env.REDIS_DB
});
```

## 🚨 Логування безпеки

### 1. **Типи подій**

#### Безпека:
- `access_denied` - відмовлено в доступі
- `rate_limit_exceeded` - перевищено ліміт запитів
- `invalid_input` - невалідні вхідні дані
- `security_violation` - порушення безпеки

#### Використання:
- `command_executed` - команда виконана
- `ai_query_processed` - AI запит оброблено
- `file_accessed` - доступ до файлу

### 2. **Формат логів**

```javascript
logSecurityEvent('access_denied', {
  userId: interaction.user.id,
  userTag: interaction.user.tag,
  command: 'ai',
  reason: 'insufficient_permissions',
  timestamp: new Date().toISOString()
});
```

## 🔧 Налаштування безпеки

### 1. **Створення ролей Discord**

#### Кроки:
1. Відкрийте налаштування сервера
2. Перейдіть до розділу "Ролі"
3. Створіть ролі:
   - `Адміністратор` (червоний колір)
   - `Бот-Користувач` (синій колір)
   - `Sheets-Доступ` (зелений колір)
   - `AI-Доступ` (фіолетовий колір)
   - `Експорт-Доступ` (жовтий колір)

#### Права ролей:
- **Адміністратор:** Всі права
- **Бот-Користувач:** Базові команди
- **Sheets-Доступ:** Робота з Google Sheets
- **AI-Доступ:** AI-функції
- **Експорт-Доступ:** Експорт даних

### 2. **Налаштування Google API**

#### Створення Service Account:
1. Перейдіть до [Google Cloud Console](https://console.cloud.google.com/)
2. Створіть новий проект або виберіть існуючий
3. Увімкніть Google Sheets API та Google Drive API
4. Створіть Service Account
5. Завантажте JSON ключ
6. Надайте доступ до конкретних файлів/папок

#### Налаштування прав доступу:
```javascript
// Приклад надання доступу до файлу
// В Google Drive: Правий клік на файлі → Надати доступ → Додати email Service Account
```

### 3. **Налаштування Redis**

#### Встановлення Redis:
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install redis-server

# Windows (WSL або Docker)
docker run -d -p 6379:6379 redis:alpine
```

#### Налаштування пароля:
```bash
# Редагування redis.conf
sudo nano /etc/redis/redis.conf

# Додати пароль
requirepass your_strong_password

# Перезапуск Redis
sudo systemctl restart redis
```

## 📊 Моніторинг безпеки

### 1. **Метрики безпеки**

#### Prometheus метрики:
```javascript
// Приклад метрик
const securityMetrics = {
  access_denied_total: new prometheus.Counter({
    name: 'bot_access_denied_total',
    help: 'Total number of access denials'
  }),
  rate_limit_exceeded_total: new prometheus.Counter({
    name: 'bot_rate_limit_exceeded_total',
    help: 'Total number of rate limit violations'
  })
};
```

### 2. **Алерти**

#### Налаштування алертів:
- Більше 10 відмов в доступі за хвилину
- Перевищення rate limit більше 5 разів за хвилину
- Невалідні запити більше 20 за хвилину

## 🚨 Реагування на інциденти

### 1. **Типи інцидентів**

#### Критичні:
- Неавторизований доступ до адміністративних функцій
- Масові спроби обходу rate limiting
- Підозріла активність з AI-функціями

#### Середні:
- Перевищення лімітів запитів
- Невалідні вхідні дані
- Помилки автентифікації

### 2. **План реагування**

#### Кроки:
1. **Виявлення** - моніторинг логів та метрик
2. **Оцінка** - визначення рівня загрози
3. **Реагування** - блокування користувача/функції
4. **Аналіз** - дослідження причин
5. **Відновлення** - повернення функціональності

#### Команди для адміністраторів:
```javascript
// Блокування користувача
await blockUser(userId, reason, duration);

// Вимкнення функції
await disableFeature('ai_analysis', reason);

// Очищення кешу
await clearUserCache(userId);
```

## 🔄 Оновлення безпеки

### 1. **Регулярні перевірки**

#### Щотижнево:
- Перегляд логів безпеки
- Аналіз метрик
- Оновлення залежностей

#### Щомісячно:
- Аудит прав доступу
- Перевірка API ключів
- Тестування системи безпеки

### 2. **Оновлення залежностей**

```bash
# Перевірка вразливостей
npm audit

# Оновлення пакетів
npm update

# Оновлення до нових версій
npm outdated
npm install package@latest
```

## 📞 Підтримка безпеки

### Контакти:
- **Адміністратор:** @admin
- **Технічна підтримка:** support@example.com
- **Екстрені випадки:** emergency@example.com

### Ресурси:
- [Discord Developer Portal](https://discord.com/developers/docs)
- [Google Cloud Security](https://cloud.google.com/security)
- [Redis Security](https://redis.io/topics/security)

---

**⚠️ Важливо:** Цей гід повинен регулярно оновлюватися відповідно до нових загроз та змін у системі безпеки. 