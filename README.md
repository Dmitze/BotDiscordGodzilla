# 🦖 Discord AI Assistant Bot - Godzilla

**Потужний Discord бот з AI функціоналом для Збройних Сил України**
**Оновлено: 13.09.2025** 📅

<div align="center">

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)
[![Discord.js](https://img.shields.io/badge/Discord.js-v14-blue?style=for-the-badge)](https://discord.js.org/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0-blue?style=for-the-badge)](https://www.typescriptlang.org/)
[![Node.js](https://img.shields.io/badge/Node.js-18%2B-green?style=for-the-badge)](https://nodejs.org/)
[![AI](https://img.shields.io/badge/AI-Ollama%20%26%20OpenAI-orange?style=for-the-badge)](https://ollama.ai/)
[![Database](https://img.shields.io/badge/Database-SQLite3%20%26%20Redis-blue?style=for-the-badge)](https://www.sqlite.org/)
[![Docker](https://img.shields.io/badge/Docker-Containerization-blue?style=for-the-badge)](https://www.docker.com/)
[![Testing](https://img.shields.io/badge/Testing-Jest%20%26%20Supertest-green?style=for-the-badge)](https://jestjs.io/)

</div>

## 🎯 Призначення

Discord AI Assistant Bot (Godzilla) - це інноваційний бот, розроблений для автоматизації роботи з документами, аналізу даних та підтримки операційної діяльності. Спеціально адаптований для потреб ЗСУ та критично важливих організацій.

Бот інтегрується з Google Sheets, Google Drive, Discord та AI сервісами (OpenAI, Ollama) для надання потужних інструментів аналізу та автоматизації.

## 🚀 Основні можливості

### 🤖 AI асистент
- Природномовний аналіз даних 📊
- Генерація звітів та рекомендацій 📝
- Контекстна пам'ять для розмов 💬
- Підтримка кількох мов 🌍

### 🔍 Розумний пошук
- Гнучкий пошук по всіх документах 🔎
- Фільтрація за датами, типами, пріоритетами 📅
- Пагінація результатів 📄
- Гібридний пошук (FTS + векторний) 🧠

### 📄 Управління документами
- Робота з Google Sheets та Google Drive ☁️
- Читання різних форматів файлів (PDF, DOCX, TXT) 📎
- AI-аналіз вмісту файлів 🤖
- Автоматична індексація документів 📇

### 📊 Аналітика
- Статистика використання бота 📈
- Аналіз даних з таблиць 📊
- Експорт результатів у різних форматах 📤
- Візуалізація даних 📉

### ⚡ Операції
- Управління військовими операціями 🎯
- Координація між підрозділами 🤝
- Розвідувальні дані 🕵️
- Моніторинг виконання завдань 📋

### 🔒 Безпека
- Максимальний рівень захисту даних 🛡️
- Контроль доступу через ролі Discord 👮
- Rate limiting та валідація даних 🚦
- Аудит логування всіх дій 📜

## 📚 Документація

Повна документація доступна у папці [docs/](docs/):

### 🌐 Мови документації
- [🇺🇦 Українська](docs/uk/README.md) - повна документація українською
- [🇺🇸 English](docs/en/README.md) - complete documentation in English

### 📖 Основні розділи
- [Швидкий старт](docs/guides/QUICK_START.md) ⚡
- [Налаштування](docs/guides/SETUP.md) ⚙️
- [Гід користувача](docs/guides/USAGE_GUIDE.md) 📖
- [Архітектура](docs/ARCHITECTURE.md) 🏗️
- [Безпека](docs/security/SECURITY_GUIDE.md) 🔐
- [API документація](docs/api/API_DOCUMENTATION.md) 🔌

### 🔍 Пошук по документації
Доступний інтерактивний пошук: [search.html](docs/search.html) 🔍

## 🏗️ Архітектура

### Основні компоненти

1. **Ядро бота** (`src/core/Bot.ts`) 🤖
   - Головний клас бота
   - Ініціалізація та управління сервісами
   - Обробка подій Discord

2. **Команди** (`src/commands/`) 💬
   - Модульна система команд
   - Валідація вхідних даних
   - Структурне логування

3. **Сервіси** (`src/services/`) ⚙️
   - GoogleService - робота з Google API
   - AIService - інтеграція з AI моделями
   - CacheService - кешування даних
   - MetricsService - збір метрик

4. **Пошук** (`src/search/`) 🔍
   - Гібридний пошук (FTS + векторний)
   - Індексація документів
   - RAG пайплайн

5. **RAG** (`src/rag/`) 🧠
   - Пошук релевантних фрагментів
   - Підготовка контексту
   - Генерація відповідей

### Технологічний стек

- **Мова**: TypeScript 5.0+ 💻
- **Платформа**: Node.js 20.x (LTS) ⚡
- **Фреймворк**: Discord.js 14.x 🎮
- **База даних**: SQLite3 (FTS5), Redis (кеш) 🗄️
- **AI/ML**: Ollama (локально), OpenAI API (опційно) 🤖
- **Інтеграції**: Google Sheets API, Google Drive API ☁️
- **Моніторинг**: Prometheus + Grafana 📊
- **Тестування**: Jest, Supertest 🧪
- **Контейнеризація**: Docker 🐳
- **Розгортання**: PM2, Docker Compose 🚀

## ⚙️ Встановлення та налаштування

### Системні вимоги

- Node.js 18+ 🟩
- Git 🐱
- Docker (опційно) 🐳
- Google Cloud обліковий запис 🌐
- Discord Developer обліковий запис 🎮

### Швидкий старт

1. **Клонування проекту** 📦
   ```bash
   git clone https://github.com/Dmitze/BotDiscordGodzilla.git
   cd BotDiscordGodzilla
   ```

2. **Встановлення залежностей** ⬇️
   ```bash
   npm install
   ```

3. **Налаштування змінних середовища** ⚙️
   ```bash
   cp .env.example .env
   # Відредагуйте .env файл з вашими налаштуваннями
   ```

4. **Налаштування Discord бота** 🎮
   - Перейдіть на Discord Developer Portal
   - Створіть новий додаток та бота
   - Скопіюйте токен
   - Увімкніть необхідні Intents

5. **Налаштування Google API** ☁️
   - Створіть проект у Google Cloud Console
   - Увімкніть Google Sheets API та Google Drive API
   - Створіть Service Account та завантажте JSON ключ

6. **Запуск бота** ▶️
   ```bash
   npm run deploy
   npm start
   ```

## 📋 Команди бота

### 🔍 Пошук та аналіз

```
/пошук запит:"особовий склад" тип_документа:"накази"
/розумний-пошук кількість_вище:100 ціна_нижче:1000
/ai запит:проаналізуй залишки та дай рекомендації
```

### 📄 Управління документами

```
/документи особовий-склад список
/документи техніка додати назва:"Танк Т-72"
/файли завантажити файл:document.pdf
```

### 📊 Аналітика

```
/статистика
/аналітика звіт тип:general
/продуктивність моніторинг
```

### ⚡ Операції

```
/операції ситуація поточний_стан
/операції завдання створити опис:"Патрулювання"
/операції координація зв'язок_з_штабом
```

## 🧪 Тестування

```bash
npm run test              # Запуск всіх тестів 🧪
npm run test:unit         # Модульні тести 🔧
npm run test:integration  # Інтеграційні тести 🔗
npm run test:coverage     # Звіт про покриття 📊
```

## 📦 Розгортання

### Docker 🐳

```bash
docker-compose up -d
```

### PM2 ⚡

```bash
npm run build
pm2 start dist/index.js
```

## 🤝 Внесок у розвиток

1. Форкніть репозиторій 🍴
2. Створіть feature branch 🌿
3. Зробіть коміт змін 💾
4. Відправте Pull Request 📤

## 📃 Ліцензія

Цей проект ліцензовано за ліцензією MIT - див. файл [LICENSE](LICENSE) для деталей.

## 📞 Контакти

- Автор: Дмитро Шивачов (Dmitze) 👨‍💻
- Email: dmitzeshivachov@outlook.com 📧
- GitHub: [@Dmitze](https://github.com/Dmitze) 🐱
- Discord: dmitryshivachov3756 🎮
- Telegram: [@DmitryShiva](https://t.me/DmitryShiva) 📱

---

<div align="center">
  <sub>Потужний. Надійний. Український.</sub> 🇺🇦
</div>