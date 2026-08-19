# 🦖 Discord AI Assistant Bot - Godzilla

**Потужний Discord бот з AI функціоналом для автоматизації бізнесу та управління командою**
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

Discord AI Assistant Bot (Godzilla) — це інноваційний корпоративний бот, розроблений для автоматизації роботи з документами, аналізу даних, щоденного менеджменту команд (Daily Standups) та підтримки операційної бізнес-діяльності. Адаптований для потреб компаній, стартапів та організацій будь-якого масштабу.

Бот інтегрується з Google Sheets, Google Drive, Discord та AI сервісами (OpenAI, Ollama) для надання потужних інструментів аналізу та автоматизації у вашому робочому просторі.

## 🚀 Основні можливості

### 🤖 AI Асистент та Комунікація
- Природномовний аналіз бізнес-даних 📊
- Самарі довгих переписок (тредів) зі створенням списку задач (Action Items) 📝
- Контекстна пам'ять для розмов 💬
- Підтримка кількох мов 🌍

### 🔍 Розумний пошук та Документи
- Гнучкий пошук по корпоративних документах 🔎
- Фільтрація за датами, типами, пріоритетами 📅
- Гібридний пошук (FTS + векторний) 🧠
- Робота з Google Sheets та Google Drive ☁️
- Автоматична індексація документів 📇

### 👥 Командна робота (Team Management)
- Організація Daily Standups команди прямо в Discord 🗓️
- Відстеження блокерів (Blockers) та планів на день 🎯
- Автоматична генерація щоденних звітів 📋
- Координація між відділами компанії 🤝

### 📊 Бізнес-аналітика
- Статистика використання ресурсів 📈
- Аналіз даних з таблиць (продажі, KPI) 📊
- Експорт результатів у різних форматах 📤

### 🔒 Безпека
- Максимальний рівень захисту комерційних даних 🛡️
- Контроль доступу через ролі Discord (Permissions) 👮
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

## 🏗️ Архітектура

### Технологічний стек

- **Мова**: TypeScript 5.0+ 💻
- **Платформа**: Node.js 20.x (LTS) ⚡
- **Фреймворк**: Discord.js 14.x 🎮
- **База даних**: SQLite3 (FTS5), Redis (кеш) 🗄️
- **AI/ML**: Ollama (локально), OpenAI API (опційно) 🤖
- **Інтеграції**: Google Sheets API, Google Drive API ☁️

## ⚙️ Встановлення та налаштування

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

4. **Запуск бота** ▶️
   ```bash
   npm run deploy
   npm start
   ```

## 📋 Команди бота

### 👥 Командна робота та Менеджмент
```
/standup trigger      # Запускає опитування для Daily Standup у поточному каналі
/summarize count:50   # Читає 50 останніх повідомлень і робить коротке AI-самарі + Action Items
```

### 🔍 Пошук та аналіз
```
/пошук запит:"маркетинг" тип_документа:"звіти"
/розумний-пошук кількість_вище:100 ціна_нижче:1000
/ai запит:проаналізуй продажі за місяць та дай рекомендації
```

### 📄 Управління документами
```
/файли завантажити файл:report.pdf
/документи відділ_продажів список
```

### 📊 Аналітика
```
/статистика
/аналітика звіт тип:general
```

## 🧪 Тестування

```bash
npm run test              # Запуск всіх тестів 🧪
npm run test:unit         # Модульні тести 🔧
npm run test:integration  # Інтеграційні тести 🔗
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

---

<div align="center">
  <sub>Потужний. Надійний. Український.</sub> 🇺🇦
</div>
