# 🛠️ Гайд розробника
# Посібник розробника Discord AI Assistant Bot

## 🎯 Вступ

Цей посібник призначений для розробників, які хочуть зробити внесок у розвиток Discord AI Assistant Bot або розширити його функціональність.

## 🏗️ Архітектура проекту

### Структура каталогів

```
src/
├── commands/          # Команди бота
├── services/          # Бізнес-логіка
├── core/             # Ядро системи
├── config/           # Конфігурація
├── utils/            # Утиліти
└── tests/            # Тести
```

### Основні компоненти

- **Bot** - головний клас бота
- **CommandManager** - управління командами
- **ServiceContainer** - контейнер сервісів
- **ErrorHandler** - обробка помилок
- **EventManager** - управління подіями

## 🧪 Розробка сервісів

### Створення нового сервісу

Всі сервіси повинні наслідувати клас `BaseService`:

```typescript
import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';

export class MyNewService extends BaseService {
  constructor(config: BotConfig) {
    super('MyNewService', config);
  }

  protected async onInitialize(): Promise<void> {
    // Логіка ініціалізації
  }

  protected async onShutdown(): Promise<void> {
    // Логіка завершення роботи
  }

  protected async onHealthCheck(): Promise<HealthStatus> {
    // Перевірка стану сервісу
    return { healthy: true, service: this.name };
  }
}
```

## 📝 Розробка команд

### Створення нової команди

Команди розміщуються в каталозі [src/commands](file:///c%3A/Users/dmitz/Documents/GitHub/BotDiscordGodzilla/src/commands):

```typescript
import { SlashCommandBuilder } from '@discordjs/builders';
import type { Command } from '@/types';

export const myCommand: Command = {
  data: new SlashCommandBuilder()
    .setName('mycommand')
    .setDescription('Опис команди')
    .addStringOption(option =>
      option.setName('параметр')
        .setDescription('Опис параметра')
        .setRequired(true)),
  
  async execute(interaction) {
    // Логіка виконання команди
    await interaction.reply('Відповідь команди');
  }
};
```

## 🔌 Інтеграція з Google API

### Налаштування доступу

1. Створіть Service Account у Google Cloud Console
2. Надайте необхідні права доступу
3. Завантажте ключ Service Account
4. Налаштуйте змінні середовища

### Приклад використання GoogleService

```typescript
const googleService = new GoogleService(config);
const sheetData = await googleService.getSheetData(spreadsheetId, range);
```

## 🤖 Інтеграція з AI

### Підтримувані провайдери

- **OpenAI** - хмарні моделі OpenAI
- **Ollama** - локальні моделі

### Приклад використання AIService

```typescript
const aiService = new AIService(config);
const response = await aiService.generateResponse('Ваш запит тут');
```

## 🧪 Тестування

### Одиничне тестування

```bash
npm run test:unit
```

### Інтеграційне тестування

```bash
npm run test:integration
```

### E2E тестування

```bash
npm run test:e2e
```

## 🐳 Розробка з Docker

### Локальний запуск

```bash
docker-compose up -d
```

### Перебудова образу

```bash
docker-compose build
```

## 📊 Моніторинг та метрики

### Prometheus метрики

Бот експортує метрики для Prometheus:
- Використання пам'яті
- Час відповіді
- Кількість запитів

### Логування

Використовується бібліотека Winston для структурованого логування.

## 🔧 Налаштування середовища розробки

### Необхідні інструменти

- Node.js 18+
- npm або yarn
- Docker (опційно)
- Google Cloud обліковий запис
- Discord Developer обліковий запис

### Встановлення залежностей

```bash
npm install
```

### Запуск у режимі розробки

```bash
npm run dev
```

## 📤 Внесок у проект

### Git workflow

1. Створіть fork репозиторію
2. Створіть feature branch
3. Внесіть зміни
4. Напишіть тести
5. Створіть Pull Request

### Стиль коду

- Використовується TypeScript
- ESLint для перевірки коду
- Prettier для форматування

## 📞 Підтримка

Якщо у вас виникли питання щодо розробки:

1. Перевірте документацію
2. Створіть issue у репозиторії GitHub
3. Зверніться до спільноти розробників

© 2025 Dmitry Shivachov (Dmitze). Всі права захищені.