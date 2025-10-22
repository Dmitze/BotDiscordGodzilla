# 📁 Вихідний код (src)

Цей каталог містить вихідний код бота, розділений за логічними модулями.

## 🏗️ Основні компоненти

### 📂 /commands
- **BaseCommand.ts** - базовий клас для всіх команд
- **SearchCommand.ts** - пошук та фільтрація даних
- **AIAssistantCommand.ts** - AI асистент
- **DocumentsCommand.ts** - управління документами
- **FileManagerCommand.ts** - робота з файлами
- **OperationsCommand.ts** - операційні процеси
- **AnalyticsCommand.ts** - аналітика та звіти
- **PerformanceCommand.ts** - моніторинг продуктивності
- **EnhancedSearchCommand.ts** - розширений пошук
- **statistics.ts** - статистика використання

### 📂 /services
- **AIService.ts** - AI функціональність (OpenAI, Ollama)
- **GoogleService.ts** - робота з Google API
- **CacheService.ts** - кешування даних
- **MetricsService.ts** - метрики та моніторинг
- **SchedulerService.ts** - планувальник завдань

### 📂 /core
- **Bot.ts** - головний клас бота
- **CommandManager.ts** - управління командами
- **ServiceContainer.ts** - контейнер сервісів
- **ErrorHandler.ts** - централізована обробка помилок

### 📂 /utils
- **logger.ts** - логування подій
- **formatters.ts** - форматування даних
- **security.ts** - функції безпеки
- **pagination.ts** - робота з пагінацією

### 📂 /config
- **index.ts** - конфігурація додатку
- **constants.ts** - константи
- **enums.ts** - перерахування

## 🔗 Пов'язана документація
- [Архітектура](../ARCHITECTURE.md)
- [Гайд розробника](../DEVELOPER_GUIDE.md)
- [API документація](../API_OVERVIEW.md)
