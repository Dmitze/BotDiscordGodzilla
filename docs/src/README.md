# 🏗️ **SRC - ОСНОВНА АРХІТЕКТУРА**

## 📁 **Структура папки src/**

Ця папка містить всю основну логіку Discord AI Assistant Bot. Архітектура побудована на принципах модульності, сервісної архітектури та чистої архітектури.

---

## 🎯 **ОСНОВНІ КОМПОНЕНТИ**

### **🤖 commands/** - Команди бота
Містить всі команди Discord бота, організовані за функціональністю:
- **[BaseCommand.ts](commands/BaseCommand.ts)** - базовий клас для всіх команд
- **[SearchCommand.ts](commands/SearchCommand.ts)** - пошук та фільтрація даних
- **[AIAssistantCommand.ts](commands/AIAssistantCommand.ts)** - AI асистент
- **[DocumentsCommand.ts](commands/DocumentsCommand.ts)** - управління документами
- **[FileManagerCommand.ts](commands/FileManagerCommand.ts)** - робота з файлами
- **[OperationsCommand.ts](commands/OperationsCommand.ts)** - військові операції
- **[AnalyticsCommand.ts](commands/AnalyticsCommand.ts)** - аналітика та звіти
- **[PerformanceCommand.ts](commands/PerformanceCommand.ts)** - моніторинг продуктивності
- **[EnhancedSearchCommand.ts](commands/EnhancedSearchCommand.ts)** - розширений пошук
- **[statistics.ts](commands/statistics.ts)** - статистика використання

### **🛠️ services/** - Бізнес-логіка
Сервіси, що відповідають за основну функціональність:
- **[AIService.ts](services/AIService.ts)** - AI функціональність (OpenAI, Ollama)
- **[GoogleService.ts](services/GoogleService.ts)** - робота з Google API
- **[CacheService.ts](services/CacheService.ts)** - кешування даних
- **[MetricsService.ts](services/MetricsService.ts)** - метрики та моніторинг
- **[SchedulerService.ts](services/SchedulerService.ts)** - планувальник завдань

### **⚙️ core/** - Ядро системи
Основні компоненти архітектури:
- **[Bot.ts](core/Bot.ts)** - головний клас бота
- **[CommandManager.ts](core/CommandManager.ts)** - управління командами
- **[ServiceContainer.ts](core/ServiceContainer.ts)** - контейнер сервісів
- **[ErrorHandler.ts](core/ErrorHandler.ts)** - централізована обробка помилок
- **[EventManager.ts](core/EventManager.ts)** - управління подіями
- **[ServiceManager.ts](core/ServiceManager.ts)** - управління сервісами
- **[BaseService.ts](core/BaseService.ts)** - базовий клас для сервісів

### **🔧 config/** - Конфігурація
Налаштування та конфігурація системи:
- **[Config.ts](config/Config.ts)** - основна конфігурація
- **[Config.js](config/Config.js)** - JavaScript конфігурація
- **[environments.ts](config/environments.ts)** - налаштування середовищ
- **[docker/Dockerfile](config/docker/Dockerfile)** - Docker конфігурація
- **[environment/env.example](config/environment/env.example)** - приклад змінних середовища

### **🛠️ utils/** - Утиліти
Допоміжні функції та утиліти:
- **[logger.ts](utils/logger.ts)** - система логування
- **[errorHandler.ts](utils/errorHandler.ts)** - обробка помилок
- **[formatters.ts](utils/formatters.ts)** - форматування даних
- **[pagination.ts](utils/pagination.ts)** - пагінація
- **[security.ts](utils/security.ts)** - безпека
- **[performanceOptimizer.ts](utils/performanceOptimizer.ts)** - оптимізація продуктивності
- **[retry.ts](utils/retry.ts)** - механізм повторних спроб
- **[queueManager.ts](utils/queueManager.ts)** - управління чергами
- **[clusterManager.ts](utils/clusterManager.ts)** - управління кластерами
- **[fileProcessor.ts](utils/fileProcessor.ts)** - обробка файлів
- **[exportHelpers.ts](utils/exportHelpers.ts)** - експорт даних
- **[formulaProcessor.ts](utils/formulaProcessor.ts)** - обробка формул
- **[uiHelpers.ts](utils/uiHelpers.ts)** - UI допоміжники
- **[aiEnhanced.ts](utils/aiEnhanced.ts)** - розширені AI функції

### **🧪 tests/** - Тестування
Тести всіх компонентів системи:
- **[setup.ts](tests/setup.ts)** - налаштування тестів
- **[unit/**](tests/unit/)** - unit тести
- **[integration/**](tests/integration/)** - integration тести
- **[e2e/**](tests/e2e/)** - end-to-end тести
- **[performance/**](tests/performance/)** - тести продуктивності
- **[load/**](tests/load/)** - навантажувальні тести

### **📋 types/** - Типи TypeScript
- **[index.ts](types/index.ts)** - основні типи системи

### **📜 scripts/** - Скрипти
- **[deployCommands.ts](scripts/deployCommands.ts)** - розгортання команд

---

## 🏗️ **АРХІТЕКТУРНІ ПРИНЦИПИ**

### **🎯 SOLID принципи**
- **Single Responsibility** - кожен клас має одну відповідальність
- **Open/Closed** - відкритий для розширення, закритий для модифікації
- **Liskov Substitution** - підкласи можуть замінювати базові класи
- **Interface Segregation** - інтерфейси розділені на менші частини
- **Dependency Inversion** - залежності від абстракцій, не від конкретних класів

### **🔄 Dependency Injection**
- **ServiceContainer** - централізований контейнер сервісів
- **Constructor Injection** - ін'єкція залежностей через конструктор
- **Interface-based Design** - проектування на основі інтерфейсів

### **🛡️ Error Handling**
- **Centralized Error Handling** - централізована обробка помилок
- **Graceful Degradation** - плавне зниження функціональності
- **Comprehensive Logging** - детальне логування всіх операцій

### **⚡ Performance**
- **Caching Strategy** - стратегія кешування
- **Async/Await** - асинхронна обробка
- **Memory Management** - управління пам'яттю
- **Resource Pooling** - пул ресурсів

---

## 🔧 **ТЕХНІЧНІ ДЕТАЛІ**

### **📦 Залежності**
- **Discord.js** - Discord API
- **OpenAI** - AI функціональність
- **Google APIs** - Google Sheets/Drive
- **Redis** - кешування
- **Winston** - логування
- **Jest** - тестування

### **🔄 Життєвий цикл**
1. **Ініціалізація** - завантаження конфігурації
2. **Реєстрація** - реєстрація команд та сервісів
3. **Підключення** - підключення до Discord
4. **Обробка** - обробка повідомлень
5. **Моніторинг** - моніторинг стану системи

### **📊 Метрики**
- **Response Time** - час відповіді
- **Memory Usage** - використання пам'яті
- **Error Rate** - частота помилок
- **Command Usage** - статистика використання команд

---

## 🚀 **РОЗРОБКА**

### **📝 Створення нової команди**
```typescript
import { BaseCommand } from './BaseCommand';

export class NewCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(config, {
      name: 'new-command',
      description: 'Опис команди',
      options: [
        // опції команди
      ]
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    // логіка команди
  }
}
```

### **🛠️ Створення нового сервісу**
```typescript
import { BaseService } from '../core/BaseService';

export class NewService extends BaseService {
  constructor(container: ServiceContainer) {
    super(container);
  }

  async initialize(): Promise<void> {
    // ініціалізація сервісу
  }

  async execute(): Promise<any> {
    // виконання логіки
  }
}
```

---

## 📚 **ДОКУМЕНТАЦІЯ**

### **📖 Детальна документація**
- **[Архітектура](../docs/architecture/ARCHITECTURE.md)** - технічна архітектура
- **[API документація](../docs/api/API_DOCUMENTATION.md)** - API довідник
- **[Команди](../docs/api/COMMANDS_REFERENCE.md)** - довідник команд

### **🧪 Тестування**
- **[Гід тестування](../docs/guides/TESTING_GUIDE.md)** - як тестувати
- **[Unit тести](tests/unit/)** - unit тести
- **[Integration тести](tests/integration/)** - integration тести

---

## 🤝 **КОНТАКТИ**

**👨‍💻 Автор:** Dmitry Shivachov (Dmitze)  
**📧 Email:** dmitze_shivachov@outlook.com  
**🌐 GitHub:** https://github.com/Dmitze  
**💬 Discord:** dmitry_shivachov3756  
**📱 Telegram:** https://t.me/Dmitry_Shiva  

---

**🦖 Godzilla Bot - Потужний, Надійний, Український!** 