# 📁 Команди бота (`src/commands/`)

Я розробив та підтримую модуль команд, який відповідає за взаємодію користувача з ботом через Slash-команди Discord. Кожна команда типізована, має чітку валідацію аргументів, структурні логи та інкапсульовану бізнес-логіку.

## 🔗 Швидкі посилання
- Архітектура команд: ../../docs/architecture/NEW_COMMANDS_ARCHITECTURE.md
- Довідник команд: ../../docs/api/COMMANDS_REFERENCE.md
- Метрики і валідація команд: modules/README.md

## 📦 Вміст каталогу
- `AIAssistantCommand.ts` — AI асистент (натуральна мова, аналіз, відповіді)
- `AnalyticsCommand.ts` — розширена аналітика
- `DocumentsCommand.ts` — робота з документами
- `EnhancedSearchCommand.ts` — покращений пошук
- `FileManagerCommand.ts` — файлові операції
- `OperationsCommand.ts` — операційна діяльність
- `PerformanceCommand.ts` — продуктивність та моніторинг
- `SearchCommand.ts` — пошук
- `statistics.ts` — статистика
- `BaseCommand.ts`, `BaseCommandRefactored.ts` — базові класи команд
- `modules/` — підмодулі метрик і валідації команд

## 🧭 Принципи
- Строга типізація, `exactOptionalPropertyTypes`
- Структурне логування з `component` і meta
- Безпечна валідація вводу та обробка помилок

## 🧩 Як додати нову команду
1) Створіть файл `MyCommand.ts` на основі `BaseCommand.ts`
2) Зареєструйте команду в `CommandManager.ts`
3) Додайте валідацію в `modules/CommandValidator.ts`
4) Додайте метрики в `modules/CommandMetrics.ts`

---

### 📞 Контакти
- Discord: dmitry_shivachov3756  
- Telegram: https://t.me/Dmitry_Shiva  
- Email: dmitze_shivachov@outlook.com  
- GitHub: https://github.com/Dmitze  
- Проект: https://github.com/Dmitze/BotDiscordGodzilla
