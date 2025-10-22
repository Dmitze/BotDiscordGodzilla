# 🏗️ Архітектура нових сервісів Discord AI Assistant Bot

## 📋 Зміст

- [🎯 Огляд](#-огляд)
- [🧠 ContextMemoryService](#-contextmemoryservice)
- [ ResponseCacheService](#-responsecacheservice)
- [📚 KnowledgeBaseService](#-knowledgebaseservice)
- [🔍 EnhancedRagService](#-enhancedragservice)
- [🔄 Інтеграція сервісів](#-інтеграція-сервісів)
- [🧪 Тестування](#-тестування)

---

## 🎯 Огляд

У версії 3.0.0 бот отримав чотири нові ключові сервіси, які значно покращують його можливості:

1. **ContextMemoryService** - зберігає контекст користувача та історію запитів
2. **ResponseCacheService** - кешування відповідей з TTL для покращення продуктивності
3. **KnowledgeBaseService** - управління базою знань з категоризацією та тегуванням
4. **EnhancedRagService** - покращений RAG з автоматичною індексацією Google Drive

---

## 🧠 ContextMemoryService

### Призначення
Зберігає історію запитів користувачів та їхні переваги для контекстної обробки.

### Архітектура
``mermaid
graph TB
    A[ContextMemoryService] --> B[LRU Cache]
    A --> C[User Preferences]
    A --> D[Query History]
    B --> E[Memory Storage]
    C --> F[Preference Tracking]
    D --> G[Query Context]
```

### Основні функції
- Зберігання останніх 5 запитів користувача
- Відстеження користувацьких переваг (мова, домен, стиль відповідей)
- Формування контекстних запитів для AI
- Автоматичне очищення старих даних

### API
```typescript
class ContextMemoryService {
  addQuery(userId: string, query: string): void
  getRecentQueries(userId: string, limit?: number): QueryContext[]
  updateUserPreferences(userId: string, preferences: UserPreferences): void
  buildContextualPrompt(userId: string, currentQuery: string): string
  getStats(): ContextStats
}
```

---

## 💾 ResponseCacheService

### Призначення
Кешування відповідей з TTL (30 хвилин) для покращення продуктивності та зменшення навантаження на AI.

### Архітектура
``mermaid
graph TB
    A[ResponseCacheService] --> B[Map-based Storage]
    A --> C[TTL Management]
    A --> D[Pattern-based Keys]
    B --> E[Cache Entries]
    C --> F[Expiration Cleanup]
    D --> G[Key Generation]
```

### Основні функції
- Кешування відповідей з TTL 30 хвилин
- Пошук за патернами ключів
- Статистика використання кешу
- Автоматичне очищення прострочених записів

### API
```typescript
class ResponseCacheService {
  get<T>(key: string): T | null
  set<T>(key: string, value: T, ttlMinutes?: number): void
  delete(key: string): boolean
  clear(): void
  getStats(): CacheStats
  extendTTL(key: string, additionalMinutes: number): boolean
}
```

---

## 📚 KnowledgeBaseService

### Призначення
Управління структурованою базою знань з категоризацією, тегуванням та пошуком.

### Архітектура
``mermaid
graph TB
    A[KnowledgeBaseService] --> B[Entry Management]
    A --> C[Search Engine]
    A --> D[Tag System]
    B --> E[Knowledge Entries]
    C --> F[Semantic Search]
    C --> G[Keyword Search]
    D --> H[Tag Indexing]
```

### Основні функції
- Створення та управління записами знань
- Категоризація та тегування знань
- Пошук за ключовими словами та семантичний пошук
- Інтеграція з AI для обробки знань

### API
```typescript
class KnowledgeBaseService {
  addEntry(entry: KnowledgeEntry): string
  getEntry(id: string): KnowledgeEntry | null
  search(query: string, options?: SearchOptions): KnowledgeEntry[]
  updateEntry(id: string, updates: Partial<KnowledgeEntry>): boolean
  deleteEntry(id: string): boolean
  getTrendingTopics(limit?: number): TrendingTopic[]
  getStats(): KnowledgeStats
}
```

---

## 🔍 EnhancedRagService

### Призначення
Покращений RAG сервіс з автоматичною індексацією документів Google Drive.

### Архітектура
``mermaid
graph TB
    A[EnhancedRagService] --> B[Auto Indexing]
    A --> C[RAG Pipeline]
    A --> D[Google Drive Integration]
    B --> E[File Monitoring]
    B --> F[Indexing Scheduler]
    C --> G[Document Retrieval]
    C --> H[Context Augmentation]
    D --> I[Drive API]
    D --> J[File Processing]
```

### Основні функції
- Автоматична індексація документів Google Drive
- Планування індексації за розкладом
- Підтримка різних типів файлів (PDF, Google Docs, Sheets, Word, Excel, текст)
- Інтеграція з існуючим RAG функціоналом

### API
```typescript
class EnhancedRagService extends RagService {
  search(query: string, options?: SearchOptions): Promise<SearchResult[]>
  triggerManualIndexing(folderId?: string): Promise<void>
  getIndexingStats(): IndexingStats
  updateAutoIndexConfig(config: Partial<AutoIndexConfig>): void
  shutdown(): Promise<void>
}
```

---

## 🔄 Інтеграція сервісів

### Потік даних
``mermaid
sequenceDiagram
    participant U as Користувач
    participant C as ContextMemoryService
    participant K as KnowledgeBaseService
    participant R as EnhancedRagService
    participant Cache as ResponseCacheService
    participant AI as AIService
    
    U->>C: Новий запит
    C->>C: Зберігає запит в історію
    C->>C: Формує контекстний запит
    C->>Cache: Перевіряє кеш
    alt Кеш знайдено
        Cache-->>U: Повертає кешовану відповідь
    else Кеш не знайдено
        C->>K: Пошук в базі знань
        K-->>C: Результати пошуку
        C->>R: RAG пошук
        R-->>C: Результати RAG
        C->>AI: Запит до AI з контекстом
        AI-->>C: Відповідь AI
        C->>Cache: Зберігає відповідь в кеш
        C->>U: Повертає відповідь користувачу
    end
```

### ServiceManager інтеграція
``typescript
// Реєстрація нових сервісів
container.register('contextMemory', ContextMemoryService);
container.register('responseCache', ResponseCacheService);
container.register('knowledgeBase', KnowledgeBaseService);
container.register('enhancedRag', EnhancedRagService);
```

---

## 🧪 Тестування

### Покриття тестами
- Unit тести для кожного сервісу
- Інтеграційні тести для взаємодії сервісів
- E2E тести для повного робочого процесу

### Ключові тести
1. **ContextMemoryService**
   - Зберігання та отримання історії запитів
   - Управління перевагами користувача
   - Формування контекстних запитів

2. **ResponseCacheService**
   - Кешування та отримання значень
   - TTL управління
   - Пошук за патернами

3. **KnowledgeBaseService**
   - Створення та управління записами
   - Пошук за ключовими словами
   - Семантичний пошук

4. **EnhancedRagService**
   - Автоматична індексація
   - Пошук документів
   - Інтеграція з Google Drive

### Метрики тестування
- Покриття тестами: 95%+
- Час відповіді: < 1.5 секунди
- Успішність інтеграційних тестів: 100%

# Нові сервіси для взаємодії з Google документами

## Огляд реалізованих покращень

У цьому документі описано нові сервіси та команди, реалізовані для покращення взаємодії з Google документами:

### 1. Покращений навігатор по документах
- **Команда**: `/drive-navigate`
- **Покращення**: Додано розширені фільтри за типами файлів, датами, розміром
- **Файли**: `src/commands/DriveNavigateCommand.ts`

### 2. Розумна класифікація документів
- **Сервіс**: `SmartDocumentClassifier`
- **Функціонал**: Автоматична категоризація документів за типом вмісту
- **Файли**: `src/services/SmartDocumentClassifier.ts`

### 3. Покращений пошук в документах
- **Команда**: `/drive-search`
- **Функціонал**: Розширений пошук з фасетними фільтрами
- **Файли**: `src/commands/EnhancedDriveSearchCommand.ts`

### 4. Інтерактивні картки документів
- **Компонент**: `DocumentCardBuilder`
- **Функціонал**: Розширені картки з прев'ю вмісту та статистикою
- **Файли**: `src/ui/DocumentCardBuilder.ts`

### 5. Сповіщення про зміни
- **Сервіс**: `DriveChangesService`
- **Функціонал**: Моніторинг змін в документах та сповіщення
- **Файли**: `src/services/DriveChangesService.ts`

### 6. Багатомовна підтримка
- **Сервіс**: `MultilingualDocumentProcessor`
- **Функціонал**: Автоматичне визначення мови та переклад документів
- **Файли**: `src/services/MultilingualDocumentProcessor.ts`

### 7. Аналітика використання
- **Сервіс**: `DocumentAnalyticsService`
- **Функціонал**: Статистика використання та рекомендації
- **Файли**: `src/services/DocumentAnalyticsService.ts`

### 8. Згадування документів в чаті
- **Сервіс**: `DocumentMentionHandler`
- **Функціонал**: Згадування документів за назвою та швидке прикріплення
- **Файли**: `src/services/DocumentMentionHandler.ts`

### 9. Автоматична обробка нових документів
- **Сервіс**: `AutomatedDocumentProcessor`
- **Функціонал**: Тригери на нові файли та автоматична обробка
- **Файли**: `src/services/AutomatedDocumentProcessor.ts`

### 10. Експорт та імпорт
- **Сервіс**: `DocumentExportImportService`
- **Функціонал**: Експорт результатів пошуку та документів у різних форматах
- **Файли**: `src/services/DocumentExportImportService.ts`

## Детальний опис сервісів

### SmartDocumentClassifier
Сервіс для автоматичної класифікації документів на основі їх вмісту та метаданих. Підтримує кілька категорій документів:
- Накази
- Звіти
- персонал
- Матеріально-технічне забезпечення
- Фінансові документи
- Операційні документи
- Навчальні матеріали
- Комунікації

### EnhancedDriveSearchCommand
Команда для розширеного пошуку документів з підтримкою:
- Фільтрації за типом файлу
- Фільтрації за датою створення/зміни
- Фільтрації за розміром файлу
- Сортування результатів
- Пагінації

### DocumentCardBuilder
Компонент для створення інтерактивних карток документів з:
- Прев'ю вмісту
- Статистикою документа (розмір, дата зміни, автор)
- Швидкими діями (аналіз, експорт, тегування)
- Тегами документа

### DriveChangesService
Сервіс для моніторингу змін у документах Google Drive:
- Відстеження створення нових файлів
- Відстеження змін в існуючих файлах
- Сповіщення про зміни в Discord
- Історія змін

### MultilingualDocumentProcessor
Сервіс для багатомовної обробки документів:
- Автоматичне визначення мови документа
- Переклад документів зі збереженням форматування
- Підтримка кількох мов (українська, англійська, російська, польська, німецька, французька, іспанська)

### DocumentAnalyticsService
Сервіс для аналітики використання документів:
- Відстеження доступу до документів
- Аналіз пошукових патернів користувачів
- Генерація персоналізованих рекомендацій
- Статистика найпопулярніших документів

### DocumentMentionHandler
Сервіс для обробки згадок документів у чаті:
- Розпізнавання згадок документів у повідомленнях
- Автоматичне надсилання інформації про документ
- Швидке прикріплення документів до повідомлень

### AutomatedDocumentProcessor
Сервіс для автоматичної обробки нових документів:
- Налаштовувані тригери на нові файли
- Автоматичний аналіз та класифікація документів
- Автоматичне тегування на основі вмісту
- Сповіщення про нові документи

### DocumentExportImportService
Сервіс для експорту та імпорту документів:
- Експорт результатів пошуку у різних форматах (PDF, DOCX, XLSX, CSV, TXT, JSON)
- Експорт окремих документів
- Імпорт документів з локальних файлів
- Синхронізація з локальними файлами
- Резервне копіювання важливих документів

## Взаємодія сервісів

Всі нові сервіси інтегровані з існуючою архітектурою бота та використовують спільні компоненти:
- GoogleService для взаємодії з Google Drive API
- AIService для обробки мови та генерації відповідей
- SchedulerService для планування завдань
- CacheService для кешування результатів

## Використання

Для використання нових функцій користувачі може використовувати наступні команди:
- `/drive-navigate` - навігація по документах з фільтрацією
- `/drive-search` - розширений пошук документів
- Згадки документів у повідомленнях (наприклад, "дивіться файл report.pdf")
- Автоматичні сповіщення про нові документи

Усі сервіси автоматично ініціалізуються при запуску бота та інтегровані з системою логування та моніторингу.
