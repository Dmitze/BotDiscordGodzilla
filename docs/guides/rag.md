# RAG (Retrieval-Augmented Generation) у BotDiscordGodzilla

У цьому гіді описано, як увімкнути та використовувати RAG-пайплайн: Retriever → Augmenter → Generator.

## Компоненти

- **Retriever** (`src/rag/Retriever.ts`): дістає релевантні документи з `SearchIndex`.
  Підтримуються режими: `fts` (за замовчуванням) та `hybrid` (FTS + ембеддінги з косинусною схожістю).
- **Augmenter** (`src/rag/Augmenter.ts`): формує контекстні уривки, дотримується ліміту токенів, застосовує PII-маскування.
- **RagPipeline** (`src/rag/RagPipeline.ts`): збирає промпт українською з джерелами та викликає `AIService`.
- **RagService** (`src/services/RagService.ts`): сервіс-обгортка для використання у командах і чаті.

> Security: усі кнопки/селекти у UI використовують підписані `customId` (HMAC + TTL) для захисту від підміни та повторного відтворення. Деталі у `docs/guides/SECURITY.md`.

## Інтеграція

- **/ai** (`src/commands/aiAssistant.ts`): спершу намагається `RagService`, інакше — звичайний `AIService`.
- **Чат QnA** (`src/chat/ChatRouter.ts` → `replyQna()`): відповідає через RAG з цитуванням джерел.

## Змінні середовища

Додайте до `.env` (див. приклад у `env.example`):

- `RETRIEVER_K` — top‑K документів (дефолт 6)
- `RETRIEVER_ALPHA` — вага для гібридного режиму (0..1; дефолт 0.5). Чим більше — тим більший вплив ембеддінгів.
- `EMBEDDINGS_ENABLE` — вмикає гібридний режим (ембеддінги). `true|false` (дефолт false)
- `EMBEDDINGS_PROVIDER` — `openai|mock` (дефолт `mock` для локального запуску без ключа)
- `EMBEDDINGS_MODEL` — модель ембеддінгів (наприклад, `text-embedding-3-small`)
- `RAG_MAX_CONTEXT_TOKENS` — бюджет токенів для контексту (дефолт 1200)
- `AI_MAX_TOKENS` — обмеження токенів відповіді (дефолт 512)
- `SEARCH_INDEX_PATH`, `SEARCH_FTS_TOKENIZER`, `SEARCH_BATCH_SIZE` — налаштування індексу (SQLite FTS)

За використання `EMBEDDINGS_PROVIDER=openai` потрібен дійсний `OPENAI_API_KEY` (див. розділ AI у `env.example`).

### Перемикання токенайзера FTS і міграції

- `SEARCH_FTS_TOKENIZER=porter|unicode61`. Під час зміни значення індекс FTS буде автоматично перебудовано.
- Для швидкої перебудови використовується кеш сегментів (`segment_cache`) — нормалізований текст зберігається під час індексації та застосовується під час реконструкції `documents_fts`.

PII-маскування налаштовується у `src/config/security.ts` через `SECURITY_PII_*` змінні.

## Фільтри пошуку та Drive Q&A

- `SearchFilters` підтримує `fileId?: string[]` — можна обмежити пошук конкретним(и) файлом(ами).
- Кнопка `Question` у картці файлу (див. `src/ui/FileCardBuilder.ts` та обробку в `FileManagerCommand`) використовує RAG із `filters.fileId = [driveFileId]`. Якщо RAG недоступний, бот бере контекстні уривки зі `SearchIndex` тим самим фільтром. Лише якщо індекс недоступний/порожній — виконується експорт Google Drive як запасний шлях.
- Ліміти:
  - `RAG_MAX_CONTEXT_TOKENS` — бюджет контексту для RAG
  - `DRIVE_QA_MAX_TOKENS` — ліміт токенів відповіді (override `AI_MAX_TOKENS`)
  - `DRIVE_QA_MAX_CONTEXT_CHARS` — максимальна довжина текстового контексту зі сніпетів/експорту

## Приклади викликів

```ts
const rag = new RagService(searchIndex, aiService);
const ans = await rag.answer(
  'Що відомо про проект?',
  { k: 6, mode: 'hybrid' },// RetrieverOptions: 'fts' | 'hybrid'
  { maskPII: true },       // AugmentOptions
  { maxTokens: 512 },      // GenerateWithContextOptions
);
console.log(ans.answer, ans.chunks);
```

## Тести

- Юніт-тести у `src/tests/unit/services/`:
  - `RagPipeline.test.ts`: мок `SearchIndex` і `AIService`, перевірка PII-маскування, укр. промпту та лімітів.
  - `RagService.test.ts`: інтеграційний шлях через сервіс, перевірка цитувань і маскування.

## План подальших покращень

- Типобезпечний `getService()`/DI та зниження складності методів у командах/роутері.
- Більше евристик для відбору уривків (re-ranking, diversity, dedup).
