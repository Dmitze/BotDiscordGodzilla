# RAG (Retrieval-Augmented Generation) у BotDiscordGodzilla

У цьому гіді описано, як увімкнути та використовувати RAG-пайплайн: Retriever → Augmenter → Generator.

## Компоненти
- **Retriever** (`src/rag/Retriever.ts`): дістає релевантні документи з `SearchIndex` (FTS, гібрид у планах).
- **Augmenter** (`src/rag/Augmenter.ts`): формує контекстні уривки, дотримується ліміту токенів, застосовує PII-маскування.
- **RagPipeline** (`src/rag/RagPipeline.ts`): збирає промпт українською з джерелами та викликає `AIService`.
- **RagService** (`src/services/RagService.ts`): сервіс-обгортка для використання у командах і чаті.

## Інтеграція
- **/ai** (`src/commands/aiAssistant.ts`): спершу намагається `RagService`, інакше — звичайний `AIService`.
- **Чат QnA** (`src/chat/ChatRouter.ts` → `replyQna()`): відповідає через RAG з цитуванням джерел.

## Змінні середовища
Додайте до `.env` (див. приклад у `env.example`):

- `RETRIEVER_K` — top‑K документів (дефолт 6)
- `RETRIEVER_ALPHA` — вага для гібридного режиму (планується)
- `EMBEDDINGS_ENABLE` — вмикає гібридний режим (ембеддінги) — планується
- `RAG_MAX_CONTEXT_TOKENS` — бюджет токенів для контексту (дефолт 1200)
- `AI_MAX_TOKENS` — обмеження токенів відповіді (дефолт 512)
- `SEARCH_INDEX_PATH`, `SEARCH_FTS_TOKENIZER`, `SEARCH_BATCH_SIZE` — налаштування індексу (SQLite FTS)

PII-маскування налаштовується у `src/config/security.ts` через `SECURITY_PII_*` змінні.

## Приклади викликів
```ts
const rag = new RagService(searchIndex, aiService);
const ans = await rag.answer(
  'Що відомо про проект?',
  { k: 6 },                // RetrieverOptions
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
- Гібридний ретрівер з ембеддінгами (OpenAI чи локальні моделі) + `RETRIEVER_ALPHA` злиття скорів.
- Типобезпечний `getService()`/DI та зниження складності методів у командах/роутері.
- Більше евристик для відбору уривків (re-ranking, diversity, dedup).
