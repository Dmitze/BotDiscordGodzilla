# RAG (Retrieval-Augmented Generation) in BotDiscordGodzilla

This guide describes how to enable and use the RAG pipeline: Retriever → Augmenter → Generator.

## Components

- **Retriever** (`src/rag/Retriever.ts`): retrieves relevant documents from `SearchIndex`.
  Supported modes: `fts` (default) and `hybrid` (FTS + embeddings with cosine similarity).
- **Augmenter** (`src/rag/Augmenter.ts`): forms contextual snippets, adheres to token limits, applies PII masking.
- **RagPipeline** (`src/rag/RagPipeline.ts`): assembles Ukrainian prompt with sources and calls `AIService`.
- **RagService** (`src/services/RagService.ts`): service wrapper for use in commands and chat.

> Security: all UI buttons/selects use signed `customId` (HMAC + TTL) to protect against tampering and replay. Details in `docs/guides/SECURITY.md`.

## Integration

- **/ai** (`src/commands/aiAssistant.ts`): first tries `RagService`, otherwise uses regular `AIService`.
- **QnA Chat** (`src/chat/ChatRouter.ts` → `replyQna()`): responds through RAG with cited sources.

## Environment Variables

Add to `.env` (see example in `env.example`):

- `RETRIEVER_K` — top-K documents (default 6)
- `RETRIEVER_ALPHA` — weight for hybrid mode (0..1; default 0.5). Higher values give more influence to embeddings.
- `EMBEDDINGS_ENABLE` — enables hybrid mode (embeddings). `true|false` (default false)
- `EMBEDDINGS_PROVIDER` — `openai|mock` (default `mock` for local run without key)
- `EMBEDDINGS_MODEL` — embedding model (e.g., `text-embedding-3-small`)
- `RAG_MAX_CONTEXT_TOKENS` — context token budget (default 1200)
- `AI_MAX_TOKENS` — response token limit (default 512)
- `SEARCH_INDEX_PATH`, `SEARCH_FTS_TOKENIZER`, `SEARCH_BATCH_SIZE` — index settings (SQLite FTS)

When using `EMBEDDINGS_PROVIDER=openai`, a valid `OPENAI_API_KEY` is required (see AI section in `env.example`).

### FTS Tokenizer Switching and Migration

- `SEARCH_FTS_TOKENIZER=porter|unicode61`. When changing the value, the FTS index will be automatically rebuilt.
- For fast rebuild, a segment cache (`segment_cache`) is used — normalized text is stored during indexing and applied during `documents_fts` reconstruction.

PII masking is configured in `src/config/security.ts` through `SECURITY_PII_*` variables.

## Search Filters and Drive Q&A

- `SearchFilters` supports `fileId?: string[]` — you can limit search to specific file(s).
- The `Question` button in the file card (see `src/ui/FileCardBuilder.ts` and processing in `FileManagerCommand`) uses RAG with `filters.fileId = [driveFileId]`. If RAG is unavailable, the bot takes contextual snippets from `SearchIndex` with the same filter. Only if the index is unavailable/empty — Google Drive export is executed as a fallback.
- Limits:
  - `RAG_MAX_CONTEXT_TOKENS` — context budget for RAG
  - `DRIVE_QA_MAX_TOKENS` — response token limit (overrides `AI_MAX_TOKENS`)
  - `DRIVE_QA_MAX_CONTEXT_CHARS` — maximum length of text context from snippets/export

## Call Examples

```ts
const rag = new RagService(searchIndex, aiService);
const ans = await rag.answer(
  'What is known about the project?',
  { k: 6, mode: 'hybrid' },// RetrieverOptions: 'fts' | 'hybrid'
  { maskPII: true },       // AugmentOptions
  { maxTokens: 512 },      // GenerateWithContextOptions
);
console.log(ans.answer, ans.chunks);
```

## Tests

- Unit tests in `src/tests/unit/services/`:
  - `RagPipeline.test.ts`: mock `SearchIndex` and `AIService`, check PII masking, Ukrainian prompt, and limits.
  - `RagService.test.ts`: integration path through service, check citations and masking.

## Future Improvements Plan

- Type-safe `getService()`/DI and reducing method complexity in commands/router.
- More heuristics for snippet selection (re-ranking, diversity, dedup).