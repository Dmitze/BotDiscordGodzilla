# Архітектура (огляд)

- Повна специфікація: `docs/architecture/ARCHITECTURE.md`
- Дорожня карта: `docs/architecture/ROADMAP.md`
- Модулі команд (історичний документ): `docs/archive/NEW_COMMANDS_ARCHITECTURE.md`

## Карта проєкту

```mermaid
flowchart LR
  Client(Discord Client) -->|Slash/Interaction| BotCore
  BotCore[Core] --> Commands
  BotCore --> Services
  Services --> GoogleAPI
  Services --> AI(LLM API)
  Services --> Cache[(Cache)]
  Commands --> Search[SearchCommand]
  Commands --> Docs[DocCommand]
```

## Системна діаграма (взаємодії)

```mermaid
sequenceDiagram
  participant U as User
  participant D as Discord Gateway
  participant C as Bot Core (src/core)
  participant CMD as Commands (src/commands)
  participant S as Services (src/services)
  participant SI as SearchIndex (src/search)
  participant RAG as RagPipeline (src/rag)

  U->>D: Slash /interaction
  D->>C: InteractionCreate
  C->>CMD: Route to BaseCommand
  CMD->>S: ServiceManager.get(...)
  alt Search
    CMD->>SI: search(query, mode)
    SI-->>CMD: results + metadata
  else RAG
    CMD->>RAG: retrieve+augment(query)
    RAG->>S: EmbeddingsService.compute()
    RAG->>SI: hybridSearch(fts+cosine)
    RAG-->>CMD: context chunks + citations
  end
  CMD-->>C: Render components (signed)
  C-->>D: Reply (ephemeral)
```

## Компоненти (мапа каталогів)

- `src/core/` — ядро: `ServiceManager`, `CommandRouter`, життєвий цикл, конфіг.
- `src/commands/` — `BaseCommand`, `SearchCommand`, `DocCommand`, `StatisticsCommand`.
- `src/services/` — `GoogleService`, `EmbeddingsService`, кеш/логер тощо.
- `src/search/` — FTS/SQLite індекс, токенізатор, DDL/міграції.
- `src/rag/` — Retriever/Augmenter/Generator, `RagService`.
- `src/ui/` — побудова карток/кнопок, `signComponentId`.
- `src/tests/` — unit/integration/e2e, сетап моку безпеки.

## Потоки даних: FTS / Embeddings / RAG

```mermaid
flowchart TB
  subgraph Indexing
    A[DriveIndexerService] --> P[Parsers (PDF/DOCX/TXT/Sheets)]
    P --> N[Normalizer]
    N --> SC[segment_cache]
    N --> FTS[SQLite FTS Index]
    N --> EMB[Embeddings (optional)]
  end
  subgraph Retrieval
    Q[Query] -->|FTS| FTS
    Q -->|Cosine| EMB
    FTS & EMB --> HR[Hybrid Retriever (alpha,k)]
  end
  HR --> AUG[Augmenter: select chunks + cite]
  AUG --> GEN[Generator: LLM]
  GEN --> OUT[Answer + Citations]
```

Параметри керуються через `.env` (див. `env.example` та `docs/guides/rag.md`).

## Безпека та i18n

- `signComponentId` з HMAC + TTL для всіх компонентів UI; у тестовому режимі — legacy-сумісність.
- Централізований логер; `no-console` окрім `src/scripts/`.
- Локаль за замовчуванням: `uk` (українська).
