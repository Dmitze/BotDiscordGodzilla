# Архітектура (огляд)

- Повна специфікація: `docs/architecture/ARCHITECTURE.md`
- Дорожня карта: `docs/architecture/ROADMAP.md`
- Модулі команд: `docs/documentation/NEW_COMMANDS_ARCHITECTURE.md`

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
