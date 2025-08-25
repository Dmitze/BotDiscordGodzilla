# API та Команди (огляд)

- Довідник команд: [api/COMMANDS_REFERENCE.md](api/COMMANDS_REFERENCE.md)
- API документація: [api/API_DOCUMENTATION.md](api/API_DOCUMENTATION.md)

## Короткий довідник команд

| Команда | Призначення | Основні опції | Довідник |
| --- | --- | --- | --- |
| `/пошук` | Пошук у FTS/Google/гібрид | `запит`, `режим` | [api/COMMANDS_REFERENCE.md#search](api/COMMANDS_REFERENCE.md#search) |
| `/ai` | AI асистент/RAG | `запит`, `режим`, `макс_токенів` | [api/COMMANDS_REFERENCE.md#ai](api/COMMANDS_REFERENCE.md#ai) |
| `/док` | Документація/гайди | `розділ` | [api/COMMANDS_REFERENCE.md#doc](api/COMMANDS_REFERENCE.md#doc) |
| `/stats` | Статистика використання | `період` | [api/COMMANDS_REFERENCE.md#statistics](api/COMMANDS_REFERENCE.md#statistics) |

Примітка: усі інтерактивні компоненти підписані через `signComponentId` (HMAC+TTL).

## Приклади
```bash
/пошук запит:"особовий склад"
/ai запит:"проаналізуй звіт"
```

### Приклади з опціями

```bash
/пошук запит:"FTS vs hybrid" режим:hybrid
/ai запит:"Підготуй резюме" макс_токенів:512
```

## Пов'язані документи

- Архітектура: [architecture/ARCHITECTURE.md](architecture/ARCHITECTURE.md)
- RAG/Пошук: [guides/rag.md](guides/rag.md)
- Безпека компонентів: [security/SECURITY_GUIDE.md](security/SECURITY_GUIDE.md)
