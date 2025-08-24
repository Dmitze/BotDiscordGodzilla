# API та Команди (огляд)

- Довідник команд: `docs/api/COMMANDS_REFERENCE.md`
- API документація: `docs/api/API_DOCUMENTATION.md`

## Короткий довідник команд

| Команда | Призначення | Основні опції | Довідник |
| --- | --- | --- | --- |
| `/пошук` | Пошук у FTS/Google/гібрид | `запит`, `режим` | `docs/api/COMMANDS_REFERENCE.md#search` |
| `/ai` | AI асистент/RAG | `запит`, `режим`, `макс_токенів` | `docs/api/COMMANDS_REFERENCE.md#ai` |
| `/док` | Документація/гайди | `розділ` | `docs/api/COMMANDS_REFERENCE.md#doc` |
| `/stats` | Статистика використання | `період` | `docs/api/COMMANDS_REFERENCE.md#statistics` |

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

- Архітектура: `docs/ARCHITECTURE.md`
- RAG/Пошук: `docs/guides/rag.md`
- Безпека компонентів: `docs/security/SECURITY_GUIDE.md`
