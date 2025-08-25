# Шаблон README модуля

> Скопіюйте цей файл у відповідну папку як `README.md` і заповніть секції.

## Призначення

Короткий опис модуля, його ролі в системі та цінності для користувачів/розробників.

## Структура

```text
<module-root>/
  ├─ src/                 # вихідний код
  ├─ tests/               # тести (unit/integration)
  ├─ fixtures/            # тестові дані (за потреби)
  └─ README.md            # цей файл
```

## Ключові модулі та інтерфейси

- `.../SomeService.ts` — основна логіка сервісу (інжектується через ServiceManager)
- `.../types.ts` — типи/інтерфейси публічного API модуля
- `.../utils.ts` — хелпери/утиліти

## Залежності й інтеграції

- Залежить від: `src/core/*`, `src/services/*` (вказати конкретні)
- Використовує конфігурацію з `.env` (див. `env.example`)
- Інтеграція з іншими підсистемами: RAG/FTS/Embeddings/Discord UI (вказати)

## Конфігурація

Перелік ключових змінних середовища й налаштувань:

- `MODULE_ENABLED` — вмикає/вимикає модуль (true/false)
- `MODULE_LIMIT` — ліміт або розмір батчів

Приклад `.env`:

```env
MODULE_ENABLED=true
MODULE_LIMIT=100
```

## Скрипти та команди

```bash
npm run test:unit  # юніт-тести для модуля
npm run lint       # лінт коду
```

## Тестування

- Покриття: не менше X%
- Юніт-тести: `tests/unit/*`
- Інтеграційні тести: `tests/integration/*`
- E2E (якщо є): короткий опис сценаріїв

## Приклади використання

```ts
import { SomeService } from './SomeService';

const svc = new SomeService(/* deps */);
await svc.doWork();
```

## Безпека та i18n

- Використовуйте підпис компонентів `signComponentId` для UI/дій (якщо застосовно)
- Локаль за замовчуванням: `uk` (українська)
- Не використовуйте `console.*` поза `src/scripts/`; застосовуйте централізований логер

## Діагностика та логування

- Рівні логів: `debug|info|warn|error`
- Метрики (за потреби): Prometheus-лічильники/таймери

## Повʼязані документи

- Архітектура: `docs/ARCHITECTURE.md`
- Гайд розробника: `docs/DEVELOPER_GUIDE.md`
- Безпека: `SECURITY.md`, `docs/security/SECURITY_GUIDE.md`
- RAG/пошук: `docs/guides/rag.md`
