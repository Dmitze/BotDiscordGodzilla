# Гайд розробника

## Технології

- Node.js 18+, TypeScript, Discord.js, Jest, ESLint/Prettier
- Архітектура: команди (`src/commands`), сервіси (`src/services`), ядро (`src/core`)

## Як працювати з кодом

- Налаштування: див. `README.md` (швидкий старт)
- Структура: див. `architecture/ARCHITECTURE.md`
- README підпапок: `src/.../README.md`

## Тести та якість

- `npm test`, `npm run lint`, `npm run type-check`
- Покриття не зменшуємо; нова логіка — з тестами.

## Безпека

- Політика: `SECURITY.md`, деталі: `security/SECURITY_GUIDE.md`
- Секрети у `.env`, не комітьте ключі.

## Швидкий старт

1. Встановіть залежності:

```bash
npm ci
```

1. Сконфігуруйте змінні середовища:

- Скопіюйте `env.example` → `.env`
- Заповніть ключі Discord/Google/AI (див. `guides/SETUP.md`)

1. Запуск у dev-режимі:

```bash
npm run dev
```

1. Реєстрація команд (за потреби):

```bash
npm run bot:register-commands
```

## Корисні npm-скрипти

```bash
# Типи та збірка
npm run type-check
npm run build

# Тести
npm test
npm run test:unit
npm run test:integration
npm run test:e2e

# Лінт та форматування
npm run lint
npm run format

# Команди бота
npm run bot:register-commands
npm run bot:clear-commands
```

## CI/CD (приклад GitHub Actions)

```yaml
name: ci
on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]
jobs:
  build-test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 18
          cache: npm
      - run: npm ci
      - run: npm run type-check
      - run: npm run lint
      - run: npm test -- --ci
```

## Релізи

- Семантичні версії: `MAJOR.MINOR.PATCH`
- CHANGELOG: `changelog/CHANGELOG.md`
- Теги релізів у Git: `vX.Y.Z`
- Перевірте безпеку перед релізом: `npm audit` + оновлення залежностей

## Профілювання та логування

- Централізований логер; `console.*` заборонено поза `src/scripts/`
- Рівні логів: `debug|info|warn|error`; у проді — `info+`
- Метрики/спостережуваність: Prometheus-метрики безпеки (див. `security/SECURITY_GUIDE.md`)
- Трейсинг компонентних дій через кореляційні ID у логері

## Корисні посилання

- Архітектура: `architecture/ARCHITECTURE.md`
- Безпека: `SECURITY.md`, `security/SECURITY_GUIDE.md`
- RAG/пошук: `guides/rag.md`
