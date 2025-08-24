# Гайд розробника

## Технології

- Node.js 18+, TypeScript, Discord.js, Jest, ESLint/Prettier
- Архітектура: команди (`src/commands`), сервіси (`src/services`), ядро (`src/core`)

## Як працювати з кодом

- Налаштування: див. `README.md` (швидкий старт)
- Структура: див. `docs/ARCHITECTURE.md`
- README підпапок: `docs/src/.../README.md`

## Тести та якість

- `npm test`, `npm run lint`, `npm run type-check`
- Покриття не зменшуємо; нова логіка — з тестами.

## Безпека

- Політика: `SECURITY.md`, деталі: `docs/security/SECURITY_GUIDE.md`
- Секрети у `.env`, не комітьте ключі.
