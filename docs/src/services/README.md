# 🧠 Сервіси (`src/services/`)

Я розробив сервісний шар, який інкапсулює бізнес-логіку, інтеграції та крос-секційні можливості (кеш, метрики, планувальник, AI, Google API).

## 🔗 Швидкі посилання
- Архітектура системи: ../../docs/architecture/ARCHITECTURE.md
- Звіт міграції сервісів: ../../docs/documentation/PHASE3_SERVICES_MIGRATION_REPORT.md

## 📦 Вміст каталогу
- `AIService.ts` — AI можливості (санітизація вводу, безпека)
- `GoogleService.ts` — Google Sheets/Drive API, типізація, кеш-інвалідація
- `CacheService.ts` — кеш з TTL, idempotent API
- `MetricsService.ts` — Prometheus метрики, безпечний старт/стоп
- `SchedulerService.ts` — планування та фонові задачі

## 🛡️ Стандарти
- Структурний `logger` з `component: '...'`
- Перевірка аргументів, чіткі типи і контракти
- Граційне завершення, обробка помилок, healthchecks

---

### 📞 Контакти
- Discord: dmitry_shivachov3756  
- Telegram: https://t.me/Dmitry_Shiva  
- Email: dmitze_shivachov@outlook.com  
- GitHub: https://github.com/Dmitze  
- Проект: https://github.com/Dmitze/BotDiscordGodzilla
