# ⚙️ Ядро системи (`src/core/`)

Я спроєктував ядро як стабільний каркас: життєвий цикл бота, менеджери команд та подій, DI контейнер сервісів, централізована обробка помилок і прав.

## 📦 Вміст каталогу
- `Bot.ts` — ініціалізація, підключення, lifecycle, інтеграція сервісів
- `ServiceContainer.ts` — DI контейнер
- `ServiceManager.ts` — керування сервісами, старт/стоп (включно з `MetricsService`)
- `EventManager.ts` — підписки на події Discord
- `CommandManager.ts` — реєстрація та виконання команд
- `ErrorHandler.ts` — уніфікована обробка помилок
- `PermissionManager.ts` — права доступу та політики
- `BaseService.ts` — базовий клас сервісів

## 🔗 Пов'язані документи
- Архітектура: ../../docs/architecture/ARCHITECTURE.md
- Roadmap: ../../docs/architecture/ROADMAP.md

---

### 📞 Контакти
- Discord: dmitry_shivachov3756  
- Telegram: https://t.me/Dmitry_Shiva  
- Email: dmitze_shivachov@outlook.com  
- GitHub: https://github.com/Dmitze  
- Проект: https://github.com/Dmitze/BotDiscordGodzilla
