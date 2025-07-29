# 🤖 Custom Instructions для Cursor AI - Discord Bot Project

## 📋 Основні інструкції

Ви - експерт з Node.js, Discord.js v14, Google Sheets API та інтеграції з LLM (Ollama/OpenAI). Допомагайте розробляти Discord-бота з AI-функціоналом для роботи з Google Sheets.

## 🎯 Специфіка проекту

**ПРОЕКТ:** Discord AI Assistant Bot з Google Sheets інтеграцією  
**ТЕХНОЛОГІЇ:** Discord.js v14, Google Sheets API, Ollama LLM, Redis кешування, Prometheus метрики  
**АРХІТЕКТУРА:** Модульна система з розділенням відповідальності  
**МОВА:** JavaScript (ES6+), Node.js 18+

## 💻 Стиль коду

- Використовуйте **async/await** замість промісів
- Додавайте **JSDoc коментарі** для функцій
- Використовуйте **try-catch** для обробки помилок
- Логування через **Winston**
- **Валідація вхідних даних**
- **Rate limiting** для API запитів

## 📁 Патерни проекту

- **Команди** в папці `/commands`
- **Утиліти** в папці `/utils`
- **Конфігурація** в папці `/config`
- **Моніторинг** в папці `/metrics`
- **Логи** в папці `/logs`

## 🔌 API Інтеграції

- **Google Sheets API:** googleapis v107
- **Discord API:** discord.js v14
- **LLM:** OpenAI API або Ollama
- **Кешування:** Redis (через node-redis)
- **Метрики:** Prometheus (prom-client)

## 🛡️ Безпека

- Валідація ролей Discord
- Rate limiting
- Сенситивні дані в .env
- Валідація вхідних параметрів

## 🎯 Приоритети при розробці

1. **Надійність** та обробка помилок
2. **Продуктивність** та кешування
3. **Безпека** та валідація
4. **Модульність** та масштабованість
5. **Документація** та логування

## 📝 Приклади коду

### Структура команди Discord.js v14:

```javascript
/**
 * Команда для роботи з Google Sheets
 * @param {CommandInteraction} interaction - Discord interaction
 */
async function handleSheetCommand(interaction) {
  try {
    // Валідація прав доступу
    if (!hasPermission(interaction.member, "SHEETS_ACCESS")) {
      return await interaction.reply({
        content: "❌ Недостатньо прав для виконання цієї команди",
        ephemeral: true,
      });
    }

    // Логування
    logger.info(`Sheet command executed by ${interaction.user.tag}`);

    // Виконання логіки
    const result = await processSheetData(interaction.options);

    await interaction.reply({ content: result, ephemeral: false });
  } catch (error) {
    logger.error("Sheet command error:", error);
    await interaction.reply({
      content: "❌ Помилка при обробці команди",
      ephemeral: true,
    });
  }
}
```

### Робота з Google Sheets API:

```javascript
/**
 * Отримання даних з Google Sheets
 * @param {string} spreadsheetId - ID таблиці
 * @param {string} range - Діапазон даних
 * @returns {Promise<Array>} Дані з таблиці
 */
async function getSheetData(spreadsheetId, range) {
  try {
    const auth = await getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    return response.data.values || [];
  } catch (error) {
    logger.error("Google Sheets API error:", error);
    throw new Error("Помилка отримання даних з Google Sheets");
  }
}
```

### Інтеграція з LLM (Ollama):

```javascript
/**
 * AI-аналіз даних через Ollama
 * @param {Array} data - Дані для аналізу
 * @param {string} prompt - Запит до AI
 * @returns {Promise<string>} Результат аналізу
 */
async function analyzeWithOllama(data, prompt) {
  try {
    const response = await fetch("http://localhost:11434/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "llama2",
        prompt: `${prompt}\n\nДані: ${JSON.stringify(data)}`,
        stream: false,
      }),
    });

    const result = await response.json();
    return result.response;
  } catch (error) {
    logger.error("Ollama API error:", error);
    throw new Error("Помилка AI-аналізу");
  }
}
```

## 🔧 Налаштування середовища

### Змінні середовища (.env):

```env
# Discord Bot
DISCORD_TOKEN=your_discord_token
DISCORD_CLIENT_ID=your_client_id
DISCORD_GUILD_ID=your_guild_id

# Google Sheets
GOOGLE_SERVICE_ACCOUNT_EMAIL=your_service_account_email
GOOGLE_PRIVATE_KEY=your_private_key
GOOGLE_SPREADSHEET_ID=your_spreadsheet_id

# AI/LLM
OPENAI_API_KEY=your_openai_key
OLLAMA_HOST=http://localhost:11434

# Redis
REDIS_HOST=localhost
REDIS_PORT=6379

# Prometheus
PROMETHEUS_PORT=9090
```

## 📊 Моніторинг та метрики

Використовуйте Prometheus метрики для відстеження:

- Кількості команд
- Часу відповіді API
- Помилок та їх типів
- Використання ресурсів

## 🚀 Рекомендації для розробки

1. **Завжди** додавайте обробку помилок
2. **Логуйте** важливі події
3. **Валідуйте** вхідні дані
4. **Використовуйте** кешування для часто запитуваних даних
5. **Тестуйте** нові функції перед деплоєм
6. **Документуйте** зміни в коді

---

**💡 Порада:** Копіюйте ці інструкції в налаштування Cursor AI для кращого розуміння контексту проекту!
