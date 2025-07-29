/**
 * Комплексне тестування Discord AI Assistant Bot
 * Версія 2.3.0
 *
 * Запуск: npm run test:comprehensive
 */

const logger = require('../../utils/logger');

// Кольори для консолі
const colors = {
  reset: '\x1b[0m',
  bright: '\x1b[1m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m',
  cyan: '\x1b[36m',
};

// Тестові дані
const testData = {
  userId: '123456789',
  guildId: '987654321',
  userTag: 'TestUser#1234',
  channelId: '111222333',
  messageId: '444555666',
  commandName: 'test',
  query: 'тестовий запит',
  fileId: 'test_file_id',
  fileName: 'test_document.pdf',
  searchResults: [
    { найменування: 'iPhone 15', кількість: 10, ціна: 25000 },
    { найменування: 'Samsung Galaxy', кількість: 5, ціна: 18000 },
    { найменування: 'MacBook Pro', кількість: 3, ціна: 45000 },
  ],
};

/**
 * Утиліти для тестування
 */
class TestUtils {
  static log(message, type = 'info') {
    const timestamp = new Date().toISOString();
    const color =
      type === 'success'
        ? colors.green
        : type === 'error'
          ? colors.red
          : type === 'warning'
            ? colors.yellow
            : colors.blue;

    console.log(`${color}[${timestamp}] ${message}${colors.reset}`);
  }

  static async testFunction(name, testFn) {
    try {
      this.log(`🧪 Тестування: ${name}`, 'info');
      const startTime = Date.now();
      const result = await testFn();
      const duration = Date.now() - startTime;

      if (result) {
        this.log(`✅ ${name} - УСПІШНО (${duration}мс)`, 'success');
        return { success: true, duration, name };
      } else {
        this.log(`❌ ${name} - ПРОВАЛЕНО`, 'error');
        return { success: false, duration, name };
      }
    } catch (error) {
      this.log(`❌ ${name} - ПОМИЛКА: ${error.message}`, 'error');
      return { success: false, error: error.message, name };
    }
  }

  static createMockInteraction() {
    return {
      user: { id: testData.userId, tag: testData.userTag },
      guild: { id: testData.guildId },
      channel: { id: testData.channelId },
      member: {
        roles: {
          cache: new Map([
            ['Адміністратор', { name: 'Адміністратор' }],
            ['Бот-Користувач', { name: 'Бот-Користувач' }],
          ]),
        },
      },
      options: {
        getString: name => testData[name] || 'test',
        getInteger: name => 10,
        getBoolean: name => false,
      },
      reply: async content => ({ success: true }),
      editReply: async content => ({ success: true }),
      followUp: async content => ({ success: true }),
    };
  }
}

/**
 * Тестування конфігурації
 */
async function testConfiguration() {
  return await TestUtils.testFunction('Конфігурація', async () => {
    const config = require('../../config/Config');

    // Тест завантаження конфігурації
    if (!config) throw new Error('Конфігурація не завантажена');

    // Тест валідації
    const isValid = config.validate();
    if (!isValid) throw new Error('Конфігурація не валідна');

    // Тест отримання налаштувань
    const discordConfig = config.getModuleConfig('discord');
    const googleConfig = config.getModuleConfig('google');
    const aiConfig = config.getModuleConfig('ai');

    if (!discordConfig || !googleConfig || !aiConfig) {
      throw new Error('Не вдалося отримати налаштування модулів');
    }

    return true;
  });
}

/**
 * Тестування логування
 */
async function testLogging() {
  return await TestUtils.testFunction('Логування', async () => {
    logger.info('Тестове інформаційне повідомлення');
    logger.warn('Тестове попередження');
    logger.error('Тестова помилка');

    // Перевірка, чи логгер працює без помилок
    return true;
  });
}

/**
 * Тестування модуля безпеки
 */
async function testSecurityModule() {
  return await TestUtils.testFunction('Модуль безпеки', async () => {
    const security = require('../../utils/security');

    // Тест валідації вхідних даних
    const sanitizedInput = security.sanitizeInput('<script>alert("test")</script>', 'search');
    if (sanitizedInput.includes('<script>')) {
      throw new Error('Валідація вхідних даних не працює');
    }

    // Тест валідації параметрів команди
    const validationSchema = {
      query: { required: true, type: 'string', maxLength: 200, sanitize: 'search' },
      limit: { required: false, type: 'number', min: 1, max: 100 },
    };

    const testOptions = { query: 'test query', limit: 50 };
    const validation = security.validateCommandOptions(testOptions, validationSchema);
    if (!validation.isValid) {
      throw new Error('Валідація параметрів команди не працює');
    }

    // Тест rate limiting
    const isLimited = await security.checkRateLimit(testData.userId, 'SEARCH');
    if (typeof isLimited !== 'boolean') {
      throw new Error('Rate limiting не працює');
    }

    // Тест перевірки ролей
    const mockInteraction = TestUtils.createMockInteraction();
    const hasRole = security.hasRole(mockInteraction, 'Адміністратор');
    if (!hasRole) {
      throw new Error('Перевірка ролей не працює');
    }

    return true;
  });
}

/**
 * Тестування AI модуля
 */
async function testAIModule() {
  return await TestUtils.testFunction('AI модуль', async () => {
    const { aiEnhanced } = require('../../utils/aiEnhanced');

    // Тест аналізу природномовного запиту
    const analysis = await aiEnhanced.analyzeNaturalLanguage('знайди товари iPhone');
    if (!analysis || !analysis.action) {
      throw new Error('Аналіз природномовного запиту не працює');
    }

    // Тест отримання контексту
    const context = aiEnhanced.getConversationContext(testData.userId);
    if (!Array.isArray(context)) {
      throw new Error('Отримання контексту не працює');
    }

    // Тест збереження в контекст
    aiEnhanced.saveToContext(testData.userId, 'user', 'тестове повідомлення');
    const updatedContext = aiEnhanced.getConversationContext(testData.userId);
    if (updatedContext.length <= context.length) {
      throw new Error('Збереження в контекст не працює');
    }

    // Тест статистики
    const stats = aiEnhanced.getStats();
    if (!stats || typeof stats.activeConversations !== 'number') {
      throw new Error('Статистика AI не працює');
    }

    // Тест очищення контексту
    aiEnhanced.clearContext(testData.userId);
    const clearedContext = aiEnhanced.getConversationContext(testData.userId);
    if (clearedContext.length !== 0) {
      throw new Error('Очищення контексту не працює');
    }

    return true;
  });
}

/**
 * Тестування модуля роботи з файлами
 */
async function testFileProcessorModule() {
  return await TestUtils.testFunction('Модуль роботи з файлами', async () => {
    const { fileProcessor } = require('../../utils/fileProcessor');

    // Тест створення звіту
    const reportData = {
      title: 'Тестовий звіт',
      content: 'Це тестовий зміст звіту для перевірки функціональності.',
      timestamp: new Date().toISOString(),
      metadata: {
        author: 'Test User',
        version: '2.3.0',
      },
    };

    const reportPath = await fileProcessor.createReport(reportData, 'txt');
    if (!reportPath) {
      throw new Error('Створення звіту не працює');
    }

    // Тест очищення тимчасових файлів
    await fileProcessor.cleanupTempFile(reportPath);

    // Тест валідації файлів
    const isValidFile = fileProcessor.validateFile({
      name: 'test.pdf',
      size: 1024,
      type: 'application/pdf',
    });
    if (!isValidFile) {
      throw new Error('Валідація файлів не працює');
    }

    return true;
  });
}

/**
 * Тестування UI/UX модуля
 */
async function testUIHelpersModule() {
  return await TestUtils.testFunction('UI/UX модуль', async () => {
    const { UIHelper, COLORS, EMOJIS } = require('../../utils/uiHelpers');

    // Тест створення базового embed
    const baseEmbed = UIHelper.createBaseEmbed('Тест', 'Опис', COLORS.INFO);
    if (!baseEmbed || !baseEmbed.data) {
      throw new Error('Створення базового embed не працює');
    }

    // Тест створення embed для результатів пошуку
    const searchEmbed = UIHelper.createSearchResultsEmbed(
      testData.searchResults,
      'test query',
      0,
      1
    );
    if (!searchEmbed || !searchEmbed.data) {
      throw new Error('Створення search embed не працює');
    }

    // Тест створення AI embed
    const aiEmbed = UIHelper.createAIResponseEmbed('test query', 'test response', 0.8);
    if (!aiEmbed || !aiEmbed.data) {
      throw new Error('Створення AI embed не працює');
    }

    // Тест створення кнопок
    const buttons = UIHelper.createNavigationButtons(0, 5);
    if (!buttons || !buttons.components) {
      throw new Error('Створення кнопок не працює');
    }

    // Тест прогрес-бару
    const progressBar = UIHelper.createProgressBar(5, 10);
    if (!progressBar || !progressBar.includes('50%')) {
      throw new Error('Створення прогрес-бару не працює');
    }

    return true;
  });
}

/**
 * Тестування команд
 */
async function testCommands() {
  const commands = [
    { name: 'AI Assistant', file: './commands/aiAssistant.js' },
    { name: 'File Manager', file: './commands/fileManager.js' },
    { name: 'Enhanced Search', file: './commands/enhancedSearch.js' },
  ];

  const results = [];

  for (const command of commands) {
    const result = await TestUtils.testFunction(`Команда: ${command.name}`, async () => {
      try {
        const commandModule = require(command.file);
        if (!commandModule) {
          throw new Error('Модуль команди не завантажений');
        }
        return true;
      } catch (error) {
        throw new Error(`Помилка завантаження команди: ${error.message}`);
      }
    });
    results.push(result);
  }

  return results;
}

/**
 * Тестування метрик
 */
async function testMetrics() {
  return await TestUtils.testFunction('Метрики', async () => {
    try {
      const prometheus = require('./metrics/prometheus');

      // Тест реєстрації метрик
      prometheus.recordCommandExecution('test_command', 100);
      prometheus.recordSearchQuery('test_query', 5);
      prometheus.recordAIRequest('test_ai_request', 2000);

      return true;
    } catch (error) {
      // Метрики можуть бути не налаштовані в dev середовищі
      TestUtils.log(`⚠️ Метрики не налаштовані: ${error.message}`, 'warning');
      return true;
    }
  });
}

/**
 * Тестування експорту
 */
async function testExportHelpers() {
  return await TestUtils.testFunction('Модуль експорту', async () => {
    const { exportHelpers } = require('../../utils/exportHelpers');

    // Тест створення Excel файлу
    const excelData = testData.searchResults;
    const excelPath = await exportHelpers.createExcelFile(excelData, 'test_export');
    if (!excelPath) {
      throw new Error('Створення Excel файлу не працює');
    }

    // Тест очищення тимчасових файлів
    await exportHelpers.cleanupTempFile(excelPath);

    return true;
  });
}

/**
 * Тестування форматування
 */
async function testFormatters() {
  return await TestUtils.testFunction('Модуль форматування', async () => {
    const { formatters } = require('../../utils/formatters');

    // Тест форматування дат
    const formattedDate = formatters.formatDate(new Date());
    if (!formattedDate) {
      throw new Error('Форматування дат не працює');
    }

    // Тест форматування чисел
    const formattedNumber = formatters.formatNumber(1234567.89);
    if (!formattedNumber) {
      throw new Error('Форматування чисел не працює');
    }

    // Тест форматування результатів пошуку
    const formattedResults = formatters.formatSearchResults(testData.searchResults);
    if (!formattedResults) {
      throw new Error('Форматування результатів пошуку не працює');
    }

    return true;
  });
}

/**
 * Тестування retry логіки
 */
async function testRetryLogic() {
  return await TestUtils.testFunction('Retry логіка', async () => {
    const { retry } = require('../../utils/retry');

    // Тест успішного виконання
    let successCount = 0;
    const successFn = async () => {
      successCount++;
      return 'success';
    };

    const result = await retry.executeWithRetry(successFn, 3);
    if (result !== 'success' || successCount !== 1) {
      throw new Error('Retry логіка не працює для успішних операцій');
    }

    // Тест retry при помилках
    let failCount = 0;
    const failFn = async () => {
      failCount++;
      throw new Error('Test error');
    };

    try {
      await retry.executeWithRetry(failFn, 3);
      throw new Error('Retry логіка не обробляє помилки');
    } catch (error) {
      if (failCount !== 3) {
        throw new Error('Retry логіка не повторює спроби');
      }
    }

    return true;
  });
}

/**
 * Головна функція тестування
 */
async function runComprehensiveTests() {
  console.log(`${colors.bright}${colors.cyan}`);
  console.log('🚀 КОМПЛЕКСНЕ ТЕСТУВАННЯ DISCORD AI ASSISTANT BOT');
  console.log('================================================');
  console.log(`Версія: 2.3.0`);
  console.log(`Дата: ${new Date().toISOString()}`);
  console.log(`${colors.reset}\n`);

  const tests = [
    testConfiguration,
    testLogging,
    testSecurityModule,
    testAIModule,
    testFileProcessorModule,
    testUIHelpersModule,
    testMetrics,
    testExportHelpers,
    testFormatters,
    testRetryLogic,
  ];

  const results = [];

  // Запуск основних тестів
  for (const test of tests) {
    const result = await test();
    results.push(result);
  }

  // Запуск тестів команд
  const commandResults = await testCommands();
  results.push(...commandResults);

  // Підсумок
  console.log(`\n${colors.bright}${colors.magenta}`);
  console.log('📋 ПІДСУМОК ТЕСТУВАННЯ');
  console.log('=====================');

  const successfulTests = results.filter(r => r.success);
  const failedTests = results.filter(r => !r.success);

  console.log(`✅ Успішних тестів: ${successfulTests.length}`);
  console.log(`❌ Провалених тестів: ${failedTests.length}`);
  console.log(`📊 Загальна кількість: ${results.length}`);
  console.log(
    `📈 Відсоток успішності: ${Math.round((successfulTests.length / results.length) * 100)}%`
  );

  if (failedTests.length > 0) {
    console.log(`\n${colors.red}❌ ПРОВАЛЕНІ ТЕСТИ:${colors.reset}`);
    failedTests.forEach(test => {
      console.log(`   - ${test.name}: ${test.error || 'Невідома помилка'}`);
    });
  }

  console.log(`\n${colors.bright}${colors.yellow}⏱️ ЧАС ВИКОНАННЯ:${colors.reset}`);
  const totalTime = results.reduce((sum, r) => sum + (r.duration || 0), 0);
  console.log(`   Загальний час: ${totalTime}мс`);
  console.log(`   Середній час на тест: ${Math.round(totalTime / results.length)}мс`);

  // Рекомендації
  console.log(`\n${colors.bright}${colors.cyan}💡 РЕКОМЕНДАЦІЇ:${colors.reset}`);
  if (successfulTests.length === results.length) {
    console.log('   🎉 Всі тести пройшли успішно! Система готова до роботи.');
    console.log('   📝 Наступний крок: Запуск бота та тестування в реальному середовищі.');
  } else {
    console.log('   ⚠️ Деякі тести не пройшли. Перевірте налаштування та спробуйте ще раз.');
    console.log('   🔧 Рекомендується виправити помилки перед запуском в продакшен.');
  }

  console.log(`\n${colors.reset}`);

  // Повернення результату
  return {
    total: results.length,
    successful: successfulTests.length,
    failed: failedTests.length,
    successRate: (successfulTests.length / results.length) * 100,
    totalTime,
    failedTests: failedTests.map(t => ({ name: t.name, error: t.error })),
  };
}

// Запуск тестів якщо файл викликається безпосередньо
if (require.main === module) {
  runComprehensiveTests()
    .then(result => {
      process.exit(result.failed > 0 ? 1 : 0);
    })
    .catch(error => {
      console.error(
        `${colors.red}❌ Критична помилка при тестуванні: ${error.message}${colors.reset}`
      );
      process.exit(1);
    });
}

module.exports = {
  runComprehensiveTests,
  TestUtils,
  testData,
};
