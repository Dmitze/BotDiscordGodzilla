/**
 * Тестовий файл для валідації інтеграції всіх модулів
 * Запускається командою: npm run test:integration
 */

const logger = require('../../utils/logger');

// Тестові дані
const testData = {
  userId: '123456789',
  guildId: '987654321',
  userTag: 'TestUser#1234',
  commandName: 'test',
  query: 'тестовий запит',
  fileId: 'test_file_id',
  fileName: 'test_document.pdf',
};

/**
 * Тестування модуля безпеки
 */
async function testSecurityModule() {
  console.log('🔒 Тестування модуля безпеки...');

  try {
    const security = require('../../utils/security');

    // Тест валідації вхідних даних
    const sanitizedInput = security.sanitizeInput('<script>alert("test")</script>', 'search');
    console.log('✅ Валідація вхідних даних:', sanitizedInput);

    // Тест валідації параметрів команди
    const validationSchema = {
      query: { required: true, type: 'string', maxLength: 200, sanitize: 'search' },
      limit: { required: false, type: 'number', min: 1, max: 100 },
    };

    const testOptions = { query: 'test query', limit: 50 };
    const validation = security.validateCommandOptions(testOptions, validationSchema);
    console.log('✅ Валідація параметрів команди:', validation.isValid);

    // Тест rate limiting
    const isLimited = await security.checkRateLimit(testData.userId, 'SEARCH');
    console.log('✅ Rate limiting:', isLimited ? 'Ліміт перевищено' : 'Ліміт не перевищено');

    console.log('✅ Модуль безпеки працює коректно\n');
    return true;
  } catch (error) {
    console.error('❌ Помилка в модулі безпеки:', error.message);
    return false;
  }
}

/**
 * Тестування AI модуля
 */
async function testAIModule() {
  console.log('🤖 Тестування AI модуля...');

  try {
    const { aiEnhanced } = require('../../utils/aiEnhanced');

    // Тест аналізу природномовного запиту
    const analysis = await aiEnhanced.analyzeNaturalLanguage('знайди товари iPhone');
    console.log('✅ Аналіз природномовного запиту:', analysis.action);

    // Тест отримання контексту
    const context = aiEnhanced.getConversationContext(testData.userId);
    console.log('✅ Контекст розмови:', context.length, 'повідомлень');

    // Тест збереження в контекст
    aiEnhanced.saveToContext(testData.userId, 'user', 'тестове повідомлення');
    const updatedContext = aiEnhanced.getConversationContext(testData.userId);
    console.log('✅ Збереження в контекст:', updatedContext.length, 'повідомлень');

    // Тест статистики
    const stats = aiEnhanced.getStats();
    console.log('✅ Статистика AI:', stats.activeConversations, 'активних розмов');

    console.log('✅ AI модуль працює коректно\n');
    return true;
  } catch (error) {
    console.error('❌ Помилка в AI модулі:', error.message);
    return false;
  }
}

/**
 * Тестування модуля роботи з файлами
 */
async function testFileProcessorModule() {
  console.log('📁 Тестування модуля роботи з файлами...');

  try {
    const { fileProcessor } = require('../../utils/fileProcessor');

    // Тест створення звіту
    const reportData = {
      title: 'Тестовий звіт',
      content: 'Це тестовий зміст звіту для перевірки функціональності.',
      timestamp: new Date().toISOString(),
    };

    const reportPath = await fileProcessor.createReport(reportData, 'txt');
    console.log('✅ Створення звіту:', reportPath ? 'Успішно' : 'Помилка');

    // Тест очищення тимчасових файлів
    if (reportPath) {
      await fileProcessor.cleanupTempFile(reportPath);
      console.log('✅ Очищення тимчасових файлів: Успішно');
    }

    console.log('✅ Модуль роботи з файлами працює коректно\n');
    return true;
  } catch (error) {
    console.error('❌ Помилка в модулі роботи з файлами:', error.message);
    return false;
  }
}

/**
 * Тестування конфігурації
 */
function testConfiguration() {
  console.log('⚙️ Тестування конфігурації...');

  try {
    const config = require('../../config/Config');

    // Тест валідації конфігурації
    const isValid = config.validate();
    console.log('✅ Валідація конфігурації:', isValid);

    // Тест отримання налаштувань
    const discordConfig = config.getModuleConfig('discord');
    console.log('✅ Discord конфігурація:', !!discordConfig.token);

    const googleConfig = config.getModuleConfig('google');
    console.log('✅ Google конфігурація:', !!googleConfig.spreadsheetId);

    const aiConfig = config.getModuleConfig('ai');
    console.log('✅ AI конфігурація:', aiConfig.provider);

    console.log('✅ Конфігурація працює коректно\n');
    return true;
  } catch (error) {
    console.error('❌ Помилка в конфігурації:', error.message);
    return false;
  }
}

/**
 * Тестування логування
 */
function testLogging() {
  console.log('📝 Тестування логування...');

  try {
    logger.info('Тестове інформаційне повідомлення');
    logger.warn('Тестове попередження');
    logger.error('Тестова помилка');

    console.log('✅ Логування працює коректно\n');
    return true;
  } catch (error) {
    console.error('❌ Помилка в логуванні:', error.message);
    return false;
  }
}

/**
 * Тестування метрик
 */
function testMetrics() {
  console.log('📊 Тестування метрик...');

  try {
    const prometheus = require('./metrics/prometheus');

    // Тест реєстрації метрик
    prometheus.recordCommandExecution('test_command', 100);
    prometheus.recordSearchQuery('test_query', 5);
    prometheus.recordAIRequest('test_ai_request', 2000);

    console.log('✅ Метрики працюють коректно\n');
    return true;
  } catch (error) {
    console.error('❌ Помилка в метриках:', error.message);
    return false;
  }
}

/**
 * Головна функція тестування
 */
async function runIntegrationTests() {
  console.log('🚀 Запуск інтеграційних тестів...\n');

  const results = {
    security: false,
    ai: false,
    files: false,
    config: false,
    logging: false,
    metrics: false,
  };

  // Запуск тестів
  results.config = testConfiguration();
  results.logging = testLogging();
  results.metrics = testMetrics();
  results.security = await testSecurityModule();
  results.ai = await testAIModule();
  results.files = await testFileProcessorModule();

  // Підсумок
  console.log('📋 ПІДСУМОК ТЕСТУВАННЯ:');
  console.log('========================');

  Object.entries(results).forEach(([module, result]) => {
    const status = result ? '✅' : '❌';
    console.log(`${status} ${module.toUpperCase()}: ${result ? 'ПАСЕ' : 'ПОМИЛКА'}`);
  });

  const passedTests = Object.values(results).filter(Boolean).length;
  const totalTests = Object.keys(results).length;

  console.log(`\n📊 Результат: ${passedTests}/${totalTests} тестів пройшли успішно`);

  if (passedTests === totalTests) {
    console.log('🎉 Всі тести пройшли успішно! Система готова до роботи.');
    process.exit(0);
  } else {
    console.log('⚠️ Деякі тести не пройшли. Перевірте налаштування та спробуйте ще раз.');
    process.exit(1);
  }
}

// Запуск тестів якщо файл викликається безпосередньо
if (require.main === module) {
  runIntegrationTests().catch(error => {
    console.error('❌ Критична помилка при тестуванні:', error);
    process.exit(1);
  });
}

module.exports = {
  runIntegrationTests,
  testSecurityModule,
  testAIModule,
  testFileProcessorModule,
  testConfiguration,
  testLogging,
  testMetrics,
};
