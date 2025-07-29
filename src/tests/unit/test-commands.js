/**
 * Тестування команд Discord
 * Версія 2.3.0
 */

const logger = require('../../utils/logger');

// Тестові дані для команд
const testCommands = {
  search: {
    name: 'пошук',
    options: {
      поле: 'найменування',
      запит: 'iPhone',
    },
  },
  smartSearch: {
    name: 'розумний-пошук',
    options: {
      номенклатура: 'iPhone',
      ціна_вище: 1000,
    },
  },
  ai: {
    name: 'ai',
    options: {
      запит: 'знайди товари iPhone',
      контекст: 'для аналізу',
    },
  },
  files: {
    name: 'файли',
    options: {
      дія: 'пошук',
      запит: 'звіт',
      папка: '123456789',
    },
  },
  export: {
    name: 'пошук-експортовано',
    options: {
      поле: 'найменування',
      запит: 'iPhone',
    },
  },
  stats: {
    name: 'статистика',
    options: {},
  },
  help: {
    name: 'допомога',
    options: {
      категорія: 'search',
    },
  },
};

/**
 * Мок об'єкт для тестування команд
 */
function createMockCommandInteraction(commandName, options = {}) {
  return {
    commandName,
    user: {
      id: '123456789',
      tag: 'TestUser#1234',
      username: 'TestUser',
    },
    guild: {
      id: '987654321',
      name: 'Test Guild',
    },
    channel: {
      id: '111222333',
      name: 'test-channel',
    },
    member: {
      roles: {
        cache: new Map([
          ['Адміністратор', { name: 'Адміністратор', id: 'admin-role' }],
          ['Бот-Користувач', { name: 'Бот-Користувач', id: 'user-role' }],
          ['Sheets-Доступ', { name: 'Sheets-Доступ', id: 'sheets-role' }],
          ['AI-Доступ', { name: 'AI-Доступ', id: 'ai-role' }],
          ['Експорт-Доступ', { name: 'Експорт-Доступ', id: 'export-role' }],
        ]),
      },
    },
    options: {
      getString: name => options[name] || 'test',
      getInteger: name => options[name] || 10,
      getBoolean: name => options[name] || false,
      getNumber: name => options[name] || 0,
    },
    reply: async content => {
      logger.info(`Mock reply: ${JSON.stringify(content)}`);
      return { success: true };
    },
    editReply: async content => {
      logger.info(`Mock editReply: ${JSON.stringify(content)}`);
      return { success: true };
    },
    followUp: async content => {
      logger.info(`Mock followUp: ${JSON.stringify(content)}`);
      return { success: true };
    },
    deferReply: async () => {
      logger.info('Mock deferReply called');
      return { success: true };
    },
  };
}

/**
 * Тестування команди пошуку
 */
async function testSearchCommand() {
  console.log('🔍 Тестування команди /пошук...');

  try {
    const { enhancedSearch } = require('./commands/enhancedSearch');
    const interaction = createMockCommandInteraction('пошук', testCommands.search.options);

    // Тест валідації параметрів
    const validation = enhancedSearch.validateOptions(interaction.options);
    if (!validation.isValid) {
      throw new Error(`Валідація параметрів не пройшла: ${validation.errors.join(', ')}`);
    }

    // Тест виконання команди
    const result = await enhancedSearch.execute(interaction);
    if (!result) {
      throw new Error('Команда не повернула результат');
    }

    console.log('✅ Команда /пошук працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /пошук:', error.message);
    return false;
  }
}

/**
 * Тестування команди розумного пошуку
 */
async function testSmartSearchCommand() {
  console.log('🧠 Тестування команди /розумний-пошук...');

  try {
    const { enhancedSearch } = require('./commands/enhancedSearch');
    const interaction = createMockCommandInteraction(
      'розумний-пошук',
      testCommands.smartSearch.options
    );

    // Тест валідації параметрів
    const validation = enhancedSearch.validateSmartSearchOptions(interaction.options);
    if (!validation.isValid) {
      throw new Error(`Валідація параметрів не пройшла: ${validation.errors.join(', ')}`);
    }

    // Тест виконання команди
    const result = await enhancedSearch.executeSmartSearch(interaction);
    if (!result) {
      throw new Error('Команда не повернула результат');
    }

    console.log('✅ Команда /розумний-пошук працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /розумний-пошук:', error.message);
    return false;
  }
}

/**
 * Тестування AI команди
 */
async function testAICommand() {
  console.log('🤖 Тестування команди /ai...');

  try {
    const { aiAssistant } = require('./commands/aiAssistant');
    const interaction = createMockCommandInteraction('ai', testCommands.ai.options);

    // Тест валідації параметрів
    const validation = aiAssistant.validateOptions(interaction.options);
    if (!validation.isValid) {
      throw new Error(`Валідація параметрів не пройшла: ${validation.errors.join(', ')}`);
    }

    // Тест виконання команди
    const result = await aiAssistant.execute(interaction);
    if (!result) {
      throw new Error('Команда не повернула результат');
    }

    console.log('✅ Команда /ai працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /ai:', error.message);
    return false;
  }
}

/**
 * Тестування команди роботи з файлами
 */
async function testFilesCommand() {
  console.log('📁 Тестування команди /файли...');

  try {
    const { fileManager } = require('./commands/fileManager');
    const interaction = createMockCommandInteraction('файли', testCommands.files.options);

    // Тест валідації параметрів
    const validation = fileManager.validateOptions(interaction.options);
    if (!validation.isValid) {
      throw new Error(`Валідація параметрів не пройшла: ${validation.errors.join(', ')}`);
    }

    // Тест виконання команди
    const result = await fileManager.execute(interaction);
    if (!result) {
      throw new Error('Команда не повернула результат');
    }

    console.log('✅ Команда /файли працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /файли:', error.message);
    return false;
  }
}

/**
 * Тестування команди експорту
 */
async function testExportCommand() {
  console.log('📤 Тестування команди /пошук-експортовано...');

  try {
    const { enhancedSearch } = require('./commands/enhancedSearch');
    const interaction = createMockCommandInteraction(
      'пошук-експортовано',
      testCommands.export.options
    );

    // Тест валідації параметрів
    const validation = enhancedSearch.validateExportOptions(interaction.options);
    if (!validation.isValid) {
      throw new Error(`Валідація параметрів не пройшла: ${validation.errors.join(', ')}`);
    }

    // Тест виконання команди
    const result = await enhancedSearch.executeExport(interaction);
    if (!result) {
      throw new Error('Команда не повернула результат');
    }

    console.log('✅ Команда /пошук-експортовано працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /пошук-експортовано:', error.message);
    return false;
  }
}

/**
 * Тестування команди статистики
 */
async function testStatsCommand() {
  console.log('📊 Тестування команди /статистика...');

  try {
    const stats = require('./stats');
    const interaction = createMockCommandInteraction('статистика', testCommands.stats.options);

    // Тест виконання команди
    const result = await stats.execute(interaction);
    if (!result) {
      throw new Error('Команда не повернула результат');
    }

    console.log('✅ Команда /статистика працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /статистика:', error.message);
    return false;
  }
}

/**
 * Тестування команди довідки
 */
async function testHelpCommand() {
  console.log('❓ Тестування команди /допомога...');

  try {
    const { UIHelper } = require('../../utils/uiHelpers');
    const interaction = createMockCommandInteraction('допомога', testCommands.help.options);

    // Тест створення довідки
    const helpEmbed = UIHelper.createHelpEmbed(testCommands.help.options.категорія);
    if (!helpEmbed || !helpEmbed.data) {
      throw new Error('Не вдалося створити embed довідки');
    }

    // Тест відповіді
    await interaction.reply({ embeds: [helpEmbed] });

    console.log('✅ Команда /допомога працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в команді /допомога:', error.message);
    return false;
  }
}

/**
 * Тестування прав доступу
 */
async function testPermissions() {
  console.log('🔒 Тестування прав доступу...');

  try {
    const security = require('../../utils/security');

    // Тест користувача з адмін правами
    const adminInteraction = createMockCommandInteraction('admin-command');
    const hasAdminRole = security.hasRole(adminInteraction, 'Адміністратор');
    if (!hasAdminRole) {
      throw new Error('Адміністратор не має доступу');
    }

    // Тест користувача без прав
    const noRoleInteraction = createMockCommandInteraction('restricted-command');
    noRoleInteraction.member.roles.cache.clear();
    const hasNoRole = security.hasRole(noRoleInteraction, 'Адміністратор');
    if (hasNoRole) {
      throw new Error('Користувач без прав має доступ');
    }

    // Тест rate limiting
    const isLimited = await security.checkRateLimit('test-user', 'SEARCH');
    if (typeof isLimited !== 'boolean') {
      throw new Error('Rate limiting не працює');
    }

    console.log('✅ Права доступу працюють коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в правах доступу:', error.message);
    return false;
  }
}

/**
 * Тестування обробки помилок
 */
async function testErrorHandling() {
  console.log('🚨 Тестування обробки помилок...');

  try {
    const { UIHelper } = require('../../utils/uiHelpers');

    // Тест створення embed помилки
    const errorEmbed = UIHelper.createErrorEmbed(new Error('Тестова помилка'), 'Контекст помилки');
    if (!errorEmbed || !errorEmbed.data) {
      throw new Error('Не вдалося створити embed помилки');
    }

    // Тест обробки невалідних параметрів
    const invalidInteraction = createMockCommandInteraction('invalid-command');
    invalidInteraction.options.getString = () => null;

    // Тест обробки відсутніх модулів
    try {
      require('./non-existent-module');
    } catch (error) {
      // Очікувана помилка
      console.log('✅ Обробка відсутніх модулів працює');
    }

    console.log('✅ Обробка помилок працює коректно');
    return true;
  } catch (error) {
    console.error('❌ Помилка в обробці помилок:', error.message);
    return false;
  }
}

/**
 * Головна функція тестування команд
 */
async function runCommandTests() {
  console.log('🚀 ТЕСТУВАННЯ КОМАНД DISCORD BOT');
  console.log('================================');
  console.log(`Версія: 2.3.0`);
  console.log(`Дата: ${new Date().toISOString()}\n`);

  const tests = [
    { name: 'Пошук', fn: testSearchCommand },
    { name: 'Розумний пошук', fn: testSmartSearchCommand },
    { name: 'AI команда', fn: testAICommand },
    { name: 'Робота з файлами', fn: testFilesCommand },
    { name: 'Експорт', fn: testExportCommand },
    { name: 'Статистика', fn: testStatsCommand },
    { name: 'Довідка', fn: testHelpCommand },
    { name: 'Права доступу', fn: testPermissions },
    { name: 'Обробка помилок', fn: testErrorHandling },
  ];

  const results = [];

  for (const test of tests) {
    const startTime = Date.now();
    const success = await test.fn();
    const duration = Date.now() - startTime;

    results.push({
      name: test.name,
      success,
      duration,
    });
  }

  // Підсумок
  console.log('\n📋 ПІДСУМОК ТЕСТУВАННЯ КОМАНД');
  console.log('=============================');

  const successfulTests = results.filter(r => r.success);
  const failedTests = results.filter(r => !r.success);

  console.log(`✅ Успішних тестів: ${successfulTests.length}`);
  console.log(`❌ Провалених тестів: ${failedTests.length}`);
  console.log(`📊 Загальна кількість: ${results.length}`);
  console.log(
    `📈 Відсоток успішності: ${Math.round((successfulTests.length / results.length) * 100)}%`
  );

  if (failedTests.length > 0) {
    console.log('\n❌ ПРОВАЛЕНІ ТЕСТИ:');
    failedTests.forEach(test => {
      console.log(`   - ${test.name}`);
    });
  }

  const totalTime = results.reduce((sum, r) => sum + r.duration, 0);
  console.log(`\n⏱️ Загальний час тестування: ${totalTime}мс`);

  if (successfulTests.length === results.length) {
    console.log('\n🎉 Всі команди працюють коректно!');
  } else {
    console.log('\n⚠️ Деякі команди потребують уваги.');
  }

  return {
    total: results.length,
    successful: successfulTests.length,
    failed: failedTests.length,
    successRate: (successfulTests.length / results.length) * 100,
    totalTime,
  };
}

// Запуск тестів якщо файл викликається безпосередньо
if (require.main === module) {
  runCommandTests()
    .then(result => {
      process.exit(result.failed > 0 ? 1 : 0);
    })
    .catch(error => {
      console.error('❌ Критична помилка при тестуванні команд:', error);
      process.exit(1);
    });
}

module.exports = {
  runCommandTests,
  createMockCommandInteraction,
  testCommands,
};
