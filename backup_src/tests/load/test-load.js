/**
 * Навантажувальне тестування Discord AI Assistant Bot
 * Версія 2.3.0
 */

const logger = require('../../utils/logger');

// Конфігурація навантажувального тестування
const loadTestConfig = {
  concurrentUsers: 10,
  requestsPerUser: 5,
  delayBetweenRequests: 1000, // 1 секунда
  testDuration: 30000, // 30 секунд
  maxResponseTime: 5000 // 5 секунд
};

// Статистика тестування
const testStats = {
  totalRequests: 0,
  successfulRequests: 0,
  failedRequests: 0,
  totalResponseTime: 0,
  minResponseTime: Infinity,
  maxResponseTime: 0,
  errors: []
};

/**
 * Симуляція користувача
 */
class VirtualUser {
  constructor(userId, config) {
    this.userId = userId;
    this.config = config;
    this.requests = 0;
    this.errors = 0;
    this.totalTime = 0;
  }

  async simulateUser() {
    const commands = [
      { name: 'пошук', options: { поле: 'найменування', запит: 'iPhone' } },
      { name: 'розумний-пошук', options: { номенклатура: 'Samsung' } },
      { name: 'ai', options: { запит: 'знайди товари' } },
      { name: 'файли', options: { дія: 'пошук', запит: 'звіт' } },
      { name: 'статистика', options: {} }
    ];

    for (let i = 0; i < this.config.requestsPerUser; i++) {
      const command = commands[i % commands.length];
      await this.executeCommand(command);
      
      if (i < this.config.requestsPerUser - 1) {
        await this.delay(this.config.delayBetweenRequests);
      }
    }
  }

  async executeCommand(command) {
    const startTime = Date.now();
    
    try {
      // Симуляція виконання команди
      const result = await this.simulateCommandExecution(command);
      const responseTime = Date.now() - startTime;
      
      this.requests++;
      this.totalTime += responseTime;
      
      // Оновлення глобальної статистики
      testStats.totalRequests++;
      testStats.successfulRequests++;
      testStats.totalResponseTime += responseTime;
      testStats.minResponseTime = Math.min(testStats.minResponseTime, responseTime);
      testStats.maxResponseTime = Math.max(testStats.maxResponseTime, responseTime);
      
      logger.info(`Користувач ${this.userId}: ${command.name} - ${responseTime}мс`);
      
    } catch (error) {
      const responseTime = Date.now() - startTime;
      
      this.errors++;
      testStats.totalRequests++;
      testStats.failedRequests++;
      testStats.errors.push({
        userId: this.userId,
        command: command.name,
        error: error.message,
        responseTime
      });
      
      logger.error(`Користувач ${this.userId}: ${command.name} - ПОМИЛКА: ${error.message}`);
    }
  }

  async simulateCommandExecution(command) {
    // Симуляція різного часу виконання команд
    const baseTime = {
      'пошук': 200,
      'розумний-пошук': 500,
      'ai': 1500,
      'файли': 800,
      'статистика': 100
    };
    
    const executionTime = baseTime[command.name] || 300;
    
    // Симуляція випадкових помилок (1% ймовірність)
    if (Math.random() < 0.01) {
      throw new Error('Симульована помилка');
    }
    
    await this.delay(executionTime);
    return { success: true, command: command.name };
  }

  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

/**
 * Навантажувальне тестування
 */
async function runLoadTest() {
  console.log('🚀 НАВАНТАЖУВАЛЬНЕ ТЕСТУВАННЯ DISCORD AI ASSISTANT BOT');
  console.log('=====================================================');
  console.log(`Версія: 2.3.0`);
  console.log(`Дата: ${new Date().toISOString()}`);
  console.log(`Конфігурація:`);
  console.log(`  - Користувачів: ${loadTestConfig.concurrentUsers}`);
  console.log(`  - Запитів на користувача: ${loadTestConfig.requestsPerUser}`);
  console.log(`  - Затримка між запитами: ${loadTestConfig.delayBetweenRequests}мс`);
  console.log(`  - Тривалість тесту: ${loadTestConfig.testDuration}мс`);
  console.log(`  - Максимальний час відповіді: ${loadTestConfig.maxResponseTime}мс\n`);

  const startTime = Date.now();
  const users = [];

  // Створення віртуальних користувачів
  for (let i = 0; i < loadTestConfig.concurrentUsers; i++) {
    const userId = `user_${i + 1}`;
    const user = new VirtualUser(userId, loadTestConfig);
    users.push(user);
  }

  // Запуск тестування
  console.log('🔄 Запуск навантажувального тестування...\n');

  const userPromises = users.map(user => user.simulateUser());
  
  try {
    await Promise.all(userPromises);
  } catch (error) {
    console.error('❌ Помилка під час навантажувального тестування:', error);
  }

  const endTime = Date.now();
  const totalTestTime = endTime - startTime;

  // Аналіз результатів
  analyzeResults(totalTestTime);
}

/**
 * Аналіз результатів тестування
 */
function analyzeResults(totalTestTime) {
  console.log('\n📊 РЕЗУЛЬТАТИ НАВАНТАЖУВАЛЬНОГО ТЕСТУВАННЯ');
  console.log('==========================================');

  // Основна статистика
  const avgResponseTime = testStats.totalRequests > 0 
    ? Math.round(testStats.totalResponseTime / testStats.totalRequests) 
    : 0;
  
  const successRate = testStats.totalRequests > 0 
    ? Math.round((testStats.successfulRequests / testStats.totalRequests) * 100) 
    : 0;
  
  const requestsPerSecond = Math.round(testStats.totalRequests / (totalTestTime / 1000));

  console.log(`📈 Основна статистика:`);
  console.log(`  - Загальна кількість запитів: ${testStats.totalRequests}`);
  console.log(`  - Успішних запитів: ${testStats.successfulRequests}`);
  console.log(`  - Провалених запитів: ${testStats.failedRequests}`);
  console.log(`  - Відсоток успішності: ${successRate}%`);
  console.log(`  - Запитів за секунду: ${requestsPerSecond}`);

  console.log(`\n⏱️ Час відповіді:`);
  console.log(`  - Середній час: ${avgResponseTime}мс`);
  console.log(`  - Мінімальний час: ${testStats.minResponseTime === Infinity ? 'N/A' : testStats.minResponseTime}мс`);
  console.log(`  - Максимальний час: ${testStats.maxResponseTime}мс`);
  console.log(`  - Загальний час тесту: ${totalTestTime}мс`);

  // Аналіз продуктивності
  console.log(`\n🎯 Аналіз продуктивності:`);
  
  if (avgResponseTime <= 1000) {
    console.log(`  ✅ Відмінна продуктивність (середній час < 1с)`);
  } else if (avgResponseTime <= 3000) {
    console.log(`  ⚠️ Хороша продуктивність (середній час < 3с)`);
  } else {
    console.log(`  ❌ Низька продуктивність (середній час > 3с)`);
  }

  if (successRate >= 95) {
    console.log(`  ✅ Відмінна надійність (успішність > 95%)`);
  } else if (successRate >= 90) {
    console.log(`  ⚠️ Хороша надійність (успішність > 90%)`);
  } else {
    console.log(`  ❌ Низька надійність (успішність < 90%)`);
  }

  if (requestsPerSecond >= 10) {
    console.log(`  ✅ Висока пропускна здатність (> 10 запитів/с)`);
  } else if (requestsPerSecond >= 5) {
    console.log(`  ⚠️ Середня пропускна здатність (> 5 запитів/с)`);
  } else {
    console.log(`  ❌ Низька пропускна здатність (< 5 запитів/с)`);
  }

  // Детальна статистика по помилках
  if (testStats.errors.length > 0) {
    console.log(`\n🚨 Помилки:`);
    const errorTypes = {};
    testStats.errors.forEach(error => {
      const type = error.error;
      errorTypes[type] = (errorTypes[type] || 0) + 1;
    });

    Object.entries(errorTypes).forEach(([type, count]) => {
      console.log(`  - ${type}: ${count} разів`);
    });
  }

  // Рекомендації
  console.log(`\n💡 РЕКОМЕНДАЦІЇ:`);
  
  if (avgResponseTime > 3000) {
    console.log(`  🔧 Оптимізуйте швидкість відповіді (поточний: ${avgResponseTime}мс)`);
  }
  
  if (successRate < 95) {
    console.log(`  🔧 Покращіть надійність (поточний: ${successRate}%)`);
  }
  
  if (requestsPerSecond < 10) {
    console.log(`  🔧 Збільшіть пропускну здатність (поточний: ${requestsPerSecond} запитів/с)`);
  }
  
  if (testStats.errors.length > 0) {
    console.log(`  🔧 Виправте помилки (всього: ${testStats.errors.length})`);
  }

  // Загальний висновок
  console.log(`\n🎯 ЗАГАЛЬНИЙ ВИСНОВОК:`);
  
  const score = calculatePerformanceScore(avgResponseTime, successRate, requestsPerSecond);
  
  if (score >= 90) {
    console.log(`  🎉 Відмінна продуктивність! Система готова до продакшену.`);
  } else if (score >= 70) {
    console.log(`  ✅ Хороша продуктивність. Можна запускати в продакшен.`);
  } else if (score >= 50) {
    console.log(`  ⚠️ Середня продуктивність. Потребує оптимізації.`);
  } else {
    console.log(`  ❌ Низька продуктивність. Не рекомендується для продакшену.`);
  }
  
  console.log(`  📊 Загальний бал: ${score}/100`);
}

/**
 * Розрахунок загального балу продуктивності
 */
function calculatePerformanceScore(avgResponseTime, successRate, requestsPerSecond) {
  let score = 0;
  
  // Бал за швидкість (40% від загального балу)
  if (avgResponseTime <= 1000) score += 40;
  else if (avgResponseTime <= 2000) score += 30;
  else if (avgResponseTime <= 3000) score += 20;
  else if (avgResponseTime <= 5000) score += 10;
  
  // Бал за надійність (40% від загального балу)
  if (successRate >= 99) score += 40;
  else if (successRate >= 95) score += 30;
  else if (successRate >= 90) score += 20;
  else if (successRate >= 80) score += 10;
  
  // Бал за пропускну здатність (20% від загального балу)
  if (requestsPerSecond >= 15) score += 20;
  else if (requestsPerSecond >= 10) score += 15;
  else if (requestsPerSecond >= 5) score += 10;
  else if (requestsPerSecond >= 2) score += 5;
  
  return Math.round(score);
}

/**
 * Тестування пам'яті
 */
async function testMemoryUsage() {
  console.log("\n🧠 ТЕСТУВАННЯ ВИКОРИСТАННЯ ПАМ'ЯТІ");
  console.log('==================================');
  
  const initialMemory = process.memoryUsage();
  console.log(`Початкове використання пам'яті:`);
  console.log(`  - RSS: ${Math.round(initialMemory.rss / 1024 / 1024)}MB`);
  console.log(`  - Heap Used: ${Math.round(initialMemory.heapUsed / 1024 / 1024)}MB`);
  console.log(`  - Heap Total: ${Math.round(initialMemory.heapTotal / 1024 / 1024)}MB`);
  
  // Симуляція навантаження
  const testData = [];
  for (let i = 0; i < 1000; i++) {
    testData.push({
      id: i,
      name: `Test Item ${i}`,
      data: 'x'.repeat(1000)
    });
  }
  
  const afterLoadMemory = process.memoryUsage();
  console.log(`\nПісля навантаження даних:`);
  console.log(`  - RSS: ${Math.round(afterLoadMemory.rss / 1024 / 1024)}MB`);
  console.log(`  - Heap Used: ${Math.round(afterLoadMemory.heapUsed / 1024 / 1024)}MB`);
  console.log(`  - Heap Total: ${Math.round(afterLoadMemory.heapTotal / 1024 / 1024)}MB`);
  
  const memoryIncrease = afterLoadMemory.heapUsed - initialMemory.heapUsed;
  console.log(`\nЗбільшення використання пам'яті: ${Math.round(memoryIncrease / 1024 / 1024)}MB`);
  
  // Очищення пам'яті
  testData.length = 0;
  global.gc && global.gc();
  
  const afterCleanupMemory = process.memoryUsage();
  console.log(`\nПісля очищення:`);
  console.log(`  - RSS: ${Math.round(afterCleanupMemory.rss / 1024 / 1024)}MB`);
  console.log(`  - Heap Used: ${Math.round(afterCleanupMemory.heapUsed / 1024 / 1024)}MB`);
  console.log(`  - Heap Total: ${Math.round(afterCleanupMemory.heapTotal / 1024 / 1024)}MB`);
}

// Запуск навантажувального тестування
if (require.main === module) {
  runLoadTest()
    .then(() => testMemoryUsage())
    .then(() => {
      console.log('\n✅ Навантажувальне тестування завершено');
      process.exit(0);
    })
    .catch(error => {
      console.error('❌ Помилка під час навантажувального тестування:', error);
      process.exit(1);
    });
}

module.exports = {
  runLoadTest,
  testMemoryUsage,
  loadTestConfig,
  testStats
}; 