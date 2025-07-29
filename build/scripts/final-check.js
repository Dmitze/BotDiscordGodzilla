/**
 * ✅ Фінальна перевірка Discord AI Assistant Bot
 * Версія: 2.3.0
 */

const fs = require('fs-extra');
const path = require('path');
const { execSync } = require('child_process');

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

// Функції логування
function log(message, color = 'reset') {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

function logInfo(message) {
  log(`[INFO] ${message}`, 'blue');
}

function logSuccess(message) {
  log(`[SUCCESS] ${message}`, 'green');
}

function logWarning(message) {
  log(`[WARNING] ${message}`, 'yellow');
}

function logError(message) {
  log(`[ERROR] ${message}`, 'red');
}

// Результати перевірки
const results = {
  passed: 0,
  failed: 0,
  warnings: 0,
  total: 0,
};

// Функція перевірки файлу
function checkFile(filePath, description) {
  results.total++;

  if (fs.existsSync(filePath)) {
    logSuccess(`✓ ${description} (${filePath})`);
    results.passed++;
    return true;
  } else {
    logError(`✗ ${description} (${filePath}) - ФАЙЛ НЕ ЗНАЙДЕНО`);
    results.failed++;
    return false;
  }
}

// Функція перевірки директорії
function checkDirectory(dirPath, description) {
  results.total++;

  if (fs.existsSync(dirPath) && fs.statSync(dirPath).isDirectory()) {
    logSuccess(`✓ ${description} (${dirPath})`);
    results.passed++;
    return true;
  } else {
    logError(`✗ ${description} (${dirPath}) - ДИРЕКТОРІЯ НЕ ЗНАЙДЕНА`);
    results.failed++;
    return false;
  }
}

// Функція перевірки змінної середовища
function checkEnvironmentVariable(varName, description) {
  results.total++;

  if (process.env[varName]) {
    logSuccess(`✓ ${description} (${varName})`);
    results.passed++;
    return true;
  } else {
    logWarning(`⚠ ${description} (${varName}) - НЕ ВСТАНОВЛЕНО`);
    results.warnings++;
    return false;
  }
}

// Функція перевірки .env файлу
function checkEnvFile() {
  results.total++;

  const envPath = path.join(process.cwd(), '.env');

  if (fs.existsSync(envPath)) {
    const envContent = fs.readFileSync(envPath, 'utf8');
    const requiredVars = ['DISCORD_TOKEN', 'CLIENT_ID', 'GUILD_ID'];
    const missingVars = [];

    for (const varName of requiredVars) {
      if (!envContent.includes(`${varName}=`)) {
        missingVars.push(varName);
      }
    }

    if (missingVars.length === 0) {
      logSuccess(`✓ .env файл налаштований правильно`);
      results.passed++;
      return true;
    } else {
      logWarning(`⚠ .env файл існує, але відсутні змінні: ${missingVars.join(', ')}`);
      results.warnings++;
      return false;
    }
  } else {
    logError(`✗ .env файл не знайдено`);
    results.failed++;
    return false;
  }
}

// Функція перевірки npm залежностей
function checkNpmDependencies() {
  results.total++;

  try {
    const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
    const requiredDeps = ['discord.js', 'dotenv', 'googleapis', 'openai', 'winston'];

    const missingDeps = [];

    for (const dep of requiredDeps) {
      if (!packageJson.dependencies[dep] && !packageJson.devDependencies[dep]) {
        missingDeps.push(dep);
      }
    }

    if (missingDeps.length === 0) {
      logSuccess(`✓ Всі необхідні npm залежності встановлені`);
      results.passed++;
      return true;
    } else {
      logError(`✗ Відсутні npm залежності: ${missingDeps.join(', ')}`);
      results.failed++;
      return false;
    }
  } catch (error) {
    logError(`✗ Помилка перевірки package.json: ${error.message}`);
    results.failed++;
    return false;
  }
}

// Функція перевірки Node.js версії
function checkNodeVersion() {
  results.total++;

  try {
    const nodeVersion = process.version;
    const majorVersion = parseInt(nodeVersion.slice(1).split('.')[0]);

    if (majorVersion >= 18) {
      logSuccess(`✓ Node.js версія ${nodeVersion} підходить`);
      results.passed++;
      return true;
    } else {
      logError(`✗ Node.js версія ${nodeVersion} занадто стара. Потрібна версія 18+`);
      results.failed++;
      return false;
    }
  } catch (error) {
    logError(`✗ Помилка перевірки версії Node.js: ${error.message}`);
    results.failed++;
    return false;
  }
}

// Функція перевірки тестів
function runTests() {
  logInfo('Запуск тестів...');

  const tests = [
    { name: 'AI тести', command: 'npm test' },
    { name: 'Інтеграційні тести', command: 'node test-integration.js' },
    { name: 'Комплексні тести', command: 'node test-comprehensive.js' },
    { name: 'Тести команд', command: 'node test-commands.js' },
  ];

  let allTestsPassed = true;

  for (const test of tests) {
    results.total++;

    try {
      logInfo(`Запуск ${test.name}...`);
      execSync(test.command, { stdio: 'pipe', timeout: 30000 });
      logSuccess(`✓ ${test.name} пройшли успішно`);
      results.passed++;
    } catch (error) {
      logError(`✗ ${test.name} не пройшли: ${error.message}`);
      results.failed++;
      allTestsPassed = false;
    }
  }

  return allTestsPassed;
}

// Функція перевірки конфігурації
function checkConfiguration() {
  logInfo('Перевірка конфігурації...');

  // Перевірка основних файлів
  checkFile('package.json', 'Конфігурація проекту');
  checkFile('index.js', 'Головний файл бота');
  checkFile('deploy-commands.js', 'Скрипт реєстрації команд');
  checkFile('env.example', 'Приклад змінних середовища');

  // Перевірка директорій
  checkDirectory('commands', 'Директорія команд');
  checkDirectory('utils', 'Директорія утиліт');
  checkDirectory('config', 'Директорія конфігурації');
  checkDirectory('metrics', 'Директорія метрик');
  checkDirectory('logs', 'Директорія логів');
  checkDirectory('scripts', 'Директорія скриптів');

  // Перевірка документації
  checkFile('README.md', 'Основна документація');
  checkFile('USAGE_GUIDE.md', 'Посібник користувача');
  checkFile('SETUP.md', 'Інструкції налаштування');
  checkFile('SECURITY_GUIDE.md', 'Безпека');
  checkFile('LAUNCH_INSTRUCTIONS.md', 'Інструкції запуску');

  // Перевірка тестів
  checkFile('test-integration.js', 'Інтеграційні тести');
  checkFile('test-comprehensive.js', 'Комплексні тести');
  checkFile('test-commands.js', 'Тести команд');
  checkFile('test-load.js', 'Навантажувальні тести');

  // Перевірка утиліт
  checkFile('src/utils/performanceOptimizer.js', 'Оптимізатор продуктивності');
  checkFile('src/utils/queueManager.js', 'Менеджер черг');
  checkFile('src/utils/clusterManager.js', 'Менеджер кластера');

  // Перевірка команд
  checkFile('src/commands/performanceMonitor.js', 'Команда моніторингу');

  // Перевірка скриптів розгортання
  checkFile('build/scripts/deploy.sh', 'Скрипт розгортання (Linux/macOS)');
  checkFile('build/scripts/deploy.ps1', 'Скрипт розгортання (Windows)');
  checkFile('build/scripts/start.js', 'Скрипт запуску');
  checkFile('src/config/Config.js', 'Конфігурація середовищ');
}

// Функція перевірки середовища
function checkEnvironment() {
  logInfo('Перевірка середовища...');

  // Перевірка Node.js версії
  checkNodeVersion();

  // Перевірка npm залежностей
  checkNpmDependencies();

  // Перевірка .env файлу
  checkEnvFile();

  // Перевірка змінних середовища
  checkEnvironmentVariable('DISCORD_TOKEN', 'Discord токен');
  checkEnvironmentVariable('CLIENT_ID', 'Discord Client ID');
  checkEnvironmentVariable('GUILD_ID', 'Discord Guild ID');
  checkEnvironmentVariable('GOOGLE_API_KEY', 'Google API ключ');
  checkEnvironmentVariable('OPENAI_API_KEY', 'OpenAI API ключ');
}

// Функція перевірки продуктивності
function checkPerformance() {
  logInfo('Перевірка продуктивності...');

  // Перевірка файлів оптимізації
  checkFile('src/utils/performanceOptimizer.js', 'Оптимізатор продуктивності');
  checkFile('src/utils/queueManager.js', 'Система черг');
  checkFile('src/utils/clusterManager.js', 'Кластеризація');

  // Перевірка метрик
  checkDirectory('data/metrics', 'Директорія метрик');

  // Перевірка логів
  checkDirectory('data/logs', 'Директорія логів');

  // Створення тестових файлів логів
  const logFiles = ['data/logs/bot.log', 'data/logs/error.log'];
  for (const logFile of logFiles) {
    if (!fs.existsSync(logFile)) {
      fs.ensureFileSync(logFile);
      logInfo(`Створено файл логу: ${logFile}`);
    }
  }
}

// Функція перевірки безпеки
function checkSecurity() {
  logInfo('Перевірка безпеки...');

  // Перевірка .gitignore
  const gitignorePath = path.join(process.cwd(), '.gitignore');
  if (fs.existsSync(gitignorePath)) {
    const gitignoreContent = fs.readFileSync(gitignorePath, 'utf8');

    if (gitignoreContent.includes('.env')) {
      logSuccess('✓ .env файл в .gitignore');
      results.passed++;
    } else {
      logWarning('⚠ .env файл не в .gitignore');
      results.warnings++;
    }

    if (gitignoreContent.includes('node_modules')) {
      logSuccess('✓ node_modules в .gitignore');
      results.passed++;
    } else {
      logWarning('⚠ node_modules не в .gitignore');
      results.warnings++;
    }

    if (gitignoreContent.includes('logs')) {
      logSuccess('✓ logs в .gitignore');
      results.passed++;
    } else {
      logWarning('⚠ logs не в .gitignore');
      results.warnings++;
    }
  } else {
    logWarning('⚠ .gitignore файл не знайдено');
    results.warnings++;
  }

  // Перевірка ESLint конфігурації
  checkFile('.eslintrc.json', 'ESLint конфігурація');

  // Перевірка Prettier конфігурації
  checkFile('.prettierrc', 'Prettier конфігурація');
}

// Функція виведення результатів
function printResults() {
  console.log('\n' + '='.repeat(60));
  log('РЕЗУЛЬТАТИ ФІНАЛЬНОЇ ПЕРЕВІРКИ', 'bright');
  console.log('='.repeat(60));

  const total = results.passed + results.failed + results.warnings;
  const successRate = total > 0 ? Math.round((results.passed / total) * 100) : 0;

  log(`✅ Успішно: ${results.passed}`, 'green');
  log(`❌ Помилки: ${results.failed}`, 'red');
  log(`⚠️ Попередження: ${results.warnings}`, 'yellow');
  log(`📊 Всього: ${total}`, 'blue');
  log(`📈 Відсоток успіху: ${successRate}%`, 'cyan');

  console.log('\n' + '='.repeat(60));

  if (results.failed === 0 && successRate >= 90) {
    log('🎉 ВСІ ПЕРЕВІРКИ ПРОЙШЛИ УСПІШНО!', 'green');
    log('🚀 Бот готовий до запуску в продакшені!', 'green');
  } else if (results.failed === 0) {
    log('⚠️ ПЕРЕВІРКИ ПРОЙШЛИ З ПОПЕРЕДЖЕННЯМИ', 'yellow');
    log('🔧 Рекомендується виправити попередження', 'yellow');
  } else {
    log('❌ ЗНАЙДЕНО ПОМИЛКИ', 'red');
    log('🔧 Необхідно виправити помилки перед запуском', 'red');
  }

  console.log('='.repeat(60));
}

// Функція рекомендацій
function printRecommendations() {
  console.log('\n📋 РЕКОМЕНДАЦІЇ:');

  if (results.failed > 0) {
    log('🔧 КРИТИЧНІ ДІЇ:', 'red');
    log('1. Виправте всі помилки перед запуском');
    log('2. Перевірте наявність всіх необхідних файлів');
    log('3. Налаштуйте змінні середовища');
  }

  if (results.warnings > 0) {
    log('⚠️ РЕКОМЕНДОВАНІ ДІЇ:', 'yellow');
    log('1. Додайте .env в .gitignore');
    log('2. Налаштуйте всі опціональні змінні');
    log('3. Перевірте конфігурацію безпеки');
  }

  log('🚀 НАСТУПНІ КРОКИ:', 'cyan');
  log('1. Запустіть: npm run setup');
  log('2. Налаштуйте .env файл');
  log('3. Запустіть: npm run deploy:guild');
  log('4. Запустіть бота: npm start');

  log('📚 КОРИСНІ РЕСУРСИ:', 'blue');
  log('- README.md - основна документація');
  log('- SETUP.md - інструкції налаштування');
  log('- USAGE_GUIDE.md - посібник користувача');
  log('- FAQ_SUPPORT.md - часто задавані питання');
}

// Головна функція
function main() {
  console.log(colors.cyan + '✅ Фінальна перевірка Discord AI Assistant Bot v2.3.0' + colors.reset);
  console.log(colors.cyan + '='.repeat(60) + colors.reset);

  // Скидання результатів
  results.passed = 0;
  results.failed = 0;
  results.warnings = 0;
  results.total = 0;

  // Виконання перевірок
  checkConfiguration();
  checkEnvironment();
  checkPerformance();
  checkSecurity();

  // Запуск тестів (опціонально)
  const runTestsFlag = process.argv.includes('--tests');
  if (runTestsFlag) {
    runTests();
  }

  // Виведення результатів
  printResults();
  printRecommendations();

  // Повернення коду виходу
  if (results.failed > 0) {
    process.exit(1);
  } else if (results.warnings > 0) {
    process.exit(2);
  } else {
    process.exit(0);
  }
}

// Запуск головної функції
if (require.main === module) {
  main();
}

module.exports = {
  checkFile,
  checkDirectory,
  checkEnvironmentVariable,
  checkEnvFile,
  checkNpmDependencies,
  checkNodeVersion,
  runTests,
  checkConfiguration,
  checkEnvironment,
  checkPerformance,
  checkSecurity,
  printResults,
  printRecommendations,
};
