/**
 * 🚀 Запуск Discord AI Assistant Bot в продакшені
 * Версія: 2.3.0
 */

const { spawn } = require('child_process');
const fs = require('fs-extra');
const path = require('path');

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

// Функція перевірки готовності
function checkReadiness() {
  logInfo('Перевірка готовності до запуску...');

  const checks = [
    { name: 'package.json', path: 'package.json' },
    { name: 'index.js', path: 'index.js' },
    { name: '.env файл', path: '.env' },
    { name: 'logs директорія', path: 'logs' },
    { name: 'node_modules', path: 'node_modules' },
  ];

  let allChecksPassed = true;

  for (const check of checks) {
    if (fs.existsSync(check.path)) {
      logSuccess(`✓ ${check.name}`);
    } else {
      logError(`✗ ${check.name} не знайдено`);
      allChecksPassed = false;
    }
  }

  return allChecksPassed;
}

// Функція перевірки змінних середовища
function checkEnvironmentVariables() {
  logInfo('Перевірка змінних середовища...');

  const requiredVars = [
    'DISCORD_TOKEN',
    'CLIENT_ID',
    'GUILD_ID',
    'GOOGLE_API_KEY',
    'OPENAI_API_KEY',
  ];

  let allVarsPresent = true;

  for (const varName of requiredVars) {
    if (process.env[varName]) {
      logSuccess(`✓ ${varName}`);
    } else {
      logError(`✗ ${varName} не встановлено`);
      allVarsPresent = false;
    }
  }

  return allVarsPresent;
}

// Функція створення директорій
function createDirectories() {
  logInfo('Створення необхідних директорій...');

  const directories = ['logs', 'metrics', 'tmp', 'config'];

  for (const dir of directories) {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
      logInfo(`Створено директорію: ${dir}`);
    }
  }
}

// Функція налаштування логування
function setupLogging() {
  logInfo('Налаштування логування...');

  const logFiles = ['logs/bot.log', 'logs/error.log', 'logs/combined.log'];

  for (const logFile of logFiles) {
    if (!fs.existsSync(logFile)) {
      fs.ensureFileSync(logFile);
      logInfo(`Створено файл логу: ${logFile}`);
    }
  }
}

// Функція реєстрації Discord команд
function deployCommands() {
  return new Promise((resolve, reject) => {
    logInfo('Реєстрація Discord команд...');

    const child = spawn('node', ['deploy-commands.js'], {
      stdio: 'pipe',
    });

    let output = '';

    child.stdout.on('data', data => {
      output += data.toString();
    });

    child.stderr.on('data', data => {
      output += data.toString();
    });

    child.on('close', code => {
      if (code === 0) {
        logSuccess('Discord команди зареєстровані успішно');
        resolve();
      } else {
        logError(`Помилка реєстрації команд: ${output}`);
        reject(new Error(`Exit code: ${code}`));
      }
    });

    child.on('error', error => {
      logError(`Помилка запуску deploy-commands.js: ${error.message}`);
      reject(error);
    });
  });
}

// Функція запуску бота
function startBot() {
  return new Promise((resolve, reject) => {
    logInfo('Запуск Discord бота...');

    const child = spawn('node', ['index.js'], {
      stdio: 'inherit',
      env: { ...process.env, NODE_ENV: 'production' },
    });

    child.on('error', error => {
      logError(`Помилка запуску бота: ${error.message}`);
      reject(error);
    });

    child.on('exit', code => {
      if (code === 0) {
        logSuccess('Бот завершив роботу успішно');
        resolve();
      } else {
        logError(`Бот завершився з кодом: ${code}`);
        reject(new Error(`Exit code: ${code}`));
      }
    });

    // Обробка сигналів завершення
    process.on('SIGINT', () => {
      logInfo('Отримано сигнал SIGINT, завершення...');
      child.kill('SIGINT');
    });

    process.on('SIGTERM', () => {
      logInfo('Отримано сигнал SIGTERM, завершення...');
      child.kill('SIGTERM');
    });

    // Передача управління дочірньому процесу
    resolve(child);
  });
}

// Функція запуску з PM2
function startWithPM2() {
  return new Promise((resolve, reject) => {
    logInfo('Запуск бота з PM2...');

    // Перевірка наявності PM2
    const pm2Check = spawn('pm2', ['--version'], { stdio: 'pipe' });

    pm2Check.on('error', () => {
      logError('PM2 не встановлено. Встановіть: npm install -g pm2');
      reject(new Error('PM2 not installed'));
    });

    pm2Check.on('close', code => {
      if (code === 0) {
        // Створення PM2 конфігурації якщо не існує
        if (!fs.existsSync('ecosystem.config.js')) {
          const pm2Config = `module.exports = {
  apps: [{
    name: 'discord-bot',
    script: 'index.js',
    instances: 'max',
    exec_mode: 'cluster',
    env: {
      NODE_ENV: 'production'
    },
    error_file: './logs/err.log',
    out_file: './logs/out.log',
    log_file: './logs/combined.log',
    time: true,
    max_memory_restart: '500M',
    restart_delay: 4000,
    max_restarts: 10
  }]
};`;

          fs.writeFileSync('ecosystem.config.js', pm2Config);
          logInfo('Створено PM2 конфігурацію');
        }

        // Запуск з PM2
        const pm2Process = spawn('pm2', ['start', 'ecosystem.config.js'], {
          stdio: 'inherit',
        });

        pm2Process.on('error', error => {
          logError(`Помилка запуску PM2: ${error.message}`);
          reject(error);
        });

        pm2Process.on('close', code => {
          if (code === 0) {
            logSuccess('Бот запущений з PM2');
            logInfo('Команди PM2:');
            logInfo('  pm2 status - статус процесів');
            logInfo('  pm2 logs - перегляд логів');
            logInfo('  pm2 stop discord-bot - зупинка');
            logInfo('  pm2 restart discord-bot - перезапуск');
            resolve();
          } else {
            logError(`PM2 завершився з кодом: ${code}`);
            reject(new Error(`PM2 exit code: ${code}`));
          }
        });
      } else {
        logError('PM2 не встановлено');
        reject(new Error('PM2 not installed'));
      }
    });
  });
}

// Функція запуску з Docker
function startWithDocker() {
  return new Promise((resolve, reject) => {
    logInfo('Запуск бота з Docker...');

    // Перевірка наявності Docker
    const dockerCheck = spawn('docker', ['--version'], { stdio: 'pipe' });

    dockerCheck.on('error', () => {
      logError('Docker не встановлено');
      reject(new Error('Docker not installed'));
    });

    dockerCheck.on('close', code => {
      if (code === 0) {
        // Запуск з Docker Compose
        const dockerProcess = spawn('docker-compose', ['up', '-d'], {
          stdio: 'inherit',
        });

        dockerProcess.on('error', error => {
          logError(`Помилка запуску Docker: ${error.message}`);
          reject(error);
        });

        dockerProcess.on('close', code => {
          if (code === 0) {
            logSuccess('Бот запущений з Docker');
            logInfo('Команди Docker:');
            logInfo('  docker-compose logs -f - перегляд логів');
            logInfo('  docker-compose down - зупинка');
            logInfo('  docker-compose restart - перезапуск');
            resolve();
          } else {
            logError(`Docker завершився з кодом: ${code}`);
            reject(new Error(`Docker exit code: ${code}`));
          }
        });
      } else {
        logError('Docker не встановлено');
        reject(new Error('Docker not installed'));
      }
    });
  });
}

// Функція показу статусу
function showStatus() {
  logInfo('Статус запуску:');
  logInfo('1. Перевірка готовності - завершено');
  logInfo('2. Налаштування середовища - завершено');
  logInfo('3. Створення директорій - завершено');
  logInfo('4. Налаштування логування - завершено');
  logInfo('5. Реєстрація команд - завершено');
  logInfo('6. Запуск бота - в процесі...');
}

// Функція показу довідки
function showHelp() {
  console.log(`
${colors.cyan}🚀 Запуск Discord AI Assistant Bot в продакшені v2.3.0${colors.reset}
${colors.cyan}================================================${colors.reset}

${colors.yellow}Використання:${colors.reset}
  node scripts/launch-production.js [опції]

${colors.yellow}Опції:${colors.reset}
  --pm2              Запуск з PM2 (кластеризація)
  --docker           Запуск з Docker
  --no-deploy        Пропустити реєстрацію команд
  --help, -h         Показати цю довідку

${colors.yellow}Приклади:${colors.reset}
  node scripts/launch-production.js              # Звичайний запуск
  node scripts/launch-production.js --pm2        # Запуск з PM2
  node scripts/launch-production.js --docker     # Запуск з Docker
  node scripts/launch-production.js --no-deploy  # Без реєстрації команд

${colors.yellow}Перед запуском:${colors.reset}
  1. Переконайтеся, що .env файл налаштований
  2. Встановіть залежності: npm install
  3. Перевірте готовність: node scripts/final-check.js

${colors.yellow}Моніторинг:${colors.reset}
  - Логи: tail -f logs/bot.log
  - Метрики: curl http://localhost:9090/metrics
  - Discord команда: /продуктивність статистика

${colors.yellow}Підтримка:${colors.reset}
  - README.md - основна документація
  - FAQ_SUPPORT.md - часто задавані питання
  - LAUNCH_INSTRUCTIONS.md - інструкції запуску
`);
}

// Головна функція
async function main() {
  const args = process.argv.slice(2);

  // Парсинг аргументів
  const options = {
    pm2: args.includes('--pm2'),
    docker: args.includes('--docker'),
    noDeploy: args.includes('--no-deploy'),
    help: args.includes('--help') || args.includes('-h'),
  };

  // Показ довідки
  if (options.help) {
    showHelp();
    return;
  }

  console.log(
    colors.cyan + '🚀 Запуск Discord AI Assistant Bot в продакшені v2.3.0' + colors.reset
  );
  console.log(colors.cyan + '='.repeat(60) + colors.reset);

  try {
    // Перевірка готовності
    if (!checkReadiness()) {
      logError('Система не готова до запуску');
      process.exit(1);
    }

    // Перевірка змінних середовища
    if (!checkEnvironmentVariables()) {
      logError('Не всі необхідні змінні середовища встановлені');
      process.exit(1);
    }

    // Створення директорій
    createDirectories();

    // Налаштування логування
    setupLogging();

    // Реєстрація команд (якщо не пропущено)
    if (!options.noDeploy) {
      await deployCommands();
    }

    // Показ статусу
    showStatus();

    // Запуск бота
    if (options.docker) {
      await startWithDocker();
    } else if (options.pm2) {
      await startWithPM2();
    } else {
      await startBot();
    }

    logSuccess('🎉 Бот успішно запущений в продакшені!');
  } catch (error) {
    logError(`Помилка запуску: ${error.message}`);
    process.exit(1);
  }
}

// Запуск головної функції
if (require.main === module) {
  main();
}

module.exports = {
  checkReadiness,
  checkEnvironmentVariables,
  createDirectories,
  setupLogging,
  deployCommands,
  startBot,
  startWithPM2,
  startWithDocker,
  showHelp,
};
