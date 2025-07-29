/**
 * 🚀 Скрипт запуску Discord AI Assistant Bot
 * Версія: 2.3.0
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs-extra');

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

// Перевірка наявності файлів
function checkFiles() {
  const requiredFiles = ['package.json', 'index.js', '.env'];

  const missingFiles = [];

  for (const file of requiredFiles) {
    if (!fs.existsSync(file)) {
      missingFiles.push(file);
    }
  }

  if (missingFiles.length > 0) {
    logError(`Відсутні файли: ${missingFiles.join(', ')}`);
    return false;
  }

  return true;
}

// Перевірка змінних середовища
function checkEnvironment() {
  const envPath = path.join(process.cwd(), '.env');

  if (!fs.existsSync(envPath)) {
    logError('Файл .env не знайдено');
    logInfo('Скопіюйте env.example в .env та налаштуйте змінні');
    return false;
  }

  const envContent = fs.readFileSync(envPath, 'utf8');
  const requiredVars = ['DISCORD_TOKEN', 'CLIENT_ID', 'GUILD_ID'];
  const missingVars = [];

  for (const varName of requiredVars) {
    if (!envContent.includes(`${varName}=`)) {
      missingVars.push(varName);
    }
  }

  if (missingVars.length > 0) {
    logWarning(`Відсутні змінні в .env: ${missingVars.join(', ')}`);
  }

  return true;
}

// Створення директорій
function createDirectories() {
  const directories = ['data/logs', 'data/metrics', 'data/tmp', 'data/cache'];

  for (const dir of directories) {
    const dirPath = path.join(process.cwd(), dir);
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
      logInfo(`Створено директорію: ${dir}`);
    }
  }
}

// Запуск бота в звичайному режимі
function startNormal() {
  logInfo('Запуск бота в звичайному режимі...');

  const child = spawn('node', ['index.js'], {
    stdio: 'inherit',
    env: { ...process.env, NODE_ENV: 'production' },
  });

  child.on('error', error => {
    logError(`Помилка запуску: ${error.message}`);
    process.exit(1);
  });

  child.on('exit', code => {
    if (code !== 0) {
      logError(`Бот завершився з кодом: ${code}`);
      process.exit(code);
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
}

// Запуск бота в режимі розробки
function startDevelopment() {
  logInfo('Запуск бота в режимі розробки...');

  const child = spawn('node', ['index.js'], {
    stdio: 'inherit',
    env: { ...process.env, NODE_ENV: 'development' },
  });

  child.on('error', error => {
    logError(`Помилка запуску: ${error.message}`);
    process.exit(1);
  });

  child.on('exit', code => {
    if (code !== 0) {
      logError(`Бот завершився з кодом: ${code}`);
      process.exit(code);
    }
  });

  // Автоматичний перезапуск при зміні файлів
  let restartTimeout;
  const watchFiles = ['index.js', 'commands/', 'utils/', 'config/'];

  for (const file of watchFiles) {
    const filePath = path.join(process.cwd(), file);
    if (fs.existsSync(filePath)) {
      fs.watch(filePath, { recursive: true }, () => {
        clearTimeout(restartTimeout);
        restartTimeout = setTimeout(() => {
          logInfo('Файли змінені, перезапуск...');
          child.kill('SIGTERM');
        }, 1000);
      });
    }
  }

  // Обробка сигналів завершення
  process.on('SIGINT', () => {
    logInfo('Отримано сигнал SIGINT, завершення...');
    child.kill('SIGINT');
  });

  process.on('SIGTERM', () => {
    logInfo('Отримано сигнал SIGTERM, завершення...');
    child.kill('SIGTERM');
  });
}

// Запуск бота в тестовому режимі
function startTesting() {
  logInfo('Запуск бота в тестовому режимі...');

  const child = spawn('node', ['index.js'], {
    stdio: 'inherit',
    env: { ...process.env, NODE_ENV: 'testing' },
  });

  child.on('error', error => {
    logError(`Помилка запуску: ${error.message}`);
    process.exit(1);
  });

  child.on('exit', code => {
    if (code !== 0) {
      logError(`Бот завершився з кодом: ${code}`);
      process.exit(code);
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
}

// Запуск бота з PM2
function startPM2() {
  logInfo('Запуск бота з PM2...');

  // Перевірка наявності PM2
  const pm2Check = spawn('pm2', ['--version'], { stdio: 'pipe' });

  pm2Check.on('error', () => {
    logError('PM2 не встановлено. Встановіть: npm install -g pm2');
    process.exit(1);
  });

  pm2Check.on('close', code => {
    if (code === 0) {
      // Запуск з PM2
      const pm2Process = spawn('pm2', ['start', 'ecosystem.config.js'], {
        stdio: 'inherit',
      });

      pm2Process.on('error', error => {
        logError(`Помилка запуску PM2: ${error.message}`);
        process.exit(1);
      });

      pm2Process.on('close', code => {
        if (code === 0) {
          logSuccess('Бот запущений з PM2');
          logInfo('Команди PM2:');
          logInfo('  pm2 status - статус процесів');
          logInfo('  pm2 logs - перегляд логів');
          logInfo('  pm2 stop discord-bot - зупинка');
          logInfo('  pm2 restart discord-bot - перезапуск');
        } else {
          logError(`PM2 завершився з кодом: ${code}`);
          process.exit(code);
        }
      });
    } else {
      logError('PM2 не встановлено');
      process.exit(1);
    }
  });
}

// Запуск бота з Docker
function startDocker() {
  logInfo('Запуск бота з Docker...');

  // Перевірка наявності Docker
  const dockerCheck = spawn('docker', ['--version'], { stdio: 'pipe' });

  dockerCheck.on('error', () => {
    logError('Docker не встановлено');
    process.exit(1);
  });

  dockerCheck.on('close', code => {
    if (code === 0) {
      // Запуск з Docker Compose
      const dockerProcess = spawn('docker-compose', ['up', '-d'], {
        stdio: 'inherit',
      });

      dockerProcess.on('error', error => {
        logError(`Помилка запуску Docker: ${error.message}`);
        process.exit(1);
      });

      dockerProcess.on('close', code => {
        if (code === 0) {
          logSuccess('Бот запущений з Docker');
          logInfo('Команди Docker:');
          logInfo('  docker-compose logs -f - перегляд логів');
          logInfo('  docker-compose down - зупинка');
          logInfo('  docker-compose restart - перезапуск');
        } else {
          logError(`Docker завершився з кодом: ${code}`);
          process.exit(code);
        }
      });
    } else {
      logError('Docker не встановлено');
      process.exit(1);
    }
  });
}

// Показ довідки
function showHelp() {
  console.log(`
${colors.cyan}🚀 Discord AI Assistant Bot - Скрипт запуску v2.3.0${colors.reset}
${colors.cyan}================================================${colors.reset}

${colors.yellow}Використання:${colors.reset}
  node scripts/start.js [опції]

${colors.yellow}Опції:${colors.reset}
  --dev, -d          Запуск в режимі розробки (з автоперезапуском)
  --test, -t         Запуск в тестовому режимі
  --pm2              Запуск з PM2 (кластеризація)
  --docker           Запуск з Docker
  --help, -h         Показати цю довідку

${colors.yellow}Приклади:${colors.reset}
  node scripts/start.js              # Звичайний запуск
  node scripts/start.js --dev        # Режим розробки
  node scripts/start.js --pm2        # Запуск з PM2
  node scripts/start.js --docker     # Запуск з Docker

${colors.yellow}Перед запуском:${colors.reset}
  1. Встановіть залежності: npm install
  2. Налаштуйте .env файл
  3. Зареєструйте команди: node deploy-commands.js

${colors.yellow}Документація:${colors.reset}
  README.md - основна документація
  LAUNCH_INSTRUCTIONS.md - інструкції запуску
  FAQ_SUPPORT.md - часто задавані питання
`);
}

// Головна функція
function main() {
  const args = process.argv.slice(2);

  // Парсинг аргументів
  const options = {
    dev: args.includes('--dev') || args.includes('-d'),
    test: args.includes('--test') || args.includes('-t'),
    pm2: args.includes('--pm2'),
    docker: args.includes('--docker'),
    help: args.includes('--help') || args.includes('-h'),
  };

  // Показ довідки
  if (options.help) {
    showHelp();
    return;
  }

  // Перевірка файлів
  if (!checkFiles()) {
    process.exit(1);
  }

  // Перевірка середовища
  if (!checkEnvironment()) {
    process.exit(1);
  }

  // Створення директорій
  createDirectories();

  // Запуск відповідного режиму
  if (options.docker) {
    startDocker();
  } else if (options.pm2) {
    startPM2();
  } else if (options.test) {
    startTesting();
  } else if (options.dev) {
    startDevelopment();
  } else {
    startNormal();
  }
}

// Запуск головної функції
if (require.main === module) {
  main();
}

module.exports = {
  startNormal,
  startDevelopment,
  startTesting,
  startPM2,
  startDocker,
  showHelp,
};
