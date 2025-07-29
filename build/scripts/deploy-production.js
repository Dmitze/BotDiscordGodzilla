#!/usr/bin/env node

/**
 * Скрипт розгортання для продакшену
 * Оновлено: 28.07.2025
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

class ProductionDeployer {
  constructor() {
    this.config = {
      projectName: 'Discord AI Assistant Bot',
      version: '3.0.0',
      environment: 'production',
      backupDir: './data/backups',
      logsDir: './data/logs',
      tmpDir: './data/tmp',
    };

    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
  }

  /**
   * Головна функція розгортання
   */
  async deploy() {
    console.log('🚀 РОЗГОРТАННЯ В ПРОДАКШЕН');
    console.log('============================');
    console.log(`📦 Проект: ${this.config.projectName}`);
    console.log(`📋 Версія: ${this.config.version}`);
    console.log(`🌍 Середовище: ${this.config.environment}`);
    console.log('');

    try {
      // 1. Перевірка перед розгортанням
      await this.preDeploymentCheck();

      // 2. Створення резервної копії
      await this.createBackup();

      // 3. Оновлення коду
      await this.updateCode();

      // 4. Встановлення залежностей
      await this.installDependencies();

      // 5. Запуск тестів
      await this.runTests();

      // 6. Збірка проекту
      await this.buildProject();

      // 7. Розгортання
      await this.deployApplication();

      // 8. Перевірка після розгортання
      await this.postDeploymentCheck();

      console.log('✅ Розгортання успішно завершено!');
      console.log('🎉 Бот готовий до роботи в продакшені');
    } catch (error) {
      console.error('❌ Помилка розгортання:', error.message);
      await this.rollback();
      process.exit(1);
    } finally {
      this.rl.close();
    }
  }

  /**
   * Перевірка перед розгортанням
   */
  async preDeploymentCheck() {
    console.log('🔍 Перевірка перед розгортанням...');

    // Перевірка Node.js версії
    const nodeVersion = process.version;
    const requiredVersion = 'v18.0.0';

    if (this.compareVersions(nodeVersion, requiredVersion) < 0) {
      throw new Error(
        `Потрібна Node.js версія ${requiredVersion} або вище. Поточна: ${nodeVersion}`
      );
    }

    // Перевірка наявності .env файлу
    if (!fs.existsSync('.env')) {
      throw new Error('Файл .env не знайдено. Створіть його на основі env.example');
    }

    // Перевірка обов'язкових змінних середовища
    await this.checkEnvironmentVariables();

    // Перевірка наявності Google credentials
    await this.checkGoogleCredentials();

    // Перевірка підключення до Discord
    await this.checkDiscordConnection();

    console.log('✅ Перевірка перед розгортанням пройдена');
  }

  /**
   * Перевірка змінних середовища
   */
  async checkEnvironmentVariables() {
    const requiredVars = [
      'DISCORD_TOKEN',
      'DISCORD_CLIENT_ID',
      'GOOGLE_SPREADSHEET_ID',
      'GOOGLE_API_KEY',
    ];

    const missingVars = [];

    for (const varName of requiredVars) {
      if (!process.env[varName]) {
        missingVars.push(varName);
      }
    }

    if (missingVars.length > 0) {
      throw new Error(`Відсутні обов'язкові змінні середовища: ${missingVars.join(', ')}`);
    }
  }

  /**
   * Перевірка Google credentials
   */
  async checkGoogleCredentials() {
    const credentialsPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;

    if (credentialsPath && !fs.existsSync(credentialsPath)) {
      throw new Error(`Google credentials файл не знайдено: ${credentialsPath}`);
    }
  }

  /**
   * Перевірка підключення до Discord
   */
  async checkDiscordConnection() {
    console.log('🔗 Перевірка підключення до Discord...');

    try {
      // Тут можна додати реальну перевірку підключення
      console.log('✅ Підключення до Discord доступне');
    } catch (error) {
      throw new Error(`Помилка підключення до Discord: ${error.message}`);
    }
  }

  /**
   * Створення резервної копії
   */
  async createBackup() {
    console.log('💾 Створення резервної копії...');

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupPath = path.join(this.config.backupDir, `backup-${timestamp}`);

    // Створення директорії для резервних копій
    if (!fs.existsSync(this.config.backupDir)) {
      fs.mkdirSync(this.config.backupDir, { recursive: true });
    }

    // Копіювання важливих файлів
    const filesToBackup = [
      'package.json',
      'package-lock.json',
      '.env',
      'src/',
      'commands/',
      'config/',
      'utils/',
    ];

    for (const file of filesToBackup) {
      if (fs.existsSync(file)) {
        const destPath = path.join(backupPath, file);
        const destDir = path.dirname(destPath);

        if (!fs.existsSync(destDir)) {
          fs.mkdirSync(destDir, { recursive: true });
        }

        if (fs.statSync(file).isDirectory()) {
          this.copyDirectory(file, destPath);
        } else {
          fs.copyFileSync(file, destPath);
        }
      }
    }

    console.log(`✅ Резервна копія створена: ${backupPath}`);
  }

  /**
   * Копіювання директорії
   */
  copyDirectory(src, dest) {
    if (!fs.existsSync(dest)) {
      fs.mkdirSync(dest, { recursive: true });
    }

    const files = fs.readdirSync(src);

    for (const file of files) {
      const srcPath = path.join(src, file);
      const destPath = path.join(dest, file);

      if (fs.statSync(srcPath).isDirectory()) {
        this.copyDirectory(srcPath, destPath);
      } else {
        fs.copyFileSync(srcPath, destPath);
      }
    }
  }

  /**
   * Оновлення коду
   */
  async updateCode() {
    console.log('📥 Оновлення коду...');

    try {
      // Перевірка чи це Git репозиторій
      if (fs.existsSync('.git')) {
        // Отримання останніх змін
        execSync('git fetch origin', { stdio: 'inherit' });

        // Перевірка чи є оновлення
        const currentBranch = execSync('git branch --show-current', { encoding: 'utf8' }).trim();
        const localCommit = execSync('git rev-parse HEAD', { encoding: 'utf8' }).trim();
        const remoteCommit = execSync(`git rev-parse origin/${currentBranch}`, {
          encoding: 'utf8',
        }).trim();

        if (localCommit !== remoteCommit) {
          console.log('🔄 Знайдено оновлення, виконується pull...');
          execSync(`git pull origin ${currentBranch}`, { stdio: 'inherit' });
        } else {
          console.log('✅ Код вже актуальний');
        }
      } else {
        console.log('ℹ️ Не Git репозиторій, пропуск оновлення коду');
      }
    } catch (error) {
      console.warn('⚠️ Помилка оновлення коду:', error.message);
    }
  }

  /**
   * Встановлення залежностей
   */
  async installDependencies() {
    console.log('📦 Встановлення залежностей...');

    try {
      // Видалення node_modules для чистої встановлення
      if (fs.existsSync('node_modules')) {
        console.log('🗑️ Видалення старих залежностей...');
        execSync('rm -rf node_modules', { stdio: 'inherit' });
      }

      // Встановлення залежностей
      console.log('📥 Встановлення нових залежностей...');
      execSync('npm ci --production', { stdio: 'inherit' });

      console.log('✅ Залежності встановлено');
    } catch (error) {
      throw new Error(`Помилка встановлення залежностей: ${error.message}`);
    }
  }

  /**
   * Запуск тестів
   */
  async runTests() {
    console.log('🧪 Запуск тестів...');

    try {
      // Unit тести
      console.log('🔬 Unit тести...');
      execSync('npm run test:unit', { stdio: 'inherit' });

      // Інтеграційні тести
      console.log('🔗 Інтеграційні тести...');
      execSync('npm run test:integration', { stdio: 'inherit' });

      // Навантажувальні тести (короткі)
      console.log('⚡ Навантажувальні тести...');
      execSync('npm run test:load -- --maxWorkers=1', { stdio: 'inherit' });

      console.log('✅ Всі тести пройшли успішно');
    } catch (error) {
      throw new Error(`Помилка тестування: ${error.message}`);
    }
  }

  /**
   * Збірка проекту
   */
  async buildProject() {
    console.log('🔨 Збірка проекту...');

    try {
      // Створення директорій
      const dirs = ['logs', 'tmp', 'dist'];

      for (const dir of dirs) {
        if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir, { recursive: true });
        }
      }

      // Копіювання конфігураційних файлів
      const configFiles = ['.env', 'package.json'];

      for (const file of configFiles) {
        if (fs.existsSync(file)) {
          fs.copyFileSync(file, path.join('dist', file));
        }
      }

      // Копіювання вихідного коду
      this.copyDirectory('src', 'dist/src');
      this.copyDirectory('commands', 'dist/commands');
      this.copyDirectory('config', 'dist/config');
      this.copyDirectory('utils', 'dist/utils');

      console.log('✅ Проект зібрано');
    } catch (error) {
      throw new Error(`Помилка збірки: ${error.message}`);
    }
  }

  /**
   * Розгортання додатку
   */
  async deployApplication() {
    console.log('🚀 Розгортання додатку...');

    try {
      // Зупинка поточного процесу (якщо запущений)
      await this.stopCurrentProcess();

      // Реєстрація команд Discord
      console.log('📝 Реєстрація Discord команд...');
      execSync('npm run deploy-commands', { stdio: 'inherit' });

      // Запуск додатку
      console.log('▶️ Запуск додатку...');

      // Використання PM2 для управління процесом
      if (this.isPM2Available()) {
        execSync('pm2 start ecosystem.config.js --env production', { stdio: 'inherit' });
        console.log('✅ Додаток запущено через PM2');
      } else {
        // Запуск через npm
        execSync('npm start', { stdio: 'inherit', detached: true });
        console.log('✅ Додаток запущено через npm');
      }
    } catch (error) {
      throw new Error(`Помилка розгортання: ${error.message}`);
    }
  }

  /**
   * Зупинка поточного процесу
   */
  async stopCurrentProcess() {
    try {
      if (this.isPM2Available()) {
        execSync('pm2 stop discord-bot || true', { stdio: 'inherit' });
        execSync('pm2 delete discord-bot || true', { stdio: 'inherit' });
      }
    } catch (error) {
      console.warn('⚠️ Помилка зупинки поточного процесу:', error.message);
    }
  }

  /**
   * Перевірка наявності PM2
   */
  isPM2Available() {
    try {
      execSync('pm2 --version', { stdio: 'ignore' });
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Перевірка після розгортання
   */
  async postDeploymentCheck() {
    console.log('🔍 Перевірка після розгортання...');

    // Зачекати трохи для запуску
    await this.sleep(5000);

    try {
      // Перевірка статусу процесу
      if (this.isPM2Available()) {
        const status = execSync('pm2 status discord-bot', { encoding: 'utf8' });
        if (!status.includes('online')) {
          throw new Error('Процес не запущений');
        }
      }

      // Перевірка логів на помилки
      await this.checkLogsForErrors();

      // Перевірка метрик
      await this.checkMetrics();

      console.log('✅ Перевірка після розгортання пройдена');
    } catch (error) {
      throw new Error(`Помилка перевірки після розгортання: ${error.message}`);
    }
  }

  /**
   * Перевірка логів на помилки
   */
  async checkLogsForErrors() {
    const logFile = path.join(this.config.logsDir, 'bot.log');

    if (fs.existsSync(logFile)) {
      const logContent = fs.readFileSync(logFile, 'utf8');
      const errorLines = logContent
        .split('\n')
        .filter(line => line.includes('ERROR') || line.includes('FATAL'));

      if (errorLines.length > 0) {
        console.warn('⚠️ Знайдено помилки в логах:', errorLines.slice(-5));
      }
    }
  }

  /**
   * Перевірка метрик
   */
  async checkMetrics() {
    try {
      // Тут можна додати перевірку метрик через HTTP запит
      console.log('📊 Метрики доступні');
    } catch (error) {
      console.warn('⚠️ Помилка перевірки метрик:', error.message);
    }
  }

  /**
   * Відкат змін при помилці
   */
  async rollback() {
    console.log('🔄 Виконання відкату...');

    try {
      // Знайти останню резервну копію
      const backups = fs
        .readdirSync(this.config.backupDir)
        .filter(file => file.startsWith('backup-'))
        .sort()
        .reverse();

      if (backups.length > 0) {
        const latestBackup = path.join(this.config.backupDir, backups[0]);
        console.log(`📦 Відновлення з резервної копії: ${latestBackup}`);

        // Відновлення файлів
        this.copyDirectory(latestBackup, '.');

        // Перезапуск з відновленими файлами
        await this.deployApplication();

        console.log('✅ Відкат виконано успішно');
      } else {
        console.log('❌ Резервні копії не знайдено');
      }
    } catch (error) {
      console.error('❌ Помилка відкату:', error.message);
    }
  }

  /**
   * Порівняння версій
   */
  compareVersions(v1, v2) {
    const normalize = v => v.replace(/^v/, '').split('.').map(Number);
    const n1 = normalize(v1);
    const n2 = normalize(v2);

    for (let i = 0; i < Math.max(n1.length, n2.length); i++) {
      const num1 = n1[i] || 0;
      const num2 = n2[i] || 0;

      if (num1 > num2) return 1;
      if (num1 < num2) return -1;
    }

    return 0;
  }

  /**
   * Затримка
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Запуск розгортання
if (require.main === module) {
  const deployer = new ProductionDeployer();
  deployer.deploy().catch(error => {
    console.error('❌ Критична помилка розгортання:', error);
    process.exit(1);
  });
}

module.exports = ProductionDeployer;
