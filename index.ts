/**
 * Головний entry point для Discord AI Assistant Bot
 * Рефакторована архітектура v3.0.0
 * TypeScript версія
 * 
 * Цей файл є єдиним entry point для запуску додатку
 * Всі інші модулі імпортуються через src/index.ts
 */

import { main } from './src/index';
import logger from './src/utils/logger';

/**
 * Головна функція запуску
 */
async function startApplication(): Promise<void> {
  try {
    logger.info('🚀 Запуск Discord AI Assistant Bot v3.0.0...');
    logger.info(`📅 Дата запуску: ${new Date().toISOString()}`);
    logger.info(`🌍 Середовище: ${process.env.NODE_ENV || 'development'}`);
    logger.info('');

    // Запуск головного модуля
    await main();

    logger.info('✅ Додаток успішно запущено');
  } catch (error) {
    logger.error('❌ Критична помилка при запуску додатку', { error });
    process.exit(1);
  }
}

// Запуск додатку
if (require.main === module) {
  startApplication();
}

// Експорт для зовнішнього використання
export {
  startApplication,
};