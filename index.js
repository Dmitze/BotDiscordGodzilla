/**
 * Головний entry point для Discord AI Assistant Bot
 * Рефакторована архітектура v3.0.0
 * 
 * Цей файл є єдиним entry point для запуску додатку
 * Всі інші модулі імпортуються через src/index.js
 */

// Імпорт головного модуля додатку
const { main } = require('./src/index');

/**
 * Головна функція запуску
 */
async function startApplication() {
  try {
    console.log('🚀 Запуск Discord AI Assistant Bot v3.0.0...');
    console.log('📅 Дата запуску:', new Date().toISOString());
    console.log('🌍 Середовище:', process.env.NODE_ENV || 'development');
    console.log('');

    // Запуск головного модуля
    await main();

    console.log('✅ Додаток успішно запущено');
  } catch (error) {
    console.error('❌ Критична помилка при запуску додатку:', error);
    process.exit(1);
  }
}

// Запуск додатку
if (require.main === module) {
  startApplication();
}

// Експорт для зовнішнього використання
module.exports = {
  startApplication,
};
