// Setup файл для Jest тестів

import { config } from 'dotenv';

// Завантаження змінних середовища
config({ path: '.env.test' });

// Мок для process.env
process.env['NODE_ENV'] = 'test';

// Базові налаштування для тестів
console.log('🧪 Тестове середовище ініціалізовано');

// Глобальная очистка ресурсов логгера после всех тестов
import logger from '../utils/logger';
afterAll(async () => {
  try {
    await logger.cleanup();
  } catch (e) {
    // ignore
  }
});
