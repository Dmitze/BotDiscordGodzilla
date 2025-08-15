// Setup файл для Jest тестів

import { config } from 'dotenv';
// ВАЖЛИВО: Мокаємо SecurityManager до імпорту тестованих модулів, щоб уникнути setInterval
jest.mock('../utils/security', () => {
  class SecurityManagerMock {
    public initialize(): void { /* no-op */ }
    public cleanup(): void { /* no-op */ }
  }
  return {
    SecurityManager: SecurityManagerMock,
    default: new SecurityManagerMock(),
  };
});

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
