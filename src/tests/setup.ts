// Setup файл для Jest тестів

import { config } from 'dotenv';

// Завантаження змінних середовища
config({ path: '.env.test' });

// Мок для process.env
process.env['NODE_ENV'] = 'test';

// Базові налаштування для тестів
console.log('🧪 Тестове середовище ініціалізовано'); 