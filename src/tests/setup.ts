// Setup файл для Jest тестів

import { config } from 'dotenv';

// Завантаження змінних середовища
config({ path: '.env.test' });

// Базові налаштування для тестів
console.log('🧪 Тестове середовище ініціалізовано'); 