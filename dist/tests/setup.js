"use strict";
// Setup файл для Jest тестів
Object.defineProperty(exports, "__esModule", { value: true });
const dotenv_1 = require("dotenv");
// Завантаження змінних середовища
(0, dotenv_1.config)({ path: '.env.test' });
// Мок для process.env
process.env['NODE_ENV'] = 'test';
// Базові налаштування для тестів
console.log('🧪 Тестове середовище ініціалізовано');
//# sourceMappingURL=setup.js.map