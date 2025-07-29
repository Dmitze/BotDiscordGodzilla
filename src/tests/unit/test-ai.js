// Тестовий файл для перевірки AI-функціоналу
require('dotenv').config();

const { 
  analyzeData, 
  naturalLanguageSearch, 
  generateRecommendations, 
  generateSmartReport 
} = require('./aiHelpers');

// Тестові дані
const testHeaders = ['Найменування номенклатури', 'Серійний номер', 'Контрагент', 'Кількість', 'Ціна', 'Вартість'];
const testData = [
  ['Компьютерна мишка', 'SN001', 'Компанія А', '50', '150', '7500'],
  ['Клавіатура', 'SN002', 'Компанія А', '30', '200', '6000'],
  ['Монітор', 'SN003', 'ООО Рога і Копита', '10', '2000', '20000'],
  ['Принтер HP', 'SN004', 'ТОВ Бізнес', '5', '3000', '15000'],
  ['Сканер', 'SN005', 'Компанія А', '15', '500', '7500']
];

async function testAI() {
  console.log('🤖 Тестування AI-функціоналу...\n');

  try {
    // Тест 1: AI-аналіз
    console.log('📊 Тест 1: AI-аналіз даних');
    const analysis = await analyzeData(testData, testHeaders);
    console.log(analysis);
    console.log('\n' + '='.repeat(50) + '\n');

    // Тест 2: Природномовний пошук
    console.log('🔍 Тест 2: Природномовний пошук');
    const searchResult = await naturalLanguageSearch('покажи товари дешевше 1000', testData, testHeaders);
    console.log('Результати пошуку:', searchResult.results.length);
    console.log('Пояснення:', searchResult.explanation);
    console.log('\n' + '='.repeat(50) + '\n');

    // Тест 3: AI-рекомендації
    console.log('💡 Тест 3: AI-рекомендації');
    const recommendations = await generateRecommendations(testData, testHeaders);
    console.log(recommendations);
    console.log('\n' + '='.repeat(50) + '\n');

    // Тест 4: AI-звіт
    console.log('📋 Тест 4: AI-звіт');
    const report = await generateSmartReport(testData, testHeaders, 'general');
    console.log(report);

  } catch (error) {
    console.error('❌ Помилка при тестуванні:', error.message);
  }
}

// Запуск тесту тільки якщо є OpenAI API ключ
if (process.env.OPENAI_API_KEY) {
  testAI();
} else {
  console.log('⚠️ OpenAI API ключ не знайдено в змінних середовища');
  console.log('Додайте OPENAI_API_KEY до файлу .env для тестування AI-функціоналу');
} 