const OpenAI = require('openai');
const { EmbedBuilder } = require('discord.js');

// Ініціалізація OpenAI
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// === ПОКРАЩЕНИЙ AI-АНАЛІЗ ДАНИХ ===
async function analyzeDataEnhanced(data, headers) {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для аналізу';
    }

    // Підготовка даних для аналізу
    const analysisData = data.slice(0, 100); // Збільшуємо до 100 рядків
    const dataSummary = {
      totalRows: data.length,
      sampleRows: analysisData.length,
      columns: headers,
      sampleData: analysisData.slice(0, 10) // Перші 10 рядків для прикладу
    };

    const prompt = `
Проаналізуй дані з таблиці та надай детальний аналіз українською мовою:

Дані для аналізу:
- Всього рядків: ${dataSummary.totalRows}
- Зразок даних: ${dataSummary.sampleRows} рядків
- Колонки: ${headers.join(', ')}

Зразок даних:
${dataSummary.sampleData.map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай детальний аналіз з такими розділами:
1. 📊 Загальна статистика (кількість записів, унікальні значення)
2. 📈 Основні тренди та патерни
3. ⚠️ Потенційні аномалії та проблеми
4. 💡 Практичні рекомендації
5. 🎯 Ключові метрики
6. 🔮 Прогнози та тенденції

Будь детальним та корисним. Використовуй емодзі для кращого форматування.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 800,
      temperature: 0.3
    });

    return completion.choices[0].message.content;
  } catch (error) {
    console.error('Помилка покращеного AI-аналізу:', error);
    return '❌ Помилка при аналізі даних';
  }
}

// === РОЗУМНИЙ ПОШУК З КОНТЕКСТОМ ===
async function smartSearchWithContext(query, data, headers, context = {}) {
  try {
    if (!data || data.length === 0) {
      return { results: [], explanation: 'Немає даних для пошуку' };
    }

    const prompt = `
Користувач шукає в таблиці: "${query}"

Контекст пошуку:
- Попередні запити: ${context.previousQueries?.join(', ') || 'немає'}
- Улюблені поля: ${context.favoriteFields?.join(', ') || 'немає'}
- Частота використання: ${context.usageFrequency || 'не відомо'}

Колонки таблиці: ${headers.join(', ')}

Зразок даних (перші 15 рядків):
${data.slice(0, 15).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Проаналізуй запит з урахуванням контексту та визнач:
1. Які поля потрібно перевірити
2. Які умови пошуку застосувати
3. Як інтерпретувати результат
4. Додаткові рекомендації для користувача

Поверни JSON у форматі:
{
  "searchFields": ["поле1", "поле2"],
  "searchConditions": ["умова1", "умова2"],
  "explanation": "Пояснення пошуку",
  "recommendations": ["рекомендація1", "рекомендація2"],
  "confidence": 0.95
}
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 400,
      temperature: 0.2
    });

    const aiResponse = completion.choices[0].message.content;
    let searchConfig;
    
    try {
      const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        searchConfig = JSON.parse(jsonMatch[0]);
      } else {
        throw new Error('Не вдалося парсити JSON');
      }
    } catch (parseError) {
      console.error('Помилка парсингу AI відповіді:', parseError);
      return { results: [], explanation: 'Помилка обробки AI-запиту' };
    }

    // Виконуємо пошук на основі AI-рекомендацій
    const results = performSmartSearch(data, headers, searchConfig);
    
    return {
      results: results.slice(0, 20), // Обмежуємо до 20 результатів
      explanation: searchConfig.explanation,
      recommendations: searchConfig.recommendations || [],
      confidence: searchConfig.confidence || 0.8
    };
  } catch (error) {
    console.error('Помилка розумного пошуку:', error);
    return { results: [], explanation: 'Помилка при пошуку' };
  }
}

// === ФУНКЦІЯ ВИКОНАННЯ РОЗУМНОГО ПОШУКУ ===
function performSmartSearch(data, headers, searchConfig) {
  const results = [];
  const searchFields = searchConfig.searchFields || [];
  
  for (const row of data) {
    let matches = 0;
    let totalChecks = 0;
    
    for (const field of searchFields) {
      const colIndex = getColumnIndex(headers, field);
      if (colIndex !== -1) {
        totalChecks++;
        const value = (row[colIndex] || '').toString().toLowerCase();
        
        // Перевіряємо умови пошуку
        for (const condition of searchConfig.searchConditions || []) {
          if (value.includes(condition.toLowerCase())) {
            matches++;
            break;
          }
        }
      }
    }
    
    // Якщо більше 50% полів відповідають умовам
    if (totalChecks > 0 && matches / totalChecks >= 0.5) {
      results.push(row);
    }
  }
  
  return results;
}

// === ПОКРАЩЕНІ AI-РЕКОМЕНДАЦІЇ ===
async function generateEnhancedRecommendations(data, headers, userContext = {}) {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для аналізу';
    }

    const analysisData = data.slice(0, 50);
    const dataSummary = {
      totalRows: data.length,
      sampleRows: analysisData.length,
      columns: headers,
      sampleData: analysisData.slice(0, 5)
    };

    const prompt = `
Проаналізуй дані та надай персоналізовані рекомендації:

Дані для аналізу:
- Всього рядків: ${dataSummary.totalRows}
- Зразок даних: ${dataSummary.sampleRows} рядків
- Колонки: ${headers.join(', ')}

Контекст користувача:
- Рівень досвіду: ${userContext.experienceLevel || 'не відомо'}
- Роль: ${userContext.role || 'користувач'}
- Цілі: ${userContext.goals || 'не вказано'}

Зразок даних:
${dataSummary.sampleData.map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай детальні рекомендації з такими розділами:
1. 🎯 Стратегічні рекомендації
2. 📊 Операційні покращення
3. 💰 Фінансові оптимізації
4. 🔧 Технічні покращення
5. 📈 Метрики для відстеження
6. 🚀 Наступні кроки

Будь конкретним та практичним. Використовуй емодзі для форматування.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 600,
      temperature: 0.4
    });

    return completion.choices[0].message.content;
  } catch (error) {
    console.error('Помилка покращених AI-рекомендацій:', error);
    return '❌ Помилка при генерації рекомендацій';
  }
}

// === ДОПОМІЖНА ФУНКЦІЯ ===
function getColumnIndex(headers, field) {
  const headerMap = {
    назва: [
      'найменування номенклатури',
      'назва',
      'наименование номенклатуры',
      'найменування'
    ],
    серія: [
      'серійний номер',
      'серйіний номер',
      'серийный номер',
      'серія'
    ],
    контрагент: [
      'контрагент',
      'постачальник',
      'поставщик'
    ],
    кількість: [
      'кількість',
      'залишок',
      'остаток',
      'количество'
    ],
    ціна: [
      'ціна',
      'цена',
      'вартість',
      'стоимость'
    ]
  };
  
  for (let i = 0; i < headers.length; i++) {
    const headerName = (headers[i] || '').toLowerCase().replace(/\s+/g, ' ').trim();
    if (headerMap[field]?.some(h => h.toLowerCase() === headerName)) {
      return i;
    }
  }
  return -1;
}

module.exports = {
  analyzeDataEnhanced,
  smartSearchWithContext,
  generateEnhancedRecommendations,
  performSmartSearch
}; 