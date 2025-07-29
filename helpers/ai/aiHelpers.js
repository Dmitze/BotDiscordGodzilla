const OpenAI = require('openai');
const { EmbedBuilder } = require('discord.js');

// Ініціалізація OpenAI
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// === AI-АНАЛІЗ ДАНИХ ===
async function analyzeData(data, headers) {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для аналізу';
    }

    // Підготовка даних для аналізу
    const analysisData = data.slice(0, 50); // Беремо перші 50 рядків для аналізу
    const dataSummary = {
      totalRows: data.length,
      sampleRows: analysisData.length,
      columns: headers,
      sampleData: analysisData.slice(0, 5) // Перші 5 рядків для прикладу
    };

    const prompt = `
Проаналізуй дані з таблиці та надай корисну інформацію:

Дані для аналізу:
- Всього рядків: ${dataSummary.totalRows}
- Зразок даних: ${dataSummary.sampleRows} рядків
- Колонки: ${headers.join(', ')}

Зразок даних:
${dataSummary.sampleData.map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай аналіз українською мовою з такими розділами:
1. Загальна статистика
2. Основні тренди
3. Потенційні аномалії
4. Рекомендації

Будь лаконічним та корисним.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 500,
      temperature: 0.3
    });

    return completion.choices[0].message.content;
  } catch (error) {
    console.error('Помилка AI-аналізу:', error);
    return '❌ Помилка при аналізі даних';
  }
}

// === ПРИРОДНОМОВНИЙ ПОШУК ===
async function naturalLanguageSearch(query, data, headers) {
  try {
    if (!data || data.length === 0) {
      return { results: [], explanation: 'Немає даних для пошуку' };
    }

    const prompt = `
Користувач шукає в таблиці: "${query}"

Колонки таблиці: ${headers.join(', ')}

Зразок даних (перші 10 рядків):
${data.slice(0, 10).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Проаналізуй запит та визнач:
1. Які поля потрібно перевірити
2. Які умови пошуку застосувати
3. Як інтерпретувати результат

Поверни JSON у форматі:
{
  "searchFields": ["поле1", "поле2"],
  "searchConditions": ["умова1", "умова2"],
  "explanation": "Пояснення пошуку"
}
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 300,
      temperature: 0.2
    });

    const aiResponse = completion.choices[0].message.content;
    let searchConfig;
    
    try {
      // Спробуємо парсити JSON
      const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        searchConfig = JSON.parse(jsonMatch[0]);
      } else {
        throw new Error('Не вдалося парсити JSON');
      }
    } catch (parseError) {
      // Якщо не вдалося парсити, використовуємо простий пошук
      searchConfig = {
        searchFields: ['назва', 'контрагент'],
        searchConditions: [query.toLowerCase()],
        explanation: 'Простий пошук за назвою та контрагентом'
      };
    }

    // Виконуємо пошук на основі AI-рекомендацій
    const results = data.filter(row => {
      return searchConfig.searchFields.some(field => {
        const colIndex = getColumnIndex(headers, field);
        if (colIndex === -1) return false;
        
        const cellValue = (row[colIndex] || '').toString().toLowerCase();
        return searchConfig.searchConditions.some(condition => 
          cellValue.includes(condition.toLowerCase())
        );
      });
    });

    return {
      results: results.slice(0, 20), // Обмежуємо до 20 результатів
      explanation: searchConfig.explanation
    };
  } catch (error) {
    console.error('Помилка природномовного пошуку:', error);
    return { results: [], explanation: 'Помилка при пошуку' };
  }
}

// === AI-РЕКОМЕНДАЦІЇ ===
async function generateRecommendations(data, headers) {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для рекомендацій';
    }

    const prompt = `
На основі даних з таблиці надай практичні рекомендації:

Дані:
- Колонки: ${headers.join(', ')}
- Кількість записів: ${data.length}

Зразок даних:
${data.slice(0, 10).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай 3-5 практичних рекомендацій українською мовою для:
1. Оптимізації процесів
2. Покращення ефективності
3. Виявлення проблем
4. Розвитку бізнесу

Будь конкретним та корисним.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 400,
      temperature: 0.4
    });

    return completion.choices[0].message.content;
  } catch (error) {
    console.error('Помилка генерації рекомендацій:', error);
    return '❌ Помилка при генерації рекомендацій';
  }
}

// === РОЗУМНІ ЗВІТИ ===
async function generateSmartReport(data, headers, reportType = 'general') {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для звіту';
    }

    const reportPrompts = {
      general: 'Створи загальний звіт з основними метриками та висновками',
      inventory: 'Створи звіт по залишках з рекомендаціями по поповненню',
      sales: 'Створи звіт по продажах з аналізом трендів',
      suppliers: 'Створи звіт по постачальниках з оцінкою ефективності'
    };

    const prompt = `
Створи професійний звіт на основі даних:

Тип звіту: ${reportPrompts[reportType]}
Колонки: ${headers.join(', ')}
Кількість записів: ${data.length}

Зразок даних:
${data.slice(0, 15).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Створи структурований звіт українською мовою з:
1. Виконавчим резюме
2. Основними метриками
3. Ключовими висновками
4. Рекомендаціями

Використову markdown форматування для структури.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 600,
      temperature: 0.3
    });

    return completion.choices[0].message.content;
  } catch (error) {
    console.error('Помилка генерації звіту:', error);
    return '❌ Помилка при створенні звіту';
  }
}

// === ДОПОМІЖНА ФУНКЦІЯ ДЛЯ ПОШУКУ ІНДЕКСУ КОЛОНКИ ===
function getColumnIndex(headers, field) {
  const headerMap = {
    назва: ['найменування номенклатури', 'назва', 'наименование номенклатуры', 'найменування'],
    серія: ['серійний номер', 'серйіний номер', 'серийный номер', 'серія'],
    контрагент: ['контрагент', 'постачальник', 'поставщик'],
    кількість: ['кількість', 'залишок', 'остаток', 'количество'],
    ціна: ['ціна', 'цена', 'вартість', 'стоимость'],
    вартість: ['вартість', 'стоимость']
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
  analyzeData,
  naturalLanguageSearch,
  generateRecommendations,
  generateSmartReport,
  getColumnIndex
}; 