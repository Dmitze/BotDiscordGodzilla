/**
 * Покращені AI допоміжні функції для Discord AI Assistant Bot
 * TypeScript версія
 */

import OpenAI from 'openai';
import { EmbedBuilder } from 'discord.js';

// Ініціалізація OpenAI
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

interface DataSummary {
  totalRows: number;
  sampleRows: number;
  columns: string[];
  sampleData: any[][];
}

interface SearchContext {
  previousQueries?: string[];
  favoriteFields?: string[];
  usageFrequency?: string;
  [key: string]: any;
}

interface EnhancedSearchConfig {
  searchFields: string[];
  searchConditions: string[];
  explanation: string;
  recommendations: string[];
  confidence: number;
}

interface UserContext {
  preferences?: string[];
  history?: string[];
  expertise?: string;
  [key: string]: any;
}

// === ПОКРАЩЕНИЙ AI-АНАЛІЗ ДАНИХ ===
async function analyzeDataEnhanced(data: any[][], headers: string[]): Promise<string> {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для аналізу';
    }

    // Підготовка даних для аналізу
    const analysisData = data.slice(0, 100); // Збільшуємо до 100 рядків
    const dataSummary: DataSummary = {
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

    return completion.choices[0].message.content || '❌ Помилка при аналізі даних';
  } catch (error) {
    console.error('Помилка покращеного AI-аналізу:', error);
    return '❌ Помилка при аналізі даних';
  }
}

// === РОЗУМНИЙ ПОШУК З КОНТЕКСТОМ ===
async function smartSearchWithContext(
  query: string, 
  data: any[][], 
  headers: string[], 
  context: SearchContext = {}
): Promise<{ results: any[][]; explanation: string; recommendations: string[] }> {
  try {
    if (!data || data.length === 0) {
      return { results: [], explanation: 'Немає даних для пошуку', recommendations: [] };
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

    const aiResponse = completion.choices[0].message.content || '';
    let searchConfig: EnhancedSearchConfig;
    
    try {
      // Спробуємо парсити JSON
      const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        searchConfig = JSON.parse(jsonMatch[0]);
      } else {
        throw new Error('Invalid JSON response');
      }
    } catch (parseError) {
      console.error('Помилка парсингу AI відповіді:', parseError);
      // Fallback до простого пошуку
      searchConfig = {
        searchFields: headers,
        searchConditions: [query.toLowerCase()],
        explanation: 'Простий пошук по всіх полях',
        recommendations: [],
        confidence: 0.5
      };
    }

    // Виконуємо пошук на основі AI аналізу
    const results = performSmartSearch(data, headers, searchConfig);
    
    return {
      results,
      explanation: searchConfig.explanation,
      recommendations: searchConfig.recommendations
    };
  } catch (error) {
    console.error('Помилка розумного пошуку з контекстом:', error);
    return { results: [], explanation: 'Помилка при пошуку', recommendations: [] };
  }
}

// === РОЗУМНИЙ ПОШУК ===
function performSmartSearch(data: any[][], headers: string[], config: EnhancedSearchConfig): any[][] {
  const results: any[][] = [];
  
  for (const row of data) {
    let matches = false;
    let matchScore = 0;
    
    for (const field of config.searchFields) {
      const fieldIndex = getColumnIndex(headers, field);
      if (fieldIndex === -1) continue;
      
      const cellValue = String(row[fieldIndex] || '').toLowerCase();
      
      for (const condition of config.searchConditions) {
        const conditionLower = condition.toLowerCase();
        
        // Точний збіг
        if (cellValue === conditionLower) {
          matchScore += 10;
          matches = true;
        }
        // Частковий збіг
        else if (cellValue.includes(conditionLower)) {
          matchScore += 5;
          matches = true;
        }
        // Збіг по словах
        else if (conditionLower.split(' ').some(word => cellValue.includes(word))) {
          matchScore += 2;
          matches = true;
        }
      }
    }
    
    if (matches) {
      results.push([...row, matchScore]); // Додаємо score для сортування
    }
  }
  
  // Сортуємо за релевантністю
  results.sort((a, b) => (b[b.length - 1] as number) - (a[a.length - 1] as number));
  
  // Видаляємо score з результатів
  return results.map(row => row.slice(0, -1)).slice(0, 20);
}

// === ПОКРАЩЕНІ РЕКОМЕНДАЦІЇ ===
async function generateEnhancedRecommendations(
  data: any[][], 
  headers: string[], 
  userContext: UserContext = {}
): Promise<string[]> {
  try {
    if (!data || data.length === 0) {
      return ['Немає даних для генерації рекомендацій'];
    }

    const prompt = `
Проаналізуй дані та надай персоналізовані рекомендації:

Колонки: ${headers.join(', ')}
Кількість рядків: ${data.length}

Контекст користувача:
- Уподобання: ${userContext.preferences?.join(', ') || 'не вказано'}
- Історія: ${userContext.history?.join(', ') || 'не вказано'}
- Експертиза: ${userContext.expertise || 'не вказано'}

Зразок даних:
${data.slice(0, 10).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай 5 персоналізованих рекомендацій українською мовою у форматі:
1. Рекомендація 1 (з поясненням)
2. Рекомендація 2 (з поясненням)
...

Будь конкретним та практичним. Враховуй контекст користувача.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 500,
      temperature: 0.4
    });

    const response = completion.choices[0].message.content || '';
    const recommendations = response
      .split('\n')
      .filter(line => line.trim().match(/^\d+\./))
      .map(line => line.replace(/^\d+\.\s*/, ''))
      .slice(0, 5);

    return recommendations.length > 0 ? recommendations : ['Немає конкретних рекомендацій'];
  } catch (error) {
    console.error('Помилка генерації покращених рекомендацій:', error);
    return ['Помилка при генерації рекомендацій'];
  }
}

// === ДОПОМІЖНІ ФУНКЦІЇ ===
function getColumnIndex(headers: string[], field: string): number {
  const normalizedField = field.toLowerCase().trim();
  return headers.findIndex(header => 
    header.toLowerCase().includes(normalizedField) || 
    normalizedField.includes(header.toLowerCase())
  );
}

// === ЕКСПОРТ ФУНКЦІЙ ===
export {
  analyzeDataEnhanced,
  smartSearchWithContext,
  performSmartSearch,
  generateEnhancedRecommendations,
  getColumnIndex,
}; 