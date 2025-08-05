/**
 * AI допоміжні функції для Discord AI Assistant Bot
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

interface SearchConfig {
  searchFields: string[];
  searchConditions: string[];
  explanation: string;
}

interface AIAnalysisResult {
  analysis: string;
  confidence: number;
  suggestions: string[];
}

// === AI-АНАЛІЗ ДАНИХ ===
async function analyzeData(data: any[][], headers: string[]): Promise<string> {
  try {
    if (!data || data.length === 0) {
      return '❌ Немає даних для аналізу';
    }

    // Підготовка даних для аналізу
    const analysisData = data.slice(0, 50); // Беремо перші 50 рядків для аналізу
    const dataSummary: DataSummary = {
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

    return completion.choices[0].message.content || '❌ Помилка при аналізі даних';
  } catch (error) {
    console.error('Помилка AI-аналізу:', error);
    return '❌ Помилка при аналізі даних';
  }
}

// === ПРИРОДНОМОВНИЙ ПОШУК ===
async function naturalLanguageSearch(query: string, data: any[][], headers: string[]): Promise<{ results: any[][]; explanation: string }> {
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

    const aiResponse = completion.choices[0].message.content || '';
    let searchConfig: SearchConfig;
    
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
        explanation: 'Простий пошук по всіх полях'
      };
    }

    // Виконуємо пошук на основі AI аналізу
    const results = performSmartSearch(data, headers, searchConfig);
    
    return {
      results,
      explanation: searchConfig.explanation
    };
  } catch (error) {
    console.error('Помилка природномовного пошуку:', error);
    return { results: [], explanation: 'Помилка при пошуку' };
  }
}

// === РОЗУМНИЙ ПОШУК ===
function performSmartSearch(data: any[][], headers: string[], config: SearchConfig): any[][] {
  const results: any[][] = [];
  
  for (const row of data) {
    let matches = false;
    
    for (const field of config.searchFields) {
      const fieldIndex = getColumnIndex(headers, field);
      if (fieldIndex === -1) continue;
      
      const cellValue = String(row[fieldIndex] || '').toLowerCase();
      
      for (const condition of config.searchConditions) {
        if (cellValue.includes(condition.toLowerCase())) {
          matches = true;
          break;
        }
      }
      
      if (matches) break;
    }
    
    if (matches) {
      results.push(row);
    }
  }
  
  return results.slice(0, 20); // Обмежуємо результати
}

// === ГЕНЕРАЦІЯ РЕКОМЕНДАЦІЙ ===
async function generateRecommendations(data: any[][], headers: string[]): Promise<string[]> {
  try {
    if (!data || data.length === 0) {
      return ['Немає даних для генерації рекомендацій'];
    }

    const prompt = `
Проаналізуй дані та надай 5 корисних рекомендацій:

Колонки: ${headers.join(', ')}
Кількість рядків: ${data.length}

Зразок даних:
${data.slice(0, 10).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай рекомендації українською мовою у форматі:
1. Рекомендація 1
2. Рекомендація 2
...

Будь конкретним та практичним.
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 400,
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
    console.error('Помилка генерації рекомендацій:', error);
    return ['Помилка при генерації рекомендацій'];
  }
}

// === ГЕНЕРАЦІЯ РОЗУМНОГО ЗВІТУ ===
async function generateSmartReport(data: any[][], headers: string[], reportType: string = 'general'): Promise<AIAnalysisResult> {
  try {
    if (!data || data.length === 0) {
      return {
        analysis: 'Немає даних для звіту',
        confidence: 0,
        suggestions: []
      };
    }

    const reportPrompts: { [key: string]: string } = {
      general: 'Надай загальний аналіз даних',
      trends: 'Проаналізуй тренди та зміни',
      anomalies: 'Знайди аномалії та викиди',
      summary: 'Створи короткий зміст',
      detailed: 'Надай детальний аналіз'
    };

    const prompt = `
${reportPrompts[reportType] || reportPrompts.general}:

Колонки: ${headers.join(', ')}
Кількість рядків: ${data.length}

Зразок даних:
${data.slice(0, 15).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай аналіз українською мовою з оцінкою впевненості (0-100%).
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 600,
      temperature: 0.3
    });

    const response = completion.choices[0].message.content || '';
    
    // Спроба витягти confidence score
    const confidenceMatch = response.match(/впевненість[:\s]*(\d+)%/i);
    const confidence = confidenceMatch ? parseInt(confidenceMatch[1]) : 70;

    return {
      analysis: response,
      confidence,
      suggestions: await generateRecommendations(data, headers)
    };
  } catch (error) {
    console.error('Помилка генерації звіту:', error);
    return {
      analysis: 'Помилка при генерації звіту',
      confidence: 0,
      suggestions: []
    };
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
  analyzeData,
  naturalLanguageSearch,
  generateRecommendations,
  generateSmartReport,
  performSmartSearch,
  getColumnIndex,
}; 