/**
 * AI допоміжні функції для Discord AI Assistant Bot
 * TypeScript версія
 */

import OpenAI from 'openai';

// Конфіг: офлайн за замовчуванням (динамічна перевірка)
function isAIEnabled(): boolean {
  return Boolean(process.env['OPENAI_API_KEY']) && process.env['AI_PROVIDER'] !== 'disabled';
}
<<<<<<< HEAD
=======

// Лінива ініціалізація OpenAI при першому зверненні
let _openai: OpenAI | null = null;
function getOpenAI(): OpenAI {
  if (!_openai) {
    _openai = new OpenAI({ apiKey: process.env['OPENAI_API_KEY'] as string });
  }
  return _openai;
}

// Утиліти
function extractJson(text: string): any | null {
  const fenced = text.match(/```json\s*([\s\S]*?)```/i);
  const raw = fenced?.[1] || text.match(/\{[\s\S]*\}/)?.[0];
  if (!raw) return null;
  try { return JSON.parse(raw); } catch { return null; }
}

>>>>>>> 1e192943 (ai: offline-by-default, lazy OpenAI init, robust extractJson; add offline jest tests)

// Лінива ініціалізація OpenAI при першому зверненні
let _openai: OpenAI | null = null;
function getOpenAI(): OpenAI {
  if (!_openai) {
    _openai = new OpenAI({ apiKey: process.env['OPENAI_API_KEY'] as string });
  }
  return _openai;
}

// Утиліти
function extractJson(text: string): any | null {
  const fenced = text.match(/```json\s*([\s\S]*?)```/i);
  const raw = fenced?.[1] || text.match(/\{[\s\S]*\}/)?.[0];
  if (!raw) return null;
  try { return JSON.parse(raw); } catch { return null; }
}

// Інтерфейси
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

// === ВИЗНАЧЕННЯ ТИПІВ ДАНИХ / МЕТРИК ===
function isPercentString(v: unknown): boolean {
  if (typeof v === 'number') return false;
  const s = String(v ?? '').trim();
  return /^[-+]?\d[\d\s.,]*%$/.test(s);
}

function toNumberMaybe(v: unknown): number | null {
  if (v == null) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = String(v).trim();
  if (!s) return null;
  // percentage -> 0..1
  if (isPercentString(s)) {
    const num = s.replace(/%/g, '').replace(/\s/g, '').replace(/,(?=\d{1,2}$)/, '.');
    const n = Number(num);
    return Number.isFinite(n) ? n / 100 : null;
  }
  // localized number
  if (/^[-+]?\d[\d\s.,]*$/.test(s)) {
    const norm = s.replace(/\s/g, '').replace(/,(?=\d{1,2}$)/, '.');
    const n = Number(norm);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function isNumericColumn(data: any[][], colIndex: number, sample: number = 30): boolean {
  const limit = Math.min(sample, data.length);
  let numeric = 0, total = 0;
  for (let i = 0; i < limit; i++) {
    const val = data[i]?.[colIndex];
    const num = toNumberMaybe(val);
    if (num !== null) numeric++;
    total++;
  }
  return total > 0 && numeric / total >= 0.6; // 60%+ числових значень
}

function detectMetricColumns(data: any[][], headers: string[]): { name: string; index: number; kind: 'number' | 'percent' }[] {
  const res: { name: string; index: number; kind: 'number' | 'percent' }[] = [];
  for (let i = 0; i < headers.length; i++) {
    if (!isNumericColumn(data, i)) continue;
    // estimate percent vs number
    let percHits = 0, checked = 0;
    for (let r = 0; r < Math.min(30, data.length); r++) {
      const v = data[r]?.[i];
      if (v != null) {
        checked++;
        if (isPercentString(v)) percHits++;
      }
    }
    const kind: 'number' | 'percent' = checked > 0 && percHits / checked > 0.3 ? 'percent' : 'number';
    res.push({ name: headers[i], index: i, kind });
  }
  return res;
}

// === АГРЕГАТИ ПО КОЛОНКАХ ===
function aggregateColumn(data: any[][], colIndex: number): { count: number; sum: number; min: number | null; max: number | null; avg: number | null } {
  let count = 0;
  let sum = 0;
  let min: number | null = null;
  let max: number | null = null;
  for (const row of data) {
    const num = toNumberMaybe(row?.[colIndex]);
    if (num === null) continue;
    count++;
    sum += num;
    min = min === null ? num : Math.min(min, num);
    max = max === null ? num : Math.max(max, num);
  }
  const avg = count > 0 ? sum / count : null;
  return { count, sum, min, max, avg };
}

function sumColumn(data: any[][], colIndex: number): number { return aggregateColumn(data, colIndex).sum; }
function avgColumn(data: any[][], colIndex: number): number | null { return aggregateColumn(data, colIndex).avg; }
function minColumn(data: any[][], colIndex: number): number | null { return aggregateColumn(data, colIndex).min; }
function maxColumn(data: any[][], colIndex: number): number | null { return aggregateColumn(data, colIndex).max; }

// === КОРОТКІ ПОЯСНЕННЯ ===
function explainColumns(headers: string[], metrics: { name: string; index: number; kind: 'number' | 'percent' }[]): string {
  if (metrics.length === 0) return 'Не знайдено метричних колонок';
  const parts = metrics.map(m => `${m.name} (${m.kind === 'percent' ? '%': 'число'})`);
  return `Виявлено метричні колонки: ${parts.join(', ')}`;
}

function explainTrends(data: any[][], headers: string[], metricIdx: number): string {
  // Проста евристика: порівняти середнє першої та другої половини
  if (data.length < 4) return 'Даних недостатньо для виявлення тренду';
  const mid = Math.floor(data.length / 2);
  const avg1 = avgColumn(data.slice(0, mid), metricIdx) ?? 0;
  const avg2 = avgColumn(data.slice(mid), metricIdx) ?? 0;
  if (avg2 > avg1 * 1.05) return `Спостерігається зростання показника '${headers[metricIdx]}'`;
  if (avg2 < avg1 * 0.95) return `Спостерігається зниження показника '${headers[metricIdx]}'`;
  return `Суттєвих змін у '${headers[metricIdx]}' не виявлено`;
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

    if (!isAIEnabled()) {
      return '⚠️ AI вимкнено (офлайн режим). Доступний лише базовий аналіз.';
    }
    const completion = await getOpenAI().chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 500,
      temperature: 0.3
    });

    return completion.choices?.[0]?.message?.content || '❌ Помилка при аналізі даних';
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

    if (!isAIEnabled()) {
      return { results: [], explanation: 'AI вимкнено (офлайн режим). Використовуйте простий пошук.' };
    }
    const completion = await getOpenAI().chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 300,
      temperature: 0.2
    });

    const aiResponse = completion.choices?.[0]?.message?.content || '';
    let searchConfig: SearchConfig;

    try {
      const parsed = extractJson(aiResponse);
      if (parsed) {
        searchConfig = parsed as SearchConfig;
      } else {
        throw new Error('Invalid JSON response');
      }
    } catch (parseError) {
      console.error('Помилка парсингу AI відповіді:', parseError);
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

    if (!isAIEnabled()) {
      return ['AI вимкнено (офлайн режим). Доступні лише базові рекомендації.'];
    }
    const completion = await getOpenAI().chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 400,
      temperature: 0.4
    });

    const response = completion.choices?.[0]?.message?.content || '';
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
${reportPrompts[reportType] || reportPrompts['general']}:

Колонки: ${headers.join(', ')}
Кількість рядків: ${data.length}

Зразок даних:
${data.slice(0, 15).map((row, i) => `Рядок ${i + 1}: ${row.join(' | ')}`).join('\n')}

Надай аналіз українською мовою з оцінкою впевненості (0-100%).
`;

    if (!isAIEnabled()) {
      return {
        analysis: 'AI вимкнено (офлайн режим). Повертаю базовий звіт.',
        confidence: 0,
        suggestions: []
      };
    }
    const completion = await getOpenAI().chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 600,
      temperature: 0.3
    });

    const response = completion.choices?.[0]?.message?.content || '';

    // Спроба витягти confidence score
    const confidenceMatch = response.match(/впевненість[:\s]*(\d+)%/i);
    const confidence = confidenceMatch ? parseInt(confidenceMatch[1] || '0', 10) : 70;

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
  // metrics & aggregates
  isNumericColumn,
  detectMetricColumns,
  aggregateColumn,
  sumColumn,
  avgColumn,
  minColumn,
  maxColumn,
  explainColumns,
  explainTrends,
};