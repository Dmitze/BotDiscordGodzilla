/**
 * Допоміжні функції для пошуку
 * TypeScript версія
 */

import { EmbedBuilder } from 'discord.js';

// --- Utilities ---
function parseLocaleNumber(v: string): number {
  const s = String(v || '').replace(/\s+/g, '').replace(',', '.');
  const n = parseFloat(s);
  return Number.isNaN(n) ? NaN : n;
}

function truncateEmbed(text: string, max = 3800): string {
  return text.length > max ? text.slice(0, max - 10) + '\n…[truncated]' : text;
}

async function fetchWithTimeout(url: string, ms = 10000, init?: RequestInit, retries = 0): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ms);
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } catch (err) {
    if (retries > 0) {
      const backoff = Math.min(1200, 200 * Math.pow(2, (2 - retries)));
      await new Promise((r) => setTimeout(r, backoff));
      return fetchWithTimeout(url, ms, init, retries - 1);
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

interface SearchCache {
  [userId: string]: {
    results: any[][];
    headers: string[];
    timestamp: number;
  };
}

interface HeaderMap {
  [key: string]: string[];
}

// === Маппинг под реальные заголовки твоей таблицы ===
function getColumnIndex(headers: string[], field: string): number {
  const headerMap: HeaderMap = {
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
    ],
    вартість: [
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

async function getSheetData(range: string = process.env['SHEET_NAME'] || 'Аркуш1'): Promise<any[][]> {
  const SHEET_ID = process.env['SHEET_ID'];
  const GOOGLE_API_KEY = process.env['GOOGLE_API_KEY'];

  if (!SHEET_ID || !GOOGLE_API_KEY) {
    console.error('❌ Відсутні необхідні змінні середовища SHEET_ID або GOOGLE_API_KEY');
    return [];
  }

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(range)}?key=${GOOGLE_API_KEY}`;
  
  try {
    const res = await fetchWithTimeout(url, 10000);
    if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
    const data: any = await res.json();
    return (data && Array.isArray(data.values)) ? data.values : [];
  } catch (err) {
    console.error('⚠️ Не вдалося отримати дані:', err instanceof Error ? err.message : 'Unknown error');
    return [];
  }
}

// --- КЭШ для поиска и пагинации ---
const searchCache: SearchCache = {};
const CACHE_TTL = 5 * 60 * 1000; // 5 хвилин
const itemsPerPage = 10;
const MAX_CACHE_ENTRIES = 100;

function evictIfNeeded() {
  const keys = Object.keys(searchCache);
  if (keys.length > MAX_CACHE_ENTRIES) {
    const first = keys[0] as string;
    if (first) delete searchCache[first];
  }
}

function cacheSearchResults(userId: string, results: any[][], headers: string[]): void {
  searchCache[userId] = {
    results,
    headers,
    timestamp: Date.now()
  };
  evictIfNeeded();
}

function getCachedResults(userId: string): { results: any[][]; headers: string[] } | null {
  const cached = searchCache[userId];
  if (!cached || Date.now() - cached.timestamp > CACHE_TTL) {
    return null;
  }
  return cached;
}

function generatePageEmbed(results: any[][], page: number, headers: string[]): EmbedBuilder {
  const totalPages = Math.max(1, Math.ceil(results.length / itemsPerPage));
  const paginatedResults = results.slice(page * itemsPerPage, (page + 1) * itemsPerPage);
  
  let output = '| Найм. номенклатури | Кількість | Ціна |\n|---------------------|-----------|--------|\n';
  
  for (let i = 0; i < paginatedResults.length && i < itemsPerPage; i++) {
    const row = paginatedResults[i] || [];
    const name = row[getColumnIndex(headers, 'назва')] ?? '—';
    const quantity = row[getColumnIndex(headers, 'кількість')] ?? '—';
    const price = row[getColumnIndex(headers, 'ціна')] ?? '—';
    output += `| ${String(name).padEnd(19).slice(0, 19)} | ${String(quantity)} | ${String(price)} |\n`;
  }

  return new EmbedBuilder()
    .setTitle(`🔍 Результати пошуку (${results.length})`)
    .setDescription(truncateEmbed(`\`\`\`md\n${output}\`\`\``))
    .setFooter({ text: `Сторінка ${page + 1}/${totalPages}` })
    .setColor(3066993);
}

/**
 * Очищення застарілих записів кешу
 */
function cleanupCache(): void {
  const now = Date.now();
  const keysToDelete: string[] = [];

  for (const [userId, cached] of Object.entries(searchCache)) {
    if (now - cached.timestamp > CACHE_TTL) {
      keysToDelete.push(userId);
    }
  }

  keysToDelete.forEach(key => delete searchCache[key]);

  if (keysToDelete.length > 0) {
    console.log(`🧹 Очищено ${keysToDelete.length} застарілих записів кешу`);
  }
}

/**
 * Отримання статистики кешу
 */
function getCacheStats(): { totalEntries: number; oldestEntry: number; newestEntry: number } {
  const entries = Object.values(searchCache);
  const timestamps = entries.map(entry => entry.timestamp);
  
  return {
    totalEntries: entries.length,
    oldestEntry: timestamps.length > 0 ? Math.min(...timestamps) : 0,
    newestEntry: timestamps.length > 0 ? Math.max(...timestamps) : 0,
  };
}

/**
 * Пошук з фільтрами
 */
function searchWithFilters(
  data: any[][], 
  headers: string[], 
  filters: Record<string, string | number>
): any[][] {
  if (!data || data.length === 0) return [];

  return data.filter(row => {
    for (const [field, value] of Object.entries(filters)) {
      const columnIndex = getColumnIndex(headers, field);
      if (columnIndex === -1) continue;

      const cellValue = String(row[columnIndex] || '').toLowerCase();
      const filterValue = String(value).toLowerCase();

      if (!cellValue.includes(filterValue)) {
        return false;
      }
    }
    return true;
  });
}

/**
 * Сортування результатів
 */
function sortResults(
  data: any[][], 
  headers: string[], 
  sortBy: string, 
  order: 'asc' | 'desc' = 'asc'
): any[][] {
  if (!data || data.length === 0) return [];

  const columnIndex = getColumnIndex(headers, sortBy);
  if (columnIndex === -1) return data;

  return [...data].sort((a, b) => {
    const aValue = a[columnIndex] || '';
    const bValue = b[columnIndex] || '';

    // Спроба числового сортування
    const aNum = parseLocaleNumber(String(aValue));
    const bNum = parseLocaleNumber(String(bValue));

    if (!isNaN(aNum) && !isNaN(bNum)) {
      return order === 'asc' ? aNum - bNum : bNum - aNum;
    }

    // Строкове сортування
    const comparison = aValue.localeCompare(bValue);
    return order === 'asc' ? comparison : -comparison;
  });
}

export {
  getColumnIndex,
  getSheetData,
  cacheSearchResults,
  getCachedResults,
  generatePageEmbed,
  cleanupCache,
  getCacheStats,
  searchWithFilters,
  sortResults,
  itemsPerPage,
  CACHE_TTL,
  searchCache,
}; 