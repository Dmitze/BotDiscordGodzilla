const { EmbedBuilder } = require('discord.js');

// === Маппинг под реальные заголовки твоей таблицы ===
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

async function getSheetData(range = process.env.SHEET_NAME || 'Аркуш1') {
  const SHEET_ID = process.env.SHEET_ID;
  const GOOGLE_API_KEY = process.env.GOOGLE_API_KEY;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(range)}?key=${GOOGLE_API_KEY}`;
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
    const data = await res.json();
    return data.values || [];
  } catch (err) {
    console.error('⚠️ Не вдалося отримати дані:', err.message);
    return [];
  }
}

// --- КЭШ для поиска и пагинации ---
const searchCache = {};
const CACHE_TTL = 5 * 60 * 1000;
const itemsPerPage = 10;

function cacheSearchResults(userId, results, headers) {
  searchCache[userId] = {
    results,
    headers,
    timestamp: Date.now()
  };
}
function getCachedResults(userId) {
  const cached = searchCache[userId];
  if (!cached || Date.now() - cached.timestamp > CACHE_TTL) {
    return null;
  }
  return cached;
}
function generatePageEmbed(results, page, headers) {
  const totalPages = Math.max(1, Math.ceil(results.length / itemsPerPage));
  const paginatedResults = results.slice(page * itemsPerPage, (page + 1) * itemsPerPage);
  let output = '| Найм. номенклатури | Кількість | Ціна |\n|---------------------|-----------|--------|\n';
  for (let i = 0; i < paginatedResults.length && i < itemsPerPage; i++) {
    const row = paginatedResults[i];
    const name = row[getColumnIndex(headers, 'назва')] || '—';
    const quantity = row[getColumnIndex(headers, 'кількість')] || '—';
    const price = row[getColumnIndex(headers, 'ціна')] || '—';
    output += `| ${name.padEnd(19).slice(0, 19)} | ${quantity} | ${price} |\n`;
  }
  return new EmbedBuilder()
    .setTitle(`🔍 Результати пошуку (${results.length})`)
    .setDescription(`\`\`\`md\n${output}\`\`\``)
    .setFooter({ text: `Сторінка ${page + 1}/${totalPages}` })
    .setColor(3066993);
}

module.exports = {
  getColumnIndex,
  getSheetData,
  cacheSearchResults,
  getCachedResults,
  generatePageEmbed,
  itemsPerPage,
  CACHE_TTL,
  searchCache
};