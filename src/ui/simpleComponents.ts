import {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  type MessageActionRowComponentBuilder,
} from 'discord.js';

export type SimpleBuildIdFn = (args: { action: 'ai' | 'search' | 'ocr' | 'close' | 'page'; page?: number; fileId?: string }) => string;

/**
 * Створення спрощеного інтерфейсу з трьома основними кнопками
 */
export function buildSimpleActionRow(params: {
  buildId: SimpleBuildIdFn;
  showSearch?: boolean;
  showOCR?: boolean;
}): ActionRowBuilder<MessageActionRowComponentBuilder> {
  const { buildId, showSearch = true, showOCR = true } = params;
  
  const row = new ActionRowBuilder<MessageActionRowComponentBuilder>();
  
  // Кнопка AI
  row.addComponents(
    new ButtonBuilder()
      .setCustomId(buildId({ action: 'ai' }))
      .setLabel('🤖 AI Асистент')
      .setStyle(ButtonStyle.Primary)
  );
  
  // Кнопка пошуку (якщо потрібна)
  if (showSearch) {
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(buildId({ action: 'search' }))
        .setLabel('🔍 Пошук у Google Drive')
        .setStyle(ButtonStyle.Secondary)
    );
  }
  
  // Кнопка OCR (якщо потрібна)
  if (showOCR) {
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(buildId({ action: 'ocr' }))
        .setLabel('📝 Розпізнати текст')
        .setStyle(ButtonStyle.Secondary)
    );
  }
  
  return row;
}

/**
 * Створення кнопок навігації для пагінації
 */
export function buildSimplePaginationRow(params: {
  buildId: SimpleBuildIdFn;
  currentPage: number;
  totalPages: number;
  baseAction: 'search' | 'ocr';
  fileId?: string;
}): ActionRowBuilder<MessageActionRowComponentBuilder> {
  const { buildId, currentPage, totalPages, baseAction, fileId } = params;
  
  const row = new ActionRowBuilder<MessageActionRowComponentBuilder>();
  
  // Кнопка "Назад"
  row.addComponents(
    new ButtonBuilder()
      .setCustomId(buildId(fileId !== undefined 
        ? { action: baseAction, page: Math.max(1, currentPage - 1), fileId: fileId }
        : { action: baseAction, page: Math.max(1, currentPage - 1) }))
      .setLabel('⬅️ Назад')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(currentPage === 1)
  );
  
  // Інформація про сторінку
  row.addComponents(
    new ButtonBuilder()
      .setCustomId(buildId({ action: 'page' }))
      .setLabel(`${currentPage}/${totalPages}`)
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(true)
  );
  
  // Кнопка "Далі"
  row.addComponents(
    new ButtonBuilder()
      .setCustomId(buildId(fileId !== undefined
        ? { action: baseAction, page: Math.min(totalPages, currentPage + 1), fileId: fileId }
        : { action: baseAction, page: Math.min(totalPages, currentPage + 1) }))
      .setLabel('Далі ➡️')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(currentPage === totalPages)
  );
  
  return row;
}

/**
 * Створення кнопки закриття
 */
export function buildCloseRow(params: {
  buildId: SimpleBuildIdFn;
}): ActionRowBuilder<MessageActionRowComponentBuilder> {
  const { buildId } = params;
  
  return new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(
    new ButtonBuilder()
      .setCustomId(buildId({ action: 'close' }))
      .setLabel('❌ Закрити')
      .setStyle(ButtonStyle.Danger)
  );
}

/**
 * Функція для розбиття тексту на частини для Discord
 */
export function chunkTextForDiscord(text: string, maxLength: number = 1900): string[] {
  if (text.length <= maxLength) {
    return [text];
  }

  const chunks: string[] = [];
  let currentChunk = '';
  
  // Розбиваємо текст на рядки
  const lines = text.split('\n');
  
  for (const line of lines) {
    // Якщо рядок занадто довгий, розбиваємо його
    if (line.length > maxLength) {
      // Якщо в нас вже є поточний чанк, зберігаємо його
      if (currentChunk) {
        chunks.push(currentChunk);
        currentChunk = '';
      }
      
      // Розбиваємо довгий рядок на частини
      let remainingLine = line;
      while (remainingLine.length > maxLength) {
        // Знаходимо останній пробіл перед лімітом, щоб не розривати слова
        let breakPoint = maxLength;
        const lastSpace = remainingLine.lastIndexOf(' ', maxLength);
        if (lastSpace > 0) {
          breakPoint = lastSpace;
        }
        
        chunks.push(remainingLine.substring(0, breakPoint));
        remainingLine = remainingLine.substring(breakPoint).trim();
      }
      
      // Залишок рядка додаємо до поточного чанку
      if (remainingLine) {
        currentChunk = remainingLine;
      }
      continue;
    }
    
    // Перевіряємо, чи додавання рядка не перевищить ліміт
    if (currentChunk.length + line.length + 1 > maxLength) {
      // Зберігаємо поточний чанк і починаємо новий
      if (currentChunk) {
        chunks.push(currentChunk);
        currentChunk = '';
      }
    }
    
    // Додаємо рядок до поточного чанку
    if (currentChunk) {
      currentChunk += '\n' + line;
    } else {
      currentChunk = line;
    }
  }
  
  // Додаємо останній чанк, якщо він є
  if (currentChunk) {
    chunks.push(currentChunk);
  }
  
  return chunks;
}

/**
 * Функція для форматування таблиць для Discord
 */
export function formatTableForDiscord(tableData: any[][], maxCellLength: number = 50): string[] {
  if (tableData.length === 0) {
    return [''];
  }

  // Визначаємо максимальну ширину для кожної колонки
  const columnWidths: number[] = [];
  // Fix: Check if tableData[0] exists before accessing its length
  if (tableData.length > 0 && tableData[0]) {
    for (let i = 0; i < tableData[0].length; i++) {
      let maxWidth = 0;
      for (const row of tableData) {
        // Fix: Check if row exists and has enough elements
        if (row && i < row.length) {
          const cellLength = String(row[i] || '').length;
          maxWidth = Math.max(maxWidth, Math.min(cellLength, maxCellLength));
        }
      }
      columnWidths.push(Math.min(maxWidth, maxCellLength));
    }
  }

  // Форматуємо таблицю
  let formattedTable = '';
  for (const row of tableData) {
    // Fix: Check if row exists before processing
    if (!row) continue;
    
    let formattedRow = '|';
    for (let i = 0; i < columnWidths.length; i++) {
      const cell = i < row.length ? String(row[i] || '') : '';
      // Обрізаємо довгі значення
      const truncatedCell = cell.length > maxCellLength ? cell.substring(0, maxCellLength - 3) + '...' : cell;
      // Вирівнюємо по лівому краю
      // Fix: Ensure columnWidths[i] is defined before using it
      const width = columnWidths[i] || 0;
      formattedRow += ` ${truncatedCell.padEnd(width)} |`;
    }
    formattedTable += formattedRow + '\n';
    
    // Додаємо роздільник після заголовка
    if (row === tableData[0]) {
      let separator = '|';
      for (const width of columnWidths) {
        // Fix: Ensure width is defined before using it
        separator += ` ${'-'.repeat(width || 0)} |`;
      }
      formattedTable += separator + '\n';
    }
  }

  // Розбиваємо таблицю на чанки
  return chunkTextForDiscord(formattedTable, 1900);
}