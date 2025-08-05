/**
 * Утиліти для форматування даних
 * TypeScript версія
 */

interface Metrics {
  [key: string]: number | string;
}

interface Stats {
  total: number;
  success: number;
  errors: number;
  avgTime: number;
  [key: string]: any;
}

class DataFormatters {
  /**
   * Форматування числа з роздільниками
   */
  static formatNumber(num: number | null | undefined, locale: string = 'uk-UA'): string {
    if (num === null || num === undefined) return '—';
    
    const number = parseFloat(num.toString());
    if (isNaN(number)) return '—';
    
    return new Intl.NumberFormat(locale).format(number);
  }

  /**
   * Форматування валюти
   */
  static formatCurrency(amount: number | null | undefined, currency: string = 'UAH', locale: string = 'uk-UA'): string {
    if (amount === null || amount === undefined) return '—';
    
    const number = parseFloat(amount.toString());
    if (isNaN(number)) return '—';
    
    return new Intl.NumberFormat(locale, {
      style: 'currency',
      currency: currency
    }).format(number);
  }

  /**
   * Форматування дати
   */
  static formatDate(date: Date | string | null | undefined, locale: string = 'uk-UA'): string {
    if (!date) return '—';
    
    const dateObj = new Date(date);
    if (isNaN(dateObj.getTime())) return '—';
    
    return new Intl.DateTimeFormat(locale, {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }).format(dateObj);
  }

  /**
   * Форматування часу роботи
   */
  static formatUptime(ms: number): string {
    const days = Math.floor(ms / (1000 * 60 * 60 * 24));
    const hours = Math.floor((ms % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((ms % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((ms % (1000 * 60)) / 1000);

    const parts: string[] = [];
    if (days > 0) parts.push(`${days}д`);
    if (hours > 0) parts.push(`${hours}г`);
    if (minutes > 0) parts.push(`${minutes}хв`);
    if (seconds > 0) parts.push(`${seconds}с`);

    return parts.join(' ') || '0с';
  }

  /**
   * Форматування розміру файлу
   */
  static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Б';
    
    const k = 1024;
    const sizes = ['Б', 'КБ', 'МБ', 'ГБ'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  /**
   * Форматування таблиці для Discord
   */
  static formatTable(data: any[][], headers: string[], maxRows: number = 10): string {
    if (!data || data.length === 0) {
      return 'Немає даних для відображення';
    }

    // Обмежуємо кількість рядків
    const limitedData = data.slice(0, maxRows);
    
    // Створюємо заголовок таблиці
    let table = '| ' + headers.join(' | ') + ' |\n';
    table += '|' + headers.map(() => '---').join('|') + '|\n';
    
    // Додаємо дані
    for (const row of limitedData) {
      const formattedRow = row.map(cell => {
        const cellStr = String(cell || '—');
        return cellStr.length > 20 ? cellStr.substring(0, 17) + '...' : cellStr;
      });
      table += '| ' + formattedRow.join(' | ') + ' |\n';
    }
    
    if (data.length > maxRows) {
      table += `\n... та ще ${data.length - maxRows} рядків`;
    }
    
    return table;
  }

  /**
   * Форматування прогрес-бару
   */
  static formatProgress(current: number, total: number, width: number = 20): string {
    if (total === 0) return '[' + '█'.repeat(width) + '] 0%';
    
    const percentage = Math.min(100, (current / total) * 100);
    const filled = Math.round((percentage / 100) * width);
    const empty = width - filled;
    
    const bar = '█'.repeat(filled) + '░'.repeat(empty);
    return `[${bar}] ${Math.round(percentage)}%`;
  }

  /**
   * Форматування статусу
   */
  static formatStatus(status: string, showIcon: boolean = true): string {
    const statusMap: { [key: string]: { icon: string; color: string } } = {
      'success': { icon: '✅', color: 'green' },
      'error': { icon: '❌', color: 'red' },
      'warning': { icon: '⚠️', color: 'yellow' },
      'info': { icon: 'ℹ️', color: 'blue' },
      'loading': { icon: '⏳', color: 'gray' },
      'pending': { icon: '⏸️', color: 'orange' },
    };
    
    const statusInfo = statusMap[status.toLowerCase()] || { icon: '❓', color: 'gray' };
    return showIcon ? `${statusInfo.icon} ${status}` : status;
  }

  /**
   * Форматування метрик
   */
  static formatMetrics(metrics: Metrics): string {
    const lines: string[] = [];
    
    for (const [key, value] of Object.entries(metrics)) {
      const formattedKey = key.replace(/([A-Z])/g, ' $1').toLowerCase();
      const formattedValue = typeof value === 'number' 
        ? this.formatNumber(value) 
        : String(value);
      
      lines.push(`**${formattedKey}:** ${formattedValue}`);
    }
    
    return lines.join('\n');
  }

  /**
   * Форматування помилки
   */
  static formatError(error: Error | string, includeDetails: boolean = false): string {
    const errorMessage = error instanceof Error ? error.message : error;
    const errorStack = error instanceof Error ? error.stack : '';
    
    let formatted = `❌ **Помилка:** ${errorMessage}`;
    
    if (includeDetails && errorStack) {
      formatted += `\n\`\`\`\n${errorStack}\n\`\`\``;
    }
    
    return formatted;
  }

  /**
   * Форматування часу виконання
   */
  static formatExecutionTime(startTime: number): string {
    const duration = Date.now() - startTime;
    
    if (duration < 1000) {
      return `${duration}мс`;
    } else if (duration < 60000) {
      return `${(duration / 1000).toFixed(2)}с`;
    } else {
      const minutes = Math.floor(duration / 60000);
      const seconds = Math.floor((duration % 60000) / 1000);
      return `${minutes}хв ${seconds}с`;
    }
  }

  /**
   * Форматування списку
   */
  static formatList(items: string[], title: string | null = null, maxItems: number = 10): string {
    if (!items || items.length === 0) {
      return 'Список порожній';
    }
    
    const limitedItems = items.slice(0, maxItems);
    const formattedItems = limitedItems.map((item, index) => `${index + 1}. ${item}`);
    
    let result = formattedItems.join('\n');
    
    if (title) {
      result = `**${title}**\n${result}`;
    }
    
    if (items.length > maxItems) {
      result += `\n... та ще ${items.length - maxItems} елементів`;
    }
    
    return result;
  }

  /**
   * Форматування статистики
   */
  static formatStats(stats: Stats): string {
    const lines: string[] = [];
    
    lines.push(`📊 **Загальна статистика**`);
    lines.push(`• Всього запитів: ${this.formatNumber(stats.total)}`);
    lines.push(`• Успішних: ${this.formatNumber(stats.success)}`);
    lines.push(`• Помилок: ${this.formatNumber(stats.errors)}`);
    lines.push(`• Середній час: ${this.formatExecutionTime(stats.avgTime)}`);
    
    const successRate = stats.total > 0 ? (stats.success / stats.total) * 100 : 0;
    lines.push(`• Відсоток успіху: ${successRate.toFixed(1)}%`);
    
    return lines.join('\n');
  }

  /**
   * Форматування дати та часу
   */
  static formatDateTime(date: Date | string, locale: string = 'uk-UA'): string {
    if (!date) return '—';
    
    const dateObj = new Date(date);
    if (isNaN(dateObj.getTime())) return '—';
    
    return new Intl.DateTimeFormat(locale, {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    }).format(dateObj);
  }

  /**
   * Форматування відсотків
   */
  static formatPercentage(value: number, total: number, decimals: number = 1): string {
    if (total === 0) return '0%';
    
    const percentage = (value / total) * 100;
    return `${percentage.toFixed(decimals)}%`;
  }

  /**
   * Обрізання тексту
   */
  static truncateText(text: string, maxLength: number, suffix: string = '...'): string {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - suffix.length) + suffix;
  }

  /**
   * Капіталізація першої літери
   */
  static capitalizeFirst(text: string): string {
    if (!text) return text;
    return text.charAt(0).toUpperCase() + text.slice(1);
  }
}

export default DataFormatters;
export { DataFormatters }; 