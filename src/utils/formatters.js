/**
 * Утиліти для форматування даних
 */

class DataFormatters {
  /**
   * Форматування числа з роздільниками
   */
  static formatNumber(num, locale = 'uk-UA') {
    if (num === null || num === undefined) return '—';
    
    const number = parseFloat(num);
    if (isNaN(number)) return '—';
    
    return new Intl.NumberFormat(locale).format(number);
  }

  /**
   * Форматування валюти
   */
  static formatCurrency(amount, currency = 'UAH', locale = 'uk-UA') {
    if (amount === null || amount === undefined) return '—';
    
    const number = parseFloat(amount);
    if (isNaN(number)) return '—';
    
    return new Intl.NumberFormat(locale, {
      style: 'currency',
      currency: currency
    }).format(number);
  }

  /**
   * Форматування дати
   */
  static formatDate(date, locale = 'uk-UA') {
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
  static formatUptime(ms) {
    const days = Math.floor(ms / (1000 * 60 * 60 * 24));
    const hours = Math.floor((ms % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((ms % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((ms % (1000 * 60)) / 1000);

    const parts = [];
    if (days > 0) parts.push(`${days}д`);
    if (hours > 0) parts.push(`${hours}г`);
    if (minutes > 0) parts.push(`${minutes}хв`);
    if (seconds > 0) parts.push(`${seconds}с`);

    return parts.join(' ') || '0с';
  }

  /**
   * Форматування розміру файлу
   */
  static formatFileSize(bytes) {
    if (bytes === 0) return '0 Б';
    
    const k = 1024;
    const sizes = ['Б', 'КБ', 'МБ', 'ГБ'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  /**
   * Форматування таблиці для Discord
   */
  static formatTable(data, headers, maxRows = 10) {
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
        // Обмежуємо довжину комірки
        return cellStr.length > 20 ? cellStr.substring(0, 17) + '...' : cellStr;
      });
      table += '| ' + formattedRow.join(' | ') + ' |\n';
    }
    
    // Додаємо інформацію про кількість рядків
    if (data.length > maxRows) {
      table += `\n*Показано ${maxRows} з ${data.length} записів*`;
    }
    
    return table;
  }

  /**
   * Форматування прогресу
   */
  static formatProgress(current, total, width = 20) {
    const percentage = total > 0 ? (current / total) : 0;
    const filled = Math.round(width * percentage);
    const empty = width - filled;
    
    const filledBar = '█'.repeat(filled);
    const emptyBar = '░'.repeat(empty);
    
    return `${filledBar}${emptyBar} ${Math.round(percentage * 100)}%`;
  }

  /**
   * Форматування статусу
   */
  static formatStatus(status, showIcon = true) {
    const statusMap = {
      'online': { icon: '🟢', text: 'Онлайн' },
      'offline': { icon: '🔴', text: 'Офлайн' },
      'error': { icon: '❌', text: 'Помилка' },
      'warning': { icon: '⚠️', text: 'Попередження' },
      'success': { icon: '✅', text: 'Успішно' },
      'loading': { icon: '⏳', text: 'Завантаження' }
    };
    
    const statusInfo = statusMap[status.toLowerCase()] || { icon: '❓', text: 'Невідомо' };
    
    return showIcon ? `${statusInfo.icon} ${statusInfo.text}` : statusInfo.text;
  }

  /**
   * Форматування метрик
   */
  static formatMetrics(metrics) {
    const lines = [];
    
    for (const [key, value] of Object.entries(metrics)) {
      let formattedValue;
      
      if (typeof value === 'number') {
        formattedValue = this.formatNumber(value);
      } else if (typeof value === 'boolean') {
        formattedValue = value ? '✅' : '❌';
      } else {
        formattedValue = String(value);
      }
      
      lines.push(`${key}: ${formattedValue}`);
    }
    
    return lines.join('\n');
  }

  /**
   * Форматування помилки для користувача
   */
  static formatError(error, includeDetails = false) {
    let message = '❌ Помилка: ';
    
    if (error.message) {
      message += error.message;
    } else {
      message += 'Невідома помилка';
    }
    
    if (includeDetails && error.stack) {
      message += '\n\n**Деталі:**\n```\n' + error.stack.split('\n').slice(0, 3).join('\n') + '\n```';
    }
    
    return message;
  }

  /**
   * Форматування часу виконання
   */
  static formatExecutionTime(startTime) {
    const endTime = Date.now();
    const duration = endTime - startTime;
    
    if (duration < 1000) {
      return `${duration}мс`;
    } else if (duration < 60000) {
      return `${(duration / 1000).toFixed(1)}с`;
    } else {
      return `${(duration / 60000).toFixed(1)}хв`;
    }
  }

  /**
   * Форматування списку для Discord
   */
  static formatList(items, title = null, maxItems = 10) {
    if (!items || items.length === 0) {
      return 'Список порожній';
    }
    
    let result = '';
    if (title) {
      result += `**${title}**\n\n`;
    }
    
    const limitedItems = items.slice(0, maxItems);
    
    for (let i = 0; i < limitedItems.length; i++) {
      result += `${i + 1}. ${limitedItems[i]}\n`;
    }
    
    if (items.length > maxItems) {
      result += `\n*Показано ${maxItems} з ${items.length} елементів*`;
    }
    
    return result;
  }

  /**
   * Форматування статистики
   */
  static formatStats(stats) {
    const lines = [];
    
    if (stats.total) {
      lines.push(`📊 **Всього:** ${this.formatNumber(stats.total)}`);
    }
    
    if (stats.successful !== undefined) {
      lines.push(`✅ **Успішно:** ${this.formatNumber(stats.successful)}`);
    }
    
    if (stats.failed !== undefined) {
      lines.push(`❌ **Помилки:** ${this.formatNumber(stats.failed)}`);
    }
    
    if (stats.percentage !== undefined) {
      lines.push(`📈 **Відсоток успіху:** ${stats.percentage.toFixed(1)}%`);
    }
    
    if (stats.average !== undefined) {
      lines.push(`⏱️ **Середній час:** ${this.formatExecutionTime(stats.average)}`);
    }
    
    return lines.join('\n');
  }
}

module.exports = DataFormatters; 