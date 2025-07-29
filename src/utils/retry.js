/**
 * Утиліта для повторних спроб з експоненційною затримкою
 */

class RetryHandler {
  constructor(maxRetries = 3, baseDelay = 1000) {
    this.maxRetries = maxRetries;
    this.baseDelay = baseDelay;
  }

  /**
   * Виконує функцію з повторними спробами
   * @param {Function} fn - Функція для виконання
   * @param {Object} options - Опції
   * @returns {Promise} Результат виконання
   */
  async execute(fn, options = {}) {
    const {
      maxRetries = this.maxRetries,
      baseDelay = this.baseDelay,
      onRetry = null,
      shouldRetry = this.defaultShouldRetry
    } = options;

    let lastError;
    
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        return await fn();
      } catch (error) {
        lastError = error;
        
        if (attempt === maxRetries || !shouldRetry(error)) {
          throw error;
        }

        const delay = this.calculateDelay(attempt, baseDelay);
        
        if (onRetry) {
          onRetry(error, attempt, delay);
        }

        await this.sleep(delay);
      }
    }
  }

  /**
   * Розраховує затримку з експоненційним збільшенням
   */
  calculateDelay(attempt, baseDelay) {
    return Math.min(baseDelay * Math.pow(2, attempt), 30000); // Максимум 30 секунд
  }

  /**
   * Затримка виконання
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * За замовчуванням повторюємо помилки мережі та 5xx
   */
  defaultShouldRetry(error) {
    // Повторюємо мережеві помилки
    if (error.code === 'ECONNRESET' || 
        error.code === 'ETIMEDOUT' || 
        error.code === 'ENOTFOUND') {
      return true;
    }

    // Повторюємо HTTP помилки 5xx
    if (error.status >= 500 && error.status < 600) {
      return true;
    }

    // Повторюємо помилки Google API з кодом 429 (rate limit)
    if (error.status === 429) {
      return true;
    }

    return false;
  }
}

/**
 * Спеціалізований обробник для Google Sheets API
 */
class GoogleSheetsRetryHandler extends RetryHandler {
  constructor() {
    super(3, 2000); // 3 спроби, базова затримка 2 секунди
  }

  defaultShouldRetry(error) {
    // Повторюємо помилки Google Sheets API
    if (error.status === 429) { // Rate limit
      return true;
    }
    if (error.status >= 500 && error.status < 600) { // Server errors
      return true;
    }
    if (error.code === 'ECONNRESET' || error.code === 'ETIMEDOUT') {
      return true;
    }
    
    return false;
  }
}

/**
 * Спеціалізований обробник для OpenAI API
 */
class OpenAIRetryHandler extends RetryHandler {
  constructor() {
    super(2, 1000); // 2 спроби, базова затримка 1 секунда
  }

  defaultShouldRetry(error) {
    // Повторюємо помилки OpenAI API
    if (error.status === 429) { // Rate limit
      return true;
    }
    if (error.status >= 500 && error.status < 600) { // Server errors
      return true;
    }
    
    return false;
  }
}

module.exports = {
  RetryHandler,
  GoogleSheetsRetryHandler,
  OpenAIRetryHandler
}; 