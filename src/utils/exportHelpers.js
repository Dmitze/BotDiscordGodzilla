const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const DataFormatters = require('./formatters');
const config = require('../config/Config');

class ExportHelpers {
  constructor() {
    this.tmpDir = './data/tmp';
    this.ensureTmpDir();
  }

  /**
   * Створення тимчасової папки
   */
  ensureTmpDir() {
    if (!fs.existsSync(this.tmpDir)) {
      fs.mkdirSync(this.tmpDir, { recursive: true });
    }
  }

  /**
   * Експорт в Excel з метаданими
   */
  async exportToExcel(data, headers, options = {}) {
    const {
      filename = 'export',
      sheetName = 'Дані',
      includeMetadata = config.export.includeMetadata,
      metadata = {}
    } = options;

    const workbook = XLSX.utils.book_new();

    // Додаємо метадані як окремий аркуш
    if (includeMetadata) {
      const metadataSheet = this.createMetadataSheet(metadata);
      XLSX.utils.book_append_sheet(workbook, metadataSheet, 'Метадані');
    }

    // Створюємо основний аркуш з даними
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...data]);

    // Налаштовуємо стилі для заголовків
    worksheet['!cols'] = headers.map(() => ({ width: 20 }));
    
    // Додаємо основний аркуш
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    // Генеруємо унікальне ім'я файлу
    const timestamp = Date.now();
    const filePath = path.join(this.tmpDir, `${filename}_${timestamp}.xlsx`);

    // Зберігаємо файл
    XLSX.writeFile(workbook, filePath);

    // Записуємо метрики експорту
    const fileSize = fs.statSync(filePath).size;
    this.recordExportMetrics('xlsx', fileSize);

    return {
      filePath,
      fileSize,
      format: 'xlsx',
      rows: data.length,
      columns: headers.length
    };
  }

  /**
   * Експорт в CSV з метаданими
   */
  async exportToCSV(data, headers, options = {}) {
    const {
      filename = 'export',
      includeMetadata = config.export.includeMetadata,
      metadata = {}
    } = options;

    let csvContent = '';

    // Додаємо метадані як коментарі
    if (includeMetadata) {
      csvContent += this.createMetadataCSV(metadata);
    }

    // Додаємо заголовки
    csvContent += headers.map(header => `"${header}"`).join(',') + '\n';

    // Додаємо дані
    for (const row of data) {
      const csvRow = row.map(cell => {
        const cellStr = String(cell || '');
        // Екрануємо лапки в CSV
        return `"${cellStr.replace(/"/g, '""')}"`;
      });
      csvContent += csvRow.join(',') + '\n';
    }

    // Генеруємо унікальне ім'я файлу
    const timestamp = Date.now();
    const filePath = path.join(this.tmpDir, `${filename}_${timestamp}.csv`);

    // Зберігаємо файл
    fs.writeFileSync(filePath, csvContent, 'utf8');

    // Записуємо метрики експорту
    const fileSize = fs.statSync(filePath).size;
    this.recordExportMetrics('csv', fileSize);

    return {
      filePath,
      fileSize,
      format: 'csv',
      rows: data.length,
      columns: headers.length
    };
  }

  /**
   * Створення аркушу з метаданими
   */
  createMetadataSheet(metadata) {
    const metadataData = [
      ['Метадані експорту'],
      [''],
      ['Дата експорту', new Date().toLocaleString('uk-UA')],
      ['Час експорту', new Date().toISOString()],
      ['Версія бота', '2.1.0'],
      [''],
      ['Параметри експорту']
    ];

    // Додаємо додаткові метадані
    for (const [key, value] of Object.entries(metadata)) {
      metadataData.push([key, value]);
    }

    return XLSX.utils.aoa_to_sheet(metadataData);
  }

  /**
   * Створення метаданих для CSV
   */
  createMetadataCSV(metadata) {
    let csv = '';
    
    // Додаємо коментарі з метаданими
    csv += `# Метадані експорту\n`;
    csv += `# Дата експорту: ${new Date().toLocaleString('uk-UA')}\n`;
    csv += `# Час експорту: ${new Date().toISOString()}\n`;
    csv += `# Версія бота: 2.1.0\n`;
    
    // Додаємо додаткові метадані
    for (const [key, value] of Object.entries(metadata)) {
      csv += `# ${key}: ${value}\n`;
    }
    
    csv += '\n';
    return csv;
  }

  /**
   * Експорт результатів пошуку
   */
  async exportSearchResults(results, headers, searchFilters, options = {}) {
    const metadata = {
      'Тип експорту': 'Результати пошуку',
      'Кількість результатів': results.length,
      'Фільтри пошуку': JSON.stringify(searchFilters),
      'Користувач': options.userId || 'Невідомо',
      'Сервер': options.guildId || 'Невідомо'
    };

    const exportOptions = {
      ...options,
      metadata,
      filename: `search_results_${options.userId || 'unknown'}`
    };

    if (options.format === 'csv') {
      return await this.exportToCSV(results, headers, exportOptions);
    } else {
      return await this.exportToExcel(results, headers, exportOptions);
    }
  }

  /**
   * Експорт всієї таблиці
   */
  async exportFullTable(data, headers, options = {}) {
    const metadata = {
      'Тип експорту': 'Повна таблиця',
      'Кількість рядків': data.length,
      'Кількість колонок': headers.length,
      'Користувач': options.userId || 'Невідомо',
      'Сервер': options.guildId || 'Невідомо'
    };

    const exportOptions = {
      ...options,
      metadata,
      filename: `full_table_${options.userId || 'unknown'}`
    };

    if (options.format === 'csv') {
      return await this.exportToCSV(data, headers, exportOptions);
    } else {
      return await this.exportToExcel(data, headers, exportOptions);
    }
  }

  /**
   * Створення звіту з аналізом даних
   */
  async exportAnalysisReport(data, headers, analysis, options = {}) {
    const metadata = {
      'Тип експорту': 'Звіт аналізу',
      'Кількість рядків': data.length,
      'Тип аналізу': analysis.type || 'Загальний',
      'Дата аналізу': new Date().toISOString(),
      'Користувач': options.userId || 'Невідомо'
    };

    // Створюємо дані для звіту
    const reportData = [
      ['Звіт аналізу даних'],
      [''],
      ['Тип аналізу', analysis.type || 'Загальний'],
      ['Дата аналізу', new Date().toLocaleString('uk-UA')],
      ['Кількість записів', data.length],
      [''],
      ['Результати аналізу']
    ];

    // Додаємо результати аналізу
    if (analysis.results) {
      for (const [key, value] of Object.entries(analysis.results)) {
        reportData.push([key, value]);
      }
    }

    const exportOptions = {
      ...options,
      metadata,
      filename: `analysis_report_${options.userId || 'unknown'}`,
      sheetName: 'Аналіз'
    };

    if (options.format === 'csv') {
      return await this.exportToCSV(reportData, ['Параметр', 'Значення'], exportOptions);
    } else {
      return await this.exportToExcel(reportData, ['Параметр', 'Значення'], exportOptions);
    }
  }

  /**
   * Очищення старих файлів
   */
  cleanupOldFiles() {
    try {
      const files = fs.readdirSync(this.tmpDir);
      const now = Date.now();
      const maxAge = config.export.tempFileTtl;

      for (const file of files) {
        const filePath = path.join(this.tmpDir, file);
        const stats = fs.statSync(filePath);
        
        if (now - stats.mtime.getTime() > maxAge) {
          fs.unlinkSync(filePath);
          console.log(`🗑️ Видалено старий файл: ${file}`);
        }
      }
    } catch (error) {
      console.error('Помилка очищення файлів:', error);
    }
  }

  /**
   * Запис метрик експорту
   */
  recordExportMetrics(format, fileSize) {
    // Тут можна додати запис в метрики Prometheus
    console.log(`📊 Експорт: ${format}, розмір: ${DataFormatters.formatFileSize(fileSize)}`);
  }

  /**
   * Отримання статистики експорту
   */
  getExportStats() {
    try {
      const files = fs.readdirSync(this.tmpDir);
      const stats = {
        totalFiles: files.length,
        totalSize: 0,
        formats: {}
      };

      for (const file of files) {
        const filePath = path.join(this.tmpDir, file);
        const fileStats = fs.statSync(filePath);
        const format = path.extname(file).substring(1);

        stats.totalSize += fileStats.size;
        stats.formats[format] = (stats.formats[format] || 0) + 1;
      }

      return {
        ...stats,
        totalSizeFormatted: DataFormatters.formatFileSize(stats.totalSize)
      };
    } catch (error) {
      console.error('Помилка отримання статистики експорту:', error);
      return { totalFiles: 0, totalSize: 0, formats: {} };
    }
  }

  /**
   * Валідація розміру файлу
   */
  validateFileSize(fileSize) {
    const maxSize = config.export.maxFileSize;
    if (fileSize > maxSize) {
      throw new Error(`Файл занадто великий: ${DataFormatters.formatFileSize(fileSize)} > ${DataFormatters.formatFileSize(maxSize)}`);
    }
    return true;
  }
}

module.exports = ExportHelpers; 