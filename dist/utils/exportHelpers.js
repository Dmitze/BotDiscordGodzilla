"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const xlsx_1 = __importDefault(require("xlsx"));
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const formatters_1 = __importDefault(require("./formatters"));
// Локальні дефолти для експорту (уникаємо залежності від глобальної Config)
const DEFAULT_INCLUDE_METADATA = true;
const DEFAULT_TEMP_FILE_TTL_MS = 60 * 60 * 1000; // 1 година
const DEFAULT_MAX_EXPORT_FILE_SIZE = 50 * 1024 * 1024; // 50 MB
class ExportHelpers {
    constructor() {
        this.tmpDir = './data/tmp';
        this.ensureTmpDir();
    }
    /**
     * Створення тимчасової папки
     */
    ensureTmpDir() {
        if (!fs_1.default.existsSync(this.tmpDir)) {
            fs_1.default.mkdirSync(this.tmpDir, { recursive: true });
        }
    }
    /**
     * Експорт в Excel з метаданими
     */
    async exportToExcel(data, headers, options = {}) {
        const { filename = 'export', sheetName = 'Дані', includeMetadata = DEFAULT_INCLUDE_METADATA, metadata = {} } = options;
        const workbook = xlsx_1.default.utils.book_new();
        // Додаємо метадані як окремий аркуш
        if (includeMetadata) {
            const metadataSheet = this.createMetadataSheet(metadata);
            xlsx_1.default.utils.book_append_sheet(workbook, metadataSheet, 'Метадані');
        }
        // Створюємо основний аркуш з даними
        const worksheet = xlsx_1.default.utils.aoa_to_sheet([headers, ...data]);
        // Налаштовуємо стилі для заголовків
        worksheet['!cols'] = headers.map(() => ({ width: 20 }));
        // Додаємо основний аркуш
        xlsx_1.default.utils.book_append_sheet(workbook, worksheet, sheetName);
        // Генеруємо унікальне ім'я файлу
        const timestamp = Date.now();
        const filePath = path_1.default.join(this.tmpDir, `${filename}_${timestamp}.xlsx`);
        // Зберігаємо файл
        xlsx_1.default.writeFile(workbook, filePath);
        // Записуємо метрики експорту
        const fileSize = fs_1.default.statSync(filePath).size;
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
        const { filename = 'export', includeMetadata = DEFAULT_INCLUDE_METADATA, metadata = {} } = options;
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
        const filePath = path_1.default.join(this.tmpDir, `${filename}_${timestamp}.csv`);
        // Зберігаємо файл
        fs_1.default.writeFileSync(filePath, csvContent, 'utf8');
        // Записуємо метрики експорту
        const fileSize = fs_1.default.statSync(filePath).size;
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
        return xlsx_1.default.utils.aoa_to_sheet(metadataData);
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
        }
        else {
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
        }
        else {
            return await this.exportToExcel(data, headers, exportOptions);
        }
    }
    /**
     * Створення звіту з аналізом даних
     */
    async exportAnalysisReport(data, _headers, analysis, options = {}) {
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
        }
        else {
            return await this.exportToExcel(reportData, ['Параметр', 'Значення'], exportOptions);
        }
    }
    /**
     * Очищення старих файлів
     */
    cleanupOldFiles() {
        try {
            const files = fs_1.default.readdirSync(this.tmpDir);
            const now = Date.now();
            const maxAge = DEFAULT_TEMP_FILE_TTL_MS;
            for (const file of files) {
                const filePath = path_1.default.join(this.tmpDir, file);
                const stats = fs_1.default.statSync(filePath);
                if (now - stats.mtime.getTime() > maxAge) {
                    fs_1.default.unlinkSync(filePath);
                    console.log(`🗑️ Видалено старий файл: ${file}`);
                }
            }
        }
        catch (error) {
            console.error('Помилка очищення файлів:', error);
        }
    }
    /**
     * Запис метрик експорту
     */
    recordExportMetrics(format, fileSize) {
        // Тут можна додати запис в метрики Prometheus
        console.log(`📊 Експорт: ${format}, розмір: ${formatters_1.default.formatFileSize(fileSize)}`);
    }
    /**
     * Отримання статистики експорту
     */
    getExportStats() {
        try {
            const files = fs_1.default.readdirSync(this.tmpDir);
            const stats = {
                totalFiles: files.length,
                totalSize: 0,
                formats: {}
            };
            for (const file of files) {
                const filePath = path_1.default.join(this.tmpDir, file);
                const fileStats = fs_1.default.statSync(filePath);
                const format = path_1.default.extname(file).substring(1);
                stats.totalSize += fileStats.size;
                stats.formats[format] = (stats.formats[format] || 0) + 1;
            }
            return {
                ...stats,
                totalSizeFormatted: formatters_1.default.formatFileSize(stats.totalSize)
            };
        }
        catch (error) {
            console.error('Помилка отримання статистики експорту:', error);
            return { totalFiles: 0, totalSize: 0, formats: {} };
        }
    }
    /**
     * Валідація розміру файлу
     */
    validateFileSize(fileSize) {
        const maxSize = DEFAULT_MAX_EXPORT_FILE_SIZE;
        if (fileSize > maxSize) {
            throw new Error(`Файл занадто великий: ${formatters_1.default.formatFileSize(fileSize)} > ${formatters_1.default.formatFileSize(maxSize)}`);
        }
        return true;
    }
}
exports.default = ExportHelpers;
//# sourceMappingURL=exportHelpers.js.map