// Типы конфигурации для интеграции Google Drive
// Версия 1.0.0

export interface DriveConfig {
  folderId: string;
  pageSize: number; // 5..100
  allowedMime: string[]; // ["*"] = все, либо список допустимых MIME
  fileMaxSizeMb: number; // лимит размера файла для индексации/загрузки
  enableTextIndex: boolean; // включить текстовый индекс (OCR/извлечение текста)
  indexCron: string; // расписание индексатора
  maxConcurrency: number; // 1..10 параллельных задач
  ttlListSec: number; // TTL кэша для листингов
  ttlTextSec: number; // TTL кэша для текста/контента
  ownerAllowlist: string[]; // допустимые владельцы файлов
  hideWebLink: boolean; // скрывать webViewLink в ответах бота
  rateQps?: number; // лимит запросов в секунду к Google API (по умолчанию 5)
  rateBurst?: number; // размер бурста токен‑бакета (по умолчанию 10)
  // Настройки парсинга/извлечения текста
  parseTimeoutMs?: number;
  parseRetryAttempts?: number;
  parseRetryDelayMs?: number;
}

// Нормализованный объект файла Drive для внутреннего использования
export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  size?: number; // bytes
  modifiedTime?: string; // ISO
  owners?: string[]; // display names/emails если доступны
  parents?: string[];
  webViewLink?: string;
  iconLink?: string;
  isShortcut?: boolean;
  shortcutDetails?: {
    targetId: string;
    targetMimeType?: string;
  };
}

// Запрос на листинг/поиск
export interface DriveListQuery {
  folderId: string;
  query?: string; // case-insensitive contains по имени
  mimeIncludes?: string[]; // пусто = любые
  ownerAllowlist?: string[]; // ограничение по владельцу
  pageSize?: number; // 5..100
  pageToken?: string;
  recursive?: boolean;
  maxDepth?: number; // при recursive=true, по умолчанию 20
  // Новые фильтры
  dateFrom?: string; // ISO: modifiedTime >= dateFrom
  dateTo?: string;   // ISO: modifiedTime <= dateTo
  sizeMin?: number;  // bytes: size >= sizeMin
  sizeMax?: number;  // bytes: size <= sizeMax
  // Сортировка
  sortBy?: 'name' | 'modifiedTime' | 'size';
  sortDir?: 'asc' | 'desc';
  // Подсветка изменений относительно предыдущего запроса в рамках сессии
  highlightChanges?: boolean;
  sessionKey?: string; // идентификатор чата/пользователя/контекста
}

// Результат листинга/поиска с пагинацией
export interface DriveListResult {
  files: DriveFile[];
  nextPageToken?: string;
  total?: number; // если доступно (обычно нет без отдельного запроса)
  // Сводка изменений (опционально)
  changes?: {
    addedIds: string[];
    removedIds: string[];
    modified: Array<{
      id: string;
      fields: Array<'name' | 'mimeType' | 'size' | 'modifiedTime' | 'owners' | 'parents' | 'webViewLink'>;
    }>;
  };
}
