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
}

// Результат листинга/поиска с пагинацией
export interface DriveListResult {
  files: DriveFile[];
  nextPageToken?: string;
  total?: number; // если доступно (обычно нет без отдельного запроса)
}
